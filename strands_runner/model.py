"""LiteLLM model factory for the Strands runner.

Configuration via env vars (matches the CAMEL runner's main.py convention so a
single .env serves both runners):

    LLM_BASE_URL    OpenAI-compatible endpoint (e.g. https://your-endpoint/v1)
    LLM_API_KEY     API key for that endpoint
    LLM_MODEL       Model id (required; MODEL_NAME is accepted as a fallback)
"""
from __future__ import annotations

import asyncio
import logging
import os

import litellm
from strands.models.litellm import LiteLLMModel
from strands.types.exceptions import ContextWindowOverflowException

logger = logging.getLogger(__name__)
logging.getLogger("strands.models.openai").setLevel(logging.ERROR)


# OpenRouter says "endpoint's maximum context length" while LiteLLM's detector
# only matches "model's maximum context length" — so the error is mapped to
# BadRequestError instead of ContextWindowExceededError, and Strands'
# ConversationManager never trims. We re-raise as ContextWindowOverflowException.
_OVERFLOW_MARKERS = (
    "this endpoint's maximum context length is",
    "this model's maximum context length is",
    "model's maximum context limit",
    "is longer than the model's context length",
    "input tokens exceed the configured limit",
    "exceeds the maximum number of tokens allowed",
)


def _is_context_overflow(err: BaseException) -> bool:
    msg = str(err).lower()
    return any(m in msg for m in _OVERFLOW_MARKERS)


_RETRYABLE_STREAM_ERRORS = (
    litellm.APIConnectionError,
    litellm.Timeout,
    litellm.ServiceUnavailableError,
    litellm.InternalServerError,
    litellm.exceptions.MidStreamFallbackError,
)


class _LiteLLMModelWithRetry(LiteLLMModel):
    """Three cross-cutting concerns on top of LiteLLMModel:

    1. Re-raise upstream context-overflow as ContextWindowOverflowException so
       the conversation manager can trim history.
    2. Retry transient upstream failures, but only before any chunks are yielded
       (retrying mid-stream would duplicate output).
    3. Idle-timeout each chunk read so a stalled SSE upstream can't block forever.
    """

    _MAX_ATTEMPTS = 3
    _BACKOFF_S = (0.5, 1.5)
    _IDLE_TIMEOUT_S = float(os.environ.get("LLM_STREAM_IDLE_TIMEOUT", "300"))

    def format_request(self, *args, **kwargs):
        """Coerce assistant/tool message content from content-block lists to a
        plain string.

        Strands' OpenAI formatter serializes message content as a list of typed
        blocks (e.g. ``[{"type": "text", "text": ...}]``). Some strict
        OpenAI-compatible gateways reject array content for ``assistant``/``tool``
        roles with 422 ``content: Input should be a valid string`` (``user`` arrays
        are accepted). Flatten text blocks into one string for those two roles; leave
        ``user``/``system`` untouched so image content keeps working.
        """
        request = super().format_request(*args, **kwargs)
        for msg in request.get("messages", []):
            if msg.get("role") in ("assistant", "tool") and isinstance(msg.get("content"), list):
                msg["content"] = "".join(
                    b.get("text", "") for b in msg["content"] if isinstance(b, dict)
                )
        return request

    async def stream(self, *args, **kwargs):
        for attempt in range(self._MAX_ATTEMPTS):
            yielded = False
            try:
                inner = super().stream(*args, **kwargs).__aiter__()
                while True:
                    try:
                        chunk = await asyncio.wait_for(
                            inner.__anext__(), timeout=self._IDLE_TIMEOUT_S
                        )
                    except StopAsyncIteration:
                        return
                    yielded = True
                    yield chunk
            except asyncio.TimeoutError as e:
                if yielded or attempt == self._MAX_ATTEMPTS - 1:
                    raise RuntimeError(
                        f"LLM stream stalled (no token in {self._IDLE_TIMEOUT_S:.0f}s)"
                    ) from e
                delay = self._BACKOFF_S[attempt]
                logger.warning(
                    "LLM stream idle on attempt %d/%d, retrying in %.1fs",
                    attempt + 1, self._MAX_ATTEMPTS, delay,
                )
                await asyncio.sleep(delay)
            except litellm.BadRequestError as e:
                if _is_context_overflow(e):
                    raise ContextWindowOverflowException(e) from e
                raise
            except _RETRYABLE_STREAM_ERRORS as e:
                if yielded or attempt == self._MAX_ATTEMPTS - 1:
                    raise
                delay = self._BACKOFF_S[attempt]
                logger.warning(
                    "LLM transient %s on attempt %d/%d, retrying in %.1fs: %s",
                    type(e).__name__, attempt + 1, self._MAX_ATTEMPTS, delay, e,
                )
                await asyncio.sleep(delay)


def make_model(model_id: str | None = None) -> LiteLLMModel:
    """Build the LiteLLM model from env. See module docstring for env vars."""
    resolved = model_id or os.environ.get("LLM_MODEL") or os.environ.get("MODEL_NAME")
    base_url = os.environ.get("LLM_BASE_URL") or os.environ.get("MODEL_API_URL")
    api_key = os.environ.get("LLM_API_KEY") or os.environ.get("MODEL_API_KEY")

    if not resolved:
        raise ValueError("LLM_MODEL (or MODEL_NAME) must be set")
    if not base_url:
        raise ValueError("LLM_BASE_URL (or MODEL_API_URL) must be set")
    if not api_key:
        raise ValueError("LLM_API_KEY (or MODEL_API_KEY) must be set")

    # Some internal gateways sit behind a TLS-inspecting zero-trust proxy that
    # presents a cert our container CA bundle doesn't trust (curl exit 60 /
    # "unable to get local issuer certificate"). Opt out per-endpoint via
    # LLM_SSL_VERIFY=false. litellm honors the GLOBAL flag (url_utils reads
    # `litellm.ssl_verify`); the per-call kwarg is ignored by this provider.
    if os.environ.get("LLM_SSL_VERIFY", "true").strip().lower() in ("0", "false", "no", "off"):
        litellm.ssl_verify = False
        logger.warning("LLM_SSL_VERIFY off — TLS verification disabled for %s", base_url)

    ctx_window = int(os.environ.get("LLM_CONTEXT_WINDOW", "262144"))
    max_tokens = int(os.environ.get("LLM_MAX_TOKENS", "32768"))
    # Sampling profile: "clean" (the default) builds a plain OpenAI-compatible
    # request — safe for any hosted provider. "vllm" additionally sends the
    # vLLM-only knobs (top_k/min_p/repetition_penalty/chat_template_kwargs) +
    # presence_penalty for self-hosted vLLM endpoints (e.g. Qwen); hosted
    # providers reject or silently drop those knobs, so opt in explicitly.
    profile = os.environ.get("LLM_PARAM_PROFILE", "clean").strip().lower()
    reasoning_effort = os.environ.get("LLM_REASONING_EFFORT", "").strip()

    if profile == "vllm":
        # Qwen thinking-mode preset. temp=0.6 is the floor for thinking — below
        # that the model collapses to short CoTs. top_k/min_p/repetition_penalty
        # are vLLM-only so they go through extra_body. preserve_thinking carries
        # thinking traces across turns for KV reuse.
        params = {
            "temperature": float(os.environ.get("LLM_TEMPERATURE", "0.6")),
            "top_p": 0.95,
            # Qwen loops/degenerates (e.g. "!!!" floods) under the official
            # precise-coding preset (presence_penalty=0). Default to the general
            # thinking value 1.5; tune via env without a rebuild.
            "presence_penalty": float(os.environ.get("LLM_PRESENCE_PENALTY", "1.5")),
            "max_tokens": max_tokens,
            "extra_body": {
                "top_k": 20,
                "min_p": 0.0,
                "repetition_penalty": float(os.environ.get("LLM_REPETITION_PENALTY", "1.0")),
                "chat_template_kwargs": {"preserve_thinking": True},
            },
        }
        # Some models gate a hidden reasoning pass on reasoning_effort (no-op
        # where unsupported). Goes through extra_body: litellm rejects
        # reasoning_effort as a top-level param for openai-compatible providers
        # (UnsupportedParamsError), but passes extra_body through raw.
        if reasoning_effort:
            params["extra_body"]["reasoning_effort"] = reasoning_effort
    else:
        # Clean OpenAI-compatible profile (Anthropic / MiniMax / generic hosted).
        params = {"temperature": float(os.environ.get("LLM_TEMPERATURE", "1.0")),
                  "max_tokens": max_tokens}
        top_p = os.environ.get("LLM_TOP_P", "").strip()
        if top_p:
            params["top_p"] = float(top_p)
        if reasoning_effort:
            # OpenRouter's unified reasoning param (Anthropic thinking / others).
            params["extra_body"] = {"reasoning": {"effort": reasoning_effort}}

    # OpenRouter provider routing — pin upstream provider/quantization via env.
    # LLM_OR_PROVIDER: comma-list of provider tags (e.g. "minimax/fp8"), goes to
    # provider.order; LLM_OR_ALLOW_FALLBACKS=false locks routing to those only.
    or_provider = os.environ.get("LLM_OR_PROVIDER", "").strip()
    if or_provider:
        prov = {"order": [p.strip() for p in or_provider.split(",") if p.strip()]}
        fb = os.environ.get("LLM_OR_ALLOW_FALLBACKS", "").strip().lower()
        if fb:
            prov["allow_fallbacks"] = fb in ("1", "true", "yes", "on")
        params.setdefault("extra_body", {})["provider"] = prov

    return _LiteLLMModelWithRetry(
        client_args={"api_key": api_key, "base_url": base_url},
        model_id=f"openai/{resolved}",
        context_window_limit=ctx_window,
        params=params,
    )
