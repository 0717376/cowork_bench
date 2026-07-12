"""
CAMEL ModelFactory wrapper.
Provider is chosen via MODEL_PROVIDER; credentials via MODEL_API_KEY / MODEL_API_URL.
"""
import os
from camel.models import ModelFactory
from camel.types import ModelPlatformType
from configs.global_configs import global_configs

# provider name → (ModelPlatformType, api_key_fn, default_base_url)
_PROVIDER_MAP = {
    "aihubmix":   (ModelPlatformType.OPENAI_COMPATIBLE_MODEL,
                   lambda: global_configs.aihubmix_key,
                   "https://aihubmix.com/v1"),
    "openai":     (ModelPlatformType.OPENAI,
                   lambda: global_configs.openai_official_key, None),
    "anthropic":  (ModelPlatformType.ANTHROPIC,
                   lambda: global_configs.anthropic_official_key, None),
    "deepseek":   (ModelPlatformType.DEEPSEEK,
                   lambda: global_configs.deepseek_official_key, None),
    "openrouter": (ModelPlatformType.OPENROUTER,
                   lambda: global_configs.openrouter_key, None),
    "qwen":       (ModelPlatformType.QWEN,
                   lambda: global_configs.qwen_official_key, None),
    "gemini":     (ModelPlatformType.GEMINI,
                   lambda: global_configs.google_official_key, None),
    "openai_compatible": (ModelPlatformType.OPENAI_COMPATIBLE_MODEL,
                          lambda: os.environ.get("MODEL_API_KEY", ""), None),
}


def resolve_provider(provider: str = None) -> str:
    """Single source of truth for which model provider to use.

    Priority: MODEL_PROVIDER env > the `provider` arg (config / --provider).
    Returns a provider key (e.g. 'openai_compatible') or a raw CAMEL
    ModelPlatformType name, or None if nothing is set.
    """
    return os.environ.get("MODEL_PROVIDER") or provider


def build_model(model_name: str, provider: str):
    """Build a CAMEL BaseModelBackend.

    The provider is resolved via resolve_provider() (MODEL_PROVIDER env > the
    `provider` arg) — a single source of truth, so the selected platform always
    matches the provider reported in logs.
    Credentials MODEL_API_KEY / MODEL_API_URL override the per-provider defaults.
    """
    provider = resolve_provider(provider)
    if not provider:
        raise ValueError("No model provider set (use MODEL_PROVIDER or --provider).")

    key = provider.lower()
    if key in _PROVIDER_MAP:
        platform, key_fn, default_url = _PROVIDER_MAP[key]
    else:
        # Escape hatch: accept a raw CAMEL ModelPlatformType name.
        try:
            platform = ModelPlatformType[provider.upper()]
        except KeyError:
            raise ValueError(
                f"Unknown provider '{provider}'. Supported: {list(_PROVIDER_MAP.keys())}"
            )
        key_fn = lambda: ""
        default_url = None

    api_key = os.environ.get("MODEL_API_KEY") or key_fn()
    url = os.environ.get("MODEL_API_URL") or default_url

    kwargs = dict(model_platform=platform, model_type=model_name, api_key=api_key)
    if url:
        # Anthropic SDK uses base_url without /v1 suffix; others need /v1
        if platform == ModelPlatformType.ANTHROPIC:
            kwargs["url"] = url.rstrip("/").rstrip("/v1")
        else:
            kwargs["url"] = url.rstrip("/") + ("/v1" if not url.rstrip("/").endswith("/v1") else "")

    return ModelFactory.create(**kwargs)
