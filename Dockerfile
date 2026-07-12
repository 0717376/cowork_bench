FROM ubuntu:22.04
ENV DEBIAN_FRONTEND=noninteractive

# Base system deps
RUN apt-get update && apt-get install -y \
    curl wget git ca-certificates gnupg \
    python3 python3-pip rsync postgresql-client \
    libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 \
    libcups2 libdrm2 libdbus-1-3 libatspi2.0-0 \
    libx11-6 libxcomposite1 libxdamage1 libxext6 \
    libxfixes3 libxrandr2 libgbm1 libxcb1 \
    libxkbcommon0 libpango-1.0-0 libcairo2 libasound2 \
    build-essential libffi-dev libssl-dev \
    libxml2-dev libxslt1-dev zlib1g-dev \
    libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

# uv
RUN curl -LsSf https://astral.sh/uv/install.sh | sh
ENV PATH="/root/.local/bin:$PATH"

# Node.js 22 (bundled npm 10.x is current enough; npm@latest upgrade is
# skipped because it intermittently corrupts arborist deps).
RUN curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y nodejs \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /workspace

# Create venv OUTSIDE /workspace so it is not shadowed by the volume mount
# camel-ai brings: openai, pyyaml, httpx, tiktoken, etc.
RUN uv venv /opt/venv && uv pip install --python /opt/venv/bin/python \
    "camel-ai" \
    "anthropic" \
    "psycopg2-binary" \
    "openpyxl" \
    "python-docx" \
    "python-pptx" \
    "termcolor" \
    "aiofiles" \
    "psutil" \
    "addict" \
    "arxiv" \
    "bibtexparser" \
    "canvasapi" \
    "prompt_toolkit" \
    "strands-agents==1.42.0" \
    "litellm>=1.82.6" \
    "httpx" \
    "beautifulsoup4" \
    "markdownify" \
    "pypdf" \
    "PyPDF2" \
    "pdfplumber"

ENV PATH="/opt/venv/bin:$PATH"
ENV VIRTUAL_ENV="/opt/venv"

# Install Playwright browser
RUN playwright install chromium || true

# Build every Node-based and Python MCP server under /opt/local_servers.
# A subdir is treated as a Node server if it has package.json, and as a
# Python (uv) server if it has pyproject.toml — both can coexist.
COPY local_servers/ /opt/local_servers/
RUN set -e; for dir in /opt/local_servers/*/; do \
    if [ -f "$dir/package.json" ]; then \
        echo "=== NODE $dir ===" && cd "$dir" && npm install && \
        if grep -q '"build"' package.json; then npm run build; fi && \
        cd /workspace; \
    fi; \
done

# uv sync MUST succeed for every Python MCP. A silent `|| true` here let
# missing system deps slip through and only surface when the first task
# using that MCP returned 0 tool calls.
RUN set -e; for dir in /opt/local_servers/*/; do \
    if [ -f "$dir/pyproject.toml" ]; then \
        echo "=== PY $dir ===" && cd "$dir" && uv sync && cd /workspace; \
    fi; \
done

# Copy project code
COPY . .

# Pre-build task runtime env (/workspace/.venv) so per-task `uv sync` is a no-op.
RUN uv sync --frozen 2>/dev/null || uv sync

# Runtime only (after all build-time syncs): uv run uses the baked .venv and
# never hits the network at MCP startup. Fixes transient Connect timeouts.
ENV UV_NO_SYNC=1

CMD ["/bin/bash"]
