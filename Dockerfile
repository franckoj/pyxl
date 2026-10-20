FROM python:3.11-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    RUSTUP_HOME=/usr/local/rustup \
    CARGO_HOME=/usr/local/cargo \
    PATH=/opt/venv/bin:/usr/local/cargo/bin:$PATH \
    VIRTUAL_ENV=/opt/venv \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Install build dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    pkg-config \
    git \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Install Rust toolchain with clippy and rustfmt
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y --default-toolchain stable \
    && rustup component add clippy rustfmt \
    && chmod -R a+w /usr/local/rustup /usr/local/cargo

# Create virtual environment and install dependencies
RUN python3 -m venv /opt/venv && \
    pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir \
    "maturin>=1.5.0" \
    "pytest>=7.4.0" \
    "openpyxl>=3.1.2" \
    "psutil>=5.9.0" \
    "tabulate>=0.9.0" \
    "matplotlib>=3.7.0" \
    "python-calamine>=0.9.0"

WORKDIR /workspace

CMD ["/bin/bash"]
