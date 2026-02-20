FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive

RUN dpkg --add-architecture i386 && \
  apt-get update && apt-get install -y \
  curl ca-certificates git zip unzip \
  python3 python3-pip python3-venv \
  wine wine64 wine32 \
  xvfb \
  libgtk-3-0 libnss3 libxss1 libasound2t64 libgbm1 libdrm2 libx11-xcb1 \
  && rm -rf /var/lib/apt/lists/*

# Node (use 20 by default; change if needed)
RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
  apt-get update && apt-get install -y nodejs && \
  rm -rf /var/lib/apt/lists/*

WORKDIR /workspace

# Install Windows Python under Wine (Python 3.11 for compatibility)
RUN mkdir -p /opt/winpython && \
  curl -L -o /opt/winpython/python-3.11.9-amd64.exe https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe && \
  WINE_BIN="$(command -v wine64 || true)" && \
  if [ -z "$WINE_BIN" ]; then WINE_BIN="$(command -v wine || true)"; fi && \
  if [ -z "$WINE_BIN" ]; then echo "Wine binary not found (wine64/wine)"; exit 1; fi && \
  xvfb-run -a "$WINE_BIN" /opt/winpython/python-3.11.9-amd64.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0 && \
  rm -f /opt/winpython/python-3.11.9-amd64.exe
