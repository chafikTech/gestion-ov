#!/usr/bin/env bash
set -euo pipefail

cd /workspace

# Windows Python installed under Wine
WIN_PY="C:\\Program Files\\Python311\\python.exe"

if command -v wine64 >/dev/null 2>&1; then
  WINE_BIN="wine64"
elif command -v wine >/dev/null 2>&1; then
  WINE_BIN="wine"
else
  echo "[win-build] ERROR: neither wine64 nor wine is available"
  exit 1
fi

echo "[win-build] Node: $(node -v)"
echo "[win-build] NPM:  $(npm -v)"
echo "[win-build] Wine: $(command -v "$WINE_BIN")"

# Install JS deps
npm ci

echo "[win-build] Installing backend deps (Windows Python under Wine)..."
"$WINE_BIN" "$WIN_PY" -m pip install --upgrade pip
"$WINE_BIN" "$WIN_PY" -m pip install -r "Z:\\workspace\\backend\\requirements.txt"
"$WINE_BIN" "$WIN_PY" -m pip install pyinstaller

echo "[win-build] Building backend exe (onedir)..."
"$WINE_BIN" "$WIN_PY" -m PyInstaller --noconfirm --onedir --name mybackend "Z:\\workspace\\backend\\main.py"

echo "[win-build] Copy backend bundle into resources/bin/mybackend/..."
rm -rf resources/bin/mybackend
mkdir -p resources/bin
cp -r dist/mybackend resources/bin/mybackend

# Hard guard: block accidental ELF backend from Linux build artifacts
if [[ ! -f "resources/bin/mybackend/mybackend.exe" ]]; then
  echo "[win-build] ERROR: resources/bin/mybackend/mybackend.exe is missing"
  exit 1
fi

echo "[win-build] Building Windows installer..."
xvfb-run -a npm run dist

echo "[win-build] Done. Check dist/ for Windows artifacts."
