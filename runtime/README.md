# Bundled Python Runtime

This folder is used to bundle Python inside the packaged Electron app.

At runtime, the app looks for Python in:

- `runtime/python/<platform-arch>/venv/...`
- `runtime/python/<platform-arch>/...`

Examples of `<platform-arch>`:

- `win32-x64`
- `linux-x64`
- `darwin-arm64`

Important:

- Runtime must match the target OS.
- `linux-x64` runtime will not work on Windows builds.
- `win32-x64` runtime will not work on Linux builds.

## Quick Setup (current machine target)

Run:

```bash
npm run prepare:python-bundle
```

This creates:

`runtime/python/<current-platform-arch>/venv`

and installs `python-docx` in that venv.

## Using an Existing Venv

```bash
node scripts/prepare-bundled-python.js --venv-dir /path/to/venv --target win32-x64
```

## Build

Then build normally:

```bash
npm run build:win
npm run build:linux
```

`electron-builder` copies `runtime/` via `extraResources`, so packaged apps can use bundled Python without requiring system Python.

## Cross-Build Note (Linux -> Windows)

If you build Windows installers from Linux:

1. Prepare a Windows venv on a Windows machine.
2. Copy it into this project as `runtime/python/win32-x64/venv`.
3. Build Windows installer (`npm run build:win`).
