#!/usr/bin/env node

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const projectRoot = path.resolve(__dirname, '..');
const backendDir = path.join(projectRoot, 'backend');
const requirementsPath = path.join(backendDir, 'requirements.txt');
const entryPoint = path.join('backend', 'main.py');
const isWindows = process.platform === 'win32';

function getPythonCandidates() {
  const explicit = String(process.env.PYTHON_BIN || process.env.GOV_PYTHON_BIN || '').trim();
  const candidates = [];

  if (explicit) {
    candidates.push(explicit);
  }

  const inActiveVenv = String(process.env.VIRTUAL_ENV || '').trim();
  if (inActiveVenv) {
    candidates.push(
      path.join(inActiveVenv, 'Scripts', 'python.exe'),
      path.join(inActiveVenv, 'bin', 'python3'),
      path.join(inActiveVenv, 'bin', 'python')
    );
  }

  const projectVenv = path.join(projectRoot, '.venv');
  candidates.push(
    path.join(projectVenv, 'Scripts', 'python.exe'),
    path.join(projectVenv, 'bin', 'python3'),
    path.join(projectVenv, 'bin', 'python')
  );

  candidates.push(process.platform === 'win32' ? 'python' : 'python3');
  candidates.push('python');

  return candidates;
}

function resolvePythonBin() {
  const candidates = getPythonCandidates();
  for (const candidate of candidates) {
    if (!candidate) continue;

    const looksLikePath =
      candidate.includes(path.sep) ||
      candidate.includes('/') ||
      candidate.includes('\\') ||
      candidate.endsWith('.exe');

    if (!looksLikePath || fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return process.platform === 'win32' ? 'python' : 'python3';
}

function fail(message) {
  console.error(`[build-python-backend] ${message}`);
  process.exit(1);
}

function info(message) {
  console.log(`[build-python-backend] ${message}`);
}

function ensureRequirementsFile() {
  if (fs.existsSync(requirementsPath)) {
    return;
  }

  const defaultRequirements = ['pyinstaller>=6.0', 'python-docx>=1.1.0', ''].join('\n');
  fs.mkdirSync(backendDir, { recursive: true });
  fs.writeFileSync(requirementsPath, defaultRequirements, 'utf8');
  info(`Created missing ${path.relative(projectRoot, requirementsPath)}`);
}

function getExtraHiddenImports() {
  const raw = String(process.env.PYI_HIDDEN_IMPORTS || '').trim();
  if (!raw) return [];
  return raw
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function getPyInstallerMode() {
  const configured = String(process.env.PYI_MODE || 'onedir').trim().toLowerCase();
  if (configured === 'onefile') return 'onefile';
  return 'onedir';
}

function getAddDataSeparator() {
  return isWindows ? ';' : ':';
}

function defaultDataMappings() {
  const candidates = [
    { source: path.join(projectRoot, 'backend', 'templates'), target: 'backend/templates' },
    { source: path.join(projectRoot, 'backend', 'assets'), target: 'backend/assets' },
    { source: path.join(projectRoot, 'backend', 'models'), target: 'backend/models' },
    { source: path.join(projectRoot, 'src', 'python', 'templates'), target: 'src/python/templates' },
    { source: path.join(projectRoot, 'src', 'python', 'assets'), target: 'src/python/assets' },
    { source: path.join(projectRoot, 'src', 'python', 'models'), target: 'src/python/models' }
  ];

  return candidates.filter((item) => fs.existsSync(item.source));
}

function parseExtraDataMappings() {
  const raw = String(process.env.PYI_ADD_DATA || '').trim();
  if (!raw) return [];

  const mappings = [];
  for (const token of raw.split(',')) {
    const item = token.trim();
    if (!item) continue;
    const idx = item.indexOf('=');
    if (idx <= 0 || idx === item.length - 1) {
      fail(
        `Invalid PYI_ADD_DATA entry "${item}". ` +
        `Use format: sourcePath=targetPath,sourcePath2=targetPath2`
      );
    }
    mappings.push({
      source: path.resolve(projectRoot, item.slice(0, idx).trim()),
      target: item.slice(idx + 1).trim()
    });
  }

  return mappings;
}

function buildPyInstallerArgs() {
  const mode = getPyInstallerMode();
  const args = [
    '--noconfirm',
    mode === 'onefile' ? '--onefile' : '--onedir',
    '--name',
    'mybackend',
    '--paths',
    '.',
    '--hidden-import',
    'src.python.generate_document',
    '--hidden-import',
    'src.python.generate_role',
    entryPoint
  ];

  const addDataSeparator = getAddDataSeparator();
  const dataMappings = [...defaultDataMappings(), ...parseExtraDataMappings()];
  for (const mapping of dataMappings) {
    args.unshift(`${mapping.source}${addDataSeparator}${mapping.target}`);
    args.unshift('--add-data');
  }

  const additional = getExtraHiddenImports();
  for (const hiddenImport of additional) {
    args.unshift(hiddenImport);
    args.unshift('--hidden-import');
  }

  return args;
}

function getExpectedBackendExecutablePath() {
  const mode = getPyInstallerMode();
  if (mode === 'onefile') {
    return path.join(projectRoot, 'dist', isWindows ? 'mybackend.exe' : 'mybackend');
  }

  return path.join(projectRoot, 'dist', 'mybackend', isWindows ? 'mybackend.exe' : 'mybackend');
}

function ensurePythonModules(pythonBin) {
  const checks = [
    { importName: 'PyInstaller', display: 'PyInstaller' },
    { importName: 'docx', display: 'python-docx' }
  ];

  for (const check of checks) {
    const result = spawnSync(pythonBin, ['-c', `import ${check.importName}`], {
      cwd: projectRoot,
      stdio: 'pipe'
    });

    const spawnFailed = !!result.error && typeof result.status !== 'number';
    if (spawnFailed || result.status !== 0) {
      fail(
        `Missing ${check.display} in interpreter "${pythonBin}". ` +
        `Install build deps with: ${pythonBin} -m pip install -r backend/requirements.txt`
      );
    }
  }
}

function main() {
  ensureRequirementsFile();
  const pythonBin = resolvePythonBin();
  ensurePythonModules(pythonBin);

  const args = buildPyInstallerArgs();
  const commandArgs = ['-m', 'PyInstaller', ...args];
  info(`Running: ${pythonBin} ${commandArgs.join(' ')}`);

  const result = spawnSync(pythonBin, commandArgs, {
    cwd: projectRoot,
    stdio: 'inherit'
  });

  const spawnFailed = !!result.error && typeof result.status !== 'number';
  if (spawnFailed) {
    fail(
      `Unable to execute Python build command (${pythonBin} -m PyInstaller): ${result.error.message}. ` +
      `Install build deps with: python -m pip install -r backend/requirements.txt`
    );
  }
  if (result.status !== 0) {
    fail(`PyInstaller failed with exit code ${result.status}`);
  }

  const outputExe = getExpectedBackendExecutablePath();
  if (!fs.existsSync(outputExe)) {
    fail(`Expected output not found: ${outputExe}`);
  }

  info(`Built ${path.relative(projectRoot, outputExe)}`);
}

main();
