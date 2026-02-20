const { spawn } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const { app } = require('electron');
const { getBackendPath } = require('./backendPath');
const { getBackendWorkingDirectory, appendBackendLog } = require('./backendRuntime');

const autoInstallStatusByPython = new Map();
let loggedBackendExePath = false;

function normalizePathIfExists(candidate) {
  if (!candidate) return null;
  const value = String(candidate).trim();
  if (!value) return null;
  return fs.existsSync(value) ? value : null;
}

function getRuntimeTag() {
  return `${process.platform}-${process.arch}`;
}

function getBundledPythonCandidates() {
  const runtimeTag = getRuntimeTag();
  const bases = [];

  const appBase = app.getAppPath();
  if (appBase) {
    bases.push(path.join(appBase, 'runtime', 'python', runtimeTag));
  }

  if (app.isPackaged) {
    bases.push(path.join(process.resourcesPath, 'runtime', 'python', runtimeTag));
    bases.push(path.join(process.resourcesPath, 'app.asar.unpacked', 'runtime', 'python', runtimeTag));
  }

  const uniqueBases = Array.from(new Set(bases));
  const candidates = [];
  for (const base of uniqueBases) {
    candidates.push(
      path.join(base, 'venv', 'Scripts', 'python.exe'),
      path.join(base, 'venv', 'Scripts', 'python3.exe'),
      path.join(base, 'python', 'python.exe'),
      path.join(base, 'python.exe'),
      path.join(base, 'venv', 'bin', 'python3'),
      path.join(base, 'venv', 'bin', 'python'),
      path.join(base, 'bin', 'python3'),
      path.join(base, 'bin', 'python')
    );
  }
  return candidates;
}

function resolveBundledPythonBinary() {
  const explicitBundled = normalizePathIfExists(process.env.GOV_BUNDLED_PYTHON_BIN);
  if (explicitBundled) {
    return explicitBundled;
  }

  for (const candidate of getBundledPythonCandidates()) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }
  return null;
}

function resolveConfiguredPythonBinary() {
  const configured = (process.env.GOV_PYTHON_BIN || process.env.PYTHON_BIN || '').toString().trim();
  if (!configured) return null;

  const looksLikePath =
    configured.includes(path.sep) ||
    configured.includes('/') ||
    configured.includes('\\') ||
    configured.endsWith('.exe');
  if (!looksLikePath) {
    return configured;
  }

  return fs.existsSync(configured) ? configured : null;
}

function resolvePythonBinary() {
  const configured = resolveConfiguredPythonBinary();
  if (configured) {
    return { pythonBin: configured, source: 'configured' };
  }

  const bundled = resolveBundledPythonBinary();
  if (bundled) {
    return { pythonBin: bundled, source: 'bundled' };
  }

  return {
    pythonBin: process.platform === 'win32' ? 'python' : 'python3',
    source: 'system'
  };
}

function resolvePythonScriptPath(relativeScriptPath) {
  const devPath = path.join(app.getAppPath(), relativeScriptPath);
  if (!app.isPackaged) {
    return devPath;
  }

  const unpackedPath = path.join(process.resourcesPath, 'app.asar.unpacked', relativeScriptPath);
  if (fs.existsSync(unpackedPath)) {
    return unpackedPath;
  }

  const asarPath = path.join(process.resourcesPath, 'app.asar', relativeScriptPath);
  if (fs.existsSync(asarPath)) {
    return asarPath;
  }

  return devPath;
}

function isMissingPythonDocxError(message = '') {
  const msg = String(message || '').toLowerCase();
  return (
    msg.includes('missing python dependency for word generation') ||
    msg.includes("no module named 'docx'") ||
    msg.includes('no module named docx') ||
    msg.includes("no module named 'lxml'") ||
    msg.includes('no module named lxml')
  );
}

function validateGeneratorResponse(parsed, stderr, code, sourceLabel) {
  const hasDocxPayload = !!(parsed && parsed.docxFilePath && parsed.docxFileName);
  const isSuccess = parsed && (parsed.success === true || (parsed.success === undefined && hasDocxPayload));

  if (!parsed || !isSuccess) {
    const message =
      (parsed && (parsed.message || parsed.error))
        ? (parsed.message || parsed.error)
        : (stderr || `${sourceLabel} a échoué (exit ${code})`);
    throw new Error(message);
  }

  if (!parsed.docxFilePath || !parsed.docxFileName) {
    throw new Error(`${sourceLabel}: réponse invalide (docx manquant)`);
  }

  const docxPath = String(parsed.docxFilePath || '');
  if (!docxPath.toLowerCase().endsWith('.docx')) {
    throw new Error(`${sourceLabel}: réponse invalide (chemin DOCX)`);
  }
  if (!fs.existsSync(docxPath)) {
    throw new Error(`${sourceLabel}: fichier DOCX introuvable`);
  }
}

function runBackendJsonOnce(backendExePath, requestPayload) {
  return new Promise((resolve, reject) => {
    if (!loggedBackendExePath) {
      console.log(`[backend] python bridge executable path: ${backendExePath}`);
      loggedBackendExePath = true;
    }

    const backendCwd = getBackendWorkingDirectory(backendExePath);
    appendBackendLog(`REQ start exe="${backendExePath}" cwd="${backendCwd}"`);

    const child = spawn(backendExePath, [], {
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
      cwd: backendCwd
    });

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (data) => {
      stdout += data.toString('utf8');
    });

    child.stderr.on('data', (data) => {
      const text = data.toString('utf8');
      stderr += text;
      const cleaned = text.trim();
      if (cleaned) {
        appendBackendLog(`REQ stderr: ${cleaned}`);
      }
    });

    child.on('error', (error) => {
      appendBackendLog(
        `REQ spawn error exe="${backendExePath}" cwd="${backendCwd}" ` +
        `code=${error.code || 'UNKNOWN'} message="${error.message}"`
      );
      reject(
        new Error(
          `Impossible de lancer le backend exécutable (${backendExePath}). ` +
          `Code=${error.code || 'UNKNOWN'} Message=${error.message}`
        )
      );
    });

    child.on('close', (code) => {
      appendBackendLog(`REQ close exit=${code}`);
      const trimmed = stdout.trim();
      let parsed;
      try {
        parsed = trimmed ? JSON.parse(trimmed) : null;
      } catch (error) {
        appendBackendLog(
          `REQ invalid JSON exit=${code} stdout="${trimmed.slice(0, 800)}" stderr="${stderr.slice(0, 800)}"`
        );
        reject(
          new Error(
            `Backend Python exécutable: sortie invalide (exit ${code}). ` +
            `stdout=${trimmed.slice(0, 500)} stderr=${stderr.slice(0, 500)}`
          )
        );
        return;
      }

      try {
        validateGeneratorResponse(parsed, stderr, code, 'Backend Python exécutable');
        appendBackendLog(
          `REQ success docx="${parsed.docxFilePath || ''}" file="${parsed.docxFileName || ''}"`
        );
        resolve(parsed);
      } catch (validationError) {
        appendBackendLog(`REQ validation failed: ${validationError.message}`);
        reject(validationError);
      }
    });

    child.stdin.write(JSON.stringify(requestPayload));
    child.stdin.end();
  });
}

function runPythonScriptJsonOnce(pythonBin, scriptPath, payload) {
  return new Promise((resolve, reject) => {
    const child = spawn(pythonBin, [scriptPath], {
      stdio: ['pipe', 'pipe', 'pipe']
    });

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (data) => {
      stdout += data.toString('utf8');
    });

    child.stderr.on('data', (data) => {
      stderr += data.toString('utf8');
    });

    child.on('error', (error) => {
      reject(new Error(`Impossible de lancer Python (${pythonBin}): ${error.message}`));
    });

    child.on('close', (code) => {
      const trimmed = stdout.trim();
      let parsed;
      try {
        parsed = trimmed ? JSON.parse(trimmed) : null;
      } catch (error) {
        reject(
          new Error(
            `Générateur Python: sortie invalide (exit ${code}). stdout=${trimmed.slice(0, 500)} stderr=${stderr.slice(0, 500)}`
          )
        );
        return;
      }

      try {
        validateGeneratorResponse(parsed, stderr, code, 'Générateur Python');
        resolve(parsed);
      } catch (validationError) {
        reject(validationError);
      }
    });

    child.stdin.write(JSON.stringify(payload));
    child.stdin.end();
  });
}

function runCommand(pythonBin, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(pythonBin, args, {
      stdio: ['ignore', 'pipe', 'pipe']
    });

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (data) => {
      stdout += data.toString('utf8');
    });

    child.stderr.on('data', (data) => {
      stderr += data.toString('utf8');
    });

    child.on('error', (error) => {
      reject(new Error(`Commande Python introuvable (${pythonBin}): ${error.message}`));
    });

    child.on('close', (code) => {
      if (code === 0) {
        resolve({ stdout, stderr, code });
        return;
      }
      reject(
        new Error(
          `Commande échouée (exit ${code}): ${pythonBin} ${args.join(' ')}; stderr=${stderr.trim().slice(0, 400)}`
        )
      );
    });
  });
}

async function installPythonDocx(pythonBin, { preferBundledInstall = false } = {}) {
  const attempts = preferBundledInstall
    ? [
        ['-m', 'pip', 'install', 'python-docx'],
        ['-m', 'pip', 'install', '--break-system-packages', 'python-docx']
      ]
    : [
        ['-m', 'pip', 'install', '--user', 'python-docx'],
        ['-m', 'pip', 'install', '--user', '--break-system-packages', 'python-docx'],
        ['-m', 'pip', 'install', 'python-docx']
      ];

  let lastError = null;
  for (const args of attempts) {
    try {
      await runCommand(pythonBin, args);
      return;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error('Installation automatique de python-docx impossible');
}

async function runWithPythonFallback(relativeScriptPath, payload) {
  const { pythonBin, source } = resolvePythonBinary();
  const scriptPath = resolvePythonScriptPath(relativeScriptPath);

  if (!fs.existsSync(scriptPath)) {
    throw new Error(`Python script introuvable: ${scriptPath}`);
  }

  try {
    return await runPythonScriptJsonOnce(pythonBin, scriptPath, payload);
  } catch (error) {
    if (!isMissingPythonDocxError(error.message)) {
      throw error;
    }

    const status = autoInstallStatusByPython.get(pythonBin) || { attempted: false, succeeded: false };
    if (status.attempted && !status.succeeded) {
      throw new Error(
        `${error.message}\n` +
        `python-docx est manquant pour ${pythonBin}. ` +
        `Installez-le manuellement avec l'une de ces commandes:\n` +
        `- ${pythonBin} -m pip install python-docx\n` +
        `- ${pythonBin} -m pip install --user python-docx`
      );
    }

    if (!status.attempted) {
      status.attempted = true;
      autoInstallStatusByPython.set(pythonBin, status);
      try {
        await installPythonDocx(pythonBin, { preferBundledInstall: source === 'bundled' });
        status.succeeded = true;
        autoInstallStatusByPython.set(pythonBin, status);
      } catch (installError) {
        status.succeeded = false;
        autoInstallStatusByPython.set(pythonBin, status);
        throw new Error(
          `${error.message}\n` +
          `Installation automatique de python-docx échouée. ` +
          `Installez manuellement avec l'une de ces commandes:\n` +
          `- ${pythonBin} -m pip install python-docx\n` +
          `- ${pythonBin} -m pip install --user python-docx\n` +
          `Détail: ${installError.message}`
        );
      }
    }

    return await runPythonScriptJsonOnce(pythonBin, scriptPath, payload);
  }
}

async function runPythonJson(relativeScriptPath, payload) {
  const backendExePath = getBackendPath();
  if (fs.existsSync(backendExePath)) {
    return runBackendJsonOnce(backendExePath, {
      command: 'generate',
      script: relativeScriptPath,
      payload
    });
  }

  if (app.isPackaged) {
    throw new Error(
      `Backend exécutable introuvable: ${backendExePath}. ` +
      `Le package doit inclure resources/bin/mybackend (onedir) ou resources/bin/mybackend.exe via extraResources.`
    );
  }

  console.warn(
    `[backend] backend exécutable introuvable en mode dev (${backendExePath}). ` +
    `Fallback vers le mode script Python.`
  );
  return runWithPythonFallback(relativeScriptPath, payload);
}

module.exports = {
  getBackendPath,
  resolvePythonBinary,
  resolvePythonScriptPath,
  runPythonJson
};
