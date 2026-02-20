#!/usr/bin/env node

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

function fail(message) {
  console.error(`[prepare-bundled-python] ${message}`);
  process.exit(1);
}

function info(message) {
  console.log(`[prepare-bundled-python] ${message}`);
}

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === '--python') {
      args.python = argv[i + 1];
      i += 1;
      continue;
    }
    if (token === '--target') {
      args.target = argv[i + 1];
      i += 1;
      continue;
    }
    if (token === '--venv-dir') {
      args.venvDir = argv[i + 1];
      i += 1;
      continue;
    }
    if (token === '--skip-install') {
      args.skipInstall = true;
      continue;
    }
    if (token === '--help' || token === '-h') {
      args.help = true;
      continue;
    }
    fail(`Unknown option: ${token}`);
  }
  return args;
}

function printHelp() {
  console.log(
    [
      'Prepare bundled Python runtime for Electron packaging.',
      '',
      'Usage:',
      '  node scripts/prepare-bundled-python.js [--python <python-bin>] [--target <platform-arch>] [--skip-install]',
      '  node scripts/prepare-bundled-python.js --venv-dir <path-to-existing-venv> [--target <platform-arch>]',
      '',
      'Examples:',
      '  node scripts/prepare-bundled-python.js',
      '  node scripts/prepare-bundled-python.js --python python3.11',
      '  node scripts/prepare-bundled-python.js --venv-dir C:\\python_runtime_venv --target win32-x64',
      '',
      'Notes:',
      '  - Default target is current platform+arch (for example linux-x64).',
      '  - Runtime is written to runtime/python/<target>/venv.',
      '  - Without --venv-dir, the script creates a new venv and installs python-docx.',
      '  - Use --skip-install if the machine is offline; install python-docx in that venv later.'
    ].join('\n')
  );
}

function run(cmd, args, { cwd } = {}) {
  const result = spawnSync(cmd, args, {
    cwd,
    stdio: 'inherit'
  });
  if (result.error) {
    fail(`Failed to run "${cmd} ${args.join(' ')}": ${result.error.message}`);
  }
  if (result.status !== 0) {
    fail(`Command failed (${result.status}): ${cmd} ${args.join(' ')}`);
  }
}

function detectTarget() {
  return `${process.platform}-${process.arch}`;
}

function getVenvPythonPath(venvDir) {
  if (process.platform === 'win32') {
    return path.join(venvDir, 'Scripts', 'python.exe');
  }
  return path.join(venvDir, 'bin', 'python3');
}

function copyVenv(sourceVenvDir, destinationVenvDir) {
  if (!fs.existsSync(sourceVenvDir)) {
    fail(`Source venv not found: ${sourceVenvDir}`);
  }
  fs.rmSync(destinationVenvDir, { recursive: true, force: true });
  fs.mkdirSync(path.dirname(destinationVenvDir), { recursive: true });
  fs.cpSync(sourceVenvDir, destinationVenvDir, { recursive: true });
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printHelp();
    return;
  }

  const projectRoot = path.resolve(__dirname, '..');
  const target = String(args.target || detectTarget()).trim();
  if (!target) {
    fail('Invalid target');
  }
  if (!args.venvDir && target !== detectTarget()) {
    fail(
      `Target ${target} differs from current runtime ${detectTarget()}. ` +
      'Use --venv-dir with a prebuilt venv from that target machine.'
    );
  }

  const runtimeRoot = path.join(projectRoot, 'runtime', 'python', target);
  const runtimeVenvDir = path.join(runtimeRoot, 'venv');

  if (args.venvDir) {
    const sourceVenvDir = path.resolve(args.venvDir);
    info(`Copying existing venv from ${sourceVenvDir}`);
    copyVenv(sourceVenvDir, runtimeVenvDir);
    info(`Bundled runtime ready at ${runtimeVenvDir}`);
    return;
  }

  const pythonBin = String(args.python || (process.platform === 'win32' ? 'python' : 'python3')).trim();
  if (!pythonBin) {
    fail('Invalid python binary');
  }

  fs.mkdirSync(runtimeRoot, { recursive: true });
  fs.rmSync(runtimeVenvDir, { recursive: true, force: true });

  info(`Creating venv for target ${target} using ${pythonBin}`);
  run(pythonBin, ['-m', 'venv', runtimeVenvDir], { cwd: projectRoot });

  const venvPython = getVenvPythonPath(runtimeVenvDir);
  if (!fs.existsSync(venvPython)) {
    fail(`Venv Python not found after creation: ${venvPython}`);
  }

  if (!args.skipInstall) {
    info('Upgrading pip in bundled venv');
    run(venvPython, ['-m', 'pip', 'install', '--upgrade', 'pip'], { cwd: projectRoot });

    info('Installing python-docx in bundled venv');
    run(venvPython, ['-m', 'pip', 'install', 'python-docx'], { cwd: projectRoot });
  } else {
    info('Skipping package install (--skip-install)');
  }

  info(`Bundled runtime ready at ${runtimeVenvDir}`);
  info(`Build app normally (for example: npm run build:win)`);
}

main();
