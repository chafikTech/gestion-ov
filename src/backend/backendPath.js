const fs = require('node:fs');
const path = require('node:path');
const { app } = require('electron');

function getProjectRoot() {
  if (app && typeof app.getAppPath === 'function') {
    return app.getAppPath();
  }
  return path.resolve(__dirname, '..', '..');
}

function getBackendPath() {
  if (app.isPackaged) {
    const packagedCandidates = [
      path.join(process.resourcesPath, 'bin', 'mybackend', 'mybackend.exe'),
      path.join(process.resourcesPath, 'bin', 'mybackend', 'mybackend'),
      path.join(process.resourcesPath, 'bin', 'mybackend.exe')
    ];

    for (const candidate of packagedCandidates) {
      if (fs.existsSync(candidate)) {
        return candidate;
      }
    }

    return packagedCandidates[0];
  }

  const projectRoot = getProjectRoot();
  const candidates = [
    path.join(projectRoot, 'resources', 'bin', 'mybackend', 'mybackend.exe'),
    path.join(projectRoot, 'resources', 'bin', 'mybackend', 'mybackend'),
    path.join(projectRoot, 'dist', 'mybackend', 'mybackend.exe'),
    path.join(projectRoot, 'dist', 'mybackend', 'mybackend'),
    path.join(projectRoot, 'resources', 'bin', 'mybackend.exe'),
    path.join(projectRoot, 'dist', 'mybackend.exe'),
    path.join(projectRoot, 'resources', 'bin', 'mybackend'),
    path.join(projectRoot, 'dist', 'mybackend')
  ];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return candidates[0];
}

module.exports = {
  getBackendPath
};
