param(
    [string]$InstallRoot = "$env:LOCALAPPDATA\Programs\Gestion des Ouvriers Occasionnels"
)

$ErrorActionPreference = "Stop"

function Write-Step([string]$Message) {
    Write-Host "[bundle-python] $Message"
}

function Require-Path([string]$PathValue, [string]$Hint) {
    if (-not (Test-Path -LiteralPath $PathValue)) {
        throw "Path not found: $PathValue`n$Hint"
    }
}

Write-Step "Install root: $InstallRoot"
Require-Path $InstallRoot "Confirm the app is installed, or pass -InstallRoot with the correct folder."

$resourcesDir = Join-Path $InstallRoot "resources"
Require-Path $resourcesDir "Expected resources folder under the installed app."

$runtimeDir = Join-Path $resourcesDir "runtime\python\win32-x64"
$venvDir = Join-Path $runtimeDir "venv"
$venvPython = Join-Path $venvDir "Scripts\python.exe"

New-Item -ItemType Directory -Path $runtimeDir -Force | Out-Null

$pythonCommand = $null
if (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonCommand = "py -3"
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonCommand = "python"
} else {
    throw "Python was not found on this machine. Install Python 3 first, then rerun this script."
}

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Step "Creating venv at $venvDir"
    if ($pythonCommand -eq "py -3") {
        & py -3 -m venv $venvDir
    } else {
        & python -m venv $venvDir
    }
}

Require-Path $venvPython "Failed to create bundled Python venv."

Write-Step "Upgrading pip"
& $venvPython -m pip install --upgrade pip

Write-Step "Installing python-docx"
& $venvPython -m pip install python-docx

Write-Step "Verifying imports"
& $venvPython -c "import docx, lxml; print('ok')"

Write-Step "Bundled Python is ready."
Write-Step "Restart the app and generate documents again."
