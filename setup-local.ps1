$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $root

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    throw "Python 3.10 or later is required. Install Python and run this script again."
}

$pythonVersion = & python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
if ([version]$pythonVersion -lt [version]"3.10") {
    throw "Python 3.10 or later is required; found $pythonVersion."
}

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    & python -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "Could not create the local Python environment." }
}

& .\.venv\Scripts\python.exe -m pip install --disable-pip-version-check -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "Dependency installation failed." }

Write-Host "Local setup complete. Start the app with .\run-local.ps1"
