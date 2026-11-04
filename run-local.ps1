$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $root
$streamlit = Join-Path $root ".venv\Scripts\streamlit.exe"

if (-not (Test-Path $streamlit)) {
    throw "Local dependencies are not installed. Run .\setup-local.ps1 first."
}

& $streamlit run app.py --server.address 127.0.0.1 --server.port 8501
if ($LASTEXITCODE -ne 0) { throw "The local validation app stopped with exit code $LASTEXITCODE." }
