$ErrorActionPreference = 'Stop'
try {
    $projectRoot = Split-Path -Parent $PSScriptRoot
    $runtime = & (Join-Path $PSScriptRoot 'Set up runtime.ps1')
    & $runtime.PythonPath -B -E (Join-Path $projectRoot 'app\agent.py') --ensure-month
    exit $LASTEXITCODE
} catch {
    Write-Error $_
    exit 1
}
