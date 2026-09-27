$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonPath = Join-Path $projectRoot 'runtime\python.exe'
if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    $pythonPath = Join-Path $projectRoot '.venv\Scripts\python.exe'
}
if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    throw 'The included Python runtime is missing. Install TNG Expense Agent again, or follow the source setup instructions in README.md.'
}
# Never look for Codex, Node, a global Python, or a paid spreadsheet runtime.
[pscustomobject]@{ PythonPath = $pythonPath }
