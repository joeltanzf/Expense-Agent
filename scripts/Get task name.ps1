# A folder-specific name prevents this copy from replacing another installation's task.
$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd('\').ToLowerInvariant()
if ($env:TNG_AGENT_INSTALL) {$projectRoot = [IO.Path]::GetFullPath($env:TNG_AGENT_INSTALL).TrimEnd('\').ToLowerInvariant()}
$sha = [Security.Cryptography.SHA256]::Create()
try {
    $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($projectRoot))
    $suffix = ([BitConverter]::ToString($hash)).Replace('-', '').Substring(0, 10).ToLowerInvariant()
} finally { $sha.Dispose() }
Write-Output ('TNG Expense Agent monthly sheets - ' + $suffix)
