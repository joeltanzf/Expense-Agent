$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$null = & (Join-Path $PSScriptRoot 'Set up runtime.ps1')
$taskName = & (Join-Path $PSScriptRoot 'Get task name.ps1')
$userId = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$launcherPath = Join-Path $PSScriptRoot 'Run monthly check.ps1'
if ($env:TNG_AGENT_INSTALL) {$launcherPath = Join-Path $env:TNG_AGENT_INSTALL 'TNG Expense Agent.exe'}
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing -and -not (($existing.Actions.Arguments -join ' ').Contains($launcherPath)) -and -not (($existing.Actions.Execute -join ' ').Equals($launcherPath,[StringComparison]::OrdinalIgnoreCase))) {
    throw 'A different folder already owns this scheduled task. No task was changed.'
}
$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$action = New-ScheduledTaskAction -Execute $powershell -Argument ('-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $launcherPath + '"') -WorkingDirectory $projectRoot
if ($launcherPath.EndsWith('.exe')) {
    $action = New-ScheduledTaskAction -Execute $launcherPath -Argument '--ensure-month' -WorkingDirectory $env:TNG_AGENT_INSTALL
}
$daily = New-ScheduledTaskTrigger -Daily -At '00:05'
$logon = New-ScheduledTaskTrigger -AtLogOn -User $userId
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited
$null = Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($daily, $logon) -Settings $settings -Principal $principal -Description ('Creates monthly sheets for ' + $projectRoot) -Force
Write-Output ('Automatic monthly sheet checks enabled: ' + $taskName)
