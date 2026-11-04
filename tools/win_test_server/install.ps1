# Register in the already logged-on user's desktop so repair dialogs are observable.
# Run through ssh win10; no password or SYSTEM/session-0 service is used.
param([string]$Directory = 'C:\vo', [string]$Listen = '127.0.0.1:8765')
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$name = 'VibeOffice Office 2016 Test API'
$exe = Join-Path $Directory 'win_test_server.exe'
if (!(Test-Path -LiteralPath $exe)) { throw "Missing $exe" }
$existing = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
if ($existing -and $existing.State -eq 'Running') { throw 'The API is running. Wait for active/pending=0, stop its task, then replace the executable and reinstall.' }
$user = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$action = New-ScheduledTaskAction -Execute $exe -Argument ('-dir "{0}" -listen "{1}"' -f $Directory, $Listen) -WorkingDirectory $Directory
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $name -Action $action -Principal $principal -Trigger $trigger -Settings $settings -Force | Out-Null
Start-ScheduledTask -TaskName $name
Write-Output "Started $name as $user in $Directory, listening on $Listen"
