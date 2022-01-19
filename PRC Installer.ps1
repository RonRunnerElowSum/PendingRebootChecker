$TaskName = "(MSP) Pending Reboot Checker"
$PSFileURL = "'https://raw.githubusercontent.com/RonRunnerElowSum/PendingRebootChecker/Prod-Branch/PRC.ps1'"

$ScheduledTaskCmd = "Invoke-WebRequest -URI $PSFileURL -UseBasicParsing | Invoke-Expression; PunchIt"
$ScheduledTaskArg = "-WindowStyle Hidden -Command `"& {$ScheduledTaskCmd}`""

function CreateSchedTask () {
    $Action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument $ScheduledTaskArg
    $Trigger = @()
    $Trigger += New-ScheduledTaskTrigger -AtLogon
    $Trigger += New-ScheduledTaskTrigger -Daily -At 12pm
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -WakeToRun -StartWhenAvailable
    $Principal = New-ScheduledTaskPrincipal -GroupId "BUILTIN\Users"
    Register-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -TaskName $TaskName -Description "Monitors for pending reboots." | Out-Null
    if(!(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)){
        Write-Warning "Failed to create schedule task...exiting..."
    }
    else{
        Write-Host "Successfully installed Pending Reboot Checker!"
    }
}

function PunchIt () {
    if(!(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)){
        CreateSchedTask
    }
    else{
        Write-Host "Pending Reboot Checker scheduled task already exists..."
    }
}
