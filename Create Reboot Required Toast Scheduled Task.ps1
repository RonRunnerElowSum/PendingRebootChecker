$TaskName = "(MSP) Throw Reboot Required Toast Notification"
$PSFileURL = "'https://raw.githubusercontent.com/RonRunnerElowSum/PendingRebootChecker/Prod-Branch/PRC.ps1'"

$ScheduledTaskCmd = "Invoke-WebRequest -URI $PSFileURL -UseBasicParsing | Invoke-Expression; ThrowToastNotification"
$ScheduledTaskArg = "-WindowStyle Hidden -NoExit -Command `"& {$ScheduledTaskCmd}`""

function CreateSchedTask () {
    $Action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument $ScheduledTaskArg
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    $Principal = New-ScheduledTaskPrincipal -GroupId "BUILTIN\Users"
    Register-ScheduledTask -Action $Action -Settings $Settings -Principal $Principal -TaskName $TaskName -Description "Throws toast notification that requests a restart of Windows." | Out-Null
    if(!(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)){
        Write-Warning "Failed to create schedule task...exiting..."
    }
    else{
        Write-Host "Successfully installed ($TaskName)!"
    }
}

function PunchIt () {
    if(!(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)){
        CreateSchedTask
    }
    else{
        Write-Host "The scheduled task ($TaskName) already exists..."
    }
}
