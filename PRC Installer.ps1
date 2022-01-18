$PRCps1URL = "https://raw.githubusercontent.com/RonRunnerElowSum/PendingRebootChecker/Prod-Branch/PRC.ps1"
$PRCLaunchString = "Invoke-WebRequest -URI $PRCps1URL -UseBasicParsing | Invoke-Expression; PunchIt"

function CreateSchedTask () {
    $Action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument "-WindowStyle Hidden $PRCLaunchString"
    $Trigger = @()
    $Trigger += New-ScheduledTaskTrigger -AtLogon
    $Trigger += New-ScheduledTaskTrigger -Daily -At 12pm
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -WakeToRun -StartWhenAvailable
    $Principal = New-ScheduledTaskPrincipal -UserId $Env:Username -LogonType Interactive
    Register-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -TaskName "(MSP) Pending Reboot Checker" -Description "Monitors for pending reboots." | Out-Null
    if(!(Get-ScheduledTask -TaskName "(MSP) Pending Reboot Checker" -ErrorAction SilentlyContinue)){
        Write-Warning "Failed to create schedule task...exiting..."
        EXIT
    }
    else{
        Write-Host "Successfully installed Pending Reboot Checker!"
        EXIT
    }
}

function PunchIt () {
    if(!(Get-ScheduledTask -TaskName "(MSP) Pending Reboot Checker" -ErrorAction SilentlyContinue)){
        CreateSchedTask
    }
    else{
        Write-Host "Pending Reboot Checker scheduled task already exists..."
    }
}
