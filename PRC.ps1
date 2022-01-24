function RestartMachine () {
    $CurrentMonthYear = Get-Date -Format MMyyyy
    Write-PRCLog "Reboot request approved by $Env:USERNAME"
    Move-Item -Path "C:\Windows\Temp\MSP\Logs\PendingRebootChecker\PRCLog-$CurrentMonthYear.log" -Destination "C:\Windows\Temp\MSP\Logs\PendingRebootChecker\_Archive\PRCLog-$CurrentMonthYear.log" | Out-Null
    Restart-Computer -Force
}

function RebootDeny () {
    Write-PRCLog "Reboot request denied by $Env:USERNAME"
}

function Write-PRCLog ($PRCLogEntryValue) {
    $CurrentMonthYear = Get-Date -Format MMyyyy
    if(!(Test-Path -Path "C:\Windows\Temp")){New-Item -Path "C:\Windows" -Name "Temp" -ItemType "Directory" | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\MSP")){New-Item -Path "C:\Windows\Temp" -Name "MSP" -ItemType "Directory" | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\MSP\Logs")){New-Item -Path "C:\Windows\Temp\MSP" -Name "Logs" -ItemType "Directory" | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\MSP\Logs\PendingRebootChecker")){New-Item -Path "C:\Windows\Temp\MSP\Logs" -Name "PendingRebootChecker" -ItemType "Directory" | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\MSP\Logs\PendingRebootChecker\_Archive")){New-Item -Path "C:\Windows\Temp\MSP\Logs\PendingRebootChecker" -Name "_Archive" -ItemType "Directory" | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\MSP\Logs\PendingRebootChecker\PRCLog-$CurrentMonthYear.log")){New-Item -Path "C:\Windows\Temp\MSP\Logs\PendingRebootChecker" -Name "PRCLog-$CurrentMonthYear.log" -ItemType "File" | Out-Null}
    Add-Content -Path "C:\Windows\Temp\MSP\Logs\PendingRebootChecker\PRCLog-$CurrentMonthYear.log" -Value "$(Get-Date) -- $PRCLogEntryValue"
}

function RebootConf () {
    Write-PRCLog "Initial reboot request approved...confirming it's safe to reboot..."
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
    if([Microsoft.VisualBasic.Interaction]::MsgBox('Are you sure it is OK to restart?  This will close all open files and applications.  Save all of your work before restarting.', 'YesNo,MsgBoxSetForeground,Exclamation', 'IT Maintenance') -eq "No"){
        RebootDeny
    }
    else{
        RestartMachine
    }
}

function ThrowToastNotification () {
    Write-PRCLog "Throwing reboot required toast notification..."
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
    $ToastNotification = New-Object System.Windows.Forms.NotifyIcon
    $ToastNotification.Icon = [System.Drawing.SystemIcons]::Information
    $ToastNotification.BalloonTipText = "Your computer needs to restart in order to finishing installing updates. Please restart at your earliest convenience."
    $ToastNotification.BalloonTipTitle = "Reboot Required"
    $ToastNotification.BalloonTipIcon = "Warning"
    $ToastNotification.Visible = $True
    $ToastNotification.ShowBalloonTip(50000)
    Unregister-Event -SourceIdentifier click_event -ErrorAction SilentlyContinue
    Register-ObjectEvent $ToastNotification BalloonTipClicked -SourceIdentifier click_event -Action {
        Write-PRCLog "Toast notification clicked...prompting to restart..."
        #[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        if([Microsoft.VisualBasic.Interaction]::MsgBox('Your computer needs to restart in order to finishing installing updates.  Restart now?', 'YesNo,MsgBoxSetForeground,Information', 'IT Maintenance') -eq "No"){
            RebootDeny
        }
        else{
            RebootConf
        }
    } | Out-Null
    Wait-Event -Timeout 10 -SourceIdentifier click_event > $null
    Unregister-Event -SourceIdentifier click_event -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 10
    $ToastNotification.Dispose()
}

function PunchIt () {
    Write-PRCLog "Starting..."
    $PendingRebootStatus = Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    $WMIInfo = Get-WMIObject -Class Win32_OperatingSystem
    $LastBootTime = $WMIInfo.ConvertToDateTime($WMIInfo.LastBootUpTime)
    $SysUpTime = (Get-Date) - $LastBootTime
    if(($PendingRebootStatus -eq "True") -or (7 -lt ($SysUpTime.Days))){
        Write-PRCLog "$Env:ComputerName has a pending reboot..."
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        if([Microsoft.VisualBasic.Interaction]::MsgBox('Your computer needs to restart in order to finishing installing updates.  Restart now?', 'YesNo,MsgBoxSetForeground,Information', 'IT Maintenance') -eq "No"){
            RebootDeny
        }
        else{
            RebootConf
        }
    }
    else{
        Write-PRCLog "$env:ComputerName is not currently in a pending reboot state..."
    }
}
