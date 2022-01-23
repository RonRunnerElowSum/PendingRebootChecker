function RestartMachine () {
    $CurrentMonthYear = Get-Date -Format MMyyyy
    Write-PRCLog "Reboot request approved by $Env:USERNAME"
    Move-Item -Path "C:\Windows\Temp\MSP\Logs\PendingRebootChecker\PRCLog-$CurrentMonthYear.log" -Destination "C:\Windows\Temp\MSP\Logs\PendingRebootChecker\_Archive\PRCLog-$CurrentMonthYear.log" | Out-Null
    Restart-Computer -Force
}

function RebootDeny () {
    if($Form.ishandlecreated){$Form.Close()}
    if($RebootConfForm.ishandlecreated){$RebootConfForm.Close()}
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
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $RebootConfForm = New-Object System.Windows.Forms.Form
    $RebootConfForm.Text = ' IT Maintenance'
    $RebootConfForm.Size = New-Object System.Drawing.Size(700,175)
    $RebootConfForm.StartPosition = 'CenterScreen'
    $RebootConfForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
    $RebootConfForm.ControlBox = $False
    $RebootConfForm.Topmost = $True
    
    $RebootConfYesButton = New-Object System.Windows.Forms.Button
    $RebootConfYesButton.Location = New-Object System.Drawing.Size(200,60)
    $RebootConfYesButton.Size = New-Object System.Drawing.Size(350,50)
    $RebootConfYesButton.Text = "Yes, all work has been saved and it's okay to restart"
    $RebootConfYesButton.Add_Click({RestartMachine})
    
    $RebootConfCancelButton = New-Object System.Windows.Forms.Button
    $RebootConfCancelButton.Location = New-Object System.Drawing.Size(560,60)
    $RebootConfCancelButton.Size = New-Object System.Drawing.Size(80,50)
    $RebootConfCancelButton.Text = "Cancel"
    $RebootConfCancelButton.Add_Click({RebootDeny})
    
    $RebootConfPRCLabel = New-Object System.Windows.Forms.Label
    $RebootConfPRCLabel.Location = New-Object System.Drawing.Point(10,20)
    $RebootConfPRCLabel.AutoSize = $True
    $Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Bold)
    $RebootConfPRCLabel.Font = $Font
    $RebootConfPRCLabel.Text = "Are you sure you want to restart? This will close all open applications. Save all of your work before restarting."
    $RebootConfPRCLabel.ForeColor = 'Black'
    
    $RebootConfForm.Controls.Add($RebootConfYesButton)
    $RebootConfForm.Controls.Add($RebootConfCancelButton)
    $RebootConfForm.Controls.Add($RebootConfPRCLabel)
    $RebootConfForm.ShowDialog()
}

function PunchIt () {
    Write-PRCLog "Starting..."
    $PendingRebootStatus = Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    $WMIInfo = Get-WMIObject -Class Win32_OperatingSystem
    $LastBootTime = $WMIInfo.ConvertToDateTime($WMIInfo.LastBootUpTime)
    $SysUpTime = (Get-Date) - $LastBootTime
    if(($PendingRebootStatus -eq "True") -or (7 -lt ($SysUpTime.Days))){
        Write-PRCLog "$Env:ComputerName has a pending reboot..."
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $Form = New-Object System.Windows.Forms.Form
        $Form.Text = ' IT Maintenance'
        $Form.Size = New-Object System.Drawing.Size(600,175)
        $Form.StartPosition = 'CenterScreen'
        $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
        $Form.ControlBox = $False
        $Form.Topmost = $True
        
        $YesButton = New-Object System.Windows.Forms.Button
        $YesButton.Location = New-Object System.Drawing.Size(350,60)
        $YesButton.Size = New-Object System.Drawing.Size(90,50)
        $YesButton.Text = "Yes, restart"
        $YesButton.Add_Click({RebootConf})
        
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Size(450,60)
        $CancelButton.Size = New-Object System.Drawing.Size(80,50)
        $CancelButton.Text = "Cancel"
        $CancelButton.Add_Click({RebootDeny})
        
        $PRCLabel = New-Object System.Windows.Forms.Label
        $PRCLabel.Location = New-Object System.Drawing.Point(10,20)
        $PRCLabel.AutoSize = $True
        $Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Bold)
        $PRCLabel.Font = $Font
        $PRCLabel.Text = "Your computer needs to restart in order to finishing installing updates. Restart now?"
        $PRCLabel.ForeColor = 'Black'
        
        $Form.Controls.Add($YesButton)
        $Form.Controls.Add($CancelButton)
        $Form.Controls.Add($PRCLabel)
        $Form.ShowDialog()
    }
    else{
        Write-PRCLog "$env:ComputerName is not currently in a pending reboot state..."
    }
}
