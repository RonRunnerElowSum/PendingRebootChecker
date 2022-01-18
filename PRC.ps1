function RestartMachine () {
    if(!(Test-Path -Path "C:\Windows\Temp")){New-Item -Path "C:\Windows" -Name "Temp" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\PRC")){New-Item -Path "C:\Windows\Temp" -Name "PRC" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\PRC\_Archive")){New-Item -Path "C:\Windows\Temp\PRC" -Name "_Archive" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Approved.txt")){
        New-Item -Path "C:\Windows\Temp\PRC" -Name "Pending-Reboot-Checker-Approved.txt" -ItemType "file" | Out-Null
    }
    $DateTime = (Get-Date)
    $ShortDateTime = (Get-Date -Format MMddyyyyHHmm)
    Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Approved.txt" "Reboot request approved by $Env:USERNAME on $DateTime"
    Move-Item -Path "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" -Destination "C:\Windows\Temp\PRC\_Archive\Pending-Reboot-Checker-Denies_$ShortDateTime.txt" | Out-Null
    Restart-Computer -Force
}

function RebootDeny () {
    if($Form.ishandlecreated){$Form.Close()}
    if($RebootConfForm.ishandlecreated){$RebootConfForm.Close()}
    if(!(Test-Path -Path "C:\Windows\Temp")){New-Item -Path "C:\Windows" -Name "Temp" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\PRC")){New-Item -Path "C:\Windows\Temp" -Name "PRC" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\PRC\_Archive")){New-Item -Path "C:\Windows\Temp\PRC" -Name "_Archive" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null}
    if(!(Test-Path -Path "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt")){
        $DateTime = (Get-Date)
        New-Item -Path "C:\Windows\Temp\PRC" -Name "Pending-Reboot-Checker-Denies.txt" -ItemType "file" | Out-Null
        Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "1 --- Reboot request denied by $Env:USERNAME on $DateTime"
    }
    else{
        if($Form.ishandlecreated){$Form.Close()}
        if($RebootConfForm.ishandlecreated){$RebootConfForm.Close()}
        $RebootDenyCount = (Get-Content -Path "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" | Select-Object -last 1)
        $DateTime = (Get-Date)
        if($RebootDenyCount | Select-String "1 ---"){Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "2 --- Reboot request denied by $Env:USERNAME on $DateTime"}
        if($RebootDenyCount | Select-String "2 ---"){Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "3 --- Reboot request denied by $Env:USERNAME on $DateTime"}
        if($RebootDenyCount | Select-String "3 ---"){Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "4 --- Reboot request denied by $Env:USERNAME on $DateTime"}
        if($RebootDenyCount | Select-String "4 ---"){Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "5 --- Reboot request denied by $Env:USERNAME on $DateTime"}
        if($RebootDenyCount | Select-String "5 ---"){Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "6 --- Reboot request denied by $Env:USERNAME on $DateTime"}
        if($RebootDenyCount | Select-String "6 ---"){Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "7 --- Reboot request denied by $Env:USERNAME on $DateTime"}
        if($RebootDenyCount | Select-String "7 ---"){Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "8 --- Reboot request denied by $Env:USERNAME on $DateTime"}
        if($RebootDenyCount | Select-String "8 ---"){Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "9 --- Reboot request denied by $Env:USERNAME on $DateTime"}
        if($RebootDenyCount | Select-String "9 ---"){
            Add-Content "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" "10 --- Reboot request denied by $Env:USERNAME on $DateTime"
            $ShortDateTime = (Get-Date -Format MMddyyyyHHmm)
            Move-Item -Path "C:\Windows\Temp\PRC\Pending-Reboot-Checker-Denies.txt" -Destination "C:\Windows\Temp\PRC\_Archive\Pending-Reboot-Checker-Denies_$ShortDateTime.txt" | Out-Null
        }
    }
}

function RebootConf () {
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

function InstallPSWindowsUpdate () {
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -Force | Out-Null
    Install-Module PSWindowsUpdate -Force | Out-Null
    Import-Module PSWindowsUpdate -Force | Out-Null
    if(!(Get-Module | Select-Object -ExpandProperty Name | Select-String PSWindowsUpdate)){
        Write-Warning "The module PSWindowsUpdate failed to install...exiting..."
        EXIT
    }
}

function PunchIt () {
    if(!(Get-Module -Name "PSWindowsUpdate")){
        InstallPSWindowsUpdate
    }
    $PendingRebootStatus = Get-WURebootStatus -Silent -CancelReboot
    $WMIInfo = Get-WMIObject -Class Win32_OperatingSystem
    $LastBootTime = $WMIInfo.ConvertToDateTime($WMIInfo.LastBootUpTime)
    $SysUpTime = (Get-Date) - $LastBootTime
    if(($PendingRebootStatus -eq "True") -or (7 -lt ($SysUpTime.Days))){
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
        EXIT
    }
}
