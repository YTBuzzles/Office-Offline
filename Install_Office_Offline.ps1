# add assembly 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# This file should be started as administrator
if (![bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")){
    echo "=====ERROR====="
    echo ""
    echo "irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/hello.ps1 | iex"
    echo ""
    echo "Make sure to run powershell"
    echo "as administrator"
    echo ""
    pause
    exit
}


######################################################
################ GUI #################################

[System.Windows.Forms.Application]::EnableVisualStyles()
$window = New-Object System.Windows.Forms.Form
$window.ClientSize = '550, 500'

$Instructions = New-Object System.Windows.Forms.TextBox
$Instructions.Text = 
"Select which programs you want to install, more common ones are highlighted in blue.
They will all be installed automatically and activated using a script from https://massgrave.dev/"
$Instructions.Location = New-Object System.Drawing.Point(10, 5)
$Instructions.Width = 630
$Instructions.Height = 100
$Instructions.font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$Instructions.Multiline = $true

$checkbox0 = New-Object System.Windows.Forms.CheckBox
$checkbox0.Name = "Access2021Volume"
$checkbox0.Text = "Access"
$checkbox0.Location = New-Object System.Drawing.Point(20,110)

$checkbox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Name = "Excel2021Volume"
$checkbox1.Text = "Excel"
$checkbox1.BackColor = "#cee5ed"
$checkbox1.ForeColor = "#107aa1"
$checkbox1.Location = New-Object System.Drawing.Point(20,135)

$checkbox2 = New-Object System.Windows.Forms.CheckBox
$checkbox2.Name = "OneNoteVolume"
$checkbox2.Text = "OneNote"
$checkbox2.Location = New-Object System.Drawing.Point(20,160)

$checkbox3 = New-Object System.Windows.Forms.CheckBox
$checkbox3.Name = "Outlook2021Volume"
$checkbox3.Text = "Outlook"
$checkbox3.Location = New-Object System.Drawing.Point(20,185)

$checkbox4 = New-Object System.Windows.Forms.CheckBox
$checkbox4.Name = "PowerPoint2021Volume"
$checkbox4.Text = "PowerPoint"
$checkbox4.BackColor = "#cee5ed"
$checkbox4.ForeColor = "#107aa1"
$checkbox4.Location = New-Object System.Drawing.Point(20,210)

$checkbox5 = New-Object System.Windows.Forms.CheckBox
$checkbox5.Name = "Publisher2019Volume"
$checkbox5.Text = "Publisher"
$checkbox5.BackColor = "#cee5ed"
$checkbox5.ForeColor = "#107aa1"
$checkbox5.Location = New-Object System.Drawing.Point(20,235)

$checkbox6 = New-Object System.Windows.Forms.CheckBox
$checkbox6.Name = "SkypeForBusiness2021Volume"
$checkbox6.Text = "SkypeForBuisness"
$checkbox6.Width = 300
$checkbox6.Location = New-Object System.Drawing.Point(20,260)

$checkbox7 = New-Object System.Windows.Forms.CheckBox
$checkbox7.Name = "Word2021Volume"
$checkbox7.Text = "Word"
$checkbox7.BackColor = "#cee5ed"
$checkbox7.ForeColor = "#107aa1"
$checkbox7.Location = New-Object System.Drawing.Point(20,285)

$checkbox8 = New-Object System.Windows.Forms.CheckBox
$checkbox8.Name = "ProjectPro2021Volume"
$checkbox8.Text = "Project Pro"
$checkbox8.Location = New-Object System.Drawing.Point(20,310)

$checkbox9 = New-Object System.Windows.Forms.CheckBox
$checkbox9.Name = "ProjectStd2021Volume"
$checkbox9.Text = "Project Standard"
$checkbox9.Width = 300
$checkbox9.Location = New-Object System.Drawing.Point(20,335)

$checkbox10 = New-Object System.Windows.Forms.CheckBox
$checkbox10.Name = "VisioPro2021Volume"
$checkbox10.Text = "Visio Pro"
$checkbox10.Location = New-Object System.Drawing.Point(20,360)

$checkbox11 = New-Object System.Windows.Forms.CheckBox
$checkbox11.Name = "VisioStd2021Volume"
$checkbox11.Text = "Visio Standard"
$checkbox11.Width = 300
$checkbox11.Location = New-Object System.Drawing.Point(20,385)

$checkbox12 = New-Object System.Windows.Forms.CheckBox
$checkbox12.Name = "OneDrive" # ExcludedApps=OneDrive line in file
$checkbox12.Text = "OneDrive Desktop"
$checkbox12.Width = 300
$checkbox12.Location = New-Object System.Drawing.Point(20,410)

$checkbox13 = New-Object System.Windows.Forms.CheckBox
$checkbox13.Text = "Microsoft Teams" # not sure for this one
$checkbox13.Width = 300
$checkbox13.Location = New-Object System.Drawing.Point(20,435)

$checkbox14 = New-Object System.Windows.Forms.CheckBox
$checkbox14.Text = "Activate"
$checkbox14.Width = 300
$checkbox14.Location = New-Object System.Drawing.Point(300,200)
$checkbox14.Checked = $true

$installButton                       = New-Object system.Windows.Forms.Button
$installButton.BackColor             = "#ffffff"
$installButton.text                  = "Install"
$installButton.width                 = 110
$installButton.height                = 50
$installButton.location              = New-Object System.Drawing.Point(400,340)
$installButton.Font                  = 'Microsoft Sans Serif,10'
$installButton.ForeColor             = "#000"

$DownloadFiles                       = New-Object system.Windows.Forms.Button
$DownloadFiles.BackColor             = "#ffffff"
$DownloadFiles.text                  = "Download"
$DownloadFiles.width                 = 110
$DownloadFiles.height                = 50
$DownloadFiles.location              = New-Object System.Drawing.Point(400,250)
$DownloadFiles.Font                  = 'Microsoft Sans Serif,10'
$DownloadFiles.ForeColor             = "#000"

$exit                       = New-Object system.Windows.Forms.Button
$exit.BackColor             = "#ffffff"
$exit.text                  = "Exit"
$exit.width                 = 70
$exit.height                = 50
$exit.location              = New-Object System.Drawing.Point(440,440)
$exit.Font                  = 'Microsoft Sans Serif,10'
$exit.ForeColor             = "#000"

function exiting {
$window.Dispose()
echo "All finished"
}
$exit.Add_Click({exiting})


function updateDownload {
if ($downloaded){
$DownloadFiles.ForeColor = "#1d8223"
$DownloadFiles.text = "Done"
}
}
updateDownload

function download {
if (!$downloaded){
# Run Generator file which creates script to download
# offline install files with choices already selected
.\YAOCTRU_Generator.cmd
echo ""
echo "created download script"
echo ""
echo "downloading installation files"
Start-Sleep -Seconds 3

# Run generated file named "XX.X.XXXXX.XXXXX_x64_en-US_Monthly_curl.bat"
echo(gci -Filter "*en-US_Monthly*").Name
echo ""
Start-Process -FilePath ((gci -Filter "*en-US_Monthly*").Name) -Wait

echo "Installation files downloaded"
echo ""
echo "Choosing Configuration"
echo ""
updateDownload
}
}
$DownloadFiles.Add_Click({download})

$changeDirectory                       = New-Object system.Windows.Forms.Button
$changeDirectory.BackColor             = "#ffffff"
$changeDirectory.text                  = "Directory"
$changeDirectory.width                 = 300
$changeDirectory.height                = 50
$changeDirectory.location              = New-Object System.Drawing.Point(200,120)
$changeDirectory.Font                  = 'Microsoft Sans Serif,10'
$changeDirectory.ForeColor             = "#000"

$window.Controls.AddRange(@($checkbox0,$checkbox1,$checkbox2,$checkbox3,$checkbox4,$checkbox5,$checkbox6,$checkbox7,$checkbox8,$checkbox9,$checkbox10,
    $checkbox11,$checkbox12,$checkbox13,$Instructions,$installButton,$checkbox14,$DownloadFiles,$changeDirectory,$exit))

$window.CancelButton = $installButton

function install{    
    # Use name to add to SKU list
    # now for logic
    $skus = "SKUs="
    for ($i=0; $i -lt 12; $i++)
    {
    
        $varName = "checkbox" + $i
        $variable = (Get-Variable -Include $varName)
        if ($variable.Value.CheckState){
            $skus += $variable.Value.Name.ToString() + ","
        }
    }
    if ($skus[$skus.Length-1] -eq ','){
    $skus = $skus.Remove($skus.Length-1,1)
    }

    # edit file from here
    # first delete old config file
    if ((gci -Filter "C2R_Config").Name.Length -gt 0){
        Remove-Item -Path (gci -Filter "C2R_Config").Name
    }


    $file = (gci -Filter "permconfig.ini").Name
    $filecontent = Get-Content -Path $file -Raw

    # if onedrive and teams remove ExcludedApps= also for just onedrive
    # if just teams, ExcludedApps=OneDrive
    # if both excluded, ExcludedApps=OneDrive
    if ($checkbox12.CheckState){
        $filecontent = $filecontent.Replace("ExcludedApps=", "")
    }
    if ($checkbox13.CheckState -and !$checkbox12.CheckState){
        $filecontent = $filecontent.Replace("ExcludedApps=", "ExcludedApps=OneDrive")
    }
    if (!$checkbox12.CheckState -and !$checkbox13.CheckState){
        $filecontent = $filecontent.Replace("ExcludedApps=", "ExcludedApps=OneDrive")
    }

    $filecontent.Replace("SKUs=",$skus) | Set-Content -Path "C2R_Config_20230627-1649.ini"
    
    # Now run the installer
    #Start-Process $PSScriptRoot/YAOCTRI_Installer.cmd -Verb RunAs -Wait
    
    # Now activate the office products
    if ($checkbox14.Checked){
    iex "&{$(irm https://massgrave.dev/get)} /KMS-Office"
    }
}
$InstallButton.Add_Click({install})


function changeDirectory(){
# now ask for directory to run installer in, not where office products will be installed
$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FileBrowser.SelectedPath = "C:\"
$null = $FileBrowser.ShowDialog()

# change to chosen directory
cd $FileBrowser.SelectedPath
$changeDirectory.Text = $FileBrowser.SelectedPath

# check if the installation files are already downloaded
Start-Sleep -Seconds 5
$downloaded = $true
if ((gci -Filter "C2R_Monthly").Name -ne "C2R_Monthly"){
$downloaded = $false
}
# now download all other helper functions from github (already downloaded ones will be overwritten)
irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/permconfig.ini -OutFile "permconfig.ini"
irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/YAOCTRI_Installer.cmd -OutFile "YAOCTRI_Installer.cmd"
irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/YAOCTRU_Generator.cmd -OutFile "YAOCTRU_Generator.cmd"
updateDownload
}
$changeDirectory.Add_Click({changeDirectory})


$SKU = $window.ShowDialog()

####################################################################
################### END OF GUI #####################################
