# add assembly 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# This file should be started as administrator
if (![bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")){
    echo "=====ERROR====="
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
"1. Change installer files directory if needed
2. If download button doesn't say done click it
3. Choose desired applications
4. Choose activation method
5. Click install"

$Instructions.Location = New-Object System.Drawing.Point(10, 5)
$Instructions.Width = 530
$Instructions.Height = 110
$Instructions.font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$Instructions.Multiline = $true

$checkbox0 = New-Object System.Windows.Forms.CheckBox
$checkbox0.Name = "Access2021Volume"
$checkbox0.Text = "Access"
$checkbox0.Location = New-Object System.Drawing.Point(20,120)

$checkbox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Name = "Excel2021Volume"
$checkbox1.Text = "Excel"
$checkbox1.BackColor = "#cee5ed"
$checkbox1.ForeColor = "#107aa1"
$checkbox1.Location = New-Object System.Drawing.Point(20,145)

$checkbox2 = New-Object System.Windows.Forms.CheckBox
$checkbox2.Name = "OneNoteVolume"
$checkbox2.Text = "OneNote"
$checkbox2.Location = New-Object System.Drawing.Point(20,170)

$checkbox3 = New-Object System.Windows.Forms.CheckBox
$checkbox3.Name = "Outlook2021Volume"
$checkbox3.Text = "Outlook"
$checkbox3.Location = New-Object System.Drawing.Point(20,195)

$checkbox4 = New-Object System.Windows.Forms.CheckBox
$checkbox4.Name = "PowerPoint2021Volume"
$checkbox4.Text = "PowerPoint"
$checkbox4.BackColor = "#cee5ed"
$checkbox4.ForeColor = "#107aa1"
$checkbox4.Location = New-Object System.Drawing.Point(20,220)

$checkbox5 = New-Object System.Windows.Forms.CheckBox
$checkbox5.Name = "Publisher2019Volume"
$checkbox5.Text = "Publisher"
$checkbox5.BackColor = "#cee5ed"
$checkbox5.ForeColor = "#107aa1"
$checkbox5.Location = New-Object System.Drawing.Point(20,245)

$checkbox6 = New-Object System.Windows.Forms.CheckBox
$checkbox6.Name = "SkypeForBusiness2021Volume"
$checkbox6.Text = "SkypeForBuisness"
$checkbox6.Width = 300
$checkbox6.Location = New-Object System.Drawing.Point(20,270)

$checkbox7 = New-Object System.Windows.Forms.CheckBox
$checkbox7.Name = "Word2021Volume"
$checkbox7.Text = "Word"
$checkbox7.BackColor = "#cee5ed"
$checkbox7.ForeColor = "#107aa1"
$checkbox7.Location = New-Object System.Drawing.Point(20,295)

$checkbox8 = New-Object System.Windows.Forms.CheckBox
$checkbox8.Name = "ProjectPro2021Volume"
$checkbox8.Text = "Project Pro"
$checkbox8.Location = New-Object System.Drawing.Point(20,320)

$checkbox9 = New-Object System.Windows.Forms.CheckBox
$checkbox9.Name = "ProjectStd2021Volume"
$checkbox9.Text = "Project Standard"
$checkbox9.Width = 300
$checkbox9.Location = New-Object System.Drawing.Point(20,345)

$checkbox10 = New-Object System.Windows.Forms.CheckBox
$checkbox10.Name = "VisioPro2021Volume"
$checkbox10.Text = "Visio Pro"
$checkbox10.Location = New-Object System.Drawing.Point(20,370)

$checkbox11 = New-Object System.Windows.Forms.CheckBox
$checkbox11.Name = "VisioStd2021Volume"
$checkbox11.Text = "Visio Standard"
$checkbox11.Width = 300
$checkbox11.Location = New-Object System.Drawing.Point(20,395)

$checkbox12 = New-Object System.Windows.Forms.CheckBox
$checkbox12.Name = "OneDrive" # ExcludedApps=OneDrive line in file
$checkbox12.Text = "OneDrive Desktop"
$checkbox12.Width = 300
$checkbox12.Location = New-Object System.Drawing.Point(20,420)

$checkbox13 = New-Object System.Windows.Forms.CheckBox
$checkbox13.Text = "Microsoft Teams" # not sure for this one
$checkbox13.Width = 300
$checkbox13.Location = New-Object System.Drawing.Point(20,445)

$checkbox14 = New-Object System.Windows.Forms.CheckBox
$checkbox14.Text = "Permanent Activation"
$checkbox14.Width = 300
$checkbox14.Location = New-Object System.Drawing.Point(300,190)
$checkbox14.Checked = $true

$checkbox15 = New-Object System.Windows.Forms.CheckBox
$checkbox15.Text = "Activate With Renewal"
$checkbox15.Width = 300
$checkbox15.Location = New-Object System.Drawing.Point(300,220)
$checkbox15.Checked = $false

$checkbox16 = New-Object System.Windows.Forms.CheckBox
$checkbox16.Text = "Activate for 180 days"
$checkbox16.Width = 300
$checkbox16.Location = New-Object System.Drawing.Point(300,250)
$checkbox16.Checked = $false

$installButton                       = New-Object system.Windows.Forms.Button
$installButton.BackColor             = "#ffffff"
$installButton.text                  = "Install"
$installButton.width                 = 110
$installButton.height                = 50
$installButton.location              = New-Object System.Drawing.Point(400,350)
$installButton.Font                  = 'Microsoft Sans Serif,10'
$installButton.ForeColor             = "#000"

$DownloadFiles                       = New-Object system.Windows.Forms.Button
$DownloadFiles.BackColor             = "#ffffff"
$DownloadFiles.text                  = "Download"
$DownloadFiles.width                 = 130
$DownloadFiles.height                = 50 
$DownloadFiles.location              = New-Object System.Drawing.Point(390,280)
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

$changeDirectory                       = New-Object system.Windows.Forms.Button
$changeDirectory.BackColor             = "#ffffff"
$changeDirectory.text                  = "Installer Directory"
$changeDirectory.width                 = 300
$changeDirectory.height                = 50
$changeDirectory.location              = New-Object System.Drawing.Point(200,130)
$changeDirectory.Font                  = 'Microsoft Sans Serif,10'
$changeDirectory.ForeColor             = "#000"

$window.Controls.AddRange(@($checkbox0,$checkbox1,$checkbox2,$checkbox3,$checkbox4,$checkbox5,$checkbox6,$checkbox7,$checkbox8,$checkbox9,$checkbox10,
    $checkbox11,$checkbox12,$checkbox13,$Instructions,$installButton,$checkbox14,$DownloadFiles,$changeDirectory,$exit,$checkbox15,$checkbox16))

function updateDownload {

# check if the installation files are already downloaded
if ((gci -Filter "C2R_Monthly").Name -eq "C2R_Monthly" -and ("$(pwd)" + "\C2R_Monthly" | Get-ChildItem | Measure-Object).count -gt 0){
$DownloadFiles.ForeColor = "#1d8223"
$DownloadFiles.text = "Done"
}
else {
$DownloadFiles.ForeColor = "#000000"
$DownloadFiles.text = "Download"
}
}

function chooseDirectory(){
# now ask for directory to run installer in, not where office products will be installed
$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FileBrowser.Description = "Choose where to put installer files.
This is not where programs will be installed"
$FileBrowser.SelectedPath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$null = $FileBrowser.ShowDialog()

# change to chosen directory
cd $FileBrowser.SelectedPath
$changeDirectory.Text = $FileBrowser.SelectedPath

updateDownload
}
$changeDirectory.Add_Click({chooseDirectory})


function download {

if (!((gci -Filter "C2R_Monthly").Name -eq "C2R_Monthly" -and ("$(pwd)" + "\C2R_Monthly" | Get-ChildItem | Measure-Object).count -gt 0)){
    # first download the generator, configurator and installer files
    irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/YAOCTRU_Generator.cmd -OutFile "YAOCTRU_Generator.cmd"

    $DownloadFiles.ForeColor = "#1d8223"
    $DownloadFiles.text = "DOWNLOADING"

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
    updateDownload }

}
$DownloadFiles.Add_Click({download})

# Make activation checkboxes flip when one is clicked

$checkbox14.Add_Click({flipchkbox(1)})
$checkbox15.Add_Click({flipchkbox(2)})
$checkbox16.Add_Click({flipchkbox(3)})

function flipchkbox($num) {
    $ParameterName
    if ($num -eq 1){
        $checkbox15.Checked = $false
        $checkbox16.Checked = $false
    }
    if ($num -eq 2){
        $checkbox14.Checked = $false
        $checkbox16.Checked = $false
    }
    if ($num -eq 3){
        $checkbox15.Checked = $false
        $checkbox14.Checked = $false
    }
}



################################# Improve this I guess ##########################################
function install{
    
if ((gci -Filter "C2R_Monthly").Name -eq "C2R_Monthly" -and ("$(pwd)" + "\C2R_Monthly" | Get-ChildItem | Measure-Object).count -gt 0 -and $DownloadFiles.text -eq "Done") {

    irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/YAOCTRI_Configurator.cmd -OutFile "YAOCTRI_Configurator.cmd"
    irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/YAOCTRI_Installer.cmd -OutFile "YAOCTRI_Installer.cmd"

    $installButton.text = "Installing"
    $installButton.ForeColor = "#1d8223"
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

    # first delete old config file
    if ((gci -Filter "*.ini").Name.Length -gt 0){
        Remove-Item -Path (gci -Filter "*.ini").Name
    }

    # now run .ini generator
    $configurator = "$(pwd)" + "/YAOCTRI_Configurator.cmd"
    Start-Process $configurator -Verb RunAs -Wait

    # now edit contents
    $file = (gci -Filter "*.ini").Name
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
    $filecontent.Replace("SKUs=",$skus) | Set-Content -Path $file
    
    # Now run the installer
    $installer = "$(pwd)" + "/YAOCTRI_Installer.cmd"
    Start-Process $installer -Verb RunAs -Wait

    $installButton.text = "Install"
    $installButton.ForeColor = "#000000"
    
    # Now activate the office products
    if ($checkbox14.Checked){
    Invoke-Expression "&{$(irm https://massgrave.dev/get)} /Ohook"}
    elseif ($checkbox15.Checked){
    Invoke-Expression "&{$(irm https://massgrave.dev/get)} /KMS-Office /KMS-RenewalTask"
    } elseif ($checkbox16.Checked){
    Invoke-Expression "&{$(irm https://massgrave.dev/get)} /KMS-Office"
    }
}
}
$InstallButton.Add_Click({install})


function exiting {
$window.Dispose()
echo "All finished"
}
$exit.Add_Click({exiting})

chooseDirectory
$SKU = $window.ShowDialog()

####################################################################
################### END OF GUI #####################################
