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

# now ask for directory to run installer in, not where office products will be installed
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
$null = $FileBrowser.ShowDialog()

# change to chosen directory
cd $FileBrowser.FileName

# now download all other helper functions from github
irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/config.ps1 -OutFile "config.ps1"
irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/permconfig.ini -OutFile "permconfig.ini"
irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/YAOCTRI_Installer.cmd -OutFile "YAOCTRI_Installer.cmd"
irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/YAOCTRU_Generator.cmd -OutFile "YAOCTRU_Generator.cmd"

Start-Sleep -Seconds 5
# check if the installation files are already downloaded
if ((gci -Filter "C2R_Monthly").Name -ne "C2R_Monthly")
{

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

}

# run config gui
& "$PSScriptRoot/config.ps1" -Wait


# Now run the installer, later when I do gui can just use option to install at end of configurator
Start-Process $PSScriptRoot/YAOCTRI_Installer.cmd -Verb RunAs -Wait

# Now activate the office products
iex "&{$(irm https://massgrave.dev/get)} /KMS-Office"

echo "All finished"
pause
exit

