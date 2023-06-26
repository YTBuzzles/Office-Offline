# This file should be started as administrator

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

# Now run the Configurator file
# basically just want to create a line that looks like this
# SKUs=Access2021Volume,Excel2021Volume,Outlook2021Volume,PowerPoint2021Volume,Publisher2021Volume,SkypeForBusiness2021Volume,Word2021Volume,OneNoteVolume
#
# For the moment I will just use a generic .ini configuration, so no need to run configurator


# Now run the installer, later when I do gui can just use option to install at end of configurator
Start-Process -FilePath YAOCTRI_Installer.cmd -Wait

# Now activate the office products
& ([ScriptBlock]::Create((irm https://massgrave.dev/get))) /KMS-Office

echo "All finished"
pause
exit

