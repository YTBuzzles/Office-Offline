## Office-Offline
This is an Offline Office install with a gui
forked from abbodi1406's Office Click-To-Run Installer (https://github.com/abbodi1406/BatUtil)

It uses massgravel's Microsoft Activation Scripts (MAS) - Microsoft Activation Scripts (MAS)
(https://github.com/massgravel/Microsoft-Activation-Scripts)

# Run Script
To run the installer open powershell as administrator and run the command below
```
irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/Install_Office_Offline.ps1 | iex
```

- Checkboxes
  - tick which programs you wish to install (More common ones are highlighted in blue)<br>
- Activate
  - Activate with Ohook method is a permanent activation<br>
            to install manually use
            ```iex "&{$(irm https://massgrave.dev/get)} /Ohook"```<br> or to uninstall manually
            use: ```iex "&{$(irm https://massgrave.dev/get)} /Ohook-Uninstall"```
  - Activate with renewal activates and installs a script that checks every 7 days if<br>
            if needs to be activated again. Install Manually:<br>
            ```iex "&{$(irm https://massgrave.dev/get)} /KMS-Office /KMS-RenewalTask"```
  - Activate for 180 days activates a single time for 180 days, if you wish to use this<br>
            when the activation stops you can run the command below to activate again<br>
            ```iex "&{$(irm https://massgrave.dev/get)} /KMS-Office"```<br>
- Directory
  - the directory button is for choosing the location of install files, not where the<br>
                actual applications are installed. It defaults to Downloads folder. The download is<br>
                over 3GB so if you want to run the script again without downloading, consider choosing<br>
                a more permament location<br>
- Download
  - After choosing the Directory if hasn't changed to "Done" in green, press it to start<br>
                download of the install files<br>
- Install
  - Before pressing make sure you have chosen the applications you wish to install. It will<br>
                install them all at once<br>
- Exit
  - Once you are done press Exit or window X button to close out of app<br>

Available Products
Volume (2021)          |                       |
---------------------- |---------------------- |
Access                 |Word                   |
Excel                  |Project Pro            |
OneNote                |Project Standard       |
Outlook                |Visio Pro              |
PowerPoint             |Visio Standard         |
Publisher              |Onerive Desktop        |
Skype For Buisness     |Microsoft Teams        |
