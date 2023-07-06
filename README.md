# Office-Offline

This is an Offline Office install with a gui
forked from abbodi1406's Office Click-To-Run Installer (https://github.com/abbodi1406/BatUtil)

It uses massgravel's Microsoft Activation Scripts (MAS) - Microsoft Activation Scripts (MAS)
(https://github.com/massgravel/Microsoft-Activation-Scripts)
* the activation command is ``` iex "&{$(irm https://massgrave.dev/get)} /KMS-Office" ```

To run the installer open powershell as administrator and run the command below
```
irm https://raw.githubusercontent.com/YTBuzzles/Office-Offline/main/Install_Office_Offline.ps1 | iex
```
It will ask for a directory which is where the install scripts will run,
any office product will be installed into "C:\Program Files"


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
