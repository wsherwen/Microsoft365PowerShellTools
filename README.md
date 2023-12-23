# Microsft 365 PowerShell Tools
The script will check for currently installed Microsoft PowerShell Tools, if it cannot locate any it'll then proceed to install.

The script isn't trusted, so you have to set the excuation policy to: set-executionpolicy unrestricted or run as "powershell.exe -ExecutionPolicy Bypass -file Microsoft365Tools.ps1 -Mode Install"
Once done, via the administrator powershell window, run the script with Local Administrative rights.

# Intune Instllation
To package this application, you will need to do the following:

Save the powershell script: Microsoft365Tools.ps1 to a local folder:
C:\Temp\Microsoft365Tools

Run the intunewin package tool
Set the source directory to: C:\Temp\Microsoft365Tools
Set the instllation file to: C:\Temp\Microsoft365Tools\Microsoft365Tools.ps1
Set the output directory to: C:\Temp\
Enter No for the catalog question.

Once packaged, upload the outputted Microsoft365Tools.intunewin file to Intune.

Set the App Information to:

App information
Edit
Name
Microsoft 365 PowerShell Tools
Description
Microsoft365Tools.ps1

Publisher Warren Sherwen
App Version 1.0
Category Development & Design or Computer management
Information URL https://github.com/wsherwen/365-PowerShell-Tools/edit/main/README.md
Developer Warren Sherwen

Install command
powershell.exe -ExecutionPolicy Bypass -file Microsoft365Tools.ps1 -Mode Install
Uninstall command
powershell.exe -ExecutionPolicy Bypass -file Microsoft365Tools.ps1 -Mode Uninstall
Installation time required (mins) 5
Allow available uninstall Yes
Install behavior System
Device restart behavior App install may force a device restart

Operating system architecture x86,x64
Minimum operating system Windows 10 1607

For the detection rules, select custom script and upload the CheckMS365Tools.ps1

For the scope, set it the user base you wish to deploy the modules too.

Install Log File: $env:windir\Temp\Logs\MS365ModuleInstall.log
Check Log File: $env:windir\Temp\Logs\MS365CheckInstalledModules.log
