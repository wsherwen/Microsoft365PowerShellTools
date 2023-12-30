MS365ModuleInstall.ps1: A PowerShell script to easily install or uninstall essential Microsoft 365 PowerShell modules.

Key Features:

Installs or uninstalls multiple modules efficiently.
Ensures current NuGet version for compatibility.
Generates detailed log files for tracking actions.
Moves installed modules to a common location for easy management.
Supports the following modules:
MicrosoftTeams
MSOnline
AzureAD
ExchangeOnlineManagement
Microsoft.Online.Sharepoint.PowerShell
SharePointPnPPowerShellOnline
ORCA
WhiteboardAdmin
Usage Instructions:

Open PowerShell as an Administrator.

Run the script with the appropriate mode:

To install modules:
PowerShell
.\MS365ModuleInstall.ps1 -Mode Install
Use code with caution. Learn more
To uninstall modules:
PowerShell
.\MS365ModuleInstall.ps1 -Mode Uninstall
Use code with caution. Learn more
Review the log file (optional):

The log file is created at $env:windir\Temp\Logs\MS365ModuleInstall.log.
Additional Notes:

The script requires an active internet connection to download modules.
It may prompt for confirmation during installation or uninstallation.
It's recommended to run as an administrator to ensure proper permissions.
