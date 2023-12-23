# Author: Warren Sherwen
# Last Edit: Warren Sherwen
# Verison: 1.0

$Logfile = "$env:windir\Temp\Logs\MS365CheckInstalledModules.log"
Function LogWrite{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring
   write-output $logstring
   
}
function Get-TimeStamp {
    return "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

if (!(Test-Path "$env:windir\Temp\Logs\"))
{
   mkdir $env:windir\Temp\Logs
   LogWrite "$(Get-TimeStamp): Script has started."
   LogWrite "$(Get-TimeStamp): Log directory created."
}
else
{
    LogWrite "$(Get-TimeStamp): Script has started."
    LogWrite "$(Get-TimeStamp): Log directory exists."
}

$modulesToCheck = @(
    "MicrosoftTeams", "MSOnline", "AzureAD", "AzureADPreview", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", "SharePointPnPPowerShellOnline", "ORCA", "WhiteboardAdmin"
)

$O365ModuleInstalled = $true
LogWrite "$(Get-TimeStamp): Checking for installed modules."


foreach ($module in $modulesToCheck) {
    if (-not (Get-Module -Name $module -ListAvailable -ErrorAction SilentlyContinue)) {
        LogWrite "$(Get-TimeStamp): The $module was not installed."
        LogWrite "$(Get-TimeStamp): Checking for the next module."
        $O365ModuleInstalled = $false
    } else {
        LogWrite "$(Get-TimeStamp): The $module is installed."
        LogWrite "$(Get-TimeStamp): Checking for the next module."
    }
}

if ($O365ModuleInstalled -eq $false) {
    LogWrite "$(Get-TimeStamp): Script is ending, a moduke was not found please check the log file."
    Write-Host "A Module was not found, please refer to the log file."
    exit 1 
} else {
    LogWrite "$(Get-TimeStamp): Script is ending, all modules found."
    Write-Host "All O365 Modiles Found."
    exit 0
}
