$Logfile = "$env:windir\Temp\Logs\MS365ModuleInstall.log"
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

$O365Modules = @("MicrosoftTeams", "MSOnline", "AzureADPreview", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", "SharePointPnPPowerShellOnline", "ORCA", "WhiteboardAdmin")

LogWrite "$(Get-TimeStamp): Setting the installtion policy for PSGallery to trusted."
Set-PSRepository -name "PSGallery" -InstallationPolicy Trusted
LogWrite "$(Get-TimeStamp): The installtion policy for PSGallery is now set to trusted."

LogWrite "$(Get-TimeStamp): Checking currently installed modules for MS365."
ForEach ($Module in $O365Modules) {
   if (Get-Module -ListAvailable -Name $Module) {
    LogWrite "$(Get-TimeStamp): The module $Module is already installed."
} 
else {
    Install-Module -Name $Module -AllowClobber
    LogWrite "$(Get-TimeStamp): The module $Module was not found, now installing $Module."
}
}

LogWrite "$(Get-TimeStamp): The script has been executed, now exiting..."
exit

