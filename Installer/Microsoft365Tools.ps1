# Author: Warren Sherwen
# Last Edit: Warren Sherwen
# Verison: 2.0
# Change: Amended to cover Microsoft Graph

Param(
[Parameter(Mandatory=$true)]
[ValidateSet("Install", "Uninstall")]
[String[]]
$Mode
)

$Logfile = "$env:APPDATA\Temp\Logs\MS365ModuleInstall.log"
Function LogWrite{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring
   write-output $logstring
   
}
function Get-TimeStamp {
    return "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

if (!(Test-Path "$env:APPDATA\Temp\Logs\"))
{
   mkdir $env:APPDATA\Temp\Logs
   LogWrite "$(Get-TimeStamp): Script has started."
   LogWrite "$(Get-TimeStamp): Log directory created."
}
else
{
    LogWrite "$(Get-TimeStamp): Script has started."
    LogWrite "$(Get-TimeStamp): Log directory exists."
}

$O365Modules = @("MicrosoftTeams", "MSOnline", "AzureAD", "ExchangeOnlineManagement", "ORCA", "WhiteboardAdmin", "Microsoft.Graph*")

LogWrite "$(Get-TimeStamp): Checking for current NuGet Version Level."
$PackageProviderVersion = (Get-PackageProvider -ForceBootstrap -Name NuGet).Version
LogWrite "$(Get-TimeStamp): Current NuGet Verison: $PackageProviderVersion." 

if ($PackageProviderVersion -lt '2.8.5.201') {
    LogWrite "$(Get-TimeStamp): Updating NuGet."
    Install-PackageProvider -Name NuGet -Force
    LogWrite "$(Get-TimeStamp): NuGet Updated."
    $PackageProviderVersion = (Get-PackageProvider -Name NuGet).Version
    LogWrite "$(Get-TimeStamp): New NuGet Verison: $PackageProviderVersion."
} else {
    LogWrite "$(Get-TimeStamp): Current NuGet Verison: $PackageProviderVersion no update required"
}

If ($Mode -eq "Install") {
    LogWrite "$(Get-TimeStamp): Checking currently installed modules for MS365."
    ForEach ($Module in $O365Modules) {
         LogWrite "$(Get-TimeStamp): Checking for the module: $Module."
            if ($module = "Microsoft.Graph*") {
                    $MSGraph = Find-Module -name 'Microsoft.Graph.*'
                     ForEach ($Graph in $MSGraph) {
                            $GraphName = $Graph.Name
                            if (Get-Module -ListAvailable -Name $Graph.name) {
                            LogWrite "$(Get-TimeStamp): The module $GraphName is already installed."
                        } else {
                            LogWrite "$(Get-TimeStamp): The module $GraphName was not found, now install $GraphName."
                            Install-Module -Name $Graph.Name -AllowClobber -Force -Scope CurrentUser
                            LogWrite "$(Get-TimeStamp): The module $GraphName was installed."
                        }                
                    }
        } else {
                if (Get-Module -ListAvailable -Name $Module) {
                LogWrite "$(Get-TimeStamp): The module $Module is already installed."
            } else {
                LogWrite "$(Get-TimeStamp): The module $Module was not found, now installing $Module."
                Install-Module -Name $Module -AllowClobber -Scope CurrentUser -Force 
                LogWrite "$(Get-TimeStamp): The module $Module was installed."
            }
        }
    }
}

If ($Mode -eq "Uninstall") {
    ForEach ($Module in $O365Modules) {
        if (Get-Module -ListAvailable -Name $Module) {
            if ($module = "Microsoft.Graph*") {
                $MSGraph = Get-Module -ListAvailable -name Microsoft.Graph*
                ForEach ($Graph in $MSGraph) {
                    $GraphName = $Graph.Name
                    LogWrite "$(Get-TimeStamp): The module $GraphName was found, now removing $GraphName."
                    Uninstall-Module -Name $Graph.Name -RequiredVersion $Graph.Version -Force
                    LogWrite "$(Get-TimeStamp): The module $GraphName was removed."
                }
            }else {
            LogWrite "$(Get-TimeStamp): The module $Module was found, now removing $Module."
            Uninstall-Module -Name $Module -AllVersions -Scope CurrentUser -Force
            LogWrite "$(Get-TimeStamp): The module $Module was removed."
            }
        } 
            else {
                LogWrite "$(Get-TimeStamp): The module $Module was not found, nothing to remove."
            }
            LogWrite "$(Get-TimeStamp): The script has been executed, now exiting..."
        }
}

LogWrite "$(Get-TimeStamp): The script has been executed, now exiting..."
Exit
