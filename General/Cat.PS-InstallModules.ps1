<# Cat.PS-InstallModules.ps1 

This script contains the Module related commands for the PowerShell modules used by the scripts in the 
CatPnP\CatPowerShell Repo.  
#>

# Install all modules that are currently used by script repository
Install-Module Microsoft.Online.SharePoint.PowerShell #-Scope AllUsers
Install-Module MicrosoftTeams -AllowClobber #-Scope AllUsers -RequiredVersion 6.8.0 # There is an issue with 7.0.0. Try the latest if it doesn't go back to 6.8.0
Install-Module PnP.PowerShell #-Scope AllUsers
# PowerShell 5 for ShareGate: Install-Module -Name PnP.PowerShell -RequiredVersion 1.12.0 -Scope AllUsers
        # You may need to update TLS: [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module Az #-Scope AllUsers
#Install-Module Microsoft.Graph -Scope AllUsers
#Install-Module AzureAD #-Scope AllUsers
#Install-Module MSOnline #-Scope AllUsers
#Install-Module ExchangeOnlineManagement #-Scope AllUsers
#Install-Module -Name Microsoft.PowerApps.Administration.PowerShell #-Scope AllUsers
#Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber #-Scope AllUsers 
#Install-Module SQLServer #-Scope AllUsers
# Desired State Configurator: Install-Module Microsoft365DSC -Scope AllUsers
    # Desired State Configurator: Update-M365DSCDependencies -Scope AllUsers
    # Desired State Configurator: Install-Module PSDscResources -Scope AllUsers

# Update all modules that are currently used by script repository
Update-Module Microsoft.Online.SharePoint.PowerShell
Update-Module MicrosoftTeams
Update-Module PnP.PowerShell
Update-Module Az
#Update-Module Microsoft.Graph
#Update-Module AzureAD
#Update-Module MSOnline
#Update-Module ExchangeOnlineManagement
#Update-Module -Name Microsoft.PowerApps.Administration.PowerShell
#Update-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber
#Update-Module SQLServer
# Desired State Configurator: Update-Module Microsoft365DSC # Desired State Configurator
    # Desired State Configurator: Update-M365DSCDependencies
    # Desired State Configurator: Update-Module PSDscResources

# Import all modules that are currently used by script repository
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
Import-Module MicrosoftTeams
Import-Module PnP.PowerShell
Import-Module Az
#Import-Module Microsoft.Graph
#Import-Module AzureAD
#Import-Module MSOnline
#Import-Module ExchangeOnlineManagement
#Import-Module -Name Microsoft.PowerApps.Administration.PowerShell
#Import-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber
#Import-Module SQLServer
# Desired State Configurator: Import-Module Microsoft365DSC

# Remove all modules that are currently used by script repository
# Must remove before Uninstall
Remove-Module Microsoft.Online.SharePoint.PowerShell
Remove-Module MicrosoftTeams
Remove-Module PnP.PowerShell
Remove-Module Az
#Remove-Module Microsoft.Graph
#Remove-Module AzureAD
#Remove-Module MSOnline
#Remove-Module ExchangeOnlineManagement
#Remove-Module -Name Microsoft.PowerApps.Administration.PowerShell
#Remove-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber
#Remove-Module SQLServer
# Desired State Configurator: Remove-Module Microsoft365DSC

# Uninstall all modules that are currently used by script repository
# Must remove before Uninstall
Uninstall-Module Microsoft.Online.SharePoint.PowerShell -AllVersions -Force
Uninstall-Module MicrosoftTeams -AllVersions -Force
Uninstall-Module PnP.PowerShell -AllVersions -Force
Uninstall-Module Az -AllVersions -Force
#Uninstall-Module Microsoft.Graph -AllVersions -Force
#Uninstall-Module AzureAD -AllVersions -Force
#Uninstall-Module MSOnline -AllVersions -Force
#Uninstall-Module ExchangeOnlineManagement -AllVersions -Force
#Uninstall-Module -Name Microsoft.PowerApps.Administration.PowerShell -AllVersions -Force
#Uninstall-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -AllVersions -Force
#Uninstall-Module SQLServer -AllVersions -Force
# Desired State Configurator: Uninstall-Module Microsoft365DSC -AllVersions -Force