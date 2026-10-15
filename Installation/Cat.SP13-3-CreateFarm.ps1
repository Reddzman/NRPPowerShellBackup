#***********************************************************************************************************************************************

# This script assumes that the "Farm-Settings.xml" file has the correct settings.  If you have multiple farms, you can configure them
# differenlty by creating multiple versions of this file and then editing the "$InputFile" variable each time you run the script.

#***********************************************************************************************************************************************

# Start Time
$Start = (get-date -format "MM-dd-yyyy HH:mm")

# Checking Admin privilages
function Use-RunAs 
{    
    # Check if script is running as Adminstrator and if not use RunAs 
    # Use Check Switch to check if admin 
     
    param([Switch]$Check) 
     
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") 
         
    if ($Check) { return $IsAdmin }     
 
    if ($MyInvocation.ScriptName -ne "") 
    {  
        if (-not $IsAdmin)  
        {  
            try 
            {  
                $arg = "-file `"$($MyInvocation.ScriptName)`"" 
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'  
            } 
            catch 
            { 
                Write-Warning "Error - Failed to restart script with runas"  
                break               
            } 
            exit # Quit this session of powershell 
        }  
    }  
    else  
    {  
        Write-Warning "Error - Script must be saved as a .ps1 file first"  
        break  
    }  
} 
Use-RunAs

# Script variables
$date = get-date -format "yyyy-MM-dd_HH-mm"
$Task = "Farm Configuration"
$Reset = $False #  This variable can be used to determine if an IISReset is necessary.  If you set it to True durning the script it will reset at the end.
$ScriptErrors = @()
$Logfile = "d:\Installation\Scripts\_Log Files\FarmCreation_$date.log.txt"
$InputFile = "d:\Installation\Scripts\_Input Files\Clients\Cat.FarmConfig-Adams12-Prod.xml"

# Creating Log file
write-host "____________________________________________________"
out-file -filepath $Logfile -inputobject "____________________________________________________"
write-host "Starting $Task - $Start"
out-file -filepath $Logfile -append -inputobject "Starting $Task - $Start"

#***********************************************************************************************************************************************

#                                         Variables

#***********************************************************************************************************************************************

# Importing XML
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If ((test-path $InputFile) -eq $True)
{
    write-host ("   Importing XML file " + $InputFile)
    out-file -filepath $Logfile -inputobject ("   Importing XML file " + $InputFile)
    [xml]$ConfigXML = get-content $InputFile -ErrorVariable +ScriptErrors
}
Else
{
    write-host ($InputFile + " doesn't exist. Terminating Script")
    out-file -filepath $Logfile -inputobject ($InputFile + " doesn't exist. Terminating Script")
    write-host "____________________________________________________"
    out-file -filepath $Logfile -append -inputobject "____________________________________________________"
    $finish = (get-date -format "MM-dd-yyyy HH:mm")
    write-host "$Task Complete - $finish"
    out-file -filepath $Logfile -append -inputobject "$Task Complete - $finish"
    $Prompt = new-object -comobject wscript.shell 
    out-file -filepath $Logfile -append -inputobject ($Prompt.popup("Failed to import XML file.  Click OK to complete the script.", 0)) 
}

# SQL Variables
If ($ConfigXML.Configuration.PSConfigSettings.SQL.ConfigAlias -eq "Yes")
{
    [boolean]$ConfigSQL = $True
}
Else
{
    [boolean]$ConfigSQL = $False
}
$SQLAlias = $ConfigXML.Configuration.PSConfigSettings.SQL.SQLAlias
$SQLSrvr = $ConfigXML.Configuration.PSConfigSettings.SQL.SQLSrvr
$TCPAlias = ("DBMSSOCN," + $SQLSrvr)
$x86 = "HKLM:\Software\Microsoft\MSSQLServer\Client\ConnectTo"
$x64 = "HKLM:\Software\Wow6432Node\Microsoft\MSSQLServer\Client\ConnectTo"

# Farm Variables
$FarmAcct = (New-Object System.Management.Automation.PSCredential ($ConfigXML.Configuration.PSConfigSettings.FarmConfig.FarmAcct),(ConvertTo-SecureString -String $ConfigXML.Configuration.PSConfigSettings.FarmConfig.FarmAcctPass -AsPlainText -Force) -ErrorVariable +ScriptErrors)
$ConfigDB = $ConfigXML.Configuration.PSConfigSettings.FarmConfig.ConfigDB
$Passphrase = (ConvertTo-SecureString $ConfigXML.Configuration.PSConfigSettings.FarmConfig.Passphrase -AsPlainText -force -ErrorVariable +ScriptErrors)

# Central Admin Variables
$AdminDB = $ConfigXML.Configuration.PSConfigSettings.CentralAdmin.DB
$CAPort = $ConfigXML.Configuration.PSConfigSettings.CentralAdmin.Port
$AuthProvider = $ConfigXML.Configuration.PSConfigSettings.CentralAdmin.Auth

# Outputing variables
write-host "____________________________________________________"
write-host "Intial Configuration variables Successfully assigned"
write-host "____________________________________________________"
write-host "----------------------------------------------------"
write-host "Script Variables"
write-host "----------------------------------------------------"
write-host "   Log file: $Logfile"
write-host "----------------------------------------------------"
write-host "SQL Variables"
write-host "----------------------------------------------------"
write-host "   Configure SQL Alias: $ConfigSQL"
write-host "   SQL Alias: $SQLAlias"
write-host "   SQL Server: $SQLSrvr"
write-host "   TCP Alias: $TCPAlias"
write-host "   x86 Registry Key: $x86"
write-host "   x64 Registry Key: $x64"
write-host "----------------------------------------------------"
write-host "Farm Variables"
write-host "----------------------------------------------------"
write-host "   Farm Account: $FarmAcct"
write-host "   Configuration Database Name: $ConfigDB"
write-host "   SharePoint Paraphrase: $ConfigXML.Configuration.PSConfigSettings.FarmConfig.Passphrase"
write-host "----------------------------------------------------"
write-host "Central Administration Variables"
write-host "----------------------------------------------------"
write-host "   Admin Content Database Name: $AdminDB"
write-host "   Central Administration Port: $CAPort"
write-host "   Authentication Provider: $AuthProvider"
write-host "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "Intial Configuration variables Successfully assigned"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Script Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Log file: $Logfile"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "SQL Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Configure SQL Alias: $ConfigSQL"
out-file -filepath $Logfile -append -inputobject "   SQL Alias: $SQLAlias"
out-file -filepath $Logfile -append -inputobject "   SQL Server: $SQLSrvr"
out-file -filepath $Logfile -append -inputobject "   TCP Alias: $TCPAlias"
out-file -filepath $Logfile -append -inputobject "   x86 Registry Key: $x86"
out-file -filepath $Logfile -append -inputobject "   x64 Registry Key: $x64"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Farm Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Farm Account: $FarmAcct"
out-file -filepath $Logfile -append -inputobject "   Configuration Database Name: $ConfigDB"
out-file -filepath $Logfile -append -inputobject "   SharePoint Paraphrase: $ConfigXML.Configuration.PSConfigSettings.FarmConfig.Passphrase"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Central Administration Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Admin Content Database Name: $AdminDB"
out-file -filepath $Logfile -append -inputobject "   Central Administration Port: $CAPort"
out-file -filepath $Logfile -append -inputobject "   Authentication Provider: $AuthProvider"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"

# Main Body
# Configuring SQL Aliases
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If ($ConfigSQL)
{
    write-host "Configuring SQL Aliases."
    out-file -filepath $Logfile -append -inputobject "Configuring SQL Aliases."
    If (test-path $x86)
    {
        write-host ("   " + $x86 + " already exists.")
        out-file -filepath $Logfile -append -inputobject ("   " + $x86 + " already exists.")
    }
    Else
    {
        write-host ("   Creating Key " + $x86 + ".")
        out-file -filepath $Logfile -append -inputobject ("   Creating " + $x86 + ".")
        out-file -filepath $Logfile -append -inputobject (New-Item $x86 -ErrorVariable +ScriptErrors)
    }
    If (test-path $x64)
    {
        write-host ("   " + $x64 + " already exists.")
        out-file -filepath $Logfile -append -inputobject ("   " + $x64 + " already exists.")
    }
    Else
    {
        write-host ("   Creating Key " + $x64 + ".")
        out-file -filepath $Logfile -append -inputobject ("   Creating " + $x64 + ".")
        out-file -filepath $Logfile -append -inputobject (New-Item $x64 -ErrorVariable +ScriptErrors)
    }
    $Alias = Get-ItemProperty $x86 $SQLAlias -ErrorAction SilentlyContinue
    If ($Alias -ne $null)
    {
        write-host "   SQL Alias already exists."
        out-file -filepath $Logfile -append -inputobject "   SQL Alias already exists."
    }
    Else
    {
        write-host "   Creating x86 SQL Alias."
        out-file -filepath $Logfile -append -inputobject "   Creating x86 SQL Alias."
        out-file -filepath $Logfile -append -inputobject (New-ItemProperty -Path $x86 -Name $SQLAlias -PropertyType String -Value $TCPAlias -ErrorVariable +ScriptErrors)
    }
    $Alias = Get-ItemProperty $x64 $SQLAlias -ErrorAction SilentlyContinue
    If ($Alias -ne $null)
    {
        write-host "   SQL Alias already exists."
        out-file -filepath $Logfile -append -inputobject "   SQL Alias already exists."
    }
    Else
    {
        write-host "   Creating x64 SQL Alias."
        out-file -filepath $Logfile -append -inputobject "   Creating x64 SQL Alias."
        out-file -filepath $Logfile -append -inputobject (New-ItemProperty -Path $x64 -Name $SQLAlias -PropertyType String -Value $TCPAlias -ErrorVariable +ScriptErrors)
    }
}
Else
{
    write-host "*** SQL Aliases will not be configured."
    out-file -filepath $Logfile -append -inputobject "*** SQL Aliases will not be configured."
}

# Adding PowerShell snap-in for SharePoint
write-host "____________________________________________________"
write-host "Adding SharePoint Snapin."
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "Adding SharePoint Snapin."
Add-PsSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue -ErrorVariable +ScriptErrors

# Configuring New SharePoint Farm
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
# **** If the farm hasn't been created yet, the following command will generate an error that can't be supressed.  You can safely ignore it. ****
$farm = Get-SPFarm -ErrorAction SilentlyContinue -ErrorVariable +ScriptErrors
if ($farm -ne $null) 
{
    write-host "SharePoint Farm already exists."
    out-file -filepath $Logfile -append -inputobject "SharePoint Farm already exists."
}
Else
{
    write-warning "The above error is expected because no farm has been created at this time."
    write-host "Creating SharePoint Farm."
    out-file -filepath $Logfile -append -inputobject "Creating SharePoint Farm."
    $Reset = $True
    out-file -filepath $Logfile -append -inputobject (New-SPConfigurationDatabase -DatabaseName $configDb -DatabaseServer $SQLAlias -AdministrationContentDatabaseName $AdminDB -passphrase $passphrase -FarmCredentials $FarmAcct -SkipRegisterAsDistributedCacheHost -ErrorVariable +ScriptErrors)
    $farm = Get-SPFarm -ErrorVariable +ScriptErrors
    if (!$farm -or $farm.Status -ne "Online") 
    {
        write-host "   Farm was not created or is not running."
        out-file -filepath $Logfile -append -inputobject "   Farm was not created or is not running."
        exit
        write-host "____________________________________________________"
        out-file -filepath $Logfile -append -inputobject "____________________________________________________"
        $finish = (get-date -format "MM-dd-yyyy HH:mm")
        write-host "$Task Complete - $finish"
        out-file -filepath $Logfile -append -inputobject "$Task Complete - $finish"
        $Prompt = new-object -comobject wscript.shell 
        out-file -filepath $Logfile -append -inputobject ($Prompt.popup("Failed to create the farm.  Click OK to complete the script.", 0)) 
    }
    Else
    {
        write-host "   Farm was successfully created."
        out-file -filepath $Logfile -append -inputobject "   Farm was successfully created."
    }
    write-host "   Securing SharePoint Resources"
    out-file -filepath $Logfile -append -inputobject "   Securing SharePoint Resources."
    out-file -filepath $Logfile -append -inputobject (Initialize-SPResourceSecurity -ErrorVariable +ScriptErrors)

    write-host "   Installing SharePoint Services."
    out-file -filepath $Logfile -append -inputobject "   Installing SharePoint Services."
    out-file -filepath $Logfile -append -inputobject (Install-SPService -ErrorVariable +ScriptErrors)

    write-host "   Installing SharePoint Features."
    out-file -filepath $Logfile -append -inputobject "   Installing SharePoint Features."
    out-file -filepath $Logfile -append -inputobject (Install-SPFeature -AllExistingFeatures -ErrorVariable +ScriptErrors)

    write-host "   Installing SharePoint Help Collection."
    out-file -filepath $Logfile -append -inputobject "   Installing SharePoint Help Collection."
    out-file -filepath $Logfile -append -inputobject (Install-SPHelpCollection -All -ErrorVariable +ScriptErrors)

    write-host "   Creating Central Administration Web Application."
    out-file -filepath $Logfile -append -inputobject "   Creating Central Administration Web Application."
    out-file -filepath $Logfile -append -inputobject (New-SPCentralAdministration -Port $CAPort -WindowsAuthProvider $AuthProvider -ErrorVariable +ScriptErrors)

    write-host "   Installing Central Administration Content."
    out-file -filepath $Logfile -append -inputobject "   Installing Central Administration Content."
    out-file -filepath $Logfile -append -inputobject (Install-SPApplicationContent -ErrorVariable +ScriptErrors)
}

# Logging Errors to log file
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "Errors during Script:"
out-file -filepath $Logfile -append -inputobject ($ScriptErrors)

# Finishing Script
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
$finish = (get-date -format "MM-dd-yyyy HH:mm")
write-host "$Task Complete - $finish."
out-file -filepath $Logfile -append -inputobject "$Task Complete - $finish."

$Prompt = new-object -comobject wscript.shell 
out-file -filepath $Logfile -append -inputobject ($Prompt.popup("Farm Configuration is complete.  Click OK to complete the script and open Central Administration.  After adding all other servers to the SharePoint Farm, run the farm configuration script.", 0))
& "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\BIN\psconfigui.exe" -cmd showcentraladmin