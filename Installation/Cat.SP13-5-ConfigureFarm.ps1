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
$Logfile = "d:\Installation\Scripts\_Log Files\FarmConfiguration_$date.log.txt"
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

# Global Variables
$FQDNExt = $ConfigXML.Configuration.ServerConfig.FQDN
$SQLAlias = $ConfigXML.Configuration.PSConfigSettings.SQL.SQLAlias

# Cache Config Variables
$CacheSuper = $ConfigXML.Configuration.GlobalSettings.CacheAccts.Super
$CacheRead = $ConfigXML.Configuration.GlobalSettings.CacheAccts.Reader
$ClaimSuper = ($ConfigXML.Configuration.GlobalSettings.CacheAccts.Claim + $ConfigXML.Configuration.GlobalSettings.CacheAccts.Super)
$ClaimRead = ($ConfigXML.Configuration.GlobalSettings.CacheAccts.Claim + $ConfigXML.Configuration.GlobalSettings.CacheAccts.Reader)

# Outgoing Email Variables
$SMTP = $ConfigXML.Configuration.GlobalSettings.OutgoingEmail.SMTPSrv
$From = $ConfigXML.Configuration.GlobalSettings.OutgoingEmail.From
$Reply = $ConfigXML.Configuration.GlobalSettings.OutgoingEmail.ReplyTo
[Int]$CharSet = $ConfigXML.Configuration.GlobalSettings.OutgoingEmail.CharSet 

# Site Structure Variables
If ($ConfigXML.Configuration.SiteStructure.CreateWebApps -eq "Yes")
{
    [boolean]$SSWebApp = $True
}
Else
{
    [boolean]$SSWebApp = $False
}
If ($ConfigXML.Configuration.SiteStructure.CreateSiteCollections -eq "Yes")
{
    [boolean]$SSSites = $True
}
Else
{
    [boolean]$SSSites = $False
}
If ($ConfigXML.Configuration.SiteStructure.CreateSubSites -eq "Yes")
{
    [boolean]$SSWebs = $True
}
Else
{
    [boolean]$SSWebs = $False
}

# Diagnostic Logging Variables
$DiagLoc = $ConfigXML.Configuration.GlobalSettings.DiagLogging.Location
If ($ConfigXML.Configuration.GlobalSettings.DiagLogging.MaxEnable -eq "Yes")
{
    [boolean]$DiagMaxEnable = $True
}
Else
{
    [boolean]$DiagMaxEnable = $False
}
[int]$DiagMaxSpace = $ConfigXML.Configuration.GlobalSettings.DiagLogging.MaxSpace
[int]$DiagDays = $ConfigXML.Configuration.GlobalSettings.DiagLogging.DaysToStore
$DiagTrace = $ConfigXML.Configuration.GlobalSettings.DiagLogging.TraceSeverity
If ($ConfigXML.Configuration.GlobalSettings.DiagLogging.FloodEnable -eq "Yes")
{
    [boolean]$DiagFloodEnable = $True
}
Else
{
    [boolean]$DiagFloodEnable = $False
}

# Request Management Variables
If ($ConfigXML.Configuration.Services.RequestManagement.Config -eq "Yes")
{
    [boolean]$RMEnable = $True
}
Else
{
    [boolean]$RMEnable = $False
}
If ($ConfigXML.Configuration.Services.RequestManagement.RoutingEnabled -eq "Yes")
{
    [boolean]$RMRouting = $True
}
Else
{
    [boolean]$RMRouting = $False
}
If ($ConfigXML.Configuration.Services.RequestManagement.ThrottlingEnabled -eq "Yes")
{
    [boolean]$RMThrottling = $True
}
Else
{
    [boolean]$RMThrottling = $False
}
$RMScheme = $ConfigXML.Configuration.Services.RequestManagement.WeightScheme

# State Service Application Varaibles
$StateSA = $ConfigXML.Configuration.ServiceApps.StateSrv.Name
$StateDB = $ConfigXML.Configuration.ServiceApps.StateSrv.DB

# Usage & Health Service Application Variables
$UsageSA = $ConfigXML.Configuration.ServiceApps.UsageHealthSrv.Name
$UsageDB = $ConfigXML.Configuration.ServiceApps.UsageHealthSrv.DB
If ($ConfigXML.Configuration.ServiceApps.UsageHealthSrv.Enabled -eq "Yes")
{
    [boolean]$UsageEnable = $True
}
Else
{
    [boolean]$UsageEnable = $False
}
$UsageLoc = $ConfigXML.Configuration.ServiceApps.UsageHealthSrv.Location

# Subscription Settings Variables
$SubscriptionName = $ConfigXML.Configuration.ServiceApps.SubscriptionSettings.Name
$SubscriptionPool = $ConfigXML.Configuration.ServiceApps.SubscriptionSettings.Pool
$SubscriptionAcct = $ConfigXML.Configuration.ServiceApps.SubscriptionSettings.Acct
$SubscriptionDB = $ConfigXML.Configuration.ServiceApps.SubscriptionSettings.DB

# Secure Store Variables
If ($ConfigXML.Configuration.ServiceApps.SecureStore.Config -eq "Yes")
{
    [boolean]$SecureConfig = $True
}
Else
{
    [boolean]$SecureConfig = $False
}
$SecureName = $ConfigXML.Configuration.ServiceApps.SecureStore.Name
$SecureAcct = $ConfigXML.Configuration.ServiceApps.SecureStore.Acct
$SecurePool = $ConfigXML.Configuration.ServiceApps.SecureStore.Pool
$SecureDB = $ConfigXML.Configuration.ServiceApps.SecureStore.DB
If ($ConfigXML.Configuration.ServiceApps.SecureStore.Auditing -eq "Yes")
{
    [boolean]$SecureAudit = $True
}
Else
{
    [boolean]$SecureAudit = $False
}
$SecureDays = $ConfigXML.Configuration.ServiceApps.SecureStore.AuditSize

# Managed Metadata Service Application variables
If ($ConfigXML.Configuration.ServiceApps.MetadataSrv.ConfigMetadata -eq "Yes")
{
    [boolean]$ConfigMetadata = $True
}
Else
{
    [boolean]$ConfigMetadata = $False
}
$MetadataSA = $ConfigXML.Configuration.ServiceApps.MetadataSrv.Name
$MetadataAcct = $ConfigXML.Configuration.ServiceApps.MetadataSrv.Acct
$MetadataPool = $ConfigXML.Configuration.ServiceApps.MetadataSrv.Pool
$MetadataSrvr = $ConfigXML.Configuration.ServiceApps.MetadataSrv.Srvr
$MetadataDB = $ConfigXML.Configuration.ServiceApps.MetadataSrv.DB
$CTHub = $ConfigXML.Configuration.ServiceApps.MetadataSrv.Hub

# Excel Calculations Service Application variables
If ($ConfigXML.Configuration.ServiceApps.ExcelCalc.Config -eq "Yes")
{
    [boolean]$ExcelEnable = $True
}
Else
{
    [boolean]$ExcelEnable = $False
}
$ExcelName = $ConfigXML.Configuration.ServiceApps.ExcelCalc.Name
$ExcelAcct = $ConfigXML.Configuration.ServiceApps.ExcelCalc.Acct
$ExcelPool = $ConfigXML.Configuration.ServiceApps.ExcelCalc.Pool

# Search Service Application Variables
If ($ConfigXML.Configuration.ServiceApps.SearchSrv.ConfigSearch -eq "Yes")
{
    [boolean]$ConfigSearch = $True
}
Else
{
    [boolean]$ConfigSearch = $False
}
$SearchPool = $ConfigXML.Configuration.ServiceApps.SearchSrv.Pool
$SearchAcct = $ConfigXML.Configuration.ServiceApps.SearchSrv.Acct
$SearchSA = $ConfigXML.Configuration.ServiceApps.SearchSrv.Name
$SearchDB = $ConfigXML.Configuration.ServiceApps.SearchSrv.DB
$SearchIndex = $ConfigXML.Configuration.ServiceApps.SearchSrv.Index.StorageLocation

# User Profile Application Variables
If ($ConfigXML.Configuration.ServiceApps.UserProfiles.Config -eq "Yes")
{
    [boolean]$UPSConfig = $True
}
Else
{
    [boolean]$UPSConfig = $False
}
$UPSName = $ConfigXML.Configuration.ServiceApps.UserProfiles.Name
$UPSPool = $ConfigXML.Configuration.ServiceApps.UserProfiles.Pool
$UPSAcct = $ConfigXML.Configuration.ServiceApps.UserProfiles.Acct
$UPSDB = $ConfigXML.Configuration.ServiceApps.UserProfiles.DB
$UPSWA = $ConfigXML.Configuration.ServiceApps.UserProfiles.WebApp
$UPSPath = $ConfigXML.Configuration.ServiceApps.UserProfiles.ManagedPath
If ($ConfigXML.Configuration.ServiceApps.UserProfiles.NetBios -eq "Yes")
{
    [boolean]$UPSNet = $True
}
Else
{
    [boolean]$UPSNet = $False
}

# App Management Service Application Variables
If ($ConfigXML.Configuration.ServiceApps.AppManagement.Config -eq "Yes")
{
    [boolean]$AMConfig = $True
}
Else
{
    [boolean]$AMConfig = $False
}
$AMName = $ConfigXML.Configuration.ServiceApps.AppManagement.Name
$AMAcct = $ConfigXML.Configuration.ServiceApps.AppManagement.Acct
$AMPool = $ConfigXML.Configuration.ServiceApps.AppManagement.Pool
$AMDB = $ConfigXML.Configuration.ServiceApps.AppManagement.DB
$AMDomain = $ConfigXML.Configuration.ServiceApps.AppManagement.Domain
$AMPrefix = $ConfigXML.Configuration.ServiceApps.AppManagement.Prefix
If ($ConfigXML.Configuration.ServiceApps.AppManagement.InternetFacing -eq "Yes")
{
    [boolean]$AMEndPoint = $True
}
Else
{
    [boolean]$AMEndPoint = $False
}

# Solution Variables
$SolutionLocation = $ConfigXML.Configuration.Solutions.Location

# Branding variables
If ($ConfigXML.Configuration.Solutions.Branding.Deploy -eq "Yes")
{
    [boolean]$BrandingConfig = $True
}
Else
{
    [boolean]$BrandingConfig = $False
}
$BrandingWSP = $ConfigXML.Configuration.Solutions.Branding.WSP
If ($ConfigXML.Configuration.Solutions.Branding.AllSites -eq "Yes")
{
    [boolean]$BrandingAll = $True
}
Else
{
    [boolean]$BrandingAll = $False
}

# Outputing variables
write-host "____________________________________________________"
write-host "Intial Configuration variables Successfully assigned"
write-host "____________________________________________________"
write-host "----------------------------------------------------"
write-host "Script Variables"
write-host "----------------------------------------------------"
write-host "   Log file: $Logfile"
write-host "----------------------------------------------------"
write-host "Global Variables"
write-host "----------------------------------------------------"
write-host "   FQDN Extension: $FQDNExt"
write-host "   SQL Alias: $SQLAlias"
write-host "----------------------------------------------------"
write-host "Outgoing Email Variables"
write-host "----------------------------------------------------"
write-host "   SMTP Server: $SMTP"
write-host "   From Address: $From"
write-host "   Reply To Address: $Reply"
write-host "   Character Set: $CharSet"
write-host "----------------------------------------------------"
write-host "Cache User Variables"
write-host "----------------------------------------------------"
write-host "   Super User: $CacheSuper"
write-host "   Read User: $CacheRead"
write-host "   Claim Super User: $ClaimSuper"
write-host "   Claim Read User: $ClaimRead"
write-host "----------------------------------------------------"
write-host "Diagnostic Logging Variables"
write-host "----------------------------------------------------"
write-host "   Log file Location: $DiagLoc"
write-host "   Enable Max Storage: $DiagMaxEnable"
write-host "   Max Storage Space: $DiagMaxSpace"
write-host "   Days to Store Log Files: $DiagDays"
write-host "   Minimum Trace Severity: $DiagTrace"
write-host "   Enable Flood Protection: $DiagFloodEnable"
write-host "----------------------------------------------------"
write-host "Site Structure Variables"
write-host "----------------------------------------------------"
write-host "   Create Web Applications: $SSWebApp"
write-host "   Create Site Collections: $SSSites"
write-host "   Create Sub Sites: $SSWebs"
write-host "----------------------------------------------------"
write-host "Request Management Variables"
write-host "----------------------------------------------------"
write-host "   Configure Request Mangement: $RMEnable"
write-host "   Routing Enabled: $RMRouting"
write-host "   Throttling Enabled: $RMThrottling"
write-host "   Routing Weight Scheme: $RMScheme"
write-host "----------------------------------------------------"
write-host "State Service Application Variables"
write-host "----------------------------------------------------"
write-host "   State Service: $StateSA"
write-host "   State Service DB: $StateDB"
write-host "----------------------------------------------------"
write-host "Usage & Health Service Application Variables"
write-host "----------------------------------------------------"
write-host "   Usage & Health Service: $UsageSA"
write-host "   Usage & Health Service DB: $UsageDB"
write-host "   Enable Data Collection: $UsageEnable"
write-host "   Location: $UsageLoc"
write-host "----------------------------------------------------"
write-host "Subscription Settings Service Application Variables"
write-host "----------------------------------------------------"
write-host "   Subscription Settings: $SubscriptionName"
write-host "   Subscription Settings App Pool: $SubscriptionPool"
write-host "   Subscription Settings Account: $SubscriptionAcct"
write-host "   Subscription Settings Database: $SubscriptionDB"
write-host "----------------------------------------------------"
write-host "Secure Store Service Application Variables"
write-host "----------------------------------------------------"
write-host "   Configure Secure Store: $SecureConfig"
write-host "   Secure Store Service Application: $SecureName"
write-host "   Secure Store Service Account: $SecureAcct"
write-host "   Secure Store App Pool: $SecurePool"
write-host "   Secure Store Database: $SecureDB"
write-host "   Secure Store Auditing: $SecureAudit"
write-host "   Secure Store Audit Maximum Days: $SecureDays"
write-host "----------------------------------------------------"
write-host "Managed Metadata Service Application Variables"
write-host "----------------------------------------------------"
write-host "   Configure Managed Metadata Service Application: $ConfigMetadata"
write-host "   Managed Metadata Service Application: $MetadataSA"
write-host "   Managed Metadata Service Account: $MetadataAcct"
write-host "   Managed Metadata Application Pool: $MetadataPool"
write-host "   Managed Metadata Database: $MetadataDB"
write-host "   Content Type Hub: $CTHub"
write-host "----------------------------------------------------"
write-host "Excel Calculations Service Application Variables"
write-host "----------------------------------------------------"
write-host "   Configure Excel Calculations service application: $ExcelEnable"
write-host "   Excel Calculations Service: $ExcelName"
write-host "   Excel Calculations Account: $ExcelAcct"
write-host "   Excel Calculations Application Pool: $ExcelPool"
write-host "----------------------------------------------------"
write-host "Search Service Application Variables"
write-host "----------------------------------------------------"
write-host "   Configure Search Service Application: $ConfigSearch"
write-host "   Search Application Pool: $SearchPool"
write-host "   Search Service Account: $SearchAcct"
write-host "   Search Service Application: $SearchSA"
write-host "   Search Database: $SearchDB"
write-host "----------------------------------------------------"
write-host "User Profile Service Application Variables"
write-host "----------------------------------------------------"
write-host "   Configure UPS: $UPSConfig"
write-host "   UPS Service Application: $UPSName"
write-host "   UPS Application Pool: $UPSPool"
write-host "   UPS Account: $UPSAcct"
write-host "   UPS Database Prefix: $UPSDB"
write-host "   UPS Web Application: $UPSWA"
write-host "   UPS Managed Path: $UPSPath"
write-host "   Configure NetBIOS: $UPSNet"
write-host "----------------------------------------------------"
write-host "App Management Service Application Variables"
write-host "----------------------------------------------------"
write-host "   Configure App Mangement: $AMConfig"
write-host "   App Management Service Application: $AMName"
write-host "   App Management Account: $AMAcct"
write-host "   App Management Application Pool: $AMPool"
write-host "   App Management Database: $AMDB"
write-host "   App Domain: $AMDomain"
write-host "   App Prefix: $AMPrefix"
write-host "   Allow Internet End Points: $AMEndPoint"
write-host "----------------------------------------------------"
write-host "Solution Variables Variables"
write-host "----------------------------------------------------"
write-host "   Storage Location: $SolutionLocation"
write-host "----------------------------------------------------"
write-host "Branding Variables"
write-host "----------------------------------------------------"
write-host "   Branding WSP: $BrandingWSP"
write-host "   Brand all Sites: $BrandingAll"
write-host "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "Intial Configuration variables Successfully assigned"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Script Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Log file: $Logfile"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Global Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   FQDN Extension: $FQDNExt"
out-file -filepath $Logfile -append -inputobject "   SQL Alias: $SQLAlias"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Outgoing Email Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   SMTP Server: $SMTP"
out-file -filepath $Logfile -append -inputobject "   From Address: $From"
out-file -filepath $Logfile -append -inputobject "   Reply To Address: $Reply"
out-file -filepath $Logfile -append -inputobject "   Character Set: $CharSet"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Cache User Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Super User: $CacheSuper"
out-file -filepath $Logfile -append -inputobject "   Read User: $CacheRead"
out-file -filepath $Logfile -append -inputobject "   Claim Super User: $ClaimSuper"
out-file -filepath $Logfile -append -inputobject "   Claim Read User: $ClaimRead"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Diagnostic Logging Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Log file Location: $DiagLoc"
out-file -filepath $Logfile -append -inputobject "   Enable Max Storage: $DiagMaxEnable"
out-file -filepath $Logfile -append -inputobject "   Max Storage Space: $DiagMaxSpace"
out-file -filepath $Logfile -append -inputobject "   Days to Store Log Files: $DiagDays"
out-file -filepath $Logfile -append -inputobject "   Minimum Trace Severity: $DiagTrace"
out-file -filepath $Logfile -append -inputobject "   Enable Flood Protection: $DiagFloodEnable"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Site Structure Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Create Web Applications: $SSWebApp"
out-file -filepath $Logfile -append -inputobject "   Create Site Collections: $SSSites"
out-file -filepath $Logfile -append -inputobject "   Create Sub Sites: $SSWebs"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Request Management Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Configure Request Mangement: $RMEnable"
out-file -filepath $Logfile -append -inputobject "   Routing Enabled: $RMRouting"
out-file -filepath $Logfile -append -inputobject "   Throttling Enabled: $RMThrottling"
out-file -filepath $Logfile -append -inputobject "   Routing Weight Scheme: $RMScheme"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "State Service Application Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   State Service: $StateSA"
out-file -filepath $Logfile -append -inputobject "   State Service DB: $StateDB"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Usage & Health Service Application Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Usage & Health Service: $UsageSA"
out-file -filepath $Logfile -append -inputobject "   Usage & Health Service DB: $UsageDB"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Subscription Settings Service Application Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Subscription Settings: $SubscriptionName"
out-file -filepath $Logfile -append -inputobject "   Subscription Settings App Pool: $SubscriptionPool"
out-file -filepath $Logfile -append -inputobject "   Subscription Settings Account: $SubscriptionAcct"
out-file -filepath $Logfile -append -inputobject "   Subscription Settings Database: $SubscriptionDB"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Secure Store Service Application Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Configure Secure Store: $SecureConfig"
out-file -filepath $Logfile -append -inputobject "   Secure Store Service Application: $SecureName"
out-file -filepath $Logfile -append -inputobject "   Secure Store Service Account: $SecureAcct"
out-file -filepath $Logfile -append -inputobject "   Secure Store App Pool: $SecurePool"
out-file -filepath $Logfile -append -inputobject "   Secure Store Database: $SecureDB"
out-file -filepath $Logfile -append -inputobject "   Secure Store Auditing: $SecureAudit"
out-file -filepath $Logfile -append -inputobject "   Secure Store Audit Maximum Days: $SecureDays"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Managed Metadata Service Application Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Configure Managed Metadata Service Application: $ConfigMetadata"
out-file -filepath $Logfile -append -inputobject "   Managed Metadata Service Application: $MetadataSA"
out-file -filepath $Logfile -append -inputobject "   Managed Metadata Service Account: $MetadataAcct"
out-file -filepath $Logfile -append -inputobject "   Managed Metadata Application Pool: $MetadataPool"
out-file -filepath $Logfile -append -inputobject "   Managed Metadata Database: $MetadataDB"
out-file -filepath $Logfile -append -inputobject "   Content Type Hub: $CTHub"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Excel Calculations Service Application Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Configure Excel Calculations service application: $ExcelEnable"
out-file -filepath $Logfile -append -inputobject "   Excel Calculations Service: $ExcelName"
out-file -filepath $Logfile -append -inputobject "   Excel Calculations Account: $ExcelAcct"
out-file -filepath $Logfile -append -inputobject "   Excel Calculations Application Pool: $ExcelPool"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Search Service Application Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Configure Search Service Application: $ConfigSearch"
out-file -filepath $Logfile -append -inputobject "   Search Application Pool: $SearchPool"
out-file -filepath $Logfile -append -inputobject "   Search Service Account: $SearchAcct"
out-file -filepath $Logfile -append -inputobject "   Search Service Application: $SearchSA"
out-file -filepath $Logfile -append -inputobject "   Search Database: $SearchDB"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "User Profile Service Application Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Configure UPS: $UPSConfig"
out-file -filepath $Logfile -append -inputobject "   UPS Service Application: $UPSName"
out-file -filepath $Logfile -append -inputobject "   UPS Application Pool: $UPSPool"
out-file -filepath $Logfile -append -inputobject "   UPS Account: $UPSAcct"
out-file -filepath $Logfile -append -inputobject "   UPS Database Prefix: $UPSDB"
out-file -filepath $Logfile -append -inputobject "   UPS Web Application: $UPSWA"
out-file -filepath $Logfile -append -inputobject "   UPS Managed Path: $UPSPath"
out-file -filepath $Logfile -append -inputobject "   Configure NetBIOS: $UPSNet"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "App Management Service Application Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Configure App Mangement: $AMConfig"
out-file -filepath $Logfile -append -inputobject "   App Management Service Application: $AMName"
out-file -filepath $Logfile -append -inputobject "   App Management Account: $AMAcct"
out-file -filepath $Logfile -append -inputobject "   App Management Application Pool: $AMPool"
out-file -filepath $Logfile -append -inputobject "   App Management Database: $AMDB"
out-file -filepath $Logfile -append -inputobject "   App Domain: $AMDomain"
out-file -filepath $Logfile -append -inputobject "   App Prefix: $AMPrefix"
out-file -filepath $Logfile -append -inputobject "   Allow Internet End Points: $AMEndPoint"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Solution Variables Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Storage Location: $SolutionLocation"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Branding Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Branding WSP: $BrandingWSP"
out-file -filepath $Logfile -append -inputobject "   Brand all Sites: $BrandingAll"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"

# Main Body
# Adding PowerShell snap-in for SharePoint
write-host "____________________________________________________"
write-host "Adding SharePoint Snapin."
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "Adding SharePoint Snapin."
Add-PsSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue -ErrorVariable +ScriptErrors

# Getting Farm configuration
$Farm = Get-SPFarm

# Global Settings
# Configuring Outgoing Email
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
$CA = Get-SPWebApplication -IncludeCentralAdministration | Where { $_.IsAdministrationWebApplication }
If (($CA.OutboundMailCodePage.Equals($CharSet)) -and ($CA.OutboundMailReplyToAddress -eq $Reply) -and ($CA.OutboundMailSenderAddress -eq $From) -and ($CA.OutboundMailServiceInstance.Server.Address -eq $SMTP))
{
    write-host "Outgoing Email has already been configured."
    out-file -filepath $Logfile -append -inputobject "Outgoing Email has already been configured." 
}
Else
{
    write-host "Configuring Outgoing Email."
    out-file -filepath $Logfile -append -inputobject "Configuring Outgoing Email." 
    $CA.UpdateMailSettings($SMTP, $From, $Reply, $CharSet) 
}

# Configuring Diagnostic Logging
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
$Log = Get-SPDiagnosticConfig
$Trace = ((get-sploglevel)[0])
If (($Log.LogLocation -eq $DiagLoc) -and ($Log.LogMaxDiskSpaceUsageEnabled -eq $DiagMaxEnable) -and ($Log.LogDiskSpaceUsageGB -eq $DiagMaxSpace) -and ($Log.DaysToKeepLogs -eq $DiagDays) -and ($Log.EventLogFloodProtectionEnabled -eq $DiagFloodEnable) -and ($Trace.TraceSeverity -eq $DiagTrace))
{
    write-host "Diagnostic Logging has already been configured"
    out-file -filepath $Logfile -append -inputobject "Diagnostic Logging has already been configured" 
}
Else
{
    write-host "Configuring Diagnostic Logging"
    out-file -filepath $Logfile -append -inputobject "Configuring Diagnostic Logging" 
    Set-SPDiagnosticConfig -LogLocation $DiagLoc -LogMaxDiskSpaceUsageEnabled:$DiagMaxEnable -LogDiskSpaceUsageGB $DiagMaxSpace -DaysToKeepLogs $DiagDays -EventLogFloodProtectionEnabled:$DiagFloodEnable
    Set-SPLogLevel -TraceSeverity $DiagTrace
}

# Creating Managed Accounts
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
write-host "Creating Managed Accounts."
out-file -filepath $Logfile -append -inputobject "Creating Managed Accounts."
foreach ($User in $ConfigXML.Configuration.GlobalSettings.ManagedAccts.Acct)
{
    $ManAcct = Get-SPManagedAccount $User.User -erroraction SilentlyContinue
    If ($ManAcct -ne $null)
    {
        write-host ("   " + $User.User + " already exists.")
        out-file -filepath $Logfile -append -inputobject ("   " + $User.User + " already exists.")
    }
    Else
    {
        write-host ("   Creating " + $User.User + ".")
        out-file -filepath $Logfile -append -inputobject ("   Creating " + $User.User + ".")
        $Acct = (New-Object System.Management.Automation.PSCredential ($User.User),(ConvertTo-SecureString -String $User.Pass -AsPlainText -Force) -ErrorVariable +ScriptErrors)
        out-file -filepath $Logfile -append -inputobject (New-SPManagedAccount $Acct -ErrorVariable +ScriptErrors)
    }
}

# Site Structure
# Creating Web Applications
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If($SSWebApp)
{
    write-host "Creating Web Applications."
    out-file -filepath $Logfile -append -inputobject "Creating Web Applications."
    write-host "   Importing Web Administration module."
    out-file -filepath $Logfile -append -inputobject "   Importing Web Administration module."
    Import-Module WebAdministration -ErrorVariable +ScriptErrors
    foreach ($App in $ConfigXML.Configuration.SiteStructure.WebApps.App)
    {
        $WAAcct = $App.Acct
        $WA = $App.Name
        $HostHeader = $App.HostHeader
        If ($App.SSL -eq "Yes")
        {
            [boolean]$SSL = $True
        }
        Else
        {
            [boolean]$SSL = $False
        }
        $Port = $App.Port
        $WAPool = $App.Pool
        $WADB = $App.DB
        $FromEmail = $App.FromEmail
        $Acct = Get-SPManagedAccount $WAAcct -ErrorVariable +ScriptErrors
        $WebApp = Get-SPWebApplication $WA -erroraction SilentlyContinue
        If ($WebApp -ne $null)
        {
            write-host ("   " + $WA + " already exists.")
            out-file -filepath $Logfile -append -inputobject ("   " + $WA + " already exists.")
        }
        Else
        {
            write-host ("   Creating Web App: " + $WA + ".")
            out-file -filepath $Logfile -append -inputobject ("   Creating Web App: " + $WA + ".")
            If ($App.ClaimsBased -eq "Yes")
            {
                $Provider = New-SPAuthenticationProvider
                If (Test-Path IIS:\\AppPools\$WAPool)
                {
                    $WebApp = New-SPWebApplication -Name $WA -Port $Port -HostHeader $HostHeader -ApplicationPool $WAPool -DatabaseName $WADB -AuthenticationProvider $Provider -SecureSocketslayer:$SSL -ErrorVariable +ScriptErrors
                    $Reset = $True
                }
                Else
                {
                    $WebApp = New-SPWebApplication -Name $WA -Port $Port -HostHeader $HostHeader -ApplicationPool $WAPool -ApplicationPoolAccount $Acct -DatabaseName $WADB -AuthenticationProvider $Provider -SecureSocketslayer:$SSL -ErrorVariable +ScriptErrors
                    $Reset = $True
                }
            }
            Else
            {
                If (Test-Path IIS:\\AppPools\$WAPool)
                {
                    # For SSL add "-SecureSocketslayer:$True" to the New Web App Command
                    $WebApp = New-SPWebApplication -Name $WA -Port $Port -HostHeader $HostHeader -ApplicationPool $WAPool -DatabaseName $WADB -SecureSocketslayer:$SSL -ErrorVariable +ScriptErrors
                    $Reset = $True
                }
                Else
                {
                    # For SSL add "-SecureSocketslayer:$True" to the New Web App Command
                    $WebApp = New-SPWebApplication -Name $WA -Port $Port -HostHeader $HostHeader -ApplicationPool $WAPool -ApplicationPoolAccount $Acct -DatabaseName $WADB -SecureSocketslayer:$SSL -ErrorVariable +ScriptErrors
                    $Reset = $True
                }
            }
        }
        foreach ($Path in $App.ManagedPaths.Path)
        {
            If ($Path.Type -eq "Explicit")
            {
                [boolean]$MPType = $True
            }
            Else
            {
                [boolean]$MPType = $False
            }
            $Managed = Get-SPManagedPath -WebApplication $WebApp | Where{$_.Name -eq $Path.URL} -erroraction SilentlyContinue
            If($Managed -ne $null)
            {
                write-host ("      Managed Path Already Exists for " + $Path.URL + ".")
                out-file -filepath $Logfile -append -inputobject ("      Managed Path Already Exists for " + $Path.URL + ".")
            }
            Else
            {
                write-host ("      Creating managed path for " + $Path.URL + ".")
                out-file -filepath $Logfile -append -inputobject ("      Creating managed path for " + $Path.URL + ".")
                out-file -filepath $Logfile -append -inputobject (New-SPManagedPath $Path.URL -WebApplication $WebApp -Explicit:$MPType -ErrorVariable +ScriptErrors)
            }
        }
        foreach ($CDB in $App.AdditionalContentDBs.DB)
        {
            $DB = Get-SPContentDatabase -WebApplication $WebApp | Where{$_.Name -eq $CDB} -erroraction SilentlyContinue
            If($DB -ne $null)
            {
                write-host ("      Content Database Already Exists for " + $CDB + ".")
                out-file -filepath $Logfile -append -inputobject ("      Content Database Already Exists for " + $CDB + ".")
            }
            Else
            {
                write-host ("      Creating database for " + $CDB + ".")
                out-file -filepath $Logfile -append -inputobject ("      Creating database for " + $CDB + ".")
                out-file -filepath $Logfile -append -inputobject (New-SPContentDatabase $CDB -WebApplication $WebApp -ErrorVariable +ScriptErrors)
            }
        }
    } 
    If ($Reset)
    {
        write-host "   Reseting IIS."
        out-file -filepath $Logfile -append -inputobject "   Reseting IIS."
        out-file -filepath $Logfile -append -inputobject (IISReset)
        $Reset = $False
    }

    # Configuring Web Application Specific Email Settings
    write-host "____________________________________________________"
    out-file -filepath $Logfile -append -inputobject "____________________________________________________"
    write-host "Configuring Web Application specific email settings."
    out-file -filepath $Logfile -append -inputobject "Configuring Web Application specific email settings."
    foreach ($App in $ConfigXML.Configuration.SiteStructure.WebApps.App)
    {
        $WA = $App.Name
        $FromEmail = $App.FromEmail
        $WebApp = Get-SPWebApplication $WA -erroraction SilentlyContinue
        If (($WebApp.OutboundMailSenderAddress -eq $FromEmail) -or ($FromEmail -eq ""))
        {
            write-host ("   " + $WA + " email already Configured.")
            out-file -filepath $Logfile -append -inputobject ("   " + $WA + " email already Configured.")
        }
        Else
        {
            write-host ("   Configuring " + $WA + " email.")
            out-file -filepath $Logfile -append -inputobject ("   Configuring " + $WA + " email.")
            $WebApp.OutboundMailSenderAddress = $FromEmail
            $WebApp.Update()
        }
    }

    # Configuring Cache Users on each Web Application
    write-host "____________________________________________________"
    out-file -filepath $Logfile -append -inputobject "____________________________________________________"
    write-host "Configuring Cache Users."
    out-file -filepath $Logfile -append -inputobject "Configuring Cache Users."
    $WAs = Get-SPWebApplication -IncludeCentralAdministration
    foreach ($WA in $WAs) 
    {
        $Read = $WA.Properties["portalsuperreaderaccount"]
        $Super = $WA.Properties["portalsuperuseraccount"]
        If ($WA.UseClaimsAuthentication)
        {
            If (($Read -eq $ClaimRead) -and ($Super -eq $ClaimSuper))
            {
                write-host ("   Cache users are already configured on " + $WA.DisplayName + ".")
                out-file -filepath $Logfile -append -inputobject ("   Cache users are already configured on " + $WA.DisplayName + ".")
            }
            Else
            {
                write-host ("   Configuring User Policies for " + $WA.DisplayName + ".")
                out-file -filepath $Logfile -append -inputobject ("   Configuring User Policies for " + $WA.DisplayName + ".")
                $Reset = $True
                $policy = $WA.Policies.Add($ClaimSuper, $ClaimSuper)
                $policy.PolicyRoleBindings.Add($WA.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullControl))
                $WA.Update()
                $policy = $WA.Policies.Add($ClaimRead, $CLaimRead)
                $policy.PolicyRoleBindings.Add($WA.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullRead))
                $WA.Update()
                write-host ("   Configuring Configuring Super User on " + $WA.DisplayName + ".")
                out-file -filepath $Logfile -append -inputobject ("   Configuring Configuring Super User on " + $WA.DisplayName + ".")
                $WA.Properties["portalsuperuseraccount"] = $ClaimSuper
                write-host ("   Configuring Configuring Super Reader on " + $WA.DisplayName + " using " + $ClaimRead + ".")
                out-file -filepath $Logfile -append -inputobject ("   Configuring Configuring Super Reader on " + $WA.DisplayName + ".")
                $WA.Properties["portalsuperreaderaccount"] = $ClaimRead
                $WA.Update()
            }
        }
        Else
        {
            If (($Read -eq $CacheRead) -and ($Super -eq $CacheSuper))
            {
                write-host ("   Cache users are already configured on " + $WA.DisplayName + ".")
                out-file -filepath $Logfile -append -inputobject ("   Cache users are already configured on " + $WA.DisplayName + ".")
            }
            Else
            {
                write-host ("   Configuring User Policies for " + $WA.DisplayName + ".")
                out-file -filepath $Logfile -append -inputobject ("   Configuring User Policies for " + $WA.DisplayName + ".")
                $policy = $WA.Policies.Add($CacheSuper, $CacheSuper)
                $policy.PolicyRoleBindings.Add($WA.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullControl))
                $WA.Update()
                $policy = $WA.Policies.Add($CacheRead, $CacheRead)
                $policy.PolicyRoleBindings.Add($WA.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullRead))
                $WA.Update()
                write-host ("   Configuring Configuring Super User on " + $WA.DisplayName + " using " + $CacheSuper + ".")
                out-file -filepath $Logfile -append -inputobject ("   Configuring Configuring Super User on " + $WA.DisplayName + ".")
                $WA.Properties["portalsuperuseraccount"] = $CacheSuper
                write-host ("   Configuring Configuring Super Reader on " + $WA.DisplayName + " using " + $CacheRead + ".")
                out-file -filepath $Logfile -append -inputobject ("   Configuring Configuring Super Reader on " + $WA.DisplayName + ".")
                $WA.Properties["portalsuperreaderaccount"] = $CacheRead
                $WA.Update()
            }
        }
    }
}
Else
{
    write-host "*** Not creating Web Applications."
    out-file -filepath $Logfile -append -inputobject "*** Not creating Web Applications."
}

# Creating Site Collections
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If($SSSites)
{
    write-host "Creating Site Collections."
    out-file -filepath $Logfile -append -inputobject "Creating Site Collections."
    foreach ($Site in $ConfigXML.Configuration.SiteStructure.SiteCollections.Site)
    {
        $Name = $Site.Name
        $URL = $Site.URL
        $Owner1 = $Site.Owner1
        $owner2 = $Site.Owner2
        $Template = $Site.Template
        $Language = $Site.Language
        $DB = $Site.DB
        $SC = Get-SPSite $URL -erroraction SilentlyContinue
        If ($SC -ne $null)
        {
            write-host ("   A site collection already exists at " + $URL + ".")
            out-file -filepath $Logfile -append -inputobject ("   A site collection already exists at " + $URL + ".")
        }
        Else
        {
            write-host ("   Creating Site Collection: " + $Name)
            out-file -filepath $Logfile -append -inputobject ("   Creating Site Collection: " + $Name)
            out-file -filepath $Logfile -append -inputobject (New-SPSite $URL -OwnerAlias $Owner1 -SecondaryOwnerAlias $Owner2 -ContentDatabase $DB -Language $Language -Template $Template -Name $Name -ErrorVariable +ScriptErrors)
        }
    }
}
Else
{
    write-host "*** Not creating Site Collections."
    out-file -filepath $Logfile -append -inputobject "*** Not creating Site Collections."
}

# Creating Sub Sites
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If($SSWebs)
{
    write-host "Creating Sub Sites."
    out-file -filepath $Logfile -append -inputobject "Creating Sub Sites."
    foreach ($Web in $ConfigXML.Configuration.SiteStructure.SubSites.Web)
    {
        $Name = $Web.Name
        $URL = $Web.URL
        $Desc = $Web.Description
        If ($Web.UniquePermissions -eq "Yes")
        {
            [boolean]$UniquePerm = $True
        }
        Else
        {
            [boolean]$UniquePerm = $False
        }
        If ($Web.UseParentTopNav -eq "Yes")
        {
            [boolean]$ParentNav = $True
        }
        Else
        {
            [boolean]$ParentNav = $False
        }
        $Temp = $Web.Template
        If ($Web.PublicSite -eq "Yes")
        {
            [boolean]$Public = $True
        }
        Else
        {
            [boolean]$Public = $False
        }
        $GrpPreFix = $Web.SecurityGroupPrefix
        $SubSite = Get-SPWeb $URL -erroraction SilentlyContinue
        If ($SubSite -ne $null)
        {
            write-host ("   A sub site already exists at " + $URL + ".")
            out-file -filepath $Logfile -append -inputobject ("   A sub site already exists at " + $URL + ".")
        }
        Else
        {
            write-host ("   Creating sub site " + $Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Creating sub site " + $Name + ".")
            $SS = New-SPWeb $URL -Confirm:$False -Description $Desc -Name $Name -Template $Temp -UniquePermissions:$UniquePerm -UseParentTopNav:$ParentNav
    
            # Creating Security Groups
            If($UniquePerm)
            {
                write-host "      Creating Owners Group..."
                out-file -filepath $Logfile -append -inputobject "         Creating Owners Group..."
                $Owners = $GrpPreFix + " Owners"
                $ODesc = ("This is a SharePoint security group. You can utilize this to provide users with Full Control right to the " + $Name + ".")
                $SS.SiteGroups.Add($Owners, $SS.Site.Owner, $null, $ODesc)
                $OGroup = $SS.SiteGroups[$Owners]
                $OGroup.Owner = $OGroup
                $OGroup.Update()
                write-host "      Creating Members Group..."
                out-file -filepath $Logfile -append -inputobject "      Creating Members Group..."
                $Members = $GrpPreFix + " Members"
                $MDesc = ("This is a SharePoint security group. You can utilize this to provide users with Contribute rights to the " + $Name + ".")
                $SS.SiteGroups.Add($Members, $OGroup, $null, $MDesc)
                $MGroup = $SS.SiteGroups[$Members]
                If ($Public)
                {
                    write-host "      Re-using Site Collection Visitors Group..."
                    out-file -filepath $Logfile -append -inputobject "      Re-using Site Collection Visitors Group..."
                    $VGroup = $SS.Site.RootWeb.SiteGroups | Where{$_.Name -like "*Visitors" -and $_.Roles -like "*Read*"}
                }
                Else
                {
                    write-host "      Creating Visitors Group..."
                    out-file -filepath $Logfile -append -inputobject "      Creating Visitors Group..."
                    $Visitors = $GrpPreFix + " Visitors"
                    $VDesc = ("This is a SharePoint security group. You can utilize this to provide users with Read rights to the " + $Name + ".")
                    $SS.SiteGroups.Add($Visitors, $OGroup, $null, $VDesc)
                    $VGroup = $SS.SiteGroups[$Visitors]   
                }
    
                # Assigning Permissions to Groups
                write-host "      Assigning Permission to SharePoint Groups..."
                out-file -filepath $Logfile -append -inputobject "         Assigning Permission to SharePoint Groups..."
                $OAssign = New-Object Microsoft.SharePoint.SPRoleAssignment($OGroup)
                $Olevel = $SS.Site.RootWeb.RoleDefinitions["Full Control"]
                $OAssign.RoleDefinitionBindings.Add($OLevel)
                $SS.RoleAssignments.Add($OAssign)
                $MAssign = New-Object Microsoft.SharePoint.SPRoleAssignment($MGroup)
                $Mlevel = $SS.Site.RootWeb.RoleDefinitions["Contribute"]
                $MAssign.RoleDefinitionBindings.Add($MLevel)
                $SS.RoleAssignments.Add($MAssign)
                $VAssign = New-Object Microsoft.SharePoint.SPRoleAssignment($VGroup)
                $Vlevel = $SS.Site.RootWeb.RoleDefinitions["Read"]
                $VAssign.RoleDefinitionBindings.Add($VLevel)
                $SS.RoleAssignments.Add($VAssign)
                $SS.Update()
            }
        }
    }
}
Else
{
    write-host "*** Not creating Sub Sites."
    out-file -filepath $Logfile -append -inputobject "*** Not creating Sub Sites."
}

# Services
# Microsoft Foundation Web Application
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
write-host "Configuring Microsoft SharePoint Foundation Web Application service"
out-file -filepath $Logfile -append -inputobject "Configuring Microsoft SharePoint Foundation Web Application service"
$Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
ForEach($Server in $Servers)
{
    $Status = "Disable"
    ForEach($EnabledServer in $ConfigXML.Configuration.Services.FoundationWebApplication.Server)
    {
        If($Server.Name -eq $EnabledServer)
        {
            $Status = "Enable"
        }
    }
    $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "Microsoft SharePoint Foundation Web Application" -and $_.Server -eq $Server}
    If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
    {
        write-host ("   Disabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
        $Reset = $True
    }
    ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
    {
        write-host ("   Service already disabled on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
    }
    ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
    {
        write-host ("   Service already enabled on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
    }
    Else
    {
        write-host ("   Enabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
        $Reset = $True
    }
}

# Claims to Windows Token Service
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
write-host "Configuring Claims to Windows Token Service"
out-file -filepath $Logfile -append -inputobject "Configuring Claims to Windows Token Service"
$Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
ForEach($Server in $Servers)
{
    $Status = "Disable"
    ForEach($EnabledServer in $ConfigXML.Configuration.Services.WindowsToken.Server)
    {
        If($Server.Name -eq $EnabledServer)
        {
            $Status = "Enable"
        }
    }
    $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "Claims to Windows Token Service" -and $_.Server -eq $Server}
    If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
    {
        write-host ("   Disabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
    }
    ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
    {
        write-host ("   Service already disabled on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
    }
    ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
    {
        write-host ("   Service already enabled on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
    }
    Else
    {
        write-host ("   Enabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
    }
}

# Microsoft SharePoint Foundation Sandboxed Code Service
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
write-host "Configuring Microsoft SharePoint Foundation Sandboxed Code Service"
out-file -filepath $Logfile -append -inputobject "Configuring Microsoft SharePoint Foundation Sandboxed Code Service"
$Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
ForEach($Server in $Servers)
{
    $Status = "Disable"
    ForEach($EnabledServer in $ConfigXML.Configuration.Services.SandboxedCode.Server)
    {
        If($Server.Name -eq $EnabledServer)
        {
            $Status = "Enable"
        }
    }
    $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "Microsoft SharePoint Foundation Sandboxed Code Service" -and $_.Server -eq $Server}
    If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
    {
        write-host ("   Disabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
    }
    ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
    {
        write-host ("   Service already disabled on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
    }
    ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
    {
        write-host ("   Service already enabled on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
    }
    Else
    {
        write-host ("   Enabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
    }
}

# Request Management
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If ($RMEnable)
{
    
    write-host "Configuring Request Management service."
    out-file -filepath $Logfile -append -inputobject "Configuring Request Management service."
    $Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
    ForEach($Server in $Servers)
    {
        $Status = "Disable"
        $Foundation = Get-SPServiceInstance | Where{$_.TypeName -eq "Microsoft SharePoint Foundation Web Application" -and $_.Server -eq $Server}
        ForEach($EnabledServer in $ConfigXML.Configuration.Services.RequestManagement.Servers.Server)
        {
            If($Server.Name -eq $EnabledServer)
            {
                $Status = "Enable"
            }
        }
        $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "Request Management" -and $_.Server -eq $Server}
        If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
        {
            If ($Foundation.Status -eq "Online")
            {
                write-host ("   Disabling service on " + $Server.Name + ".")
                out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
                out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
            }
            Else
            {
                write-host ("   *** Microsoft SharePoint Foundation Web Application is not enabled on " + $Server.Name + ".")
                out-file -filepath $Logfile -append -inputobject ("   *** Microsoft SharePoint Foundation Web Application is not enabled on " + $Server.Name + ".")
            }
        }
        ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
        {
            write-host ("   Service already disabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
        }
        ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Service already enabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
        }
        Else
        {
            If ($Foundation.Status -eq "Online")
            {
                write-host ("   Enabling service on " + $Server.Name + ".")
                out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
                out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
            }
            Else
            {
                write-host ("   *** Microsoft SharePoint Foundation Web Application is not enabled on " + $Server.Name + ".")
                out-file -filepath $Logfile -append -inputobject ("   *** Microsoft SharePoint Foundation Web Application is not enabled on " + $Server.Name + ".")
            }
        }
    }
    $WebApps = Get-SPWebApplication
    ForEach($WebApp in $WebApps)
    {
        $Config = Get-SPRequestManagementSettings $WebApp
        write-host ("   Configuring Request Management settings on " + $WebApp.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Configuring Request Management settings on " + $WebApp.Name + ".")
        If($RMRouting -eq $Config.RoutingEnabled)
        {
            write-host ("      Routing Enabled is already set to " + $RMRouting + ".")
            out-file -filepath $Logfile -append -inputobject ("      Routing Enabled is already set to " + $RMRouting + ".")
        }
        Else
        {
            write-host ("      Setting Routing Enabled to " + $RMRouting + ".")
            out-file -filepath $Logfile -append -inputobject ("      Setting Routing Enabled to " + $RMRouting + ".")
            Set-SPRequestManagementSettings $WebApp -RoutingEnabled $RMRouting
        }
        If($RMThrottling -eq $Config.ThrottlingEnabled)
        {
            write-host ("      Throttling Enabled is already set to " + $RMThrottling + ".")
            out-file -filepath $Logfile -append -inputobject ("      Throttling Enabled is already set to " + $RMThrottling + ".")
        }
        Else
        {
            write-host ("      Setting Throttling Enabled to " + $RMThrottling + ".")
            out-file -filepath $Logfile -append -inputobject ("      Setting Throttling to " + $RMThrottling + ".")
            Set-SPRequestManagementSettings $WebApp -ThrottlingEnabled $RMThrottling
        }
        If(($RMScheme + "MachineWeight") -eq $Config.RoutingScheme -or($RMScheme + "Based") -eq $Config.RoutingScheme -or $RMScheme -eq $Config.RoutingScheme)
        {
            write-host ("      Routing Scheme is already set to " + $RMScheme + ".")
            out-file -filepath $Logfile -append -inputobject ("      Routing Scheme is already set to " + $RMScheme + ".")
        }
        Else
        {
            write-host ("      Setting Routing Scheme to " + $RMScheme + ".")
            out-file -filepath $Logfile -append -inputobject ("      Setting Routing Scheme to " + $RMScheme + ".")
            Set-SPRequestManagementSettings $WebApp -RoutingScheme $RMScheme
        }
    }   
}
Else
{
    write-host "*** Request Management service will not be configured"
    out-file -filepath $Logfile -append -inputobject "*** Request Management service will not be configured"
}

# Service Applications
# Creating a State Service Application
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
$SA = Get-SPStateServiceApplication $StateSA -erroraction SilentlyContinue
If ($SA -ne $null)
{
    write-host ($StateSA + " already exists.")
    out-file -filepath $Logfile -append -inputobject ($StateSA + " already exists.")
}
Else
{
    write-host "Creating State Service Application."
    out-file -filepath $Logfile -append -inputobject "Creating State Service Application."
    write-host ("   Creating database: " + $StateDB)
    out-file -filepath $Logfile -append -inputobject ("   Creating database: " + $StateDB)
    $SA_DB = New-SPStateServiceDatabase -Name $StateDB -ErrorVariable +ScriptErrors 
    write-host ("   Creating Service Application: " + $StateSA)
    out-file -filepath $Logfile -append -inputobject ("   Creating Service Application: " + $StateSA)
    $SA = New-SPStateServiceApplication -Name $StateSA -Database $SA_DB -ErrorVariable +ScriptErrors
    write-host ("   Creating Service Application Proxy: " + $StateSA + " Proxy")
    out-file -filepath $Logfile -append -inputobject ("   Creating Service Application Proxy: " + $StateSA + " Proxy")
    out-file -filepath $Logfile -append -inputobject(New-SPStateServiceApplicationProxy -Name "$StateSA Proxy" -ServiceApplication $SA -DefaultProxyGroup -ErrorVariable +ScriptErrors)
}

# Creating Usage and Health Service Application
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
$SA = Get-SPUsageApplication $UsageSA -erroraction SilentlyContinue
If ($SA -ne $null)
{
    write-host ($UsageSA + " already exists.")
    out-file -filepath $Logfile -append -inputobject ($UsageSA + " already exists.")
}
Else
{
    write-host "Creating Usage and Health Service Application."
    out-file -filepath $Logfile -append -inputobject "Creating Usage and Health Service Application."
    out-file -filepath $Logfile -append -inputobject (New-SPUsageApplication -Name $UsageSA -DatabaseName $UsageDB -ErrorVariable +ScriptErrors)
}
$Usage = Get-SPUsageService
If (($Usage.LoggingEnabled -eq $UsageEnable) -and ($Usage.UsageLogDir -eq $UsageLoc))
{
    write-host "   Data Collection has already been configured."
    out-file -filepath $Logfile -append -inputobject "   Data Collection has already been configured." 
}
Else
{
    write-host "   Configuring Data Collection."
    out-file -filepath $Logfile -append -inputobject "   Configuring Data Collection."
    Set-SPUsageService -LoggingEnabled:$UsageEnable -UsageLogLocation $UsageLoc
}

# Subscription Settings
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
write-host "Configuring Subscription Settings Service Application"
out-file -filepath $Logfile -append -inputobject "Configuring Subscription Settings Service Application"
$Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
ForEach($Server in $Servers)
{
    $Status = "Disable"
    ForEach($EnabledServer in $ConfigXML.Configuration.ServiceApps.SubscriptionSettings.Servers.Server)
    {
        If($Server.Name -eq $EnabledServer)
        {
            $Status = "Enable"
        }
    }
    $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "Microsoft SharePoint Foundation Subscription Settings Service" -and $_.Server -eq $Server}
    If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
    {
        write-host ("   Disabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
    }
    ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
    {
        write-host ("   Service already disabled on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
    }
    ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
    {
        write-host ("   Service already enabled on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
    }
    Else
    {
        write-host ("   Enabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
        out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
    }
}
$SS = Get-SPServiceApplication -Name $SubscriptionName -erroraction SilentlyContinue
If ($SS -ne $null)
{
    write-host "   Service Application already exists."
    out-file -filepath $Logfile -append -inputobject "   Service Application already exists."
}
Else
{
    $AppPool = get-SPServiceApplicationPool $SubscriptionPool -erroraction SilentlyContinue
    If ($AppPool -ne $null)
    {
        write-host "   Application pool Already Exists."
        out-file -filepath $Logfile -append -inputobject "   Application pool Already Exists."
        $Pool = $AppPool
    }
    Else
    {
        write-host "   Creating Application pool."
        out-file -filepath $Logfile -append -inputobject "   Creating Application pool."
        $Pool = New-SPServiceApplicationPool $SubscriptionPool -account $SubscriptionAcct -ErrorVariable +ScriptErrors
    }
    write-host "   Creating service application."
    out-file -filepath $Logfile -append -inputobject "   Creating service application."
    $SA = New-SPSubscriptionSettingsServiceApplication -Name $SubscriptionName -ApplicationPool $Pool -DatabaseServer $SQLAlias -DatabaseName $SubscriptionDB
    write-host "   Creating service application proxy."
    out-file -filepath $Logfile -append -inputobject "   Creating service application proxy."
    out-file -filepath $Logfile -append -inputobject (New-SPSubscriptionSettingsServiceApplicationProxy -ServiceApplication $SA -ErrorVariable +ScriptErrors)
}

# Secure Store
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If ($SecureConfig)
{
    write-host "Configuring Secure Store Service Application"
    out-file -filepath $Logfile -append -inputobject "Configuring Secure Store Service Application"
    $Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
    ForEach($Server in $Servers)
    {
        $Status = "Disable"
        ForEach($EnabledServer in $ConfigXML.Configuration.ServiceApps.SecureStore.Servers.Server)
        {
            If($Server.Name -eq $EnabledServer)
            {
                $Status = "Enable"
            }
        }
        $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "Secure Store Service" -and $_.Server -eq $Server}
        If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
        }
        ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
        {
            write-host ("   Service already disabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
        }
        ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Service already enabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
        }
        Else
        {
            write-host ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
        }
    }
    $SS = Get-SPServiceApplication -Name $SecureName -erroraction SilentlyContinue
    If ($SS -ne $null)
    {
        write-host "   Service Application already exists."
        out-file -filepath $Logfile -append -inputobject "   Service Application already exists."
    }
    Else
    {
        $AppPool = get-SPServiceApplicationPool $SecurePool -erroraction SilentlyContinue
        If ($AppPool -ne $null)
        {
            write-host "   Application pool Already Exists."
            out-file -filepath $Logfile -append -inputobject "   Application pool Already Exists."
            $Pool = $AppPool
        }
        Else
        {
            write-host "   Creating Application pool."
            out-file -filepath $Logfile -append -inputobject "   Creating Application pool."
            $Pool = New-SPServiceApplicationPool $SecurePool -account $SecureAcct -ErrorVariable +ScriptErrors
        }
        write-host "   Creating service application."
        out-file -filepath $Logfile -append -inputobject "   Creating service application."
        $SA = New-SPSecureStoreServiceApplication -Name $SecureName -ApplicationPool $Pool -DatabaseServer $SQLAlias -DatabaseName $SecureDB -AuditingEnabled:$SecureAudit -AuditlogMaxSize $SecureDays
        write-host "   Creating service application proxy."
        out-file -filepath $Logfile -append -inputobject "   Creating service application proxy."
        out-file -filepath $Logfile -append -inputobject (New-SPSecureStoreServiceApplicationProxy -Name "$SecureName Proxy" -ServiceApplication $SA -ErrorVariable +ScriptErrors)
    }
}
Else
{
    write-host "*** Secure Store Service Application will not be configured."
    out-file -filepath $Logfile -append -inputobject "*** Secure Store Service Application will not be configured."
}

# Creating Managed Metadata Service Application
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If ($ConfigMetadata)
{
    write-host "Configuring Managed Metadata Service Application"
    out-file -filepath $Logfile -append -inputobject "Configuring Managed Metadata Service Application"
    $Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
    ForEach($Server in $Servers)
    {
        $Status = "Disable"
        ForEach($EnabledServer in $ConfigXML.Configuration.ServiceApps.MetadataSrv.Servers.Server)
        {
            If($Server.Name -eq $EnabledServer)
            {
                $Status = "Enable"
            }
        }
        $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "Managed Metadata Web Service" -and $_.Server -eq $Server}
        If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
        }
        ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
        {
            write-host ("   Service already disabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
        }
        ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Service already enabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
        }
        Else
        {
            write-host ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
        }
    }
    $MM = Get-SPMetadataServiceApplication $MetadataSA -erroraction SilentlyContinue
    If ($MM -ne $null)
    {
        write-host "   Service application already exists."
        out-file -filepath $Logfile -append -inputobject "   Service application already exists."
    }
    Else
    {
        $AppPool = Get-SPServiceApplicationPool $MetadataPool -erroraction SilentlyContinue
        If ($AppPool -ne $null)
        {
            write-host "   Application pool Already Exists."
            out-file -filepath $Logfile -append -inputobject "   Application pool Already Exists."
            $Pool = $AppPool
        }
        Else
        {
            write-host "   Creating Application pool."
            out-file -filepath $Logfile -append -inputobject "   Creating Application pool."
            $Pool = New-SPServiceApplicationPool $MetadataPool -account $MetadataAcct -ErrorVariable +ScriptErrors
        }
        write-host "   Creating service application."
        out-file -filepath $Logfile -append -inputobject "   Creating service application."
        $SA = New-SPMetadataServiceApplication -Name $MetadataSA -ApplicationPool $Pool -DatabaseName $MetadataDB -hubURI $CTHub -ErrorVariable +ScriptErrors
        write-host "   Creating service application proxy."
        out-file -filepath $Logfile -append -inputobject "   Creating service application proxy."
        out-file -filepath $Logfile -append -inputobject (New-SPMetadataServiceApplicationProxy -Name "$MetadataSA Proxy" -ServiceApplication $SA -DefaultProxyGroup -ErrorVariable +ScriptErrors)
    }
}
Else
{
    write-host "*** Managed Metadata Service Application will not be configured."
    out-file -filepath $Logfile -append -inputobject "*** Managed Metadata Service Application will not be configured."
}

# Creating Excel Calculations Service Application
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If ($ExcelEnable)
{
    write-host "Configuring Excel Calculations Service Application"
    out-file -filepath $Logfile -append -inputobject "Configuring Excel Calculations Service Application"
    $Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
    ForEach($Server in $Servers)
    {
        $Status = "Disable"
        ForEach($EnabledServer in $ConfigXML.Configuration.ServiceApps.ExcelCalc.Servers.Server)
        {
            If($Server.Name -eq $EnabledServer)
            {
                $Status = "Enable"
            }
        }
        $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "Excel Calculation Services" -and $_.Server -eq $Server}
        If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
        }
        ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
        {
            write-host ("   Service already disabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
        }
        ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Service already enabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
        }
        Else
        {
            write-host ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
        }
    }
    $Ex = Get-SPExcelServiceApplication $ExcelName -erroraction SilentlyContinue
    If ($Ex -ne $null)
    {
        write-host "   Service application already exists."
        out-file -filepath $Logfile -append -inputobject "   Service application already exists."
    }
    Else
    {
        $AppPool = Get-SPServiceApplicationPool $ExcelPool -erroraction SilentlyContinue
        If ($AppPool -ne $null)
        {
            write-host "   Application pool Already Exists."
            out-file -filepath $Logfile -append -inputobject "   Application pool Already Exists."
            $Pool = $AppPool
        }
        Else
        {
            write-host "   Creating Application pool."
            out-file -filepath $Logfile -append -inputobject "   Creating Application pool."
            $Pool = New-SPServiceApplicationPool $ExcelPool -account $ExcelAcct -ErrorVariable +ScriptErrors
        }
        write-host "   Creating service application."
        out-file -filepath $Logfile -append -inputobject "   Creating service application."
        out-file -filepath $Logfile -append -inputobject (New-SPExcelServiceApplication -Name $ExcelName -ApplicationPool $Pool -ErrorVariable +ScriptErrors)
    }
}
Else
{
    write-host "*** Excel Calculations Service Application will not be configured."
    out-file -filepath $Logfile -append -inputobject "*** Excel Calculations Service Application will not be configured."
}

# Configuring Search
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If ($ConfigSearch)
{
    write-host "Configuring Search Service Application"
    out-file -filepath $Logfile -append -inputobject "Configuring Search Service Application"
    $Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
    ForEach($Server in $Servers)
    {
        $Status = "Disable"
        ForEach($EnabledServer in $ConfigXML.Configuration.ServiceApps.SearchSrv.Servers.Server)
        {
            If($Server.Name -eq $EnabledServer)
            {
                $Status = "Enable"
            }
        }
        $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "SharePoint Server Search" -and $_.Server -eq $Server}
        If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
        }
        ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
        {
            write-host ("   Service already disabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
        }
        ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Service already enabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
        }
        Else
        {
            write-host ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
        }
    }
    $Search = Get-SPEnterpriseSearchServiceApplication $SearchSA -erroraction SilentlyContinue
    If ($Search -ne $null)
    {
        write-host "   Service application already exists."
        out-file -filepath $Logfile -append -inputobject "   Service application already exists."
    }
    Else
    {
        $AppPool = Get-SPServiceApplicationPool $SearchPool -erroraction SilentlyContinue
        If ($AppPool -ne $null)
        {
            write-host "   Application pool Already Exists."
            out-file -filepath $Logfile -append -inputobject "   Application pool Already Exists."
            $Pool = $AppPool
        }
        Else
        {
            write-host "   Creating Application pool."
            out-file -filepath $Logfile -append -inputobject "   Creating Application pool."
            $Pool = New-SPServiceApplicationPool $SearchPool -Account $SearchAcct -ErrorVariable +ScriptErrors
        }
        write-host "   Creating Search Service Application."
        out-file -filepath $Logfile -append -inputobject "   Creating Search Service Application."
        $SA = New-SPEnterpriseSearchServiceApplication $SearchSA -ApplicationPool $Pool -DatabaseName $SearchDB -ErrorVariable +ScriptErrors
        write-host "   Creating Search Service Application Proxy."
        out-file -filepath $Logfile -append -inputobject "   Creating Search Service Application Proxy."
        $Proxy = New-SPEnterpriseSearchServiceApplicationProxy "$SearchSA Proxy" -SearchApplication $SA -ErrorVariable +ScriptErrors
        <# write-host "   Creating Search Topology"
            out-file -filepath $Logfile -append -inputobject "   Creating Search Topology"
        $Topology = New-SPEnterpriseSearchTopology -SearchApplication $SA
        ForEach($Server in $ConfigXML.Configuration.ServiceApps.SearchSrv.Admin.Server)
        {
            write-host ("      Creating Search Admin Component on " + $Server + ".")
            out-file -filepath $Logfile -append -inputobject ("      Creating Search Admin Component on " + $Server + ".")
            $SrvInst = Get-SPEnterpriseSearchServiceInstance -identity $Server -ErrorVariable +ScriptErrors
            out-file -filepath $Logfile -append -inputobject (New-SPEnterpriseSearchAdminComponent –SearchTopology $Topology -SearchServiceInstance $SrvInst -ErrorVariable +ScriptErrors)
        }
        ForEach($Server in $ConfigXML.Configuration.ServiceApps.SearchSrv.Content.Server)
        {
            write-host ("      Creating Search Content Processing Component on " + $Server + ".")
            out-file -filepath $Logfile -append -inputobject ("      Creating Search Content Processing Component on " + $Server + ".")
            $SrvInst = Get-SPEnterpriseSearchServiceInstance -identity $Server -ErrorVariable +ScriptErrors
            out-file -filepath $Logfile -append -inputobject (New-SPEnterpriseSearchContentProcessingComponent –SearchTopology $Topology -SearchServiceInstance $SrvInst -ErrorVariable +ScriptErrors)
        }
        ForEach($Server in $ConfigXML.Configuration.ServiceApps.SearchSrv.Analytics.Server)
        {
            write-host ("      Creating Search Analytics Component on " + $Server + ".")
            out-file -filepath $Logfile -append -inputobject ("      Creating Search Analytics Component on " + $Server + ".")
            $SrvInst = Get-SPEnterpriseSearchServiceInstance -identity $Server -ErrorVariable +ScriptErrors
            out-file -filepath $Logfile -append -inputobject (New-SPEnterpriseSearchAnalyticsProcessingComponent –SearchTopology $Topology -SearchServiceInstance $SrvInst -ErrorVariable +ScriptErrors)
        }
        ForEach($Server in $ConfigXML.Configuration.ServiceApps.SearchSrv.Crawl.Server)
        {
            write-host ("      Creating Search Crawl Component on " + $Server + ".")
            out-file -filepath $Logfile -append -inputobject ("      Creating Search Crawl Component on " + $Server + ".")
            $SrvInst = Get-SPEnterpriseSearchServiceInstance -identity $Server -ErrorVariable +ScriptErrors
            out-file -filepath $Logfile -append -inputobject (New-SPEnterpriseSearchCrawlComponent –SearchTopology $Topology -SearchServiceInstance $SrvInst -ErrorVariable +ScriptErrors) 
        }
        ForEach($Server in $ConfigXML.Configuration.ServiceApps.SearchSrv.Index.Server)
        {
            write-host ("      Creating Search Index Component on " + $Server + ".")
            out-file -filepath $Logfile -append -inputobject ("      Creating Search Index Component on " + $Server + ".")
            $SrvInst = Get-SPEnterpriseSearchServiceInstance -identity $Server -ErrorVariable +ScriptErrors
            out-file -filepath $Logfile -append -inputobject (New-SPEnterpriseSearchIndexComponent –SearchTopology $Topology -SearchServiceInstance $SrvInst -IndexPartition 0 -RootDirectory $SearchIndex -ErrorVariable +ScriptErrors)
        }
        ForEach($Server in $ConfigXML.Configuration.ServiceApps.SearchSrv.Query.Server)
        {
            write-host ("      Creating Search Query Component on " + $Server + ".")
            out-file -filepath $Logfile -append -inputobject ("      Creating Search Query Component on " + $Server + ".")
            $SrvInst = Get-SPEnterpriseSearchServiceInstance -identity $Server -ErrorVariable +ScriptErrors
            out-file -filepath $Logfile -append -inputobject (New-SPEnterpriseSearchQueryProcessingComponent –SearchTopology $Topology -SearchServiceInstance $SrvInst -ErrorVariable +ScriptErrors)
        } 
        
        write-host "   Activating Search Topology"
        out-file -filepath $Logfile -append -inputobject "   Activating Search Topology."
        Set-SPEnterpriseSearchTopology $Topology #>
    }
}
Else
{
    write-host "*** Search Service Application will not be configured."
    out-file -filepath $Logfile -append -inputobject "*** Search Service Application will not be configured."
}

# Creating User Profiles Service Application
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If ($UPSConfig)
{
    write-host "Configuring User Profile Service Application"
    out-file -filepath $Logfile -append -inputobject "Configuring User Profile Service Application"
    write-warning "The Profile Syncronization service will need to be started manually on the appropriate servers after the completion of this script.  Temporary permissions are required before that service can be started."
    out-file -filepath $Logfile -append -inputobject "*** The Profile Syncronization service will need to be started manually on the appropriate servers after the completion of this script.  Temporary permissions are required before that service can be started."
    $Path = Get-SPManagedPath -WebApplication $UPSWA | Where{$_.Name -eq $UPSPath} -erroraction SilentlyContinue
    If($Path -ne $null)
    {
        write-host "   Managed Path Already Exists."
        out-file -filepath $Logfile -append -inputobject "   Managed Path Already Exists."
    }
    Else
    {
        write-host "   Creating managed path."
        out-file -filepath $Logfile -append -inputobject "   Creating managed path."
        out-file -filepath $Logfile -append -inputobject (New-SPManagedPath $UPSPath -WebApplication $UPSWA -ErrorVariable +ScriptErrors)
    }
    $UPS = Get-SPServiceApplication | Where{$_.DisplayName -eq $UPSName}
    If ($UPS -ne $null)
    {
        write-host "   Service application already exists."
        out-file -filepath $Logfile -append -inputobject "   Service application already exists."
        $SA = $UPS
    }
    Else
    {
        $AppPool = Get-SPServiceApplicationPool $UPSPool -erroraction SilentlyContinue
        If ($AppPool -ne $null)
        {
            write-host "   Application pool Already Exists."
            out-file -filepath $Logfile -append -inputobject "   Application pool Already Exists."
            $Pool = $AppPool
        }
        Else
        {
            write-host "   Creating Application pool."
            out-file -filepath $Logfile -append -inputobject "   Creating Application pool."
            $Pool = New-SPServiceApplicationPool $UPSPool -account $UPSAcct -ErrorVariable +ScriptErrors
        }
        write-host "   Creating service application."
        out-file -filepath $Logfile -append -inputobject "   Creating service application."
        $SA = New-SPProfileServiceApplication -Name $UPSName -ApplicationPool $Pool -MySiteHostLocation $UPSWA -MySiteManagedPath $UPSPath -ProfileDBName ($UPSDB + "_Profile") -ProfileSyncDBName ($UPSDB + "_Sync") -SocialDBName ($UPSDB + "_Social") -ErrorVariable +ScriptErrors
        out-file -filepath $Logfile -append -inputobject (New-SPProfileServiceApplicationProxy -ServiceApplication $SA -Name "$UPSName Proxy")
    }
    If($UPSNet)
    {
        If($SA.NetBIOSDomainNamesEnabled -eq 1)
        {
            write-host "   NetBIOS already Enabled."
            out-file -filepath $Logfile -append -inputobject "   NetBIOS already Enabled."
        }
        Else
        {
            write-host "   Enabling NetBIOS."
            out-file -filepath $Logfile -append -inputobject "   Enabling NetBios."
            $SA.NetBIOSDomainNamesEnabled = 1
            $SA.Update()
        }
    }
    $Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
    ForEach($Server in $Servers)
    {
        $Status = "Disable"
        ForEach($EnabledServer in $ConfigXML.Configuration.ServiceApps.UserProfiles.Servers.Server)
        {
            If($Server.Name -eq $EnabledServer)
            {
                $Status = "Enable"
            }
        }
        $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "User Profile Service" -and $_.Server -eq $Server}
        If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
        }
        ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
        {
            write-host ("   Service already disabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
        }
        ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Service already enabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
        }
        Else
        {
            write-host ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
        }
    }
}
Else
{
    write-host "*** User Profile Service Application will not be configured."
    out-file -filepath $Logfile -append -inputobject "*** User Profile Service Application will not be configured."
}

# Creating App Management Service Application
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
If ($AMConfig)
{
    write-host "Configuring App Management Service Application"
    out-file -filepath $Logfile -append -inputobject "Configuring App Management Service Application"
    write-warning "The App Catalog will need to be created manually."
    out-file -filepath $Logfile -append -inputobject "*** The App Catalog will need to be created manually."
    $Servers = $Farm.Servers | Where{$_.Role -ne "Invalid"}
    ForEach($Server in $Servers)
    {
        $Status = "Disable"
        ForEach($EnabledServer in $ConfigXML.Configuration.ServiceApps.AppManagement.Servers.Server)
        {
            If($Server.Name -eq $EnabledServer)
            {
                $Status = "Enable"
            }
        }
        $SrvInst = Get-SPServiceInstance | Where{$_.TypeName -eq "App Management Service" -and $_.Server -eq $Server}
        If ($Status -eq "Disable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Disabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Stop-SPServiceInstance $SrvInst -Confirm:$False)
        }
        ElseIf($Status -eq "Disable" -and $SrvInst.Status -eq "Disabled")
        {
            write-host ("   Service already disabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already disabled on " + $Server.Name + ".")
        }
        ElseIf($Status -eq "Enable" -and $SrvInst.Status -eq "Online")
        {
            write-host ("   Service already enabled on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Service already enabled on " + $Server.Name + ".")
        }
        Else
        {
            write-host ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("   Enabling service on " + $Server.Name + ".")
            out-file -filepath $Logfile -append -inputobject (Start-SPServiceInstance $SrvInst)
        }
    }
    $AM = Get-SPServiceApplication | Where{$_.DisplayName -eq $AMName}
    If ($AM -ne $null)
    {
        write-host "   Service application already exists."
        out-file -filepath $Logfile -append -inputobject "   Service application already exists."
    }
    Else
    {
        $AppPool = Get-SPServiceApplicationPool $AMPool -erroraction SilentlyContinue
        If ($AppPool -ne $null)
        {
            write-host "   Application pool Already Exists."
            out-file -filepath $Logfile -append -inputobject "   Application pool Already Exists."
            $Pool = $AppPool
        }
        Else
        {
            write-host "   Creating Application pool."
            out-file -filepath $Logfile -append -inputobject "   Creating Application pool."
            $Pool = New-SPServiceApplicationPool $AMPool -account $AMAcct -ErrorVariable +ScriptErrors
        }
        write-host "   Creating service application."
        out-file -filepath $Logfile -append -inputobject "   Creating service application."
        $SA = New-SPAppManagementServiceApplication -Name $AMName -ApplicationPool $Pool -DatabaseName $AMDB -DatabaseServer $SQLAlias -ErrorVariable +ScriptErrors
        out-file -filepath $Logfile -append -inputobject (New-SPAppManagementServiceApplicationProxy -ServiceApplication $SA -Name "$AMName Proxy" -ErrorVariable +ScriptErrors)
        $Reset = $Trueping
    }

    # Configuring app URLs
    $Config = Get-SPAppDomain -erroraction SilentlyContinue
    If($AMDomain -eq $Config)
    {
        write-host "   App domain already configured."
        out-file -filepath $Logfile -append -inputobject "   App domain already configured."
    }
    Else
    {
        write-host ("   Setting App domain to " + $AMDomain + ".")
        out-file -filepath $Logfile -append -inputobject ("   Setting App domain to " + $AMDomain + ".")
        Set-SPAppDomain $AMDomain -ErrorVariable +ScriptErrors

    }
    $Config = Get-SPAppSiteSubscriptionName -erroraction SilentlyContinue
    If($AMPrefix -eq $Config)
    {
        write-host "   App prefix already configured."
        out-file -filepath $Logfile -append -inputobject "   App prefix already configured."
    }
    Else
    {
        write-host ("   Setting App prefix to " + $AMPrefix + ".")
        out-file -filepath $Logfile -append -inputobject ("   Setting App prefix to " + $AMPrefix + ".")
        Set-SPAppSiteSubscriptionName -Name $AMPrefix -Confirm:$False -ErrorVariable +ScriptErrors
    }
    $WebApps = Get-SPWebApplication
    write-host "   Configuring Internet Facing Endpoints on Web Applications."
    out-file -filepath $Logfile -append -inputobject "   Configuring Internet Facing Endpoints on Web Applications."
    
    # Configuring 
    ForEach($WebApp in $WebApps)
    {
        $Feature = Get-SPFeature -WebApplication $WebApp | Where{$_.DisplayName -eq "IfeDependentApps"} -erroraction SilentlyContinue
        If($Feature -ne $null -and $AMEndPoint)
        {
            write-host ("      Internet Endpoints already configured on " + $WebApp.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("      Internet Endpoints already configured on " + $WebApp.Name + ".")
        }
        ElseIf($Feature -eq $null -and $AMEndPoint)
        {
            write-host ("      Enabling Internet Endpoints for " + $WebApp.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("      Enabling Internet Endpoints for " + $WebApp.Name + ".")
            Enable-SPFeature "IfeDependentApps" -URL $WebApp.URL -ErrorVariable +ScriptErrors
        }
        ElseIf($Feature -eq $null -and $AMEndPoint -eq $False)
        {
            write-host ("      Internet Endpoints already configured on " + $WebApp.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("      Internet Endpoints already configured on " + $WebApp.Name + ".")
        }
        ElseIf($Feature -ne $null -and $AMEndPoint -eq $False)
        {
            write-host ("      Disabling Internet Endpoints for " + $WebApp.Name + ".")
            out-file -filepath $Logfile -append -inputobject ("      Enabling Internet Endpoints for " + $WebApp.Name + ".")
            Disable-SPFeature $Feature -URL $WebApp.URL -Confirm:$False -ErrorVariable +ScriptErrors
        }
    } 
}
Else
{
    write-host "*** App Management Service Application will not be configured."
    out-file -filepath $Logfile -append -inputobject "*** App Management Service Application will not be configured."
}

<# Branding Deployment 
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
$TimerJobs = $False
If ($BrandingConfig)
{
    # Adding Solution
    write-host "Deploying Branding"
    out-file -filepath $Logfile -append -inputobject "Deploying Branding"
    $Solution = Get-SPSolution $BrandingWSP -erroraction SilentlyContinue
    If ($Solution -ne $null)
    {
        write-host "   WSP already exists."
        out-file -filepath $Logfile -append -inputobject "   WSP already exists."
    }
    Else
    {
        write-host "   Adding WSP files."
        out-file -filepath $Logfile -append -inputobject "   Adding WSP files."
        out-file -filepath $Logfile -append -inputobject (Add-SPSolution $SolutionLocation$BrandingWSP -ErrorVariable +ScriptErrors)
    }

    # Installing Solution
    $Solution = Get-SPSolution $BrandingWSP -erroraction SilentlyContinue
    If ($Solution -ne $null)
    {
        If ($Solution.Deployed)
        {
            write-host "   Solution already deployed."
            out-file -filepath $Logfile -append -inputobject "   Solution already deployed."
        }
        Else
        {
            write-host "   Deploying solution."
            out-file -filepath $Logfile -append -inputobject "   Deploying solution."
            Install-SPSolution –Identity $BrandingWSP -GACDeployment -ErrorVariable +ScriptErrors
            $TimerJobs = $True
        }
    }

    # Executing Timer Jobs
    If ($TimerJobs)
    {
        write-host "   Executing timer jobs..."
        out-file -filepath $Logfile -append -inputobject "   Executing timer jobs..."
        Stop-Service "SharePoint Administration" -ErrorVariable +ScriptErrors
        Start-SPAdminJob -ErrorVariable +ScriptErrors
        Start-Service "SharePoint Administration" -ErrorVariable +ScriptErrors
        $Prompt = new-object -comobject wscript.shell 
        out-file -filepath $Logfile -append -inputobject ($Prompt.popup("Please wait for the deployment to complete.  You can check the status in Central Administrion / System Settings / Farm Solutions... Click OK when complete.", 0))
    }

    # Acitvating Features & setting Master page
    If ($BrandingAll)
    {
        $Sites = Get-SPSite -Limit All
        Foreach ($Site in $Sites)
        {
            # Activating Features     ***** You Will need to update this for your site templates *****
            If (($Site.RootWeb.WebTemplate -ne "SPSMSITEHOST") -and ($Site.RootWeb.WebTemplate -ne "SPSPERS") -and ($Site.RootWeb.WebTemplate -ne "SPSMSITE"))
            {
                write-host "   ----------------------------------------------------"
                out-file -filepath $Logfile -append -inputobject "   ----------------------------------------------------"
                write-host ("   Applying Branding to " + $Site.RootWeb.Title + ".")
                out-file -filepath $Logfile -append -inputobject ("   Applying Branding to " + $Site.RootWeb.Title + ".")
                write-host "   ----------------------------------------------------"
                out-file -filepath $Logfile -append -inputobject "   ----------------------------------------------------"

                # Branding Variables
                IF ($Site.ServerRelativeURL -eq "/")
                {
                    $BrandingFont = $ConfigXML.Configuration.Solutions.Branding.FontScheme
                    $BrandingColor = $ConfigXML.Configuration.Solutions.Branding.ColorPalette
                    If ($ConfigXML.Configuration.Solutions.Branding.BackgroundImage -eq "")
                    {
                        $BrandingImage = $null
                    }
                    Else
                    {
                        $BrandingImage = $ConfigXML.Configuration.Solutions.Branding.BackgroundImage
                    }
                    If ($Site.WebApplication.URL -eq $ConfigXML.Configuration.Solutions.Branding.ITARWebApp)
                    {
                        $BrandingMaster = $ConfigXML.Configuration.Solutions.Branding.ITARMasterPage
                    }
                    Else
                    {
                        $BrandingMaster = $ConfigXML.Configuration.Solutions.Branding.MasterPage
                    }
                    $BrandingLogo = $ConfigXML.Configuration.Solutions.Branding.ReplacementLogo
                }
                Else
                {
                    $BrandingFont = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.FontScheme)
                    $BrandingColor = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.ColorPalette)
                    If ($ConfigXML.Configuration.Solutions.Branding.BackgroundImage -eq "")
                    {
                        $BrandingImage = $null
                    }
                    Else
                    {
                        $BrandingImage = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.BackgroundImage)
                    }
                    If ($Site.WebApplication.URL -eq $ConfigXML.Configuration.Solutions.Branding.ITARWebApp)
                    {
                        $BrandingMaster = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.ITARMasterPage)
                    }
                    Else
                    {
                        $BrandingMaster = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.MasterPage)
                    }
                    $BrandingLogo = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.ReplacementLogo)
                }
                
                # Activating Site Features
                ForEach ($Feature in $ConfigXML.Configuration.Solutions.Branding.SiteFeatures.Feature)
                {
                    $Feat = Get-SPFeature $Feature -site $Site -erroraction SilentlyContinue
                    If ($Feat -ne $null)
                    {
                        write-host ("      " + $Feature + " is already activated.")
                        out-file -filepath $Logfile -append -inputobject ("      " + $Feature + " is already activated.")
                    }
                    Else
                    {
                        write-host ("      Activating feature " + $Feature + ".")
                        out-file -filepath $Logfile -append -inputobject ("      Activating feature " + $Feature + ".")
                        Enable-SPFeature $Feature -url $Site.URL -ErrorVariable +ScriptErrors
                    }
                }

                # Activating Web Features
                ForEach ($Feature in $ConfigXML.Configuration.Solutions.Branding.WebFeatures.Feature)
                {
                    $Feat = Get-SPFeature $Feature -web $Site.RootWeb -erroraction SilentlyContinue
                    If ($Feat -ne $null)
                    {
                        write-host ("      " + $Feature + " is already activated.")
                        out-file -filepath $Logfile -append -inputobject ("      " + $Feature + " is already activated.")
                    }
                    Else
                    {
                        write-host ("      Activating feature " + $Feature + ".")
                        out-file -filepath $Logfile -append -inputobject ("      Activating feature " + $Feature + ".")
                        Enable-SPFeature $Feature -url $Site.URL -ErrorVariable +ScriptErrors
                    }
                }
                
                # Applying Master Pages
                $Webs = Get-SPWeb -site $Site -limit All
                Foreach ($Web in $Webs)
                {
                    If (($BrandingMaster -ne $Web.MasterURL) -or ($BrandingMaster -ne $Web.CustomMasterURL) -or ($BrandingLogo -ne $Web.SiteLogoURL))
                    {
                        write-host ("      Applying master page and Site Logo on " + $Web.Title + ".")
                        out-file -filepath $Logfile -append -inputobject ("      Applying master page and Site Logo on " + $Web.Title + ".")
                        $Web.MasterURL = $BrandingMaster
                        $Web.CustomMasterURL = $BrandingMaster
                        $Web.SiteLogoURL = $BrandingLogo
                        $Web.Update()
                    }
                    Else
                    {
                        write-host ("      Masterpage and Site Logo are already set on " + $Web.Title + ".")
                        out-file -filepath $Logfile -append -inputobject ("      Masterpage and Site Logo are already set on " + $Web.Title + ".")
                    }
                }

                # Applying Theme
                Foreach ($Web in $Webs)
                {
                    write-host ("      Applying Theme on " + $Web.Title + ".")
                    out-file -filepath $Logfile -append -inputobject ("      Applying Theme on " + $Web.Title + ".")
                    $Web.ApplyTheme($BrandingColor, $BrandingFont, $BrandingImage, $False)
                }
                $Web.Dispose()
            }
            $Site.Dispose()
        }
    }
    Else
    {
        Foreach ($URL in $ConfigXML.Configuration.Solutions.Branding.SelectSites.URL)
        {
            $Site = Get-SPSite $URL
        
            # Activating Features     ***** You Will need to update this for your site templates *****
            If (($Site.RootWeb.WebTemplate -ne "SPSMSITEHOST") -and ($Site.RootWeb.WebTemplate -ne "SPSPERS") -and ($Site.RootWeb.WebTemplate -ne "SPSMSITE"))
            {
                write-host "   ----------------------------------------------------"
                out-file -filepath $Logfile -append -inputobject "   ----------------------------------------------------"
                write-host ("   Applying Branding to " + $Site.RootWeb.Title + ".")
                out-file -filepath $Logfile -append -inputobject ("   Applying Branding to " + $Site.RootWeb.Title + ".")
                write-host "   ----------------------------------------------------"
                out-file -filepath $Logfile -append -inputobject "   ----------------------------------------------------"

                # Branding Variables
                IF ($Site.ServerRelativeURL -eq "/")
                {
                    $BrandingFont = $ConfigXML.Configuration.Solutions.Branding.FontScheme
                    $BrandingColor = $ConfigXML.Configuration.Solutions.Branding.ColorPalette
                    If ($ConfigXML.Configuration.Solutions.Branding.BackgroundImage -eq "")
                    {
                        $BrandingImage = $null
                    }
                    Else
                    {
                        $BrandingImage = $ConfigXML.Configuration.Solutions.Branding.BackgroundImage
                    }
                    If ($Site.WebApplication.URL -eq $ConfigXML.Configuration.Solutions.Branding.ITARWebApp)
                    {
                        $BrandingMaster = $ConfigXML.Configuration.Solutions.Branding.ITARMasterPage
                    }
                    Else
                    {
                        $BrandingMaster = $ConfigXML.Configuration.Solutions.Branding.MasterPage
                    }
                    $BrandingLogo = $ConfigXML.Configuration.Solutions.Branding.ReplacementLogo
                }
                Else
                {
                    $BrandingFont = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.FontScheme)
                    $BrandingColor = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.ColorPalette)
                    If ($ConfigXML.Configuration.Solutions.Branding.BackgroundImage -eq "")
                    {
                        $BrandingImage = $null
                    }
                    Else
                    {
                        $BrandingImage = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.BackgroundImage)
                    }
                    If ($Site.WebApplication.URL -eq $ConfigXML.Configuration.Solutions.Branding.ITARWebApp)
                    {
                        $BrandingMaster = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.ITARMasterPage)
                    }
                    Else
                    {
                        $BrandingMaster = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.MasterPage)
                    }
                    $BrandingLogo = ($Site.ServerRelativeURL + $ConfigXML.Configuration.Solutions.Branding.ReplacementLogo)
                }
                
                # Activating Site Features
                ForEach ($Feature in $ConfigXML.Configuration.Solutions.Branding.SiteFeatures.Feature)
                {
                    $Feat = Get-SPFeature $Feature -site $Site -erroraction SilentlyContinue
                    If ($Feat -ne $null)
                    {
                        write-host ("      " + $Feature + " is already activated.")
                        out-file -filepath $Logfile -append -inputobject ("      " + $Feature + " is already activated.")
                    }
                    Else
                    {
                        write-host ("      Activating feature " + $Feature + ".")
                        out-file -filepath $Logfile -append -inputobject ("      Activating feature " + $Feature + ".")
                        Enable-SPFeature $Feature -url $Site.URL -ErrorVariable +ScriptErrors
                    }
                }

                # Activating Web Features
                ForEach ($Feature in $ConfigXML.Configuration.Solutions.Branding.WebFeatures.Feature)
                {
                    $Feat = Get-SPFeature $Feature -web $Site.RootWeb -erroraction SilentlyContinue
                    If ($Feat -ne $null)
                    {
                        write-host ("      " + $Feature + " is already activated.")
                        out-file -filepath $Logfile -append -inputobject ("      " + $Feature + " is already activated.")
                    }
                    Else
                    {
                        write-host ("      Activating feature " + $Feature + ".")
                        out-file -filepath $Logfile -append -inputobject ("      Activating feature " + $Feature + ".")
                        Enable-SPFeature $Feature -url $Site.URL -ErrorVariable +ScriptErrors
                    }
                }
                
                # Applying Master Pages
                $Webs = Get-SPWeb -site $Site -limit All
                Foreach ($Web in $Webs)
                {
                    If (($BrandingMaster -ne $Web.MasterURL) -or ($BrandingMaster -ne $Web.CustomMasterURL) -or ($BrandingLogo -ne $Web.SiteLogoURL))
                    {
                        write-host ("      Applying master page and Site Logo on " + $Web.Title + ".")
                        out-file -filepath $Logfile -append -inputobject ("      Applying master page and Site Logo on " + $Web.Title + ".")
                        $Web.MasterURL = $BrandingMaster
                        $Web.CustomMasterURL = $BrandingMaster
                        $Web.SiteLogoURL = $BrandingLogo
                        $Web.Update()
                    }
                    Else
                    {
                        write-host ("      Masterpage and Site Logo are already set on " + $Web.Title + ".")
                        out-file -filepath $Logfile -append -inputobject ("      Masterpage and Site Logo are already set on " + $Web.Title + ".")
                    }
                }

                # Applying Theme
                Foreach ($Web in $Webs)
                {
                    write-host ("      Applying Theme on " + $Web.Title + ".")
                    out-file -filepath $Logfile -append -inputobject ("      Applying Theme on " + $Web.Title + ".")
                    $Web.ApplyTheme($BrandingColor, $BrandingFont, $BrandingImage, $False)
                }
                $Web.Dispose()
            }
            $Site.Dispose()
        }
    }
}
Else
{
    write-host "*** No Branding is being Deployed."
    out-file -filepath $Logfile -append -inputobject "*** No Branding is being Deployed."
}#>

# Deploying Farm Solutions
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
ForEach ($Solution in $ConfigXML.Configuration.Solutions.Solution)
{
    #Solution Variables
    $TimerJobs = $False
    $WSP = $Solution.WSP
    If ($Solution.GACDeployment -eq "Yes")
    {
        [boolean]$GAC = $True
    }
    Else
    {
        [boolean]$GAC = $False
    }
    If ($Solution.AllWebApps -eq "Yes")
    {
        [boolean]$AllWA = $True
    }
    Else
    {
        [boolean]$AllWA = $False
    }
    If ($Solution.AllSites -eq "Yes")
    {
        [boolean]$AllSites = $True
    }
    Else
    {
        [boolean]$AllSites = $False
    }

    # Adding Solution
    write-host ("Deploying " + $WSP)
    out-file -filepath $Logfile -append -inputobject ("Deploying " + $WSP)
    $S = Get-SPSolution $WSP -erroraction SilentlyContinue
    If ($S -ne $null)
    {
        write-host "   WSP already exists."
        out-file -filepath $Logfile -append -inputobject "   WSP already exists."
    }
    Else
    {
        write-host "   Adding WSP files."
        out-file -filepath $Logfile -append -inputobject "   Adding WSP files."
        out-file -filepath $Logfile -append -inputobject (Add-SPSolution $SolutionLocation$WSP -ErrorVariable +ScriptErrors)
    }

    # Installing Solution
    $S = Get-SPSolution $WSP -erroraction SilentlyContinue
    If ($S -ne $null)
    {
        If ($S.Deployed)
        {
            write-host "   Solution already deployed."
            out-file -filepath $Logfile -append -inputobject "   Solution already deployed."
        }
        Else
        {
            write-host "   Deploying solution."
            out-file -filepath $Logfile -append -inputobject "   Deploying solution."
            If ($AllWA)
            {
                Install-SPSolution $S -GACDeployment:$GAC -AllWebApplications -ErrorVariable +ScriptErrors
            }
            Else
            {
                ForEach ($WA in $Solution.WebApps.URL)
                {
                    write-host ("      " + $WA)
                    out-file -filepath $Logfile -append -inputobject ("      " + $WA)
                    Install-SPSolution $S -GACDeployment:$GAC -WebApplication $WA -ErrorVariable +ScriptErrors
                }
            }
            $TimerJobs = $True
        }
    }

    # Executing Timer Jobs
    If ($TimerJobs)
    {
        write-host "   Executing timer jobs..."
        out-file -filepath $Logfile -append -inputobject "   Executing timer jobs..."
        Stop-Service "SharePoint Administration" -ErrorVariable +ScriptErrors
        Start-SPAdminJob -ErrorVariable +ScriptErrors
        Start-Service "SharePoint Administration" -ErrorVariable +ScriptErrors
        $Prompt = new-object -comobject wscript.shell 
        out-file -filepath $Logfile -append -inputobject ($Prompt.popup("Please wait for the deployment to complete.  You can check the status in Central Administrion / System Settings / Farm Solutions... Click OK when complete.", 0))
    }

    # Acitvating Features
    If ($AllSites)
    {
        $Sites = Get-SPSite -Limit All
        Foreach ($Site in $Sites)
        {
            write-host ("   Activating Features on " + $Site.RootWeb.Title + ".")
            out-file -filepath $Logfile -append -inputobject ("   Activating Features on " + $Site.RootWeb.Title + ".")
                
            # Activating Site Features
            ForEach ($Feature in $Solution.SiteFeatures.Feature)
            {
                $Feat = Get-SPFeature $Feature -site $Site -erroraction SilentlyContinue
                If ($Feat -ne $null)
                {
                    write-host ("      " + $Feature + " is already activated.")
                    out-file -filepath $Logfile -append -inputobject ("      " + $Feature + " is already activated.")
                }
                Else
                {
                    write-host ("      Activating feature " + $Feature + ".")
                    out-file -filepath $Logfile -append -inputobject ("      Activating feature " + $Feature + ".")
                    Enable-SPFeature $Feature -url $Site.URL -ErrorVariable +ScriptErrors
                }
            }

            # Activating Web Features
            ForEach ($Feature in $Solution.WebFeatures.Feature)
            {
                $Feat = Get-SPFeature $Feature -web $Site.RootWeb -erroraction SilentlyContinue
                If ($Feat -ne $null)
                {
                    write-host ("      " + $Feature + " is already activated.")
                    out-file -filepath $Logfile -append -inputobject ("      " + $Feature + " is already activated.")
                }
                Else
                {
                    write-host ("      Activating feature " + $Feature + ".")
                    out-file -filepath $Logfile -append -inputobject ("      Activating feature " + $Feature + ".")
                    Enable-SPFeature $Feature -url $Site.URL -ErrorVariable +ScriptErrors
                }
            }
            $Site.Dispose()
        }
    }
    Else
    {
        Foreach ($URL in $Solution.SelectSites.URL)
        {
            $Site = Get-SPSite $URL
        
            # Activating Features     ***** You Will need to update this for your site templates *****
            write-host ("   Applying Branding to " + $Site.RootWeb.Title + ".")
            out-file -filepath $Logfile -append -inputobject ("   Applying Branding to " + $Site.RootWeb.Title + ".")

            # Activating Site Features
            ForEach ($Feature in $Solution.SiteFeatures.Feature)
            {
                $Feat = Get-SPFeature $Feature -site $Site -erroraction SilentlyContinue
                If ($Feat -ne $null)
                {
                    write-host ("      " + $Feature + " is already activated.")
                    out-file -filepath $Logfile -append -inputobject ("      " + $Feature + " is already activated.")
                }
                Else
                {
                    write-host ("      Activating feature " + $Feature + ".")
                    out-file -filepath $Logfile -append -inputobject ("      Activating feature " + $Feature + ".")
                    Enable-SPFeature $Feature -url $Site.URL -ErrorVariable +ScriptErrors
                }
            }

            # Activating Web Features
            ForEach ($Feature in $Solution.WebFeatures.Feature)
            {
                $Feat = Get-SPFeature $Feature -web $Site.RootWeb -erroraction SilentlyContinue
                If ($Feat -ne $null)
                {
                    write-host ("      " + $Feature + " is already activated.")
                    out-file -filepath $Logfile -append -inputobject ("      " + $Feature + " is already activated.")
                }
                Else
                {
                    write-host ("      Activating feature " + $Feature + ".")
                    out-file -filepath $Logfile -append -inputobject ("      Activating feature " + $Feature + ".")
                    Enable-SPFeature $Feature -url $Site.URL -ErrorVariable +ScriptErrors
                }
            }
        }
        $Site.Dispose()
    }
}

# Resetting IIS
IF ($Reset)
{
    write-host "____________________________________________________"
    out-file -filepath $Logfile -append -inputobject "____________________________________________________"
    write-host "Resetting IIS."
    out-file -filepath $Logfile -append -inputobject "Resetting IIS."
    out-file -filepath $Logfile -append -inputobject (IISReset)
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
out-file -filepath $Logfile -append -inputobject ($Prompt.popup("Farm Configuration is complete.  Click OK to complete the script and open Central Administration.", 0))
& "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\BIN\psconfigui.exe" -cmd showcentraladmin