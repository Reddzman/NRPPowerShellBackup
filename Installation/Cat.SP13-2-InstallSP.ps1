#***********************************************************************************************************************************************

# This script assumes that the SharePoint setup files have been copied to a single location and prerequisites have been downloaded.  This should
# have been done for the Prerequiste script. If not, please perform the following tasks before running the script.
#    1. Extract the setup ISO to a network share
#    2. Download prerequisites and copy them to "prerequisiteinstallfiles" folder - List can be found here: 
#                                                                           http://technet.microsoft.com/en-us/library/cc262485.aspx#section5
#    3. Pick the appropriate "Config.xml" file from software share \ files folder, append it with the correct license key and save it to the
#       root of the software share.  Same location as SharePoint's setup.exe file
#    4. Provide the path to the software share in the "Farm-Settings.xml" file.  Configuration.Software.SharePoint

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
$Task = "Install SharePoint"
$ScriptErrors = @()
$Logfile = "c:\\scripts\\logs\\FarmInstall_$date.log.txt"
$InputFile = "c:\scripts\Farm-Settings.xml"

# Creating Log file
write-host "____________________________________________________"
out-file -filepath $Logfile -inputobject "____________________________________________________"
write-host "Starting $Task - $Start"
out-file -filepath $Logfile -append -inputobject "Starting $Task - $Start"

#***********************************************************************************************************************************************

#                                         Variables

#***********************************************************************************************************************************************

# Importing XML
If ((test-path $InputFile) -eq $True)
{
    write-host "____________________________________________________"
    out-file -filepath $Logfile -append -inputobject "____________________________________________________"
    write-host ("   Importing XML file " + $InputFile)
    out-file -filepath $Logfile -inputobject ("   Importing XML file " + $InputFile)
    [xml]$ConfigXML = get-content $InputFile -ErrorVariable +ScriptErrors
}
Else
{
    write-host "____________________________________________________"
    out-file -filepath $Logfile -append -inputobject "____________________________________________________"
    write-host ($InputFile + " doesn't exist. Terminating Script")
    out-file -filepath $Logfile -inputobject ($InputFile + " doesn't exist. Terminating Script")
    write-host "____________________________________________________"
    out-file -filepath $Logfile -append -inputobject "____________________________________________________"
    $finish = (get-date -format "MM-dd-yyyy HH:mm")
    write-host "$Task Complete - $finish"
    out-file -filepath $Logfile -append -inputobject "$Task Complete - $finish"
    $Prompt = new-object -comobject wscript.shell 
}

# Software Variables
$SPInstall = $ConfigXML.Configuration.Software.SharePoint

# Outputing variables
write-host "____________________________________________________"
write-host "Intial Configuration variables Successfully assigned"
write-host "____________________________________________________"
write-host "----------------------------------------------------"
write-host "Script Variables"
write-host "----------------------------------------------------"
write-host "   Log file: $Logfile"
write-host "----------------------------------------------------"
write-host "Software Variables"
write-host "----------------------------------------------------"
write-host "   SharePoint Install Directory: $SPInstall"
write-host "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "Intial Configuration variables Successfully assigned"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Script Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   Log file: $Logfile"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "Software Variables"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"
out-file -filepath $Logfile -append -inputobject "   SharePoint Install Directory: $SPInstall"
out-file -filepath $Logfile -append -inputobject "----------------------------------------------------"

# Main Body
# Running SharePoint Installer
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
write-host "   Installing SharePoint Server..."
out-file -filepath $Logfile -append -inputobject "   Installing SharePoint Server 2013..."
Start-Process "$SPInstall\Setup.exe" -ArgumentList "/Config $SPInstall\Config.xml" -ErrorVariable +ScriptErrors

# Logging Errors to log file
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "Errors during Script:"
out-file -filepath $Logfile -append -inputobject ($ScriptErrors)

# Finishing Script
write-host "____________________________________________________"
out-file -filepath $Logfile -append -inputobject "____________________________________________________"
$finish = (get-date -format "MM-dd-yyyy HH:mm")
write-host "$Task Complete - $finish"
out-file -filepath $Logfile -append -inputobject "$Task Complete - $finish"

$Prompt = new-object -comobject wscript.shell 
out-file -filepath $Logfile -append -inputobject ($Prompt.popup("Farm Configuration is complete.  Click OK to complete the script.", 0))