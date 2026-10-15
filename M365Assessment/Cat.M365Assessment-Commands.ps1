<#

These command will help you run the M365 Assessment tool.  They are expecting the exe file to exist in the M365Assessment 
folder. Please do no not sync the EXE in GitHub.

You have to run the command for each different type of report you want to pull from the tenant. This tool is continuously changes
so make sure you validate the current modes/reports available to you.  

The second command will convert the report generated to CSV for our internal Power BI template.

#>

#Run Syntex Report
cd C:\GitHub\CatPowerShell\M365Assessment\
$clientTenant = "nrpdev2"
$clientAppID = "caf6b11a-cf03-416c-b0f0-67a09a8e30ee"
$clientCertPath = "My|CurrentUser|4291b1e987db4d2679043168ee119329c8afd06f"
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
# Modes: <AddInsACS|Classic|InfoPath|Syntex|Workflow>
# Get Status by running exe and using the status command
# Syntex
.\microsoft365-assessment.exe start --mode syntex --authmode application --tenant "$clientTenant.sharepoint.com" --applicationid $clientAppID --certpath $clientCertPath --syntexfull

# Others
.\microsoft365-assessment.exe start --mode Workflow --authmode application --tenant "$clientTenant.sharepoint.com" --applicationid $clientAppID --certpath $clientCertPath

# Extacting CSV
# Get reports by running exe and using the list command
$report = "716932b5-8e42-49ea-b502-0ba3c05d8d2a"
$reportFolder = "Syntex"
$OutputPath = "C:\Users\npatterson.CORP\Quisitive\Client-REMAX International - Documents\Project-SharePoint Staff Augmentation\Team Documents\Discovery\M365Assessment"
.\microsoft365-assessment.exe report --id $report --mode CsvOnly --path ("{0}\{1}" -f $OutputPath, $reportFolder)