@ECHO OFF
REM	*********************************
REM	**** Cat.SP07-PreMigScan.ps1 ****
REM	*********************************

REM This batch file exports a clients 2007 environment config using STSADM.exe to preform discovery prior to a migration.
REM	NOTE:	Do to the environment this was created for I am not sure XML is the proper format for all files.

REM Dependencies
REM	-	SharePoint 2007 STSADM.exe

REM Call batch file with parameters
REM 	ScriptName ClientNameOrAbbriviation > LogFileName
REM		Cat.SP07-PreMigScan.bat {client} > Cat.SP07-PreMigScan-{date}.log

REM	**********************
REM	**** Start Script ****
REM	**********************
ECHO %DATE% %TIME% - Starting script

REM	*******************
REM	**** Variables ****
REM	*******************
ECHO.
ECHO %TIME% - Setting script variables
SET fileExtension=%DATE:~-2%%DATE:~4,2%%DATE:~7,2%%TIME:~0,2%%TIME:~3,2%%TIME:~6,2%
SET webApplications=http://apmoss.aptina.com/ http://ams.aptina.com/ http://apmysite.aptina.com/ http://sjcasp01:45001/ http://mysite.aptina.com/ http://sjcasp01:45000/
REM * If there is only one SSP you will want to comment out the SSP Loop
SET serviceApplications=SSP1 SSP2 SSP3
SET contentDatabases=moss_content_01 moss_content_02

REM	**********************
REM	**** Output Files ****
REM	**********************
ECHO.
ECHO %TIME% - Setting output file names
SET altDomains=%1-AlternateDomainsReport-%fileExtension%.xml
SET dataConnectionFiles=%1-DataConnnectionFilesReport-%fileExtension%.xml
SET deployments=%1-DeploymentsReport-%fileExtension%.xml
SET exemptUserAgents=%1-ExemptUserAgentsReport-%fileExtension%.txt
SET formTemplates=%1-FormTemplatesReport-%fileExtension%.txt
SET services=%1-ServicesReport-%fileExtension%.xml
SET solutions=%1-SolutionsReport-%fileExtension%.xml
SET features=%1-FeaturesReport-%fileExtension%.xml
SET templates=%1-TemplatesReport-%fileExtension%.xml
SET wpPacks=%1-WPPacksReport-%fileExtension%.txt
SET zoneURLs=%1-ZoneURLsReport-%fileExtension%.xml
SET contentDBs=%1-ContentDatabaseReport-%fileExtension%.xml
SET groups=%1-GroupsReport-%fileExtension%.xml
SET roles=%1-RolesReport-%fileExtension%.xml
SET Users=%1-UsersReport-%fileExtension%.xml
SET sites=%1-SitesReport-%fileExtension%.xml
SET webs=%1-WebsReport-%fileExtension%.xml
SET serviceApps=%1-ServiceAppsReport-%fileExtension%.xml
SET sspTimerJobs=%1-SSPTimerJobsReport-%fileExtension%.xml

REM	*************************
REM	**** Variable Output ****
REM	*************************
ECHO   Parameter: Client Name = %1
ECHO   Parameter: Log File Name = %2
ECHO   Variable: File Extension = %fileExtension%
ECHO   Variable: Web Applications = %webApplications%
ECHO   Variable: Service Applications = %serviceApplications%
ECHO   Variable: Alternate Domains = %altDomains%
ECHO   Variable: Data Connection Files = %dataConnectionFiles%
ECHO   Variable: Deployments = %deployments%
ECHO   Variable: Form Templates = %formTemplates%
ECHO   Variable: Services = %services%
ECHO   Variable: Solutions = %solutions%
ECHO   Variable: Features = %features%
ECHO   Variable: Templates = %templates%
ECHO   Variable: WP Packs = %wpPacks%
ECHO   Variable: Zone URLs = %zoneURLs%
ECHO   Variable: Content Databases = %contentDBs%
ECHO   Variable: Groups = %groups%
ECHO   Variable: Roles = %roles%
ECHO   Variable: Users = %users%
ECHO   Variable: Sites = %sites%
ECHO   Variable: Webs = %webs%
ECHO   Variable: Service Applications (SSP) = %serviceApps%
ECHO   Variable: SSP Timer Jbos = %sspTimerJobs%

REM	*********************
REM	**** Main Script ****
REM	*********************

REM	*************************
REM	**** SP 07 Farm Data ****
REM	*************************

ECHO.
ECHO Extracting SP07 Farm Data

REM Retrieving Alternate Domains
ECHO   %TIME% - Retrieving Alternate Domains
stsadm.exe -o enumalternatedomains > %altDomains%

REM Retrieving Data Connection Files
ECHO   %TIME% - Retrieving Data Connection Files
stsadm.exe -o enumdataconnectionfiles > %dataConnectionFiles%
ECHO     **** If this returns results Run enumdataconnectionfiledependants and getdataconnectionfileproperty for each ****

REM Retrieving Deployments
ECHO   %TIME% - Retrieving Deployments
stsadm.exe -o enumdeployments > %deployments%

REM Retrieving Exempt user Agents
ECHO   %TIME% - Retrieving Exempt user Agents
stsadm.exe -o enumexemptuseragents > %exemptUserAgents%

REM Retrieving Form Templates
ECHO   %TIME% - Retrieving Form Templates
stsadm.exe -o enumformtemplates > %formTemplates%

REM Retrieving Services
ECHO   %TIME% - Retrieving Services
stsadm.exe -o enumservices > %services%

REM Retrieving Solutions
ECHO   %TIME% - Retrieving solutions
stsadm.exe -o enumsolutions > %solutions%

REM Retrieving Features
ECHO   %TIME% - Retrieving features
stsadm.exe -o scanforfeatures > %features%

REM Retrieving Templates
ECHO   %TIME% - Retrieving templates
stsadm.exe -o enumtemplates > %templates%

REM Retrieving WP Packs
ECHO   %TIME% - Retrieving WP Packs
stsadm.exe -o enumwppacks > %wpPacks%

REM Retrieving Zone URLs
ECHO   %TIME% - Retrieving Zone URLs
stsadm.exe -o enumzoneurls > %zoneURLs%

REM	************************************
REM	**** SP 07 Web Applciation Data ****
REM	************************************

REM Looping through Web Applications to collect data
(For %%w in (%webApplications%) do (

ECHO.
ECHO Extracting SP07 Web Application Data: %%w

REM Retrieving Content Databases
ECHO   %TIME% - Retrieving Content Databases
stsadm.exe -o enumcontentdbs -Url %%w >> %contentDBs%

REM Retrieving Groups
ECHO   %TIME% - Retrieving Groups
stsadm.exe -o enumgroups -Url %%w >> %groups%

REM Retrieving Roles
ECHO   %TIME% - Retrieving Roles
stsadm.exe -o enumroles -Url %%w >> %roles%

REM Retrieving Users
ECHO   %TIME% - Retrieving Users
stsadm.exe -o enumusers -Url %%w >> %users%

REM Retrieving Sites
ECHO   %TIME% - Retrieving Sites
stsadm.exe -o enumsites -Url %%w >> %sites%

))

REM	*******************************
REM	**** SP 07 Content DB Data ****
REM	*******************************

REM Looping through Content Databases to collect data
(For %%c in (%contentDatabases%) do (

ECHO.
ECHO Extracting SP07 Content DB Data: %%c

REM Retrieving Webs
ECHO   %TIME% - Retrieving Webs
stsadm.exe -o enumallwebs -databasename %%c -includefeatures -includesetupfiles -includewebparts -includeeventreceivers -includecustomlistview >> %webs%

))

REM	************************************
REM	**** SP 07 SSP Data ****
REM	************************************

REM Looping through SSPs to collect data
(For %%s in (%serviceApplications%) do (

ECHO.
ECHO Extracting SP07 SSP Data: %%s

REM Retrieving Service Applciations (SSPs)
ECHO   %TIME% - Retrieving Service Applciations (SSPs)
stsadm.exe -o enumssp -title %%s >> %serviceApps%

REM Retrieving SSP Timer Jobs
ECHO   %TIME% - Retrieving SSP Timer Jobs
stsadm.exe -o enumssptimerjobs -title %%s >> %sspTimerJobs%
REM	enumssptimerjobs

))

REM	*************************
REM	**** Script Complete ****
REM	*************************
ECHO.
ECHO %TIME% - Completing script