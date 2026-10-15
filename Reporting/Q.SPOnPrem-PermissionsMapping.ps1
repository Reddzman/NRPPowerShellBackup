<#  Q.SP-SPOnPrem-PermissionsMapping.ps1

This script scans sharepoint permissions for the provided list mapping file and outputs a permission mapping file
that can be used by the Q.PnP-AddusersToGroups.ps1 as it's input file for applying permissions. 

Dependencies
	-	SharePoint 2013 / SharePoint 2016 Module
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1

Call script with parameters
	.\Q.SP-SPOnPrem-PermissionsMapping.ps1  -jobFile "Q.SitesList-Client.csv" -client "ClientName" -getAllWebs $false -domainName "clientdomain.com"

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.  Functions
	5.	Main Script
#>

#region Script Parameters
Param(	[string] $jobFile = "Q.SitesList-Client.csv", #You can use the sharegate list mapping file but will produce duplicates as site urls will be listed multiple times.
		[string] $client = "clientname",
		[string] $domainName = "clientdomain.com", #Used as the email domain address when building user's UPN.
		[boolean] $getAllWebs = $false # set to $true to get all sub sites if url is a site collection in csv
)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log FIles\{0}.SP-SPOnPrem-PermissionsMapping-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)
$outPath = ("..\_Output Files\Q.SPOnPrem-PermissionsMapping-{0}-{1}.csv" -f $client, $scriptStart[1])
$errorPath = ("..\_Output Files\{0}.SPOnPrem-PermissionsMapping-{1}.log" -f $client, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode | Sort-Object -Property @{Expression = "srcSiteUrl"} 
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}

#region 2.	Variables

# Used to configure UPN (i.e. if user, not a security group)
$userClaimPrefix = "i:0#.w|" + $client.ToLower() + "\"

# Skip these user accounts if found
$accExclusions = [PSObject]@(
   "sharepoint\system",
   "c:0!.s|windows",
   "service account,sp_setup",
   $client + "\srvsp_setup"
)

# Capture previous job run for avoiding duplicates
$prevJob = ""

#endregion

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   Client Tenant: {0}" -f $client)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion

#region 4. Functions
function Write-QPermissionsInfo { 
	param(
        [psobject] $pJob,
        [string] $pOutputPath
    )

    #$pJob # Uncomment for job information troubleshooting
	try {
        # Ouput List Information
        $csvRecord = New-Object -TypeName PSObject -Property (@{
			    userName		= $pJob.UserName;
                loginName		= $pJob.Principal;
			    loginType 		= $pJob.LoginType;
			    groupName		= $pJob.GroupName;
                siteUrl			= $pJob.DestSiteUrl;
                srcSiteURL		= $pJob.SrcSiteUrl;
			    groupType		= $pJob.GroupType;
			    groupRole		= $pJob.GroupRole;
                groupRoles		= $pJob.GroupRoles;
                permissions		= $pJob.Permissions;
                #groupLoginName  = $pJob.GroupLoginName;
                #groupID         = $pJob.GroupID;
                #parentWeb       = $pjob.ParentWeb;
		})

        $csvRecord | Export-Csv -Path $pOutputPath -NoTypeInformation -Append
        Remove-Variable csvRecord
        #Write-Host " - Success" -ForegroundColor Green
    } 
	catch {
        # Ouput error
        #Write-ScriptError -funcJob $job -funcAction "Unable to write to csv" -funcException $null -funcDetails $null -funcStackTrace $null -funcLogPath $errorPath
	}
}
#endregion
    
#region 5.	Main Script

# Looping through input file
:jobLoop ForEach ($job in $inFile){

	# Only scan sites that haven't already been scanned
	If($prevJob.srcSiteUrl -ne $job.srcSiteUrl) {
		# Screen Output
		Write-Host
		Write-Host ("Scanning site collection: {0}" -f $job.srcSiteURL) -ForegroundColor Magenta

		# Get Site Collection
		Write-Host ("   Retrieving site and web objects: {0}" -f $job.srcSiteURL)
		#$site = Get-SPSite -Identity $job.srcSiteURL
		#$webs = $site.AllWebs
		$site = Get-SPWeb $job.srcSiteURL

		# If web is the root web
		If($web.Url -eq $web.Site.RootWeb.Url) {

			# Get all group permissions from all webs in site collection
			If($getAllWebs) {
				#get all webs
				$webs = $site.AllWebs

				# Looping through Web objects
				ForEach ($web in $webs) {
					
					# Get all Groups for web
					Write-Host ("   Retrieving groups for web: {0}" -f $web.Url)
					$groups = $web.groups

					# Looping through Groups
					ForEach ($group in $groups) {
						# Get all users for group
						Write-Host ("      Retrieving users for group: {0}" -f $group.Name) -ForegroundColor DarkYellow
						$users = $group.users

						# Replacing "|" characters where necessary in Group Info
						If ($group.Name) {$groupName = $group.Name.Replace("|","\")} Else {$groupName = ""}
						If ($group.LoginName) {$groupLoginName = $group.LoginName.Replace("|","\")} Else {$groupLoginName = ""}
						If ($group.Roles.Name) {$groupRoles = $group.Roles.Name.Replace("|","\") -join ","} Else {$groupRoles = ""}

						# Determine the group name for SPO site
						If ($groupName.contains("Owners"))  {
							$groupNameSPO = $job.destSiteTitle + " Owners"
						} elseif ($groupName.contains("Members")) {
							$groupNameSPO = $job.destSiteTitle + " Members"
						} elseif ($groupName.contains("Visitors")) {							
							$groupNameSPO = $job.destSiteTitle + " Visitors"
						} else {
							$groupNameSPO = $groupName
						}

						# Looping through Users
						ForEach ($user in $users) {

							# Check if User is on the exclusion list
							$nonExcludedMember = $accExclusions -notcontains $user.LoginName.ToLower()
							if ($nonExcludedMember) {

								# Screen Output
								Write-Host ("         Writing user record: {0} ..." -f $user.DisplayName) -ForegroundColor DarkCyan

								# Replacing "|" characters where necessary in User Info
								If ($user.DisplayName) {$userDisplayName = $user.DisplayName.Replace("|","\")} Else {$userDisplayName = ""}
								If ($user.Roles.Name) {$userRoles = $user.Roles.Name.Replace("|","\") -join ","} Else {$userRoles = ""}
								
								# Build UPN & login type
								If ($user.LoginName.ToLower().contains($userClaimPrefix)) {
									# User Account
									$memberUPN = $user.LoginName.Replace($userClaimPrefix, '')
									$memberUPN += "@" + $domainName
									$loginType = "User"
								} else {
									# Security Group
									if($user.DisplayName.ToLower() -eq "everyone") {
										$memberUPN = "Everyone except external users"
									} else {
										$memberUPN = $user.DisplayName
									}
									$loginType = "SecurityGroup"
								}

								# Build output data object
								$permissionData = @{
									DestSiteUrl       = $job.destSiteUrl
									SiteTitle         = $web.Title
									SrcSiteUrl        = $job.srcSiteURL
									Principal         = $memberUPN
									GroupName		  = $(if($job.siteTemplate.ToLower() -eq 'msteams'){$job.destGroupAlias}else{$groupNameSPO})
									GroupLoginName	  = $groupLoginName
									UserName		  = $userDisplayName
									Permissions       = $(if($job.siteTemplate.ToLower() -eq 'msteams'){"M365Group"}else{$userRoles})
									GroupType	   	  = $(if($job.siteTemplate.ToLower() -eq 'msteams'){"M365Group"}else{"SharePoint"})
									GroupRole		  = $(if($job.siteTemplate.ToLower() -eq 'msteams'){"Member"}else{"None"})
									GroupRoles 		  = $groupRoles
									GroupID			  = $group.Id
									ParentWeb		  = $group.ParentWeb.Url
									LoginType	  	  = $loginType
								}

								# Output List information to CSV File
								Write-QPermissionsInfo -pJob $permissionData -pOutputPath $outPath

								# If migrating to MSTeams and user is an owner then add the user a second time but as "owner"
								If($job.siteTemplate.ToLower() -eq 'msteams' -and $group.Name.ToLower().contains('owners')){

									# Build output data object
									$permissionData = @{
										DestSiteUrl       = $job.destSiteUrl
										SiteTitle         = $web.Title
										SrcSiteUrl        = $job.srcSiteURL
										Principal         = $memberUPN
										GroupName		  = $job.destGroupAlias
										GroupLoginName	  = $groupLoginName
										UserName		  = $userDisplayName
										Permissions       = "M365Group"
										GroupType	   	  = "M365Group"
										GroupRole		  = "Owner"
										GroupRoles 		  = $groupRoles
										GroupID			  = $group.Id
										ParentWeb		  = $group.ParentWeb.Url
										LoginType	  	  = $loginType
									}

									# Output List information to CSV File
									Write-QPermissionsInfo -pJob $permissionData -pOutputPath $outPath
								}
							}
						}
					}

				}
			} else {

				# Get all Groups for web
				Write-Host ("   Retrieving groups for web: {0}" -f $site.Url)
				$groups = $site.groups

				# Looping through Groups
				ForEach ($group in $groups) {
					# Get all users for group
					Write-Host ("      Retrieving users for group: {0}" -f $group.Name) -ForegroundColor DarkYellow
					$users = $group.users

					# Replacing "|" charecters where necessary in Group Info
					If ($group.Name) {$groupName = $group.Name.Replace("|","\")} Else {$groupName = ""}
					If ($group.LoginName) {$groupLoginName = $group.LoginName.Replace("|","\")} Else {$groupLoginName = ""}
					If ($group.Roles.Name) {$groupRoles = $group.Roles.Name.Replace("|","\") -join ","} Else {$groupRoles = ""}

					# Determine the group name for SPO site
					If ($groupName.contains("Owners"))  {
						$groupNameSPO = $job.destSiteTitle + " Owners"
					} elseif ($groupName.contains("Members")) {
						$groupNameSPO = $job.destSiteTitle + " Members"
					} elseif ($groupName.contains("Visitors")) {							
						$groupNameSPO = $job.destSiteTitle + " Visitors"
					} else {
						$groupNameSPO = $groupName
					}

					# Looping through Users
					ForEach ($user in $users) {

						# Check if User is on the exclusion list
						$nonExcludedMember = $accExclusions -notcontains $user.LoginName.ToLower()
						if ($nonExcludedMember) {

							# Screen Output
							Write-Host ("         Writing user record: {0} ..." -f $user.DisplayName) -ForegroundColor DarkCyan

							# Replacing "|" charecters where necessary in User Info
							If ($user.DisplayName) {$userDisplayName = $user.DisplayName.Replace("|","\")} Else {$userDisplayName = ""}
							#If ($user.Email) {$userEmail = $user.Email.Replace("|","\")} Else {$userEmail = ""}
							#If ($user.LoginName) {$userLoginName = $user.LoginName.Replace("|","\")} Else {$userLoginName = ""}
							If ($user.Roles.Name) {$userRoles = $user.Roles.Name.Replace("|","\") -join ","} Else {$userRoles = ""}

							# Build UPN & login type
							If ($user.LoginName.ToLower().contains($userClaimPrefix)) {
								# User Account
								$memberUPN = $user.LoginName.Replace($userClaimPrefix, '')
								$memberUPN += "@" + $domainName
								$loginType = "User"
							} else {
								# Security Group
								if($user.DisplayName.ToLower() -eq "everyone") {
									$memberUPN = "Everyone except external users"
								} else {
									$memberUPN = $user.DisplayName
								}
								$loginType = "SecurityGroup"
							}

							# Build output data object
							$permissionData = @{
								DestSiteUrl       = $job.destSiteUrl
								SiteTitle         = $web.Title
								SrcSiteUrl        = $job.srcSiteURL
								Principal         = $memberUPN
								GroupName		  = $(if($job.siteTemplate.ToLower() -eq 'msteams'){$job.destGroupAlias}else{$groupNameSPO})
								GroupLoginName	  = $groupLoginName
								UserName		  = $userDisplayName
								Permissions       = $(if($job.siteTemplate.ToLower() -eq 'msteams'){"M365Group"}else{$userRoles})
								GroupType	   	  = $(if($job.siteTemplate.ToLower() -eq 'msteams'){"M365Group"}else{"SharePoint"})
								GroupRole		  = $(if($job.siteTemplate.ToLower() -eq 'msteams'){"Member"}else{"None"})
								GroupRoles 		  = $groupRoles
								GroupID			  = $group.Id
								ParentWeb		  = $group.ParentWeb.Url
								LoginType	  	  = $loginType
							}

							# Output List information to CSV File
							Write-QPermissionsInfo -pJob $permissionData -pOutputPath $outPath

							# If migrating to MSTeams and user is an owner then add the user a second time but as "owner"
							If($job.siteTemplate.ToLower() -eq 'msteams' -and $group.Name.ToLower().contains('owners')){

								# Build output data object
								$permissionData = @{
									DestSiteUrl       = $job.destSiteUrl
									SiteTitle         = $web.Title
									SrcSiteUrl        = $job.srcSiteURL
									Principal         = $memberUPN
									GroupName		  = $job.destGroupAlias
									GroupLoginName	  = $groupLoginName
									UserName		  = $userDisplayName
									Permissions       = "M365Group"
									GroupType	   	  = "M365Group"
									GroupRole		  = "Owner"
									GroupRoles 		  = $groupRoles
									GroupID			  = $group.Id
									ParentWeb		  = $group.ParentWeb.Url
									LoginType	  	  = $loginType
								}

								# Output List information to CSV File
								Write-QPermissionsInfo -pJob $permissionData -pOutputPath $outPath
							}
						}
					}
				}

			}
		}
	}

	Clear-Variable prevJob
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion