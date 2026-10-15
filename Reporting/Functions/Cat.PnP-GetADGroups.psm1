<#  Get-CatADGroups ( PnP)
    This function Loops through role assigments to find Active Directory groups being used and writes 
    them to a SharePoint list in the Reporting Site. The site, Lists, Content Types, and libraries must exist in advance.

    Command: Get-CatADGroups -Url $VarURL -Creds $VarCred -Assignments $VarAssignments -SiteName $VarSiteName -SiteUrl $VarSiteUrl -WebName $VarWebName -WebUrl $VarWebUrl

	Parameters
		URL - URL of Site
        Creds - Credentials for access to site
        Assignments - Array of assignments that are being reviewed
        SiteName - Name of Site Collection
        SiteURL - URL of Site Collection
        WebName - Name of Web
        WebURl - URL of Web
#>
function global:Get-CatADGroups
{
    param (
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [String]$Url,
        
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [System.Management.Automation.PSCredential]$Creds,
        
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Microsoft.SharePoint.SPRoleAssignmentCollection]$Assignments,
        
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [String]$SiteName,
        
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [String]$SiteUrl,
        
        [parameter(Mandatory=$false, Position=0, ValueFromPipeline=$true)]
        [String]$WebName,
        
        [parameter(Mandatory=$false, Position=0, ValueFromPipeline=$true)]
        [String]$WebUrl,
        
        [parameter(Mandatory=$false, Position=0, ValueFromPipeline=$true)]
        [String]$ListName,
        
        [parameter(Mandatory=$false, Position=0, ValueFromPipeline=$true)]
        [String]$ListUrl,
        
        [parameter(Mandatory=$false, Position=0, ValueFromPipeline=$true)]
        [String]$ItemTitle)
    
    ForEach ($Assignment in $Assignments)
    {
        # Checking Role
        If (($Assignment.Member.LoginName -like "i:0#.w|*") -or ($Assignment.Member.LoginName -like "c:0+.w|*"))
        {
            If ($Assignment.Member.IsDomainGroup)
            {
                # Add to SP List Domain Group
                $Values = @{"MigSCName"=$SiteName;"MigSCURL"=$SiteUrl;"MigWebName"=$WebName;"MigWebURL"=$WebUrl;"MigListName"=$ListName;"MigListURL"=$ListUrl;
                    "MigItemTitle"=$ItemTitle;"MigADGroup"=$Assignment.Member.Name;}
	            $Temp = Add-CatMigrationRecord -Url $Url -Creds $Creds -L "Migration Active Directory Groups" -ContentType "Migration Active Directory Groups Item" -Values $Values
            }
        } # If ($AssignRole.Name -eq $PowerUserRole)
        else 
        {
            # Loop through users in each SP group
            ForEach ($User in $Assignment.Member.Users)
            {
                If ($User.IsDomainGroup)
                {
                    # Add to SP List Domain Group
                    $Values = @{"MigSCName"=$SiteName;"MigSCURL"=$SiteUrl;"MigWebName"=$WebName;"MigWebURL"=$WebUrl;"MigListName"=$ListName;"MigListURL"=$ListUrl;
                    "MigItemTitle"=$ItemTitle;"MigADGroup"=$Assignment.Member.Name;}
	                $Temp = Add-CatMigrationRecord -Url $Url -Creds $Creds -L "Migration Active Directory Groups" -ContentType "Migration Active Directory Groups Item" -Values $Values
                }
            } # ForEach ($User in $Assignment.Member.Users)
        } # If ($Role.Name -eq $PowerUserRole)
    } # ForEach ($Assignment in $Site.RootWeb.RoleAssignments)
}