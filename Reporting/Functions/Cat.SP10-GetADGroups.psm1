<#  Get-CatSP10ADGroups (SP10)
    This function Loops through role assigments to find Active Directory groups being used and writes 
    them to a CSV file. The site, Lists, Content Types, and libraries must exist in advance.

    Command: Get-CatSP10ADGroups -File $VarFile -Assignments $VarAssignments -SiteName $VarSiteName -SiteUrl $VarSiteUrl -WebName $VarWebName -WebUrl $VarWebUrl -ListName $VarListName -LURL $VarListURL -ItemName $VarItemName

	Parameters
		File - Path to CSV File
        Assignments - Array of assignments that are being reviewed
        SiteName - Name of Site Collection
        SiteURL - URL of Site Collection
        WebName - Name of Web
        WebURL - URL of Web
        ListName - Name of List
        LURL - URL Of List
        ItemTitle - Name of Item
#>
function global:Get-CatSP10ADGroups
{
    param (
            [String]$File,
            [Microsoft.SharePoint.SPRoleAssignmentCollection]$Assignments,
            [String]$SiteName,
            [String]$SiteUrl,
            [String]$WebName,
            [String]$WebUrl,
            [String]$ListName,
            [String]$LUrl,
            [String]$ItemTitle
    )
    
    ForEach ($Assignment in $Assignments)
    {
        # Checking to determine if it is a user or group
        If (($Assignment.Member.IsDomainGroup -eq $true) -or ($Assignment.Member.IsDomainGroup -eq $false))
        {
            If ($Assignment.Member.IsDomainGroup)
            {
                # Add to SP List Domain Group
                Out-File -FilePath $File -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}" -f 
                    $SiteName, 
                    $SiteUrl, 
                    $WebName, 
                    $WebURL, 
                    $ListName, 
                    $LURL, 
                    $ItemTitle, 
                    $Assignment.Member.UserLogin.Replace("|","\"))
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
                    Out-File -FilePath $File -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}" -f 
                        $SiteName.Replace("|","\"), 
                        $SiteUrl, 
                        $WebName.Replace("|","\"), 
                        $WebURL, 
                        $ListName.Replace("|","\"), 
                        $LURL, 
                        $ItemTitle.Replace("|","\"), 
                        $User.UserLogin.Replace("|","\"))
                }
            } # ForEach ($User in $Assignment.Member.Users)
        } # If ($Role.Name -eq $PowerUserRole)
    } # ForEach ($Assignment in $Site.RootWeb.RoleAssignments)
}