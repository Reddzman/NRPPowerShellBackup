<#  Get-CatPowerUsers (SP13)
    This function Loops through role assigments and finds power users based on provided permission level.

    Command: $PowerUsers = Get-CatPowerUsers -Assignments $VarAssignments -Level $VarLevel

    Parameters
		Assignments - Array of assignments that are being reviewed
		Level - Desired Permission level
#>
function global:Get-CatPowerUsers
{
    param (
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Microsoft.SharePoint.SPRoleAssignmentCollection]$Assignments,
        
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [String]$Level)
    
    $OutputUsers = @()
    
    ForEach ($Assignment in $Assignments)
    {
        ForEach ($Role in $Assignment.RoleDefinitionBindings)
        {
            # Determine if it is a user or a group
            If ($Role.Name -eq $Level)
            {
                # Checking Role
                If (($Assignment.Member.LoginName -like "i:0#.w|*") -or ($Assignment.Member.LoginName -like "c:0+.w|*"))
                {
                    If ($Assignment.Member.IsDomainGroup)
                    {
                        # Add Domain Group
                        $OutputUsers += $Assignment.Member.DisplayName + ";"
                    }
                    else 
                    {
                         # Add user name
                        $OutputUsers += $Assignment.Member.LoginName.Replace("i:0#.w|","").Replace("i:0#.f|","").Replace("5652;","") + ";"   
                    } # If ($User.IsDomainGroup)
                } # If ($AssignRole.Name -eq $PowerUserRole)
                else 
                {
                    # Loop through users in each SP group
                    ForEach ($User in $Assignment.Member.Users)
                    {
                        If ($User.IsDomainGroup)
                        {
                            # Add Domain Group
                            $OutputUsers += $Assignment.Member.DisplayName + ";"
                        }
                        else 
                        {
                            # Add user name
                            $OutputUsers += $User.UserLogin.Replace("i:0#.w|","").Replace("i:0#.f|","").Replace("5652;","") + ";"   
                        } # If ($User.IsDomainGroup)
                    } # ForEach ($User in $Assignment.Member.Users)
                } # else
            }
        } # If ($Role.Name -eq $PowerUserRole)
    } # ForEach ($Assignment in $Site.RootWeb.RoleAssignments)
    $OutputUsers
}