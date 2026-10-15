<#  Get-CatSP10PowerUsers (SP10)
    This function Loops through role assigments and finds power users based on provided permission level.

    Command: $PowerUsers = Get-CatSP10PowerUsers -Assignments $VarAssignments -Level $VarLevel

    Parameters
		Assignments - Array of assignments that are being reviewed
		Level - Desired Permission level
#>
function global:Get-CatSP10PowerUsers
{
    param (
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Microsoft.SharePoint.SPRoleAssignmentCollection]$Assignments,
        
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [String]$Level)
    
    ForEach ($Assignment in $Assignments)
    {
        ForEach ($Role in $Assignment.RoleDefinitionBindings)
        {
            # Determine it is the appropriate roled
            If ($Role.Name -eq $Level)
            {
                # Checking Determining if Assignment is a user or a domina group
                If (($Assignment.Member.IsDomainGroup -eq $true) -or ($Assignment.Member.IsDomainGroup -eq $false))
                {
                    $OutputUsers = ("{0}{1};" -f $OutputUsers, $Assignment.Member.UserLogin)
                } # If (($Assignment.Member.IsDomainGroup -eq $true) -or ($Assignment.Member.IsDomainGroup -eq $false))
                else 
                {
                    # Loop through users in each SP group
                    ForEach ($User in $Assignment.Member.Users)
                    {
                       $OutputUsers = ("{0}{1};" -f $OutputUsers, $User.UserLogin)
                    } # ForEach ($User in $Assignment.Member.Users)
                } # else
            } # If ($Role.Name -eq $Level)
        } # ForEach ($Role in $Assignment.RoleDefinitionBindings)
    } # ForEach ($Assignment in $Assignments)
    Return $OutputUsers
}