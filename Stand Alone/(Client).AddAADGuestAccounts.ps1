$tenantDomain = 'ClientSAN';

if(!($creds)) {
    $creds = Get-Credential;
}
Connect-AzureAD -TenantDomain "$tenantDomain.onmicrosoft.com" -Credential $creds;

$global:qUserDetails = New-Object System.Collections.Queue;
$userInfoCSV = "C:\Users\rtoland\OneDrive - Catapult Systems\AE\JustAEUsersForGuestCreation.csv"
Import-Csv -Delimiter ',' -Path $userInfoCSV | %{
    $global:qUserDetails.Enqueue($_);
}

while($global:qUserDetails.count -ne 0){
	$newUserCount = $global:qUserDetails.count 
	$global:qUserDetails.Dequeue() | %{
        $userDet = $_;

        $userDet.DisplayName.Replace('***', ',');
        $userDet.LoginName;

        #https://docs.microsoft.com/en-us/azure/active-directory/external-identities/b2b-quickstart-invite-powershell
        New-AzureADMSInvitation -InvitedUserDisplayName $userDet.DisplayName.Replace('***', ',') -InvitedUserEmailAddress $userDet.LoginName -InviteRedirectURL https://myapps.microsoft.com -SendInvitationMessage $false;
    }
}

