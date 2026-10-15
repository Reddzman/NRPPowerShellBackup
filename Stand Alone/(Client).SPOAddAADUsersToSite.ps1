Import-Module Microsoft.Online.Sharepoint.PowerShell -DisableNameChecking;
$tenantDomain = 'ClientSAN';
$adminURL="https://$tenantDomain-admin.sharepoint.com/";


#Create credential object
if(!($azCreds)) {
    $azCreds = Get-Credential;
}

if(!($spGroupCreds)) {
    $spGroupCreds = Get-Credential;
}
Connect-SPOService -url $adminURL -Credential $spGroupCreds
Connect-AzureAD -TenantDomain "$tenantDomain.onmicrosoft.com" -Credential $azCreds;
$ExchOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $spGroupCreds -Authentication Basic -AllowRedirection
Import-PSSession $ExchOnlineSession


$pnpCon = Connect-PnPOnline -Url $adminURL -UseWebLogin -ReturnConnection
Connect-PnPOnline -Url $adminURL -PnPManagementShell -LaunchBrowser 
Connect-PnPOnline -Scopes "Group.ReadWrite.All","User.ReadWrite.All","Mail.Read","Files.Read","ActivityFeed.Read" -Credentials (New-Object System.Management.Automation.PSCredential ("svc_catapult_migration@$tenandDomain.onmicrosoft.com", (ConvertTo-SecureString "R3@llyH@rdP@ssw0rd!" -AsPlainText -Force)))
Connect-PnPOnline -Scopes "Group.ReadWrite.All","User.ReadWrite.All","Mail.Read","Files.Read","ActivityFeed.Read" -Credentials $spGroupCreds

$skipList = @(
    "https://$tenanDomain.sharepoint.com/teams/Site1",
    "https://$tenanDomain.sharepoint.com/teams/Site2"
)

$global:qSPSiteURLs = New-Object System.Collections.Queue;
$global:qUserDetails = New-Object System.Collections.Queue;
$debug = $true;

if( ($null -ne $MyInvocation.MyCommand.Definition) -and 
    ('' -ne $MyInvocation.MyCommand.Definition) -and
    ($MyInvocation.MyCommand.Definition) -and
    ($MyInvocation.MyCommand.Definition -gt 0)
    ){
        #Write-Host $true;
        $global:currDir = (Split-Path $MyInvocation.MyCommand.Definition)
}

if($global:currDir){
    $global:currDir = (Split-Path $MyInvocation.MyCommand.Definition)
    $outputFolder = ($global:currDir + "\" + $tenantDomain);
    if(!(Test-Path -Path $outputFolder)){
        mkdir  $outputFolder | Out-Null;
    }
}else{
    $outputFolder = ("C:\Temp\$tenantDomain");
    mkdir  $outputFolder | Out-Null;
}
$userInfoPath = "$outputFolder\JustClientUsersReport.csv"

if(Test-Path -Path "$outputFolder\SPOSitesLog.csv"){}else{
	Out-File -FilePath "$outputFolder\SPOSitesLog.csv" -InputObject "SiteURL|DisplayName|Groups|IsGroup|IsSiteAdmin|LoginName|UserType" -Encoding unicode
}
if(Test-Path -Path "$outputFolder\SPO365GroupSitesLog.csv"){}else{
	Out-File -FilePath "$outputFolder\SPO365GroupSitesLog.csv" -InputObject "SiteURL|DisplayName|Groups|IsGroup|IsSiteAdmin|LoginName|UserType" -Encoding unicode
}
if(Test-Path -Path "$outputFolder\SPOErrorLog.csv"){}else{
	Out-File -FilePath "$outputFolder\SPOErrorLog.csv" -InputObject "SiteURL|DisplayName|Groups|IsGroup|IsSiteAdmin|LoginName|UserType|Error" -Encoding unicode
}

#Read in Users found in Sites
Import-Csv -Delimiter ',' -Path $userInfoPath | %{
    if(!($skipList.Contains($_.SiteURL))){
        $global:qUserDetails.Enqueue($_);
    }
}

#Process the queue
while($global:qUserDetails.count -ne 0){
	$newUserCount = $global:qUserDetails.count 
	$global:qUserDetails.Dequeue() | %{
        try{
            $userDet = $_;
            $userDet;

            #https://docs.microsoft.com/en-us/azure/active-directory/external-identities/b2b-quickstart-invite-powershell
            #New-AzureADMSInvitation -InvitedUserDisplayName $userDet.DisplayName -InvitedUserEmailAddress $userDet.LoginName -InviteRedirectURL https://myapps.microsoft.com -SendInvitationMessage $false;

            #Determine if Site is O365 Group or not
            if(($userDet.SiteURL -match '/sites/') -or 
                ($userDet.SiteURL -match '/teams/')
            ){ 
                $grpCheck = Get-UnifiedGroup -Identity $userDet.SiteUrl.Split('/')[$userDet.SiteUrl.Split('/').length-1] -ErrorAction SilentlyContinue
                if(($null -ne $grpCheck) -and ($grpCheck.SharePointSiteUrl -eq $userDet.SiteURL)){
                    $grp = $grpCheck;
                }else{
                    if(($null -ne $grpCheck) -and ('' -ne $grpCheck) -and ($grpCheck)){
                        Write-Host -ForegroundColor Red 'Looking up O365 Group by URL ...';
                        $grp = Get-UnifiedGroup -IncludeAllProperties | ?{$_.SharePointSiteUrl -eq $userDet.SiteURL};
                    }
                }
            }

            #Log 365Group vs SPO
            If($grp){
                Out-File -FilePath "$outputFolder\SPO365GroupSitesLog.csv" -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f 
				    $userDet.SiteURL, 
				    $userDet.DisplayName, 
				    $userDet.Groups, 
				    $userDet.IsGroup,
				    $userDet.IsSiteAdmin,
				    $userDet.LoginName,
                    $userDet.UserType)
            }else{
                Out-File -FilePath "$outputFolder\SPOSitesLog.csv" -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f 
			        $userDet.SiteURL, 
			        $userDet.DisplayName, 
			        $userDet.Groups, 
			        $userDet.IsGroup,
			        $userDet.IsSiteAdmin,
			        $userDet.LoginName,
                    $userDet.UserType)        
            }

            If ($userDet.IsSiteAdmin -eq 'False'){
                try{
                    #Add user as Member
                    If($grp){
                        Write-Host -ForegroundColor Yellow "Identified O365 Group"
                        #https://docs.microsoft.com/en-us/powershell/module/exchange/add-unifiedgrouplinks?view=exchange-ps
                        Add-UnifiedGroupLinks -Identity $grp.Id -LinkType Member -Links $userDet.LoginName; 

                        ###Add-PnPUnifiedGroupMember -Users $userDet.LoginName;
                    }else{
                        Write-Host -ForegroundColor Green "SharePoint Site Collection"
                        $userDet.SiteURL

			            switch -wildcard ($userDet.SiteURL){
				            '*/portals/*' {
                                $groups = Get-SPOSiteGroup -Site $userDet.SiteURL;
                                $group = ($groups | ?{$_.Title -match 'Contributors'})[0];				        
					            break
				            }
				            '*/teams/*' {
                                $groups = Get-SPOSiteGroup -Site $userDet.SiteURL;
                                $group = ($groups | ?{$_.Title -match 'Member'})[0];
					            break
				            }
				            default{
                                $groups = Get-SPOSiteGroup -Site $userDet.SiteURL;
                                $group = ($groups | ?{$_.Title -match 'Member'})[0];				        
					            break
				            }
			            }
                        #Add user to Site as member 
                        Add-SPOUser -Site $userDet.SiteURL -LoginName $userDet.LoginName -Group $group.LoginName ;
                    }
                }
                catch{
	                if($debug -eq $true) {write-host -ForegroundColor Red ($Error[0])}
                    Out-File -FilePath "$outputFolder\SPOErrorLog.csv" -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}" -f 
				        $userDet.SiteURL, 
				        $userDet.DisplayName, 
				        $userDet.Groups, 
				        $userDet.IsGroup,
				        $userDet.IsSiteAdmin,
				        $userDet.LoginName,
                        $userDet.UserType,
                        $Error[0])

                }
            }else{
                try{
                    #Add user to Site as Owner
                    If($grp){
                        Write-Host -ForegroundColor Yellow "Identified O365 Group"
                        #https://docs.microsoft.com/en-us/powershell/module/exchange/add-unifiedgrouplinks?view=exchange-ps
                        Add-UnifiedGroupLinks -Identity $grp.Id -LinkType Owner -Links $userDet.LoginName; 

                    }else{
                        Write-Host -ForegroundColor Green "SharePoint Site Collection"
                        $userDet.SiteURL

                        #Accomodate special managed paths
                        switch -wildcard ($userDet.SiteURL){
				            '*/portals/*' {
                                $groups = Get-SPOSiteGroup -Site $userDet.SiteURL;
                                $group = ($groups | ?{$_.Title -match 'Contributors'})[0];				        
					            break
				            }
				            '*/teams/*' {
                                $groups = Get-SPOSiteGroup -Site $userDet.SiteURL;
                                $group = ($groups | ?{$_.Title -match 'Owner'})[0];
					            break
				            }
				            default{
                                $groups = Get-SPOSiteGroup -Site $userDet.SiteURL;
                                $group = ($groups | ?{$_.Title -match 'Owner'})[0];				        
					            break
				            }
			            }
                        #Add user to Site as owner
                        Add-SPOUser -Site $userDet.SiteURL -LoginName $userDet.LoginName -Group $group.LoginName ;
                    }
                }
                catch{
	                if($debug -eq $true) {write-host -ForegroundColor Red ($Error[0])}
                    Out-File -FilePath "$outputFolder\SPOErrorLog.csv" -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}" -f 
				        $userDet.SiteURL, 
				        $userDet.DisplayName, 
				        $userDet.Groups, 
				        $userDet.IsGroup,
				        $userDet.IsSiteAdmin,
				        $userDet.LoginName,
                        $userDet.UserType,
                        $Error[0])

                }
            }
        }
        catch{
	        if($debug -eq $true) {write-host -ForegroundColor Red ($Error[0])}
        }
        finally{
			if($grp){Remove-Item variable:grp;};
			if($grpCheck){Remove-Item variable:grpCheck;};

        }

    }
}
