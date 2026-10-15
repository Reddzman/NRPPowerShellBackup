$tenantDomain = 'ClientSAN'
$adminURL="https://$tenantDomain-admin.sharepoint.com/";

Import-Module Microsoft.Online.Sharepoint.PowerShell -DisableNameChecking;
Connect-SPOService -url $adminURL ;


$global:qSPSiteURLs = New-Object System.Collections.Queue;
$global:qUserDetails = New-Object System.Collections.Queue;
$debug = $true;

if( ($null -ne $MyInvocation.MyCommand.Definition) -and 
    ('' -ne $MyInvocation.MyCommand.Definition) -and
    ($MyInvocation.MyCommand.Definition) -and
    ($MyInvocation.MyCommand.Definition -gt 0)
    ){
        Write-Host $true
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
$userInfoPath = "$outputFolder\$tenantDomain\UserReport.csv"
if(!(Test-Path -Path $userInfoPath)){
	Out-File -FilePath $userInfoPath -InputObject "SiteURL|DisplayName|Groups|IsGroup|IsSiteAdmin|LoginName|UserType" -Encoding unicode
}


Get-SPOSite -Limit ALL | select Url | %{
    $global:qSPSiteURLs.Enqueue($_);
}

try {
	while($global:qSPSiteURLs.count -ne 0){
		$newSiteCount = $global:qSPSiteURLs.count 
		$global:qSPSiteURLs.Dequeue() | %{
            if($_.Url -notmatch '-my.sharepoint.com'){
                $scUrl = $_.Url;
			    #region jobVariables
				    #region Variable for job queue param from single string URL
					    if(($_.GetType().Name -eq 'String') -and ($_.StartsWith('http'))){
						    $siteHash = $_;
						    $qMasterjobName = "Enterprise Subweb from URL String";
					    }
				    #endregion

				    #region Variables for job queue params from SPObj
					    else{
						    $siteHash = ConvertTo-Json $_ 
						    $qMasterjobName = "Enterprise Site Objects Details";
					    }
				    #endregion
									
			    #endregion
								
			    #region Process the queue
                    try{
				        #region If debug mode dot source
					        if (($debug -eq $true) -or ($debug -eq 'true')) {
                                Write-host "Getting Users from Site collection:"$scUrl -f Yellow
                                Get-SPOUser -Limit ALL -Site $scUrl | %{
                                    #if(1 -eq 2){
                                    if($_.LoginName -match '@'){
                                        if($_.LoginName -match '@perpetualreality.com'){
                                            $global:qUserDetails.Enqueue($_);
                                        }
                                    }else{
                                        $global:qUserDetails.Enqueue($_);
                                    }
                                }
					        } 
				        #endregion

	                    while($global:qUserDetails.count -ne 0){
		                    $newUserCount = $global:qUserDetails.count 
		                    $global:qUserDetails.Dequeue() | %{
                                $userDet = $_;
						        Out-File -FilePath $userInfoPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f 
							        $scUrl, 
							        $userDet.DisplayName, 
							        $userDet.Groups, 
							        $userDet.IsGroup,
							        $userDet.IsSiteAdmin,
							        $userDet.LoginName,
                                    $userDet.UserType)
                            }
                        }
									
                    } 
                    catch {
	                    if($debug -eq $true) {write-host -ForegroundColor Red ($Error[0])}
                    } 
			    #endregion
            }
		};
	}
} 
catch {
	if($debug -eq $true) {write-host -ForegroundColor Red ($Error[0])}
} 
finally {
	if($siteHash){Remove-Item variable:siteHash;};
	if($newSiteCount){Remove-Item variable:newCount;}
}


