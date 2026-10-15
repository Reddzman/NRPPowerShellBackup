# Double check OneDrive URLs
Connect-SPOService -Credential $creds -Url https://client-admin.sharepoint.com
Get-SPOSite -limit all -IncludePersonalSite $true | where {$_.url -like "*Clayton*"}



<# Delete Jobs in Batches

    ***  This script errors on the first row every time because it is the header.

    ***  Be very carefull, run one line at a time because it is really easy to delete all jobs.

#>
$jobs = skysync-cli jobs list --search "Ryan.Wahrenbrock"
$jobs.Count # This seems to get batches of 20; Also Count will always be 1 more than shown in the UI
$jobs # Always double check this list before running the next command
$jobs | Foreach {skysync-cli jobs delete $_.Substring(0,$_.IndexOf(" "))} # Be Very Very Careful; First item always errors

# Manually create job based on our CreateJob script - Line by line execution

$migrationFolder = "Box Migration"
$serverURL = "http://localhost:9090/"
$baseURL = ("{0}v1" -f $serverURL)
$contentType = "application/json"
$accessToken = skysync-cli sessions token --output-json false
$authHeader = @{"Authorization"="Bearer $accessToken";}
$templateFile = "Cat.SkySync-SPOSimTemp-client-ConflictLatest.json"
$templatePath = ("..\_Input Files\{0}" -f $templateFile)
$templateInput = Get-Content -Path $templatePath -Raw | ConvertFrom-Json
$template = $templateInput
$template.category.name = "Wave6-Services"
$template.name = ("SPO-Wave6-Services-Demand-ClientInternalRepository-{0}" -f $letter)
$template.transfer.source.connection.id = "eab374d7852d417abce3ac258a98ad22"
$template.transfer.source.impersonate_as.email = "Jason.Prost@client.com"
$template.transfer.source.target.path = ("/+Client Project Repository/Internal/{0}" -f $letter)
$template.transfer.destination.connection.id = "eab374d7852d417abce3ac258a98ad22"
$template.transfer.source.connection.id = "66a56d049e77420095094606be91d119"
$template.transfer.destination.target.path = ("/Internal/{0}" -f $letter)
$template.transfer.account_map.id = "29ff94ab6df74a9aac769238350cc33a"
$template.transfer.group_map.id = "6e687f173d5a4e3b8fde058ae447f3c1"
$spoJob = $template | ConvertTo-Json -Depth 10
$createJobMethod = "$baseUrl/jobs"
Invoke-RestMethod $createJobMethod -Headers $authHeader -Method Post -ContentType $contentType -Body $spoJob
