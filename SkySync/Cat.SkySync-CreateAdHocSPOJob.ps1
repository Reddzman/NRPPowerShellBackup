Param(  
	[String] $letter = "A"

)
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
$template.transfer.source.connection.id = "66a56d049e77420095094606be91d119"
$template.transfer.source.impersonate_as.email = "Jason.Prost@client.com"
$template.transfer.source.target.path = ("/+Client Project Repository/Internal/{0}" -f $letter)
$template.transfer.destination.connection.id = "eab374d7852d417abce3ac258a98ad22"
$template.transfer.destination.target.path = ("/Internal/{0}" -f $letter)
$template.transfer.account_map.id = "29ff94ab6df74a9aac769238350cc33a"
$template.transfer.group_map.id = "6e687f173d5a4e3b8fde058ae447f3c1"
$spoJob = $template | ConvertTo-Json -Depth 10
$jobName = $template.name
$createJobMethod = "$baseUrl/jobs"
Invoke-RestMethod $createJobMethod -Headers $authHeader -Method Post -ContentType $contentType -Body $spoJob

Write-Host ("   Starting Job: {0}" -f $jobName)
$startJob = skysync-cli jobs start --search $jobName --json | ConvertFrom-Json