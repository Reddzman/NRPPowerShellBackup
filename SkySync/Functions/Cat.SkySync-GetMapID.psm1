<# GetConnectionID

This function get the id for a specifc Conneciton based on name.

Sample Call: 
    Get-SkySyncMapID -mapType ("account_map" or "group_map") -mapName $poolName -restURL $baseURL -restHeader $authHeader -ct $contentType
#>
function Get-SkySyncMapID {
    param (
        [String] $mapType,
        [String] $mapName,
        [String] $restURL,
        $restHeader,
        [String] $ct
    )

    # Determine endpoint based on Type
    switch ($mapType) {
        account_map {
            $endPoint = ("{0}/account_maps" -f $restURL)
            
            # Get all account maps
            $mapsResponse = Invoke-RestMethod $endPoint -Method GET -Headers $restHeader
            $mapGUID = ""
            foreach ($map in $mapsResponse.account_maps){
                If ($map.name -eq $mapName){
                    #$map | ConvertTo-Json -Depth 10 # Output Map Details
                    $mapID = $map.ID
                }
            }
        }
        group_map {
            $endPoint = ("{0}/group_maps" -f $restURL)
            
            # Get all account maps
            $mapsResponse = Invoke-RestMethod $endPoint -Method GET -Headers $restHeader
            $mapGUID = ""
            foreach ($map in $mapsResponse.group_maps){
                If ($map.name -eq $mapName){
                    #$map | ConvertTo-Json -Depth 10 # Output Map Details
                    $mapID = $map.ID
                }
            }
        }
    }

    if([String]::IsNullOrEmpty($mapID)) {
        return "empty"
    }
    Else {
        return $mapID
    }
}