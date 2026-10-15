<# Create-CMListItemJob 

	This function create a content matrix paste list items job for a list or library.

	**** Content Matrix XML is server specific must regenerate based clients configuration ****

	Parameters
		jobName - Name to be used in Content Matrix
		sourceURL - URL for source list
		targetURL - URL for target list
		list - URL path to list "listpath" not "/listpath"
		$filter - dynamic xml filter based on PS export from Content Matrix
	
	Sample: Create-CMListItemsJob -jobName $curJob -sourceURL $job.sURL -targetURL $job.dURL -list $job.listPath -filter $curFilter
	
#>
function global:New-CMListItemsJob
{
	param (
		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$jobName,
		
		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$sourceURL,

		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$targetURL,
		
		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [String]$list,

		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$filter
	)

	# Loading PowerShell Modules
	Write-Host "         Loading PowerShell Modules"
	if ( $PsVersionTable.PSVersion.Major -lt 3 ) { Write-Host "Windows PowerShell Version 3.0 or later needs to be installed in order to execute Content Matrix PowerShell scripts."; exit; }
    if ( (Get-PSSnapin -Name Metalogix.System.Commands -ErrorAction SilentlyContinue) -eq $null ) { add-pssnapin Metalogix.System.Commands | out-null }
    if ( (Get-PSSnapin -Name Metalogix.SharePoint.Migration.Commands -ErrorAction SilentlyContinue) -eq $null ) { add-pssnapin Metalogix.SharePoint.Migration.Commands | out-null }
    if (Get-Command Set-MetalogixJobPrerequisites -ErrorAction SilentlyContinue){ Set-MetalogixJobPrerequisites -Value "Content Matrix Console - SharePoint Edition" }

	# Set default resolvers
	Write-Host "         Loading Default Resolvers"
	Set-MetalogixDefaultResolverSetting -Name "NodeLocation" -Value "Metalogix.Explorer.PersistentConnectionsResolver, Metalogix.Explorer, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03"
    Set-MetalogixDefaultResolverSetting -Name "ResourceTableResourceTableLink" -Value "Metalogix.Core.ConfigVariables.ResourceDatabaseTableResolver, Metalogix.Core, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03"
    Set-MetalogixDefaultResolverSetting -Name "DataResolverDataRepositoryLink" -Value "Metalogix.Client.ClientDataRepositoryResolver, Metalogix.Client, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03"
    Set-MetalogixDefaultResolverSetting -Name "ResourceDatabaseTableResolverDatabase" -Value "Data Source=CKAZSCUSSQL01;Initial Catalog=MigrationMetabase;Integrated Security=True"

	# Load source
	Write-Host "         Loading Source Collection"
	$SourceCollection = New-MetalogixSerializableObjectCollection ("
        <IXMLAbleList IXMLAbleType=`"Metalogix.Explorer.NodeCollection, Metalogix.Explorer, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`">
            <NodeCollection>
                <Location Path=`"/{0}`" DisplayUrl=`"{1}/{0}`">
                    <Connection NodeType=`"Metalogix.SharePoint.SPWeb, Metalogix.SharePoint, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" SharePointVersion=`"15.0.4833.1000`" UnderlyingAdapterType=`"WS`" ShowAllSites=`"True`" AdapterType=`"WS`" Url=`"{1}`" UserName=`"CST\CSPCP20`" SavePassword=`"True`" Password=`"MIIBrQYJKoZIhvcNAQcDoIIBnjCCAZoCAQAxggFOMIIBSgIBADAyMB4xHDAaBgNVBAMME01ldGFsb2dpeE1pZ3JhdGlvbnMCEDWAxVXwWbaZTpJ4Pi+kDScwDQYJKoZIhvcNAQEBBQAEggEATjvDB76eLB11EzzQWld9yHoTzCE5iFvdjHn7ljDwclKbvEGieMcSzI6Ueke07zOdb6QjQkTbDB8VDTWFyz9OI5la+E4Rh8mPYhWLjItbCQAsQG4dVpw9JsEvPRl843+SS5BBOgtLtospjibAbXe1wbVE60ew6LMgNsbLm6fYzeVYdvXJG0M8wWU+pR3bVKU4GnCs4hmJUXu9/PAFvUgEbtx8HZSNw4TCWDMwQC7lzsSSDh31PoCYEKJGiFJIM6kE8ePRWbswMMt5HMPTnwevhH1jXMkzmErVmO4i6/Z7speJKcvCtfIX04no1Dx/q+gJb4veCdN1Mlep1qEnEkhOqDBDBgkqhkiG9w0BBwEwFAYIKoZIhvcNAwcECEt2dtnwD0q7gCAo1t7eUvER9+fXqKeoWT6XIMPxrlFfmR22ZblpGyyALQ==`" ReadOnly=`"False`" AuthenticationType=`"Metalogix.SharePoint.Adapters.Authentication.WindowsInitializer`" />
                </Location>
            </NodeCollection>
        </IXMLAbleList>" -f $list,$sourceURL)

	# Load target
	Write-Host "         Loading Target Collection"
	$TargetCollection = New-MetalogixSerializableObjectCollection ("
        <IXMLAbleList IXMLAbleType=`"Metalogix.Explorer.NodeCollection, Metalogix.Explorer, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`">
            <NodeCollection>
                <Location Path=`"/{0}2`" DisplayUrl=`"{1}/{0}2`">
                    <Connection NodeType=`"Metalogix.SharePoint.SPWeb, Metalogix.SharePoint, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" SharePointVersion=`"2147483647.0.0.0`" UnderlyingAdapterType=`"CSOM`" ShowAllSites=`"True`" AdapterType=`"CSOM`" Url=`"{1}`" UserName=`"metalogix@MyCoucheTard.onmicrosoft.com`" SavePassword=`"True`" Password=`"MIIBpQYJKoZIhvcNAQcDoIIBljCCAZICAQAxggFOMIIBSgIBADAyMB4xHDAaBgNVBAMME01ldGFsb2dpeE1pZ3JhdGlvbnMCEDWAxVXwWbaZTpJ4Pi+kDScwDQYJKoZIhvcNAQEBBQAEggEAMKPJwwJGhyIS4zUETdquOyXr8GpGMIEt5xv/UEjPixtFbUOIpOkSZOBG9ilh1xl/fzyq1WVYP0twTp8y3tKTu4I5JBzHWKkVp1w+JyHd91t7Ymt6py09Kf8khLSuKxFc9EHmdGr1uNTRL6YlzRWhnZFyyqYb8sX3V8XDMIg10MmTkR3/UWFwyBdW2EZEKhCmbVLIXrTqR23Mj2BaTr6yK8VJaDdiFivSlAbd7yfXlzWxRz1eix6XM3NoEt+tIBxEv4gSq9TJXp07WBMYf8O60b5naJsaPCjnd5eExDgOW10qjpZ9Km+stpiIB449avXzclavla4sRy0s7aB7rJbFhTA7BgkqhkiG9w0BBwEwFAYIKoZIhvcNAwcECHiOIJhW3sIDgBiqJ2bI0sqNKslD3arG93lP76i0wfupPlg=`" ReadOnly=`"False`" AuthenticationType=`"Metalogix.SharePoint.Adapters.Authentication.Office365StandADFSInitializer`" />
                </Location>
            </NodeCollection>
        </IXMLAbleList>" -f $list,$targetURL)

	# Run the action
	Write-Host "         Creating Migraiton Jobs"
	foreach($Target in $TargetCollection) {
        $SourceCollection | Copy-MLAllListItems -Target $Target  -OverwriteFolders -CopyListItems -UseAzureOffice365Upload -EncryptAzureMigrationJobs -CopyContentZoneContent -ItemFilterExpression ( New-MetalogixSerializableObject "Metalogix.Data.Filters.IFilterExpression" "Metalogix.Actions" "$filter" ) -MapTermStores -ClearRoleAssignments -OverwriteItems -CheckModifiedDatesForItemsDocuments -ReattachPageLayouts -CopyVersions -CopySubFolders -MaximumVersionCount "2" -PreserveItemIDs -ResolveManagedMetadataByName -CopyReferencedManagedMetadata -TermstoreNameMappingTable ( New-MetalogixSerializableObject "Metalogix.DataStructures.Generic.CommonSerializableTable``2[[System.String, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089],[System.String, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089]]" "Metalogix.Core" "
            <XmlableTable>
                <XmlableEntry>
                    <Key>d8c9d001-9f46-447d-9cf4-a68958860274</Key>
                    <Value>955e65fb-2b04-4abf-96f0-3d2f7150fab9</Value>
                </XmlableEntry>
            </XmlableTable>" ) -MigrationMode "Full" -LogSkippedItems -ForceRefresh -MapMissingUsersToLoginName "archived-sp@MyCoucheTard.onmicrosoft.com" -OverrideCheckouts -Transformers ( New-MetalogixSerializableObject "Metalogix.Transformers.TransformerCollection" "Metalogix.Actions" "
            <TransformerCollection>
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ContentTypeFieldMapper, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ManagedMetadataItemValueMapper, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.DeletionPropagator, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ReferencedListItemDataUpdater, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.DocumentSetsApplicator, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ContentTypesApplicator, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ListItemColumnMapper, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.WorkflowDataUpdater, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.PublishingDataUpdater, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.InfopathItemContentXml, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.MapUsers, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.WebPartsProcessor, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                </TransformerCollection>" -Enumerate) -agentdatabase "Data Source=CKAZSCUSSQL01;Initial Catalog=MigrationMetabase;Integrated Security=True" -jobtitle "$jobName" -runremotely -Certificate "MetalogixMigrations"
    }		
} # Migrate-CMAllChildLists