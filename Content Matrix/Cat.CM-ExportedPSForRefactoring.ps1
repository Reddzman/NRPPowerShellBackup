<# Cat.CM-ExportedPSForRefactoring.ps1

This file is used for testing the PowerShell exported from content matrix.  This will alow you to test variable updates prior
to updating/creating functions for creating migration jobs.

Dependencies
	-	Content Matrix PowerShell Snapins
#>

$jobName = "PS Test- Site To Site List No Filter"
$sourceURL = "https://records.sp.cstbrands.local/sites/hr"
$targetURL = "https://mycouchetard.sharepoint.com/sites/HR-Records"
$list = "ACC09"

    if ( $PsVersionTable.PSVersion.Major -lt 3 ) { Write-Host "Windows PowerShell Version 3.0 or later needs to be installed in order to execute Content Matrix PowerShell scripts."; exit; }
    if ( (Get-PSSnapin -Name Metalogix.System.Commands -ErrorAction SilentlyContinue) -eq $null ) { add-pssnapin Metalogix.System.Commands | out-null }
    if ( (Get-PSSnapin -Name Metalogix.SharePoint.Migration.Commands -ErrorAction SilentlyContinue) -eq $null ) { add-pssnapin Metalogix.SharePoint.Migration.Commands | out-null }
    if (Get-Command Set-MetalogixJobPrerequisites -ErrorAction SilentlyContinue){ Set-MetalogixJobPrerequisites -Value "Content Matrix Console - SharePoint Edition" }

# Set default resolvers
    Set-MetalogixDefaultResolverSetting -Name "NodeLocation" -Value "Metalogix.Explorer.PersistentConnectionsResolver, Metalogix.Explorer, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03"
    Set-MetalogixDefaultResolverSetting -Name "ResourceTableResourceTableLink" -Value "Metalogix.Core.ConfigVariables.ResourceDatabaseTableResolver, Metalogix.Core, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03"
    Set-MetalogixDefaultResolverSetting -Name "DataResolverDataRepositoryLink" -Value "Metalogix.Client.ClientDataRepositoryResolver, Metalogix.Client, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03"
    Set-MetalogixDefaultResolverSetting -Name "ResourceDatabaseTableResolverDatabase" -Value "Data Source=CKAZSCUSSQL01;Initial Catalog=MigrationMetabase;Integrated Security=True"

# Load source
    $SourceCollection = New-MetalogixSerializableObjectCollection ("
        <IXMLAbleList IXMLAbleType=`"Metalogix.Explorer.NodeCollection, Metalogix.Explorer, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`">
            <NodeCollection>
                <Location Path=`"/{0}`" DisplayUrl=`"{1}/{0}`">
                    <Connection NodeType=`"Metalogix.SharePoint.SPWeb, Metalogix.SharePoint, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" SharePointVersion=`"15.0.4833.1000`" UnderlyingAdapterType=`"WS`" ShowAllSites=`"True`" AdapterType=`"WS`" Url=`"{1}`" UserName=`"CST\CSPCP20`" SavePassword=`"True`" Password=`"MIIBrQYJKoZIhvcNAQcDoIIBnjCCAZoCAQAxggFOMIIBSgIBADAyMB4xHDAaBgNVBAMME01ldGFsb2dpeE1pZ3JhdGlvbnMCEDWAxVXwWbaZTpJ4Pi+kDScwDQYJKoZIhvcNAQEBBQAEggEAM5tnklPq7o9fugt60BRJgKXsP9UgN6RT2p4mgwmI2kuzTjMP7OiWuvuxVHqdihB5EkUNoxPpjWjBZJhGSKuPVMTEoHrYW7J/EkzNbbJsISFTa1/50KPVeP3eWmEDXfwBdf+r0QT2suDPjShwcpueUWLeUqrBSw4E0iHe+h+0u4HtL1EK1heNX7Zt9y+kTpy3iif7KE5RL9I+khAzeTwLLDGFXKMic/emzfQXAU7fUJPFVGP7VN5YPBf8zQpIHIzfTkqKgxP9XayOtwEqaeBHSTaymG5Fb8yuSQVl3icriBrFlgjmed8xaclan06JJGPPkYQWZ1p51onGnIhcyDUkjzBDBgkqhkiG9w0BBwEwFAYIKoZIhvcNAwcECFdUjEjtsxk1gCBTZn4lYIczJpxVRisjGU+r6PIFgEtkTN6THHJq9YIezQ==`" ReadOnly=`"False`" AuthenticationType=`"Metalogix.SharePoint.Adapters.Authentication.WindowsInitializer`" />
                </Location>
            </NodeCollection>
        </IXMLAbleList>" -f $list,$sourceURL)

# Load target
    $TargetCollection = New-MetalogixSerializableObjectCollection ("
        <IXMLAbleList IXMLAbleType=`"Metalogix.Explorer.NodeCollection, Metalogix.Explorer, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`">
            <NodeCollection>
                <Location Path=`"`" DisplayUrl=`"{0}`">
                    <Connection NodeType=`"Metalogix.SharePoint.SPWeb, Metalogix.SharePoint, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" SharePointVersion=`"`" UnderlyingAdapterType=`"`" ShowAllSites=`"True`" AdapterType=`"CSOM`" Url=`"{0}`" UserName=`"metalogix@MyCoucheTard.onmicrosoft.com`" SavePassword=`"True`" Password=`"MIIBpQYJKoZIhvcNAQcDoIIBljCCAZICAQAxggFOMIIBSgIBADAyMB4xHDAaBgNVBAMME01ldGFsb2dpeE1pZ3JhdGlvbnMCEDWAxVXwWbaZTpJ4Pi+kDScwDQYJKoZIhvcNAQEBBQAEggEAR4SWGLehJFUNhivsSjgA9Vs2Zn7pZWCJpj2NglPRTTTRTajaPiJYSN91+m7VkGOVdNJcPYySrl8z6fraTSAmee6kp1nKW0DyZwSR709z0YLmmPB180En98xdteviuBPkiPZko9uaqKAI4nxn6zwipN4+umXJTAyDcdm34SLwK1k5Tt9DVd5Ui7ndP2MLwDgYepGu7oYBJ+nMLeY9ftQc5sopBpM7DMii03Ati0npR5OAKw11KtNLQYer2Gy/N/QtC5pIbszoQ33KGVUWEyIcCOyJae+un/hh8JW0gk2jj3kNSgFM9RLdwPATI7pf7Z71aizsNAk3iQ2jqJqJCRYVOzA7BgkqhkiG9w0BBwEwFAYIKoZIhvcNAwcECMvpnps5RwmBgBi6e0Xps99E7c/3k7GDoR09jV/gdlS2zWk=`" ReadOnly=`"False`" AuthenticationType=`"Metalogix.SharePoint.Adapters.Authentication.Office365StandADFSInitializer`" />
                </Location>
            </NodeCollection>
        </IXMLAbleList>" -f $targetURL)

# Run the action
    foreach($Target in $TargetCollection) {
        $SourceCollection | Copy-MLSharePointList -Target $Target  -OverwriteLists -CheckModifiedDatesForLists -CopyCustomizedFormPages -CopyFormWebParts -CopyListItems -UseAzureOffice365Upload -EncryptAzureMigrationJobs -CopyContentZoneContent -MapTermStores -ClearRoleAssignments -OverwriteItems -CheckModifiedDatesForItemsDocuments -ReattachPageLayouts -CopyVersions -CopySubFolders -MaximumVersionCount "2" -PreserveItemIDs -ResolveManagedMetadataByName -CopyReferencedManagedMetadata -TermstoreNameMappingTable ( New-MetalogixSerializableObject "Metalogix.DataStructures.Generic.CommonSerializableTable``2[[System.String, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089],[System.String, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089]]" "Metalogix.Core" "
            <XmlableTable>
                <XmlableEntry>
                    <Key>d8c9d001-9f46-447d-9cf4-a68958860274</Key>
                    <Value>955e65fb-2b04-4abf-96f0-3d2f7150fab9</Value>
                </XmlableEntry>
            </XmlableTable>" ) -MigrationMode "Full" -LogSkippedItems -ForceRefresh -MapMissingUsersToLoginName "archived-sp@MyCoucheTard.onmicrosoft.com" -OverrideCheckouts -Transformers ( New-MetalogixSerializableObject "Metalogix.Transformers.TransformerCollection" "Metalogix.Actions" "
            <TransformerCollection>
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ContentTypeFieldMapper, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.DocumentSetsFolderApplicator, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ReferencedFolderDataUpdater, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.FolderColumnMapper, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ManagedMetadataItemValueMapper, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.DeletionPropagator, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ReferencedListItemDataUpdater, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.DocumentSetsApplicator, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ContentTypesApplicator, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ListItemColumnMapper, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.WorkflowDataUpdater, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.PublishingDataUpdater, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.InfopathItemContentXml, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ListColumnMapper, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.BDCUpdater, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ListGuidMapper, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.MapUsers, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
                <Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.WebPartsProcessor, Metalogix.SharePoint.Actions, Version=8.9.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
            </TransformerCollection>" -Enumerate) -agentdatabase "Data Source=CKAZSCUSSQL01;Initial Catalog=MigrationMetabase;Integrated Security=True" -jobtitle "$jobName" -runremotely -Certificate "MetalogixMigrations"
    }