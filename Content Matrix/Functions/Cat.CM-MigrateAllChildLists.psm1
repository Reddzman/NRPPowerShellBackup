<# 1.	Migrate-CMAllChildLists
	This function create a content matrix incremental job for list and libraries for a specific web.

	**** Content Matrix XML is server specific must regenerate based clients configuration ****

	Parameters
		SourceURL - URL for source web
        TargetURL - URL for target web
        JobName - Name for the metalogix job
        $ContentType - Content Type Used by migration job
        $Server - Server hosting metalogix database
        $Database - Metalogix Agent Database
        $User - Username with database permissions
        $Pass - Password for database access account		
#>
function global:Migrate-CMAllChildLists
{
	param (
		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$SourceURL,
		
		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$TargetURL,

		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [String]$JobName,
        
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$ContentType,

		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$Server,

		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$Database,
		
		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$User,

		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[String]$Pass
	)

	# Loading Metalogix Snapins
	Write-Host "Loading Metalogix Snapins"
	if ( $PsVersionTable.PSVersion.Major -lt 3 ) { Write-Host "Windows PowerShell Version 3.0 or later needs to be installed in order to execute Content Matrix PowerShell scripts."; exit; }
	if ( (Get-PSSnapin -Name Metalogix.System.Commands -ErrorAction SilentlyContinue) -eq $null ) { add-pssnapin Metalogix.System.Commands | out-null }
	if ( (Get-PSSnapin -Name Metalogix.SharePoint.Migration.Commands -ErrorAction SilentlyContinue) -eq $null ) { add-pssnapin Metalogix.SharePoint.Migration.Commands | out-null }

	# Set default resolvers
	Write-Host "Setting Metalogix default resolvers"
	Set-MetalogixDefaultResolverSetting -Name "NodeLocation" -Value "Metalogix.Explorer.PersistentConnectionsResolver, Metalogix.Explorer, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03"
	Set-MetalogixDefaultResolverSetting -Name "ResourceTableResourceTableLink" -Value "Metalogix.ResourceFileTableResolver, Metalogix.Core, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03"
	Set-MetalogixDefaultResolverSetting -Name "DataResolverDataRepositoryLink" -Value "Metalogix.Client.ClientDataRepositoryResolver, Metalogix.Client, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03"

	# Load configuration settings
	Write-Host "Loading Metalogix Configuration Settings"
	Load-MetalogixConfigurationVariableSettings -FilePath "C:\Users\npatterson\AppData\Roaming\Metalogix\UserSettings.xml" -Scope UserSpecific
	Load-MetalogixConfigurationVariableSettings -FilePath "C:\Users\npatterson\AppData\Roaming\Metalogix\Content Matrix Console - SharePoint Edition\ApplicationSettings.xml" -Scope ApplicationAndUserSpecific
	Load-MetalogixConfigurationVariableSettings -FilePath "C:\ProgramData\Metalogix\EnvironmentSettings.xml" -Scope EnvironmentSpecific

	# Load source
	Write-Host "Loading Source Web"
	$SourceCollection = New-MetalogixSerializableObjectCollection ("<IXMLAbleList IXMLAbleType=`"Metalogix.Explorer.NodeCollection, Metalogix.Explorer, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`"><NodeCollection><Location Path=`"`" DisplayUrl=`"{0}`"><Connection NodeType=`"Metalogix.SharePoint.SPWeb, Metalogix.SharePoint, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" SharePointVersion=`"15.0.4667.1000`" UnderlyingAdapterType=`"WS`" ShowAllSites=`"True`" AdapterType=`"WS`" Url=`"{0}`" UserName=`"neo\npatterson`" SavePassword=`"True`" Password=`"AQAAANCMnd8BFdERjHoAwE/Cl+sBAAAAj9VIxJ6vl0Gs7ejqbEsHIAQAAAACAAAAAAADZgAAwAAAABAAAADhCRUV0V9VQAZCK9ptOF6CAAAAAASAAACgAAAAEAAAAKo58Ud1nc0Y1diajQ9z3tcYAAAA44qk2nBvrD9UBa3K+11cy4BMAz5mDUXhFAAAAMoiqBwYZEWKVpZCbI58ApERxIkj`" ReadOnly=`"False`" AuthenticationType=`"Metalogix.SharePoint.Adapters.Authentication.WindowsInitializer`" /></Location></NodeCollection></IXMLAbleList>" -f $SourceURL)

	# Load target
	Write-Host "Loading Target Web"
	$TargetCollection = New-MetalogixSerializableObjectCollection ("<IXMLAbleList IXMLAbleType=`"Metalogix.Explorer.NodeCollection, Metalogix.Explorer, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`"><NodeCollection><Location Path=`"`" DisplayUrl=`"{0}`"><Connection NodeType=`"Metalogix.SharePoint.SPWeb, Metalogix.SharePoint, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" SharePointVersion=`"`" UnderlyingAdapterType=`"`" ShowAllSites=`"True`" AdapterType=`"CSOM`" Url=`"{0}`" UserName=`"metalogix@neomaterials.onmicrosoft.com`" SavePassword=`"True`" Password=`"AQAAANCMnd8BFdERjHoAwE/Cl+sBAAAAj9VIxJ6vl0Gs7ejqbEsHIAQAAAACAAAAAAADZgAAwAAAABAAAABVzYELiXlAP1HVNgupjX0UAAAAAASAAACgAAAAEAAAAPijNljyDhJuxN9r8EIAOT0gAAAAH1GMy/upxCa30ry8TSEc1QNtQlBnhbX4CroNengeCvYUAAAAUDbdovbHaxTN93HoaMrZYK51br4=`" ReadOnly=`"False`" AuthenticationType=`"Metalogix.SharePoint.Adapters.Authentication.Office365StandADFSInitializer`" /></Location></NodeCollection></IXMLAbleList>" -f $TargetURL)

	# Run the action
	Write-Host "Running All Child Lists Migration"
	foreach($Target in $TargetCollection)
	{
		$SourceCollection | Copy-MLSharePointAllChildLists -Target $Target  -CopyListPermissions -UpdateLists "All" -CheckModifiedDatesForLists -CopyCustomizedFormPages -CopyViewWebParts -CopyFormWebParts -CopyFolderPermissions -CopyListItems -UseAzureOffice365Upload -EncryptAzureMigrationJobs -CopyDocumentWebParts -CopyContentZoneContent -ExistingWebPartsAction "Delete" -CopyRootPermissions -MapRolesByName -ClearRoleAssignments -CopyItemPermissions -CheckModifiedDatesForItemsDocuments -UpdateItems "All" -ReattachPageLayouts -CopyVersions -CopySubFolders -PreserveItemIDs -PreserveSharePointDocumentIDs -ContentTypeApplicationObjects ( New-MetalogixSerializableObject "Metalogix.DataStructures.Generic.CommonSerializableList``1[[Metalogix.SharePoint.Options.Migration.ContentTypeApplicationOptionsCollection, Metalogix.SharePoint.Options, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03]]" "Metalogix.Core" ("<XmlableSet>
			<XmlableItem Type=`"Metalogix.SharePoint.Options.Migration.ContentTypeApplicationOptionsCollection, Metalogix.SharePoint.Options, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`">
				<ContentTypeApplicationOptionsCollection Type=`"Metalogix.SharePoint.Options.Migration.ContentTypeApplicationOptionsCollection, Metalogix.SharePoint.Options, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" AddUnmappedItemsToDefaultContentType=`"True`" RemoveAllOtherContentTypes=`"False`">
					<AppliesToFilter>
						<FilterExpression>
							<Contains Property=`"Url`" CaseSensitive=`"True`" BaseFilter=`"False`" Pattern=`"neomaterials`" CultureInfoName=`"en-US`">
								<AppliesToTypes>
									<Type>Metalogix.SharePoint.SPList</Type>
								</AppliesToTypes>
							</Contains>
						</FilterExpression>
					</AppliesToFilter>
					<InternalData>
						<XmlableSet>
							<XmlableItem Type=`"Metalogix.SharePoint.Options.Migration.ContentTypeApplicationOptions, Metalogix.SharePoint.Options, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`">
								<ContentTypeApplicationOptions ContentTypeName=`"{0}`" MakeDefault=`"True`">
									<MapItemsFilter>
										<And IsImplicitGroup=`"True`">
											<FilterExpression>
												<NotEquals Property=`"FSObjType`" CaseSensitive=`"True`" BaseFilter=`"False`" Pattern=`"1`">
													<AppliesToTypes>
														<Type>Metalogix.SharePoint.SPListItem</Type>
													</AppliesToTypes>
												</NotEquals>
											</FilterExpression>
										</And>
									</MapItemsFilter>
								</ContentTypeApplicationOptions>
							</XmlableItem>
						</XmlableSet>
					</InternalData>
				</ContentTypeApplicationOptionsCollection>
			</XmlableItem>
		</XmlableSet>" -f $ContentType) -Enumerate) -MigrationMode "Incremental" -LogSkippedItems -ForceRefresh -CorrectingLinks -LinkCorrectionScope "SiteCollection" -MapGroupsByName -OverwriteGroups -MapMissingUsersToLoginName "archived@neomaterials.onmicrosoft.com" -OverrideCheckouts -Transformers ( New-MetalogixSerializableObject "Metalogix.Transformers.TransformerCollection" "Metalogix.Actions" "<TransformerCollection>
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ContentTypeFieldMapper, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.DocumentSetsFolderApplicator, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ReferencedFolderDataUpdater, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.FolderColumnMapper, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ManagedMetadataItemValueMapper, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.DeletionPropagator, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ReferencedListItemDataUpdater, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.DocumentSetsApplicator, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ContentTypesApplicator, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ListItemColumnMapper, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.WorkflowDataUpdater, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.PublishingDataUpdater, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.InfopathItemContentXml, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ListColumnMapper, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.BDCUpdater, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.ListGuidMapper, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.MapUsers, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
			<Transformer TransformerType=`"Metalogix.SharePoint.Actions.Transform.WebPartsProcessor, Metalogix.SharePoint.Actions, Version=8.6.0.3, Culture=neutral, PublicKeyToken=3b240fac3e39fc03`" ReadOnly=`"True`" />
		</TransformerCollection>" -Enumerate) -jobdatabase "Data Source=SharePointAlias;Initial Catalog=SPOMig_JobDB;Integrated Security=True"
	}
	
	#region 7.	Renaming Metalogix Job
	Write-Host "Renaming Metalogix Migration Job"
	$SQLIDQuery = ("SELECT TOP 1 JobID 	FROM [{0}].[dbo].[JobHistory] WHERE [Target]='{1}' ORDER BY [Created] DESC" -f $Database, $TargetURL)
	$JobID = Invoke-Sqlcmd -ServerInstance $Server -Query $SQLIDQuery -Username $User -Password $Pass
	$SQLUpdateQuery = ("UPDATE [{0}].[dbo].[JobHistory] SET [Title]='{1}' WHERE [JobID]='{2}'" -f $Database, $JobName, $JobID.JobID)
	$SQLUpdateResult = Invoke-Sqlcmd -ServerInstance $Server -Query $SQLUpdateQuery -Username $User -Password $Pass
	#endregion 5.	Renaming Metalogix Job
} # Migrate-CMAllChildLists