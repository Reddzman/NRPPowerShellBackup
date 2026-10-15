-- K2 provided discovery script for Solution Configuration. Before using, ensure that there is not a new version.  The intent is to plan
-- a migration to K2 Cloud.  

-- Get DB, Schema and Table info from the SQL Server
IF OBJECT_ID(N'tempdb..#TableList', 'U') IS NOT NULL
    DROP TABLE [#TableList]

CREATE TABLE #TableList(
      DatabaseName	SYSNAME NOT NULL
    , SchemaName	SYSNAME NOT NULL
    , TableName		SYSNAME NOT NULL
    )

	INSERT INTO #TableList
	SELECT
				DB_NAME() AS DatabaseName
			, OBJECT_SCHEMA_NAME(object_id) AS SchemaName
			, name AS TableName
	FROM	sys.tables
	WHERE	OBJECT_SCHEMA_NAME(object_id) IN ('HostServer','Server','ServerLog','SmartBroker','Form')
			AND
			[name] IN ('Configuration','SecurityCredentialCache','HostServerSetting','Server','ProcInst','ProcSet','ActInst','EventInst','Proc','SmartObject','Form','View')
	ORDER BY [DatabaseName], [SchemaName], [TableName];

-- Get relevant DB names from HostServer DB/Schema - used for Legacy DBs queries.
DECLARE @CurrentDBName NVARCHAR(128)
DECLARE @HSDBName NVARCHAR(128)
DECLARE @K2DBName NVARCHAR(128)
DECLARE @K2LogDBName NVARCHAR(128)
DECLARE @SmODBName NVARCHAR(128)

SET @CurrentDBName = DB_NAME()
PRINT 'Running script on: ''[' + @CurrentDBName + ']'' database'

IF EXISTS(SELECT * FROM sys.sysdatabases WHERE name=@CurrentDBName)
BEGIN
	PRINT 'Database Exists'

	IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = N'HostServer')
	BEGIN
		PRINT '''[HostServer]'' Schema Exists'
		
		IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'Configuration')
		BEGIN
			PRINT '''[HostServer].[Configuration]'' Table Exists'
			
			-- Get the necessary DBNames
			SELECT	@HSDBName=[VariableValue] FROM [HostServer].[Configuration] WHERE [VariableToken]='[HOSTSERVERDBNAME]'
			IF (@HSDBName IS NULL) SELECT @HSDBName=DB_NAME()

			SELECT	@K2DBName=[VariableValue] FROM [HostServer].[Configuration] WHERE [VariableToken]='[K2DBNAME]'
			IF (@K2DBName IS NULL) SELECT @K2DBName=DB_NAME()

			SELECT	@K2LogDBName=[VariableValue] FROM [HostServer].[Configuration] WHERE [VariableToken]='[K2LOGDBNAME]'
			IF (@K2LogDBName IS NULL) SELECT @K2LogDBName=DB_NAME()

			SELECT	@SmODBName=[VariableValue] FROM [HostServer].[Configuration] WHERE [VariableToken]='[SMARTBROKERDBNAME]'
			IF (@SmODBName IS NULL) SELECT @SmODBName=DB_NAME()

		END
		ELSE
		BEGIN
			PRINT '''[HostServer].[Configuration]'' Table cannot be found on DB: [' + @CurrentDBName + '].'
			SELECT	 '''[HostServer].[Configuration]'' Table cannot be found on DB: [' + @CurrentDBName + '].' AS [DatabaseName]
					,'''[HostServer].[Configuration]'' Table cannot be found on DB: [' + @CurrentDBName + '].' AS [SchemaName]
					,'''[HostServer].[Configuration]'' Table cannot be found on DB: [' + @CurrentDBName + '].' AS [TableName]
			UNION ALL
			SELECT * FROM #TableList
			RETURN
		END
	END
	ELSE
	BEGIN
		PRINT '''[HostServer]'' Schema cannot be found on DB: [' + @CurrentDBName + '].'
		SELECT	 '''[HostServer]'' Schema cannot be found on DB: [' + @CurrentDBName + '].' AS [DatabaseName]
				,'''[HostServer]'' Schema cannot be found on DB: [' + @CurrentDBName + '].' AS [SchemaName]
				,'''[HostServer]'' Schema cannot be found on DB: [' + @CurrentDBName + '].' AS [TableName]
		UNION ALL
		SELECT * FROM #TableList
		RETURN
	END
END
ELSE
BEGIN
	PRINT 'Database: ''[' + @CurrentDBName + ']'' cannot be found.'
	SELECT	 'Database: ''[' + @CurrentDBName + ']'' cannot be found.' AS [DatabaseName]
			,'Database: ''[' + @CurrentDBName + ']'' cannot be found.' AS [SchemaName]
			,'Database: ''[' + @CurrentDBName + ']'' cannot be found.' AS [TableName]
	UNION ALL
	SELECT * FROM #TableList
	RETURN
END

-- Variables - Global
DECLARE @SQL NVARCHAR(MAX)
-- Variables - Configuration related
DECLARE @Product NVARCHAR(100)
DECLARE @LicType NVARCHAR(100)
DECLARE @LicKey NVARCHAR(100)
DECLARE @MachKey NVARCHAR(100)
DECLARE @Processors NVARCHAR(100)
DECLARE @LicUsers NVARCHAR(100)
DECLARE @LicExpire NVARCHAR(100)
DECLARE @ProdVersion NVARCHAR(100)
DECLARE @BuildNum NVARCHAR(100)
DECLARE @SQLVersion NVARCHAR(100)
DECLARE @DBVersion NVARCHAR(100)
DECLARE @DBStruct NVARCHAR(100)
DECLARE @ExchVersion NVARCHAR(100)
DECLARE @PTAOption NVARCHAR(100)
DECLARE @LastUpdate NVARCHAR(100)
DECLARE @IsNLB NVARCHAR(100)
-- Variables - Usage
DECLARE @LicenseCountAllTime INT 
DECLARE @LicenseCountLast7Days INT 
DECLARE @LicenseCountLast14Days INT 
DECLARE @LicenseCountLast30Days INT
DECLARE @LicenseCount1year INT
DECLARE @ActiveServers NVARCHAR(100)
DECLARE @PIActive INT
DECLARE @PICompleted INT
DECLARE @WFs INT
DECLARE @SmOs INT
DECLARE @Forms INT
DECLARE @Views INT

-- Populate Variables
IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @HSDBName AND [SchemaName] = N'HostServer' AND [TableName] = N'Configuration' )
BEGIN
	SELECT	@Product=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[LICENSEDPRODUCT]';	IF (@Product IS NULL) SET @Product='UNDEFINED'
	SELECT	@LicType=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[LICENSETYPE]';		IF (@LicType IS NULL) SET @LicType='UNDEFINED'
	SELECT	@LicKey=[VariableValue]			FROM [HostServer].[Configuration] WHERE [VariableToken]='[LICENSEKEY]';			IF (@LicKey IS NULL) SET @LicKey='UNDEFINED'
	SELECT	@MachKey=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[MACHINEKEY]';			IF (@MachKey IS NULL) SET @MachKey='UNDEFINED'
	SELECT	@Processors=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[PROCESSORS]';			IF (@Processors IS NULL) SET @Processors='UNDEFINED'
	SELECT	@LicUsers=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[USERS]';				IF (@LicUsers IS NULL) SET @LicUsers='UNDEFINED'
	SELECT	@LicExpire=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[EXPIREDATE]';			IF (@LicExpire IS NULL) SET @LicExpire='UNDEFINED'
	SELECT	@ProdVersion=[VariableValue]	FROM [HostServer].[Configuration] WHERE [VariableToken]='[PRODUCTVERSION]';		IF (@ProdVersion IS NULL) SET @ProdVersion='UNDEFINED'
	SELECT	@BuildNum=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[BUILDNUMBER]';		IF (@BuildNum IS NULL) SET @BuildNum='UNDEFINED'
	SELECT	@SQLVersion=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[SQLSERVER_VERSION]';	IF (@SQLVersion IS NULL) SELECT @SQLVersion=@@VERSION 
	SELECT	@DBVersion=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[CURRENTDBVERSION]';	IF (@DBVersion IS NULL) SET @DBVersion='UNDEFINED'
	SELECT	@DBStruct=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[DATABASESTRUCTURE]';	IF (@DBStruct IS NULL) SET @DBStruct='UNDEFINED'
	SELECT	@ExchVersion=[VariableValue]	FROM [HostServer].[Configuration] WHERE [VariableToken]='[ExchangeVersion]';	IF (@ExchVersion IS NULL) SET @ExchVersion='UNDEFINED'
	SELECT	@PTAOption=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[PTA_OPTION]';			IF (@PTAOption IS NULL) SET @PTAOption='UNDEFINED'
	SELECT	@LastUpdate=[VariableValue]		FROM [HostServer].[Configuration] WHERE [VariableToken]='[LASTUPDATE]';			IF (@LastUpdate IS NULL) SET @LastUpdate='UNDEFINED'
	SELECT	@IsNLB=[VariableValue]			FROM [HostServer].[Configuration] WHERE [VariableToken]='[ISNLB]';				IF (@IsNLB IS NULL) SET @IsNLB='UNDEFINED'
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @HSDBName AND [SchemaName] = N'HostServer' AND [TableName] = N'SecurityCredentialCache' )
BEGIN
	DECLARE @LastAccessDate DATE = SYSUTCDATETIME()
	DECLARE @TodayCounter INT = (SELECT COUNT(DISTINCT scc.[UserName]) FROM [HostServer].[LicensedUser] AS lu INNER JOIN [HostServer].[SecurityCredentialCache] AS scc ON lu.[CredentialID] = scc.[CredentialID] WHERE  CONVERT(DATE,scc.[LastAccessDateUtc]) = CONVERT(DATE,SYSUTCDATETIME()))

	IF (@TodayCounter = 0)
	BEGIN
		   SET @LastAccessDate = DATEADD(DAY, -1, SYSUTCDATETIME())
		   SET @TodayCounter = (SELECT COUNT(DISTINCT scc.[UserName]) FROM [HostServer].[LicensedUser] AS lu INNER JOIN [HostServer].[SecurityCredentialCache] AS scc ON lu.[CredentialID] = scc.[CredentialID] WHERE (scc.[UserName] not like 'NT AUTHORITY%' and scc.[UserName] not like '%SYSTEM%' and scc.[UserName] not like '%$%') AND CONVERT(DATE,scc.[LastAccessDateUtc]) = CONVERT(DATE,@LastAccessDate))
	END;

		   SET @LicenseCountAllTime = (SELECT COUNT(DISTINCT scc.[UserName]) FROM [HostServer].[LicensedUser] AS lu INNER JOIN [HostServer].[SecurityCredentialCache] AS scc ON lu.[CredentialID] = scc.[CredentialID] WHERE scc.[UserName] not like 'NT AUTHORITY%' and scc.[UserName] not like '%SYSTEM%' and scc.[UserName] not like '%$%');
		   SET @LicenseCountLast7Days = (SELECT COUNT(DISTINCT scc.[UserName]) FROM [HostServer].[LicensedUser] AS lu INNER JOIN [HostServer].[SecurityCredentialCache] AS scc ON lu.[CredentialID] = scc.[CredentialID] WHERE (scc.[UserName] not like 'NT AUTHORITY%' and scc.[UserName] not like '%SYSTEM%' and scc.[UserName] not like '%$%') AND [LastAccessDateUtc] >= DATEADD(DAY, -7, @LastAccessDate))
		   SET @LicenseCountLast14Days = (SELECT COUNT(DISTINCT scc.[UserName]) FROM [HostServer].[LicensedUser] AS lu INNER JOIN [HostServer].[SecurityCredentialCache] AS scc ON lu.[CredentialID] = scc.[CredentialID] WHERE (scc.[UserName] not like 'NT AUTHORITY%' and scc.[UserName] not like '%SYSTEM%' and scc.[UserName] not like '%$%') AND [LastAccessDateUtc] >= DATEADD(DAY, -14, @LastAccessDate))
		   SET @LicenseCountLast30Days = (SELECT COUNT(DISTINCT scc.[UserName]) FROM [HostServer].[LicensedUser] AS lu INNER JOIN [HostServer].[SecurityCredentialCache] AS scc ON lu.[CredentialID] = scc.[CredentialID] WHERE (scc.[UserName] not like 'NT AUTHORITY%' and scc.[UserName] not like '%SYSTEM%' and scc.[UserName] not like '%$%') AND [LastAccessDateUtc] >= DATEADD(DAY, -30, @LastAccessDate))
		   SET @LicenseCount1year = (SELECT COUNT(DISTINCT scc.[UserName]) FROM [HostServer].[LicensedUser] AS lu INNER JOIN [HostServer].[SecurityCredentialCache] AS scc ON lu.[CredentialID] = scc.[CredentialID] WHERE (scc.[UserName] not like 'NT AUTHORITY%' and scc.[UserName] not like '%SYSTEM%' and scc.[UserName] not like '%$%') AND [LastAccessDateUtc] >= DATEADD(YEAR, -1, @LastAccessDate))

	IF (@TodayCounter = 0)
	BEGIN
		   SET @LastAccessDate = (SELECT MAX([LastAccessDateUtc]) FROM [HostServer].[SecurityCredentialCache] WHERE [UserName] LIKE '%@%')
	END
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @K2DBName AND [SchemaName] = N'Server' AND [TableName] = N'Server' )
BEGIN
	SET @SQL='SELECT @ActiveServers=COUNT(1) FROM [' + @K2DBName + '].[Server].[Server] WHERE [Running]=1'
	EXECUTE sp_executesql @SQL , N'@ActiveServers int OUTPUT', @ActiveServers = @ActiveServers output;
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @K2DBName AND [SchemaName] = N'Server' AND [TableName] = N'ProcInst' )
BEGIN
	SET @SQL='SELECT @PIActive=COUNT(1) FROM [' + @K2DBName + '].[Server].[ProcInst] WHERE [Status]<>5 --(Deleted)'
	EXECUTE sp_executesql @SQL , N'@PIActive int OUTPUT', @PIActive = @PIActive output;
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @K2LogDBName AND [SchemaName] = N'ServerLog' AND [TableName] = N'ProcInst' )
BEGIN
	SET @SQL='SELECT @PICompleted=COUNT(1) FROM [' + @K2LogDBName + '].[ServerLog].[ProcInst] WHERE [Status]=3 --(Completed)'
	EXECUTE sp_executesql @SQL , N'@PICompleted int OUTPUT', @PICompleted = @PICompleted output;
END

IF EXISTS ( SELECT * FROM #TableList WHERE ([DatabaseName] = @K2DBName AND [SchemaName] = N'Server' AND [TableName] = N'ProcSet') )
BEGIN
	SET @SQL='SELECT @WFs=COUNT(1) FROM [' + @K2DBName + '].[Server].[ProcSet] PS JOIN [' + @K2DBName + '].[Server].[Proc] P ON P.[ProcSetID] = PS.[ID] WHERE P.[Ver]=1'
	EXECUTE sp_executesql @SQL , N'@WFs int OUTPUT', @WFs = @WFs output;
END

IF EXISTS ( SELECT * FROM #TableList WHERE ([DatabaseName] = @SmODBName AND [SchemaName] = N'SmartBroker' AND [TableName] = N'SmartObject') )
BEGIN
	SET @SQL='SELECT @SmOs=COUNT(DISTINCT [Name]) FROM [' + @SmODBName + '].[SmartBroker].[SmartObject] WHERE [CreateDate] IS NOT NULL'
	EXECUTE sp_executesql @SQL , N'@SmOs int OUTPUT', @SmOs = @SmOs output;
END

IF EXISTS ( SELECT * FROM #TableList WHERE ([DatabaseName] = @HSDBName AND [SchemaName] = N'Form' AND [TableName] = N'Form') )
BEGIN
	SET @SQL='SELECT @Forms=COUNT(DISTINCT [Name]) FROM [' + @HSDBName + '].[Form].[Form]'
	EXECUTE sp_executesql @SQL , N'@Forms int OUTPUT', @Forms = @Forms output;
END

IF EXISTS ( SELECT * FROM #TableList WHERE ([DatabaseName] = @HSDBName AND [SchemaName] = N'Form' AND [TableName] = N'View') )
BEGIN
	SET @SQL='SELECT @Views=COUNT(DISTINCT [Name]) FROM [' + @HSDBName + '].[Form].[View]'
	EXECUTE sp_executesql @SQL , N'@Views int OUTPUT', @Views = @Views output;
END

-- Store Summary information
IF (Object_ID('tempdb..#Summary') IS NOT NULL)
	DROP TABLE [#Summary]

CREATE TABLE #Summary([Token] NVARCHAR(250), [Value] NVARCHAR(1024))

INSERT INTO #Summary ([Token], [Value]) 
	VALUES	 ('[Product]', @Product)
			,('[LicenseType]', @LicType)
			,('[LicenseKey]', @LicKey)
			,('[MachineKey]', @MachKey)
			,('[Processors]', @Processors)
			,('[LicensedUsers]', @LicUsers)
			,('[LicenseExpiry]', @LastUpdate)
			,('[ProductVersion]', @ProdVersion)
			,('[BuildNumber]', @BuildNum)
			,('[SQLVersion]', @SQLVersion)
			,('[DBVersion]', @DBVersion)
			,('[DBStructure]', @DBStruct)
			,('[ExchangeVersion]', @ExchVersion)
			,('[PTA_Option]', @PTAOption)
			,('[LastUpdate]', @LastUpdate)
			,('[IsNLB]', @IsNLB)
			,('[LicenseCountAllTime]', CONVERT(NVARCHAR(100), @LicenseCountAllTime))
			,('[LicenseCountLast7Days]', CONVERT(NVARCHAR(100), @LicenseCountLast7Days))
			,('[LicenseCountLast14Days]', CONVERT(NVARCHAR(100), @LicenseCountLast14Days))
			,('[LicenseCountLast30Days]', CONVERT(NVARCHAR(100), @LicenseCountLast30Days))
			,('[LicenseCountLast12Months]', CONVERT(NVARCHAR(100), @LicenseCount1year))
			,('[ActiveServers]', CONVERT(NVARCHAR(100), @ActiveServers))
			,('[WFInstsActive]', CONVERT(NVARCHAR(100), @PIActive))
			,('[WFInstsCompleted]', CONVERT(NVARCHAR(100), @PICompleted))
			,('[Total#WFDefs]', CONVERT(NVARCHAR(100), @WFs))
			,('[Total#SmOs]', CONVERT(NVARCHAR(100), @SmOs))
			,('[Total#Forms]', CONVERT(NVARCHAR(100), @Forms))
			,('[Total#Views]', CONVERT(NVARCHAR(100), @Views))

IF EXISTS ( SELECT * FROM #TableList WHERE ([DatabaseName] = @HSDBName AND [SchemaName] = N'HostServer' AND [TableName] = N'HostServerSetting') )
BEGIN
	SET @SQL = '
		INSERT INTO #Summary ([Token], [Value])
		SELECT	 ''[UpgradeHistory - '' + [Name] + '']'' AS [Token]
				,[Value]
		FROM	[' + @HSDBName + '].[HostServer].[HostServerSetting]'

	--PRINT @SQL
	EXEC SP_EXECUTESQL @SQL
END

IF EXISTS ( SELECT * FROM #TableList WHERE ([DatabaseName] = @K2DBName AND [SchemaName] = N'Server' AND [TableName] = N'ProcSet') )
BEGIN
	SET @SQL = '
		INSERT INTO #Summary ([Token], [Value])
		SELECT	 ''[WFName_'' + RIGHT(''0000'' + CONVERT(NVARCHAR(4), [PS].[ID]), 4) + '']'' AS [Token]
				,[PS].[FullName] AS [Value]
		FROM	[' + @K2DBName + '].[Server].[ProcSet] [PS]
		ORDER BY [PS].[ID]'

	--PRINT @SQL
	EXEC SP_EXECUTESQL @SQL
END

-- Display the Summary information.
SELECT	* FROM #Summary

-- Cleanup
IF (Object_ID('tempdb..#Summary') IS NOT NULL)
	DROP TABLE [#Summary]

IF OBJECT_ID(N'tempdb..#TableList', 'U') IS NOT NULL
    DROP TABLE [#TableList]