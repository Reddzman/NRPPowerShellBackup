-- K2 provided discovery script for Solution Utilization. Before using, ensure that there is not a new version.  The intent is to plan
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

-- Variable used for dynamic queries
DECLARE @Sql NVARCHAR(MAX)

-- Make sure no temp tables exists
IF (Object_ID('tempdb..#Usage') IS NOT NULL)
	DROP TABLE [#Usage]

IF (Object_ID('tempdb..#Months') IS NOT NULL)
	DROP TABLE [#Months]

IF (Object_ID('tempdb..#CumUsage') IS NOT NULL)
	DROP TABLE [#CumUsage]

-- Create new temp table
CREATE TABLE #Usage (
	[Order] INT,
	[Type] NVARCHAR(100),
	[ID] NVARCHAR(500),
	[Date] DATETIME
	)

-- Gather Usage data
SET @Sql = '
INSERT INTO #Usage([Order], [Type], [ID], [Date]) 
SELECT	[Order], [Type], [ID], [Date]
FROM	( '

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @K2DBName AND [SchemaName] = N'Server' AND [TableName] = N'ProcSet' )
BEGIN
	SET @Sql = @Sql + '
			SELECT	 1 AS [Order]
					,''#WFDefs'' AS [Type]
					,CONVERT(NVARCHAR, [PS].[ID]) AS [ID]
					,[P].[ChangeDate] AS [Date]
			FROM	[' + @K2DBName + '].[Server].[ProcSet] AS [PS]
					JOIN [' + @K2DBName + '].[Server].[Proc] [P] ON [P].[ProcSetID] = [PS].[ID]
			WHERE	[P].[Ver] = 1

			UNION ALL '
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @K2LogDBName AND [SchemaName] = N'ServerLog' AND [TableName] = N'ProcInst' )
BEGIN
	SET @Sql = @Sql + '
			SELECT	 2 AS [Order]
					,''#ProcInsts'' AS [Type]
					,CONVERT(NVARCHAR, [PI].[ID]) AS [ID]
					,[PI].[StartDate] AS [Date]
			FROM	[' + @K2LogDBName + '].[ServerLog].[ProcInst] AS [PI]
					JOIN [' + @K2LogDBName + '].[ServerLog].[Proc] AS [P] ON [P].[ID] = [PI].[ProcID]

			UNION ALL '
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @K2LogDBName AND [SchemaName] = N'ServerLog' AND [TableName] = N'ActInst' )
BEGIN
	SET @Sql = @Sql + '
			SELECT	 3 AS [Order]
					,''#ActInsts'' AS [Type]
					,CONVERT(NVARCHAR, [AI].[ID]) AS [ID]
					,[AI].[StartDate] AS [Date]
			FROM	[' + @K2LogDBName + '].[ServerLog].[ActInst] AS [AI]

			UNION ALL '
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @K2LogDBName AND [SchemaName] = N'ServerLog' AND [TableName] = N'EventInst' )
BEGIN
	SET @Sql = @Sql + '
			SELECT	 4 AS [Order]
					,''#EventInsts'' As [Type]
					,CONVERT(NVARCHAR, [EI].[ID]) AS [ID]
					,[EI].[StartDate] AS [Date]
			FROM	[' + @K2LogDBName + '].[ServerLog].[EventInst] AS [EI]

			UNION ALL '
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @SmODBName AND [SchemaName] = N'SmartBroker' AND [TableName] = N'SmartObject' )
BEGIN
	SET @Sql = @Sql + '
			SELECT	 5 AS [Order]
					,''#SmartObjects'' AS [Type]
					,[Name] AS [ID]
					,[CreateDate] AS [Date]
			FROM	[' + @SmODBName + '].[SmartBroker].[SmartObject]
			WHERE	[CreateDate] IS NOT NULL

			UNION ALL '
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @HSDBName AND [SchemaName] = N'Form' AND [TableName] = N'Form' )
BEGIN
	SET @Sql = @Sql + '
			SELECT	 6 AS [Order]
					,''#Forms'' AS [Type]
					,[Name] AS [ID]
					,[CreatedDate] AS [Date]
			FROM	[' + @HSDBName + '].[Form].[Form]

			UNION ALL '
END

IF EXISTS ( SELECT * FROM #TableList WHERE [DatabaseName] = @HSDBName AND [SchemaName] = N'Form' AND [TableName] = N'View' )
BEGIN
	SET @Sql = @Sql + '
			SELECT	 7 AS [Order]
					,''#Views'' AS [Type]
					,[Name] AS [ID]
					,[CreatedDate] AS [Date]
			FROM	[' + @HSDBName + '].[Form].[View] '
END

IF (RIGHT(@Sql,10) = N'UNION ALL ')
BEGIN
	-- Get rid of the last 'UNION ALL' and whitespace
	SET @Sql = LEFT(@Sql, LEN(@Sql) - 16)
END

SET @Sql = @Sql + '
		) AS Arb'

PRINT @Sql
PRINT CONVERT(NVARCHAR, LEN(@Sql))

IF (LEN(@Sql) > 150)
BEGIN
	EXEC SP_EXECUTESQL @Sql
END

-- Determine Date Range extents
SET DATEFORMAT mdy;
DECLARE @StartDate DATETIME, @EndDate DATETIME

SELECT TOP 1 @StartDate=[Date] FROM #Usage ORDER BY [Date] ASC
SELECT TOP 1 @EndDate=[Date] FROM #Usage ORDER BY [Date] DESC
SET @EndDate = GETDATE()

-- FROM: http://stackoverflow.com/questions/7885851/months-between-two-dates
-- Get a unique name for each period within the date range
;WITH lv0 AS (SELECT 0 g UNION ALL SELECT 0)
    ,lv1 AS (SELECT 0 g FROM lv0 a CROSS JOIN lv0 b) -- 4
    ,lv2 AS (SELECT 0 g FROM lv1 a CROSS JOIN lv1 b) -- 16
    ,lv3 AS (SELECT 0 g FROM lv2 a CROSS JOIN lv2 b) -- 256
    ,lv4 AS (SELECT 0 g FROM lv3 a CROSS JOIN lv3 b) -- 65,536
    ,lv5 AS (SELECT 0 g FROM lv4 a CROSS JOIN lv4 b) -- 4,294,967,296
    ,Tally (n) AS (SELECT 0 UNION SELECT ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) FROM lv5)
SELECT	 n+1 AS [PeriodNum]
		,DATENAME(YEAR,DATEADD(MONTH,N,@StartDate)) + '-' + RIGHT('00' + CONVERT(NVARCHAR, DATEPART(MONTH, DATEADD(MONTH,n,@StartDate))), 2) AS [PeriodName]
		,DATENAME(YEAR,DATEADD(MONTH,N,@StartDate)) AS [YearPart]
		,CONVERT(NVARCHAR, DATEPART(MONTH, DATEADD(MONTH,n,@StartDate))) AS [MonthPart]
INTO	#Months
FROM	Tally 
WHERE	N BETWEEN 0 AND DATEDIFF(mm,@StartDate,@EndDate)  
ORDER BY n;

CREATE TABLE #CumUsage (
	[ID]			INT,
	[Month]			NVARCHAR(10),
	[#WFDefs]		INT,
	[#ProcInsts]	INT,
	[#ActInsts]		INT,
	[#EventInsts]	INT,
	[#SmartObjects]	INT,
	[#Forms]		INT,
	[#Views]		INT
	)

-- Aggregate the Usage data and use Pivot table for display
SET @Sql='
WITH
ctePreAgg AS
( --=== Preaggregate the data for an improvement in performance
 SELECT [Order]			= [Order],
		[Type]			= [Type],
		[Year]			= DATEPART(yy, [Date]),
        [Month]			= DATEPART(mm, [Date]),
		[Period]		= CONVERT(NVARCHAR(4), DATEPART(yy, [Date])) + ''-'' + RIGHT(''00'' + CONVERT(NVARCHAR(2), DATEPART(mm, [Date])), 2),
        [MonthCount]	= COUNT(*)
   FROM #Usage
GROUP BY DATEPART(yy, [Date]), DATEPART(mm, [Date]), [Type], [Order]
) --=== Now, pivot the data using a high performance crosstab.
INSERT INTO	[#CumUsage] ([ID], [Month], [#WFDefs], [#ProcInsts], [#ActInsts], [#EventInsts], [#SmartObjects], [#Forms], [#Views])
SELECT	[PeriodNum] AS [ID],
		[PeriodName] AS [Month],
		[#WFDefs] = SUM(CASE WHEN [Type] = ''#WFDefs'' THEN MonthCount ELSE 0 END),
		[#ProcInsts] = SUM(CASE WHEN [Type] = ''#ProcInsts'' THEN MonthCount ELSE 0 END),
		[#ActInsts] = SUM(CASE WHEN [Type] = ''#ActInsts'' THEN MonthCount ELSE 0 END),
		[#EventInsts] = SUM(CASE WHEN [Type] = ''#EventInsts'' THEN MonthCount ELSE 0 END),
		[#SmartObjects] = SUM(CASE WHEN [Type] = ''#SmartObjects'' THEN MonthCount ELSE 0 END),
		[#Forms] = SUM(CASE WHEN [Type] = ''#Forms'' THEN MonthCount ELSE 0 END),
		[#Views] = SUM(CASE WHEN [Type] = ''#Views'' THEN MonthCount ELSE 0 END)
   FROM ctePreAgg
		RIGHT JOIN #Months M ON M.[PeriodName] = ctePreAgg.[Period]
  GROUP BY [Period], M.[PeriodName], M.[PeriodNum]
  ORDER BY M.[PeriodName]
'

--PRINT @Sql
EXECUTE sp_executesql @Sql

-- Correlated scalar query to return cumulative results
SELECT	 [A].[Month]
		--,[A].[#WFDefs]
		,(SELECT SUM([B].[#WFDefs]) FROM #CumUsage [B] WHERE [B].[ID] <= [A].[ID]) AS [#WFDefs]
		--,[A].[#ProcInsts]
		,(SELECT SUM([B].[#ProcInsts]) FROM #CumUsage [B] WHERE [B].[ID] <= [A].[ID]) AS [#ProcInsts]
		--,[A].[#ActInsts]
		,(SELECT SUM([B].[#ActInsts]) FROM #CumUsage [B] WHERE [B].[ID] <= [A].[ID]) AS [#ActInsts]
		--,[A].[#EventInsts]
		,(SELECT SUM([B].[#EventInsts]) FROM #CumUsage [B] WHERE [B].[ID] <= [A].[ID]) AS [#EventInsts]
		--,[A].[#SmartObjects]
		,(SELECT SUM([B].[#SmartObjects]) FROM #CumUsage [B] WHERE [B].[ID] <= [A].[ID]) AS [#SmartObjects]
		--,[A].[#Forms]
		,(SELECT SUM([B].[#Forms]) FROM #CumUsage [B] WHERE [B].[ID] <= [A].[ID]) AS [#Forms]
		--,[A].[#Views]
		,(SELECT SUM([B].[#Views]) FROM #CumUsage [B] WHERE [B].[ID] <= [A].[ID]) AS [#Views]
FROM	#CumUsage [A]
ORDER BY [A].[ID];

--select * from #CumUsage order by [ID]

-- Cleanup
IF (Object_ID('tempdb..#Usage') IS NOT NULL)
	DROP TABLE [#Usage]

IF (Object_ID('tempdb..#Months') IS NOT NULL)
	DROP TABLE [#Months]

IF (Object_ID('tempdb..#CumUsage') IS NOT NULL)
	DROP TABLE [#CumUsage]

IF OBJECT_ID(N'tempdb..#TableList', 'U') IS NOT NULL
    DROP TABLE [#TableList]