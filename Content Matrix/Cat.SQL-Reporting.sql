-- Site Collection Start Time
SELECT MIN (History.[Started]) as 'SiteStartTime'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'ITS%Inc%' AND [Started] >= '3/23/2018 5:00:00 PM' AND [Finished] <= '3/25/2018 8:00:00 AM'

-- Site Collection Finish Time
SELECT MAX (History.[Finished]) as 'SiteFinishTime'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'ITS%Inc%' AND [Started] >= '3/23/2018 5:00:00 PM' AND [Finished] <= '3/25/2018 8:00:00 AM'

-- Specific Site Data Moved
SELECT SUM (History.[LicensedDataUsed]) as 'SiteDataMigrated'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'ITS - % - Inc%'and [Started] >= '3/23/2018 5:00:00 PM'

-- Migration Jobs
SELECT COUNT(History.[JobID]) as 'TotalJobs'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'PM-ITS%Inc%' AND [Started] >= '3/23/2018 5:00:00 PM' AND [Finished] <= '3/25/2018 8:00:00 AM'

-- Completed Active Site Migration Jobs
SELECT COUNT(History.[JobID]) as 'CompletedJobs'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'ITS - % - Inc%' AND [Status] like 'Done'

-- Running Migration Jobs
SELECT COUNT(History.[JobID]) as 'RunningJobs'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'ITS - % - Inc%' AND [Status] like 'Running'

-- Failed Migration Jobs
SELECT COUNT(History.[JobID]) as 'FailedJobs'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'ITS - % - Inc%' AND [Status] like 'Failed'

-- Migration Jobs Complete w/ Warnings
SELECT COUNT(History.[JobID]) as 'WarningJobs'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'ITS - % - Inc%' AND [Status] like 'Done' AND [ResultsSummary] Like '%Warnings%'

-- Migration Jobs Complete w/ Failures
SELECT COUNT(History.[JobID]) as 'FailuresJobs'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'ITS - % - Perms%'and [Started] >= '3/23/2018 5:00:00 PM' AND [Status] like 'Done' AND [ResultsSummary] Like '%Failures%'

-- Migration Jobs Complete w/o Failures or Warnings
SELECT COUNT(History.[JobID]) as 'FailuresJobs'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History
	WHERE [Title] like 'ITS - % - Inc%' AND [Status] like 'Done' AND [ResultsSummary] not Like '%Failures%' AND [ResultsSummary] not like '%Warnings%'

-- Total Data Moved
SELECT SUM (History.[LicensedDataUsed]) as 'TotalDataMigrated'
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory] as History

-- Running or Queued jbos
SELECT [Title], [Source], [Target]
	FROM [Metalogix_Logging_Dev].[dbo].[JobHistory]
	WHERE [Status] = 'Running' OR [Status] = 'Queued'