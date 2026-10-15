/* This script is to help update azure upload settings */

/* Get all environment settings */
SELECT TOP (1000) [Scope]
      ,[Name]
      ,[Value]
      ,[Edition]
      ,[UserName]
  FROM [AgentDB].[dbo].[UserData]

/* Update Storage Connection String */
UPDATE [AgentDB].dbo.UserData
	SET [Value] = 'CopiedFromAzure'
	WHERE Name = 'UploadManagerAzureStorageConnectionString'

/* Update Local Temporary Storage */
UPDATE [AgentDB].dbo.UserData
	SET [Value] = 'C:\MetalogixAzureTemp'
	WHERE Name = 'UploadManagerLocalTemporaryStorageLocation'