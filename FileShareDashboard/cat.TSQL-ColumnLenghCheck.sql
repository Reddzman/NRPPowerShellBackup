/* Directories */
SELECT LEN([Name]), [Name] 
FROM dbo.directories
order BY LEN([Name]) desc

SELECT LEN([Parent]), [Parent] 
FROM dbo.directories
order BY LEN([Parent]) desc

SELECT LEN([Root]), [Root] 
FROM dbo.directories
order BY LEN([Root]) desc

SELECT LEN([Attributes]), [Attributes]
FROM dbo.directories
order BY LEN([Attributes]) desc

/* Files */
SELECT LEN([Name]), [Name] 
FROM dbo.files
order BY LEN([Name]) desc

SELECT LEN([BaseName]), [BaseName] 
FROM dbo.files
order BY LEN([BaseName]) desc

SELECT LEN([DirectoryName]), [DirectoryName] 
FROM dbo.files
order BY LEN([DirectoryName]) desc

SELECT LEN([FileName]), [FileName] 
FROM dbo.files
order BY LEN([FileName]) desc

SELECT LEN([Extension]), [Extension] 
FROM dbo.files
order BY LEN([Extension]) desc

SELECT LEN([IsReadOnly]), [IsReadOnly] 
FROM dbo.files
order BY LEN([IsReadOnly]) desc

SELECT LEN([LastWriteTime]), [LastWriteTime] 
FROM dbo.files
order BY LEN([LastWriteTime]) desc

SELECT LEN([LastWriteTimeUtc]), [LastWriteTimeUtc] 
FROM dbo.files
order BY LEN([LastWriteTimeUtc]) desc

/* Permissions */
SELECT LEN([Owner]), [Owner] 
FROM dbo.permissions
order BY LEN([Owner]) desc

SELECT LEN([Group]), [Group] 
FROM dbo.permissions
order BY LEN([Group]) desc

SELECT LEN([FileSystemRights]), [FileSystemRights] 
FROM dbo.permissions
order BY LEN([FileSystemRights]) desc

SELECT LEN([IdentityReference]), [IdentityReference] 
FROM dbo.permissions
order BY LEN([IdentityReference]) desc

SELECT LEN([AccessControlType]), [AccessControlType] 
FROM dbo.permissions
order BY LEN([AccessControlType]) desc