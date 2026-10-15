SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[files](
	[ServerUNC] [nvarchar](260) NULL,
	[ShareName] [nvarchar](260) NULL,
	[Name] [nvarchar](260) NULL,
	[BaseName] [nvarchar](260) NULL,
	[Attributes] [nvarchar](100) NULL,
	[CreationTime] [DATETIME2] NULL,
	[CreationTimeUtc] [DATETIME2] NULL,
	[PSPath] [nvarchar](max) NULL,
	[DirectoryName] [nvarchar](260) NULL,
	[Extension] [nvarchar](50) NULL,
	[IsReadOnly] [nvarchar](10) NULL,
	[LastAccessTime] [DATETIME2] NULL,
	[LastAccessTimeUtc] [DATETIME2] NULL,
	[LastWriteTime] [DATETIME2] NULL,
	[LastWriteTimeUtc] [DATETIME2] NULL,
	[Length] [bigint] NULL,
	[FileName] [nvarchar](260) NULL,
	[client] [nvarchar](50) NULL,
	[scandate] [nvarchar](15) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

SET ANSI_PADDING ON
GO
CREATE NONCLUSTERED INDEX [IX_file] ON [dbo].[files]
(
	[client] ASC
)
INCLUDE([scandate]) WITH (STATISTICS_NORECOMPUTE = OFF, DROP_EXISTING = OFF, ONLINE = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO


ALTER TABLE [dbo].[files]
ALTER COLUMN [Name] [nvarchar](max) NULL
ALTER TABLE [dbo].[files]
ALTER COLUMN [BaseName] [nvarchar](max) NULL
ALTER TABLE [dbo].[files]
ALTER COLUMN [DirectoryName] [nvarchar](max) NULL
ALTER TABLE [dbo].[files]
ALTER COLUMN [FileName] [nvarchar](max) NULL
ALTER TABLE [dbo].[files]
ALTER COLUMN [Extension] [nvarchar](260) NULL
ALTER TABLE [dbo].[files]
ALTER COLUMN [CreationTime] DATETIME2
ALTER TABLE [dbo].[files]
ALTER COLUMN [CreationTimeUtc] DATETIME2
ALTER TABLE [dbo].[files]
ALTER COLUMN [LastAccessTime] DATETIME2
ALTER TABLE [dbo].[files]
ALTER COLUMN [LastAccessTimeUtc] DATETIME2
ALTER TABLE [dbo].[files]
ALTER COLUMN [LastWriteTime] DATETIME2
ALTER TABLE [dbo].[files]
ALTER COLUMN [LastWriteTimeUtc] DATETIME2