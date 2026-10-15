/* Creating Table */
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[directories](
	[ServerUNC] [nvarchar](260) NULL,
	[ShareName] [nvarchar](260) NULL,
	[Name] [nvarchar](max) NULL,
	[Parent] [nvarchar](max) NULL,
	[Root] [nvarchar](520) NULL,
	[FullName] [nvarchar](max) NULL,
	[LastAccessTime] [smalldatetime] NULL,
	[LastAccessTimeUtc] [smalldatetime] NULL,
	[LastWriteTime] [smalldatetime] NULL,
	[LastWriteTimeUtc] [smalldatetime] NULL,
	[Attributes] [nvarchar](100) NULL,
	[CreationTime] [smalldatetime] NULL,
	[CreationTimeUtc] [smalldatetime] NULL,
	[client] [nvarchar](50) NULL,
	[scandate] [nvarchar](15) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO

/* Create Index */
CREATE NONCLUSTERED INDEX [IX_directory] ON [dbo].[directories]
(
	[client] ASC
)
INCLUDE([scandate]) WITH (STATISTICS_NORECOMPUTE = OFF, DROP_EXISTING = OFF, ONLINE = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO

/* Table updates after creation - Table creation already updated */
ALTER TABLE [dbo].[directories]
ALTER COLUMN [Name] [nvarchar](max) NULL
ALTER TABLE [dbo].[directories]
ALTER COLUMN [Parent] [nvarchar](max) NULL
ALTER TABLE [dbo].[directories]
ALTER COLUMN [Root] [nvarchar](520) NULL
