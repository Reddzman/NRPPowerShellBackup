SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[permissions](
	[ServerUNC] [nvarchar](260) NULL,
	[ShareName] [nvarchar](260) NULL,
	[Owner] [nvarchar](260) NULL,
	[Group] [nvarchar](260) NULL,
	[Path] [nvarchar](max) NULL,
	[FileSystemRights] [nvarchar](340) NULL,
	[AccessControlType] [nvarchar](15) NULL,
	[IdentityReference] [nvarchar](100) NULL,
	[IsInherited] [nvarchar](10) NULL,
	[client] [nvarchar](50) NULL,
	[scandate] [nvarchar](15) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

SET ANSI_PADDING ON
GO
CREATE NONCLUSTERED INDEX [IX_permission] ON [dbo].[permissions]
(
	[client] ASC
)
INCLUDE([scandate]) WITH (STATISTICS_NORECOMPUTE = OFF, DROP_EXISTING = OFF, ONLINE = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO