/* DirectoryReport */
let
    Client = ClientName,
    Date = ScanDate,
    Source = Sql.Database("sqlsvr-spmigration.database.windows.net","sql-spmigration",
    [Query="SELECT ServerUNC, ShareName, Name, Parent, Root, FullName, LastAccessTime, LastAccessTimeUtc, LastWriteTime, 
    LastWriteTimeUtc, Attributes, CreationTime, CreationTimeUtc, client, scandate FROM dbo.directories WHERE client='" & Client & "' AND scandate='" & Date & "'"])
in
    #"Source"

/* Files Report */
let
    Client = ClientName,
    Date = ScanDate,
    Source = Sql.Database("sqlsvr-spmigration.database.windows.net","sql-spmigration",
    [Query="SELECT ServerUNC, ShareName, Name, BaseName, Attributes, CreationTime, CreationTimeUtc, DirectoryName,
    Extension, IsReadOnly, LastAccessTime, LastAccessTimeUtc, LastWriteTime, LastWriteTimeUtc, Length, FileName, client, scandate 
    FROM dbo.files WHERE client='" & Client & "' AND scandate='" & Date & "'"])
in
    #"Source"

/* Permissions Report */
let
    Client = ClientName,
    Date = ScanDate,
    Source = Sql.Database("sqlsvr-spmigration.database.windows.net","sql-spmigration",
    [Query="SELECT ServerUNC, ShareName, Owner, [Group], Path, FileSystemRights, AccessControlType, IdentityReference, IsInherited, client, scandate 
    FROM dbo.permissions WHERE client='" & Client & "' AND scandate='" & Date & "'"])
in
    #"Source"