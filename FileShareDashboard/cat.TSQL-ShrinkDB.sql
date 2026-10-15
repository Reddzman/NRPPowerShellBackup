USE [sql-spmigration]

/* This query will help you determine if it is worth shrinking the DB and how
small you can go. */
SELECT file_id, type_desc,
       CAST(FILEPROPERTY(name, 'SpaceUsed') AS decimal(19,4)) * 8 / 1024. AS space_used_mb,
       CAST(size/128.0 - CAST(FILEPROPERTY(name, 'SpaceUsed') AS int)/128.0 AS decimal(19,4)) AS space_unused_mb,
       CAST(size AS decimal(19,4)) * 8 / 1024. AS space_allocated_mb,
       CAST(max_size AS decimal(19,4)) * 8 / 1024. AS max_size_mb
FROM sys.database_files;
GO

/* Use the SELECT Above to confirm file_id (#1), and the size you want to shrink
to (#1) in the SHRINKFILE command below. */
DBCC SHRINKFILE (2,200);
GO

/* Use the SELECT Above to confirm file_id (#1), and the size you want to shrink
to (#1) in the SHRINKFILE command below. */
DBCC SHRINKFILE (3,94000);
GO

/* Use the SELECT Above to confirm file_id (#1), and the size you want to shrink
to (#1) in the SHRINKFILE command below. */
DBCC SHRINKFILE (1,155000);
GO

/* This is just here to confirm the end result */
SELECT file_id, type_desc,
       CAST(FILEPROPERTY(name, 'SpaceUsed') AS decimal(19,4)) * 8 / 1024. AS space_used_mb,
       CAST(size/128.0 - CAST(FILEPROPERTY(name, 'SpaceUsed') AS int)/128.0 AS decimal(19,4)) AS space_unused_mb,
       CAST(size AS decimal(19,4)) * 8 / 1024. AS space_allocated_mb,
       CAST(max_size AS decimal(19,4)) * 8 / 1024. AS max_size_mb
FROM sys.database_files;
GO
