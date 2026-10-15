/* Delete specific client data; use select to test results first */

/* Directories Table */
SELECT *
  FROM [dbo].[directories]
  WHERE [Client] = 'Client'

DELETE
  FROM [dbo].[directories]
  WHERE [Client] = 'Client'

/* Files Table */
SELECT *
  FROM [dbo].[files]
  WHERE [Client] = 'Client'

DELETE
  FROM [dbo].[files]
  WHERE [Client] = 'Client'

/* Permissions Table */
SELECT *
  FROM [dbo].[permissions]
  WHERE [Client] = 'Client'

DELETE
  FROM [dbo].[permissions]
  WHERE [Client] = 'Client'