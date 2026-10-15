-- Select Database to Query
USE SPOMig_JobDB;
GO

-- Get Estimated StartTime & Current Time
DECLARE @StartTime as DateTime;
DECLARE @Date as DATE;
DECLARE @MinusHours as INT;
SET @MinusHours = -18;
SET @StartTime = DATEADD(hh,@MinusHours,GETDATE());
SET @Date = DATEADD(hh,@MinusHours,GETDATE());

-- Rename All jobs created since StartTime
UPDATE [dbo].[JobHistory]
	SET [Title] = (
		CASE
			WHEN [Target] = 'https://neomaterials.sharepoint.com/' THEN CONCAT('Homepage - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/finance' THEN CONCAT('Finance - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/tax' THEN CONCAT('Tax - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/legal' THEN CONCAT('Legal - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/legal/legacy' THEN CONCAT('Legal Legacy - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/hr' THEN CONCAT('HR - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/it' THEN CONCAT('IT - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/seg/co' THEN CONCAT('Chem & Ox - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/co-sales' THEN CONCAT('CO Sales & Mark - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/co-ops' THEN CONCAT('CO Sales & Ops - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/co-water' THEN CONCAT('CO Water Treat - Nightly Inc - ',@Date)
			WHEN [Target] = 'https://neomaterials.sharepoint.com/sites/rm-silmet' THEN CONCAT('RM Silmet - Nightly Inc - ',@Date)
		END
	)
	WHERE [Started] >= @StartTime;
GO