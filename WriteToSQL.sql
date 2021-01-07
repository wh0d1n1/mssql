IF NOT EXISTS (SELECT (1) FROM master..sysdatabases WHERE name = 'SE')
	BEGIN
		CREATE DATABASE SE
	END
GO
USE SE
GO
IF EXISTS (SELECT (1) FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[usp_UseOA]')
				AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
	BEGIN
		DROP PROCEDURE [dbo].[usp_UseOA]
	END
GO
CREATE PROCEDURE usp_UseOA (
   @File varchar(1000)
 , @Str varchar(1000)
) 
AS
DECLARE @FS int
        , @OLEResult int
        , @FileID int
EXECUTE @OLEResult = sp_OACreate 
 'Scripting.FileSystemObject'
 , @FS OUT
IF @OLEResult <> 0 
	BEGIN
		PRINT
'Error: Scripting.FileSystemObject'
	END
-- Opens the file specified by the @File input parameter 
execute @OLEResult = sp_OAMethod 
   @FS
   , 'OpenTextFile'
	, @FileID OUT
	, @File
	, 8
	, 1
-- Prints error if non 0 return code during sp_OAMethod OpenTextFile execution 
IF @OLEResult <> 0 
	BEGIN
		PRINT 'Error: OpenTextFile'
	END
-- Appends the string value line to the file specified by the @File input parameter
execute @OLEResult = sp_OAMethod 
   @FileID
	, 'WriteLine'
	, Null
	, @Str
-- Prints error if non 0 return code during sp_OAMethod WriteLine execution 
IF @OLEResult <> 0 
	BEGIN
		PRINT 'Error : WriteLine'
	END
EXECUTE @OLEResult = sp_OADestroy @FileID
EXECUTE @OLEResult = sp_OADestroy @FS
go
--Execution Script
DECLARE
@file as varchar(1000)

, @i INT
SET @i = 1
SET @file = 'D:\test5.csv'
WHILE @i <= 100
 BEGIN
   -- executes this stored procedure for each @i value
   EXEC SE..usp_UseOA @file, @i
   SET @i = @i + 1
 END
 go