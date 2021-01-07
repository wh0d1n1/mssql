  
IF NOT EXISTS (SELECT * FROM sys.configurations WHERE name ='show advanced options' AND value=1)
  	BEGIN
	EXECUTE sp_configure 'show advanced options', 1;  
	 RECONFIGURE
	 
     EXECUTE sp_configure 'xp_cmdshell', 1;
     RECONFIGURE 

     end 

DECLARE @cmd VARCHAR(255)
SET @cmd = 'osql -E /Q "SET NOCOUNT ON SELECT GETDATE() DECLARE @i INT SET @i = 1 WHILE @i <= 100 BEGIN PRINT @i SET @i = @i + 1 END SELECT GETDATE()" /o D:\test2.csv' 
EXEC xp_cmdshell @cmd