IF Object_Id('dbo.F_ExtractSubString') IS NOT NULL 
DROP FUNCTION dbo.F_ExtractSubString
	go 
create FUNCTION [dbo].[F_ExtractSubString](@String VARCHAR(MAX),@NroSubString INT,@Separator VARCHAR(5))
RETURNS VARCHAR(MAX) AS
BEGIN
    DECLARE @St INT = 0, @End INT = 0, @Ret VARCHAR(MAX)
    SET @String = @String + @Separator
    WHILE CHARINDEX(@Separator, @String, @End + 1) > 0 AND @NroSubString > 0
    BEGIN
        SET @St = @End + 1
        SET @End = CHARINDEX(@Separator, @String, @End + 1)
        SET @NroSubString = @NroSubString - 1
    END
    IF @NroSubString > 0
        SET @Ret = ''
    ELSE
        SET @Ret = SUBSTRING(@String, @St, @End - @St)
    RETURN @Ret
END

GO
