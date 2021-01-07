IF Object_Id('dbo.JSONPathsAndValues') IS NOT NULL DROP FUNCTION dbo.JSONPathsAndValues;
  GO
  CREATE FUNCTION dbo.JSONPathsAndValues
    (@JSONData NVARCHAR(MAX))
  RETURNS @TheHierarchyMetadata TABLE
    (
    -- columns returned by the function
    element_id INT NOT NULL,
    Depth INT NOT NULL,
    Thepath NVARCHAR(2000),
    ValueType VARCHAR(10) NOT NULL,
    TheValue NVARCHAR(MAX) NOT NULL
    )
  AS
    -- body of the function
    BEGIN
      DECLARE @ii INT = 1, @rowcount INT = -1;
      DECLARE @null INT = 0, @string INT = 1, @int INT = 2, --
        @boolean INT = 3, @array INT = 4, @object INT = 5;
      DECLARE @TheHierarchy TABLE
        (
        element_id INT IDENTITY(1, 1) PRIMARY KEY,
        Depth INT NOT NULL, /* effectively, the recursion level. =the depth of nesting*/
        Thepath NVARCHAR(2000) NOT NULL,
        TheName NVARCHAR(2000) NOT NULL,
        TheValue NVARCHAR(MAX) NOT NULL,
        ValueType VARCHAR(10) NOT NULL
        );
      INSERT INTO @TheHierarchy
        (Depth, Thepath, TheName, TheValue, ValueType)
        SELECT @ii, '$', '$', @JSONData, 'object';
      WHILE @rowcount <> 0
        BEGIN
          SELECT @ii = @ii + 1;
          INSERT INTO @TheHierarchy
            (Depth, Thepath, TheName, TheValue, ValueType)
            SELECT @ii,
              CASE WHEN [Key] NOT LIKE '%[^0-9]%' THEN Thepath + '[' + [Key] + ']' --nothing but numbers 
                WHEN [Key] LIKE '%[$ ]%' THEN Thepath + '."' + [Key] + '"' --got a space in it
                ELSE Thepath + '.' + [Key] END, [Key], Coalesce(Value,''),
              CASE Type WHEN @string THEN 'string'
                WHEN @null THEN 'null'
                WHEN @int THEN 'int'
                WHEN @boolean THEN 'boolean'
                WHEN @int THEN 'int'
                WHEN @array THEN 'array' ELSE 'object' END
            FROM @TheHierarchy AS m
              CROSS APPLY OpenJson(TheValue) AS o
            WHERE ValueType IN
          ('array', 'object') AND Depth = @ii - 1;
          SELECT @rowcount = @@RowCount;
        END;
      INSERT INTO @TheHierarchyMetadata
        SELECT element_id, Depth, Thepath, ValueType, TheValue 
        FROM @TheHierarchy
        WHERE ValueType NOT IN
      ('array', 'object');
      RETURN;
    END;
  GO
