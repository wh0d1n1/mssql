SET ANSI_NULLS		ON;
SET QUOTED_IDENTIFIER 	ON;			

IF NOT EXISTS	

(
SELECT * 
FROM 
	sys.configurations 
WHERE 
	name ='Ole Automation Procedures' 
	AND value = 1
)

BEGIN
     EXECUTE sp_configure 'Ole Automation Procedures', 1;
     RECONFIGURE;  
END 
			
GO
IF 
	Object_Id('dbo.GetWebService','P') IS NOT NULL 
  	DROP procedure	dbo.GetWebService		
GO

  CREATE PROCEDURE	dbo.GetWebService
					@TheURL			VARCHAR(255),				--	the url of the web service
					@TheResponse	NVARCHAR(4000) OUTPUT		--	the resulting JSON
  AS
BEGIN

      DECLARE 
		@obj INT
	  , @hr INT
	  , @status INT
	  , @message VARCHAR(255)
	  ;


      EXEC @hr = sp_OACreate 'MSXML2.ServerXMLHttp', @obj OUT;

	  IF @hr <> 0
      SET @message = 'sp_OAMethod Open failed';
      IF @hr = 0 

	  EXEC 
	  @hr = sp_OAMethod @obj, 'open', NULL, 'GET', @TheURL, false;

      SET @message = 'sp_OAMethod setRequestHeader failed';

      IF @hr = 0 
					EXEC @hr = sp_OAMethod 
											@obj
												,'setRequestHeader'
												,NULL
												,'Content-Type'
												,'application/x-www-form-urlencoded'
												;

					SET @message = 'sp_OAMethod Send failed';
      IF @hr = 0 
					EXEC @hr = sp_OAMethod 
											@obj
												, send
												, NULL
												, ''
												;
					SET @message = 'sp_OAMethod read status failed';


      IF @hr = 0 EXEC @hr = sp_OAGetProperty @obj, 'status', @status OUT;
      IF @status <> 200 
		BEGIN
                          SELECT @message = 'sp_OAMethod http status ' + Str(@status), @hr = -1;
        END;
      SET @message = 'sp_OAMethod read response failed';
      IF @hr = 0
        BEGIN
          EXEC @hr = sp_OAGetProperty @obj, 'responseText', @Theresponse OUT;
          END;
      EXEC sp_OADestroy @obj;
      --IF @hr <> 0 RAISERROR(@message, 16, 1);
      END;
  GO

     DECLARE @Theresponse	NVARCHAR(4000) 
	 DECLARE @TheURL		NVARCHAR(255)

Begin

set @TheURL = 'http://www.wikipedia.org/'


     EXECUTE dbo.GetWebService @TheURL, @Theresponse OUTPUT;
     PRINT @Theresponse 
END
