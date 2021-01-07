IF NOT EXISTS (SELECT * FROM sys.configurations WHERE name ='Ole Automation Procedures' AND value=1)
  	BEGIN
     EXECUTE sp_configure 'Ole Automation Procedures', 1;
     RECONFIGURE;  
	 EXECUTE sp_configure 'show advanced options', 1;  
	 RECONFIGURE;  
     end 
  SET ANSI_NULLS ON;
  SET QUOTED_IDENTIFIER ON;
  GO
  IF Object_Id('dbo.get_http','P') IS NOT NULL 
  	drop procedure dbo.get_http
	go 
   create procedure get_http @url varchar(2000),
						@output text output
as
begin
/*
exec get_http 'https://www.bing.com', 
*/

    declare @hr int;
    declare @win int;
    declare @errorMessage varchar(2000);
	 

    begin try

      EXEC @hr=sp_OACreate 'WinHttp.WinHttpRequest.5.1',@win OUT 
      IF @hr <> 0
      begin;
        set @errorMessage = concat('sp_OACreate failed ', convert(varchar(20),cast(@hr as varbinary(4)),1));
        throw 60000, @errorMessage, 1;
      end;

      EXEC @hr=sp_OAMethod @win, 'Open',NULL,'GET',@url,'false'
      IF @hr <> 0
      begin;
        set @errorMessage = concat('Open failed ', convert(varchar(20),cast(@hr as varbinary(4)),1));
        throw 60000, @errorMessage, 1;
      end;

      --Option is an indexed property, so newvalue = 2048 and index = 9
      --sp_OASetProperty objecttoken , propertyname , newvalue [ , index... ] 
      EXEC @hr=sp_OASetProperty @win, 'Option', 2048, 9
      IF @hr <> 0
      begin;
        set @errorMessage = concat('set Option failed ', convert(varchar(20),cast(@hr as varbinary(4)),1) );
        throw 60000, @errorMessage, 1;
      end;

      EXEC @hr=sp_OAMethod @win,'Send'
      IF @hr <> 0
      begin;
        set @errorMessage = concat('Send failed ', convert(varchar(20),cast(@hr as varbinary(4)),1));
        throw 60000, @errorMessage, 1;
      end;

      declare @status int
      EXEC @hr=sp_OAGetProperty @win,'Status', @status out
      IF @hr <> 0
      begin;
        set @errorMessage = concat('get Status failed ', convert(varchar(20),cast(@hr as varbinary(4)),1));
        throw 60000, @errorMessage, 1;
      end;

      if @status <> 200
      begin;
        set @errorMessage = concat('web request failed ', @status);
        throw 60000, @errorMessage, 1;
      end;

      declare @response table(text nvarchar(max));

      insert into @response(text)
      EXEC @hr=sp_OAGetProperty @win,'ResponseText';
      IF @hr <> 0
      begin;
        set @errorMessage = concat('get ResponseText failed ', convert(varchar(20),cast(@hr as varbinary(4)),1));
        throw 60000, @errorMessage, 1;
      end;

	        select  @output= [text]

      from @response


      EXEC @hr=sp_OADestroy @win 
      IF @hr <> 0 EXEC sp_OAGetErrorInfo @win;

    end try
    begin catch
      declare @error varchar(200) = error_message()
      declare @source varchar(200);
      declare @description varchar(200);
      declare @helpfile varchar(200);
      declare @helpid int;

      exec sp_OAGetErrorInfo @win, @source out, @description out, @helpfile out, @helpid out;
      declare @msg varchar(max) = concat('COM Failure ', @error,' ',@source,' ',@description)

      EXEC @hr=sp_OADestroy @win; 
      --IF @hr <> 0 EXEC sp_OAGetErrorInfo @win;
      throw 60000, @msg, 1;

            RETURN;
    end catch
end
go


declare @output  VARCHAR(MAX)
declare @baseurl VARCHAR(MAX) = 'https://catalog.data.gov/api/3/action/package_search?facet.field='
declare @search varchar(max) = '&q="business license"'
declare @searchterm varchar(max)  = '["name"]'
DECLARE @RowCount varchar(max) = '200'
DECLARE @start varchar(max) = '200'
declare @starts varchar(max) = '&start=' + @start
declare @rows varchar(max) = '&rows='+ @RowCount 
declare @API varchar(max) = @baseurl
--+ @format 
+ @searchterm
+@rows
+ @search



exec dbo.get_http @API,
@output output 


SELECT 
x.alljson
,x.extras
,[dbo].[F_ExtractSubString]([dbo].[F_ExtractSubString](replace(replace(replace(r.qa,'{',''),nchar(39),''),'}',''),1,','),2,': ') as updated
,[dbo].[F_ExtractSubString]([dbo].[F_ExtractSubString](replace(replace(replace(r.qa,'{',''),nchar(39),''),'}',''),2,','),2,': ') as openness_score
,[dbo].[F_ExtractSubString]([dbo].[F_ExtractSubString](replace(replace(replace(r.qa,'{',''),nchar(39),''),'}',''),3,','),2,': ') as archival_timestamp
,replace([dbo].[F_ExtractSubString]([dbo].[F_ExtractSubString](replace(replace(replace(r.qa,'{',''),nchar(39),''),'}',''),4,','),2,': '),'u','') as [format]
,[dbo].[F_ExtractSubString]([dbo].[F_ExtractSubString](replace(replace(replace(r.qa,'{',''),nchar(39),''),'}',''),5,','),2,': ') as created
,[dbo].[F_ExtractSubString]([dbo].[F_ExtractSubString](replace(replace(replace(r.qa,'{',''),nchar(39),''),'}',''),6,','),2,': ') as resource_timestamp
,[dbo].[F_ExtractSubString]([dbo].[F_ExtractSubString](replace(replace(replace(r.qa,'{',''),nchar(39),''),'}',''),7,','),2,': ')
+[dbo].[F_ExtractSubString]([dbo].[F_ExtractSubString](replace(replace(replace(r.qa,'{',''),nchar(39),''),'}',''),8,','),2,': ') as openness_score_reason
,x.resources
,x.tags
,x.license_title					
,x.maintainer						
,x.private						
,x.maintainer_email				
,x.num_tags						
,x.id								
,x.metadata_created				
,x.metadata_modified								
,x.state							
,x.creator_user_id				
,x.type							
,x.num_resources					
,x.license_id						
,x.name							
,x.isopen							
,x.notes							
,x.owner_org						
,x.license_url					
,x.title							
,x.revision_id					
,x.tracking_summary_total			
,x.tracking_summary_recent		
,x.description					
,x.created						
,x.title							
,x.name							
,x.is_organization				
,x.state							
,x.image_url						
,x.revision_id					
,x.type							
,x.id								
,x.approval_status				
,r.package_id
,r.id				
,r.state			
,r.archiver		
,r.description	
,r.format			
,r.mimetype		
,r.name			
,r.created		
,r.url						
,r.position		
,r.revision_id	
,t.state			
,t.display_name	
,t.id				
,t.name			
,t2.total
,t2.recent

from openjson(@output,'$.result.results') 
with(
alljson nVARCHAR(max) '$' as json
,resources NVARCHAR(MAX) '$.resources' AS JSON  
,tags								nVARCHAR(max) '$.tags' as json
,license_title									nVARCHAR(4000) '$.license_title'
,maintainer										nVARCHAR(4000) '$.maintainer'
,private										nVARCHAR(4000) '$.private'
,maintainer_email								nVARCHAR(4000) '$.maintainer_email'
,num_tags										nVARCHAR(4000) '$.num_tags'
,id												nVARCHAR(4000) '$.id'
,metadata_created								nVARCHAR(4000) '$.metadata_created'
,metadata_modified								nVARCHAR(4000) '$.metadata_modified'
,state											nVARCHAR(4000) '$.state'
,version										nVARCHAR(4000) '$.version'
,creator_user_id								nVARCHAR(4000) '$.creator_user_id'
,type											nVARCHAR(4000) '$.type'
,num_resources									nVARCHAR(4000) '$.num_resources'
,license_id										nVARCHAR(4000) '$.license_id'
,name											nVARCHAR(4000) '$.name'
,isopen											nVARCHAR(4000) '$.isopen'
,notes											nVARCHAR(4000) '$.notes'
,owner_org										nVARCHAR(4000) '$.owner_org'
,license_url									nVARCHAR(4000) '$.license_url'
,title											nVARCHAR(4000) '$.title'
,revision_id									nVARCHAR(4000) '$.revision_id'
,tracking_summary_total							nVARCHAR(4000) '$.tracking_summary.total'
,tracking_summary_recent						nVARCHAR(4000) '$.tracking_summary.recent'
,description									nVARCHAR(4000) '$.organization.description'
,created										nVARCHAR(4000) '$.organization.created'
,title											nVARCHAR(4000) '$.organization.title'
,name											nVARCHAR(4000) '$.organization.name'
,is_organization								nVARCHAR(4000) '$.organization.is_organization'
,state											nVARCHAR(4000) '$.organization.state'
,image_url										nVARCHAR(4000) '$.organization.image_url'
,revision_id									nVARCHAR(4000) '$.organization.revision_id'
,type											nVARCHAR(4000) '$.organization.type'
,id												nVARCHAR(4000) '$.organization.id'
,approval_status								nVARCHAR(4000) '$.organization.approval_status'
,extras											nVARCHAR(max) '$.extras' as json
) x


cross apply openjson(isnull(x.resources,'{}'),'$')  
with (
package_id nvarchar(max) '$.package_id'
,id									nVARCHAR(4000) '$.id'
,size								nVARCHAR(4000) '$.size'
,state									nVARCHAR(4000) '$.state'
,archiver									nVARCHAR(4000) '$.archiver'
,description									nVARCHAR(4000) '$.description'
,format									nVARCHAR(4000) '$.format'
,mimetype									nVARCHAR(4000) '$.mimetype'
,name									nVARCHAR(4000) '$.name'
,created									nVARCHAR(4000) '$.created'
,url									nVARCHAR(4000) '$.url'
,qa									VARCHAR(8000) '$.qa' 
,position									nVARCHAR(4000) '$.position'
,revision_id									nVARCHAR(4000) '$.revision_id'
,tracking_summary nVARCHAR(max) '$.tracking_summary' as json
) r

cross apply openjson(isnull(x.tags,'{}'),'$') 

with(
vocabulary_id									nVARCHAR(4000) '$.vocabulary_id'
,state									nVARCHAR(4000) '$.state'
,display_name									nVARCHAR(4000) '$.display_name'
,id									nVARCHAR(4000) '$.id'
,name									nVARCHAR(4000) '$.name'
)

t

cross apply openjson(isnull(r.tracking_summary,'{}'),'$')
with(total nvarchar(max) '$.total'
,recent nvarchar(max) '$.recent') t2




DECLARE @JSONData varchar(max)

exec dbo.get_http 'http://us-city.census.okfn.org/api/entries.json'

,@JSONData output 



SELECT 
*
from OPENJSON(@JSONData, '$.results')
with(
id						nVARCHAR(4000)					'$.id'
,[site]					nVARCHAR(4000)					'$.site'
,[timestamp]				nVARCHAR(4000)				'$.timestamp'
,[year]					nVARCHAR(4000)					'$.year'
,place					nVARCHAR(4000)					'$.place'
,dataset				nVARCHAR(4000)					'$.dataset'
,[exists]					nVARCHAR(4000)				'$.exists'
,digital				nVARCHAR(4000)					'$.digital'
,[public]					nVARCHAR(4000)				'$.public'
,[online]					nVARCHAR(4000)				'$.online'
,free					nVARCHAR(4000)					'$.free'
,machinereadable		nVARCHAR(4000)					'$.machinereadable'
,[bulk]					nVARCHAR(4000)					'$.bulk'
,openlicense			nVARCHAR(4000)					'$.openlicense'
,uptodate				nVARCHAR(4000)					'$.uptodate'
,[url]					nVARCHAR(4000)					'$.url'
,[format]					nVARCHAR(4000)				'$.format'
,licenseurl				nVARCHAR(4000)					'$.licenseurl'
,officialtitle			nVARCHAR(4000)					'$.officialtitle'
,publisher				nVARCHAR(4000)					'$.publisher'
,reviewed				nVARCHAR(4000)					'$.reviewed'
,reviewResult			nVARCHAR(4000)					'$.reviewResult'
,reviewComments			nVARCHAR(4000)					'$.reviewComments'
,details				nVARCHAR(4000)					'$.details'
,isCurrent				nVARCHAR(4000)					'$.isCurrent'
,isOpen					nVARCHAR(4000)					'$.isOpen'
,submitter				nVARCHAR(4000)					'$.submitter'
,reviewer				nVARCHAR(4000)					'$.reviewer'
,score					nVARCHAR(4000)					'$.score')
