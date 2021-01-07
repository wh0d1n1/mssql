
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
