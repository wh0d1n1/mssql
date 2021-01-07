
DECLARE @JSON VARCHAR(MAX)

SELECT @JSON = BulkColumn
FROM OPENROWSET 
(BULK 'C:\Users\Travis Padilla\Documents\cmsdata.json', SINGLE_CLOB) 
AS j



Select reports.landingpage
,reports.keyword
,reports.[description]
,reports.identifier
,reports.title
,downloadspecial.*
,downloadspecial2.*
--,downloads1.[VALUE] [Text/Csv]
--,downloads2.[value] 
--,downloads3.[value] 
--,downloads4.[value] 
,case

FROM OPENJSON (@JSON, '$.dataset') 
with(
	accessLevel nvarchar(max) 
	, landingPage nvarchar(max) '$.landingPage'
	, bureauCode nvarchar(max) '$.bureauCode'
	, issued nvarchar(max) '$.issued'
	, modified nvarchar(max) '$.modified'
	, keyword nvarchar(max) '$.keyword'
	, identifier nvarchar(max) '$.identifier'
	, [description] nvarchar(max) '$.description'
	, title nvarchar(max) '$.title'
	, programCode nvarchar(max) '$.programCode'
	,distribution nvarchar(max) '$.distribution' as json
	, Downloadurl nvarchar(max) '$.distribution[0]' as json
	, Downloadurl2 nvarchar(max) '$.distribution[1]' as json
	, Downloadurl3 nvarchar(max) '$.distribution[2]' as json
	, Downloadurl4 nvarchar(max) '$.distribution[3]' as json
	--, theme nvarchar(max) '$.theme'
	--, license nvarchar(max) '$.license'
	--, systemOfRecords nvarchar(max) '$.systemOfRecords'
	--, accrualPeriodicity nvarchar(max) '$.accrualPeriodicity'
	--,temporal nvarchar(max) '$.temporal'
	) as reports
	cross apply openjson(reports.distribution, '$') 
	with (downloadspecial nvarchar(max) '$' as json) as downloadspecial
	cross apply openjson(downloadspecial.downloadspecial) as downloadspecial2
	--with( downoadurl nvarchar(max) )											
	--with (downloadurlspec nvarchar(max)) as downloadspecial2
	--cross apply openjson(reports.Downloadurl, '$') 
	--as downloads1
	--cross apply openjson(reports.Downloadurl2, '$') 
	--as downloads2
	--cross apply openjson(reports.Downloadurl3, '$') 
	--as downloads3
	--cross apply openjson(reports.Downloadurl4, '$') 
	--as downloads4
where 
downloadspecial2.[value] LIKE '%http%'
--and
--downloads2.[value] LIKE '%http%'
--and
--downloads3.[value] LIKE '%http%'
--and
--downloads4.[value] LIKE '%http%'


