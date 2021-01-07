Declare @JSON varchar(max)

SELECT @JSON = BulkColumn
FROM OPENROWSET (BULK 'C:\Users\Travis Padilla\Documents\download.json', SINGLE_CLOB) as j

SELECT * FROM OPENJSON (@JSON) 
With (Results Varchar(4000) '$.Results.taxonomies',
		Taxonomies			Varchar(1000) '$.Taxonomies',
		TAXONOMY_GROUP		varchar(4000) '$.Results.TAXONOMIES."TAXONOMY_GROUP"',
		license				Varchar(4000) '$.Results.TAXONOMIES."license"',
		[Primary]			varchar(4000) '$.Results.TAXONOMIES."primary"',
		[state]				varchar(4000) '$.Results.TAXONOMIES."state"',
		code				varchar(4000) '$.Results.TAXONOMIES."code"'
)as Dataset;