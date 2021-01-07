--Declare @JSON varchar(max)

--SELECT @JSON = BulkColumn
--FROM OPENROWSET (BULK 'C:\Users\Travis Padilla\Documents\cmsdata.json', SINGLE_CLOB) as j

--SELECT * FROM OPENJSON (@JSON, '$.dataset') 
--with(
--		taxonomy_group varchar(255) '$.taxonomies.taxonomy_group',
--		license		   varchar(255) '$.Order.Date',  
--        Customer	   varchar(200) '$.AccountNumber',  
--        Quantity       int          '$.Item.Quantity' )

EXEC writeFile @fileName = 'C:\SQL_DATA\test.txt', 
               @fileContents = 'this is a test
some more text
go
go
even more';


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
--SELECT @JSON=BulkColumn --FROM OPENROWSET (BULK 'C:\sqlshack\Results.JSON', SINGLE_CLOB) import
SELECT * INTO #JSONTable
FROM OPENJSON (@JSON)

SELECT * from #JSONTable
