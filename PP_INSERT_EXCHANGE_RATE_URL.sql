/*
sp_configure 'show advanced options', 1;
GO
RECONFIGURE;
GO
sp_configure 'Ole Automation Procedures', 1;
GO
RECONFIGURE;
GO
*/

DECLARE @OBJECT INT
DECLARE @RESPONSETEXT VARCHAR(8000)
DECLARE @URL VARCHAR(300)
DECLARE @XML AS TABLE(MYXML XML)

SET @URL = 'https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx'

EXEC sp_OACreate 'MSXML2.XMLHTTP', @OBJECT OUT; 
EXEC sp_OAMethod @OBJECT, 'open', NULL, 'get', @URL, 'false' 
EXEC sp_OAMethod @OBJECT, 'send' 
EXEC sp_OAMethod @OBJECT, 'responseText', @RESPONSETEXT OUTPUT      
EXEC sp_OADestroy @OBJECT 

INSERT @XML
SELECT @RESPONSETEXT

SELECT DateTime.value('(.)[1]', 'DATETIME2(0)') AS DATETIME,
      Exrate.value('@CurrencyCode', 'NVARCHAR(100)') AS CURRENCY_CODE,
      Exrate.value('@CurrencyName', 'NVARCHAR(100)') AS CURRENCY_NAME
	, CAST(REPLACE(REPLACE(Exrate.value('@Buy', 'NVARCHAR(50)'), '-', ''), ',', '') AS FLOAT) AS BUY
	, CAST(REPLACE(REPLACE(Exrate.value('@Transfer', 'NVARCHAR(50)'), '-', ''), ',', '') AS FLOAT) AS TRANSFER
	, CAST(REPLACE(REPLACE(Exrate.value('@Sell', 'NVARCHAR(50)'), '-', ''), ',', '') AS FLOAT) AS SELL
FROM @xml
CROSS APPLY [MyXml].nodes('ExrateList') AS MyXml(ExrateList)
OUTER APPLY ExrateList.nodes('Exrate') AS ExrateList(Exrate)
CROSS APPLY [MyXml].nodes('/ExrateList/DateTime') AS Datetime(DateTime)