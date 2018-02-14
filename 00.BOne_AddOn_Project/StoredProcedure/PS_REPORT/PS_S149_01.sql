IF OBJECT_ID('PS_S149_01') IS NOT NULL
BEGIN
	DROP PROC PS_S149_01
END
GO
--EXEC PS_S149_01 '1'
CREATE PROC PS_S149_01
(
	@DocEntry INT
)
AS
BEGIN
	SELECT
		CONVERT(NVARCHAR,OQUT.DocEntry) AS OQUTNo,
		CONVERT(NVARCHAR,OQUT.DocDate,111) AS OQUTDate,
		OBPL.BPLName AS BPLId,
		(SELECT CompnyName FROM OADM) AS CompanyName,
		(SELECT Manager FROM OADM) AS Manager,
		(SELECT CompnyAddr FROM OADM) AS CompanyAddress,
		(SELECT Phone1 FROM [OADM]) AS Phone,
		(SELECT Fax FROM [OADM]) AS Fax,
		CONVERT(NVARCHAR,ROW_NUMBER() OVER(ORDER BY OQUT.DocEntry)) AS RowNumber,
		QUT1.ItemCode AS ItemCode,
		QUT1.Dscription AS ItemName,
		OITM.SalUnitMsr AS Unit,
		QUT1.Quantity AS Quantity,
		CASE WHEN OQUT.CurSource IN('L','S')
		THEN QUT1.Price
		WHEN OQUT.CurSource IN('C')
		THEN 
			CASE WHEN OQUT.DocCur = 'KRW'
			THEN
				QUT1.Price
			ELSE
				QUT1.PriceBefDi
			END
		END AS Price,		
		CASE WHEN OQUT.CurSource IN('L','S') --현지,시스템
		THEN QUT1.LineTotal
		WHEN OQUT.CurSource IN('C') --BP
		THEN 
			CASE WHEN OQUT.DocCur = 'KRW'
			THEN
				QUT1.LineTotal
			ELSE
				QUT1.TotalFrgn
			END
		END AS LineTotal,
		CONVERT(NVARCHAR,QUT1.U_Note) AS Note,
		(SELECT SUM(QUT1.Quantity)
		 FROM [QUT1] WHERE DocEntry = OQUT.DocEntry) AS TQuantity,
		(SELECT CASE WHEN OQUT.CurSource IN('L','S') THEN SUM(QUT1.LineTotal) WHEN OQUT.CurSource IN('C') THEN CASE WHEN OQUT.DocCur = 'KRW' THEN SUM(QUT1.LineTotal) ELSE SUM(QUT1.TotalFrgn) END END
		 FROM [QUT1] WHERE DocEntry = OQUT.DocEntry) AS TLineTotal				
	FROM
		[OQUT] OQUT
		LEFT JOIN [QUT1] QUT1 ON OQUT.DocEntry = QUT1.DocEntry
		LEFT JOIN [OBPL] OBPL ON OQUT.U_BPLId = OBPL.BPLId		
		LEFT JOIN [OITM] OITM ON QUT1.ItemCode = OITM.ItemCode		
	WHERE
		OQUT.DocEntry = @DocEntry
END
