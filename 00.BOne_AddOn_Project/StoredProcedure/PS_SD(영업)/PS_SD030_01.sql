IF OBJECT_ID('PS_SD030_01') IS NOT NULL
BEGIN
	DROP PROC PS_SD030_01
END
GO

CREATE PROC PS_SD030_01
(
	@OrderNum NVARCHAR(40),
	@DocType NVARCHAR(10)
)
AS
BEGIN
	SELECT
		CONVERT(VARCHAR,ORDR.DocEntry) + '-' + CONVERT(VARCHAR,RDR1.LineNum) AS OrderNum,
		RDR1.ItemCode AS ItemCode,
		OITM.ItemName AS ItemName,
		OITM.ItmsGrpCod AS ItemGpCd,
		OITM.U_ItmBsort AS ItmBsort,
		OITM.U_ItmMsort AS ItmMsort,
		OITM.U_Unit1 AS Unit1,
		OITM.U_Size AS Size,
		OITM.U_ItemType AS ItemType,
		OITM.U_Quality AS Quality,
		OITM.U_Mark AS Mark,
		OITM.U_SbasUnit AS SbasUnit,
		RDR1.U_Qty AS SjQty,
		RDR1.Quantity AS SjWeight,
		RDR1.U_Qty - ISNULL(PS_SD030.Qty,0) AS Qty,
		OITM.U_Unweight AS UnWeight, --����(������ ������ ǥ���ϰ� ������ ǥ������ ����)
		RDR1.Quantity - ISNULL(PS_SD030.Weight,0) AS Weight, --�߷�
		RDR1.Currency AS Currency,
		CASE WHEN RDR1.Currency = 'KRW' THEN RDR1.Price ELSE RDR1.PriceBefDi END AS Price,
		(CASE WHEN RDR1.Currency = 'KRW' THEN RDR1.Price ELSE RDR1.PriceBefDi END) * (RDR1.Quantity - ISNULL(PS_SD030.Weight,0)) AS LinTotal,
		RDR1.WhsCode AS WhsCode,
		OWHS.WhsName AS WhsName,
		'' AS Comments,
		RDR1.U_TrType AS TrType,
		ORDR.DocEntry AS ORDRNum,
		RDR1.LineNum AS RDR1Num,		
		'O' AS Status,
		'' AS LineId
	FROM
		[ORDR] ORDR
		LEFT JOIN [RDR1] RDR1 ON ORDR.DocEntry = RDR1.DocEntry
		LEFT JOIN [OITM] OITM ON RDR1.ItemCode = OITM.ItemCode	
		LEFT JOIN [OWHS] OWHS ON RDR1.WhsCode = OWHS.WhsCode
		LEFT JOIN 
		(SELECT
			PS_SD030L.U_ORDRNum AS ORDRNum,
			PS_SD030L.U_RDR1Num AS RDR1Num,
			SUM(PS_SD030L.U_Qty) AS Qty,
			SUM(PS_SD030L.U_Weight) AS Weight
		FROM
			[@PS_SD030H] PS_SD030H
			LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry
		WHERE
			PS_SD030H.Canceled = 'N'
		GROUP BY
			PS_SD030L.U_ORDRNum,
			PS_SD030L.U_RDR1Num
		) PS_SD030 ON PS_SD030.ORDRNum = ORDR.DocEntry AND PS_SD030.RDR1Num = RDR1.LineNum
	WHERE
		CONVERT(VARCHAR,ORDR.DocEntry) + '-' + CONVERT(VARCHAR,RDR1.LineNum) = @OrderNum
		AND RDR1.Quantity - ISNULL(PS_SD030.Weight,0) > 0
		AND ORDR.Canceled = 'N'
		AND ORDR.DocStatus = 'O'
		AND RDR1.LineStatus = 'O'
		AND (@DocType = 1 OR (@DocType = 2 AND OITM.U_ItmBsort IN('105','106')))--���,����
		AND (@DocType = 1 OR (@DocType = 2 AND ORDR.BPLId = '2')) -- ������ ���������϶�
END

--EXEC PS_SD030_01 '1-0'