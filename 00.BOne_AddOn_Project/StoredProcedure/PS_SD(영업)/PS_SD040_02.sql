IF OBJECT_ID('PS_SD040_02') IS NOT NULL
BEGIN
	DROP PROC PS_SD040_02
END
GO
--EXEC PS_SD040_01 '2-1',1
CREATE PROC PS_SD040_02
(
	@CardCode NVARCHAR(100),
	@BPLId NVARCHAR(1),
	@TradeType NVARCHAR(1),
	@DocType NVARCHAR(1)
)
AS
BEGIN
	SELECT
		CONVERT(VARCHAR,PS_SD030H.DocEntry) + '-' + CONVERT(VARCHAR,PS_SD030L.LineId) AS SD030Num,
		CONVERT(VARCHAR,PS_SD030L.U_OrderNum) AS OrderNum,
		PS_SD030L.U_ItemCode AS ItemCode,
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
		PS_SD030L.U_Qty AS SjQty,
		PS_SD030L.U_Weight AS SjWeight,
		PS_SD030L.U_Qty - ISNULL(PS_SD040.Qty,0) AS Qty,
		OITM.U_Unweight AS UnWeight, --����(������ ������ ǥ���ϰ� ������ ǥ������ ����)
		PS_SD030L.U_Weight - ISNULL(PS_SD040.Weight,0) AS Weight, --�߷�
		PS_SD030L.U_Currency AS Currency,
		PS_SD030L.U_Price AS Price,
		PS_SD030L.U_Price * (PS_SD030L.U_Weight - ISNULL(PS_SD040.Weight,0)) AS LinTotal,
		PS_SD030L.U_WhsCode AS WhsCode,
		OWHS.WhsName AS WhsName,
		PS_SD030L.U_Comments AS Comments,
		PS_SD030L.DocEntry AS SD030H,
		PS_SD030L.LineId AS SD030L,
		PS_SD030L.U_TrType AS TrType,
		PS_SD030L.U_ORDRNum AS ORDRNum,
		PS_SD030L.U_RDR1Num AS RDR1Num,
		'O' AS Status,
		'' AS LineId,
		CASE WHEN PS_SD030H.U_TrType = '1' THEN '' ELSE (SELECT ORDR.U_LotNo FROM [ORDR]ORDR WHERE DocEntry = PS_SD030L.U_ORDRNum) END AS LotNo		
	FROM
		[@PS_SD030H] PS_SD030H
		LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry
		LEFT JOIN [OITM] OITM ON PS_SD030L.U_ItemCode = OITM.ItemCode	
		LEFT JOIN [OWHS] OWHS ON PS_SD030L.U_WhsCode = OWHS.WhsCode
		LEFT JOIN 
		(SELECT
			PS_SD040L.U_SD030H AS SD030HNum,
			PS_SD040L.U_SD030L AS SD030LNum,
			SUM(PS_SD040L.U_Qty) AS Qty,
			SUM(PS_SD040L.U_Weight) AS Weight
		FROM
			[@PS_SD040H] PS_SD040H
			LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry
		WHERE
			PS_SD040H.Canceled = 'N'
		GROUP BY
			PS_SD040L.U_SD030H,
			PS_SD040L.U_SD030L
		) PS_SD040 ON PS_SD040.SD030HNum = PS_SD030H.DocEntry AND PS_SD040.SD030LNum = PS_SD030L.LineId
	WHERE
		PS_SD030H.U_CardCode = @CardCode
		AND PS_SD030H.U_BPLId = @BPLId
		AND PS_SD030H.U_TrType = @TradeType
		AND PS_SD030L.U_Weight - ISNULL(PS_SD040.Weight,0) > 0
		AND PS_SD030H.Canceled = 'N'
		AND PS_SD030H.Status = 'O'
		and PS_SD030H.U_DocType = '1'--���Ͽ�û
		AND PS_SD030H.U_ProgStat = '1' --���Ͽ�û
		--AND PS_SD030L.U_Status = 'O'
		
END

--EXEC PS_SD040_01 '1-0'