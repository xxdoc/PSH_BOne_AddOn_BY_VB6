IF OBJECT_ID('PS_SD040_05') IS NOT NULL
BEGIN
	DROP PROC PS_SD040_05
END
GO
CREATE PROC PS_SD040_05
(
	@BPLId NVARCHAR(100),
	@DCardCod NVARCHAR(100)
)
AS
BEGIN
	SELECT 
		PS_PP090H.U_PackNo AS PackNo,
		PS_PP090L.U_LotNo AS LotNo,
		PS_PP090L.U_ItemCode AS ItemCode,
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
		PS_PP090L.U_Qty AS SjQty,
		PS_PP090L.U_Weight AS SjWeight,
		PS_PP090L.U_Qty AS Qty, --선택수량
		PS_PP090L.U_Weight AS Weight, --선택중량
		'KRW' AS Currency,
		0 AS Price,
		0 AS LineTotal,
		OITM.DfltWH AS WhsCode, --입고창고
		(SELECT WhsName FROM [OWHS] WHERE WhsCode = OITM.DfltWH) AS WhsName,
		'O' AS Status,
		'' AS LineId		
	FROM 
		[@PS_PP090H] PS_PP090H
		LEFT JOIN [@PS_PP090L] PS_PP090L ON PS_PP090H.DocEntry = PS_PP090L.DocEntry
		LEFT JOIN [@PS_QM020H] PS_QM020H ON PS_PP090L.U_LotNo = PS_QM020H.U_LotNo AND PS_PP090L.U_ItemCode = PS_QM020H.U_ItemCode
		LEFT JOIN [OITM] OITM ON PS_PP090L.U_ItemCode = OITM.ItemCode		
	WHERE
		PS_PP090H.Canceled = 'N'
		AND PS_PP090H.U_BPLId = @BPLId
		--AND PS_QM020H.U_CardCode = @DCardCod		
		--AND PS_QM020H.DocEntry IS NOT NULL
		AND (SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040L.U_PackNo = PS_PP090H.U_PackNo AND PS_SD040H.Canceled = 'N') <= 0 
END