IF OBJECT_ID('PS_SD380_10') IS NOT NULL
BEGIN
	DROP PROC PS_SD380_10
END
GO
--EXEC PS_SD380_10 '20101100001','20101116001','Y'
CREATE PROC PS_SD380_10
(
	@FrPackNo NVARCHAR(100),
	@ToPackNo NVARCHAR(100),
	@CardType NVARCHAR(10)--거래처타입 'Y' = 품목배치번호 'N' = 원소재배치번호
)
AS
BEGIN
	SELECT
		CONVERT(NVARCHAR,@FrPackNo) AS FrPackNo,
		CONVERT(NVARCHAR,@ToPackNo) AS ToPackNo,
		CONVERT(NVARCHAR,PS_SD040L.U_ItemCode) AS ItemCode,
		CONVERT(NVARCHAR,PS_SD040L.U_ItemName) AS ItemName,
		CONVERT(NVARCHAR,OITM.U_Size) AS Size,
		CONVERT(NVARCHAR,PS_SD040L.U_PackNo) AS PackNo,
		--원소재배치번호7자리 + 창원:S 동래:V + 품목호칭 + 
		CONVERT(NVARCHAR,ISNULL((SELECT LEFT(U_BatchNum,7) FROM [@PS_PP030L] WHERE DocEntry = PS_PP030H.DocEntry),'')
		+ '-' + CASE WHEN PS_SD040H.U_BPLId = '1' THEN 'S' WHEN PS_SD040H.U_BPLId = '2' THEN 'V' END
		+ OITM.U_CallSize + SUBSTRING(PS_SD040L.U_LotNo,5,LEN(PS_SD040L.U_LotNo)-4)) AS Info01, --품목호칭,작지
		CONVERT(NVARCHAR,CASE WHEN PS_SD040H.U_BPLId = '1' THEN 'S' WHEN PS_SD040H.U_BPLId = '2' THEN 'V' END
		+ OITM.U_CallSize 
		+ CASE WHEN @CardType = 'Y' THEN PS_SD040L.U_LotNo
		WHEN @CardType = 'N' THEN (SELECT U_BatchNum FROM [@PS_PP030L] WHERE DocEntry = PS_PP030H.DocEntry) END) AS Info02,
		PS_SD040L.U_Weight AS Weight
		--PS_SD040L.U_LotNo AS BatchNum
	FROM
		[@PS_SD040H] PS_SD040H
		LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry
		LEFT JOIN [@PS_PP030H] PS_PP030H ON PS_SD040L.U_LotNo = PS_PP030H.U_OrdNum
		LEFT JOIN [OITM] OITM ON PS_SD040L.U_ItemCode = OITM.ItemCode
	WHERE
		PS_SD040L.U_PackNo >= @FrPackNo
		AND PS_SD040L.U_PackNo <= @ToPackNo
		AND PS_SD040H.Canceled = 'N'
		AND PS_SD040H.U_BPLId IN('1','2') --창원,동래
		AND (SELECT U_ItmBsort FROM OITM WHERE ItemCode = PS_SD040L.U_ItemCode) IN('104') --멀티
		AND ISNULL(PS_SD040L.U_PackNo,'') <> '' --포장번호가 있어야함(멀티)
	
END