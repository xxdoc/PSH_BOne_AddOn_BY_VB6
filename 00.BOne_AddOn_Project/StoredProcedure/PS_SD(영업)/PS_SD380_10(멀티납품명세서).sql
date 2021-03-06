USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_SD380_10]    Script Date: 03/08/2011 08:05:10 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--EXEC PS_SD380_10 '20110208001','20110208001'
ALTER PROC [dbo].[PS_SD380_10]
(
	@FrPackNo NVARCHAR(100),
	@ToPackNo NVARCHAR(100)
	--@CardType NVARCHAR(10)--거래처타입 'Y' = 품목배치번호 'N' = 원소재배치번호
)
AS
BEGIN
Select	Convert(Nvarchar(20), a.U_PackNo) As PackNo
		,Convert(Nvarchar(20), b.U_ItemCode) As ItemCode
		,Convert(Nvarchar(100), z.FrgnName) As ItemName
        ,Convert(Nvarchar(20), e.U_CardName) As CardName
		,Convert(Nvarchar(100), z.U_Size) As Size
		,Convert(Nvarchar(100),k.U_BatchNum) AS BatchNum
		,Convert(Nvarchar,(Case When a.U_BPLId = '1' and LEFT(b.U_LotNo,'6') < '201103' Then 'S' + z.U_CallSize + (select DS.lotno FROM Z_DSMDFRY DS Where DS.custlotno = k.U_BatchNum)
		     When a.U_BPLId = '1' and LEFT(b.U_LotNo,'6') >= '201103' Then 'S' + z.U_CallSize + b.U_Lotno
		     When a.U_BPLId = '2' Then 'V' + z.U_CallSize + b.U_LotNo
		End)) As Info02
		--,CONVERT(NVARCHAR,CASE WHEN a.U_BPLId = '1' THEN 'S' WHEN a.U_BPLId = '2' THEN 'V' END
		--+ z.U_CallSize + b.U_LotNo) AS Info02
		--,a.U_InDate As PackDate
		,b.U_Weight As Weight
		
  From	[@PS_PP090H] a 
		Inner Join [@PS_PP090L] b On a.DocEntry = b.DocEntry
		--Inner Join [Z_PS_PP091] c On a.U_PackNo = c.PackNo
		--Inner Join [OBPL] d On d.BPLId = a.U_BPLId
		Inner Join [@PS_QM020H] e On e.U_ItemCode = b.U_ItemCode And e.U_OrdNum = b.U_LotNo
		Inner Join [@PS_PP030H] f On f.U_OrdNum = b.U_LotNo
		Inner Join [@PS_PP030L] k On f.DocEntry = k.DocEntry
		Inner Join [OITM] z On z.ItemCode = b.U_ItemCode
		Inner Join (Select	a.DocEntry, Sum(b.U_Qty) As SumQty, Sum(b.u_Weight) As NetWt
					  From	[@PS_PP090H] a Inner Join [@PS_PP090L] b On a.DocEntry = b.DocEntry
					Group by a.DocEntry) g On g.DocEntry = a.DocEntry
 where a.U_PackNo >= @FrPackNo
   AND a.U_PackNo <= @ToPackNo		
   AND a.Canceled = 'N'
   AND a.U_BPLId IN('1','2') --창원,동래
   AND b.U_ItmBsort = '104'
   AND ISNULL(a.U_PackNo,'') <> '' --포장번호가 있어야함(멀티)

Order by PackNo, b.U_LineNum

--select * from [@PS_PP090H] a Inner Join [@PS_PP090L] b On a.DocEntry = b.DocEntry
--where a.U_PackNo Between '20110125001' and '20110125007'

	--SELECT
	--	CONVERT(NVARCHAR,@FrPackNo) AS FrPackNo,
	--	CONVERT(NVARCHAR,@ToPackNo) AS ToPackNo,
	--	CONVERT(NVARCHAR,PS_SD040L.U_ItemCode) AS ItemCode,
	--	--CONVERT(NVARCHAR,PS_SD040L.U_ItemName) AS ItemName,
	--	CONVERT(NVARCHAR,OITM.FrgnName) AS ItemName,
	--	CONVERT(NVARCHAR,OITM.U_Size) AS Size,
	--	CONVERT(NVARCHAR,PS_SD040L.U_PackNo) AS PackNo,
	--	--원소재배치번호7자리 + 창원:S 동래:V + 품목호칭 + 
	--	--CONVERT(NVARCHAR,ISNULL((SELECT LEFT(U_BatchNum,7) FROM [@PS_PP030L] WHERE DocEntry = PS_PP030H.DocEntry),'')
	--	--+ '-' + CASE WHEN PS_SD040H.U_BPLId = '1' THEN 'S' WHEN PS_SD040H.U_BPLId = '2' THEN 'V' END
	--	--+ OITM.U_CallSize + SUBSTRING(PS_SD040L.U_LotNo,5,LEN(PS_SD040L.U_LotNo)-4)) AS Info01, --품목호칭,작지
	--	CONVERT(NVARCHAR,CASE WHEN PS_SD040H.U_BPLId = '1' THEN 'S' WHEN PS_SD040H.U_BPLId = '2' THEN 'V' END
	--	+ OITM.U_CallSize 
	--	+ CASE WHEN @CardType = 'Y' THEN PS_SD040L.U_LotNo
	--	WHEN @CardType = 'N' THEN (SELECT U_BatchNum FROM [@PS_PP030L] WHERE DocEntry = PS_PP030H.DocEntry) END) AS Info02,
	--	PS_SD040L.U_Weight AS Weight
	--	--PS_SD040L.U_LotNo AS BatchNum
	--FROM
	--	[@PS_SD040H] PS_SD040H
	--	LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry
	--	LEFT JOIN [@PS_PP030H] PS_PP030H ON PS_SD040L.U_LotNo = PS_PP030H.U_OrdNum
	--	left Join [@PS_PP030L] PS_PP030L On PS_PP030H.DocEntry = PS_PP030L.DocEntry
	--	LEFT JOIN [OITM] OITM ON PS_SD040L.U_ItemCode = OITM.ItemCode
	--WHERE
	--	PS_SD040L.U_PackNo >= @FrPackNo
	--	AND PS_SD040L.U_PackNo <= @ToPackNo
	--	AND PS_SD040H.Canceled = 'N'
	--	AND PS_SD040H.U_BPLId IN('1','2') --창원,동래
	--	AND (SELECT U_ItmBsort FROM OITM WHERE ItemCode = PS_SD040L.U_ItemCode) IN('104') --멀티
	--	AND ISNULL(PS_SD040L.U_PackNo,'') <> '' --포장번호가 있어야함(멀티)
	
END


--select * from [@PS_PP090L]