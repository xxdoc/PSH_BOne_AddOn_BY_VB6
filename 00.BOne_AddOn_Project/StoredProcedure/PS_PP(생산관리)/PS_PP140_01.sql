USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP140_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 제품입고대장 (휘팅-창원)																	*/
/*  Create Date    : 2010.11.23                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP140_01]
--Create PROC [dbo].[PS_PP140_01]
(
	@DocDateFr Date,
	@DocDateTo Date,
	@CardName  Nvarchar(50),
	@ItemMsort Nvarchar(10),
	@ItemName Nvarchar(50),
	@Spec1Fr Nvarchar(10),
	@Spec1To Nvarchar(10),
	@Mark Nvarchar(10),
	@ItemType Nvarchar(10)
)
AS

BEGIN

--SET @DocDateFr = ISNULL(@DocDateFr,'')
--SET @DocDateTo = ISNULL(@DocDateTo,'')
--SET @CardCode = ISNULL(@CardCode,'')
--SET @ItemMsort = ISNULL(@ItemMsort,'')
--SET @ItemName = ISNULL(@ItemName,'')
--SET @Spec1Fr = ISNULL(@Spec1Fr,'')
--SET @Spec1To = ISNULL(@Spec1To,'')
--SET @Mark = ISNULL(@Mark,'')
--SET @ItemType = ISNULL(@ItemType,'')

select @Spec1Fr = CASE When ISNULL(@Spec1Fr,'') = '' Then '0' Else @Spec1Fr  end
select @Spec1To = CASE When ISNULL(@Spec1To,'') = '' Then '9999' Else @Spec1To end

-- 생산완료 
select G.div,
	   SUBSTRING(CONVERT(CHAR(8),G.DocDate,112),1,4) + '-' + SUBSTRING(CONVERT(CHAR(8),G.DocDate,112),6,2) + '-' + SUBSTRING(CONVERT(CHAR(8),G.DocDate,112),7,2)AS DocDate,
	   G.DocEntry,
	   G.ItemCode,
	   G.ItemName,
	   G.OrdNum,
	   G.CardCode,
	   G.CardName,
	   G.YQty,
	   G.YWeight,
	   G.Div
from (

SELECT H.div,
	   H.DocDate,
	   H.DocEntry,
	   H.ItemCode,
	   H.ItemName,
	   H.OrdNum,
	   H.CardCode,
	   H.CardName,
	   H.YQty,
	   H.YWeight
FROM (
SELECT div = '포장',
	   SUBSTRING(CONVERT(CHAR(8),A.U_DocDate,112),1,4) + '-' + SUBSTRING(CONVERT(CHAR(8),A.U_DocDate,112),5,2) + '-' + SUBSTRING(CONVERT(CHAR(8),A.U_DocDate,112),7,2)AS DocDate,
	   Convert(Char(10),A.DocEntry)   AS DocEntry, 
	   CONVERT(Nvarchar(20),B.U_ItemCode) AS ItemCode, 
	   CONVERT(Nvarchar(50),B.U_ItemName) AS ItemName, 
	   CONVERT(CHAR(30),B.U_OrdNum) AS OrdNum,
	   CardCode = (select CardCode FROM ORDR WHERE DocNum = C.U_SjNum),
	   CardName = (select CONVERT(CHAR(50),CardName) FROM ORDR WHERE DocNum = C.U_SjNum),
	   B.U_YQty	   AS YQty,
	   B.U_YWeight AS YWeight
  FROM [@PS_PP080H] AS A INNER JOIN [@PS_PP080L] AS B
			ON A.DocEntry  = B.DocEntry
	   INNER JOIN [@PS_PP030H] AS C
			ON B.U_OrdNum = C.U_OrdNum
	   INNER JOIN OITM AS D
			ON B.U_ItemCode = D.ItemCode 
 WHERE A.U_BPLId = '1'   
   AND D.U_ItmBsort = '101'
   AND A.U_DocDate BETWEEN @DocDateFr AND @DocDateTo
   AND Isnull(D.U_ItmMsort, '') LIKE @ItemMsort
   AND Isnull(B.U_ItemName, '') LIKE @ItemName
   AND Isnull(D.U_Spec1, '') BETWEEN @Spec1Fr AND @Spec1To
   AND Isnull(D.U_Mark, '') LIKE @Mark
   AND Isnull(D.U_ItemType, '') LIKE @ItemType
   ) H
 WHERE isnull(H.CardName,'') LIKE(@CardName)
UNION ALL

-- 벌크포장
SELECT H.div,
	   H.DocDate AS DocDate,
	   H.DocEntry,
	   H.ItemCode,
	   H.ItemName,
	   H.OrdNum,
	   H.CardCode,
	   H.CardName,
	   H.SelQty,
	   H.SelWt
FROM (
SELECT div = '벌크',
	   SUBSTRING(CONVERT(CHAR(8),A.U_DocDate,112),1,4) + '-' + SUBSTRING(CONVERT(CHAR(8),A.U_DocDate,112),5,2) + '-' + SUBSTRING(CONVERT(CHAR(8),A.U_DocDate,112),7,2)AS DocDate,
	   Convert(Char(10),A.DocEntry)   AS DocEntry, 
	   CONVERT(Nvarchar(20),B.U_ItemCode) AS ItemCode, 
	   CONVERT(Nvarchar(50),B.U_ItemName) AS ItemName, 
	   CONVERT(CHAR(30),B.U_OrdNum) AS OrdNum,
	   CardCode = (select CardCode FROM ORDR WHERE DocNum = C.U_SjNum),
	   CardName = (select CONVERT(CHAR(50),CardName) FROM ORDR WHERE DocNum = C.U_SjNum),
	   B.U_SelQty  AS SelQty,
	   B.U_SelWt   AS SelWt
  FROM [@PS_PP070H] AS A INNER JOIN [@PS_PP070L] AS B
			ON A.DocEntry  = B.DocEntry
		INNER JOIN [@PS_PP030H] AS C
			ON A.DocEntry = C.DocEntry
		INNER JOIN OITM AS D
			ON B.U_ItemCode = D.ItemCode
WHERE A.U_BPLId = '1'   
   AND D.U_ItmBsort = '101'
   AND A.U_DocDate BETWEEN @DocDateFr AND @DocDateTo
   AND Isnull(D.U_ItmMsort, '') LIKE @ItemMsort
   AND Isnull(B.U_ItemName, '') LIKE @ItemName
   AND Isnull(D.U_Spec1, '') BETWEEN @Spec1Fr AND @Spec1To
   AND Isnull(D.U_Mark, '') LIKE @Mark
   AND Isnull(D.U_ItemType, '') LIKE @ItemType
  ) H
  WHERE isnull(H.CardName,'') LIKE(@CardName)
   ) G
 
ORDER BY G.DocDate, G.DocEntry

End

--EXEC [dbo].[PS_PP140_01] '20101101','20101231',NULL,NULL,NULL,NULL,NULL,NULL,NULL
--  EXEC [dbo].[PS_PP140_01] '20101101','20101231','%','%','%','%','%','%','%'
  
--  EXEC [dbo].[PS_PP140_01] '20100101', '20101125', '%','%','%','%','%','%','%'

--EXEC [PS_PP140_01] '20101101', '20101125', '한%','%','%', '0000000', 'ZZZZZZZ','%','%'




