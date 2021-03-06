USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM131_01]    Script Date: 11/04/2010 12:53:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : 재고관리 > 외주반출등록(원재료반출)[PS_MM131_01]				                            */
/*  Create Date    : 2010.10.14                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_MM131_01]
(
	@ItemCode AS NVARCHAR(20)
)
AS
BEGIN
SELECT
	PP030H.U_OrdNum AS OrdNum,		--작지번호
	PP030H.U_ItemCode AS ItemCode,											--품목코드
	PP030H.U_ItemName AS ItemName,											--품목이름
	PP030H.U_Size AS Size,													--규격
	PP030H.UnWeight AS UnWeight,											--단중
	--PP030H.SelWt AS SelQty,													--작지수량	
	--PP030H.SelWt * PP030H.UnWeight AS SelWt,								--작지중량
	'' AS OutQty,
	'' AS OutWt,
	PP030H.DfltWH AS OutWhCd,
	PP030H.WhsName AS OutWhNm,
	PP030H.U_Mark AS Mark,
	'904' AS InWhCd,
	'외주처-원재료' AS InWhNm,													
	PP030H.OnHand - PP030M.OutQty AS PosQty,								--출고가능수량	
	(PP030H.OnHand * PP030H.UnWeight) - PP030M.OutWt AS PosWt,				--출고가능중량
	PP030M.DocEntry AS PP030HNo,
	PP030M.U_Sequence AS PP030MNo
FROM																		
	(
	SELECT
		PP030H.DocEntry,
		PP030H.U_OrdNum,
		PP030H.U_OrdSub1,
		PP030H.U_OrdSub2,
		PP030H.U_ItemCode,
		PP030H.U_ItemName,
		OITM.U_Size,
		OITM.DfltWH,
		OWHS.WhsName,
		OITM.U_Mark,
		SUM(ISNULL(OITM.OnHand, 0)) AS OnHand,
		SUM(ISNULL(OITM.U_UnWeight, 1)) AS UnWeight,
		SUM(ISNULL(PP030H.U_SelWt, 0)) AS SelWt
	FROM
		[@PS_PP030H] AS PP030H INNER JOIN
		[OITM] AS OITM ON PP030H.U_ItemCode = OITM.ItemCode INNER JOIN
		[OWHS] AS OWHS ON OITM.DfltWH = OWHS.WhsCode
	WHERE 
		PP030H.U_BPLId = '1' AND
	 	(OITM.ItemClass = @ItemCode OR @ItemCode = '') 
	 --AND
		--PP030H.U_DocDate =
		
	GROUP BY PP030H.DocEntry, PP030H.U_OrdNum, PP030H.U_OrdSub1, PP030H.U_OrdSub2, PP030H.U_ItemCode, 
			 PP030H.U_ItemName, OITM.U_Size, OITM.DfltWH, OWHS.WhsName, OITM.U_Mark
			
	) PP030H INNER JOIN
	(
	SELECT
		PP030M.DocEntry,
		PP030M.U_Sequence,
		SUM(ISNULL(MM130L.U_OutQty, 0)) AS OutQty,
		SUM(ISNULL(MM130L.U_OutWt, 0))	AS OutWt
	FROM 
		[@PS_PP030M] AS PP030M LEFT JOIN
		[@PS_MM130L] AS MM130L ON PP030M.DocEntry = MM130L.U_PP030HNo AND PP030M.U_Sequence = MM130L.U_PP030Mno
	WHERE PP030M.U_WorkGbn = '30' AND
		  PP030M.U_Sequence = '1' 		  
	GROUP BY 
		PP030M.DocEntry, PP030M.U_Sequence	  
	) PP030M ON PP030H.DocEntry = PP030M.DocEntry

 END
 
-- EXEC PS_MM131_01 ''

