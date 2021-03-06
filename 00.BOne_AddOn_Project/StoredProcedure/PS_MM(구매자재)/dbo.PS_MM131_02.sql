USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM131_02]    Script Date: 11/04/2010 12:53:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : 재고관리 > 외주반출등록(재공반출)[PS_MM131_02]				                            */
/*  Create Date    : 2010.10.18                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_MM131_02]
(
	@ItemCode AS NVARCHAR(20)
)
AS
BEGIN
SELECT 	PP030H.DocEntry AS PP030HNo,
		PP030H.U_OrdNum AS OrdNum,													--작지번호
		PP030H.U_ItemCode AS ItemCode,												--품목코드
		PP030H.U_ItemName AS ItemName,												--품목명
		OITM.U_Size AS Size,														--규격
		ISNULL(OITM.U_UnWeight, 1) AS UnWeight,										--단중
		ISNULL(PP040L.U_PQty, 0) AS SelQty,											--작지수량		
		ISNULL(PP040L.U_PQty, 0) * ISNULL(OITM.U_UnWeight, 1) AS SelWt,				--작지중량
		SUM(ISNULL(MM130L.U_OutQty, 0)) AS OutQty,									--기출고수량
		SUM(ISNULL(MM130L.U_OutWt, 0))	AS OutWt,									--기출고수량
		OITM.DfltWH AS OutWhCd,														--출고창고
		OWHS.WhsName AS OutWhNm,													--출고창고명
		Mark.Name AS Mark,															--인증기호
		'' AS InWhCd,																--입고창고
		'' AS InWhNm,																--입고창고명													
		ISNULL(PP040L.U_PQty, 0) - SUM(ISNULL(MM130L.U_OutQty, 0)) AS PosQty,		--출고가능수량	
		--(ISNULL(PP030H.U_SelWt, 0) * ISNULL(OITM.U_UnWeight, 0)) -
		--SUM(ISNULL(MM130L.U_OutWt, 0)) AS PosWt,									--출고가능중량
		MIN(PP030M.U_Sequence) + 1 AS PP030MNo											--작지라인
  FROM   
		 [@PS_PP030M] AS PP030M 
		 INNER JOIN [@PS_PP030H] AS PP030H ON PP030M.DocEntry = PP030H.DocEntry
		 LEFT JOIN [@PS_PP040L] AS PP040L ON PP030M.DocEntry = PP040L.U_PP030HNo AND PP030M.LineId = PP040L.U_PP030MNo 
		 LEFT JOIN [OITM] AS OITM ON PP030H.U_ItemCode = OITM.ItemCode 
		 LEFT JOIN [OWHS] AS OWHS ON OITM.DfltWH = OWHS.WhsCode 
		 LEFT JOIN [@PSH_MARK] AS MARK ON OITM.U_Mark = MARK.Code  
		 LEFT JOIN [@PS_MM130L] AS MM130L ON MM130L.U_PP030HNo = PP030M.DocEntry AND  MM130L.U_PP030MNo = PP030M.U_Sequence
  WHERE  CONVERT(Nvarchar(10), PP030M.DocEntry ) + '-' + CONVERT(NVARCHAR(10), PP030M.U_Sequence ) IN
	(SELECT  CONVERT(NVARCHAR(10), DocEntry ) + '-' + CONVERT(NVARCHAR(10), MIN(U_Sequence) - 1 )
 	   FROM  [@PS_PP030M]
	  WHERE  DocEntry NOT IN (SELECT DocEntry FROM [@PS_PP030M] WHERE  U_Sequence = '1' AND U_WorkGbn = '30')
   	    AND  (U_Sequence <> '1' And U_WorkGbn = '30') 
   GROUP BY  DocEntry) 
   AND   PP030H.U_OrdGbn <> '10'
   AND   (OITM.ItemCode = @ItemCode OR @ItemCode = '')
GROUP BY PP030H.DocEntry, PP030H.U_OrdNum, PP030H.U_ItemCode, PP030H.U_ItemName, PP040L.U_PQty,
		 OITM.U_Size, OITM.DfltWH, OITM.U_UnWeight, OWHS.WhsName, MARK.Name, PP030M.U_Sequence
END

-- EXEC PS_MM131_02 ''