USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP035_01]    Script Date: 11/04/2010 12:55:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 작업지시리스트조회 > 작지조회 SELECT[PS_PP035]				                                */
/*  Create Date    : 2010.10.10                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP035_01]
(
    @BPLId    AS NVARCHAR(20),
    @Canceled AS NVARCHAR(20),
    @OrdNum   AS NVARCHAR(20),
    @OrdGbn   AS NVARCHAR(20),
    @ItemCode AS NVARCHAR(30),
    @WorkDtFr AS NVARCHAR(8),   
    @WorkDtTo AS NVARCHAR(8)
)
AS
BEGIN

	SELECT
		A1.U_BPLId AS BPLId, 
		A1.DocEntry AS DocEntry,
		A1.U_OrdGbn AS OrdGbn,
		A1.U_OrdNum AS OrdNum,
		A1.U_OrdSub1 AS OrdSub1,
		A1.U_OrdSub2 AS OrdSub2,
		A1.Canceled AS Canceled,	
		B1.CardCode AS CardCode,
		B1.CardName AS CardName,	
		A1.U_ItemCode AS ItemCode,	
		A1.U_ItemName AS ItemName,
		A1.U_ReqWt AS ReqWt,	
		A1.U_SelWt AS SelWt,
		A1.U_CntcCode AS CntcCode,	
		A1.U_CntcName AS CntcName,	
		IsNull(A1.U_DocDate, '') AS DocDate,	
		A1.U_DueDate AS DueDate,
		C1.U_Quality AS Quality,
		C1.U_Unweight AS Unweight,	
		D1.U_CpName AS CpName
	FROM
		-- 공통
		(
		SELECT
			A2.U_BPLId,
			A2.DocEntry,
			A2.U_OrdGbn,
			A2.U_OrdNum,
			A2.U_OrdSub1,
			A2.U_OrdSub2,
			A2.Canceled,		
			A2.U_ItemCode,	
			A2.U_ItemName,
			A2.U_ReqWt,	
			A2.U_SelWt,	
			A2.U_CntcCode,
			A2.U_CntcName,	
			A2.U_DocDate,	
			A2.U_DueDate
		FROM
			[@PS_PP030H] AS A2
		) A1 LEFT JOIN 
		
		-- 고객이름
		(
		SELECT
			A2.DocEntry,
			B2.CardCode,
			B2.CardName
		FROM
			[@PS_PP030H] AS A2
			LEFT JOIN [ORDR] AS B2 ON A2.U_SjNum = B2.DocEntry
		) B1 ON A1.DocEntry = B1.DocEntry LEFT JOIN
		
		-- 아이템
		(
		SELECT
			A2.DocEntry,
			B2.U_Quality,
			B2.U_Unweight
		FROM 
			[@PS_PP030H] AS A2
			LEFT JOIN [OITM] AS B2 ON A2.U_ItemCode = B2.ItemCode
		) C1 ON A1.DocEntry = C1.DocEntry LEFT JOIN
		
		-- 첫공정명
		(		
		SELECT
			A2.DocEntry,
			B2.U_CpName				
		FROM
			[@PS_PP030H] AS A2
			LEFT JOIN [@PS_PP030M] AS B2 ON A2.DocEntry = B2.DocEntry
		WHERE
			B2.U_Sequence = '1'	
		) D1 ON A1.DocEntry = D1.DocEntry
	WHERE (A1.U_BPLId = @BPLId OR @BPLId= '')
	  AND (A1.Canceled = @Canceled OR @Canceled = '')
	  AND (A1.U_OrdNum = @OrdNum OR @OrdNum = '')
	  AND (A1.U_OrdGbn = @OrdGbn OR @OrdGbn = '')
	  AND (A1.U_ItemCode = @ItemCode OR @ItemCode = '')
	  AND A1.U_DocDate BETWEEN @WorkDtFr AND @WorkDtTo
 End
-- EXEC PS_PP035_01 '', '', '', '', '', '', '', ''
