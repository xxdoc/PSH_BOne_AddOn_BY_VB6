USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP410_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 지원공수 현황(마이크로텍)																	*/
/*  Create Date    : 2010.11.26                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP410_01]
--Create PROC [dbo].[PS_PP410_01]
(
	@DocDateFr Date,
	@DocDateTo Date
)
AS

BEGIN

SELECT  C.U_OrdGbn			AS OrdGbn,                 --대분류
        E.Name,
		B.U_CpCode			AS CpCode, 
		B.U_CpName			AS CpName,
		C.U_DocDate			AS DocDate,
		D.U_WorkCode		AS WorkCode,
		D.U_WorkName		AS WorkName,
		C.U_OrdNum			AS OrdNum,
		ISNULL(B.U_PQty,0)	AS PQty,
		B.U_WorkTime		AS WorkTime,
		(A.U_PsmtP * B.U_WorkTime) AS PRICE
  FROM [@PS_PP001L] AS A INNER JOIN [@PS_PP040L] AS B
			ON A.U_CpCode = B.U_CpCode
		INNER JOIN [@PS_PP040H] AS C
			ON B.DocEntry = C.DocEntry
		INNER JOIN [@PS_PP040M] AS D
			ON B.DocEntry = D.DocEntry	
		INNER JOIN [@PSH_ITMBSORT] AS E
			ON C.U_OrdGbn = E.Code		
 WHERE C.U_OrdType = '20'
   AND C.U_DocDate BETWEEN @DocDateFr AND @DocDateTo
			
ORDER BY C.U_OrdGbn, B.U_CpCode

END

--EXEC [dbo].[PS_PP410_01] '20101101', '20101130'
