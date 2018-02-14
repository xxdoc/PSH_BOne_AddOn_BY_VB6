USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP350_02]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 부품작업일지																	*/
/*  Create Date    : 2010.11.26                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP350_02]
--Create PROC [dbo].[PS_PP350_02]
(
	@DocDateFr Date,
	@DocDateTo Date
)
AS

BEGIN

SELECT	G.ItemCode,
		G.ItemName,
		G.DocDate,
		G.Mark,
		G.CpCode,
		G.CpName,
		G.WorkCode,
		G.WorkName,
		G.INSU,
		G.YQty,
		G.NQty,
		G.WorkTime,
		D.Name       
FROM(

 SELECT	CONVERT(CHAR(20),A.U_ItemCode)		AS ItemCode, 
		CONVERT(CHAR(60),A.U_ItemName)		AS ItemName, 
		A.U_DocDate			AS DocDate,
		Mark = (SELECT U_Mark FROM OITM WHERE ItemCode = A.U_ItemCode),
		CONVERT(CHAR(20),B.U_CpCode)			AS CpCode, 
		CONVERT(CHAR(60),B.U_CpName)			AS CpName,
		CONVERT(CHAR(10),C.U_WorkCode)		AS WorkCode,
		CONVERT(CHAR(10),C.U_WorkName)		AS WorkName,
		(B.U_YQty + B.U_NQty) AS INSU,
		CONVERT(NUMERIC(19,6),B.U_YQty)			AS YQty,
		CONVERT(NUMERIC(19,6),B.U_NQty)			AS NQty,
		CONVERT(NUMERIC(19,6),B.U_WorkTime)		AS WorkTime
  FROM [@PS_PP040H] AS A INNER JOIN [@PS_PP040L] AS B
        ON A.DocEntry = B.DocEntry
	INNER JOIN [@PS_PP040M] AS C
        ON B.DocEntry = C.DocEntry
        ) G
    INNER JOIN [@PSH_MARK] AS D
		ON G.Mark = D.Code
  WHERE G.DocDate BETWEEN @DocDateFr AND @DocDateTo

ORDER BY G.ItemCode, G.CpCode

End		  



--EXEC [dbo].[PS_PP350_02] '20101101', '20101130'

EXEC [PS_PP350_02] '20101101', '20101126'