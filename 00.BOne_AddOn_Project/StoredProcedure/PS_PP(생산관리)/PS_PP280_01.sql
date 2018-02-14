USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP280_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 단중비교현황(휘팅)																	*/
/*  Create Date    : 2010.11.30                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP280_01]
--Create PROC [dbo].[PS_PP280_01]
(
	@DocDateFr Date,
	@DocDateTo Date
)
AS

BEGIN

SELECT	CONVERT(NVARCHAR(30),B.U_ItemCode) AS ItemCode, 
		CONVERT(NVARCHAR(70),B.U_ItemName) As ItemName, 
		SUM(CONVERT(NUMERIC(9,2),C.U_UnWeight)) AS UnWeight,
		SUM(CONVERT(NUMERIC(9,2),U_PQty)) AS SENGSU,
		SUM(CONVERT(NUMERIC(9,2),U_Pweight)) AS SENGJUNG,
		SUM(CONVERT(NUMERIC(9,2),U_Pweight)/CONVERT(NUMERIC(9,2),U_PQty)) AS POSIL	--포장실단중 
  FROM [@PS_PP080H] AS A INNER JOIN [@PS_PP080L] AS B
			ON A.DocEntry = B.DocEntry
        INNER JOIN OITM AS C
			ON B.U_ItemCode = C.ItemCode
 WHERE  A.U_DocDate BETWEEN @DocDateFr AND @DocDateTo
   AND  C.U_ItmBsort = '101'
 
 GROUP BY CONVERT(NVARCHAR(30),B.U_ItemCode),CONVERT(NVARCHAR(70),B.U_ItemName)
 
order by ItemCode

End

--EXEC [dbo].[PS_PP280_01] '20101101', '20101130'