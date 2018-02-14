USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP035_06]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : End Bearing 공정카드																				*/
/*  Create Date    : 2010.11.18                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--EXEC [PS_PP035_06]
ALTER PROC [dbo].[PS_PP035_06]
--Create PROC [dbo].[PS_PP035_06]
(
	@Seq	As Nvarchar(1)
)
AS
BEGIN
IF OBJECT_ID('Temp_LBG13') IS NULL
	BEGIN
		CREATE TABLE [Temp_LBG13]
		( 
		DocEntry  numeric
		)
	END	
If @Seq = 'M'

SELECT CONVERT(Nvarchar(10), A.DocEntry) aS DocEntry,
	   CONVERT(NVARCHAR(20),(A.U_OrdNum + '-' + A.U_OrdSub1 + '-' + A.U_OrdSub2)) AS JAKBUN,
	   CONVERT(NVARCHAR(50),Max(C.U_CpName)) AS CpName,
	   CONVERT(NVARCHAR(30),A.U_ItemCode) AS ItemCode030H,
	   ItemName030H  = (select ItemName FROM OITM WHERE ItemCode = A.U_ItemCode),
	   Size030H = (select (CONVERT(CHAR(30),U_Size)) FROM OITM WHERE ItemCode = A.U_ItemCode),
	   CONVERT(NVARCHAR(30),B.U_ItemCode) AS ItemCode030L,
	   ItemName030L = (select ItemName FROM OITM WHERE ItemCode = B.U_ItemCode),
	   Size030L = (select (CONVERT(CHAR(30),U_Size)) from OITM WHERE ItemCode = B.U_ItemCode),
	   A.U_DocDate AS DocDate,
	   A.U_DueDate AS DueDate,
	   CONVERT(NUMERIC(19,3),A.U_SelWt) AS SelWt,
	   CONVERT(NVARCHAR(50),B.U_BatchNum) AS BatchNum,
	   CONVERT(CHAR(10),RIGHT(A.U_OrdNum,5)) AS YMD                                         --LOT - NO 뒤 5자리
  FROM [@PS_PP030H] AS A INNER JOIN [@PS_PP030L] AS B
        ON A.DocEntry = B.DocEntry
        INNER JOIN [@PS_PP030M] AS C
        ON A.DocEntry = C.DocEntry
 WHERE  A.DocEntry in (select DocEntry FROM Temp_LBG13)
 Group by A.DocEntry, A.U_OrdNum, A.U_OrdSub1, A.U_OrdSub2, A.U_ItemCode, B.U_ItemCode, A.U_DocDate, A.U_DueDate, A.U_SelWt,
		  B.U_BatchNum, A.U_OrdNum
       
If @Seq = 'S'

SELECT  CONVERT(Nvarchar(10), A.DocEntry) aS DocEntry,
		CONVERT(CHAR(30),B.U_CpCode) AS CpCode,
	    CONVERT(CHAR(50),B.U_CpName) AS CpName
  FROM [@PS_PP030H] AS A INNER JOIN [@PS_PP030M] AS B
        ON A.DocEntry = B.DocEntry
 WHERE  A.DocEntry in (select DocEntry FROM Temp_LBG13)

End   

--EXEC [PS_PP035_06] 'M'

--EXEC [PS_PP035_06] 'S'

--EXEC [PS_PP035_06] 'M'