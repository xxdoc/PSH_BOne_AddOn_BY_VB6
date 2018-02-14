USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP035_07]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : M/G공정카드																			*/
/*  Create Date    : 2010.11.19                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--EXEC [PS_PP035_07]
ALTER PROC [dbo].[PS_PP035_07]
--Create PROC [dbo].[PS_PP035_07]
--(
--	@Seq	As Nvarchar(1)
--)
AS
BEGIN
IF OBJECT_ID('Temp_LBG14') IS NULL
	BEGIN
		CREATE TABLE [Temp_LBG14]
		( 
		DocEntry  numeric
		)
	END	

SELECT CONVERT(Nvarchar(10), A.DocEntry) aS DocEntry,
       (CONVERT(Nvarchar(30),D.U_CallSize) + CONVERT(Nvarchar(20),Substring(A.U_OrdNum,5,7))) AS LotNo, 
	   ItemName030L = (select ItemName FROM OITM WHERE ItemCode = B.U_ItemCode),
	   ((CONVERT(NVARCHAR(50),B.U_BatchNum)) + '(' + (CONVERT(CHAR(10),RIGHT(A.U_OrdNum,5))) + ')') AS SoLotNo,
	   ItemName030H  = (select ItemName FROM OITM WHERE ItemCode = A.U_ItemCode),
       A.U_SelWt AS SelWt,
	   A.U_DocDate AS DocDate,
	   CONVERT(Nvarchar(10),A.U_MulGbn1) AS MulGbn1,
     CASE WHEN U_MulGbn1 = '10'
          THEN  '탈지'
          WHEN U_MulGbn1 = '20'
          THEN  '비탈지'
     END AS MulType1,
	    CONVERT(Nvarchar(10),A.U_MulGbn2) AS MulGbn2,     
     CASE WHEN U_MulGbn2 = '10'
          THEN  '시계'
          WHEN U_MulGbn2 = '20'
          THEN  '반시계'
     END AS MulType2,
        CONVERT(Nvarchar(10),A.U_MulGbn3) AS MulGbn3,     
     CASE WHEN U_MulGbn3 = '10'
          THEN  '배면'
          WHEN U_MulGbn3 = '20'
          THEN  '상면'
     END AS MulType3,
		CONVERT(Nvarchar(30),C.U_CpCode) AS CpCode,
	    CpName = (select CONVERT(Nvarchar(50),U_ShortNam) FROM [@PS_PP001L] WHERE U_CpCode = C.U_CpCode),
	    CONVERT(Nvarchar(20),E.U_StdTime) AS StdTime
  FROM [@PS_PP030H] AS A INNER JOIN [@PS_PP030L] AS B
       ON A.DocEntry = B.DocEntry
       INNER JOIN [@PS_PP030M] AS C
       ON A.DocEntry = C.DocEntry
       INNER JOIN OITM AS D	
       ON A.U_ItemCode = D.ItemCode
       INNER JOIN [@PS_PP004H] AS E
       ON A.U_ItemCode = E.U_ItemCode 
       And C.U_CpCode = E.U_CpCode
 WHERE  A.DocEntry in (select DocEntry FROM Temp_LBG14)
 Group by A.DocEntry, D.U_CallSize, A.U_OrdNum, B.U_ItemCode, A.U_SelWt, A.U_DocDate, B.U_BatchNum, A.U_ItemCode,
          A.U_MulGbn1, A.U_MulGbn2, A.U_MulGbn3, C.U_CpCode, E.U_StdTime
End   


-- EXEC [PS_PP035_07] 
  