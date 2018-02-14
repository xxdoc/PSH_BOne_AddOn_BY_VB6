USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP035_03]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 자주검사 CHECK SHEET																				*/
/*  Create Date    : 2010.11.09                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--EXEC [PS_PP035_03]
--Create PROC [dbo].[PS_PP035_03]

ALTER PROC [dbo].[PS_PP035_03]

AS
BEGIN

IF OBJECT_ID('Temp_LBG10') IS NULL
	BEGIN
		CREATE TABLE [Temp_LBG10]
		( 
		DocEntry  numeric
		)
	END	

SELECT CONVERT(NVARCHAR(30),(A.U_OrdNum + A.U_OrdSub1 + A.U_OrdSub2)) AS JAKBUN, 
       A.DocEntry, 
       B.FrgnName, 
       CONVERT(NVARCHAR(50),B.U_SIZE) AS SIZE,
       FLOOR(A.U_SelWt) AS SELWT, 
       converT(char(8),A.U_DueDate,112) AS DATE, C.CardName
  FROM [@PS_PP030H] AS A INNER JOIN OITM AS B
       ON A.U_ItemCode = B.ITEMCODE
       inner join ORDR AS C
       on A.U_SjNum = C.DocEntry
 WHERE A.DocEntry in (select DocEntry FROM Temp_LBG10)
end