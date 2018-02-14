USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP030_04]    Script Date: 11/09/2010 16:08:16 ******/
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
ALTER PROC [dbo].[PS_PP030_04]
--(
--	@DocEntry As NVARCHAR(10)
--)
AS
BEGIN

CREATE TABLE [Temp_LBG10]
( 
DocEntry  numeric
)

SELECT (A.U_OrdNum + A.U_OrdSub1 + A.U_OrdSub2) AS JAKBUN, A.DocEntry, B.FrgnName, B.U_SIZE, FLOOR(A.U_SelWt) AS SELWT, 
       converT(char(8),A.U_DueDate,112) as DATE, C.CardName
  FROM [@PS_PP030H] AS A INNER JOIN OITM AS B
       ON A.U_ItemCode = B.ITEMCODE
       inner join ORDR as C
       on A.U_SjNum = C.DocEntry
 WHERE A.DocEntry in (select DocEntry FROM Temp_LBG10)
end