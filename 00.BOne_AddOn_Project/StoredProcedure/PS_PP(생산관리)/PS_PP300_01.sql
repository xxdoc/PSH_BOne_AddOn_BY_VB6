USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP300_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 휘팅서울포장현황																	*/
/*  Create Date    : 2010.12.06                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP300_01]
--Create PROC [dbo].[PS_PP300_01]
(
    @YYYY Nvarchar(10)
)

AS

BEGIN

SELECT CONVERT(NVARCHAR(30),G.Spec1) AS Spec1,
       SUM(NPkWt) AS NPkWtTo,
	   ym01 = SUM(CHARINDEX(substring(G.DocDate,5,2), '01') * G.NPkWt),
	   ym02 = SUM(CHARINDEX(substring(G.DocDate,5,2), '02') * G.NPkWt),
	   ym03 = SUM(CHARINDEX(substring(G.DocDate,5,2), '03') * G.NPkWt),
	   ym04 = SUM(CHARINDEX(substring(G.DocDate,5,2), '04') * G.NPkWt),
	   ym05 = SUM(CHARINDEX(substring(G.DocDate,5,2), '05') * G.NPkWt),
	   ym06 = SUM(CHARINDEX(substring(G.DocDate,5,2), '06') * G.NPkWt),
	   ym07 = SUM(CHARINDEX(substring(G.DocDate,5,2), '07') * G.NPkWt),
	   ym08 = SUM(CHARINDEX(substring(G.DocDate,5,2), '08') * G.NPkWt),
	   ym09 = SUM(CHARINDEX(substring(G.DocDate,5,2), '09') * G.NPkWt),
	   ym10 = SUM(CHARINDEX(substring(G.DocDate,5,2), '10') * G.NPkWt),
	   ym11 = SUM(CHARINDEX(substring(G.DocDate,5,2), '11') * G.NPkWt),
	   ym12 = SUM(CHARINDEX(substring(G.DocDate,5,2), '12') * G.NPkWt) 
  FROM       

  ( SELECT C.U_Spec1 AS Spec1, 
         CONVERT(CHAR(8),B.U_DocDate,112) AS DocDate, 
         SUM(A.U_NPkWt) AS NPkWt
  FROM [@PS_PP077H] AS A INNER JOIN [@PS_PP040H] AS B
			ON A.U_PP040No = B.DocEntry
		INNER JOIN OITM AS C
			ON A.U_ItemCode = C.ItemCode
        
  GROUP BY C.U_Spec1, B.U_DocDate ) G
 
 WHERE SUBSTRING(DocDate,1,4) =  @YYYY
 
 GROUP BY G.Spec1
 
 END
 
--EXEC [dbo].[PS_PP300_01] '2010'