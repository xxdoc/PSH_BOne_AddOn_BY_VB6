USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP633_01]    Script Date: 12/09/2010 13:15:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 제품재고 LIST(세부내역)     																	*/
/*  Create Date    : 2010.12.09                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP633_01]
--Create PROC [dbo].[PS_PP633_01]
(
	@Location	NVARCHAR(1),
	@DocDate		Date,
	@ItemCode	NVARCHAR(30)
)
AS

BEGIN

SELECT	CONVERT(NVARCHAR(20),A.U_ItemCode)	AS ItemCode,
		CONVERT(NVARCHAR(60),A.U_ItemName)	AS ItemName,
		CONVERT(NVARCHAR(100),A.U_OrdNum)    AS OrdNum,
		CONVERT(NVARCHAR(30),A.U_LotNo)		AS LotNo,
  	    CASE WHEN A.U_MulGbn1 = '10' THEN '탈지'
		     WHEN A.U_MulGbn1 = '20' THEN '비탈지' 
		END AS MulGbn1 , 	
		B.Quantity		AS Quantity,
		SUM(CASE WHEN CONVERT(NVARCHAR(5),B.Direction) = '1' THEN -1
				 WHEN B.Direction = '0' THEN 1
			END)  AS Coil,
		B.DocDate AS DocDate 
  FROM [@PS_PP030H] AS A INNER JOIN IBT1 AS B
			ON A.U_ItemCode = B.ItemCode  AND
			   A.U_OrdNum   = B.BatchNum
 WHERE A.U_OrdGbn = '104'		   
   AND A.U_BPLId = @Location
   AND A.U_ItemCode LIKE @ItemCode
   AND B.DocDate <= @DocDate

GROUP BY A.U_ItemCode, A.U_ItemName, A.U_OrdNum, A.U_LotNo, A.U_MulGbn1, B.Quantity, B.DocDate

HAVING SUM(CASE WHEN CONVERT(NVARCHAR(5),B.Direction) = '1' THEN -1
				WHEN B.Direction = '0' THEN 1
		   END) > 0
			  
 End


-- EXEC [PS_PP633_01] '1', '20101121', '%'