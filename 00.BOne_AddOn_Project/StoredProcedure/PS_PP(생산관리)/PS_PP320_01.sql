USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP320_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/******************************************************************************************************************/
/*  Module         : PP								    														*/
/*  Description    : 부품 주문번호별 수주잔량 현황    															*/
/*  Create Date    : 2010.11.23                                                                                   */
/*  Modified Date  :										       													*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/******************************************************************************************************************/
ALTER PROC [dbo].[PS_PP320_01]
--Create PROC [dbo].[PS_PP320_01]
(
	@DocDate  Date,
	@ItmBsort Nvarchar(10),
	@ItmMsort Nvarchar(10)
)
AS

BEGIN

SELECT  A.U_LotNo		AS	LotNo,                                          --주문번호
        B.ItemCode		AS	ItemCode,  
        B.Dscription	AS	Dsription,                                      --제품명                                                             
		CONVERT(CHAR(10),A.DocDate,120)	AS	DocDate,										--수주일자
		B.Quantity,															--수주수량
		D.U_YQty		AS YWeight,											--생산수량
		ISNULL((select SUM(f.Quantity) from ODLN E,
					      DLN1 F
					where E.DocEntry = F.DocEntry
					  and E.U_LotNo = A.U_LotNo
					  and F.ItemCode = B.ItemCode ),0) As SUMQTY,
		(B.Quantity - ISNULL((select SUM(F.Quantity) from ODLN E,
					      DLN1 F
					where E.DocEntry = F.DocEntry
					  and E.U_LotNo = A.U_LotNo
					  and F.ItemCode = B.ItemCode ),0)) AS JANYANG			  		  
  FROM ORDR AS A INNER JOIN RDR1 AS B
		ON A.DocEntry = A.DocEntry
       INNER JOIN OITM  AS C
		ON B.ItemCode = C.ItemCode
       INNER JOIN [@PS_PP080L] AS D
		ON D.U_BatchNum = A.U_LotNo	     
 WHERE C.U_ItmBsort = '102'
   AND C.U_ItmBsort = @ItmBsort
   AND C.U_ItmMsort = @ItmMsort
   AND A.DocDate > = @DocDate
   
End   

--EXEC [dbo].[PS_PP320_01] '20101112','102','10201'