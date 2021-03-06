USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP320_01]    Script Date: 03/29/2011 19:54:47 ******/
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
	@ItmMsort Nvarchar(10),
	@Seq	As Nvarchar(1)
)
AS

BEGIN

If @Seq = 'S'

Select CONVERT(NVARCHAR(20),t1.Lotno) As Lotno,
	   t1.ItemCode,
	   CONVERT(NVARCHAR(60),t1.Dscription) As Dsription,
	   t1.DocDate,
	   t1.Quantity,
	   YQty = (select sum(Case When BaseType = 59 Then Quantity Else Quantity * -1 End) from IBT1 Where BaseType in (59,60) And ItemCode = t1.ItemCode And BatchNum = t1.Lotno), --생산수량
	   SUMQTY = (select sum(Case When BaseType = 15 Then Quantity Else Quantity * -1 End) from IBT1 Where BaseType in (15,16) And ItemCode = t1.ItemCode And BatchNum = t1.Lotno),
	   JANYANG = t1.Quantity - (select sum(Case When BaseType = 15 Then Quantity Else Quantity * -1 End) from IBT1 Where BaseType in (15,16) And ItemCode = t1.ItemCode And BatchNum = t1.Lotno),
        CASE WHEN TrType = '1' THEN ''
             WHEN TrType = '2' THEN '군납'
        END AS TYPE			  			  		  
From (
Select A.U_Lotno As Lotno,
	   B.ItemCode,
	   B.Dscription,
	   Max(A.DocDate) As DocDate,
	   sum(B.Quantity) As Quantity,
	   Max(B.U_TrType) As TrType
 FROM ORDR AS A INNER JOIN RDR1 AS B
		ON A.DocEntry = B.DocEntry
       INNER JOIN OITM  AS C
		ON B.ItemCode = C.ItemCode     
 WHERE C.U_ItmBsort = '102'
   And A.Canceled = 'N'
   AND C.U_ItmMsort Like @ItmMsort + '%'
   Group by A.U_Lotno,
	   B.ItemCode,
	   B.Dscription
) t1				  
Where t1.Quantity - (select sum(Case When BaseType = 15 Then Quantity Else Quantity * -1 End) from IBT1 Where BaseType in (15,16) And ItemCode = t1.ItemCode And BatchNum = t1.Lotno) > 0
Order by t1.ItemCode, t1.DocDate, t1.Lotno
--SELECT  CONVERT(NVARCHAR(20),A.U_LotNo)		AS	LotNo,                                                  --주문번호
--        CONVERT(NVARCHAR(60),B.ItemCode)	AS	ItemCode,  
--        CONVERT(NVARCHAR(60),B.Dscription)	AS	Dsription,                                              --제품명                                                             
--		A.DocDate		AS	DocDate,						        --수주일자
--		CONVERT(NUMERIC(19,2),B.Quantity) AS Quantity,																    --수주수량
--		YQty = (select sum(Case When BaseType = 59 Then Quantity Else Quantity * -1 End) from IBT1 Where BaseType in (59,60) And ItemCode = B.ItemCode And BatchNum = A.U_Lotno), --생산수량
--		--SELECT SUM(CONVERT(NUMERIC(19,2),U_YQty)) FROM [@PS_PP080L] WHERE U_BatchNum = A.U_LotNo),	--생산수량
--		ISNULL((select SUM(CONVERT(NUMERIC(19,2),f.Quantity)) from ODLN E,
--					      DLN1 F
--					where E.DocEntry = F.DocEntry
--					  and E.U_LotNo = A.U_LotNo
--					  and F.ItemCode = B.ItemCode ),0) As SUMQTY,
--		(CONVERT(NUMERIC(19,2),B.Quantity) - ISNULL((select SUM(CONVERT(NUMERIC(19,2),F.Quantity)) from ODLN E,
--					      DLN1 F
--					where E.DocEntry = F.DocEntry
--					  and E.U_LotNo = A.U_LotNo
--					  and F.ItemCode = B.ItemCode ),0)) AS JANYANG,
--        CASE WHEN U_TrType = '1' THEN ''
--             WHEN U_TrType = '2' THEN '군납'
--        END AS TYPE			  			  		  
--  FROM ORDR AS A INNER JOIN RDR1 AS B
--		ON A.DocEntry = B.DocEntry
--       INNER JOIN OITM  AS C
--		ON B.ItemCode = C.ItemCode     
-- WHERE C.U_ItmBsort = '102'
--   And A.Canceled = 'N'
--   AND C.U_ItmMsort Like @ItmMsort + '%'
--   AND (B.Quantity - ISNULL((select SUM(F.Quantity) from ODLN E,
--					      DLN1 F
--					where E.DocEntry = F.DocEntry
--					  and E.U_LotNo = A.U_LotNo
--					  and F.ItemCode = B.ItemCode ),0)) > 0 
					  
If @Seq = 'T'

Select CONVERT(NVARCHAR(20),t1.Lotno) As Lotno,
	   t1.ItemCode,
	   CONVERT(NVARCHAR(60),t1.Dscription) As Dsription,
	   t1.DocDate,
	   t1.Quantity,
	   YQty = (select sum(Case When BaseType = 59 Then Quantity Else Quantity * -1 End) from IBT1 Where BaseType in (59,60) And ItemCode = t1.ItemCode And BatchNum = t1.Lotno), --생산수량
	   SUMQTY = (select sum(Case When BaseType = 15 Then Quantity Else Quantity * -1 End) from IBT1 Where BaseType in (15,16) And ItemCode = t1.ItemCode And BatchNum = t1.Lotno),
	   JANYANG = t1.Quantity - (select sum(Case When BaseType = 15 Then Quantity Else Quantity * -1 End) from IBT1 Where BaseType in (15,16) And ItemCode = t1.ItemCode And BatchNum = t1.Lotno),
        CASE WHEN TrType = '1' THEN ''
             WHEN TrType = '2' THEN '군납'
        END AS TYPE			  			  		  
From (
Select A.U_Lotno As Lotno,
	   B.ItemCode,
	   B.Dscription,
	   Max(A.DocDate) As DocDate,
	   sum(B.Quantity) As Quantity,
	   Max(B.U_TrType) As TrType
 FROM ORDR AS A INNER JOIN RDR1 AS B
		ON A.DocEntry = B.DocEntry
       INNER JOIN OITM  AS C
		ON B.ItemCode = C.ItemCode     
 WHERE C.U_ItmBsort = '102'
   And A.Canceled = 'N'
   AND C.U_ItmMsort Like @ItmMsort + '%'
   Group by A.U_Lotno,
	   B.ItemCode,
	   B.Dscription
) t1				  
Order by t1.ItemCode, t1.DocDate, t1.Lotno

--SELECT  CONVERT(NVARCHAR(20),A.U_LotNo)		AS	LotNo,                                                  --주문번호
--        CONVERT(NVARCHAR(60),B.ItemCode)	AS	ItemCode,  
--        CONVERT(NVARCHAR(60),B.Dscription)	AS	Dsription,                                              --제품명                                                             
--		A.DocDate		AS	DocDate,						        --수주일자
--		CONVERT(NUMERIC(19,2),B.Quantity) AS Quantity,																    --수주수량
--		YQty = (SELECT SUM(CONVERT(NUMERIC(19,2),U_YQty)) FROM [@PS_PP080L] WHERE U_BatchNum = A.U_LotNo),	--생산수량
--		ISNULL((select SUM(CONVERT(NUMERIC(19,2),f.Quantity)) from ODLN E,
--					      DLN1 F
--					where E.DocEntry = F.DocEntry
--					  and E.U_LotNo = A.U_LotNo
--					  and F.ItemCode = B.ItemCode ),0) As SUMQTY,
--		(CONVERT(NUMERIC(19,2),B.Quantity) - ISNULL((select SUM(CONVERT(NUMERIC(19,2),F.Quantity)) from ODLN E,
--					      DLN1 F
--					where E.DocEntry = F.DocEntry
--					  and E.U_LotNo = A.U_LotNo
--					  and F.ItemCode = B.ItemCode ),0)) AS JANYANG,
--        CASE WHEN U_TrType = '1' THEN ''
--             WHEN U_TrType = '2' THEN '군납'
--        END AS TYPE			  			  		  
--  FROM ORDR AS A INNER JOIN RDR1 AS B
--		ON A.DocEntry = B.DocEntry
--       INNER JOIN OITM  AS C
--		ON B.ItemCode = C.ItemCode     
-- WHERE C.U_ItmBsort = '102'
--   AND C.U_ItmMsort = @ItmMsort

End   


--EXEC [dbo].[PS_PP320_01] '10201', 'S'
--EXEC [dbo].[PS_PP320_01] '10201', 'T'
