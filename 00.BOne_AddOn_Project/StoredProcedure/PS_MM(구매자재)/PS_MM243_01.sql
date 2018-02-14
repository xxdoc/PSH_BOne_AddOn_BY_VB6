USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM243_01]    Script Date: 12/06/2010 22:17:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : 임가공 원재료 입고현황																	*/
/*  Create Date    : 2010.12.07                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_MM243_01]
--Create PROC [dbo].[PS_MM243_01]
(
	@Location	NVARCHAR(1),
	@DocDateFr	Date,
	@DocDateTo	Date,
	@ItemCode	NVARCHAR(30),
	@ItmBsort	NVARCHAR(30),
	@BatchNum	NVARCHAR(30)
)
AS

BEGIN

SELECT CONVERT(NVARCHAR(10),D.U_ItmBsort) AS ItmBsor,
       CONVERT(NVARCHAR(30),B.ItemCode) AS ItemCode030L,
	   ItemName030L = (select ItemName FROM OITM WHERE ItemCode = B.ItemCode),
  	   CONVERT(NVARCHAR(50),A.U_BaseLot) AS BatchNum,
  	   B.DocDate AS DocDate,
  	   CONVERT(NUMERIC(19,2),B.Quantity) AS Quantity,
  	   0 AS BanNap,												--반납수량
  	   (B.Quantity - 0) AS ToTal								--총입고중량 - 반납수량
  FROM OBTN AS A INNER JOIN IGN1 AS B
			ON A.ItemCode = B.ItemCode   
	   INNER JOIN OWHS AS C
			ON B.WhsCode = C.WhsCode   
	   INNER JOIN OITM AS D
			ON B.ItemCode = D.ItemCode
 WHERE D.U_ItmBsort IN ('302','309')
   AND C.Location = @Location
   AND B.DocDate BETWEEN @DocDateFr AND @DocDateTo
   AND B.ItemCode Like @ItemCode
   AND D.U_ItmBsort = @ItmBsort
   AND ISNULL(A.U_BaseLot,'') = @BatchNum
   
 ORDER BY B.ItemCode, B.DocDate ASC
   
END   

--EXEC [dbo].[PS_MM243_01] '1','20101101','20101130','502010004','302',''