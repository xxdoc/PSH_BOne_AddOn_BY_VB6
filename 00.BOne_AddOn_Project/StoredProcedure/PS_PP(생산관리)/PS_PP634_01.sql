USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP634_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 출고현황(품목별)       																	*/
/*  Create Date    : 2010.11.11                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/

ALTER PROC [dbo].PS_PP634_01
--Create PROC [dbo].[PS_PP634_01]
(
	@BPLId		AS   Nvarchar(1),
	@DocDateFr  AS   Nvarchar(8),
	@DocDateTo  AS   Nvarchar(8),
	@ItemCode	AS	 Nvarchar(30)
)
AS
BEGIN

Create Table #Temp01
(	ItemCode		Nvarchar(30),
	ItemName		Nvarchar(60),
	GUBUN			Nvarchar(10),
	BaseType		Nvarchar(3),	
	TalQuantity 	Numeric(19,1),
	BTalQuantity 	Numeric(19,1) )

Declare	@ItemCode1		Nvarchar(30),
		@ItemName		Nvarchar(60),
		@GUBUN			Nvarchar(10),
		@BaseType		Nvarchar(3),
		@TalQuantity	Numeric(19,1),
		@BTalQuantity	Numeric(19,1)

Declare @TempTalQuantity Numeric(19,1), 
		@TempBTalQuantity Numeric(19,1)

DECLARE CUR_1 CURSOR FOR

SELECT  Z.ItemCode AS ItemCode,
		Z.ItemName AS ItemName,
  	    CASE WHEN C.U_MulGbn1 = '10' THEN '탈지'
		     WHEN C.U_MulGbn1 = '20' THEN '비탈지' 
		END AS GUBUN,
		A.BaseType		AS BaseType,
		SUM( CASE WHEN C.U_MulGbn1 = '10' THEN  A.Quantity END ) AS TalQuantity, 
		SUM( CASE WHEN C.U_MulGbn1 = '20' THEN  A.Quantity END ) AS BTalQuantity      
  FROM	OIBT As Z
		INNER JOIN IBT1 AS A On Z.ItemCode = A.ItemCode And Z.BatchNum = A.BatchNum And Z.WhsCode = A.WhsCode
		INNER JOIN [@PS_PP030L] AS B
					ON A.ItemCode  = B.U_ItemCode AND
					   A.BatchNum  = B.U_BatchNum
				 INNER JOIN [@PS_PP030H] AS C
					ON B.DocEntry = C.DocEntry
 WHERE A.BaseType BETWEEN '13' AND '16'
   AND C.U_OrdGbn = '104'
   AND C.U_BPLId = @BPLId
   AND A.ItemCode Like @ItemCode
   AND A.DocDate between @DocDateFr AND @DocDateTo
 
GROUP BY Z.ItemCode, Z.ItemName, C.U_MulGbn1, A.BaseType
   
ORDER BY Z.ItemCode

OPEN CUR_1

FETCH NEXT FROM CUR_1 INTO @ItemCode1, @ItemName, @GUBUN, @TalQuantity, @BTalQuantity

WHILE	@@FETCH_STATUS = 0

BEGIN	
	If @GUBUN = '10' AND (@BaseType = '13' or @BaseType = '15')
	 BEGIN
	  set @TempTalQuantity = @TalQuantity + @TalQuantity
	 END
	ELSE IF @GUBUN = '10' AND (@BaseType = '14' or @BaseType = '16')
	 BEGIN
	  set @TempTalQuantity = @TempTalQuantity - @TalQuantity
	 END
	ELSE If @GUBUN = '20' AND (@BaseType = '13' or @BaseType = '15')
	 BEGIN
	  set @TempBTalQuantity = @TempBTalQuantity + @BTalQuantity
	 END
	ELSE IF @GUBUN = '10' AND (@BaseType = '14' or @BaseType = '16')
	 BEGIN
	  set @TempBTalQuantity = @TempBTalQuantity - @BTalQuantity
	 END		
		Insert	#Temp01 Values (@ItemCode, @ItemName, @GUBUN, @TalQuantity, @BTalQuantity)
     
 FETCH NEXT FROM CUR_1 INTO @ItemCode1, @ItemName, @GUBUN, @TalQuantity, @BTalQuantity
End

CLOSE	CUR_1

DEALLOCATE CUR_1

SELECT  ItemCode,
		ItemName,
	    GUBUN,
		TalQuantity,
		BTalQuantity  From #Temp01
End


-- EXEC [PS_PP634_01] '1', '20101101', '20101130', '101010002'