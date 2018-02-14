USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP635_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 출고현황(세부내역)       																	*/
/*  Create Date    : 2010.11.11                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/

ALTER PROC [dbo].PS_PP635_01
--Create PROC [dbo].[PS_PP635_01]
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
	DocDate			Nvarchar(8),
	CardCode		Nvarchar(20),
	CardName		Nvarchar(60),
	OrdNum			Nvarchar(20),
	BatchNum		Nvarchar(20),
	BaseType		Nvarchar(3),
	Quantity		Numeric(19,1),
	GUBUN			Nvarchar(10) )

Declare	@ItemCode1		Nvarchar(30),
		@ItemName		Nvarchar(60),
		@DocDate		Nvarchar(8),
		@CardCode		Nvarchar(20),
		@CardName		Nvarchar(60),
		@OrdNum			Nvarchar(20),
		@BatchNum		Nvarchar(20),
		@BaseType		Nvarchar(3),
		@Quantity		Numeric(19,1),
		@GUBUN			Nvarchar(10)

Declare @TempTalQuantity  Numeric(19,1), 
		@TempBTalQuantity Numeric(19,1),
		@BaseQuantity	  Numeric(19,1)
		
		
Declare @TempBaseType Nvarchar(3)
        		

DECLARE CUR_1 CURSOR FOR

SELECT CONVERT(NVARCHAR(30),Z.ItemCode)	AS ItemCode, 
       CONVERT(NVARCHAR(60),Z.ItemName)	AS ItemName,
       A.DocDate  AS DocDate,
       CONVERT(NVARCHAR(20),A.CardCode) AS CardCode,
       CONVERT(NVARCHAR(60),A.CardName) AS CardName,
       CONVERT(NVARCHAR(20),C.U_OrdNum) AS OrdNum,								 --Lot-No
       CONVERT(NVARCHAR(20),A.BatchNum) AS BatchNum,
       A.BaseType		AS BaseType,
	   A.Quantity AS Quantity,      
       CASE WHEN C.U_MulGbn1 = '10' THEN 'Y'
		    WHEN C.U_MulGbn1 = '20' THEN 'N' 
	   END AS GUBUN 	
  FROM	OIBT AS Z
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

 
GROUP BY Z.ItemCode, Z.ItemName, C.U_MulGbn1, A.DocDate, A.CardCode, A.CardName, C.U_OrdNum, A.BatchNum, 
		 A.BaseType, A.Quantity

Order by Z.ItemCode, A.DocDate, A.CardCode DESC

OPEN CUR_1

FETCH NEXT FROM CUR_1 INTO @ItemCode1, @ItemName, @DocDate,	@CardCode, @CardName, @OrdNum,
						   @BatchNum, @BaseType, @Quantity, @GUBUN

WHILE	@@FETCH_STATUS = 0

BEGIN
If @TempBaseType <> @BaseType Begin
	If @BaseType = '13' or @BaseType = '15' or @BaseType = '14' or @BaseType = '16'
	 BEGIN
	  set @TempTalQuantity = @Quantity
		Insert	#Temp01 Values (@ItemCode, @ItemName, @DocDate,	@CardCode, @CardName, @OrdNum,
						        @BatchNum, @BaseType, @Quantity, @GUBUN)	 
	 END
End

If @TempBaseType = @BaseType Begin
	If @BaseType = '13' or @BaseType = '15'
	 BEGIN
	  set @TempTalQuantity = @TempTalQuantity + @Quantity
	 END
	ELSE IF @BaseType = '14' or @BaseType = '16'
	 BEGIN
	  set @TempTalQuantity = @TempTalQuantity - @Quantity
	 END    	 
		Update	#Temp01 Set @TempTalQuantity = @BaseQuantity + @Quantity
		 Where BaseType = @BaseType

End

	Set @TempBaseType = @BaseType
	Set @BaseQuantity = @TempTalQuantity

     
 FETCH NEXT FROM CUR_1 INTO @ItemCode1, @ItemName, @DocDate, @CardCode, @CardName, @OrdNum,
						    @BatchNum, @BaseType, @Quantity, @GUBUN
End

CLOSE	CUR_1

DEALLOCATE CUR_1

SELECT  ItemCode,
		ItemName,
		DocDate,
		CardCode,
		CardName,
		OrdNum,
		BatchNum,
		BaseType
		Quantity,
		GUBUN    From #Temp01
End


-- EXEC [PS_PP635_01] '1', '20101101', '20101130', '101010002'
