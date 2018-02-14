USE [PSHDB_Zchoi]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP035_05]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 작업지시서																		*/
/*  Create Date    : 2010.11.11                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--  EXEC [PS_PP035_05] 'S'
--  EXEC [PS_PP035_05] 'M'
ALTER PROC [dbo].[PS_PP035_05]
--Create PROC [dbo].[PS_PP035_05]
(
	@Seq	As Nvarchar(1)
)
AS
BEGIN
IF OBJECT_ID('Temp_LBG12') IS NULL BEGIN
	CREATE TABLE [Temp_LBG12]
	( 
	DocEntry  int
	)
END
If @Seq = 'M' BEGIN		
Create Table #Temp01
(	DocEntry		Nvarchar(10),
	JAKBUN			Nvarchar(20),
	FrgnName		Nvarchar(30),
	MarkName		Nvarchar(20),
	SIZE			Nvarchar(50),
	CallSize		Nvarchar(20),
	TypeName		Nvarchar(20),
	UnWeight		Numeric(19,3),
	CardName		Nvarchar(100),
	Selwt			Numeric(19,3),
	Jungrang		Numeric(19,3),
	HComments		Nvarchar(100),
	ItemName		Nvarchar(1000),
	Comments		Nvarchar(100) )

Declare	@DocEntry		Nvarchar(10),
		@JAKBUN			Nvarchar(20),
		@FrgnName		Nvarchar(30),
		@MarkName		Nvarchar(20),
		@SIZE			Nvarchar(50),
		@CallSize		Nvarchar(20),
		@TypeName		Nvarchar(20),
		@UnWeight		Numeric(19,3),
		@CardName		Nvarchar(100),
		@Selwt			Numeric(19,3),
		@Jungrang		Numeric(19,3),
		@HComments		Nvarchar(100),
		@ItemName		Nvarchar(1000),
		@Comments		Nvarchar(100)
		
Declare @BaseDocEntry	Nvarchar(10), @BaseItemName	Nvarchar(100), @TempItemName Nvarchar(100)

Set @BaseDocEntry = 0
Set @BaseItemName = ''
Set @TempItemName = ''
DECLARE CUR_1 CURSOR FOR
	SELECT	CONVERT(Nvarchar(10), A.DocEntry) aS DocEntry,
			CONVERT(NVARCHAR(20),(A.U_OrdNum + A.U_OrdSub1 + A.U_OrdSub2)) AS JAKBUN,
			FrgnName = ( SELECT MAX(FrgnName) FROM OITM WHERE Itemcode = A.U_Itemcode ), 
		    MarkName =  ( SELECT MAX(Name) FROM [@PSH_MARK] AS A INNER JOIN OITM AS B 
								ON A.Code = B.U_Mark),
			SIZE = ( SELECT MAX(U_Size) FROM OITM WHERE Itemcode = A.U_Itemcode ),
			CallSize = ( SELECT MAX(U_CallSize) FROM OITM WHERE Itemcode = A.U_Itemcode ),
			TypeName = ( SELECT MAX(Name) FROM[@PSH_SHAPE] AS A INNER JOIN OITM AS B
								ON A.Code = B. U_ItemType),
			UnWeight = ( SELECT MAX(U_UnWeight) FROM OITM WHERE Itemcode = A.U_Itemcode ),
			CardName = ( SELECT MAX(U_CardName) FROM [@PS_SD010H]AS A INNER JOIN [@PS_PP030H] AS B
								ON A.DocEntry = B.U_BaseNum),
			A.U_SelWt AS Selwt,             
			 Jungrang = ( SELECT ( (A.U_SelWt) *  (H.U_UnWeight) / 1000 ) FROM OITM AS H
								WHERE A.U_Itemcode = H.ItemCode ),		                            --지시수량   
			CONVERT(nvarchar(100),A.U_Comments) AS PP030HComments,
			ItemName = (SELECT MAX(U_ItemName) FROM [@PS_PP030L] WHERE DocEntry = A.DocEntry),			    	--사용원재료
			Comments = (select Comments from ORDR WHERE DocEntry = A.U_SjNum)
	  FROM	[@PS_PP030H] AS A 
	 WHERE  A.DocEntry in (select DocEntry FROM Temp_LBG12)
OPEN CUR_1
FETCH NEXT FROM CUR_1 INTO 	@DocEntry, @JAKBUN, @FrgnName, @MarkName, @SIZE, @CallSize, @TypeName, @UnWeight,
							@CardName, @Selwt, @Jungrang, @HComments, @ItemName, @Comments
WHILE	@@FETCH_STATUS = 0
BEGIN	
	If @BaseDocEntry <> @DocEntry Begin
		Set @TempItemName = @ItemName
		Insert	#Temp01 Values (@DocEntry, @JAKBUN, @FrgnName, @MarkName, @SIZE, @CallSize, @TypeName, @UnWeight,
								@CardName, @Selwt, @Jungrang, @HComments, @ItemName, @Comments)
	End
	If @BaseDocEntry = @DocEntry Begin
		Set @TempItemName = @BaseItemName + ' ,  ' + @ItemName
		Update	#Temp01	Set		ItemName = @TempItemName
		 Where	DocEntry = @DocEntry
	End						
	
	Set @BaseDocEntry = @Docentry
	Set @BaseItemName = @TempItemName
	------Print @BaseDocEntry
	------Print @ItemName
	------Print @BaseItemName
	------Print @TempItemName
  ------ EXEC [PS_PP035_05] 'M'
FETCH NEXT FROM CUR_1 INTO 	@DocEntry, @JAKBUN, @FrgnName, @MarkName, @SIZE, @CallSize, @TypeName, @UnWeight,
							@CardName, @Selwt, @Jungrang, @HComments, @ItemName, @Comments
END
CLOSE	CUR_1
DEALLOCATE CUR_1

Select DocEntry,
	JAKBUN,
	FrgnName,
	MarkName,
	SIZE,
	CallSize,
	TypeName,
	UnWeight,
	CardName,
	Selwt,
	Jungrang,
	HComments,
	ItemName,
	Comments From #Temp01
END
If @Seq = 'S' BEGIN
SELECT CONVERT(Nvarchar(10), A.DocEntry) aS DocEntry,
       CONVERT(NVARCHAR(20),B.U_CpName) AS CpName,
       B.LineId,
       CONVERT(NVARCHAR(8),B.U_WorkGbn) AS WorkGbn,
     CASE WHEN U_WorkGbn = '30'
          THEN  '외주'
     ELSE  ''
     END AS WorkGbn     
  FROM [@PS_PP030H] AS A INNER JOIN [@PS_PP030M] AS B
        ON A.DocEntry = B.DocEntry 
 WHERE A.DocEntry in (select DocEntry FROM Temp_LBG12)
END
END