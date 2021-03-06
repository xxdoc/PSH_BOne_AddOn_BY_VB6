USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP091_02]    Script Date: 03/08/2011 09:46:57 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 패킹리스트조회 및 출력 [PS_PP091]                                                          */
/*  Create Date    : 2010.11.03                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Ryu Yung Jo																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP091_02]
--Create  PROC [dbo].[PS_PP091_02]
AS

SET NOCOUNT ON
--BEGIN ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
--Create Table [Z_PS_PP091]
--( PackNo Nvarchar(20) )
--Creat Table
CREATE	TABLE #Temp01 (
	PackNo01			Nvarchar(20),-- Not NULL,
	BPLId01				Nvarchar(2),
	BPLName01			Nvarchar(100),
	PackDate01			DateTime,
	FinalArrival01		Nvarchar(100),
	ItemName01			Nvarchar(100),
	LotNo01_01			Nvarchar(20),
	Qty01_01			Numeric(19,6) default 0,
	Weight01_01			Numeric(19,6) default 0,
	Status01_01			Nvarchar(50),
	LotNo01_02			Nvarchar(20),
	Qty01_02			Numeric(19,6) default 0,
	Weight01_02			Numeric(19,6) default 0,
	Status01_02			Nvarchar(50),
	LotNo01_03			Nvarchar(20),
	Qty01_03			Numeric(19,6) default 0,
	Weight01_03			Numeric(19,6) default 0,
	Status01_03			Nvarchar(50),
	LotNo01_04			Nvarchar(20),
	Qty01_04			Numeric(19,6) default 0,
	Weight01_04			Numeric(19,6) default 0,
	Status01_04			Nvarchar(50),
	LotNo01_05			Nvarchar(20),
	Qty01_05			Numeric(19,6) default 0,
	Weight01_05			Numeric(19,6) default 0,
	Status01_05			Nvarchar(50),
	SumQty01			Int,
	NetWt01				Numeric(19,6),
	TareWt01			Numeric(19,6),
	GrossWt01			Numeric(19,6),
	UserName01_01		Nvarchar(20),
	UserName01_02		Nvarchar(20),
	UserName01_03		Nvarchar(20),
	UserName01_04		Nvarchar(20),
	UserName01_05		Nvarchar(20),
		
	PackNo02			Nvarchar(20),-- Not NULL,
	BPLId02				Nvarchar(2),
	BPLName02			Nvarchar(100),
	PackDate02			DateTime,
	FinalArrival02		Nvarchar(100),
	ItemName02			Nvarchar(100),
	LotNo02_01			Nvarchar(20),
	Qty02_01			Numeric(19,6) default 0,
	Weight02_01			Numeric(19,6) default 0,
	Status02_01			Nvarchar(50),
	LotNo02_02			Nvarchar(20),
	Qty02_02			Numeric(19,6) default 0,
	Weight02_02			Numeric(19,6) default 0,
	Status02_02			Nvarchar(50),
	LotNo02_03			Nvarchar(20),
	Qty02_03			Numeric(19,6) default 0,
	Weight02_03			Numeric(19,6) default 0,
	Status02_03			Nvarchar(50),
	LotNo02_04			Nvarchar(20),
	Qty02_04			Numeric(19,6) default 0,
	Weight02_04			Numeric(19,6) default 0,
	Status02_04			Nvarchar(50),
	LotNo02_05			Nvarchar(20),
	Qty02_05			Numeric(19,6) default 0,
	Weight02_05			Numeric(19,6) default 0,
	Status02_05			Nvarchar(50),
	SumQty02			Int,
	NetWt02				Numeric(19,6),
	TareWt02			Numeric(19,6),
	GrossWt02			Numeric(19,6),
	UserName02_01		Nvarchar(20),
	UserName02_02		Nvarchar(20),
	UserName02_03		Nvarchar(20),
	UserName02_04		Nvarchar(20),
	UserName02_05		Nvarchar(20),
	
	Page				Int Not NULL )

Declare	@Cnt Int, @LineCnt Int, @PageCnt Int, @BefPackNo Nvarchar(20)
Declare	@PackNo Nvarchar(20), @BPLName Nvarchar(100), @PackDate DateTime, @FinalArrival Nvarchar(100), @ItemName Nvarchar(100), @CardName Nvarchar(100),
		@LotNo Nvarchar(20), @Weight Numeric(19,6), @Status Nvarchar(50), @SumQty Int, @NetWt Numeric(19,6), @TareWt Numeric(19,6), @Qty Numeric(19,6), @BPLId Nvarchar(2),
		@GrossWt Numeric(19,6), @UserName01 Nvarchar(20), @UserName02 Nvarchar(20), @UserName03 Nvarchar(20), @UserName04 Nvarchar(20), @UserName05 Nvarchar(20)
		
Set @Cnt = 1
Set @LineCnt = 0
Set @PageCnt = 0
Set @BefPackNo = ''  
DECLARE CUR_1 CURSOR FOR
	Select	Convert(Nvarchar(20), a.U_PackNo) As PackNo
			,d.BPLId
			,d.BPLName
			,a.U_InDate  As PackDate
			,Convert(Nvarchar(100), b.U_ItemName) As ItemName
			--,Left(e.U_LotNo, 7) + '-' + (CASE When d.BPLId = '1' Then 'S' When d.BPLId = '2' Then 'V' End) + z.U_CallSize
			-- + RIGHT(b.U_LotNo, 7) As LotNo
			,Left(e.U_LotNo, 7) + '-' + (CASE When d.BPLId = '1' Then 'S' When d.BPLId = '2' Then 'V' End) + z.U_CallSize
			 + Case When d.BPLId = '1' and left(b.U_LotNo,6) < '201103' Then (select Right(DS.LOTNO,7) from Z_DSMDFRY DS WHERE DS.custlotno = g.U_BatchNum)
			        When d.BPLId = '1' and left(b.U_LotNo,6) >= '201103' Then RIGHT(b.U_LotNo, 7)
			   Else RIGHT(b.U_LotNo, 7) End	 As LotNo
			,b.U_Qty As Qty
			,b.U_Weight As Weight
			,b.U_Qty As SumQty
			,b.u_Weight As NetWt
			--,Sum(b.U_Qty) As SumQty
			--,Sum(b.u_Weight) As NetWt
			,CONVERT(Nvarchar(100), e.U_CardName) As CardName
			,Case When IsNull(f.U_MulGbn1, '') = '10' Then 'D/G' When IsNull(f.U_MulGbn1, '') = '20' Then 'UN D/G' Else '' End As Status
	  From	[@PS_PP090H] a 
			Inner Join [@PS_PP090L] b On a.DocEntry = b.DocEntry
			Inner Join [Z_PS_PP091] c On a.U_PackNo = c.PackNo
			Inner Join [OBPL] d On d.BPLId = a.U_BPLId
			Inner Join [@PS_QM020H] e On e.U_OrdNum = b.U_LotNo
			Inner Join [@PS_PP030H] f On f.U_OrdNum = b.U_LotNo
			Inner Join [@PS_PP030L] g On f.DocEntry = g.DocEntry
			Inner Join [OITM] z On z.ItemCode = b.U_ItemCode
	--Group by a.U_PackNo, d.BPLName, a.U_InDate, b.U_ItemName,
	--		 --Left(e.U_LotNo, 7) + '-' + (CASE When d.BPLId = '1' Then 'S' When d.BPLId = '2' Then 'V' End) + z.U_CallSize
	--		 --+ Case When d.BPLId = '1' and left(b.U_LotNo,6) < '201103' Then (select Right(DS.LOTNO,7) from Z_DSDMFRY DS WHERE DS.custlotno = g.U_BatchNum)
	--		 --       When d.BPLId = '1' and left(b.U_LotNo,6) >= '201103' Then RIGHT(b.U_LotNo, 7)
	--		 --  Else RIGHT(b.U_LotNo, 7) End,
	--		 Left(e.U_LotNo, 7) + '-' + (CASE When d.BPLId = '1' Then 'S' When d.BPLId = '2' Then 'V' End) + z.U_CallSize
	--		 + RIGHT(b.U_LotNo, 7),
	--		 b.U_Weight, e.U_CardName, U_Qty, U_MulGbn1, d.BPLId, b.U_LineNum
	Order by PackNo, b.U_LineNum
OPEN CUR_1
FETCH NEXT FROM CUR_1 INTO 	@PackNo, @BPLId, @BPLName, @PackDate, @ItemName, @LotNo, @Qty, @Weight, @SumQty, @NetWt, @CardName, @Status
WHILE	@@FETCH_STATUS = 0
BEGIN			
	IF @BefPackNo <> @PackNo BEGIN
    	Set @LineCnt = 1
		--Set @PageCnt = 1 
	END		
	
	IF @BefPackNo = @PackNo and @Cnt > 1 Begin 
		set @Cnt = @Cnt - 1 
	End
	
	IF @LineCnt % 5 = 1 AND @LineCnt <> 1 BEGIN    	
		Set @LineCnt = 1
	END
	--print @Cnt
	--print @LineCnt
	--print @PageCnt
	
	IF @Cnt % 2 = 1 BEGIN
		IF @LineCnt = 1 BEGIN
			Set @PageCnt = @PageCnt + 1
			Insert Into #Temp01 (PackNo01, BPLId01, BPLName01, PackDate01, ItemName01, LotNo01_01, Qty01_01, Weight01_01, SumQty01, NetWt01, FinalArrival01, Status01_01, Page)
			Values				(@PackNo, @BPLId, @BPLName, @packDate, @ItemName, @LotNo, @Qty, @Weight, @SumQty, @NetWt, @CardName, @Status, @PageCnt)
		END
		IF @LineCnt = 2 BEGIN
			Update #Temp01 Set LotNo01_02 = @LotNo, Weight01_02 = @Weight, Status01_02 = @Status, Qty01_02 = @Qty
			 Where PackNo01 = @PackNo And Page = @PageCnt
		END
		IF @LineCnt = 3 BEGIN
			Update #Temp01 Set LotNo01_03 = @LotNo, Weight01_03 = @Weight, Status01_03 = @Status, Qty01_03 = @Qty
			 Where PackNo01 = @PackNo And Page = @PageCnt
		END
		IF @LineCnt = 4 BEGIN
			Update #Temp01 Set LotNo01_04 = @LotNo, Weight01_04 = @Weight, Status01_04 = @Status, Qty01_04 = @Qty
			 Where PackNo01 = @PackNo And Page = @PageCnt
		END
		IF @LineCnt = 5 BEGIN
			Update #Temp01 Set LotNo01_05 = @LotNo, Weight01_05 = @Weight, Status01_05 = @Status, Qty01_05 = @Qty
			 Where PackNo01 = @PackNo And Page = @PageCnt
		END		
	END
	
	IF @Cnt % 2 = 0 BEGIN
		IF @LineCnt = 1 BEGIN
			Update #Temp01 Set PackNo02 = @PackNo, BPLId02 = @BPLId, BPLName02 = @BPLName, PackDate02 = @PackDate, ItemName02 = @ItemName,							   						   
							   LotNo02_01 = @LotNo, Weight02_01 = @Weight, Status02_01 = @Status,
							   SumQty02 = @SumQty, NetWt02 = @NetWt, Qty02_01 = @Qty, FinalArrival02 = @CardName
			 Where Page = @PageCnt
		END
		IF @LineCnt = 2 BEGIN
			Update #Temp01 Set LotNo02_02 = @LotNo, Weight02_02 = @Weight, Status02_02 = @Status, Qty02_02 = @Qty
			 Where PackNo02 = @PackNo And Page = @PageCnt
		END
		IF @LineCnt = 3 BEGIN
			Update #Temp01 Set LotNo02_03 = @LotNo, Weight02_03 = @Weight, Status02_03 = @Status, Qty02_03 = @Qty
			 Where PackNo02 = @PackNo And Page = @PageCnt
		END
		IF @LineCnt = 4 BEGIN
			Update #Temp01 Set LotNo02_04 = @LotNo, Weight02_04 = @Weight, Status02_04 = @Status, Qty02_04 = @Qty
			 Where PackNo02 = @PackNo And Page = @PageCnt
		END
		IF @LineCnt = 5 BEGIN
			Update #Temp01 Set LotNo02_05 = @LotNo, Weight02_05 = @Weight, Status02_05 = @Status, Qty02_05 = @Qty
			 Where PackNo02 = @PackNo And Page = @PageCnt
		END		
	END
	
	Set @Cnt = @Cnt + 1
	Set @LineCnt = @LineCnt + 1
	Set @BefPackNo = @PackNo
FETCH NEXT FROM CUR_1 INTO 	@PackNo, @BPLId, @BPLName, @PackDate, @ItemName, @LotNo, @Qty, @Weight, @SumQty, @NetWt, @CardName, @Status
END	

CLOSE	CUR_1
DEALLOCATE CUR_1

Select * From #Temp01
--Order by PackNo01
--THE END //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-- EXEC [PS_PP091_02] 
-- select * from [Z_PS_PP091]

-- update [@ps_mm005h] set  U_Status = 'O', U_CntcCode = '1', U_DeptCode = '3'








