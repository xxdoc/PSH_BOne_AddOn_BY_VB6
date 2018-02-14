SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 생산관리																				    */
/*  Description    : 휘팅이동등록 - 출고원부/반출증																*/
/*  ALTER  Date    : 2010.11.15																					*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_PP075_02]
ALTER     PROC [dbo].[PS_PP075_02]
(
  @_DocNum			as int
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
create table #PS_PP075
(
	Seq			int identity ,
	DocNum		int,
	MovDocNo	nvarchar(20),
	BPLId		nvarchar(5),
	BPLName		nvarchar(200),
	RegiDate	datetime,
	CardCode	nvarchar(20),
	CardName	nvarchar(100),
	CarNo		nvarchar(30),
	TransCom	nvarchar(150),
	DeliArea	nvarchar(150),
	Fee			numeric(19,6),
	Comments	nvarchar(254),
	
	No01		int,
	ItemCode01	nvarchar(20),
	ItemName01	nvarchar(200),
	Size01		nvarchar(200),
	TypeCode01	nvarchar(5),
	TypeName01	nvarchar(200),
	Mark01		nvarchar(100),
	Box01		int,
	Qty01		numeric(19,6),
	Weight01	numeric(19,6),

	No02		int,	
	ItemCode02	nvarchar(20),
	ItemName02	nvarchar(200),
	Size02		nvarchar(200),
	TypeCode02	nvarchar(5),
	TypeName02	nvarchar(200),
	Mark02		nvarchar(100),
	Box02		int,
	Qty02		numeric(19,6),
	Weight02	numeric(19,6),
	
	No03		int,	
	ItemCode03	nvarchar(20),
	ItemName03	nvarchar(200),
	Size03		nvarchar(200),
	TypeCode03	nvarchar(5),
	TypeName03	nvarchar(200),
	Mark03		nvarchar(100),
	Box03		int,	
	Qty03		numeric(19,6),
	Weight03	numeric(19,6),

	No04		int,	
	ItemCode04	nvarchar(20),
	ItemName04	nvarchar(200),
	Size04		nvarchar(200),
	TypeCode04	nvarchar(5),
	TypeName04	nvarchar(200),
	Mark04		nvarchar(100),
	Box04		int,
	Qty04		numeric(19,6),
	Weight04	numeric(19,6),

	No05		int,	
	ItemCode05	nvarchar(20),
	ItemName05	nvarchar(200),
	Size05		nvarchar(200),
	TypeCode05	nvarchar(5),
	TypeName05	nvarchar(200),
	Mark05		nvarchar(100),
	Box05		int,	
	Qty05		numeric(19,6),
	Weight05	numeric(19,6),

	No06		int,	
	ItemCode06	nvarchar(20),
	ItemName06	nvarchar(200),
	Size06		nvarchar(200),
	TypeCode06	nvarchar(5),
	TypeName06	nvarchar(200),
	Mark06		nvarchar(100),
	Box06		int,
	Qty06		numeric(19,6),
	Weight06	numeric(19,6),

	No07		int,	
	ItemCode07	nvarchar(20),
	ItemName07	nvarchar(200),
	Size07		nvarchar(200),
	TypeCode07	nvarchar(5),
	TypeName07	nvarchar(200),
	Mark07		nvarchar(100),
	Box07		int,	
	Qty07		numeric(19,6),
	Weight07	numeric(19,6),

	No08		int,	
	ItemCode08	nvarchar(20),
	ItemName08	nvarchar(200),
	Size08		nvarchar(200),
	TypeCode08	nvarchar(5),
	TypeName08	nvarchar(200),
	Mark08		nvarchar(100),
	Box08		int,
	Qty08		numeric(19,6),
	Weight08	numeric(19,6),

	No09		int,	
	ItemCode09	nvarchar(20),
	ItemName09	nvarchar(200),
	Size09		nvarchar(200),
	TypeCode09	nvarchar(5),
	TypeName09	nvarchar(200),
	Mark09		nvarchar(100),
	Box09		int,
	Qty09		numeric(19,6),
	Weight09	numeric(19,6),

	No10		int,	
	ItemCode10	nvarchar(20),
	ItemName10	nvarchar(200),
	Size10		nvarchar(200),
	TypeCode10	nvarchar(5),
	TypeName10	nvarchar(200),
	Mark10		nvarchar(100),
	Box10		int,	
	Qty10		numeric(19,6),
	Weight10	numeric(19,6)
)
-----------------------------------------------------------------------------------------------------------------------------------------
declare	@DocNum		int,
		@MovDocNo	nvarchar(20),
		@BPLId		nvarchar(5),
		@BPLName	nvarchar(200),
		@RegiDate	datetime,
		@CardCode	nvarchar(20),
		@CardName	nvarchar(100),
		@CarNo		nvarchar(30),
		@TransCom	nvarchar(150),
		@DeliArea	nvarchar(150),
		@Fee		numeric(19,6),
		@Comments	nvarchar(254),
		@ItemCode	nvarchar(20),
		@ItemName	nvarchar(200),
		@Size		nvarchar(200),
		@ItemType	nvarchar(5),
		@Name		nvarchar(200),
		@Mark		nvarchar(200),
		@Qty		numeric(19,6),
		@Weight		numeric(19,6)

		,@LoopCnt	int
		,@Index		int
		,@NoCnt		int
		
		
DECLARE CUR0 CURSOR LOCAL FOR

select	/*Header*/
		a.DocNum,
		a.U_MovDocNo,
		a.U_BPLId,
		c.BPLName,
		a.U_RegiDate,
		a.U_CardCode,
		a.U_CardName,
		a.U_CarNo,
		a.U_TransCom,
		a.U_DeliArea,
		a.U_Fee,
		a.U_Comments,
		/*Line*/
		b.U_ItemCode,
		b.U_ItemName,
		b.U_Size,
		d.U_ItemType,
		e.Name,
		b.U_Mark,
		b.U_Qty,
		b.U_Weight
from [@PS_PP075H] a inner join [@PS_PP075L] b on a.docentry=b.docentry
					left  join [OBPL] c on a.U_BPLId=c.BPLId
					left  join [OITM] d on b.U_ItemCode=d.ItemCode
					left  join [@PSH_SHAPE] e on d.U_ItemType=e.Code

where DocNum = @_DocNum

set @LoopCnt=1
set @NoCnt	=1

OPEN CUR0
FETCH NEXT FROM CUR0 INTO @DocNum,@MovDocNo,@BPLId,@BPLName,@RegiDate,@CardCode,@CardName,@CarNo,@TransCom,@DeliArea,@Fee,@Comments,@ItemCode,@ItemName,@Size,@ItemType,@Name,@Mark,@Qty,@Weight
WHILE (@@FETCH_STATUS = 0) BEGIN

	if @LoopCnt > 10 SET @LoopCnt = 1
						
	if @LoopCnt = 1 --//첫번째이면 INSERT
		begin
			insert into #PS_PP075 (DocNum,MovDocNo,BPLId,BPLName,RegiDate,CardCode,CardName,CarNo,TransCom,DeliArea,Fee,Comments,
								   No01,ItemCode01,ItemName01,Size01,TypeCode01,TypeName01,Mark01,Box01,Qty01,Weight01)
			values(@DocNum,@MovDocNo,@BPLId,@BPLName,@RegiDate,@CardCode,@CardName,@CarNo,@TransCom,@DeliArea,@Fee,@Comments,
				   @NoCnt,@ItemCode,@ItemName,@Size,@ItemType,@Name,@Mark,1,@Qty,@Weight)
								   
			set @Index = (select max(Seq) from #PS_PP075)
		end
		
	else if @LoopCnt = 2 --//두번째부터는 현재행에 update, 10번째까지								   
		begin
			update #PS_PP075
			set No02=@NoCnt,
				ItemCode02=@ItemCode,
				ItemName02=@ItemName,
				Size02=@Size,
				TypeCode02=@ItemType,
				TypeName02=@Name,
				Mark02=@Mark,
				Box02=1,
				Qty02=@Qty,
				Weight02=@Weight
			where Seq = @Index
		end
	else if @LoopCnt = 3
		begin
			update #PS_PP075
			set No03=@NoCnt,
				ItemCode03=@ItemCode,
				ItemName03=@ItemName,
				Size03=@Size,
				TypeCode03=@ItemType,
				TypeName03=@Name,
				Mark03=@Mark,
				Box03=1,
				Qty03=@Qty,
				Weight03=@Weight
			where Seq = @Index
		end
		
	else if @LoopCnt = 4
		begin
			update #PS_PP075
			set No04=@NoCnt,
				ItemCode04=@ItemCode,
				ItemName04=@ItemName,
				Size04=@Size,
				TypeCode04=@ItemType,
				TypeName04=@Name,
				Mark04=@Mark,
				Box04=1,				
				Qty04=@Qty,
				Weight04=@Weight
			where Seq = @Index
		end		
	else if @LoopCnt = 5
		begin
			update #PS_PP075
			set No05=@NoCnt,
				ItemCode05=@ItemCode,
				ItemName05=@ItemName,
				Size05=@Size,
				TypeCode05=@ItemType,
				TypeName05=@Name,
				Mark05=@Mark,
				Box05=1,				
				Qty05=@Qty,
				Weight05=@Weight
			where Seq = @Index
		end		
	else if @LoopCnt = 6
		begin
			update #PS_PP075
			set No06=@NoCnt,
				ItemCode06=@ItemCode,
				ItemName06=@ItemName,
				Size06=@Size,
				TypeCode06=@ItemType,
				TypeName06=@Name,
				Mark06=@Mark,
				Box06=1,				
				Qty06=@Qty,
				Weight06=@Weight
			where Seq = @Index
		end		
	else if @LoopCnt = 7
		begin
			update #PS_PP075
			set No07=@NoCnt,
				ItemCode07=@ItemCode,
				ItemName07=@ItemName,
				Size07=@Size,
				TypeCode07=@ItemType,
				TypeName07=@Name,
				Mark07=@Mark,
				Box07=1,				
				Qty07=@Qty,
				Weight07=@Weight
			where Seq = @Index
		end		
	else if @LoopCnt = 8
		begin
			update #PS_PP075
			set No08=@NoCnt,
				ItemCode08=@ItemCode,
				ItemName08=@ItemName,
				Size08=@Size,
				TypeCode08=@ItemType,
				TypeName08=@Name,
				Mark08=@Mark,
				Box08=1,				
				Qty08=@Qty,
				Weight08=@Weight
			where Seq = @Index
		end		
	else if @LoopCnt = 9
		begin
			update #PS_PP075
			set No09=@NoCnt,
				ItemCode09=@ItemCode,
				ItemName09=@ItemName,
				Size09=@Size,
				TypeCode09=@ItemType,
				TypeName09=@Name,
				Mark09=@Mark,
				Box09=1,
				Qty09=@Qty,
				Weight09=@Weight
			where Seq = @Index
		end		
	else if @LoopCnt = 10	--//10번까지는UPDATE
		begin
			update #PS_PP075
			set No10=@NoCnt,
				ItemCode01=@ItemCode,
				ItemName10=@ItemName,
				Size10=@Size,
				TypeCode10=@ItemType,
				TypeName10=@Name,
				Mark10=@Mark,
				Box10=1,
				Qty10=@Qty,
				Weight10=@Weight
			where Seq = @Index
		end		
	else
		begin
			set @LoopCnt = 1
		end
		
	set @LoopCnt = @LoopCnt +1
	set @NoCnt   = @NoCnt   +1		

FETCH NEXT FROM CUR0 INTO @DocNum,@MovDocNo,@BPLId,@BPLName,@RegiDate,@CardCode,@CardName,@CarNo,@TransCom,@DeliArea,@Fee,@Comments,@ItemCode,@ItemName,@Size,@ItemType,@Name,@Mark,@Qty,@Weight
END

CLOSE CUR0
DEALLOCATE CUR0

-----------------------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------------------------------------------
--select * from #PS_PP075

--//결과 Data
select	Seq,
		DocNum,
		MovDocNo,
		BPLId,
		BPLName,
		RegiDate,
		CardCode,
		CardName,
		CarNo,
		TransCom,
		DeliArea,
		Fee,
		Comments,
		
		No01,ItemCode01,ItemName01,Size01,TypeCode01,TypeName01,Mark01,Box01,Qty01,Weight01,
		No02,ItemCode02,ItemName02,Size02,TypeCode02,TypeName02,Mark02,Box02,Qty02,Weight02,
		No03,ItemCode03,ItemName03,Size03,TypeCode03,TypeName03,Mark03,Box03,Qty03,Weight03,
		No04,ItemCode04,ItemName04,Size04,TypeCode04,TypeName04,Mark04,Box04,Qty04,Weight04,
		No05,ItemCode05,ItemName05,Size05,TypeCode05,TypeName05,Mark05,Box05,Qty05,Weight05,
		No06,ItemCode06,ItemName06,Size06,TypeCode06,TypeName06,Mark06,Box06,Qty06,Weight06,
		No07,ItemCode07,ItemName07,Size07,TypeCode07,TypeName07,Mark07,Box07,Qty07,Weight07,
		No08,ItemCode08,ItemName08,Size08,TypeCode08,TypeName08,Mark08,Box08,Qty08,Weight08,
		No09,ItemCode09,ItemName09,Size09,TypeCode09,TypeName09,Mark09,Box09,Qty09,Weight09,
		No10,ItemCode10,ItemName10,Size10,TypeCode10,TypeName10,Mark10,Box10,Qty10,Weight10,
		
		isnull(Box01,0)+isnull(Box02,0)+isnull(Box03,0)+isnull(Box04,0)+isnull(Box05,0)
		+isnull(Box06,0)+isnull(Box07,0)+isnull(Box08,0)+isnull(Box09,0)+isnull(Box10,0) as TotalBox,
		
		isnull(Qty01,0)+isnull(Qty02,0)+isnull(Qty03,0)+isnull(Qty04,0)+isnull(Qty05,0)
		+isnull(Qty06,0)+isnull(Qty07,0)+isnull(Qty08,0)+isnull(Qty09,0)+isnull(Qty10,0) as TotalQty,

		isnull(Weight01,0)+isnull(Weight02,0)+isnull(Weight03,0)+isnull(Qty04,0)+isnull(Weight05,0)
		+isnull(Weight06,0)+isnull(Weight07,0)+isnull(Weight08,0)+isnull(Weight09,0)+isnull(Weight10,0) as TotalWeight
		
 FROM #PS_PP075
-----------------------------------------------------------------------------------------------------------------------------------------
--EXEC PS_PP075_02 '7'