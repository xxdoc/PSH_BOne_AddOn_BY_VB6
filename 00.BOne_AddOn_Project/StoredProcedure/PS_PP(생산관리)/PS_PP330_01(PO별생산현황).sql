USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP330_01]    Script Date: 04/08/2011 15:17:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/******************************************************************************************************************/
/*  Module         : PP								    														*/
/*  Description    : PO별 생산현황(휘팅)  															*/
/*  Create Date    : 2011.07.25                                                                                   */
/*  Modified Date  :										       													*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/******************************************************************************************************************/
ALTER PROC [dbo].[PS_PP330_01]
--Create PROC [dbo].[PS_PP330_01]
(
	@ItmBsort   NvarChar(10),
	@ItmMsort   NvarChar(10),
	@DocDateFr  datetime,
	@DocDateTo  datetime
   
)
AS

BEGIN
Create Table #Temp01 (
	OrdNum		Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	DueDate		DateTime,
	ItemCode	Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	ItemName	Nvarchar(50),
	Size		Nvarchar(50),
	Mark		Nvarchar(10),
	MarkName	Nvarchar(10),
	ItemType	Nvarchar(10),
	TypeName	Nvarchar(10),
	CallSize	Nvarchar(20),
	Qty			Numeric(8,0),   
	Wgt			Numeric(19,6),
	jgQty		Numeric(19,6),
	jgWgt		Numeric(19,6),
	mkDate		DateTime,
	mkQty		Numeric(8,0),
	mkWgt		Numeric(19,6)
)

--본사이동대기, 서울포장대기
Create Table #Temp02 (
	OrdNum		Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	SEOULDE		Numeric(19,6),
	SEOULDEWt	Numeric(19,6),
	SEOULJE		Numeric(19,6),
	SEOULJEWt	Numeric(19,6)
)

Insert Into #Temp01( OrdNum, DueDate, ItemCode, ItemName, Size, Mark, MarkName, ItemType, TypeName, Callsize, Qty, Wgt)
select c.U_OrdNum,
	   a.U_DueDate, 
	   a.U_ItemCode, 
	   b.FrgnName,
	   b.U_Size,
	   b.U_Mark,
	   MarkName = (Select Name from [@PSH_MARK] Where Code = b.U_Mark),
	   b.U_ItemType,
	   TypeName = (Select Name from [@PSH_SHAPE] Where Code = b.U_ItemType),
	   b.U_CallSize,
	   a.U_ReWeight As Qty,
	   Round((a.U_ReWeight * b.U_UnWeight) / 1000,3) As Wgt
 From [@PS_SD010H] a Inner join [OITM] b On a.U_ItemCode = b.ItemCode
					 Left Join [@PS_PP030H] c On c.U_BaseNum = a.U_RegNum
where b.U_ItmBsort Like @ItmBsort + '%'
   And b.U_ItmMsort Like @ItmMsort + '%'
   And a.U_ItmBsort = '101'
   and a.U_DueDate between @DocDateFr and @DocDateTo


 --from [@PS_PP030H] a Inner Join [@PS_SD010H] b On a.U_BaseNum = b.U_RegNum
					 --Inner Join [OITM] c On a.U_ItemCode = c.ItemCode
  --where c.U_ItmBsort Like @ItmBsort + '%'
  -- And c.U_ItmMsort Like @ItmMsort + '%'
  -- And a.U_OrdGbn = '101'
  -- and b.U_DueDate between @DocDateFr and @DocDateTo


--본사출하대기
Insert into #Temp02 ( OrdNum, SEOULDE, SEOULDEWt )
select b.U_OrdNum, Isnull(Sum(b.U_SelQty),0), Isnull(Sum(b.U_SelWt),0)
from [@ps_pp070H] a Inner Join [@ps_pp070L] b On a.DocEntry = b.DocEntry
where Isnull(b.U_MovDocNo,'') = ''
  and a.Canceled = 'N'
  and Exists (select * from #Temp01 c Where b.U_OrdNum = c.OrdNum)
 Group by b.U_OrdNum

 --서울포장대기
Insert into #Temp02 ( OrdNum, SEOULJE, SEOULJEWt )
--서울이동
select c.U_OrdNum,
	   Isnull(Sum(b.U_Qty),0),
	   Isnull(Sum(b.U_Weight),0)
	from [@PS_PP075H] a inner join [@PS_PP075L] b on a.docentry=b.docentry
						inner Join [@PS_PP070L] c On b.U_PP070No = convert(nvarchar(10),c.DocEntry) + '-' + convert(nvarchar(10),c.U_LineId)
	where a.Canceled = 'N'
  and Exists (select * from #Temp01 d Where c.U_OrdNum = d.OrdNum)
   group by c.U_OrdNum	

Union All

select	b.U_OrdNum,
			Isnull(sum(a.U_NPkQty),0) * -1 As  Qty, 
			Isnull(sum(a.U_NPkWt),0) * -1 As  Wgt 
	from [@PS_PP077H] a Inner Join [@PS_PP070L] b On a.U_PP070No = b.DocEntry And a.U_PP070NoL = b.LineId
	where a.Canceled = 'N'
	and Exists (select * from #Temp01 c Where b.U_OrdNum = c.OrdNum)
group by b.U_OrdNum

Union all

select	c.U_OrdNum,
		Isnull(sum(b.U_PkQty - b.U_OPkQty),0) * -1 As Qty,
		Isnull(sum(b.U_PkWt - b.U_OPkWt),0) * -1 As Wgt
	from [@PS_PP777H] a inner join [@PS_PP777L] b on a.docentry=b.docentry				
						Inner Join [@PS_PP070L] c On b.U_PP070No = convert(nvarchar(10),c.DocEntry) + '-' + convert(nvarchar(10),c.U_LineId)
	where a.Canceled = 'N'
  and Exists (select * from #Temp01 d Where c.U_OrdNum = d.OrdNum)
group by c.U_OrdNum

--본사이동대기, 서울 포장대기 Update

--//포장대기 Update 
 Update #Temp01
    set jgQty = Isnull(g.SEOULDE,0) + Isnull(g.SEOULJE,0),
		jgWgt = Isnull(g.SEOULDEWt,0) + Isnull(g.SEOULJEWt,0)
   from (Select a.Ordnum,
				 SEOULDE = Isnull(Sum(a.SEOULDE),0),
				 SEOULDEWt = Isnull(Sum(a.SEOULDEWt),0),
				 SEOULJE = Isnull(Sum(SEOULJE),0),
				 SEOULJEWt = Isnull(Sum(SEOULJEWt),0)
		  from #Temp02 a
		 Group by a.Ordnum ) g
 Where #Temp01.OrdNum = g.OrdNum

--//생산일자, 수량, 중량 Update

Update #Temp01
   set mkDate = g.InDate,
	   mkQty = g.Qty,
	   mkWgt = g.Wgt
  from (Select a.U_PorNum As OrdNum,
			   Max(a.U_InDate) As InDate,
			   Sum(a.U_NpkQty) As Qty,
			   Sum(a.U_NpkWt) As Wgt
		  from [@PS_PP077H] a
--		 where a.Canceled = 'N'
		Group by a.U_PorNum ) g
  Where #Temp01.OrdNum = g.OrdNum


Select	a.OrdNum, --//작업지시번호
		a.DueDate, --//납기일
		a.ItemCode, --//품목코드 
		a.ItemName, --//품목명
		a.Size, --//규격
		a.MarkName, --//인증
		a.TypeName, --//형태
		a.CallSize, --//호칭
		a.Qty,  --//지시수량
		a.Wgt,  --//지시중량
		a.mkDate, --//포장완료일
		a.mkQty, --//생산수량
		a.mkWgt, --//생산중량
		a.jgQty, --//포장대기수량
		a.jgWgt, --//포장대기중량
		a.mkQty + a.jgQty As TotmkQty, --//합계수량
		a.mkWgt + a.jgWgt As TotmkWgt, --//합계중량
		a.Qty - (a.mkQty + a.jgQty) As NmkQty, --//포장기준 미완료수량
		a.Wgt - (a.mkWgt + a.jgWgt) As NmkWgt --//포장기준 미완료중량
		
  From	#Temp01 a Inner Join [OITM] b On a.ItemCode = b.ItemCode
 Order by b.U_ItmBsort, b.U_ItmMsort, b.FrgnName, b.U_Mark, b.U_ItemType,
		  Convert(Numeric(12,3),b.U_Spec1), Convert(Numeric(12,3),b.U_Spec2), Convert(Numeric(12,3),b.U_Spec3)
End

--  EXEC [dbo].[PS_PP330_01] '101', '%','20110501', '20110630'

--EXEC [dbo].[PS_PP330_01] '1','102', '%', '20100101','20110331', '%','%', ''

