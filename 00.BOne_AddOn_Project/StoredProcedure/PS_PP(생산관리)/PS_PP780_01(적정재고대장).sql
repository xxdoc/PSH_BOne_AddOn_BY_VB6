USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP780_01]    Script Date: 03/24/2011 13:03:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************************************************************************************************************/
/*  Module         : 생산관리																				    */
/*  Description    : 적정재고대장  																				*/
/*  ALTER  Date    : 2011.03.24  																				*/
/*  Modified Date  :																							*/
/*  Creator        : N.G.Y			                                                                            */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER  PROC [dbo].[PS_PP780_01]
--CREATE     PROC [dbo].[PS_PP780_01]
(
  @DocDate			as datetime,
  @ItmMsort			As Nvarchar(10),
  @Mark				As Nvarchar(10)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
declare @ymd2  char(8),
        @ymd6	char(8),
        @ymd12	char(8)

select @ymd2   = convert(char(8),dateadd(mm,-2,@DocDate),112)
select @ymd6   = convert(char(8),dateadd(mm,-6,@DocDate),112)
select @ymd12  = convert(char(8),dateadd(yy,-1,@DocDate),112)

CREATE TABLE #PS_PP780 
			(
             ItemCode        nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
             jqty			 numeric(12,0),
             bqty			 numeric(12,0),
             sqty			 numeric(12,0),
             chqty6			 numeric(12,0),
             chqty12		 numeric(12,0),
             chmaxqty		 numeric(12,0),
             chminqty		 numeric(12,0),
             pojqty			 numeric(12,0),
             apoqty			 numeric(12,0)
			)

--//서울재고 계산용
CREATE TABLE #PS_PP780_01 	(
			 MovDocNo		 nvarchar(20),
			 PP070No		 nvarchar(20),
             ItemCode        nvarchar(20),
             qty			 numeric(12,0),
             wgt			 numeric(12,3)
             )

--//현재고수량
Insert Into #PS_PP780( ItemCode, jqty)
select b.ItemCode, qty = SUM(a.InQty-a.OutQty) 
from OINM a left join OITM b on a.ItemCode=b.ItemCode
where a.DocDate <= @DocDate
and b.U_ItmBsort = '101'
and a.Warehouse like '%4'
AND isnull(a.ApplObj,0) <> '911'
And b.U_ItmMsort like @ItmMsort + '%'
And b.U_Mark like @Mark + '%'
 GROUP BY b.ItemCode
 
 --//본사출하대기
 Insert Into #PS_PP780( ItemCode, bqty)
 select b.U_ItemCode, qty = sum(b.U_SelQty)
from [@ps_pp070H] a,
	 [@ps_pp070L] b,
	 [OITM] c
where a.DocEntry = b.DocEntry
  and (Isnull(b.U_MovDocNo,'') = '' Or left(Isnull(b.U_MovDocNo,''),8) >= @DocDate)
  and a.Canceled = 'N'
  and b.U_ItemCode = c.ItemCode
  And c.U_ItmMsort like @ItmMsort + '%'
  And c.U_Mark like @Mark + '%'
group by b.U_ItemCode

--//서울이동분
Insert into #PS_PP780_01( MovDocNo, PP070No, ItemCode, qty, wgt)
select	a.U_MovDocNo,
		b.U_PP070No,
			b.U_ItemCode,
			sum(b.U_Qty) Qty,
			sum(b.U_Weight) Weight
	from [@PS_PP075H] a inner join [@PS_PP075L] b on a.docentry=b.docentry
						inner Join [OITM] c on b.U_ItemCode = c.ItemCode
	where a.Canceled <> 'Y'
	  and a.U_RegiDate < @DocDate --이동등록일
	  And c.U_ItmMsort like @ItmMsort + '%'
	  And c.U_Mark like @Mark + '%'
group by a.U_MovDocNo,
		 b.U_PP070No,
		 b.U_ItemCode

--//서울 포장분

Insert into #PS_PP780_01( MovDocNo, PP070No, ItemCode, qty, wgt)
select	a.U_MovDocNo,
		a.U_PP070No + '-' + a.U_PP070NoL As PP070No,
			a.U_ItemCode,
			sum(a.U_NPkQty) * -1 As  Qty ,
			sum(a.U_NPkWt) * -1 As Weight
	from [@PS_PP077H] a,
		 [OITM] b
	where a.Canceled <> 'Y'
	  and a.U_InDate < @DocDate --이동등록일
	  and a.U_ItemCode = b.ItemCode
	  And b.U_ItmMsort like @ItmMsort + '%'
	  And b.U_Mark like @Mark + '%'
group by a.U_MovDocNo,
		 a.U_PP070No,
		 a.U_PP070NoL,
		 a.U_ItemCode

--//반품등록
Insert into #PS_PP780_01( MovDocNo, PP070No, ItemCode, qty, wgt)
select	b.U_MovDocNo,
		b.U_PP070No,
		b.U_ItemCode,
		Isnull(sum(b.U_PkQty - b.U_OPkQty),0) * -1 As Qty,
		Isnull(sum(b.U_PkWt - b.U_OPkWt),0) * -1 As Weight
	from [@PS_PP777H] a inner join [@PS_PP777L] b on a.docentry=b.docentry				
						inner join [OITM] c On b.U_ItemCode = c.ItemCode
	where a.Canceled <> 'Y'
	  and a.U_DocDate < @DocDate --이동등록일
	  And c.U_ItmMsort like @ItmMsort + '%'
	  And c.U_Mark like @Mark + '%'
group by b.U_MovDocNo,
		 b.U_PP070No,
		 b.U_ItemCode

--//서울 포장대기
Insert Into #PS_PP780( ItemCode, sqty)
Select g.ItemCode, Isnull(Sum(g.qty),0)
from (
SELECT MovDocNo, PP070No, ItemCode, Isnull(sum(qty),0) As qty, Isnull(sum(wgt),0) As wgt 
FROM #PS_PP780_01
Group by MovDocNo, PP070No, ItemCode
having sum(wgt) > 0
) g
group by g.ItemCode

--//6개월 평균출고량
Insert Into #PS_PP780( ItemCode, chqty6)
Select b.U_ItemCode, round(Sum(b.U_Weight) / 6,0)
from [@PS_SD040H] a,
	 [@PS_SD040L] b,
	 [OITM] c
where a.DocEntry = b.DocEntry
  and b.U_ItemCode = c.ItemCode
  And c.U_ItmMsort like @ItmMsort + '%'
  And c.U_Mark like @Mark + '%'
  and b.U_ItmBsort = '101'
  and a.U_BPLId = '4'
  and a.Canceled = 'N'
  and a.U_DocDate Between @ymd6 and @DocDate
group by b.U_ItemCode

Insert Into #PS_PP780( ItemCode, chqty12)
Select b.U_ItemCode, round(Sum(b.U_Weight) / 12,0)
from [@PS_SD040H] a,
	 [@PS_SD040L] b,
	 [OITM] c
where a.DocEntry = b.DocEntry
  and b.U_ItemCode = c.ItemCode
  And c.U_ItmMsort like @ItmMsort + '%'
  And c.U_Mark like @Mark + '%'
  and b.U_ItmBsort = '101'
  and a.U_BPLId = '4'
  and a.Canceled = 'N'
  and a.U_DocDate Between @ymd6 and @DocDate
group by b.U_ItemCode

--//1년중 최고, 최소 출고수량
Insert Into #PS_PP780( ItemCode, chmaxqty, chminqty)
select g.U_ItemCode, max(qty), min(qty)
from (
Select ym = Convert(Char(6),a.U_DocDate,112), b.U_ItemCode, qty = Sum(b.U_Weight)
from [@PS_SD040H] a,
	 [@PS_SD040L] b,
	 [OITM] c
where a.DocEntry = b.DocEntry
  and b.U_ItemCode = c.ItemCode
  And c.U_ItmMsort like @ItmMsort + '%'
  And c.U_Mark like @Mark + '%'
  and b.U_ItmBsort = '101'
  and a.U_BPLId = '4'
  and a.Canceled = 'N'
  and a.U_DocDate Between @ymd12 and @DocDate
group by Convert(Char(6),a.U_DocDate,112), b.U_ItemCode
) g
group by g.U_ItemCode

--//3개월 po잔량 => 3개월PO량 - 창원생산 - 벌크생산
Insert Into #PS_PP780( ItemCode, pojqty)
select g.U_ItemCode, sum(g.qty)
From (
select a.U_ItemCode, qty = Sum(a.U_ReWeight)
from [@PS_SD010H] a,
	 [OITM] b
Where a.U_ItmBsort = '101'
  and a.U_ItemCode = b.ItemCode
  And b.U_ItmMsort like @ItmMsort + '%'
  And b.U_Mark like @Mark + '%'
  and a.Canceled = 'N'
  and a.U_DueDate between @ymd2 and @DocDate
  group by a.U_ItemCode
  
union all

select e.U_ItemCode, qty = sum(e.U_YQty) * -1
from [@PS_SD010H] a,
	 [@PS_PP030H] b,
	 [@PS_PP030M] c,
	 [@PS_PP080H] d,
	 [@PS_PP080L] e,
	 [OITM] f
WHERE a.U_RegNum = b.U_BaseNum
  and b.DocEntry = c.DocEntry
  and d.DocEntry = e.DocEntry
  and e.U_ItemCode = f.ItemCode
  And f.U_ItmMsort like @ItmMsort + '%'
  And f.U_Mark like @Mark + '%'
  and Convert(nvarchar(10),b.DocEntry) + '-' + Convert(nvarchar(10),c.U_Sequence) = e.U_PP030No
  and a.U_DueDate between @ymd2 and @DocDate
  and d.U_BPLId = '1'
  and a.U_ItmBsort = '101'
  and d.Canceled = 'N'
  group by e.U_ItemCode
  
  Union all
  
select e.U_ItemCode, qty = sum(e.U_SelQty) * -1
from [@PS_SD010H] a,
	 [@PS_PP030H] b,
	 [@PS_PP030M] c,
	 [@PS_PP070H] d,
	 [@PS_PP070L] e,
	 [OITM] f
WHERE a.U_RegNum = b.U_BaseNum
  and b.DocEntry = c.DocEntry
  and d.DocEntry = e.DocEntry
  and e.U_ItemCode = f.ItemCode
  And f.U_ItmMsort like @ItmMsort + '%'
  And f.U_Mark like @Mark + '%'
  and Convert(nvarchar(10),b.DocEntry) + '-' + Convert(nvarchar(10),c.U_Sequence) = e.U_PP030No
  and a.U_DueDate between @ymd2 and @DocDate
  and d.U_BPLId = '1'
  and a.U_ItmBsort = '101'
  group by e.U_ItemCode
  ) g
 Group by g.U_ItemCode 
  having sum(g.qty) > 0

--//3개월이후 PO량

Insert Into #PS_PP780( ItemCode, apoqty)
select a.U_ItemCode, qty = Sum(a.U_ReWeight)
from [@PS_SD010H] a,
	 [OITM] b
Where a.U_ItmBsort = '101'
  and a.U_ItemCode = b.ItemCode
  And b.U_ItmMsort like @ItmMsort + '%'
  And b.U_Mark like @Mark + '%'
  and a.Canceled = 'N'
  and a.U_DueDate between dateadd(mm,1, @DocDate) and dateadd(mm,3, @DocDate)
  group by a.U_ItemCode


----------------------------------------------------------------------------------------------------------------------------------------------------------------------------   
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------   
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
/* DATA RETURN */
Select t1.ItemCode,
	   t2.ItemName,
	   janil = Case When t1.chqty6 > 0 Then round(t1.jqty / (chqty6/22),1)
				Else 1 End , --잔여재고일
	   t1.jqty, --//현재고
	   t1.sqty, --//서울포장대기
	   t1.bqty, --//창원이동대기
	   totjqty = t1.jqty + t1.sqty + t1.bqty, --//총재고
	   jjqty = chqty6 * t2.U_Stkjj, --//적정재고
	   comqty = Case When chqty6 * t2.U_Stkjj > 0 Then t1.jqty + t1.sqty + t1.bqty - (chqty6 * t2.U_Stkjj)
				Else 1 End , --//과부족
	   t1.pojqty, --//3개월PO잔량
	   t1.apoqty, --//3개월이후 po량
	   t2.U_Stkjj, --//적정재고 배수
	   hjqty = Case When t1.chqty6 > 0 Then round((t1.jqty + t1.sqty + t1.bqty) / t1.chqty6,1)
				Else 1 End, --//적정현재고 배수
	   t1.chqty6, --//6개월평균출고량
	   t1.chqty12,--//12개월평균출고량
	   t1.chmaxqty, --//1년중 최대출고량
	   t1.chminqty --//1년중 최소출고량
from (
SELECT	a.ItemCode,
        jqty = Isnull(sum(a.jqty),0),
		bqty = Isnull(sum(a.bqty),0),
		sqty = Isnull(Sum(a.sqty),0),
		chqty6 = Isnull(sum(a.chqty6),0),
		chqty12 = Isnull(sum(a.chqty12),0),
		chmaxqty = Isnull(sum(a.chmaxqty),0),
		chminqty = Isnull(sum(a.chminqty),0),
		pojqty = Isnull(sum(a.pojqty),0),
		apoqty = Isnull(sum(a.apoqty),0)
		
FROM #PS_PP780 a
Group by a.ItemCode
) t1,
[OITM] t2
where t1.ItemCode = t2.ItemCode
  
----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_PP780_01] '20110325', '10102', '02'
