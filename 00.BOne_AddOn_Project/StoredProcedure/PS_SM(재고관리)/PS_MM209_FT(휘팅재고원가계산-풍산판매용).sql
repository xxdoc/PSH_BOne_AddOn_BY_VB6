USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_CO210_01]    Script Date: 08/24/2011 11:43:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : 원가																				    */
/*  Description    : 휘팅재고원가계산																					*/
/*  질의관리자>REPORT-재고>590.수불명세서([MDC_InOut_QUERY_Detail] ) 참고하여 생성함							*/
/*  ALTER  Date    : 2011.08.26																					*/
/*  Modified Date  :																							*/
/*  Creator        : N.G.Y		                                                                                */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/



--Create PROC [dbo].[PS_CO210_01]
ALTER     PROC [dbo].[PS_CO210_01]
(
  
  @YM              as char(6),
  @AddAmt			as Numeric(12,0)
  
)
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------

DECLARE   @BPLId     Nvarchar(5),
		  @FrDate	 datetime,
		  @ToDate	 datetime,
		  @AcctCode	 Nvarchar(15),
		  @WareHouse Nvarchar(10),
		  @ItmBsort  Nvarchar(10),
		  @ItmMsort  Nvarchar(10)

set @BPLId = '1' --창원사업장
Set @AcctCode = '11502100' --제품
Set @WareHouse = '000' -- 창고
Set @ItmBsort = '101'
Set @ItmMsort = '%'

DECLARE   @DocDate     datetime,
          @ItemCode    nvarchar(20),
          @WhsCode     nvarchar(8),
          @Quantity    numeric(19,6),
          @Amount      numeric(19,6),
          @OutQty      numeric(19,6),
          @OutAmt      numeric(19,6)

Declare @totiamt  numeric(19,0),  --총 입고 제조원가
		@totiwgt numeric(19,3),   --총 입고중량
		@Flangeamt numeric(19,0),  --후렌지 철판금액
		@danga	  numeric(19,6)--,	   --입고평균단가(kg당)
		--@chamt	   numeric(19,0)	--배부 차이금액

SELECT @FrDate=F_RefDate,@ToDate=T_RefDate FROM OFPR WHERE CONVERT(CHAR(6),F_RefDate,112) = @YM
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--EXEC MDCp_AddOnTransType @FrDate,@ToDate
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
select * INTO #ZOINM from OINM where DocDate <= @ToDate and isnull(ApplObj,0) <> '911'

--if @BPLId='1' and @WareHouse='000' and @AcctCode='11506100' begin --임시:사업장-창원,창고-전체,재고계정-원재료일경우 재고이전(67) 제외
--	delete from #ZOINM where TransType = '67'
--end
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- 중량 수불명세서 (휘팅 적용)
--UPDATE a SET InQty=(InQty*b.U_UnWeight/1000),OutQty=(OutQty*b.U_UnWeight/1000)
--  FROM #ZOINM a
--  JOIN OITM b on b.ItemCode = a.ItemCode
-- WHERE b.U_ItmBSort = '101' AND b.U_UnWeight <> 0 AND @Wgt = 'Y'

----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
select 	a.ItemCode, a.FrgnName + Isnull((select Case When Isnull(Convert(nvarchar(20),Name),'') = '' Then '' Else '(' + Convert(nvarchar(20),Name) + ')' End  from [@PSH_QUALITY] Where Code = a.U_Quality),'')
                               + Isnull((select Case When Isnull(Convert(nvarchar(20),Name),'') = '' Then '' Else '(' + Convert(nvarchar(20),Name) + ')' End  from [@PSH_SHAPE] Where Code = a.U_ItemType),'')
                               + Isnull((select Case When Isnull(Convert(nvarchar(20),Name),'') = '' Then '' Else '(' + Convert(nvarchar(20),Name) + ')' End  from [@PSH_MARK] Where Code = a.U_MARK),'') As FrgnName ,
		a.InvntryUom ,
		convert(char,a.U_Size) as Size,

		--//*이월수량.금액*//--
		ISNULL(a.iwqty,0) as iwqty,
		Round(ISNULL(a.iwqty,0) * b.U_UnWeight / 1000,3) As iwwgt, --중량
		ISNULL(a.iwamt,0) as iwamt,
		
		--//*정상입고 수량*//--
		--<<원 재 료>>--
		case when @AcctCode='11506100' then ISNULL(a.i1qty,0) --구매
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.i2qty,0) --생산
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.i1qty,0) --구매
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.i1qty,0) --구매
			 end as i1qty,
			 
		--정상입고 중량
		Round(case when @AcctCode='11506100' then ISNULL(a.i1qty,0) --구매
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.i2qty,0) --생산
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.i1qty,0) --구매
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.i1qty,0) --구매
			 end * b.U_UnWeight / 1000,3) as i1wgt,
			 
		--//*정상입고 금액*//--
		--<<원 재 료>>--
		case when @AcctCode='11506100' then ISNULL(a.i1amt,0) --구매
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.i2amt,0) --생산
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.i1amt,0) --구매
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.i1amt,0) --구매
			 end as i1amt,
			 
		--//*타계정입고 수량*//--
		--<<원 재 료>>--
		case when @AcctCode='11506100' then ISNULL(a.i2qty,0)+ISNULL(a.i3qty,0)+ISNULL(a.i4qty,0)+ISNULL(a.i5qty,0) --생산,타계정,실사,기타입고
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.i1qty,0)+ISNULL(a.i3qty,0)+ISNULL(a.i4qty,0)+ISNULL(a.i5qty,0) --구매,타계정,실사,기타입고
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.i2qty,0)+ISNULL(a.i3qty,0)+ISNULL(a.i4qty,0)+ISNULL(a.i5qty,0) --생산,타계정,실사,기타입고
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.i2qty,0)+ISNULL(a.i3qty,0)+ISNULL(a.i4qty,0)+ISNULL(a.i5qty,0) --생산,타계정,실사,기타입고
			 end as i2qty,
			 
		--타계정입고 중량
		Round(case when @AcctCode='11506100' then ISNULL(a.i2qty,0)+ISNULL(a.i3qty,0)+ISNULL(a.i4qty,0)+ISNULL(a.i5qty,0) --생산,타계정,실사,기타입고
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.i1qty,0)+ISNULL(a.i3qty,0)+ISNULL(a.i4qty,0)+ISNULL(a.i5qty,0) --구매,타계정,실사,기타입고
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.i2qty,0)+ISNULL(a.i3qty,0)+ISNULL(a.i4qty,0)+ISNULL(a.i5qty,0) --생산,타계정,실사,기타입고
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.i2qty,0)+ISNULL(a.i3qty,0)+ISNULL(a.i4qty,0)+ISNULL(a.i5qty,0) --생산,타계정,실사,기타입고
			 end * b.U_UnWeight / 1000, 3) as i2wgt,			 
			 
		--//*타계정입고 금액*//--
		--<<원 재 료>>--
		case when @AcctCode='11506100' then ISNULL(a.i2amt,0)+ISNULL(a.i3amt,0)+ISNULL(a.i4amt,0)+ISNULL(a.i5amt,0) --생산,타계정,실사,기타입고
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.i1amt,0)+ISNULL(a.i3amt,0)+ISNULL(a.i4amt,0)+ISNULL(a.i5amt,0) --구매,타계정,실사,기타입고
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.i2amt,0)+ISNULL(a.i3amt,0)+ISNULL(a.i4amt,0)+ISNULL(a.i5amt,0) --생산,타계정,실사,기타입고
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.i2amt,0)+ISNULL(a.i3amt,0)+ISNULL(a.i4amt,0)+ISNULL(a.i5amt,0) --생산,타계정,실사,기타입고
			 end as i2amt,			 
			 
		--//*정상출고 수량*//--
		--<<원 재 료>>--
		case when @AcctCode='11506100' then ISNULL(a.o2qty,0) --생산
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.o1qty,0) --판매
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.o1qty,0) --판매
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.o2qty,0) --생산
			 end as o1qty,
			 
		--정상출고 중량
		Round(case when @AcctCode='11506100' then ISNULL(a.o2qty,0) --생산
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.o1qty,0) --판매
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.o1qty,0) --판매
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.o2qty,0) --생산
			 end * b.U_UnWeight / 1000, 3) as o1wgt,
			 
		--//*정상출고 금액*//--
		--<<원 재 료>>--
		case when @AcctCode='11506100' then ISNULL(a.o2amt,0) --생산
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.o1amt,0) --판매
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.o1amt,0) --판매
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.o2amt,0) --생산
			 end as o1amt,
			 
		--//*타계정출고 수량*//--
		--<<원 재 료>>--
		case when @AcctCode='11506100' then ISNULL(a.o1qty,0)+ISNULL(a.o3qty,0)+ISNULL(a.o4qty,0)+ISNULL(a.o5qty,0) --판매,타계정,실사,기타출고
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.o2qty,0)+ISNULL(a.o3qty,0)+ISNULL(a.o4qty,0)+ISNULL(a.o5qty,0) --생산,타계정,실사,기타출고
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.o2qty,0)+ISNULL(a.o3qty,0)+ISNULL(a.o4qty,0)+ISNULL(a.o5qty,0) --생산,타계정,실사,기타출고
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.o1qty,0)+ISNULL(a.o3qty,0)+ISNULL(a.o4qty,0)+ISNULL(a.o5qty,0) --판매,타계정,실사,기타출고
			 end as o2qty,	
			 
		--타계정출고 중량	 	
		Round(case when @AcctCode='11506100' then ISNULL(a.o1qty,0)+ISNULL(a.o3qty,0)+ISNULL(a.o4qty,0)+ISNULL(a.o5qty,0) --판매,타계정,실사,기타출고
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.o2qty,0)+ISNULL(a.o3qty,0)+ISNULL(a.o4qty,0)+ISNULL(a.o5qty,0) --생산,타계정,실사,기타출고
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.o2qty,0)+ISNULL(a.o3qty,0)+ISNULL(a.o4qty,0)+ISNULL(a.o5qty,0) --생산,타계정,실사,기타출고
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.o1qty,0)+ISNULL(a.o3qty,0)+ISNULL(a.o4qty,0)+ISNULL(a.o5qty,0) --판매,타계정,실사,기타출고
			 end * b.U_UnWeight / 1000, 3) as o2wgt,
			 
		--//*타계정출고 금액*//--
		--<<원 재 료>>--
		case when @AcctCode='11506100' then ISNULL(a.o1amt,0)+ISNULL(a.o3amt,0)+ISNULL(a.o4amt,0)+ISNULL(a.o5amt,0) --판매,타계정,실사,기타출고
		--<<제    품>>--
			 when @AcctCode='11502100' then ISNULL(a.o2amt,0)+ISNULL(a.o3amt,0)+ISNULL(a.o4amt,0)+ISNULL(a.o5amt,0) --생산,타계정,실사,기타출고
		--<<상    품>>--			 
			 when @AcctCode='11501100' then ISNULL(a.o2amt,0)+ISNULL(a.o3amt,0)+ISNULL(a.o4amt,0)+ISNULL(a.o5amt,0) --생산,타계정,실사,기타출고
		--<<저 장 품>>--			 
			 when @AcctCode='11507100' then ISNULL(a.o1amt,0)+ISNULL(a.o3amt,0)+ISNULL(a.o4amt,0)+ISNULL(a.o5amt,0) --판매,타계정,실사,기타출고
			 end as o2amt,					 

		ISNULL(a.jgqty,0) as jgqty,
		Round(ISNULL(a.jgqty,0) * b.U_UnWeight / 1000,3) as jgwgt,
		ISNULL(a.jgamt,0) as jgamt

 Into #MM209_ft
FROM
(
	SELECT	 t0.ItemCode, t1.ItemName, FrgnName = t1.FrgnName, t1.U_Quality, t1.InvntryUom, t1.U_Size, t1.U_Mark, t1.U_ItemType,

			--//이월
			 sum(case when t0.docdate < @FrDate then isnull(t0.inqty,0)-isnull(t0.outqty,0) end) as iwqty,
			 sum(case when t0.docdate < @FrDate then round(t0.transvalue,2) end) as iwamt,

			--//입고
			 --1.구매
			 sum(case when t0.docdate >= @FrDate and t0.transtype in (18,19,20,21,69,162) then isnull(t0.inqty,0)-isnull(t0.outqty,0) end) as i1qty,
			 sum(case when t0.docdate >= @FrDate and t0.transtype in (18,19,20,21,69,162,931) then round(t0.transvalue,2) end) as i1amt,
			 --2.생산
			 sum(case when t0.docdate >= @FrDate and t0.transtype in (59,60) and t0.applobj IN (202,901) and t0.AppObjType = 'P' then isnull(t0.inqty,0)-isnull(t0.outqty,0) end) as i2qty,
			 sum(case when t0.docdate >= @FrDate and t0.transtype in (59,60) and t0.applobj IN (202,901) and t0.AppObjType = 'P' then round(t0.transvalue,2) end) as i2amt,
			 --3.재고이전
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 67 and isnull(t0.inqty,0) > 0 then isnull(t0.inqty,0) end) as i3qty,
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 67 and isnull(t0.inqty,0) > 0 then round(t0.transvalue,2) end) as i3amt,
			 --4.재고실사
			 --sum(case when t0.docdate >= @FrDate and t0.transtype = 10000071 and isnull(t0.inqty,0) > 0 then isnull(t0.inqty,0) end) as i4qty,
			 --sum(case when t0.docdate >= @FrDate and t0.transtype = 10000071 and isnull(t0.inqty,0) > 0 then round(t0.transvalue,2) end) as i4amt,
			 0 as i4qty,
			 0 as i4amt,
			 --5.기타입고
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 59 and isnull(t0.applobj,0) NOT IN (202,901) then isnull(t0.inqty,0) end) as i5qty,
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 59 and isnull(t0.applobj,0) NOT IN (202,901) then round(t0.transvalue,2) end) as i5amt,

		
			--//출고
			 --1.판매
			 sum(case when t0.docdate >= @FrDate and t0.transtype in (13,14,15,16) then isnull(t0.outqty,0)-isnull(t0.inqty,0) end) as o1qty,
			 sum(case when t0.docdate >= @FrDate and t0.transtype in (13,14,15,16,932) then round(-t0.transvalue,2) end) as o1amt,
			 --2.생산
			 sum(case when t0.docdate >= @FrDate and t0.transtype in (59,60) and t0.applobj IN (202,901) and t0.AppObjType = 'C' then isnull(t0.outqty,0)-isnull(t0.inqty,0) end) as o2qty,
			 sum(case when t0.docdate >= @FrDate and t0.transtype in (59,60) and t0.applobj IN (202,901) and t0.AppObjType = 'C' then round(-t0.transvalue,2) end) as o2amt,
			 --3.재고이전
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 67 and isnull(t0.outqty,0) > 0 then isnull(t0.outqty,0) end) as o3qty,
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 67 and isnull(t0.outqty,0) > 0 then round(-t0.transvalue,2) end) as o3amt,
			 --4.재고실사
			 --sum(case when t0.docdate >= @FrDate and t0.transtype = 10000071 and isnull(t0.outqty,0) > 0	then isnull(t0.outqty,0) end) as o4qty,
			 --sum(case when t0.docdate >= @FrDate and t0.transtype = 10000071 and isnull(t0.outqty,0) > 0 then round(-t0.transvalue,2) end) as o4amt,
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 10000071 then isnull(t0.outqty-t0.inqty,0) end) as o4qty,
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 10000071 then round(-t0.transvalue,2) end) as o4amt,
			 --5.기타출고
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 60 and isnull(t0.applobj,0) NOT IN (202,901) then isnull(t0.outqty,0) end) as o5qty,
			 sum(case when t0.docdate >= @FrDate and t0.transtype = 60 and isnull(t0.applobj,0) NOT IN (202,901) then round(-t0.transvalue,2) end) as o5amt,

			--//재고
			 sum(isnull(t0.inqty,0)-isnull(t0.outqty,0)) as jgqty,
			 sum(round(t0.transvalue,2)) as jgamt

--	FROM OINM t0 
	FROM #ZOINM t0
	JOIN OITM t1 ON t1.ItemCode = t0.ItemCode
	JOIN OITB t2 ON t2.ItmsGrpCod = t1.ItmsGrpCod
	JOIN OWHS t3 ON t3.WhsCode = t0.warehouse
    LEFT JOIN OWHS t4 ON t4.WhsCode = t0.Ref2 AND t0.TransType = 67
	
	WHERE t0.docdate <= @ToDate
	  AND t2.U_InvntAct = @AcctCode
	  --포장사업팀의 재고를 사업장별로 보기위해 수정 : 6 => 사상 + 온산
	  And (Case When @BPLId = '3' Or @BPLId = '5' Then right(t3.WhsCode,1) Else t3.BPLId End = (Case When @BPLId = '6' Then '3' Else @BPLId End) Or @BPLId = '0')
	  --AND (t3.BPLId = @BPLId or @BPLId = '0')
	  AND (t0.warehouse = @WareHouse or @WareHouse = '000')
	  AND t1.U_ItmBsort like @ItmBsort + '%'
	  AND t1.U_ItmMsort like @ItmMsort + '%'
	  --AND t3.BPLId <> ISNULL(t4.BPLId,'')
	  --AND t0.ItemCode = '5A0100043'
	Group By t0.ItemCode, t1.ItemName, t1.FrgnName, t1.U_Quality, t1.InvntryUom, t1.U_Size, U_Mark, t1.U_ItemType
	--GROUP BY t0.itemcode,t1.ItemName,t1.FrgnName,t1.InvntryUom,t1.U_Size

) a Inner Join OITM b On a.ItemCode = b.ItemCode

WHERE (isnull(a.iwqty,0)<>0 or isnull(a.iwamt,0)<>0 or isnull(a.i1qty,0)<>0 or isnull(a.i1amt,0)<>0 or isnull(a.i2qty,0)<>0 or isnull(a.i2amt,0)<>0 or 
       isnull(a.i3qty,0)<>0 or isnull(a.i3amt,0)<>0 or isnull(a.i4qty,0)<>0 or isnull(a.i4amt,0)<>0 or isnull(a.i5qty,0)<>0 or isnull(a.i5amt,0)<>0 or 
       isnull(a.o1qty,0)<>0 or isnull(a.o1amt,0)<>0 or isnull(a.o2qty,0)<>0 or isnull(a.o2amt,0)<>0 or isnull(a.o3qty,0)<>0 or isnull(a.o3amt,0)<>0 or 
       isnull(a.o4qty,0)<>0 or isnull(a.o4amt,0)<>0 or isnull(a.o5qty,0)<>0 or isnull(a.o5amt,0)<>0 or isnull(a.jgqty,0)<>0 or isnull(a.jgamt,0)<>0 )

ORDER BY a.ItemCode

--총입고 제조원가 + 추가비용
Select @totiamt = Isnull(Sum(i1amt),0) + Isnull(@AddAmt,0)
 From #MM209_ft a Inner Join OITM b On a.ItemCode = b.ItemCode And b.U_ItmBsort = '101'
 
 --총입고 중량
Select @totiwgt = Isnull(Sum(i1wgt),0)
 From #MM209_ft a Inner Join OITM b On a.ItemCode = b.ItemCode And b.U_ItmBsort = '101'

--후렌지 철판금액
Select @Flangeamt = Sum(a.i1qty * c.Price)
from #MM209_ft a Inner Join OITM b On a.ItemCode = b.ItemCode And b.U_ItmBsort = '101' and b.U_ItmMsort = '10107'
				 Inner Join Z_FLANGE_PRICE_TEMP c On a.ItemCode = c.ItemCode
Where a.i1qty <> 0

--총입고 제조원가 금액에서 철판금액 차감
Set @totiamt = @totiamt - @Flangeamt

--kg당 입고단가
Set @danga = round(@totiamt / @totiwgt ,6)

Update #MM209_ft
   set jgamt = Round(jgWgt * @danga,0)
 From OITM a
where #MM209_ft.ItemCode = a.ItemCode
  and a.U_ItmBsort = '101'

--후렌지 철판단가 적용
Update #MM209_ft
   set jgamt = jgamt + (jgqty * b.Price)
  From OITM a,
	   Z_FLANGE_PRICE_TEMP b
 Where #MM209_ft.ItemCode = a.ItemCode
   And a.ItemCode = b.ItemCode
   And a.U_ItmBsort = '101' and a.U_ItmMsort = '10107'
   And #MM209_ft.jgqty > 0

--Select @chamt = Sum(jgamt) - (@totiamt +  @Flangeamt)
--  From #MM209_ft a Inner Join OITM b On a.ItemCode = b.ItemCode And b.U_ItmBsort = '101'

--Update #MM209_ft
--   set jgamt = jgamt - @chamt
-- Where #MM209_ft.ItemCode = (select max(a.ItemCode) From #MM209_ft a Inner Join OITM b On a.ItemCode = b.ItemCode And b.U_ItmBsort = '101' and a.jgamt <> 0)
 
--select @totiwgt, @totiamt, @danga
select a.ItemCode,
	   b.ItemName, 
	   Qty = a.jgqty, 
	   Wgt = a.jgwgt, 
	   QPrice = Case When a.jgqty > 0 Then round(a.jgamt / a.jgqty,0) Else 0 End, 
	   WPrice = Case When a.jgwgt > 0 Then round(a.jgamt / a.jgwgt,0) Else 0 End, 
	   Amt = a.jgamt
	   --Sum(a.jgamt)
 from #MM209_ft a Inner Join OITM b On a.ItemCode = b.ItemCode And b.U_ItmBsort = '101'
 Where a.jgqty > 0
--select a.ItemCode,
--	   a.FrgnName,
--	   a.InvntryUom,
--	   a.Size,
--	   a.iwqty,
--	   a.iwwgt,
--	   a.iwamt,
--	   a.i1qty,
--	   a.i1wgt,
--	   a.i1amt,
--	   a.i2qty,
--	   a.i2wgt,
--	   a.i2amt,
--	   a.o1qty,
--	   a.o1wgt,
--	   a.o1amt,
--	   a.o2qty,
--	   a.o2wgt,
--	   a.o2amt,
--	   a.jgqty,
--	   a.jgwgt,
--	   a.jgamt
-- from #MM209_ft a Inner Join OITM b On a.ItemCode = b.ItemCode And b.U_ItmBsort = '101'

			

--Select WhsCode, WhsName, BPLId from OWHS
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_CO210_01] '201107',0 --제품

--@BPLId			as Nvarchar(5),
--  @FrDate			as datetime,
--  @ToDate			as datetime,
--  @AcctCode			as Nvarchar(15),
--  @WareHouse		as Nvarchar(10),
--  @Wgt              as char(1),
--  @ItmBsort			as Nvarchar(10),
--  @ItmMsort

