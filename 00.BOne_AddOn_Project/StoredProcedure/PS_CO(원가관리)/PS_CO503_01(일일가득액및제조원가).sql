USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_CO503_01]    Script Date: 04/25/2011 23:36:24 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/******************************************************************************************************************/
/*  Module         : PP								    														*/
/*  Description    : 일일 가득액 및 생산원가				  															*/
/*  Create Date    : 2011.05.16                                                                                   */
/*  Modified Date  :										       													*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/******************************************************************************************************************/
--ALTER PROC [dbo].[PS_CO503_01]
ALTER PROC [dbo].[PS_CO503_01]
(
	@DocDate	Date
   
)
AS

BEGIN
--판매
Create Table #Temp01 (
	ItmBsort	Nvarchar(10) Collate Korean_Wansung_Unicode_CI_AS,
	ItmBname	Nvarchar(20),
	ItemCode	Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	ItemName	Nvarchar(50),
	Qty			Numeric(19,2), --판매수량
	Wgt			Numeric(19,2), --판매중량
	Amt			Numeric(19,0), --판매금액
	Wamt		Numeric(19,0), --원재료금액
	gaduk		Numeric(19,0)  --가득액
)

--생산
Create Table #Temp02 (
	ItmBsort	Nvarchar(10) Collate Korean_Wansung_Unicode_CI_AS,
	ItmBname	Nvarchar(20),
	ItemCode	Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	ItemName	Nvarchar(50),
	MQty		Numeric(12,0), --생산수량
	MWgt		Numeric(12,2),  --생산중량
	MAmt		Numeric(12,0)
)

Create Table #Temp03 (
	ItmBsort	Nvarchar(10) Collate Korean_Wansung_Unicode_CI_AS,
	Code		Nvarchar(10),
	Name		Nvarchar(20),
	Value		Numeric(12,0)
)

Create Table #Temp04 (
	ItemCode	Nvarchar(10) Collate Korean_Wansung_Unicode_CI_AS,
	Price		Numeric(12,0)
)


Insert Into #Temp03
Select a.U_ItmBsort,
	   b.U_Code,
	   b.U_Name,
	   b.U_Value   
From [@PS_CO502H] a Inner Join [@PS_CO502L] b On a.DocEntry = b.DocEntry
Where a.U_DocDate = (Select Max(x.U_DocDate)
					From [@PS_CO502H] x
				   Where x.U_DocDate <= @DocDate)

--휘팅규격별 단가(평균할인율 53%적용)
Insert Into #Temp04
SELECT ItemCode, ROUND( PRICE *  ( (100-53.0)/100 ) ,0) S_PRICE
  FROM (SELECT	T0.ItemCode,
				Case When T1.U_ItmMsort='10101' then  (CASE WHEN T0.U_TikAngle=90 THEN  T1.U_price1 ELSE T1.U_Price2 END) --엘보우의 경우
					 WHEN T1.U_ItmMsort='10107' then  (CASE WHEN T0.U_TikAngle=10 THEN  T1.U_price1 ELSE T1.U_Price2 END) --후랜지의 경우
					  else T1.U_Price1 end as PRICE
		  FROM OITM T0 INNER JOIN [@PS_MM060L] T1 ON T0.U_ItmBsort=T1.U_ItmBsort AND T0.U_ItmMsort=T1.U_ItmMsort
					--규격이 현재규격보다작은것들중에서 제일큰것
					AND T1.U_Spec1=
					(SELECT 
						MAX(CONVERT(FLOAT,Tx.U_Spec1)) 
					FROM 
						[@PS_MM060L] Tx 
					WHERE 
						Tx.U_ItmBsort= T0.U_ItmBsort AND Tx.U_ItmMsort=T0.U_ItmMsort and Convert(float,U_Spec1) <= Convert(float,T0.U_Spec1)
					)
				) X1
				

----실동공수, 가동율
--Create Table #Temp03 (
--	ItmBsort	Nvarchar(10) Collate Korean_Wansung_Unicode_CI_AS,
--	ItmBname	Nvarchar(20),
--	WTime		Numeric(12,2), --실동시간
--	NTime		Numeric(12,2), --비가동시간
--	GaRate		Numeric(12,2)  --가동율
--)


Insert Into #Temp01

--일일납품
Select x.ItmBsort,
	   x.ItmBname,
	   x.ItemCode,
	   x.ItemName,
	   sum(x.Qty) As Qty,
	   Sum(x.Wgt) As Wgt,
	   Sum(x.Linetotal) As Amt,
	   Case When x.ItmBsort = '101' Then round(Sum((x.Wgt * (wdanga101 + 700 + 425))),0)
			When x.ItmBsort = '102' Then Sum(x.Qty * wdanga102)
			When x.ItmBsort = '104' Then 0
	   End As wamt,
	   Case When x.ItmBsort = '101' Then round(Sum(x.Linetotal - (x.Wgt * (wdanga101 + 700 + 425))),0)
			When x.ItmBsort = '102' Then Sum(x.LIneTotal - (x.Qty * wdanga102))
			When x.ItmBsort = '104' Then Sum(x.LIneTotal)
	   End As gaduk
from (
select t3.U_ItmBsort As ItmBsort,
	   ItmBname = (select Name From [@PSH_ITMBSORT] Where Code = t3.U_ItmBsort),
	   t3.ItemCode,
	   t3.ItemName,
	   Qty = Case When t3.U_ItmBsort = '101' Then t2.Quantity
				  When t3.U_ItmBsort = '102' Then t2.Quantity
				  When t3.U_ItmBsort = '104' Then 0
			 End,
	   Wgt = Case When t3.U_ItmBsort = '101' Then Round((t2.Quantity * t3.U_UnWeight) / 1000,3)
				  When t3.U_ItmBsort = '102' Then 0
				  When t3.U_ItmBsort = '104' Then t2.Quantity
			 End,
	   t2.Linetotal,
	   wdanga102 = Isnull((Select b.U_Price
					  From [@PS_CO501H] a Inner Join [@PS_CO501L] b On a.DocEntry = b.DocEntry
					 Where b.U_ItemCode = t3.ItemCode
					   And a.U_DocDate = (Select max(c.U_DocDate)
										   From [@PS_CO501H] c
										  where c.U_DocDate <= @DocDate)),0), --부품 원재료단가
	   wdanga101 = Isnull((Select b.U_CopPrice
					  From [@PS_MM001H] a Inner Join [@PS_MM001L] b On a.DocEntry = b.DocEntry
					 Where a.U_DocDate = (Select max(c.U_DocDate)
										   From [@PS_MM001H] c Inner Join [@PS_MM001L] d On c.DocEntry = d.DocEntry
										  Where Convert(char(8),c.U_DocDate,112) Like Convert(char(6),dateAdd(mm, -1, @DocDate),112) + '%'
											and d.U_CopPrice > 0)),0) --휘팅전기동가
												
					  
from ODLN t1 Inner Join DLN1 t2 On t1.DocEntry = t2.DocEntry
			 inner Join OITM t3 On t2.ItemCode = t3.ItemCode
where t1.DocDate = @DocDate
  and t1.BPLId = '1'
  And t3.U_ItmBsort In ('101','102', '104')
) x
Group by x.ItmBsort,
	   x.ItmBname,
	   x.ItemCode,
	   x.ItemName

Insert Into #Temp01
--일일반품
Select x.ItmBsort,
	   x.ItmBname,
	   x.ItemCode,
	   x.ItemName,
	   sum(x.Qty) * -1 As Qty,
	   Sum(x.Wgt) * -1 As Wgt,
	   Sum(x.Linetotal) * -1 As Amt,
	   (Case When x.ItmBsort = '101' Then round(Sum((x.Wgt * (wdanga101 + 700 + 425))),0)
			When x.ItmBsort = '102' Then Sum(x.Qty * wdanga102)
			When x.ItmBsort = '104' Then 0
	   End) * -1 As wamt,
	   (Case When x.ItmBsort = '101' Then round(Sum(x.Linetotal - (x.Wgt * (wdanga101 + 700 + 425))),0)
			When x.ItmBsort = '102' Then Sum(x.LIneTotal - (x.Qty * wdanga102))
			When x.ItmBsort = '104' Then Sum(x.LIneTotal)
	   End) * -1 As gaduk
from (
select t3.U_ItmBsort As ItmBsort,
	   ItmBname = (select Name From [@PSH_ITMBSORT] Where Code = t3.U_ItmBsort),
	   t3.ItemCode,
	   t3.ItemName,
	   Qty = Case When t3.U_ItmBsort = '101' Then t2.Quantity
				  When t3.U_ItmBsort = '102' Then t2.Quantity
				  When t3.U_ItmBsort = '104' Then 0
			 End,
	   Wgt = Case When t3.U_ItmBsort = '101' Then Round((t2.Quantity * t3.U_UnWeight) / 1000,3)
				  When t3.U_ItmBsort = '102' Then 0
				  When t3.U_ItmBsort = '104' Then t2.Quantity
			 End,
	   t2.Linetotal,
	   wdanga102 = Isnull((Select b.U_Price
					  From [@PS_CO501H] a Inner Join [@PS_CO501L] b On a.DocEntry = b.DocEntry
					 Where b.U_ItemCode = t3.ItemCode
					   And a.U_DocDate = (Select max(c.U_DocDate)
										   From [@PS_CO501H] c
										  where c.U_DocDate <= @DocDate)),0), --부품 원재료단가
	   wdanga101 = Isnull((Select b.U_CopPrice
					  From [@PS_MM001H] a Inner Join [@PS_MM001L] b On a.DocEntry = b.DocEntry
					 Where a.U_DocDate = (Select max(c.U_DocDate)
										   From [@PS_MM001H] c Inner Join [@PS_MM001L] d On c.DocEntry = d.DocEntry
										  Where Convert(char(8),c.U_DocDate,112) Like Convert(char(6),dateAdd(mm, -1, @DocDate),112) + '%'
											and d.U_CopPrice > 0)),0) --휘팅전기동가
												
					  
from ORDN t1 Inner Join RDN1 t2 On t1.DocEntry = t2.DocEntry
			 inner Join OITM t3 On t2.ItemCode = t3.ItemCode
where t1.DocDate = @DocDate
  and t1.BPLId = '1'
  And t3.U_ItmBsort In ('101','102', '104')
) x
Group by x.ItmBsort,
	   x.ItmBname,
	   x.ItemCode,
	   x.ItemName

--//휘팅
Insert Into #Temp02(ItmBsort, ItmBname, ItemCode, ItemName, MQty, MWgt, MAmt)  --생산
select c.U_ItmBsort,
	   ItmBname = (Select Name From [@PSH_ITMBSORT] Where Code = c.U_ItmBsort),
	   c.ItemCode,
	   c.ItemName,
	   b.U_YQty,
	   Case When c.U_ItmBsort = '104' Then b.U_YQty Else b.U_YWeight End,
	   Amt = Isnull((Select Price From #Temp04 t
				Where t.ItemCode = c.ItemCode),0) * b.U_YQty
from [@PS_PP080H] a Inner Join [@PS_PP080L] b On a.DocEntry = b.DocEntry
					Inner Join OITM c On b.U_Itemcode = c.ItemCode
where a.Canceled = 'N'
  and a.U_BPLId = '1'
  and c.U_ItmBsort = '101'
  And a.U_DocDate = @DocDate
  
--//부품
Insert Into #Temp02(ItmBsort, ItmBname, ItemCode, ItemName, MQty, MWgt, MAmt)  --생산
select c.U_ItmBsort,
	   ItmBname = (Select Name From [@PSH_ITMBSORT] Where Code = c.U_ItmBsort),
	   c.ItemCode,
	   c.ItemName,
	   b.U_YQty,
	   Case When c.U_ItmBsort = '104' Then b.U_YQty Else b.U_YWeight End,
	   Amt = Isnull((Select Price From ORDR t Inner Join RDR1 t1 On t.DocEntry = t1.DocEntry
				Where t1.DocEntry = b.U_ORDRNo and t1.LineNum = b.U_RDR1No
				  And t.Canceled = 'N'),0) * b.U_YQty
from [@PS_PP080H] a Inner Join [@PS_PP080L] b On a.DocEntry = b.DocEntry
					Inner Join OITM c On b.U_Itemcode = c.ItemCode
where a.Canceled = 'N'
  and a.U_BPLId = '1'
  and c.U_ItmBsort = '102'
  And a.U_DocDate = @DocDate

--멀티
Insert Into #Temp02(ItmBsort, ItmBname, ItemCode, ItemName, MQty, MWgt, MAmt)  --생산
select c.U_ItmBsort,
	   ItmBname = (Select Name From [@PSH_ITMBSORT] Where Code = c.U_ItmBsort),
	   c.ItemCode,
	   c.ItemName,
	   b.U_YQty,
	   Case When c.U_ItmBsort = '104' Then b.U_YQty Else b.U_YWeight End,
	   Amt = Isnull((Select Price From [ITM1] t
				Where t.ItemCode = b.U_ItemCode and t.PriceList = 1),0) * b.U_YQty
from [@PS_PP080H] a Inner Join [@PS_PP080L] b On a.DocEntry = b.DocEntry
					Inner Join OITM c On b.U_Itemcode = c.ItemCode
where a.Canceled = 'N'
  and a.U_BPLId = '1'
  and c.U_ItmBsort = '104'
  And a.U_DocDate = @DocDate


  
Insert Into #Temp02(ItmBsort, ItmBname, ItemCode, ItemName, MQty, MWgt, MAmt)  --벌크생산
Select c.U_ItmBsort,
	   ItmBname = (Select Name From [@PSH_ITMBSORT] Where Code = c.U_ItmBsort),
	   c.ItemCode,
	   c.ItemName,
	   b.U_SelQTy,
	   b.U_SelWt,
	   Amt = Isnull((Select Price From #Temp04 t
				Where t.ItemCode = c.ItemCode),0) * b.U_SelQTy
From [@PS_PP070H] a Inner Join [@PS_PP070L] b On a.DocEntry = b.DocEntry
					Inner Join [OITM] c On b.U_ItemCode = c.ItemCode
Where a.Canceled = 'N'
  and a.U_DocDate = @DocDate

--노무공수

--Insert Into #Temp03 (ItmBsort, ItmBname, Wtime, Ntime )
--Select x.ItmBsort,
--	   x.ItmBname,
--	   Wtime = Sum(x.WorkTime),
--	   Ntime = Sum(x.Ntime)
--From ( 
--select a.U_OrdGbn As ItmBsort,
--	   ItmBname = (Select Name from [@PSH_ITMBSORT] Where Code = a.U_OrdGbn),
--	   a.DocEntry,
--	   Worktime = (select Sum(Case When Isnull(Convert(Numeric(6,2), t1.U_YTime),0) = 0 Then Convert(Numeric(6,2),t.U_BaseTime)
--					   Else Convert(Numeric(6,2), t1.U_YTime) End )
--				  From [@PS_PP040H] t Inner Join [@PS_PP040M] t1 On t.DocEntry = t1.DocEntry
--				 Where t.DocEntry = a.DocEntry ), --개인 근로시간
--	   NTime = (Select Sum(Case When Isnull(Convert(Numeric(6,2), t.U_Ntime),0) = 0 Then 0 Else Isnull(convert(Numeric(8,2),t.U_Ntime),0) End)
--				  From [@PS_PP040M] t
--				 Where a.DocEntry = t.DocEntry )
	   
--from [@PS_PP040H] a 
--where a.Canceled = 'N'
--  and a.U_DocDate = '20110412'--@DocDate
--  and a.U_OrdGbn in ('101','102','104')
--  and a.U_OrdType = '10'
--  ) x
--  Group by x.ItmBsort,
--	   x.ItmBname
	   
--//품목코드별 실동시간 계산 -- 품목그룹별 이기때문에 사용안함 -----------------------------------------------------
--Select y.ItmBsort,
--	   y.ItmBname,
--	   Sum(y.Wtime),
--	   Sum(z.Ntime)
--From (
--Select x.DocEntry,
--	   x.ItmBsort,
--	   x.ItmBname,
--	   Wtime = Round(Sum((WorkTime / SWtime) * PTime),2) --노무공수
--From ( 
--select c.U_ItmBsort As ItmBsort,
--	   ItmBname = (Select Name from [@PSH_ITMBSORT] Where Code = c.U_ItmBsort),
--	   a.DocEntry,
--	   b.U_ItemCode As ItemCode,
--	   c.ItemName, 
--	   b.U_OrdNum As OrdNum, 
--	   Isnull(Convert(Numeric(8,2),a.U_BaseTime),0) As BaseTime, 
--	   Isnull(Convert(Numeric(8,2),b.U_WorkTime),0) As WorkTime,
--	   SWtime = (select sum(convert(Numeric(8,2),t.U_WorkTime))
--				   From [@PS_PP040L] t
--				  Where a.DocEntry = t.DocEntry), --실동시간 합계
--	   Ptime = (select Sum(Case When Isnull(Convert(Numeric(6,2), t1.U_YTime),0) = 0 Then Convert(Numeric(6,2),t.U_BaseTime)
--					   Else Convert(Numeric(6,2), t1.U_YTime) End )
--				  From [@PS_PP040H] t Inner Join [@PS_PP040M] t1 On t.DocEntry = t1.DocEntry
--				 Where t.DocEntry = a.DocEntry ) --개인 근로시간
	   
--from [@PS_PP040H] a Inner Join [@PS_PP040L] b On a.DocEntry = b.DocEntry
--					Inner Join OITM c On b.U_ItemCode = c.ItemCode
--where a.Canceled = 'N'
--  and a.U_DocDate = '20110412'--@DocDate
--  and c.U_ItmBsort in ('101','102','104')
--  and a.U_OrdType = '10'
--  ) x
--  Group by x.DocEntry,
--	   x.ItmBsort,
--	   x.ItmBname
--  )y,
--  (Select t.DocEntry, Ntime = Sum(Case When Isnull(Convert(Numeric(6,2), t.U_Ntime),0) = 0 Then 0 Else Isnull(convert(Numeric(8,2),t.U_Ntime),0) End)
--	 From [@PS_PP040M] t
--	Group by t.DocEntry ) z
--Where y.DocEntry = z.DocEntry
--  Group by y.ItmBsort,
--	   y.ItmBname
--------------------------------------------------------------------------------------------------------------------------------------------
--Insert Into #Temp03 (ItmBsort, ItmBname, Wtime, Ntime )
----비가동시간
--Select t.ItmBsort,
--	   ItmBname = (Select Name From [@PSH_ITMBSORT] Where Code = t.ItmBsort),
--	   t.WTime,
--	   t.NTime
--From (
--Select a.U_WorkGbn As ItmBsort,
--	   WTime = Sum(Charindex(a.U_OrdType, '10') * Isnull(Convert(Numeric(8,2),a.U_WorkTime),0)), 
--	   NTime = Sum(Charindex(a.U_OrdType, '20') * Isnull(Convert(Numeric(8,2),a.U_WorkTime),0))
--From [@PS_PP060H] a
--Where a.U_DocDate = @DocDate
--  and a.U_WorkGbn In ('101','102','104')
-- Group by a.U_WorkGbn
--) t


Select x.ItmBsort,
	   x.ItmBname,
	   x.ItmMsort,
	   x.ItmMname,
	   x.ItemCode,
	   x.ItemName,
	   SQty = Sum(x.Qty),
	   SWgt = Sum(x.Wgt),
	   SAmt = Sum(x.Amt),
	   WAmt = Sum(x.Wamt),
	   Gaduk = Sum(x.Gaduk),
	   MQty = Sum(x.MQty),
	   MWgt = Sum(x.MWgt),
	   MAmt = Sum(x.MAmt),
	   BuAmt = Round(Sum(x.BuAmt),0),		--보조제료비
	   PayAmt1 = Round(Sum(x.PayAmt1),0),		--직접노무비
	   PayAmt2 = Round(Sum(x.PayAmt2),0),		--간접노무비
	   UniAmt1 = Round(Sum(x.UniAmt1),0),		--직접제조경비
	   UniAmt2 = Round(Sum(x.UniAmt2),0),		--간접제조경비
	   Wonga = Round(Sum(x.BuAmt),0) + Round(Sum(x.PayAmt1),0) + Round(Sum(x.PayAmt2),0) + Round(Sum(x.UniAmt1),0) + Round(Sum(x.UniAmt2),0),
	   x.div
From (
Select t.ItmBsort,
	   t.ItmBname,
	   ItmMsort = t1.U_ItmMsort,
	   ItmMname = (Select U_CodeName From [@PSH_ITMMSORT] Where U_Code = t1.U_ItmMsort),
	   t.ItemCode,
	   t.ItemName,
	   t.Qty,
	   t.Wgt,
	   t.Amt,
	   t.Wamt,
	   t.Gaduk,
	   MQty = 0,
	   MWgt = 0,
	   MAmt = 0,
	   BuAmt = t.Wgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '02'),		--보조제료비(생산량 * 평균단가)
	   PayAmt1 = (t.Wgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '06')),		--직접노무비
	   PayAmt2 = (t.Wgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '06')) * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '03') / 100,		--간접노무비
	   UniAmt1 = (t.Wgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '04')),		--직접제조경비
	   UniAmt2 = (t.Wgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '04')) * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '05') / 100,		--간접제조경비
	   div = '1'
from #Temp01 t Inner Join OITM t1 On t.ItemCode = t1.ItemCode
Where t.ItmBsort in ('101','104') --휘팅, 멀티

Union all

Select t.ItmBsort,
	   t.ItmBname,
	   ItmMsort = t1.U_ItmMsort,
	   ItmMname = (Select U_CodeName From [@PSH_ITMMSORT] Where U_Code = t1.U_ItmMsort),
	   t.ItemCode,
	   t.ItemName,
	   t.Qty,
	   t.Wgt,
	   t.Amt,
	   t.Wamt,
	   t.Gaduk,
	   MQty = 0,
	   MWgt = 0,
	   MAmt = 0,
	   BuAmt = t.Amt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '02') / 100,		--보조제료비(생산량 * 평균단가)
	   PayAmt1 = (t.Amt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '06')) / 100,		--직접노무비
	   PayAmt2 = ((t.Amt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '06')) / 100) * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '03') / 100,		--간접노무비
	   UniAmt1 = (t.Amt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '04')) / 100,		--직접제조경비
	   UniAmt2 = ((t.Amt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '04')) / 100) * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '05') / 100,		--간접제조경비
	   div = '1'
from #Temp01 t Inner Join OITM t1 On t.ItemCode = t1.ItemCode
Where t.ItmBsort in ('102') --부품
) x
Group by x.ItmBsort,
	   x.ItmBname,
	   x.ItmMsort,
	   x.ItmMname,
	   x.ItemCode,
	   x.ItemName,
	   x.div


Union all


Select x.ItmBsort,
	   x.ItmBname,
	   x.ItmMsort,
	   x.ItmMname,
	   x.ItemCode,
	   x.ItemName,
	   SQty = Sum(x.Qty),
	   SWgt = Sum(x.Wgt),
	   SAmt = Sum(x.Amt),
	   WAmt = Sum(x.Wamt),
	   Gaduk = Sum(x.Gaduk),
	   MQty = Sum(x.MQty),
	   MWgt = Sum(x.MWgt),
	   MAmt = Sum(x.MAmt),
	   BuAmt = Round(Sum(x.BuAmt),0),		--보조제료비
	   PayAmt1 = Round(Sum(x.PayAmt1),0),		--직접노무비
	   PayAmt2 = Round(Sum(x.PayAmt2),0),		--간접노무비
	   UniAmt1 = Round(Sum(x.UniAmt1),0),		--직접제조경비
	   UniAmt2 = Round(Sum(x.UniAmt2),0),		--간접제조경비
	   Wonga = Round(Sum(x.BuAmt),0) + Round(Sum(x.PayAmt1),0) + Round(Sum(x.PayAmt2),0) + Round(Sum(x.UniAmt1),0) + Round(Sum(x.UniAmt2),0),
	   x.div
From (

Select t.ItmBsort,
	   t.ItmBname,
	   ItmMsort = t1.U_ItmMsort,
	   ItmMname = (Select U_CodeName From [@PSH_ITMMSORT] Where U_Code = t1.U_ItmMsort),
	   t.ItemCode,
	   t.ItemName,
	   Qty = 0,
	   Wgt = 0,
	   Amt = 0,
	   Wamt = 0,
	   Gaduk = 0,
	   MQty = t.MQty,
	   MWgt = t.MWgt,
	   MAmt = t.MAmt,
	   BuAmt = t.MWgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '02'),		--보조제료비(생산량 * 평균단가)
	   PayAmt1 = (t.MWgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '06')),		--직접노무비
	   PayAmt2 = (t.MWgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '06')) * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '03') / 100,		--간접노무비
	   UniAmt1 = (t.MWgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '04')),		--직접제조경비
	   UniAmt2 = (t.MWgt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '04')) * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '05') / 100,		--간접제조경비
	   div = '2'
from #Temp02 t Inner Join OITM t1 On t.ItemCode = t1.ItemCode
Where t.ItmBsort in ('101','104') --휘팅, 멀티

Union All

Select t.ItmBsort,
	   t.ItmBname,
	   ItmMsort = t1.U_ItmMsort,
	   ItmMname = (Select U_CodeName From [@PSH_ITMMSORT] Where U_Code = t1.U_ItmMsort),
	   t.ItemCode,
	   t.ItemName,
	   Qty = 0,
	   Wgt = 0,
	   Amt = 0,
	   Wamt = 0,
	   Gaduk = 0,
	   MQty = t.MQty,
	   MWgt = t.MWgt,
	   MAmt = t.MAmt,
	   BuAmt = t.MAmt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '02')/ 100,		--보조제료비(생산량 * 평균단가)
	   PayAmt1 = (t.MAmt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '06')) / 100,		--직접노무비
	   PayAmt2 = (t.MAmt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '06')) / 100 * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '03') / 100,		--간접노무비
	   UniAmt1 = (t.MAmt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '04')) / 100,		--직접제조경비
	   UniAmt2 = (t.MAmt * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '04')) / 100 * (select Value From #Temp03 Where ItmBsort = t.ItmBsort And Code = '05') / 100,		--간접제조경비
	   div = '2'
from #Temp02 t Inner Join OITM t1 On t.ItemCode = t1.ItemCode
Where t.ItmBsort in ('102') --부품
) x

Group by x.ItmBsort,
	   x.ItmBname,
	   x.ItmMsort,
	   x.ItmMname,
	   x.ItemCode,
	   x.ItemName,
	   x.div

Order by x.div, x.ItmBsort, x.ItmMsort, x.ItemCode



End



--  EXEC [dbo].[PS_CO503_01] '20110412'