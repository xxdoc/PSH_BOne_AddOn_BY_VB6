USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_CO504_01]    Script Date: 05/06/2011 14:03:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/******************************************************************************************************************/
/*  Module         : PP								    														*/
/*  Description    : 일일 생산 및 판매원가				  															*/
/*  Create Date    : 2011.05.03                                                                                  */
/*  Modified Date  :										       													*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/******************************************************************************************************************/
--Create PROC [dbo].[PS_CO504_01]
ALTER PROC [dbo].[PS_CO504_01]
(
	@DocDate	Date,
	@Gubun		Char(1) --//출력구분
   
)
AS

BEGIN
--판매
Create Table #Temp01 (
	ItmBsort	Nvarchar(10) Collate Korean_Wansung_Unicode_CI_AS,
	ItmBname	Nvarchar(20),
	ItemCode	Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	ItemName	Nvarchar(50),
	MQty		Numeric(19,2), --월판매 누계수량 
	MWgt		Numeric(19,2), --월판매 누계중량 
	MAmt		Numeric(19,0), --월판매 누계금액 
	Qty			Numeric(19,2), --일판매수량
	Wgt			Numeric(19,2), --일판매중량
	Amt			Numeric(19,0)  --일판매금액
)

--생산
Create Table #Temp02 (
	ItmBsort	Nvarchar(10) Collate Korean_Wansung_Unicode_CI_AS,
	ItmBname	Nvarchar(20),
	ItemCode	Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	ItemName	Nvarchar(50),
	MQty		Numeric(12,0), --월생산누계수량
	MWgt		Numeric(12,2),  --월생산누계중량
	MAmt		Numeric(12,0), --월생산 누계금액
	Qty			Numeric(19,2), --일생산수량
	Wgt			Numeric(19,2), --일생산중량
	Amt			Numeric(19,0)  --일생산금액
)

Create Table #Temp04 (
	ItemCode	Nvarchar(10) Collate Korean_Wansung_Unicode_CI_AS,
	Price		Numeric(12,0)
)


Create Table #Temp_co504 (
	ItmBsort	Nvarchar(10),
	ItmBname	Nvarchar(20),
	ItmMsort	Nvarchar(10),
	ItmMname	Nvarchar(20),
	ItemCode	Nvarchar(20),
	ItemName	Nvarchar(50),
	SQty		Numeric(12,0), --판매수량
	SWgt		Numeric(12,2),  --판매중량
	SAmt		Numeric(12,0), --판매금액
	SmQty		Numeric(12,0), --월판매누계수량
	SmWgt		Numeric(12,2),  --월판매누계중량
	SmAmt		Numeric(12,0), --월판매누계금액
	MQty		Numeric(12,0), --생산수량
	MWgt		Numeric(12,2),  --생산중량
	MAmt		Numeric(12,0), --생산금액
	MmQty		Numeric(12,0), --월생산누계수량
	MmWgt		Numeric(12,2),  --월생산누계중량
	MmAmt		Numeric(12,0), --월생산 누계금액
	div			Char(1)
)
	

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
				

Insert Into #Temp01

--일일납품
Select x.ItmBsort,
	   x.ItmBname,
	   x.ItemCode,
	   x.ItemName,
	   sum(x.MQty) As MQty,
	   Sum(x.MWgt) As MWgt,
	   Sum(x.MAmt) As MAmt,
	   sum(x.Qty) As Qty,
	   Sum(x.Wgt) As Wgt,
	   Sum(x.Amt) As Amt
from (
select t3.U_ItmBsort As ItmBsort,
	   ItmBname = (select Name From [@PSH_ITMBSORT] Where Code = t3.U_ItmBsort),
	   t3.ItemCode,
	   t3.ItemName,
	   MQty = Case When t3.U_ItmBsort = '101' Then t2.Quantity
				  When t3.U_ItmBsort = '102' Then t2.Quantity
				  When t3.U_ItmBsort = '104' Then 0
			 End,
	   MWgt = Case When t3.U_ItmBsort = '101' Then Round((t2.Quantity * t3.U_UnWeight) / 1000,3)
				  When t3.U_ItmBsort = '102' Then 0
				  When t3.U_ItmBsort = '104' Then t2.Quantity
			 End,
	   MAmt = t2.Linetotal,
	   Qty = Case When t1.DocDate = @DocDate Then
				  Case When t3.U_ItmBsort = '101' Then t2.Quantity
					   When t3.U_ItmBsort = '102' Then t2.Quantity
					   When t3.U_ItmBsort = '104' Then 0
				  End
			 Else 0
			 End,
	   Wgt = Case When t1.DocDate = @DocDate Then
				  Case When t3.U_ItmBsort = '101' Then Round((t2.Quantity * t3.U_UnWeight) / 1000,3)
					   When t3.U_ItmBsort = '102' Then 0
					   When t3.U_ItmBsort = '104' Then t2.Quantity
				  End
			 Else 0
			 End,
	   Amt = Case When t1.DocDate = @DocDate Then
				  t2.Linetotal
			 Else 0
			 End
from ODLN t1 Inner Join DLN1 t2 On t1.DocEntry = t2.DocEntry
			 inner Join OITM t3 On t2.ItemCode = t3.ItemCode
where t1.DocDate Between Convert(Char(6),@DocDate,112) + '01' and @DocDate
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
	   sum(x.MQty) * -1 As MQty,
	   Sum(x.MWgt) * -1 As MWgt,
	   Sum(x.MAmt) * -1 As MAmt,
	   sum(x.Qty) * -1 As Qty,
	   Sum(x.Wgt) * -1 As Wgt,
	   Sum(x.Amt) * -1 As Amt
from (
select t3.U_ItmBsort As ItmBsort,
	   ItmBname = (select Name From [@PSH_ITMBSORT] Where Code = t3.U_ItmBsort),
	   t3.ItemCode,
	   t3.ItemName,
	   MQty = Case When t3.U_ItmBsort = '101' Then t2.Quantity
				  When t3.U_ItmBsort = '102' Then t2.Quantity
				  When t3.U_ItmBsort = '104' Then 0
			 End,
	   MWgt = Case When t3.U_ItmBsort = '101' Then Round((t2.Quantity * t3.U_UnWeight) / 1000,3)
				  When t3.U_ItmBsort = '102' Then 0
				  When t3.U_ItmBsort = '104' Then t2.Quantity
			 End,
	   MAmt = t2.Linetotal,
	   Qty = Case When t1.DocDate = @DocDate Then
				  Case When t3.U_ItmBsort = '101' Then t2.Quantity
					   When t3.U_ItmBsort = '102' Then t2.Quantity
					   When t3.U_ItmBsort = '104' Then 0
				  End
			 Else 0
			 End,
	   Wgt = Case When t1.DocDate = @DocDate Then
				  Case When t3.U_ItmBsort = '101' Then Round((t2.Quantity * t3.U_UnWeight) / 1000,3)
					   When t3.U_ItmBsort = '102' Then 0
					   When t3.U_ItmBsort = '104' Then t2.Quantity
				  End
			 Else 0
			 End,
	   Amt = Case When t1.DocDate = @DocDate Then
				  t2.Linetotal
			 Else 0
			 End				
					  
from ORDN t1 Inner Join RDN1 t2 On t1.DocEntry = t2.DocEntry
			 inner Join OITM t3 On t2.ItemCode = t3.ItemCode
where t1.DocDate Between Convert(Char(6),@DocDate,112) + '01' and @DocDate
  and t1.BPLId = '1'
  And t3.U_ItmBsort In ('101','102', '104')
) x
Group by x.ItmBsort,
	   x.ItmBname,
	   x.ItemCode,
	   x.ItemName


--//휘팅생산
Insert Into #Temp02(ItmBsort, ItmBname, ItemCode, ItemName, MQty, MWgt, MAmt, Qty, Wgt, Amt)  --생산
Select t.ItmBsort, t.ItmBname, t.ItemCode, t.ItemName,
	   MQty = t.Qty,
	   MWgt = t.Wgt,
	   MAmt = t.Amt,
	   Qty  = Case When t.Docdate = @DocDate Then t.Qty Else 0 End,
	   Wgt  = Case When t.Docdate = @DocDate Then t.Wgt Else 0 End,
	   Amt  = Case When t.Docdate = @DocDate Then t.Amt Else 0 End
From (
select a.U_DocDate As DocDate,
	   c.U_ItmBsort As ItmBsort,
	   ItmBname = (Select Name From [@PSH_ITMBSORT] Where Code = c.U_ItmBsort),
	   c.ItemCode,
	   c.ItemName,
	   Qty = b.U_YQty,
	   Wgt = Case When c.U_ItmBsort = '104' Then b.U_YQty Else b.U_YWeight End,
	   Amt = Isnull((Select Price From #Temp04 t
				Where t.ItemCode = c.ItemCode),0) * b.U_YQty
from [@PS_PP080H] a Inner Join [@PS_PP080L] b On a.DocEntry = b.DocEntry
					Inner Join OITM c On b.U_Itemcode = c.ItemCode
where a.Canceled = 'N'
  and a.U_BPLId = '1'
  and c.U_ItmBsort = '101'
  And a.U_DocDate Between Convert(Char(6),@DocDate,112) + '01' and @DocDate
 ) t 
 
--//부품
Insert Into #Temp02(ItmBsort, ItmBname, ItemCode, ItemName, MQty, MWgt, MAmt, Qty, Wgt, Amt)  --생산
Select t.ItmBsort, t.ItmBname, t.ItemCode, t.ItemName,
	   MQty = t.Qty,
	   MWgt = t.Wgt,
	   MAmt = t.Amt,
	   Qty  = Case When t.Docdate = @DocDate Then t.Qty Else 0 End,
	   Wgt  = Case When t.Docdate = @DocDate Then t.Wgt Else 0 End,
	   Amt  = Case When t.Docdate = @DocDate Then t.Amt Else 0 End
From (
select a.U_DocDate As DocDate,
	   c.U_ItmBsort As ItmBsort,
	   ItmBname = (Select Name From [@PSH_ITMBSORT] Where Code = c.U_ItmBsort),
	   c.ItemCode,
	   c.ItemName,
	   Qty = b.U_YQty,
	   Wgt = Case When c.U_ItmBsort = '104' Then b.U_YQty Else b.U_YWeight End,
	   Amt = Isnull((Select Price From ORDR t Inner Join RDR1 t1 On t.DocEntry = t1.DocEntry
				Where t1.DocEntry = b.U_ORDRNo and t1.LineNum = b.U_RDR1No
				  And t.Canceled = 'N'),0) * b.U_YQty
from [@PS_PP080H] a Inner Join [@PS_PP080L] b On a.DocEntry = b.DocEntry
					Inner Join OITM c On b.U_Itemcode = c.ItemCode
where a.Canceled = 'N'
  and a.U_BPLId = '1'
  and c.U_ItmBsort = '102'
  And a.U_DocDate Between Convert(Char(6),@DocDate,112) + '01' and @DocDate
) t

--멀티
Insert Into #Temp02(ItmBsort, ItmBname, ItemCode, ItemName, MQty, MWgt, MAmt, Qty, Wgt, Amt)  --생산
Select t.ItmBsort, t.ItmBname, t.ItemCode, t.ItemName,
	   MQty = t.Qty,
	   MWgt = t.Wgt,
	   MAmt = t.Amt,
	   Qty  = Case When t.Docdate = @DocDate Then t.Qty Else 0 End,
	   Wgt  = Case When t.Docdate = @DocDate Then t.Wgt Else 0 End,
	   Amt  = Case When t.Docdate = @DocDate Then t.Amt Else 0 End
From (
select a.U_DocDate As DocDate,
	   c.U_ItmBsort As ItmBsort,
	   ItmBname = (Select Name From [@PSH_ITMBSORT] Where Code = c.U_ItmBsort),
	   c.ItemCode,
	   c.ItemName,
	   Qty = b.U_YQty,
	   Wgt = Case When c.U_ItmBsort = '104' Then b.U_YQty Else b.U_YWeight End,
	   Amt = Isnull((Select Price From [ITM1] t
				Where t.ItemCode = b.U_ItemCode and t.PriceList = 1),0) * b.U_YQty
from [@PS_PP080H] a Inner Join [@PS_PP080L] b On a.DocEntry = b.DocEntry
					Inner Join OITM c On b.U_Itemcode = c.ItemCode
where a.Canceled = 'N'
  and a.U_BPLId = '1'
  and c.U_ItmBsort = '104'
  And a.U_DocDate Between Convert(Char(6),@DocDate,112) + '01' and @DocDate
) t
  
Insert Into #Temp02(ItmBsort, ItmBname, ItemCode, ItemName, MQty, MWgt, MAmt, Qty, Wgt, Amt)  --벌크생산
Select t.ItmBsort, t.ItmBname, t.ItemCode, t.ItemName,
	   MQty = t.Qty,
	   MWgt = t.Wgt,
	   MAmt = t.Amt,
	   Qty  = Case When t.Docdate = @DocDate Then t.Qty Else 0 End,
	   Wgt  = Case When t.Docdate = @DocDate Then t.Wgt Else 0 End,
	   Amt  = Case When t.Docdate = @DocDate Then t.Amt Else 0 End
From (
Select a.U_DocDate As DocDate,
	   c.U_ItmBsort As ItmBsort,
	   ItmBname = (Select Name From [@PSH_ITMBSORT] Where Code = c.U_ItmBsort),
	   c.ItemCode,
	   c.ItemName,
	   Qty = b.U_SelQTy,
	   Wgt = b.U_SelWt,
	   Amt = Isnull((Select Price From #Temp04 t
				Where t.ItemCode = c.ItemCode),0) * b.U_SelQTy
From [@PS_PP070H] a Inner Join [@PS_PP070L] b On a.DocEntry = b.DocEntry
					Inner Join [OITM] c On b.U_ItemCode = c.ItemCode
Where a.Canceled = 'N'
  and a.U_DocDate Between Convert(Char(6),@DocDate,112) + '01' and @DocDate
) t


Insert Into #temp_co504 ( ItmBsort, ItmBname, ItmMsort, ItmMname, ItemCode, ItemName, SQty, SWgt, SAmt, SmQty, SmWgt, SmAmt, MQty, MWgt, MAmt, MmQty, MmWgt, MmAmt, div)
	Select x.ItmBsort,
		   x.ItmBname,
		   x.ItmMsort,
		   x.ItmMname,
		   x.ItemCode,
		   x.ItemName,
		   SQty = Sum(x.SQty),
		   SWgt = Sum(x.SWgt),
		   SAmt = Sum(x.SAmt),
		   SmQty = Sum(x.SmQty),
		   SmWgt = Sum(x.SmWgt),
		   SmAmt = Sum(x.SmAmt),
		   MQty = Sum(x.MQty),
		   MWgt = Sum(x.MWgt),
		   MAmt = Sum(x.MAmt),
		   MmQty = Sum(x.MmQty),
		   MmWgt = Sum(x.MmWgt),
		   MmAmt = Sum(x.MmAmt),
		   x.div
	From (
	Select t.ItmBsort,
		   t.ItmBname,
		   ItmMsort = t1.U_ItmMsort,
		   ItmMname = (Select U_CodeName From [@PSH_ITMMSORT] Where U_Code = t1.U_ItmMsort),
		   t.ItemCode,
		   t.ItemName,
		   SmQty = t.MQty,
		   SmWgt = t.MWgt,
		   SmAmt = t.MAmt,
		   SQty  = t.Qty,
		   SWgt  = t.Wgt,
		   SAmt  = t.Amt,
		   MmQty = 0,
		   MmWgt = 0,
		   MmAmt = 0,
		   MQty = 0,
		   MWgt = 0,
		   MAmt = 0,
		   div = '1'
	from #Temp01 t Inner Join OITM t1 On t.ItemCode = t1.ItemCode
	--Where t.ItmBsort in ('101','104') --휘팅, 멀티

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
		   SQty = Sum(x.SQty),
		   SWgt = Sum(x.SWgt),
		   SAmt = Sum(x.SAmt),
		   SmQty = Sum(x.SmQty),
		   SmWgt = Sum(x.SmWgt),
		   SmAmt = Sum(x.SmAmt),
		   MQty = Sum(x.MQty),
		   MWgt = Sum(x.MWgt),
		   MAmt = Sum(x.MAmt),
		   MmQty = Sum(x.MmQty),
		   MmWgt = Sum(x.MmWgt),
		   MmAmt = Sum(x.MmAmt),
		   x.div
	From (

	Select t.ItmBsort,
		   t.ItmBname,
		   ItmMsort = t1.U_ItmMsort,
		   ItmMname = (Select U_CodeName From [@PSH_ITMMSORT] Where U_Code = t1.U_ItmMsort),
		   t.ItemCode,
		   t.ItemName,
		   SmQty = 0,
		   SmWgt = 0,
		   SmAmt = 0,
		   SQty = 0,
		   SWgt = 0,
		   SAmt = 0,
		   MmQty = t.MQty,
		   MmWgt = t.MWgt,
		   MmAmt = t.MAmt,
		   MQty = t.Qty,
		   MWgt = t.Wgt,
		   MAmt = t.Amt,
		   div = '2'
	from #Temp02 t Inner Join OITM t1 On t.ItemCode = t1.ItemCode
	--Where t.ItmBsort in ('101','104') --휘팅, 멀티

	) x

	Group by x.ItmBsort,
		   x.ItmBname,
		   x.ItmMsort,
		   x.ItmMname,
		   x.ItemCode,
		   x.ItemName,
		   x.div

	Order by x.div, x.ItmBsort, x.ItmMsort, x.ItemCode


If @Gubun = '1' --개별
Begin
Select t.ItmBsort,
	   t.ItmBname,
	   t.ItmMsort,
	   t.ItmMname,
	   t.ItemCode,
	   t.ItemName,
	   t.SmQty,
	   t.SmWgt,
	   t.SmAmt,
	   t.SQty,
	   t.SWgt,
	   t.SAmt,
	   t.MmQty,
	   t.MmWgt,
	   t.MmAmt,
	   t.MQty,
	   t.MWgt,
	   t.MAmt,
	   t.div
 From #temp_co504 t
End

Else
Begin

Select x.ItmBsort,
	   x.ItmBname,
	   x.ItmMsort,
	   x.ItmMname,
	   x.ItemCode,
	   x.ItemName,
	   x.SmQty,
	   x.SmWgt,
	   x.SmAmt,
	   x.SQty,
	   x.SWgt,
	   x.SAmt,
	   x.MmQty,
	   x.MmWgt,
	   x.MmAmt,
	   x.MQty,
	   x.MWgt,
	   x.MAmt,
	   x.div
From (

Select t.ItmBsort,
	   t.ItmBname,
	   t.ItmMsort,
	   t.ItmMname,
	   ItemCode = '',
	   ItemName = '',
	   SmQty = Sum(t.SmQty),
	   SmWgt = Sum(t.SmWgt),
	   SmAmt = Sum(t.SmAmt),
	   SQty = Sum(t.SQty),
	   SWgt = Sum(t.SWgt),
	   SAmt = Sum(t.SAmt),
	   MmQty = Sum(t.MmQty),
	   MmWgt = Sum(t.MmWgt),
	   MmAmt = Sum(t.MmAmt),
	   MQty = Sum(t.MQty),
	   MWgt = Sum(t.MWgt),
	   MAmt = Sum(t.MAmt),
	   div = '1'
 From #temp_co504 t
--Where t.ItmBsort = '101'
 Group by t.ItmBsort,
	   t.ItmBname,
	   t.ItmMsort,
	   t.ItmMname

Union all

Select t.ItmBsort,
	   t.ItmBname,
	   ItmMsort = '',
	   ItmMname = '',
	   ItemCode = '',
	   ItemName = '소계',
	   SmQty = Sum(t.SmQty),
	   SmWgt = Sum(t.SmWgt),
	   SmAmt = Sum(t.SmAmt),
	   SQty = Sum(t.SQty),
	   SWgt = Sum(t.SWgt),
	   SAmt = Sum(t.SAmt),
	   MmQty = Sum(t.MmQty),
	   MmWgt = Sum(t.MmWgt),
	   MmAmt = Sum(t.MmAmt),
	   MQty = Sum(t.MQty),
	   MWgt = Sum(t.MWgt),
	   MAmt = Sum(t.MAmt),
	   div = '2'
 From #temp_co504 t
--Where t.ItmBsort = '101'
 Group by t.ItmBsort,
	   t.ItmBname
	   
Union all

Select ItmBsort = '99',
	   ItmBname = '',
	   ItmMsort = '',
	   ItmMname = '',
	   ItemCode = '',
	   ItemName = '총계',
	   SmQty = Sum(t.SmQty),
	   SmWgt = Sum(t.SmWgt),
	   SmAmt = Sum(t.SmAmt),
	   SQty = Sum(t.SQty),
	   SWgt = Sum(t.SWgt),
	   SAmt = Sum(t.SAmt),
	   MmQty = Sum(t.MmQty),
	   MmWgt = Sum(t.MmWgt),
	   MmAmt = Sum(t.MmAmt),
	   MQty = Sum(t.MQty),
	   MWgt = Sum(t.MWgt),
	   MAmt = Sum(t.MAmt),
	   div = '3'
 From #temp_co504 t
) x
Order by x.ItmBsort,
	     x.div,
	   x.ItmMsort
	   
End

End

--  EXEC [dbo].[PS_CO504_01] '20110412', '1'