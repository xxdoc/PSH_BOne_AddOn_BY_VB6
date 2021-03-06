USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP080_05]    Script Date: 04/03/2011 10:25:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROC [dbo].[PS_PP080_05]
--Create PROC [dbo].[PS_PP080_05]
(
	@ItemCode NVARCHAR(20),
	@PP030HNo NVARCHAR(20)
)
AS
BEGIN

Select Cost01 = Sum(t2.Cost01),
	   Cost02 = Sum(t2.Cost02),
	   Cost03 = Sum(t2.Cost03),
	   Cost04 = Sum(t2.Cost04),
	   Cost05 = Sum(t2.Cost05),
	   Cost06 = Sum(t2.Cost06),
	   CostTot = Sum(t2.CostTot),
	   aCost01 = Sum(t2.aCost01),
	   aCost02 = Sum(t2.aCost02),
	   aCost03 = Sum(t2.aCost03),
	   aCost04 = Sum(t2.aCost04),
	   aCost05 = Sum(t2.aCost05),
	   aCost06 = Sum(t2.aCost06),
	   aCostTot = Sum(t2.aCostTot),
	   tSaleQty = Sum(t2.tSaleQty),
	   tSaleAmt = Sum(t2.tSaleAmt)
from (

	select Cost01 = Sum(t1.Cost01),
	   Cost02 = Sum(t1.Cost02),
	   Cost03 = Sum(t1.Cost03),
	   Cost04 = Sum(t1.Cost04),
	   Cost05 = Sum(t1.Cost05),
	   Cost06 = Sum(t1.Cost06),
	   CostTot = Sum(t1.Cost01 + t1.Cost02 + t1.Cost03 + t1.Cost04 + t1.Cost05 + t1.Cost06),
	   aCost01 = 0,
	   aCost02 = 0,
	   aCost03 = 0,
	   aCost04 = 0,
	   aCost05 = 0,
	   aCost06 = 0,
	   aCostTot = 0,
	   tSaleQTy = 0,
	   tSaleAmt = 0
	from (
	SELECT Sum(Charindex(c.U_Purchase,'10') * LineTotal) As Cost01, -- 원자재비
		   Sum(Charindex(c.U_Purchase,'40') * LineTotal) As Cost02, -- 외주제작비
		   Sum(Charindex(c.U_Purchase,'30') * LineTotal) As Cost03, -- 외주가공비
		   Cost04 = 0,
		   Cost05 = 0,
		   Cost06 = 0
	 FROM OPDN a Inner Join PDN1 b On a.DocEntry = b.DocEntry
				 Inner Join [@PS_MM070H] c On c.U_GRDocNum = a.DocEntry
	   and b.U_OrdNum = @ItemCode

	Union all

	select Sum(Charindex(e.U_Purchase,'10') * b.LineTotal * -1) As Cost01, -- 원자재비
		   Sum(Charindex(e.U_Purchase,'40') * b.LineTotal * -1) As Cost02, -- 외주제작비
		   Sum(Charindex(e.U_Purchase,'30') * b.LineTotal * -1) As Cost03, -- 외주가공비
		   Cost04 = 0,
		   Cost05 = 0,
		   Cost06 = 0
	from ORPD a Inner Join RPD1 b On a.DocEntry =  b.DocEntry
				Inner Join PDN1 c On b.BaseEntry = c.DocEntry And b.BaseLine = c.LineNum
				Inner Join OPDN d On c.DocEntry = d.DocEntry
				Inner Join [@PS_MM070H] e On e.U_GRDocNum = d.DocEntry
	Where b.U_OrdNum = @ItemCode
	
	
	Union all
	
	select Cost01 = 0,
			Cost02 = 0,
			Cost03 = 0,
			Cost04 = 0,
			Cost05 = Isnull(Sum(Convert(Decimal(4,1),b.U_WorkTime) * c.U_Price),0), --자체가공비
			Cost06 = 0
	 from [@ps_pp040H] a Inner Join [@ps_pp040L] b On a.DocEntry = b.DocEntry
 						 Inner Join [@PS_PP001L] c On b.U_CpCode = c.U_CpCode
	where a.U_OrdType = '10'
	  and a.Canceled = 'N'
	  and b.U_PP030HNo = @PP030HNo
	  and b.U_CpCode Not In ('CP21301', 'CP21302')
	  
	  Union all
	  
	  select Cost01 = 0,
		   Cost02 = 0,
		   Cost03 = 0,
		   Cost04 = 0,
		   Cost05 = 0,
		   Cost06 = Isnull(Sum(Convert(Decimal(4,1),b.U_WorkTime) * c.U_PsmtP),0) --PSMT가공비
	 from [@ps_pp040H] a Inner Join [@ps_pp040L] b On a.DocEntry = b.DocEntry
 						 Inner Join [@PS_PP001L] c On b.U_CpCode = c.U_CpCode
	where a.U_OrdType = '20'
	  and a.Canceled = 'N'
	  and b.U_PP030HNo = @PP030HNo
	  and b.U_CpCode Not In ('CP21301', 'CP21302')

	Union all

	select Cost01 = 0,
		   Cost02 = 0,
		   Cost03 = 0,
		   Cost04 = Isnull(Sum(Convert(Decimal(4,1),b.U_PQty) * c.U_Price),0), --설계비
		   Cost05 = 0,
		   Cost06 = 0
	 from [@ps_pp040H] a Inner Join [@ps_pp040L] b On a.DocEntry = b.DocEntry
 						 Inner Join [@PS_PP001L] c On b.U_CpCode = c.U_CpCode
	where a.U_OrdType = '70'
	  and a.Canceled = 'N'
	  and b.U_PP030HNo = @PP030HNo
	  and b.U_CpCode In ('CP21301', 'CP21302')
	   ) t1
	   
	
	Union all
	
	select Cost01 = 0,
		   Cost02 = 0,
		   Cost03 = 0,
		   Cost04 = 0,
		   Cost05 = 0,
		   Cost06 = 0,
		   CostTot = 0,
		   aCost01 = Isnull(Sum(b.U_Cost01),0),
		   aCost02 = Isnull(Sum(b.U_Cost02),0),
		   aCost03 = Isnull(Sum(b.U_Cost03),0),
		   aCost04 = Isnull(Sum(b.U_Cost04),0),
		   aCost05 = Isnull(Sum(b.U_Cost05),0),
		   aCost06 = Isnull(Sum(b.U_Cost06),0),
		   aCostTot = Isnull(Sum(b.U_CostTot),0),
		   tSaleQty = Isnull(Sum(b.U_PQty),0),
		   tSaleAmt = Isnull(Sum(b.U_SaleAmt),0)
	  from [@PS_PP080H] a Inner Join [@PS_PP080L] b On a.Docentry = b.DocEntry
	 where a.Canceled = 'N'
	   and b.U_PP030No = @PP030HNo
  	) t2
END

--exec [PS_PP080_05] CP201103005, '0'