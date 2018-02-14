USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_SD610_01]    Script Date: 08/22/2013 18:32:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/******************************************************************************************************************/
/*  Module         : SD								    														*/
/*  Description    : 내부거래출고 현황				  															*/
/*  Create Date    : 2013.08.22                                                                                 */
/*  Modified Date  :									       													*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/******************************************************************************************************************/
ALTER PROC [dbo].[PS_SD610_01]
--Create PROC [dbo].[PS_SD610_01]
(
	@BPLId        NVARCHAR(01),
	@CardCode	  NVARCHAR(10),
	@DocDateFr    datetime,
	@DocDateTo    datetime
)
AS

Select g.DocDate,
	   g.CardCode,
	   g.CardName,
	   g.ItemCode,
	   g.ItemName,
	   g.U_OutSize,
	   Qty = Sum(g.Qty),
	   Danga = Round(Sum(g.Amt) / Sum(g.Qty),0),
	   Amt = Sum(g.Amt)
From (

Select t0.DocDate,
	   t1.CardCode,
	   t1.CardName,
	   t0.ItemCode,
	   t2.ItemName,
	   t2.U_OutSize,
	   Qty = Sum(t0.OutQty),
	   Amt = Sum(t0.TransValue) * -1
  From OINM t0 Inner join (Select a.DocEntry, b.ItemCode, CardCode = a.U_CardCode, CardName = a.CardName
							from OIGE a inner Join IGE1 b On a.DocEntry = b.DocEntry
						 				  Inner Join OWHS c ON c.WhsCode = b.WhsCode
										  Inner Join OITM d On b.ItemCode = d.Itemcode
							Where a.DocDate between @DocDateFr and @DocDateTo
							  And (Case When @BPLId = '3' Or @BPLId = '5' Then right(c.WhsCode,1) Else a.BPLId End = (Case When @BPLId = '6' Then '3' Else @BPLId End) )
							  and a.U_CtrlType = 'T' ) t1 On t0.TransType = '60' And t0.Base_Ref = t1.DocEntry and t0.ItemCode = t1.ItemCode
				Inner Join OITM t2 ON t0.ItemCode = t2.ItemCode
Where t0.DocDate between @DocDateFr and @DocDateTo
  And Isnull(t1.CardCode,'') like @CardCode + '%'
Group by t0.Docdate, 
		t1.CardCode,
	   t1.CardName,
	   t0.ItemCode,
	   t2.ItemName,
	   t2.U_OutSize
	   
Union all

Select t0.DocDate,
	   t1.CardCode,
	   t1.CardName,
	   t0.ItemCode,
	   t2.ItemName,
	   t2.U_OutSize,
	   Qty = Sum(t0.InQty) * -1,
	   Amt = Sum(t0.TransValue) * -1
  From OINM t0 Inner join (Select a.DocEntry, b.ItemCode, CardCode = a.U_CardCode, CardName = a.CardName
							from OIGN a inner Join IGN1 b On a.DocEntry = b.DocEntry
						 				  Inner Join OWHS c ON c.WhsCode = b.WhsCode
										  Inner Join OITM d On b.ItemCode = d.Itemcode
										  Inner Join OIGE e On e.DocEntry = a.U_CancDoc
							Where a.DocDate between @DocDateFr and @DocDateTo
							  and (Case When @BPLId = '3' Or @BPLId = '5' Then right(c.WhsCode,1) Else a.BPLId End = (Case When @BPLId = '6' Then '3' Else @BPLId End) )
							  and a.U_CtrlType = 'T' ) t1 On t0.TransType = '59' And t0.Base_Ref = t1.DocEntry and t0.ItemCode = t1.ItemCode
				Inner Join OITM t2 ON t0.ItemCode = t2.ItemCode
Where t0.DocDate between @DocDateFr and @DocDateTo
  And Isnull(t1.CardCode,'') like @CardCode + '%'
Group by t0.Docdate, 
		t1.CardCode,
	   t1.CardName,
	   t0.ItemCode,
	   t2.ItemName,
	   t2.U_OutSize
) g
Group by g.DocDate,
	   g.CardCode,
	   g.CardName,
	   g.ItemCode,
	   g.ItemName,
	   g.U_OutSize
	   
--exec [PS_SD610_01] '3', '%', '20130701', '20130731'