SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************************************************************************************************************/
/*  Module         : 구매관리																				    */
/*  Description    : 검수현황																					*/
/*  ALTER  Date    : 2010.11.17  																				*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_MM204_01]
ALTER     PROC [dbo].[PS_MM204_01]
(
  @StrDate			as datetime,
  @EndDate			as datetime,
  @SItemCode		as nvarchar(20),
  @EItemCode		as nvarchar(20)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
Select	T0.DocDate,
		'입고PO' As Type,
		Convert(Nvarchar,  T0.DocNum) + '-' + Convert(Nvarchar, T1.LineNum+1) As DocLine,
		T0.CardCode,
		T0.CardName,
		T1.ItemCode,
		T1.Dscription,
		convert(nvarchar(100),T3.U_Size) as Size,
		T4.Name as Mark,
		T5.Name as ItemType,
		T1.U_Qty as Qty,
		T1.Quantity,
		T3.InvntryUom,
		T1.Price,
		T1.Currency,
		Case When T0.DocCur = 'KRW' Then T1.LineTotal Else T1.TotalFrgn End As LineTotal,
		T0.DocCur
  From	[OPDN] T0 Inner Join [PDN1] T1 On T1.DocEntry = T0.DocEntry 
				  Left  Join [OITM] T3 On T1.ItemCode = T3.ItemCode
				  Left	Join [@PSH_Mark] T4 On T3.U_Mark = T4.Code
    			  Left  Join [@PSH_SHAPE] T5 on T3.U_ItemType = T5.Code
 Where	T0.DocDate Between @StrDate And @EndDate
   And	T1.ItemCode Between @SItemCode And @EItemCode
   --And	T1.WhsCode Like @WhsCode
--	Order by T0.DocDate, T0.CardCode, T1.ItemCode, T1.Quantity

Union All

Select	T0.DocDate,
		'반품' As Type,
		Convert(Nvarchar,  T0.DocNum) + '-' + Convert(Nvarchar, T1.LineNum+1) As DocLine,
		T0.CardCode,
		T0.CardName,
		T1.ItemCode,
		T1.Dscription,
		convert(nvarchar(100),T3.U_Size) as Size,	
		T4.Name as Mark,	
		T5.Name as ItemType,
		- T1.U_Qty as Qty,						
		- T1.Quantity,
		T3.InvntryUom,
		T1.Price,
		T1.Currency,
		Case When T0.DocCur = 'KRW' Then - T1.LineTotal Else - T1.TotalFrgn End As LineTotal,
		T0.DocCur
  From	[ORPD] T0 Inner Join [RPD1] T1 On T1.DocEntry = T0.DocEntry 
				  Left  Join [OITM] T3 On T1.ItemCode = T3.ItemCode
				  Left	Join [@PSH_Mark] T4 On T3.U_Mark = T4.Code
    			  Left  Join [@PSH_SHAPE] T5 on T3.U_ItemType = T5.Code				  
 Where	T0.DocDate Between @StrDate And @EndDate
   And	T1.ItemCode Between @SItemCode And @EItemCode
   --And	T1.WhsCode Like @WhsCode
--	Order by T0.DocDate, T0.CardCode, T1.ItemCode, T1.Quantity

Union All

Select	T0.DocDate,
		'A/P송장' As Type,
		Convert(Nvarchar,  T0.DocNum) + '-' + Convert(Nvarchar, T1.LineNum+1) As DocLine,
		T0.CardCode,
		T0.CardName,
		T1.ItemCode,
		T1.Dscription,
		convert(nvarchar(100),T3.U_Size) as Size,	
		T4.Name as Mark,	
		T5.Name as ItemType,
		T1.U_Qty as Qty,				
		T1.Quantity,
		T3.InvntryUom,
		T1.Price,
		T1.Currency,
		Case When T0.DocCur = 'KRW' Then T1.LineTotal Else T1.TotalFrgn End As LineTotal,
		T0.DocCur
  From	[OPCH] T0 Inner Join [PCH1] T1 On T1.DocEntry = T0.DocEntry And T1.BaseType <> '20'
				  Left  Join [OITM] T3 On T1.ItemCode = T3.ItemCode
				  Left	Join [@PSH_Mark] T4 On T3.U_Mark = T4.Code	
    			  Left  Join [@PSH_SHAPE] T5 on T3.U_ItemType = T5.Code				  			  
 Where	T0.DocDate Between @StrDate And @EndDate
   And	T1.ItemCode Between @SItemCode And @EItemCode
   --And	T1.WhsCode Like @WhsCode
--	Order by T0.DocDate, T0.CardCode, T1.ItemCode, T1.Quantity

Union All

Select	T0.DocDate,
		'A/P대변메모' As Type,
		Convert(Nvarchar,  T0.DocNum) + '-' + Convert(Nvarchar, T1.LineNum+1) As DocLine,
		T0.CardCode,
		T0.CardName,
		T1.ItemCode,
		T1.Dscription,
		convert(nvarchar(100),T3.U_Size) as Size,		
		T4.Name as Mark,		
		T5.Name as ItemType,	
		- T1.U_Qty as Qty,	
		- T1.Quantity,
		T3.InvntryUom,
		T1.Price,
		T1.Currency,
		Case When T0.DocCur = 'KRW' Then - T1.LineTotal Else - T1.TotalFrgn End As LineTotal,
		T0.DocCur
  From	[ORPC] T0 Inner Join [RPC1] T1 On T1.DocEntry = T0.DocEntry 
    			  Left  Join [OITM] T3 On T1.ItemCode = T3.ItemCode
				  Left	Join [@PSH_Mark] T4 On T3.U_Mark = T4.Code   
    			  Left  Join [@PSH_SHAPE] T5 on T3.U_ItemType = T5.Code				   			  
 Where	T0.DocDate Between @StrDate And @EndDate
   And	T1.ItemCode Between @SItemCode And @EItemCode
   --And	T1.WhsCode Like @WhsCode
Order by T0.DocDate, T0.CardCode, T1.ItemCode

----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_MM204_01] '20100101', '20101130', '1', 'ZZZZZZZZ'