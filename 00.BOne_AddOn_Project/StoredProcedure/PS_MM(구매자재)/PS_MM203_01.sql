SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************************************************************************************************************/
/*  Module         : 구매관리																				    */
/*  Description    : 입고현황																					*/
/*  ALTER  Date    : 2010.11.22  																				*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_MM203_01]
ALTER     PROC [dbo].[PS_MM203_01]
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
select	a.DocNum,
		convert(nvarchar(100),a.U_CardCode) U_CardCode, 
		convert(nvarchar(100),a.U_CardName)	U_CardName,
		a.U_BPLId, c.BPLName,
		a.U_CntcCode, a.U_CntcName,
		a.U_POType,
		a.U_POStatus,
		a.U_DocDate,
		--a.U_DueDateFr, a.U_DueDateTo
		
		convert(nvarchar(100),b.U_ItemCode) U_ItemCode, 
		convert(nvarchar(100),b.U_ItemName) U_ItemName,
		d.InvntryUom,
		convert(nvarchar(100),d.U_Size) as Size,
		e.Name as Mark,
		f.Name as ItemType,
		
		b.U_Qty,
		b.U_Weight,
		b.U_UnWeight,
		b.U_Price,
		b.U_LinTotal,
		b.U_WhsCode,
		b.U_WhsName
						
from [@PS_MM050H] a inner join [@PS_MM050L] b on a.docentry=b.docentry
				    left  join [OBPL] c on a.U_BPLId=c.BPLId
					left  join [OITM] d On b.U_ItemCode=d.ItemCode
					left  join [@PSH_Mark] e On d.U_Mark=e.Code
    			    left  join [@PSH_SHAPE] f on d.U_ItemType=f.Code				    
----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_MM203_01] '20100101', '20101130', '1', 'ZZZZZZZZ'