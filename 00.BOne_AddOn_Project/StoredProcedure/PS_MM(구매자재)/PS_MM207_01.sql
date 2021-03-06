USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM207_01]    Script Date: 11/25/2010 21:13:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************************************************************************************************************/
/*  Module         : 구매관리																				    */
/*  Description    : 출고대장    																				*/
/*  ALTER  Date    : 2010.11.22  																				*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_MM207_01]
ALTER     PROC [dbo].[PS_MM207_01]
(
  @StrDate			as datetime,
  @EndDate			as datetime,
  @SItemCode		as nvarchar(20),
  @EItemCode		as nvarchar(20),
  @OptBtnValue		as nvarchar(1)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
--//일반출고
if @OptBtnValue='A' begin

	select	convert(char(10),a.DocDate,120) DocDate,
			a.ItemCode,
			a.Dscription,
			convert(char(100),b.U_Size) Size,
			b.InvntryUom,
			convert(char(100),c.U_CodeName) Msort,
			sum(a.OutQty-a.InQty) as OutQty,
			sum(-isnull(a.transvalue,0)) as transvalue
	from [OINM] a inner join [OITM] b on a.ItemCode=b.ItemCode
				  left  join [@PSH_ITMMSORT] c on b.U_ItmMsort=c.U_Code
	where a.transtype in (59,60)
	and   a.DocDate between @StrDate and @EndDate
	and   a.ItemCode between @SItemCode and @EItemCode
	group by a.DocDate,a.ItemCode,a.Dscription,b.U_Size,b.InvntryUom,c.U_CodeName
	
--//기타자재출고
end else if @OptBtnValue='B' begin 

	select	convert(char(100),a.U_CardCode) U_CardCode,
			convert(char(100),a.U_CardName) U_CardName,
			a.U_InDate,
			convert(char(100),a.U_Purpose) U_Purpose,
			convert(char(100),a.U_OutNum) U_OutNum,
			convert(char(100),b.U_ItemCode) U_ItemCode,
			convert(char(100),b.U_ItemName) U_ItemName,
			convert(char(100),b.U_Size) U_Size,
			convert(char(100),b.U_Unit) U_Unit,
			b.U_Qty,
			b.U_Weight,
			b.U_Price
	from [@PS_MM090H] a inner join [@PS_MM090L] b on a.docentry=b.docentry
	where a.Canceled<>'Y'
	and	  a.U_InDate between @StrDate and @EndDate
	and	  b.U_ItemCode between @SItemCode and @EItemCode

--//외주출고
end else if @OptBtnValue='C' begin

	select	a.DocEntry,
			convert(char(100),a.U_CardCode) as CardCode,
			convert(char(100),a.U_CardName) as CardName,
			convert(char(10),a.U_DocDate,120) as DocDate,
			convert(char(100),b.U_OrdNum) as OrdNum,
			convert(char(100),b.U_ItemCode) as ItemCode,
			convert(char(100),b.U_ItemName) as ItemName,
			convert(char(100),b.U_Size) as Size,
			convert(char(100),b.U_OutItmCd) as OutItmCd,
			convert(char(100),b.U_OutItmNm) as OutItmNm,
			convert(char(100),C.U_Size) as OutSize,
			b.U_OutQty as OutQty,
			b.U_OutWt as OutWt
			
	from [@PS_MM130H] a inner join [@PS_MM130L] b on a.docentry=b.docentry
						left  join [OITM] c on b.U_OutItmCd=c.ItemCode
	where a.Canceled<>'Y'
	and	  a.U_DocDate between @StrDate and @EndDate
	and   b.U_ItemCode between @SItemCode and @EItemCode
end


----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_MM207_01] '20100101', '20101130', '1', 'ZZZZZZZZ', 'A'