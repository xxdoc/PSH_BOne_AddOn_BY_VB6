USE [PSHDB_TEST2]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP077_02]    Script Date: 11/01/2010 19:12:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 생산관리																				    */
/*  Description    : 휘팅서울포장등록																				*/
/*  ALTER  Date    : 2010.10.26																					*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_PP077_02]
ALTER     PROC [dbo].[PS_PP077_02]
(
  @MovDocNo		as nvarchar(15),
  @PP070No		as nvarchar(15),
  @ItemCode		as nvarchar(20)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
--select * from [@PS_PP075H]
--select * from [@PS_PP075L]	
--select U_MovDocNo,* from [@PS_PP070L]

select	a.U_MovDocNo,
		left(b.U_PP070No,charindex('-',b.U_PP070No)-1),
		substring(b.U_PP070No,charindex('-',b.U_PP070No)+1,len(b.U_PP070No)-charindex('-',b.U_PP070No)),
		c.U_OrdNum,
		b.U_ItemCode,
		b.U_ItemName,
		sum(b.U_Qty), 
		sum(b.U_Weight),
		sum(b.U_Weight)-isnull(sum(d.NPkWt),0) NPkWt,
		sum(d.NPkQty) OPkQty,
		sum(d.NPkWt) OPkWt
from [@PS_PP075H] a inner join [@PS_PP075L] b on a.docentry=b.docentry
					left  join [@PS_PP070L] c on a.U_MovDocNo=c.U_MovDocNo and left(b.U_PP070No,charindex('-',b.U_PP070No)-1)=c.DocEntry and substring(b.U_PP070No,charindex('-',b.U_PP070No)+1,len(b.U_PP070No)-charindex('-',b.U_PP070No)) = c.LineId
					left  join (
								select U_MovDocNo, U_PP070No, U_PP070NoL, sum(U_NPkQty )NPkQty, sum(U_NPkWt) NPkWt
								from [@PS_PP077H]
								group by U_MovDocNo, U_PP070No, U_PP070NoL
								) d on a.U_MovDocNo=d.U_MovDocNo and left(b.U_PP070No,charindex('-',b.U_PP070No)-1)=d.U_PP070No and substring(b.U_PP070No,charindex('-',b.U_PP070No)+1,len(b.U_PP070No)-charindex('-',b.U_PP070No))=d.U_PP070NoL
					
					
where a.Canceled <> 'Y'
and	  a.U_MovDocNo like @MovDocNo		
and	  left(b.U_PP070No,charindex('-',b.U_PP070No)-1) like @PP070No 	
and   b.U_ItemCode=@ItemCode
group by a.U_MovDocNo,
		 left(b.U_PP070No,charindex('-',b.U_PP070No)-1),
		 substring(b.U_PP070No,charindex('-',b.U_PP070No)+1,len(b.U_PP070No)-charindex('-',b.U_PP070No)),
		 c.U_OrdNum,
		 b.U_ItemCode,
		 b.U_ItemName		
-----------------------------------------------------------------------------------------------------------------------------------------
--EXEC PS_PP077_02 '20101101001', '1003'
