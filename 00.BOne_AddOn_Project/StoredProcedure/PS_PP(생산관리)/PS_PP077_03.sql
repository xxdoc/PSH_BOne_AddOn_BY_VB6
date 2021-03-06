USE [PSHDB_TEST2]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP077_03]    Script Date: 10/29/2010 11:02:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 생산관리																				    */
/*  Description    : 휘팅서울포장등록																				*/
/*  ALTER  Date    : 2010.10.28																				*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_PP077_03]
ALTER     PROC [dbo].[PS_PP077_03]
(
  @MovDocNo		as nvarchar(15),
  @PP070No		as nvarchar(15),
  @PP070NoL		as nvarchar(10)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
--select * from [@PS_PP070H]
--select * from [@PS_PP070L]	
-----------------------------------------------------------------------------------------------------------------------------------------

select	U_OrdNum,
		U_OrdSub1,
		U_OrdSub2,
		U_PP030HNo,
		U_PP030MNo,
		U_BPLId,
		U_ItemCode,
		U_ItemName,
		U_CpCode,
		U_CpName
 
from [@PS_PP070L]
where U_MovDocNo  = @MovDocNo
and   DocEntry    = @PP070No
and	  LineId      = @PP070NoL



-----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_PP077_03] '20101025004','1','1'