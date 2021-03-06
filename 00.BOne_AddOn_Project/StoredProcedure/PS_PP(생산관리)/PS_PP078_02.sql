USE [PSHDB_TEST2]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP078_02]    Script Date: 11/01/2010 19:17:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 생산관리																				    */
/*  Description    : 휘팅서울포장등록취소																		*/
/*  ALTER  Date    : 2010.10.26																					*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_PP078_02]
ALTER     PROC [dbo].[PS_PP078_02]
(
  @MovDocNo		as nvarchar(15),
  @PP070No		as nvarchar(15),
  @ItemCode		as nvarchar(20)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
select	Canceled,
		U_MovDocNo,
		U_PP070No,
		U_PP070NoL,
		U_PorNum,
		U_ItemCode,
		U_ItemName,
		U_PkQty,
		U_PkWt,
		U_NPkQty,
		U_NPkWt,
		U_OPkQty,
		U_OPkWt,
		U_InDate,
		U_PP040No,
		DocNum
from	[@PS_PP077H]
where	Canceled <> 'Y'
and		U_MovDocNo = @MovDocNo
and		U_PP070No = @PP070No
and		U_ItemCode=@ItemCode
-----------------------------------------------------------------------------------------------------------------------------------------
--EXEC PS_PP078_02 '20101101003','1003','101010006'