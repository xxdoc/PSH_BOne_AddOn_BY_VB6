SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 품질관리																				    */
/*  Description    : 원소재성분입력																				*/
/*  ALTER  Date    : 2010.11.09																					*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_QM030_01]
ALTER     PROC [dbo].[PS_QM030_01]
(
  @LotNo		as nvarchar(32),
  @ItemCode		as nvarchar(20)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
select	a.ItemCode,
		a.DistNumber,
		a.U_ChemC_Fe,
		a.U_ChemC_P,
		a.InDate,
		b.Quantity

from OBTN a inner join ITL1 b on a.ItemCode=b.ItemCode and a.SysNumber=b.SysNumber
			inner join OITL c on b.LogEntry=c.LogEntry

where c.DocType='59'
and	  left(a.DistNumber,8) = @LotNo
and	  a.ItemCode = @ItemCode

-----------------------------------------------------------------------------------------------------------------------------------------
--EXEC PS_QM030_01 'AABBCCDD', '502010069'