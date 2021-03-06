SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 생산관리																				    */
/*  Description    : 휘팅서울포장등록																				*/
/*  ALTER  Date    : 2010.10.29																				*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_PP077_04]
ALTER     PROC [dbo].[PS_PP077_04]
(
  @PorNum		as nvarchar(15)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
--select * from [@PS_PP030H]
--select * from [@PS_PP030M]
--update [@PS_PP030M] set u_reportyn='N' where U_CpCode in ('CP30112','CP30114')
-----------------------------------------------------------------------------------------------------------------------------------------


select	a.U_OrdNum,
		b.U_Sequence,
		b.U_CpBCode,
		b.U_CpBName,
		b.U_CpCode,
		b.U_CpName,
		b.U_ReportYN,
		b.LineId
from [@PS_PP030H] a inner join [@PS_PP030M] b on a.docentry=b.docentry
where a.U_OrdNum=@PorNum
and   b.U_CpCode in ('CP30112','CP30114')
and   b.U_ReportYN = 'N'

-----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_PP077_04] '20101101002'