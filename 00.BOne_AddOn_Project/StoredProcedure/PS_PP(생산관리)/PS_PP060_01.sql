USE [PSHDB_TEST]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP060_01]    Script Date: 10/15/2010 20:34:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO

/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 작업외공수등록 [PS_PP060]                                                                     */
/*  Create Date    : 2010.10.14                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Noh Geun Yong																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP060_01]
--CREATE PROCEDURE [dbo].[PS_PP060_01]
 (
 
	@BPLId			As Nvarchar(10),
	@CpGbn			As Nvarchar(10),	
	@CpCode			As Nvarchar(10),
	@CntcCode		As Nvarchar(10),
	@WorkGbn		As Nvarchar(10),
	@DocDateFr		As Nvarchar(8),
	@DocDateTo		As NVarchar(8)
)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON

    -- Insert statements for procedure here
	Select a.DocEntry,
		a.DocNum,
		a.U_BPLId,
		a.U_LineNum,
		a.U_CpGbn,
		a.U_CpCode,
		a.U_CpName,
		a.U_CntcCode,
		a.U_CntcName,
		a.U_ItmBsort,
		a.U_DocDate,
		a.U_WorkNote,
		a.U_WorkTime,
		a.U_WorkGbn
From	[@PS_PP060H] a
Where	IsNull(a.U_BPLId, 0) Like @BPLId
  And   IsNull(a.U_CpGbn,'') Like @CpGbn
  And   ISNULL(a.U_CpCode,'') like @CpCode
  And	IsNull(a.U_CntcCode, '') Like @CntcCode
  And	IsNull(a.U_WorkGbn, '') Like @WorkGbn
  And	a.U_DocDate Between @DocDateFr And @DocDateTo
Order by a.DocNum, a.U_LineNum
END
