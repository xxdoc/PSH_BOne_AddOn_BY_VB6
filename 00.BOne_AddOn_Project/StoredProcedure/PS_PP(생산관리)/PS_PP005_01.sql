USE [PSHDB_TEST2]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP005_01]    Script Date: 10/25/2010 13:19:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO

/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 제품원재료관계등록 [PS_PP005]                                                                     */
/*  Create Date    : 2010.10.22                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP005_01]
--CREATE PROCEDURE [dbo].[PS_PP005_01]
 (
 
	@ItmBsort   	As Nvarchar(10),
	@ItmMsort		As Nvarchar(10),	
	@ItemCod1		As NVarchar(20),
	@ItemCod2       As NVarchar(20)
)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON

    -- Insert statements for procedure here
	Select a.DocEntry,
		a.DocNum,
		a.U_Chk,
		a.U_DocNum,
		a.U_ItemCod1,
		a.U_ItemNam1,
		c.U_CodeName,
		a.U_ItemCod2,
		a.U_ItemNam2,
		a.U_InDate,
		a.U_MoDate,
		c.U_CodeName As U_itmMSort
From	[@PS_PP005H] a inner join [OITM] b
   on a.U_ItemCod1 = b.ItemCode 
       inner join [@PSH_ITMMSORT] c
   on c.U_Code = b.U_ItmMsort
Where	IsNull(a.U_ItemCod1, '') Like @ItemCod1
  And   IsNull(a.U_ItemCod2, '') Like @ItemCod2
  And   IsNull(b.U_ItmBSort,'') Like @ItmBSort
  And   ISNULL(b.U_ItmMSort,'') like @ItmMSort
Order by U_ItemCod1
END