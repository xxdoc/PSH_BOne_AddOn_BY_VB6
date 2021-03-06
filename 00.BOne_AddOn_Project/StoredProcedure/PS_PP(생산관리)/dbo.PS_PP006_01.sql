USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP006_01]    Script Date: 11/04/2010 12:56:18 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 생산 > 외주업체별 가공비단가 SELECT[PS_PP006]				                                */
/*  Create Date    : 2010.10.26                                                                               */
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP006_01]
(
    @CardCode  AS NVARCHAR(20),
    @ItemCode  AS NVARCHAR(20),
    @CtrDate   AS NVARCHAR(20)
)
AS
BEGIN
	SELECT Code, U_eCardCod, U_eCardNam, U_ItemCode, U_ItemName, U_Cprice, 
		   CONVERT(NVARCHAR(8), U_CtrDate, 112) AS U_CtrDate
	  FROM [@PS_PP006H]
	 WHERE (U_eCardCod = @CardCode OR @CardCode = '')
	   AND (U_ItemCode = @ItemCode OR @ItemCode = '')
	   AND (U_CtrDate = @CtrDate OR @CtrDate = '')
 End
 
-- EXEC PS_PP006_01 '', '', ''
