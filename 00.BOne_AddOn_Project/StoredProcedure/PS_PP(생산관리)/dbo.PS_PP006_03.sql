USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP006_03]    Script Date: 11/04/2010 12:56:57 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 생산 > 외주업체별 가공비단가 UPDATE [PS_PP006]			 	                                */
/*  Create Date    : 2010.10.26																					*/
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP006_03]
(
	@Code		AS NVARCHAR(20),
    @eCardCod   AS NVARCHAR(20),
    @eCardNam   AS NVARCHAR(50),
    @ItemCode   AS NVARCHAR(20),
    @ItemName   AS NVARCHAR(50),
    @Cprice		AS NUMERIC(10,6),
    @CtrDate    AS NVARCHAR(8)
)
AS

BEGIN
	UPDATE	[@PS_PP006H]
	   SET	U_eCardCod = @eCardCod,
			U_eCardNam = @eCardNam,
			U_ItemCode = @ItemCode,
			U_ItemName = @ItemName,
			U_CPrice   = @Cprice,
			U_CtrDate  = @CtrDate
	 Where	Code = @Code
End

-- EXEC PS_PP006_03 '4', '10008', '(주)동방 부산지점', '101010008', 'ELBOW 22.5˚', '333', '19000101'
