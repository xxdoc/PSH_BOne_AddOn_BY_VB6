USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP006_02]    Script Date: 11/04/2010 12:56:36 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 생산 > 외주업체별 가공비단가 INSERT [PS_PP006]			 	                                */
/*  Create Date    : 2010.10.26																					*/
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP006_02]
(
	@Code		AS NVARCHAR(20),
    @eCardCod   AS NVARCHAR(20),
    @eCardNam   AS NVARCHAR(50),
    @ItemCode   AS NVARCHAR(20),
    @ItemName   AS NVARCHAR(50),
    @Cprice		AS NUMERIC(10,6),
    @CtrDate    AS DATETIME
)
AS

BEGIN
	INSERT INTO [@PS_PP006H]
				(Code, DocEntry, U_eCardCod, U_eCardNam, U_ItemCode, U_ItemName, U_Cprice, U_CtrDate)
	VALUES (@Code, @Code, @eCardCod, @eCardNam, @ItemCode, @ItemName, @Cprice, @CtrDate)	
End

-- EXEC PS_PP006_02 '', '', '', '', '', ''
