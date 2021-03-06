USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP035_02]    Script Date: 11/04/2010 12:55:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 작업지시리스트조회 > 작지 테이블 UPDATE[PS_PP035]				                                */
/*  Create Date    : 2010.10.10                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP035_02]
(
	@DocEntry AS INTEGER,
    @Canceled AS NVARCHAR(20),
    @SelWt	  AS NUMERIC(19,6),
    @CntcCode AS NVARCHAR(20),
    @CntcName AS NVARCHAR(20),
    @DocDate  AS DATETIME,
    @DueDate  AS DATETIME
)
AS
BEGIN
	UPDATE	[@PS_PP030H]
	   SET	Canceled = @Canceled,
			U_SelWt = @SelWt,
			U_CntcCode = @CntcCode,
			U_CntcName = @CntcName,
			U_DocDate = @DocDate,
			U_DueDate = @DueDate
	 Where	DocNum = @DocEntry
 END
 
-- EXEC PS_PP035_02 '2', '', '', '', '', '', ''
-- EXEC PS_PP035_02 '1', 'Y', '40', '177', '강정석', '', ''

