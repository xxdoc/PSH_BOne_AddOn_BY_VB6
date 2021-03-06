USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM060_02]    Script Date: 11/04/2010 12:53:10 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : 작업지시등록 > 히스토리 테이블 INSERT[PS_MM060]			                                */
/*  Create Date    : 2010.10.02                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_MM060_02]
(
	@U_inDate   AS DateTime,
	@Code       AS Nvarchar(20)
)
AS
BEGIN
UPDATE [@PS_MM060L]
   SET U_ItmBcode = A.U_ItmBcode,
	   U_ItmBname = A.U_ItmBname,
	   U_ItmMcode = A.U_ItmMcode,
	   U_ItmMname = A.U_ItmMname,
	   U_inDate = @U_inDate
  FROM [@PS_MM060H] AS A INNER JOIN [@PS_MM060L] AS B
    ON A.Code = B.Code
 WHERE @Code = '' Or A.Code = @Code
END

-- EXEC PS_MM060_02 '', '', '', '', ''
