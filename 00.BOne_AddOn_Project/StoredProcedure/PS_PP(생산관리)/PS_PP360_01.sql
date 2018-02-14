USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP360_01]    Script Date: 11/04/2010 19:58:24 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO

/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 공정코드 LIST                                                                     */
/*  Create Date    : 2010.12.03                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP360_01]
--CREATE PROCEDURE [dbo].[PS_PP360_01]
(
   @CpBCode AS NVARCHAR(20)
)
AS
declare @CpBCode01 as nvarchar(20),
		@CpBCode02 as nvarchar(20)

IF @CpBCode=''
  BEGIN
	SET @CpBCode01 ='CP101'
	SET @CpBCode02 ='CP999'
  END
  
IF @CpBCode<>''
  BEGIN
	SET @CpBCode01 = @CpBCode
	SET @CpBCode02 = @CpBCode
  END

SELECT CONVERT(NVARCHAR(20),A.U_CpBCode) AS CpBCode,
	   CONVERT(NVARCHAR(60),A.U_CpBName) AS CpBName,
	   CONVERT(NVARCHAR(20),B.U_CpCode)  AS CpCode,
	   CONVERT(NVARCHAR(60),B.U_CpName)  AS CpName,
	   CONVERT(NVARCHAR(20),B.U_DeptName) AS DeptName,
	   CONVERT(NVARCHAR(30),B.U_PartName) AS U_PartName,
	   CONVERT(NVARCHAR(30),B.U_WkClsNam)	AS WkClsNam,
	   SUM(B.U_PsmtP) AS PsmtP,
	   SUM(B.U_Price) AS Price,
	   CONVERT(NVARCHAR(10),B.U_Unit) AS Unit
  FROM [@PS_PP001H] AS A INNER JOIN [@PS_PP001L] AS B
		ON A.Code = B.Code
 
 WHERE A.U_CpBCode BETWEEN @CpBCode01 AND @CpBCode02
 		
GROUP BY A.U_CpBCode, A.U_CpBName, B.U_CpCode, B.U_CpName, B.U_DeptName, B.U_PartName, B.U_WkClsNam, B.U_Unit

ORDER BY A.U_CpBCode, B.U_CpCode

--exec [PS_PP360_01] 'CP101'
