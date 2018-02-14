/*==========================================================================================
        프로시저명      : PH_PY104
        프로시저설명    : 고정수당공제금액일괄등록_조회
        만든이          : 
        작업일자        : 2012-11-12
        작업지시자      : 
        작업지시일자    : 
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104' AND xtype = 'P'))
	DROP PROCEDURE PH_PY104
GO

CREATE PROC PH_PY104 (
	@CLTCOD		AS NVARCHAR(10),	-- 사업장
    @TeamCode	AS NVARCHAR(10),	-- 부서
    @RspCode	AS NVARCHAR(10),	-- 담당
    @PAYTYP		AS NVARCHAR(10),	-- 금여형태
    @JIGCODF	AS NVARCHAR(10),	-- 직급From
    @JIGCODT	AS NVARCHAR(10),	-- 직급To
    @HOBONGF	AS NVARCHAR(10),	-- 호봉From
    @HOBONGT	AS NVARCHAR(10)		-- 호봉To
)

AS

--SELECT @SDate, @EDate
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

	SELECT	--'LineNum'	= CONVERT(INT,ROW_NUMBER() OVER (PARTITION BY T0.DOCENTRY ORDER BY T0.DOCENTRY) -1 ),
			'Code'		= Code,						--사번
			'FullName'	= U_FullName					--성명			
	FROM [@PH_PY001A]
	WHERE U_CLTCOD = @CLTCOD
	AND (@TeamCode = '' OR U_TeamCode = @TeamCode) 
	AND (@RspCode = '' OR U_RspCode = @RspCode)
	AND (@PAYTYP = '' OR U_PAYTYP = @PAYTYP)
	AND (U_JIGCOD BETWEEN @JIGCODF and @JIGCODT)
	AND (U_HOBONG BETWEEN @HOBONGF and @HOBONGT)
	