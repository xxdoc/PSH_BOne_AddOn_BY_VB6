/*==========================================================================================
        프로시저명      : PH_PY104_Grid2
        프로시저설명    : 고정수당공제금액일괄등록_조회
        만든이          : 
        작업일자        : 2012-11-12
        작업지시자      : 
        작업지시일자    : 
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104_Grid2' AND xtype = 'P'))
	DROP PROCEDURE PH_PY104_Grid2
GO

CREATE PROC PH_PY104_Grid2 (
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
			'Code'		= Code,							--사번
			'FullName'	= U_FullName					--성명			
	INTO #TEMP
	FROM [@PH_PY001A]
	WHERE U_CLTCOD = @CLTCOD
	AND (@TeamCode = '%' OR (@TeamCode <> '%' AND U_TeamCode = @TeamCode) )
	AND (@RspCode = '%' OR (@RspCode <> '%' AND U_RspCode = @RspCode))
	AND (@PAYTYP = '%' OR (@PAYTYP <> '%' AND U_PAYTYP = @PAYTYP))
	AND (U_JIGCOD BETWEEN @JIGCODF and @JIGCODT)
	AND (U_HOBONG BETWEEN @HOBONGF and @HOBONGT)
	AND U_Status <> '5'
	
IF NOT (EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104_TEMP2' AND xtype = 'U'))
begin	
	create table PH_PY104_TEMP2(
		CODE NVARCHAR(10), 
		NAME NVARCHAR(10)
	)
end

BEGIN
	INSERT INTO PH_PY104_TEMP2 
	SELECT Code, FullName
	FROM #TEMP
END

--Exec PH_PY104_Grid2 '1','1200','',''

--go
/*
 select Code, U_FullName from [@PH_PY001A] where U_CLTCOD = '1' AND U_PAYTYP='4'
 Exec PH_PY104_Grid2 '1','%','%','2','0000000000','ZZZZZZZZZZ','0000000000','ZZZZZZZZZZ'
 select * from PH_PY104_TEMP2
 delete PH_PY104_TEMP2
*/
