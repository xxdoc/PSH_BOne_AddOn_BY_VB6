/*==========================================================================================
        프로시저명      : PH_PY104_Grid1
        프로시저설명    : 고정수당공제금액일괄등록_조회
        만든이          : 
        작업일자        : 2012-11-12
        작업지시자      : 
        작업지시일자    : 
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104_Grid1' AND xtype = 'P'))
	DROP PROCEDURE PH_PY104_Grid1
GO

CREATE PROC PH_PY104_Grid1 (
	@GBN		AS NVARCHAR(10),	-- 사업장
    @CSUCOD		AS NVARCHAR(10),	-- 부서
    @CSUNAM		AS NVARCHAR(10),	-- 담당
    @SEQ		AS NVARCHAR(10)		-- 금여형태
)

AS
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

IF not (EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104_TEMP' AND xtype = 'U'))
begin	
	create table PH_PY104_TEMP(
		GUBUN NVARCHAR(10), 
		CSUCOD NVARCHAR(10), 
		CSUNAM NVARCHAR(10), 
		SEQ NVARCHAR(10)
	)
	
end

	
INSERT into PH_PY104_TEMP(GUBUN,CSUCOD,CSUNAM,SEQ) values (@GBN,@CSUCOD,@CSUNAM,@SEQ)

--go
/*
 EXEC PH_PY104_Grid1 '수당','E11','test3','2'
 select * from PH_PY104_TEMP
 drop table PH_PY104_TEMP
 delete PH_PY104_TEMP
*/
