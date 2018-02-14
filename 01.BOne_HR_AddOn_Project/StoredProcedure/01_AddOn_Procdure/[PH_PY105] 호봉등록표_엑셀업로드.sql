/*==========================================================================================
        프로시저명      : PH_PY105
        프로시저설명    : 호봉등록표_엑셀업로드
        만든이          : 
        작업일자        : 2012-11-14
        작업지시자      : 
        작업지시일자    : 
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY105' AND xtype = 'P'))
	DROP PROCEDURE PH_PY105
GO

CREATE PROC PH_PY105 (
    @JIGCOD		AS NVARCHAR(10),	-- 직급코드
    @HOBCOD		AS NVARCHAR(10),	-- 호봉코드
    @HOBNAM		AS NVARCHAR(20),	-- 호봉이름
    @STDAMT		AS NVARCHAR(30),	-- 급여기본
    @BNSAMT		AS NVARCHAR(30),	-- 상여기본
    @EXTAMT01	AS NVARCHAR(20),	-- 제수당01
    @EXTAMT02	AS NVARCHAR(20),	-- 제수당02
    @EXTAMT03	AS NVARCHAR(20),	-- 제수당03
    @EXTAMT04	AS NVARCHAR(20),	-- 제수당04
    @EXTAMT05	AS NVARCHAR(20),	-- 제수당05
    @EXTAMT06	AS NVARCHAR(20),	-- 제수당06
    @EXTAMT07	AS NVARCHAR(20),	-- 제수당07
    @EXTAMT08	AS NVARCHAR(20),	-- 제수당08
    @EXTAMT09	AS NVARCHAR(20),	-- 제수당09
    @EXTAMT10	AS NVARCHAR(20)		-- 제수당10
)

AS

--SELECT @SDate, @EDate
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

BEGIN
	INSERT INTO PH_PY105_TEMP 
	(JIGCOD, HOBCOD	, HOBNAM, STDAMT, BNSAMT, EXTAMT01, EXTAMT02, EXTAMT03, EXTAMT04, EXTAMT05, EXTAMT06,
	 EXTAMT07, EXTAMT08, EXTAMT09, EXTAMT10)
	VALUES (@JIGCOD, @HOBCOD, @HOBNAM, @STDAMT, @BNSAMT, @EXTAMT01, @EXTAMT02, @EXTAMT03, @EXTAMT04,
			@EXTAMT05, @EXTAMT06, @EXTAMT07, @EXTAMT08, @EXTAMT09, @EXTAMT10)

END

--go
/*
 EXEC PH_PY105 '','','',''
 select * from PH_PY105_TEMP
 delete PH_PY104_TEMP2
*/
