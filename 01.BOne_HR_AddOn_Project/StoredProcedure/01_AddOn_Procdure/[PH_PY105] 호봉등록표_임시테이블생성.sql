/*==========================================================================================
        프로시저명      : PH_PY105_TEMP_CHK
        프로시저설명    : 호봉등록표_엑셀업로드_임시테이블 생성
        만든이          : 
        작업일자        : 2012-11-14
        작업지시자      : 
        작업지시일자    : 
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY105_TEMP_CHK' AND xtype = 'P'))
	DROP PROCEDURE PH_PY105_TEMP_CHK
GO

CREATE PROC PH_PY105_TEMP_CHK
AS

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

IF NOT (EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY105_TEMP' AND xtype = 'U'))
	begin	
		create table PH_PY105_TEMP(
			JIGCOD		 NVARCHAR(10),	-- 직급코드
			HOBCOD		 NVARCHAR(10),	-- 호봉코드
			HOBNAM		 NVARCHAR(20),	-- 호봉이름
			STDAMT		 NVARCHAR(30),	-- 급여기본
			BNSAMT		 NVARCHAR(30),	-- 상여기본
			EXTAMT01	 NVARCHAR(20),	-- 제수당01
			EXTAMT02	 NVARCHAR(20),	-- 제수당02
			EXTAMT03	 NVARCHAR(20),	-- 제수당03
			EXTAMT04	 NVARCHAR(20),	-- 제수당04
			EXTAMT05	 NVARCHAR(20),	-- 제수당05
			EXTAMT06	 NVARCHAR(20),	-- 제수당06
			EXTAMT07	 NVARCHAR(20),	-- 제수당07
			EXTAMT08	 NVARCHAR(20),	-- 제수당08
			EXTAMT09	 NVARCHAR(20),	-- 제수당09
			EXTAMT10	 NVARCHAR(20)	-- 제수당10
		)
		
	end
else
	begin
		DELETE PH_PY105_TEMP
	end
