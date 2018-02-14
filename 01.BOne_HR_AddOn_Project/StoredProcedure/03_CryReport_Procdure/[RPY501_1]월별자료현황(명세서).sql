IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY501_1' AND xtype = 'P'))
	DROP PROCEDURE RPY501_1
GO

CREATE  PROC RPY501_1
	(
		@JSNYER 	AS Nvarchar(4),		--작업연월
		@STRMON 	AS Nvarchar(2), 	--시작월
		@ENDMON 	AS Nvarchar(2), 	--종료월
		@JOBGBN		AS Nvarchar(1), 	--작업구분(1연말정산,2중도정산,3전체)
		@CLTCOD		AS Nvarchar(8), 	--자사코드
		@MSTDPT		AS Nvarchar(8), 	--부서
	    @MSTCOD 	AS Nvarchar(8) 	   	--사원번호		
	)
	

 AS
	/*==========================================================================================
		프로시저명		: RPY501_1
		프로시저설명	: 월별자료현황(명세서)
		만든이			: 최동권
		작업일자		: 2009-12-28
		작업지시자		: 함미경
		작업지시일자	: 2009-12-28
		작업목적		: 
		작업내용		: 
	===========================================================================================*/
	--DROP PROC RPY501_1
	--Exec RPY501_1 '2013', '01', '12', '3', N'%', N'%',  N'%'
	--Exec RPY501_1 '2013', '01', '12', '3', '%', '%', '%'

	SET NOCOUNT ON
-- <1.임시테이블 생성 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    

	CREATE TABLE #RPY501_1 (
	        CLTCOD     nvarchar(8),
	        CLTNAM     nvarchar(50),
	        MSTCOD     nvarchar(8),
	        MSTNAM     nvarchar(50),
	        GWAPAY     Numeric(19,6),
		    BIGWA1     Numeric(19,6),
		    BIGWA2     Numeric(19,6),
	        BIGWA3     Numeric(19,6),
		    BIGWU3     Numeric(19,6),
	        BIGWA4     Numeric(19,6),
		    BIGWA5     Numeric(19,6),
	        BIGWA6     Numeric(19,6),
		    BIGWA7     Numeric(19,6),

			BIGG01     Numeric(19,6),
	        BIGH01     Numeric(19,6),
	        BIGH05     Numeric(19,6),
	        BIGH06     Numeric(19,6),
	        BIGH07     Numeric(19,6),

			BIGH08     Numeric(19,6),
	        BIGH09     Numeric(19,6),
	        BIGH10     Numeric(19,6),
	        BIGH11     Numeric(19,6),
	        BIGH12     Numeric(19,6),

			BIGH13     Numeric(19,6),
	        BIGI01     Numeric(19,6),
	        BIGK01     Numeric(19,6),
	        BIGM01     Numeric(19,6),
	        BIGM02     Numeric(19,6),

			BIGM03     Numeric(19,6),
	        BIGO01     Numeric(19,6),
	        BIGQ01     Numeric(19,6),
	        BIGS01     Numeric(19,6),
	        BIGT01     Numeric(19,6),

			BIGX01     Numeric(19,6),
	        BIGY01     Numeric(19,6),
	        BIGY02     Numeric(19,6),
	        BIGY03     Numeric(19,6),
	        BIGY20     Numeric(19,6),
	        BIGY21     Numeric(19,6),
	        BIGZ01     Numeric(19,6),

	        GWASEE     Numeric(19,6),
	        GWABNS     Numeric(19,6),
			JIGTOT     Numeric(19,6),
	        KUKAMT     Numeric(19,6),
	        MEDAMT     Numeric(19,6),
	        GBHAMT     Numeric(19,6),
	        GABGUN     Numeric(19,6),
	        JUMINN     Numeric(19,6),
	        TOTGON     Numeric(19,6) ) 
				        
-- <2.월별 자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
	INSERT INTO [#RPY501_1]
	SELECT	CLTCOD	=	T1.U_CLTCOD,
			CLTNAM	=	T3.Name,
			MSTCOD	=	T1.U_MstCode,
		    MSTNAM	=	T1.U_MstName,
		    GWAPAY	=	SUM(T0.U_GWAPAY),						-- 과세표준
		    BIGWA1	=	SUM(ISNULL(T0.U_BIGWA01,0)),						-- 비과세
		    BIGWA2	=	SUM(ISNULL(T0.U_BIGWA02,0)),
		    BIGWA3	=	SUM(ISNULL(T0.U_BIGWA03,0)),
		    BIGWU3	=	SUM(ISNULL(T0.U_BIGWU03,0)),
		    BIGWA4	=	SUM(ISNULL(T0.U_BIGWA04,0)),
		    BIGWA5	=	SUM(ISNULL(T0.U_BIGWA05,0)),
		    BIGWA6	=	SUM(ISNULL(T0.U_BIGWA06,0)),
		    BIGWA7	=	SUM(ISNULL(T0.U_BIGWA07,0)),

			BIGG01  =   SUM(ISNULL(T0.U_BIGG01,0)),
	        BIGH01  =   SUM(ISNULL(T0.U_BIGH01,0)),
	        BIGH05  =   SUM(ISNULL(T0.U_BIGH05,0)),
	        BIGH06  =   SUM(ISNULL(T0.U_BIGH06,0)),
	        BIGH07  =   SUM(ISNULL(T0.U_BIGH07,0)),
			BIGH08  =   SUM(ISNULL(T0.U_BIGH08,0)),
	        BIGH09  =   SUM(ISNULL(T0.U_BIGH09,0)),
	        BIGH10  =   SUM(ISNULL(T0.U_BIGH10,0)),
	        BIGH11  =   SUM(ISNULL(T0.U_BIGH11,0)),
	        BIGH12  =   SUM(ISNULL(T0.U_BIGH12,0)),
			BIGH13  =   SUM(ISNULL(T0.U_BIGH13,0)),
	        BIGI01  =   SUM(ISNULL(T0.U_BIGI01,0)),
	        BIGK01  =   SUM(ISNULL(T0.U_BIGK01,0)),
	        BIGM01  =   SUM(ISNULL(T0.U_BIGM01,0)),
	        BIGM02  =   SUM(ISNULL(T0.U_BIGM02,0)),
			BIGM03  =   SUM(ISNULL(T0.U_BIGM03,0)),
	        BIGO01  =   SUM(ISNULL(T0.U_BIGO01,0)),
	        BIGQ01  =   SUM(ISNULL(T0.U_BIGQ01,0)),
	        BIGS01  =   SUM(ISNULL(T0.U_BIGS01,0)),
	        BIGT01  =   SUM(ISNULL(T0.U_BIGT01,0)),
			BIGX01  =   SUM(ISNULL(T0.U_BIGX01,0)),
	        BIGY01  =   SUM(ISNULL(T0.U_BIGY01,0)),
	        BIGY02  =   SUM(ISNULL(T0.U_BIGY02,0)),
	        BIGY03  =   SUM(ISNULL(T0.U_BIGY03,0)),
	        BIGY20  =   SUM(ISNULL(T0.U_BIGY20,0)),
	        BIGY21  =   SUM(ISNULL(T0.U_BIGY21,0)),
	        BIGZ01  =   SUM(ISNULL(T0.U_BIGZ01,0)),

		    GWASEE	=	0,	-- 급여총액
		    GWABNS	=	SUM(T0.U_GWABNS + T0.U_INJBNS),				-- 상여총액
		    JIGTOT	=	SUM(T0.U_JIGTOTAL),							-- 총계
		    KUKAMT	=	SUM(T0.U_KUKAMT),							-- 국민연금
		    MEDAMT	=	SUM(T0.U_MEDAMT + ISNULL(T0.U_NGYAMT,0)),	-- 건강보험
		    GBHAMT	=	SUM(T0.U_GBHAMT),							-- 고용보험
		    GABGUN	=	SUM(T0.U_GABGUN),							-- 갑근세
		    JUMINN	=	SUM(T0.U_JUMIN),							-- 주민세
		    TOTGON	=	SUM(T0.U_KUKAMT + T0.U_MEDAMT + T0.U_GBHAMT + T0.U_GABGUN + T0.U_JUMIN + ISNULL(T0.U_NGYAMT,0))
	FROM	[@ZPY343L] T0 	
			INNER JOIN [@ZPY343H] T1 ON T0.DocEntry  = T1.DocEntry
			INNER JOIN [@PH_PY001A] T2 ON T1.U_MstCode = T2.Code
			LEFT JOIN [@PH_PY005A] T3 ON T1.U_CLTCOD = T3.U_CLTCode
	WHERE 	T1.U_JsnYear    = @JSNYER
	AND		T0.U_LineNum   >= @STRMON
	AND		T0.U_LineNum   <= @ENDMON 
	AND		T1.U_CLTCOD  LIKE @CLTCOD
	AND		T1.U_DptCode LIKE @MSTDPT 
	AND		T1.U_MstCode LIKE @MSTCOD
	AND		T2.U_Status    LIKE CASE @JOBGBN WHEN '1' THEN '1' 
						  				   WHEN '2' THEN '4'
										   ELSE '%' END
	GROUP	BY T1.U_CLTCOD, T3.NAME, T1.U_MstCode,  T1.U_MstName                                                 
	ORDER	BY T1.U_CLTCOD, T1.U_MstName,  T1.U_MstCode

    UPDATE  [#RPY501_1]
    SET     GWASEE  =   ISNULL(GWAPAY,0) 
                      + ISNULL(BIGWA1,0) + ISNULL(BIGWA2,0) + ISNULL(BIGWA3,0) + ISNULL(BIGWU3,0)
                      + ISNULL(BIGWA4,0) + ISNULL(BIGWA5,0) + ISNULL(BIGWA6,0) + ISNULL(BIGWA7,0)
                      + ISNULL(BIGG01,0) + ISNULL(BIGH01,0) + ISNULL(BIGH05,0) + ISNULL(BIGH06,0)
                      + ISNULL(BIGH07,0) + ISNULL(BIGH08,0) + ISNULL(BIGH09,0) + ISNULL(BIGH10,0)
                      + ISNULL(BIGH11,0) + ISNULL(BIGH12,0) + ISNULL(BIGH13,0) + ISNULL(BIGI01,0)
                      + ISNULL(BIGK01,0) + ISNULL(BIGM01,0) + ISNULL(BIGM02,0) + ISNULL(BIGM03,0)
                      + ISNULL(BIGO01,0) + ISNULL(BIGQ01,0) + ISNULL(BIGS01,0) + ISNULL(BIGT01,0)
                      + ISNULL(BIGX01,0) + ISNULL(BIGY01,0) + ISNULL(BIGY02,0) + ISNULL(BIGY03,0)
                      + ISNULL(BIGY20,0) + ISNULL(BIGY21,0) + ISNULL(BIGZ01,0)
                      
-- <3.월별 자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
	SELECT * FROM [#RPY501_1] ORDER BY CLTCOD, MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF
