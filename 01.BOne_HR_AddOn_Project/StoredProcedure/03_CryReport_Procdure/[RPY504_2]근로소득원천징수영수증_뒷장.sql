IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY504_2' AND xtype = 'P'))
	DROP PROCEDURE RPY504_2
GO

CREATE  PROC RPY504_2
	(
		@JSNYER 	AS Nvarchar(4),		--작업연월
		@STRMON 	AS Nvarchar(2), 	--시작월
		@ENDMON 	AS Nvarchar(2), 	--종료월
		@JOBGBN		AS Nvarchar(1), 	--작업구분(1연말정산,2중도정산,3전체)
		@CLTCOD		AS Nvarchar(8), 	--자사코드
		@MSTDPT		AS Nvarchar(8), 	--부서
	    @MSTCOD 	AS Nvarchar(8) 	   	--사원번호		
	)
	
-- Exec RPY504_2 '2013', '01', '12', '3', '%','%', '%', '%'
 AS
	/*==========================================================================================
		프로시저명		: RPY504_2
		프로시저설명	: 근로소득원천징수영수증-뒷장
		만든이			: 함미경
		작업일자		: 2007-01-30
		작업지시자		: 함미경
		작업지시일자	: 2007-01-30
		작업목적		: 
		작업내용		: 
	===========================================================================================*/
	-- DROP PROC RPY504_2
	-- Exec RPY504_2  '2009', '01', '12', '3', N'%',N'%', '%'

	SET NOCOUNT ON

-- <1.임시테이블 생성 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
	CREATE TABLE #RPY504_2 (
			U_CLTCOD	nvarchar(8)  COLLATE Korean_Wansung_Unicode_CI_AS NOT NULL ,
			U_CLTNAM	nvarchar(20) COLLATE Korean_Wansung_Unicode_CI_AS NOT NULL ,
			U_MSTCOD	nvarchar(10) COLLATE Korean_Wansung_Unicode_CI_AS NOT NULL ,
			U_MSTNAM	nvarchar(40) COLLATE Korean_Wansung_Unicode_CI_AS NULL,
			U_MSTDPT	nvarchar(8),
			U_DPTNAM	nvarchar(40),
			U_LINEID	Int,
			U_CHKCOD	nvarchar(1),
			U_CHKINT	Int,
			U_FAMNAM	nvarchar(20),
			U_FAMPER	nvarchar(14),
			U_CHKBAS	nvarchar(1),
			U_CHKJAN	nvarchar(1),
			U_CHKCHL	nvarchar(1),
			U_CHKBUY	nvarchar(1),
			U_CHKJEL	nvarchar(1),
			U_CHKDAJ	nvarchar(1),
			U_CHKCHS	nvarchar(1),
			U_BOHAMT1	Numeric(19,6),
			U_BOHAMT2	Numeric(19,6),
			U_MEDAMT1	Numeric(19,6),
			U_MEDAMT2	Numeric(19,6),
			U_EDCAMT1	Numeric(19,6),
			U_EDCAMT2	Numeric(19,6),
			U_CADAMT1	Numeric(19,6),
			U_CADAMT2	Numeric(19,6),
			U_CSHAMT1	Numeric(19,6),
			U_GBUAMT1	Numeric(19,6),
			U_GBUAMT2	Numeric(19,6),
			U_CSHCAD1	Numeric(19,6),
			U_CSHCAD2	Numeric(19,6)
			) 
	CREATE TABLE #RPY504_3 (
			U_CLTCOD	nvarchar(8)  COLLATE Korean_Wansung_Unicode_CI_AS NOT NULL ,
			U_MSTCOD	nvarchar(10) COLLATE Korean_Wansung_Unicode_CI_AS NOT NULL ,
			U_BOHAMT2	Numeric(19,6)
			) 

-- <2.1정산자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
	--2.1) 본인명세(정산한 사원 모두 생성)
	INSERT INTO [#RPY504_2]	
	SELECT	U_CLTCOD	=	T0.U_CLTCOD, 
			U_CLTNAM	=	(SELECT U_CodeNM FROM [@PS_HR200L] WHERE Code = 'P144' AND U_Code = T0.U_CLTCOD),
			U_MSTCOD	=	T0.Code, 
			U_MSTNAM	=	T0.U_Fullname, 
			U_MSTDPT	=	T0.U_TeamCode,
			U_DPTNAM	=	(select U_CodeNm From [@PS_HR200L] where Code='1' AND U_Code = T0.U_TeamCode),
			U_LINEID	=	1, 
			U_CHKCOD	=	'0', 
			U_CHKINT	=	T0.U_INTGBN,  
			U_FAMNAM	=	T1.U_MSTNAM, 
			U_FAMPER	=	T0.U_govID, 
			U_CHKBAS	=	'Y', 
			U_CHKJAN	=	T0.U_BJNGAE, 
			U_CHKCHL	=	'N', 
			U_CHKBUY	=	T0.U_MZBURI,
			U_CHKJEL	=	'N', 
			U_CHKDAJ	=	'N', 
			U_CHKCHS	=	'N', 
			U_BOHAMT1	=	0, 
--			U_BOHAMT2	=	ISNULL(T6.U_MedAmt,0) + ISNULL(T6.U_GBHAmt,0) + ISNULL(T6.U_NGYAMT,0), 
			U_BOHAMT2	=	0, 
			U_MEDAMT1	=	0, 
			U_MEDAMT2	=	0, 
			U_EDCAMT1	=	0, 
			U_EDCAMT2	=	0, 
			U_CADAMT1	=	0, 
			U_CADAMT2	=	0, 
			U_CSHAMT1	=	0, 
			U_GBUAMT1	=	0, 
			U_GBUAMT2	=	0,
			U_CSHCAD1	=	0,
			U_CSHCAD2	=	0

	FROM	[@PH_PY001A] T0 
			INNER JOIN [@ZPY504H] T1 ON T0.Code = T1.U_MSTCOD AND T1.U_JSNYER = @JSNYER
			--INNER JOIN [OHEM] T2 ON T0.U_EmpID = T2.EmpID
			--INNER JOIN [OUDP] T3 ON T2.Dept = T3.Code
			--INNER JOIN [@ZPY106H] T4 ON T2.U_CLTCOD = T4.CODE
			LEFT  JOIN [@PH_PY005A] T4 ON T0.U_CLTCOD = T4.U_CLTCode
--			INNER JOIN [@ZPY343H] T5 ON T1.U_MSTCOD = T5.U_MstCode AND T1.U_JSNYER = T5.U_JsnYear
--			INNER JOIN [@ZPY343L] T6 ON T5.DocEntry = T6.DocEntry  AND T6.U_LineNum = '13'
	WHERE 	T1.U_JSNYER    = @JSNYER
	AND		T1.U_CLTCOD LIKE @CLTCOD                        
	AND		T0.U_TeamCode LIKE @MSTDPT                        
	AND		T0.Code LIKE @MSTCOD
	AND		(@JOBGBN ='3' OR (@JOBGBN <> '3' AND T1.U_JSNGBN = @JOBGBN))
	AND		T1.U_JSNMON BETWEEN @STRMON AND @ENDMON
/*
	AND		T2.Status LIKE CASE @JOBGBN WHEN '1' THEN '1' 
										WHEN '2' THEN '4'
										ELSE '%' END
*/										
-- <2.2가족명세 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
	DECLARE	@U_OLDMST		 AS Nvarchar(8)
	DECLARE	@U_MSTCOD		 AS Nvarchar(8)
	DECLARE	@U_MSTNAM		 AS Nvarchar(40)
	DECLARE	@U_MSTDPT		 AS Nvarchar(8)
	DECLARE	@U_DPTNAM		 AS Nvarchar(40)
	DECLARE	@U_LineID		 AS Int	

	DECLARE	@U_CHKCOD	nvarchar(1), 	@U_CHKINT Int, 	@U_FAMNAM	nvarchar(20), 	@U_FAMPER	nvarchar(14),
			@U_CHKBAS	nvarchar(1), 	@U_CHKJAN	nvarchar(1), 	@U_CHKCHL	nvarchar(1), 	@U_BOHAMT1	Numeric(19,6),
			@U_BOHAMT2	Numeric(19,6),	@U_MEDAMT1	Numeric(19,6),	@U_MEDAMT2	Numeric(19,6),  @U_EDCAMT1	Numeric(19,6),
			@U_EDCAMT2	Numeric(19,6),	@U_CADAMT1	Numeric(19,6),	@U_CADAMT2	Numeric(19,6),	@U_CSHAMT1	Numeric(19,6),
			@U_GBUAMT1	Numeric(19,6),  @U_GBUAMT2	Numeric(19,6),	@U_CHKBUY	nvarchar(1),	@U_CHKJEL	nvarchar(1),
			@U_CHKDAJ	nvarchar(1),	@U_CHKCHS	nvarchar(1),	@U_CSHCAD1	Numeric(19,6),	@U_CSHCAD2	Numeric(19,6)
			
	SET @U_OLDMST = ''
	DECLARE Cur_MDC CURSOR FOR
	
	SELECT 	U_MSTCOD	=	T0.Code , 
			U_MSTNAM	=	T4.U_FullName, 
			U_MSTDPT	=	T4.U_TeamCode,
			U_DPTNAM	=	(SELECT U_CodeNM FROM [@PS_HR200L] WHERE Code ='1' AND U_Code = T4.U_TeamCode),
			U_CHKCOD	=	T0.U_CHKCOD, 
			U_CHKINT	=	T0.U_CHKINT, 
			U_FAMNAM	=	T0.U_FAMNAM, 
			U_FAMPER	=	T0.U_FAMPER, 
			U_CHKBAS	=	ISNULL(T0.U_CHKBAS, 'N'), 
			U_CHKJAN	=	ISNULL(T0.U_CHKJAN, 'N'),  
			U_CHKCHL	=	ISNULL(T0.U_CHKCHL, 'N'), 
			U_CHKBUY	=	'N', 
			U_CHKJEL	=	'N', 
			U_CHKDAJ	=	'N', 
			U_CHKCHS	=	'N',
			U_BOHAMT1	=	0, 
			U_BOHAMT2	=	0, 
			U_MEDAMT1	=	0, 
			U_MEDAMT2	=	0, 
			U_EDCAMT1	=	0, 
			U_EDCAMT2	=	0, 
			U_CADAMT1	=	0, 
			U_CADAMT2	=	0, 
			U_CSHAMT1	=	0, 
			U_GBUAMT1	=	0, 
			U_GBUAMT2	=	0,
			U_CSHCAD1	=	0,
			U_CSHCAD2	=	0
	FROM 	[@PH_PY001D] T0 
			--INNER JOIN [@ZPY121H]  T1 ON T0.Code = T1.Code
			INNER JOIN [#RPY504_2] T2 ON T0.Code COLLATE Korean_Wansung_Unicode_CI_AS = T2.U_MSTCOD COLLATE Korean_Wansung_Unicode_CI_AS
			--INNER JOIN [OHEM] T4 ON T1.U_MSTCOD = T4.U_MSTCOD
			INNER JOIN [@PH_PY001A] T4 ON T0.Code = T4.Code
			INNER JOIN [OUDP] T3 ON T4.U_TeamCode = T3.Code
	WHERE  	ISNULL(T0.U_Chkbas, 'N') = 'Y'  -- 기본공제대상자만
	AND   	ISNULL(T0.U_ChkCod, '') <> ''  -- 관계코드가 있는사람만
	ORDER 	BY T2.U_MSTCOD, T0.U_CHKCOD, T0.U_LineNum	
	
	OPEN Cur_MDC	
	
	FETCH NEXT FROM Cur_MDC INTO  	@U_MSTCOD, @U_MSTNAM, @U_MSTDPT, @U_DPTNAM, @U_CHKCOD, @U_CHKINT, @U_FAMNAM, @U_FAMPER, @U_CHKBAS, @U_CHKJAN, 
									@U_CHKCHL, @U_CHKBUY, @U_CHKJEL, @U_CHKDAJ, @U_CHKCHS, @U_BOHAMT1,@U_MEDAMT1,@U_EDCAMT1,@U_CADAMT1,@U_GBUAMT1, 
									@U_BOHAMT2,@U_MEDAMT2,@U_EDCAMT2,@U_CADAMT2,@U_GBUAMT2,@U_CSHAMT1,@U_CSHCAD1,@U_CSHCAD2

	WHILE @@FETCH_STATUS = 0	--성공
		BEGIN
			--2.3) 가족명세
			IF (@U_OLDMST <> @U_MSTCOD)
				BEGIN
					SET @U_LineID = 1
				END	

			SET @U_LineID = @U_LineID + 1

			INSERT INTO [#RPY504_2]	
			SELECT 	T0.U_CLTCOD, T1.NAME,   
					@U_MSTCOD,  @U_MSTNAM,  @U_MSTDPT,  @U_DPTNAM,  @U_LineID,  @U_CHKCOD,  @U_CHKINT,
					@U_FAMNAM,  @U_FAMPER, 	@U_CHKBAS,  @U_CHKJAN,  @U_CHKCHL,  @U_CHKBUY,  @U_CHKJEL,
					@U_CHKDAJ,  @U_CHKCHS,  @U_BOHAMT1, @U_BOHAMT2, @U_MEDAMT1, @U_MEDAMT2, @U_EDCAMT1,
					@U_EDCAMT2, @U_CADAMT1, @U_CADAMT2, @U_CSHAMT1, @U_GBUAMT1, @U_GBUAMT2, @U_CSHCAD1, @U_CSHCAD2
			FROM 	[@PH_PY001A] T0 
					INNER JOIN [@PH_PY005A] T1 ON T0.U_CLTCOD = T1.U_CLTCode
			WHERE 	T0.Code = @U_MSTCOD
			
			SET @U_OLDMST = @U_MSTCOD
	--다음 레코드로 이동							
	FETCH NEXT FROM Cur_MDC INTO  	@U_MSTCOD, @U_MSTNAM, @U_MSTDPT, @U_DPTNAM, @U_CHKCOD, @U_CHKINT, @U_FAMNAM, @U_FAMPER, @U_CHKBAS, @U_CHKJAN,
									@U_CHKCHL, @U_CHKBUY, @U_CHKJEL, @U_CHKDAJ, @U_CHKCHS, @U_BOHAMT1,@U_MEDAMT1,@U_EDCAMT1,@U_CADAMT1,@U_GBUAMT1, 
									@U_BOHAMT2,@U_MEDAMT2,@U_EDCAMT2,@U_CADAMT2,@U_GBUAMT2,@U_CSHAMT1,@U_CSHCAD1,@U_CSHCAD2 
	END

	CLOSE Cur_MDC
	DEALLOCATE Cur_MDC	

	--<2.3.대상자 조회>ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
	SELECT T1.U_MSTCOD
	INTO #RPY504_4
  	FROM [#RPY504_2] T1
  	GROUP BY T1.U_MSTCOD
  	
	--<2.4.소득자료에 있는분 제거>ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
	DELETE 	[#RPY504_2] 
	FROM 	[#RPY504_4] T0  
			INNER JOIN [@ZPY501H] T1  ON T0.U_MSTCOD COLLATE Korean_Wansung_Unicode_CI_AS = T1.U_MSTCOD COLLATE Korean_Wansung_Unicode_CI_AS 
                                     AND T1.U_JSNYER COLLATE Korean_Wansung_Unicode_CI_AS = @JSNYER COLLATE Korean_Wansung_Unicode_CI_AS 

	--<2.5.소득자료데이터로 변경 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    	
  	INSERT INTO [#RPY504_2]	
	SELECT	U_CLTCOD	=	T4.U_CLTCOD, 
			U_CLTNAM	=	T5.NAME, 
			U_MSTCOD	=	T1.U_MSTCOD, 
			U_MSTNAM	=	T1.U_MSTNAM, 
			U_MSTDPT	=	T2.U_TeamCode,
			U_DPTNAM	=	(SELECT U_CodeNM FROM [@PS_HR200L] WHERE Code ='1' AND U_Code = T2.U_TeamCode),
			U_LINEID	=	T0.U_LineNum, 
			U_CHKCOD	=	T0.U_CHKCOD, 
			U_CHKINT	=	T0.U_CHKINT, 
			U_FAMNAM	=	T0.U_FAMNAM, 
			U_FAMPER	=	T0.U_FAMPER,
			U_CHKBAS	=	T0.U_CHKBAS, 
			U_CHKJAN	=	T0.U_CHKJAN, 
			U_CHKCHL	=	T0.U_CHKCHL, 
			U_CHKBUY	=	T0.U_CHKBUY, 
			U_CHKJEL	=	T0.U_CHKJEL, 
			U_CHKDAJ	=	T0.U_CHKDAJ, 
			U_CHKCHS	=	ISNULL(T0.U_CHKCHS,0),
			U_BOHAMT1	=	ISNULL(T0.U_BOHAMT1,0), 
			--2011.11.24 HMK 아래주석처리되어있는걸 2011년이후 보험료2만 가져오도록변경함. 2011년 초에 로직추가변경된 내용으로 보임.
			U_BOHAMT2	=	CASE WHEN @JSNYER >= '2011' THEN ISNULL(T0.U_BOHAMT2,0) ELSE ISNULL(T0.U_BOHAMT2,0) + ISNULL(T0.U_BOHAMT3,0) END, 
--			U_BOHAMT2	=	ISNULL(T0.U_BOHAMT2,0) + ISNULL(T0.U_BOHAMT3,0), 
			U_MEDAMT1	=	T0.U_MEDAMT1, 
			U_MEDAMT2	=	T0.U_MEDAMT2, 
			U_EDCAMT1	=	T0.U_EDCAMT1, 
			U_EDCAMT2	=	T0.U_EDCAMT2,
			U_CADAMT1	=	T0.U_CADAMT1, 
			U_CADAMT2	=	T0.U_CADAMT2, 
			U_CSHAMT1	=	T0.U_CSHAMT1, 
			U_GBUAMT1	=	T0.U_GBUAMT1, 
			U_GBUAMT2	=	ISNULL(T0.U_GBUAMT2,0) + ISNULL(T0.U_GBUAMT3,0),
			U_CSHCAD1	=	T0.U_CSHCAD1,
			U_CSHCAD2	=	T0.U_CSHCAD2
			
	FROM 	[@ZPY501L] T0 
			INNER JOIN [@ZPY501H] T1 ON T0.DocEntry = T1.DocEntry
		   	INNER JOIN [@ZPY504H] T4 ON T4.U_JSNYER = T1.U_JSNYER AND T4.U_MSTCOD = T1.U_MSTCOD AND T1.U_CLTCOD = T4.U_CLTCOD
			--INNER JOIN [OHEM] T2 ON T1.U_MSTCOD = T2.U_MSTCOD
			INNER JOIN [@PH_PY001A] T2 ON T1.U_MSTCOD = T2.Code
			--INNER JOIN [OUDP] T3 ON T2.U_TeamCode = T3.Code
			--INNER JOIN [@ZPY106H] T5 ON T4.U_CLTCOD = T5.CODE
			LEFT JOIN [@PH_PY005A] T5 ON T4.U_CLTCOD = T5.U_CLTCode
	WHERE 	T1.U_JSNYER = @JSNYER
	AND		T4.U_CLTCOD LIKE @CLTCOD
	AND		T2.U_TeamCode LIKE @MSTDPT 
	AND		T1.U_MSTCOD LIKE @MSTCOD
	AND		ISNULL(T0.U_FamNam,'') <> ''
	AND		(@JOBGBN ='3' OR (@JOBGBN <> '3' AND T4.U_JSNGBN = @JOBGBN))	
	AND		T4.U_JSNMON BETWEEN @STRMON AND @ENDMON

/*

	AND		T2.Status LIKE CASE @JOBGBN WHEN '1' THEN '1' 
										WHEN '2' THEN '4'
										ELSE '%' END
*/
	UPDATE	T0
	SET		U_CHKBAS = CASE T0.U_CHKBAS WHEN 'Y' THEN '1' WHEN 'N' THEN '0' ELSE T0.U_CHKBAS END,
			U_CHKJAN = CASE T0.U_CHKJAN WHEN 'Y' THEN '1' WHEN 'N' THEN '0' ELSE T0.U_CHKJAN END,
			U_CHKCHL = CASE T0.U_CHKCHL WHEN 'Y' THEN '1' WHEN 'N' THEN '0' ELSE T0.U_CHKCHL END,
			U_CHKBUY = CASE T0.U_CHKBUY WHEN 'Y' THEN '1' WHEN 'N' THEN '0' ELSE T0.U_CHKBUY END,
			U_CHKJEL = CASE T0.U_CHKJEL WHEN 'Y' THEN '1' WHEN 'N' THEN '0' ELSE T0.U_CHKJEL END,
			U_CHKDAJ = CASE T0.U_CHKDAJ WHEN 'Y' THEN '1' WHEN 'N' THEN '0' ELSE T0.U_CHKDAJ END,
			U_CHKCHS = CASE T0.U_CHKCHS WHEN 'Y' THEN '1' WHEN 'N' THEN '0' ELSE T0.U_CHKCHS END
	FROM	[#RPY504_2] T0

	DECLARE @BOHAMT3 AS NUMERIC(19,6)

	INSERT	INTO #RPY504_3
	SELECT	U_CLTCOD  = A0.CLTCOD,
			U_MSTCOD  = A0.MSTCOD,
			U_BOHAMT2 = CASE WHEN ISNULL(SUM(A0.MEDAMT),0) < 0 THEN 0 ELSE ISNULL(SUM(A0.MEDAMT),0) END
                      + CASE WHEN ISNULL(SUM(A0.GBHAMT),0) < 0 THEN 0 ELSE ISNULL(SUM(A0.GBHAMT),0) END
	FROM (
			SELECT	CLTCOD = T0.U_CLTCOD,
					MSTCOD = T0.U_MstCode,
                    MEDAMT = ISNULL(T1.U_MedAmt,0) + ISNULL(T1.U_NGYAMT,0),
                    GBHAMT = ISNULL(T1.U_GBHAmt,0)
			FROM	[@ZPY343H] T0
					INNER JOIN [@ZPY343L] T1 ON T0.DocEntry = T1.DocEntry
			WHERE	T0.U_CLTCOD LIKE @CLTCOD
			AND		T0.U_JsnYear = @JSNYER
			AND		T0.U_MstCode LIKE @MSTCOD
			AND		T1.U_LineNum = '13'

			UNION	ALL

			SELECT	CLTCOD = T0.U_CLTCOD,
					MSTCOD = T0.U_MSTCOD,
                    MEDAMT = ISNULL(T1.U_JONMED,0),
                    GBHAMT = ISNULL(T1.U_JONGBH,0)
			FROM	[@ZPY502H] T0
					INNER JOIN [@ZPY502L] T1 ON T0.DocEntry = T1.DocEntry
			WHERE	T0.U_CLTCOD LIKE @CLTCOD
			AND		T0.U_JSNYER = @JSNYER
			AND		T0.U_MSTCOD LIKE @MSTCOD
	) A0
	GROUP	BY A0.CLTCOD, A0.MSTCOD

	UPDATE	T0
	SET		U_BOHAMT2 = ISNULL(T0.U_BOHAMT2,0) + ISNULL(T1.U_BOHAMT2,0)
	FROM	[#RPY504_2] T0
			INNER JOIN [#RPY504_3] T1 ON T0.U_CLTCOD = T1.U_CLTCOD AND T0.U_MSTCOD = T1.U_MSTCOD
	WHERE	U_CHKCOD = '0'
  
-- <3.정산자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
	SELECT * FROM [#RPY504_2] ORDER BY  U_CLTCOD, U_MSTCOD, U_CHKCOD, U_LINEID



--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF
