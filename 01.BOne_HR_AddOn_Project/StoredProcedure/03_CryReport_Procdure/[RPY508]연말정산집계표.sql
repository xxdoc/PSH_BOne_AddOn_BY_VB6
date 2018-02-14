IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY508' AND xtype = 'P'))
	DROP PROCEDURE RPY508
GO

CREATE PROC RPY508
    (
        @JSNYER     AS Nvarchar(4),     --작업연월
        @JOBGBN     AS Nvarchar(1),     --작업구분(1연말정산,2중도정산,3전체)
        @CLTCOD     AS Nvarchar(8),     --자사코드
        @MSTDPT     AS Nvarchar(8)      --부서

    )
    

 AS
    /*==========================================================================================
        프로시저명      : RPY508
        프로시저설명    : 연말정산집계표
        만든이          : 최동권
        작업일자        : 2008-05-19
        작업지시자      : 함미경
        작업지시일자    : 2009-07-29
        작업목적        : 자사코드추가
        작업내용        : 
    ===========================================================================================*/
    -- DROP PROC RPY508
    -- Exec RPY508  '2007', '3', N'%', N'%'

    SET NOCOUNT ON

-- <1.임시테이블 생성 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    

    CREATE TABLE #RPY508 (
            EMPCNT  NUMERIC(19,6),
            PAYAMT  NUMERIC(19,6),
            BNSAMT  NUMERIC(19,6),
            BTXAM2  NUMERIC(19,6),
            BTXAM1  NUMERIC(19,6),
            PAYAL1  NUMERIC(19,6),
            PAYAL2  NUMERIC(19,6),
            INCOME  NUMERIC(19,6),
            PILGNL  NUMERIC(19,6),
            GNLOSD  NUMERIC(19,6),
            INJBAS  NUMERIC(19,6),
            BAEWOO  NUMERIC(19,6),
            BUYNSU  NUMERIC(19,6),
            GYNGLO  NUMERIC(19,6),
            JANGAE  NUMERIC(19,6),
            MZBURI  NUMERIC(19,6),
            BUYN06  NUMERIC(19,6),
            DAGYSU  NUMERIC(19,6),
            INJBWO  NUMERIC(19,6),
            INJBYN  NUMERIC(19,6),
            INJGYN  NUMERIC(19,6),
            INJJAE  NUMERIC(19,6),
            INJBNJ  NUMERIC(19,6),
            INJSON  NUMERIC(19,6),
            INJADD  NUMERIC(19,6),
            BHMCNT  NUMERIC(19,6),
            MEDCNT  NUMERIC(19,6),
            SCHCNT  NUMERIC(19,6),
            HUSCNT  NUMERIC(19,6),
            GBUCNT  NUMERIC(19,6),
            PILBHM  NUMERIC(19,6),
            PILMED  NUMERIC(19,6),
            PILSCH  NUMERIC(19,6),
            PILHUS  NUMERIC(19,6),
            PILGBU  NUMERIC(19,6),
            PILTOT  NUMERIC(19,6),
            GONCNT  NUMERIC(19,6),
            YUNGON  NUMERIC(19,6),
            CHAGAM  NUMERIC(19,6),
            GYNCNT  NUMERIC(19,6),
            YUNCNT  NUMERIC(19,6),
            INVCNT  NUMERIC(19,6),
            CADCNT  NUMERIC(19,6),
            USJCNT  NUMERIC(19,6),
            GITGYN  NUMERIC(19,6),
            GITYUN  NUMERIC(19,6),
            GITINV  NUMERIC(19,6),
            GITCAD  NUMERIC(19,6),
            GITUSJ  NUMERIC(19,6),
            TAXCNT  NUMERIC(19,6),
            TAXSTD  NUMERIC(19,6),
            SANTAX  NUMERIC(19,6),
            TAXGNL  NUMERIC(19,6),
            BROCNT  NUMERIC(19,6),
            FRGCNT  NUMERIC(19,6),
            NABCNT  NUMERIC(19,6),
            POLCNT  NUMERIC(19,6),
            TAXBRO  NUMERIC(19,6),
            TAXFRG  NUMERIC(19,6),
            TAXNAB  NUMERIC(19,6),
            TAXGBU  NUMERIC(19,6),
            TAXTOT  NUMERIC(19,6),
            GAMSOD  NUMERIC(19,6),
            GAMJOS  NUMERIC(19,6),
            GAMTOT  NUMERIC(19,6),
            GULCNT  NUMERIC(19,6),
            GULGAB  NUMERIC(19,6),
            GULNON  NUMERIC(19,6),
            GULJUM  NUMERIC(19,6),
            NANCNT  NUMERIC(19,6),
            NANGAB  NUMERIC(19,6),
            NANNON  NUMERIC(19,6),
            NANJUM  NUMERIC(19,6),
            NALCNT  NUMERIC(19,6),
            NALGAB  NUMERIC(19,6),
            NALNON  NUMERIC(19,6),
            NALJUM  NUMERIC(19,6),
            JSUCNT  NUMERIC(19,6),
            JSUGAB  NUMERIC(19,6),
            JSUNON  NUMERIC(19,6),
            JSUJUM  NUMERIC(19,6),
            HWACNT  NUMERIC(19,6),
            HWAGAB  NUMERIC(19,6),
            HWANON  NUMERIC(19,6),
            HWAJUM  NUMERIC(19,6),
			CHLSAN	NUMERIC(19,6),
			INJCHL	NUMERIC(19,6),
			KUKCNT	NUMERIC(19,6),
			KUKGON	NUMERIC(19,6),
			RETCNT	NUMERIC(19,6),
			GITRET	NUMERIC(19,6),
			JHECNT	NUMERIC(19,6),
			PILJHE	NUMERIC(19,6),
			HUNCNT	NUMERIC(19,6),	-- 결혼.이사.장례 인원수  (2009년이후 사용안함)
			PILHUN	NUMERIC(19,6),	-- 결혼.이사.장례 공제금액
			SGICNT	NUMERIC(19,6),
			GITSGI	NUMERIC(19,6),
			GHSCNT	NUMERIC(19,6),
			GITHUS	NUMERIC(19,6),
			JFDCNT	NUMERIC(19,6),
			GITJFD	NUMERIC(19,6)
            ) 

-- <2.정산자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
    INSERT  INTO [#RPY508]
    SELECT  EMPCNT  =   COUNT(T0.U_MSTCOD),                                             -- 총인원
            PAYAMT  =   ISNULL(SUM(T0.U_PAYAMT),0),                                     -- 급여총액
            BNSAMT  =   ISNULL(SUM(T0.U_BNSAMT),0) + ISNULL(SUM(T0.U_INBAMT),0)         -- 상여총액
                    +   ISNULL(SUM(T0.U_JUSAMT),0),
            BTXAM2  =   ISNULL(SUM(T0.U_BIGWA1),0) + ISNULL(SUM(T0.U_BIGWA2),0)         -- 비과세계(면제분 포함)
                    +   ISNULL(SUM(T0.U_BIGWA3),0) + ISNULL(SUM(T0.U_BIGWA5),0) 
                    +   ISNULL(SUM(T0.U_BIGWA6),0) + ISNULL(SUM(T0.U_BIGWU3),0) + ISNULL(SUM(T0.U_BIGWA4),0),
            BTXAM1  =   ISNULL(SUM(T0.U_BIGTOT),0),                                     -- 비과세계(면제분 미포함)
            PAYAL1  =   ISNULL(SUM(T0.U_INCOME),0) + ISNULL(SUM(T0.U_BIGWA1),0)         -- 총금액(면제분 포함)
                    +   ISNULL(SUM(T0.U_BIGWA2),0) + ISNULL(SUM(T0.U_BIGWA3),0) 
                    +   ISNULL(SUM(T0.U_BIGWA5),0) + ISNULL(SUM(T0.U_BIGWA6),0)
                    +   ISNULL(SUM(T0.U_BIGWU3),0) + ISNULL(SUM(T0.U_BIGWA4),0),
            PAYAL2  =   ISNULL(SUM(T0.U_INCOME),0) + ISNULL(SUM(T0.U_BIGTOT),0),        -- 총금액(면제분 미포함)
            
            INCOME  =   ISNULL(SUM(T0.U_INCOME),0),                                     -- 근로소득
            PILGNL  =   ISNULL(SUM(T0.U_PILGNL),0),                                     -- 근로소득공제
            GNLOSD  =   ISNULL(SUM(T0.U_GNLOSD),0),                                     -- 근로소득금액
            INJBAS  =   ISNULL(SUM(T0.U_INJBAS),0),                                     -- 본인공제금액
            
            BAEWOO  =   ISNULL(SUM(T0.U_BAEWOO),0),                                     -- 배우자인원
            BUYNSU  =   ISNULL(SUM(T0.U_BUYNSU),0),                                     -- 부양가족인원
            GYNGLO  =   ISNULL(SUM(T0.U_GYNGLO),0),                                     -- 경로우대 인원
            JANGAE  =   ISNULL(SUM(T0.U_JANGAE),0),                                     -- 장애자 인원
            MZBURI  =   ISNULL(SUM(T0.U_MZBURI),0),                                     -- 부녀자 인원
            BUYN06  =   ISNULL(SUM(T0.U_BUYN06),0),                                     -- 6세이하 자녀인원
            DAGYSU  =   ISNULL(SUM(T0.U_DAGYSU),0),                                     -- 다자녀 인원
            
            INJBWO  =   ISNULL(SUM(T0.U_INJBWO),0),                                     -- 배우자공제금액
            INJBYN  =   ISNULL(SUM(T0.U_INJBYN),0),                                     -- 부양가족공제금액
            INJGYN  =   ISNULL(SUM(T0.U_INJGYN),0),                                     -- 경로우대 공제금액
            INJJAE  =   ISNULL(SUM(T0.U_INJJAE),0),                                     -- 장애인 공제금액
            INJBNJ  =   ISNULL(SUM(T0.U_INJBNJ),0),                                     -- 부녀자 공제금액
            INJSON  =   ISNULL(SUM(T0.U_INJSON),0),                                     -- 6세이하 자녀공제 금액
            INJADD  =   ISNULL(SUM(T0.U_INJADD),0),                                     -- 다자녀 공제금액
            
            BHMCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILBHM > 0 THEN 1 ELSE 0 END),0),     -- 보험료 인원
            MEDCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILMED > 0 THEN 1 ELSE 0 END),0),     -- 의료비 인원
            SCHCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILSCH > 0 THEN 1 ELSE 0 END),0),     -- 교육비 인원
            HUSCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILHUS > 0 THEN 1 ELSE 0 END),0),     -- 주택자금 인원
            GBUCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILGBU > 0 THEN 1 ELSE 0 END),0),     -- 기부금 인원
            
            PILBHM  =   ISNULL(SUM(T0.U_PILBHM),0),                                     -- 보험료 공제금액
            PILMED  =   ISNULL(SUM(T0.U_PILMED),0),                                     -- 의료비 공제금액
            PILSCH  =   ISNULL(SUM(T0.U_PILSCH),0),                                     -- 교육비 공제금액
            PILHUS  =   ISNULL(SUM(T0.U_PILHUS),0),                                     -- 주택자금 공제금액
            PILGBU  =   ISNULL(SUM(T0.U_PILGBU),0),                                     -- 기부금 공제금액
            PILTOT  =   ISNULL(SUM(T0.U_PILTOT),0) + ISNULL(SUM(T0.U_PILGON),0),        -- 계 또는 표준공제
            
            GONCNT  =   ISNULL(SUM(CASE WHEN T0.U_YUNGON > 0 THEN 1 ELSE 0 END),0),
            YUNGON  =   ISNULL(SUM(T0.U_KUKGON),0)         -- 연금보험료 공제금액
                    +   ISNULL(SUM(T0.U_GITRET),0),

            CHAGAM  =   ISNULL(SUM(T0.U_CHAGAM),0),                                     -- 차감소득금액
            GYNCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITGYN > 0 THEN 1 ELSE 0 END),0),     -- 개인연금소득공제 인원
            YUNCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITYUN > 0 THEN 1 ELSE 0 END),0),     -- 연금저축소득공제 인원
            INVCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITINV > 0 THEN 1 ELSE 0 END),0),     -- 투자조합소득공제 인원
            CADCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITCAD > 0 THEN 1 ELSE 0 END),0),     -- 신용카드소득공제 인원
            USJCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITUSJ > 0 THEN 1 ELSE 0 END),0),     -- 우리사주조합소득공제 인원
            GITGYN  =   ISNULL(SUM(T0.U_GITGYN),0),                                     -- 개인연금소득공제 금액
            GITYUN  =   ISNULL(SUM(T0.U_GITYUN),0),                                     -- 연금저축소득공제 금액
            GITINV  =   ISNULL(SUM(T0.U_GITINV),0),                                     -- 투자조합소득공제 금액
            GITCAD  =   ISNULL(SUM(T0.U_GITCAD),0),                                     -- 신용카드소득공제 금액
            GITUSJ  =   ISNULL(SUM(T0.U_GITUSJ),0),                                     -- 우리사주조합소득공제 금액
            
            TAXCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXSTD > 0 THEN 1 ELSE 0 END),0),     -- 종합소득과세표준 인원
            TAXSTD  =   ISNULL(SUM(T0.U_TAXSTD),0),                                     -- 종합소득과세표준
            SANTAX  =   ISNULL(SUM(T0.U_SANTAX),0),                                     -- 산출세액

            TAXGNL  =   ISNULL(SUM(T0.U_TAXGNL),0),                                     -- 근로소득세액공제
            BROCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXBRO > 0 THEN 1 ELSE 0 END),0),     -- 주택차입금인원
            FRGCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXFRG > 0 THEN 1 ELSE 0 END),0),     -- 외국납부인원
            NABCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXNAB > 0 THEN 1 ELSE 0 END),0),     -- 납세조합인원
            POLCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXGBU > 0 THEN 1 ELSE 0 END),0),     -- 기부정치자금 인원
            TAXBRO  =   ISNULL(SUM(T0.U_TAXBRO),0),                                     -- 주택차입금 세액공제
            TAXFRG  =   ISNULL(SUM(T0.U_TAXFRG),0),                                     -- 외국납부 세액공제
            TAXNAB  =   ISNULL(SUM(T0.U_TAXNAB),0),                                     -- 납세조합 세액공제
            TAXGBU  =   ISNULL(SUM(T0.U_TAXGBU),0),                                     -- 기부정치자금 세액공제
            TAXTOT  =   ISNULL(SUM(T0.U_TAXTOT),0),                                     -- 세액공제 계
            
            GAMSOD  =   ISNULL(SUM(T0.U_GAMSOD),0),                                     -- 소득세법 세액감면
            GAMJOS  =   ISNULL(SUM(T0.U_GAMJOS),0),                                     -- 조세특례제한법 세액감면
            GAMTOT  =   ISNULL(SUM(T0.U_GAMTOT),0),                                     -- 감면세액 계
            
            GULCNT  =   ISNULL(SUM(CASE WHEN T0.U_GULGAB > 0 THEN 1 ELSE 0 END),0),     -- 결정세액인원
            GULGAB  =   ISNULL(SUM(T0.U_GULGAB),0),                                     -- 결정소득세
            GULNON  =   ISNULL(SUM(T0.U_GULNON),0),                                     -- 결정농특세
            GULJUM  =   ISNULL(SUM(T0.U_GULJUM),0),                                     -- 결정주민세
            
            NANCNT  =   ISNULL(SUM(CASE WHEN T0.U_NANGAB > 0 THEN 1 ELSE 0 END),0),
            NANGAB  =   ISNULL(SUM(T0.U_NANGAB),0),                                     -- 현근무지 기납부 소득세
            NANNON  =   ISNULL(SUM(T0.U_NANNON),0),                                     -- 현근무지 기납부 농특세
            NANJUM  =   ISNULL(SUM(T0.U_NANJUM),0),                                     -- 현근무지 기납부 주민세
            
            NALCNT  =   ISNULL(SUM(CASE WHEN T0.U_JONGAB > 0 THEN 1 ELSE 0 END),0),
            NALGAB  =   ISNULL(SUM(T0.U_JONGAB),0),                                     -- 전근무지 기납부 소득세
            NALNON  =   ISNULL(SUM(T0.U_JONNON),0),                                     -- 전근무지 기납부 농특세
            NALJUM  =   ISNULL(SUM(T0.U_JONJUM),0),                                     -- 전근무지 기납부 주민세
            
            JSUCNT  =   ISNULL(SUM(CASE WHEN T0.U_CHAGAB > 0 THEN 1                ELSE 0 END),0),  -- 징수 세액인원
            JSUGAB  =   ISNULL(SUM(CASE WHEN T0.U_CHAGAB > 0 THEN T0.U_CHAGAB      ELSE 0 END),0),  -- 징수 소득세
            JSUNON  =   ISNULL(SUM(CASE WHEN T0.U_CHANON > 0 THEN T0.U_CHANON      ELSE 0 END),0),  -- 징수 농특세
            JSUJUM  =   ISNULL(SUM(CASE WHEN T0.U_CHAJUM > 0 THEN T0.U_CHAJUM      ELSE 0 END),0),  -- 징수 주민세
            
            HWACNT  =   ISNULL(SUM(CASE WHEN T0.U_CHAGAB < 0 THEN 1                ELSE 0 END),0),  -- 환급 소득세
            HWAGAB  =   ISNULL(SUM(CASE WHEN T0.U_CHAGAB < 0 THEN T0.U_CHAGAB * -1 ELSE 0 END),0),  -- 환급 소득세
            HWANON  =   ISNULL(SUM(CASE WHEN T0.U_CHANON < 0 THEN T0.U_CHANON * -1 ELSE 0 END),0),  -- 환급 농특세
            HWAJUM  =   ISNULL(SUM(CASE WHEN T0.U_CHAJUM < 0 THEN T0.U_CHAJUM * -1 ELSE 0 END),0),   -- 환급 주민세,

            CHLSAN  =   ISNULL(SUM(CASE WHEN T0.U_INJCHL > 0 THEN 1                ELSE 0 END),0),  -- 출산보육대상인원
            INJCHL  =   ISNULL(SUM(T0.U_INJCHL),0),  -- 출산보육공제
            KUKCNT  =   ISNULL(SUM(CASE WHEN T0.U_KUKGON > 0 THEN 1                ELSE 0 END),0),  -- 국민연금공제인원
            KUKGON  =   ISNULL(SUM(T0.U_KUKGON),0),  -- 국민연금 공제
            RETCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITRET > 0 THEN 1                ELSE 0 END),0),  -- 퇴직연금공제인원
            GITRET  =   ISNULL(SUM(T0.U_GITRET),0),  -- 퇴직연금 공제
            JHECNT  =   ISNULL(SUM(CASE WHEN T0.U_PILJHE > 0 THEN 1                ELSE 0 END),0),  -- 장기주택차입인원
            PILJHE  =   ISNULL(SUM(T0.U_PILJHE),0),  -- 장기주택차입공제
            HUNCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILHUN > 0 THEN 1                ELSE 0 END),0),  -- 혼인장례이사(2009년이후 사용안함)
            PILHUN  =   ISNULL(SUM(T0.U_PILHUN),0),  -- 혼인장례이사공제
            SGICNT  =   ISNULL(SUM(CASE WHEN T0.U_GITSGI > 0 THEN 1                ELSE 0 END),0),  -- 소기업공제
            GITSGI  =   ISNULL(SUM(T0.U_GITSGI),0),  -- 소기업공제
            GHSCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITHUS > 0 THEN 1                ELSE 0 END),0),  -- 주택마련저축
            GITHUS  =   ISNULL(SUM(T0.U_GITHUS),0),  -- 주택마련저축공제
            JFDCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITJFD > 0 THEN 1                ELSE 0 END),0),  -- 장기주식형저축
            GITJFD  =   ISNULL(SUM(T0.U_GITJFD),0)  -- 장기주식형저축공제

    FROM    [@ZPY504H]  T0
            --INNER JOIN [OHEM] T1 ON T0.U_MSTCOD = T1.U_MSTCOD
            INNER JOIN [@PH_PY001A] T1 ON T0.U_MstCod = T1.Code
            --INNER JOIN [OUDP] T2 ON T1.Dept     = T2.Code
    WHERE   T0.U_JSNYER     =       @JSNYER
    AND     (T0.U_JSNGBN    =       @JOBGBN
    OR      @JOBGBN         =       '3')
    AND     T0.U_CLTCOD     LIKE    @CLTCOD                        
    AND     T1.U_TeamCode   LIKE    @MSTDPT                        

-- <3.정산자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
    
    SELECT * FROM [#RPY508] 
    
--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF
