IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY508_2009' AND xtype = 'P'))
	DROP PROCEDURE RPY508_2009
GO


CREATE PROC RPY508_2009 (
        @JSNYER     AS Nvarchar(4),     --작업연월
        @JOBGBN     AS Nvarchar(1),     --작업구분(1연말정산,2중도정산,3전체)
        @CLTCOD     AS Nvarchar(8),     --자사코드
        @MSTDPT     AS Nvarchar(8)      --부서
    ) 

 AS
    /*==========================================================================================
        프로시저명      : RPY508_2009
        프로시저설명    : 연말정산집계표
        만든이          : 최동권
        작업일자        : 2008-05-19
        작업지시자      : 함미경
        작업지시일자    : 2009-07-29
        작업목적        : 자사코드추가
        작업내용        : 
    ===========================================================================================*/
    -- DROP PROC RPY508_2009
    -- Exec RPY508_2009 '2013','3','%','%'

    SET NOCOUNT ON

-- <1.임시테이블 생성 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    

    CREATE TABLE #RPY508_2009 (
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
			HU2CNT  NUMERIC(19,6),
            GBUCNT  NUMERIC(19,6),
            PILBHM  NUMERIC(19,6),
            PILMED  NUMERIC(19,6),
            PILSCH  NUMERIC(19,6),
            PILHUS  NUMERIC(19,6),
			PILHU2  NUMERIC(19,6),
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
			JH2CNT	NUMERIC(19,6),
			PILJH2	NUMERIC(19,6),
			JH3CNT	NUMERIC(19,6),
			PILJH3	NUMERIC(19,6),
			HUNCNT	NUMERIC(19,6),
			PILHUN	NUMERIC(19,6),
			SGICNT	NUMERIC(19,6),
			GITSGI	NUMERIC(19,6),
			GHSCNT	NUMERIC(19,6),
			GITHUS	NUMERIC(19,6),
			JFDCNT	NUMERIC(19,6),
			GITJFD	NUMERIC(19,6),
			GYUCNT	NUMERIC(19,6),
			GITGYU	NUMERIC(19,6),
			WOLCNT	NUMERIC(19,6),
			PILWOL	NUMERIC(19,6)
            ) 

-- <2.정산자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
    INSERT  INTO [#RPY508_2009]
    SELECT  EMPCNT  =   COUNT(T0.U_MSTCOD),                                             -- 총인원
            PAYAMT  =   SUM(ISNULL(T0.U_PAYAMT,0)),                                     -- 급여총액
            BNSAMT  =   SUM(ISNULL(T0.U_BNSAMT,0)) + SUM(ISNULL(T0.U_INBAMT,0))         -- 상여총액(상여 + 인정상여 + 스톡옵션 + 우리사주)
                    +   SUM(ISNULL(T0.U_JUSAMT,0)) + SUM(ISNULL(T0.U_URIAMT,0)),
            BTXAM2  =   SUM(ISNULL(T0.U_BIGWA1,0)) + SUM(ISNULL(T0.U_BIGWA2,0))         -- 비과세계(면제분 포함)
                    +   SUM(ISNULL(T0.U_BIGWA3,0)) + SUM(ISNULL(T0.U_BIGWA5,0)) 
                    +   SUM(ISNULL(T0.U_BIGWA6,0)) + SUM(ISNULL(T0.U_BIGWU3,0)) + SUM(ISNULL(T0.U_BIGWA4,0)) 

                    +   SUM(ISNULL(T0.U_BIGG01,0)) + SUM(ISNULL(T0.U_BIGH01,0)) 
                    +   SUM(ISNULL(T0.U_BIGH05,0)) + SUM(ISNULL(T0.U_BIGH06,0)) 
                    +   SUM(ISNULL(T0.U_BIGH07,0)) + SUM(ISNULL(T0.U_BIGH08,0)) 
                    +   SUM(ISNULL(T0.U_BIGH09,0)) + SUM(ISNULL(T0.U_BIGH10,0)) 
                    +   SUM(ISNULL(T0.U_BIGH11,0)) + SUM(ISNULL(T0.U_BIGH12,0)) 
                    +   SUM(ISNULL(T0.U_BIGH13,0)) + SUM(ISNULL(T0.U_BIGI01,0)) 
                    +   SUM(ISNULL(T0.U_BIGK01,0)) + SUM(ISNULL(T0.U_BIGM01,0)) 
                    +   SUM(ISNULL(T0.U_BIGM02,0)) + SUM(ISNULL(T0.U_BIGM03,0)) 
                    +   SUM(ISNULL(T0.U_BIGO01,0)) + SUM(ISNULL(T0.U_BIGQ01,0)) 
                    +   SUM(ISNULL(T0.U_BIGS01,0)) + SUM(ISNULL(T0.U_BIGT01,0)) 
                    +   SUM(ISNULL(T0.U_BIGX01,0)) + SUM(ISNULL(T0.U_BIGY01,0)) 
                    +   SUM(ISNULL(T0.U_BIGY02,0)) + SUM(ISNULL(T0.U_BIGY03,0)) 
                    +   SUM(ISNULL(T0.U_BIGY20,0)) + SUM(ISNULL(T0.U_BIGZ01,0)),
            BTXAM1  =   SUM(ISNULL(T0.U_BIGTOT,0)),                                     -- 비과세계(면제분 미포함)
            PAYAL1  =   SUM(ISNULL(T0.U_INCOME,0)) + SUM(ISNULL(T0.U_BIGWA1,0))         -- 총금액(면제분 포함)
                    +   SUM(ISNULL(T0.U_BIGWA2,0)) + SUM(ISNULL(T0.U_BIGWA3,0)) 
                    +   SUM(ISNULL(T0.U_BIGWA5,0)) + SUM(ISNULL(T0.U_BIGWA6,0))
                    +   SUM(ISNULL(T0.U_BIGWU3,0)) + SUM(ISNULL(T0.U_BIGWA4,0))

                    +   SUM(ISNULL(T0.U_BIGG01,0)) + SUM(ISNULL(T0.U_BIGH01,0)) 
                    +   SUM(ISNULL(T0.U_BIGH05,0)) + SUM(ISNULL(T0.U_BIGH06,0)) 
                    +   SUM(ISNULL(T0.U_BIGH07,0)) + SUM(ISNULL(T0.U_BIGH08,0)) 
                    +   SUM(ISNULL(T0.U_BIGH09,0)) + SUM(ISNULL(T0.U_BIGH10,0)) 
                    +   SUM(ISNULL(T0.U_BIGH11,0)) + SUM(ISNULL(T0.U_BIGH12,0)) 
                    +   SUM(ISNULL(T0.U_BIGH13,0)) + SUM(ISNULL(T0.U_BIGI01,0)) 
                    +   SUM(ISNULL(T0.U_BIGK01,0)) + SUM(ISNULL(T0.U_BIGM01,0)) 
                    +   SUM(ISNULL(T0.U_BIGM02,0)) + SUM(ISNULL(T0.U_BIGM03,0)) 
                    +   SUM(ISNULL(T0.U_BIGO01,0)) + SUM(ISNULL(T0.U_BIGQ01,0)) 
                    +   SUM(ISNULL(T0.U_BIGS01,0)) + SUM(ISNULL(T0.U_BIGT01,0)) 
                    +   SUM(ISNULL(T0.U_BIGX01,0)) + SUM(ISNULL(T0.U_BIGY01,0)) 
                    +   SUM(ISNULL(T0.U_BIGY02,0)) + SUM(ISNULL(T0.U_BIGY03,0)) 
                    +   SUM(ISNULL(T0.U_BIGY20,0)) + SUM(ISNULL(T0.U_BIGZ01,0)),
            PAYAL2  =   SUM(ISNULL(T0.U_INCOME,0)) + SUM(ISNULL(T0.U_BIGTOT,0)),        -- 총금액(면제분 미포함)
            
            INCOME  =   SUM(ISNULL(T0.U_INCOME,0)),                                     -- 근로소득
            PILGNL  =   SUM(ISNULL(T0.U_PILGNL,0)),                                     -- 근로소득공제
            GNLOSD  =   SUM(ISNULL(T0.U_GNLOSD,0)),                                     -- 근로소득금액
            INJBAS  =   SUM(ISNULL(T0.U_INJBAS,0)),                                     -- 본인공제금액
            
            BAEWOO  =   SUM(ISNULL(T0.U_BAEWOO,0)),                                     -- 배우자인원
            BUYNSU  =   SUM(ISNULL(T0.U_BUYNSU,0)),                                     -- 부양가족인원
            GYNGLO  =   SUM(ISNULL(T0.U_GYNGLO,0)),                                     -- 경로우대 인원
            JANGAE  =   SUM(ISNULL(T0.U_JANGAE,0)),                                     -- 장애자 인원
            MZBURI  =   SUM(ISNULL(T0.U_MZBURI,0)),                                     -- 부녀자 인원
            BUYN06  =   SUM(ISNULL(T0.U_BUYN06,0)),                                     -- 6세이하 자녀인원
            DAGYSU  =   SUM(ISNULL(T0.U_DAGYSU,0)),                                     -- 다자녀 인원
            
            INJBWO  =   SUM(ISNULL(T0.U_INJBWO,0)),                                     -- 배우자공제금액
            INJBYN  =   SUM(ISNULL(T0.U_INJBYN,0)),                                     -- 부양가족공제금액
            INJGYN  =   SUM(ISNULL(T0.U_INJGYN,0)),                                     -- 경로우대 공제금액
            INJJAE  =   SUM(ISNULL(T0.U_INJJAE,0)),                                     -- 장애인 공제금액
            INJBNJ  =   SUM(ISNULL(T0.U_INJBNJ,0)),                                     -- 부녀자 공제금액
            INJSON  =   SUM(ISNULL(T0.U_INJSON,0)),                                     -- 6세이하 자녀공제 금액
            INJADD  =   SUM(ISNULL(T0.U_INJADD,0)),                                     -- 다자녀 공제금액
            
            BHMCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILBHM > 0 OR T0.U_PILJHM > 0
                                          OR T0.U_PILMBH > 0 OR T0.U_PILGBH > 0
                                         THEN 1 ELSE 0 END,0)),     -- 보험료 인원
            MEDCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILMED > 0 THEN 1 ELSE 0 END,0)),     -- 의료비 인원
            SCHCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILSCH > 0 THEN 1 ELSE 0 END,0)),     -- 교육비 인원
            HUSCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILHUS > 0 THEN 1 ELSE 0 END,0)),     -- 주택자금 인원
			HU2CNT  =   SUM(ISNULL(CASE WHEN ISNULL(T0.U_PILHU2, 0) > 0 THEN 1 ELSE 0 END,0)),     -- 주택자금 인원
            GBUCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILGBU > 0 THEN 1 ELSE 0 END,0)),     -- 기부금 인원
            
            PILBHM  =   SUM(ISNULL(T0.U_PILBHM,0)) + SUM(ISNULL(T0.U_PILJHM,0)) 
                    +   SUM(ISNULL(T0.U_PILMBH,0)) + SUM(ISNULL(T0.U_PILGBH,0)),        -- 보험료 공제금액
            PILMED  =   SUM(ISNULL(T0.U_PILMED,0)),                                     -- 의료비 공제금액
            PILSCH  =   SUM(ISNULL(T0.U_PILSCH,0)),                                     -- 교육비 공제금액
            PILHUS  =   SUM(ISNULL(T0.U_PILHUS,0)),                                     -- 주택자금 공제금액
			PILHU2  =   SUM(ISNULL(T0.U_PILHU2,0)),                                     -- 주택자금 공제금액
            PILGBU  =   SUM(ISNULL(T0.U_PILGBU,0)),                                     -- 기부금 공제금액
            PILTOT  =   SUM(ISNULL(T0.U_PILTOT,0)) + SUM(ISNULL(T0.U_PILGON,0)),        -- 계 또는 표준공제
            
            GONCNT  =   SUM(ISNULL(CASE WHEN T0.U_YUNGON > 0 OR T0.U_YUNGO1 > 0 
                                          OR T0.U_YUNGO2 > 0 OR T0.U_YUNGO3 > 0 THEN 1 ELSE 0 END,0)),
            YUNGON  =   SUM(ISNULL(T0.U_YUNGON,0)) + SUM(ISNULL(T0.U_YUNGO1,0)) 
                    +   SUM(ISNULL(T0.U_YUNGO2,0)) + SUM(ISNULL(T0.U_YUNGO3,0)),          -- 연금보험료 공제금액

            CHAGAM  =   SUM(ISNULL(T0.U_CHAGAM,0)),                                     -- 차감소득금액
            GYNCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITGYN > 0 THEN 1 ELSE 0 END,0)),     -- 개인연금소득공제 인원
            YUNCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITYUN > 0 THEN 1 ELSE 0 END,0)),     -- 연금저축소득공제 인원
            INVCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITINV > 0 THEN 1 ELSE 0 END,0)),     -- 투자조합소득공제 인원
            CADCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITCAD > 0 THEN 1 ELSE 0 END,0)),     -- 신용카드소득공제 인원
            USJCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITUSJ > 0 THEN 1 ELSE 0 END,0)),     -- 우리사주조합소득공제 인원
            GITGYN  =   SUM(ISNULL(T0.U_GITGYN,0)),                                     -- 개인연금소득공제 금액
            GITYUN  =   SUM(ISNULL(T0.U_GITYUN,0)),                                     -- 연금저축소득공제 금액
            GITINV  =   SUM(ISNULL(T0.U_GITINV,0)),                                     -- 투자조합소득공제 금액
            GITCAD  =   SUM(ISNULL(T0.U_GITCAD,0)),                                     -- 신용카드소득공제 금액
            GITUSJ  =   SUM(ISNULL(T0.U_GITUSJ,0)),                                     -- 우리사주조합소득공제 금액
            
            TAXCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXSTD > 0 THEN 1 ELSE 0 END,0)),     -- 종합소득과세표준 인원
            TAXSTD  =   SUM(ISNULL(T0.U_TAXSTD,0)),                                     -- 종합소득과세표준
            SANTAX  =   SUM(ISNULL(T0.U_SANTAX,0)),                                     -- 산출세액

            TAXGNL  =   SUM(ISNULL(T0.U_TAXGNL,0)),                                     -- 근로소득세액공제
            BROCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXBRO > 0 THEN 1 ELSE 0 END,0)),     -- 주택차입금인원
            FRGCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXFRG > 0 THEN 1 ELSE 0 END,0)),     -- 외국납부인원
            NABCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXNAB > 0 THEN 1 ELSE 0 END,0)),     -- 납세조합인원
            POLCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXGBU > 0 THEN 1 ELSE 0 END,0)),     -- 기부정치자금 인원
            TAXBRO  =   SUM(ISNULL(T0.U_TAXBRO,0)),                                     -- 주택차입금 세액공제
            TAXFRG  =   SUM(ISNULL(T0.U_TAXFRG,0)),                                     -- 외국납부 세액공제
            TAXNAB  =   SUM(ISNULL(T0.U_TAXNAB,0)),                                     -- 납세조합 세액공제
            TAXGBU  =   SUM(ISNULL(T0.U_TAXGBU,0)),                                     -- 기부정치자금 세액공제
            TAXTOT  =   SUM(ISNULL(T0.U_TAXTOT,0)),                                     -- 세액공제 계
            
            GAMSOD  =   SUM(ISNULL(T0.U_GAMSOD,0)),                                     -- 소득세법 세액감면
            GAMJOS  =   SUM(ISNULL(T0.U_GAMJOS,0)),                                     -- 조세특례제한법 세액감면
            GAMTOT  =   SUM(ISNULL(T0.U_GAMTOT,0)),                                     -- 감면세액 계
            
            GULCNT  =   SUM(ISNULL(CASE WHEN T0.U_GULGAB > 0 THEN 1 ELSE 0 END,0)),     -- 결정세액인원
            GULGAB  =   SUM(ISNULL(T0.U_GULGAB,0)),                                     -- 결정소득세
            GULNON  =   SUM(ISNULL(T0.U_GULNON,0)),                                     -- 결정농특세
            GULJUM  =   SUM(ISNULL(T0.U_GULJUM,0)),                                     -- 결정주민세
            
            NANCNT  =   SUM(ISNULL(CASE WHEN T0.U_NANGAB > 0 THEN 1 ELSE 0 END,0)),
            NANGAB  =   SUM(ISNULL(T0.U_NANGAB,0)),                                     -- 현근무지 기납부 소득세
            NANNON  =   SUM(ISNULL(T0.U_NANNON,0)),                                     -- 현근무지 기납부 농특세
            NANJUM  =   SUM(ISNULL(T0.U_NANJUM,0)),                                     -- 현근무지 기납부 주민세
            
            NALCNT  =   SUM(ISNULL(CASE WHEN T0.U_JONGAB > 0 THEN 1 ELSE 0 END,0)),
            NALGAB  =   SUM(ISNULL(T0.U_JONGAB,0)),                                     -- 전근무지 기납부 소득세
            NALNON  =   SUM(ISNULL(T0.U_JONNON,0)),                                     -- 전근무지 기납부 농특세
            NALJUM  =   SUM(ISNULL(T0.U_JONJUM,0)),                                     -- 전근무지 기납부 주민세
            
            JSUCNT  =   SUM(ISNULL(CASE WHEN T0.U_CHAGAB > 0 THEN 1                ELSE 0 END,0)),  -- 징수 세액인원
            JSUGAB  =   SUM(ISNULL(CASE WHEN T0.U_CHAGAB > 0 THEN T0.U_CHAGAB      ELSE 0 END,0)),  -- 징수 소득세
            JSUNON  =   SUM(ISNULL(CASE WHEN T0.U_CHANON > 0 THEN T0.U_CHANON      ELSE 0 END,0)),  -- 징수 농특세
            JSUJUM  =   SUM(ISNULL(CASE WHEN T0.U_CHAJUM > 0 THEN T0.U_CHAJUM      ELSE 0 END,0)),  -- 징수 주민세
            
            HWACNT  =   SUM(ISNULL(CASE WHEN T0.U_CHAGAB < 0 THEN 1                ELSE 0 END,0)),  -- 환급 소득세
            HWAGAB  =   SUM(ISNULL(CASE WHEN T0.U_CHAGAB < 0 THEN T0.U_CHAGAB * -1 ELSE 0 END,0)),  -- 환급 소득세
            HWANON  =   SUM(ISNULL(CASE WHEN T0.U_CHANON < 0 THEN T0.U_CHANON * -1 ELSE 0 END,0)),  -- 환급 농특세
            HWAJUM  =   SUM(ISNULL(CASE WHEN T0.U_CHAJUM < 0 THEN T0.U_CHAJUM * -1 ELSE 0 END,0)),   -- 환급 주민세,

            CHLSAN  =   SUM(ISNULL(CASE WHEN T0.U_INJCHL > 0 THEN 1                ELSE 0 END,0)),  -- 출산보육대상인원
            INJCHL  =   SUM(ISNULL(T0.U_INJCHL,0)),                                                 -- 출산보육공제
            KUKCNT  =   SUM(ISNULL(CASE WHEN T0.U_KUKGON > 0 THEN 1                ELSE 0 END,0)),  -- 국민연금공제인원
            KUKGON  =   SUM(ISNULL(T0.U_KUKGON,0)),                                                 -- 국민연금 공제
            RETCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITRET > 0 OR T0.U_GITRE2 > 0 THEN 1 ELSE 0 END,0)),  -- 퇴직연금공제인원
            GITRET  =   SUM(ISNULL(T0.U_GITRET,0)) + SUM(ISNULL(T0.U_GITRE2,0)),                    -- 퇴직연금 공제
            JHECNT  =   SUM(ISNULL(CASE WHEN T0.U_PILJHE > 0 THEN 1                ELSE 0 END,0)),  -- 장기주택차입인원
            PILJHE  =   SUM(ISNULL(T0.U_PILJHE,0)),                                                 -- 장기주택차입공제
            JH2CNT  =   SUM(ISNULL(CASE WHEN T0.U_PILJH2 > 0 THEN 1                ELSE 0 END,0)),  -- 장기주택차입인원
            PILJH2  =   SUM(ISNULL(T0.U_PILJH2,0)),                                                 -- 장기주택차입공제
            JH3CNT  =   SUM(ISNULL(CASE WHEN T0.U_PILJH3 > 0 THEN 1                ELSE 0 END,0)),  -- 장기주택차입인원
            PILJH3  =   SUM(ISNULL(T0.U_PILJH3,0)),                                                 -- 장기주택차입공제
            HUNCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILHUN > 0 THEN 1                ELSE 0 END,0)),  -- 혼인장례이사
            PILHUN  =   SUM(ISNULL(T0.U_PILHUN,0)),                                                 -- 혼인장례이사공제
            SGICNT  =   SUM(ISNULL(CASE WHEN T0.U_GITSGI > 0 THEN 1                ELSE 0 END,0)),  -- 소기업공제
            GITSGI  =   SUM(ISNULL(T0.U_GITSGI,0)),                                                 -- 소기업공제
            GHSCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITHUS > 0 OR T0.U_GITHU1 > 0 
                                          OR T0.U_GITHU2 > 0 OR T0.U_GITHU3 > 0 THEN 1 ELSE 0 END,0)),  -- 주택마련저축
            GITHUS  =   SUM(ISNULL(T0.U_GITHUS,0)) + SUM(ISNULL(T0.U_GITHU1,0)) 
                    +   SUM(ISNULL(T0.U_GITHU2,0)) + SUM(ISNULL(T0.U_GITHU3,0)),                                                 -- 주택마련저축공제
            JFDCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITJFD > 0 THEN 1                ELSE 0 END,0)),  -- 장기주식형저축
            GITJFD  =   SUM(ISNULL(T0.U_GITJFD,0)),                                                 -- 장기주식형저축공제
            GYUCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITGYU > 0 THEN 1                ELSE 0 END,0)),  -- 고용유지중소기업
            GITGYU  =   SUM(ISNULL(T0.U_GITGYU,0)),                                                  -- 고용유지중소기업소득공제
			WOLCNT	=	SUM(ISNULL(CASE WHEN T0.U_PILWOL > 0 THEN 1				   ELSE 0 END,0)),	-- 월세액
			PILWOL	=	SUM(ISNULL(T0.U_PILWOL,0))													-- 월세액공제

    FROM    [@ZPY504H]  T0
            INNER JOIN [@PH_PY001A] T1 ON T0.U_MstCod = T1.Code
            --INNER JOIN [OUDP] T2 ON T1.Dept     = T2.Code
    WHERE   T0.U_JSNYER     =       @JSNYER
    AND     (T0.U_JSNGBN    =       @JOBGBN
    OR      @JOBGBN         =       '3')
    AND     T0.U_CLTCOD     LIKE    @CLTCOD         
    AND     T1.U_TeamCode     LIKE    @MSTDPT                        

-- <3.정산자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
    
    SELECT * FROM [#RPY508_2009] 
    
--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF
