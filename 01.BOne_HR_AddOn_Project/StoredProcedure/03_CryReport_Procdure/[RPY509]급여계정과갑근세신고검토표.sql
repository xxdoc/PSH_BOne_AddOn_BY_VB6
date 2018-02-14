IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY509' AND xtype = 'P'))
	DROP PROCEDURE RPY509
GO

CREATE PROC RPY509 (
        @JSNYER     AS Nvarchar(4),     --작업연월
        @CLTCOD     AS Nvarchar(8),     --자사코드
        @MSTDPT     AS Nvarchar(8),     --부서
        @MSTCOD     AS Nvarchar(8)      --사원번호
    ) 

 AS
    /*==========================================================================================
        프로시저명      : RPY509
        프로시저설명    : 연말정산집계표
        만든이          : 최동권
        작업일자        : 2008-05-19
        작업지시자      : 함미경
        작업지시일자    : 2009-07-29
        작업목적        : 
        작업내용        : 자사코드추가
    ===========================================================================================*/
    -- DROP PROC RPY509
    -- Exec RPY509  '2009', N'%', N'%', N'%', N'%'

    SET NOCOUNT ON

---------------------------------------------------------------------------------------------------
-- 1.임시테이블 생성 
---------------------------------------------------------------------------------------------------
    CREATE TABLE #RPY509 (
            U_LINENO    NVARCHAR(2),
            U_CNTHLD    NUMERIC(19,6),
            U_PAYHLD    NUMERIC(19,6),
            U_BNSHLD    NUMERIC(19,6),
            U_BTXHLD    NUMERIC(19,6),
            U_JIGHLD    NUMERIC(19,6),
            U_GABHLD    NUMERIC(19,6),
            U_JUMHLD    NUMERIC(19,6),
            U_MEDHLD    NUMERIC(19,6),
            U_GBHHLD    NUMERIC(19,6),
            U_KUKHLD    NUMERIC(19,6),
            U_TOTHLD    NUMERIC(19,6),
            U_CNTRET    NUMERIC(19,6),
            U_PAYRET    NUMERIC(19,6),
            U_BNSRET    NUMERIC(19,6),
            U_BTXRET    NUMERIC(19,6),
            U_JIGRET    NUMERIC(19,6),
            U_GABRET    NUMERIC(19,6),
            U_JUMRET    NUMERIC(19,6),
            U_MEDRET    NUMERIC(19,6),
            U_GBHRET    NUMERIC(19,6),
            U_KUKRET    NUMERIC(19,6),
            U_TOTRET    NUMERIC(19,6),
            U_EMPCNT    NUMERIC(19,6),
            U_GWAPAY    NUMERIC(19,6),
            U_GWABNS    NUMERIC(19,6),
            U_BTXAMT    NUMERIC(19,6),
            U_JIGTOT    NUMERIC(19,6),
            U_GABGUN    NUMERIC(19,6),
            U_JUMINN    NUMERIC(19,6),
            U_MEDAMT    NUMERIC(19,6),
            U_GBHAMT    NUMERIC(19,6),
            U_KUKAMT    NUMERIC(19,6),
            U_TOTGON    NUMERIC(19,6),
            
            U_EMPCN1    NUMERIC(19,6),
            U_GWAPA1    NUMERIC(19,6),
            U_GWABN1    NUMERIC(19,6),
            U_BTXAM1    NUMERIC(19,6),
            U_JIGTO1    NUMERIC(19,6),
            U_GABGU1    NUMERIC(19,6),
            U_JUMIN1    NUMERIC(19,6),
            U_MEDAM1    NUMERIC(19,6),
            U_GBHAM1    NUMERIC(19,6),
            U_KUKAM1    NUMERIC(19,6),
            U_TOTGO1    NUMERIC(19,6),
            U_EMPCN2    NUMERIC(19,6),
            U_GWAPA2    NUMERIC(19,6),
            U_GWABN2    NUMERIC(19,6),
            U_BTXAM2    NUMERIC(19,6),
            U_JIGTO2    NUMERIC(19,6),
            U_GABGU2    NUMERIC(19,6),
            U_JUMIN2    NUMERIC(19,6),
            U_MEDAM2    NUMERIC(19,6),
            U_GBHAM2    NUMERIC(19,6),
            U_KUKAM2    NUMERIC(19,6),
            U_TOTGO2    NUMERIC(19,6)
            ) 

---------------------------------------------------------------------------------------------------
-- 2.월별 내역 조회
---------------------------------------------------------------------------------------------------
	-- 2.1) 월별 소득자료 조회 
    SELECT  U_LINENO    =   T0.U_LINENUM,
            U_MSTCOD    =   T1.U_MSTCODE,
            U_GWAPAY    =   T0.U_GWAPAY,
            U_GWABNS    =   ISNULL(T0.U_GWABNS ,0) + ISNULL(T0.U_INJBNS ,0) + ISNULL(T0.U_JUSBNS ,0) + ISNULL(T0.U_URIBNS ,0),
            U_BTXAMT    =   ISNULL(T0.U_BIGWA01,0) + ISNULL(T0.U_BIGWA02,0) + ISNULL(T0.U_BIGWA03,0) + ISNULL(T0.U_BIGWA04,0)
                          + ISNULL(T0.U_BIGWA05,0) + ISNULL(T0.U_BIGWA06,0) + ISNULL(T0.U_BIGWU03,0) + ISNULL(T0.U_BIGWA07,0)
                          + ISNULL(T0.U_BIGG01 ,0) + ISNULL(T0.U_BIGH01 ,0) + ISNULL(T0.U_BIGH05 ,0) + ISNULL(T0.U_BIGH06 ,0)
                          + ISNULL(T0.U_BIGH07 ,0) + ISNULL(T0.U_BIGH08 ,0) + ISNULL(T0.U_BIGH09 ,0) + ISNULL(T0.U_BIGH10 ,0)
                          + ISNULL(T0.U_BIGH11 ,0) + ISNULL(T0.U_BIGH12 ,0) + ISNULL(T0.U_BIGH13 ,0) + ISNULL(T0.U_BIGI01 ,0)
                          + ISNULL(T0.U_BIGK01 ,0) + ISNULL(T0.U_BIGM01 ,0) + ISNULL(T0.U_BIGM02 ,0) + ISNULL(T0.U_BIGM03 ,0)
                          + ISNULL(T0.U_BIGO01 ,0) + ISNULL(T0.U_BIGQ01 ,0) + ISNULL(T0.U_BIGS01 ,0) + ISNULL(T0.U_BIGT01 ,0)
                          + ISNULL(T0.U_BIGX01 ,0) + ISNULL(T0.U_BIGY01 ,0) + ISNULL(T0.U_BIGY02 ,0) + ISNULL(T0.U_BIGY03 ,0)
                          + ISNULL(T0.U_BIGY20 ,0) + ISNULL(T0.U_BIGY21 ,0) + ISNULL(T0.U_BIGZ01 ,0),
            U_JIGTOT    =   T0.U_JIGTOTAL,
            U_GABGUN    =   ISNULL(T0.U_GABGUN,0),
            U_JUMINN    =   ISNULL(T0.U_JUMIN ,0),
            U_MEDAMT    =   ISNULL(T0.U_MEDAMT,0) + ISNULL(T0.U_NGYAMT ,0),
            U_GBHAMT    =   ISNULL(T0.U_GBHAMT,0),
            U_KUKAMT    =   ISNULL(T0.U_KUKAMT,0),
            U_TOTGON    =   ISNULL(T0.U_KUKAMT,0) + ISNULL(T0.U_GBHAMT,0) + ISNULL(T0.U_MEDAMT,0) + ISNULL(T0.U_NGYAMT,0)
                          + ISNULL(T0.U_GBUAMT,0) + ISNULL(T0.U_JUMIN ,0) + ISNULL(T0.U_GABGUN,0) + ISNULL(T0.U_NONTK ,0),
            U_RETRMM    =   CASE WHEN T2.U_TERMDATE IS NULL OR CONVERT(NVARCHAR(4),T2.U_TERMDATE,112) > @JSNYER 
                                 THEN '00'
                                 WHEN CONVERT(NVARCHAR(4),T2.U_TERMDATE,112) = @JSNYER 
                                 THEN SUBSTRING(CONVERT(NVARCHAR(10),T2.U_TERMDATE,112),5,2)
                                 ELSE '01' END,
            U_JIGDAT	=  ISNULL(T0.U_JIGDATE, '')                     
    INTO    [#RPY509_1]         
    FROM    [@ZPY343L] T0
            INNER JOIN [@ZPY343H] T1 ON T0.DOCENTRY  = T1.DOCENTRY
            INNER JOIN [@PH_PY001A] T2 ON T1.U_MSTCODE = T2.Code
            --INNER JOIN [OUDP]     T3 ON T2.Dept      = T3.Code
    WHERE   T1.U_JSNYEAR                   =       @JSNYER
    AND     T0.U_LINENUM                   <       '13'
    AND     T0.U_JIGDATE                   IS NOT  NULL
    AND     T1.U_CLTCOD                    LIKE    @CLTCOD                        
    AND     T2.U_TeamCode                    LIKE    @MSTDPT                        
    AND     T2.Code							LIKE    @MSTCOD

    -- 2.2) 재직자/중퇴자 분리
    SELECT  U_LINENO    =   T0.U_LINENO,
            U_CNTHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN 1           ELSE 0 END,  --재직자
            U_PAYHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_GWAPAY ELSE 0 END,
            U_BNSHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_GWABNS ELSE 0 END,
            U_BTXHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_BTXAMT ELSE 0 END,
            U_JIGHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_JIGTOT ELSE 0 END,
            U_GABHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_GABGUN ELSE 0 END,
            U_JUMHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_JUMINN ELSE 0 END,
            U_MEDHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_MEDAMT ELSE 0 END,
            U_GBHHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_GBHAMT ELSE 0 END,
            U_KUKHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_KUKAMT ELSE 0 END,
            U_TOTHLD    =   CASE WHEN T0.U_LINENO <> T0.U_RETRMM AND T0.U_JIGDAT <> '' THEN T0.U_TOTGON ELSE 0 END,

            U_CNTRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN 1           ELSE 0 END,  --중퇴자
            U_PAYRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_GWAPAY ELSE 0 END,
            U_BNSRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_GWABNS ELSE 0 END,
            U_BTXRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_BTXAMT ELSE 0 END,
            U_JIGRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_JIGTOT ELSE 0 END,
            U_GABRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_GABGUN ELSE 0 END,
            U_JUMRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_JUMINN ELSE 0 END,
            U_MEDRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_MEDAMT ELSE 0 END,
            U_GBHRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_GBHAMT ELSE 0 END,
            U_KUKRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_KUKAMT ELSE 0 END,
            U_TOTRET    =   CASE WHEN T0.U_LINENO =  T0.U_RETRMM THEN T0.U_TOTGON ELSE 0 END,

            U_EMPCNT    =   CASE WHEN T0.U_JIGDAT <> '' THEN 1 ELSE 0 END,
            U_GWAPAY    =   T0.U_GWAPAY,
            U_GWABNS    =   T0.U_GWABNS,
            U_BTXAMT    =   T0.U_BTXAMT,
            U_JIGTOT    =   T0.U_JIGTOT,
            U_GABGUN    =   T0.U_GABGUN,
            U_JUMINN    =   T0.U_JUMINN,
            U_MEDAMT    =   T0.U_MEDAMT,
            U_GBHAMT    =   T0.U_GBHAMT,
            U_KUKAMT    =   T0.U_KUKAMT,
            U_TOTGON    =   T0.U_TOTGON
    INTO    [#RPY509_2]         
    FROM    [#RPY509_1] T0      

    --3.3) 월별내역 SUM
    SELECT  U_LINENO    =   T0.U_LINENO,
            U_CNTHLD    =   SUM(T0.U_CNTHLD),
            U_PAYHLD    =   SUM(T0.U_PAYHLD),
            U_BNSHLD    =   SUM(T0.U_BNSHLD),
            U_BTXHLD    =   SUM(T0.U_BTXHLD),
            U_JIGHLD    =   SUM(T0.U_JIGHLD),
            U_GABHLD    =   SUM(T0.U_GABHLD),
            U_JUMHLD    =   SUM(T0.U_JUMHLD),
            U_MEDHLD    =   SUM(T0.U_MEDHLD),
            U_GBHHLD    =   SUM(T0.U_GBHHLD),
            U_KUKHLD    =   SUM(T0.U_KUKHLD),
            U_TOTHLD    =   SUM(T0.U_TOTHLD),
            U_CNTRET    =   SUM(T0.U_CNTRET),
            U_PAYRET    =   SUM(T0.U_PAYRET),
            U_BNSRET    =   SUM(T0.U_BNSRET),
            U_BTXRET    =   SUM(T0.U_BTXRET),
            U_JIGRET    =   SUM(T0.U_JIGRET),
            U_GABRET    =   SUM(T0.U_GABRET),
            U_JUMRET    =   SUM(T0.U_JUMRET),
            U_MEDRET    =   SUM(T0.U_MEDRET),
            U_GBHRET    =   SUM(T0.U_GBHRET),
            U_KUKRET    =   SUM(T0.U_KUKRET),
            U_TOTRET    =   SUM(T0.U_TOTRET),
            U_EMPCNT    =   SUM(T0.U_EMPCNT),
            U_GWAPAY    =   SUM(T0.U_GWAPAY),
            U_GWABNS    =   SUM(T0.U_GWABNS),
            U_BTXAMT    =   SUM(T0.U_BTXAMT),
            U_JIGTOT    =   SUM(T0.U_JIGTOT),
            U_GABGUN    =   SUM(T0.U_GABGUN),
            U_JUMINN    =   SUM(T0.U_JUMINN),
            U_MEDAMT    =   SUM(T0.U_MEDAMT),
            U_GBHAMT    =   SUM(T0.U_GBHAMT),
            U_KUKAMT    =   SUM(T0.U_KUKAMT),
            U_TOTGON    =   SUM(T0.U_TOTGON)
            
    INTO    [#RPY509_6]
    FROM    [#RPY509_2] T0
    GROUP   BY  T0.U_LINENO

---------------------------------------------------------------------------------------------------
-- 3. 전근무지 내역 조회
---------------------------------------------------------------------------------------------------
    -- 3.1) 전근무지 내역
    SELECT  U_LINENO    =   T0.U_LINENUM,
            U_MSTCOD    =   T1.U_MSTCOD,
            U_GWAPAY    =   ISNULL(T0.U_JONPAY,0),
            U_GWABNS    =   ISNULL(T0.U_JONBNS,0) + ISNULL(T0.U_INJBNS,0) + ISNULL(T0.U_JONJUS,0) + ISNULL(T0.U_URIBNS,0),
            U_BTXAMT    =   ISNULL(T0.U_JBTTOT,0),
            U_JIGTOT    =   ISNULL(T0.U_JONPAY,0) + ISNULL(T0.U_JONBNS,0) + ISNULL(T0.U_INJBNS,0) + ISNULL(T0.U_JONJUS,0) + ISNULL(T0.U_URIBNS,0),
            U_GABGUN    =   ISNULL(T0.U_JONGAB,0),
            U_JUMINN    =   ISNULL(T0.U_JONJUM,0),
            U_MEDAMT    =   ISNULL(T0.U_JONMED,0),
            U_GBHAMT    =   ISNULL(T0.U_JONGBH,0),
            U_KUKAMT    =   ISNULL(T0.U_JONKUK,0),
            U_TOTGON    =   ISNULL(T0.U_JONKUK,0) + ISNULL(T0.U_JONKUK,0) + ISNULL(T0.U_JONGBH,0) + ISNULL(T0.U_JONMED,0) 
                          + ISNULL(T0.U_JONJUM,0) + ISNULL(T0.U_JONGAB,0) + ISNULL(T0.U_JONKUE,0)
    INTO    [#RPY509_3] 
    FROM    [@ZPY502L] T0
            INNER JOIN [@ZPY502H] T1 ON T0.DOCENTRY = T1.DOCENTRY
            INNER JOIN [@PH_PY001A] T2 ON T1.U_MstCod = T2.Code          
    WHERE   T1.U_JSNYER = @JSNYER
    AND     T1.U_CLTCOD       LIKE    @CLTCOD         
    AND     T2.U_TeamCode     LIKE    @MSTDPT                        
    AND     T2.Code     LIKE    @MSTCOD
    
    -- 3.2) 전근무지를 개인별 항번으로 분리 
    SELECT  U_EMPCN1    =   CASE WHEN T0.U_LINENO = '1' THEN 1        ELSE 0 END,
            U_GWAPA1    =   CASE WHEN T0.U_LINENO = '1' THEN U_GWAPAY ELSE 0 END,
            U_GWABN1    =   CASE WHEN T0.U_LINENO = '1' THEN U_GWABNS ELSE 0 END,
            U_BTXAM1    =   CASE WHEN T0.U_LINENO = '1' THEN U_BTXAMT ELSE 0 END,
            U_JIGTO1    =   CASE WHEN T0.U_LINENO = '1' THEN U_JIGTOT ELSE 0 END,
            U_GABGU1    =   CASE WHEN T0.U_LINENO = '1' THEN U_GABGUN ELSE 0 END,
            U_JUMIN1    =   CASE WHEN T0.U_LINENO = '1' THEN U_JUMINN ELSE 0 END,
            U_MEDAM1    =   CASE WHEN T0.U_LINENO = '1' THEN U_MEDAMT ELSE 0 END,
            U_GBHAM1    =   CASE WHEN T0.U_LINENO = '1' THEN U_GBHAMT ELSE 0 END,
            U_KUKAM1    =   CASE WHEN T0.U_LINENO = '1' THEN U_KUKAMT ELSE 0 END,
            U_TOTGO1    =   CASE WHEN T0.U_LINENO = '1' THEN U_TOTGON ELSE 0 END,
            
            U_EMPCN2    =   CASE WHEN T0.U_LINENO = '2' THEN 1        ELSE 0 END,
            U_GWAPA2    =   CASE WHEN T0.U_LINENO = '2' THEN U_GWAPAY ELSE 0 END,
            U_GWABN2    =   CASE WHEN T0.U_LINENO = '2' THEN U_GWABNS ELSE 0 END,
            U_BTXAM2    =   CASE WHEN T0.U_LINENO = '2' THEN U_BTXAMT ELSE 0 END,
            U_JIGTO2    =   CASE WHEN T0.U_LINENO = '2' THEN U_JIGTOT ELSE 0 END,
            U_GABGU2    =   CASE WHEN T0.U_LINENO = '2' THEN U_GABGUN ELSE 0 END,
            U_JUMIN2    =   CASE WHEN T0.U_LINENO = '2' THEN U_JUMINN ELSE 0 END,
            U_MEDAM2    =   CASE WHEN T0.U_LINENO = '2' THEN U_MEDAMT ELSE 0 END,
            U_GBHAM2    =   CASE WHEN T0.U_LINENO = '2' THEN U_GBHAMT ELSE 0 END,
            U_KUKAM2    =   CASE WHEN T0.U_LINENO = '2' THEN U_KUKAMT ELSE 0 END,
            U_TOTGO2    =   CASE WHEN T0.U_LINENO = '2' THEN U_TOTGON ELSE 0 END
    INTO    [#RPY509_4] 
    FROM    [#RPY509_3] T0

    -- 3.3) 전근무지 합계
    SELECT  U_EMPCN1    =   SUM(T0.U_EMPCN1),
            U_GWAPA1    =   SUM(T0.U_GWAPA1),
            U_GWABN1    =   SUM(T0.U_GWABN1),
            U_BTXAM1    =   SUM(T0.U_BTXAM1),
            U_JIGTO1    =   SUM(T0.U_JIGTO1),
            U_GABGU1    =   SUM(T0.U_GABGU1),
            U_JUMIN1    =   SUM(T0.U_JUMIN1),
            U_MEDAM1    =   SUM(T0.U_MEDAM1),
            U_GBHAM1    =   SUM(T0.U_GBHAM1),
            U_KUKAM1    =   SUM(T0.U_KUKAM1),
            U_TOTGO1    =   SUM(T0.U_TOTGO1),
            U_EMPCN2    =   SUM(T0.U_EMPCN2),
            U_GWAPA2    =   SUM(T0.U_GWAPA2),
            U_GWABN2    =   SUM(T0.U_GWABN2),
            U_BTXAM2    =   SUM(T0.U_BTXAM2),
            U_JIGTO2    =   SUM(T0.U_JIGTO2),
            U_GABGU2    =   SUM(T0.U_GABGU2),
            U_JUMIN2    =   SUM(T0.U_JUMIN2),
            U_MEDAM2    =   SUM(T0.U_MEDAM2),
            U_GBHAM2    =   SUM(T0.U_GBHAM2),
            U_KUKAM2    =   SUM(T0.U_KUKAM2),
            U_TOTGO2    =   SUM(T0.U_TOTGO2)
    INTO    [#RPY509_5]
    FROM    [#RPY509_4] T0

---------------------------------------------------------------------------------------------------
-- 4. 자료 취합
---------------------------------------------------------------------------------------------------
    INSERT  INTO [#RPY509]
    SELECT  U_LINENO    =   T0.U_LINENO,
            U_CNTHLD    =   T0.U_CNTHLD,
            U_PAYHLD    =   T0.U_PAYHLD,
            U_BNSHLD    =   T0.U_BNSHLD,
            U_BTXHLD    =   T0.U_BTXHLD,
            U_JIGHLD    =   T0.U_JIGHLD,
            U_GABHLD    =   T0.U_GABHLD,
            U_JUMHLD    =   T0.U_JUMHLD,
            U_MEDHLD    =   T0.U_MEDHLD,
            U_GBHHLD    =   T0.U_GBHHLD,
            U_KUKHLD    =   T0.U_KUKHLD,
            U_TOTHLD    =   T0.U_TOTHLD,
            U_CNTRET    =   T0.U_CNTRET,
            U_PAYRET    =   T0.U_PAYRET,
            U_BNSRET    =   T0.U_BNSRET,
            U_BTXRET    =   T0.U_BTXRET,
            U_JIGRET    =   T0.U_JIGRET,
            U_GABRET    =   T0.U_GABRET,
            U_JUMRET    =   T0.U_JUMRET,
            U_MEDRET    =   T0.U_MEDRET,
            U_GBHRET    =   T0.U_GBHRET,
            U_KUKRET    =   T0.U_KUKRET,
            U_TOTRET    =   T0.U_TOTRET,
            U_EMPCNT    =   T0.U_EMPCNT,
            U_GWAPAY    =   T0.U_GWAPAY,
            U_GWABNS    =   T0.U_GWABNS,
            U_BTXAMT    =   T0.U_BTXAMT,
            U_JIGTOT    =   T0.U_JIGTOT,
            U_GABGUN    =   T0.U_GABGUN,
            U_JUMINN    =   T0.U_JUMINN,
            U_MEDAMT    =   T0.U_MEDAMT,
            U_GBHAMT    =   T0.U_GBHAMT,
            U_KUKAMT    =   T0.U_KUKAMT,
            U_TOTGON    =   T0.U_TOTGON,
            U_EMPCN1    =   T1.U_EMPCN1,
            U_GWAPA1    =   T1.U_GWAPA1,
            U_GWABN1    =   T1.U_GWABN1,
            U_BTXAM1    =   T1.U_BTXAM1,
            U_JIGTO1    =   T1.U_JIGTO1,
            U_GABGU1    =   T1.U_GABGU1,
            U_JUMIN1    =   T1.U_JUMIN1,
            U_MEDAM1    =   T1.U_MEDAM1,
            U_GBHAM1    =   T1.U_GBHAM1,
            U_KUKAM1    =   T1.U_KUKAM1,
            U_TOTGO1    =   T1.U_TOTGO1,
            U_EMPCN2    =   T1.U_EMPCN2,
            U_GWAPA2    =   T1.U_GWAPA2,
            U_GWABN2    =   T1.U_GWABN2,
            U_BTXAM2    =   T1.U_BTXAM2,
            U_JIGTO2    =   T1.U_JIGTO2,
            U_GABGU2    =   T1.U_GABGU2,
            U_JUMIN2    =   T1.U_JUMIN2,
            U_MEDAM2    =   T1.U_MEDAM2,
            U_GBHAM2    =   T1.U_GBHAM2,
            U_KUKAM2    =   T1.U_KUKAM2,
            U_TOTGO2    =   T1.U_TOTGO2
            
    FROM    [#RPY509_6] T0,
            [#RPY509_5] T1

---------------------------------------------------------------------------------------------------
-- 5. 결과 조회
---------------------------------------------------------------------------------------------------
    SELECT * FROM [#RPY509] ORDER BY U_LINENO
    
--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF


