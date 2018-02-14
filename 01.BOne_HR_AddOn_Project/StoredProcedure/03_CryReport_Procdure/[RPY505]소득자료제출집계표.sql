IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY505' AND xtype = 'P'))
	DROP PROCEDURE RPY505
GO

CREATE PROC RPY505
    (
        @JSNFYMM    AS Nvarchar(6),     --귀속년월(From)
        @JSNTYMM    AS Nvarchar(6),     --귀속년월(To)
        @SINFYMM    AS Nvarchar(6),     --원천신고년월(From)
        @SINTYMM    AS Nvarchar(6),     --원천신고년월(To)
        @JOBGBN     AS Nvarchar(1),     --작업구분(1연말정산,2중도정산,3전체)
        @CLTCOD     AS Nvarchar(8),     --징수의무자
        @MSTDPT     AS Nvarchar(8),     --부서
        @MSTCOD     AS Nvarchar(8),     --사원번호      
        @PRTTYP     AS Nvarchar(2)      --출력구분(1근로소득,2퇴직소득,3사업소득,4기타소득,
                                        --         5비거주자,6이자소득,7배당소득)
    )
      

 AS
    /*==========================================================================================
        프로시저명      : RPY502
        프로시저설명    : 소득자료제출집계표
        만든이          : 함미경
        작업일자        : 2007-01-30
        작업지시자      : 함미경
        작업지시일자    : 2007-01-30
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/
    -- DROP PROC RPY505
    /*
    Exec RPY505  '200801', '200812', '', '', '3', N'1', N'%',  N'%', '1'
    GO
    Exec RPY505  '200901', '201012', '', '', '3', N'1', N'%',  N'%', '2'
    GO
    Exec RPY505  '200901', '201012', '', '', '3', N'1', N'%',  N'%', '3'
    GO
    Exec RPY505  '200801', '200812', '', '', '3', N'1', N'%',  N'%', '4'
    GO
    Exec RPY505  '200801', '200812', '', '', '3', N'1', N'%',  N'%', '5'
    GO
    Exec RPY505  '200801', '200812', '', '', '3', N'1', N'%',  N'%', '6'
    GO
    Exec RPY505  '200901', '201012', '', '', '3', N'1', N'%',  N'%', '7'
    GO

    Exec RPY505  '', '', '200801', '200912', '3', '1', '%',  '%', '1'
    GO
    Exec RPY505  '', '', '200801', '200912', '3', N'1', N'%',  N'%', '2'
    GO
    Exec RPY505  '', '', '200801', '200912', '3', N'1', N'%',  N'%', '3'
    GO
    Exec RPY505  '', '', '200801', '200912', '3', N'1', N'%',  N'%', '4'
    GO
    Exec RPY505  '', '', '200801', '200912', '3', N'1', N'%',  N'%', '5'
    GO
    Exec RPY505  '', '', '200801', '200912', '3', N'1', N'%',  N'%', '6'
    GO
    Exec RPY505  '', '', '200801', '200912', '3', N'1', N'%',  N'%', '7'
    GO
	Exec RPY505 '200801', '200812', '200801', '200912', '3', '1', '%', '%', '1'

    */
    SET NOCOUNT ON

---------------------------------------------------------------------------------------------------
-- 1.임시테이블 생성, 인수세팅
---------------------------------------------------------------------------------------------------

    CREATE TABLE #RPY505 (
            JSNYER     nvarchar(4),
            INCOME     Numeric(19,6),
            TOTCNT     Numeric(19,6),
            GULGAB     Numeric(19,6),
            GULBUB     Numeric(19,6),
            GULNON     Numeric(19,6),
            GULJUM     Numeric(19,6),
            BUSNUM     nvarchar(12), 
            PERNUM     nvarchar(14), 
            CLTNAM     nvarchar(50), 
            COMPRT     nvarchar(30), 
            POSADD     nvarchar(100), 
            TELNUM     nvarchar(20), 
            TAXNAM     nvarchar(20) 
            ) 

    IF @JSNFYMM = ''
        SET @JSNFYMM = '190001'
    IF @JSNTYMM = ''
        SET @JSNTYMM = '299912'
    IF @SINFYMM = ''
        SET @SINFYMM = '190001'
    IF @SINTYMM = ''
        SET @SINTYMM = '299912'

---------------------------------------------------------------------------------------------------
-- 2.정산자료 조회 
---------------------------------------------------------------------------------------------------
    -- 1) 근로소득
    IF  (@PRTTYP = '1')
        INSERT INTO [#RPY505] 
        SELECT  JSNYER  =   T0.U_JSNYER,
                INCOME  =   ISNULL(T0.INCOME, 0) ,
                TOTCNT  =   ISNULL(T0.TOTCNT, 0) ,
                GULGAB  =   ISNULL(T0.GULGAB, 0) ,
                GULBUB  =   ISNULL(T0.GULBUB, 0) ,
                GULNON  =   ISNULL(T0.GULNON, 0) ,
                GULJUM  =   ISNULL(T0.GULJUM, 0) ,              
                BUSNUM  =   ISNULL(T1.U_BUSNum, ''),
                PERNUM  =   ISNULL(T1.U_PerNum, ''),
                CLTNAM  =   ISNULL(T1.U_CLTName, ''),
                COMPRT  =   ISNULL(T1.U_ComPrt, ''),
                POSADD  =   ISNULL(T1.U_PosAdd, ''),
                TELNUM  =   ISNULL(T1.U_TelNum, ''),
                TAXNAM  =   ISNULL(T1.U_TaxName, '')
        FROM
        (
            SELECT  T0.U_JSNYER,
                    SUM(T0.U_INCOME+T0.U_BIGTOT)    AS INCOME,
                    ISNULL(COUNT(T0.U_MSTCOD),0)    AS TOTCNT,
                    SUM(T0.U_GULGAB)                AS GULGAB,
                    SUM(0)                          AS GULBUB,
                    SUM(T0.U_GULNON)                AS GULNON,
                    SUM(T0.U_GULJUM)                AS GULJUM,
                    MAX(T1.U_CLTCOD)                AS U_CLTCOD
            FROM    [@ZPY504H] T0
                    --INNER JOIN [OHEM] T1 ON T0.U_EmpID  =  T1.EmpID
                    INNER JOIN [@PH_PY001A] T1 ON T0.U_EmpID  =  T1.U_EmpID
                    --INNER JOIN [OUDP] T2 ON T1.Dept     =  T2.Code
                                     
            WHERE   T0.U_JSNYER + T0.U_JSNMON BETWEEN @JSNFYMM AND @JSNTYMM
            AND     T0.U_JIGYMM BETWEEN @SINFYMM AND @SINTYMM 
            AND    (T0.U_JSNGBN =    @JOBGBN   OR    @JOBGBN     =    '3')
            AND     ISNULL(T1.U_CLTCOD,'') =  @CLTCOD    
            AND     ISNULL(T1.U_TeamCode,'') LIKE @MSTDPT                        
            AND     T0.U_MSTCOD LIKE @MSTCOD
            AND    (T0.U_INCOME+T0.U_BIGTOT) > 0
            
--          AND  ('3' = @JOBGBN             -- 전체
--          OR   ('2' = @JOBGBN             -- 중도정산
--          AND  (ISNULL(CONVERT(CHAR(10),TermDate,20),'') <>   '' 
--          AND   ISNULL(CONVERT(CHAR(10),TermDate,20),'') <    @JSNYER + '-' + @ENDMON + '-31')) 
--          OR   ('1' = @JOBGBN             -- 연말정산
--          AND  (ISNULL(CONVERT(CHAR(10),TermDate,20),'') =    '' 
--          OR    ISNULL(CONVERT(CHAR(10),TermDate,20),'') >=   @JSNYER + '-' + @ENDMON + '-31')))

            GROUP BY T0.U_JSNYER
        ) T0  RIGHT JOIN [@PH_PY005A]  T1 ON T0.U_CLTCOD = T1.U_CLTCode
        WHERE T1.CODE = @CLTCOD

    -- 2) 퇴직소득
    ELSE IF @PRTTYP = '2' 
        INSERT INTO [#RPY505]
        SELECT  JSNYER  =   CONVERT(NVARCHAR(4),T0.U_ENDINT,112),
                INCOME  =   SUM(T0.U_RETPAY),
                TOTCNT  =   ISNULL(COUNT(DISTINCT T1.U_MSTCOD),0),
                GULGAB  =   SUM(T0.U_CHAGAB),
                GULBUB  =   0,
                GULNON  =   0,
                GULJUM  =   SUM(T0.U_CHAJUM),
                BUSNUM  =   MAX(T3.U_BUSNUM),
                PERNUM  =   MAX(T3.U_PERNUM),
                CLTNAM  =   MAX(T3.U_CLTNAME),
                COMPRT  =   MAX(T3.U_COMPRT),
                POSADD  =   MAX(T3.U_POSADD),
                TELNUM  =   MAX(T3.U_TELNUM),
                TAXNAM  =   MAX(T3.U_TAXNAME)
        FROM    [@ZPY401H] T0   
                INNER JOIN [OHEM]     T1 ON T0.U_MSTCOD = T1.U_MSTCOD
                INNER JOIN [OUDP]     T2 ON T1.Dept     =  T2.Code
                INNER JOIN [@ZPY106H] T3 ON T1.U_CLTCOD = T3.CODE
        WHERE   CONVERT(NVARCHAR(6),T0.U_ENDINT,112) BETWEEN @JSNFYMM AND @JSNTYMM 
        AND     T0.U_SINYMM BETWEEN @SINFYMM AND @SINTYMM
--      AND     (T0.U_JSNGBN =    @JOBGBN   OR    @JOBGBN     =    '3')
        AND     ISNULL(T1.U_CLTCOD,'') =  @CLTCOD      
        AND     ISNULL(T2.U_MSTDPT,'') LIKE @MSTDPT
        AND     T0.U_MSTCOD LIKE    @MSTCOD
        /* 정산종료일기준으로 
        AND     CONVERT(NVARCHAR(6),T0.U_ENDINT,112)    >=  @JSNYER + @STRMON
        AND     CONVERT(NVARCHAR(6),T0.U_ENDINT,112)    <=  @JSNYER + @ENDMON   */  
        GROUP   BY  CONVERT(NVARCHAR(4),T0.U_ENDINT,112)

    -- 3) 사업소득(거주자)
    ELSE IF @PRTTYP = '3'
        INSERT INTO [#RPY505]
        SELECT  JSNYER  =   T0.U_JOBYER,
                INCOME  =   SUM(T1.U_AMOUNT),
                TOTCNT  =   ISNULL(COUNT(DISTINCT T2.CODE),0),
                GULGAB  =   SUM(T1.U_GULGAB),
                GULBUB  =   0,
                GULNON  =   0,
                GULJUM  =   SUM(T1.U_GULJUM),
                BUSNUM  =   MAX(T3.U_BUSNUM),
                PERNUM  =   MAX(T3.U_PERNUM),
                CLTNAM  =   MAX(T3.U_CLTNAME),
                COMPRT  =   MAX(T3.U_COMPRT),
                POSADD  =   MAX(T3.U_POSADD),
                TELNUM  =   MAX(T3.U_TELNUM),
                TAXNAM  =   MAX(T3.U_TAXNAME)
        FROM    [@ZPY531H] T0
                INNER JOIN [@ZPY531L] T1 ON T0.DOCENTRY = T1.DOCENTRY
                INNER JOIN [@ZPY530H] T2 ON T0.U_MSTCOD = T2.CODE
                INNER JOIN [@ZPY106H] T3 ON T2.U_CLTCOD = T3.CODE
        WHERE   T0.U_JOBYER + T1.U_JOBMON BETWEEN @JSNFYMM AND @JSNTYMM 
        AND     T1.U_SINYMM BETWEEN @SINFYMM AND @SINTYMM
--      AND     (T0.U_JSNGBN =    @JOBGBN   OR    @JOBGBN     =    '3')        
        AND     ISNULL(T2.U_CLTCOD,'') =  @CLTCOD      
        AND     ISNULL(T2.U_PNLCOD,'') LIKE @MSTDPT
        AND     T2.CODE     LIKE    @MSTCOD
        AND     T2.U_DWEGBN =       '1' -- 거주자만.. ONLY

--      AND     T0.U_JOBYER =       @JSNYER
--      AND     T1.U_JOBMON >=      @STRMON
--      AND     T1.U_JOBMON <=      @ENDMON
        GROUP   BY  T0.U_JOBYER

    -- 4) 기타소득(거주자)
    ELSE IF @PRTTYP = '4'
        INSERT INTO [#RPY505]
        SELECT  JSNYER  =   T0.U_JOBYER,
                INCOME  =   SUM(T1.U_INCOME),
                TOTCNT  =   ISNULL(COUNT(DISTINCT T2.CODE),0),
                GULGAB  =   SUM(T1.U_GULGAB),
                GULBUB  =   SUM(T1.U_GULCOM),
                GULNON  =   SUM(T1.U_GULNON),
                GULJUM  =   SUM(T1.U_GULJUM),
                BUSNUM  =   MAX(T3.U_BUSNUM),
                PERNUM  =   MAX(T3.U_PERNUM),
                CLTNAM  =   MAX(T3.U_CLTNAME),
                COMPRT  =   MAX(T3.U_COMPRT),
                POSADD  =   MAX(T3.U_POSADD),
                TELNUM  =   MAX(T3.U_TELNUM),
                TAXNAM  =   MAX(T3.U_TAXNAME)
        FROM    [@ZPY541H] T0
                INNER JOIN [@ZPY541L] T1 ON T0.DOCENTRY = T1.DOCENTRY
                INNER JOIN [@ZPY540H] T2 ON T0.U_MSTCOD = T2.CODE
                INNER JOIN [@ZPY106H] T3 ON T2.U_CLTCOD = T3.CODE
        WHERE   T0.U_JOBYER + T1.U_JOBMON BETWEEN @JSNFYMM AND @JSNTYMM 
        AND     T1.U_SINYMM BETWEEN @SINFYMM AND @SINTYMM
        AND     ISNULL(T2.U_CLTCOD,'') =  @CLTCOD      
        AND     ISNULL(T2.U_PNLCOD,'') LIKE @MSTDPT
        AND     T2.CODE     LIKE    @MSTCOD
        AND     T2.U_DWEGBN =       '1' -- 거주자만.. ONLY        
    
--      AND     T0.U_JOBYER =       @JSNYER
--      AND     T1.U_JOBMON >=      @STRMON
--      AND     T1.U_JOBMON <=      @ENDMON

        GROUP   BY  T0.U_JOBYER

    -- 5) 사업.기타소득(비거주자)
    ELSE IF @PRTTYP = '5'
        INSERT INTO [#RPY505]
        SELECT  JSNYER  =   T0.U_JOBYER,
                INCOME  =   SUM(T1.U_INCOME),
                TOTCNT  =   ISNULL(COUNT(DISTINCT T2.CODE),0),
                GULGAB  =   SUM(T1.U_GULGAB),
                GULBUB  =   SUM(T1.U_GULCOM),
                GULNON  =   SUM(T1.U_GULNON),
                GULJUM  =   SUM(T1.U_GULJUM),
                BUSNUM  =   MAX(T3.U_BUSNUM),
                PERNUM  =   MAX(T3.U_PERNUM),
                CLTNAM  =   MAX(T3.U_CLTNAME),
                COMPRT  =   MAX(T3.U_COMPRT),
                POSADD  =   MAX(T3.U_POSADD),
                TELNUM  =   MAX(T3.U_TELNUM),
                TAXNAM  =   MAX(T3.U_TAXNAME)

        FROM    [@ZPY541H] T0   
                INNER JOIN [@ZPY541L] T1 ON T0.DOCENTRY = T1.DOCENTRY
                INNER JOIN [@ZPY540H] T2 ON T0.U_MSTCOD = T2.CODE
                INNER JOIN [@ZPY106H] T3 ON T2.U_CLTCOD = T3.CODE
        WHERE   T0.U_JOBYER + T1.U_JOBMON BETWEEN @JSNFYMM AND @JSNTYMM
        AND     T1.U_SINYMM               BETWEEN @SINFYMM AND @SINTYMM
        AND     ISNULL(T2.U_CLTCOD,'')  =       @CLTCOD      
        AND     ISNULL(T2.U_PNLCOD,'')  LIKE    @MSTDPT
        AND     T2.CODE                 LIKE    @MSTCOD
        AND     T2.U_DWEGBN =       '2' -- 비거주자만.. ONLY                
        GROUP   BY  T0.U_JOBYER

    -- 6) 이자소득
    ELSE IF @PRTTYP = '6'
        INSERT INTO [#RPY505]
        SELECT  JSNYER  =   T0.U_JOBYER,
                INCOME  =   SUM(T1.U_AMOUNT),
                TOTCNT  =   ISNULL(COUNT(DISTINCT T2.CODE), 0),
                GULGAB  =   SUM(T1.U_GULGAB),
                GULBUB  =   SUM(T1.U_GULCOM),
                GULNON  =   SUM(T1.U_GULNON),
                GULJUM  =   SUM(T1.U_GULJUM),
                BUSNUM  =   MAX(T3.U_BUSNUM),
                PERNUM  =   MAX(T3.U_PERNUM),
                CLTNAM  =   MAX(T3.U_CLTNAME),
                COMPRT  =   MAX(T3.U_COMPRT),
                POSADD  =   MAX(T3.U_POSADD),
                TELNUM  =   MAX(T3.U_TELNUM),
                TAXNAM  =   MAX(T3.U_TAXNAME)
        FROM    [@ZPY542H] T0   
                INNER JOIN [@ZPY542L] T1 ON T0.DOCENTRY = T1.DOCENTRY
                INNER JOIN [@ZPY540H] T2 ON T0.U_MSTCOD = T2.CODE
                INNER JOIN [@ZPY106H] T3 ON T2.U_CLTCOD = T3.CODE
        WHERE   T0.U_JOBYER + T1.U_JOBMON BETWEEN @JSNFYMM AND @JSNTYMM
        AND     T1.U_SINYMM BETWEEN @SINFYMM AND @SINTYMM
        AND     ISNULL(T2.U_CLTCOD,'') =  @CLTCOD      
        AND     ISNULL(T2.U_PNLCOD,'') LIKE @MSTDPT
        AND     T2.CODE     LIKE    @MSTCOD
        AND     T1.U_CODTYP =       '1' -- 1.이자만 ONLY..
        GROUP   BY  T0.U_JOBYER

    -- 7) 배당소득
    ELSE IF @PRTTYP = '7'
        INSERT INTO [#RPY505]
        SELECT  JSNYER  =   T0.U_JOBYER,
                INCOME  =   SUM(T1.U_AMOUNT),
                TOTCNT  =   ISNULL(COUNT(DISTINCT T2.CODE), 0),
                GULGAB  =   SUM(T1.U_GULGAB),
                GULBUB  =   SUM(T1.U_GULCOM),
                GULNON  =   SUM(T1.U_GULNON),
                GULJUM  =   SUM(T1.U_GULJUM),
                BUSNUM  =   MAX(T3.U_BUSNUM),
                PERNUM  =   MAX(T3.U_PERNUM),
                CLTNAM  =   MAX(T3.U_CLTNAME),
                COMPRT  =   MAX(T3.U_COMPRT),
                POSADD  =   MAX(T3.U_POSADD),
                TELNUM  =   MAX(T3.U_TELNUM),
                TAXNAM  =   MAX(T3.U_TAXNAME)

        FROM    [@ZPY542H] T0   
                INNER JOIN [@ZPY542L] T1 ON T0.DOCENTRY = T1.DOCENTRY
                INNER JOIN [@ZPY540H] T2 ON T0.U_MSTCOD = T2.CODE
                INNER JOIN [@ZPY106H] T3 ON T2.U_CLTCOD = T3.CODE
        WHERE   T0.U_JOBYER + T1.U_JOBMON BETWEEN @JSNFYMM AND @JSNTYMM
        AND     T1.U_SINYMM BETWEEN @SINFYMM AND @SINTYMM
        AND     ISNULL(T2.U_CLTCOD,'') =  @CLTCOD      
        AND     ISNULL(T2.U_PNLCOD,'') LIKE @MSTDPT
        AND     T2.CODE     LIKE    @MSTCOD
        AND     T1.U_CODTYP = '2' -- 2.배당만 ONLY..
        GROUP   BY  T0.U_JOBYER

-- <3.정산자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
    
    SELECT * FROM [#RPY505] ORDER BY JSNYER
    
--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF
