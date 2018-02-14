IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY510' AND xtype = 'P'))
	DROP PROCEDURE RPY510
GO

CREATE PROC RPY510
    (
        @JSNYER     AS Nvarchar(4),     --작업연월
        @STRMON     AS Nvarchar(2),     --시작월
        @ENDMON     AS Nvarchar(2),     --종료월
        @JOBGBN     AS Nvarchar(1),     --작업구분(1연말정산,2중도정산,3전체)
        @CLTCOD     AS Nvarchar(8),     --자사코드
        @MSTDPT     AS Nvarchar(8),     --부서
        @MSTCOD     AS Nvarchar(8),     --사원번호      
        @AMTYN      AS Nvarchar(8)      --전체 출력 여부(N:비과세금액이 없어도 출력, Y:비과세 금액이 없으면 출력안함)
    )
   

 AS
    /*==========================================================================================
        프로시저명      : RPY510
        프로시저설명    : 비과세 근로소득 명세서
        만든이          : 최동권
        작업일자        : 2009-01-05
        작업지시자      : 함미경
        작업지시일자    : 
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/
    --DROP PROC RPY510
    --Exec RPY510 '2013', '01', '12', '3', '%', '%', '%', 'Y'
    SET NOCOUNT ON

---------------------------------------------------------------------------------------------------
-- 1.임시테이블 생성 
---------------------------------------------------------------------------------------------------

    -- 1.1) 메인 임시테이블
    CREATE TABLE #RPY510 (
        JSNYER      nvarchar(4),
        MSTCOD      nvarchar(8),
        MSTNAM      nvarchar(50),
        CLTCOD      nvarchar(8),
        STRINT      nvarchar(10),
        ENDINT      nvarchar(10),
        J01NAM      nvarchar(40),
        J02NAM      nvarchar(40),
        J01NBR      nvarchar(12),
        J02NBR      nvarchar(12),
        BIGTOT      Numeric(19,6),
        J01TOT      Numeric(19,6),
        J02TOT      Numeric(19,6),
        PERNBR      nvarchar(20),
        ADDRES      nvarchar(100),
        CLTNAM      nvarchar(50),
        COMPRT      nvarchar(30),
        BUSNUM      nvarchar(12),
        PERNUM      nvarchar(14),
        POSADD      nvarchar(100),
        BIRDAY      nvarchar(10),
        BIGM01      Numeric(19,6),  BIGM02      Numeric(19,6),  BIGM03      Numeric(19,6),  BIGO01      Numeric(19,6),
        BIGQ01      Numeric(19,6),  BIGX01      Numeric(19,6),  BIGH06      Numeric(19,6),  BIGH07      Numeric(19,6),
        BIGH08      Numeric(19,6),  BIGH09      Numeric(19,6),  BIGH10      Numeric(19,6),  BIGG01      Numeric(19,6),
        BIGH11      Numeric(19,6),  BIGH12      Numeric(19,6),  BIGH13      Numeric(19,6),  BIGH01      Numeric(19,6),
        BIGK01      Numeric(19,6),  BIGS01      Numeric(19,6),  BIGT01      Numeric(19,6),  BIGY01      Numeric(19,6),
        BIGY02      Numeric(19,6),  BIGY03      Numeric(19,6),  BIGY20      Numeric(19,6),  BIGY21      Numeric(19,6),  
        BIGZ01      Numeric(19,6),  BIGH05      Numeric(19,6),  BIGI01      Numeric(19,6),  
        BIGWA2      Numeric(19,6),  BIGWA4      Numeric(19,6),  BIGWA7      Numeric(19,6),
        JBTM011     Numeric(19,6),  JBTM021     Numeric(19,6),  JBTM031     Numeric(19,6),  JBTO011     Numeric(19,6),
        JBTQ011     Numeric(19,6),  JBTX011     Numeric(19,6),  JBTH061     Numeric(19,6),  JBTH071     Numeric(19,6),
        JBTH081     Numeric(19,6),  JBTH091     Numeric(19,6),  JBTH101     Numeric(19,6),  JBTG011     Numeric(19,6),
        JBTH111     Numeric(19,6),  JBTH121     Numeric(19,6),  JBTH131     Numeric(19,6),  JBTH011     Numeric(19,6),
        JBTK011     Numeric(19,6),  JBTS011     Numeric(19,6),  JBTT011     Numeric(19,6),  JBTY011     Numeric(19,6),
        JBTY021     Numeric(19,6),  JBTY031     Numeric(19,6),  JBTY201     Numeric(19,6),  JBTY211     Numeric(19,6),  
        JBTZ011     Numeric(19,6),  JBTH051     Numeric(19,6),  JBTI011     Numeric(19,6),
        JBIG021     Numeric(19,6),  JBIG041     Numeric(19,6),
        JBTM012     Numeric(19,6),  JBTM022     Numeric(19,6),  JBTM032     Numeric(19,6),  JBTO012     Numeric(19,6),
        JBTQ012     Numeric(19,6),  JBTX012     Numeric(19,6),  JBTH062     Numeric(19,6),  JBTH072     Numeric(19,6),
        JBTH082     Numeric(19,6),  JBTH092     Numeric(19,6),  JBTH102     Numeric(19,6),  JBTG012     Numeric(19,6),
        JBTH112     Numeric(19,6),  JBTH122     Numeric(19,6),  JBTH132     Numeric(19,6),  JBTH012     Numeric(19,6),
        JBTK012     Numeric(19,6),  JBTS012     Numeric(19,6),  JBTT012     Numeric(19,6),  JBTY012     Numeric(19,6),
        JBTY022     Numeric(19,6),  JBTY032     Numeric(19,6),  JBTY202     Numeric(19,6),  JBTY212     Numeric(19,6),  
        JBTZ012     Numeric(19,6),  JBTH052     Numeric(19,6),  JBTI012     Numeric(19,6),
        JBIG022     Numeric(19,6),  JBIG042     Numeric(19,6)
        ) 

--SELECT * FROM [#RPY510_1]
---------------------------------------------------------------------------------------------------
-- 2.정산자료 조회 
---------------------------------------------------------------------------------------------------
    INSERT INTO [#RPY510]
    SELECT  JSNYER    =   T0.U_JSNYER,
            MSTCOD    =   T0.U_MSTCOD,
            MSTNAM    =   T0.U_MSTNAM,
            CLTCOD    =   T0.U_CLTCOD,
            STRINT    =   CONVERT(CHAR(10),T0.U_STRINT,120),
            ENDINT    =   CONVERT(CHAR(10),T0.U_ENDINT,120),
            J01NAM    =   T0.U_J01NAM,
            J02NAM    =   T0.U_J02NAM,
            J01NBR    =   T0.U_J01NBR,
            J02NBR    =   T0.U_J02NBR,
            BIGTOT    =   0,
            J01TOT    =   0,
            J02TOT    =   0,
            PERNBR    =   ISNULL(T2.GovID, ''),
            ADDRES    =   ISNULL(T2.HomeStreet, ''),
            CLTNAM    =   ISNULL(T1.U_CLTName, ''),
            COMPRT    =   ISNULL(T1.U_ComPrt, ''),
            BUSNUM    =   ISNULL(T1.U_BusNum, ''),
            PERNUM    =   ISNULL(T1.U_PerNum, ''),
            POSADD    =   ISNULL(T1.U_PosAdd, ''),
            BIRDAY    =   CONVERT(CHAR(10),T2.birthDate,120),   
            
            ISNULL(T5.BIGM01,0),    ISNULL(T5.BIGM02,0),    ISNULL(T5.BIGM03,0),    ISNULL(T5.BIGO01,0),
            ISNULL(T5.BIGQ01,0),    ISNULL(T5.BIGX01,0),    ISNULL(T5.BIGH06,0),    ISNULL(T5.BIGH07,0),
            ISNULL(T5.BIGH08,0),    ISNULL(T5.BIGH09,0),    ISNULL(T5.BIGH10,0),    ISNULL(T5.BIGG01,0),
            ISNULL(T5.BIGH11,0),    ISNULL(T5.BIGH12,0),    ISNULL(T5.BIGH13,0),    ISNULL(T5.BIGH01,0),
            ISNULL(T5.BIGK01,0),    ISNULL(T5.BIGS01,0),    ISNULL(T5.BIGT01,0),    ISNULL(T5.BIGY01,0),
            ISNULL(T5.BIGY02,0),    ISNULL(T5.BIGY03,0),    ISNULL(T5.BIGY20,0),    ISNULL(T5.BIGY21,0),    
            ISNULL(T5.BIGZ01,0),    ISNULL(T5.BIGH05,0),    ISNULL(T5.BIGI01,0),
            ISNULL(T5.BIGWA2,0),    ISNULL(T5.BIGWA4,0),    ISNULL(T5.BIGWA7,0),

            ISNULL(T4.JBTM011,0),   ISNULL(T4.JBTM021,0),   ISNULL(T4.JBTM031,0),   ISNULL(T4.JBTO011,0),
            ISNULL(T4.JBTQ011,0),   ISNULL(T4.JBTX011,0),   ISNULL(T4.JBTH061,0),   ISNULL(T4.JBTH071,0),
            ISNULL(T4.JBTH081,0),   ISNULL(T4.JBTH091,0),   ISNULL(T4.JBTH101,0),   ISNULL(T4.JBTG011,0),
            ISNULL(T4.JBTH111,0),   ISNULL(T4.JBTH121,0),   ISNULL(T4.JBTH131,0),   ISNULL(T4.JBTH011,0),
            ISNULL(T4.JBTK011,0),   ISNULL(T4.JBTS011,0),   ISNULL(T4.JBTT011,0),   ISNULL(T4.JBTY011,0),
            ISNULL(T4.JBTY021,0),   ISNULL(T4.JBTY031,0),   ISNULL(T4.JBTY201,0),   ISNULL(T4.JBTY211,0),   
            ISNULL(T4.JBTZ011,0),   ISNULL(T4.JBTH051,0),   ISNULL(T4.JBTI011,0),
            ISNULL(T4.JBIG021,0),   ISNULL(T4.JBIG041,0),
            
            ISNULL(T4.JBTM012,0),   ISNULL(T4.JBTM022,0),   ISNULL(T4.JBTM032,0),   ISNULL(T4.JBTO012,0),
            ISNULL(T4.JBTQ012,0),   ISNULL(T4.JBTX012,0),   ISNULL(T4.JBTH062,0),   ISNULL(T4.JBTH072,0),
            ISNULL(T4.JBTH082,0),   ISNULL(T4.JBTH092,0),   ISNULL(T4.JBTH102,0),   ISNULL(T4.JBTG012,0),
            ISNULL(T4.JBTH112,0),   ISNULL(T4.JBTH122,0),   ISNULL(T4.JBTH132,0),   ISNULL(T4.JBTH012,0),
            ISNULL(T4.JBTK012,0),   ISNULL(T4.JBTS012,0),   ISNULL(T4.JBTT012,0),   ISNULL(T4.JBTY012,0),
            ISNULL(T4.JBTY022,0),   ISNULL(T4.JBTY032,0),   ISNULL(T4.JBTY202,0),   ISNULL(T4.JBTY212,0),   
            ISNULL(T4.JBTZ012,0),   ISNULL(T4.JBTH052,0),   ISNULL(T4.JBTI012,0),
            ISNULL(T4.JBIG022,0),   ISNULL(T4.JBIG042,0)
    FROM    [@ZPY504H] T0   
            LEFT  JOIN [@ZPY106H] T1 ON T0.U_CLTCOD = T1.U_CLTCode
            INNER JOIN [@PH_PY001A] T2 ON T0.U_EmpID = T2.U_EmpID

            LEFT  JOIN (
            SELECT  A0.U_MSTCOD,
                    A0.U_JSNYER,
                    A0.U_CLTCOD,
                    JBTM011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTM01,0) ELSE 0 END), 
                    JBTM021 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTM02,0) ELSE 0 END), 
                    JBTM031 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTM03,0) ELSE 0 END), 
                    JBTO011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTO01,0) ELSE 0 END),
                    JBTQ011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTQ01,0) ELSE 0 END), 
                    JBTX011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTX01,0) ELSE 0 END), 
                    JBTH061 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH06,0) ELSE 0 END), 
                    JBTH071 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH07,0) ELSE 0 END),
                    JBTH081 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH08,0) ELSE 0 END), 
                    JBTH091 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH09,0) ELSE 0 END), 
                    JBTH101 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH10,0) ELSE 0 END), 
                    JBTG011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTG01,0) ELSE 0 END),
                    JBTH111 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH11,0) ELSE 0 END), 
                    JBTH121 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH12,0) ELSE 0 END), 
                    JBTH131 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH13,0) ELSE 0 END), 
                    JBTH011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH01,0) ELSE 0 END),
                    JBTK011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTK01,0) ELSE 0 END), 
                    JBTS011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTS01,0) ELSE 0 END), 
                    JBTT011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTT01,0) ELSE 0 END), 
                    JBTY011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTY01,0) ELSE 0 END),
                    JBTY021 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTY02,0) ELSE 0 END), 
                    JBTY031 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTY03,0) ELSE 0 END), 
                    JBTY201 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTY20,0) ELSE 0 END), 
                    JBTY211 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTY21,0) ELSE 0 END), 
                    JBTZ011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTZ01,0) ELSE 0 END),
                    JBTH051 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH05,0) ELSE 0 END), 
                    JBTI011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTI01,0) ELSE 0 END),
                    JBIG021 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JONBT2,0) ELSE 0 END),
                    JBIG041 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JONBT4,0) ELSE 0 END),
                    
                    JBTM012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTM01,0) ELSE 0 END), 
                    JBTM022 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTM02,0) ELSE 0 END), 
                    JBTM032 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTM03,0) ELSE 0 END), 
                    JBTO012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTO01,0) ELSE 0 END),
                    JBTQ012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTQ01,0) ELSE 0 END), 
                    JBTX012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTX01,0) ELSE 0 END), 
                    JBTH062 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH06,0) ELSE 0 END), 
                    JBTH072 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH07,0) ELSE 0 END),
                    JBTH082 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH08,0) ELSE 0 END), 
                    JBTH092 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH09,0) ELSE 0 END), 
                    JBTH102 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH10,0) ELSE 0 END), 
                    JBTG012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTG01,0) ELSE 0 END),
                    JBTH112 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH11,0) ELSE 0 END), 
                    JBTH122 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH12,0) ELSE 0 END), 
                    JBTH132 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH13,0) ELSE 0 END), 
                    JBTH012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH01,0) ELSE 0 END),
                    JBTK012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTK01,0) ELSE 0 END), 
                    JBTS012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTS01,0) ELSE 0 END), 
                    JBTT012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTT01,0) ELSE 0 END), 
                    JBTY012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTY01,0) ELSE 0 END),
                    JBTY022 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTY02,0) ELSE 0 END), 
                    JBTY032 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTY03,0) ELSE 0 END), 
                    JBTY202 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTY20,0) ELSE 0 END), 
                    JBTY212 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTY21,0) ELSE 0 END), 
                    JBTZ012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTZ01,0) ELSE 0 END),
                    JBTH052 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH05,0) ELSE 0 END), 
                    JBTI012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTI01,0) ELSE 0 END),
                    JBIG022 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JONBT2,0) ELSE 0 END),
                    JBIG042 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JONBT4,0) ELSE 0 END)
            FROM    [@ZPY502H] A0
                    INNER JOIN [@ZPY502L] A1 ON A0.DocEntry = A1.DocEntry
            WHERE   A0.U_JSNYER = @JSNYER
            AND     A0.U_MSTCOD LIKE @MSTCOD 
			AND     A0.U_CLTCOD LIKE @CLTCOD
			GROUP   BY A0.U_MSTCOD,  A0.U_JSNYER, A0.U_CLTCOD
            ) T4 ON T0.U_JSNYER = T4.U_JSNYER AND T0.U_MSTCOD = T4.U_MSTCOD AND T0.U_CLTCOD = T4.U_CLTCOD
		LEFT JOIN (
			SELECT	MSTCOD = A0.U_MstCode,
                    JSNYER = A0.U_JsnYear,
                    CLTCOD = A0.U_CLTCOD,
                    BIGM01 = ISNULL(A1.U_BIGM01,0), 
                    BIGM02 = ISNULL(A1.U_BIGM02,0), 
                    BIGM03 = ISNULL(A1.U_BIGM03,0), 
                    BIGO01 = ISNULL(A1.U_BIGO01,0),
                    BIGQ01 = ISNULL(A1.U_BIGQ01,0), 
                    BIGX01 = ISNULL(A1.U_BIGX01,0), 
                    BIGH06 = ISNULL(A1.U_BIGH06,0), 
                    BIGH07 = ISNULL(A1.U_BIGH07,0),
                    BIGH08 = ISNULL(A1.U_BIGH08,0), 
                    BIGH09 = ISNULL(A1.U_BIGH09,0), 
                    BIGH10 = ISNULL(A1.U_BIGH10,0), 
                    BIGG01 = ISNULL(A1.U_BIGG01,0),
                    BIGH11 = ISNULL(A1.U_BIGH11,0), 
                    BIGH12 = ISNULL(A1.U_BIGH12,0), 
                    BIGH13 = ISNULL(A1.U_BIGH13,0), 
                    BIGH01 = ISNULL(A1.U_BIGH01,0),
                    BIGK01 = ISNULL(A1.U_BIGK01,0), 
                    BIGS01 = ISNULL(A1.U_BIGS01,0), 
                    BIGT01 = ISNULL(A1.U_BIGT01,0), 
                    BIGY01 = ISNULL(A1.U_BIGY01,0),
                    BIGY02 = ISNULL(A1.U_BIGY02,0), 
                    BIGY03 = ISNULL(A1.U_BIGY03,0), 
                    BIGY20 = ISNULL(A1.U_BIGY20,0), 
                    BIGY21 = ISNULL(A1.U_BIGY21,0), 
                    BIGZ01 = ISNULL(A1.U_BIGZ01,0),
                    BIGH05 = ISNULL(A1.U_BIGH05,0), 
                    BIGI01 = ISNULL(A1.U_BIGI01,0), 
                    BIGWA2 = ISNULL(A1.U_BiGwa02,0), 
                    BIGWA4 = ISNULL(A1.U_BiGwa04,0), 
                    BIGWA7 = ISNULL(A1.U_BiGwa07,0)
			FROM	[@ZPY343H] A0
					INNER JOIN [@ZPY343L] A1 ON A0.DocEntry = A1.DocEntry
			WHERE	A0.U_JsnYear = @JSNYER
            AND     A0.U_MstCode LIKE @MSTCOD 
			AND     A0.U_CLTCOD LIKE @CLTCOD
			AND		A1.U_LineNum = '13'
            ) T5 ON T0.U_JSNYER = T5.JSNYER AND T0.U_MSTCOD = T5.MSTCOD AND T0.U_CLTCOD = T5.CLTCOD

    WHERE   T0.U_JSNYER =    @JSNYER
    AND     T0.U_CLTCOD LIKE @CLTCOD
    AND     T2.U_TeamCode LIKE @MSTDPT                        
    AND     T0.U_MSTCOD LIKE @MSTCOD
    AND     (@JOBGBN ='3' OR (@JOBGBN <> '3' AND T0.U_JSNGBN = @JOBGBN))
    AND     T0.U_JSNMON  BETWEEN @STRMON AND @ENDMON
    ORDER BY  T0.U_MSTNAM,  T0.U_MSTCOD

    UPDATE  [#RPY510]
    SET     BIGTOT  =   ISNULL(BIGM01,0)  + ISNULL(BIGM02,0)  + ISNULL(BIGM03,0)  + ISNULL(BIGO01,0) + 
                        ISNULL(BIGQ01,0)  + ISNULL(BIGX01,0)  + ISNULL(BIGH06,0)  + ISNULL(BIGH07,0) + 
                        ISNULL(BIGH08,0)  + ISNULL(BIGH09,0)  + ISNULL(BIGH10,0)  + ISNULL(BIGG01,0) + 
                        ISNULL(BIGH11,0)  + ISNULL(BIGH12,0)  + ISNULL(BIGH13,0)  + ISNULL(BIGH01,0) + 
                        ISNULL(BIGK01,0)  + ISNULL(BIGS01,0)  + ISNULL(BIGT01,0)  + ISNULL(BIGY01,0) + 
                        ISNULL(BIGY02,0)  + ISNULL(BIGY03,0)  + ISNULL(BIGY20,0)  + ISNULL(BIGY21,0) + 
                        ISNULL(BIGZ01,0)  + ISNULL(BIGH05,0)  + ISNULL(BIGI01,0)  + ISNULL(BIGWA2,0),
            J01TOT  =   ISNULL(JBTM011,0) + ISNULL(JBTM021,0) + ISNULL(JBTM031,0) + ISNULL(JBTO011,0) + 
                        ISNULL(JBTQ011,0) + ISNULL(JBTX011,0) + ISNULL(JBTH061,0) + ISNULL(JBTH071,0) + 
                        ISNULL(JBTH081,0) + ISNULL(JBTH091,0) + ISNULL(JBTH101,0) + ISNULL(JBTG011,0) + 
                        ISNULL(JBTH111,0) + ISNULL(JBTH121,0) + ISNULL(JBTH131,0) + ISNULL(JBTH011,0) + 
                        ISNULL(JBTK011,0) + ISNULL(JBTS011,0) + ISNULL(JBTT011,0) + ISNULL(JBTY011,0) + 
                        ISNULL(JBTY021,0) + ISNULL(JBTY031,0) + ISNULL(JBTY201,0) + ISNULL(JBTY211,0) + 
                        ISNULL(JBTZ011,0) + ISNULL(JBTH051,0) + ISNULL(JBTI011,0),
            J02TOT  =   ISNULL(JBTM012,0) + ISNULL(JBTM022,0) + ISNULL(JBTM032,0) + ISNULL(JBTO012,0) + 
                        ISNULL(JBTQ012,0) + ISNULL(JBTX012,0) + ISNULL(JBTH062,0) + ISNULL(JBTH072,0) + 
                        ISNULL(JBTH082,0) + ISNULL(JBTH092,0) + ISNULL(JBTH102,0) + ISNULL(JBTG012,0) + 
                        ISNULL(JBTH112,0) + ISNULL(JBTH122,0) + ISNULL(JBTH132,0) + ISNULL(JBTH012,0) + 
                        ISNULL(JBTK012,0) + ISNULL(JBTS012,0) + ISNULL(JBTT012,0) + ISNULL(JBTY012,0) + 
                        ISNULL(JBTY022,0) + ISNULL(JBTY032,0) + ISNULL(JBTY202,0) + ISNULL(JBTY212,0) + 
                        ISNULL(JBTZ012,0) + ISNULL(JBTH052,0) + ISNULL(JBTI012,0) 

    IF ISNULL(@AMTYN, 'N') = 'Y'
    BEGIN
        DELETE 
        FROM    [#RPY510]
        WHERE   BIGTOT <= 0
        AND     J01TOT <= 0
        AND     J02TOT <= 0
		AND		BIGWA2 <= 0
		AND		BIGWA7 <= 0
		AND		JBIG021 <= 0
		AND		JBIG022 <= 0
    END

-- <3. 자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
    SELECT * FROM [#RPY510] ORDER BY CLTCOD, MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF
