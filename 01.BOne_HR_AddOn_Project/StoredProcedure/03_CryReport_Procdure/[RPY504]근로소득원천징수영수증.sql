IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY504' AND xtype = 'P'))
	DROP PROCEDURE RPY504
GO

CREATE PROC RPY504
    (
        @JSNYER     AS Nvarchar(4),     --작업연월
        @STRMON     AS Nvarchar(2),     --시작월
        @ENDMON     AS Nvarchar(2),     --종료월
        @JOBGBN     AS Nvarchar(1),     --작업구분(1연말정산,2중도정산,3전체)
        @CLTCOD     AS Nvarchar(8),     --자사코드
        @MSTDPT     AS Nvarchar(8),     --부서
        @MSTCOD     AS Nvarchar(8)      --사원번호      
    )
   

 AS
    /*==========================================================================================
        프로시저명      : RPY504
        프로시저설명    : 근로소득원천징수영수증
        만든이          : 함미경
        작업일자        : 2007-01-30
        작업지시자      : 함미경
        작업지시일자    : 2007-01-30
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/
    --DROP PROC RPY504
    --Exec RPY504  '2009', '01', '12', '3', N'%', N'%',  N'%'
    SET NOCOUNT ON

---------------------------------------------------------------------------------------------------
-- 1.임시테이블 생성 
---------------------------------------------------------------------------------------------------
	DECLARE @BTXNAM_G01 AS nvarchar(100),   @BTXNAM_H01 AS nvarchar(100)
	DECLARE @BTXNAM_H05 AS nvarchar(100),   @BTXNAM_H06 AS nvarchar(100)
	DECLARE @BTXNAM_H07 AS nvarchar(100),   @BTXNAM_H08 AS nvarchar(100)
	DECLARE @BTXNAM_H09 AS nvarchar(100),   @BTXNAM_H10 AS nvarchar(100)
	DECLARE @BTXNAM_H11 AS nvarchar(100),   @BTXNAM_H12 AS nvarchar(100)
	DECLARE @BTXNAM_H13 AS nvarchar(100),   @BTXNAM_I01 AS nvarchar(100)
	DECLARE @BTXNAM_K01 AS nvarchar(100),   @BTXNAM_M01 AS nvarchar(100)
	DECLARE @BTXNAM_M02 AS nvarchar(100),   @BTXNAM_M03 AS nvarchar(100)
	DECLARE @BTXNAM_O01 AS nvarchar(100),   @BTXNAM_Q01 AS nvarchar(100)
	DECLARE @BTXNAM_S01 AS nvarchar(100),   @BTXNAM_T01 AS nvarchar(100)
	DECLARE @BTXNAM_X01 AS nvarchar(100),   @BTXNAM_Y01 AS nvarchar(100)
	DECLARE @BTXNAM_Y02 AS nvarchar(100),   @BTXNAM_Y03 AS nvarchar(100)
	DECLARE @BTXNAM_Y20 AS nvarchar(100),   @BTXNAM_Z01 AS nvarchar(100)

    -- 1.1) 메인 임시테이블
    CREATE TABLE #RPY504 (
        U_MSTCOD    nvarchar(8),
        U_MSTNAM    nvarchar(50),
        U_CLTCOD    nvarchar(8),
		U_MSTDPT	nvarchar(8),
		U_DPTNAM	nvarchar(40),
        U_STRINT    nvarchar(10),
        U_ENDINT    nvarchar(10),
        U_STRGAM    nvarchar(10),
        U_ENDGAM    nvarchar(10),
		U_JSTR01    nvarchar(10),   -- 종전근무지1 근무기간
        U_JEND01    nvarchar(10),
        U_JGFR01    nvarchar(10),   -- 종전근무지1 감면기간
        U_JGTO01    nvarchar(10),
		U_JSTR02    nvarchar(10),   -- 종전근무지2 근무기간
        U_JEND02    nvarchar(10),
        U_JGFR02    nvarchar(10),   -- 종전근무지2 감면기간
        U_JGTO02    nvarchar(10),
        U_J01NAM    nvarchar(40),
        U_J02NAM    nvarchar(40),
        U_J01NBR    nvarchar(12),
        U_J02NBR    nvarchar(12),
        U_JPAY01    Numeric(19,6),  
        U_JBNS01    Numeric(19,6),      
        U_JPAY02    Numeric(19,6),      
        U_JBNS02    Numeric(19,6),  
        U_JONGAB    Numeric(19,6),  
        U_JONJUM    Numeric(19,6),  
        U_JONNON    Numeric(19,6),
        U_JONGA1    Numeric(19,6),  
        U_JONJU1    Numeric(19,6),  
        U_JONNO1    Numeric(19,6),
        U_JONGA2    Numeric(19,6),  
        U_JONJU2    Numeric(19,6),  
        U_JONNO2    Numeric(19,6),
        U_PAYAMT    Numeric(19,6),
        U_BNSAMT    Numeric(19,6),
        U_INBAMT    Numeric(19,6),
        U_TOTAMT    Numeric(19,6),
        U_BIGWA1    Numeric(19,6),
        U_BIGWA2    Numeric(19,6),
        U_BIGWA3    Numeric(19,6),
        U_BIGWU3    Numeric(19,6),
        U_BIGWA4    Numeric(19,6),
        U_BIGWA5    Numeric(19,6),
        U_BIGWA6    Numeric(19,6),
        U_BIGTOT    Numeric(19,6),
        U_INCOME    Numeric(19,6),
        U_PILGNL    Numeric(19,6),
        U_GNLOSD    Numeric(19,6),
        U_INJBAS    Numeric(19,6),
        U_INJBWO    Numeric(19,6),
        U_INJBYN    Numeric(19,6),
        U_INJGYN    Numeric(19,6),
        U_INJJAE    Numeric(19,6),
        U_INJBNJ    Numeric(19,6),
        U_INJSON    Numeric(19,6),
        U_INJADD    Numeric(19,6),
        U_INJCHL    Numeric(19,6),
        U_KUKGON    Numeric(19,6),
        U_PILBHM    Numeric(19,6),
        U_PILMED    Numeric(19,6),
        U_PILSCH    Numeric(19,6),
        U_PILHUS    Numeric(19,6),
        U_PILJHE    Numeric(19,6),
        U_PILGBU    Numeric(19,6),
        U_PILHUN    Numeric(19,6),
        U_PILTOT    Numeric(19,6),
        U_PILGON    Numeric(19,6),
        U_CHAGAM    Numeric(19,6),
        U_GITGYN    Numeric(19,6),
        U_GITYUN    Numeric(19,6),
        U_GITSGI    Numeric(19,6),
        U_GITHUS    Numeric(19,6),
        U_GITINV    Numeric(19,6),
        U_GITCAD    Numeric(19,6),
        U_GITUSJ    Numeric(19,6),
        U_GITRET    Numeric(19,6),
        U_GITJFD    Numeric(19,6),
        U_GITTOT    Numeric(19,6),
        U_TAXSTD    Numeric(19,6),
        U_SANTAX    Numeric(19,6),
        U_TAXGNL    Numeric(19,6),
        U_TAXNAB    Numeric(19,6),
        U_TAXBRO    Numeric(19,6),
        U_TAXGBU    Numeric(19,6),
        U_TAXFRG    Numeric(19,6),
        U_TAXTOT    Numeric(19,6),
        U_GAMSOD    Numeric(19,6),
        U_GAMJOS    Numeric(19,6),
        U_GAMTOT    Numeric(19,6),
        U_GULGAB    Numeric(19,6),
        U_GULJUM    Numeric(19,6),
        U_GULNON    Numeric(19,6),
        U_NANGAB    Numeric(19,6),
        U_NANJUM    Numeric(19,6),
        U_NANNON    Numeric(19,6),
        U_CHAGAB    Numeric(19,6),
        U_CHAJUM    Numeric(19,6),
        U_CHANON    Numeric(19,6),
        U_CSHSAV    Numeric(19,6),
        U_INTGBN    nvarchar(1),
        U_DWEGBN    nvarchar(1),
        U_BUYNSU    Numeric(19,6),
        U_GYNGLO    Numeric(19,6),
        U_JANGAE    Numeric(19,6),
        U_BUYN06    Numeric(19,6),
        U_DAGYSU    Numeric(19,6),
        U_CHLSAN    Numeric(19,6),
        U_PERNBR    nvarchar(20),
        U_ADDRES    nvarchar(100),
        U_CLTNAM    nvarchar(50),
        U_COMPRT    nvarchar(30),
        U_BUSNUM    nvarchar(12),
        U_PERNUM    nvarchar(14),
        U_POSADD    nvarchar(100),
        U_TAXNAM    nvarchar(20),
        U_FRGTAX    nvarchar(1),
        U_Countr    nvarchar(8),
        U_MEDAMT    Numeric(19,6),
        U_GBHAMT    Numeric(19,6),
        U_JUSAMT    Numeric(19,6),
        U_JINJ01    Numeric(19,6), 
        U_JINJ02    Numeric(19,6), 
        U_JJUS01    Numeric(19,6), 
        U_JJUS02    Numeric(19,6),
        U_YUNGON    Numeric(19,6),
        U_URIAMT    Numeric(19,6),
        U_JURI01    Numeric(19,6),
        U_JURI02    Numeric(19,6),
		U_GITGYU	Numeric(19,6),
        U_BTXCOD1   nvarchar(10),   U_BTXNAM1   nvarchar(100),  U_BTXAMT1   Numeric(19,6),  U_JBTAMT11  Numeric(19,6),   U_JBTAMT12  Numeric(19,6),
        U_BTXCOD2   nvarchar(10),   U_BTXNAM2   nvarchar(100),  U_BTXAMT2   Numeric(19,6),  U_JBTAMT21  Numeric(19,6),   U_JBTAMT22  Numeric(19,6),
        U_BTXCOD3   nvarchar(10),   U_BTXNAM3   nvarchar(100),  U_BTXAMT3   Numeric(19,6),  U_JBTAMT31  Numeric(19,6),   U_JBTAMT32  Numeric(19,6),
        U_BTXCOD4   nvarchar(10),   U_BTXNAM4   nvarchar(100),  U_BTXAMT4   Numeric(19,6),  U_JBTAMT41  Numeric(19,6),   U_JBTAMT42  Numeric(19,6),
        U_BTXCOD5   nvarchar(10),   U_BTXNAM5   nvarchar(100),  U_BTXAMT5   Numeric(19,6),  U_JBTAMT51  Numeric(19,6),   U_JBTAMT52  Numeric(19,6),
        U_BTXCOD6   nvarchar(10),   U_BTXNAM6   nvarchar(100),  U_BTXAMT6   Numeric(19,6),  U_JBTAMT61  Numeric(19,6),   U_JBTAMT62  Numeric(19,6),
        U_BTXCOD7   nvarchar(10),   U_BTXNAM7   nvarchar(100),  U_BTXAMT7   Numeric(19,6),  U_JBTAMT71  Numeric(19,6),   U_JBTAMT72  Numeric(19,6),
        U_BTXCOD8   nvarchar(10),   U_BTXNAM8   nvarchar(100),  U_BTXAMT8   Numeric(19,6),  U_JBTAMT81  Numeric(19,6),   U_JBTAMT82  Numeric(19,6),
		U_GUKNAM	nvarchar(10),
		U_GUKCOD	nvarchar(10),
		U_HUSMAN	nvarchar(8),
		U_JSNGBN	nvarchar(8),
		U_YUNGO1	Numeric(19,6),	U_YUNGO2	Numeric(19,6),	U_YUNGO3	Numeric(19,6),
		U_GITRE2	Numeric(19,6),
		U_PILJHM	Numeric(19,6),	U_PILMBH	Numeric(19,6),	U_PILGBH	Numeric(19,6),
		U_PILWOL	Numeric(19,6),
		U_GITHU1	Numeric(19,6),	U_GITHU2	Numeric(19,6),	U_GITHU3	Numeric(19,6),
        ) 

    -- 1.2) 부양가족 임시테이블
    CREATE TABLE #RPY504_1(
        U_MSTCOD    nvarchar(8),
        U_INTGBN    nvarchar(1),
        U_DWEGBN    nvarchar(1),
        U_FRGTAX    nvarchar(1),
        U_BUYNSU    Numeric(19,6),
        U_GYNGLO    Numeric(19,6),
        U_JANGAE    Numeric(19,6),
        U_BUYN06    Numeric(19,6),
        U_DAGYSU    Numeric(19,6),
        U_CHLSAN    Numeric(19,6)
        ) 

    -- 1.3) 비과세 항목1
    CREATE TABLE #RPY504_2 (
        DOCNUM  INT,
        LINEID  INT,
        LINENUM INT,
        BTXCOD  NVARCHAR(10),
        BTXNAM  NVARCHAR(100),
        BIGAMT  Numeric(19,6),
        JBTAMT1 Numeric(19,6),
        JBTAMT2 Numeric(19,6) )

    -- 1.4) 비과세 항목2
    CREATE TABLE #RPY504_3 (
        DOCNUM  INT,
        BIGM01  Numeric(19,6),  BIGM02  Numeric(19,6),  BIGM03  Numeric(19,6),  BIGO01  Numeric(19,6),
        BIGQ01  Numeric(19,6),  BIGX01  Numeric(19,6),  BIGH06  Numeric(19,6),  BIGH07  Numeric(19,6),
        BIGH08  Numeric(19,6),  BIGH09  Numeric(19,6),  BIGH10  Numeric(19,6),  BIGG01  Numeric(19,6),
        BIGH11  Numeric(19,6),  BIGH12  Numeric(19,6),  BIGH13  Numeric(19,6),  BIGH01  Numeric(19,6),
        BIGK01  Numeric(19,6),  BIGS01  Numeric(19,6),  BIGT01  Numeric(19,6),  BIGY01  Numeric(19,6),
        BIGY02  Numeric(19,6),  BIGY03  Numeric(19,6),  BIGY20  Numeric(19,6),  BIGZ01  Numeric(19,6),
        BIGH05  Numeric(19,6),  BIGI01  Numeric(19,6),
        JBTM011 Numeric(19,6),  JBTM021 Numeric(19,6),  JBTM031 Numeric(19,6),  JBTO011 Numeric(19,6),
        JBTQ011 Numeric(19,6),  JBTX011 Numeric(19,6),  JBTH061 Numeric(19,6),  JBTH071 Numeric(19,6),
        JBTH081 Numeric(19,6),  JBTH091 Numeric(19,6),  JBTH101 Numeric(19,6),  JBTG011 Numeric(19,6),
        JBTH111 Numeric(19,6),  JBTH121 Numeric(19,6),  JBTH131 Numeric(19,6),  JBTH011 Numeric(19,6),
        JBTK011 Numeric(19,6),  JBTS011 Numeric(19,6),  JBTT011 Numeric(19,6),  JBTY011 Numeric(19,6),
        JBTY021 Numeric(19,6),  JBTY031 Numeric(19,6),  JBTY201 Numeric(19,6),  JBTZ011 Numeric(19,6),
        JBTH051 Numeric(19,6),  JBTI011 Numeric(19,6),
        JBTM012 Numeric(19,6),  JBTM022 Numeric(19,6),  JBTM032 Numeric(19,6),  JBTO012 Numeric(19,6),
        JBTQ012 Numeric(19,6),  JBTX012 Numeric(19,6),  JBTH062 Numeric(19,6),  JBTH072 Numeric(19,6),
        JBTH082 Numeric(19,6),  JBTH092 Numeric(19,6),  JBTH102 Numeric(19,6),  JBTG012 Numeric(19,6),
        JBTH112 Numeric(19,6),  JBTH122 Numeric(19,6),  JBTH132 Numeric(19,6),  JBTH012 Numeric(19,6),
        JBTK012 Numeric(19,6),  JBTS012 Numeric(19,6),  JBTT012 Numeric(19,6),  JBTY012 Numeric(19,6),
        JBTY022 Numeric(19,6),  JBTY032 Numeric(19,6),  JBTY202 Numeric(19,6),  JBTZ012 Numeric(19,6),
        JBTH052 Numeric(19,6),  JBTI012 Numeric(19,6) )

    -- 1.5) 비과세 항목3
    CREATE TABLE #RPY504_4 (
        DOCNUM  INT,
        BTXCOD1 nvarchar(10),   BTXNAM1 nvarchar(100),  BTXAMT1 Numeric(19,6),  JBTAMT11 Numeric(19,6),  JBTAMT12 Numeric(19,6),
        BTXCOD2 nvarchar(10),   BTXNAM2 nvarchar(100),  BTXAMT2 Numeric(19,6),  JBTAMT21 Numeric(19,6),  JBTAMT22 Numeric(19,6),
        BTXCOD3 nvarchar(10),   BTXNAM3 nvarchar(100),  BTXAMT3 Numeric(19,6),  JBTAMT31 Numeric(19,6),  JBTAMT32 Numeric(19,6),
        BTXCOD4 nvarchar(10),   BTXNAM4 nvarchar(100),  BTXAMT4 Numeric(19,6),  JBTAMT41 Numeric(19,6),  JBTAMT42 Numeric(19,6),
        BTXCOD5 nvarchar(10),   BTXNAM5 nvarchar(100),  BTXAMT5 Numeric(19,6),  JBTAMT51 Numeric(19,6),  JBTAMT52 Numeric(19,6),
        BTXCOD6 nvarchar(10),   BTXNAM6 nvarchar(100),  BTXAMT6 Numeric(19,6),  JBTAMT61 Numeric(19,6),  JBTAMT62 Numeric(19,6),
        BTXCOD7 nvarchar(10),   BTXNAM7 nvarchar(100),  BTXAMT7 Numeric(19,6),  JBTAMT71 Numeric(19,6),  JBTAMT72 Numeric(19,6),
        BTXCOD8 nvarchar(10),   BTXNAM8 nvarchar(100),  BTXAMT8 Numeric(19,6),  JBTAMT81 Numeric(19,6),  JBTAMT82 Numeric(19,6) )

---------------------------------------------------------------------------------------------------
-- 2.1 부양가족 내역
---------------------------------------------------------------------------------------------------

    --2.1) 가족사항 가져오기
    INSERT INTO [#RPY504_1]
    SELECT  U_MSTCOD =  T1.Code, 
            U_INTGBN =  T1.U_INTGBN,
            U_DWEGBN =  T1.U_DWEGBN,
            U_FRGTAX =  T1.U_FRGTAX,
            U_BUYNSU =  ISNULL(T1.U_BUYNSU,0),
            U_GYNGLO =  ISNULL(T1.U_GYNGLO+T1.U_GYNGL2,0),
            U_JANGAE =  ISNULL(T1.U_JANGAE,0),
            U_BUYN06 =  ISNULL(T1.U_BUYN06,0),
            U_DAGYSU =  ISNULL(T1.U_DAGYSU,0),
            U_CHLSAN =  0
    FROM    [@PH_PY001A] T1 
    WHERE   T1.Code LIKE @MSTCOD

    -- 2.3) 가족사항 업데이트
    UPDATE  [#RPY504_1] 
    SET     U_BUYNSU = ISNULL(T1.U_BUYN20,0) + ISNULL(T1.U_BUYN60,0),
            U_GYNGLO = ISNULL(T1.U_GYNGLO,0) + ISNULL(T1.U_GYNGL2,0),
            U_JANGAE = ISNULL(T1.U_JANGAE,0),
            U_BUYN06 = ISNULL(T1.U_BUYN06,0),
            U_DAGYSU = ISNULL(T1.U_DAGYSU,0),
            U_CHLSAN = ISNULL(T1.U_CHLSAN,0)
    FROM    [@ZPY501H] T1 
            INNER JOIN  [#RPY504_1] T0 ON T1.U_MSTCOD COLLATE Korean_Wansung_CI_AS 
                                        = T0.U_MSTCOD COLLATE Korean_Wansung_CI_AS
    WHERE   T1.U_JSNYER =       @JSNYER
    AND     T1.U_CLTCOD LIKE    @CLTCOD
    AND     T1.U_MSTCOD LIKE    @MSTCOD

---------------------------------------------------------------------------------------------------
-- 3. 비과세 내역(비과세 내역은 비과세금액이 존재하는 9건만 뿌려준다)
---------------------------------------------------------------------------------------------------
	-- 3.1) 비과세 항목명 조회
	-- 3.1.1) 2008년 이전은 표시하지 않아도 됨
	IF @JSNYER <= '2008'
	BEGIN
		SET @BTXNAM_G01 = ''   SET @BTXNAM_H01 = ''
		SET @BTXNAM_H05 = ''   SET @BTXNAM_H06 = ''
		SET @BTXNAM_H07 = ''   SET @BTXNAM_H08 = ''
		SET @BTXNAM_H09 = ''   SET @BTXNAM_H10 = ''
		SET @BTXNAM_H11 = ''   SET @BTXNAM_H12 = ''
		SET @BTXNAM_H13 = ''   SET @BTXNAM_I01 = ''
		SET @BTXNAM_K01 = ''   SET @BTXNAM_M01 = ''
		SET @BTXNAM_M02 = ''   SET @BTXNAM_M03 = ''
		SET @BTXNAM_O01 = ''   SET @BTXNAM_Q01 = ''
		SET @BTXNAM_S01 = ''   SET @BTXNAM_T01 = ''
		SET @BTXNAM_X01 = ''   SET @BTXNAM_Y01 = ''
		SET @BTXNAM_Y02 = ''   SET @BTXNAM_Y03 = ''
		SET @BTXNAM_Y20 = ''   SET @BTXNAM_Z01 = ''
	END
	ELSE IF @JSNYER >= '2009'
	BEGIN
		-- 3.1.2) 비과세항목 설정에서 간략명칭 조회
		SELECT  @BTXNAM_G01 =  MAX(CASE WHEN T0.U_BTXCOD = 'G01' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_H01 =  MAX(CASE WHEN T0.U_BTXCOD = 'H01' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_H05 =  MAX(CASE WHEN T0.U_BTXCOD = 'H05' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_H06 =  MAX(CASE WHEN T0.U_BTXCOD = 'H06' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_H07 =  MAX(CASE WHEN T0.U_BTXCOD = 'H07' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_H08 =  MAX(CASE WHEN T0.U_BTXCOD = 'H08' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_H09 =  MAX(CASE WHEN T0.U_BTXCOD = 'H09' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_H10 =  MAX(CASE WHEN T0.U_BTXCOD = 'H10' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_H11 =  MAX(CASE WHEN T0.U_BTXCOD = 'H11' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_H12 =  MAX(CASE WHEN T0.U_BTXCOD = 'H12' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_H13 =  MAX(CASE WHEN T0.U_BTXCOD = 'H13' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_I01 =  MAX(CASE WHEN T0.U_BTXCOD = 'I01' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_K01 =  MAX(CASE WHEN T0.U_BTXCOD = 'K01' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_M01 =  MAX(CASE WHEN T0.U_BTXCOD = 'M01' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_M02 =  MAX(CASE WHEN T0.U_BTXCOD = 'M02' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_M03 =  MAX(CASE WHEN T0.U_BTXCOD = 'M03' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_O01 =  MAX(CASE WHEN T0.U_BTXCOD = 'O01' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_Q01 =  MAX(CASE WHEN T0.U_BTXCOD = 'Q01' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_S01 =  MAX(CASE WHEN T0.U_BTXCOD = 'S01' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_T01 =  MAX(CASE WHEN T0.U_BTXCOD = 'T01' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_X01 =  MAX(CASE WHEN T0.U_BTXCOD = 'X01' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_Y01 =  MAX(CASE WHEN T0.U_BTXCOD = 'Y01' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_Y02 =  MAX(CASE WHEN T0.U_BTXCOD = 'Y02' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_Y03 =  MAX(CASE WHEN T0.U_BTXCOD = 'Y03' THEN U_BTXNAM ELSE '' END),
				@BTXNAM_Y20 =  MAX(CASE WHEN T0.U_BTXCOD = 'Y20' THEN U_BTXNAM ELSE '' END), 
				@BTXNAM_Z01 =  MAX(CASE WHEN T0.U_BTXCOD = 'Z01' THEN U_BTXNAM ELSE '' END)
		FROM    [@ZPY117L] T0 
		WHERE   T0.CODE = (SELECT MAX(CODE) FROM [@ZPY117L] T1 WHERE CODE <= @JSNYER) 
		AND     T0.CODE >= '2009' 

		-- 3.1.3) 만약 비과세 항목설정이 등록되지 않은 경우 내부적으로 갖고 있는 기본값으로 뿌린다
		SET @BTXNAM_G01 =  ISNULL(@BTXNAM_G01,N'학자금')
		SET @BTXNAM_H01 =  ISNULL(@BTXNAM_H01,N'무보수위원')
		SET @BTXNAM_H05 =  ISNULL(@BTXNAM_H05,N'경호.승선수당')
		SET @BTXNAM_H06 =  ISNULL(@BTXNAM_H06,N'연구보조비')
		SET @BTXNAM_H07 =  ISNULL(@BTXNAM_H07,N'연구보조비')
		SET @BTXNAM_H08 =  ISNULL(@BTXNAM_H08,N'연구보조비')
		SET @BTXNAM_H09 =  ISNULL(@BTXNAM_H09,N'연구보조비')
		SET @BTXNAM_H10 =  ISNULL(@BTXNAM_H10,N'연구보조비')
		SET @BTXNAM_H11 =  ISNULL(@BTXNAM_H11,N'취재수당')
		SET @BTXNAM_H12 =  ISNULL(@BTXNAM_H12,N'벽지수당')
		SET @BTXNAM_H13 =  ISNULL(@BTXNAM_H13,N'재해급여')
		SET @BTXNAM_I01 =  ISNULL(@BTXNAM_I01,N'외국정부근무자')
		SET @BTXNAM_K01 =  ISNULL(@BTXNAM_K01,N'외국주둔군인')
		SET @BTXNAM_M01 =  ISNULL(@BTXNAM_M01,N'국외근로')
		SET @BTXNAM_M02 =  ISNULL(@BTXNAM_M02,N'국외근로')
		SET @BTXNAM_M03 =  ISNULL(@BTXNAM_M03,N'국외근로')
		SET @BTXNAM_O01 =  ISNULL(@BTXNAM_O01,N'야간근로수당')
		SET @BTXNAM_Q01 =  ISNULL(@BTXNAM_Q01,N'보육수당')
		SET @BTXNAM_S01 =  ISNULL(@BTXNAM_S01,N'주식매수선택권')
		SET @BTXNAM_T01 =  ISNULL(@BTXNAM_T01,N'외국인기술자')
		SET @BTXNAM_X01 =  ISNULL(@BTXNAM_X01,N'외국인근로자')
		SET @BTXNAM_Y01 =  ISNULL(@BTXNAM_Y01,N'우리사주 배정')
		SET @BTXNAM_Y02 =  ISNULL(@BTXNAM_Y02,N'우리사주 인출')
		SET @BTXNAM_Y03 =  ISNULL(@BTXNAM_Y03,N'우리사주 인출')
		SET @BTXNAM_Y20 =  ISNULL(@BTXNAM_Y20,N'주택자금보조금')
		SET @BTXNAM_Z01 =  ISNULL(@BTXNAM_Z01,N'해저광물개발')

	END

    -- 3.2) 비과세내역 조회 
    INSERT  INTO [#RPY504_3]
    SELECT  T0.DOCENTRY,
            ISNULL(T0.U_BIGM01,0),  ISNULL(T0.U_BIGM02,0),  ISNULL(T0.U_BIGM03,0), ISNULL(T0.U_BIGO01,0),
            ISNULL(T0.U_BIGQ01,0),  ISNULL(T0.U_BIGX01,0),  ISNULL(T0.U_BIGH06,0), ISNULL(T0.U_BIGH07,0),
            ISNULL(T0.U_BIGH08,0),  ISNULL(T0.U_BIGH09,0),  ISNULL(T0.U_BIGH10,0), ISNULL(T0.U_BIGG01,0),
            ISNULL(T0.U_BIGH11,0),  ISNULL(T0.U_BIGH12,0),  ISNULL(T0.U_BIGH13,0), ISNULL(T0.U_BIGH01,0),
            ISNULL(T0.U_BIGK01,0),  ISNULL(T0.U_BIGS01,0),  ISNULL(T0.U_BIGT01,0), ISNULL(T0.U_BIGY01,0),
            ISNULL(T0.U_BIGY02,0),  ISNULL(T0.U_BIGY03,0),  ISNULL(T0.U_BIGY20,0), ISNULL(T0.U_BIGZ01,0),
            ISNULL(T0.U_BIGH05,0),  ISNULL(T0.U_BIGI01,0),

            ISNULL(T4.JBTM011,0),   ISNULL(T4.JBTM021,0),   ISNULL(T4.JBTM031,0),   ISNULL(T4.JBTO011,0),
            ISNULL(T4.JBTQ011,0),   ISNULL(T4.JBTX011,0),   ISNULL(T4.JBTH061,0),   ISNULL(T4.JBTH071,0),
            ISNULL(T4.JBTH081,0),   ISNULL(T4.JBTH091,0),   ISNULL(T4.JBTH101,0),   ISNULL(T4.JBTG011,0),
            ISNULL(T4.JBTH111,0),   ISNULL(T4.JBTH121,0),   ISNULL(T4.JBTH131,0),   ISNULL(T4.JBTH011,0),
            ISNULL(T4.JBTK011,0),   ISNULL(T4.JBTS011,0),   ISNULL(T4.JBTT011,0),   ISNULL(T4.JBTY011,0),
            ISNULL(T4.JBTY021,0),   ISNULL(T4.JBTY031,0),   ISNULL(T4.JBTY201,0),   ISNULL(T4.JBTZ011,0),
            ISNULL(T4.JBTH051,0),   ISNULL(T4.JBTI011,0),
            
            ISNULL(T4.JBTM012,0),   ISNULL(T4.JBTM022,0),   ISNULL(T4.JBTM032,0),   ISNULL(T4.JBTO012,0),
            ISNULL(T4.JBTQ012,0),   ISNULL(T4.JBTX012,0),   ISNULL(T4.JBTH062,0),   ISNULL(T4.JBTH072,0),
            ISNULL(T4.JBTH082,0),   ISNULL(T4.JBTH092,0),   ISNULL(T4.JBTH102,0),   ISNULL(T4.JBTG012,0),
            ISNULL(T4.JBTH112,0),   ISNULL(T4.JBTH122,0),   ISNULL(T4.JBTH132,0),   ISNULL(T4.JBTH012,0),
            ISNULL(T4.JBTK012,0),   ISNULL(T4.JBTS012,0),   ISNULL(T4.JBTT012,0),   ISNULL(T4.JBTY012,0),
            ISNULL(T4.JBTY022,0),   ISNULL(T4.JBTY032,0),   ISNULL(T4.JBTY202,0),   ISNULL(T4.JBTZ012,0),
            ISNULL(T4.JBTH052,0),   ISNULL(T4.JBTI012,0)

    FROM    [@ZPY504H] T0 
            --INNER JOIN [OHEM]       T2 ON T0.U_EmpID  = T2.EmpID
            INNER JOIN [@PH_PY001A] T2 ON T0.U_EmpID = T2.U_EmpID
            --INNER JOIN [OUDP]       T3 ON T2.Dept     = T3.Code
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
                    JBTZ011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTZ01,0) ELSE 0 END),
                    JBTH051 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTH05,0) ELSE 0 END), 
                    JBTI011 = SUM(CASE WHEN A1.U_LINENUM = '1' THEN ISNULL(A1.U_JBTI01,0) ELSE 0 END),

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
                    JBTZ012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTZ01,0) ELSE 0 END),
                    JBTH052 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTH05,0) ELSE 0 END), 
                    JBTI012 = SUM(CASE WHEN A1.U_LINENUM = '2' THEN ISNULL(A1.U_JBTI01,0) ELSE 0 END)
            FROM    [@ZPY502H] A0
                    INNER JOIN [@ZPY502L] A1 ON A0.DOCENTRY = A1.DOCENTRY
            WHERE   A0.U_JSNYER = @JSNYER
            AND     A0.U_MSTCOD LIKE @MSTCOD 
            GROUP   BY A0.U_MSTCOD,  A0.U_JSNYER, A0.U_CLTCOD
            ) T4 ON T0.U_JSNYER = T4.U_JSNYER AND T0.U_MSTCOD = T4.U_MSTCOD AND T0.U_CLTCOD = T4.U_CLTCOD
    WHERE   T0.U_JSNYER =    @JSNYER 
    AND     T0.U_CLTCOD LIKE @CLTCOD 
    AND     T2.U_TeamCode LIKE @MSTDPT                        
    AND     T0.U_MSTCOD LIKE @MSTCOD 
    --AND     ISNULL(CONVERT(Nvarchar(8),T2.Branch), '')  LIKE @Branch   
    AND     (@JOBGBN ='3' OR (@JOBGBN <> '3' AND T0.U_JSNGBN = @JOBGBN)) 
    AND     T0.U_JSNMON  BETWEEN @STRMON AND @ENDMON

    -- 3.3) 지급명세서 출력순서에 따라 하나씩 임시테이블에 저장(열 -> 행으로 변경)
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 1,  0, 'M01', @BTXNAM_M01, T0.BIGM01, T0.JBTM011, T0.JBTM012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 2,  0, 'M02', @BTXNAM_M02, T0.BIGM02, T0.JBTM021, T0.JBTM022 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 3,  0, 'M03', @BTXNAM_M03, T0.BIGM03, T0.JBTM031, T0.JBTM032 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 4,  0, 'O01', @BTXNAM_O01, T0.BIGO01, T0.JBTO011, T0.JBTO012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 5,  0, 'Q01', @BTXNAM_Q01, T0.BIGQ01, T0.JBTQ011, T0.JBTQ012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 6,  0, 'X01', @BTXNAM_X01, T0.BIGX01, T0.JBTX011, T0.JBTX012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 7,  0, 'H06', @BTXNAM_H06, T0.BIGH06, T0.JBTH061, T0.JBTH062 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 8,  0, 'H07', @BTXNAM_H07, T0.BIGH07, T0.JBTH071, T0.JBTH072 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 9,  0, 'H08', @BTXNAM_H08, T0.BIGH08, T0.JBTH081, T0.JBTH082 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 10, 0, 'H09', @BTXNAM_H09, T0.BIGH09, T0.JBTH091, T0.JBTH092 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 11, 0, 'H10', @BTXNAM_H10, T0.BIGH10, T0.JBTH101, T0.JBTH102 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 12, 0, 'G01', @BTXNAM_G01, T0.BIGG01, T0.JBTG011, T0.JBTG012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 13, 0, 'H11', @BTXNAM_H11, T0.BIGH11, T0.JBTH111, T0.JBTH112 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 14, 0, 'H12', @BTXNAM_H12, T0.BIGH12, T0.JBTH121, T0.JBTH122 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 15, 0, 'H13', @BTXNAM_H13, T0.BIGH13, T0.JBTH131, T0.JBTH132 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 16, 0, 'H01', @BTXNAM_H01, T0.BIGH01, T0.JBTH011, T0.JBTH012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 17, 0, 'K01', @BTXNAM_K01, T0.BIGK01, T0.JBTK011, T0.JBTK012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 18, 0, 'S01', @BTXNAM_S01, T0.BIGS01, T0.JBTS011, T0.JBTS012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 19, 0, 'T01', @BTXNAM_T01, T0.BIGT01, T0.JBTT011, T0.JBTT012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 20, 0, 'Y01', @BTXNAM_Y01, T0.BIGY01, T0.JBTY011, T0.JBTY012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 21, 0, 'Y02', @BTXNAM_Y02, T0.BIGY02, T0.JBTY021, T0.JBTY022 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 22, 0, 'Y03', @BTXNAM_Y03, T0.BIGY03, T0.JBTY031, T0.JBTY032 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 23, 0, 'Y20', @BTXNAM_Y20, T0.BIGY20, T0.JBTY201, T0.JBTY202 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 24, 0, 'Z01', @BTXNAM_Z01, T0.BIGZ01, T0.JBTZ011, T0.JBTZ012 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 25, 0, 'H05', @BTXNAM_H05, T0.BIGH05, T0.JBTH051, T0.JBTH052 FROM [#RPY504_3] T0
    INSERT  INTO [#RPY504_2] SELECT  T0.DOCNUM, 26, 0, 'I01', @BTXNAM_I01, T0.BIGI01, T0.JBTI011, T0.JBTI012 FROM [#RPY504_3] T0

    -- 3.4) 금액이 존재하지 않는 비과세내역 삭제
    DELETE  FROM [#RPY504_2]
    WHERE   BIGAMT  <= 0
    AND     JBTAMT1 <= 0
    AND     JBTAMT2 <= 0

    -- 3.5) 순서정렬
    UPDATE  T0
    SET     LINENUM =   (SELECT COUNT(*) FROM [#RPY504_2] T1 WHERE T0.DOCNUM = T1.DOCNUM AND T0.LINEID >= T1.LINEID)
    FROM    [#RPY504_2] T0

    -- 3.6) 사원별로 1줄로 만든다
    -- (SubReport로 만들 수 있으나, 향후 서식이 어떻게 바뀔지 장담할 수 없으므로 한정된 칸에 대해서만 데이터 채움)
    INSERT  INTO [#RPY504_4]
    SELECT  T0.DOCNUM,
            BTXCOD1  =   MAX(CASE WHEN T0.LINENUM = 1 THEN T0.BTXCOD  ELSE NULL END),
            BTXNAM1  =   MAX(CASE WHEN T0.LINENUM = 1 THEN T0.BTXNAM  ELSE NULL END),
            BTXAMT1  =   SUM(CASE WHEN T0.LINENUM = 1 THEN T0.BIGAMT  ELSE 0 END),
            JBTAMT11 =   SUM(CASE WHEN T0.LINENUM = 1 THEN T0.JBTAMT1 ELSE 0 END),
            JBTAMT12 =   SUM(CASE WHEN T0.LINENUM = 1 THEN T0.JBTAMT2 ELSE 0 END),

            BTXCOD2  =   MAX(CASE WHEN T0.LINENUM = 2 THEN T0.BTXCOD  ELSE NULL END),
            BTXNAM2  =   MAX(CASE WHEN T0.LINENUM = 2 THEN T0.BTXNAM  ELSE NULL END),
            BTXAMT2  =   SUM(CASE WHEN T0.LINENUM = 2 THEN T0.BIGAMT  ELSE 0 END),
            JBTAMT21 =   SUM(CASE WHEN T0.LINENUM = 2 THEN T0.JBTAMT1 ELSE 0 END),
            JBTAMT22 =   SUM(CASE WHEN T0.LINENUM = 2 THEN T0.JBTAMT2 ELSE 0 END),

            BTXCOD3  =   MAX(CASE WHEN T0.LINENUM = 3 THEN T0.BTXCOD  ELSE NULL END),
            BTXNAM3  =   MAX(CASE WHEN T0.LINENUM = 3 THEN T0.BTXNAM  ELSE NULL END),
            BTXAMT3  =   SUM(CASE WHEN T0.LINENUM = 3 THEN T0.BIGAMT  ELSE 0 END),
            JBTAMT31 =   SUM(CASE WHEN T0.LINENUM = 3 THEN T0.JBTAMT1 ELSE 0 END),
            JBTAMT32 =   SUM(CASE WHEN T0.LINENUM = 3 THEN T0.JBTAMT2 ELSE 0 END),

            BTXCOD4  =   MAX(CASE WHEN T0.LINENUM = 4 THEN T0.BTXCOD  ELSE NULL END),
            BTXNAM4  =   MAX(CASE WHEN T0.LINENUM = 4 THEN T0.BTXNAM  ELSE NULL END),
            BTXAMT4  =   SUM(CASE WHEN T0.LINENUM = 4 THEN T0.BIGAMT  ELSE 0 END),
            JBTAMT41 =   SUM(CASE WHEN T0.LINENUM = 4 THEN T0.JBTAMT1 ELSE 0 END),
            JBTAMT42 =   SUM(CASE WHEN T0.LINENUM = 4 THEN T0.JBTAMT2 ELSE 0 END),

            BTXCOD5  =   MAX(CASE WHEN T0.LINENUM = 5 THEN T0.BTXCOD  ELSE NULL END),
            BTXNAM5  =   MAX(CASE WHEN T0.LINENUM = 5 THEN T0.BTXNAM  ELSE NULL END),
            BTXAMT5  =   SUM(CASE WHEN T0.LINENUM = 5 THEN T0.BIGAMT  ELSE 0 END),
            JBTAMT51 =   SUM(CASE WHEN T0.LINENUM = 5 THEN T0.JBTAMT1 ELSE 0 END),
            JBTAMT52 =   SUM(CASE WHEN T0.LINENUM = 5 THEN T0.JBTAMT2 ELSE 0 END),

            BTXCOD6  =   MAX(CASE WHEN T0.LINENUM = 6 THEN T0.BTXCOD  ELSE NULL END),
            BTXNAM6  =   MAX(CASE WHEN T0.LINENUM = 6 THEN T0.BTXNAM  ELSE NULL END),
            BTXAMT6  =   SUM(CASE WHEN T0.LINENUM = 6 THEN T0.BIGAMT  ELSE 0 END),
            JBTAMT61 =   SUM(CASE WHEN T0.LINENUM = 6 THEN T0.JBTAMT1 ELSE 0 END),
            JBTAMT62 =   SUM(CASE WHEN T0.LINENUM = 6 THEN T0.JBTAMT2 ELSE 0 END),

            BTXCOD7  =   MAX(CASE WHEN T0.LINENUM = 7 THEN T0.BTXCOD  ELSE NULL END),
            BTXNAM7  =   MAX(CASE WHEN T0.LINENUM = 7 THEN T0.BTXNAM  ELSE NULL END),
            BTXAMT7  =   SUM(CASE WHEN T0.LINENUM = 7 THEN T0.BIGAMT  ELSE 0 END),
            JBTAMT71 =   SUM(CASE WHEN T0.LINENUM = 7 THEN T0.JBTAMT1 ELSE 0 END),
            JBTAMT72 =   SUM(CASE WHEN T0.LINENUM = 7 THEN T0.JBTAMT2 ELSE 0 END),

            BTXCOD8  =   MAX(CASE WHEN T0.LINENUM = 8 THEN T0.BTXCOD  ELSE NULL END),
            BTXNAM8  =   MAX(CASE WHEN T0.LINENUM = 8 THEN T0.BTXNAM  ELSE NULL END),
            BTXAMT8  =   SUM(CASE WHEN T0.LINENUM = 8 THEN T0.BIGAMT  ELSE 0 END),
            JBTAMT81 =   SUM(CASE WHEN T0.LINENUM = 8 THEN T0.JBTAMT1 ELSE 0 END),
            JBTAMT82 =   SUM(CASE WHEN T0.LINENUM = 8 THEN T0.JBTAMT2 ELSE 0 END)

    FROM    [#RPY504_2] T0
    GROUP   BY T0.DOCNUM
    ORDER   BY T0.DOCNUM

---------------------------------------------------------------------------------------------------
-- 4.연말정산 자료 총괄 조회
---------------------------------------------------------------------------------------------------

    INSERT INTO [#RPY504]
    SELECT  U_MSTCOD    =   T0.U_MSTCOD,
            U_MSTNAM    =   T0.U_MSTNAM,
            U_CLTCOD    =   T0.U_CLTCOD,
			U_MSTDPT	=	T2.U_TeamCode,
			U_DPTNAM	=	(SELECT U_codeNM FROM [@PS_HR200L] WHERE Code = '1' AND U_Code = T2.U_TeamCode),
            U_STRINT    =   CASE WHEN T0.U_CLTCOD <> T2.U_CLTCOD THEN CONVERT(CHAR(10),T0.U_STRINT,120)
								 WHEN T2.startDate > T0.U_STRINT THEN CONVERT(CHAR(10),T2.startDate,120) 
                                 ELSE CONVERT(CHAR(10),T0.U_STRINT,120) END,
            U_ENDINT    =   CONVERT(CHAR(10),T0.U_ENDINT,120),
            U_STRGAM    =   CONVERT(CHAR(10),T0.U_STRGAM,120),
            U_ENDGAM    =   CONVERT(CHAR(10),T0.U_ENDGAM,120),
            U_JSTR01    =   ISNULL(T7.JSTR01, ''),
            U_JEND01    =   ISNULL(T7.JEND01, ''),
            U_JGFR01    =   ISNULL(T7.JGFR01, ''),
            U_JGTO01    =   ISNULL(T7.JGTO01, ''),
            U_JSTR02    =   ISNULL(T7.JSTR02, ''),
            U_JEND02    =   ISNULL(T7.JEND02, ''),
            U_JGFR02    =   ISNULL(T7.JGFR02, ''),
            U_JGTO02    =   ISNULL(T7.JGTO02, ''),
            U_J01NAM    =   T0.U_J01NAM,
            U_J02NAM    =   T0.U_J02NAM,
            U_J01NBR    =   T0.U_J01NBR,
            U_J02NBR    =   T0.U_J02NBR,
            U_JPAY01    =   T0.U_JPAY01,
            U_JBNS01    =   T0.U_JBNS01,
            U_JPAY02    =   T0.U_JPAY02,
            U_JBNS02    =   T0.U_jBNS02,
            U_JONGAB    =   T0.U_JONGAB,
            U_JONJUM    =   T0.U_JONJUM,
            U_JONNON    =   T0.U_JONNON,
            U_JONGA1    =   ISNULL(T7.JONGA1,0),
            U_JONJU1    =   ISNULL(T7.JONJU1,0),
            U_JONNO1    =   ISNULL(T7.JONNO1,0),
            U_JONGA2    =   ISNULL(T7.JONGA2,0),
            U_JONJU2    =   ISNULL(T7.JONJU2,0),
            U_JONNO2    =   ISNULL(T7.JONNO2,0),
            U_PAYAMT    =   T0.U_PAYAMT,
            U_BNSAMT    =   T0.U_BNSAMT,
            U_INBAMT    =   T0.U_INBAMT,
            U_TOTAMT    =   T0.U_TOTAMT,
            U_BIGWA1    =   T0.U_BIGWA1,
            U_BIGWA2    =   T0.U_BIGWA2,
            U_BIGWA3    =   T0.U_BIGWA3,
            U_BIGWU3    =   ISNULL(T0.U_BIGWU3,0),
            U_BIGWA4    =   ISNULL(T0.U_BIGWA4,0),
            U_BIGWA5    =   ISNULL(T0.U_BIGWA5,0),
            U_BIGWA6    =   ISNULL(T0.U_BIGWA6,0),
            U_BIGTOT    =   T0.U_BIGTOT,
            U_INCOME    =   T0.U_INCOME,
            U_PILGNL    =   T0.U_PILGNL,
            U_GNLOSD    =   T0.U_GNLOSD,
            U_INJBAS    =   T0.U_INJBAS,
            U_INJBWO    =   T0.U_INJBWO,
            U_INJBYN    =   T0.U_INJBYN,
            U_INJGYN    =   T0.U_INJGYN,
            U_INJJAE    =   T0.U_INJJAE,
            U_INJBNJ    =   T0.U_INJBNJ,
            U_INJSON    =   T0.U_INJSON,
            U_INJADD    =   T0.U_INJADD,
            U_INJCHL    =   ISNULL(T0.U_INJCHL,0),
            U_KUKGON    =   T0.U_KUKGON,
            U_PILBHM    =   T0.U_PILBHM,
            U_PILMED    =   T0.U_PILMED,
            U_PILSCH    =   T0.U_PILSCH,
            U_PILHUS    =   ISNULL(T0.U_PILHUS,0),
            U_PILJHE    =   T0.U_PILJHE,
            U_PILGBU    =   T0.U_PILGBU,
            U_PILHUN    =   T0.U_PILHUN,
            U_PILTOT    =   T0.U_PILTOT,
            U_PILGON    =   T0.U_PILGON,
            U_CHAGAM    =   T0.U_CHAGAM,
            U_GITGYN    =   T0.U_GITGYN,
            U_GITYUN    =   T0.U_GITYUN,
            U_GITSGI    =   ISNULL(T0.U_GITSGI, 0),
            U_GITHUS    =   ISNULL(T0.U_GITHUS, 0),
            U_GITINV    =   T0.U_GITINV,
            U_GITCAD    =   T0.U_GITCAD,
            U_GITUSJ    =   T0.U_GITUSJ,
            U_GITRET    =   T0.U_GITRET,
            U_GITJFD    =   ISNULL(T0.U_GITJFD, 0),
            U_GITTOT    =   T0.U_GITTOT,
            U_TAXSTD    =   T0.U_TAXSTD,
            U_SANTAX    =   T0.U_SANTAX,
            U_TAXGNL    =   T0.U_TAXGNL,
            U_TAXNAB    =   T0.U_TAXNAB,
            U_TAXBRO    =   T0.U_TAXBRO,
            U_TAXGBU    =   T0.U_TAXGBU,
            U_TAXFRG    =   T0.U_TAXFRG,
            U_TAXTOT    =   T0.U_TAXTOT,
            U_GAMSOD    =   T0.U_GAMSOD,
            U_GAMJOS    =   T0.U_GAMJOS,
            U_GAMTOT    =   T0.U_GAMTOT,
            U_GULGAB    =   T0.U_GULGAB,
            U_GULJUM    =   T0.U_GULJUM,
            U_GULNON    =   T0.U_GULNON,
            U_NANGAB    =   T0.U_NANGAB,
            U_NANJUM    =   T0.U_NANJUM,
            U_NANNON    =   T0.U_NANNON,
            U_CHAGAB    =   T0.U_CHAGAB,
            U_CHAJUM    =   T0.U_CHAJUM,
            U_CHANON    =   T0.U_CHANON,
            U_CSHSAV    =   T0.U_CSHSAV,
            U_INTGBN    =   T1.U_INTGBN,
            U_DWEGBN    =   T1.U_DWEGBN,
            U_BUYNSU    =   ISNULL(T1.U_BUYNSU,0),
            U_GYNGLO    =   ISNULL(T1.U_GYNGLO,0),
            U_JANGAE    =   ISNULL(T1.U_JANGAE,0),
            U_BUYN06    =   ISNULL(T1.U_BUYN06,0),
            U_DAGYSU    =   ISNULL(T1.U_DAGYSU,0),
            U_CHLSAN    =   ISNULL(T1.U_CHLSAN,0),
            U_PERNBR    =   ISNULL(T2.GovID, ''),
            U_ADDRES    =   ISNULL(T2.HomeStreet, ''),
            U_CLTNAM    =   ISNULL(T4.U_CLTName, ''),
            U_COMPRT    =   ISNULL(T4.U_ComPrt, ''),
            U_BUSNUM    =   ISNULL(T4.U_BusNum, ''),
            U_PERNUM    =   ISNULL(T4.U_PerNum, ''),
            U_POSADD    =   ISNULL(T4.U_PosAdd, ''),
            U_TAXNAM    =   ISNULL(T4.U_TAXName, ''),
            U_FRGTAX    =   T1.U_FRGTAX,    
            U_Countr    =   T2.HomeCountr,
            U_MEDAMT    =   ISNULL(T5.U_MEDAMT, 0),
            U_GBHAMT    =   ISNULL(T5.U_GBHAMT, 0),
            U_JUSAMT    =   ISNULL(T0.U_JUSAMT,0),
            U_JINJ01    =   ISNULL(T0.U_JINJ01,0), 
            U_JINJ02    =   ISNULL(T0.U_JINJ02,0), 
            U_JJUS01    =   ISNULL(T0.U_JJUS01,0), 
            U_JJUS02    =   ISNULL(T0.U_JJUS02,0),
            U_YUNGON    =   ISNULL(T0.U_YUNGON,0),
            U_URIAMT    =   ISNULL(T0.U_URIAMT,0),
            U_JURI01    =   ISNULL(T0.U_JURI01,0),
            U_JURI02    =   ISNULL(T0.U_JURI02,0),
			U_GITGYU	=	ISNULL(T0.U_GITGYU,0),

            U_BTXCOD1   =   ISNULL(T6.BTXCOD1,''),
            U_BTXNAM1   =   ISNULL(T6.BTXNAM1,''),
            U_BTXAMT1   =   ISNULL(T6.BTXAMT1,0),
            U_JBTAMT11  =   ISNULL(T6.JBTAMT11,0),
            U_JBTAMT12  =   ISNULL(T6.JBTAMT12,0),

            U_BTXCOD2   =   ISNULL(T6.BTXCOD2,''),
            U_BTXNAM2   =   ISNULL(T6.BTXNAM2,''),
            U_BTXAMT2   =   ISNULL(T6.BTXAMT2,0),
            U_JBTAMT21  =   ISNULL(T6.JBTAMT21,0),
            U_JBTAMT22  =   ISNULL(T6.JBTAMT22,0),

            U_BTXCOD3   =   ISNULL(T6.BTXCOD3,''),
            U_BTXNAM3   =   ISNULL(T6.BTXNAM3,''),
            U_BTXAMT3   =   ISNULL(T6.BTXAMT3,0),
            U_JBTAMT31  =   ISNULL(T6.JBTAMT31,0),
            U_JBTAMT32  =   ISNULL(T6.JBTAMT32,0),

            U_BTXCOD4   =   ISNULL(T6.BTXCOD4,''),
            U_BTXNAM4   =   ISNULL(T6.BTXNAM4,''),
            U_BTXAMT4   =   ISNULL(T6.BTXAMT4,0),
            U_JBTAMT41  =   ISNULL(T6.JBTAMT41,0),
            U_JBTAMT42  =   ISNULL(T6.JBTAMT42,0),

            U_BTXCOD5   =   ISNULL(T6.BTXCOD5,''),
            U_BTXNAM5   =   ISNULL(T6.BTXNAM5,''),
            U_BTXAMT5   =   ISNULL(T6.BTXAMT5,0),
            U_JBTAMT51  =   ISNULL(T6.JBTAMT51,0),
            U_JBTAMT52  =   ISNULL(T6.JBTAMT52,0),

            U_BTXCOD6   =   ISNULL(T6.BTXCOD6,''),
            U_BTXNAM6   =   ISNULL(T6.BTXNAM6,''),
            U_BTXAMT6   =   ISNULL(T6.BTXAMT6,0),
            U_JBTAMT61  =   ISNULL(T6.JBTAMT61,0),
            U_JBTAMT62  =   ISNULL(T6.JBTAMT62,0),

            U_BTXCOD7   =   ISNULL(T6.BTXCOD7,''),
            U_BTXNAM7   =   ISNULL(T6.BTXNAM7,''),
            U_BTXAMT7   =   ISNULL(T6.BTXAMT7,0),
            U_JBTAMT71  =   ISNULL(T6.JBTAMT71,0),
            U_JBTAMT72  =   ISNULL(T6.JBTAMT72,0),

            U_BTXCOD8   =   ISNULL(T6.BTXCOD8,''),
            U_BTXNAM8   =   ISNULL(T6.BTXNAM8,''),
            U_BTXAMT8   =   ISNULL(T6.BTXAMT8,0),
            U_JBTAMT81  =   ISNULL(T6.JBTAMT81,0),
            U_JBTAMT82  =   ISNULL(T6.JBTAMT82,0),
			U_GUKNAM	=	CASE WHEN T1.U_INTGBN = '1' THEN '' ELSE (SELECT A0.Name FROM OCRY A0 WHERE A0.Code = ISNULL(T2.citizenshp, '')) END,
			U_GUKCOD	=	CASE WHEN T1.U_INTGBN = '1' THEN '' ELSE ISNULL(T2.citizenshp, '') END,
			U_HUSMAN	=	ISNULL(T8.U_HUSMAN,'2'),
			U_JSNGBN	=	ISNULL(T0.U_JSNGBN,'2'),
			U_YUNGO1	=	ISNULL(T0.U_YUNGO1,0),
			U_YUNGO2	=	ISNULL(T0.U_YUNGO2,0),
			U_YUNGO3	=	ISNULL(T0.U_YUNGO3,0),
			U_GITRE2	=	ISNULL(T0.U_GITRE2,0),
			U_PILJHM	=	ISNULL(T0.U_PILJHM,0),
			U_PILMBH	=	ISNULL(T0.U_PILMBH,0),
			U_PILGBH	=	ISNULL(T0.U_PILGBH,0),
			U_PILWOL	=	ISNULL(T0.U_PILWOL,0),
			U_GITHU1	=	ISNULL(T0.U_GITHU1,0),
			U_GITHU2	=	ISNULL(T0.U_GITHU2,0),
			U_GITHU3	=	ISNULL(T0.U_GITHU3,0)

    FROM    [@ZPY504H] T0   
            INNER JOIN [#RPY504_1]  T1 ON T0.U_MSTCOD  COLLATE Korean_Wansung_CI_AS = T1.U_MSTCOD COLLATE Korean_Wansung_CI_AS
            --INNER JOIN [OHEM]       T2 ON T0.U_EmpID  = T2.EmpID
            INNER JOIN [@PH_PY001A] T2 ON T0.U_EmpID = T2.EmpID
            --INNER JOIN [OUDP]       T3 ON T2.Dept     = T3.Code
            INNER JOIN [@ZPY501H]   T8 ON T0.U_MSTCOD = T8.U_MSTCOD AND T0.U_JSNYER = T8.U_JSNYER AND T0.U_CLTCOD = T8.U_CLTCOD
            --LEFT  JOIN [@ZPY106H]   T4 ON T0.U_CLTCOD = T4.U_CLTCode
            LEFT  JOIN [@PH_PY005A] T4 ON T0.U_CLTCOD = T4.U_CLTCode
            LEFT  JOIN [#RPY504_4]  T6 ON T0.DOCENTRY = T6.DOCNUM
            LEFT  JOIN (	-- 건강보험, 고용보험
            SELECT  U_JSNYER    =   T1.U_JsnYear, 
                    U_CLTCOD    =   T1.U_CLTCOD,
                    U_MSTCOD    =   T1.U_MstCode,
                    U_MEDAMT    =   T0.U_MEDAMT + ISNULL(U_NGYAMT,0), 
                    U_GBHAMT    =   T0.U_GBHAMT 
            FROM    [@ZPY343L] T0 
                    INNER JOIN [@ZPY343H] T1 ON T0.DocEntry = T1.DocEntry   AND T0.U_LineNum = '13'
            WHERE   T1.U_JsnYear   = @JSNYER 
            AND     T1.U_CLTCOD LIKE @CLTCOD
            ) T5 ON T0.U_JSNYER = T5.U_JSNYER AND T0.U_MSTCOD = T5.U_MSTCOD AND T0.U_CLTCOD = T5.U_CLTCOD
            LEFT  JOIN (	-- 종(전)근무지 근무기간, 감면기간
            SELECT  A0.U_MSTCOD,
                    A0.U_JSNYER,
                    A0.U_CLTCOD,
                    JSTR01 = MAX(CASE WHEN A1.U_LINENUM = '1' THEN CONVERT(CHAR(10),A1.U_JONSTR,120) ELSE '' END), 
                    JEND01 = MAX(CASE WHEN A1.U_LINENUM = '1' THEN CONVERT(CHAR(10),A1.U_JONEND,120) ELSE '' END), 
                    JGFR01 = MAX(CASE WHEN A1.U_LINENUM = '1' THEN CONVERT(CHAR(10),A1.U_JONGFR,120) ELSE '' END), 
                    JGTO01 = MAX(CASE WHEN A1.U_LINENUM = '1' THEN CONVERT(CHAR(10),A1.U_JONGTO,120) ELSE '' END),
                    JONGA1 = MAX(CASE WHEN A1.U_LINENUM = '1' THEN A1.U_JONGAB ELSE 0 END),
                    JONJU1 = MAX(CASE WHEN A1.U_LINENUM = '1' THEN A1.U_JONJUM ELSE 0 END),
                    JONNO1 = MAX(CASE WHEN A1.U_LINENUM = '1' THEN A1.U_JONNON ELSE 0 END),

                    JSTR02 = MAX(CASE WHEN A1.U_LINENUM = '2' THEN CONVERT(CHAR(10),A1.U_JONSTR,120) ELSE '' END), 
                    JEND02 = MAX(CASE WHEN A1.U_LINENUM = '2' THEN CONVERT(CHAR(10),A1.U_JONEND,120) ELSE '' END), 
                    JGFR02 = MAX(CASE WHEN A1.U_LINENUM = '2' THEN CONVERT(CHAR(10),A1.U_JONGFR,120) ELSE '' END), 
                    JGTO02 = MAX(CASE WHEN A1.U_LINENUM = '2' THEN CONVERT(CHAR(10),A1.U_JONGTO,120) ELSE '' END),
                    JONGA2 = MAX(CASE WHEN A1.U_LINENUM = '2' THEN A1.U_JONGAB ELSE 0 END),
                    JONJU2 = MAX(CASE WHEN A1.U_LINENUM = '2' THEN A1.U_JONJUM ELSE 0 END),
                    JONNO2 = MAX(CASE WHEN A1.U_LINENUM = '2' THEN A1.U_JONNON ELSE 0 END)
            FROM    [@ZPY502H] A0
                    INNER JOIN [@ZPY502L] A1 ON A0.DOCENTRY = A1.DOCENTRY
            WHERE   A0.U_JSNYER = @JSNYER
            AND     A0.U_MSTCOD LIKE @MSTCOD 
            GROUP   BY A0.U_MSTCOD,  A0.U_JSNYER, A0.U_CLTCOD
            ) T7 ON T0.U_JSNYER = T7.U_JSNYER AND T0.U_MSTCOD = T7.U_MSTCOD AND T0.U_CLTCOD = T7.U_CLTCOD
    WHERE   T0.U_JSNYER =    @JSNYER
    AND     T0.U_CLTCOD LIKE @CLTCOD
    AND     T3.U_MSTDPT LIKE @MSTDPT                        
    AND     T0.U_MSTCOD LIKE @MSTCOD
    --AND     ISNULL(CONVERT(Nvarchar(8),T2.Branch), '')  LIKE @Branch
    AND     (@JOBGBN ='3' OR (@JOBGBN <> '3' AND T0.U_JSNGBN = @JOBGBN))
    AND     T0.U_JSNMON  BETWEEN @STRMON AND @ENDMON
/*
    AND     T2.Status LIKE CASE @JOBGBN WHEN '1' THEN '1' 
                                        WHEN '2' THEN '4'
                                        ELSE '%' END
*/
    ORDER BY  T0.U_MSTNAM,  T0.U_MSTCOD

---------------------------------------------------------------------------------------------------
-- 5. 결과조회
---------------------------------------------------------------------------------------------------
    SELECT * FROM [#RPY504] ORDER BY U_MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF
