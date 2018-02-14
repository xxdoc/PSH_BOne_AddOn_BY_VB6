/*==========================================================================================
		함수명		: PH_PY106
		함수설명	: 수당계산공식
		만든이			: 함미경
		작업일자		: 2007-07-05
		작업지시자		: 함미경
		작업지시일자	: 2007-01-11
		작업목적		: 업체별로 해당급여 수식에 따라 수식계산을 해줍니다.
		작업내용		: 
	===========================================================================================*/
		/*ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
		SELECT ROUND(124.573, 0, 0)   A: 125.000 반올림(숫자, 반올림자릿수, 0:반올림, -1:버림)
		SELECT ROUND(124.573, 0, -1)  --A: 124.000 1원미만 버림1원자름
			
		SELECT ROUND(125.573/10,0,-1)*10 			A:120.000 10원미만자름
		SELECT ROUND(125.5730000, -1, -1)  A:120.000    10원미만자름	
	ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY106' AND xtype = 'P'))
	DROP PROCEDURE PH_PY106
GO

create             PROC PH_PY106
	(
		@JOBDAT 	 AS Nvarchar(6)	,	--작업연월
		@MSTCOD 	 AS Nvarchar(8)	,	--사원번호
		@GONSIL	 	 AS Nvarchar(2000)	--계산공식
	)
--WITH ENCRYPTION
 AS
	

	
--< 1. 사용가능한 환경 만들기 >
	--1.1) 급여고정수당들
		SELECT	Code AS U_MSTCOD, 
				SUM(CASE U_LineNum WHEN '1' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD01, 
				SUM(CASE U_LineNum WHEN '2' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD02, 
				SUM(CASE U_LineNum WHEN '3' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD03, 
				SUM(CASE U_LineNum WHEN '4' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD04, 
				SUM(CASE U_LineNum WHEN '5' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD05, 
				SUM(CASE U_LineNum WHEN '6' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD06, 
				SUM(CASE U_LineNum WHEN '7' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD07, 
				SUM(CASE U_LineNum WHEN '8' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD08, 
				SUM(CASE U_LineNum WHEN '9' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD09, 
				SUM(CASE U_LineNum WHEN '10' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD10, 
				SUM(CASE U_LineNum WHEN '11' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD11, 
				SUM(CASE U_LineNum WHEN '12' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD12, 
				SUM(CASE U_LineNum WHEN '13' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD13, 
				SUM(CASE U_LineNum WHEN '14' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD14, 
				SUM(CASE U_LineNum WHEN '15' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_CSUD15 
		INTO #PH_PY001B
		FROM [@PH_PY001B]
		WHERE (@MSTCOD = '%' OR (@MSTCOD <> '%' AND Code =  @MSTCOD ))
		GROUP BY  Code

	--1.2) 급여고정공제들
--ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    							     								
	CREATE TABLE #PH_PY001C (
	        U_MSTCOD    nvarchar(8),
	        U_GOND01    Numeric(19,6),
	        U_GOND02    Numeric(19,6),
	        U_GOND03    Numeric(19,6),
	        U_GOND04    Numeric(19,6),
	        U_GOND05    Numeric(19,6),
	        U_GOND06    Numeric(19,6),
	        U_GOND07    Numeric(19,6),
	        U_GOND08    Numeric(19,6),
	        U_GOND09    Numeric(19,6),
	        U_GOND10    Numeric(19,6),
	        U_GOND11    Numeric(19,6),
			U_GOND12    Numeric(19,6),
			U_GOND13    Numeric(19,6)
	        ) 
		INSERT INTO [#PH_PY001C]
		SELECT	Code AS U_MSTCOD, 
				SUM(CASE U_LineNum WHEN '1' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND01, 
				SUM(CASE U_LineNum WHEN '2' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND02, 
				SUM(CASE U_LineNum WHEN '3' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND03, 
				SUM(CASE U_LineNum WHEN '4' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND04, 
				SUM(CASE U_LineNum WHEN '5' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND05, 
				SUM(CASE U_LineNum WHEN '6' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND06, 
				SUM(CASE U_LineNum WHEN '7' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND07, 
				SUM(CASE U_LineNum WHEN '8' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND08, 
				SUM(CASE U_LineNum WHEN '9' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND09, 
				SUM(CASE U_LineNum WHEN '10' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND10, 
				SUM(CASE U_LineNum WHEN '11' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND11,
				SUM(CASE U_LineNum WHEN '12' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND12,
				SUM(CASE U_LineNum WHEN '13' THEN ISNULL(U_FILD03, 0) ELSE 0 END) AS U_GOND13   
		FROM [@PH_PY001C]
		WHERE (@MSTCOD = '%' OR (@MSTCOD <> '%' AND Code =  @MSTCOD ))
		GROUP BY  Code

	--1.2)근태자료들
		SELECT  T1.* 
		INTO #ZPY230L
		FROM [@ZPY230H] T0 INNER JOIN [@ZPY230L] T1 ON T0.DocEntry = T1.DocEntry
		WHERE (@MSTCOD = '%' OR (@MSTCOD <> '%' AND T1.U_MSTCOD = @MSTCOD))
		AND	  (@JOBDAT = '%' OR (@JOBDAT <> '%' AND T0.U_GNTYMM = @JOBDAT))   

--< 2. 변수들 해당 필드로 변경해주세요 > ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ		
		DECLARE @Tmp_GONSIL as VarChar(8000)
		DECLARE	@JSUCNT		 AS Int	--총제수당갯수
		DECLARE	@QueryString AS VarChar(8000)  
		
		SET @JSUCNT = 31
		
		SET @Tmp_GONSIL = @GONSIL
		--수정할꺼 수정하고
		--SET @Tmp_GONSIL = REPLACE(@Tmp_GONSIL,


--<3. 계산이 필요한 얘들> ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ		



--< 2. 변수들 해당 필드로 변경해주세요 > ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ		
		
		
		SET @Tmp_GONSIL = @GONSIL
		--수정할꺼 수정하고
		--SET @Tmp_GONSIL = REPLACE(@Tmp_GONSIL,


		SET @QueryString = 'SELECT '
		SET @QueryString = @QueryString + @Tmp_GONSIL
--< 2. 쿼리문 생성이여 > ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ		
--	PRINT @QueryString

--	SET @QueryString = ' SELECT * FROM [OHEM] T0 WHERE (@MSTCOD = '%' OR (@MSTCOD <> '%' AND T0.U_MSTCOD = @MSTCOD))'

	SET @QueryString = RTRIM(@QueryString) + 
			 ' FROM [@PH_PY001A] T0 LEFT JOIN [@PH_PY001B] T2 ON T0.Code = T2.Code COLLATE Korean_Wansung_CI_AS
						LEFT JOIN [@PH_PY001C] T3 ON T0.Code = T3.Code COLLATE Korean_Wansung_CI_AS
						LEFT JOIN  [@PS_HR200L] T4 ON T0.U_TeamCode = T4.U_Code AND T4.Code = ''1'' AND T4.U_UseYN = ''Y''
						LEFT JOIN  [@PS_HR200L] T5 ON T0.U_Position = T5.U_COde AND T4.Code = ''P129'' AND T5.U_UseYN = ''Y''
						LEFT JOIN [#ZPY230L] T6 ON T0.Code = T6.U_MSTCOD
			WHERE T0.Code = ' +  '''' + @MSTCOD + ''''
--	SET @QueryString =@QueryString + '' + @MSTCOD + ''

--SELECT @QueryString
	EXEC(@QueryString)

/*
	--제수당 총사용갯수
	DECLARE	@JSUCNT		 AS numeric(19, 6)	--총제수당갯수
	
	SET @JSUCNT = 0	
	
	SELECT @JSUCNT = ISNULL(MAX(T0.U_INSLIN), 0)
	FROM [@ZPY111L] T0 INNER JOIN [@ZPY111H] T1 ON T0.Code = T1.Code -- WHERE T0.Code = N'YES'
	AND     T0.U_STDTYP='Y'   --제수당여부
	AND     T0.U_HOBUSE='Y'	--호봉참조여부
	AND     T0.U_FIXGBN='Y'	--고정변동구분	
	--AND     T1.CODE='YES'	


--< 2. 호봉표에서 해당연월의 호봉자료 사원마스터에 적용 > ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ	
	DECLARE @SposID      AS Nvarchar(10)               
	DECLARE @EposID      AS Nvarchar(10)            

	DECLARE @STPCOD      AS Nvarchar(10)               
	DECLARE @PosID      AS Nvarchar(10)            
	DECLARE @HOBONG      AS Nvarchar(10)               
	DECLARE	@STDAMT		 AS numeric(19, 6)	
	DECLARE	@JSUA01		 AS numeric(19, 6)		   	
	DECLARE	@JSUA02	 	 AS numeric(19, 6)
	DECLARE	@JSUA03 	 AS numeric(19, 6)	         	
    DECLARE	@JSUA04 	 AS numeric(19, 6)	
    DECLARE	@JSUA05		 AS numeric(19, 6)	
	DECLARE	@JSUA06		 AS numeric(19, 6)		   	
	DECLARE	@JSUA07	 	 AS numeric(19, 6)
	DECLARE	@JSUA08 	 AS numeric(19, 6)	         	
    DECLARE	@JSUA09 	 AS numeric(19, 6)	
    DECLARE	@JSUA10		 AS numeric(19, 6)	
	

	SELECT	T1.U_STPCOD,
			T2.posID,
			T0.U_HOBCOD,
			ISNULL(U_STDAMT, 0) AS STDAMT,
			ISNULL(U_JSUA01, 0) AS JSUA01,
			ISNULL(U_JSUA02, 0) AS JSUA02,
			ISNULL(U_JSUA03, 0) AS JSUA03,
			ISNULL(U_JSUA04, 0) AS JSUA04,
			ISNULL(U_JSUA05, 0) AS JSUA05,
			ISNULL(U_JSUA06, 0) AS JSUA06,
			ISNULL(U_JSUA07, 0) AS JSUA07,
			ISNULL(U_JSUA08, 0) AS JSUA08,
			ISNULL(U_JSUA09, 0) AS JSUA09,
			ISNULL(U_JSUA10, 0) AS JSUA10
	FROM [@PH_PY105B] T0	INNER JOIN  [@PH_PY105A]	T1 ON T0.DocEntry = T1.DocEntry
						INNER JOIN	[OHPS]	T2 ON T1.U_STPCOD = T2.U_MSTSTP
	WHERE T1.U_JOBDAT =@JOBDAT
	ORDER BY T1.U_STPCOD, T0.U_HOBCOD
*/
	
-- 최종수정일


--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
--SET NOCOUNT OFF


--Exec PH_PY106  '200706', 'SJ067', 'ROUND(T1.U_STDAMT/30,0,0)'
--Exec PH_PY106  '200705', '20041101', '4229*T4.U_HGNTIM*1.5'
--SET NOCOUNT ON
--Exec PH_PY106  '200706', 'SJ067', 'CASE WHEN 1 =1 THEN (CASE 5 WHEN 0 THEN 0 WHEN 1 THEN 10000 WHEN 2 THEN 15000 WHEN 3 THEN 20000 WHEN 4 THEN 25000 ELSE 30000 END) /  30 * 15  ELSE (CASE 5 WHEN 0 THEN 0 WHEN 1 THEN 10000 WHEN 2 THEN 15000 WHEN 3 THEN 20000 WHEN 4 THEN 25000 ELSE 30000 END) END'


