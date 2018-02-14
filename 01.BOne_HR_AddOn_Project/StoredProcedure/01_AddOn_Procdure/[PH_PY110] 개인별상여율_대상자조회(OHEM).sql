CREATE  PROC PH_PY110 (
		@JOBYMM		AS Nvarchar(6)		-- 귀속년월
	)
	--WITH Encryption  
 AS
	/*==========================================================================================
		프로시저명		: PH_PY110
		프로시저설명	: 개인별 상여율 대상자 조회
		만든이			: 최동권
		작업일자		: 2010-04-02
		작업지시자		: 최동권
		작업지시일자	: 2010-04-02
		작업목적		: 개인별 상여율 등록 대상자를 조회
		작업내용		: 
	===========================================================================================*/
	--DROP PROC PH_PY110
	--Exec PH_PY110  '200701', '%', '%', '1', '%'

	SET NOCOUNT ON
---------------------------------------------------------------------------------------------------
-- 1. 변수 선언, 임시테이블 생성
---------------------------------------------------------------------------------------------------
	DECLARE @PH_PY110_CODE	AS NVARCHAR(6)
	DECLARE @PAYSEL			AS NVARCHAR(10)
	DECLARE @STRDAT			AS DATETIME
	DECLARE @ENDDAT			AS DATETIME

	CREATE TABLE #PH_PY110 (
		MSTCOD	NVARCHAR(8),
		MSTNAM	NVARCHAR(50),
		EMPID	NVARCHAR(10),
		MSTBRK	SMALLINT,
		BRKNAM	NVARCHAR(50),
		MSTDPT	NVARCHAR(8),
		DPTNAM	NVARCHAR(50),
		STRDAT  DATETIME,
        ENDDAT  DATETIME)

---------------------------------------------------------------------------------------------------
-- 2. 급여지급대상별로 급여기준일 설정 데이터 조회
---------------------------------------------------------------------------------------------------
	SET @PH_PY110_CODE = (
		SELECT	MAX(Code)
		FROM	[@PH_PY107A]
		WHERE	Code <= @JOBYMM)

	DECLARE EMP_LIST CURSOR FOR
	SELECT	U_PAYSEL,
			-- 근태기간 종료일
			U_STRDAT = SUBSTRING(DBO.Func_PAYTerm(LEFT(@JOBYMM,4)+'-'+SUBSTRING(@JOBYMM,5,2)+'-01', U_PAYSEL),10,8),
			-- 상여지급 제외일
			U_ENDDAT = SUBSTRING(DBO.Func_PAYTerm(LEFT(@JOBYMM,4)+'-'+SUBSTRING(@JOBYMM,5,2)+'-01', U_PAYSEL),28,8)
	FROM	[@PH_PY107B] T0
	WHERE	Code = @PH_PY110_CODE

	OPEN EMP_LIST
	FETCH NEXT FROM EMP_LIST
	INTO @PAYSEL, @STRDAT, @ENDDAT

---------------------------------------------------------------------------------------------------
-- 3. 급여지급대상자별 기준일 범위내 재직자 조회
---------------------------------------------------------------------------------------------------
	WHILE @@FETCH_STATUS = 0
	BEGIN
	--	SELECT @PAYSEL, @STRDAT, @ENDDAT
		INSERT INTO #PH_PY110
		SELECT	T0.U_MSTCOD,
				lastName + firstName,
				CONVERT(NVARCHAR(10),empID),
				T0.branch,
				T3.Name,
				--T2.U_MSTDPT,
				T2.Name,
				T0.startDate,
				T0.termDate
		FROM	OHEM T0
				INNER JOIN [@PH_PY001A] T1 ON T0.U_MSTCOD = T1.Code
				LEFT  JOIN [OUDP] T2 ON T0.dept = T2.Code
				LEFT  JOIN [OUBR] T3 ON T0.branch = T3.Code
		WHERE	T1.U_PAYSEL = @PAYSEL
		AND		T0.startDate <= @STRDAT
		AND		(T0.termDate >  @ENDDAT OR T0.termDate IS NULL)
		AND		T1.U_BNSSEL = 'Y'

		FETCH NEXT FROM EMP_LIST
		INTO @PAYSEL, @STRDAT, @ENDDAT
	END
	CLOSE EMP_LIST
	DEALLOCATE EMP_LIST

---------------------------------------------------------------------------------------------------
-- 4. 결과조회
---------------------------------------------------------------------------------------------------
	SELECT * FROM #PH_PY110 ORDER BY MSTBRK, MSTDPT, MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

