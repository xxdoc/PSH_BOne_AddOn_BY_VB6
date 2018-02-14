/*==========================================================================================
		프로시저명		: PH_PY110
		프로시저설명	: 개인별 상여율 대상자 조회
		작업일자		: 2012-11-27
		작업목적		: 개인별 상여율 등록 대상자를 조회
		작업내용		: 
	===========================================================================================*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY110' AND xtype = 'P'))
	DROP PROCEDURE PH_PY110
GO

CREATE  PROC PH_PY110 (
		@CLTCOD		AS Nvarchar(1),		-- 사업장
		@JOBYMM		AS Nvarchar(6),		-- 귀속년월
		@JOBTRG		AS Nvarchar(1)
		
	)
 AS
	

	SET NOCOUNT ON
---------------------------------------------------------------------------------------------------
-- 1. 변수 선언, 임시테이블 생성
---------------------------------------------------------------------------------------------------
	DECLARE @PH_PY110_CODE	AS NVARCHAR(6)
	DECLARE @PAYSEL			AS NVARCHAR(10)
	DECLARE @STRDAT			AS DATETIME
	DECLARE @ENDDAT			AS DATETIME

	CREATE TABLE #PH_PY110 (
		MSTCOD	NVARCHAR(8),	--사번
		MSTNAM	NVARCHAR(50),	--사원명
		EMPID	NVARCHAR(10),	--사원순번
		DPTCOD	NVARCHAR(8),	--부서코드
		DPTNAM	NVARCHAR(50),	--부서이름
		STRDAT  DATETIME,		--입사일자
        ENDDAT  DATETIME)		--퇴직일자

---------------------------------------------------------------------------------------------------
-- 2. 급여지급대상별로 급여기준일 설정 데이터 조회
---------------------------------------------------------------------------------------------------
	SET @PH_PY110_CODE = (
		SELECT	MAX(U_YM)
		FROM	[@PH_PY107A]
		WHERE	U_YM <= @JOBYMM
				AND U_CLTCOD = @CLTCOD
				)

	DECLARE EMP_LIST CURSOR FOR
	SELECT	U_PAYSEL,
			-- 근태기간 종료일
			U_STRDAT = SUBSTRING(DBO.Func_PAYTerm(LEFT(@JOBYMM,4)+'-'+SUBSTRING(@JOBYMM,5,2)+'-01', U_PAYSEL),10,8),
			-- 상여지급 제외일
			U_ENDDAT = SUBSTRING(DBO.Func_PAYTerm(LEFT(@JOBYMM,4)+'-'+SUBSTRING(@JOBYMM,5,2)+'-01', U_PAYSEL),28,8)
	--SELECT *	
	FROM	[@PH_PY107B] T0
	WHERE	substring(Code,2,7) = @PH_PY110_CODE
			AND U_PAYSEL = @JOBTRG
	
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
		SELECT	T1.Code,									--사번
				T1.U_FullName,								--풀네임
				CONVERT(NVARCHAR(10),T1.U_empID),			--사번순번
				T1.U_TeamCode,								--부서
				(SELECT U_CodeNm FROM [@PS_HR200L] where Code='1' AND U_CODE = T1.U_TeamCode),	--부서명
				T1.U_startDat,								--입사일자
				T1.U_termDate								--퇴직일자
		FROM	[@PH_PY001A] T1
		WHERE	T1.U_PAYSEL = @PAYSEL
		AND		T1.U_StartDat <= @STRDAT
		AND		(T1.U_TermDate >  @ENDDAT OR T1.U_TermDate IS NULL)
		AND		T1.U_BNSSEL = 'Y'
		AND		U_CLTCOD = @CLTCOD

		FETCH NEXT FROM EMP_LIST
		INTO @PAYSEL, @STRDAT, @ENDDAT
	END
	CLOSE EMP_LIST
	DEALLOCATE EMP_LIST

---------------------------------------------------------------------------------------------------
-- 4. 결과조회
---------------------------------------------------------------------------------------------------
	SELECT distinct MSTCOD,MSTNAM,EMPID,DPTCOD,DPTNAM,STRDAT,ENDDAT FROM #PH_PY110 ORDER BY DPTCOD, MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF


-- Exec PH_PY110  '1','201211','1'