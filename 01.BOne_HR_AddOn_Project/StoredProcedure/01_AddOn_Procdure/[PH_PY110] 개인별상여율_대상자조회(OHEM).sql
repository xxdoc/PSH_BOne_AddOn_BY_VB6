CREATE  PROC PH_PY110 (
		@JOBYMM		AS Nvarchar(6)		-- �ͼӳ��
	)
	--WITH Encryption  
 AS
	/*==========================================================================================
		���ν�����		: PH_PY110
		���ν�������	: ���κ� ���� ����� ��ȸ
		������			: �ֵ���
		�۾�����		: 2010-04-02
		�۾�������		: �ֵ���
		�۾���������	: 2010-04-02
		�۾�����		: ���κ� ���� ��� ����ڸ� ��ȸ
		�۾�����		: 
	===========================================================================================*/
	--DROP PROC PH_PY110
	--Exec PH_PY110  '200701', '%', '%', '1', '%'

	SET NOCOUNT ON
---------------------------------------------------------------------------------------------------
-- 1. ���� ����, �ӽ����̺� ����
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
-- 2. �޿����޴�󺰷� �޿������� ���� ������ ��ȸ
---------------------------------------------------------------------------------------------------
	SET @PH_PY110_CODE = (
		SELECT	MAX(Code)
		FROM	[@PH_PY107A]
		WHERE	Code <= @JOBYMM)

	DECLARE EMP_LIST CURSOR FOR
	SELECT	U_PAYSEL,
			-- ���±Ⱓ ������
			U_STRDAT = SUBSTRING(DBO.Func_PAYTerm(LEFT(@JOBYMM,4)+'-'+SUBSTRING(@JOBYMM,5,2)+'-01', U_PAYSEL),10,8),
			-- ������ ������
			U_ENDDAT = SUBSTRING(DBO.Func_PAYTerm(LEFT(@JOBYMM,4)+'-'+SUBSTRING(@JOBYMM,5,2)+'-01', U_PAYSEL),28,8)
	FROM	[@PH_PY107B] T0
	WHERE	Code = @PH_PY110_CODE

	OPEN EMP_LIST
	FETCH NEXT FROM EMP_LIST
	INTO @PAYSEL, @STRDAT, @ENDDAT

---------------------------------------------------------------------------------------------------
-- 3. �޿����޴���ں� ������ ������ ������ ��ȸ
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
-- 4. �����ȸ
---------------------------------------------------------------------------------------------------
	SELECT * FROM #PH_PY110 ORDER BY MSTBRK, MSTDPT, MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

