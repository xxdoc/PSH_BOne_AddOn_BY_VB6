/*==========================================================================================
		���ν�����		: PH_PY110
		���ν�������	: ���κ� ���� ����� ��ȸ
		�۾�����		: 2012-11-27
		�۾�����		: ���κ� ���� ��� ����ڸ� ��ȸ
		�۾�����		: 
	===========================================================================================*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY110' AND xtype = 'P'))
	DROP PROCEDURE PH_PY110
GO

CREATE  PROC PH_PY110 (
		@CLTCOD		AS Nvarchar(1),		-- �����
		@JOBYMM		AS Nvarchar(6),		-- �ͼӳ��
		@JOBTRG		AS Nvarchar(1)
		
	)
 AS
	

	SET NOCOUNT ON
---------------------------------------------------------------------------------------------------
-- 1. ���� ����, �ӽ����̺� ����
---------------------------------------------------------------------------------------------------
	DECLARE @PH_PY110_CODE	AS NVARCHAR(6)
	DECLARE @PAYSEL			AS NVARCHAR(10)
	DECLARE @STRDAT			AS DATETIME
	DECLARE @ENDDAT			AS DATETIME

	CREATE TABLE #PH_PY110 (
		MSTCOD	NVARCHAR(8),	--���
		MSTNAM	NVARCHAR(50),	--�����
		EMPID	NVARCHAR(10),	--�������
		DPTCOD	NVARCHAR(8),	--�μ��ڵ�
		DPTNAM	NVARCHAR(50),	--�μ��̸�
		STRDAT  DATETIME,		--�Ի�����
        ENDDAT  DATETIME)		--��������

---------------------------------------------------------------------------------------------------
-- 2. �޿����޴�󺰷� �޿������� ���� ������ ��ȸ
---------------------------------------------------------------------------------------------------
	SET @PH_PY110_CODE = (
		SELECT	MAX(U_YM)
		FROM	[@PH_PY107A]
		WHERE	U_YM <= @JOBYMM
				AND U_CLTCOD = @CLTCOD
				)

	DECLARE EMP_LIST CURSOR FOR
	SELECT	U_PAYSEL,
			-- ���±Ⱓ ������
			U_STRDAT = SUBSTRING(DBO.Func_PAYTerm(LEFT(@JOBYMM,4)+'-'+SUBSTRING(@JOBYMM,5,2)+'-01', U_PAYSEL),10,8),
			-- ������ ������
			U_ENDDAT = SUBSTRING(DBO.Func_PAYTerm(LEFT(@JOBYMM,4)+'-'+SUBSTRING(@JOBYMM,5,2)+'-01', U_PAYSEL),28,8)
	--SELECT *	
	FROM	[@PH_PY107B] T0
	WHERE	substring(Code,2,7) = @PH_PY110_CODE
			AND U_PAYSEL = @JOBTRG
	
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
		SELECT	T1.Code,									--���
				T1.U_FullName,								--Ǯ����
				CONVERT(NVARCHAR(10),T1.U_empID),			--�������
				T1.U_TeamCode,								--�μ�
				(SELECT U_CodeNm FROM [@PS_HR200L] where Code='1' AND U_CODE = T1.U_TeamCode),	--�μ���
				T1.U_startDat,								--�Ի�����
				T1.U_termDate								--��������
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
-- 4. �����ȸ
---------------------------------------------------------------------------------------------------
	SELECT distinct MSTCOD,MSTNAM,EMPID,DPTCOD,DPTNAM,STRDAT,ENDDAT FROM #PH_PY110 ORDER BY DPTCOD, MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF


-- Exec PH_PY110  '1','201211','1'