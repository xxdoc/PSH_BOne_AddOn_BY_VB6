/*==========================================================================================
		���ν�����		: PH_PY103
		���ν�������	: �����׸� Ÿ��Ʋ����
		������			: 
		�۾�����		: 2012-11-20
		�۾�������		: 
		�۾���������	: 
		�۾�����		: ������,������,�޿��� �׸� Ÿ��Ʋ��ȸ��
		�۾�����		: 
	===========================================================================================*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY103' AND xtype = 'P'))
	DROP PROCEDURE PH_PY103
GO

CREATE             PROC PH_PY103
	(
		@CLTCOD		AS Nvarchar(1),		--�����
		@YM 		AS Nvarchar(6),		--�۾�����
		@FIXGBN 	AS Nvarchar(1), 	--������������
		@CSUCOD		AS Nvarchar(10)
	)
--WITH ENCRYPTION
 AS
	

	SET NOCOUNT ON


--< 1. �����׸� ��������� > �ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�		
	DECLARE	@U_JOBYMM		 AS Nvarchar(6)
	
	SELECT TOP 1 @U_JOBYMM = T0.U_YM
	FROM [@PH_PY103A] T0
	WHERE T0.U_YM <= @YM 
	AND   T0.U_CLTCOD = @CLTCOD
	ORDER BY T0.Code DESC

	IF ISNULL(@U_JOBYMM, '') = ''
	BEGIN
		SET @U_JOBYMM = @YM
	END 	
	
--< 2. �����׸� Ÿ��Ʋ ���� > �ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�			
	SELECT 	ISNULL(T0.U_CSUCOD, '') AS U_CSUCOD, 
			ISNULL(T0.U_CSUNAM,'')  AS U_CSUNAM, 
			ISNULL(T1.Code,'') AS Code, 
			ISNULL(T0.U_SILCUN,'') AS U_SILCUN, 
			ISNULL(T0.U_BNSUSE,'N') AS U_BNSUSE, 
			ISNULL(T0.U_ROUNDT, 'R') AS U_ROUNDT, 
			ISNULL(T0.U_LENGTH, 1) AS U_LENGTH
	FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code -- WHERE T0.Code = N'YES'
	WHERE T1.U_CLTCOD = @CLTCOD
	AND  T1.U_YM = @U_JOBYMM 	
	AND  (@FIXGBN = '' OR (@FIXGBN <> '' AND T0.U_FIXGBN = @FIXGBN))	--������������(Y����,N����)
	--AND  T0.U_CSUCOD LIKE @CSUCOD
	AND	 ISNULL(T0.U_LINSEQ,'') <> ''
	ORDER BY T0.U_LINSEQ


--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

--Exec PH_PY103  '1','201212',  '', ''