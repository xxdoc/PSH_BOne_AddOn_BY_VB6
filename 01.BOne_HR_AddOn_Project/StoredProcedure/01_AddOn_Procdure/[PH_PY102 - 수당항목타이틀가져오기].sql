/*==========================================================================================
		���ν�����		: PH_PY102
		���ν�������	: �����׸� Ÿ��Ʋ����
		������			: 
		�۾�����		: 2012-11-20
		�۾�������		: 
		�۾���������	: 
		�۾�����		: ������,������,�޿��� �׸� Ÿ��Ʋ��ȸ��
		�۾�����		: 
	===========================================================================================*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY102' AND xtype = 'P'))
	DROP PROCEDURE PH_PY102
GO

CREATE            PROC PH_PY102
	(
		@CLTCOD		 AS Nvarchar(1),		--�����
		@YM 		 AS Nvarchar(6),		--�۾�����
		@STDTYP 	 AS Nvarchar(1),		--�����翩��
		@HOBUSE 	 AS Nvarchar(1),		--ȣ����������
		@FIXGBN 	 AS Nvarchar(1), 		--������������
		@CSUCOD		 AS Nvarchar(10)
	)
--WITH ENCRYPTION
 AS
	SET NOCOUNT ON
--< 1. �����׸� ��������� > �ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�		
	DECLARE	@U_JOBYMM		 AS Nvarchar(6)
	
	SELECT TOP 1 @U_JOBYMM = T0.U_YM
	FROM [@PH_PY102A] T0
	WHERE T0.U_YM <= @YM 
	AND   T0.U_CLTCOD = @CLTCOD
	ORDER BY T0.U_YM DESC
	IF ISNULL(@U_JOBYMM, '') = ''
	BEGIN
		SET @U_JOBYMM = @YM
	END 	

--< 2. �����׸� Ÿ��Ʋ ���� > �ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�			
	SELECT 	ISNULL(T0.U_CSUCOD,'') AS U_CSUCOD, 
			ISNULL(T0.U_CSUNAM,'')  AS U_CSUNAM, 
			ISNULL(T1.Code,'') AS Code, 
			ISNULL(T0.U_MONPAY,'') AS U_MONPAY,
			ISNULL(T0.U_KUMAMT, 0) AS U_KUMAMT, 
			ISNULL(T0.U_CSUGBN, 30) AS U_CSUGBN,
			ISNULL(T0.U_GWATYP, '') AS U_GWATYP,
			ISNULL(T0.U_GBHGBN, '') AS U_GBHGBN,
			ISNULL(T0.U_ROUNDT,'R') AS U_ROUNDT,
			ISNULL(T0.U_LENGTH,'1') AS U_LENGTH,
			ISNULL(T0.U_BNSUSE, 'N') AS U_BNSUSE,
			ISNULL(T0.U_INSLIN,'') AS U_INSLIN,
			ISNULL(T0.U_LINSEQ,'') AS U_LINSEQ,
			ISNULL(T0.U_BTXCOD, '') AS U_BTXCOD

	INTO #PH_PY102
	--select *
	FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code 
	WHERE T1.U_CLTCOD = @CLTCOD
	AND  T1.U_YM = @U_JOBYMM	
	AND  (@HOBUSE = '' OR (@HOBUSE <> '' AND T0.U_HOBUSE = @HOBUSE))	--ȣ����������
	AND	 (@FIXGBN = '' OR (@FIXGBN <> '' AND T0.U_FIXGBN = @FIXGBN))		--������������(Y����,N����)
	AND	 (@FIXGBN = '' OR (@FIXGBN <> '' AND LEFT(T0.U_CSUCOD,1) <> 'A'))	--�⺻��,�󿩱�����
	--AND  T0.U_CSUCOD LIKE @CSUCOD
	AND	 ISNULL(T0.U_LINSEQ,'') <> ''
	ORDER BY T0.U_LINSEQ
	IF @FIXGBN = 'Y' --�������ϰ�� �����λ������ ���� ��µǵ���
	BEGIN
		SELECT * FROM [#PH_PY102] ORDER BY U_INSLIN
	END
	ELSE
	BEGIN
		SELECT * FROM [#PH_PY102] ORDER BY U_LINSEQ
	END		


--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

--GO
--Exec PH_PY102  '1','201212', '', '', '', ''

--
