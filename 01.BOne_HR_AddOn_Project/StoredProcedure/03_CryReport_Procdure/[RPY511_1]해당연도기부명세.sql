IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY511_1' AND xtype = 'P'))
	DROP PROCEDURE RPY511_1
GO

CREATE PROC RPY511_1 (
		@JSNYER		AS NVARCHAR(6),		--�ͼӳ⵵
		@CLTCOD     AS Nvarchar(8),     --�ڻ��ڵ�
		@MSTDPT     AS Nvarchar(8),     --�μ�
	    @MSTCOD 	AS NVARCHAR(8) 	   	--�����ȣ			
	) 

AS
    /*==========================================================================================
        ���ν�����      : RPY511_1
        ���ν�������    : ��α� ����_1
        ������          : ����ȣ
        �۾�����        : 2011-02-24
        �۾�������      : 
        �۾���������    : 
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/
    -- DROP PROC RPY511_1
    -- Exec RPY511_1 '2009', N'%', N'%', '106001'

    SET NOCOUNT ON

-- <1.�ӽ����̺� ���� >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�

	CREATE TABLE #RPY511_1 (
		DocEntry	Int,
		U_GBUYOU	NVARCHAR(40),
		U_GBUCOD	NVARCHAR(2),
		U_GBUNAE	NVARCHAR(10),
		U_GBUNAM	NVARCHAR(40),
		U_GBUNBR	NVARCHAR(14),
		U_GWANGE	NVARCHAR(1),
		U_FAMNAM	NVARCHAR(20),
		U_PERNBR	NVARCHAR(14),
		U_GBUCNT	NUMERIC(19,6),
		U_GBUAMT	NUMERIC(19,6)
	)

-- <2.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�

	INSERT	INTO [#RPY511_1]
	SELECT	DocEntry    =   T0.DocEntry,
			U_GBUYOU	=	CASE WHEN T1.U_GBUCOD = '10' THEN N'����'
								 WHEN T1.U_GBUCOD = '20' THEN N'��ġ�ڱ�'
								 WHEN T1.U_GBUCOD = '30' THEN N'��Ư�� 73'
								 WHEN T1.U_GBUCOD = '31' THEN N'��Ư�� 73 �� 11'
								 WHEN T1.U_GBUCOD = '40' THEN N'����'
								 WHEN T1.U_GBUCOD = '41' THEN N'����'
								 WHEN T1.U_GBUCOD = '42' THEN N'�츮����'
								 WHEN T1.U_GBUCOD = '50' THEN N'��������'
								 ELSE '' END,
			U_GBUCOD	=	T1.U_GBUCOD,
			U_GBUNAE	=	N'����',
			U_GBUNAM	=	MAX(T1.U_GBUNAM),
			U_GBUNBR	=	T1.U_GBUNBR,
			U_GWANGE	=	MAX(T1.U_GWANGE),
			U_FAMNAM	=	MAX(T1.U_FAMNAM),
			U_PERNBR	=	T1.U_PERNBR,
			U_GBUCNT	=	SUM(T1.U_GBUCNT),
			U_GBUAMT	=	SUM(T1.U_GBUAMT)
	FROM	[@ZPY505H] T0 
			INNER JOIN [@ZPY505L] T1 ON T0.DocEntry = T1.DocEntry
			INNER JOIN [@PH_PY001A] T2 ON T0.U_MSTCOD = T2.Code
			LEFT JOIN [@PH_PY005A] T3 ON T0.U_CLTCOD = T3.U_CLTCode			
	WHERE	T0.U_JSNYER = @JSNYER
	AND		T2.U_TeamCode LIKE @MSTDPT 
	AND		T0.U_MSTCOD LIKE @MSTCOD
	GROUP	BY T0.DocEntry, T1.U_GBUCOD, T1.U_GBUNBR, T1.U_PERNBR

-- <3.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    
    SELECT * FROM [#RPY511_1]

	SET NOCOUNT OFF