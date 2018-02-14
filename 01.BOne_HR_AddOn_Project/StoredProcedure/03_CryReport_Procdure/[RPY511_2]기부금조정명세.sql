IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY511_2' AND xtype = 'P'))
	DROP PROCEDURE RPY511_2
GO

CREATE PROC RPY511_2 (
		@JSNYER		AS NVARCHAR(6),		--�ͼӳ⵵
		@CLTCOD     AS Nvarchar(8),     --�ڻ��ڵ�
		@MSTDPT     AS Nvarchar(8),     --�μ�
	    @MSTCOD 	AS NVARCHAR(8) 	   	--�����ȣ	
	) 

AS
    /*==========================================================================================
        ���ν�����      : RPY511_2
        ���ν�������    : ��α� ����_2
        ������          : ����ȣ
        �۾�����        : 2011-02-24
        �۾�������      : 
        �۾���������    : 
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/
    -- DROP PROC RPY511_2
    -- Exec RPY511_2 '2009', N'%', N'%','106001'

    SET NOCOUNT ON

-- <1.�ӽ����̺� ���� >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�

	CREATE TABLE #RPY511_2 (
		DocEntry	Int,
		U_GBUCOD	NVARCHAR(2),
		U_GIBUYY	NVARCHAR(6),
		U_GBUAMT	NUMERIC(19,6),
		U_BEFAMT	NUMERIC(19,6),
		U_DAEAMT	NUMERIC(19,6),
		U_CURAMT	NUMERIC(19,6),
		U_DELAMT	NUMERIC(19,6),
		U_CHAAMT	NUMERIC(19,6)

	)

-- <2.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�

	INSERT	INTO [#RPY511_2]
	SELECT	DocEntry    =   T0.DocEntry,
			U_GBUCOD	=	T1.U_GBUCOD,
			U_GIBUYY	=	LEFT(T1.U_GBUYMM,4),
			U_GBUAMT	=	SUM(T1.U_GBUAMT),
			U_BEFAMT	=	SUM(CASE WHEN T0.U_JSNYER > LEFT(T1.U_GBUYMM,4) THEN ISNULL(T1.U_GBUAMT,0) - ISNULL(T1.U_BEFAMT,0) ELSE 0 END),
			U_DAEAMT	=	SUM(CASE WHEN T0.U_JSNYER > LEFT(T1.U_GBUYMM,4) THEN T1.U_BEFAMT ELSE T1.U_GBUAMT END),
			U_CURAMT	=	SUM(T1.U_CURAMT),
			U_DELAMT	=	SUM(CASE WHEN T0.U_JSNYER > LEFT(T1.U_GBUYMM,4) THEN T1.U_BEFAMT ELSE T1.U_GBUAMT END)
			            -   ISNULL(SUM(T1.U_CURAMT),0) - ISNULL(SUM(T1.U_CHAAMT),0),
			U_CHAAMT	=	SUM(T1.U_CHAAMT)
	FROM	[@ZPY505H] T0 
			INNER JOIN [@ZPY505L] T1 ON T0.DocEntry = T1.DocEntry 
			INNER JOIN [@PH_PY001A] T2 ON T0.U_MSTCOD = T2.Code
			LEFT JOIN [@PH_PY005A] T3 ON T0.U_CLTCOD = T3.U_CLTCode		
	WHERE	T0.U_JSNYER = @JSNYER
	AND		T2.U_TeamCode LIKE @MSTDPT 
	AND		T0.U_MSTCOD LIKE @MSTCOD
	GROUP	BY T0.DocEntry, T1.U_GBUCOD, LEFT(T1.U_GBUYMM,4)

-- <3.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    
    SELECT * FROM [#RPY511_2]

	SET NOCOUNT OFF