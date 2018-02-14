IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY401' AND xtype = 'P'))
	DROP PROCEDURE RPY401
GO

CREATE  PROC RPY401
	(
		@STRDAT 	AS Nvarchar(10), 	--��������
		@ENDDAT 	AS Nvarchar(10), 	--��������
		@JOBGBN		AS Nvarchar(1), 	--�۾�����(1��������,2�ߵ�����,3��ü)
		@Branch		AS Nvarchar(8), 	--����
		@MSTDPT		AS Nvarchar(8), 	--�μ�
	    @MSTCOD 	AS Nvarchar(8), 	   	--�����ȣ
	    @STRJIG 	AS Nvarchar(10), 	--���޽�������
		@ENDJIG 	AS Nvarchar(10) 	--������������		
	)
	

 AS
    /*==========================================================================================
        ���ν�����      : RPY401
        ���ν�������    : �����ҵ��õ¡�������� - ��1����
        ������          : �Թ̰�
        �ۼ�����        : 2007-11-30
        �۾�������      : �Թ̰�
        �۾���������    : 2007-11-30
        �۾�����        : �����ҵ� ��õ¡���������� ���
        �۾�����        : 
		������/��������	: �ֵ��� / 2009-02-05
		����������/����	: �Թ̰� / 2009-02-04
		��������		: �Ի��� <> ��������� AND �������� AND ��������� �����ϴ� ��� 
						  ��2���Ŀ� ���;��ϹǷ� �̸� ��� �����ϴ� ��� ��ȸ���� �ʵ��� ����
		��������: 2009.05-28 �����ҵ漼�װ������� ���Ĥ�����.
    ===========================================================================================*/
	--DROP PROC RPY401
	--Exec RPY401  '2008-01-01', '2009-12-31', '3', N'%', N'%', N'%'
SET NOCOUNT ON
--<1.�ӽ����̺����>�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    

	CREATE TABLE #RPY401 (
        DocNum	   	Int,
        U_MSTCOD   	NvarChar(10),	U_MSTNAM   	NvarChar(20),
        U_EmpID   	NvarChar(10),	U_INTGBN   	NvarChar(1),
	    U_DWEGBN   	NvarChar(1),	U_FRGTAX   	NvarChar(1),
		U_FRGCOD   	NvarChar(2),	U_FRGNAM   	NvarChar(30),
		U_BUSNUM   	NvarChar(12),	U_CLTNAM   	NvarChar(50),
		U_COMPRT   	NvarChar(30),	U_PERNUM   	NvarChar(20),
		U_POSADD   	NvarChar(100),	U_PERNBR   	NvarChar(20),
		U_ADDRES   	NvarChar(100),	
		U_STRINT   	NvarChar(10), 	U_ENDINT   	NvarChar(10),
		U_RETRES	NvarChar(8), 	U_SPCGBN	NvarChar(1),
		--�ٹ�ó���ҵ�
		U_TJKPAY   	Numeric(19,6),	U_SUDAMT	Numeric(19,6),
		U_YILPAY  	Numeric(19,6),	U_RETPAY   	Numeric(19,6),	
		U_BTXPAY	Numeric(19,6), 	
		U_J01NAM	NvarChar(50), 	U_J01NBR	NvarChar(12),	
		U_JS1SPC	NvarChar(1),	U_JRET01   	Numeric(19,6),
		U_JSUD01	Numeric(19,6), 	U_JYIL01	Numeric(19,6),
		U_JTOT01   	Numeric(19,6), 	U_BTXP01		Numeric(19,6), 	
		U_MYNTOT	Numeric(19,6), 	U_MYNWON	Numeric(19,6), 		
		U_MYNBUL	Numeric(19,6), 	U_MYNGON	Numeric(19,6), 		
		U_JYNTOT	Numeric(19,6), 	U_JYNWON	Numeric(19,6), 		
		U_JYNBUL	Numeric(19,6),	U_JYNGON	Numeric(19,6),	
		U_SH1JIG	Numeric(19,6),	U_SH3JIG	Numeric(19,6),
		U_SH1TIL	Numeric(19,6),	U_SH1SUR	Numeric(19,6),	
		U_SH1GON	Numeric(19,6),	U_SH1GWA	Numeric(19,6),	
		U_SH1YAG	Numeric(19,6),	U_SH1YAS	Numeric(19,6),	
		U_STRRET   	NvarChar(10), U_ENDRET   	NvarChar(10),	
		U_GNMMON	INT,	U_EXPMON	Numeric(19,6), 
		U_JINDAT  	NvarChar(10), 	U_JOTDAT   	NvarChar(10),
		U_GNMDAY  	INT,	U_JEXMON  	INT,		
		U_JMMMON   	INT,	U_GNMYER   	INT,	
		U_RETGON	Numeric(19,6),	U_TAXSTD	Numeric(19,6),
		U_YTXSTD	Numeric(19,6),	U_YSANTX	Numeric(19,6),	
		U_SANTAX	Numeric(19,6),	U_SPCGON	Numeric(19,6),	
		U_TAXGON	Numeric(19,6),		
		U_GULGAB	Numeric(19,6),	U_GULJUM  	Numeric(19,6),
		U_JONGAB   	Numeric(19,6),	U_JONJUM	Numeric(19,6),
		U_CHAGAB  	Numeric(19,6),	U_CHAJUM   	Numeric(19,6),
		U_TAXNAM   	NvarChar(20)
		)       

-- <2.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
	INSERT INTO [#RPY401]
	SELECT	T0.DocNum,		T0.U_MSTCOD,	T0.U_MSTNAM,	T0.U_EmpID, 
			T5.U_INTGBN,	T5.U_DWEGBN,	T5.U_FRGTAX,	T2.HomeCountr AS U_FRGCOD, 
			'' AS U_FRGNAM, ISNULL(T4.U_BusNum, '') AS U_BUSNUM,
			ISNULL(T4.U_CLTName, '') AS U_CLTNAM,	ISNULL(T4.U_ComPrt, '') AS U_COMPRT,
			ISNULL(T4.U_PerNum, '') AS U_PERNUM,	ISNULL(T4.U_PosAdd, '') AS U_POSADD, 
			ISNULL(T2.GovID, '') AS U_PERNBR,		ISNULL(T2.HomeStreet, '') AS U_ADDRES,
			ISNULL(CONVERT(CHAR(10), T0.U_STRINT,20),'') AS U_STRINT, 
			ISNULL(CONVERT(CHAR(10), T0.U_ENDINT,20),'') AS U_ENDINT, 
			ISNULL(T0.U_RETRES,''), ISNULL(T0.U_SPCGBN,'N'), 
			T0.U_TJKPAY+T0.U_BHMAMT,	
			T0.U_SUDAMT,	
			T0.U_YILPAY,
			T0.U_RETPAY,
			T0.U_BTXPAY,
			T0.U_J01NAM,	
			T0.U_J01NBR, 
			ISNULL(T0.U_JS1SPC,'N'),
			T0.U_JRET01,
			T0.U_JSUD01,	
			T0.U_JYIL01,	
			T0.U_JTOT01, 
			T0.U_BTXP01,
			T0.U_MYNTOT,	T0.U_MYNWON,	T0.U_MYNBUL,	T0.U_MYNGON,	
			T0.U_JYNTOT,T0.U_JYNWON,	T0.U_JYNBUL,	T0.U_JYNGON,	
			T0.U_SH1JIG, 
			T0.U_SH3JIG, 
			T0.U_SH1TIL,	T0.U_SH1SUR,	T0.U_SH1GON,	T0.U_SH1GWA, 
			T0.U_SH1YAG,	T0.U_SH1YAS,
			ISNULL(CONVERT(CHAR(10), T0.U_STRRET,20),'') AS U_STRRET, 
			ISNULL(CONVERT(CHAR(10), T0.U_ENDRET,20),'') AS U_ENDRET, 
			T0.U_GNMMON,	
			T0.U_EXPMON, 	
			ISNULL(CONVERT(CHAR(10), T0.U_JINDAT,20),'') AS U_JINDAT, 
			ISNULL(CONVERT(CHAR(10), T0.U_JOTDAT,20),'') AS U_JOTDAT,
			T0.U_GNMDAY,
			T0.U_JEXMON,
			T0.U_JMMMON,	
			T0.U_GNMYER, 
			T0.U_RETGON, 
			T0.U_TAXSTD, 
			T0.U_YTXSTD,
			T0.U_YSANTX, 
			T0.U_SANTAX,	
			ISNULL(T0.U_SPCGON,0), ISNULL(T0.U_TAXGON,0), 
			T0.U_GULGAB, T0.U_GULJUM,	
			T0.U_JONGAB, T0.U_JONJUM,
			T0.U_CHAGAB, T0.U_CHAJUM,			
			ISNULL(T4.U_TAXName, '') AS U_TAXNAM	
			
	FROM	[@ZPY401H] T0 	
			INNER JOIN [OHEM] T2 ON T0.U_EmpID = T2.EmpID
			INNER JOIN [OUDP] T3 ON T2.Dept = T3.Code
			INNER JOIN [@ZPY127H] T5 ON T0.U_MSTCOD = T5.U_MSTCOD
			LEFT JOIN [@ZPY106H] T4 ON T4.CODE = T2.U_CLTCOD
	WHERE 	T0.U_ENDRET BETWEEN @STRDAT AND @ENDDAT
	AND		T0.U_JIGBIL BETWEEN @STRJIG AND @ENDJIG
	AND		ISNULL(T2.Branch, '')  LIKE @Branch
	AND		T3.U_MSTDPT LIKE @MSTDPT                        
	AND		T0.U_MSTCOD LIKE @MSTCOD
	--// �Ի��� <> ��������� AND �������� AND ��������� �����ϴ� ��� 
	--// ��2���Ŀ� ���;���. ���� �ϳ��� ���� ���ϴ� �͵鸸 ��1���Ŀ� ������ �� 
	AND	   (T0.U_INPDAT = T0.U_STRRET OR T0.U_SUDAMT = 0)
	AND		T0.U_JSNGBN LIKE CASE @JOBGBN WHEN '1' THEN '1' 
									 WHEN '2' THEN '2'
									 ELSE '%' END
	AND		T0.U_ENDRET <= '2009-12-31' -- 2009������� ���
	ORDER BY  T0.U_MSTNAM,  T0.U_MSTCOD
	
	-- <3.��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
	SELECT * FROM [#RPY401] ORDER BY U_MSTCOD, U_ENDRET


--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

