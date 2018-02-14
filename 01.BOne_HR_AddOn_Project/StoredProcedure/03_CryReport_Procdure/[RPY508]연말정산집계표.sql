IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY508' AND xtype = 'P'))
	DROP PROCEDURE RPY508
GO

CREATE PROC RPY508
    (
        @JSNYER     AS Nvarchar(4),     --�۾�����
        @JOBGBN     AS Nvarchar(1),     --�۾�����(1��������,2�ߵ�����,3��ü)
        @CLTCOD     AS Nvarchar(8),     --�ڻ��ڵ�
        @MSTDPT     AS Nvarchar(8)      --�μ�

    )
    

 AS
    /*==========================================================================================
        ���ν�����      : RPY508
        ���ν�������    : ������������ǥ
        ������          : �ֵ���
        �۾�����        : 2008-05-19
        �۾�������      : �Թ̰�
        �۾���������    : 2009-07-29
        �۾�����        : �ڻ��ڵ��߰�
        �۾�����        : 
    ===========================================================================================*/
    -- DROP PROC RPY508
    -- Exec RPY508  '2007', '3', N'%', N'%'

    SET NOCOUNT ON

-- <1.�ӽ����̺� ���� >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    

    CREATE TABLE #RPY508 (
            EMPCNT  NUMERIC(19,6),
            PAYAMT  NUMERIC(19,6),
            BNSAMT  NUMERIC(19,6),
            BTXAM2  NUMERIC(19,6),
            BTXAM1  NUMERIC(19,6),
            PAYAL1  NUMERIC(19,6),
            PAYAL2  NUMERIC(19,6),
            INCOME  NUMERIC(19,6),
            PILGNL  NUMERIC(19,6),
            GNLOSD  NUMERIC(19,6),
            INJBAS  NUMERIC(19,6),
            BAEWOO  NUMERIC(19,6),
            BUYNSU  NUMERIC(19,6),
            GYNGLO  NUMERIC(19,6),
            JANGAE  NUMERIC(19,6),
            MZBURI  NUMERIC(19,6),
            BUYN06  NUMERIC(19,6),
            DAGYSU  NUMERIC(19,6),
            INJBWO  NUMERIC(19,6),
            INJBYN  NUMERIC(19,6),
            INJGYN  NUMERIC(19,6),
            INJJAE  NUMERIC(19,6),
            INJBNJ  NUMERIC(19,6),
            INJSON  NUMERIC(19,6),
            INJADD  NUMERIC(19,6),
            BHMCNT  NUMERIC(19,6),
            MEDCNT  NUMERIC(19,6),
            SCHCNT  NUMERIC(19,6),
            HUSCNT  NUMERIC(19,6),
            GBUCNT  NUMERIC(19,6),
            PILBHM  NUMERIC(19,6),
            PILMED  NUMERIC(19,6),
            PILSCH  NUMERIC(19,6),
            PILHUS  NUMERIC(19,6),
            PILGBU  NUMERIC(19,6),
            PILTOT  NUMERIC(19,6),
            GONCNT  NUMERIC(19,6),
            YUNGON  NUMERIC(19,6),
            CHAGAM  NUMERIC(19,6),
            GYNCNT  NUMERIC(19,6),
            YUNCNT  NUMERIC(19,6),
            INVCNT  NUMERIC(19,6),
            CADCNT  NUMERIC(19,6),
            USJCNT  NUMERIC(19,6),
            GITGYN  NUMERIC(19,6),
            GITYUN  NUMERIC(19,6),
            GITINV  NUMERIC(19,6),
            GITCAD  NUMERIC(19,6),
            GITUSJ  NUMERIC(19,6),
            TAXCNT  NUMERIC(19,6),
            TAXSTD  NUMERIC(19,6),
            SANTAX  NUMERIC(19,6),
            TAXGNL  NUMERIC(19,6),
            BROCNT  NUMERIC(19,6),
            FRGCNT  NUMERIC(19,6),
            NABCNT  NUMERIC(19,6),
            POLCNT  NUMERIC(19,6),
            TAXBRO  NUMERIC(19,6),
            TAXFRG  NUMERIC(19,6),
            TAXNAB  NUMERIC(19,6),
            TAXGBU  NUMERIC(19,6),
            TAXTOT  NUMERIC(19,6),
            GAMSOD  NUMERIC(19,6),
            GAMJOS  NUMERIC(19,6),
            GAMTOT  NUMERIC(19,6),
            GULCNT  NUMERIC(19,6),
            GULGAB  NUMERIC(19,6),
            GULNON  NUMERIC(19,6),
            GULJUM  NUMERIC(19,6),
            NANCNT  NUMERIC(19,6),
            NANGAB  NUMERIC(19,6),
            NANNON  NUMERIC(19,6),
            NANJUM  NUMERIC(19,6),
            NALCNT  NUMERIC(19,6),
            NALGAB  NUMERIC(19,6),
            NALNON  NUMERIC(19,6),
            NALJUM  NUMERIC(19,6),
            JSUCNT  NUMERIC(19,6),
            JSUGAB  NUMERIC(19,6),
            JSUNON  NUMERIC(19,6),
            JSUJUM  NUMERIC(19,6),
            HWACNT  NUMERIC(19,6),
            HWAGAB  NUMERIC(19,6),
            HWANON  NUMERIC(19,6),
            HWAJUM  NUMERIC(19,6),
			CHLSAN	NUMERIC(19,6),
			INJCHL	NUMERIC(19,6),
			KUKCNT	NUMERIC(19,6),
			KUKGON	NUMERIC(19,6),
			RETCNT	NUMERIC(19,6),
			GITRET	NUMERIC(19,6),
			JHECNT	NUMERIC(19,6),
			PILJHE	NUMERIC(19,6),
			HUNCNT	NUMERIC(19,6),	-- ��ȥ.�̻�.��� �ο���  (2009������ ������)
			PILHUN	NUMERIC(19,6),	-- ��ȥ.�̻�.��� �����ݾ�
			SGICNT	NUMERIC(19,6),
			GITSGI	NUMERIC(19,6),
			GHSCNT	NUMERIC(19,6),
			GITHUS	NUMERIC(19,6),
			JFDCNT	NUMERIC(19,6),
			GITJFD	NUMERIC(19,6)
            ) 

-- <2.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    INSERT  INTO [#RPY508]
    SELECT  EMPCNT  =   COUNT(T0.U_MSTCOD),                                             -- ���ο�
            PAYAMT  =   ISNULL(SUM(T0.U_PAYAMT),0),                                     -- �޿��Ѿ�
            BNSAMT  =   ISNULL(SUM(T0.U_BNSAMT),0) + ISNULL(SUM(T0.U_INBAMT),0)         -- ���Ѿ�
                    +   ISNULL(SUM(T0.U_JUSAMT),0),
            BTXAM2  =   ISNULL(SUM(T0.U_BIGWA1),0) + ISNULL(SUM(T0.U_BIGWA2),0)         -- �������(������ ����)
                    +   ISNULL(SUM(T0.U_BIGWA3),0) + ISNULL(SUM(T0.U_BIGWA5),0) 
                    +   ISNULL(SUM(T0.U_BIGWA6),0) + ISNULL(SUM(T0.U_BIGWU3),0) + ISNULL(SUM(T0.U_BIGWA4),0),
            BTXAM1  =   ISNULL(SUM(T0.U_BIGTOT),0),                                     -- �������(������ ������)
            PAYAL1  =   ISNULL(SUM(T0.U_INCOME),0) + ISNULL(SUM(T0.U_BIGWA1),0)         -- �ѱݾ�(������ ����)
                    +   ISNULL(SUM(T0.U_BIGWA2),0) + ISNULL(SUM(T0.U_BIGWA3),0) 
                    +   ISNULL(SUM(T0.U_BIGWA5),0) + ISNULL(SUM(T0.U_BIGWA6),0)
                    +   ISNULL(SUM(T0.U_BIGWU3),0) + ISNULL(SUM(T0.U_BIGWA4),0),
            PAYAL2  =   ISNULL(SUM(T0.U_INCOME),0) + ISNULL(SUM(T0.U_BIGTOT),0),        -- �ѱݾ�(������ ������)
            
            INCOME  =   ISNULL(SUM(T0.U_INCOME),0),                                     -- �ٷμҵ�
            PILGNL  =   ISNULL(SUM(T0.U_PILGNL),0),                                     -- �ٷμҵ����
            GNLOSD  =   ISNULL(SUM(T0.U_GNLOSD),0),                                     -- �ٷμҵ�ݾ�
            INJBAS  =   ISNULL(SUM(T0.U_INJBAS),0),                                     -- ���ΰ����ݾ�
            
            BAEWOO  =   ISNULL(SUM(T0.U_BAEWOO),0),                                     -- ������ο�
            BUYNSU  =   ISNULL(SUM(T0.U_BUYNSU),0),                                     -- �ξ簡���ο�
            GYNGLO  =   ISNULL(SUM(T0.U_GYNGLO),0),                                     -- ��ο�� �ο�
            JANGAE  =   ISNULL(SUM(T0.U_JANGAE),0),                                     -- ����� �ο�
            MZBURI  =   ISNULL(SUM(T0.U_MZBURI),0),                                     -- �γ��� �ο�
            BUYN06  =   ISNULL(SUM(T0.U_BUYN06),0),                                     -- 6������ �ڳ��ο�
            DAGYSU  =   ISNULL(SUM(T0.U_DAGYSU),0),                                     -- ���ڳ� �ο�
            
            INJBWO  =   ISNULL(SUM(T0.U_INJBWO),0),                                     -- ����ڰ����ݾ�
            INJBYN  =   ISNULL(SUM(T0.U_INJBYN),0),                                     -- �ξ簡�������ݾ�
            INJGYN  =   ISNULL(SUM(T0.U_INJGYN),0),                                     -- ��ο�� �����ݾ�
            INJJAE  =   ISNULL(SUM(T0.U_INJJAE),0),                                     -- ����� �����ݾ�
            INJBNJ  =   ISNULL(SUM(T0.U_INJBNJ),0),                                     -- �γ��� �����ݾ�
            INJSON  =   ISNULL(SUM(T0.U_INJSON),0),                                     -- 6������ �ڳ���� �ݾ�
            INJADD  =   ISNULL(SUM(T0.U_INJADD),0),                                     -- ���ڳ� �����ݾ�
            
            BHMCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILBHM > 0 THEN 1 ELSE 0 END),0),     -- ����� �ο�
            MEDCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILMED > 0 THEN 1 ELSE 0 END),0),     -- �Ƿ�� �ο�
            SCHCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILSCH > 0 THEN 1 ELSE 0 END),0),     -- ������ �ο�
            HUSCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILHUS > 0 THEN 1 ELSE 0 END),0),     -- �����ڱ� �ο�
            GBUCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILGBU > 0 THEN 1 ELSE 0 END),0),     -- ��α� �ο�
            
            PILBHM  =   ISNULL(SUM(T0.U_PILBHM),0),                                     -- ����� �����ݾ�
            PILMED  =   ISNULL(SUM(T0.U_PILMED),0),                                     -- �Ƿ�� �����ݾ�
            PILSCH  =   ISNULL(SUM(T0.U_PILSCH),0),                                     -- ������ �����ݾ�
            PILHUS  =   ISNULL(SUM(T0.U_PILHUS),0),                                     -- �����ڱ� �����ݾ�
            PILGBU  =   ISNULL(SUM(T0.U_PILGBU),0),                                     -- ��α� �����ݾ�
            PILTOT  =   ISNULL(SUM(T0.U_PILTOT),0) + ISNULL(SUM(T0.U_PILGON),0),        -- �� �Ǵ� ǥ�ذ���
            
            GONCNT  =   ISNULL(SUM(CASE WHEN T0.U_YUNGON > 0 THEN 1 ELSE 0 END),0),
            YUNGON  =   ISNULL(SUM(T0.U_KUKGON),0)         -- ���ݺ���� �����ݾ�
                    +   ISNULL(SUM(T0.U_GITRET),0),

            CHAGAM  =   ISNULL(SUM(T0.U_CHAGAM),0),                                     -- �����ҵ�ݾ�
            GYNCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITGYN > 0 THEN 1 ELSE 0 END),0),     -- ���ο��ݼҵ���� �ο�
            YUNCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITYUN > 0 THEN 1 ELSE 0 END),0),     -- ��������ҵ���� �ο�
            INVCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITINV > 0 THEN 1 ELSE 0 END),0),     -- �������ռҵ���� �ο�
            CADCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITCAD > 0 THEN 1 ELSE 0 END),0),     -- �ſ�ī��ҵ���� �ο�
            USJCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITUSJ > 0 THEN 1 ELSE 0 END),0),     -- �츮�������ռҵ���� �ο�
            GITGYN  =   ISNULL(SUM(T0.U_GITGYN),0),                                     -- ���ο��ݼҵ���� �ݾ�
            GITYUN  =   ISNULL(SUM(T0.U_GITYUN),0),                                     -- ��������ҵ���� �ݾ�
            GITINV  =   ISNULL(SUM(T0.U_GITINV),0),                                     -- �������ռҵ���� �ݾ�
            GITCAD  =   ISNULL(SUM(T0.U_GITCAD),0),                                     -- �ſ�ī��ҵ���� �ݾ�
            GITUSJ  =   ISNULL(SUM(T0.U_GITUSJ),0),                                     -- �츮�������ռҵ���� �ݾ�
            
            TAXCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXSTD > 0 THEN 1 ELSE 0 END),0),     -- ���ռҵ����ǥ�� �ο�
            TAXSTD  =   ISNULL(SUM(T0.U_TAXSTD),0),                                     -- ���ռҵ����ǥ��
            SANTAX  =   ISNULL(SUM(T0.U_SANTAX),0),                                     -- ���⼼��

            TAXGNL  =   ISNULL(SUM(T0.U_TAXGNL),0),                                     -- �ٷμҵ漼�װ���
            BROCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXBRO > 0 THEN 1 ELSE 0 END),0),     -- �������Ա��ο�
            FRGCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXFRG > 0 THEN 1 ELSE 0 END),0),     -- �ܱ������ο�
            NABCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXNAB > 0 THEN 1 ELSE 0 END),0),     -- ���������ο�
            POLCNT  =   ISNULL(SUM(CASE WHEN T0.U_TAXGBU > 0 THEN 1 ELSE 0 END),0),     -- �����ġ�ڱ� �ο�
            TAXBRO  =   ISNULL(SUM(T0.U_TAXBRO),0),                                     -- �������Ա� ���װ���
            TAXFRG  =   ISNULL(SUM(T0.U_TAXFRG),0),                                     -- �ܱ����� ���װ���
            TAXNAB  =   ISNULL(SUM(T0.U_TAXNAB),0),                                     -- �������� ���װ���
            TAXGBU  =   ISNULL(SUM(T0.U_TAXGBU),0),                                     -- �����ġ�ڱ� ���װ���
            TAXTOT  =   ISNULL(SUM(T0.U_TAXTOT),0),                                     -- ���װ��� ��
            
            GAMSOD  =   ISNULL(SUM(T0.U_GAMSOD),0),                                     -- �ҵ漼�� ���װ���
            GAMJOS  =   ISNULL(SUM(T0.U_GAMJOS),0),                                     -- ����Ư�����ѹ� ���װ���
            GAMTOT  =   ISNULL(SUM(T0.U_GAMTOT),0),                                     -- ���鼼�� ��
            
            GULCNT  =   ISNULL(SUM(CASE WHEN T0.U_GULGAB > 0 THEN 1 ELSE 0 END),0),     -- ���������ο�
            GULGAB  =   ISNULL(SUM(T0.U_GULGAB),0),                                     -- �����ҵ漼
            GULNON  =   ISNULL(SUM(T0.U_GULNON),0),                                     -- ������Ư��
            GULJUM  =   ISNULL(SUM(T0.U_GULJUM),0),                                     -- �����ֹμ�
            
            NANCNT  =   ISNULL(SUM(CASE WHEN T0.U_NANGAB > 0 THEN 1 ELSE 0 END),0),
            NANGAB  =   ISNULL(SUM(T0.U_NANGAB),0),                                     -- ���ٹ��� �ⳳ�� �ҵ漼
            NANNON  =   ISNULL(SUM(T0.U_NANNON),0),                                     -- ���ٹ��� �ⳳ�� ��Ư��
            NANJUM  =   ISNULL(SUM(T0.U_NANJUM),0),                                     -- ���ٹ��� �ⳳ�� �ֹμ�
            
            NALCNT  =   ISNULL(SUM(CASE WHEN T0.U_JONGAB > 0 THEN 1 ELSE 0 END),0),
            NALGAB  =   ISNULL(SUM(T0.U_JONGAB),0),                                     -- ���ٹ��� �ⳳ�� �ҵ漼
            NALNON  =   ISNULL(SUM(T0.U_JONNON),0),                                     -- ���ٹ��� �ⳳ�� ��Ư��
            NALJUM  =   ISNULL(SUM(T0.U_JONJUM),0),                                     -- ���ٹ��� �ⳳ�� �ֹμ�
            
            JSUCNT  =   ISNULL(SUM(CASE WHEN T0.U_CHAGAB > 0 THEN 1                ELSE 0 END),0),  -- ¡�� �����ο�
            JSUGAB  =   ISNULL(SUM(CASE WHEN T0.U_CHAGAB > 0 THEN T0.U_CHAGAB      ELSE 0 END),0),  -- ¡�� �ҵ漼
            JSUNON  =   ISNULL(SUM(CASE WHEN T0.U_CHANON > 0 THEN T0.U_CHANON      ELSE 0 END),0),  -- ¡�� ��Ư��
            JSUJUM  =   ISNULL(SUM(CASE WHEN T0.U_CHAJUM > 0 THEN T0.U_CHAJUM      ELSE 0 END),0),  -- ¡�� �ֹμ�
            
            HWACNT  =   ISNULL(SUM(CASE WHEN T0.U_CHAGAB < 0 THEN 1                ELSE 0 END),0),  -- ȯ�� �ҵ漼
            HWAGAB  =   ISNULL(SUM(CASE WHEN T0.U_CHAGAB < 0 THEN T0.U_CHAGAB * -1 ELSE 0 END),0),  -- ȯ�� �ҵ漼
            HWANON  =   ISNULL(SUM(CASE WHEN T0.U_CHANON < 0 THEN T0.U_CHANON * -1 ELSE 0 END),0),  -- ȯ�� ��Ư��
            HWAJUM  =   ISNULL(SUM(CASE WHEN T0.U_CHAJUM < 0 THEN T0.U_CHAJUM * -1 ELSE 0 END),0),   -- ȯ�� �ֹμ�,

            CHLSAN  =   ISNULL(SUM(CASE WHEN T0.U_INJCHL > 0 THEN 1                ELSE 0 END),0),  -- ��꺸������ο�
            INJCHL  =   ISNULL(SUM(T0.U_INJCHL),0),  -- ��꺸������
            KUKCNT  =   ISNULL(SUM(CASE WHEN T0.U_KUKGON > 0 THEN 1                ELSE 0 END),0),  -- ���ο��ݰ����ο�
            KUKGON  =   ISNULL(SUM(T0.U_KUKGON),0),  -- ���ο��� ����
            RETCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITRET > 0 THEN 1                ELSE 0 END),0),  -- �������ݰ����ο�
            GITRET  =   ISNULL(SUM(T0.U_GITRET),0),  -- �������� ����
            JHECNT  =   ISNULL(SUM(CASE WHEN T0.U_PILJHE > 0 THEN 1                ELSE 0 END),0),  -- ������������ο�
            PILJHE  =   ISNULL(SUM(T0.U_PILJHE),0),  -- ����������԰���
            HUNCNT  =   ISNULL(SUM(CASE WHEN T0.U_PILHUN > 0 THEN 1                ELSE 0 END),0),  -- ȥ������̻�(2009������ ������)
            PILHUN  =   ISNULL(SUM(T0.U_PILHUN),0),  -- ȥ������̻����
            SGICNT  =   ISNULL(SUM(CASE WHEN T0.U_GITSGI > 0 THEN 1                ELSE 0 END),0),  -- �ұ������
            GITSGI  =   ISNULL(SUM(T0.U_GITSGI),0),  -- �ұ������
            GHSCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITHUS > 0 THEN 1                ELSE 0 END),0),  -- ���ø�������
            GITHUS  =   ISNULL(SUM(T0.U_GITHUS),0),  -- ���ø����������
            JFDCNT  =   ISNULL(SUM(CASE WHEN T0.U_GITJFD > 0 THEN 1                ELSE 0 END),0),  -- ����ֽ�������
            GITJFD  =   ISNULL(SUM(T0.U_GITJFD),0)  -- ����ֽ����������

    FROM    [@ZPY504H]  T0
            --INNER JOIN [OHEM] T1 ON T0.U_MSTCOD = T1.U_MSTCOD
            INNER JOIN [@PH_PY001A] T1 ON T0.U_MstCod = T1.Code
            --INNER JOIN [OUDP] T2 ON T1.Dept     = T2.Code
    WHERE   T0.U_JSNYER     =       @JSNYER
    AND     (T0.U_JSNGBN    =       @JOBGBN
    OR      @JOBGBN         =       '3')
    AND     T0.U_CLTCOD     LIKE    @CLTCOD                        
    AND     T1.U_TeamCode   LIKE    @MSTDPT                        

-- <3.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    
    SELECT * FROM [#RPY508] 
    
--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF
