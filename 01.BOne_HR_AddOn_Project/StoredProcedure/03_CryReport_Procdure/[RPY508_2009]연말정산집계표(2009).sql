IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY508_2009' AND xtype = 'P'))
	DROP PROCEDURE RPY508_2009
GO


CREATE PROC RPY508_2009 (
        @JSNYER     AS Nvarchar(4),     --�۾�����
        @JOBGBN     AS Nvarchar(1),     --�۾�����(1��������,2�ߵ�����,3��ü)
        @CLTCOD     AS Nvarchar(8),     --�ڻ��ڵ�
        @MSTDPT     AS Nvarchar(8)      --�μ�
    ) 

 AS
    /*==========================================================================================
        ���ν�����      : RPY508_2009
        ���ν�������    : ������������ǥ
        ������          : �ֵ���
        �۾�����        : 2008-05-19
        �۾�������      : �Թ̰�
        �۾���������    : 2009-07-29
        �۾�����        : �ڻ��ڵ��߰�
        �۾�����        : 
    ===========================================================================================*/
    -- DROP PROC RPY508_2009
    -- Exec RPY508_2009 '2013','3','%','%'

    SET NOCOUNT ON

-- <1.�ӽ����̺� ���� >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    

    CREATE TABLE #RPY508_2009 (
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
			HU2CNT  NUMERIC(19,6),
            GBUCNT  NUMERIC(19,6),
            PILBHM  NUMERIC(19,6),
            PILMED  NUMERIC(19,6),
            PILSCH  NUMERIC(19,6),
            PILHUS  NUMERIC(19,6),
			PILHU2  NUMERIC(19,6),
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
			JH2CNT	NUMERIC(19,6),
			PILJH2	NUMERIC(19,6),
			JH3CNT	NUMERIC(19,6),
			PILJH3	NUMERIC(19,6),
			HUNCNT	NUMERIC(19,6),
			PILHUN	NUMERIC(19,6),
			SGICNT	NUMERIC(19,6),
			GITSGI	NUMERIC(19,6),
			GHSCNT	NUMERIC(19,6),
			GITHUS	NUMERIC(19,6),
			JFDCNT	NUMERIC(19,6),
			GITJFD	NUMERIC(19,6),
			GYUCNT	NUMERIC(19,6),
			GITGYU	NUMERIC(19,6),
			WOLCNT	NUMERIC(19,6),
			PILWOL	NUMERIC(19,6)
            ) 

-- <2.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    INSERT  INTO [#RPY508_2009]
    SELECT  EMPCNT  =   COUNT(T0.U_MSTCOD),                                             -- ���ο�
            PAYAMT  =   SUM(ISNULL(T0.U_PAYAMT,0)),                                     -- �޿��Ѿ�
            BNSAMT  =   SUM(ISNULL(T0.U_BNSAMT,0)) + SUM(ISNULL(T0.U_INBAMT,0))         -- ���Ѿ�(�� + ������ + ����ɼ� + �츮����)
                    +   SUM(ISNULL(T0.U_JUSAMT,0)) + SUM(ISNULL(T0.U_URIAMT,0)),
            BTXAM2  =   SUM(ISNULL(T0.U_BIGWA1,0)) + SUM(ISNULL(T0.U_BIGWA2,0))         -- �������(������ ����)
                    +   SUM(ISNULL(T0.U_BIGWA3,0)) + SUM(ISNULL(T0.U_BIGWA5,0)) 
                    +   SUM(ISNULL(T0.U_BIGWA6,0)) + SUM(ISNULL(T0.U_BIGWU3,0)) + SUM(ISNULL(T0.U_BIGWA4,0)) 

                    +   SUM(ISNULL(T0.U_BIGG01,0)) + SUM(ISNULL(T0.U_BIGH01,0)) 
                    +   SUM(ISNULL(T0.U_BIGH05,0)) + SUM(ISNULL(T0.U_BIGH06,0)) 
                    +   SUM(ISNULL(T0.U_BIGH07,0)) + SUM(ISNULL(T0.U_BIGH08,0)) 
                    +   SUM(ISNULL(T0.U_BIGH09,0)) + SUM(ISNULL(T0.U_BIGH10,0)) 
                    +   SUM(ISNULL(T0.U_BIGH11,0)) + SUM(ISNULL(T0.U_BIGH12,0)) 
                    +   SUM(ISNULL(T0.U_BIGH13,0)) + SUM(ISNULL(T0.U_BIGI01,0)) 
                    +   SUM(ISNULL(T0.U_BIGK01,0)) + SUM(ISNULL(T0.U_BIGM01,0)) 
                    +   SUM(ISNULL(T0.U_BIGM02,0)) + SUM(ISNULL(T0.U_BIGM03,0)) 
                    +   SUM(ISNULL(T0.U_BIGO01,0)) + SUM(ISNULL(T0.U_BIGQ01,0)) 
                    +   SUM(ISNULL(T0.U_BIGS01,0)) + SUM(ISNULL(T0.U_BIGT01,0)) 
                    +   SUM(ISNULL(T0.U_BIGX01,0)) + SUM(ISNULL(T0.U_BIGY01,0)) 
                    +   SUM(ISNULL(T0.U_BIGY02,0)) + SUM(ISNULL(T0.U_BIGY03,0)) 
                    +   SUM(ISNULL(T0.U_BIGY20,0)) + SUM(ISNULL(T0.U_BIGZ01,0)),
            BTXAM1  =   SUM(ISNULL(T0.U_BIGTOT,0)),                                     -- �������(������ ������)
            PAYAL1  =   SUM(ISNULL(T0.U_INCOME,0)) + SUM(ISNULL(T0.U_BIGWA1,0))         -- �ѱݾ�(������ ����)
                    +   SUM(ISNULL(T0.U_BIGWA2,0)) + SUM(ISNULL(T0.U_BIGWA3,0)) 
                    +   SUM(ISNULL(T0.U_BIGWA5,0)) + SUM(ISNULL(T0.U_BIGWA6,0))
                    +   SUM(ISNULL(T0.U_BIGWU3,0)) + SUM(ISNULL(T0.U_BIGWA4,0))

                    +   SUM(ISNULL(T0.U_BIGG01,0)) + SUM(ISNULL(T0.U_BIGH01,0)) 
                    +   SUM(ISNULL(T0.U_BIGH05,0)) + SUM(ISNULL(T0.U_BIGH06,0)) 
                    +   SUM(ISNULL(T0.U_BIGH07,0)) + SUM(ISNULL(T0.U_BIGH08,0)) 
                    +   SUM(ISNULL(T0.U_BIGH09,0)) + SUM(ISNULL(T0.U_BIGH10,0)) 
                    +   SUM(ISNULL(T0.U_BIGH11,0)) + SUM(ISNULL(T0.U_BIGH12,0)) 
                    +   SUM(ISNULL(T0.U_BIGH13,0)) + SUM(ISNULL(T0.U_BIGI01,0)) 
                    +   SUM(ISNULL(T0.U_BIGK01,0)) + SUM(ISNULL(T0.U_BIGM01,0)) 
                    +   SUM(ISNULL(T0.U_BIGM02,0)) + SUM(ISNULL(T0.U_BIGM03,0)) 
                    +   SUM(ISNULL(T0.U_BIGO01,0)) + SUM(ISNULL(T0.U_BIGQ01,0)) 
                    +   SUM(ISNULL(T0.U_BIGS01,0)) + SUM(ISNULL(T0.U_BIGT01,0)) 
                    +   SUM(ISNULL(T0.U_BIGX01,0)) + SUM(ISNULL(T0.U_BIGY01,0)) 
                    +   SUM(ISNULL(T0.U_BIGY02,0)) + SUM(ISNULL(T0.U_BIGY03,0)) 
                    +   SUM(ISNULL(T0.U_BIGY20,0)) + SUM(ISNULL(T0.U_BIGZ01,0)),
            PAYAL2  =   SUM(ISNULL(T0.U_INCOME,0)) + SUM(ISNULL(T0.U_BIGTOT,0)),        -- �ѱݾ�(������ ������)
            
            INCOME  =   SUM(ISNULL(T0.U_INCOME,0)),                                     -- �ٷμҵ�
            PILGNL  =   SUM(ISNULL(T0.U_PILGNL,0)),                                     -- �ٷμҵ����
            GNLOSD  =   SUM(ISNULL(T0.U_GNLOSD,0)),                                     -- �ٷμҵ�ݾ�
            INJBAS  =   SUM(ISNULL(T0.U_INJBAS,0)),                                     -- ���ΰ����ݾ�
            
            BAEWOO  =   SUM(ISNULL(T0.U_BAEWOO,0)),                                     -- ������ο�
            BUYNSU  =   SUM(ISNULL(T0.U_BUYNSU,0)),                                     -- �ξ簡���ο�
            GYNGLO  =   SUM(ISNULL(T0.U_GYNGLO,0)),                                     -- ��ο�� �ο�
            JANGAE  =   SUM(ISNULL(T0.U_JANGAE,0)),                                     -- ����� �ο�
            MZBURI  =   SUM(ISNULL(T0.U_MZBURI,0)),                                     -- �γ��� �ο�
            BUYN06  =   SUM(ISNULL(T0.U_BUYN06,0)),                                     -- 6������ �ڳ��ο�
            DAGYSU  =   SUM(ISNULL(T0.U_DAGYSU,0)),                                     -- ���ڳ� �ο�
            
            INJBWO  =   SUM(ISNULL(T0.U_INJBWO,0)),                                     -- ����ڰ����ݾ�
            INJBYN  =   SUM(ISNULL(T0.U_INJBYN,0)),                                     -- �ξ簡�������ݾ�
            INJGYN  =   SUM(ISNULL(T0.U_INJGYN,0)),                                     -- ��ο�� �����ݾ�
            INJJAE  =   SUM(ISNULL(T0.U_INJJAE,0)),                                     -- ����� �����ݾ�
            INJBNJ  =   SUM(ISNULL(T0.U_INJBNJ,0)),                                     -- �γ��� �����ݾ�
            INJSON  =   SUM(ISNULL(T0.U_INJSON,0)),                                     -- 6������ �ڳ���� �ݾ�
            INJADD  =   SUM(ISNULL(T0.U_INJADD,0)),                                     -- ���ڳ� �����ݾ�
            
            BHMCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILBHM > 0 OR T0.U_PILJHM > 0
                                          OR T0.U_PILMBH > 0 OR T0.U_PILGBH > 0
                                         THEN 1 ELSE 0 END,0)),     -- ����� �ο�
            MEDCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILMED > 0 THEN 1 ELSE 0 END,0)),     -- �Ƿ�� �ο�
            SCHCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILSCH > 0 THEN 1 ELSE 0 END,0)),     -- ������ �ο�
            HUSCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILHUS > 0 THEN 1 ELSE 0 END,0)),     -- �����ڱ� �ο�
			HU2CNT  =   SUM(ISNULL(CASE WHEN ISNULL(T0.U_PILHU2, 0) > 0 THEN 1 ELSE 0 END,0)),     -- �����ڱ� �ο�
            GBUCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILGBU > 0 THEN 1 ELSE 0 END,0)),     -- ��α� �ο�
            
            PILBHM  =   SUM(ISNULL(T0.U_PILBHM,0)) + SUM(ISNULL(T0.U_PILJHM,0)) 
                    +   SUM(ISNULL(T0.U_PILMBH,0)) + SUM(ISNULL(T0.U_PILGBH,0)),        -- ����� �����ݾ�
            PILMED  =   SUM(ISNULL(T0.U_PILMED,0)),                                     -- �Ƿ�� �����ݾ�
            PILSCH  =   SUM(ISNULL(T0.U_PILSCH,0)),                                     -- ������ �����ݾ�
            PILHUS  =   SUM(ISNULL(T0.U_PILHUS,0)),                                     -- �����ڱ� �����ݾ�
			PILHU2  =   SUM(ISNULL(T0.U_PILHU2,0)),                                     -- �����ڱ� �����ݾ�
            PILGBU  =   SUM(ISNULL(T0.U_PILGBU,0)),                                     -- ��α� �����ݾ�
            PILTOT  =   SUM(ISNULL(T0.U_PILTOT,0)) + SUM(ISNULL(T0.U_PILGON,0)),        -- �� �Ǵ� ǥ�ذ���
            
            GONCNT  =   SUM(ISNULL(CASE WHEN T0.U_YUNGON > 0 OR T0.U_YUNGO1 > 0 
                                          OR T0.U_YUNGO2 > 0 OR T0.U_YUNGO3 > 0 THEN 1 ELSE 0 END,0)),
            YUNGON  =   SUM(ISNULL(T0.U_YUNGON,0)) + SUM(ISNULL(T0.U_YUNGO1,0)) 
                    +   SUM(ISNULL(T0.U_YUNGO2,0)) + SUM(ISNULL(T0.U_YUNGO3,0)),          -- ���ݺ���� �����ݾ�

            CHAGAM  =   SUM(ISNULL(T0.U_CHAGAM,0)),                                     -- �����ҵ�ݾ�
            GYNCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITGYN > 0 THEN 1 ELSE 0 END,0)),     -- ���ο��ݼҵ���� �ο�
            YUNCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITYUN > 0 THEN 1 ELSE 0 END,0)),     -- ��������ҵ���� �ο�
            INVCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITINV > 0 THEN 1 ELSE 0 END,0)),     -- �������ռҵ���� �ο�
            CADCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITCAD > 0 THEN 1 ELSE 0 END,0)),     -- �ſ�ī��ҵ���� �ο�
            USJCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITUSJ > 0 THEN 1 ELSE 0 END,0)),     -- �츮�������ռҵ���� �ο�
            GITGYN  =   SUM(ISNULL(T0.U_GITGYN,0)),                                     -- ���ο��ݼҵ���� �ݾ�
            GITYUN  =   SUM(ISNULL(T0.U_GITYUN,0)),                                     -- ��������ҵ���� �ݾ�
            GITINV  =   SUM(ISNULL(T0.U_GITINV,0)),                                     -- �������ռҵ���� �ݾ�
            GITCAD  =   SUM(ISNULL(T0.U_GITCAD,0)),                                     -- �ſ�ī��ҵ���� �ݾ�
            GITUSJ  =   SUM(ISNULL(T0.U_GITUSJ,0)),                                     -- �츮�������ռҵ���� �ݾ�
            
            TAXCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXSTD > 0 THEN 1 ELSE 0 END,0)),     -- ���ռҵ����ǥ�� �ο�
            TAXSTD  =   SUM(ISNULL(T0.U_TAXSTD,0)),                                     -- ���ռҵ����ǥ��
            SANTAX  =   SUM(ISNULL(T0.U_SANTAX,0)),                                     -- ���⼼��

            TAXGNL  =   SUM(ISNULL(T0.U_TAXGNL,0)),                                     -- �ٷμҵ漼�װ���
            BROCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXBRO > 0 THEN 1 ELSE 0 END,0)),     -- �������Ա��ο�
            FRGCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXFRG > 0 THEN 1 ELSE 0 END,0)),     -- �ܱ������ο�
            NABCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXNAB > 0 THEN 1 ELSE 0 END,0)),     -- ���������ο�
            POLCNT  =   SUM(ISNULL(CASE WHEN T0.U_TAXGBU > 0 THEN 1 ELSE 0 END,0)),     -- �����ġ�ڱ� �ο�
            TAXBRO  =   SUM(ISNULL(T0.U_TAXBRO,0)),                                     -- �������Ա� ���װ���
            TAXFRG  =   SUM(ISNULL(T0.U_TAXFRG,0)),                                     -- �ܱ����� ���װ���
            TAXNAB  =   SUM(ISNULL(T0.U_TAXNAB,0)),                                     -- �������� ���װ���
            TAXGBU  =   SUM(ISNULL(T0.U_TAXGBU,0)),                                     -- �����ġ�ڱ� ���װ���
            TAXTOT  =   SUM(ISNULL(T0.U_TAXTOT,0)),                                     -- ���װ��� ��
            
            GAMSOD  =   SUM(ISNULL(T0.U_GAMSOD,0)),                                     -- �ҵ漼�� ���װ���
            GAMJOS  =   SUM(ISNULL(T0.U_GAMJOS,0)),                                     -- ����Ư�����ѹ� ���װ���
            GAMTOT  =   SUM(ISNULL(T0.U_GAMTOT,0)),                                     -- ���鼼�� ��
            
            GULCNT  =   SUM(ISNULL(CASE WHEN T0.U_GULGAB > 0 THEN 1 ELSE 0 END,0)),     -- ���������ο�
            GULGAB  =   SUM(ISNULL(T0.U_GULGAB,0)),                                     -- �����ҵ漼
            GULNON  =   SUM(ISNULL(T0.U_GULNON,0)),                                     -- ������Ư��
            GULJUM  =   SUM(ISNULL(T0.U_GULJUM,0)),                                     -- �����ֹμ�
            
            NANCNT  =   SUM(ISNULL(CASE WHEN T0.U_NANGAB > 0 THEN 1 ELSE 0 END,0)),
            NANGAB  =   SUM(ISNULL(T0.U_NANGAB,0)),                                     -- ���ٹ��� �ⳳ�� �ҵ漼
            NANNON  =   SUM(ISNULL(T0.U_NANNON,0)),                                     -- ���ٹ��� �ⳳ�� ��Ư��
            NANJUM  =   SUM(ISNULL(T0.U_NANJUM,0)),                                     -- ���ٹ��� �ⳳ�� �ֹμ�
            
            NALCNT  =   SUM(ISNULL(CASE WHEN T0.U_JONGAB > 0 THEN 1 ELSE 0 END,0)),
            NALGAB  =   SUM(ISNULL(T0.U_JONGAB,0)),                                     -- ���ٹ��� �ⳳ�� �ҵ漼
            NALNON  =   SUM(ISNULL(T0.U_JONNON,0)),                                     -- ���ٹ��� �ⳳ�� ��Ư��
            NALJUM  =   SUM(ISNULL(T0.U_JONJUM,0)),                                     -- ���ٹ��� �ⳳ�� �ֹμ�
            
            JSUCNT  =   SUM(ISNULL(CASE WHEN T0.U_CHAGAB > 0 THEN 1                ELSE 0 END,0)),  -- ¡�� �����ο�
            JSUGAB  =   SUM(ISNULL(CASE WHEN T0.U_CHAGAB > 0 THEN T0.U_CHAGAB      ELSE 0 END,0)),  -- ¡�� �ҵ漼
            JSUNON  =   SUM(ISNULL(CASE WHEN T0.U_CHANON > 0 THEN T0.U_CHANON      ELSE 0 END,0)),  -- ¡�� ��Ư��
            JSUJUM  =   SUM(ISNULL(CASE WHEN T0.U_CHAJUM > 0 THEN T0.U_CHAJUM      ELSE 0 END,0)),  -- ¡�� �ֹμ�
            
            HWACNT  =   SUM(ISNULL(CASE WHEN T0.U_CHAGAB < 0 THEN 1                ELSE 0 END,0)),  -- ȯ�� �ҵ漼
            HWAGAB  =   SUM(ISNULL(CASE WHEN T0.U_CHAGAB < 0 THEN T0.U_CHAGAB * -1 ELSE 0 END,0)),  -- ȯ�� �ҵ漼
            HWANON  =   SUM(ISNULL(CASE WHEN T0.U_CHANON < 0 THEN T0.U_CHANON * -1 ELSE 0 END,0)),  -- ȯ�� ��Ư��
            HWAJUM  =   SUM(ISNULL(CASE WHEN T0.U_CHAJUM < 0 THEN T0.U_CHAJUM * -1 ELSE 0 END,0)),   -- ȯ�� �ֹμ�,

            CHLSAN  =   SUM(ISNULL(CASE WHEN T0.U_INJCHL > 0 THEN 1                ELSE 0 END,0)),  -- ��꺸������ο�
            INJCHL  =   SUM(ISNULL(T0.U_INJCHL,0)),                                                 -- ��꺸������
            KUKCNT  =   SUM(ISNULL(CASE WHEN T0.U_KUKGON > 0 THEN 1                ELSE 0 END,0)),  -- ���ο��ݰ����ο�
            KUKGON  =   SUM(ISNULL(T0.U_KUKGON,0)),                                                 -- ���ο��� ����
            RETCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITRET > 0 OR T0.U_GITRE2 > 0 THEN 1 ELSE 0 END,0)),  -- �������ݰ����ο�
            GITRET  =   SUM(ISNULL(T0.U_GITRET,0)) + SUM(ISNULL(T0.U_GITRE2,0)),                    -- �������� ����
            JHECNT  =   SUM(ISNULL(CASE WHEN T0.U_PILJHE > 0 THEN 1                ELSE 0 END,0)),  -- ������������ο�
            PILJHE  =   SUM(ISNULL(T0.U_PILJHE,0)),                                                 -- ����������԰���
            JH2CNT  =   SUM(ISNULL(CASE WHEN T0.U_PILJH2 > 0 THEN 1                ELSE 0 END,0)),  -- ������������ο�
            PILJH2  =   SUM(ISNULL(T0.U_PILJH2,0)),                                                 -- ����������԰���
            JH3CNT  =   SUM(ISNULL(CASE WHEN T0.U_PILJH3 > 0 THEN 1                ELSE 0 END,0)),  -- ������������ο�
            PILJH3  =   SUM(ISNULL(T0.U_PILJH3,0)),                                                 -- ����������԰���
            HUNCNT  =   SUM(ISNULL(CASE WHEN T0.U_PILHUN > 0 THEN 1                ELSE 0 END,0)),  -- ȥ������̻�
            PILHUN  =   SUM(ISNULL(T0.U_PILHUN,0)),                                                 -- ȥ������̻����
            SGICNT  =   SUM(ISNULL(CASE WHEN T0.U_GITSGI > 0 THEN 1                ELSE 0 END,0)),  -- �ұ������
            GITSGI  =   SUM(ISNULL(T0.U_GITSGI,0)),                                                 -- �ұ������
            GHSCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITHUS > 0 OR T0.U_GITHU1 > 0 
                                          OR T0.U_GITHU2 > 0 OR T0.U_GITHU3 > 0 THEN 1 ELSE 0 END,0)),  -- ���ø�������
            GITHUS  =   SUM(ISNULL(T0.U_GITHUS,0)) + SUM(ISNULL(T0.U_GITHU1,0)) 
                    +   SUM(ISNULL(T0.U_GITHU2,0)) + SUM(ISNULL(T0.U_GITHU3,0)),                                                 -- ���ø����������
            JFDCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITJFD > 0 THEN 1                ELSE 0 END,0)),  -- ����ֽ�������
            GITJFD  =   SUM(ISNULL(T0.U_GITJFD,0)),                                                 -- ����ֽ����������
            GYUCNT  =   SUM(ISNULL(CASE WHEN T0.U_GITGYU > 0 THEN 1                ELSE 0 END,0)),  -- ��������߼ұ��
            GITGYU  =   SUM(ISNULL(T0.U_GITGYU,0)),                                                  -- ��������߼ұ���ҵ����
			WOLCNT	=	SUM(ISNULL(CASE WHEN T0.U_PILWOL > 0 THEN 1				   ELSE 0 END,0)),	-- ������
			PILWOL	=	SUM(ISNULL(T0.U_PILWOL,0))													-- �����װ���

    FROM    [@ZPY504H]  T0
            INNER JOIN [@PH_PY001A] T1 ON T0.U_MstCod = T1.Code
            --INNER JOIN [OUDP] T2 ON T1.Dept     = T2.Code
    WHERE   T0.U_JSNYER     =       @JSNYER
    AND     (T0.U_JSNGBN    =       @JOBGBN
    OR      @JOBGBN         =       '3')
    AND     T0.U_CLTCOD     LIKE    @CLTCOD         
    AND     T1.U_TeamCode     LIKE    @MSTDPT                        

-- <3.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    
    SELECT * FROM [#RPY508_2009] 
    
--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF
