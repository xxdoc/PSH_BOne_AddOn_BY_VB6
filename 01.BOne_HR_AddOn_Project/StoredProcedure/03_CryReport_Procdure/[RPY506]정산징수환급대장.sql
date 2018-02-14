IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY506' AND xtype = 'P'))
	DROP PROCEDURE RPY506
GO

CREATE  PROC RPY506 (
        @JSNYER     AS Nvarchar(4),     --�۾�����
        @STRMON     AS Nvarchar(2),     --���ۿ�
        @ENDMON     AS Nvarchar(2),     --�����
        @JOBGBN     AS Nvarchar(1),     --�۾�����(1��������,2�ߵ�����,3��ü)
        @CLTCOD     AS Nvarchar(8),     --�ڻ��ڵ�
        @MSTDPT     AS Nvarchar(8),     --�μ�
        @MSTCOD     AS Nvarchar(8)      --�����ȣ      
    )
   

 AS
    /*==========================================================================================
        ���ν�����      : RPY506
        ���ν�������    : ����¡��ȯ�޴���
        ������          : �Թ̰�
        �۾�����        : 2007-01-30
        �۾�������      : �Թ̰�
        �۾���������    : 2007-01-30
        �۾�����        : 
        �۾�����        : (2009.10.22 �Թ̰�) �ܱ��κ�����ϰ�� �����ݾ��� +�ܱ��κ��2��� ����.���μ���
    ===========================================================================================*/
	--DROP PROC RPY506
   
    SET NOCOUNT ON

-- <1.�ӽ����̺� ���� >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
--1.2)
    CREATE TABLE #RPY506 (
            JSNYER   nvarchar(4),
            MSTCOD   nvarchar(8),
            MSTNAM   nvarchar(50),   
            PAYAMT   Numeric(19,6),
            BNSAMT   Numeric(19,6),		--5
            TOTAMT   Numeric(19,6),
            BIGTOT   Numeric(19,6),
            INCOME   Numeric(19,6),
            GULGAB   Numeric(19,6),
            GULNON   Numeric(19,6),		--10
            GULJUM   Numeric(19,6),
            NANGAB   Numeric(19,6),
            NANNON   Numeric(19,6), 
            NANJUM   Numeric(19,6),
            CHAGAB   Numeric(19,6),		--15
            CHANON   Numeric(19,6),
            CHAJUM   Numeric(19,6),
            JONPAY   Numeric(19,6),
            JONBNS   Numeric(19,6),
            JONTOT   Numeric(19,6),		--20
            TOTPAY   Numeric(19,6),
            TOTBNS   Numeric(19,6),
            TOTTAL   Numeric(19,6),
            MSTBRK   nvarchar(8),
            BRKNAM   nvarchar(50),		
            CLTCOD   nvarchar(8),
            CLTNAM   nvarchar(50),
            MSTDPT   nvarchar(10),
            DPTNAM   nvarchar(20),
            MSTSTP   nvarchar(10),
            STPNAM   nvarchar(20),
            ) 
            
-- <2.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    INSERT INTO [#RPY506]
    SELECT  JSNYER   =	T0.U_JSNYER, 
            MSTCOD   =	T0.U_MSTCOD,
            MSTNAM   =	MAX(T0.U_MSTNAM),
			PAYAMT   =	SUM(T0.U_PAYAMT+T0.U_BIGTOT-ISNULL(T0.U_BIGWU3,0)-ISNULL(T1.JONBTX,0)),  --����
			TOTAMT   =  SUM(T0.U_TOTAMT+T0.U_BIGTOT-ISNULL(T0.U_BIGWU3,0)-ISNULL(T1.JONBTX,0)), --�����Ѱ�
			BIGTOT   =  SUM(T0.U_BIGTOT),                --�����
			TOTTAL   =  SUM(T0.U_TOTAMT +ISNULL(T1.JONTOT,0)+T0.U_BIGTOT),    
			--�����޿� (�ѱ޿�+�����-�����������ܺ����-���������)
            --PAYAMT   =	SUM(ISNULL(T0.U_PAYAMT,0)+ISNULL(T0.U_BIGTOT,0)+ISNULL(T0.U_BIGWA2,0)-ISNULL(T0.U_BIGWU3,0)-ISNULL(T1.JONBTX,0)),  
			--������
            --BNSAMT   =	SUM(ISNULL(T0.U_BNSAMT,0)+ISNULL(T0.U_INBAMT,0)+ISNULL(T0.U_JUSAMT,0)+ISNULL(T0.U_URIAMT,0)),
			--�����Ѱ�
            --TOTAMT   =  SUM(ISNULL(T0.U_TOTAMT,0)+ISNULL(T0.U_BIGTOT,0)+ISNULL(T0.U_BIGWA2,0)-ISNULL(T0.U_BIGWU3,0)-ISNULL(T1.JONBTX,0)), 
			--�����=��������������Ѱ�+�����������ܺ����
            --BIGTOT   =  SUM(ISNULL(T0.U_BIGTOT,0)+ISNULL(T0.U_BIGWA2,0)),    
			--����
            INCOME   =  SUM(T0.U_INCOME),                
            GULGAB   =  SUM(T0.U_GULGAB),
            GULNON   =  SUM(T0.U_GULNON),
            GULJUM   =  SUM(T0.U_GULJUM),
            NANGAB   =  SUM(T0.U_NANGAB+T0.U_JONGAB),
            NANNON   =  SUM(T0.U_NANNON+T0.U_JONNON),
            NANJUM   =  SUM(T0.U_NANJUM+T0.U_JONJUM),
            CHAGAB   =  SUM(T0.U_CHAGAB),
            CHANON   =  SUM(T0.U_CHANON),
            CHAJUM   =  SUM(T0.U_CHAJUM),
            JONPAY   =  ISNULL(SUM(T1.JONPAY+T1.JONBTX), 0),  --�����޿�
            JONBNS   =  ISNULL(SUM(T1.JONBNS), 0),            --������
            JONTOT   =  ISNULL(SUM(T1.JONTOT+T1.JONBTX), 0),  --�����Ѱ�
            TOTPAY   =  SUM(T0.U_PAYAMT +ISNULL(T1.JONPAY,0)+ISNULL(T0.U_BIGTOT,0)+ISNULL(T0.U_BIGWA2,0)-ISNULL(T0.U_BIGWU3,0)),						--�޿���
            TOTBNS   =  SUM(T0.U_BNSAMT +ISNULL(T0.U_INBAMT,0)+ISNULL(T0.U_JUSAMT,0)+ISNULL(T0.U_URIAMT,0)+ISNULL(T1.JONBNS,0)),    --�󿩰�
            TOTTAL   =  SUM(T0.U_TOTAMT +ISNULL(T1.JONTOT,0)+ISNULL(T0.U_BIGTOT,0)+ISNULL(T0.U_BIGWA2,0)-ISNULL(T0.U_BIGWU3,0)),
            MSTBRK   =  '', --ISNULL(MAX(T2.Branch),'')   ,
            BRKNAM   =  '', --ISNULL(MAX(T2.U_Name),'')     ,
            CLTCOD   =  ISNULL(MAX(T0.U_CLTCOD),'') ,
            CLTNAM   =  ISNULL(MAX(T4.U_CLTName),''),
            MSTDPT   =  ISNULL(MAX(T2.U_TeamCode),''), 
            DPTNAM   =  '', --ISNULL(MAX(T3.NAME),'')     ,
            MSTSTP   =  '', --ISNULL(MAX(T5.U_MSTSTP),'') ,
            STPNAM   =  '' --ISNULL(MAX(T5.NAME),'')     
    
    FROM  	[@ZPY504H] T0 
    		LEFT JOIN (
            SELECT  U_JSNYER	=	T1.U_JSNYER, 
                    U_CLTCOD	=	T1.U_CLTCOD,
                    U_MSTCOD	=	T1.U_MSTCOD,
                    JONPAY		=	SUM(ISNULL(T0.U_JONPAY,0)),
                    JONBNS		=	SUM(ISNULL(T0.U_JONBNS,0)+ISNULL(T0.U_INJBNS,0)+ISNULL(T0.U_JONJUS,0)+ISNULL(T0.U_URIBNS,0)),
                    JONBTX		=	SUM(ISNULL(T0.U_JONBT1,0)+ISNULL(T0.U_JONBT2,0)+ISNULL(T0.U_JONBT3,0)),
                    JONTOT		=	SUM(ISNULL(T0.U_JONPAY,0)+ISNULL(T0.U_JONBNS,0)+ISNULL(T0.U_INJBNS,0)+ISNULL(T0.U_JONJUS,0)+ISNULL(T0.U_URIBNS,0))
            FROM 	[@ZPY502L] T0 
            		INNER JOIN [@ZPY502H] T1 ON T0.DocEntry = T1.DocEntry
            WHERE   T1.U_JSNYER = @JSNYER
            AND     T1.U_CLTCOD LIKE @CLTCOD
            GROUP BY T1.U_JSNYER, T1.U_CLTCOD, T1.U_MSTCOD
            )                     T1 ON T0.U_JSNYER = T1.U_JSNYER 
                                    AND T0.U_MSTCOD = T1.U_MSTCOD 
                                    AND T0.U_CLTCOD = T1.U_CLTCOD
            INNER JOIN [@PH_PY001A]     T2 ON T0.U_EmpID  = T2.U_EmpID
            --INNER JOIN [OUDP]     T3 ON T2.Dept     = T3.Code
            INNER JOIN [OHPS]     T5 ON T2.U_position = T5.posID
            --LEFT  JOIN [@ZPY106H] T4 ON T0.U_CLTCOD = T4.Code
            LEFT JOIN [@PH_PY005A] T4 ON T0.U_CLTCOD = T4.U_CLTCode                                    
            --LEFT  JOIN [OUBR]     T6 ON T2.Branch   = T6.Code
    WHERE 	T0.U_JSNYER                    =    @JSNYER
    AND   	T0.U_CLTCOD                    LIKE @CLTCOD                        
    AND   	T2.U_TeamCode				   LIKE @MSTDPT                        
    AND   	T0.U_MSTCOD                    LIKE @MSTCOD
    AND    (T0.U_INCOME+T0.U_BIGTOT) > 0           
    AND    (@JOBGBN ='3' OR (@JOBGBN <> '3' AND T0.U_JSNGBN = @JOBGBN))
    AND 	T0.U_JSNMON  BETWEEN @STRMON AND @ENDMON
/*
    AND (('2' = @JOBGBN AND (ISNULL(CONVERT(CHAR(10),TermDate,20),'')<>'' AND 
                         ISNULL(CONVERT(CHAR(10),TermDate,20),'') < @JSNYER + '-' + @ENDMON + '-31')
          ) 
    OR  
          ('1' = @JOBGBN AND (ISNULL(CONVERT(CHAR(10),TermDate,20),'')='' OR 
                         ISNULL(CONVERT(CHAR(10),TermDate,20),'') >= @JSNYER + '-' + @ENDMON + '-31')
          ) 
    OR   ('3' = @JOBGBN))
/*  
    AND   T2.Status LIKE CASE @JOBGBN   WHEN '1' THEN '1' 
                                        WHEN '2' THEN '4'
                                        ELSE '%' END
*/*/
    
    GROUP 	BY T0.U_JSNYER, T0.U_CLTCOD,T0.U_MSTCOD
    ORDER 	BY CLTCOD, MSTDPT,T0.U_MSTCOD

-- <3.�����ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    SELECT * FROM [#RPY506] ORDER BY JSNYER, MSTBRK, MSTDPT, MSTSTP, MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

go
--Exec RPY506 '2013', '01', '12', '3', '%', '%', '%'