IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY501' AND xtype = 'P'))
	DROP PROCEDURE RPY501
GO

CREATE  PROC RPY501
    (
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
        ���ν�����      : RPY501
        ���ν�������    : �����ڷ���Ȳ(����ǥ)
        ������          : �ֵ���
        �۾�����        : 2009-12-28
        �۾�������      : �Թ̰�
        �۾���������    : 2009-12-28
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/
    --DROP PROC RPY501
   

    SET NOCOUNT ON
-- <1.�ӽ����̺� ���� >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    

    CREATE TABLE #RPY501 (
            CLTCOD     nvarchar(8),
            CLTNAM     nvarchar(50),
            MSTCOD     nvarchar(8),
            MSTNAM     nvarchar(50),
            GWAPAY     Numeric(19,6),
            BIGWA1     Numeric(19,6),
            BIGWA2     Numeric(19,6),
            GWASEE     Numeric(19,6),
            GWABNS     Numeric(19,6),
            JIGTOT     Numeric(19,6),
            KUKAMT     Numeric(19,6),
            MEDAMT     Numeric(19,6),
            GBHAMT     Numeric(19,6),
            GABGUN     Numeric(19,6),
            JUMINN     Numeric(19,6),
            TOTGON     Numeric(19,6) ) 
                        
-- <2.���� �ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    INSERT INTO [#RPY501]
    SELECT  CLTCOD  =   T1.U_CLTCOD,
            CLTNAM  =   T3.Name,
            MSTCOD  =   T1.U_MstCode,
            MSTNAM  =   T1.U_MstName,
            GWAPAY  =   SUM(T0.U_GWAPAY),                       -- ����ǥ��
            BIGWA1  =   SUM(ISNULL(T0.U_BIGWA01,0) + ISNULL(T0.U_BIGWA03,0) + ISNULL(T0.U_BIGWU03,0)  
                          + ISNULL(T0.U_BIGWA04,0) + ISNULL(T0.U_BIGWA05,0) + ISNULL(T0.U_BIGWA06,0)  
                          + ISNULL(T0.U_BIGG01, 0) + ISNULL(T0.U_BIGH01, 0) + ISNULL(T0.U_BIGH05, 0) + ISNULL(T0.U_BIGH06, 0)  
                          + ISNULL(T0.U_BIGH07, 0) + ISNULL(T0.U_BIGH08, 0) + ISNULL(T0.U_BIGH09, 0) + ISNULL(T0.U_BIGH10, 0)  
                          + ISNULL(T0.U_BIGH11, 0) + ISNULL(T0.U_BIGH12, 0) + ISNULL(T0.U_BIGH13, 0) + ISNULL(T0.U_BIGI01, 0)  
                          + ISNULL(T0.U_BIGK01, 0) + ISNULL(T0.U_BIGM01, 0) + ISNULL(T0.U_BIGM02, 0) + ISNULL(T0.U_BIGM03, 0)  
                          + ISNULL(T0.U_BIGO01, 0) + ISNULL(T0.U_BIGQ01, 0) + ISNULL(T0.U_BIGS01, 0) + ISNULL(T0.U_BIGT01, 0)  
                          + ISNULL(T0.U_BIGX01, 0) + ISNULL(T0.U_BIGY01, 0) + ISNULL(T0.U_BIGY02, 0) + ISNULL(T0.U_BIGY03, 0)  
                          + ISNULL(T0.U_BIGY20, 0) + ISNULL(T0.U_BIGY21, 0) + ISNULL(T0.U_BIGZ01, 0)),
            BIGWA2  =   SUM(ISNULL(T0.U_BIGWA02,0) + ISNULL(T0.U_BIGWA07,0)),
            GWASEE  =   0,  -- �޿��Ѿ�
            GWABNS  =   SUM(T0.U_GWABNS + T0.U_INJBNS),             -- ���Ѿ�
            JIGTOT  =   SUM(T0.U_JIGTOTAL),                         -- �Ѱ�
            KUKAMT  =   SUM(T0.U_KUKAMT),                           -- ���ο���
            MEDAMT  =   SUM(T0.U_MEDAMT + ISNULL(T0.U_NGYAMT,0)),   -- �ǰ�����
            GBHAMT  =   SUM(T0.U_GBHAMT),                           -- ��뺸��
            GABGUN  =   SUM(T0.U_GABGUN),                           -- ���ټ�
            JUMINN  =   SUM(T0.U_JUMIN),                            -- �ֹμ�
            TOTGON  =   SUM(T0.U_KUKAMT + T0.U_MEDAMT + T0.U_GBHAMT + T0.U_GABGUN + T0.U_JUMIN + ISNULL(T0.U_NGYAMT,0))
    FROM    [@ZPY343L] T0   
            INNER JOIN [@ZPY343H] T1 ON T0.DocEntry  = T1.DocEntry
            INNER JOIN [@PH_PY001A] T2 ON T1.U_MstCode = T2.Code
            INNER JOIN [@PH_PY005A] T3 ON T1.U_CLTCOD = T3.U_CLTCode
    WHERE   T1.U_JsnYear    = @JSNYER
    AND     T0.U_LineNum   >= @STRMON
    AND     T0.U_LineNum   <= @ENDMON                        
    AND     T1.U_CLTCOD  LIKE @CLTCOD                       
    AND     T1.U_DptCode LIKE @MSTDPT                        
    AND     T1.U_MstCode LIKE @MSTCOD
    AND     T2.U_Status    LIKE CASE @JOBGBN WHEN '1' THEN '1' 
                                           WHEN '2' THEN '4'
                                           ELSE '%' END
    GROUP   BY T1.U_CLTCOD, T3.NAME, T1.U_MstCode,  T1.U_MstName                                                 
    ORDER   BY T1.U_CLTCOD, T1.U_MstName,  T1.U_MstCode
    
    UPDATE  [#RPY501]
    SET     GWASEE  =   ISNULL(GWAPAY,0) 
                      + ISNULL(BIGWA1,0) + ISNULL(BIGWA2,0) 

-- <3.���� �ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    SELECT * FROM [#RPY501] ORDER BY CLTCOD, MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

--go
-- Exec RPY501  '2013', '01', '12', '3', N'%',  N'%', N'%'