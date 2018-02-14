IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY502' AND xtype = 'P'))
	DROP PROCEDURE RPY502
GO

CREATE  PROC RPY502 (
        @JSNYER AS Nvarchar(4),     --�۾�����
        @CLTCOD AS Nvarchar(8),     --�ڻ��ڵ�
        @MSTDPT AS Nvarchar(8),     --�μ�
        @MSTCOD AS Nvarchar(8)      --�����ȣ      
    )
      
 AS
    /*==========================================================================================
        ���ν�����      : RPY502
        ���ν�������    : �����ٹ�����Ȳ
        ������          : �Թ̰�
        �۾�����        : 2007-01-30
        �۾�������      : �Թ̰�
        �۾���������    : 2007-01-30
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/
    --DROP PROC RPY502
    --Exec RPY502 '2013', N'%', N'%', N'%'

    SET NOCOUNT ON

-- <1.�ӽ����̺� ���� >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    

    CREATE TABLE #RPY502 (
            MSTCOD     nvarchar(8),
            MSTNAM     nvarchar(50),
            JONNAM     nvarchar(40),
            JONNBR     nvarchar(12),
            JONPAY     Numeric(19,6),
            JONBNS     Numeric(19,6),
            INJBNS     Numeric(19,6),
			JONJUS     Numeric(19,6),
			URIBNS     Numeric(19,6),
            JONGAB     Numeric(19,6),
            JONJUM     Numeric(19,6),
            JONMED     Numeric(19,6),
            JONGBH     Numeric(19,6),
            JONKUK     Numeric(19,6),
            JONBT1     Numeric(19,6)
            ) 

-- <2.�����ٹ��� �ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    INSERT INTO [#RPY502]
    SELECT  MSTCOD  =   T1.U_MSTCOD,
            MSTNAM  =   T1.U_MSTNAM,
            JONNAM  =   T0.U_JONNAM,
            JONNBR  =   T0.U_JONNBR,
            JONPAY  =   ISNULL(T0.U_JONPAY,0),
            JONBNS  =   ISNULL(T0.U_JONBNS,0),
            INJBNS  =   ISNULL(T0.U_INJBNS,0),
            JONJUS  =   ISNULL(T0.U_JONJUS,0),
            URIBNS  =   ISNULL(T0.U_URIBNS,0),
            JONGAB  =   ISNULL(T0.U_JONGAB,0),
            JONJUM  =   ISNULL(T0.U_JONJUM,0),
            JONMED  =   ISNULL(T0.U_JONMED,0),
            JONGBH  =   ISNULL(T0.U_JONGBH,0),
            JONKUK  =   ISNULL(T0.U_JONKUK,0),
            JONBT1  =   ISNULL(T0.U_JONBT1,0) + ISNULL(T0.U_JONBT2,0) + ISNULL(T0.U_JONBT3,0)
                      + ISNULL(T0.U_JONBU3,0) + ISNULL(T0.U_JONBT4,0) + ISNULL(T0.U_JONBT5,0) + ISNULL(T0.U_JONBT6,0)
                      + ISNULL(T0.U_JBTG01,0) + ISNULL(T0.U_JBTH01,0) + ISNULL(T0.U_JBTH05,0) + ISNULL(T0.U_JBTH06,0) 
                      + ISNULL(T0.U_JBTH07,0) + ISNULL(T0.U_JBTH08,0) + ISNULL(T0.U_JBTH09,0) + ISNULL(T0.U_JBTH10,0) 
                      + ISNULL(T0.U_JBTH11,0) + ISNULL(T0.U_JBTH12,0) + ISNULL(T0.U_JBTH13,0) + ISNULL(T0.U_JBTI01,0) 
                      + ISNULL(T0.U_JBTK01,0) + ISNULL(T0.U_JBTM01,0) + ISNULL(T0.U_JBTM02,0) + ISNULL(T0.U_JBTM03,0) 
                      + ISNULL(T0.U_JBTO01,0) + ISNULL(T0.U_JBTQ01,0) + ISNULL(T0.U_JBTS01,0) + ISNULL(T0.U_JBTT01,0) 
                      + ISNULL(T0.U_JBTX01,0) + ISNULL(T0.U_JBTY01,0) + ISNULL(T0.U_JBTY02,0) + ISNULL(T0.U_JBTY03,0) 
                      + ISNULL(T0.U_JBTY20,0) + ISNULL(T0.U_JBTY21,0) + ISNULL(T0.U_JBTZ01,0)
    FROM    [@ZPY502L] T0 
            INNER JOIN [@ZPY502H] T1 ON T0.DocEntry = T1.DocEntry
            INNER JOIN [@PH_PY001A] T2 ON T1.U_MstCod = T2.Code
            
    WHERE   T1.U_JSNYER = @JSNYER       --�⵵
    AND     T1.U_CLTCOD LIKE @CLTCOD    --�ڻ� 
    AND     T2.U_TeamCode LIKE @MSTDPT    --�μ�
    AND     T1.U_MSTCOD LIKE @MSTCOD    --���
    ORDER   BY T1.U_CLTCOD, T1.U_MSTCOD

-- <3.�����ٹ��� �ڷ� ��ȸ >�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�    
    SELECT * FROM [#RPY502] ORDER BY MSTCOD

--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

