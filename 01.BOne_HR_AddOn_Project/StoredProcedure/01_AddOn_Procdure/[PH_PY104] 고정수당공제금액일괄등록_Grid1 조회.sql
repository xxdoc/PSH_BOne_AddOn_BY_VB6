/*==========================================================================================
        ���ν�����      : PH_PY104_Grid1
        ���ν�������    : ������������ݾ��ϰ����_��ȸ
        ������          : 
        �۾�����        : 2012-11-12
        �۾�������      : 
        �۾���������    : 
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104_Grid1' AND xtype = 'P'))
	DROP PROCEDURE PH_PY104_Grid1
GO

CREATE PROC PH_PY104_Grid1 (
	@GBN		AS NVARCHAR(10),	-- �����
    @CSUCOD		AS NVARCHAR(10),	-- �μ�
    @CSUNAM		AS NVARCHAR(10),	-- ���
    @SEQ		AS NVARCHAR(10)		-- �ݿ�����
)

AS
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

IF not (EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104_TEMP' AND xtype = 'U'))
begin	
	create table PH_PY104_TEMP(
		GUBUN NVARCHAR(10), 
		CSUCOD NVARCHAR(10), 
		CSUNAM NVARCHAR(10), 
		SEQ NVARCHAR(10)
	)
	
end

	
INSERT into PH_PY104_TEMP(GUBUN,CSUCOD,CSUNAM,SEQ) values (@GBN,@CSUCOD,@CSUNAM,@SEQ)

--go
/*
 EXEC PH_PY104_Grid1 '����','E11','test3','2'
 select * from PH_PY104_TEMP
 drop table PH_PY104_TEMP
 delete PH_PY104_TEMP
*/
