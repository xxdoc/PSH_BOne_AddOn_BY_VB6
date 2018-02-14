/*==========================================================================================
        ���ν�����      : PH_PY105_TEMP_CHK
        ���ν�������    : ȣ�����ǥ_�������ε�_�ӽ����̺� ����
        ������          : 
        �۾�����        : 2012-11-14
        �۾�������      : 
        �۾���������    : 
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY105_TEMP_CHK' AND xtype = 'P'))
	DROP PROCEDURE PH_PY105_TEMP_CHK
GO

CREATE PROC PH_PY105_TEMP_CHK
AS

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

IF NOT (EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY105_TEMP' AND xtype = 'U'))
	begin	
		create table PH_PY105_TEMP(
			JIGCOD		 NVARCHAR(10),	-- �����ڵ�
			HOBCOD		 NVARCHAR(10),	-- ȣ���ڵ�
			HOBNAM		 NVARCHAR(20),	-- ȣ���̸�
			STDAMT		 NVARCHAR(30),	-- �޿��⺻
			BNSAMT		 NVARCHAR(30),	-- �󿩱⺻
			EXTAMT01	 NVARCHAR(20),	-- ������01
			EXTAMT02	 NVARCHAR(20),	-- ������02
			EXTAMT03	 NVARCHAR(20),	-- ������03
			EXTAMT04	 NVARCHAR(20),	-- ������04
			EXTAMT05	 NVARCHAR(20),	-- ������05
			EXTAMT06	 NVARCHAR(20),	-- ������06
			EXTAMT07	 NVARCHAR(20),	-- ������07
			EXTAMT08	 NVARCHAR(20),	-- ������08
			EXTAMT09	 NVARCHAR(20),	-- ������09
			EXTAMT10	 NVARCHAR(20)	-- ������10
		)
		
	end
else
	begin
		DELETE PH_PY105_TEMP
	end
