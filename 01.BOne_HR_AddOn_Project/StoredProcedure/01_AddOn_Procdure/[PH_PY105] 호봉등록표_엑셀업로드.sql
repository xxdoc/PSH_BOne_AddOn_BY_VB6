/*==========================================================================================
        ���ν�����      : PH_PY105
        ���ν�������    : ȣ�����ǥ_�������ε�
        ������          : 
        �۾�����        : 2012-11-14
        �۾�������      : 
        �۾���������    : 
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY105' AND xtype = 'P'))
	DROP PROCEDURE PH_PY105
GO

CREATE PROC PH_PY105 (
    @JIGCOD		AS NVARCHAR(10),	-- �����ڵ�
    @HOBCOD		AS NVARCHAR(10),	-- ȣ���ڵ�
    @HOBNAM		AS NVARCHAR(20),	-- ȣ���̸�
    @STDAMT		AS NVARCHAR(30),	-- �޿��⺻
    @BNSAMT		AS NVARCHAR(30),	-- �󿩱⺻
    @EXTAMT01	AS NVARCHAR(20),	-- ������01
    @EXTAMT02	AS NVARCHAR(20),	-- ������02
    @EXTAMT03	AS NVARCHAR(20),	-- ������03
    @EXTAMT04	AS NVARCHAR(20),	-- ������04
    @EXTAMT05	AS NVARCHAR(20),	-- ������05
    @EXTAMT06	AS NVARCHAR(20),	-- ������06
    @EXTAMT07	AS NVARCHAR(20),	-- ������07
    @EXTAMT08	AS NVARCHAR(20),	-- ������08
    @EXTAMT09	AS NVARCHAR(20),	-- ������09
    @EXTAMT10	AS NVARCHAR(20)		-- ������10
)

AS

--SELECT @SDate, @EDate
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

BEGIN
	INSERT INTO PH_PY105_TEMP 
	(JIGCOD, HOBCOD	, HOBNAM, STDAMT, BNSAMT, EXTAMT01, EXTAMT02, EXTAMT03, EXTAMT04, EXTAMT05, EXTAMT06,
	 EXTAMT07, EXTAMT08, EXTAMT09, EXTAMT10)
	VALUES (@JIGCOD, @HOBCOD, @HOBNAM, @STDAMT, @BNSAMT, @EXTAMT01, @EXTAMT02, @EXTAMT03, @EXTAMT04,
			@EXTAMT05, @EXTAMT06, @EXTAMT07, @EXTAMT08, @EXTAMT09, @EXTAMT10)

END

--go
/*
 EXEC PH_PY105 '','','',''
 select * from PH_PY105_TEMP
 delete PH_PY104_TEMP2
*/
