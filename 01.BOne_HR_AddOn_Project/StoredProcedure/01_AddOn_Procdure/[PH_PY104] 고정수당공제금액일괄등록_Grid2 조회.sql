/*==========================================================================================
        ���ν�����      : PH_PY104_Grid2
        ���ν�������    : ������������ݾ��ϰ����_��ȸ
        ������          : 
        �۾�����        : 2012-11-12
        �۾�������      : 
        �۾���������    : 
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104_Grid2' AND xtype = 'P'))
	DROP PROCEDURE PH_PY104_Grid2
GO

CREATE PROC PH_PY104_Grid2 (
	@CLTCOD		AS NVARCHAR(10),	-- �����
    @TeamCode	AS NVARCHAR(10),	-- �μ�
    @RspCode	AS NVARCHAR(10),	-- ���
    @PAYTYP		AS NVARCHAR(10),	-- �ݿ�����
    @JIGCODF	AS NVARCHAR(10),	-- ����From
    @JIGCODT	AS NVARCHAR(10),	-- ����To
    @HOBONGF	AS NVARCHAR(10),	-- ȣ��From
    @HOBONGT	AS NVARCHAR(10)		-- ȣ��To
)

AS

--SELECT @SDate, @EDate
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

	SELECT	--'LineNum'	= CONVERT(INT,ROW_NUMBER() OVER (PARTITION BY T0.DOCENTRY ORDER BY T0.DOCENTRY) -1 ),
			'Code'		= Code,							--���
			'FullName'	= U_FullName					--����			
	INTO #TEMP
	FROM [@PH_PY001A]
	WHERE U_CLTCOD = @CLTCOD
	AND (@TeamCode = '%' OR (@TeamCode <> '%' AND U_TeamCode = @TeamCode) )
	AND (@RspCode = '%' OR (@RspCode <> '%' AND U_RspCode = @RspCode))
	AND (@PAYTYP = '%' OR (@PAYTYP <> '%' AND U_PAYTYP = @PAYTYP))
	AND (U_JIGCOD BETWEEN @JIGCODF and @JIGCODT)
	AND (U_HOBONG BETWEEN @HOBONGF and @HOBONGT)
	AND U_Status <> '5'
	
IF NOT (EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104_TEMP2' AND xtype = 'U'))
begin	
	create table PH_PY104_TEMP2(
		CODE NVARCHAR(10), 
		NAME NVARCHAR(10)
	)
end

BEGIN
	INSERT INTO PH_PY104_TEMP2 
	SELECT Code, FullName
	FROM #TEMP
END

--Exec PH_PY104_Grid2 '1','1200','',''

--go
/*
 select Code, U_FullName from [@PH_PY001A] where U_CLTCOD = '1' AND U_PAYTYP='4'
 Exec PH_PY104_Grid2 '1','%','%','2','0000000000','ZZZZZZZZZZ','0000000000','ZZZZZZZZZZ'
 select * from PH_PY104_TEMP2
 delete PH_PY104_TEMP2
*/
