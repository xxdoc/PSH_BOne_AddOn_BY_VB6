/*==========================================================================================
        ���ν�����      : PH_PY104
        ���ν�������    : ������������ݾ��ϰ����_��ȸ
        ������          : 
        �۾�����        : 2012-11-12
        �۾�������      : 
        �۾���������    : 
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY104' AND xtype = 'P'))
	DROP PROCEDURE PH_PY104
GO

CREATE PROC PH_PY104 (
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
			'Code'		= Code,						--���
			'FullName'	= U_FullName					--����			
	FROM [@PH_PY001A]
	WHERE U_CLTCOD = @CLTCOD
	AND (@TeamCode = '' OR U_TeamCode = @TeamCode) 
	AND (@RspCode = '' OR U_RspCode = @RspCode)
	AND (@PAYTYP = '' OR U_PAYTYP = @PAYTYP)
	AND (U_JIGCOD BETWEEN @JIGCODF and @JIGCODT)
	AND (U_HOBONG BETWEEN @HOBONGF and @HOBONGT)
	