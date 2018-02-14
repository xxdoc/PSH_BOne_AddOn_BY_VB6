/*==========================================================================================
        ���ν�����      : PH_PY004
        ���ν�������    : �ٹ��������ȭ��_��ȸ
        ������          : 
        �۾�����        : 2012-11-05
        �۾�������      : 
        �۾���������    : 
        �۾�����        : 
        �۾�����        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY004' AND xtype = 'P'))
	DROP PROCEDURE PH_PY004
GO

CREATE PROC PH_PY004 (
	@CLTCOD		AS NVARCHAR(10),	-- �����
    @TeamCode	AS NVARCHAR(10),	-- �μ�
    @RspCode	AS NVARCHAR(10),	-- ���
    @ShiftDat	AS NVARCHAR(10)		-- �ٹ�����
)

AS

--SELECT @SDate, @EDate
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

	SELECT	--'LineNum'	= CONVERT(INT,ROW_NUMBER() OVER (PARTITION BY T0.DOCENTRY ORDER BY T0.DOCENTRY) -1 ),
			'TeamCode'	= T1.U_CodeNm,					--�μ�
			'RspCode'	= T2.U_CodeNm,					--���
			'Code'		= T0.Code,						--���
			'FullName'	= T0.U_FullName,				--����
			'Position'	= T3.name,						--��å
			'GNMUJO'	= T0.U_GNMUJO					--�ٹ���
	FROM [@PH_PY001A] T0 LEFT OUTER JOIN [@PS_HR200L] T1 ON T0.U_TeamCode = T1.U_Code AND T1.CODE = '1'
						 LEFT OUTER JOIN [@PS_HR200L] T2 ON T0.U_RspCode = T2.U_Code AND T2.Code = '2'
						 LEFT OUTER JOIN [OHPS]		  T3 ON T0.U_position = T3.posID
						 LEFT OUTER JOIN [@PS_HR200L] T4 ON T0.U_GNMUJO = T4.U_Code AND T4.Code = 'P155'
	WHERE U_CLTCOD = @CLTCOD
	AND (@TeamCode = '' OR U_TeamCode = @TeamCode) 
	AND (@RspCode = '' OR U_RspCode = @RspCode)
	AND (@ShiftDat = '' OR U_ShiftDat = @ShiftDat)
