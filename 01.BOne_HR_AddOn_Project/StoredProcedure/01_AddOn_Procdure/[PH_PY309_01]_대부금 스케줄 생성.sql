/*==========================================================================
	���ν�����		:	PH_PY309_01
	���ν�������	:	��αݵ�� �� ��ȯ������ ����
	������			:	�۸��
	�۾�����			:	2013.01.14
	����������		:	
	�۾�������		:	�۸��
	�۾���������	:	2013.01.14
	�۾�����			:	
	�۾�����			:	
	�⺻�۲�			:	���� ���, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY309_01]
(
	@prmLoanAmt		NUMERIC(19,6), --����ݾ�
	@prmLoanDate	DATETIME, --��������
	@prmRpmtPrd		INT --��ȯ�Ⱓ
)
AS
SET NOCOUNT ON

----/////�׽�Ʈ�뺯�������/////
--DECLARE @prmLoanAmt		NUMERIC(19,6)
--DECLARE @prmLoanDate	DATETIME
--DECLARE @prmRpmtPrd		INT

--SET @prmLoanAmt	= 2000000
--SET @prmLoanDate	= '20130305'
--SET @prmRpmtPrd	= 5
----/////�׽�Ʈ�뺯�������/////

CREATE TABLE #TEMP01
(
	Cnt			INT, --ȸ��
	PayDate		VARCHAR(10), --�޿����޳��
	RpmtAmt	INT, --����ȯ��
	TotRpmt		INT --������ȯ
)

DECLARE @MinPrd AS INT
DECLARE @MaxPrd AS INT

SELECT	@MinPrd = 1,
			@MaxPrd = @prmRpmtPrd

DECLARE @Cnt			INT
DECLARE @PayDate	VARCHAR(10)
DECLARE @RpmtAmt	INT
DECLARE @TotRpmt	INT

WHILE @MinPrd <= @MaxPrd
	BEGIN
	
		SELECT	@Cnt = @MinPrd,
					@PayDate = CONVERT(VARCHAR(7), DATEADD(Month, @MinPrd - 1, @prmLoanDate), 120),
					@RpmtAmt = @prmLoanAmt / @prmRpmtPrd,
					@TotRpmt = 0
	
		INSERT INTO #TEMP01
		(
			Cnt,
			PayDate,
			RpmtAmt,
			TotRpmt
		)
		VALUES
		(
			@Cnt,
			@PayDate,
			@RpmtAmt,
			@TotRpmt
		)
	
		SET @MinPrd = @MinPrd + 1
	END

SELECT	Cnt,
			PayDate,
			RpmtAmt,
			TotRpmt
FROM		#TEMP01

DROP TABLE #TEMP01

SET NOCOUNT OFF
