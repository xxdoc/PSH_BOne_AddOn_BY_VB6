
/*
	PH_PY310 ��αݹ�����ȸ
*/





DECLARE @CLTCOD		VARCHAR(20)
DECLARE @CntcCode		VARCHAR(20)

SET @CLTCOD		= $[@PH_PY310A.U_CLTCOD.0]
SET @CntcCode	= $[@PH_PY310A.U_CntcCode.0]


SELECT		T0.DocEntry AS [������ȣ],
				T1.U_CLTName AS [�����],
				T0.U_CntcCode AS [�����ȣ],
				T0.U_CntcName AS [�������],
				T0.U_LoanDate AS [��������],
				T0.U_LoanAmt AS [����ݾ�],
				T0.U_RpmtPrd AS [��ȯ�Ⱓ]
FROM			[@PH_PY309A] AS T0
				LEFT JOIN
				[@PH_PY005A] AS T1
					ON T0.U_CLTCOD = T1.U_CLTCode
WHERE		T0.U_CLTCOD = @CLTCOD
				AND T0.U_CntcCode = CASE WHEN @CntcCode = '' THEN T0.U_CntcCode ELSE @CntcCode END


--��α� ��ȯ���� ���̺�
CREATE TABLE Z_PH_PY310
(
	CLTCOD		VARCHAR(5), --�����
	CntcCode	VARCHAR(20), --���
	LoanDoc		INT, --��αݹ�����ȣ
	RpmtDate	DATETIME, --��ȯ����
	RpmtAmt	NUMERIC(19,6), --��ȯ�ݾ�
	RpmtInt		NUMERIC(19,6), --��ȯ����
	RmainAmt	NUMERIC(19,6) --��ȯ�ܾ�
)
GO




--�׽�Ʈ�� ������ �Է�
INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20120705', 400000, 13150, 3600000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20120806', 400000, -10520, 3200000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20120905', 400000, 10520, 2800000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20121005', 400000, 9200, 2400000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20121105', 400000, 8150, 2000000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20121205', 400000, 6570, 1600000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20130107', 400000, 5780, 1200000)
GO












/*==========================================================================
	���ν�����		:	PH_PY310_01
	���ν�������	:	��αݻ�ȯ���� ��ȸ
	������			:	�۸��
	�۾�����			:	2013.01.18
	����������		:	
	�۾�������		:	�۸��
	�۾���������	:	2013.01.18
	�۾�����			:	
	�۾�����			:	
	�⺻�۲�			:	���� ���, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY310_01]
(
	@LoanDoc INT --��αݵ�� ������ȣ
)
AS
SET NOCOUNT ON

----/////�׽�Ʈ�뺯�������/////
--DECLARE @LoanDoc INT

--SET @LoanDoc = 2
----/////�׽�Ʈ�뺯�������/////


--��ȯ���� ��ȸ
SELECT		CONVERT(VARCHAR(10), RpmtDate, 112) AS [RpmtDate],
				RpmtAmt AS [RpmtAmt],
				RpmtInt AS [RpmtInt],
				RmainAmt AS [RmainAmt],
				'N' AS [AddYN]
FROM			Z_PH_PY310
WHERE		LoanDoc = @LoanDoc
ORDER BY	RpmtDate

SET NOCOUNT OFF








/*==========================================================================
	���ν�����		:	PH_PY310_02
	���ν�������	:	��������, ��αݾ�, �ѻ�ȯ�ݾ�, ��ȯ�ܾ� ��ȸ
	������			:	�۸��
	�۾�����			:	2013.01.19
	����������		:	
	�۾�������		:	�۸��
	�۾���������	:	2013.01.19
	�۾�����			:	
	�۾�����			:	
	�⺻�۲�			:	���� ���, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY310_02]
(
	@LoanDoc INT --��αݵ�� ������ȣ
)
AS
SET NOCOUNT ON

----/////�׽�Ʈ�뺯�������/////
--DECLARE @LoanDoc INT
--SET @LoanDoc = 2
----/////�׽�Ʈ�뺯�������/////

CREATE TABLE #Z_PH_PY310
(
	LoanDoc		INT,
	TRpmtAmt	NUMERIC(19,6)	
)

INSERT		#Z_PH_PY310
SELECT		LoanDoc AS [LoanDoc],
				SUM(RpmtAmt) AS [RpmtAmt]
FROM			Z_PH_PY310
WHERE		LoanDoc = @LoanDoc
GROUP BY	LoanDoc


SELECT		CONVERT(VARCHAR(10), T0.U_LoanDate, 112) AS [LoanDate], --��������
				T0.U_LoanAmt AS [LoanAmt], --����ݾ�
				T1.TRpmtAmt AS [TRpmtAmt], --�ѻ�ȯ�ݾ�
				T0.U_LoanAmt - T1.TRpmtAmt AS [RmainAmt] --��ȯ�ܾ�
FROM			[@PH_PY309A] AS T0
				LEFT JOIN
				#Z_PH_PY310 AS T1
					ON T0.DocEntry = T1.LoanDoc
WHERE		T0.DocEntry = @LoanDoc

DROP TABLE #Z_PH_PY310

SET NOCOUNT OFF




/*==========================================================================
	���ν�����		:	PH_PY310_03
	���ν�������	:	��αݰ�����ȯ ������ Z_PH_PY310�� INSERT, @PH_PY309B�� UPDATE
	������			:	�۸��
	�۾�����			:	2013.01.21
	����������		:	
	�۾�������		:	�۸��
	�۾���������	:	2013.01.21
	�۾�����			:	
	�۾�����			:	
	�⺻�۲�			:	���� ���, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY310_03]
(
	@CLTCOD		VARCHAR(5), --�����
	@CntcCode		VARCHAR(20), --���
	@LoanDoc		INT, --��αݹ�����ȣ
	@RpmtDate		DATETIME, --��ȯ����
	@RpmtAmt		NUMERIC(19,6), --��ȯ�ݾ�
	@RpmtInt		NUMERIC(19,6), --��ȯ����
	@RmainAmt	NUMERIC(19,6) --��ȯ�ܾ�
)
AS
SET NOCOUNT ON

--��ȯ�� ����
DECLARE @PayMonth AS VARCHAR(7)
SELECT @PayMonth = CONVERT(VARCHAR(7), @RpmtDate, 120)

--Z_PH_PY310�� INSERT
INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
(@CLTCOD, @CntcCode, @LoanDoc, @RpmtDate, @RpmtAmt, @RpmtInt, @RmainAmt)


--@PH_PY309B�� UPDATE
UPDATE	[@PH_PY309B]
SET		U_TotRpmt = @RpmtAmt,
			U_RpmtYN = 'Y'
WHERE	DocEntry = @LoanDoc
			AND U_PayDate = @PayMonth

SET NOCOUNT OFF












--//////////////////////////�Է� ������ ������//////////////////////////
BEGIN TRAN
UPDATE	ONNM
SET		AutoKey = 1
WHERE	ObjectCode = 'PH_PY310'

DELETE 
FROM		[@PH_PY310A]

DELETE
FROM		[Z_PH_PY310]
WHERE	CONVERT(VARCHAR(7), RpmtDate, 120) = '2013-02'
--//////////////////////////�Է� ������ ������//////////////////////////