/*==========================================================================
	���ν�����		:	[PH_PY301_01]
	���ν�������	:	���ڱ� ����Ƚ�� ��ȸ
	�ۼ���			:	Song Myounggyu
	�۾�����			:	2012.11.21
	������			:	
	������������	:	
	�۾�������		:	
	�۾���������	:	
	�۾�����			:	
	�۾�����			:	
	�⺻�۲�			:	�������, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY301_01]
(
	@GovID		AS VARCHAR(20), --�ֹε�Ϲ�ȣ
	@SchCls		AS VARCHAR(5), --�б�
	@DocEntry	AS INT --������ȣ
)
AS
SET NOCOUNT ON

----/////�׽�Ʈ�뺯�������/////
--DECLARE @GovID		AS VARCHAR(20)
--DECLARE @SchCls		AS VARCHAR(5)

--SET @GovID	= '930531-1823716'
--SET @SchCls	= '01'
----/////�׽�Ʈ�뺯�������/////

IF @SchCls = '03'
	BEGIN
		SET @SchCls = '02' --�������а� ���б��� �ڵ带 �����ϰ� ó�� �� �������п��� ���б��� �����ϴ� ��츦 ó���ϱ� ����
	END

DECLARE @TEMP_TABLE AS TABLE
(
	DocEntry	INT,
	Name		NVARCHAR(50),
	GovID		VARCHAR(20),
	SchCls	VARCHAR(5)
)

INSERT		@TEMP_TABLE
SELECT		DocEntry AS [DocEntry],
				U_Name AS [Name],
				U_GovID AS [GovID],
				CASE
					WHEN U_SchCls = '02' THEN '02'
					WHEN U_SchCls = '03' THEN '02'
					ELSE U_SchCls
				END SchCls --�������а� ���б��� �ڵ带 �����ϰ� ó�� �� �������п��� ���б��� �����ϴ� ��츦 ó���ϱ� ����
FROM			[@PH_PY301B] AS T0
WHERE		T0.U_GovID = @GovID


SELECT	COUNT(*) AS [PayCount]
FROM		@TEMP_TABLE
WHERE	SchCls = @SchCls
			AND DocEntry <> @DocEntry --���� ������ȣ ����



SET NOCOUNT OFF