/*==========================================================================
	���ν�����		:	[PH_PY305_01]
	���ν�������	:	���ڱݽ�û�� ��ȸ
	�ۼ���			:	Song Myounggyu
	�۾�����			:	2012.11.22
	������			:	
	������������	:	
	�۾�������		:	
	�۾���������	:	
	�۾�����			:	
	�۾�����			:	
	�⺻�۲�			:	�������, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY305_01]
(
	@CLTCOD	AS VARCHAR(1), --�����
	@SCode		AS VARCHAR(20), --�����ȣ
	@StdYear	AS VARCHAR(4), --�⵵
	@Quarter	AS VARCHAR(5) --�б�
)
AS
SET NOCOUNT ON

----/////�׽�Ʈ�뺯�������/////
--DECLARE @CLTCOD	AS VARCHAR(1) --�����
--DECLARE @SCode		AS VARCHAR(20) --�����ȣ
--DECLARE @StdYear	AS VARCHAR(4) --�⵵
--DECLARE @Quarter	AS VARCHAR(5) --�б�

--SET @CLTCOD		= '1'
--SET @SCode		= '11880102'
--SET @StdYear		= '2012'
--SET @Quarter		= '01'
----/////�׽�Ʈ�뺯�������/////

SELECT		--//////////��û��_S//////////
				T0.U_CntcCode AS [CntcCode], --����ڵ�
				T0.U_CntcName AS [CntcName], --�������
				T2.U_TeamCode AS [TeamCode], --�μ��ڵ�
				T3.U_CodeNm AS [TeamName], --�μ���
				CONVERT(VARCHAR(10), T2.U_startDat, 112) AS [startDat], --�Ի�����
				--//////////��û��_E//////////
				--//////////�ڳ�_S//////////
				T1.U_Name AS [Name], --����
				dbo.FUNC_Split(T1.U_GovID, '-', 2) AS [BirthDat], --�������
				T1.U_Sex AS [Sex], --����
				T1.U_SchName AS [SchName], --�б���
				T1.U_Grade AS [Grade], --�г�
				T1.U_EntFee AS [EntFee], --���б�
				T1.U_Tuition AS [Tuition] --��ϱ�
				--//////////�ڳ�_E//////////
FROM			[@PH_PY301A] AS T0 --���ڱݽ�û���H
				INNER JOIN
				[@PH_PY301B] AS T1 --���ڱݽ�û���L
					ON T0.DocEntry = T1.DocEntry
				LEFT JOIN
				[@PH_PY001A] AS T2 --���������
					ON T0.U_CntcCode = T2.Code
				LEFT JOIN
				[@PS_HR200L] AS T3 --�λ��ڵ�
					ON T2.U_TeamCode = T3.U_Code
					AND T3.Code = '1'
WHERE		T0.U_CLTCOD = @CLTCOD --�����
				AND T0.U_CntcCode = @SCode --�����ȣ
				AND T0.U_StdYear = @StdYear --�⵵
				AND T0.U_Quarter = @Quarter --�б�



SET NOCOUNT OFF