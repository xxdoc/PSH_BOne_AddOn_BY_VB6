/*==========================================================================
	���ν�����		:	[PH_PY306_01]
	���ν�������	:	���ڱݽ�û����(���κ�) ��ȸ
	�ۼ���			:	Song Myounggyu
	�۾�����			:	2012.11.28
	������			:	
	������������	:	
	�۾�������		:	
	�۾���������	:	
	�۾�����			:	
	�۾�����			:	
	�⺻�۲�			:	�������, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY306_01]
(
	@CLTCOD	AS VARCHAR(1), --�����
	@SCode	AS VARCHAR(20) --�����ȣ
)
AS
SET NOCOUNT ON

--DECLARE @CLTCOD	AS VARCHAR(1) --�����
--DECLARE @SCode	AS VARCHAR(20) --�����ȣ

--SET @CLTCOD		= '1'
--SET @SCode	= '11880102'

--���ʵ����� �ӽ������ ���̺� ����
DECLARE @TEMP AS TABLE
(
	CntcCode	VARCHAR(20),
	CntcName	NVARCHAR(50),
	TeamCode	VARCHAR(20),
	TeamName	NVARCHAR(50),
	StartDate	VARCHAR(10),
	Name			NVARCHAR(50),
	StdYear		VARCHAR(4),
	SchCls		VARCHAR(5),
	SchClsName	NVARCHAR(20),
	Grade			VARCHAR(5),
	GradeName	NVARCHAR(20),
	EntFee		NUMERIC(19,6),
	Tuition		NUMERIC(19,6),
	[Quarter]		VARCHAR(5)
)


INSERT		@TEMP
SELECT		--//////////��û��_S//////////
				T0.U_CntcCode AS [CntcCode], --����ڵ�
				T0.U_CntcName AS [CntcName], --�������
				T2.U_TeamCode AS [TeamCode], --�μ��ڵ�
				T3.U_CodeNm AS [TeamName], --�μ���
				CONVERT(VARCHAR(10), T2.U_startDat, 112) AS [startDat], --�Ի�����
				--//////////��û��_E//////////
				--//////////�ڳ�_S//////////
				T1.U_Name AS [Name], --����
				T0.U_StdYear AS [StdYear], --�⵵
				T1.U_SchCls AS [SchCls], --�б�(Code)
				T4.U_CodeNm AS [SchClsName], --�б�(Name)
				T1.U_Grade AS [Grade], --�г�
				CASE
					WHEN T1.U_Grade = '01' THEN '1�г�'
					WHEN T1.U_Grade = '02' THEN '2�г�'
					WHEN T1.U_Grade = '03' THEN '3�г�'
					WHEN T1.U_Grade = '04' THEN '4�г�'
				END AS [GreadName], --�г�
				T1.U_EntFee AS [EntFee], --���б�
				T1.U_Tuition AS [Tuition], --��ϱ�
				T0.U_Quarter AS [Quarter] --�б�
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
				LEFT JOIN
				[@PS_HR200L] AS T4 --�λ��ڵ�
					ON T1.U_SchCls = T4.U_Code
					AND T4.Code = 'P222'
WHERE		T0.U_CLTCOD = @CLTCOD --�����
				AND T0.U_CntcCode = @SCode --�����ȣ


--���б��������� Ȯ�� ���(2012.11.230 �۸��)
SELECT		TeamName,
				CntcCode,
				CntcName,
				StartDate,
				Name,
				StdYear,
				SchCls,
				SchClsName,
				Grade,
				SUM
				(
					CASE
						WHEN [Quarter] = '01' THEN EntFee + Tuition
						ELSE 0 
					END
				) AS [Quarter1],
				SUM
				(
					CASE
						WHEN [Quarter] = '02' THEN EntFee + Tuition
						ELSE 0 
					END
				) AS [Quarter2],
				SUM
				(
					CASE
						WHEN [Quarter] = '03' THEN EntFee + Tuition
						ELSE 0 
					END
				) AS [Quarter3],
				SUM
				(
					CASE
						WHEN [Quarter] = '04' THEN EntFee + Tuition
						ELSE 0 
					END
				) AS [Quarter4],
				SUM(EntFee + Tuition) AS [Total]
FROM			@TEMP
GROUP BY	TeamName,
				CntcCode,
				CntcName,
				StartDate,
				Name,
				StdYear,
				SchCls,
				SchClsName,
				Grade
ORDER BY	StdYear,
				SchCls,
				Grade

--����
--�⵵
--�б�
--�г�
--1/4�б�
--2/4�б�
--3/4�б�
--4/4/�б�
--��


SET NOCOUNT OFF