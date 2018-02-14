/*==========================================================================
	프로시저명		:	[PH_PY306_01]
	프로시저설명	:	학자금신청내역(개인별) 조회
	작성자			:	Song Myounggyu
	작업일자			:	2012.11.28
	수정자			:	
	최종수정일자	:	
	작업지시자		:	
	작업지시일자	:	
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY306_01]
(
	@CLTCOD	AS VARCHAR(1), --사업장
	@SCode	AS VARCHAR(20) --사원번호
)
AS
SET NOCOUNT ON

--DECLARE @CLTCOD	AS VARCHAR(1) --사업장
--DECLARE @SCode	AS VARCHAR(20) --사원번호

--SET @CLTCOD		= '1'
--SET @SCode	= '11880102'

--최초데이터 임시저장용 테이블 변수
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
SELECT		--//////////신청인_S//////////
				T0.U_CntcCode AS [CntcCode], --사원코드
				T0.U_CntcName AS [CntcName], --사원성명
				T2.U_TeamCode AS [TeamCode], --부서코드
				T3.U_CodeNm AS [TeamName], --부서명
				CONVERT(VARCHAR(10), T2.U_startDat, 112) AS [startDat], --입사일자
				--//////////신청인_E//////////
				--//////////자녀_S//////////
				T1.U_Name AS [Name], --성명
				T0.U_StdYear AS [StdYear], --년도
				T1.U_SchCls AS [SchCls], --학교(Code)
				T4.U_CodeNm AS [SchClsName], --학교(Name)
				T1.U_Grade AS [Grade], --학년
				CASE
					WHEN T1.U_Grade = '01' THEN '1학년'
					WHEN T1.U_Grade = '02' THEN '2학년'
					WHEN T1.U_Grade = '03' THEN '3학년'
					WHEN T1.U_Grade = '04' THEN '4학년'
				END AS [GreadName], --학년
				T1.U_EntFee AS [EntFee], --입학금
				T1.U_Tuition AS [Tuition], --등록금
				T0.U_Quarter AS [Quarter] --분기
				--//////////자녀_E//////////
FROM			[@PH_PY301A] AS T0 --학자금신청등록H
				INNER JOIN
				[@PH_PY301B] AS T1 --학자금신청등록L
					ON T0.DocEntry = T1.DocEntry
				LEFT JOIN
				[@PH_PY001A] AS T2 --사원마스터
					ON T0.U_CntcCode = T2.Code
				LEFT JOIN
				[@PS_HR200L] AS T3 --인사코드
					ON T2.U_TeamCode = T3.U_Code
					AND T3.Code = '1'
				LEFT JOIN
				[@PS_HR200L] AS T4 --인사코드
					ON T1.U_SchCls = T4.U_Code
					AND T4.Code = 'P222'
WHERE		T0.U_CLTCOD = @CLTCOD --사업장
				AND T0.U_CntcCode = @SCode --사원번호


--입학금제외인지 확인 요망(2012.11.230 송명규)
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

--성명
--년도
--학교
--학년
--1/4분기
--2/4분기
--3/4분기
--4/4/분기
--계


SET NOCOUNT OFF