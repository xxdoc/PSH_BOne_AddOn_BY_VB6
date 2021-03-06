/*==========================================================================
	프로시저명		:	[PH_PY307_01]
	프로시저설명	:	학자금신청내역(분기별_개인별)
	작성자			:	Song Myounggyu
	작업일자			:	2012.12.10
	수정자			:	
	최종수정일자	:	
	작업지시자		:	
	작업지시일자	:	
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY307_01]
(
	@CLTCOD	AS VARCHAR(1), --사업장
	@StdYear	AS VARCHAR(4), --년도
	@Quarter	AS VARCHAR(5), --분기
	@Count		AS VARCHAR(5) --회차
)
AS
SET NOCOUNT ON

----/////테스트용변수선언부/////
--DECLARE @CLTCOD	AS VARCHAR(1) --사업장
--DECLARE @StdYear	AS VARCHAR(4) --년도
--DECLARE @Quarter	AS VARCHAR(5) --분기
--DECLARE @Count		AS VARCHAR(5) --회차

--SET @CLTCOD	= '1'
--SET @StdYear	= '2012'
--SET @Quarter	= '01'
--SET @Count		= '01'
----/////테스트용변수선언부/////



DECLARE @TEMP01 AS TABLE
(
	CntcCode	VARCHAR(20), --사원코드
	CntcName	NVARCHAR(50), --사원성명
	TeamCode	VARCHAR(20), --팀코드
	TeamName	NVARCHAR(50), --팀명
	RspCode		VARCHAR(20), --담당코드
	RspName	NVARCHAR(50), --담당명
	startDat		VARCHAR(8), --입사일자
	Name			NVARCHAR(50), --성명
	GovID			VARCHAR(15), --주민등록번호
	BirDat			VARCHAR(10), --생년월일
	StdYear		VARCHAR(4), --년도
	SchCls		VARCHAR(5), --학교코드
	SchName		NVARCHAR(10), --학교
	Grade			VARCHAR(5), --학년코드
	GradeName	NVARCHAR(10), --학년
	EntFee		NUMERIC(19,6), --입학금
	Tuition		NUMERIC(19,6), --등록금
	Quater		VARCHAR(5), --분기
	[Count]		VARCHAR(5) --회차
)

INSERT		@TEMP01
SELECT		--//////////신청인_S//////////
				T0.U_CntcCode AS [CntcCode], --사원코드
				T0.U_CntcName AS [CntcName], --사원성명
				T2.U_TeamCode AS [TeamCode], --부서코드
				ISNULL(T3.U_CodeNm,'') AS [TeamName], --부서명
				T2.U_RspCode AS [RspCode], --담당코드
				ISNULL(T4.U_CodeNm,'') AS [RspName], --담당명
				CONVERT(VARCHAR(10), T2.U_startDat, 112) AS [startDat], --입사일자
				--//////////신청인_E//////////
				--//////////자녀_S//////////
				T1.U_Name AS [Name], --성명
				REPLACE(T1.U_GovID, '-', '') AS [GovID], --주민등록번호
				CONVERT(VARCHAR(10), T6.U_BirDat, 112) AS [BirDat], --생년월일
				T0.U_StdYear AS [StdYear], --년도
				T1.U_SchCls AS [SchCls], --학교(Code)
				T5.U_CodeNm AS [SchClsName], --학교(Name)
				T1.U_Grade AS [Grade], --학년
				CASE
					WHEN T1.U_Grade = '01' THEN '1학년'
					WHEN T1.U_Grade = '02' THEN '2학년'
					WHEN T1.U_Grade = '03' THEN '3학년'
					WHEN T1.U_Grade = '04' THEN '4학년'
				END AS [GreadName], --학년
				T1.U_EntFee AS [EntFee], --입학금
				T1.U_Tuition AS [Tuition], --등록금
				T0.U_Quarter AS [Quarter], --분기
				CASE
					WHEN T1.U_Count = '01' THEN '1차'
					WHEN T1.U_Count = '02' THEN '2차'
				END AS [Count] --회차
				--//////////자녀_E//////////
FROM			[@PH_PY301A] AS T0 --학자금신청등록H
				INNER JOIN
				[@PH_PY301B] AS T1 --학자금신청등록L
					ON T0.DocEntry = T1.DocEntry
				LEFT JOIN
				[@PH_PY001A] AS T2 --사원마스터
					ON T0.U_CntcCode = T2.Code
				LEFT JOIN
				[@PS_HR200L] AS T3 --인사코드(팀)
					ON T2.U_TeamCode = T3.U_Code
					AND T3.Code = '1'
				LEFT JOIN
				[@PS_HR200L] AS T4 --인사코드(담당)
					ON T2.U_RspCode = T4.U_Code
					AND T4.Code = '2'
				LEFT JOIN
				[@PS_HR200L] AS T5 --인사코드
					ON T1.U_SchCls = T5.U_Code
					AND T5.Code = 'P222'
				LEFT JOIN
				[@PH_PY001D] AS T6 --가족사항 테이블
					ON REPLACE(T1.U_GovID, '-', '') = T6.U_FamPer
					AND T0.U_CntcCode = T6.Code
WHERE		T0.U_CLTCOD = @CLTCOD --사업장
				AND T0.U_StdYear = @StdYear --년도
				AND T0.U_Quarter = @Quarter --분기
				AND T1.U_Count = @Count --회차

DECLARE @TEMP02 AS TABLE
(
	Team			NVARCHAR(50), --신청인부서
	CntcName	NVARCHAR(50), --신청인성명
	Name1		NVARCHAR(50), --자녀성명고등
	BirDat1		VARCHAR(15), --자녀생년월일고등
	Grade1		NVARCHAR(5), --자녀학년고등
	Amt1			NUMERIC(19,6), --자녀학자금고등
	Name2		NVARCHAR(50), --자녀성명전문대
	BirDat2		VARCHAR(15), --자녀생년월일전문대
	Grade2		NVARCHAR(5), --자녀학년전문대
	Amt2			NUMERIC(19,6), --자녀학자금전문대
	Count2		NVARCHAR(5), --회차
	Name3		NVARCHAR(50), --자녀성명대학
	BirDat3		VARCHAR(15), --자녀생년월일대학
	Grade3		NVARCHAR(5), --자녀학년대학
	Amt3			NUMERIC(19,6), --자녀학자금대학
	Count3		NVARCHAR(5), --회차
	Total			NUMERIC(19,6) --계
)

INSERT		@TEMP02
SELECT		T0.TeamName + CASE WHEN T0.RspName = '' THEN '' ELSE '-' END + T0.RspName AS [Team], --[신청인]부서
				T0.CntcName AS [CntcName], --[신청인]성명
				MAX
				(
					CASE
						WHEN T0.SchCls = '01' THEN T0.Name
						ELSE ''
					END
				) AS [Name1], --[자녀]성명(고등)
				MAX
				(
					CASE
						WHEN T0.SchCls = '01' THEN T0.BirDat
						ELSE ''
					END
				) AS [BirDat1], --[자녀]생년월일(고등)
				MAX
				(
					CASE
						WHEN T0.SchCls = '01' THEN T0.GradeName
						ELSE ''
					END 
				) AS [Grade1], --[자녀]학년(고등)
				SUM
				(
					CASE
						WHEN T0.SchCls = '01' THEN T0.EntFee + T0.Tuition
						ELSE 0
					END
				) AS [Amt1], --[자녀]학자금(고등)
				MAX
				(
					CASE
						WHEN T0.SchCls = '02' THEN T0.Name
						ELSE ''
					END 
				) AS [Name2], --[자녀]성명(전문대)
				MAX
				(
					CASE
						WHEN T0.SchCls = '02' THEN T0.BirDat
						ELSE ''
					END
				) AS [BirDat2], --[자녀]생년월일(전문대)
				MAX
				(
					CASE
						WHEN T0.SchCls = '02' THEN T0.GradeName
						ELSE ''
					END
				) AS [Grade2], --[자녀]학년(전문대)
				SUM
				(
					CASE
						WHEN T0.SchCls = '02' THEN T0.EntFee + T0.Tuition
						ELSE 0
					END
				) AS [Amt2], --[자녀]학자금(전문대)
				MAX
				(
					CASE
						WHEN T0.SchCls = '02' THEN T0.[Count]
						ELSE ''
					END
				) AS [Count2], --[자녀]회차
				MAX
				(
					CASE
						WHEN T0.SchCls = '03' THEN T0.Name
						ELSE ''
					END 
				) AS [Name3], --[자녀]성명(대학)
				MAX
				(
					CASE
						WHEN T0.SchCls = '03' THEN T0.BirDat
						ELSE ''
					END
				) AS [BirDat3], --[자녀]생년월일(대학)
				MAX
				(
					CASE
						WHEN T0.SchCls = '03' THEN T0.GradeName
						ELSE ''
					END
				) AS [Grade3], --[자녀]학년(대학)
				SUM
				(
					CASE
						WHEN T0.SchCls = '03' THEN T0.EntFee + T0.Tuition
						ELSE 0
					END
				) AS [Amt3], --[자녀]학자금(대학)
				MAX
				(
					CASE
						WHEN T0.SchCls = '03' THEN T0.[Count]
						ELSE ''
					END
				) AS [Count3], --[자녀]회차
				SUM(T0.EntFee + T0.Tuition) AS [Total] --[자녀]계
FROM			@TEMP01 AS T0
GROUP BY	T0.TeamName + CASE WHEN T0.RspName = '' THEN '' ELSE '-' END + T0.RspName,
				T0.CntcName

SELECT		T0.Team, --신청인부서
				T0.CntcName, --신청인성명
				T0.Name1, --자녀성명고등
				T0.BirDat1, --자녀생년월일고등
				T0.Grade1, --자녀학년고등
				T0.Amt1, --자녀학자금고등
				T0.Name2, --자녀성명전문대
				T0.BirDat2, --자녀생년월일전문대
				T0.Grade2, --자녀학년전문대
				T0.Amt2, --자녀학자금전문대
				T0.Count2, --회차
				T0.Name3, --자녀성명대학
				T0.BirDat3, --자녀생년월일대학
				T0.Grade3, --자녀학년대학
				T0.Amt3, --자녀학자금대학
				T0.Count3, --회차
				T0.Total --계
FROM			@TEMP02 AS T0
				


SET NOCOUNT OFF