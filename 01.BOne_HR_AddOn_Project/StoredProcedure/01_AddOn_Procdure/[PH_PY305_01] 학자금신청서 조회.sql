/*==========================================================================
	프로시저명		:	[PH_PY305_01]
	프로시저설명	:	학자금신청서 조회
	작성자			:	Song Myounggyu
	작업일자			:	2012.11.22
	수정자			:	
	최종수정일자	:	
	작업지시자		:	
	작업지시일자	:	
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY305_01]
(
	@CLTCOD	AS VARCHAR(1), --사업장
	@SCode		AS VARCHAR(20), --사원번호
	@StdYear	AS VARCHAR(4), --년도
	@Quarter	AS VARCHAR(5) --분기
)
AS
SET NOCOUNT ON

----/////테스트용변수선언부/////
--DECLARE @CLTCOD	AS VARCHAR(1) --사업장
--DECLARE @SCode		AS VARCHAR(20) --사원번호
--DECLARE @StdYear	AS VARCHAR(4) --년도
--DECLARE @Quarter	AS VARCHAR(5) --분기

--SET @CLTCOD		= '1'
--SET @SCode		= '11880102'
--SET @StdYear		= '2012'
--SET @Quarter		= '01'
----/////테스트용변수선언부/////

SELECT		--//////////신청인_S//////////
				T0.U_CntcCode AS [CntcCode], --사원코드
				T0.U_CntcName AS [CntcName], --사원성명
				T2.U_TeamCode AS [TeamCode], --부서코드
				T3.U_CodeNm AS [TeamName], --부서명
				CONVERT(VARCHAR(10), T2.U_startDat, 112) AS [startDat], --입사일자
				--//////////신청인_E//////////
				--//////////자녀_S//////////
				T1.U_Name AS [Name], --성명
				dbo.FUNC_Split(T1.U_GovID, '-', 2) AS [BirthDat], --생년월일
				T1.U_Sex AS [Sex], --성별
				T1.U_SchName AS [SchName], --학교명
				T1.U_Grade AS [Grade], --학년
				T1.U_EntFee AS [EntFee], --입학금
				T1.U_Tuition AS [Tuition] --등록금
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
WHERE		T0.U_CLTCOD = @CLTCOD --사업장
				AND T0.U_CntcCode = @SCode --사원번호
				AND T0.U_StdYear = @StdYear --년도
				AND T0.U_Quarter = @Quarter --분기



SET NOCOUNT OFF