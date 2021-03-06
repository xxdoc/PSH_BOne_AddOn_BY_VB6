/*==========================================================================
	프로시저명		:	[PH_PY302_01]
	프로시저설명	:	학자금지급완료처리를 위한 전체 금액 조회
	작성자			:	Song Myounggyu
	작업일자			:	2012.11.20
	수정자			:	
	최종수정일자	:	
	작업지시자		:	
	작업지시일자	:	
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY302_01]
(
	@CLTCOD	AS VARCHAR(1), --사업장
	@StdYear	AS VARCHAR(4), --년도
	@Quarter	AS VARCHAR(5) --분기
)
AS
SET NOCOUNT ON

----/////테스트용변수선언부/////
--DECLARE @CLTCOD	AS VARCHAR(1) --사업장
--DECLARE @StdYear	AS VARCHAR(4) --년도
--DECLARE @Quarter	AS VARCHAR(5) --분기

--SET @CLTCOD	= '1'
--SET @StdYear	= '2012'
--SET @Quarter	= '01'
----/////테스트용변수선언부/////


SELECT		T1.U_Count AS [Count], --회차
				SUM(T1.U_EntFee) AS [EntFee], --입학금
				SUM(T1.U_Tuition) AS [Tuition], --등록금
				SUM(T1.U_EntFee) + SUM(T1.U_Tuition) AS [Total],
				T1.U_PayYN AS [PayYN]
FROM			[@PH_PY301A] AS T0
				INNER JOIN
				[@PH_PY301B] AS T1
					ON T0.DocEntry = T1.DocEntry
WHERE		T0.[Status] = 'O'
				AND T0.U_CLTCOD = @CLTCOD
				AND T0.U_StdYear = @StdYear
				AND T0.U_Quarter = @Quarter
GROUP BY	T1.U_Count,
				T1.U_PayYN
ORDER BY	T1.U_Count


SET NOCOUNT OFF