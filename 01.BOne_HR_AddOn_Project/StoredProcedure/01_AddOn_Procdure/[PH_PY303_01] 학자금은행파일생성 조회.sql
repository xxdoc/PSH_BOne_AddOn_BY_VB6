/*==========================================================================
	프로시저명		:	[PH_PY303_01]
	프로시저설명	:	학자금은행파일생성을 위한 조회 쿼리
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
ALTER PROC [dbo].[PH_PY303_01]
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


SELECT		T0.U_CntcCode AS [CntcCode], --사번
				T0.U_CntcName AS [CntcName], --성명
				SUM(T1.U_EntFee + T1.U_Tuition) AS [Amount], --금액
				T2.U_BANK1 AS [BankName], --은행명
				T2.U_ACCTNO1 AS [AcctNo] --계좌번호
FROM			[@PH_PY301A] AS T0
				INNER JOIN
				[@PH_PY301B] AS T1
					ON T0.DocEntry = T1.DocEntry
				LEFT JOIN
				[@PH_PY001A] AS T2
					ON T0.U_CntcCode = T2.Code
WHERE		T0.U_CLTCOD = @CLTCOD
				AND T0.U_StdYear = @StdYear
				AND T0.U_Quarter = @Quarter
				AND T1.U_Count = @Count
GROUP BY	T0.U_CntcCode,
				T0.U_CntcName,
				T2.U_BANK1,
				T2.U_ACCTNO1
ORDER BY	T0.U_CntcCode

SET NOCOUNT OFF