/*==========================================================================
	프로시저명		:	[PH_PY302_02]
	프로시저설명	:	학자금지급완료처리 - 지급완료여부 UPDATE
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
CREATE PROC [dbo].[PH_PY302_02]
(
	@CLTCOD	AS VARCHAR(1), --사업장
	@StdYear	AS VARCHAR(4), --년도
	@Quarter	AS VARCHAR(5), --분기
	@Count		AS VARCHAR(5), --회차
	@PayYN		AS VARCHAR(1) --지급완료여부
)
AS
SET NOCOUNT ON

UPDATE	[@PH_PY301B]
SET		T1.U_PayYN = @PayYN
FROM		[@PH_PY301A] AS T0
			INNER JOIN
			[@PH_PY301B] AS T1
				ON T0.DocEntry = T1.DocEntry
WHERE	T0.[Status] = 'O'
			AND T0.U_CLTCOD = @CLTCOD
			AND T0.U_StdYear = @StdYear
			AND T0.U_Quarter = @Quarter
			AND T1.U_Count = @Count

SET NOCOUNT OFF