
/*==========================================================================
	프로시저명		:	PH_PY313_01
	프로시저설명	:	대부금 계산 조회
	만든이			:	송명규
	작업일자			:	2013.01.18
	최종수정일		:	
	작업지시자		:	송명규
	작업지시일자	:	2013.01.18
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은 고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY313_01]
(
	@CLTCOD		VARCHAR(1),
	@RpmtDate		DATETIME,
	@CntcCode		VARCHAR(20),
	@RegYN			VARCHAR(1)
)
AS
SET NOCOUNT ON

----/////테스트용변수선언부/////
--DECLARE @CLTCOD		VARCHAR(1)
--DECLARE @RpmtDate	DATETIME
--DECLARE @CntcCode		VARCHAR(20)
--DECLARE @RegYN			VARCHAR(1)

--SET @CLTCOD		= '2'
--SET @RpmtDate	= '20120605'
--SET @CntcCode	= ''
--SET @RegYN		= 'Y'
----/////테스트용변수선언부/////


CREATE TABLE #PH_PY310_TEMP
(
	LoanDoc		INT,
	RpmtDate	DATETIME,
	RmainAmt	NUMERIC(19,6),
	LineId			INT
)

INSERT		#PH_PY310_TEMP
SELECT		LoanDoc,
				MAX(RpmtDate) AS [RpmtDate],
				RmainAmt,
				LineId
FROM			[Z_PH_PY310]
GROUP BY	LoanDoc,
				RmainAmt,
				LineId

DECLARE @RpmtInt NUMERIC(10,2)


SELECT		T0.DocEntry AS [LoanDoc], --대부금등록문서번호
				T0.U_CntcCode AS [CntcCode], --사번
				T0.U_CntcName AS [CntcName], --성명
				--T1.U_PayDate AS [PayDate], --대상급여분(월)-데이터 확인용
				--CONVERT(DATETIME, T1.U_PayDate + '-25') AS [PayDate2], --대상급여분(월)-데이터 확인용
				REPLACE(CONVERT(VARCHAR(10), T0.U_LoanDate, 120), '-', '.') AS [LoanDate], --대출일자
				T0.U_LoanAmt AS [LoanAmt], --대출금액
				CASE 
					WHEN T2.RpmtDate IS NULL THEN REPLACE(CONVERT(VARCHAR(10), T0.U_LoanDate, 120), '-','.')
					ELSE REPLACE(CONVERT(VARCHAR(10), T2.RpmtDate, 120), '-','.')
				END AS [PrRpmtDt], --이전상환일자
				CASE
					WHEN T2.RmainAmt IS NULL THEN T0.U_LoanAmt
					ELSE T2.RmainAmt
				END AS [PrRmainAmt], --이전상환잔액
				CASE
					WHEN T2.RpmtDate IS NULL THEN DATEDIFF(d, T0.U_LoanDate, @RpmtDate)
					ELSE DATEDIFF(d, T2.RpmtDate, @RpmtDate)
				END AS [UseDt], --사용일수
				T1.U_RpmtAmt AS [RpmtAmt], --상환금액
				ROUND
				(
					CASE
						WHEN T2.RmainAmt IS NULL THEN T0.U_LoanAmt
						ELSE T2.RmainAmt
					END --이전상환잔액
					*
					CASE
						WHEN T2.RpmtDate IS NULL THEN DATEDIFF(d, T0.U_LoanDate, @RpmtDate)
						ELSE DATEDIFF(d, T2.RpmtDate, @RpmtDate)
					END --사용일수
					/ 365 * (T0.U_IntRate * 0.01), -1, -1
				)
				AS [RpmtInt], --상환이자(이전상환잔액 * 사용일수 / 365 * 이자율) ※절사
				CASE
					WHEN T2.RmainAmt IS NULL THEN T0.U_LoanAmt
					ELSE T2.RmainAmt
				END
				-
				T1.U_RpmtAmt AS [RmainAmt], --상환잔액(이전상환잔액 - 상환금액)
				T1.U_RpmtYN AS [RegYN], --등록여부
				T1.LineId AS [LineId]
FROM			[@PH_PY309A] AS T0
				INNER JOIN
				[@PH_PY309B] AS T1
					ON T0.DocEntry = T1.DocEntry
				LEFT JOIN
				#PH_PY310_TEMP T2
					ON T0.DocEntry = T2.LoanDoc
					AND T1.LineId = T2.LineId
WHERE		T0.U_CLTCOD = @CLTCOD
				AND T1.U_RpmtAmt > 0 --상환금액이 존재하고
				AND T1.U_RpmtYN = @RegYN --등록여부에 따른 조회
				AND CONVERT(DATETIME, T1.U_PayDate + '-25') < @RpmtDate
				AND T0.U_CntcCode = CASE WHEN @CntcCode = '' THEN T0.U_CntcCode ELSE @CntcCode END

DROP TABLE #PH_PY310_TEMP

SET NOCOUNT OFF




/*
	PH_PY313_사원조회
*/


DECLARE @BPLId AS VARCHAR(1)
SET @BPLId = $[$CLTCOD.0.0]

SELECT	T0.Code AS [사원번호],
			T0.U_FullName AS [사원성명]
FROM		[@PH_PY001A] AS T0
WHERE	T0.U_CLTCOD = @BPLId
			AND T0.U_Status <> '5'
















/*==========================================================================
	프로시저명		:	PH_PY313_02
	프로시저설명	:	대부금 계산 데이터 저장
	만든이			:	송명규
	작업일자			:	2013.02.19
	최종수정일		:	
	작업지시자		:	송명규
	작업지시일자	:	2013.02.19
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은 고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY313_02]
(
	@CLTCOD		VARCHAR(1), --사업장
	@CntcCode		VARCHAR(20), --사번
    @LoanDoc		INT, --대부금문서번호
    @RpmtDate		DATETIME, --상환일자
    @RpmtAmt		NUMERIC(19,6), --상환금액
    @RpmtInt		NUMERIC(19,6), --상환이자
    @RmainAmt	NUMERIC(19,6), --상환잔액
    @LineId			INT, --대부금라인번호
    @RegYN			VARCHAR(1) --등록여부
)
AS
SET NOCOUNT ON

----/////테스트용변수선언부/////
--DECLARE @CLTCOD		VARCHAR(1) --사업장
--DECLARE @CntcCode		VARCHAR(20) --사번
--DECLARE @LoanDoc		INT --대부금문서번호
--DECLARE @RpmtDate	DATETIME --상환일자
--DECLARE @RpmtAmt		NUMERIC(19,6) --상환금액
--DECLARE @RpmtInt		NUMERIC(19,6) --상환이자
--DECLARE @RmainAmt	NUMERIC(19,6) --상환잔액
--DECLARE @LineId			INT --대부금라인번호
--DECLARE @RegYN			VARCHAR(1) --등록여부

--SET @CLTCOD		= ''
--SET @CntcCode	= ''	
--SET @LoanDoc		= ''
--SET @RpmtDate	= ''
--SET @RpmtAmt	= ''
--SET @RpmtInt		= ''
--SET @RmainAmt	= ''
--SET @LineId			= ''
--SET @RegYN		= ''
----/////테스트용변수선언부/////


IF @RegYN = 'Y'
	BEGIN
	
		--대부금상환 테이블에 INSERT
		INSERT [Z_PH_PY310]
		(
			CLTCOD,
			CntcCode,
			LoanDoc,
			RpmtDate,
			RpmtAmt,
			RpmtInt,
			RmainAmt,
			LineId
		)
		VALUES
		(
			@CLTCOD,
			@CntcCode,
			@LoanDoc,
			@RpmtDate,
			@RpmtAmt,
			@RpmtInt,
			@RmainAmt,
			@LineId
		)
		
		--대부금스케줄 테이블에 UPDATE
		UPDATE	[@PH_PY309B]
		SET		U_RpmtYN = @RegYN
		WHERE	DocEntry = @LoanDoc
					AND LineId = @LineId
	
	END
ELSE IF @RegYN = 'N'
	BEGIN
	
		--대부금상환 테이블에 DELETE
		DELETE
		FROM		[Z_PH_PY310]
		WHERE	LoanDoc = @LoanDoc
					AND LineId = @LineId
		
		--대부금스케줄 테이블에 UPDATE
		UPDATE	[@PH_PY309B]
		SET		U_RpmtYN = @RegYN
		WHERE	DocEntry = @LoanDoc
					AND LineId = @LineId
	
	END

SET NOCOUNT OFF
