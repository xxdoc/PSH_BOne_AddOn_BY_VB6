/*==========================================================================
	프로시저명		:	PH_PY309_01
	프로시저설명	:	대부금등록 시 상환스케줄 생성
	만든이			:	송명규
	작업일자			:	2013.01.14
	최종수정일		:	
	작업지시자		:	송명규
	작업지시일자	:	2013.01.14
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은 고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY309_01]
(
	@prmLoanAmt		NUMERIC(19,6), --대출금액
	@prmLoanDate	DATETIME, --대출일자
	@prmRpmtPrd		INT --상환기간
)
AS
SET NOCOUNT ON

----/////테스트용변수선언부/////
--DECLARE @prmLoanAmt		NUMERIC(19,6)
--DECLARE @prmLoanDate	DATETIME
--DECLARE @prmRpmtPrd		INT

--SET @prmLoanAmt	= 2000000
--SET @prmLoanDate	= '20130305'
--SET @prmRpmtPrd	= 5
----/////테스트용변수선언부/////

CREATE TABLE #TEMP01
(
	Cnt			INT, --회차
	PayDate		VARCHAR(10), --급여지급년월
	RpmtAmt	INT, --월상환액
	TotRpmt		INT --개별상환
)

DECLARE @MinPrd AS INT
DECLARE @MaxPrd AS INT

SELECT	@MinPrd = 1,
			@MaxPrd = @prmRpmtPrd

DECLARE @Cnt			INT
DECLARE @PayDate	VARCHAR(10)
DECLARE @RpmtAmt	INT
DECLARE @TotRpmt	INT

WHILE @MinPrd <= @MaxPrd
	BEGIN
	
		SELECT	@Cnt = @MinPrd,
					@PayDate = CONVERT(VARCHAR(7), DATEADD(Month, @MinPrd - 1, @prmLoanDate), 120),
					@RpmtAmt = @prmLoanAmt / @prmRpmtPrd,
					@TotRpmt = 0
	
		INSERT INTO #TEMP01
		(
			Cnt,
			PayDate,
			RpmtAmt,
			TotRpmt
		)
		VALUES
		(
			@Cnt,
			@PayDate,
			@RpmtAmt,
			@TotRpmt
		)
	
		SET @MinPrd = @MinPrd + 1
	END

SELECT	Cnt,
			PayDate,
			RpmtAmt,
			TotRpmt
FROM		#TEMP01

DROP TABLE #TEMP01

SET NOCOUNT OFF
