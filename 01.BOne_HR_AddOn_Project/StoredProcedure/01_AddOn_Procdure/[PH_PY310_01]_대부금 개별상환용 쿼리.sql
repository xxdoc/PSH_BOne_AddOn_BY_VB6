
/*
	PH_PY310 대부금문서조회
*/





DECLARE @CLTCOD		VARCHAR(20)
DECLARE @CntcCode		VARCHAR(20)

SET @CLTCOD		= $[@PH_PY310A.U_CLTCOD.0]
SET @CntcCode	= $[@PH_PY310A.U_CntcCode.0]


SELECT		T0.DocEntry AS [문서번호],
				T1.U_CLTName AS [사업장],
				T0.U_CntcCode AS [사원번호],
				T0.U_CntcName AS [사원성명],
				T0.U_LoanDate AS [대출일자],
				T0.U_LoanAmt AS [대출금액],
				T0.U_RpmtPrd AS [상환기간]
FROM			[@PH_PY309A] AS T0
				LEFT JOIN
				[@PH_PY005A] AS T1
					ON T0.U_CLTCOD = T1.U_CLTCode
WHERE		T0.U_CLTCOD = @CLTCOD
				AND T0.U_CntcCode = CASE WHEN @CntcCode = '' THEN T0.U_CntcCode ELSE @CntcCode END


--대부금 상환내역 테이블
CREATE TABLE Z_PH_PY310
(
	CLTCOD		VARCHAR(5), --사업장
	CntcCode	VARCHAR(20), --사번
	LoanDoc		INT, --대부금문서번호
	RpmtDate	DATETIME, --상환일자
	RpmtAmt	NUMERIC(19,6), --상환금액
	RpmtInt		NUMERIC(19,6), --상환이자
	RmainAmt	NUMERIC(19,6) --상환잔액
)
GO




--테스트용 데이터 입력
INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20120705', 400000, 13150, 3600000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20120806', 400000, -10520, 3200000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20120905', 400000, 10520, 2800000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20121005', 400000, 9200, 2400000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20121105', 400000, 8150, 2000000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20121205', 400000, 6570, 1600000)
GO

INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
('2', '8680022', '2', '20130107', 400000, 5780, 1200000)
GO












/*==========================================================================
	프로시저명		:	PH_PY310_01
	프로시저설명	:	대부금상환내역 조회
	만든이			:	송명규
	작업일자			:	2013.01.18
	최종수정일		:	
	작업지시자		:	송명규
	작업지시일자	:	2013.01.18
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은 고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY310_01]
(
	@LoanDoc INT --대부금등록 문서번호
)
AS
SET NOCOUNT ON

----/////테스트용변수선언부/////
--DECLARE @LoanDoc INT

--SET @LoanDoc = 2
----/////테스트용변수선언부/////


--상환내역 조회
SELECT		CONVERT(VARCHAR(10), RpmtDate, 112) AS [RpmtDate],
				RpmtAmt AS [RpmtAmt],
				RpmtInt AS [RpmtInt],
				RmainAmt AS [RmainAmt],
				'N' AS [AddYN]
FROM			Z_PH_PY310
WHERE		LoanDoc = @LoanDoc
ORDER BY	RpmtDate

SET NOCOUNT OFF








/*==========================================================================
	프로시저명		:	PH_PY310_02
	프로시저설명	:	대출일자, 대부금액, 총상환금액, 상환잔액 조회
	만든이			:	송명규
	작업일자			:	2013.01.19
	최종수정일		:	
	작업지시자		:	송명규
	작업지시일자	:	2013.01.19
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은 고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY310_02]
(
	@LoanDoc INT --대부금등록 문서번호
)
AS
SET NOCOUNT ON

----/////테스트용변수선언부/////
--DECLARE @LoanDoc INT
--SET @LoanDoc = 2
----/////테스트용변수선언부/////

CREATE TABLE #Z_PH_PY310
(
	LoanDoc		INT,
	TRpmtAmt	NUMERIC(19,6)	
)

INSERT		#Z_PH_PY310
SELECT		LoanDoc AS [LoanDoc],
				SUM(RpmtAmt) AS [RpmtAmt]
FROM			Z_PH_PY310
WHERE		LoanDoc = @LoanDoc
GROUP BY	LoanDoc


SELECT		CONVERT(VARCHAR(10), T0.U_LoanDate, 112) AS [LoanDate], --대출일자
				T0.U_LoanAmt AS [LoanAmt], --대출금액
				T1.TRpmtAmt AS [TRpmtAmt], --총상환금액
				T0.U_LoanAmt - T1.TRpmtAmt AS [RmainAmt] --상환잔액
FROM			[@PH_PY309A] AS T0
				LEFT JOIN
				#Z_PH_PY310 AS T1
					ON T0.DocEntry = T1.LoanDoc
WHERE		T0.DocEntry = @LoanDoc

DROP TABLE #Z_PH_PY310

SET NOCOUNT OFF




/*==========================================================================
	프로시저명		:	PH_PY310_03
	프로시저설명	:	대부금개별상환 내역을 Z_PH_PY310에 INSERT, @PH_PY309B에 UPDATE
	만든이			:	송명규
	작업일자			:	2013.01.21
	최종수정일		:	
	작업지시자		:	송명규
	작업지시일자	:	2013.01.21
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은 고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY310_03]
(
	@CLTCOD		VARCHAR(5), --사업장
	@CntcCode		VARCHAR(20), --사번
	@LoanDoc		INT, --대부금문서번호
	@RpmtDate		DATETIME, --상환일자
	@RpmtAmt		NUMERIC(19,6), --상환금액
	@RpmtInt		NUMERIC(19,6), --상환이자
	@RmainAmt	NUMERIC(19,6) --상환잔액
)
AS
SET NOCOUNT ON

--상환월 지정
DECLARE @PayMonth AS VARCHAR(7)
SELECT @PayMonth = CONVERT(VARCHAR(7), @RpmtDate, 120)

--Z_PH_PY310에 INSERT
INSERT INTO Z_PH_PY310
(CLTCOD, CntcCode, LoanDoc, RpmtDate, RpmtAmt, RpmtInt, RmainAmt)
VALUES
(@CLTCOD, @CntcCode, @LoanDoc, @RpmtDate, @RpmtAmt, @RpmtInt, @RmainAmt)


--@PH_PY309B에 UPDATE
UPDATE	[@PH_PY309B]
SET		U_TotRpmt = @RpmtAmt,
			U_RpmtYN = 'Y'
WHERE	DocEntry = @LoanDoc
			AND U_PayDate = @PayMonth

SET NOCOUNT OFF












--//////////////////////////입력 데이터 삭제용//////////////////////////
BEGIN TRAN
UPDATE	ONNM
SET		AutoKey = 1
WHERE	ObjectCode = 'PH_PY310'

DELETE 
FROM		[@PH_PY310A]

DELETE
FROM		[Z_PH_PY310]
WHERE	CONVERT(VARCHAR(7), RpmtDate, 120) = '2013-02'
--//////////////////////////입력 데이터 삭제용//////////////////////////