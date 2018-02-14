/*==========================================================================
	프로시저명		:	[PH_PY301_01]
	프로시저설명	:	학자금 지급횟수 조회
	작성자			:	Song Myounggyu
	작업일자			:	2012.11.21
	수정자			:	
	최종수정일자	:	
	작업지시자		:	
	작업지시일자	:	
	작업목적			:	
	작업내용			:	
	기본글꼴			:	맑은고딕, 8
==========================================================================*/
ALTER PROC [dbo].[PH_PY301_01]
(
	@GovID		AS VARCHAR(20), --주민등록번호
	@SchCls		AS VARCHAR(5), --학교
	@DocEntry	AS INT --문서번호
)
AS
SET NOCOUNT ON

----/////테스트용변수선언부/////
--DECLARE @GovID		AS VARCHAR(20)
--DECLARE @SchCls		AS VARCHAR(5)

--SET @GovID	= '930531-1823716'
--SET @SchCls	= '01'
----/////테스트용변수선언부/////

IF @SchCls = '03'
	BEGIN
		SET @SchCls = '02' --전문대학과 대학교의 코드를 동일하게 처리 → 전문대학에서 대학교로 편입하는 경우를 처리하기 위함
	END

DECLARE @TEMP_TABLE AS TABLE
(
	DocEntry	INT,
	Name		NVARCHAR(50),
	GovID		VARCHAR(20),
	SchCls	VARCHAR(5)
)

INSERT		@TEMP_TABLE
SELECT		DocEntry AS [DocEntry],
				U_Name AS [Name],
				U_GovID AS [GovID],
				CASE
					WHEN U_SchCls = '02' THEN '02'
					WHEN U_SchCls = '03' THEN '02'
					ELSE U_SchCls
				END SchCls --전문대학과 대학교의 코드를 동일하게 처리 → 전문대학에서 대학교로 편입하는 경우를 처리하기 위함
FROM			[@PH_PY301B] AS T0
WHERE		T0.U_GovID = @GovID


SELECT	COUNT(*) AS [PayCount]
FROM		@TEMP_TABLE
WHERE	SchCls = @SchCls
			AND DocEntry <> @DocEntry --현재 문서번호 제외



SET NOCOUNT OFF