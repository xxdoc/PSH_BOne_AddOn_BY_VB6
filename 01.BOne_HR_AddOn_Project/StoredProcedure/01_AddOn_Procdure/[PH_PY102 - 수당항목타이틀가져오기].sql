/*==========================================================================================
		프로시저명		: PH_PY102
		프로시저설명	: 수당항목 타이틀구함
		만든이			: 
		작업일자		: 2012-11-20
		작업지시자		: 
		작업지시일자	: 
		작업목적		: 고정급,변동급,급여에 항목 타이틀조회용
		작업내용		: 
	===========================================================================================*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY102' AND xtype = 'P'))
	DROP PROCEDURE PH_PY102
GO

CREATE            PROC PH_PY102
	(
		@CLTCOD		 AS Nvarchar(1),		--사업장
		@YM 		 AS Nvarchar(6),		--작업연월
		@STDTYP 	 AS Nvarchar(1),		--제수당여부
		@HOBUSE 	 AS Nvarchar(1),		--호봉참조여부
		@FIXGBN 	 AS Nvarchar(1), 		--고정변동구분
		@CSUCOD		 AS Nvarchar(10)
	)
--WITH ENCRYPTION
 AS
	SET NOCOUNT ON
--< 1. 수당항목 적용월구함 > ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ		
	DECLARE	@U_JOBYMM		 AS Nvarchar(6)
	
	SELECT TOP 1 @U_JOBYMM = T0.U_YM
	FROM [@PH_PY102A] T0
	WHERE T0.U_YM <= @YM 
	AND   T0.U_CLTCOD = @CLTCOD
	ORDER BY T0.U_YM DESC
	IF ISNULL(@U_JOBYMM, '') = ''
	BEGIN
		SET @U_JOBYMM = @YM
	END 	

--< 2. 수당항목 타이틀 구함 > ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ			
	SELECT 	ISNULL(T0.U_CSUCOD,'') AS U_CSUCOD, 
			ISNULL(T0.U_CSUNAM,'')  AS U_CSUNAM, 
			ISNULL(T1.Code,'') AS Code, 
			ISNULL(T0.U_MONPAY,'') AS U_MONPAY,
			ISNULL(T0.U_KUMAMT, 0) AS U_KUMAMT, 
			ISNULL(T0.U_CSUGBN, 30) AS U_CSUGBN,
			ISNULL(T0.U_GWATYP, '') AS U_GWATYP,
			ISNULL(T0.U_GBHGBN, '') AS U_GBHGBN,
			ISNULL(T0.U_ROUNDT,'R') AS U_ROUNDT,
			ISNULL(T0.U_LENGTH,'1') AS U_LENGTH,
			ISNULL(T0.U_BNSUSE, 'N') AS U_BNSUSE,
			ISNULL(T0.U_INSLIN,'') AS U_INSLIN,
			ISNULL(T0.U_LINSEQ,'') AS U_LINSEQ,
			ISNULL(T0.U_BTXCOD, '') AS U_BTXCOD

	INTO #PH_PY102
	--select *
	FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code 
	WHERE T1.U_CLTCOD = @CLTCOD
	AND  T1.U_YM = @U_JOBYMM	
	AND  (@HOBUSE = '' OR (@HOBUSE <> '' AND T0.U_HOBUSE = @HOBUSE))	--호봉참조여부
	AND	 (@FIXGBN = '' OR (@FIXGBN <> '' AND T0.U_FIXGBN = @FIXGBN))		--고정변동구분(Y고정,N변동)
	AND	 (@FIXGBN = '' OR (@FIXGBN <> '' AND LEFT(T0.U_CSUCOD,1) <> 'A'))	--기본급,상여금제외
	--AND  T0.U_CSUCOD LIKE @CSUCOD
	AND	 ISNULL(T0.U_LINSEQ,'') <> ''
	ORDER BY T0.U_LINSEQ
	IF @FIXGBN = 'Y' --고정급일경우 고정인사순서에 따라 출력되도록
	BEGIN
		SELECT * FROM [#PH_PY102] ORDER BY U_INSLIN
	END
	ELSE
	BEGIN
		SELECT * FROM [#PH_PY102] ORDER BY U_LINSEQ
	END		


--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

--GO
--Exec PH_PY102  '1','201212', '', '', '', ''

--
