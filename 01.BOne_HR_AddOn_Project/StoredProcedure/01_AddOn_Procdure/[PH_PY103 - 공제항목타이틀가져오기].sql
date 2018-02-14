/*==========================================================================================
		프로시저명		: PH_PY103
		프로시저설명	: 공제항목 타이틀구함
		만든이			: 
		작업일자		: 2012-11-20
		작업지시자		: 
		작업지시일자	: 
		작업목적		: 고정급,변동급,급여에 항목 타이틀조회용
		작업내용		: 
	===========================================================================================*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY103' AND xtype = 'P'))
	DROP PROCEDURE PH_PY103
GO

CREATE             PROC PH_PY103
	(
		@CLTCOD		AS Nvarchar(1),		--사업장
		@YM 		AS Nvarchar(6),		--작업연월
		@FIXGBN 	AS Nvarchar(1), 	--고정변동구분
		@CSUCOD		AS Nvarchar(10)
	)
--WITH ENCRYPTION
 AS
	

	SET NOCOUNT ON


--< 1. 공제항목 적용월구함 > ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ		
	DECLARE	@U_JOBYMM		 AS Nvarchar(6)
	
	SELECT TOP 1 @U_JOBYMM = T0.U_YM
	FROM [@PH_PY103A] T0
	WHERE T0.U_YM <= @YM 
	AND   T0.U_CLTCOD = @CLTCOD
	ORDER BY T0.Code DESC

	IF ISNULL(@U_JOBYMM, '') = ''
	BEGIN
		SET @U_JOBYMM = @YM
	END 	
	
--< 2. 공제항목 타이틀 구함 > ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ			
	SELECT 	ISNULL(T0.U_CSUCOD, '') AS U_CSUCOD, 
			ISNULL(T0.U_CSUNAM,'')  AS U_CSUNAM, 
			ISNULL(T1.Code,'') AS Code, 
			ISNULL(T0.U_SILCUN,'') AS U_SILCUN, 
			ISNULL(T0.U_BNSUSE,'N') AS U_BNSUSE, 
			ISNULL(T0.U_ROUNDT, 'R') AS U_ROUNDT, 
			ISNULL(T0.U_LENGTH, 1) AS U_LENGTH
	FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code -- WHERE T0.Code = N'YES'
	WHERE T1.U_CLTCOD = @CLTCOD
	AND  T1.U_YM = @U_JOBYMM 	
	AND  (@FIXGBN = '' OR (@FIXGBN <> '' AND T0.U_FIXGBN = @FIXGBN))	--고정변동구분(Y고정,N변동)
	--AND  T0.U_CSUCOD LIKE @CSUCOD
	AND	 ISNULL(T0.U_LINSEQ,'') <> ''
	ORDER BY T0.U_LINSEQ


--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

--Exec PH_PY103  '1','201212',  '', ''