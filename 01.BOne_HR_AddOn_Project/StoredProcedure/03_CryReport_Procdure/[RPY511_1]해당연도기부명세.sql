IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY511_1' AND xtype = 'P'))
	DROP PROCEDURE RPY511_1
GO

CREATE PROC RPY511_1 (
		@JSNYER		AS NVARCHAR(6),		--귀속년도
		@CLTCOD     AS Nvarchar(8),     --자사코드
		@MSTDPT     AS Nvarchar(8),     --부서
	    @MSTCOD 	AS NVARCHAR(8) 	   	--사원번호			
	) 

AS
    /*==========================================================================================
        프로시저명      : RPY511_1
        프로시저설명    : 기부금 명세서_1
        만든이          : 송정호
        작업일자        : 2011-02-24
        작업지시자      : 
        작업지시일자    : 
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/
    -- DROP PROC RPY511_1
    -- Exec RPY511_1 '2009', N'%', N'%', '106001'

    SET NOCOUNT ON

-- <1.임시테이블 생성 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ

	CREATE TABLE #RPY511_1 (
		DocEntry	Int,
		U_GBUYOU	NVARCHAR(40),
		U_GBUCOD	NVARCHAR(2),
		U_GBUNAE	NVARCHAR(10),
		U_GBUNAM	NVARCHAR(40),
		U_GBUNBR	NVARCHAR(14),
		U_GWANGE	NVARCHAR(1),
		U_FAMNAM	NVARCHAR(20),
		U_PERNBR	NVARCHAR(14),
		U_GBUCNT	NUMERIC(19,6),
		U_GBUAMT	NUMERIC(19,6)
	)

-- <2.정산자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ

	INSERT	INTO [#RPY511_1]
	SELECT	DocEntry    =   T0.DocEntry,
			U_GBUYOU	=	CASE WHEN T1.U_GBUCOD = '10' THEN N'법정'
								 WHEN T1.U_GBUCOD = '20' THEN N'정치자금'
								 WHEN T1.U_GBUCOD = '30' THEN N'조특법 73'
								 WHEN T1.U_GBUCOD = '31' THEN N'조특법 73 ① 11'
								 WHEN T1.U_GBUCOD = '40' THEN N'지정'
								 WHEN T1.U_GBUCOD = '41' THEN N'종교'
								 WHEN T1.U_GBUCOD = '42' THEN N'우리사주'
								 WHEN T1.U_GBUCOD = '50' THEN N'공제제외'
								 ELSE '' END,
			U_GBUCOD	=	T1.U_GBUCOD,
			U_GBUNAE	=	N'금전',
			U_GBUNAM	=	MAX(T1.U_GBUNAM),
			U_GBUNBR	=	T1.U_GBUNBR,
			U_GWANGE	=	MAX(T1.U_GWANGE),
			U_FAMNAM	=	MAX(T1.U_FAMNAM),
			U_PERNBR	=	T1.U_PERNBR,
			U_GBUCNT	=	SUM(T1.U_GBUCNT),
			U_GBUAMT	=	SUM(T1.U_GBUAMT)
	FROM	[@ZPY505H] T0 
			INNER JOIN [@ZPY505L] T1 ON T0.DocEntry = T1.DocEntry
			INNER JOIN [@PH_PY001A] T2 ON T0.U_MSTCOD = T2.Code
			LEFT JOIN [@PH_PY005A] T3 ON T0.U_CLTCOD = T3.U_CLTCode			
	WHERE	T0.U_JSNYER = @JSNYER
	AND		T2.U_TeamCode LIKE @MSTDPT 
	AND		T0.U_MSTCOD LIKE @MSTCOD
	GROUP	BY T0.DocEntry, T1.U_GBUCOD, T1.U_GBUNBR, T1.U_PERNBR

-- <3.정산자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
    
    SELECT * FROM [#RPY511_1]

	SET NOCOUNT OFF