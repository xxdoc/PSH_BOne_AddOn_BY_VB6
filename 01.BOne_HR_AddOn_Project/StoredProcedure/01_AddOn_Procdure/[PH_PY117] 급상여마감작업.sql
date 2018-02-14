/*==========================================================================================
		프로시저명		: PH_PY117
		프로시저설명	: 급상여마감작업
		만든이			: 
		작업일자		: 2013-01-07
		작업지시자		: 
		작업지시일자	: 
		작업목적		: 급상여마감작업
		작업내용		: 
	===========================================================================================*/
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY117' AND xtype = 'P'))
	DROP PROCEDURE PH_PY117
GO

CREATE            PROC PH_PY117
	(
		@CLTCOD		 AS Nvarchar(1),		--사업장
		@YM 		 AS Nvarchar(6),		--작업연월
		@JOBTYP 	 AS Nvarchar(1),		--지급종류
		@JOBGBN 	 AS Nvarchar(1),		--지급구분
		@PAYSEL 	 AS Nvarchar(1), 		--지급대상
		@MSTCOD		 AS Nvarchar(10),		--사원
		@TEAMCODE	 AS Nvarchar(10),		--부서
		@RSPCODE	 AS Nvarchar(10),		--담당
		@ENDCHK		 AS Nvarchar(1)			--마감
	)
--WITH ENCRYPTION
 AS
	SET NOCOUNT ON
	
	SELECT '' as MAGAM, T1.U_CodeNM , T2.U_CodeNM , T0.U_MSTCOD, T0.U_MSTNAM,
	T0.U_TOTPAY, T0.U_TOTGON, T0.U_SILJIG,
	T0.U_CSUD01, T0.U_CSUD02, T0.U_CSUD03, T0.U_CSUD04, T0.U_CSUD05, T0.U_CSUD06, T0.U_CSUD07, T0.U_CSUD08, T0.U_CSUD09, T0.U_CSUD10, 
	T0.U_CSUD11, T0.U_CSUD12, T0.U_CSUD13, T0.U_CSUD14, T0.U_CSUD15, T0.U_CSUD16, T0.U_CSUD17, T0.U_CSUD18, T0.U_CSUD19, T0.U_CSUD20, 
	T0.U_CSUD21, T0.U_CSUD22, T0.U_CSUD23, T0.U_CSUD24, T0.U_CSUD25, T0.U_CSUD26, T0.U_CSUD27, T0.U_CSUD28, T0.U_CSUD29, T0.U_CSUD30, 
	T0.U_CSUD31, T0.U_CSUD22, T0.U_CSUD33, T0.U_CSUD34, T0.U_CSUD35, T0.U_CSUD36, 
	T0.U_GONG01, T0.U_GONG02, T0.U_GONG03, T0.U_GONG04, T0.U_GONG05, T0.U_GONG06, T0.U_GONG07, T0.U_GONG08, T0.U_GONG09, T0.U_GONG10, 
	T0.U_GONG11, T0.U_GONG12, T0.U_GONG13, T0.U_GONG14, T0.U_GONG15, T0.U_GONG16, T0.U_GONG17, T0.U_GONG18, T0.U_GONG19, T0.U_GONG20, 
	T0.U_GONG21, T0.U_GONG22, T0.U_GONG23, T0.U_GONG24, T0.U_GONG25, T0.U_GONG26, T0.U_GONG27, T0.U_GONG28, T0.U_GONG29, T0.U_GONG30, 
	T0.U_GONG31, T0.U_GONG32, T0.U_GONG33, T0.U_GONG34, T0.U_GONG35, T0.U_GONG36
	FROM [@PH_PY112A] T0 
	INNER JOIN [@PS_HR200L] T1 ON T0.U_TeamCode = T1.U_Code AND T1.Code = '1' 
	INNER JOIN [@PS_HR200L] T2 ON T0.U_RspCode = T2.U_Code AND T2.Code = '2'
	WHERE T0.U_YM = @YM
	/* 해당 연월이 아닌경우
	(T0.U_YM = @YM OR (T0.U_YM <> @YM AND T0.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY112A] WHERE U_YM < @YM															
															AND   U_CLTCOD = @CLTCOD
															AND   U_JOBTYP = @JOBTYP
															AND   U_JOBGBN = @JOBGBN
															AND   (U_JOBTRG = @PAYSEL OR (U_JOBTRG <> @PAYSEL AND U_JOBTRG LIKE @PAYSEL)))))
	*/										
	AND   T0.U_CLTCOD = @CLTCOD
	AND   T0.U_JOBTYP = @JOBTYP
	AND   T0.U_JOBGBN = @JOBGBN
	AND   (T0.U_JOBTRG = @PAYSEL OR (T0.U_JOBTRG <> @PAYSEL AND T0.U_JOBTRG LIKE @PAYSEL))
	AND   (T0.U_MSTCOD = @MSTCOD OR (T0.U_MSTCOD <> @MSTCOD AND T0.U_MSTCOD LIKE @MSTCOD))
	AND   (T0.U_TeamCode = @TEAMCODE OR (T0.U_TeamCode <> @TEAMCODE AND T0.U_TeamCode LIKE @TEAMCODE))
	AND   (T0.U_RspCode = @RSPCODE OR (T0.U_RspCode <> @RSPCODE AND T0.U_RspCode LIKE @RSPCODE))
	AND   T0.U_ENDCHK = @ENDCHK
	ORDER BY T0.U_TeamCode, T0.U_RspCode, T0.U_MSTCOD DESC
	
--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

--GO
-- Exec PH_PY117 '1','201211','1','1','%','%','%','%','N'

