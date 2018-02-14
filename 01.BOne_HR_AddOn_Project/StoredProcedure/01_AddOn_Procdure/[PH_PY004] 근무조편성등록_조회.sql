/*==========================================================================================
        프로시저명      : PH_PY004
        프로시저설명    : 근무조편성등록화면_조회
        만든이          : 
        작업일자        : 2012-11-05
        작업지시자      : 
        작업지시일자    : 
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/

IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY004' AND xtype = 'P'))
	DROP PROCEDURE PH_PY004
GO

CREATE PROC PH_PY004 (
	@CLTCOD		AS NVARCHAR(10),	-- 사업장
    @TeamCode	AS NVARCHAR(10),	-- 부서
    @RspCode	AS NVARCHAR(10),	-- 담당
    @ShiftDat	AS NVARCHAR(10)		-- 근무형태
)

AS

--SELECT @SDate, @EDate
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

	SELECT	--'LineNum'	= CONVERT(INT,ROW_NUMBER() OVER (PARTITION BY T0.DOCENTRY ORDER BY T0.DOCENTRY) -1 ),
			'TeamCode'	= T1.U_CodeNm,					--부서
			'RspCode'	= T2.U_CodeNm,					--담당
			'Code'		= T0.Code,						--사번
			'FullName'	= T0.U_FullName,				--성명
			'Position'	= T3.name,						--직책
			'GNMUJO'	= T0.U_GNMUJO					--근무조
	FROM [@PH_PY001A] T0 LEFT OUTER JOIN [@PS_HR200L] T1 ON T0.U_TeamCode = T1.U_Code AND T1.CODE = '1'
						 LEFT OUTER JOIN [@PS_HR200L] T2 ON T0.U_RspCode = T2.U_Code AND T2.Code = '2'
						 LEFT OUTER JOIN [OHPS]		  T3 ON T0.U_position = T3.posID
						 LEFT OUTER JOIN [@PS_HR200L] T4 ON T0.U_GNMUJO = T4.U_Code AND T4.Code = 'P155'
	WHERE U_CLTCOD = @CLTCOD
	AND (@TeamCode = '' OR U_TeamCode = @TeamCode) 
	AND (@RspCode = '' OR U_RspCode = @RspCode)
	AND (@ShiftDat = '' OR U_ShiftDat = @ShiftDat)
