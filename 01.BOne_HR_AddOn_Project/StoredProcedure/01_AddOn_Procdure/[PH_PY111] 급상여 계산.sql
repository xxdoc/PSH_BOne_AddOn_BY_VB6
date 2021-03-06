
-- =============================================
-- Procedure ID : PH_PY111
-- Author       : Minho Choi
-- Create date  : 2012.12.05
-- Description  : 급상여자료생성
-- EXEC PH_PY111 15
-- =============================================
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY111' AND xtype = 'P'))
	DROP PROCEDURE PH_PY111
GO

CREATE   PROCEDURE [dbo].[PH_PY111]
         @iDocEntry  AS int
AS

DECLARE	 @CltCod   nvarchar(10),
         @YM       nvarchar(06),
         @JobTyp   nvarchar(08),
         @JobGbn   nvarchar(08),
         @JobTrg   nvarchar(08),
         @Jigbil   datetime,
         @StrDpt   nvarchar(10),
         @EndDpt   nvarchar(10),
         @MstCod   nvarchar(08),
         @UserSign int,
         @RetChk   char(1)

--PRINT 'Display TEST Error Message' RETURN --cmh

--- PH_PY111 테이블에서 자료를 가져와 값 설정 ---
SELECT
 @CltCod   = U_CLTCOD
,@YM       = U_YM
,@JobTyp   = U_JOBTYP
,@JobGbn   = U_JOBGBN
,@JobTrg   = U_JOBTRG
,@Jigbil   = U_JIGBIL
,@StrDpt   = ISNULL(U_STRDPT,'0000')
,@EndDpt   = ISNULL(U_ENDDPT,'ZZZZ')
,@MstCod   = ISNULL(U_MSTCOD,'')
,@UserSign = UserSign
,@RetChk   = U_RETCHK
FROM [@PH_PY111A]
WHERE DocEntry = @iDocEntry
----------------- --------------- ---------------

DECLARE	 @dCLTCOD   nvarchar(10),
         @dMSTCOD   nvarchar(08),
         @dMSTNAM   nvarchar(20),
         @dSTDAMT   numeric(19,6),
         @dEmpID    nvarchar(10),
         @dTeamCode nvarchar(08),
         @dRspCode  nvarchar(08),
         @dPayTyp   nvarchar(08),
         @dJigCod   nvarchar(08),
         @dHoBong   nvarchar(08),
         @dBAEWOO   nvarchar(08),
         @dBUYNSU   int,
         @dJANGAE   int,
         @dGYNGL2   int,
         
         @dSILCUN   nvarchar(max),
         @dREMARK   nvarchar(max),
         @dCSUCOD   nvarchar(100),
         @dLINSEQ   smallint,
         @dLineNum  smallint,
         @dCSUCHK   nvarchar(08),
         @dGWATYP   nvarchar(08),
         @dDEDCHK   nvarchar(08),
         @dSequence int,
         @dLENGTH   nvarchar(08),
         @dROUNDT   nvarchar(08),
         @dICTCHK   nvarchar(08),
         @dWorkType nvarchar(100),
         @dSTATUS   nvarchar(100),
         
         @dOrder    nvarchar(100)
                  
DECLARE  @FrDate    datetime,
         @ToDate    datetime,
         @CSUCOD    nvarchar(08),
         @GONCOD    nvarchar(08),
         @COMCOD    nvarchar(08),
         @PayTyp    char(1),
         @GNSGBN    nvarchar(08),
         @101COD    nvarchar(08),
         @106COD    nvarchar(08),
         @109COD    nvarchar(08)

DECLARE  @AutoKey   int
DECLARE  @CNT       int
DECLARE  @Code      nvarchar(08)
DECLARE  @sql1      nvarchar(max)
DECLARE  @fild      nvarchar(100)
DECLARE  @RCNT      int

IF NOT EXISTS(SELECT * FROM SYSOBJECTS WHERE Name = 'ZPH_PY112A') 
    SELECT * INTO ZPH_PY112A FROM [@PH_PY112A]

SET @StrDpt = CASE WHEN @StrDpt = '%' THEN '0000' ELSE @StrDpt END
SET @EndDpt = CASE WHEN @EndDpt = '%' THEN 'ZZZZ' ELSE @EndDpt END
SET @MstCod = CASE WHEN @MstCod = '' THEN '%' ELSE @MstCod END

SET @FrDate = @YM+'01'
SET @ToDate = DATEADD(month,1,@FrDate)-1

-- 기존 DATA 삭제
DELETE T0
FROM [ZPH_PY112A] T0
JOIN [@PH_PY112A] T1 ON T1.Code = T0.Code
WHERE T1.U_CLTCOD = @CltCod 
AND T1.U_YM = @YM 
AND T1.U_JOBTYP = @JobTyp 
AND T1.U_JOBGBN = @JobGbn
AND T1.U_JOBTRG = @JobTrg 
AND T1.U_TeamCode BETWEEN @StrDpt AND @EndDpt 
AND T1.U_MSTCOD LIKE @MstCod
AND T1.U_ENDCHK = 'N'  -- 마감되지 않은것

DELETE FROM [@PH_PY112A]
WHERE U_CLTCOD = @CltCod 
AND U_YM = @YM 
AND U_JOBTYP = @JobTyp 
AND U_JOBGBN = @JobGbn
AND U_JOBTRG = @JobTrg 
AND U_TeamCode BETWEEN @StrDpt AND @EndDpt 
AND U_MSTCOD LIKE @MstCod
AND U_ENDCHK = 'N'  -- 마감되지 않은것

SELECT @CSUCOD=MAX(Code) FROM [@PH_PY102A] WHERE U_CLTCOD = @CltCod AND U_YM <= @YM
SELECT @GONCOD=MAX(Code) FROM [@PH_PY103A] WHERE U_CLTCOD = @CltCod AND U_YM <= @YM
SELECT @101COD=MAX(Code) FROM [PH_PY101V] WHERE Code <= @YM
SELECT @109COD=Code FROM [@PH_PY109A] WHERE U_CLTCOD = @CltCod AND U_YM = @YM AND U_JOBTYP = @JobTyp AND U_JOBGBN = @JobGbn AND U_JOBTRG = @JobTrg

SELECT * INTO #A112A FROM [@PH_PY112A] WHERE Code IS NULL --임시테이블생성
SET @CNT = 1

INSERT #A112A --[@PH_PY112A] 
       (Code,Name,DocEntry,Object,UserSign,CreateDate,CreateTime,DataSource
       ,U_CLTCOD,U_CLTNAM,U_YM,U_JOBTYP,U_TYPNAM --05
       ,U_JOBGBN,U_GBNNAM,U_JOBTRG,U_JIGBIL,U_ENDCHK --10
       ,U_RETCHK,U_CSUCOD,U_GONCOD,U_MSTCOD,U_MSTNAM --15
       ,U_EmpID,U_TeamCode,U_TeamName,U_RspCode,U_RspName --20
       ,U_PAYTYP,U_JIGCOD,U_HOBONG,U_STDAMT,U_INPDAT --25
       ,U_OUTDAT,U_BAEWOO,U_BUYNSU,U_JANGAE,U_GYNGL2 --30
       ,U_CSUD01,U_CSUD02,U_CSUD03,U_CSUD04,U_CSUD05 --35
       ,U_CSUD06,U_CSUD07,U_CSUD08,U_CSUD09,U_CSUD10 --40
       ,U_CSUD11,U_CSUD12,U_CSUD13,U_CSUD14,U_CSUD15 --45
       ,U_CSUD16,U_CSUD17,U_CSUD18,U_CSUD19,U_CSUD20 --50
       ,U_CSUD21,U_CSUD22,U_CSUD23,U_CSUD24,U_CSUD25 --55
       ,U_CSUD26,U_CSUD27,U_CSUD28,U_CSUD29,U_CSUD30 --60
       ,U_CSUD31,U_CSUD32,U_CSUD33,U_CSUD34,U_CSUD35 --65
       ,U_CSUD36,U_TOTPAY,U_GWASEE,U_bGWASEE --69
       ,U_GONG01,U_GONG02,U_GONG03,U_GONG04,U_GONG05 --74
       ,U_GONG06,U_GONG07,U_GONG08,U_GONG09,U_GONG10 --79
       ,U_GONG11,U_GONG12,U_GONG13,U_GONG14,U_GONG15 --84
       ,U_GONG16,U_GONG17,U_GONG18,U_GONG19,U_GONG20 --89
       ,U_GONG21,U_GONG22,U_GONG23,U_GONG24,U_GONG25 --94
       ,U_GONG26,U_GONG27,U_GONG28,U_GONG29,U_GONG30 --99
       ,U_GONG31,U_GONG32,U_GONG33,U_GONG34,U_GONG35 --104
       ,U_GONG36,U_TOTGON,U_SILJIG,U_AVRPAY,U_NABTAX --109
       ,U_BNSRAT,U_APPRAT,U_GNSYER,U_GNSMON,U_TAXTRM --114
       ,U_BONUSS,U_DAYAMT,U_BASAMT --117
       ,U_BTX01,U_BTX02,U_BTX03,U_BTX04,U_BTX05 --122
       ,U_BTX06,U_BTX07,U_BTX08,U_BTX09,U_BTX10 --127
       ,U_BTX11,U_BTX12,U_BTX13,U_BTX14,U_BTX15 --132
       ,U_BTX16,U_BTX17,U_BTX18,U_BTX19,U_BTX20 --137
       ,U_BTX21,U_BTX22,U_BTX23,U_BTX24,U_BTX25 --142
       ,U_BTX26,U_BTX27,U_BTX28,U_BTX29,U_BTX30 --147
       ,U_BTX31,U_BTX32,U_BTX33,U_BTX34,U_BTX35 --152
       ,U_BTX36,U_BTXFRE,U_TAXFRE --155
       ,U_Status,U_WorkType,U_Order --158
       )
SELECT '','',ROW_NUMBER()OVER(order by T0.U_TeamCode,T0.U_RspCode,T0.Code)
,'PH_PY112',@UserSign,CONVERT(char(8),GETDATE(),112),DATEPART(hh,GETDATE())*100+DATEPART(mi,GETDATE()),'I'
,T0.U_CLTCOD,'',@YM,@JOBTYP,'' --05
,@JOBGBN,'',@JOBTRG,@JIGBIL,'N' --10
,@RetChk,@CSUCOD,@GONCOD,T0.Code,T0.U_fullName --15
,T0.U_empID,T0.U_TeamCode,'',T0.U_RspCode,'' --20
,T0.U_PAYTYP,T0.U_JIGCOD,T0.U_HOBONG,T0.U_STDAMT,NULL --25
,NULL,T0.U_BAEWOO,T0.U_BUYN20,T0.U_JANGAE,T0.U_GYNGL2 --30
,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 --50
,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 --70
,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 --90
,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 --110
,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 --130
,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 --150
,0,0,0,0,0,T0.U_Status,NULL,0 --158
FROM [@PH_PY001A] T0
WHERE T0.U_CLTCOD = @CltCod
AND T0.U_PAYSEL = @JobTrg
AND T0.U_Status <> '5'
ORDER BY T0.U_TeamCode,T0.U_RspCode,T0.Code

-- 재직상태/휴직구분/휴직차수 갱신 --
UPDATE #A112A SET U_Status = '9' WHERE @JobGbn = '4' --잔여급여
UPDATE T0 SET U_Status = '8' --복직
FROM #A112A T0
JOIN [@PH_PY017B] T1 ON T1.Code = @CltCod+@YM AND T1.U_MSTCOD = T0.U_MSTCOD
WHERE T0.U_Status = '1' AND T1.U_EtcDAY4 > 0

UPDATE T0 SET U_WorkType = T2.WorkType --휴직구분
FROM #A112A T0
JOIN [@PH_PY017B] T1 ON T1.Code = @CltCod+@YM AND T1.U_MSTCOD = T0.U_MSTCOD
JOIN (SELECT MSTCOD,WorkType,COUNT(*)CNT
      FROM [ZPH_PY008] 
      WHERE PosDate BETWEEN @FrDate and @ToDate AND WorkType IN ('F01','F02','F03','F05')
      GROUP BY MSTCOD,WorkType) T2 ON T2.MSTCOD = T0.U_MSTCOD
WHERE T0.U_Status IN ('3','4','8') AND T1.U_EtcDAY4 > 0

SELECT T0.Code,T0.AppDate
INTO #LEAVE --휴직시작일 TABLE
FROM (
SELECT G.Code,MAX(U_appDate)AppDate
FROM [@PH_PY001A] A
JOIN [@PH_PY001G] G ON G.Code = A.Code AND A.U_CltCod = @CltCod
WHERE U_appDate <= @ToDate AND U_appType IN ('27','28','29','30') --휴직
GROUP BY G.Code) T0 
LEFT JOIN [@PH_PY001G] T1 ON T1.Code = T0.Code AND T1.U_appType = '34' --복직
                         AND T1.U_AppDate BETWEEN T0.AppDate AND @ToDate
WHERE T1.Code IS NULL

UPDATE T0 SET U_Order=DATEDIFF(month,T1.AppDate,@ToDate)+1 --휴직월차수
FROM #A112A T0
JOIN #LEAVE T1 ON T1.Code = T0.U_MSTCOD
JOIN [@PH_PY017B] T2 ON T2.Code = @CltCod+@YM AND T2.U_MSTCOD = T0.U_MSTCOD
WHERE T2.U_EtcDAY4 > 0
-- ================================ --
SELECT @CNT=COUNT(*) FROM #A112A

-- 월근태 집계에 위해수당 갱신 --
UPDATE T0 SET U_WHMAMT = ISNULL(T1.DGAMT,0)
FROM [@PH_PY017B] T0
JOIN (SELECT T0.MSTCOD,SUM(ISNULL(DangerNu,0)*ISNULL(T1.U_Num1,0))DGAMT
      FROM [ZPH_PY008] T0
      JOIN [@PS_HR200L] T1 ON T1.Code = 'P220' AND T1.U_Code = T0.DangerCD AND T1.U_Char2 = T0.CLTCOD
      WHERE T0.PosDate BETWEEN @FrDate AND @ToDate
      GROUP BY T0.MSTCOD) T1 ON T1.MSTCOD = T0.U_MSTCOD
WHERE T0.Code = @CltCod+@YM

-- 지급항목 시작 --
SET @RCNT = 1
--
WHILE @RCNT < 3 BEGIN
SET    @PayTyp = @RCNT

SELECT @COMCOD=MAX(Code) FROM [@PH_PY106A] WHERE U_CLTCOD = @CltCod AND U_YM <= @YM AND U_PAYTYP = @PayTyp
SELECT @GNSGBN=U_GNSGBN FROM [@PH_PY106A]  WHERE Code = @COMCOD

-- 입사일자,근속년수 갱신 --
IF @GNSGBN = '1' 
    UPDATE T0 SET U_INPDAT = CONVERT(CHAR(8),T1.U_GrpDat,112) 
    FROM #A112A T0 JOIN [@PH_PY001A] T1 ON T1.U_MSTCOD = T0.U_MSTCOD AND T1.U_PayTyp = @GNSGBN
IF @GNSGBN = '2' 
    UPDATE T0 SET U_INPDAT = CONVERT(CHAR(8),T1.U_StartDat,112) 
    FROM #A112A T0 JOIN [@PH_PY001A] T1 ON T1.U_MSTCOD = T0.U_MSTCOD AND T1.U_PayTyp = @GNSGBN

UPDATE #A112A SET U_GNSYER = DATEDIFF(year,U_INPDAT,@FrDate)

UPDATE #A112A SET U_GNSYER = U_GNSYER - 1 WHERE U_GNSYER > 0 AND DATEADD(year,U_GNSYER,U_INPDAT) > @FrDate

-- 고정수당 시작 --
DECLARE CUR1 CURSOR FOR
SELECT T0.U_CSUCOD,T0.U_LineNum,T0.U_CSUCHK,T0.U_GWATYP,T0.U_DEDCHK,T0.U_ICTCHK
FROM [@PH_PY102B] T0
WHERE @JobTyp = '1' AND T0.Code = @CSUCOD AND T0.U_FIXGBN = 'Y'
UNION ALL
SELECT T0.U_CSUCOD,T0.U_LineNum,T0.U_CSUCHK,T0.U_GWATYP,T0.U_DEDCHK,T0.U_ICTCHK
FROM [@PH_PY102B] T0
WHERE @JobTyp = '2' AND T0.Code = @CSUCOD AND T0.U_FIXGBN = 'Y' AND T0.U_BNSUSE = 'Y'

OPEN  CUR1
FETCH NEXT FROM CUR1 INTO @dCSUCOD,@dLineNum,@dCSUCHK,@dGWATYP,@dDEDCHK,@dICTCHK
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @dSILCUN = 'T2.U_FILD03'
    SET @sql1 = 'UPDATE T0 SET U_CSUD'+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')+'='+@dSILCUN
    -- 통상임금
    IF @dCSUCHK = 'Y'
        SET @sql1 = @sql1+',U_BASAMT=U_BASAMT+'+@dSILCUN
    
    SET @sql1 = @sql1+' FROM [#A112A] T0'
    SET @sql1 = @sql1+' JOIN [@PH_PY001A] T1 ON T1.Code = T0.U_MSTCOD'
    SET @sql1 = @sql1+' JOIN [@PH_PY001B] T2 ON T2.Code = T0.U_MSTCOD AND T2.U_FILD01 = '''+@dCSUCOD+''''
    SET @sql1 = @sql1+' WHERE T0.U_PAYTYP = '+@PayTyp

    EXEC sp_executesql @sql1
FETCH NEXT FROM CUR1 INTO @dCSUCOD,@dLineNum,@dCSUCHK,@dGWATYP,@dDEDCHK,@dICTCHK
END
CLOSE CUR1
DEALLOCATE CUR1
-- 고정수당 종료 --

-- 변동수당 시작 --
DECLARE CUR2 CURSOR FOR
SELECT T1.U_Sequence,T2.U_LineNum,T2.U_CSUCHK,T2.U_GWATYP,T2.U_DEDCHK,T2.U_ICTCHK
FROM [@PH_PY109Z] T1
JOIN [@PH_PY102B] T2 ON T2.U_CSUCOD = T1.U_PDCode
WHERE @JobTyp = '1' AND T1.Code = @109COD AND T2.Code = @CSUCOD
UNION ALL
SELECT T1.U_Sequence,T2.U_LineNum,T2.U_CSUCHK,T2.U_GWATYP,T2.U_DEDCHK,T2.U_ICTCHK
FROM [@PH_PY109Z] T1
JOIN [@PH_PY102B] T2 ON T2.U_CSUCOD = T1.U_PDCode
WHERE @JobTyp = '2' AND T1.Code = @109COD AND T2.Code = @CSUCOD AND T2.U_BNSUSE = 'Y'

OPEN  CUR2
FETCH NEXT FROM CUR2 INTO @dSequence,@dLineNum,@dCSUCHK,@dGWATYP,@dDEDCHK,@dICTCHK
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @dSILCUN = 'T2.U_AMT'+REPLACE(STR(ISNULL(@dSequence,0),2),' ','0')
    SET @sql1 = 'UPDATE T0 SET U_CSUD'+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')+'='+@dSILCUN
    -- 통상임금
    IF @dCSUCHK = 'Y'
        SET @sql1 = @sql1+',U_BASAMT=U_BASAMT+'+@dSILCUN
        
    SET @sql1 = @sql1+' FROM [#A112A] T0'
    SET @sql1 = @sql1+' JOIN [@PH_PY109B] T2 ON T2.Code = '''+@109COD+''' AND T2.U_MSTCOD = T0.U_MSTCOD'
    SET @sql1 = @sql1+' WHERE T0.U_PAYTYP = '+@PayTyp

    EXEC sp_executesql @sql1
FETCH NEXT FROM CUR2 INTO @dSequence,@dLineNum,@dCSUCHK,@dGWATYP,@dDEDCHK,@dICTCHK
END
CLOSE CUR2
DEALLOCATE CUR2
-- 변동수당 종료 --

-- 계산수당 시작 --
DECLARE CUR3 CURSOR FOR
SELECT T1.U_SILCUN,T1.U_REMARK,T0.U_CSUCOD,T0.U_LineNum,T0.U_CSUCHK
,T0.U_GWATYP,T0.U_DEDCHK,T0.U_LENGTH,T0.U_ROUNDT,T0.U_ICTCHK,T1.U_LINSEQ
FROM [@PH_PY102B] T0
JOIN [@PH_PY106B] T1 ON T1.U_CSUCOD = T0.U_CSUCOD
WHERE @JobTyp = '1' AND T0.Code = @CSUCOD AND T1.Code = @COMCOD
AND CAST(T1.U_SILCUN AS nvarchar(max)) <> '0'
UNION ALL
SELECT CASE WHEN CAST(T1.U_SILCOD AS nvarchar(max)) <> '' THEN T1.U_SILCOD ELSE T1.U_SILCUN END
,T1.U_REMARK,T0.U_CSUCOD,T0.U_LineNum,T0.U_CSUCHK
,T0.U_GWATYP,T0.U_DEDCHK,T0.U_LENGTH,T0.U_ROUNDT,T0.U_ICTCHK,T1.U_LINSEQ
FROM [@PH_PY102B] T0
JOIN [@PH_PY106B] T1 ON T1.U_CSUCOD = T0.U_CSUCOD
WHERE @JobTyp = '2' AND T0.Code = @CSUCOD AND T1.Code = @COMCOD
AND T0.U_BNSUSE = 'Y'
AND (CAST(T1.U_SILCOD AS nvarchar(max)) <> '' OR CAST(T1.U_SILCUN AS nvarchar(max)) <> '0')
ORDER BY T0.U_CSUCHK DESC,T1.U_LINSEQ

OPEN  CUR3
FETCH NEXT FROM CUR3 INTO @dSILCUN,@dREMARK,@dCSUCOD,@dLineNum,@dCSUCHK
                         ,@dGWATYP,@dDEDCHK,@dLENGTH,@dROUNDT,@dICTCHK,@dLINSEQ
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @dSILCUN = REPLACE(@dSILCUN,'X21',''''+@YM+'''')
    SET @dSILCUN = REPLACE(@dSILCUN,'X22',''''+@JobTyp+'''')
    SET @dSILCUN = REPLACE(@dSILCUN,'X24',''''+@GNSGBN+'''')
    
    SET @dSILCUN = '('+@dSILCUN+')/'+@dLENGTH -- 적용자릿수지정
    IF @dROUNDT = 'R' SET @dSILCUN = 'ROUND('+@dSILCUN+',0)'
    IF @dROUNDT = 'F' SET @dSILCUN = 'FLOOR('+@dSILCUN+')'
    IF @dROUNDT = 'C' SET @dSILCUN = 'CEILING('+@dSILCUN+')'
    SET @dSILCUN = '('+@dSILCUN+')*'+@dLENGTH -- 적용자릿수지정

    SET @dREMARK = '('+@dREMARK+')/'+@dLENGTH -- 적용자릿수지정
    IF @dROUNDT = 'R' SET @dREMARK = 'ROUND('+@dREMARK+',0)'
    IF @dROUNDT = 'F' SET @dREMARK = 'FLOOR('+@dREMARK+')'
    IF @dROUNDT = 'C' SET @dREMARK = 'CEILING('+@dREMARK+')'
    SET @dREMARK = '('+@dREMARK+')*'+@dLENGTH -- 적용자릿수지정

    SET @sql1 = 'UPDATE T0 SET U_CSUD'+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')+'=U_CSUD'
    SET @sql1 = @sql1+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')+'+'+@dSILCUN
    -- 통상임금
    IF @dCSUCHK = 'Y'
        IF @dCSUCOD = 'E07' --통상임금 교대수당은 계산식의 비고란의 식을 사용함
        SET @sql1 = @sql1+',U_BASAMT=U_BASAMT+'+@dREMARK
        ELSE
        SET @sql1 = @sql1+',U_BASAMT=U_BASAMT+'+@dSILCUN

    SET @sql1 = @sql1+' FROM [#A112A] T0'
    SET @sql1 = @sql1+' JOIN [@PH_PY001A] T1 ON T1.Code = T0.U_MSTCOD'
    SET @sql1 = @sql1+' LEFT JOIN [@PH_PY017B] T6 ON T6.Code = '''+@CLTCOD+@YM+''' AND T6.U_MSTCOD = T0.U_MSTCOD'
    SET @sql1 = @sql1+' WHERE T0.U_PAYTYP = '''+@PayTyp+''''

    EXEC sp_executesql @sql1
FETCH NEXT FROM CUR3 INTO @dSILCUN,@dREMARK,@dCSUCOD,@dLineNum,@dCSUCHK
                         ,@dGWATYP,@dDEDCHK,@dLENGTH,@dROUNDT,@dICTCHK,@dLINSEQ
END
CLOSE CUR3
DEALLOCATE CUR3
-- 계산수당 종료 --

SET @RCNT = @RCNT + 1
END --WHILE END
-- 지급항목 종료 --

SELECT * INTO [#Z112A] FROM [#A112A]
--================================================================================--
-- 예외(휴직,수습,잔여)적용 시작 --
SET @RCNT = 1
--
WHILE @RCNT < 3 BEGIN
SET    @PayTyp = @RCNT

SELECT @COMCOD=MAX(Code) FROM [@PH_PY106A] WHERE U_CLTCOD = @CltCod AND U_YM <= @YM AND U_PAYTYP = @PayTyp
-- 근무일,휴직일 계산
UPDATE [@PH_PY017B] SET U_EtcDAY8=(U_StdGDay+U_StdGDay+U_StdNDay)-U_EtcDAY4,U_EtcDAY9=U_EtcDAY4 WHERE Code = @CltCod+@YM
UPDATE [@PH_PY017B] SET U_EtcDAY9=30-U_EtcDAY8 WHERE Code = @CltCod+@YM AND U_EtcDAY8+U_EtcDAY9 > 30
UPDATE [@PH_PY017B] SET U_EtcDAY8=30-U_EtcDAY9 WHERE Code = @CltCod+@YM AND U_EtcDAY8+U_EtcDAY9 < 30

DECLARE CUR9 CURSOR FOR
SELECT T0.U_Status,ISNULL(T0.U_WorkType,''),ISNULL(T0.U_Order,''),T0.U_CSUCOD,T0.U_SILCUN
FROM [@PH_PY106D] T0
WHERE T0.Code = @COMCOD AND ISNULL(T0.U_Status,'') <> ''
ORDER BY T0.U_LINSEQ

OPEN  CUR9
FETCH NEXT FROM CUR9 INTO @dStatus,@dWorkType,@dOrder,@dCSUCOD,@dSILCUN
WHILE @@FETCH_STATUS = 0
BEGIN
    --SET @dCSUCOD = ''''+REPLACE(@dCSUCOD,',',''',''')+''''
    
    DECLARE CUR0 CURSOR FOR
    SELECT U_LineNum
    FROM [@PH_PY102B] 
    WHERE Code = @CSUCOD AND U_CSUCOD IN (@dCSUCOD)
    ORDER BY U_LINSEQ

	OPEN  CUR0
	FETCH NEXT FROM CUR0 INTO @dLineNum
	WHILE @@FETCH_STATUS = 0
	BEGIN
		SET @fild = 'U_CSUD'+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')
	    
		SET @sql1 = 'UPDATE T0 SET '+@fild+'=ROUND('+@fild+'*'+@dSILCUN+',0)'
		SET @sql1 = @sql1+' FROM [#A112A] T0'
		SET @sql1 = @sql1+' JOIN [@PH_PY017B] T6 ON T6.Code = '''+@CLTCOD+@YM+''' AND T6.U_MSTCOD = T0.U_MSTCOD'
		SET @sql1 = @sql1+' WHERE T0.U_PAYTYP = '''+@PayTyp+''''
		SET @sql1 = @sql1+' AND U_Status IN ('''+REPLACE(@dStatus,',',''',''')+''')'
		IF @dWorkType <> ''
		SET @sql1 = @sql1+' AND U_WorkType IN ('''+REPLACE(@dWorkType,',',''',''')+''')'
		IF @dOrder <> ''
		SET @sql1 = @sql1+' AND U_Order IN ('+@dOrder+')'
	    
		EXEC sp_executesql @sql1
		print @sql1
	FETCH NEXT FROM CUR0 INTO @dLineNum
	END
	CLOSE CUR0
	DEALLOCATE CUR0    
FETCH NEXT FROM CUR9 INTO @dStatus,@dWorkType,@dOrder,@dCSUCOD,@dSILCUN
END
CLOSE CUR9
DEALLOCATE CUR9

SET @RCNT = @RCNT + 1
END --WHILE END
-- 예외적용 종료 --
--================================================================================--
-- 누적금액 재계산 시작 --
UPDATE [#A112A] SET
 U_TOTPAY=0,U_SILJIG=0,U_GWASEE=0,U_bGWASEE=0,U_DAYAMT=0 
WHERE U_PAYTYP = @PayTyp

DECLARE CUR8 CURSOR FOR
SELECT T0.U_LineNum,T0.U_CSUCOD,T0.U_GWATYP,T0.U_DEDCHK,T0.U_ICTCHK
FROM [@PH_PY102B] T0
WHERE T0.Code = @CSUCOD

OPEN  CUR8
FETCH NEXT FROM CUR8 INTO @dLineNum,@dCSUCOD,@dGWATYP,@dDEDCHK,@dICTCHK
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @fild = 'U_CSUD'+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')
    SET @sql1 = 'UPDATE T0 SET '

    -- 지급총액
    SET @sql1 = @sql1+'U_TOTPAY=U_TOTPAY+'+@fild
    -- 실지급액
    IF @dDEDCHK = 'N' -- 공제항목으로표시
    SET @sql1 = @sql1+',U_SILJIG=U_SILJIG+'+@fild
    -- 과세,비과세
    IF @dGWATYP = '1'  -- 
        SET @sql1 = @sql1+',U_GWASEE=U_GWASEE+'+@fild
    ELSE
        SET @sql1 = @sql1+',U_bGWASEE=U_bGWASEE+'+@fild
    -- 소득세계산기준
    IF @dICTCHK = 'Y'
        SET @sql1 = @sql1+',U_DAYAMT=U_DAYAMT+'+@fild

    SET @sql1 = @sql1+' FROM [#A112A] T0' -- WHERE T0.U_PAYTYP = '''+@PayTyp+''''

    EXEC sp_executesql @sql1
FETCH NEXT FROM CUR8 INTO @dLineNum,@dCSUCOD,@dGWATYP,@dDEDCHK,@dICTCHK
END
CLOSE CUR8
DEALLOCATE CUR8
-- 누적금액 재계산 종료 --
--================================================================================--
-- 공제항목 시작 --

-- 고정공제 시작 --
DECLARE CUR4 CURSOR FOR
SELECT T0.U_CSUCOD,T0.U_LineNum
FROM [@PH_PY103B] T0
WHERE @JobTyp = '1' AND T0.Code = @GONCOD AND T0.U_FIXGBN = 'Y'
UNION ALL
SELECT T0.U_CSUCOD,T0.U_LineNum
FROM [@PH_PY103B] T0
WHERE @JobTyp = '2' AND T0.Code = @GONCOD AND T0.U_FIXGBN = 'Y' AND T0.U_BNSUSE = 'Y'

OPEN  CUR4
FETCH NEXT FROM CUR4 INTO @dCSUCOD,@dLineNum
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @dSILCUN = 'T2.U_FILD03'
    SET @sql1 = 'UPDATE T0 SET U_GONG'+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')+'='+@dSILCUN

    -- 공제총액
    SET @sql1 = @sql1+',U_TOTGON=U_TOTGON+'+@dSILCUN
    -- 실지급액
    SET @sql1 = @sql1+',U_SILJIG=U_SILJIG-'+@dSILCUN

    SET @sql1 = @sql1+' FROM [#A112A] T0'
    SET @sql1 = @sql1+' JOIN [@PH_PY001A] T1 ON T1.Code = T0.U_MSTCOD'
    SET @sql1 = @sql1+' JOIN [@PH_PY001C] T2 ON T2.Code = T0.U_MSTCOD AND T2.U_FILD01 = '''+@dCSUCOD+''''

    EXEC sp_executesql @sql1
FETCH NEXT FROM CUR4 INTO @dCSUCOD,@dLineNum
END
CLOSE CUR4
DEALLOCATE CUR4
-- 고정수당 종료 --


-- 변동공제 시작 --
DECLARE CUR2 CURSOR FOR
SELECT T1.U_Sequence,T2.U_LineNum
FROM [@PH_PY109Z] T1
JOIN [@PH_PY102B] T2 ON T2.U_CSUCOD = T1.U_PDCode
WHERE @JobTyp = '1' AND T1.Code = @109COD AND T2.Code = @CSUCOD
UNION ALL
SELECT T1.U_Sequence,T2.U_LineNum
FROM [@PH_PY109Z] T1
JOIN [@PH_PY102B] T2 ON T2.U_CSUCOD = T1.U_PDCode
WHERE @JobTyp = '2' AND T1.Code = @109COD AND T2.Code = @CSUCOD AND T2.U_BNSUSE = 'Y'

OPEN  CUR2
FETCH NEXT FROM CUR2 INTO @dSequence,@dLineNum
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @dSILCUN = 'T2.U_AMT'+REPLACE(STR(ISNULL(@dSequence,0),2),' ','0')
    SET @sql1 = 'UPDATE T0 SET U_GONG'+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')+'='+@dSILCUN

    -- 공제총액
    SET @sql1 = @sql1+',U_TOTGON=U_TOTGON+'+@dSILCUN
    -- 실지급액
    SET @sql1 = @sql1+',U_SILJIG=U_SILJIG-'+@dSILCUN

    SET @sql1 = @sql1+' FROM [#A112A] T0'
    SET @sql1 = @sql1+' JOIN [@PH_PY109B] T2 ON T2.Code = '''+@109COD+''' AND T2.U_MSTCOD = T0.U_MSTCOD'

    EXEC sp_executesql @sql1
FETCH NEXT FROM CUR2 INTO @dSequence,@dLineNum
END
CLOSE CUR2
DEALLOCATE CUR2
-- 변동공제 종료 --

-- 계산공제 시작 --
DECLARE CUR3 CURSOR FOR
SELECT T0.U_SILCUN,T0.U_CSUCOD,T0.U_LineNum,T0.U_LENGTH,T0.U_ROUNDT,T0.U_LINSEQ
FROM [@PH_PY103B] T0
WHERE @JobTyp = '1' AND T0.Code = @GONCOD AND CAST(T0.U_SILCUN AS nvarchar(1000)) <> '0'
UNION ALL
SELECT T0.U_SILCUN,T0.U_CSUCOD,T0.U_LineNum,T0.U_LENGTH,T0.U_ROUNDT,T0.U_LINSEQ
FROM [@PH_PY103B] T0
WHERE @JobTyp = '2' AND T0.Code = @GONCOD AND CAST(T0.U_SILCUN AS nvarchar(1000)) <> '0'
AND T0.U_BNSUSE = 'Y'
ORDER BY T0.U_LINSEQ

OPEN  CUR3
FETCH NEXT FROM CUR3 INTO @dSILCUN,@dCSUCOD,@dLineNum,@dLENGTH,@dROUNDT,@dLINSEQ
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @dSILCUN = REPLACE(@dSILCUN,'X21',''''+@YM+'''')

    SET @dSILCUN = '('+@dSILCUN+')/'+@dLENGTH -- 적용자릿수지정
    IF @dROUNDT = 'R' SET @dSILCUN = 'ROUND('+@dSILCUN+',0)'
    IF @dROUNDT = 'F' SET @dSILCUN = 'FLOOR('+@dSILCUN+')'
    IF @dROUNDT = 'C' SET @dSILCUN = 'CEILING('+@dSILCUN+')'
    SET @dSILCUN = '('+@dSILCUN+')*'+@dLENGTH -- 적용자릿수지정

    SET @sql1 = 'UPDATE T0 SET U_GONG'+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')+'=U_GONG'
    SET @sql1 = @sql1+REPLACE(STR(ISNULL(@dLineNum,0),2),' ','0')+'+'+@dSILCUN

    -- 공제총액
    SET @sql1 = @sql1+',U_TOTGON=U_TOTGON+'+@dSILCUN
    -- 실지급액
    SET @sql1 = @sql1+',U_SILJIG=U_SILJIG-'+@dSILCUN

    SET @sql1 = @sql1+' FROM [#A112A] T0'
    SET @sql1 = @sql1+' JOIN [@PH_PY001A] T1 ON T1.Code = T0.U_MSTCOD'
    SET @sql1 = @sql1+' LEFT JOIN [PH_PY101V] T2 ON T2.Code = '''+@101COD+''''

    EXEC sp_executesql @sql1
FETCH NEXT FROM CUR3 INTO @dSILCUN,@dCSUCOD,@dLineNum,@dLENGTH,@dROUNDT,@dLINSEQ
END
CLOSE CUR3
DEALLOCATE CUR3
-- 계산공제 종료 --

-- 공제항목 종료 --
--================================================================================--
-- Last process --
SELECT @AutoKey=AutoKey FROM ONNM WHERE ObjectCode = 'PH_PY112'

UPDATE [#A112A] SET
 Code=REPLACE(STR(ISNULL(DocEntry+@AutoKey-1,0),8),' ','0')
,Name=REPLACE(STR(ISNULL(DocEntry+@AutoKey-1,0),8),' ','0')
,DocEntry=DocEntry+@AutoKey-1

INSERT [@PH_PY112A] SELECT * FROM [#A112A]

UPDATE ONNM SET AutoKey = @AutoKey+@CNT WHERE ObjectCode = 'PH_PY112'
-------------------
UPDATE [#Z112A] SET
 Code=REPLACE(STR(ISNULL(DocEntry+@AutoKey-1,0),8),' ','0')
,Name=REPLACE(STR(ISNULL(DocEntry+@AutoKey-1,0),8),' ','0')
,DocEntry=DocEntry+@AutoKey-1

INSERT [ZPH_PY112A] SELECT * FROM [#Z112A]

/* 기 작업내용 삭제
DELETE FROM [@PH_PY112A]
UPDATE ONNM SET AutoKey = 1 WHERE ObjectCode = 'PH_PY112'
*/



