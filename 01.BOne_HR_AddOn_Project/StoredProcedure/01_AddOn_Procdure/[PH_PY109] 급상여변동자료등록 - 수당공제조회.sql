-- =============================================
-- Procedure ID : PH_PY109
-- Author       : Minho Choi
-- Create date  : 2012.12.14
-- Description  : 변동자료 문서 생성
-- EXEC PH_PY109 '1','201211','1','1','1'
-- =============================================
IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY109' AND xtype = 'P'))
	DROP PROCEDURE PH_PY109
GO

Create   PROCEDURE [dbo].[PH_PY109]
         @iCltCod   nvarchar(08), --사업장
         @iYM       nvarchar(06), --귀속년월
         @iJobTyp   nvarchar(08), --지급종류
         @iJobGbn   nvarchar(08), --지급구분
         @iJobTrg   nvarchar(08)  --지급대상
AS

DECLARE	 @DocEntry  int,
         @Code      nvarchar(08),
         @CSUCOD    nvarchar(08),
         @GONCOD    nvarchar(08)

SELECT @CSUCOD=MAX(Code) FROM [@PH_PY102A] WHERE U_CLTCOD = @iCltCod AND U_YM <= @iYM
SELECT @GONCOD=MaX(Code) FROM [@PH_PY103A] WHERE U_CLTCOD = @iCltCod AND U_YM <= @iYM

SELECT @DocEntry=AutoKey FROM ONNM WHERE ObjectCode = 'PH_PY109'

SET @Code = @iCltCod+SUBSTRING(@iYM,3,4)+@iJobTyp+@iJobGbn+@iJobTrg

-- Insert Header --
INSERT [@PH_PY109A] (Code,Name,DocEntry,Canceled,Object
                    ,LogInst,UserSign,Transfered,CreateDate,CreateTime
                    ,UpdateDate,UpdateTime,DataSource
                    ,U_CLTCOD,U_YM,U_JOBTYP,U_JOBGBN,U_JOBTRG
                    )
VALUES (@Code,@Code,@DocEntry,'N','PH_PY109'
       ,NULL,1,'N',CONVERT(char(8),GETDATE(),112),DATEPART(hh,GETDATE())*100+DATEPART(mi,GETDATE())
       ,NULL,NULL,'I'
       ,@iCltCod,@iYM,@iJobTyp,@iJobGbn,@iJobTrg)

-- Insert Characteristic --
INSERT [@PH_PY109Z] (Code,LineId,Object,LogInst,U_LineNum,U_Sequence,U_PayDud,U_PDCode,U_PDName)
SELECT @Code
      ,ROW_NUMBER()OVER(order by DataType,U_LinSeq)
      ,'PH_PY109'
      ,NULL
      ,ROW_NUMBER()OVER(order by DataType,U_LinSeq)
      ,ROW_NUMBER()OVER(order by DataType,U_LinSeq)
      ,DataType
      ,U_CSUCOD
      ,U_CSUNAM
FROM (SELECT '1'DataType,U_LinSeq,U_CSUCOD,U_CSUNAM FROM [@PH_PY102B] 
      WHERE Code = @CSUCOD AND U_FIXGBN = 'V'
      UNION ALL
      SELECT '2'DataType,U_LinSeq,U_CSUCOD,U_CSUNAM FROM [@PH_PY103B] 
      WHERE Code = @GONCOD AND U_FIXGBN = 'V') T0
      
-- Insert Lines --
INSERT [@PH_PY109B] (Code,LineId,Object,LogInst,U_LineNum,U_MSTCOD,U_MSTNAM,U_DPTCOD,U_DPTNAM
,U_AMT01,U_AMT02,U_AMT03,U_AMT04,U_AMT05,U_AMT06,U_AMT07,U_AMT08,U_AMT09,U_AMT10,U_AMT11,U_AMT12)
SELECT @Code
      ,ROW_NUMBER()OVER(order by T0.U_TeamCode,T0.U_RspCode,T0.U_ClsCode,T0.Code)
      ,'PH_PY109'
      ,NULL
      ,ROW_NUMBER()OVER(order by T0.U_TeamCode,T0.U_RspCode,T0.U_ClsCode,T0.Code)
      ,T0.Code
      ,T0.U_FullName
      ,T0.U_TeamCode
      ,T1.U_CodeNm
      ,0,0,0,0,0,0,0,0,0,0,0,0
FROM [@PH_PY001A] T0
LEFT JOIN [@PS_HR200L] T1 ON T1.Code = '1' AND T1.U_Code = T0.U_TeamCode
WHERE U_Status <> '5'
AND U_CLTCOD = @iCltCod
AND U_PAYSEL = @iJobTrg

UPDATE ONNM SET AutoKey = @DocEntry + 1 WHERE ObjectCode = 'PH_PY109' AND AutoKey = @DocEntry