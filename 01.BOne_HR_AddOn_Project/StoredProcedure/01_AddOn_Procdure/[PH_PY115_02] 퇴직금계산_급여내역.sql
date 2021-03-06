
-- =============================================
-- Procedure ID : PH_PY115_2
-- Author       : 
-- Create date  : 2012.12.05
-- Description  : 퇴직금계산 > 급여내역
-- EXEC PH_PY115_2 15
-- =============================================
--IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'PH_PY115_2' AND xtype = 'P'))
--	DROP PROCEDURE PH_PY115_2
--GO

--CREATE   PROCEDURE [dbo].[PH_PY115_2]
--         @MSTCOD  AS nvarchar(10)
--AS

DECLARE	 @CLTCOD	nvarchar(10),
		 @MSTCOD	nvarchar(10),
         @YM1		nvarchar(06),
         @YM2		nvarchar(08),
         @YM3		nvarchar(08),
         @YM4		nvarchar(08)

set @CLTCOD = '1'
set @SMTCOD = '2071201'
set @YM = '201213'

select  U_YM,
		'' as CSUCOD,
		'' as CSUNAM,
		CSUCOD2,
		CSUAMT
--INTO #TEMP2
FROM (
	select U_YM ,U_CSUD01, U_CSUD02, U_CSUD03, U_CSUD04, U_CSUD05, U_CSUD06, U_CSUD07, U_CSUD08, U_CSUD09, U_CSUD10,
		   U_CSUD11, U_CSUD12, U_CSUD13, U_CSUD14, U_CSUD15, U_CSUD16, U_CSUD17, U_CSUD18, U_CSUD19, U_CSUD20,
		   U_CSUD21, U_CSUD22, U_CSUD23, U_CSUD24, U_CSUD25, U_CSUD26, U_CSUD27, U_CSUD28, U_CSUD29, U_CSUD30,
		   U_CSUD31, U_CSUD32, U_CSUD33, U_CSUD34, U_CSUD35, U_CSUD36
	FROM YPH_PY112R
)A
UNPIVOT (
	CSUAMT FOR CSUCOD2 in (
		U_CSUD01, U_CSUD02, U_CSUD03, U_CSUD04, U_CSUD05, U_CSUD06, U_CSUD07, U_CSUD08, U_CSUD09, U_CSUD10,
		U_CSUD11, U_CSUD12, U_CSUD13, U_CSUD14, U_CSUD15, U_CSUD16, U_CSUD17, U_CSUD18, U_CSUD19, U_CSUD20,
		U_CSUD21, U_CSUD22, U_CSUD23, U_CSUD24, U_CSUD25, U_CSUD26, U_CSUD27, U_CSUD28, U_CSUD29, U_CSUD30,
		U_CSUD31, U_CSUD32, U_CSUD33, U_CSUD34, U_CSUD35, U_CSUD36 ) 
 ) PV  
 
 --select U_CSUCOD, U_CSUNAM from [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code
--WHERE T1.U_CLTCOD = (SELECT U_CLTCOD FROM [@PH_PY001A] WHERE Code = '2071201')
--AND T1.U_YM = @YM OR (T1.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY102A] WHERE U_YM < @YM))


--select * from [@PH_PY115B]
--SELECT T0.U_CSUNAM
--FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code

--select * from [@PH_PY102A]
--select * from [@PH_PY102B]
--select * from YPH_PY112R