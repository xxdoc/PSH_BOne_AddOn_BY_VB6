/*

                     수 불 명 세 서 (임가공 원자재 제외)

*/
DECLARE @BPLId    smallint,
        @FrDate   datetime,
        @ToDate   datetime,
        @AcctCode nvarchar(15) 
/* select t0.DocDate from OINM t0 */
/* SELECT * FROM JDT1 t1 */
/* SELECT * FROM OITB t2 */
SET @BPLId =  /* t1.U_BPLId */ [%0]
SET @FrDate = /* t0.DocDate */  [%1]
SET @ToDate = /* t0.DocDate */[%2]
SET @AcctCode =  /* t2.U_InvntAct */ [%3]

EXEC [MDC_InOut_QUERY_Detail] @BPLId, @FrDate, @ToDate, @AcctCode