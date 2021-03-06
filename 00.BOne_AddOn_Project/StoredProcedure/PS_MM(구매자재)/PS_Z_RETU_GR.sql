USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_Z_RETU_GR]    Script Date: 11/29/2010 18:03:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROC [dbo].[PS_Z_RETU_GR]
    (
        @iDocEntry   as int,
        @oError      as int OUTPUT,
        @oErMsg      as nvarchar(200) OUTPUT
    )
AS
/**-----------------------------------------------------
     Procedure ID       : PS_Z_RETU_GR
     Create Name        : 
            Date        : 
     Update Name        : Minho Choi
            Date        : 2010.11.29
     Inpur parameter    : @TransType  as smallint
                        : @DocEntry   as int
     Return             : output Non @Parameter
     Execute            : EXECUTE [PS_Z_RETU_GR] 22
     Company            : ㈜모닝정보
     Description        : 반품(입고취소)시 발주수량 회복
-----------------------------------------------------**/

DECLARE      @DocEntry      int,
             @LineNum       int,
             @Price         numeric(19,6),
             @Rate          numeric(19,6),
             @VatPrcnt      numeric(19,6),
             @RTQty         numeric(19,6),
             @ItemCode      nvarchar(20),
             @WhsCode       nvarchar(8)

DECLARE CUR1 CURSOR LOCAL FOR
 SELECT T3.DocEntry,T3.LineNum,T3.Price,ISNULL(T3.Rate,0),T3.VatPrcnt,T1.Quantity
       ,T3.ItemCode,T3.WhsCode
   FROM ORPD T0 JOIN RPD1 T1 ON T1.DocEntry = T0.DocEntry
                JOIN PDN1 T2 ON T2.ObjType = T1.BaseType AND T2.DocEntry = T1.BaseEntry AND T2.LineNum = T1.BaseLine
                JOIN POR1 T3 ON T3.ObjType = T2.BaseType AND T3.DocEntry = T2.BaseEntry AND T3.LineNum = T2.BaseLine
  WHERE T0.DocEntry = @iDocEntry

OPEN CUR1
FETCH NEXT FROM CUR1 INTO @DocEntry,@LineNum,@Price,@Rate,@VatPrcnt,@RTQty,@ItemCode,@WhsCode
WHILE (@@FETCH_STATUS = 0) BEGIN
    UPDATE POR1 SET
    OpenQty    = OpenQty    + @RTQty,
    OpenCreQty = OpenCreQty + @RTQty,
    VatAppld   = VatAppld   - (@RTQty*@Price*@VatPrcnt/100),
    VatAppldFC = VatAppldFC - (@RTQty*@Price*@Rate*@VatPrcnt/100),
    VatAppldSC = VatAppldSC - (@RTQty*@Price*@VatPrcnt/100)
    WHERE DocEntry = @DocEntry AND LineNum = @LineNum

    UPDATE POR1 SET
    TargetType = CASE WHEN Quantity = OpenQty THEN  -1   ELSE TargetType END,
    TrgetEntry = CASE WHEN Quantity = OpenQty THEN  NULL ELSE TrgetEntry END,
    LineStatus = 'O',
    InvntSttus = 'O'
    WHERE DocEntry = @DocEntry

    UPDATE OPOR SET
    DocStatus = 'O',
    InvntSttus = 'O',
    PaidToDate = PaidToDate - (@RTQty*@Price*(100+@VatPrcnt)/100),       
    PaidFC     = PaidFC     - (@RTQty*@Price*@Rate*(100+@VatPrcnt)/100),
    PaidSys    = PaidSys    - (@RTQty*@Price*(100+@VatPrcnt)/100),
    VatPaid    = VatPaid    - (@RTQty*@Price*@VatPrcnt/100),      
    VatPaidFC  = VatPaidFC  - (@RTQty*@Price*@Rate*@VatPrcnt/100),
    VatPaidSys = VatPaidSys - (@RTQty*@Price*@VatPrcnt/100)       
    WHERE DocEntry = @DocEntry

    UPDATE OITM SET OnOrder = OnOrder + @RTQty WHERE ItemCode = @ItemCode
    UPDATE OITW SET OnOrder = OnOrder + @RTQty WHERE ItemCode = @ItemCode AND WhsCode = @WhsCode

    FETCH NEXT FROM CUR1 INTO @DocEntry,@LineNum,@Price,@Rate,@VatPrcnt,@RTQty,@ItemCode,@WhsCode
END
CLOSE CUR1
DEALLOCATE CUR1
