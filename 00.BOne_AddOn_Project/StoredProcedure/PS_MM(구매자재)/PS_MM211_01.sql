SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************************************************************************************************************/
/*  Module         : 구매관리																				    */
/*  Description    : 수불대장    																				*/
/*  ALTER  Date    : 2010.11.23  																				*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_MM211_01]
ALTER     PROC [dbo].[PS_MM211_01]
(
  @SItemCode		as nvarchar(20),
  @EItemCode		as nvarchar(20),
  @StrDate			as datetime,
  @EndDate			as datetime
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
drop table PS_MM211
CREATE TABLE PS_MM211 
			(
             ItemCode        nvarchar(20),
             Itemname        nvarchar(100),
             VisOrder        int,
             TransNum        int,			--Key
             DocDate         datetime,
		
			 CardCode		 Nvarchar(10),
			 CardName		 Nvarchar(50),
             TransType       smallint,		--Type
             CreatedBy       int,			--Type에 따른 문서 DocEntry
			 DocLineNum		 int,			--Type에 따른 문서 Line 번호
             Warehouse       nvarchar(08),
             JrnlMemo        nvarchar(50),
             InQty           numeric(19,6),
             OutQty          numeric(19,6),
             Price           numeric(19,6),
             TransValue      numeric(19,6),
             StockQty        numeric(19,6),
             StockAmt        numeric(19,6),

             CreateDate      datetime,
             CreateTime      smallint,
             UserSign        smallint
             --U_ShotName      nvarchar(20),
             --U_Material      nvarchar(30),
             --U_Size          nvarchar(50)
			)

DECLARE   @DocDate        datetime,
          @ItemCode       nvarchar(20),
          @WhsCode        nvarchar(8),
          @Quantity       numeric(19,6),
          @Amount         numeric(19,6),
          @OutQty         numeric(19,6),
          @OutAmt         numeric(19,6),

          @TransNum       int

----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

select @TransNum=AutoKey from onnm where ObjectCode = '58'

select *
  INTO #ZOINM
  from OINM
 where DocDate <= @EndDate
   AND ItemCode BETWEEN @SItemCode AND @EItemCode
   AND isnull(ApplObj,0) <> '911'

/*
/*수입부대비용 및 단가소급정산*/
DECLARE CUR1 CURSOR LOCAL FOR
 SELECT U_DocDate,U_ItemCode,U_WhsCode,U_Quantity,U_Amount,U_OutQty,U_OutAmt
   FROM [@PSH_ZMM901]
  WHERE U_DocDate <= @EndDate
  ORDER BY U_ItemCode,U_DocDate

OPEN CUR1
FETCH NEXT FROM CUR1 INTO @DocDate,@ItemCode,@WhsCode,@Quantity,@Amount,@OutQty,@OutAmt
WHILE (@@FETCH_STATUS = 0) BEGIN
    INSERT #ZOINM (TransNum,Instance,TransType,DocDate,ItemCode
                  ,Warehouse,InQty,OutQty,TransValue,CreatedBy,DocLineNum)
    VALUES(@TransNum,0,'931',@DocDate,@ItemCode,@WhsCode,0,0,@Amount,0,0)
    IF @OutQty <> 0 BEGIN
        SET @TransNum=@TransNum+1
        INSERT #ZOINM (TransNum,Instance,TransType,DocDate,ItemCode
                      ,Warehouse,InQty,OutQty,TransValue,CreatedBy,DocLineNum)
        VALUES(@TransNum,0,'932',@DocDate,@ItemCode,@WhsCode,0,0,-@OutAmt,0,0)
    END
    SET @TransNum=@TransNum+1
    FETCH NEXT FROM CUR1 INTO @DocDate,@ItemCode,@WhsCode,@Quantity,@Amount,@OutQty,@OutAmt
END

CLOSE CUR1
DEALLOCATE CUR1
*/


INSERT PS_MM211
SELECT t0.ItemCode,t1.ItemName,0,0,@StrDate,NULL,NULL
      ,NULL,NULL,NULL,NULL,N'전기이월',0
      ,0,0,0,SUM(t0.InQty-t0.OutQty),SUM(isnull(t0.TransValue,0))
      ,NULL,NULL,NULL
      --,t1.U_ShotName,t1.U_Material,t1.U_Size
  FROM #ZOINM t0 JOIN OITM t1 ON t1.ItemCode = t0.ItemCode
 WHERE t0.DocDate < @StrDate
   AND t0.ItemCode BETWEEN @SItemCode AND @EItemCode
 GROUP BY t0.ItemCode,t1.ItemName--,t1.U_ShotName,t1.U_Material,t1.U_Size

 UNION ALL --------------------------------------------

SELECT t0.ItemCode,t1.ItemName,NULL,t0.TransNum,t0.DocDate,t0.CardCode,t0.CardName
      ,t0.TransType,t0.CreatedBy,t0.DocLineNum,t0.Warehouse,t0.JrnlMemo,t0.InQty
      ,t0.OutQty,t0.CalcPrice,t0.TransValue,(t0.InQty-t0.OutQty),t0.TransValue
      ,t0.CreateDate,t0.DocTime,t0.UserSign
      --,t1.U_ShotName,t1.U_Material,t1.U_Size
  FROM #ZOINM t0 JOIN OITM t1 ON t1.ItemCode = t0.ItemCode
 WHERE t0.DocDate BETWEEN @StrDate AND @EndDate
   AND t0.ItemCode BETWEEN @SItemCode AND @EItemCode
   AND (t0.InQty <> 0 or t0.OutQty <> 0 or t0.TransValue <> 0)

----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT * FROM PS_MM211 ORDER BY ItemCode,DocDate,TransNum
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

DECLARE      --@ItemCode       nvarchar(20),
             --@TransNum       int,
             @StockQty       numeric(19,6),
             @StockAmt       numeric(19,6),

             @_SItemCod       nvarchar(20),
             @_SStockQt       numeric(19,6),
             @_SStockAm       numeric(19,6),

             @_VisOrder       int

SET @_SItemCod = ''
/* Calculate stock value */
DECLARE CUR1 CURSOR LOCAL FOR
 SELECT ItemCode,TransNum,StockQty,StockAmt
   FROM PS_MM211 t0
  ORDER BY ItemCode,DocDate,TransNum

OPEN CUR1
FETCH NEXT FROM CUR1 INTO @ItemCode,@TransNum,@StockQty,@StockAmt
WHILE (@@FETCH_STATUS = 0) BEGIN
    IF @_SItemCod <> @ItemCode BEGIN
        SET @_SItemCod = @ItemCode
        SET @_SStockQt = @StockQty
        SET @_SStockAm = @StockAmt
        SET @_VisOrder = 0
    END
    ELSE BEGIN
        SET @_SStockQt = @_SStockQt + @StockQty
        SET @_SStockAm = @_SStockAm + @StockAmt
        SET @_VisOrder = @_VisOrder + 1
        UPDATE PS_MM211 SET VisOrder=@_VisOrder,StockQty=@_SStockQt,StockAmt=@_SStockAm
         WHERE ItemCode = @ItemCode AND TransNum = @TransNum
    END
    FETCH NEXT FROM CUR1 INTO @ItemCode,@TransNum,@StockQty,@StockAmt
END
CLOSE CUR1
DEALLOCATE CUR1


----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
/* DATA RETURN */
SELECT	ItemCode,
        Itemname,
        VisOrder,
        TransNum,			--Key
        DocDate,
		CardCode,
		CardName,
        TransType,		--Type
        CreatedBy,		--Type에 따른 문서 DocEntry
		DocLineNum,		--Type에 따른 문서 Line 번호
		case when TransNum = 0 or TransNum = 999999999 then JrnlMemo
             when TransType in (931,932) then '수입부대비용'
			 else convert(nvarchar, TransType) + '  ' + convert(nvarchar, CreatedBy) + '-' + convert(nvarchar, DocLineNum) end as TypeEntryLine,
        Warehouse,
        JrnlMemo,
        InQty,
        OutQty,
        Price,
        TransValue,
        StockQty,
        StockAmt,
        CreateDate,
        CreateTime,
        UserSign
        --U_ShotName,
        --U_Material,
        --U_Size

FROM PS_MM211
ORDER BY ItemCode,VisOrder
----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_MM211_01] '1', 'ZZZZZZZZ', '20100101', '20101130'