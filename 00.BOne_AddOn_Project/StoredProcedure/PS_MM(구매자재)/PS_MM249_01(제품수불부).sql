USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM249_01]    Script Date: 02/25/2011 20:06:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************************************************************************************************************/
/*  Module         : 구매관리																				    */
/*  Description    : 제품수불부    																				*/
/*  ALTER  Date    : 2011.02.24  																				*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER  PROC [dbo].[PS_MM249_01]
--CREATE     PROC [dbo].[PS_MM249_01]
(
  @SItemCode		as nvarchar(20),
  --@EItemCode		as nvarchar(20),
  @StrDate			as datetime,
  @EndDate			as datetime,
  @ItmGrp			as nvarchar(10), --품목그룹
  @ItmBsort			as nvarchar(10), --대분류
  @ItmMsort			as nvarchar(10), --중분류
  @BPLId			as nvarchar(5),	 --사업장
  @ItemType			as nvarchar(5),  --형태타입
  @Mark				as nvarchar(5)   --인증기호
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
DROP   TABLE PS_MM249 
CREATE TABLE PS_MM249 
			(
             ItemCode        nvarchar(20),
             Itemname        nvarchar(100),
             VisOrder        int,
             TransNum        int,			--Key
             DocDate         datetime,
		
			 CardCode		 Nvarchar(10),
			 CardName		 Nvarchar(50),
             --TransType       smallint,	--Type
             TransType       int,			--Type
             CreatedBy       int,			--Type에 따른 문서 DocEntry
			 DocLineNum		 int,			--Type에 따른 문서 Line 번호
             Warehouse       nvarchar(08),
             JrnlMemo        nvarchar(50),
             InQty           numeric(19,6),
             OutQty          numeric(19,6),
             --Price           numeric(19,6),
             --TransValue      numeric(19,6),
             StockQty        numeric(19,6),
             --StockAmt        numeric(19,6),

             CreateDate      datetime,
             CreateTime      smallint,
             UserSign        smallint,
             LastLine		 nvarchar(1),
             LastLineYYYYMM	 nvarchar(1)
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
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
select a.*
INTO #ZOINM
from OINM a left join OITM b on a.ItemCode=b.ItemCode
where a.DocDate <= @EndDate
--AND ItemCode BETWEEN @SItemCode AND @EItemCode
and b.ItmsGrpCod like @ItmGrp + '%'
and b.U_ItmBsort like @ItmBsort + '%'
and b.U_ItmMsort like @ItmMsort + '%'
and a.ItemCode like @SItemCode + '%'
and a.Warehouse like @BPLId + '%'
and b.U_ItemType like @ItemType + '%'
and b.U_Mark like @Mark + '%'
AND isnull(a.ApplObj,0) <> '911'
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------   
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------   
INSERT PS_MM249
SELECT t0.ItemCode,t1.ItemName,0,0,@StrDate,NULL,NULL
      ,NULL,NULL,NULL,NULL,N'전기이월'
      ,0,0
      --,0,0
      ,SUM(t0.InQty-t0.OutQty)
      --,SUM(isnull(t0.TransValue,0))
      ,NULL,NULL,NULL,NULL,NULL
  FROM #ZOINM t0 JOIN OITM t1 ON t1.ItemCode = t0.ItemCode
 WHERE t0.DocDate < @StrDate
   --AND t0.ItemCode BETWEEN @SItemCode AND @EItemCode
   AND t0.ItemCode like @SItemCode
 GROUP BY t0.ItemCode,t1.ItemName--,t1.U_ShotName,t1.U_Material,t1.U_Size

UNION ALL ------------------------------------------------------------------------

SELECT t0.ItemCode,t1.ItemName,NULL,t0.TransNum,t0.DocDate,t0.CardCode,t0.CardName
      ,t0.TransType,t0.CreatedBy,t0.DocLineNum,t0.Warehouse,t0.JrnlMemo
      ,t0.InQty,t0.OutQty
      --,t0.CalcPrice,t0.TransValue
      ,(t0.InQty-t0.OutQty)
      --,t0.TransValue
      ,t0.CreateDate,t0.DocTime,t0.UserSign,NULL,NULL
  FROM #ZOINM t0 JOIN OITM t1 ON t1.ItemCode = t0.ItemCode
 WHERE t0.DocDate BETWEEN @StrDate AND @EndDate
   --AND t0.ItemCode BETWEEN @SItemCode AND @EItemCode
   AND t0.ItemCode like @SItemCode
   AND (t0.InQty <> 0 or t0.OutQty <> 0)-- or t0.TransValue <> 0)

----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT * FROM PS_MM249 ORDER BY ItemCode,DocDate,TransNum
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
DECLARE      --@ItemCode       nvarchar(20),
             --@TransNum       int,
             @StockQty       numeric(19,6),
             --@StockAmt       numeric(19,6),

             @_SItemCod       nvarchar(20),
             @_SStockQt       numeric(19,6),
             --@_SStockAm       numeric(19,6),

             @_VisOrder       int

SET @_SItemCod = ''
/* Calculate stock value */
DECLARE CUR1 CURSOR LOCAL FOR
 SELECT ItemCode,TransNum,StockQty--,StockAmt
   FROM PS_MM249 t0
  ORDER BY ItemCode,DocDate,TransNum

OPEN CUR1
FETCH NEXT FROM CUR1 INTO @ItemCode,@TransNum,@StockQty--,@StockAmt
WHILE (@@FETCH_STATUS = 0) BEGIN
    IF @_SItemCod <> @ItemCode BEGIN
        SET @_SItemCod = @ItemCode
        SET @_SStockQt = @StockQty
        --SET @_SStockAm = @StockAmt
        SET @_VisOrder = 0
    END
    ELSE BEGIN
        SET @_SStockQt = @_SStockQt + @StockQty
        --SET @_SStockAm = @_SStockAm + @StockAmt
        SET @_VisOrder = @_VisOrder + 1
        UPDATE PS_MM249 SET VisOrder=@_VisOrder,StockQty=@_SStockQt--,StockAmt=@_SStockAm
         WHERE ItemCode = @ItemCode AND TransNum = @TransNum
    END
    FETCH NEXT FROM CUR1 INTO @ItemCode,@TransNum,@StockQty--,@StockAmt
END
CLOSE CUR1
DEALLOCATE CUR1
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------/* DATA UPATE */
/* 출력물의 [계]재고를 표시(Crystal Report에서 사용)하기 위해서*/
update A
set LastLine='Y'
from PS_MM249 A inner join ( select ItemCode,max(VisOrder) VisOrder
							 from PS_MM249
							 group by ItemCode
							) B on A.ItemCode=B.ItemCode and A.VisOrder=B.VisOrder
/* 출력물의 [월별계]재고를 표시(Crystal Report에서 사용)하기 위해서*/
update A
set LastLineYYYYMM='Y'
from PS_MM249 A inner join ( select ItemCode,convert(char(7),DocDate,120) YYYYMM,max(VisOrder) VisOrder
							 from PS_MM249
							 group by ItemCode,convert(char(7),DocDate,120)
							 --order by ItemCode,convert(char(7),DocDate,120)
							) B on A.ItemCode=B.ItemCode and A.VisOrder=B.VisOrder
							
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
/* DATA RETURN */
SELECT	a.ItemCode,
        a.Itemname,
        h.Name ItemType,
        i.Name Mark,
        a.VisOrder,
        a.TransNum,			--Key
        a.DocDate,
        convert(char(7),a.DocDate,120) YYYYMM,
		a.CardCode,
		a.CardName,
        a.TransType,		--Type
        a.CreatedBy,		--Type에 따른 문서 DocEntry
		a.DocLineNum,		--Type에 따른 문서 Line 번호
		case when a.TransNum = 0 or a.TransNum = 999999999 then JrnlMemo
             when a.TransType in (931,932) then '수입부대비용'
			 else convert(nvarchar, a.TransType) + ' ' + convert(nvarchar, a.CreatedBy) + '-' + convert(nvarchar, DocLineNum) end as TypeEntryLine,
        a.Warehouse,
        a.JrnlMemo,
        InQty = Case When a.TransType = 14 Or a.TransType = 16 Then 0 Else a.InQty End,  --반품, AR대변매모는 출고에서 (-)
        InQty2 = (Case When a.TransType = 14 Or a.TransType = 16 Then 0 Else a.InQty End * b.U_UnWeight)/1000,
        OutQty = Case When a.TransType = 14 Or a.TransType = 16 Then a.InQty * -1 Else a.OutQty End, --반품, AR대변매모는 출고에서 (-)
        OutQty2 = (Case When a.TransType = 14 Or a.TransType = 16 Then a.InQty * -1 Else a.OutQty End * b.U_UnWeight)/1000,
        a.StockQty,
        (a.StockQty*b.U_UnWeight)/1000 as StockQty2,
        a.CreateDate,
        a.CreateTime,
        a.UserSign,
        a.LastLine,
        a.LastLineYYYYMM
		,convert(char,b.U_ItmMsort) as ItmMCode
		,convert(char,c.U_CodeName) as ItmMName
		,convert(char,case when a.TransType='59' then f.U_PP070No
			  when a.TransType='60' then g.U_PP070No end) as PackNo  --벌크포장번호
			  
		,case when a.TransType='15' then convert(nvarchar,d.DocEntry)+'-'+convert(nvarchar,d.LineId)
			  when a.TransType='16' then convert(nvarchar,e.DocEntry)+'-'+convert(nvarchar,e.LineId) end as SD040EntryLine     --납품문서번호-라인
		--,convert(char,case when a.TransType='15' then d.LineId
		--	  when a.TransType='16' then e.LineId end) as SD040Line  --납품문서라인
		,case when a.TransType='15' then convert(nvarchar,d.U_ORDRNum)+'-'+convert(nvarchar,d.U_RDR1Num)
			  when a.TransType='16' then convert(nvarchar,e.U_ORDNNum)+'-'+convert(nvarchar,e.U_RDR1Num) end as ORDREntryLine  --판매오더문서번호
		--,convert(char,case when a.TransType='15' then d.U_RDR1Num
		--  	  when a.TransType='16' then e.U_RDR1Num end) as RDR1Num  --판매오더문서라인
			 
			 	
FROM PS_MM249 a left join [OITM] b on a.ItemCode=b.ItemCode
				left join [@PSH_ITMMSORT] c on b.U_ItmMsort=c.U_Code
				left join [@PSH_SHAPE] h on  b.U_ItemType=h.Code
				left join [@PSH_Mark] i on  b.U_Mark=i.Code
				left join [@PS_SD040L] d on a.Createdby=d.U_ODLNNum and a.DocLineNum=d.U_DLN1Num and a.TransType='15' --납품
				left join [@PS_SD040L] e on a.Createdby=e.U_ORDNNum and a.DocLineNum=e.U_RDN1Num and a.TransType='16' --반품
				left join (
							select U_PP070No,U_OIGNNo 
							from [@PS_PP077H]
							group by U_PP070No,U_OIGNNo
						   ) f on a.Createdby=f.U_OIGNNo and a.TransType='59' --포장처리등록(입고)
				left join (
							select U_PP070No,U_OIGENo 
							from [@PS_PP077H]
							group by U_PP070No,U_OIGENo
						   ) g on a.Createdby=g.U_OIGENo and a.TransType='60' --포장처리등록취소(출고)
								
ORDER BY ItemCode,ItemName,b.U_ItmMsort,VisOrder

----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_MM249_01] '%','20110101','20110105','102','101','%','%4','%','%'
--EXEC [PS_MM249_01] '%','20110101','20110131','102','101','%','%','%','%'
--EXEC [PS_MM249_01] '%','20110101','20110105','101','202','%','%4','%','%'
--EXEC [PS_MM249_01] '%','20110101','20110110','102','101','10105','%4','%','%'
--EXEC [PS_MM249_01] '%','20101210','20110110','102','101','10105','%4','%','%'
--EXEC [PS_MM249_01] '101030108','20110101','20110131','102','101','%','%4','%','%'