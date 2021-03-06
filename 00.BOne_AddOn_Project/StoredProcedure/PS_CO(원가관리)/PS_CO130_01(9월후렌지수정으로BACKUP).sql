USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_CO130_01]    Script Date: 10/04/2011 19:09:46 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Procedure ID : PS_CO130_01
-- Author       : Minho Choi
-- Create date  : 2011.01.21
-- Description  : 제품별 원가계산
-- EXEC [PS_CO130_01] '201105' , 1
-- =============================================
ALTER PROCEDURE [dbo].[PS_CO130_01]
    @iYM        AS nvarchar(6),
    @iBPLId     AS int
AS

CREATE TABLE #CO130A (  -- 재공 전기이월
    POEntry int NOT NULL
 CONSTRAINT [PK_#CO130A] PRIMARY KEY CLUSTERED 
    ( POEntry )
)

CREATE TABLE #CO130B (  -- 재공 당기생산
    POEntry int NOT NULL
 CONSTRAINT [PK_#CO130B] PRIMARY KEY CLUSTERED 
    ( POEntry )
)

CREATE TABLE #CO130X (  -- 마지막공정
    POEntry int NOT NULL,
    POLine  int
 CONSTRAINT [PK_#CO130X] PRIMARY KEY CLUSTERED 
    ( POEntry )
)
CREATE TABLE #CO130E (  -- 원가요소
    CECode  nvarchar(15) COLLATE Korean_Wansung_Unicode_CI_AS,
    CEName  nvarchar(100) COLLATE Korean_Wansung_Unicode_CI_AS
)

CREATE TABLE #CO130C (    -- 재공수불명세서
    POEntry  int,
    POLine   int,
    Sequence int,
    ItemCode nvarchar(20)  COLLATE Korean_Wansung_Unicode_CI_AS,
    ItemName nvarchar(100) COLLATE Korean_Wansung_Unicode_CI_AS,
    CpCode   nvarchar(10)  COLLATE Korean_Wansung_Unicode_CI_AS,
    CpName   nvarchar(30)  COLLATE Korean_Wansung_Unicode_CI_AS,
    OrdGbn   nvarchar(3)   COLLATE Korean_Wansung_Unicode_CI_AS,

    COQty    numeric(19,6),
    COAMT    numeric(19,6),
    InQty    numeric(19,6),
    InAMT    numeric(19,6),
    IPreAMT  numeric(19,6),
    DefQty   numeric(19,6),
    
    OutQty   numeric(19,6),
    OutAMT   numeric(19,6),
    StcQty   numeric(19,6),
    StcAMT   numeric(19,6),
    Scrap    numeric(19,6),
    
    NextAmt  numeric(19,6),
    CostAmt  numeric(19,6),
    ScrAmt   numeric(19,6),
 CONSTRAINT [PK_#CO130C] PRIMARY KEY CLUSTERED 
    ( POEntry,POLine )
)

CREATE TABLE #CO130D (    -- 재공수불명세서
    POEntry  int,
    POLine   int,
    CECode   nvarchar(15) COLLATE Korean_Wansung_Unicode_CI_AS,
    CEName   nvarchar(15) COLLATE Korean_Wansung_Unicode_CI_AS,
    COAMT    numeric(19,6),
    InAMT    numeric(19,6),
    IPreAMT  numeric(19,6),
    OutAMT   numeric(19,6),
    StcAMT   numeric(19,6),
    NextAMT  numeric(19,6),
    CostAmt  numeric(19,6),
    ScrAmt   numeric(19,6),
 CONSTRAINT [PK_#CO130D] PRIMARY KEY CLUSTERED 
    ( POEntry,POLine,CECode )
)

DECLARE @BefYM    nvarchar(6),
        @FrDate   datetime,
        @ToDate   datetime,
        @Code     nvarchar(8),
        @BefCode  nvarchar(8),

        @POEntry  int,
        @POLine   int,
        @COQty    numeric(19,6),
        @InQty    numeric(19,6),
        @OutQty   numeric(19,6),
        @DefQty   numeric(19,6),
        @StcQty   numeric(19,6),
        @Scrap    numeric(19,6),
        
        @COAMT    numeric(19,6),
        @InAMT    numeric(19,6),
        @OutAMT   numeric(19,6),
        @DefAMT   numeric(19,6),
        @StcAMT   numeric(19,6),
        @NextAMT  numeric(19,6),
        @CostAMT  numeric(19,6),
        
        @CECode   nvarchar(15),
        @CEName   nvarchar(100),
        @ItmBSort nvarchar(8),
        @ReqAmt   numeric(19,6),
        
        @xPOEntry int,
        @xPOLine  int,
        @xOrdGbn  nvarchar(3),
        @xCOQty   numeric(19,6),
        @xInQty   numeric(19,6),
        @xDefQty  numeric(19,6),
        @xOutQty  numeric(19,6),
        @xScrap   numeric(19,6),
        @xStcQty  numeric(19,6),
        @xCOAmt   numeric(19,6),
        @xInAmt   numeric(19,6),
        @xIPreAmt numeric(19,6),
        @xOutAmt  numeric(19,6),
        @xNextAmt numeric(19,6),
        @xCostAmt numeric(19,6),
        @xInVal   numeric(19,6),
        @xReqVal  numeric(19,6),
        @xLastLn  int,
        
        @CompAmt  numeric(19,6),
        @ScrPrice numeric(19,6),
        @ScrAmt   numeric(19,6),

        @sPOEntry int,
        @sPOLine  int,
        
        @LastCP   char(1)

----------------------------------------------------------------------------------------------
IF @iBPLId = 3 BEGIN 
    EXEC [PS_CO130_31] @iYM, @iBPLId 
    RETURN
END

SELECT @FrDate=F_RefDate,@ToDate=T_RefDate
  FROM OFPR
 WHERE CONVERT(CHAR(6),F_RefDate,112) = @iYM

SELECT @BefYM =  CONVERT(CHAR(6),@FrDate-1,112)

-- 스크랩 단가
--SELECT @ScrPrice = CASE @iYM WHEN '201101' THEN  9940
--                             WHEN '201102' THEN  10460
--                             WHEN '201103' THEN  10460 ELSE 0 END

select @ScrPrice = t2.U_Price
  from [@PS_MM001H] t1 Inner Join [@PS_MM001L] t2 On t1.DocEntry = t2.DocEntry
 where t1.U_DocDate = ( SELECT Max(a.U_DocDate)
						  From [@PS_MM001H] a Inner Join [@PS_MM001L] b On a.DocEntry = b.DocEntry
						 where Convert(char(6),a.U_DocDate,112) = @iYM
						   and b.U_Price > 0 )
	 



SET @Code = @iYM+STR(@iBPLId,1)
SET @BefCode = @BefYM+STR(@iBPLId,1)

-- 제품 수불 자료 생성
EXEC [PS_CO130_02] @iYM,@iBPLId
-- 원가요소 수집
INSERT #CO130E
SELECT DISTINCT CECode,CEName FROM (
SELECT CECode,CEName FROM Z_PS_CO120L WHERE Code = @Code AND Cost <> 0
UNION ALL
SELECT CECode,CEName FROM Z_PS_CO130L WHERE Code = @BefCode AND StcAmt <> 0
) ZZ

-- 원가계산 대상 작지번호 자료 작성
-- 전기재공PO
INSERT #CO130A
SELECT DISTINCT L.U_POEntry
  FROM [@PS_CO130H] H
  JOIN [@PS_CO130L] L ON L.Code = H.Code
 WHERE H.Code = @BefCode AND (L.U_StcQty <> 0 Or L.U_StcAmt <> 0)
-- 당기생산PO
INSERT #CO130B
SELECT DISTINCT L.U_POEntry
  FROM [@PS_CO120H] H
  JOIN [@PS_CO120L] L ON L.Code = H.Code
 WHERE H.Code = @Code
-- 작번별 최종공정 행번호 자료 작성
INSERT #CO130X
SELECT T0.DocEntry,MAX(T1.LineId) 
FROM [@PS_PP030H] T0 
JOIN [@PS_PP030M] T1 ON T1.DocEntry = T0.DocEntry
WHERE T0.Canceled='N'
GROUP BY T0.DocEntry

-- 당기재공
-- 기계공구外
INSERT #CO130C
SELECT ISNULL(T1.U_POEntry,T2.DocEntry)
      ,ISNULL(T1.U_POLine,T2.LineId)
      ,ISNULL(T1.U_Sequence,T2.U_Sequence)
      
      ,ISNULL(T1.U_ItemCode,T2.U_ItemCode)
      ,ISNULL(T1.U_ItemName,T2.U_ItemName)
      ,ISNULL(T1.U_CpCode,T2.U_CpCode)
      ,ISNULL(T1.U_CpName,T2.U_CpName)
      ,ISNULL(T1.U_ItmBSort,T2.U_OrdGbn)
      
      ,ISNULL(T1.U_StcQty,0),ISNULL(T1.U_StcAmt,0)  -- 이월
      ,ISNULL(T2.U_ProdQty,0),0,0 --,ISNULL(T2.U_Cost,0),0 -- 입고
      ,ISNULL(T2.U_DefQty,0)                        -- 불량
      ,0,0,0,0                                      -- 출고,재고
      ,ISNULL(T2.U_Scrap,0)                         -- 스크랩
      ,0,0,0                                        -- 기타금액
  FROM (SELECT L.U_POEntry,L.U_POLine,L.U_Sequence,L.U_ItemCode,L.U_ItemName,L.U_CpCode,L.U_CpName
              ,L.U_StcQty,L.U_StcAmt,I.U_ItmBSort
          FROM #CO130A      A
          JOIN [@PS_CO130L] L ON L.U_POEntry = A.POEntry
          JOIN [@PS_CO130H] H ON H.Code = L.Code AND H.Code = @BefCode    ---- 전월 이월 자료
          JOIN OITM I ON I.ItemCode = L.U_ItemCode AND I.U_ItmBSort NOT IN ('105','106')
  ) T1 FULL JOIN 
       (
        SELECT H.DocEntry,M.LineId,M.U_Sequence,H.U_ItemCode,H.U_ItemName,M.U_CpCode,M.U_CpName
              ,L.U_ProdQty,L.U_DefQty,L.U_Cost,L.U_Scrap,H.U_OrdGbn
          FROM #CO130B      B
          JOIN [@PS_PP030H] H ON H.DocEntry = B.POEntry AND H.U_OrdGbn NOT IN ('105','106') --H.Canceled = 'N' 취소도 포함 한다
     LEFT JOIN [@PS_PP030M] M ON M.DocEntry = H.DocEntry
     LEFT JOIN [@PS_CO120L] L ON L.U_POEntry = M.DocEntry AND L.U_POLine = M.LineId AND L.Code = @Code
         WHERE H.U_ItemCode NOT LIKE '10107%' -- FLANGE
         UNION ALL
        SELECT L.U_POEntry,L.U_POLine,0,L.U_ItemCode,L.U_ItemName,L.U_CpCode,L.U_CpName
              ,L.U_ProdQty,L.U_DefQty,L.U_Cost,0,'101'
          FROM #CO130B      B
          JOIN [@PS_CO120L] L ON L.U_POEntry = B.POEntry AND L.Code = @Code
         WHERE L.U_ItemCode LIKE '10107%'     -- FLANGE
  ) T2 ON T2.DocEntry = T1.U_POEntry AND T2.LineId = T1.U_POLine




--Update #CO130C
--   set InQty = t.Quantity
--  From [@PS_PP030H] PP030H,
--	   (Select a.BatchNum,
--				Sum(a.Quantity * (Case When a.Direction = '0' Then 1 Else -1 End)) As Quantity
--				 from IBT1 a Inner Join OITM b On a.ItemCode = b.ItemCode And b.U_ItmBsort = '104' --멀티
--				 Where a.DocDate between @FrDate and @ToDate
--				   and a.BaseType In ('59', '60')
--				  Group by a.BatchNum
--				  Having Sum(a.Quantity * (Case When a.Direction = '0' Then 1 Else -1 End)) > 0 ) t
-- Where PP030H.U_OrdNum = t.BatchNum
--   And PP030H.U_OrdGbn = '104'
--   And PP030H.DocEntry = #CO130C.PoEntry
--   And #CO130C.POLine = (Select Max(a.PoLine)
--						   From #CO130C a
--						  Where #CO130C.PoEntry = a.PoEntry)

-- 멀티 최종 완료 중량을 입고 중량으로...
UPDATE t0 SET InQty = t1.Qty	
FROM #CO130C t0
JOIN (SELECT U_PP030Hno,U_PP030MNo,SUM(b.U_YQty) Qty
      FROM [@PS_PP080H] a 
      JOIN [@PS_PP080L] b on b.DocEntry = a.DocEntry
      WHERE a.Canceled = 'N' AND a.U_DocDate BETWEEN @FrDate and @ToDate
      AND b.U_OrdGbn = '104'
      And Isnull(b.U_OIGENum,'') = ''
      GROUP BY U_PP030Hno,U_PP030MNo) t1 ON t1.U_PP030Hno = t0.POEntry AND t1.U_PP030MNo = t0.POLine

-- 입/출/재고 수량 계산
SET @sPOEntry = 0
DECLARE CUR1 CURSOR FOR
SELECT T0.POENtry,T0.POLine,COQty,InQty,DefQty,Scrap,U_ItmBSort,CASE WHEN T2.POLine IS NULL THEN 'N' ELSE 'Y' END
  FROM #CO130C T0
  JOIN OITM    T1 ON T1.ItemCode = T0.ItemCode
  LEFT JOIN #CO130X T2 ON T2.POEntry = T0.POEntry AND T2.POLine = T0.POLine
 WHERE ISNULL(T1.U_ItmBSort,'') NOT IN ('105','106')  --- 기계공구 
 ORDER BY POENtry,Sequence

OPEN  CUR1
FETCH NEXT FROM CUR1 INTO @POEntry,@POLine,@COQty,@InQty,@DefQty,@Scrap,@ItmBSort,@LastCP
WHILE	@@FETCH_STATUS = 0 BEGIN
    -- 입고수량
    IF @POENtry <> @sPOENtry BEGIN
        SET @sPOENtry = @POENtry
    END
    ELSE BEGIN
        SET @OutQty = @InQty + @DefQty + @Scrap
        
        UPDATE #CO130C SET OutQty = @OutQty WHERE POEntry = @sPOEntry AND POLine = @sPOLine
    END

    SET @sPOLine = @POLine
FETCH NEXT FROM CUR1 INTO @POEntry,@POLine,@COQty,@InQty,@DefQty,@Scrap,@ItmBSort,@LastCP
END

CLOSE CUR1
DEALLOCATE CUR1

-- 생산완료등록
UPDATE t0 SET OutQty = t1.Qty
FROM #CO130C t0
JOIN (SELECT U_PP030Hno,U_PP030MNo,SUM(b.U_YQty) Qty
      FROM [@PS_PP080H] a 
      JOIN [@PS_PP080L] b on b.DocEntry = a.DocEntry
      WHERE a.Canceled = 'N' AND a.U_DocDate BETWEEN @FrDate and @ToDate
      AND b.U_OrdGbn NOT IN ('105','106','101') 
      And Isnull(b.U_OIGENum,'') = ''
      GROUP BY U_PP030Hno,U_PP030MNo) t1 ON t1.U_PP030Hno = t0.POEntry AND t1.U_PP030MNo = t0.POLine

-- 휘팅포장완료등록
UPDATE t0 SET OutQty = t1.U_YWeight
FROM #CO130C t0 JOIN (
SELECT b.U_PP030Hno,b.U_PP030MNo,SUM(U_YWeight)U_YWeight
FROM [@PS_PP040H] a
JOIN [@PS_PP040L] b ON b.DocEntry = a.DocEntry AND b.U_CpCode='CP30114'  -- 휘팅 포장
WHERE a.Canceled='N' and a.U_DocDate BETWEEN @FrDate AND @ToDate
GROUP BY b.U_PP030Hno,b.U_PP030MNo
) t1 ON t1.U_PP030Hno = t0.POEntry AND t1.U_PP030MNo = t0.POLine

If @iBPLId = '1'
begin
--//부품 
Delete ZCO130P2 Where YM = @iYM

Insert Into ZCO130P2
Select @iYM,
	   a.POEntry,
	   a.POLine,
	   a.Sequence,
	   a.ItemCode,
	   a.ItemName,
	   a.CpCode,
	   a.CpName,
	   Sum(Isnull(a.COQTy,0)),
	   Sum(Isnull(a.COAmt,0)),
	   Sum(Isnull(a.InQty,0)),
	   Sum(Isnull(a.InAmt,0)),
	   Sum(Isnull(a.IPreAMT,0)),
	   Sum(Isnull(a.DefQty,0)),
	   Sum(Isnull(a.OutQty,0)),
	   Sum(Isnull(a.OutAmt,0)),
	   Sum(Isnull(a.CostAmt,0)),
	   Sum(Isnull(a.StcQty,0)),
	   Sum(Isnull(a.StcAmt,0)),
	   Sum(Isnull(a.NextAmt,0)),
	   ScrQTy = 0,
	   Sum(Isnull(a.ScrAmt,0))
  From #CO130C a 
  Where Not Exists ( Select * From ZCO130P1 b Where a.ItemCode = b.ItemCode )--And b.ym = @iYM )
    And a.OrdGbn = '102' --부품제품만
 Group by a.POEntry,
	   a.POLine,
	   a.Sequence,
	   a.ItemCode,
	   a.ItemName,
	   a.CpCode,
	   a.CpName
	   

Insert Into ZCO130P2
Select @iYM,
	   b.POEntry,
	   a.POLine,
	   a.Sequence,
	   a.ItemCode,
	   a.ItemName,
	   a.CpCode,
	   a.CpName,
	   Sum(Isnull(a.COQTy,0)),
	   Sum(Isnull(a.COAmt,0)),
	   Sum(Isnull(a.InQty,0)),
	   Sum(Isnull(a.InAmt,0)),
	   Sum(Isnull(a.IPreAMT,0)),
	   Sum(Isnull(a.DefQty,0)),
	   Sum(Isnull(a.OutQty,0)),
	   Sum(Isnull(a.OutAmt,0)),
	   Sum(Isnull(a.CostAmt,0)),
	   Sum(Isnull(a.StcQty,0)),
	   Sum(Isnull(a.StcAmt,0)),
	   Sum(Isnull(a.NextAmt,0)),
	   ScrQTy = 0,
	   Sum(Isnull(a.ScrAmt,0))
  From #CO130C a Inner Join ZCO130P1 b On a.ItemCode = b.ItemCode
 Where a.OrdGbn = '102'
   --And b.ym = @iYM
 Group by b.POEntry,
	   a.POLine,
	   a.Sequence,
	   a.ItemCode,
	   a.ItemName,
	   a.CpCode,
	   a.CpName
	   

Delete FROM #CO130C
WHERE OrdGbn = '102'


Insert Into #CO130C
Select POEntry,
	   POLine,
	   Sequence,
	   ItemCode,
	   ItemName,
	   CpCode,
	   CpName,
	   '102',
	   COQTy,
	   COAmt,
	   InQty,
	   InAmt,
	   IPreAMT,
	   DefQty,
	   OutQty,
	   OutAmt,
	   StcQty,
	   StcAmt,
	   Scrap = 0,
	   NextAmt,
	   CostAmt,
	   ScrAmt
  From ZCO130P2
 Where ym = @iYM
   
End



-- 부품통합
--DELETE C
--FROM #CO130C C JOIN ZCO130P1 Z ON Z.YM = @iYM AND Z.ItemCode = C.ItemCode
--WHERE C.OrdGbn = '102' AND C.POEntry <> Z.POEntry

--UPDATE C SET InQty=A.InQTy,DefQty=A.DefQty,OutQty=A.OutQTy,Scrap=A.ScrQty
--FROM #CO130C C
--JOIN ZCO130P2 A ON A.YM = @iYM AND A.POEntry = C.POEntry AND A.POLine = C.POLine

-- 재고수량
UPDATE #CO130C SET StcQty = COQty + InQty - OutQty WHERE OrdGbn NOT IN ('105','106')
-- M/G 이월수량 오류정정
UPDATE T0 SET COQty=COQty-T0.StcQty,StcQty=0
FROM #CO130C T0
--JOIN #CO130X T1 ON T1.POEntry = T0.POEntry
WHERE @iYM = '201101' 
AND T0.OrdGbn = '104' AND T0.COQty <> 0 AND T0.StcQty <> 0 AND T0.OutQty <> 0

-- 기계공구
SELECT ISNULL(C.U_POEntry,D.U_POEntry) POEntry
      ,ISNULL(C.U_POLine,D.U_POLine) POLine
      ,ISNULL(C.U_StcAmt,0) StcAmt
      ,ISNULL(D.U_Cost,0) Cost
  INTO #CO130T
  FROM (SELECT U_POEntry,U_POLine,U_StcAmt FROM [@PS_CO130L] WHERE Code = @BefCode) C
  FULL JOIN (SELECT U_POEntry,U_POLine,U_Cost FROM [@PS_CO120L] WHERE Code = @Code) D
       ON D.U_POEntry = C.U_POEntry AND D.U_POLine = C.U_POLine

INSERT #CO130C
SELECT T.POEntry
      ,T.POLine
      ,M.U_Sequence
      
      ,H.U_ItemCode
      ,H.U_ItemName
      ,M.U_CpCode
      ,M.U_CpName
      ,H.U_OrdGbn
      
      ,ISNULL(A.COQty,0),ISNULL(T.StcAmt,0)  -- 이월
      ,ISNULL(A.InQty,0),ISNULL(T.Cost,0),0  -- 입고(생산;합격량)
      ,0                                     -- 불량
      ,ISNULL(A.OutQty,0),0                  -- 출고
      ,ISNULL(A.StcQty,0)StcQty,0,0,0,0,0    -- 재고
  FROM #CO130T T
  JOIN [@PS_PP030H] H ON H.DocEntry = T.POEntry -- H.Canceled = 'N' AND 취소작번 포함
  LEFT JOIN [@PS_PP030M] M ON M.DocEntry = T.POEntry AND M.LineId = T.POLine
  LEFT JOIN Z_PS_CO130C A ON A.YYMM = CONVERT(CHAR(6),@FrDate,112) 
                          AND A.BPLId = H.U_BPLId AND A.ItemCode =  H.U_ItemCode
 WHERE H.U_OrdGbn IN ('105','106')

-- 포장사업팀
INSERT #CO130C
SELECT T3.ItemCode,0,0,T3.ItemCode,T3.ItemName,'','',T3.U_ItmBSort
,ISNULL(T2.U_StcQty,0),ISNULL(T2.U_StcAmt,0),0,T0.U_Cost,0,0
,T1.Qty,0,0,0,0
,0,0,0
FROM [@PS_CO120L] T0
JOIN OITM T3 ON T3.ItemCode = T0.U_ItemCode
JOIN (SELECT T0.ItemCode,SUM(T0.InQty-T0.OutQty)Qty
      FROM OIVL T0
      JOIN OITM T3 ON T3.ItemCode = T0.ItemCode
      JOIN OWHS T4 ON T4.WhsCode = T0.LocCode
      WHERE T0.DocDate BETWEEN @FrDate AND @ToDate
      AND T4.BPLId = 3 AND @iBPLId = 3
      AND T3.U_ItmBSort IN (108,109,110)
      AND T0.TransType IN (59,60)
      GROUP BY T0.ItemCode) T1 ON T1.ItemCode = T3.ItemCode
LEFT JOIN [@PS_CO130L] T2 ON T2.Code = @BefCode AND T2.U_ItemCode = T0.U_ItemCode
WHERE T0.Code = @Code
AND T3.U_ItmBSort IN (108,109,110)

---------=---------=---------=---------=---------=---------=---------=---------=

SELECT ISNULL(C.POEntry,A.POEntry)POEntry,A.POLine,A.CECode,SUM(Cost)Cost
INTO #CO120L
FROM [Z_PS_CO120L] A
LEFT JOIN [@PS_PP030H] B ON B.DocEntry = A.POEntry
LEFT JOIN ZCO130P1 C ON C.ItemCode = B.U_ItemCode --C.YM = @iYM AND 
WHERE A.Code = @Code 
GROUP BY ISNULL(C.POEntry,A.POEntry),A.POLine,A.CECode

SELECT ISNULL(T2.POEntry,T1.POEntry)POEntry,T1.POLine,T1.CECode,SUM(T1.StcAmt)StcAmt
INTO #CO130L
FROM [@PS_CO130L] T0
JOIN [Z_PS_CO130L] T1 ON T1.Code = T0.Code AND T1.POEntry = T0.U_POEntry AND T1.POLine = T0.U_POLine
LEFT JOIN ZCO130P1 T2 ON T2.ItemCode = T0.U_ItemCode --T2.YM = @iYM AND 
WHERE T0.Code = @BefCode
GROUP BY ISNULL(T2.POEntry,T1.POEntry),T1.POLine,T1.CECode

-- 입/출/재고 금액 계산
-- 기계공구,몰드 포장외
INSERT #CO130D
SELECT ISNULL(T3.POEntry,T0.POEntry),T0.POLine,T0.CECode,T0.CEName
      ,SUM(ISNULL(T1.StcAmt,0)),SUM(ISNULL(T2.Cost,0))
      ,0,0,0,0,0,0
  FROM (SELECT * FROM #CO130C,#CO130E) T0 
  LEFT JOIN #CO130L T1 ON T1.POEntry = T0.POEntry AND T1.POLine = T0.POLine AND T1.CECode = T0.CECode
  LEFT JOIN #CO120L T2 ON T2.POEntry = T0.POEntry AND T2.POLine = T0.POLine AND T2.CECode = T0.CECode
  LEFT JOIN ZCO130P1 T3 ON T3.ItemCode = T0.ItemCode  -- 부품통합용  --T3.YM = @iYM AND 
 WHERE T0.OrdGbn NOT IN ('105','106','108','109','110')
 GROUP BY ISNULL(T3.POEntry,T0.POEntry),T0.POLine,T0.CECode,T0.CEName






DECLARE CUR2 CURSOR FOR
SELECT CECode,CEName
  FROM #CO130E
 ORDER BY CECode

OPEN  CUR2
FETCH NEXT FROM CUR2 INTO @CECode,@CEName
WHILE	@@FETCH_STATUS = 0 BEGIN
    SET @sPOEntry = 0
    
    DECLARE CUR3 CURSOR FOR
    SELECT C.POEntry,C.POLine,C.OrdGbn,C.COQty,C.InQty,C.DefQty,C.OutQty,C.Scrap,D.COAmt,D.InAmt,L.POLine
      FROM #CO130C C
      JOIN #CO130D D ON D.POEntry = C.POEntry AND D.POLine = C.POLine
      JOIN #CO130X L ON L.POEntry = C.POENtry
     WHERE D.CECode = @CECode
     ORDER BY C.POEntry,C.POLine,C.Sequence

    OPEN  CUR3
    FETCH NEXT FROM CUR3 INTO @xPOEntry,@xPOLine,@xOrdGbn,@xCOQty,@xInQty,@xDefQty,@xOutQty,@xScrap,@xCOAmt,@xInAmt,@xLastLn
    WHILE @@FETCH_STATUS = 0 BEGIN
        IF @xPOEntry <> @sPOEntry BEGIN
            SET @sPOEntry = @xPOEntry
            SET @xIPreAmt = 0
            END
        ELSE
            SET @xIPreAmt = @xNextAmt

        IF @CECode = '55101010' AND @xOrdGbn = '101' BEGIN-- 재료비 스크랩 차감
            SET @ScrAmt = ROUND(@xScrap * @ScrPrice,0)
            --SET @xInAmt = @xInAmt - @ScrAmt
        END
        ELSE 
            SET @ScrAmt = 0

        IF (@xDefQty + @xOutQty + @xScrap) = 0 
            SET @CompAmt = 0
        ELSE
            SET @CompAmt = CASE WHEN (@xCOQty+@xInQty) = 0 THEN 0 
                                --ELSE ROUND(@xOutQty*(@xCOAmt+@xInAmt+@xIPreAmt)/(@xCOQty+@xInQty),0) END
                                ELSE ROUND(@xOutQty*(@xCOAmt+@xInAmt+@xIPreAmt-@ScrAmt)/(@xCOQty+@xInQty),0) END

        IF (@xOrdGbn = '104' AND @xDefQty <> 0 AND @xOutQty = 0) OR @xStcQty = 0 BEGIN  --멀티 불량은 전금액 매출원가로
            SET @xOutAmt  = 0
            SET @xNextAmt = 0
            SET @xCostAmt = @xCOAmt+@xInAmt+@xIPreAmt    
        END
        ELSE IF @xPOLine = @xLastLn BEGIN
            SET @xOutAmt  = @CompAmt
            SET @xNextAmt = 0
            SET @xCostAmt = 0
        END
        ELSE BEGIN
            SET @xOutAmt  = 0
            SET @xNextAmt = @CompAmt
            SET @xCostAmt = 0
        END

        UPDATE #CO130D SET IPreAmt=@xIPreAmt,OutAmt=@xOutAmt,NextAmt=@xNextAmt,CostAmt=@xCostAmt,ScrAmt=@ScrAmt
         WHERE POEntry=@xPOEntry AND POLine=@xPOLine AND CECode=@CECode

        UPDATE #CO130C SET InAmt=InAmt+@xInAmt,IPreAmt=IPreAmt+@xIPreAmt,OutAmt=OutAmt+@xOutAmt
                          ,NextAmt=NextAmt+@xNextAmt,CostAmt=CostAmt+@xCostAmt,ScrAmt=ScrAmt+@ScrAmt
         WHERE POEntry=@xPOEntry AND POLine=@xPOLine

    FETCH NEXT FROM CUR3 INTO @xPOEntry,@xPOLine,@xOrdGbn,@xCOQty,@xInQty,@xDefQty,@xOutQty,@xScrap,@xCOAmt,@xInAmt,@xLastLn
    END

    CLOSE CUR3
    DEALLOCATE CUR3

FETCH NEXT FROM CUR2 INTO @CECode,@CEName
END
CLOSE CUR2
DEALLOCATE CUR2



-- 기계공구,몰드 & 포장
INSERT #CO130D
SELECT T0.POEntry
      ,T0.POLine
      ,T0.CECode
      ,T0.CEName
      ,ISNULL(T1.StcAmt,0)
      ,ISNULL(T2.Cost,0)
      ,0,0,0,0,0,0
  FROM (SELECT * FROM #CO130C,#CO130E) T0 
  LEFT JOIN [Z_PS_CO130L] T1 ON T1.Code = @BefCode AND T1.POEntry = T0.POEntry AND T1.POLine = T0.POLine AND T1.CECode = T0.CECode
  LEFT JOIN [Z_PS_CO120L] T2 ON T2.Code = @Code AND T2.POEntry = T0.POEntry AND T2.POLine = T0.POLine AND T2.CECode = T0.CECode
 WHERE T0.OrdGbn IN ('105','106','108','109','110')

DECLARE CUR4 CURSOR FOR
SELECT a.POEntry,a.POLine,a.CECODE,a.COAMT,a.InAMT,b.OutQty,b.StcQty
  FROM #CO130D a
  JOIN #CO130C b on b.POEntry = a.POEntry and b.POLine = a.POLine
 WHERE b.OrdGbn IN ('105','106','108','109','110')
 ORDER BY a.POEntry,a.POLine

OPEN  CUR4
FETCH NEXT FROM CUR4 INTO @xPOEntry,@xPOLine,@CECode,@xCOAMT,@xInAMT,@xOutQty,@xStcQty
WHILE	@@FETCH_STATUS = 0 BEGIN
    SET @OutAmt = 0
    SET @CostAmt = 0
    IF @xOutQty <> 0 AND @xStcQty = 0 SET @OutAmt = @xCoAmt + @xInAmt
    IF @xOutQty <> 0 AND @xStcQty <> 0 BEGIN
        SET @xInVal = 0   SET @xReqVal = 0
        SELECT @xInVal=InVal,@xReqVal=ReqVal
          FROM Z_PS_CO130B
         WHERE YYMM = CONVERT(CHAR(6),@FrDate,112)
           AND POEntry = @xPOEntry
           AND POLine = @xPOLine
        SET @ReqAmt = CASE WHEN @xInVal = 0 THEN 0 ELSE (@xCOAmt + @xInAmt) * @xReqVal  / @xInVal END
        SET @OutAmt = ROUND((@xCOAmt + @xInAmt + @ReqAmt) * @xOutQty / (@xOutQty+@xStcQty),0)
    END

    IF @xStcQty = 0 AND (@xCOAMT+@xInAMT-@OutAmt) <> 0 BEGIN
        SET @CostAmt = @OutAmt + (@xCOAMT+@xInAMT-@OutAmt)
        SET @OutAmt = 0
    END

    UPDATE #CO130D SET OutAmt=@OutAmt,CostAmt=@CostAmt
     WHERE POEntry=@xPOEntry AND POLine=@xPOLine AND CECode=@CECode

    UPDATE #CO130C SET OutAmt=OutAmt+@OutAmt,CostAmt=CostAmt+@CostAmt
     WHERE POEntry=@xPOEntry AND POLine=@xPOLine

FETCH NEXT FROM CUR4 INTO @xPOEntry,@xPOLine,@CECODE,@xCOAMT,@xInAMT,@xOutQty,@xStcQty
END
CLOSE CUR4
DEALLOCATE CUR4

-- 재고 금액
UPDATE #CO130D SET StcAmt = COAmt + InAmt + IPreAmt - OutAmt - NextAmt - CostAmt - ScrAmt
UPDATE #CO130C SET StcAmt = COAmt + InAmt + IPreAmt - OutAmt - NextAmt - CostAmt - ScrAmt

---------=---------=---------=---------=---------=---------=---------=---------=
DELETE Z_PS_CO130L WHERE Code = @Code
INSERT Z_PS_CO130L
SELECT @Code,* FROM #CO130D

--최종데이터조회
--'이월수량'부터 '스크랩금액'까지 모두 0인 행은 출력하지 않음(2011.09.05 송명규 수정, 최수환 이사님 요청)
--이상 발생 시 WHERE문 주석 처리
SELECT		* 
FROM			#CO130C

WHERE		COQty <> 0 --이월수량
				OR COAMT <> 0 --이월금액
				OR InQty <> 0 --입고수량
				OR InAMT <> 0 --입고금액
				OR IPreAMT <> 0 --전공정금액
				OR DefQty <> 0 --불량수량
				OR OutQty <> 0 --출고수량
				OR OutAMT <> 0 --출고금액
				OR CostAMT <> 0 --매출원가
				OR NextAMT <> 0 --다음공정
				OR StcQty <> 0 --재고수량
				OR StcAMT <> 0 --재고금액
				OR Scrap <> 0 --스크랩
				OR ScrAMT <> 0 --스크랩금액
ORDER BY	ItemCode,
				POEntry,
				POLine
			/*	
--select sum(InAmt) from #CO130D
--Where POEntry > 1000000
*/

--select *--sum(InAmt)
-- from #CO130C
--Where POEntry > 1000000
--EXEC [PS_CO130_01] '201108' , 1