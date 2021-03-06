USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_CO120_03]    Script Date: 09/24/2013 13:34:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Procedure ID : PS_CO120_03
-- Author       : Minho Choi
-- Create date  : 2011.01.21
-- Description  : 외주가공비
-- EXEC PS_CO120_03 '201104',1
-- =============================================
ALTER PROCEDURE [dbo].[PS_CO120_03]
    @iYM        AS nvarchar(6),
    @iBPLId     AS int
AS

DECLARE		@FrDate   datetime,
				@ToDate   datetime,
				@ItemCode nvarchar(20),
				@Amount   numeric(19,6),
				@SAmt     numeric(19,6),
				@RAmt     numeric(19,6),
				@Quantity numeric(19,6),
				@tQty     numeric(19,6),
				@POEntry  int,
				@cAmt     numeric(19,6),
				@AcctCode nvarchar(15),
				@AcctName nvarchar(100),
				@Code     nvarchar(8),
				@ItmMSort nvarchar(8),
				@CCCode   nvarchar(8),
				@tActVal  numeric(19,6),
				@ProdCode nvarchar(20),
				@ActVal   numeric(19,6),
				@cAmount  numeric(19,6)
        
CREATE TABLE #CO120_03 (    -- 작업일보
    ItemCode nvarchar(20) COLLATE Korean_Wansung_Unicode_CI_AS,
    POEntry  int,
    Quantity numeric(19,6),
 CONSTRAINT [PK_#CO120_03] PRIMARY KEY CLUSTERED 
    ( ItemCode,POEntry))

CREATE TABLE #CO120A03 (    -- 제품별
    ItemCode nvarchar(20) COLLATE Korean_Wansung_Unicode_CI_AS,
    ItmMSort nvarchar(8) COLLATE Korean_Wansung_Unicode_CI_AS,
    Amount   numeric(19,6),
    NE       char(1),
 CONSTRAINT [PK_#CO120A03] PRIMARY KEY CLUSTERED 
    ( ItemCode ))

CREATE TABLE #COST3 (    -- 배부대상 비용
        CCCode    nvarchar(15) COLLATE Korean_Wansung_Unicode_CI_AS,
        Amount    numeric(19,6),
        Assign    numeric(19,6)
        )

----------------------------------------------------------------------------------------------
SET @Code = @iYM+STR(@iBPLId,1)

SELECT	@FrDate = F_RefDate,
			@ToDate = T_RefDate 
FROM		OFPR 
WHERE	CONVERT(CHAR(6),F_RefDate,112) = @iYM

SET @AcctCode = '55302010'

SELECT	@AcctName = AcctName 
FROM		OACT 
WHERE	AcctCode = @AcctCode

-- 외주가공비
DELETE 
FROM		Z_PS_CO120L 
WHERE	Code = @Code 
			AND CECode = @AcctCode
-- 기계공구
INSERT		Z_PS_CO120L
SELECT		@Code,
				DocEntry,
				0,
				@AcctCode,
				@AcctName,
				SUM(Quantity),
				SUM(LineTotal),
				'',
				'',
				'Y'
FROM			(
					SELECT		T2.DocEntry, 
									(T1.U_OrdNum+'-'+T1.U_OrdSub1+'-'+t1.U_OrdSub2) OrdKey, 
									T1.Quantity, 
									T1.LineTotal
					FROM			OPCH T0
									JOIN 
									PCH1 T1 
										ON T1.DocEntry = T0.DocEntry
									JOIN 
									[@PS_PP030H] T2 
										ON T2.U_OrdNum = T1.U_OrdNum 
										AND T2.U_OrdSub1 = T1.U_OrdSub1 
										AND T2.U_OrdSub2 = T1.U_OrdSub2
					WHERE		T0.DocDate BETWEEN @FrDate AND @ToDate 
									and t1.acctcode = '55302010'
									AND T0.BPLId = @iBPLId
									AND T2.U_OrdGbn IN ('105','106')
									AND T2.Canceled = 'N'
									
					UNION ALL
					
					SELECT		T2.DocEntry, 
									(T1.U_OrdNum+'-'+T1.U_OrdSub1+'-'+t1.U_OrdSub2), 
									-T1.Quantity, 
									-T1.LineTotal
					FROM			ORPC T0
									JOIN 
									RPC1 T1 
										ON T1.DocEntry = T0.DocEntry
									JOIN 
									[@PS_PP030H] T2 
										ON T2.U_OrdNum = T1.U_OrdNum 
										AND T2.U_OrdSub1 = T1.U_OrdSub1 
										AND T2.U_OrdSub2 = T1.U_OrdSub2
					WHERE		T0.DocDate BETWEEN @FrDate AND @ToDate 
									and t1.acctcode = '55302010'
									AND T0.BPLId = @iBPLId
									AND T2.U_OrdGbn IN ('105','106')
									AND T2.Canceled = 'N'
				) A 
GROUP BY	A.DocEntry

--//포장사업팀
If @iBPLId = '3'
	Begin
		INSERT Z_PS_CO120L
		SELECT @Code,
			   T1.U_sItemCode,
			   right(T2.U_CCCode,1),
			   @AcctCode,
			   @AcctName,
			   0,
			   SUM(CASE When T2.U_CCCode = T3.MCCCode Then T1.LineTotal Else 0 End),
			   T2.U_CCCode,
			   T2.U_CCName,
			   Max(CASE When T2.U_CCCode = T3.MCCCode Then 'Y' Else '' End)
		  FROM OPCH T0 JOIN PCH1 T1 ON T1.DocEntry = T0.DocEntry
			           Join OITM M  On T1.U_sItemCode = m.ItemCode
			           Left Join (Select L.U_ItmBsort, L.U_CCCode, L.U_CCName From [@PS_PP001H] H Inner JOin [@PS_PP001L] L On H.Code = L.Code
									Where H.U_CpBCode in ('CP701', 'CP702') ) T2 On M.U_ItmBsort = T2.U_ItmBsort
					   Left Join (Select L.U_ItmBsort, MCCCode = Max(L.U_CCCode) From [@PS_PP001H] H Inner JOin [@PS_PP001L] L On H.Code = L.Code
									Where H.U_CpBCode in ('CP701', 'CP702')
									Group by L.U_ItmBsort ) T3 On M.U_ItmBsort = T3.U_ItmBsort
		 WHERE T0.DocDate BETWEEN @FrDate AND @ToDate 
		   and t1.acctcode = @AcctCode
		   AND T0.BPLId In ('3','5')
		 Group by U_sItemCode, T2.U_CCCode, T2.U_CCName

		UNION ALL
							
		SELECT @Code,
			   T1.U_sItemCode,
			   right(T2.U_CCCode,1),
			   @AcctCode,
			   @AcctName,
			   0,
			   SUM(CASE When T2.U_CCCode = T3.MCCCode Then T1.LineTotal Else 0 End),
			   T2.U_CCCode,
			   T2.U_CCName,
			   Max(CASE When T2.U_CCCode = T3.MCCCode Then 'Y' Else '' End)
		  FROM ORPC T0 JOIN RPC1 T1 ON T1.DocEntry = T0.DocEntry
			           Join OITM M  On T1.U_sItemCode = m.ItemCode
			           Left Join (Select L.U_ItmBsort, L.U_CCCode, L.U_CCName From [@PS_PP001H] H Inner JOin [@PS_PP001L] L On H.Code = L.Code
									Where H.U_CpBCode in ('CP701', 'CP702') ) T2 On M.U_ItmBsort = T2.U_ItmBsort
					   Left Join (Select L.U_ItmBsort, MCCCode = Max(L.U_CCCode) From [@PS_PP001H] H Inner JOin [@PS_PP001L] L On H.Code = L.Code
									Where H.U_CpBCode in ('CP701', 'CP702')
									Group by L.U_ItmBsort ) T3 On M.U_ItmBsort = T3.U_ItmBsort
		 WHERE T0.DocDate BETWEEN @FrDate AND @ToDate 
		   and t1.acctcode = @AcctCode
		   AND T0.BPLId in ('3', '5')
		 Group by T1.U_sItemCode, T2.U_CCCode, T2.U_CCName
		 
   End
-- 엔드베어링 -------------------------------------------------------------
INSERT		#CO120_03
SELECT		T1.U_ItemCode,
				T1.U_PP030HNo,
				SUM(T1.U_PQty) Quantity
FROM			[@PS_PP040H] T0
				JOIN 
				[@PS_PP040L] T1 
					ON T1.DocEntry = T0.DocEntry
				JOIN 
				OITM T2 
					ON T2.ItemCode = T1.U_ItemCode
WHERE		T0.Canceled = 'N'
				AND T0.U_DocDate BETWEEN @FrDate AND @ToDate
				AND T1.U_BPLId = @iBPLId
				AND T2.U_ItmBSort = '107'
GROUP BY	T1.U_ItemCode,
				T1.U_PP030HNo
HAVING		SUM(T1.U_PQty) > 0


DECLARE CUR0 CURSOR FOR
SELECT		T1.U_sItemCode,
				SUM(T1.LineTotal)
FROM			OPCH T0
				JOIN 
				PCH1 T1 
					ON T1.DocEntry = T0.DocEntry
				JOIN 
				OITM T2 
					ON T2.ItemCode = T1.U_sItemCode
WHERE		T0.DocDate BETWEEN @FrDate AND @ToDate
				AND T0.BPLId = @iBPLId
				AND T1.AcctCode = @AcctCode
				AND T1.TargetType <> 19
				AND T2.U_ItmBSort = '107'
GROUP BY	T1.U_sItemCode

OPEN  CUR0
FETCH NEXT FROM CUR0 INTO @ItemCode,@Amount
WHILE @@FETCH_STATUS = 0
BEGIN
    SELECT @tQty = SUM(Quantity) FROM #CO120_03 WHERE ItemCode = @ItemCode
    DECLARE CUR1 CURSOR FOR
    SELECT POEntry,Quantity
      FROM #CO120_03
     WHERE ItemCode = @ItemCode

    OPEN  CUR1
    FETCH NEXT FROM CUR1 INTO @POEntry,@Quantity
    WHILE @@FETCH_STATUS = 0
    BEGIN
        SET @cAmt = ROUND(@Amount * @Quantity / @tQty,0)
        SET @tQty = @tQTy - @Quantity
        SET @Amount = @Amount - @cAmt
        INSERT Z_PS_CO120L VALUES (@Code,@POEntry,1,@AcctCode,@AcctName,@Quantity,@cAmt,'','','Y')
    FETCH NEXT FROM CUR1 INTO @POEntry,@Quantity
    END
    CLOSE CUR1
    DEALLOCATE CUR1
    -- CURSOR1 END --
FETCH NEXT FROM CUR0 INTO @ItemCode,@Amount
END
CLOSE CUR0
DEALLOCATE CUR0

-- 기타 -------------------------------------------------------------
DELETE FROM #CO120_03
INSERT #CO120_03
SELECT T1.U_ItemCode,T1.U_PP030HNo,SUM(T1.U_PQty) Quantity
  FROM [@PS_PP040H] T0
  JOIN [@PS_PP040L] T1 ON T1.DocEntry = T0.DocEntry
  JOIN [@PS_PP030M] T2 ON T2.DocEntry = T1.U_PP030HNo AND T2.LineId = T1.U_PP030MNo
  JOIN OITM         T3 ON T3.ItemCode = T1.U_ItemCode
 WHERE T0.Canceled = 'N'
   AND T0.U_DocDate BETWEEN @FrDate AND @ToDate
   AND T1.U_BPLId = @iBPLId
   And T1.U_BPLId <> '3'
   AND T2.U_WorkGbn = '30'
   AND T3.U_ItmBSort NOT IN ('105','106','107')
 GROUP BY T1.U_ItemCode,T1.U_PP030HNo
HAVING SUM(T1.U_PQty) > 0

INSERT #CO120A03
SELECT ItemCode,U_ItmMSort,SUM(LineTotal)AMT,'N'
  FROM (
SELECT T2.ItemCode,T2.U_ItmMSort,T1.LineTotal
  FROM OPCH T0
  JOIN PCH1 T1 ON T1.DocEntry =T0.DocEntry
  JOIN OITM T2 ON T2.ItemCode = T1.U_sItemCode
 WHERE T0.DocDate BETWEEN @FrDate AND @ToDate
   AND T0.BPLId = @iBPLId
   --And T0.U_BPLId <> '3'
   AND T1.AcctCode = @AcctCode
   AND T2.U_ItmBSort NOT IN ('105','106','107')
 UNION ALL
 --대변메모
SELECT T2.ItemCode,T2.U_ItmMSort,T1.LineTotal * -1
  FROM ORPC T0
  JOIN RPC1 T1 ON T1.DocEntry =T0.DocEntry
  JOIN OITM T2 ON T2.ItemCode = T1.U_sItemCode
 WHERE T0.DocDate BETWEEN @FrDate AND @ToDate
   AND T0.BPLId = @iBPLId
   --And T0.U_BPLId <> '3'
   AND T1.AcctCode = @AcctCode
   AND T2.U_ItmBSort NOT IN ('105','106','107')
) AA GROUP BY ItemCode,U_ItmMSort

-- 
UPDATE A SET NE = 'Y'
  FROM #CO120A03 A LEFT JOIN #CO120_03 B ON B.ItemCode = A.ItemCode
 WHERE B.ItemCode IS NULL

SELECT @SAmt=SUM(Amount) FROM #CO120A03 WHERE NE = 'Y'
SELECT @RAmt=SUM(Amount) FROM #CO120A03 WHERE NE = 'N'

DECLARE CUR9 CURSOR FOR
 SELECT ItemCode,Amount FROM #CO120A03 WHERE NE = 'N' and @SAmt <> 0
OPEN  CUR9
FETCH NEXT FROM CUR9 INTO @ItemCode,@Amount
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @cAmt = ROUND(@SAmt * @Amount / @RAmt,0)
    UPDATE #CO120A03 SET Amount = Amount + @cAmt WHERE ItemCode = @ItemCode
    SET @SAmt = @SAmt - @cAmt
    SET @RAmt = @Ramt - @Amount
FETCH NEXT FROM CUR9 INTO @ItemCode,@Amount
END

-- 외주가공비 배부
DECLARE CUR0 CURSOR FOR
SELECT ItemCode,ItmMSort,Amount FROM #CO120A03 WHERE NE = 'N'
OPEN  CUR0
FETCH NEXT FROM CUR0 INTO @ItemCode,@ItmMSort,@Amount
WHILE @@FETCH_STATUS = 0
BEGIN
    IF @ItmMSort = '10107' --후렌지 POLine을 0에서 1로 수정 2011-10-07 노근용
		INSERT Z_PS_CO120L VALUES (@Code,@ItemCode,1,@AcctCode,@AcctName,0,@Amount,'','','Y')
        --INSERT Z_PS_CO120L VALUES (@Code,@ItemCode,0,@AcctCode,@AcctName,0,@Amount,'','','Y')
    SELECT @tQty = SUM(Quantity) FROM #CO120_03 WHERE ItemCode = @ItemCode
    DECLARE CUR1 CURSOR FOR
    SELECT POEntry,Quantity
      FROM #CO120_03
     WHERE ItemCode = @ItemCode

    OPEN  CUR1
    FETCH NEXT FROM CUR1 INTO @POEntry,@Quantity
    WHILE @@FETCH_STATUS = 0
    BEGIN
        SET @cAmt = ROUND(@Amount * @Quantity / @tQty,0)
        SET @tQty = @tQTy - @Quantity
        SET @Amount = @Amount - @cAmt
        INSERT Z_PS_CO120L VALUES (@Code,@POEntry,1,@AcctCode,@AcctName,@Quantity,@cAmt,'','','Y')
    FETCH NEXT FROM CUR1 INTO @POEntry,@Quantity
    END
    CLOSE CUR1
    DEALLOCATE CUR1
    -- CURSOR1 END --
FETCH NEXT FROM CUR0 INTO @ItemCode,@ItmMSort,@Amount
END
CLOSE CUR0
DEALLOCATE CUR0

------------------------------- 포장사업팀 --------------------------
INSERT #COST3
SELECT ISNULL(U_ProdCode,ISNULL(U_ItmMSort,U_ItmBSort))CCCode,SUM(U_Qty)Amount,0 Assign
FROM [@PS_CO302L]
WHERE Code = @Code AND U_IType = 'S'
GROUP BY ISNULL(U_ProdCode,ISNULL(U_ItmMSort,U_ItmBSort))

-- Cursor 2 START --
DECLARE CUR2 CURSOR FOR
SELECT CCCode,Amount
  FROM #COST3

OPEN  CUR2
FETCH NEXT FROM CUR2 INTO @CCCode,@Amount
WHILE @@FETCH_STATUS = 0
BEGIN
    SELECT @tActVal=U_ActVal
      FROM [@PS_CO303L]
     WHERE Code=@Code AND (U_ProdCode=@CCCode OR U_ItmMSort=@CCCode OR U_ItmBSort=@CCCode)

    -- CURSOR3 START --
    DECLARE CUR3 CURSOR FOR
    SELECT U_ProdCode,U_ActVal,AcctCode,AcctName
      FROM [@PS_CO303L] T0
      JOIN OACT T1 ON T1.AcctCode = '55302010'
     WHERE Code=@Code AND (U_ProdCode=@CCCode OR U_ItmMSort=@CCCode OR U_ItmBSort=@CCCode)

    OPEN  CUR3
    FETCH NEXT FROM CUR3 INTO @ProdCode,@ActVal,@AcctCode,@AcctName
    WHILE @@FETCH_STATUS = 0
    BEGIN
        SET @cAmount = ROUND(@Amount * @ActVal / @tActVal,0)
        INSERT #CO120B VALUES (@ProdCode,0,@AcctCode,@AcctName,@ActVal,@cAmount,'','')
        UPDATE #COST3 SET Assign = Assign + @cAmount WHERE CCCode = @CCCode
        SET @Amount  = @Amount  - @cAmount
        SET @tActVal = @tActVal - @ActVal

    FETCH NEXT FROM CUR3 INTO @ProdCode,@ActVal,@AcctCode,@AcctName
    END
    CLOSE CUR3
    DEALLOCATE CUR3
    -- CURSOR3 END --
FETCH NEXT FROM CUR2 INTO @CCCode,@Amount
END
CLOSE CUR2
DEALLOCATE CUR2
-- Cursor 2 END --
