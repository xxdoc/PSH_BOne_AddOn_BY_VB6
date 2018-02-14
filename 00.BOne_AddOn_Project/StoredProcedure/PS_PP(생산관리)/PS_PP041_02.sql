IF OBJECT_ID('PS_PP041_02') IS NOT NULL
BEGIN
	DROP PROC PS_PP041_02
END
GO
--EXEC PS_PP041_02 '선택','선택','CP50103'
CREATE PROC PS_PP041_02
(
	@OrdMgNum NVARCHAR(100)
)
AS
BEGIN
	DECLARE @CpCode NVARCHAR(100)	
	SET @CpCode = (SELECT U_CpCode FROM [@PS_PP030M] PS_PP030M WHERE CONVERT(NVARCHAR,PS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = @OrdMgNum)
	IF (SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = @CpCode) = 'CP101' --엔드베어링
	BEGIN
		SELECT
			CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) AS OrdMgNum,
			PS_PP030M.U_Sequence AS Sequence,
			PS_PP030M.U_CpCode AS CpCode,
			PS_PP030M.U_CpName AS CpName,
			PS_PP030H.U_OrdGbn AS OrdGbn,
			PS_PP030H.U_BPLId AS BPLId,
			PS_PP030H.U_ItemCode AS ItemCode,
			PS_PP030H.U_ItemName AS ItemName,
			PS_PP030H.U_OrdNum AS OrdNum,
			PS_PP030H.U_OrdSub1 AS OrdSub1,
			PS_PP030H.U_OrdSub2 AS OrdSub2,
			PS_PP030H.DocEntry AS PP030HNo,
			PS_PP030M.LineId AS PP030MNo,
			0 AS BQty,
			(SELECT ISNULL(SUM(SUB_PS_PP040L.U_PQty),0) FROM [@PS_PP040H] SUB_PS_PP040H LEFT JOIN [@PS_PP040L] SUB_PS_PP040L ON SUB_PS_PP040H.DocEntry = SUB_PS_PP040L.DocEntry WHERE SUB_PS_PP040H.Canceled = 'N' AND SUB_PS_PP040L.U_PP030HNo = PS_PP030H.DocEntry AND SUB_PS_PP040L.U_PP030MNo = PS_PP030M.LineId) AS PSum
			
		FROM 
			[@PS_PP030H] PS_PP030H
			LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry
		WHERE
			CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = @OrdMgNum
			AND PS_PP030H.Canceled = 'N'
			AND PS_PP030M.U_ReportYN = 'Y' --일보여부가 'Y' 인것들			
	END
	ELSE IF (SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = @CpCode) = 'CP501' --멀티
	BEGIN
		DECLARE @FirstCpCode NVARCHAR(30) --첫공정코드
		SELECT TOP 1 @FirstCpCode = U_CpCode FROM [@PS_PP001L] WHERE LEFT(U_CpCode,5) = LEFT(@CpCode,5) ORDER BY LineId ASC
		IF @FirstCpCode = @CpCode
		BEGIN		
			SELECT
				CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) AS OrdMgNum,
				PS_PP030M.U_Sequence AS Sequence,
				PS_PP030M.U_CpCode AS CpCode,
				PS_PP030M.U_CpName AS CpName,
				PS_PP030H.U_OrdGbn AS OrdGbn,
				PS_PP030H.U_BPLId AS BPLId,
				PS_PP030H.U_ItemCode AS ItemCode,
				PS_PP030H.U_ItemName AS ItemName,
				PS_PP030H.U_OrdNum AS OrdNum,
				PS_PP030H.U_OrdSub1 AS OrdSub1,
				PS_PP030H.U_OrdSub2 AS OrdSub2,
				PS_PP030H.DocEntry AS PP030HNo, --작업지시 헤더
				PS_PP030M.LineId AS PP030MNo, --작업지시 공정라인
				(SELECT SUM(SUB_PS_PP030L.U_Weight) FROM [@PS_PP030L] SUB_PS_PP030L WHERE SUB_PS_PP030L.DocEntry = PS_PP030H.DocEntry) AS BQty,
				0 AS PSum
			FROM
				[@PS_PP030H] PS_PP030H
				LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry
				LEFT JOIN
				(SELECT
					PS_PP040L.U_PP030HNo AS PP030HNo,
					PS_PP040L.U_PP030MNo AS PP030MNo,
					SUM(PS_PP040L.U_PQty) AS PQty
				FROM
					[@PS_PP040H] PS_PP040H
					LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry
				WHERE
					PS_PP040H.Canceled = 'N'
				GROUP BY
					PS_PP040L.U_PP030HNo,
					PS_PP040L.U_PP030MNo
				) PS_PP040 ON PS_PP040.PP030HNo = PS_PP030H.DocEntry AND PS_PP040.PP030MNo = PS_PP030M.LineId
			WHERE
				CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = @OrdMgNum
				AND PS_PP030H.Canceled = 'N'
				AND PS_PP030M.U_ReportYN = 'Y' --일보여부가 'Y' 인것들
				AND PS_PP040.PP030HNo IS NULL --작업일보등록되지 않은값
		END
		ELSE --첫공정 이외의 공정에 대해
		BEGIN
			SELECT
				CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) AS OrdMgNum,
				PS_PP030M.U_Sequence AS Sequence,
				PS_PP030M.U_CpCode AS CpCode,
				PS_PP030M.U_CpName AS CpName,
				PS_PP030H.U_OrdGbn AS OrdGbn,
				PS_PP030H.U_BPLId AS BPLId,
				PS_PP030H.U_ItemCode AS ItemCode,
				PS_PP030H.U_ItemName AS ItemName,
				PS_PP030H.U_OrdNum AS OrdNum,
				PS_PP030H.U_OrdSub1 AS OrdSub1,
				PS_PP030H.U_OrdSub2 AS OrdSub2,
				PS_PP030H.DocEntry AS PP030HNo, --작업지시 헤더
				PS_PP030M.LineId AS PP030MNo, --작업지시 공정라인
				(SELECT 
					ISNULL(SUM(PS_PP040L.U_PQty),0)
				FROM 
					[@PS_PP040H] PS_PP040H
					LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry
				WHERE
					PS_PP040H.Canceled = 'N'
					AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) =
					(SELECT
						TOP 1 
						CONVERT(NVARCHAR,PrevPS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PrevPS_PP030M.LineId)
					FROM
						[@PS_PP030M] PrevPS_PP030M
						LEFT JOIN
						(SELECT
							CurrentPS_PP030H.DocEntry AS DocEntry,
							CurrentPS_PP030M.U_Sequence AS Sequence
						FROM
							[@PS_PP030H] CurrentPS_PP030H
							LEFT JOIN [@PS_PP030M] CurrentPS_PP030M ON CurrentPS_PP030H.DocEntry = CurrentPS_PP030M.DocEntry
						WHERE
							CONVERT(NVARCHAR,CurrentPS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,CurrentPS_PP030M.LineId) = CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId)
						) CURRENTROW ON CURRENTROW.DocEntry = PrevPS_PP030M.DocEntry
					WHERE
						PrevPS_PP030M.U_Sequence < CURRENTROW.Sequence
					ORDER BY
						U_Sequence DESC
					)
				) AS BQty,
				0 AS PSum
			FROM
				[@PS_PP030H] PS_PP030H
				LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry
				LEFT JOIN
				(SELECT
					PS_PP040L.U_PP030HNo AS PP030HNo,
					PS_PP040L.U_PP030MNo AS PP030MNo
				FROM
					[@PS_PP040H] PS_PP040H
					LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry
				WHERE
					PS_PP040H.Canceled = 'N'
				) PS_PP040 ON PS_PP040.PP030HNo = PS_PP030H.DocEntry AND PS_PP040.PP030MNo = PS_PP030M.LineId
			WHERE
				CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = @OrdMgNum
				AND PS_PP030H.Canceled = 'N'
				AND PS_PP030M.U_ReportYN = 'Y' --일보여부가 'Y' 인것들			
				AND PS_PP040.PP030HNo IS NULL --작업일보등록되지 않은값
				AND 
				(SELECT --이전공정이 존재하는경우
					COUNT(*)
				FROM 
					[@PS_PP040H] PS_PP040H
					LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry
				WHERE
					PS_PP040H.Canceled = 'N' 
					AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) =
					(SELECT
						TOP 1 
						CONVERT(NVARCHAR,PrevPS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PrevPS_PP030M.LineId)
					FROM
						[@PS_PP030M] PrevPS_PP030M
						LEFT JOIN
						(SELECT
							CurrentPS_PP030H.DocEntry AS DocEntry,
							CurrentPS_PP030M.U_Sequence AS Sequence
						FROM
							[@PS_PP030H] CurrentPS_PP030H
							LEFT JOIN [@PS_PP030M] CurrentPS_PP030M ON CurrentPS_PP030H.DocEntry = CurrentPS_PP030M.DocEntry
						WHERE
							CONVERT(NVARCHAR,CurrentPS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,CurrentPS_PP030M.LineId) = CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId)
						) CURRENTROW ON CURRENTROW.DocEntry = PrevPS_PP030M.DocEntry
					WHERE
						PrevPS_PP030M.U_Sequence < CURRENTROW.Sequence
					ORDER BY
						U_Sequence DESC
					)
				) > 0
			END
		END
	END
	
--	DECLARE @FirstCpCode NVARCHAR(30) --첫공정코드
--	SELECT TOP 1 @FirstCpCode = U_CpCode FROM [@PS_PP001L] WHERE LEFT(U_CpCode,5) = LEFT(@CpCode,5) ORDER BY LineId ASC
--	IF @FirstCpCode = @CpCode
--	BEGIN
--		IF (SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = @CpCode) = 'CP101' --엔드베어링
--		BEGIN --첫공정은 계속해서 입력가능하다.. 단 해당월이 지나면 조회대상이 아니다..
--			SELECT
--				CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) AS OrdMgNum,
--				PS_PP030M.U_Sequence AS Sequence,
--				PS_PP030M.U_CpCode AS CpCode,
--				PS_PP030M.U_CpName AS CpName,
--				PS_PP030H.U_OrdGbn AS OrdGbn,
--				PS_PP030H.U_BPLId AS BPLId,
--				PS_PP030H.U_ItemCode AS ItemCode,
--				PS_PP030H.U_ItemName AS ItemName,
--				PS_PP030H.U_OrdNum AS OrdNum,
--				PS_PP030H.U_OrdSub1 AS OrdSub1,
--				PS_PP030H.U_OrdSub2 AS OrdSub2,
--				PS_PP030H.DocEntry AS PP030HNo, --작업지시 헤더
--				PS_PP030M.LineId AS PP030MNo, --작업지시 공정라인
--				0 AS BQty,
--				(SELECT ISNULL(SUM(SUB_PS_PP040L.U_PQty),0) FROM [@PS_PP040H] SUB_PS_PP040H LEFT JOIN [@PS_PP040L] SUB_PS_PP040L ON SUB_PS_PP040H.DocEntry = SUB_PS_PP040L.DocEntry WHERE SUB_PS_PP040H.Canceled = 'N' AND SUB_PS_PP040L.U_PP030HNo = PS_PP030H.DocEntry AND SUB_PS_PP040L.U_PP030MNo = PS_PP030M.LineId) AS PSum					
--			FROM
--				[@PS_PP030H] PS_PP030H
--				LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry				
--			WHERE				
--				(@BPLId = '선택' OR PS_PP030H.U_BPLId = @BPLId)
--				AND (@OrdGbn = '선택' OR PS_PP030H.U_OrdGbn = @OrdGbn)
--				AND PS_PP030M.U_CpCode = @CpCode
--				AND PS_PP030H.Canceled = 'N'
--				AND DATEPART(MONTH,PS_PP030H.U_DocDate) = DATEPART(MONTH,GETDATE()) --작업일보의 월과 현재 월이 같으면 대상
--		END
--		ELSE IF (SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = @CpCode) = 'CP501' --멀티
--		BEGIN
--			SELECT
--				CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) AS OrdMgNum,
--				PS_PP030M.U_Sequence AS Sequence,
--				PS_PP030M.U_CpCode AS CpCode,
--				PS_PP030M.U_CpName AS CpName,
--				PS_PP030H.U_OrdGbn AS OrdGbn,
--				PS_PP030H.U_BPLId AS BPLId,
--				PS_PP030H.U_ItemCode AS ItemCode,
--				PS_PP030H.U_ItemName AS ItemName,
--				PS_PP030H.U_OrdNum AS OrdNum,
--				PS_PP030H.U_OrdSub1 AS OrdSub1,
--				PS_PP030H.U_OrdSub2 AS OrdSub2,
--				PS_PP030H.DocEntry AS PP030HNo, --작업지시 헤더
--				PS_PP030M.LineId AS PP030MNo, --작업지시 공정라인
--				(SELECT SUB_PS_PP030L.U_Weight FROM [@PS_PP030L] SUB_PS_PP030L WHERE SUB_PS_PP030L.DocEntry = PS_PP030H.DocEntry) AS BQty,
--				0 AS PSum
--			FROM
--				[@PS_PP030H] PS_PP030H
--				LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry
--				LEFT JOIN
--				(SELECT
--					PS_PP040L.U_PP030HNo AS PP030HNo,
--					PS_PP040L.U_PP030MNo AS PP030MNo,
--					SUM(PS_PP040L.U_PQty) AS PQty
--				FROM
--					[@PS_PP040H] PS_PP040H
--					LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry
--				WHERE
--					PS_PP040H.Canceled = 'N'
--				GROUP BY
--					PS_PP040L.U_PP030HNo,
--					PS_PP040L.U_PP030MNo
--				) PS_PP040 ON PS_PP040.PP030HNo = PS_PP030H.DocEntry AND PS_PP040.PP030MNo = PS_PP030M.LineId
--			WHERE
--				PS_PP040.PP030HNo IS NULL --작업일보등록되지 않은값
--				AND (@BPLId = '선택' OR PS_PP030H.U_BPLId = @BPLId)
--				AND (@OrdGbn = '선택' OR PS_PP030H.U_OrdGbn = @OrdGbn)
--				AND PS_PP030M.U_CpCode = @CpCode
--				AND PS_PP030H.Canceled = 'N'
--		END
--	END
--	ELSE --첫공정이 아닌경우
--	BEGIN
--		IF (SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = @CpCode) = 'CP101' --엔드베어링
--		BEGIN
--			SELECT
--				CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) AS OrdMgNum,
--				PS_PP030M.U_Sequence AS Sequence,
--				PS_PP030M.U_CpCode AS CpCode,
--				PS_PP030M.U_CpName AS CpName,
--				PS_PP030H.U_OrdGbn AS OrdGbn,
--				PS_PP030H.U_BPLId AS BPLId,
--				PS_PP030H.U_ItemCode AS ItemCode,
--				PS_PP030H.U_ItemName AS ItemName,
--				PS_PP030H.U_OrdNum AS OrdNum,
--				PS_PP030H.U_OrdSub1 AS OrdSub1,
--				PS_PP030H.U_OrdSub2 AS OrdSub2,
--				PS_PP030H.DocEntry AS PP030HNo, --작업지시 헤더
--				PS_PP030M.LineId AS PP030MNo, --작업지시 공정라인
--				0 AS BQty,
--				(SELECT ISNULL(SUM(SUB_PS_PP040L.U_PQty),0) FROM [@PS_PP040H] SUB_PS_PP040H LEFT JOIN [@PS_PP040L] SUB_PS_PP040L ON SUB_PS_PP040H.DocEntry = SUB_PS_PP040L.DocEntry WHERE SUB_PS_PP040H.Canceled = 'N' AND SUB_PS_PP040L.U_PP030HNo = PS_PP030H.DocEntry AND SUB_PS_PP040L.U_PP030MNo = PS_PP030M.LineId) AS PSum					
--			FROM
--				[@PS_PP030H] PS_PP030H
--				LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry				
--			WHERE
--				(@BPLId = '선택' OR PS_PP030H.U_BPLId = @BPLId)
--				AND (@OrdGbn = '선택' OR PS_PP030H.U_OrdGbn = @OrdGbn)
--				AND PS_PP030M.U_CpCode = @CpCode
--				AND PS_PP030H.Canceled = 'N'
--				AND 
--				(SELECT --이전공정이 존재하는경우
--					COUNT(*)
--				FROM 
--					[@PS_PP040H] PS_PP040H
--					LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry
--				WHERE
--					PS_PP040H.Canceled = 'N' 
--					AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) =
--					(SELECT
--						TOP 1 
--						CONVERT(NVARCHAR,PrevPS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PrevPS_PP030M.LineId)
--					FROM
--						[@PS_PP030M] PrevPS_PP030M
--						LEFT JOIN
--						(SELECT
--							CurrentPS_PP030H.DocEntry AS DocEntry,
--							CurrentPS_PP030M.U_Sequence AS Sequence
--						FROM
--							[@PS_PP030H] CurrentPS_PP030H
--							LEFT JOIN [@PS_PP030M] CurrentPS_PP030M ON CurrentPS_PP030H.DocEntry = CurrentPS_PP030M.DocEntry
--						WHERE
--							CONVERT(NVARCHAR,CurrentPS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,CurrentPS_PP030M.LineId) = CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId)
--						) CURRENTROW ON CURRENTROW.DocEntry = PrevPS_PP030M.DocEntry
--					WHERE
--						PrevPS_PP030M.U_Sequence < CURRENTROW.Sequence
--					ORDER BY
--						U_Sequence DESC
--					)
--				) > 0
--		END
--		ELSE IF (SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = @CpCode) = 'CP501' --멀티
--		BEGIN
--			SELECT
--				CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) AS OrdMgNum,
--				PS_PP030M.U_Sequence AS Sequence,
--				PS_PP030M.U_CpCode AS CpCode,
--				PS_PP030M.U_CpName AS CpName,
--				PS_PP030H.U_OrdGbn AS OrdGbn,
--				PS_PP030H.U_BPLId AS BPLId,
--				PS_PP030H.U_ItemCode AS ItemCode,
--				PS_PP030H.U_ItemName AS ItemName,
--				PS_PP030H.U_OrdNum AS OrdNum,
--				PS_PP030H.U_OrdSub1 AS OrdSub1,
--				PS_PP030H.U_OrdSub2 AS OrdSub2,
--				PS_PP030H.DocEntry AS PP030HNo, --작업지시 헤더
--				PS_PP030M.LineId AS PP030MNo, --작업지시 공정라인
--				(SELECT 
--					ISNULL(SUM(PS_PP040L.U_PQty),0)
--				FROM 
--					[@PS_PP040H] PS_PP040H
--					LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry
--				WHERE
--					PS_PP040H.Canceled = 'N'
--					AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) =
--					(SELECT
--						TOP 1 
--						CONVERT(NVARCHAR,PrevPS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PrevPS_PP030M.LineId)
--					FROM
--						[@PS_PP030M] PrevPS_PP030M
--						LEFT JOIN
--						(SELECT
--							CurrentPS_PP030H.DocEntry AS DocEntry,
--							CurrentPS_PP030M.U_Sequence AS Sequence
--						FROM
--							[@PS_PP030H] CurrentPS_PP030H
--							LEFT JOIN [@PS_PP030M] CurrentPS_PP030M ON CurrentPS_PP030H.DocEntry = CurrentPS_PP030M.DocEntry
--						WHERE
--							CONVERT(NVARCHAR,CurrentPS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,CurrentPS_PP030M.LineId) = CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId)
--						) CURRENTROW ON CURRENTROW.DocEntry = PrevPS_PP030M.DocEntry
--					WHERE
--						PrevPS_PP030M.U_Sequence < CURRENTROW.Sequence
--					ORDER BY
--						U_Sequence DESC
--					)
--				) AS BQty,
--				0 AS PSum
--			FROM
--				[@PS_PP030H] PS_PP030H
--				LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry
--				LEFT JOIN
--				(SELECT
--					PS_PP040L.U_PP030HNo AS PP030HNo,
--					PS_PP040L.U_PP030MNo AS PP030MNo
--				FROM
--					[@PS_PP040H] PS_PP040H
--					LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry
--				WHERE
--					PS_PP040H.Canceled = 'N'
--				) PS_PP040 ON PS_PP040.PP030HNo = PS_PP030H.DocEntry AND PS_PP040.PP030MNo = PS_PP030M.LineId
--			WHERE
--				PS_PP040.PP030HNo IS NULL --작업일보등록되지 않은값
--				AND (@BPLId = '선택' OR PS_PP030H.U_BPLId = @BPLId)
--				AND (@OrdGbn = '선택' OR PS_PP030H.U_OrdGbn = @OrdGbn)
--				AND PS_PP030M.U_CpCode = @CpCode
--				AND PS_PP030H.Canceled = 'N'
--				AND 
--				(SELECT --이전공정이 존재하는경우
--					COUNT(*)
--				FROM 
--					[@PS_PP040H] PS_PP040H
--					LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry
--				WHERE
--					PS_PP040H.Canceled = 'N' 
--					AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) =
--					(SELECT
--						TOP 1 
--						CONVERT(NVARCHAR,PrevPS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PrevPS_PP030M.LineId)
--					FROM
--						[@PS_PP030M] PrevPS_PP030M
--						LEFT JOIN
--						(SELECT
--							CurrentPS_PP030H.DocEntry AS DocEntry,
--							CurrentPS_PP030M.U_Sequence AS Sequence
--						FROM
--							[@PS_PP030H] CurrentPS_PP030H
--							LEFT JOIN [@PS_PP030M] CurrentPS_PP030M ON CurrentPS_PP030H.DocEntry = CurrentPS_PP030M.DocEntry
--						WHERE
--							CONVERT(NVARCHAR,CurrentPS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,CurrentPS_PP030M.LineId) = CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId)
--						) CURRENTROW ON CURRENTROW.DocEntry = PrevPS_PP030M.DocEntry
--					WHERE
--						PrevPS_PP030M.U_Sequence < CURRENTROW.Sequence
--					ORDER BY
--						U_Sequence DESC
--					)
--				) > 0
--		END
--	END
--END