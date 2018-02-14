IF OBJECT_ID('PS_PP030_02') IS NOT NULL
BEGIN
	DROP PROC PS_PP030_02
END
GO
--EXEC PS_PP030_02 '','','','10','',''
CREATE PROC PS_PP030_02
(
	@BPLId NVARCHAR(10),
	@ItmBsort NVARCHAR(10),
	@ItmMsort NVARCHAR(10),
	@ReqType NVARCHAR(10),
	@ItemCode NVARCHAR(20),
	@CardCode NVARCHAR(20)
)
AS
BEGIN
	(SELECT --기계공구,몰드
		'작번요청',
		PS_PP020H.DocEntry,
		OITM.U_ItmBSort,
		PS_PP020H.U_JakName,--U_RegNum,
		PS_PP020H.U_SubNo1,
		PS_PP020H.U_SubNo2,
		CONVERT(NVARCHAR,PS_PP020H.U_ReDate,112), --완료요구일
		PS_PP020H.U_ItemCode,
		PS_PP020H.U_ItemName,		
		PS_PP020H.U_WrWeight,
		PS_PP020H.U_WrWeight - ISNULL(PS_PP030H.WrWeight,0),		
		ISNULL(PS_PP020H.U_SjDocNum,''),			
		ISNULL(PS_PP020H.U_SjLinNum,''),
		PS_PP020H.U_CardCode,
		PS_PP020H.U_CardName,
		'',
		PS_PP020H.U_BPLId,
		PS_PP020H.U_JakMyung,
		PS_PP020H.U_JakSize,
		PS_PP020H.U_JakUnit
	FROM 
	--SELECT * FROM [@PS_PP020H]
		[@PS_PP020H] PS_PP020H
		LEFT JOIN 
		(SELECT
			U_ItemCode AS ItemCode,
			U_BaseType AS BaseType,
			U_BaseNum AS BaseNum,
			SUM(U_SelWt) AS WrWeight
		FROM
			[@PS_PP030H]
		WHERE
			Canceled = 'N'
		GROUP BY
			U_BaseNum,
			U_BaseType,
			U_ItemCode
		) PS_PP030H ON PS_PP020H.DocEntry = PS_PP030H.BaseNum AND PS_PP030H.BaseType = '작번요청' AND PS_PP030H.ItemCode = PS_PP020H.U_ItemCode
		LEFT JOIN [OITM] OITM ON PS_PP020H.U_ItemCode = OITM.ItemCode			
	WHERE
		PS_PP020H.U_WrWeight - ISNULL(PS_PP030H.WrWeight,0) > 0
		AND OITM.U_ItmBsort IN('105','106')--기계공구,몰드
		AND (OITM.U_ItmBsort = @ItmBsort OR @ItmBsort = '')
		AND (OITM.U_ItmMsort = @ItmMsort OR @ItmMsort = '')
		AND (@ReqType = '')
		--AND ((SELECT U_ReGbn FROM [@PS_SD010H] WHERE ''='') IN(@ReqType) OR @ReqType = '')
		--AND (PS_PP020H.U_ReGbn = @ReqType OR @ReqType = '')
		AND (PS_PP020H.U_ItemCode = @ItemCode OR @ItemCode = '')
		AND (PS_PP020H.U_CardCode = @CardCode OR @CardCode ='')
		AND (PS_PP020H.U_BPLId = @BPLId OR @BPLId ='')		
		AND PS_PP020H.Canceled = 'N'		
	)
	UNION ALL
	(SELECT --휘팅,부품만!
		'생산요청',
		PS_SD010H.DocEntry,
		OITM.U_ItmBSort,
		PS_SD010H.U_RegNum,
		'',
		'',
		CONVERT(NVARCHAR,PS_SD010H.U_DueDate,112),
		PS_SD010H.U_ItemCode,
		PS_SD010H.U_ItemName,		
		PS_SD010H.U_ReWeight,
		PS_SD010H.U_ReWeight - ISNULL(PS_PP030H.ReWeight,0),
		CASE WHEN ISNULL(PS_SD010H.U_SjDocLin,'') = '' THEN '' ELSE
		SUBSTRING(PS_SD010H.U_SjDocLin,1,CHARINDEX('-',PS_SD010H.U_SjDocLin)-1) END,
		CASE WHEN ISNULL(PS_SD010H.U_SjDocLin,'') = '' THEN '' ELSE
		SUBSTRING(PS_SD010H.U_SjDocLin,CHARINDEX('-',PS_SD010H.U_SjDocLin)+1,LEN(PS_SD010H.U_SjDocLin) - CHARINDEX('-',PS_SD010H.U_SjDocLin)) END,
		PS_SD010H.U_CardCode,
		PS_SD010H.U_CardName,
		PS_SD010H.U_ReGbn,
		PS_SD010H.U_BPLId,
		PS_SD010H.U_ItemName,
		OITM.U_Size,
		OITM.U_Unit1
		--SELECT InvntryUom FROM [OITM]
	FROM 
		[@PS_SD010H] PS_SD010H
		LEFT JOIN 
		(SELECT
			U_ItemCode AS ItemCode,
			U_BaseType AS BaseType,
			U_BaseNum AS BaseNum,
			SUM(U_SelWt) AS ReWeight
		FROM
			[@PS_PP030H]
		WHERE
			Canceled = 'N'
		GROUP BY
			U_BaseNum,
			U_BaseType,
			U_ItemCode
		) PS_PP030H ON PS_SD010H.DocEntry = PS_PP030H.BaseNum AND PS_PP030H.BaseType = '생산요청' AND PS_PP030H.ItemCode = PS_SD010H.U_ItemCode
		LEFT JOIN [OITM] OITM ON PS_SD010H.U_ItemCode = OITM.ItemCode			
	WHERE
		PS_SD010H.U_ReWeight - ISNULL(PS_PP030H.ReWeight,0) > 0
		AND OITM.U_ItmBsort IN('101','102') --휘팅,부품
		AND (OITM.U_ItmBsort = @ItmBsort OR @ItmBsort = '')
		AND (OITM.U_ItmMsort = @ItmMsort OR @ItmMsort = '')
		AND (PS_SD010H.U_ReGbn = @ReqType OR @ReqType = '')
		AND (PS_SD010H.U_ItemCode = @ItemCode OR @ItemCode = '')
		AND (PS_SD010H.U_CardCode = @CardCode OR @CardCode ='')
		AND (PS_SD010H.U_BPLId = @BPLId OR @BPLId ='')
		AND PS_SD010H.Canceled = 'N'		
	)	
END