IF OBJECT_ID('PS_PP004_01') IS NOT NULL
BEGIN
	DROP PROC PS_PP004_01
END
GO
CREATE PROC PS_PP004_01
(
	@ItmBsort NVARCHAR(10),
	@ItmMsort NVARCHAR(10),
	@ItemCode NVARCHAR(20)
)
AS
BEGIN
	SELECT 
		ItemCode AS ItemCode,
		ItemName AS ItemName,
		ItmsGrpCod AS ItmsGrpCod,
		U_ItmBsort AS ItmBsort,
		U_ItmMsort AS ItmMsort,
		U_Unit1 AS Unit1,
		U_Size AS Size,
		U_ItemType AS ItemType,
		U_UnWeight AS UnWeight,
		U_Quality AS Quality,
		U_Mark AS Mark,
		U_CallSize AS CallSize
	FROM 
		[OITM]
	WHERE
		(@ItmBsort = '' OR U_ItmBsort = @ItmBsort)
		AND (@ItmMsort = '' OR U_ItmMsort = @ItmMsort)
		AND (@ItemCode = '' OR ItemCode = @ItemCode)
		AND ItmsGrpCod = '102' --제품중에서만
END