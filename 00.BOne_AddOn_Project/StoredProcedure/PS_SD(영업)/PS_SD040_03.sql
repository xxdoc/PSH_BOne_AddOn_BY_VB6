IF OBJECT_ID('PS_SD040_03') IS NOT NULL
BEGIN
	DROP PROC PS_SD040_03
END
GO
CREATE PROC PS_SD040_03
(
	@ItemCode NVARCHAR(100),
	@BatchNum NVARCHAR(100),
	@WhsCode NVARCHAR(100)
)
AS
BEGIN
	SELECT
		OIBT.ItemCode AS ItemCode, 
		OIBT.WhsCode AS WhsCode,
		OIBT.BatchNum AS BatchNum,
		OIBT.Quantity AS Weight	
	FROM
		[OIBT] OIBT
		--LEFT JOIN [OBTN] OBTN ON OIBT.ItemCode = OBTN.ItemCode AND OIBT.BatchNum = OBTN.DistNumber
	WHERE
		OIBT.ItemCode = @ItemCode
		AND OIBT.BatchNum = @BatchNum
		AND OIBT.WhsCode = @WhsCode
		AND OIBT.Quantity > 0
END