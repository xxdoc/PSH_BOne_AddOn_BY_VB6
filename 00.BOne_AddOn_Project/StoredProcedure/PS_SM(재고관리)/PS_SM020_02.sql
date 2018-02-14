IF OBJECT_ID('PS_SM020_02') IS NOT NULL
BEGIN
	DROP PROC PS_SM020_02
END
GO
CREATE PROC PS_SM020_02
(
	@ItemCode VARCHAR(100)
)
AS
BEGIN
	IF((SELECT ManBtchNum FROM [OITM] WHERE ItemCode = @ItemCode) = 'Y') --배치사용품목
	BEGIN
		SELECT
			OIBT.BatchNum AS BatchNum,
			OIBT.WhsCode AS WhsCode,
			OWHS.WhsName AS WhsName,
			OBTN.U_PackNo AS PackNo,			
			OIBT.Quantity AS Weight,			
			CASE WHEN OITM.U_ItmBsort IN('104','302') THEN 1 
			WHEN OITM.U_ItmBsort IN('102') THEN OIBT.Quantity END AS SelQty,
			OIBT.Quantity AS SelWeight
		FROM
			[OIBT] OIBT
			LEFT JOIN [OBTN] OBTN ON OBTN.ItemCode = OIBT.ItemCode AND OBTN.DistNumber = OIBT.BatchNum
			LEFT JOIN [OWHS] OWHS ON OIBT.WhsCode = OWHS.WhsCode
			LEFT JOIN [OITM] OITM ON OIBT.ItemCode = OITM.ItemCode
		WHERE
			(@ItemCode = '' OR OIBT.ItemCode = @ItemCode)
			AND OIBT.Quantity > 0
	END
	ELSE
	BEGIN
		SELECT
			'' AS BatchNum,
			OITW.WhsCode AS WhsCode,
			OWHS.WhsName AS WhsName,
			'' AS PackNo,
			OITW.OnHand AS Weight,			
			'' AS SelQty,
			'' AS SelWeight
		FROM
			[OITW] OITW
			LEFT JOIN [OWHS] OWHS ON OITW.WhsCode = OWHS.WhsCode
			LEFT JOIN [OITM] OITM ON OITW.ItemCode = OITM.ItemCode
		WHERE
			(@ItemCode = '' OR OITW.ItemCode = @ItemCode)
			AND OITW.OnHand > 0
	END
END
GO