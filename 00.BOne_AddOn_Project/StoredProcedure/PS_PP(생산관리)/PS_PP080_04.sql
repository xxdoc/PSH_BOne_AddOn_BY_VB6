IF OBJECT_ID('PS_PP080_04') IS NOT NULL
BEGIN
	DROP PROC PS_PP080_04
END
GO
--EXEC PS_PP080_04 '101','62'
CREATE PROC PS_PP080_04
(
	@OrdGbn NVARCHAR(100),
	@OIGNNo NVARCHAR(100)
)
AS
BEGIN
	IF @OrdGbn IN('102','104') --부품,멀티
	BEGIN
		SELECT
			OIBT.ItemCode,
			OIBT.WhsCode,
			OIBT.Quantity,
			IBT1.Quantity
		FROM
			[OIBT] OIBT
			LEFT JOIN
			(SELECT
				IBT1.ItemCode,
				IBT1.WhsCode,
				IBT1.BatchNum,
				IBT1.Quantity			
			FROM
				[IBT1] IBT1				
			WHERE
				IBT1.BaseType = '59'
				AND IBT1.BaseEntry = @OIGNNo				
			) IBT1 ON IBT1.ItemCode = OIBT.ItemCode AND IBT1.WhsCode = OIBT.WhsCode AND IBT1.BatchNum = OIBT.BatchNum
		WHERE
			OIBT.Quantity < IBT1.Quantity
	END
	ELSE IF @OrdGbn IN('101','105','106','107') --기타
	BEGIN
		SELECT
			OITW.ItemCode,
			OITW.WhsCode,
			OITW.OnHand,
			OIGN.Quantity
		FROM
			[OITW] OITW
			LEFT JOIN
			(SELECT
				IGN1.ItemCode,
				IGN1.WhsCode,
				IGN1.Quantity			
			FROM
				[OIGN] OIGN
				LEFT JOIN [IGN1] IGN1 ON OIGN.DocEntry = IGN1.DocEntry
			WHERE
				OIGN.DocEntry = @OIGNNo
			) OIGN ON OIGN.ItemCode = OITW.ItemCode AND OIGN.WhsCode = OITW.WhsCode
		WHERE
			OITW.OnHand < OIGN.Quantity
	END
END