IF OBJECT_ID('PS_SM020_01') IS NOT NULL
BEGIN
	DROP PROC PS_SM020_01
END
GO
CREATE PROC PS_SM020_01
(
	@ItemCode NVARCHAR(100),	
	@SellItem NVARCHAR(20), --판매
	@PurchaseItem NVARCHAR(20),--구매
	@StockType NVARCHAR(20), --재고타입(1)재고있는것만,(2)전체
	@TradeType NVARCHAR(20), --거래타입(1)일반,(2)임가공
	@ItmBsort NVARCHAR(10),
	@ItmMsort NVARCHAR(10),
	@Size NVARCHAR(50),
	@ItemType NVARCHAR(10),
	@Mark NVARCHAR(10),
	@ItemName NVARCHAR(100),
	@ItemGpCd NVARCHAR(10)
)
AS
BEGIN
	SELECT
		OITM.ItemCode AS ItemCode,
		OITM.ItemName AS ItemName,
		OITM.U_CallSize AS CallSize,
		(SELECT NAME FROM [@PSH_MARK] WHERE CODE = OITM.U_Mark) AS Mark,
		OITM.OnHand AS OnHand,
		OITM.IsCommited AS IsCommited,
		OITM.OnOrder AS OnOrder,
		(OITM.OnHand - OITM.IsCommited + OITM.OnOrder) AS OnEnabled
	FROM
		[OITM] OITM
		--LEFT JOIN [OITW] OITW ON OITW.ItemCode = OITM.ItemCode
		--LEFT JOIN [OWHS] OWHS ON OITW.WhsCode = OWHS.WhsCode
	WHERE
		(@ItemCode = '' OR OITM.ItemCode LIKE('%' + @ItemCode + '%'))
		AND (@ItemName = '' OR OITM.ItemName LIKE('%' + @ItemName + '%'))
		AND (@StockType = '2' OR (@StockType = '1' AND (SELECT SUM(OITW.OnHand) FROM [OITW] OITW WHERE OITW.ItemCode = OITM.ItemCode) > 0))
		AND (@SellItem = '' OR OITM.SellItem = @SellItem)
		AND (@PurchaseItem = '' OR OITM.PrchseItem = @PurchaseItem)
		AND (@TradeType = '' OR OITM.U_TradeType = @TradeType)
		AND (@ItmBsort = '선택' OR OITM.U_ItmBsort = @ItmBsort)
		AND (@ItmMsort = '선택' OR OITM.U_ItmMsort = @ItmMsort)
		AND (@Size = '' OR OITM.U_Size LIKE('%' + @Size + '%'))
		AND (@ItemType = '선택' OR OITM.U_ItemType = @ItemType)
		AND (@Mark = '선택' OR OITM.U_Mark = @Mark)
		AND (@ItemGpCd = '선택' OR OITM.ItmsGrpCod = REPLACE(@ItemGpCd,'선택',''))
END
GO