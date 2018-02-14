IF OBJECT_ID('PS_SM010_01') IS NOT NULL
BEGIN
	DROP PROC PS_SM010_01
END
GO
CREATE PROC PS_SM010_01
(
	@ItemCode NVARCHAR(100),	
	@SellItem NVARCHAR(20), --�Ǹ�
	@PurchaseItem NVARCHAR(20),--����
	@StockType NVARCHAR(20), --���Ÿ��(1)����ִ°͸�,(2)��ü
	@TradeType NVARCHAR(20), --�ŷ�Ÿ��(1)�Ϲ�,(2)�Ӱ���
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
		OITM.ItemCode AS ǰ���ڵ�,
		OITM.ItemName AS ǰ���,
		OITM.U_CallSize AS ȣĪ�԰�,
		(SELECT NAME FROM [@PSH_MARK] WHERE CODE = OITM.U_Mark) AS ������ȣ,
		OITM.OnHand AS ���,
		OITM.IsCommited AS ����,
		OITM.OnOrder AS ����,
		(OITM.OnHand - OITM.IsCommited + OITM.OnOrder) AS ����
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
		AND (@ItmBsort = '����' OR OITM.U_ItmBsort = @ItmBsort)
		AND (@ItmMsort = '����' OR OITM.U_ItmMsort = @ItmMsort)
		AND (@Size = '' OR OITM.U_Size LIKE('%' + @Size + '%'))
		AND (@ItemType = '����' OR OITM.U_ItemType = @ItemType)
		AND (@Mark = '����' OR OITM.U_Mark = @Mark)
		AND (@ItemGpCd = '����' OR OITM.ItmsGrpCod = REPLACE(@ItemGpCd,'����',''))
END
GO