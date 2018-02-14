IF OBJECT_ID('PS_SD040_07') IS NOT NULL
BEGIN
	DROP PROC PS_SD040_07
END
GO
--EXEC PS_SD040_07 '20101116001'
CREATE PROC PS_SD040_07
(
	@PackNo NVARCHAR(100)
)
AS
BEGIN	
	DECLARE @HasStock NVARCHAR(100)
	DECLARE @StockCount INT
	DECLARE @BatchNum NVARCHAR(100)
	DECLARE @ItemCode NVARCHAR(100)
	DECLARE @WhsCode NVARCHAR(100)
	DECLARE @Quantity NUMERIC(19,6)
	
	SET @HasStock = 'Enabled' --��� �ִ»���
	SET @StockCount = 0
	
	DECLARE CURSOR01 SCROLL CURSOR FOR
	SELECT 		
		PS_PP090L.U_LotNo AS BatchNum,
		PS_PP090L.U_ItemCode AS ItemCode,
		(SELECT WhsCode FROM OIBT WHERE ItemCode = PS_PP090L.U_ItemCode AND BatchNum = PS_PP090L.U_LotNo AND Quantity = PS_PP090L.U_Weight) AS WhsCode, --�԰�â��		
		PS_PP090L.U_Weight AS Quantity
	FROM 
		[@PS_PP090H] PS_PP090H
		LEFT JOIN [@PS_PP090L] PS_PP090L ON PS_PP090H.DocEntry = PS_PP090L.DocEntry				
	WHERE		
		PS_PP090H.U_PackNo = @PackNo		
	
	OPEN CURSOR01
	FETCH NEXT FROM CURSOR01 INTO @BatchNum,@ItemCode,@WhsCode,@Quantity
	WHILE(@@FETCH_STATUS=0)
	BEGIN
		SELECT
			@StockCount = COUNT(*) --���˻�
		FROM
			[OIBT] OIBT
		WHERE
			OIBT.ItemCode = @ItemCode
			AND OIBT.BatchNum = @BatchNum
			AND OIBT.WhsCode = @WhsCode
			AND OIBT.Quantity = @Quantity
		
		IF @StockCount <= 0 --���� ��� ���ٸ�
		BEGIN
			SET @HasStock = 'Disabled' --��� ���»���
		END
		
		FETCH NEXT FROM CURSOR01 INTO @BatchNum,@ItemCode,@WhsCode,@Quantity
	END
	CLOSE CURSOR01
	DEALLOCATE CURSOR01
	
	SELECT @HasStock --�����ȯ
END