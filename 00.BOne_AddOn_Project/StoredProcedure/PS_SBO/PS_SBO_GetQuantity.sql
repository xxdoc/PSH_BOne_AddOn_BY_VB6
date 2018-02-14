IF OBJECT_ID('PS_SBO_GETQUANTITY') IS NOT NULL
BEGIN
	DROP PROCEDURE PS_SBO_GETQUANTITY
END
GO
CREATE PROC PS_SBO_GETQUANTITY
(
	@BaseType NVARCHAR(30), --EX> 17
	@BaseTable NVARCHAR(30), --EX> RDR
	@BaseEntry NVARCHAR(30), -- EX> 892
	@BaseLine NVARCHAR(30) -- EX> 0
)
AS
BEGIN
	DECLARE @Qty BIGINT
	SET @Qty = 0
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM QUT1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM RDR1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM DLN1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM RDN1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM INV1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM RIN1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM POR1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM PDN1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM RPD1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM PCH1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Qty = @Qty + (SELECT ISNULL(SUM(CONVERT(BIGINT,U_Qty)),0) FROM RPC1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	
	DECLARE @Weight NUMERIC(19,6)
	SET @Weight = 0
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM QUT1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM RDR1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM DLN1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM RDN1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM INV1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM RIN1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM POR1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM PDN1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM RPD1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM PCH1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	SET @Weight = @Weight + (SELECT ISNULL(SUM(CONVERT(NUMERIC(19,6),Quantity)),0) FROM RPC1 WHERE BaseType = @BaseType AND BaseEntry = @BaseEntry AND BaseLine = @BaseLine)
	
	DECLARE @STR NVARCHAR(1000)
	SET @STR = 
	'SELECT 
		ISNULL('+@BaseTable+'1.U_Qty,0) - '+CONVERT(NVARCHAR,@Qty)+',
		ISNULL('+@BaseTable+'1.Quantity,0) - '+CONVERT(NVARCHAR,@Weight)+'
	FROM
		[O'+@BaseTable+'] O'+@BaseTable+'
		LEFT JOIN ['+@BaseTable+'1] '+@BaseTable+'1 ON O'+@BaseTable+'.DocEntry = '+@BaseTable+'1.DocEntry
	WHERE
		O'+@BaseTable+'.DocEntry = '+ @BaseEntry + '
		AND '+@BaseTable+'1.LineNum = '+@BaseLine + ''
	EXEC(@STR)
END


	