IF OBJECT_ID('PS_PP004_03') IS NOT NULL
BEGIN
	DROP PROC PS_PP004_03
END
GO
CREATE PROC PS_PP004_03
(
	@ItemCode NVARCHAR(20)
)
AS
BEGIN
	SELECT
		U_Sequence AS Sequence,
		U_ItemCode AS ItemCode,
		U_ItemName AS ItemName,
		U_CpBCode AS CpBCode,
		U_CpBName AS CpBName,
		U_CpCode AS CpCode,
		U_CpName AS CpName,
		U_CpUnWt AS CpUnWt,
		U_ResultYN AS ResultYN,
		U_ReportYN AS ReportYN,
		U_StdTime AS StdTime		
	FROM
		[@PS_PP004H]
	WHERE
		U_ItemCode = @ItemCode
END