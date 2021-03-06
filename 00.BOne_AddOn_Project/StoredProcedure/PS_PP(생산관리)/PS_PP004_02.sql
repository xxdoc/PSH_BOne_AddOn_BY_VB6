IF OBJECT_ID('PS_PP004_02') IS NOT NULL
BEGIN
	DROP PROC PS_PP004_02
END
GO
CREATE PROC PS_PP004_02
(
	@ItmBsort NVARCHAR(10)
)
AS
BEGIN
	SELECT
		PS_PP001H.U_CpBCode AS CpBCode,
		PS_PP001H.U_CpBName AS CpBName,
		PS_PP001L.U_CpCode AS CpCode,
		PS_PP001L.U_CpName AS CpName,
		PS_PP001L.U_ItmBsort AS ItmBsort,
		PS_PP001L.U_ItmBname AS ItmBname
	FROM
		[@PS_PP001H] PS_PP001H 
		LEFT JOIN [@PS_PP001L] PS_PP001L ON PS_PP001H.Code = PS_PP001L.Code
	WHERE
		(@ItmBsort = '' OR U_ItmBsort = @ItmBsort)
END