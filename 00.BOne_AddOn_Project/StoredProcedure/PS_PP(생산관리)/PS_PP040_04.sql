IF OBJECT_ID('PS_PP040_04') IS NOT NULL
BEGIN
	DROP PROC PS_PP040_04
END
GO
--EXEC PS_PP040_04 '6-2'
--�ش����ν����� ��������Ʈ�� �����ϱ�� �Ͽ��� ������� ����������
CREATE PROC PS_PP040_04 --�ش������ �ٷ��հ������� Y,N
(	
	@OrdMgNum NVARCHAR(100)
)
AS
BEGIN
	DECLARE @CpCode NVARCHAR(100)
	SET @CpCode =
	(SELECT
		TOP 1 
		PS_PP030M.U_CpCode
	FROM
		[@PS_PP030M] PS_PP030M
		LEFT JOIN
		(SELECT
			PS_PP030H.DocEntry AS DocEntry,
			PS_PP030M.U_Sequence AS Sequence
		FROM
			[@PS_PP030H] PS_PP030H
			LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry
		WHERE
			CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = @OrdMgNum
		) CURRENTROW ON CURRENTROW.DocEntry = PS_PP030M.DocEntry
	WHERE
		PS_PP030M.U_Sequence > CURRENTROW.Sequence
	ORDER BY
		U_Sequence ASC
	)
	SELECT CASE WHEN @CpCode = 'CP30112' THEN 'Y' ELSE 'N' END --�ٷ������̸� �ް����� �ٷ�������
END