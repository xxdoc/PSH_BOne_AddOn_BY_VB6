IF OBJECT_ID('PS_PP030_03') IS NOT NULL
BEGIN
	DROP PROC PS_PP030_03
END
GO
CREATE PROC PS_PP030_03
(
	@BaseItmBsort NVARCHAR(10), --ǰ���з� ����
	@BaseInputType NVARCHAR(10), --���Ա���,�����/��ũ��
	@ItemCode NVARCHAR(20), --ǰ���ڵ�
	@ItmsGrpCod NVARCHAR(10), --ǰ�񱸺�
	@ItmBsort NVARCHAR(10), --ǰ���з�
	@ItmMsort NVARCHAR(10) --ǰ���ߺз�	
)
AS
BEGIN
	IF @BaseItmBsort IN('101','102','105','106') --����,��ǰ,���,���� �ϰ��
	BEGIN
		SELECT
			OITM.ItemCode AS ItemCode,
			OITM.ItemName AS ItemName,
			OITM.ItmsGrpCod AS ItemGpCd,
			'' AS BatchNum,			
			'' AS Weight
		FROM
			[OITM] OITM
		WHERE
			OITM.ItmsGrpCod IN('101','104') --��ǰ �����
			AND (@ItmsGrpCod = '' OR OITM.ItmsGrpCod = @ItmsGrpCod)
			AND (@ItmBsort = '' OR OITM.U_ItmBsort = @ItmBsort)
			AND (@ItmMsort = '' OR OITM.U_ItmMsort = @ItmMsort)
			AND (@ItemCode = '' OR OITM.ItemCode = @ItemCode)
	END
	ELSE IF @BaseItmBsort IN('104') --��Ƽ
	BEGIN
		SELECT
			OIBT.ItemCode AS ItemCode,
			OITM.ItemName AS ItemName,
			OITM.ItmsGrpCod AS ItemGpCd,
			OIBT.BatchNum AS BatchNum,			
			OIBT.Quantity AS Weight
		FROM
			[OIBT] OIBT
			LEFT JOIN [OBTN] OBTN ON OIBT.ItemCode = OBTN.ItemCode AND OIBT.BatchNum = OBTN.DistNumber
			LEFT JOIN [OITM] OITM ON OIBT.ItemCode = OITM.ItemCode
		WHERE
			OITM.ItmsGrpCod IN('101','104') --��ǰ �����
			AND (@ItmsGrpCod = '' OR OITM.ItmsGrpCod = @ItmsGrpCod)
			AND (@ItmBsort = '' OR OITM.U_ItmBsort = @ItmBsort)
			AND (@ItmMsort = '' OR OITM.U_ItmMsort = @ItmMsort)
			AND (@ItemCode = '' OR OITM.ItemCode = @ItemCode)
	END
	ELSE IF @BaseItmBsort IN('107') --���庣�
	BEGIN
		IF @BaseInputType = '20' --�����
		BEGIN
			SELECT
				OIBT.ItemCode AS ItemCode,
				OITM.ItemName AS ItemName,
				OITM.ItmsGrpCod AS ItemGpCd,
				OIBT.BatchNum AS BatchNum,				
				OIBT.Quantity AS Weight
			FROM
				[OIBT] OIBT
				LEFT JOIN [OBTN] OBTN ON OIBT.ItemCode = OBTN.ItemCode AND OIBT.BatchNum = OBTN.DistNumber
				LEFT JOIN [OITM] OITM ON OIBT.ItemCode = OITM.ItemCode
			WHERE
				OITM.ItmsGrpCod IN('101','104') --��ǰ �����
				AND (@ItmsGrpCod = '' OR OITM.ItmsGrpCod = @ItmsGrpCod)
				AND (@ItmBsort = '' OR OITM.U_ItmBsort = @ItmBsort)
				AND (@ItmMsort = '' OR OITM.U_ItmMsort = @ItmMsort)
				AND (@ItemCode = '' OR OITM.ItemCode = @ItemCode)
		END
		ELSE IF @BaseInputType = '30' --��ũ��
		BEGIN
			SELECT
				PS_PP030L.U_ItemCode AS ItemCode,
				PS_PP030L.U_ItemName AS ItemName,
				PS_PP030L.U_ItemGpCd AS ItemGpCd,
				PS_PP030L.U_BatchNum AS BatchNum,				
				PS_PP030L.U_Weight AS Weight
			FROM 
				[@PS_PP030L] PS_PP030L
				LEFT JOIN [OITM] OITM ON PS_PP030L.U_ItemCode = OITM.ItemCode
			WHERE
				OITM.ItmsGrpCod IN('101','104') --��ǰ �����
				AND (@ItmsGrpCod = '' OR OITM.ItmsGrpCod = @ItmsGrpCod)
				AND (@ItmBsort = '' OR OITM.U_ItmBsort = @ItmBsort)
				AND (@ItmMsort = '' OR OITM.U_ItmMsort = @ItmMsort)
				AND (@ItemCode = '' OR OITM.ItemCode = @ItemCode)
		END
	END
END