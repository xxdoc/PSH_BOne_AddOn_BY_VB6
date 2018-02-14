IF OBJECT_ID('PS_PP540_10') IS NOT NULL
BEGIN
	DROP PROC PS_PP540_10
END
GO
--EXEC PS_PP540_10 '20'
CREATE PROC PS_PP540_10
(
	@DocType NVARCHAR(100),
	@DocEntry NVARCHAR(100)
)
AS
BEGIN
	CREATE TABLE #TEMP
	(
		[IDX] INT ,
		[Doc] NVARCHAR(100),
		[DueDate] NVARCHAR(100),
		[VATRegNum] NVARCHAR(100),
		[CardName] NVARCHAR(100),
		[RepName] NVARCHAR(100),
		[Street] NVARCHAR(100),
		[CVATRegNum] NVARCHAR(100),
		[CCardName] NVARCHAR(100),
		[CRepName] NVARCHAR(100),
		[CStreet] NVARCHAR(100),
		
		[QuantitySum] NUMERIC(19,6),
		[TranCard] NVARCHAR(100),
		[TranCode] NVARCHAR(100),
		[Destin] NVARCHAR(100),
		[TranCost] NVARCHAR(100),
		[CntcCode] NVARCHAR(100),
		[HComments] NVARCHAR(100),
		
		[ItemCode01]	NVARCHAR(100),
		[ItemName01]	NVARCHAR(100),				
		[Size01] NVARCHAR(100),--규격						
		[Unit01] NVARCHAR(100),--규격
		[Quantity01]	NUMERIC(19,6),				
		[LComments01] NVARCHAR(100),--구분

		[ItemCode02]	NVARCHAR(100),
		[ItemName02]	NVARCHAR(100),				
		[Size02] NVARCHAR(100),--규격						
		[Unit02] NVARCHAR(100),--규격
		[Quantity02]	NUMERIC(19,6),				
		[LComments02] NVARCHAR(100),--구분

		[ItemCode03]	NVARCHAR(100),
		[ItemName03]	NVARCHAR(100),				
		[Size03] NVARCHAR(100),--규격						
		[Unit03] NVARCHAR(100),--규격
		[Quantity03]	NUMERIC(19,6),				
		[LComments03] NVARCHAR(100),--구분

		[ItemCode04]	NVARCHAR(100),
		[ItemName04]	NVARCHAR(100),				
		[Size04] NVARCHAR(100),--규격						
		[Unit04] NVARCHAR(100),--규격
		[Quantity04]	NUMERIC(19,6),				
		[LComments04] NVARCHAR(100),--구분

		[ItemCode05]	NVARCHAR(100),
		[ItemName05]	NVARCHAR(100),				
		[Size05] NVARCHAR(100),--규격						
		[Unit05] NVARCHAR(100),--규격
		[Quantity05]	NUMERIC(19,6),				
		[LComments05] NVARCHAR(100),--구분

		[ItemCode06]	NVARCHAR(100),
		[ItemName06]	NVARCHAR(100),				
		[Size06] NVARCHAR(100),--규격						
		[Unit06] NVARCHAR(100),--규격
		[Quantity06]	NUMERIC(19,6),				
		[LComments06] NVARCHAR(100),--구분

		[ItemCode07]	NVARCHAR(100),
		[ItemName07]	NVARCHAR(100),				
		[Size07] NVARCHAR(100),--규격						
		[Unit07] NVARCHAR(100),--규격
		[Quantity07]	NUMERIC(19,6),				
		[LComments07] NVARCHAR(100),--구분
	)
	
	DECLARE @IDX INT 
	DECLARE @ILOOPER INT
	DECLARE @Doc NVARCHAR(100)
	DECLARE @DueDate NVARCHAR(100)
	DECLARE @VATRegNum NVARCHAR(100)
	DECLARE @CardName NVARCHAR(100)
	DECLARE @RepName NVARCHAR(100)
	DECLARE @Street NVARCHAR(100)
	DECLARE @CVATRegNum NVARCHAR(100)
	DECLARE @CCardName NVARCHAR(100)
	DECLARE @CRepName NVARCHAR(100)
	DECLARE @CStreet NVARCHAR(100)

	DECLARE @QuantitySum NUMERIC(19,6)
	DECLARE @TranCard NVARCHAR(100)
	DECLARE @TranCode NVARCHAR(100)
	DECLARE @Destin NVARCHAR(100)
	DECLARE @TranCost NVARCHAR(100)
	DECLARE @CntcCode NVARCHAR(100)
	DECLARE @HComments NVARCHAR(100)

	DECLARE @ItemCode	NVARCHAR(100)
	DECLARE @ItemName	NVARCHAR(100)				
	DECLARE @Size NVARCHAR(100)--규격						
	DECLARE @Unit NVARCHAR(100)--규격
	DECLARE @Quantity	NUMERIC(19,6)
	DECLARE @LComments NVARCHAR(100)--구분
	
	IF(@DocType = '납품')
	BEGIN
		DECLARE CURSOR01 CURSOR FOR
		SELECT		
			CONVERT(NVARCHAR(100),PS_SD040H.DocEntry) AS Doc, --출고번호
			CONVERT(NVARCHAR(100),PS_SD040H.U_DueDate,111) AS DueDate, --출고일자
			CONVERT(NVARCHAR(100),OBPL.VATRegNum) AS VATRegNum,
			CONVERT(NVARCHAR(100),OBPL.BPLName) AS CardName,		
			CONVERT(NVARCHAR(100),OBPL.RepName) AS RepName,
			CONVERT(NVARCHAR(100),OBPL.Address) AS Street,
			CONVERT(NVARCHAR(100),OCRD.VATRegNum) AS CVATRegNum,
			CONVERT(NVARCHAR(100),OCRD.CardName) AS CCardName,
			CONVERT(NVARCHAR(100),OCRD.RepName) AS CRepName,
			CONVERT(NVARCHAR(100),OCRD.Address) AS CStreet,

			CONVERT(NUMERIC(19,6),(SELECT SUM(U_Weight) FROM [@PS_SD040L] WHERE DocEntry = @DocEntry)) AS QuantitySum,		
			CONVERT(NVARCHAR(100),PS_SD040H.U_TranCard) AS TranCard, --운송업체
			CONVERT(NVARCHAR(100),PS_SD040H.U_TranCode) AS TranCode, --차량번호
			CONVERT(NVARCHAR(100),PS_SD040H.U_Destin) AS Destin, --도착지
			CONVERT(NUMERIC(19,6),PS_SD040H.U_TranCost) AS TranCost, --운임
			CONVERT(NVARCHAR(100),PS_SD040H.U_CntcCode) AS CntcCode, --담당자
			CONVERT(NVARCHAR(100),PS_SD040H.U_Comments) AS HComments,		
			
			CONVERT(NVARCHAR(100),PS_SD040L.U_ItemCode) AS ItemCode,
			CONVERT(NVARCHAR(100),PS_SD040L.U_ItemName) AS ItemName,
			CONVERT(NVARCHAR(100),OITM.U_Size) AS Size,
			CONVERT(NVARCHAR(100),OITM.InvntryUom) AS Unit,
			CONVERT(NUMERIC(19,6),PS_SD040L.U_Weight) AS Quantity,
			CONVERT(NVARCHAR(100),PS_SD040L.U_Comments) AS LComments		
		FROM
			[@PS_SD040H] PS_SD040H
			LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry
			LEFT JOIN [OBPL] OBPL ON PS_SD040H.U_BPLId = OBPL.BPLId
			LEFT JOIN [OCRD] OCRD ON PS_SD040H.U_CardCode = OCRD.CardCode
			LEFT JOIN [OITM] OITM ON PS_SD040L.U_ItemCode = OITM.ItemCode
		WHERE
			PS_SD040H.DocEntry = @DocEntry
			AND PS_SD040H.U_BPLId IN('2') --동래
			AND (SELECT U_ItmBsort FROM [OITM] WHERE ItemCode = PS_SD040L.U_ItemCode) IN('105','106') --기계,몰드		
	END
	ELSE IF(@DocType = '선출')
	BEGIN
		DECLARE CURSOR01 CURSOR FOR
		SELECT		
			CONVERT(NVARCHAR(100),PS_SD030H.DocEntry) AS Doc, --출고번호
			CONVERT(NVARCHAR(100),PS_SD030H.U_DueDate,111) AS DueDate, --출고일자
			CONVERT(NVARCHAR(100),OBPL.VATRegNum) AS VATRegNum,
			CONVERT(NVARCHAR(100),OBPL.BPLName) AS CardName,		
			CONVERT(NVARCHAR(100),OBPL.RepName) AS RepName,
			CONVERT(NVARCHAR(100),OBPL.Address) AS Street,
			CONVERT(NVARCHAR(100),OCRD.VATRegNum) AS CVATRegNum,
			CONVERT(NVARCHAR(100),OCRD.CardName) AS CCardName,
			CONVERT(NVARCHAR(100),OCRD.RepName) AS CRepName,
			CONVERT(NVARCHAR(100),OCRD.Address) AS CStreet,

			CONVERT(NUMERIC(19,6),(SELECT SUM(U_Weight) FROM [@PS_SD030L] WHERE DocEntry = @DocEntry)) AS QuantitySum,		
			CONVERT(NVARCHAR(100),PS_SD030H.U_TranCard) AS TranCard, --운송업체
			CONVERT(NVARCHAR(100),PS_SD030H.U_TranCode) AS TranCode, --차량번호
			CONVERT(NVARCHAR(100),PS_SD030H.U_Destin) AS Destin, --도착지
			CONVERT(NUMERIC(19,6),PS_SD030H.U_TranCost) AS TranCost, --운임
			CONVERT(NVARCHAR(100),PS_SD030H.U_CntcCode) AS CntcCode, --담당자
			CONVERT(NVARCHAR(100),PS_SD030H.U_Comments) AS HComments,		
			
			CONVERT(NVARCHAR(100),PS_SD030L.U_ItemCode) AS ItemCode,
			CONVERT(NVARCHAR(100),PS_SD030L.U_ItemName) AS ItemName,
			CONVERT(NVARCHAR(100),OITM.U_Size) AS Size,
			CONVERT(NVARCHAR(100),OITM.InvntryUom) AS Unit,
			CONVERT(NUMERIC(19,6),PS_SD030L.U_Weight) AS Quantity,
			CONVERT(NVARCHAR(100),PS_SD030L.U_Comments) AS LComments		
		FROM
			[@PS_SD030H] PS_SD030H
			LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry
			LEFT JOIN [OBPL] OBPL ON PS_SD030H.U_BPLId = OBPL.BPLId
			LEFT JOIN [OCRD] OCRD ON PS_SD030H.U_CardCode = OCRD.CardCode
			LEFT JOIN [OITM] OITM ON PS_SD030L.U_ItemCode = OITM.ItemCode
		WHERE
			PS_SD030H.DocEntry = @DocEntry
			AND PS_SD030H.U_BPLId IN('2') --동래
			AND (SELECT U_ItmBsort FROM [OITM] WHERE ItemCode = PS_SD030L.U_ItemCode) IN('105','106') --기계,몰드		
			AND PS_SD030H.U_DocType = '2' --선출등록
			AND PS_SD030H.U_ProgStat = '3' --납품상태
	END
	
	SET @IDX = 1
	SET @ILOOPER = 1
	OPEN CURSOR01
	FETCH NEXT FROM CURSOR01 INTO @Doc,@DueDate,@VATRegNum,@CardName,@RepName,@Street,@CVATRegNum,@CCardName,@CRepName,@CStreet,@QuantitySum,@TranCard,@TranCode,@Destin,@TranCost,@CntcCode,@HComments,@ItemCode,@ItemName,@Size,@Unit,@Quantity,@LComments
	WHILE @@FETCH_STATUS = 0
	BEGIN	
		IF @ILOOPER > 7
		BEGIN
			SET @IDX = @IDX + 1
			SET @ILOOPER = 1
		END

		IF @ILOOPER = 1 
		BEGIN
			INSERT INTO #TEMP (IDX,Doc,DueDate,VATRegNum,CardName,RepName,Street,CVATRegNum,CCardName,CRepName,CStreet,QuantitySum,TranCard,TranCode,Destin,TranCost,CntcCode,HComments,ItemCode01,ItemName01,Size01,Unit01,Quantity01,LComments01)
			VALUES(@IDX,@Doc,@DueDate,@VATRegNum,@CardName,@RepName,@Street,@CVATRegNum,@CCardName,@CRepName,@CStreet,@QuantitySum,@TranCard,@TranCode,@Destin,@TranCost,@CntcCode,@HComments,@ItemCode,@ItemName,@Size,@Unit,@Quantity,@LComments)
		END
		ELSE IF @ILOOPER = 2
		BEGIN
			UPDATE #TEMP SET ItemCode02 = @ItemCode ,ItemName02 = @ItemName ,Size02 = @Size ,Unit02 = @Unit ,Quantity02 = @Quantity ,LComments02 = @LComments WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 3
		BEGIN
			UPDATE #TEMP SET ItemCode03 = @ItemCode ,ItemName03 = @ItemName ,Size03 = @Size ,Unit03 = @Unit ,Quantity03 = @Quantity ,LComments03 = @LComments WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 4
		BEGIN
			UPDATE #TEMP SET ItemCode04 = @ItemCode ,ItemName04 = @ItemName ,Size04 = @Size ,Unit04 = @Unit ,Quantity04 = @Quantity ,LComments04 = @LComments WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 5
		BEGIN
			UPDATE #TEMP SET ItemCode05 = @ItemCode ,ItemName05 = @ItemName ,Size05 = @Size ,Unit05 = @Unit ,Quantity05 = @Quantity ,LComments05 = @LComments WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 6
		BEGIN
			UPDATE #TEMP SET ItemCode06 = @ItemCode ,ItemName06 = @ItemName ,Size06 = @Size ,Unit06 = @Unit ,Quantity06 = @Quantity ,LComments06 = @LComments WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 7
		BEGIN
			UPDATE #TEMP SET ItemCode07 = @ItemCode ,ItemName07 = @ItemName ,Size07 = @Size ,Unit07 = @Unit ,Quantity07 = @Quantity ,LComments07 = @LComments WHERE IDX = @IDX
		END
		
		SET @ILOOPER = @ILOOPER + 1
	FETCH NEXT FROM CURSOR01 INTO @Doc,@DueDate,@VATRegNum,@CardName,@RepName,@Street,@CVATRegNum,@CCardName,@CRepName,@CStreet,@QuantitySum,@TranCard,@TranCode,@Destin,@TranCost,@CntcCode,@HComments,@ItemCode,@ItemName,@Size,@Unit,@Quantity,@LComments
	END
	CLOSE CURSOR01
	DEALLOCATE CURSOR01
	
	SELECT		
		Doc,
		DueDate,
		VATRegNum,
		CardName,
		RepName,
		Street,
		CVATRegNum,
		CCardName,
		CRepName,
		CStreet,

		QuantitySum,
		TranCard,
		TranCode,
		Destin,
		TranCost,
		CntcCode,
		HComments,
		
		ItemCode01,
		ItemName01,				
		Size01,--규격						
		Unit01,--규격
		Quantity01,				
		LComments01,--구분

		ItemCode02,
		ItemName02,				
		Size02,--규격						
		Unit02,--규격
		Quantity02,				
		LComments02,--구분

		ItemCode03,
		ItemName03,				
		Size03,--규격						
		Unit03,--규격
		Quantity03,				
		LComments03,--구분

		ItemCode04,
		ItemName04,				
		Size04,--규격						
		Unit04,--규격
		Quantity04,				
		LComments04,--구분

		ItemCode05,
		ItemName05,				
		Size05,--규격						
		Unit05,--규격
		Quantity05,				
		LComments05,--구분

		ItemCode06,
		ItemName06,				
		Size06,--규격						
		Unit06,--규격
		Quantity06,				
		LComments06,--구분

		ItemCode07,
		ItemName07,				
		Size07,--규격						
		Unit07,--규격
		Quantity07,				
		LComments07--구분
	FROM
		[#TEMP]
END