IF OBJECT_ID('PS_SD220_10') IS NOT NULL
BEGIN
	DROP PROC PS_SD220_10
END
GO
--EXEC PS_SD220_10 '5'
CREATE PROC PS_SD220_10
(
	@DocEntry NVARCHAR(100)
)
AS
BEGIN
	CREATE TABLE #TEMP
	(
		[IDX] INT ,
		[VATRegNum]NVARCHAR(100),--등록번호
		[CardName] NVARCHAR(100),--상호
		[RepName] NVARCHAR(100),--대표자명
		[Street] NVARCHAR(100),--주소
		
		[PARAM01] NVARCHAR(100),
		[PARAM02] NVARCHAR(100),
		[PARAM03] NVARCHAR(100),
		[PARAM04] NVARCHAR(100),
		[PARAM05] NVARCHAR(100),
		[PARAM06] NVARCHAR(100),
		[PARAM07] NVARCHAR(100),
		[PARAM08] NVARCHAR(100),
		
		[OutDocEntry] NVARCHAR(100),--출고번호
		[BPLName] NVARCHAR(100), --사업장
		[DocDate] NVARCHAR(100),--거래일자
		[CCardName] NVARCHAR(100),--고객
		[DCardName] NVARCHAR(100),--납품처
		[Comments] NVARCHAR(100),--특기사항		
		
		[ItemCode01]	NVARCHAR(100),
		[ItemName01]	NVARCHAR(100),		
		[Quality01] NVARCHAR(100),--질별
		[Size01] NVARCHAR(100),--규격		
		[ItemType01] NVARCHAR(100),--형태타입
		[Mark01] NVARCHAR(100),--인증
		[Box01]	NUMERIC(19,6),
		[Qty01]	NUMERIC(19,6),
		[Weight01]	NUMERIC(19,6),
		[Sort01] NVARCHAR(100),--구분

		[ItemCode02]	NVARCHAR(100),
		[ItemName02]	NVARCHAR(100),		
		[Quality02] NVARCHAR(100),--질별
		[Size02] NVARCHAR(100),--규격		
		[ItemType02] NVARCHAR(100),--형태타입
		[Mark02] NVARCHAR(100),--인증
		[Box02]	NUMERIC(19,6),
		[Qty02]	NUMERIC(19,6),
		[Weight02]	NUMERIC(19,6),		
		[Sort02] NVARCHAR(100),--구분

		[ItemCode03]	NVARCHAR(100),
		[ItemName03]	NVARCHAR(100),		
		[Quality03] NVARCHAR(100),--질별
		[Size03] NVARCHAR(100),--규격		
		[ItemType03] NVARCHAR(100),--형태타입
		[Mark03] NVARCHAR(100),--인증
		[Box03]	NUMERIC(19,6),
		[Qty03]	NUMERIC(19,6),
		[Weight03]	NUMERIC(19,6),		
		[Sort03] NVARCHAR(100),--구분

		[ItemCode04]	NVARCHAR(100),
		[ItemName04]	NVARCHAR(100),		
		[Quality04] NVARCHAR(100),--질별
		[Size04] NVARCHAR(100),--규격		
		[ItemType04] NVARCHAR(100),--형태타입
		[Mark04] NVARCHAR(100),--인증
		[Box04]	NUMERIC(19,6),
		[Qty04]	NUMERIC(19,6),
		[Weight04]	NUMERIC(19,6),		
		[Sort04] NVARCHAR(100),--구분

		[ItemCode05]	NVARCHAR(100),
		[ItemName05]	NVARCHAR(100),		
		[Quality05] NVARCHAR(100),--질별
		[Size05] NVARCHAR(100),--규격
		[ItemType05] NVARCHAR(100),--형태타입
		[Mark05] NVARCHAR(100),--인증
		[Box05]	NUMERIC(19,6),
		[Qty05]	NUMERIC(19,6),
		[Weight05]	NUMERIC(19,6),		
		[Sort05] NVARCHAR(100),--구분

		[ItemCode06]	NVARCHAR(100),		
		[ItemName06]	NVARCHAR(100),
		[Quality06] NVARCHAR(100),--질별
		[Size06] NVARCHAR(100),--규격		
		[ItemType06] NVARCHAR(100),--형태타입
		[Mark06] NVARCHAR(100),--인증
		[Box06]	NUMERIC(19,6),
		[Qty06]	NUMERIC(19,6),
		[Weight06]	NUMERIC(19,6),		
		[Sort06] NVARCHAR(100),--구분

		[ItemCode07]	NVARCHAR(100),
		[ItemName07]	NVARCHAR(100),		
		[Quality07] NVARCHAR(100),--질별
		[Size07] NVARCHAR(100),--규격	
		[ItemType07] NVARCHAR(100),--형태타입
		[Mark07] NVARCHAR(100),--인증
		[Box07]	NUMERIC(19,6),
		[Qty07]	NUMERIC(19,6),
		[Weight07]	NUMERIC(19,6),		
		[Sort07] NVARCHAR(100),--구분
		
		[ItemCode08]	NVARCHAR(100),
		[ItemName08]	NVARCHAR(100),		
		[Quality08] NVARCHAR(100),--질별
		[Size08] NVARCHAR(100),--규격	
		[ItemType08] NVARCHAR(100),--형태타입
		[Mark08] NVARCHAR(100),--인증
		[Box08]	NUMERIC(19,6),
		[Qty08]	NUMERIC(19,6),
		[Weight08]	NUMERIC(19,6),		
		[Sort08] NVARCHAR(100),--구분
		
		[ItemCode09]	NVARCHAR(100),
		[ItemName09]	NVARCHAR(100),		
		[Quality09] NVARCHAR(100),--질별
		[Size09] NVARCHAR(100),--규격	
		[ItemType09] NVARCHAR(100),--형태타입
		[Mark09] NVARCHAR(100),--인증
		[Box09]	NUMERIC(19,6),
		[Qty09]	NUMERIC(19,6),
		[Weight09]	NUMERIC(19,6),		
		[Sort09] NVARCHAR(100),--구분
		
		[ItemCode10]	NVARCHAR(100),
		[ItemName10]	NVARCHAR(100),		
		[Quality10] NVARCHAR(100),--질별
		[Size10] NVARCHAR(100),--규격	
		[ItemType10] NVARCHAR(100),--형태타입
		[Mark10] NVARCHAR(100),--인증
		[Box10]	NUMERIC(19,6),
		[Qty10]	NUMERIC(19,6),
		[Weight10]	NUMERIC(19,6),		
		[Sort10] NVARCHAR(100),--구분
	)

	DECLARE @IDX INT 
	DECLARE @ILOOPER INT
	DECLARE @VATRegNum NVARCHAR(100)--등록번호
	DECLARE @CardName NVARCHAR(100)--상호
	DECLARE @RepName NVARCHAR(100)--대표자명
	DECLARE @Street NVARCHAR(100)--주소
	DECLARE @Param01 NVARCHAR(100)--인수자
	DECLARE @Param02 NVARCHAR(100)
	DECLARE @Param03 NVARCHAR(100)
	DECLARE @Param04 NVARCHAR(100)
	DECLARE @Param05 NVARCHAR(100)
	DECLARE @Param06 NVARCHAR(100)
	DECLARE @Param07 NVARCHAR(100)
	DECLARE @Param08 NVARCHAR(100)
	
	DECLARE @OutDocEntry NVARCHAR(100)--출고번호
	DECLARE @BPLName NVARCHAR(100) --사업장
	DECLARE @DocDate NVARCHAR(100)--거래일자
	DECLARE @CCardName NVARCHAR(100)--고객
	DECLARE @DCardName NVARCHAR(100)--납품
	DECLARE @Comments NVARCHAR(100)--특기사항
	
	DECLARE @ItemCode	NVARCHAR(100)
	DECLARE @ItemName	NVARCHAR(100)	
	DECLARE @Quality NVARCHAR(100)--질별
	DECLARE @Size NVARCHAR(100)--규격	
	DECLARE @ItemType NVARCHAR(100)--형태타입
	DECLARE @Mark NVARCHAR(100)--인증
	DECLARE @Box	NUMERIC(19,6)
	DECLARE @Qty	NUMERIC(19,6)
	DECLARE @Weight	NUMERIC(19,6)	
	DECLARE @Sort NVARCHAR(100)--구분

	DECLARE CURSOR01 CURSOR FOR
	SELECT
		--헤더
		OBPL.VATRegNum AS VATRegNum
		,OBPL.BPLName AS CardName
		,OBPL.RepName AS RepName
		,OBPL.Address AS Street
		,PS_SD040H.U_CntcCode AS PARAM01
		,PS_SD040H.U_TranCard AS PARAM02
		,PS_SD040H.U_TranCode AS PARAM03
		,PS_SD040H.U_Destin AS PARAM04
		,PS_SD040H.U_TranCost AS PARAM05
		,'' AS PARAM06
		,'' AS PARAM07
		,'' AS PARAM08
		,PS_SD040H.DocEntry AS OutDocEntry
		,OBPL.BPLName AS BPLName
		,CONVERT(NVARCHAR,PS_SD040H.U_DocDate,111) AS DocDate
		,PS_SD040H.U_CardName AS CCardName
		,PS_SD040H.U_DCardNam AS DCardName
		,PS_SD040H.U_Comments AS Comments
		--라인
		,PS_SD040L.U_ItemCode AS ItemCode
		,OITM.ItemName AS ItemName		
		,(SELECT NAME FROM [@PSH_QUALITY] WHERE CODE = OITM.U_Quality) AS Quality
		,OITM.U_Size AS Size		
		,(SELECT NAME FROM [@PSH_SHAPE] WHERE CODE = OITM.U_ItemType) AS ItemType
		,(SELECT NAME FROM [@PSH_MARK] WHERE CODE = OITM.U_Mark) AS Mark
		,PS_SD040L.U_Qty AS Box
		,PS_SD040L.U_Weight AS Qty
		,PS_SD040L.U_Weight * OITM.U_UnWeight AS Weight		
		,'' AS Sort
	FROM
		[@PS_SD040H] PS_SD040H
		LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry 
		LEFT JOIN [OBPL] OBPL ON PS_SD040H.U_BPLId = OBPL.BPLId
		LEFT JOIN [OITM] OITM ON PS_SD040L.U_ItemCode = OITM.ItemCode		
	WHERE
		PS_SD040H.DocEntry = @DocEntry
		AND PS_SD040H.U_BPLId IN('1','4') --창원,서울
		AND (SELECT U_ItmBsort FROM [OITM] WHERE ItemCode = PS_SD040L.U_ItemCode) IN('101','102') --휘팅,부품

	SET @IDX = 1
	SET @ILOOPER = 1
	OPEN CURSOR01
	FETCH NEXT FROM CURSOR01 INTO @VATRegNum,@CardName,@RepName,@Street,@PARAM01,@PARAM02,@PARAM03,@PARAM04,@PARAM05,@PARAM06,@PARAM07,@PARAM08,@OutDocEntry,@BPLName,@DocDate,@CCardName,@DCardName,@Comments,@ItemCode,@ItemName,@Quality,@Size,@ItemType,@Mark,@Box,@Qty,@Weight,@Sort
	WHILE @@FETCH_STATUS = 0
	BEGIN	
		IF @ILOOPER > 10
		BEGIN
			SET @IDX = @IDX + 1
			SET @ILOOPER = 1
		END

		IF @ILOOPER = 1 
		BEGIN
			INSERT INTO #TEMP (IDX,VATRegNum,CardName,RepName,Street,PARAM01,PARAM02,PARAM03,PARAM04,PARAM05,PARAM06,PARAM07,PARAM08,OutDocEntry,BPLName,DocDate,CCardName,DCardName,Comments,ItemCode01,ItemName01,Quality01,Size01,ItemType01,Mark01,Box01,Qty01,Weight01,Sort01)
			VALUES(@IDX,@VATRegNum,@CardName,@RepName,@Street,@PARAM01,@PARAM02,@PARAM03,@PARAM04,@PARAM05,@PARAM06,@PARAM07,@PARAM08,@OutDocEntry,@BPLName,@DocDate,@CCardName,@DCardName,@Comments,@ItemCode,@ItemName,@Quality,@Size,@ItemType,@Mark,@Box,@Qty,@Weight,@Sort)
		END
		ELSE IF @ILOOPER = 2
		BEGIN
			UPDATE #TEMP SET ItemCode02 = @ItemCode ,ItemName02 = @ItemName ,Quality02 = @Quality ,Size02 = @Size ,ItemType02 = @ItemType ,Mark02 = @Mark ,Box02 = @Box ,Qty02 = @Qty ,Weight02 = @Weight ,Sort02 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 3
		BEGIN
			UPDATE #TEMP SET ItemCode03 = @ItemCode ,ItemName03 = @ItemName ,Quality03 = @Quality ,Size03 = @Size ,ItemType03 = @ItemType ,Mark03 = @Mark ,Box03 = @Box ,Qty03 = @Qty ,Weight03 = @Weight ,Sort03 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 4
		BEGIN
			UPDATE #TEMP SET ItemCode04 = @ItemCode ,ItemName04 = @ItemName ,Quality04 = @Quality ,Size04 = @Size ,ItemType04 = @ItemType ,Mark04 = @Mark ,Box04 = @Box ,Qty04 = @Qty ,Weight04 = @Weight ,Sort04 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 5
		BEGIN
			UPDATE #TEMP SET ItemCode05 = @ItemCode ,ItemName05 = @ItemName ,Quality05 = @Quality ,Size05 = @Size ,ItemType05 = @ItemType ,Mark05 = @Mark ,Box05 = @Box ,Qty05 = @Qty ,Weight05 = @Weight ,Sort05 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 6
		BEGIN
			UPDATE #TEMP SET ItemCode06 = @ItemCode ,ItemName06 = @ItemName ,Quality06 = @Quality ,Size06 = @Size ,ItemType06 = @ItemType ,Mark06 = @Mark ,Box06 = @Box ,Qty06 = @Qty ,Weight06 = @Weight ,Sort06 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 7
		BEGIN
			UPDATE #TEMP SET ItemCode07 = @ItemCode ,ItemName07 = @ItemName ,Quality07 = @Quality ,Size07 = @Size ,ItemType07 = @ItemType ,Mark07 = @Mark ,Box07 = @Box ,Qty07 = @Qty ,Weight07 = @Weight ,Sort07 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 8
		BEGIN
			UPDATE #TEMP SET ItemCode08 = @ItemCode ,ItemName08 = @ItemName ,Quality08 = @Quality ,Size08 = @Size ,ItemType08 = @ItemType ,Mark08 = @Mark ,Box08 = @Box ,Qty08 = @Qty ,Weight08 = @Weight ,Sort08 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 9
		BEGIN
			UPDATE #TEMP SET ItemCode09 = @ItemCode ,ItemName09 = @ItemName ,Quality09 = @Quality ,Size09 = @Size ,ItemType09 = @ItemType ,Mark09 = @Mark ,Box09 = @Box ,Qty09 = @Qty ,Weight09 = @Weight ,Sort09 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 10
		BEGIN
			UPDATE #TEMP SET ItemCode10 = @ItemCode ,ItemName10 = @ItemName ,Quality10 = @Quality ,Size10 = @Size ,ItemType10 = @ItemType ,Mark10 = @Mark ,Box10 = @Box ,Qty10 = @Qty ,Weight10 = @Weight ,Sort10 = @Sort WHERE IDX = @IDX
		END
		SET @ILOOPER = @ILOOPER + 1
	FETCH NEXT FROM CURSOR01 INTO @VATRegNum,@CardName,@RepName,@Street,@PARAM01,@PARAM02,@PARAM03,@PARAM04,@PARAM05,@PARAM06,@PARAM07,@PARAM08,@OutDocEntry,@BPLName,@DocDate,@CCardName,@DCardName,@Comments,@ItemCode,@ItemName,@Quality,@Size,@ItemType,@Mark,@Box,@Qty,@Weight,@Sort
	END
	CLOSE CURSOR01
	DEALLOCATE CURSOR01

	SELECT
		IDX,
		VATRegNum,
		CardName,
		RepName,
		Street,
		PARAM01,
		PARAM02,
		PARAM03,
		PARAM04,
		PARAM05,
		PARAM06,
		PARAM07,
		PARAM08,
		OutDocEntry,
		BPLName,
		DocDate,
		CCardName,
		DCardName,
		Comments,		
		ItemCode01,
		ItemName01,		
		Quality01,
		Size01,		
		ItemType01,
		Mark01,
		Box01,
		Qty01,
		Weight01,		
		Sort01,
		ItemCode02,
		ItemName02,		
		Quality02,
		Size02,		
		ItemType02,
		Mark02,
		Box02,
		Qty02,
		Weight02,		
		Sort02,
		ItemCode03,
		ItemName03,		
		Quality03,
		Size03,	
		ItemType03,
		Mark03,
		Box03,
		Qty03,
		Weight03,		
		Sort03,
		ItemCode04,
		ItemName04,		
		Quality04,
		Size04,	
		ItemType04,
		Mark04,
		Box04,
		Qty04,
		Weight04,		
		Sort04,
		ItemCode05,
		ItemName05,		
		Quality05,
		Size05,		
		ItemType05,
		Mark05,
		Box05,
		Qty05,
		Weight05,		
		Sort05,
		ItemCode06,
		ItemName06,		
		Quality06,
		Size06,		
		ItemType06,
		Mark06,
		Box06,
		Qty06,
		Weight06,		
		Sort06,
		ItemCode07,
		ItemName07,		
		Quality07,
		Size07,		
		ItemType07,
		Mark07,
		Box07,
		Qty07,
		Weight07,		
		Sort07,
		ItemCode08,
		ItemName08,		
		Quality08,
		Size08,		
		ItemType08,
		Mark08,
		Box08,
		Qty08,
		Weight08,		
		Sort08,
		ItemCode09,
		ItemName09,		
		Quality09,
		Size09,		
		ItemType09,
		Mark09,
		Box09,
		Qty09,
		Weight09,		
		Sort09,
		ItemCode10,
		ItemName10,		
		Quality10,
		Size10,		
		ItemType10,
		Mark10,
		Box10,
		Qty10,
		Weight10,		
		Sort10,
		(ISNULL(Qty01,0) + ISNULL(Qty02,0) + ISNULL(Qty03,0) + ISNULL(Qty04,0) + ISNULL(Qty05,0) + ISNULL(Qty06,0) + ISNULL(Qty07,0) + ISNULL(Qty08,0) + ISNULL(Qty09,0) + ISNULL(Qty10,0)) AS QtyPageSum,
		(ISNULL(Weight01,0) + ISNULL(Weight02,0) + ISNULL(Weight03,0) + ISNULL(Weight04,0) + ISNULL(Weight05,0) + ISNULL(Weight06,0) + ISNULL(Weight07,0) + ISNULL(Weight08,0) + ISNULL(Weight09,0) + ISNULL(Weight10,0)) AS WeightPageSum,
		0 AS QtyDocSum,
		0 AS WeightDocSum
	FROM 
		[#TEMP] T0
END
 


