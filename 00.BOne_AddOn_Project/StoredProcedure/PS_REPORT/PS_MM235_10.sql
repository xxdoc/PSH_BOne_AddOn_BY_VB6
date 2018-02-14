IF OBJECT_ID('PS_MM235_10') IS NOT NULL
BEGIN
	DROP PROC PS_MM235_10
END
GO
--EXEC PS_MM235_10 '2'
CREATE PROC PS_MM235_10
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
		
		[PARAM01] NVARCHAR(100),--인수자				
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
		[Purpose] NVARCHAR(100),--목적
		[Comments] NVARCHAR(100),--특기사항
		
		[ItemCode01]	NVARCHAR(100),
		[ItemName01_1]	NVARCHAR(100),
		[ItemName01_2]	NVARCHAR(100),
		[Quality01] NVARCHAR(100),--질별
		[Size01_1] NVARCHAR(100),--규격
		[Size01_2] NVARCHAR(100),--규격2
		[ItemType01] NVARCHAR(100),--형태타입
		[Mark01] NVARCHAR(100),--인증
		[Qty01]	NUMERIC(19,6),
		[Weight01]	NUMERIC(19,6),
		[BatchNum01] NVARCHAR(100),--배치번호
		[Sort01] NVARCHAR(100),--구분

		[ItemCode02]	NVARCHAR(100),
		[ItemName02_1]	NVARCHAR(100),
		[ItemName02_2]	NVARCHAR(100),
		[Quality02] NVARCHAR(100),--질별
		[Size02_1] NVARCHAR(100),--규격
		[Size02_2] NVARCHAR(100),--규격2
		[ItemType02] NVARCHAR(100),--형태타입
		[Mark02] NVARCHAR(100),--인증
		[Qty02]	NUMERIC(19,6),
		[Weight02]	NUMERIC(19,6),
		[BatchNum02] NVARCHAR(100),--배치번호
		[Sort02] NVARCHAR(100),--구분

		[ItemCode03]	NVARCHAR(100),
		[ItemName03_1]	NVARCHAR(100),
		[ItemName03_2]	NVARCHAR(100),
		[Quality03] NVARCHAR(100),--질별
		[Size03_1] NVARCHAR(100),--규격
		[Size03_2] NVARCHAR(100),--규격2
		[ItemType03] NVARCHAR(100),--형태타입
		[Mark03] NVARCHAR(100),--인증
		[Qty03]	NUMERIC(19,6),
		[Weight03]	NUMERIC(19,6),
		[BatchNum03] NVARCHAR(100),--배치번호
		[Sort03] NVARCHAR(100),--구분

		[ItemCode04]	NVARCHAR(100),
		[ItemName04_1]	NVARCHAR(100),
		[ItemName04_2]	NVARCHAR(100),
		[Quality04] NVARCHAR(100),--질별
		[Size04_1] NVARCHAR(100),--규격
		[Size04_2] NVARCHAR(100),--규격2
		[ItemType04] NVARCHAR(100),--형태타입
		[Mark04] NVARCHAR(100),--인증
		[Qty04]	NUMERIC(19,6),
		[Weight04]	NUMERIC(19,6),
		[BatchNum04] NVARCHAR(100),--배치번호
		[Sort04] NVARCHAR(100),--구분

		[ItemCode05]	NVARCHAR(100),
		[ItemName05_1]	NVARCHAR(100),
		[ItemName05_2]	NVARCHAR(100),
		[Quality05] NVARCHAR(100),--질별
		[Size05_1] NVARCHAR(100),--규격
		[Size05_2] NVARCHAR(100),--규격2
		[ItemType05] NVARCHAR(100),--형태타입
		[Mark05] NVARCHAR(100),--인증
		[Qty05]	NUMERIC(19,6),
		[Weight05]	NUMERIC(19,6),
		[BatchNum05] NVARCHAR(100),--배치번호
		[Sort05] NVARCHAR(100),--구분

		[ItemCode06]	NVARCHAR(100),
		[ItemName06_1]	NVARCHAR(100),
		[ItemName06_2]	NVARCHAR(100),
		[Quality06] NVARCHAR(100),--질별
		[Size06_1] NVARCHAR(100),--규격
		[Size06_2] NVARCHAR(100),--규격2
		[ItemType06] NVARCHAR(100),--형태타입
		[Mark06] NVARCHAR(100),--인증
		[Qty06]	NUMERIC(19,6),
		[Weight06]	NUMERIC(19,6),
		[BatchNum06] NVARCHAR(100),--배치번호
		[Sort06] NVARCHAR(100),--구분

		[ItemCode07]	NVARCHAR(100),
		[ItemName07_1]	NVARCHAR(100),
		[ItemName07_2]	NVARCHAR(100),
		[Quality07] NVARCHAR(100),--질별
		[Size07_1] NVARCHAR(100),--규격
		[Size07_2] NVARCHAR(100),--규격2
		[ItemType07] NVARCHAR(100),--형태타입
		[Mark07] NVARCHAR(100),--인증
		[Qty07]	NUMERIC(19,6),
		[Weight07]	NUMERIC(19,6),
		[BatchNum07] NVARCHAR(100),--배치번호
		[Sort07] NVARCHAR(100),--구분
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
	DECLARE @Purpose NVARCHAR(100)--목적
	DECLARE @Comments NVARCHAR(100)--특기사항
	
	DECLARE @ItemCode	NVARCHAR(100)
	DECLARE @ItemName_1	NVARCHAR(100)
	DECLARE @ItemName_2	NVARCHAR(100)
	DECLARE @Quality NVARCHAR(100)--질별
	DECLARE @Size_1 NVARCHAR(100)--규격
	DECLARE @Size_2 NVARCHAR(100)--규격2
	DECLARE @ItemType NVARCHAR(100)--형태타입
	DECLARE @Mark NVARCHAR(100)--인증
	DECLARE @Qty	NUMERIC(19,6)
	DECLARE @Weight	NUMERIC(19,6)
	DECLARE @BatchNum NVARCHAR(100)--배치번호
	DECLARE @Sort NVARCHAR(100)--구분

	DECLARE CURSOR01 CURSOR FOR
	SELECT
		--헤더
		OBPL.VATRegNum AS VATRegNum
		,OBPL.BPLName AS CardName
		,OBPL.RepName AS RepName
		,OBPL.Address AS Street
		,PS_MM130H.U_CntcCode AS PARAM01
		,PS_MM130H.U_ShipTo AS PARAM02
		,PS_MM130H.U_CarNo AS PARAM03
		,PS_MM130H.U_ArrivePl AS PARAM04
		,PS_MM130H.U_Fare AS PARAM05
		,'' AS PARAM06
		,'' AS PARAM07
		,'' AS PARAM08
		,PS_MM130H.DocEntry AS OutDocEntry
		,OBPL.BPLName AS BPLName
		,CONVERT(NVARCHAR,PS_MM130H.U_DocDate,111) AS DocDate
		,PS_MM130H.U_CardName AS CCardName
		,'' AS Purpose
		,PS_MM130H.U_Comments AS Comments
		--라인
		,PS_MM130L.U_OutItmCd AS ItemCode
		,OITM.ItemName AS ItemName_1
		,PS_MM130L.U_ItemName AS ItemName_2
		,(SELECT NAME FROM [@PSH_QUALITY] WHERE CODE = OITM.U_Quality) AS Quality
		,OITM.U_Size AS Size_1
		,OITM.U_Size AS Size_2
		,(SELECT NAME FROM [@PSH_SHAPE] WHERE CODE = OITM.U_ItemType) AS ItemType
		,(SELECT NAME FROM [@PSH_MARK] WHERE CODE = OITM.U_Mark) AS Mark
		,PS_MM130L.U_OutQty AS Qty
		,PS_MM130L.U_OutWt AS Weight
		,'' AS BatchNum
		,'' AS Sort
	FROM 
		[@PS_MM130H] PS_MM130H
		LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry
		LEFT JOIN [OBPL] OBPL ON PS_MM130H.U_BPLId = OBPL.BPLId
		LEFT JOIN [OITM] OITM ON PS_MM130L.U_OutItmCd = OITM.ItemCode		
	WHERE
		PS_MM130H.DocEntry = @DocEntry

	SET @IDX = 1
	SET @ILOOPER = 1
	OPEN CURSOR01
	FETCH NEXT FROM CURSOR01 INTO @VATRegNum,@CardName,@RepName,@Street,@PARAM01,@PARAM02,@PARAM03,@PARAM04,@PARAM05,@PARAM06,@PARAM07,@PARAM08,@OutDocEntry,@BPLName,@DocDate,@CCardName,@Purpose,@Comments,@ItemCode,@ItemName_1,@ItemName_2,@Quality,@Size_1,@Size_2,@ItemType,@Mark,@Qty,@Weight,@BatchNum,@Sort
	WHILE @@FETCH_STATUS = 0
	BEGIN	
		IF @ILOOPER > 7
		BEGIN
			SET @IDX = @IDX + 1
			SET @ILOOPER = 1
		END

		IF @ILOOPER = 1 
		BEGIN
			INSERT INTO #TEMP (IDX,VATRegNum,CardName,RepName,Street,PARAM01,PARAM02,PARAM03,PARAM04,PARAM05,PARAM06,PARAM07,PARAM08,OutDocEntry,BPLName,DocDate,CCardName,Purpose,Comments,ItemCode01,ItemName01_1,ItemName01_2,Quality01,Size01_1,Size01_2,ItemType01,Mark01,Qty01,Weight01,BatchNum01,Sort01)
			VALUES(@IDX,@VATRegNum,@CardName,@RepName,@Street,@PARAM01,@PARAM02,@PARAM03,@PARAM04,@PARAM05,@PARAM06,@PARAM07,@PARAM08,@OutDocEntry,@BPLName,@DocDate,@CCardName,@Purpose,@Comments,@ItemCode,@ItemName_1,@ItemName_2,@Quality,@Size_1,@Size_2,@ItemType,@Mark,@Qty,@Weight,@BatchNum,@Sort)
		END
		ELSE IF @ILOOPER = 2
		BEGIN
			UPDATE #TEMP SET ItemCode02 = @ItemCode ,ItemName02_1 = @ItemName_1 ,ItemName02_2 = @ItemName_2 ,Quality02 = @Quality ,Size02_1 = @Size_1 ,Size02_2 = @Size_2 ,ItemType02 = @ItemType ,Mark02 = @Mark ,Qty02 = @Qty ,Weight02 = @Weight ,BatchNum02 = @BatchNum ,Sort02 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 3
		BEGIN
			UPDATE #TEMP SET ItemCode03 = @ItemCode ,ItemName03_1 = @ItemName_1 ,ItemName03_2 = @ItemName_2 ,Quality03 = @Quality ,Size03_1 = @Size_1 ,Size03_2 = @Size_2 ,ItemType03 = @ItemType ,Mark03 = @Mark ,Qty03 = @Qty ,Weight03 = @Weight ,BatchNum03 = @BatchNum ,Sort03 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 4
		BEGIN
			UPDATE #TEMP SET ItemCode04 = @ItemCode ,ItemName04_1 = @ItemName_1 ,ItemName04_2 = @ItemName_2 ,Quality04 = @Quality ,Size04_1 = @Size_1 ,Size04_2 = @Size_2 ,ItemType04 = @ItemType ,Mark04 = @Mark ,Qty04 = @Qty ,Weight04 = @Weight ,BatchNum04 = @BatchNum ,Sort04 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 5
		BEGIN
			UPDATE #TEMP SET ItemCode05 = @ItemCode ,ItemName05_1 = @ItemName_1 ,ItemName05_2 = @ItemName_2 ,Quality05 = @Quality ,Size05_1 = @Size_1 ,Size05_2 = @Size_2 ,ItemType05 = @ItemType ,Mark05 = @Mark ,Qty05 = @Qty ,Weight05 = @Weight ,BatchNum05 = @BatchNum ,Sort05 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 6
		BEGIN
			UPDATE #TEMP SET ItemCode06 = @ItemCode ,ItemName06_1 = @ItemName_1 ,ItemName06_2 = @ItemName_2 ,Quality06 = @Quality ,Size06_1 = @Size_1 ,Size06_2 = @Size_2 ,ItemType06 = @ItemType ,Mark06 = @Mark ,Qty06 = @Qty ,Weight06 = @Weight ,BatchNum06 = @BatchNum ,Sort06 = @Sort WHERE IDX = @IDX
		END
		ELSE IF @ILOOPER = 7
		BEGIN
			UPDATE #TEMP SET ItemCode07 = @ItemCode ,ItemName07_1 = @ItemName_1 ,ItemName07_2 = @ItemName_2 ,Quality07 = @Quality ,Size07_1 = @Size_1 ,Size07_2 = @Size_2 ,ItemType07 = @ItemType ,Mark07 = @Mark ,Qty07 = @Qty ,Weight07 = @Weight ,BatchNum07 = @BatchNum ,Sort07 = @Sort WHERE IDX = @IDX
		END
		SET @ILOOPER = @ILOOPER + 1
	FETCH NEXT FROM CURSOR01 INTO @VATRegNum,@CardName,@RepName,@Street,@PARAM01,@PARAM02,@PARAM03,@PARAM04,@PARAM05,@PARAM06,@PARAM07,@PARAM08,@OutDocEntry,@BPLName,@DocDate,@CCardName,@Purpose,@Comments,@ItemCode,@ItemName_1,@ItemName_2,@Quality,@Size_1,@Size_2,@ItemType,@Mark,@Qty,@Weight,@BatchNum,@Sort
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
		Purpose,
		Comments,		
		ItemCode01,
		ItemName01_1,
		ItemName01_2,
		Quality01,
		Size01_1,
		Size01_2,
		ItemType01,
		Mark01,
		Qty01,
		Weight01,
		BatchNum01,
		Sort01,
		ItemCode02,
		ItemName02_1,
		ItemName02_2,
		Quality02,
		Size02_1,
		Size02_2,
		ItemType02,
		Mark02,
		Qty02,
		Weight02,
		BatchNum02,
		Sort02,
		ItemCode03,
		ItemName03_1,
		ItemName03_2,
		Quality03,
		Size03_1,
		Size03_2,
		ItemType03,
		Mark03,
		Qty03,
		Weight03,
		BatchNum03,
		Sort03,
		ItemCode04,
		ItemName04_1,
		ItemName04_2,
		Quality04,
		Size04_1,
		Size04_2,
		ItemType04,
		Mark04,
		Qty04,
		Weight04,
		BatchNum04,
		Sort04,
		ItemCode05,
		ItemName05_1,
		ItemName05_2,
		Quality05,
		Size05_1,
		Size05_2,
		ItemType05,
		Mark05,
		Qty05,
		Weight05,
		BatchNum05,
		Sort05,
		ItemCode06,
		ItemName06_1,
		ItemName06_2,
		Quality06,
		Size06_1,
		Size06_2,
		ItemType06,
		Mark06,
		Qty06,
		Weight06,
		BatchNum06,
		Sort06,
		ItemCode07,
		ItemName07_1,
		ItemName07_2,
		Quality07,
		Size07_1,
		Size07_2,
		ItemType07,
		Mark07,
		Qty07,
		Weight07,
		BatchNum07,
		Sort07,
		(ISNULL(Qty01,0) + ISNULL(Qty02,0) + ISNULL(Qty03,0) + ISNULL(Qty04,0) + ISNULL(Qty05,0) + ISNULL(Qty06,0) + ISNULL(Qty07,0)) AS QtyPageSum,
		(ISNULL(Weight01,0) + ISNULL(Weight02,0) + ISNULL(Weight03,0) + ISNULL(Weight04,0) + ISNULL(Weight05,0) + ISNULL(Weight06,0) + ISNULL(Weight07,0)) AS WeightPageSum,
		0 AS QtyDocSum,
		0 AS WeightDocSum
	FROM 
		[#TEMP] T0
END
 


