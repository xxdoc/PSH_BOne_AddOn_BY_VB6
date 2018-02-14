IF OBJECT_ID('PS_PP080_03') IS NOT NULL
BEGIN
	DROP PROC PS_PP080_03
END
GO
--EXEC PS_PP080_03 '3'
CREATE PROC PS_PP080_03
(
	@PP080HNo NVARCHAR(10)
)
AS
BEGIN
	DECLARE @DocDate DATETIME
	DECLARE @CntcCode NVARCHAR(100)
	DECLARE @CntcName NVARCHAR(100)
	DECLARE @PP030No NVARCHAR(100)
	DECLARE @PP030HNo NVARCHAR(100)
	DECLARE @PP030MNo NVARCHAR(100)
	DECLARE @PQty NUMERIC(19,6) --�������
	DECLARE @PWeight NUMERIC(19,6) --�����߷�
	DECLARE @YQty NUMERIC(19,6) --�հݼ���
	DECLARE @YWeight NUMERIC(19,6) --�հ��߷�
	DECLARE @NQty NUMERIC(19,6) --�ҷ�����
	DECLARE @NWeight NUMERIC(19,6) --�ҷ��߷�
	
	DECLARE CURSOR1 SCROLL CURSOR FOR
	SELECT
		PS_PP080H.U_DocDate,
		PS_PP080H.U_CntcCode,
		PS_PP080H.U_CntcName,
		PS_PP080L.U_PP030No,
		PS_PP080L.U_PP030HNo,
		PS_PP080L.U_PP030MNo,
		PS_PP080L.U_PQty,
		PS_PP080L.U_PWeight,
		PS_PP080L.U_YQTY,
		PS_PP080L.U_YWeight,
		PS_PP080L.U_NQty,
		PS_PP080L.U_NWeight
	FROM
		[@PS_PP080H] PS_PP080H
		LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry
	WHERE
		PS_PP080H.DocEntry = @PP080HNo
		AND PS_PP080L.U_OrdGbn IN('101','102') --����,��ǰ�ǰ�츸 ó��
		
	OPEN CURSOR1
	FETCH NEXT FROM CURSOR1 INTO @DocDate,@CntcCode,@CntcName,@PP030No,@PP030HNo,@PP030MNo,@PQty,@PWeight,@YQty,@YWeight,@NQty,@NWeight
	WHILE (@@FETCH_STATUS = 0) 
	BEGIN
		DECLARE @LineId INT
		DECLARE @DocEntry INT
		SET @DocEntry = (SELECT AUTOKEY FROM [ONNM] WHERE OBJECTCODE = 'PS_PP040')
		UPDATE ONNM SET AUTOKEY = AUTOKEY + 1 WHERE OBJECTCODE = 'PS_PP040'		
		DECLARE @OrdGbn NVARCHAR(10)
		SET @OrdGbn = (SELECT U_OrdGbn FROM [@PS_PP030H] WHERE DocEntry = @PP030HNo)
		IF @OrdGbn IN('101','102') --����,��ǰ
		BEGIN
			IF (SELECT COUNT(*) FROM [@PS_PP030M] WHERE DocEntry = @PP030HNo AND U_ReportYN = 'N') > 0
			BEGIN
				INSERT INTO [@PS_PP040H]
				(
					DocEntry,	DocNum,		Period,		Instance,	Series,		Handwrtten,	Canceled,	
					Object,		LogInst,		UserSign,	Transfered,	Status,		CreateDate,	CreateTime,	
					UpdateDate,	UpdateTime,	DataSource,	U_DocDate,	U_DocType,	U_OrdType,	U_OrdGbn,
					U_BPLId,	U_ItemCode,	U_ItemName,	U_OrdMgNum,	U_OrdNum,	U_OrdSub1,	U_OrdSub2,
					U_PP030HNo
					--U_UseMCode,	U_UseMName,	U_MoldCode
				)
				SELECT 
					@DocEntry AS DocEntry,
					@DocEntry AS DocNum,
					22 AS Period,
					0 AS Instance,
					-1 AS Series,
					'N' AS Handwrtten,
					'N' AS Canceled,
					'PS_PP040' AS Object,
					NULL AS LogInst,
					1 AS UserSign,
					'N' AS Transfered,
					'O' AS Status,
					CONVERT(NVARCHAR,GETDATE(),112) AS CreateDate,
					NULL AS CreateTime,
					NULL AS UpdateDate,
					NULL AS UpdateTime,
					'I' AS DataSource,
					@DocDate AS DocDate,
					'10' AS DocType, --����Ÿ�� 10 ��������
					'40' AS OrdType, --�۾�Ÿ�� 10 �Ϲ�, 20 PSMT����, 30 ����, 40 �����߰�
					PS_PP030H.U_OrdGbn AS OrdGbn,  --�۾����� 101,102,105,106�� �ش�
					PS_PP030H.U_BPLId AS BPLId, --�����
					PS_PP030H.U_ItemCode AS ItemCode, --��ǰ�ڵ�
					PS_PP030H.U_ItemName AS ItemName, --��ǰ�̸�
					'' AS OrdMgNum, --����������ȣ
					PS_PP030H.U_OrdNum AS OrdNum, --������ȣ
					PS_PP030H.U_OrdSub1 AS OrdSub1,
					PS_PP030H.U_OrdSub2 AS OrdSub2,
					PS_PP030H.DocEntry AS PP030HNo --���������ȣ
					--NULL AS UseMCode, --������ڵ�
					--NULL AS UseMName, --������̸�
					--NULL AS MoldCode --����������ȣ
				FROM
					[@PS_PP030H] PS_PP030H
				WHERE
					PS_PP030H.DocEntry = @PP030HNo											
				
				INSERT INTO [@PS_PP040L]
				(
					DocEntry,	LineId,		VisOrder,	Object,		LogInst,	U_LineNum,	U_LineId,
					U_OrdMgNum,	U_OrdGbn,	U_ItemCode,	U_ItemName,	U_OrdNum,	U_OrdSub1,	U_OrdSub2,
					U_BPLId,	U_PP030HNo,	U_PP030MNo,	U_Sequence,	U_CpCode,	U_CpName,	U_PSum,
					U_BQty,		U_PQty,		U_PWeight,	U_YQty,		U_YWeight,	U_NQty,		U_NWeight,
					U_ScrapWt,	U_WorkTime, U_MM155Dat
				)
				SELECT
					@DocEntry AS DocEntry,								
					ROW_NUMBER() OVER (ORDER BY LineId DESC) AS LineId,
					(ROW_NUMBER() OVER (ORDER BY LineId DESC))-1 AS VisOrder,
					'PS_PP040' AS Object,
					NULL AS LogInst,
					ROW_NUMBER() OVER (ORDER BY LineId DESC) AS U_LineNum,
					ROW_NUMBER() OVER (ORDER BY LineId DESC) AS U_LineId,
					CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) AS OrdMgNum,			
					PS_PP030H.U_OrdGbn AS OrdGbn,
					PS_PP030H.U_ItemCode AS ItemCde,
					PS_PP030H.U_ItemName AS ItemName,
					PS_PP030H.U_OrdNum AS OrdNum,
					PS_PP030H.U_OrdSub1 AS OrdSub1,
					PS_PP030H.U_OrdSub2 AS OrdSub2,
					PS_PP030H.U_BPLId AS BPLId,
					PS_PP030H.DocEntry AS PP030HNo,
					PS_PP030M.LineId AS PP030MNo,
					PS_PP030M.U_Sequence AS Sequence,
					PS_PP030M.U_CpCode AS CpCode,
					PS_PP030M.U_CpName AS CpName,
					0 AS PSum, --���괩�� 101,102,105,106,107 ���� ���
					0 AS BQty, --���ؼ��� 104 ���� ���
					@PQty AS PQty,
					@PWeight AS PWeight,
					@YQty AS YQty,
					@YWeight AS YWeight,
					@NQty AS NQty,
					@NWeight AS NWeight,
					0 AS ScrapWt,
					0 AS WorkTime,
					NULL AS MM155Dat --��ǰ����			
				FROM 
					[@PS_PP030H] PS_PP030H
					LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry
				WHERE
					PS_PP030H.DocEntry = @PP030HNo
					AND PS_PP030M.U_ReportYN = 'N'
					
							
				INSERT INTO [@PS_PP040M]
				(
					DocEntry,	LineId,		VisOrder,	Object,		LogInst,	U_LineNum,	U_LineId,
					U_WorkCode,	U_WorkName,	U_NCode,	U_NStart,	U_NEnd,		U_NTime,	U_LTime,
					U_YTime,	U_TTime
				)
				SELECT
					@DocEntry AS DocEntry,
					1 AS LineId,
					0 AS VisOrder,
					'PS_PP040' AS Object,
					NULL AS LogInst,
					1 AS LineNum,
					1 AS U_LineId,
					@CntcCode AS WorkCode,
					@CntcName AS WorkName,
					NULL,
					NULL,
					NULL,
					NULL,
					NULL,
					NULL,
					NULL
				
							
				INSERT INTO [@PS_PP040N]
				(
					DocEntry,	LineId,		VisOrder,	Object,		LogInst,	U_LineNum,	U_LineId,
					U_OrdMgNum,	U_CpCode,	U_CpName,	U_FailCode,	U_FailName,	U_FailQty
				)
				SELECT
					@DocEntry AS DocEntry,
					ROW_NUMBER() OVER (ORDER BY LineId DESC) AS LineId,
					(ROW_NUMBER() OVER (ORDER BY LineId DESC))-1 AS VisOrder,
					'PS_PP040' AS Object,
					NULL AS LogInst,
					ROW_NUMBER() OVER (ORDER BY LineId DESC) AS U_LineNum,
					ROW_NUMBER() OVER (ORDER BY LineId DESC) AS U_LineId,
					CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) AS OrdMgNum,							
					PS_PP030M.U_CpCode AS CpCode,
					PS_PP030M.U_CpName AS CpName,
					NULL,
					NULL,
					@NQty
				FROM 
					[@PS_PP030H] PS_PP030H
					LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry
				WHERE
					PS_PP030H.DocEntry = @PP030HNo
					AND PS_PP030M.U_ReportYN = 'N'		
				--����Ϸῡ �����Ϻ�������ȣ ������Ʈ
				UPDATE [@PS_PP080H] SET U_PP040No = @DocEntry WHERE DocEntry = @PP080HNo
			END
		END		
		FETCH NEXT FROM CURSOR1 INTO @DocDate,@CntcCode,@CntcName,@PP030No,@PP030HNo,@PP030MNo,@PQty,@PWeight,@YQty,@YWeight,@NQty,@NWeight
	END
	CLOSE CURSOR1
	DEALLOCATE CURSOR1	
END
