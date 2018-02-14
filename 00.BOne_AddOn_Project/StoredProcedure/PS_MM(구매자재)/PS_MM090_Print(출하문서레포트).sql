
GO
IF OBJECT_ID('PS_MM090_Print') IS NOT NULL
	BEGIN
		DROP Proc PS_MM090_Print
	END
GO

CREATE PROC [dbo].PS_MM090_Print
(
--@SPID INT --�ý��� ���μ��� ID
@DocEntry Nvarchar(20) --������ȣ
)
--WITH Encryption  
AS
	/*�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
		--���ν�����	: PS_MM090_Print
		--���ν�������	: �����
		--������		: �ڼ��� ����
		--�۾�����		: 2010.11.06
		--�۾�������	: �ڼ��� ����
		--�۾���������	: 2010.11.06
		--�۾�����		: �����Ÿ������¹�
		--�۾�����		: 
		--��������		: 
		--�� �� ��      : 
		--��������		: 
		EXEC PS_MM090_Print 60
		SELECT * FROM PS_MM090_TEMP
	�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�*/
/* ����ϱ� ���ؼ��� @@SPID�� ������������ �ӽ÷������� PS_MM090_TEMP ���̺��� �ʼ� */
IF OBJECT_ID('PS_MM090_TEMP') IS NULL
	BEGIN
		CREATE TABLE PS_MM090_TEMP (
					ID	INT IDENTITY(1,1),
					SPID   INT,
					DocEntry INT,
					LineNum	INT,
					Quantity NUMERIC(19,6),--����
					Weight	NUMERIC(19,6)--�߷�
			)
	END
--�������� �ӽ����̺����
CREATE TABLE #TEMP_PS_MM090_1
			([ID] INT IDENTITY ,
			[DocNum]	NUMERIC(19,6),
			[CardName]	NVARCHAR(200),	--��ü
			[InDate]	DateTime,		--������
			[PurPose]	NVARCHAR(100),	--����
			[BPLId]		NVARCHAR(10),	--�����
			[BPLName]	NVARCHAR(50),	--������
			[Comment1]	NVARCHAR(100),	--Ư�����
			[TranCard]	NVARCHAR(30),	--�����
			[TranCode]	NVARCHAR(20),	--������ȣ
			[TranCost]	NVARCHAR(100),	--����
			--[TranCost]	Numeric(19,6),	--����
			[Destin]	NVARCHAR(30),	--�������
			[OutNum]	NVARCHAR(50),	--�������ȣ
			[RptGbn]	NVARCHAR(10),	--��±���
			[Title]		Nvarchar(2),	--Ÿ��Ʋ���� 01-���� 02-�ŷ�����

			[ItemName1]	NVARCHAR(100),
			[Size1]		NVARCHAR(100),
			[Quantity1]	NUMERIC(19,6),
			[Weight1]	NUMERIC(19,6),
			[Unit1]	NVARCHAR(20),

			[ItemName2]	NVARCHAR(100),
			[Size2]		NVARCHAR(100),
			[Quantity2]	NUMERIC(19,6),
			[Weight2]	NUMERIC(19,6),
			[Unit2]	NVARCHAR(20),

			[ItemName3]	NVARCHAR(100),
			[Size3]		NVARCHAR(100),
			[Quantity3]	NUMERIC(19,6),
			[Weight3]	NUMERIC(19,6),
			[Unit3]	NVARCHAR(20),

			[ItemName4]	NVARCHAR(100),
			[Size4]		NVARCHAR(100),
			[Quantity4]	NUMERIC(19,6),
			[Weight4]	NUMERIC(19,6),
			[Unit4]	NVARCHAR(20),

			[ItemName5]	NVARCHAR(100),
			[Size5]		NVARCHAR(100),
			[Quantity5]	NUMERIC(19,6),
			[Weight5]	NUMERIC(19,6),
			[Unit5]	NVARCHAR(20),

			[ItemName6]	NVARCHAR(100),
			[Size6]		NVARCHAR(100),
			[Quantity6]	NUMERIC(19,6),
			[Weight6]	NUMERIC(19,6),
			[Unit6]	NVARCHAR(20)
)


DECLARE 
	--//���
	 @DocNum INT				--������ȣ
	,@CardName NVARCHAR(200)	--��ü��
	,@InDate	DateTime		--��������
	,@PurPose	Nvarchar(100)	--����
	,@BPLId		Nvarchar(10)	--�����
	,@BPLName	Nvarchar(50)	--������
	,@Comment1	Nvarchar(100)	--Ư�����
	,@TranCard	Nvarchar(30)	--�����
	,@TranCode	Nvarchar(20)	--������ȣ
	,@TranCost	Nvarchar(100)	--����
	,@Destin	Nvarchar(30)	--�������
	,@OutNum	Nvarchar(50)	--�������ȣ
	,@RptGbn	Nvarchar(10)	--��±���
	,@Title		Nvarchar(2)		--Ÿ��Ʋ���� 01-���� 02-�ŷ�����
	--//����
	,@ItemName NVARCHAR(100)	--ǰ��
	,@Size	   NVARCHAR(100)	--�԰�
	,@Quantity NUMERIC(19,6)	--����
	,@Weight NUMERIC(19,6)		--�߷�
	,@Unit NVARCHAR(20)			--����

	,@ILOOPER	INT
	,@INDEXID	INT
	,@BEFORE_DocNum	INT
	,@ReptCount	INT	

SET @ReptCount = 1
WHILE @ReptCount < 3 --�ι�����.
BEGIN
IF @ReptCount = 1
	SET @RptGbn = 'Type_A'	--ù��Ÿ��
ELSE
	SET @RptGbn = 'Type_B'	--��°��Ÿ��
	
DECLARE MM090_CUR1 CURSOR	FOR
	SELECT 
	/*���*/
	T0.DocNum
	,T0.U_CardName		--��ü��
	,T0.U_InDate		--��������
	,T0.U_PurPose		--����
	,T0.U_BPLId			--�����
	,T3.BPLName			--������
	,T0.U_Comment1		--Ư�����
	,T0.U_TranCard		--�����(��۾�ü)
	,T0.U_TranCode		--������ȣ
	,T0.U_TranCost		--����
	,T0.U_Destin		--����â��
	,T0.U_OutNum		--�������ȣ
	--,''	AS RptGbn		--��±���
	,T0.U_Title			--Ÿ��Ʋ����	01-����, 02-�ŷ�����
	/*����*/
	,Convert(nvarchar(100),T1.U_ItemName) AS ItemName	--ǰ��
	,T1.U_Size AS Size			--�԰�
	,T1.U_Qty AS Quantity		--����
	,T1.U_Weight AS Weight		--�߷�
	,T1.U_Unit AS Unit			--����
	FROM 
		[@PS_MM090H] T0 
		LEFT JOIN [@PS_MM090L]T1 ON T0.DocEntry = T1.DocEntry
		--LEFT JOIN 
		--(SELECT 
		--	SPID
		--	,DocEntry
		--	,LineNum
		--	,Quantity AS Quantity
		--	,Weight AS Weight
		--FROM 
		--	[PS_MM090_TEMP]
		--) T2 ON T0.DocEntry = T2.DocEntry AND T1.U_LineNum = T2.LineNum
		LEFT JOIN [OBPL] T3 ON T0.U_BPLId=T3.BPLId
	--WHERE SPID = '70'
	--WHERE SPID =@SPID
	WHERE T0.DocEntry = @DocEntry
	--WHERE T0.DocEntry = '1'

		SET @ILOOPER = 1
		OPEN MM090_CUR1
			FETCH NEXT FROM MM090_CUR1 
			INTO @DocNum,@CardName,@InDate,@PurPose,@BPLId,@BPLName,@Comment1,@TranCard,@TranCode,@TranCost,@Destin,@OutNum,@Title,@ItemName,@Size,@Quantity,@Weight,@Unit
			WHILE @@FETCH_STATUS = 0 --�����ϸ� 0����ȯ
			BEGIN	
						IF @ILOOPER > 6 OR @BEFORE_DocNum <> @DocNum 
							SET @ILOOPER = 1
						IF @ILOOPER = 1 --//ù��°�̸� INSERT
							BEGIN
								INSERT INTO #TEMP_PS_MM090_1 (DocNum,CardName,InDate,PurPose,BPLId,BPLName,Comment1,TranCard
																,TranCode,TranCost,Destin,OutNum,RptGbn,Title
																,ItemName1
																,Size1
																,Quantity1
																,Weight1
																,Unit1
																)
								VALUES(@DocNum,@CardName,@InDate,@PurPose,@BPLId,@BPLName,@Comment1,@TranCard
																,@TranCode,@TranCost,@Destin,@OutNum,@RptGbn,@Title
																,@ItemName
																,@Size
																,@Quantity
																,@Weight																
																,@Unit
																)
							SET @INDEXID = (SELECT MAX([ID]) FROM #TEMP_PS_MM090_1)
							END
						ELSE IF @ILOOPER = 2 --//�ι�°���ʹ� �����࿡ UPDATE 6��°����
							BEGIN
								UPDATE #TEMP_PS_MM090_1 
									SET	
										ItemName2 = @ItemName
										,Size2	=	@Size
										,Quantity2 = @Quantity
										,Weight2 = @Weight
										,Unit2 = @Unit
									WHERE [ID] = @INDEXID
							END
						ELSE IF @ILOOPER = 3
							BEGIN
								UPDATE #TEMP_PS_MM090_1 
									SET	
										ItemName3 = @ItemName
										,Size3	=	@Size
										,Quantity3 = @Quantity
										,Weight3 = @Weight
										,Unit3 = @Unit
									WHERE [ID] = @INDEXID
							END
						ELSE IF @ILOOPER = 4
							BEGIN
								UPDATE #TEMP_PS_MM090_1 
									SET	
										ItemName4 = @ItemName
										,Size4	=	@Size
										,Quantity4 = @Quantity
										,Weight4 = @Weight
										,Unit4 = @Unit
									WHERE [ID] = @INDEXID
							END
						ELSE IF @ILOOPER = 5
							BEGIN
								UPDATE #TEMP_PS_MM090_1 
									SET	
										ItemName5 = @ItemName
										,Size5	=	@Size
										,Quantity5 = @Quantity
										,Weight5 = @Weight
										,Unit5 = @Unit
									WHERE [ID] = @INDEXID
							END
						ELSE IF @ILOOPER = 6 --//6��������UPDATE
							BEGIN
								UPDATE #TEMP_PS_MM090_1 
									SET	
										ItemName6 = @ItemName
										,Size6	=	@Size
										,Quantity6 = @Quantity
										,Weight6 = @Weight
										,Unit6 = @Unit
									WHERE [ID] = @INDEXID
							END
						ELSE
							BEGIN
								SET @ILOOPER = 1
							END

						SET @BEFORE_DocNum = @DocNum
						SET @ILOOPER = @ILOOPER +1
			FETCH NEXT FROM MM090_CUR1 
			INTO @DocNum,@CardName,@InDate,@PurPose,@BPLId,@BPLName,@Comment1,@TranCard,@TranCode,@TranCost,@Destin,@OutNum,@Title,@ItemName,@Size,@Quantity,@Weight,@Unit --����Ŀ�����̵�
			END
		CLOSE MM090_CUR1
		DEALLOCATE MM090_CUR1 --Ŀ���� �޸𸮿��� ����

	SET @ReptCount = @ReptCount + 1
END
--��������ȸ
SELECT
			DocNum	
			,CardName	
			,InDate		--��������
			,PurPose	--����	
			,BPLId		--�����
			,[BPLName]	--������
			,[Comment1]	--Ư�����
			,[TranCard]	--�����
			,[TranCode]	--������ȣ
			,[TranCost]	--����
			,[Destin]	--�������
			,[OutNum]	--�������ȣ
			,RptGbn		--��±���
			,Title		--Ÿ��Ʋ���� 01-���� 02-�ŷ�����
			
			,Convert(nvarchar(100),ItemName1) as ItemName1
			,Size1
			,Quantity1
			,Weight1
			,Unit1

			,Convert(nvarchar(100),ItemName2) as ItemName2
			,Size2
			,Quantity2
			,Weight2
			,Unit2

			,ItemName3
			,Size3
			,Quantity3
			,Weight3
			,Unit3

			,ItemName4
			,Size4
			,Quantity4
			,Weight4
			,Unit4

			,ItemName5
			,Size5
			,Quantity5
			,Weight5
			,Unit5

			,ItemName6
			,Size6
			,Quantity6
			,Weight6
			,Unit6
			
			,CONVERT(INTEGER
			,ISNULL(Quantity1,0)
			+ISNULL(Quantity2,0)
			+ISNULL(Quantity3,0)
			+ISNULL(Quantity4,0)
			+ISNULL(Quantity5,0)
			+ISNULL(Quantity6,0)) AS TotalQty

			,ISNULL(Weight1,0)
			+ISNULL(Weight2,0)
			+ISNULL(Weight3,0)
			+ISNULL(Weight4,0)
			+ISNULL(Weight5,0)
			+ISNULL(Weight6,0) AS TotalWeight
 FROM #TEMP_PS_MM090_1
	--SELECT
	--		 Convert(NVARCHAR(100),DocNum)	AS	DocNum
	--		,Convert(NVARCHAR(100),CardName)AS	CardName
	--		,Convert(NVARCHAR(100),InDate)	AS	InDate		--��������
	--		,Convert(NVARCHAR(100),PurPose)	AS	PurPose		--����	
	--		,Convert(NVARCHAR(100),BPLId)	AS	BPLId		--�����
	--		,Convert(NVARCHAR(100),BPLName)	AS	BPLName		--������
	--		,Convert(NVARCHAR(100),Comment1)AS	Comment1	--Ư�����
	--		,Convert(NVARCHAR(100),TranCard)AS	TranCard	--�����
	--		,Convert(NVARCHAR(100),TranCode)AS	TranCode	--������ȣ
	--		,Convert(NVARCHAR(100),TranCost)AS	TranCost	--����
	--		,Convert(NVARCHAR(100),Destin)	AS	Destin		--�������
	--		,Convert(NVARCHAR(100),OutNum)	AS	OutNum		--�������ȣ
	--		,Convert(NVARCHAR(100),RptGbn)	AS	RptGbn		--��±���
	--		,Convert(NVARCHAR(100),Title)	AS	Title		--Ÿ��Ʋ���� 01-���� 02-�ŷ�����
			
	--		,Convert(NVARCHAR(100),ItemName1)	AS	ItemName1
	--		,Convert(NVARCHAR(100),Size1)		AS	Size1
	--		,Convert(NVARCHAR(100),Quantity1)	AS	Quantity1
	--		,Convert(NVARCHAR(100),Weight1)		AS	Weight1
	--		,Convert(NVARCHAR(100),Unit1)		AS	Unit1

	--		,Convert(NVARCHAR(100),ItemName2)	AS	ItemName2
	--		,Convert(NVARCHAR(100),Size2)		AS	Size2
	--		,Convert(NVARCHAR(100),Quantity2)	AS	Quantity2
	--		,Convert(NVARCHAR(100),Weight2)		AS	Weight2
	--		,Convert(NVARCHAR(100),Unit2)		AS	Unit2

	--		,Convert(NVARCHAR(100),ItemName3)	AS	ItemName3
	--		,Convert(NVARCHAR(100),Size3)		AS	Size3
	--		,Convert(NVARCHAR(100),Quantity3)	AS	Quantity3
	--		,Convert(NVARCHAR(100),Weight3)		AS	Weight3
	--		,Convert(NVARCHAR(100),Unit3)		AS	Unit3

	--		,Convert(NVARCHAR(100),ItemName4)	AS	ItemName4
	--		,Convert(NVARCHAR(100),Size4)		AS	Size4
	--		,Convert(NVARCHAR(100),Quantity4)	AS	Quantity4
	--		,Convert(NVARCHAR(100),Weight4)		AS	Weight4
	--		,Convert(NVARCHAR(100),Unit4)		AS	Unit4

	--		,Convert(NVARCHAR(100),ItemName5)	AS	ItemName5
	--		,Convert(NVARCHAR(100),Size5)		AS	Size5
	--		,Convert(NVARCHAR(100),Quantity5)	AS	Quantity5
	--		,Convert(NVARCHAR(100),Weight5)		AS	Weight5
	--		,Convert(NVARCHAR(100),Unit5)		AS	Unit5

	--		,Convert(NVARCHAR(100),ItemName6)	AS	ItemName6
	--		,Convert(NVARCHAR(100),Size6)		AS	Size6
	--		,Convert(NVARCHAR(100),Quantity6)	AS	Quantity6
	--		,Convert(NVARCHAR(100),Weight6)		AS	Weight6
	--		,Convert(NVARCHAR(100),Unit6)		AS	Unit6
	--		,CONVERT(INTEGER
	--		,ISNULL(Quantity1,0)
	--		+ISNULL(Quantity2,0)
	--		+ISNULL(Quantity3,0)
	--		+ISNULL(Quantity4,0)
	--		+ISNULL(Quantity5,0)
	--		+ISNULL(Quantity6,0)) AS TotalQty

	--		,ISNULL(Weight1,0)
	--		+ISNULL(Weight2,0)
	--		+ISNULL(Weight3,0)
	--		+ISNULL(Weight4,0)
	--		+ISNULL(Weight5,0)
	--		+ISNULL(Weight6,0) AS TotalWeight
 --FROM #TEMP_PS_MM090_1


go
