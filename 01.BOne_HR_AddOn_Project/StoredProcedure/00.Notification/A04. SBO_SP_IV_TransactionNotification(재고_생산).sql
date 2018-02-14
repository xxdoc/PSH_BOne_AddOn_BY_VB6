IF OBJECT_ID('SBO_SP_IV_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_IV_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_IV_TransactionNotification
 ��      �� : ������ TransactionNotification
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_IV_TransactionNotification] 

	  @object_type				NVARCHAR(20)	-- SBO Object Type
	, @transaction_type			NCHAR(1)		-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
	, @num_of_cols_in_key		INT
	, @list_of_key_cols_tab_del NVARCHAR(255)
	, @list_of_cols_val_tab_del NVARCHAR(255)
	, @error					INT OUTPUT
	, @error_message			NVARCHAR(200) OUTPUT

AS

BEGIN

SET NOCOUNT ON
/***************************************** Sample ********************************************************
IF @object_type = '30'
BEGIN
	IF @transaction_type IN ('A', 'U')
	BEGIN
		IF EXISTS(SELECT U_CODE FROM [@PWC_BCCSY1] WHERE CODE=@list_of_cols_val_tab_del AND U_CODE IS NULL)
		BEGIN
			SET @error = -1										-- Error Code ����
			SET @error_message = N'�ڵ带 �Է��� �ּ���.'		-- Error Message ����
			RETURN												-- Error �߻��� RETURN�� ����Ͽ� SP EXIT
		END
	END
END
****************************************** Sample ********************************************************/
--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------
DECLARE @VALUE01 NVARCHAR(MAX)

/*** OITM : ������ - ǰ�񸶽��� ***/
IF @object_type = '4'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* ǰ����� �ʼ� üũ
		IF EXISTS(SELECT ItemCode FROM OITM 
					WHERE ItemCode= @list_of_cols_val_tab_del
					AND EvalSystem = 'S'
					AND ISNULL(AvgPrice,0) = 0)
		BEGIN
			SET @error = -1
			SET @error_message = N'��ǰ/����ǰ ǰ�� ǰ����� 0�Դϴ�.'
			RETURN
		END
	END
END


/*** OIGN : ������ - �԰� ***/
IF @object_type = '59'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OIGN WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END
		
		IF EXISTS(SELECT DocEntry FROM OIGN WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'���-�μ��� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END
		--* �ܰ� 0�� üũ
		IF EXISTS(SELECT DocEntry FROM IGN1 WHERE DocEntry = @list_of_cols_val_tab_del AND ISNULL(Price,0) = 0 )
		BEGIN
			SET @error = -1
			SET @error_message = N'�ܰ��� 0���� �Է��� �� �����ϴ�. '
			RETURN
		END
	END
	
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		--* ���� �԰� - ���� �������� ������ ���� ��� ���� �޽��� ���
		IF EXISTS(
			SELECT DocEntry
			  FROM OWOR
			 WHERE DocEntry IN (SELECT BaseEntry FROM IGN1 WHERE DocEntry=@list_of_cols_val_tab_del AND BaseType='202')
			   AND PlannedQty < ISNULL(CmpltQty, 0)+ISNULL(RjctQty, 0)
		)
		BEGIN
			SET @error = -1
			SET @error_message = N'���� �԰� ������ ���� ������ ��ȹ �������� Ŭ �� �����ϴ�.'
			RETURN
		END
		
		--* ���� ����ǰ ���� ��� �Ͼ�� ���� ��� ���� �԰� ���� �ʵ��� ó��
		--IF EXISTS(
		--	SELECT OWOR.DocEntry
		--	  FROM OWOR
		--	 INNER JOIN WOR1 ON OWOR.DocEntry=WOR1.DocEntry
		--	 WHERE OWOR.DocEntry IN (SELECT BaseEntry FROM IGN1 WHERE DocEntry=@list_of_cols_val_tab_del AND BaseType='202')
		--	   AND WOR1.IssueType = 'M'
		--	   AND (ISNULL(OWOR.CmpltQty, 0) + ISNULL(OWOR.RjctQty, 0)) * WOR1.BaseQty > WOR1.IssuedQty
		--)
		--BEGIN
		--	SET @error = -1
		--	SET @error_message = N'���� ����ǰ�� ��� ���� ��� �� ���� �԰� �̷������ �˴ϴ�.'
		--	RETURN
		--END
	END
END


/*** OIGE : ������ - ��� ***/
IF @object_type = '60'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OIGE WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END
		
		IF EXISTS(SELECT DocEntry FROM OIGE WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'���-�μ��� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END
	END
END


/*** OWTQ : ������ - ��� ���� ��û ***/
IF @object_type = '1250000001'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OWTQ WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END
		
		IF EXISTS(SELECT DocEntry FROM OWTQ WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'���-�μ��� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END
	END
END


/*** OWTR : ������ - ��� ���� ***/
IF @object_type = '67'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OWTR WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END		
		
		IF EXISTS(SELECT DocEntry FROM OWTR WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'���-�μ��� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END
	END	
END


/*** OMRV : ������ - ������� ***/
IF @object_type = '162'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OMRV WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END
	END
END


/*** OWOR : ������� - ���� ���� ***/
IF @object_type = '202'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A','U')
	BEGIN
		--* ���� ǰ�� CostCenter �ʼ� üũ
		IF EXISTS(SELECT DocEntry FROM OWOR WHERE DocEntry=@list_of_cols_val_tab_del AND ISNULL(OcrCode, '')='')
		BEGIN
			SET @error = -1
			SET @error_message = N'���� ǰ�� ��� ��Ģ�� �ʼ��Դϴ�.'
			RETURN
		END
		
		--* ���� ǰ�� CostCenter �ʼ� üũ
		IF EXISTS(SELECT DocEntry FROM WOR1 WHERE DocEntry=@list_of_cols_val_tab_del AND ISNULL(OcrCode, '')='')
		BEGIN
			SET @error = -1
			SET @error_message = N'���� ǰ�� Cost Center�� �ʼ��Դϴ�.'
			RETURN
		END

		--* ���� ��ǰ�� ǰ����� �ʼ� üũ
		IF EXISTS(SELECT T0.DocEntry FROM OWOR T0 INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode
					WHERE T0.DocEntry= @list_of_cols_val_tab_del
					AND T1.EvalSystem = 'S'
					AND ISNULL(T1.AvgPrice,0) = 0)
		BEGIN
			SET @error = -1
			SET @error_message = N'���� ��ǰ�� ǰ����� 0�Դϴ�.'
			RETURN
		END
		--* ���� ��ǰ�� ǰ����� �ʼ� üũ
		ELSE IF EXISTS(SELECT T0.DocEntry FROM WOR1 T0 INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode
					WHERE T0.DocEntry=@list_of_cols_val_tab_del
					AND T1.EvalSystem = 'S'
					AND ISNULL(T1.AvgPrice,0) = 0)
		BEGIN
			SET @error = -1
			SET @error_message = N'���� ��ǰ�� ǰ����� 0�Դϴ�.'
			RETURN
		END
	END
END

IF @object_type = 'MDC_MM_CSM001'
BEGIN
	IF @transaction_type IN ('A', 'U')
	BEGIN
		SELECT @VALUE01 = LEN(@list_of_cols_val_tab_del)
		IF @VALUE01 > 4
		BEGIN
			SET @error = -1										-- Error Code ����
			SET @error_message = N'�ڵ�� 4�ڸ��� �Է��ϼ���.'	-- Error Message ����
			RETURN												-- Error �߻��� RETURN�� ����Ͽ� SP EXIT
		END
	END
END


/*** MDC_MM_CPD103 : ������� - �۾��ҷ� �� �ν� ��� ***/
IF @object_type = 'MDC_MM_CPD103'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		IF EXISTS(SELECT DocEntry FROM [@MDC_MM_CPD103H] WHERE DocEntry = @list_of_cols_val_tab_del AND U_OutEntry IS NULL)
		BEGIN
			SET @error = -1										
			SET @error_message = N'��Ͽ� ���� ��� ó���� ���������� �̷������ �ʾҽ��ϴ�.'
			RETURN
		END
		
		IF (SELECT COUNT(DocEntry) FROM [@MDC_MM_CPD103L] WHERE DocEntry = @list_of_cols_val_tab_del AND ISNULL(U_ItemCode, '') <> '') = 0 
		BEGIN
			SET @error = -1										
			SET @error_message = N'�۾� �ҷ� �� �ν� ǰ�� ������ �������� �ʽ��ϴ�.'
			RETURN
		END
	END
	
	/* ��� ��� */
	IF @transaction_type = 'C'
	BEGIN
		IF EXISTS(SELECT DocEntry FROM [@MDC_MM_CPD103H] WHERE DocEntry = @list_of_cols_val_tab_del AND U_InEntry IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'��ҿ� ���� �԰� ó���� ���������� �̷������ �ʾҽ��ϴ�.'
			RETURN
		END
	END	
END


/*** MDC_MM_CPD105 : ������� - �۾� �ܷ� �԰� ��� ***/
IF @object_type = 'MDC_MM_CPD105'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		IF EXISTS(SELECT DocEntry FROM [@MDC_MM_CPD105H] WHERE DocEntry = @list_of_cols_val_tab_del AND U_InEntry IS NULL)
		BEGIN
			SET @error = -1										
			SET @error_message = N'��Ͽ� ���� �԰� ó���� ���������� �̷������ �ʾҽ��ϴ�.'
			RETURN
		END
		
		IF (SELECT COUNT(DocEntry) FROM [@MDC_MM_CPD105L] WHERE DocEntry = @list_of_cols_val_tab_del AND ISNULL(U_ItemCode, '') <> '') = 0 
		BEGIN
			SET @error = -1										
			SET @error_message = N'�۾� �ܷ� �԰� ǰ�� ������ �������� �ʽ��ϴ�.'
			RETURN
		END
	END
	
	/* ��� ��� */
	IF @transaction_type = 'C'
	BEGIN
		IF EXISTS(SELECT DocEntry FROM [@MDC_MM_CPD105H] WHERE DocEntry = @list_of_cols_val_tab_del AND U_OutEntry IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'��ҿ� ���� ��� ó���� ���������� �̷������ �ʾҽ��ϴ�.'
			RETURN
		END
	END
END


/*** MDC_MM_CPD211 : ������� - PMS ���� �Ƿ� ����_��� ***/
IF @object_type = 'MDC_MM_CPD211'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		--* PMS Data ���� ���� üũ		 
		IF EXISTS(
			SELECT SAPOPRR.U_PrjSeqn
			  FROM (
				SELECT    PRR1.U_PrjSeqn
						, PRR1.U_ItemCode			
						, CONVERT(NVARCHAR(8), U_DueDate, 112) AS DueDate
						, PRR1.U_ReceQty
				  FROM [@MDC_MM_CPD211H] OPRR
				 INNER JOIN [@MDC_MM_CPD211L] PRR1 ON OPRR.DocEntry=PRR1.DocEntry
				 WHERE OPRR.DocEntry = @list_of_cols_val_tab_del
			 ) SAPOPRR
			 INNER JOIN (
				SELECT    PMSRQI.PJT_SEQ
						, PMSPJI.ITEM_CD
						, PMSRQI.TARGET_DT
						, PMSRQI.REQ_QTY
				  FROM [EAGON_PMS].[dbo].[TPMS_PJT_ITEM_I] PMSPJI 
				 INNER JOIN [EAGON_PMS].[dbo].[TPMS_PREQ_ITEM] PMSRQI ON PMSPJI.PJT_CD=PMSRQI.PJT_CD AND PMSPJI.PJT_SEQ=PMSRQI.PJT_SEQ
				 WHERE PMSPJI.PJT_CD = (SELECT U_PrjCode COLLATE Korean_Wansung_CI_AS FROM [@MDC_MM_CPD211H] WHERE DocEntry = @list_of_cols_val_tab_del)
				   AND PMSRQI.PREQ_NUM = (SELECT U_PreqNum COLLATE Korean_Wansung_CI_AS FROM [@MDC_MM_CPD211H] WHERE DocEntry = @list_of_cols_val_tab_del)
			 ) PMSREQI ON SAPOPRR.U_PrjSeqn=PMSREQI.PJT_SEQ -- �ϳ��� �����Ƿ� ��ȣ�� ���Ͽ� PJT_SEQ�� �ߺ� �� �� �����Ƿ� PJT_SEQ�� ����
			 WHERE SAPOPRR.U_ItemCode <> PMSREQI.ITEM_CD COLLATE Korean_Wansung_Unicode_CI_AS
				OR SAPOPRR.DueDate <> PMSREQI.TARGET_DT COLLATE Korean_Wansung_Unicode_CI_AS
				OR SAPOPRR.U_ReceQty <> PMSREQI.REQ_QTY
		)
		BEGIN
			SET @error = -1
			SET @error_message = N'PMS���� �ٸ� ����ڿ� ���Ͽ� ���� �Ƿ� ������ ����Ǿ����ϴ�.'
			RETURN
		END
	END
END

IF @object_type = 'MDC_MM_CPD104'
BEGIN
	IF @transaction_type IN ('A', 'U')
	BEGIN
		IF EXISTS(SELECT U_OcrCode FROM [@MDC_MM_CPD104L] 
				WHERE DocEntry = @list_of_cols_val_tab_del
				GROUP BY U_OcrCode HAVING COUNT(U_OcrCode) > 1)
		BEGIN
			SET @error = -1										-- Error Code ����
			SET @error_message = N'�ߺ��� �۾������� �ֽ��ϴ�.'	-- Error Message ����
			RETURN												-- Error �߻��� RETURN�� ����Ͽ� SP EXIT
		END
	END
END


/*** MDC_MM_CPD106 : ������� - ���� ���� ���� ���[PVC, �˹̴�] ***/
IF @object_type = 'MDC_MM_CPD106'
BEGIN
	/* ��� ��� */
	IF @transaction_type = 'C'
	BEGIN
		IF EXISTS(	SELECT DocEntry
					  FROM OWOR WITH (NOLOCK)
					 WHERE DocEntry IN (SELECT U_WREntry FROM [@MDC_MM_CPD106L] WHERE DocEntry = @list_of_cols_val_tab_del)
					   AND [Status] <> 'C'
		)
		BEGIN
			SET @error = -1
			SET @error_message = N'�ش� ������ ���� ������ ��� ��� ó�� ���� ���� ��� ��Ҹ� �� �� �����ϴ�.'
			RETURN
		END
	END
END


END