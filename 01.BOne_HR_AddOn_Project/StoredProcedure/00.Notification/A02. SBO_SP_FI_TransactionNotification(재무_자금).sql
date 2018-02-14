IF OBJECT_ID('SBO_SP_FI_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_FI_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_FI_TransactionNotification
 ��      �� : �繫����, �ڱݰ��� TransactionNotification
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_FI_TransactionNotification] 

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
/*** OJDT : �а� ***/
IF @object_type = '30'
BEGIN

	/* �ٸ� ������ ��ó(�Ա�, ������ ���)�� ������ ��쿡�� �ش� IF�� �� �ʿ� �۾� ���� */
	IF NOT EXISTS(SELECT TransId FROM OJDT WHERE TransId=@list_of_cols_val_tab_del AND TransType='30')
	BEGIN
		RETURN	-- �а����� �ۼ��� ��쿡�� ���� ��ȿ�� �˻簡 ���� �ǵ��� ����(������� �߰ߵ� �Ա�, ������ ��쿡�� �ش� ������ �����ǹǷ� �̿� ���� ó��)
	END
	
	DECLARE @OJDT_nvcAcctCode	NVARCHAR(15)

	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* ���Ա� ������ȣ ���� ���� �˻�
		IF EXISTS(	SELECT JDT1.U_PWC_DebtCode
					  FROM JDT1 JDT1
		              LEFT JOIN [@PWC_TRODBT] ODBT WITH (NOLOCK) ON JDT1.U_PWC_DebtCode=ODBT.Code
		             WHERE JDT1.TransId=@list_of_cols_val_tab_del
		               AND JDT1.U_PWC_DebtCode IS NOT NULL
		               AND ODBT.Code IS NULL	)
		BEGIN
			SET @error = -1
			SET @error_message = N'�������� �ʴ� ���Ա� ������ȣ�� �Է��Ͽ����ϴ�.'
			RETURN
		END
		
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(
			SELECT OJDT.TransId
			  FROM OJDT
			 INNER JOIN JDT1 ON OJDT.TransId=JDT1.TransId
			 WHERE OJDT.TransId=@list_of_cols_val_tab_del
			   AND OJDT.U_BA_TCODE IS NULL					-- �����ڻ� �߻� ���� ��� ����� ��ȿ �˻� ����
			   AND JDT1.U_PWC_BpliCode IS NULL
		)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END
		
		--* BP �ʼ� ���� ���� ����
		SET @OJDT_nvcAcctCode = ''
		SELECT TOP 1 @OJDT_nvcAcctCode = JDT1.Account
		  FROM JDT1
		 INNER JOIN OACT WITH (NOLOCK) ON JDT1.Account=OACT.AcctCode
		 WHERE JDT1.TransId=@list_of_cols_val_tab_del
		   AND OACT.U_PWC_MdBpCd = 'Y'
		   AND ISNULL(JDT1.U_PWC_CardCode, '') = ''
		   			
		IF @OJDT_nvcAcctCode <> ''
		BEGIN
			SET @error = -1
			SET @error_message = N'�ش� ������ ��� B/P�ڵ�� �ʼ��Դϴ�. [G/L ���� : ' + @OJDT_nvcAcctCode + ']'
			RETURN
		END
	END
	
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		--* ���Ա� ������ȣ ��Ȱ�� �� �Է� ���� �˻�
		IF EXISTS(	SELECT ODBT.Code
		              FROM JDT1 JDT1
		             INNER JOIN [@PWC_TRODBT] ODBT WITH (NOLOCK) ON JDT1.U_PWC_DebtCode=ODBT.Code 
		             WHERE JDT1.TransId=@list_of_cols_val_tab_del
		               AND ODBT.U_Active='N'	)
		BEGIN
			SET @error = -1
			SET @error_message = N'��Ȱ�� ó���� ���Ա� ������ȣ�� �Է��� �� �����ϴ�.'
			RETURN
		END
	END
	
	/* ���� ��� */
	IF @transaction_type = 'U'
	BEGIN
		--* ���Ա� ������ȣ ��Ȱ�� �� ���� ���� �˻�
		IF EXISTS(	SELECT DBT3.Code
					  FROM [@PWC_TRODBT] ODBT WITH (NOLOCK)
					 INNER JOIN [@PWC_TRDBT3] DBT3 WITH (NOLOCK) ON ODBT.Code=DBT3.Code
					 INNER JOIN JDT1 JDT1 ON DBT3.U_TransIdx=JDT1.TransId AND DBT3.U_JdtLineId=JDT1.Line_Id
					 WHERE JDT1.TransId = @list_of_cols_val_tab_del
					   AND ODBT.U_Active = 'N'
					   AND ODBT.Code <> ISNULL(JDT1.U_PWC_DebtCode, '')	)
		BEGIN
			SET @error = -1
			SET @error_message = N'��Ȱ�� ó���� ���Ա� ������ȣ�� ������ �� �����ϴ�.'
			RETURN
		END	
	END
END


/*** ORCT : �ڱݰ��� - �Ա� ***/
IF @object_type = '24'
BEGIN
	/* �߰�, ����, ��� ��� */
	IF @transaction_type IN ('A', 'U', 'C')
	BEGIN
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM ORCT WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END
	END
END


/*** ODPS : �ڱݰ��� - ���� ***/
IF @object_type = '25'
BEGIN

	/* �߰�, ����, ��� ��� */
	IF @transaction_type IN ('A', 'U', 'C')
	BEGIN
		--* ���� ó�� ������ ������ ��� ��ȿ�� �˻� X (������������ ó��)
		IF (SELECT TOP 1 DeposType FROM ODPS WHERE DeposId = @list_of_cols_val_tab_del) = 'B'	
		BEGIN
			RETURN
		END
		
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(SELECT DeposId FROM ODPS WHERE DeposId=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END		
	END
END


/*** OVPM : �ڱݰ��� - ���� ***/
IF @object_type = '46'
BEGIN
	/* �߰�, ����, ��� ��� */
	IF @transaction_type IN ('A', 'U', 'C')
	BEGIN
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OVPM WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END
	END
END


/*** OBOT : �ڱݰ��� - �������� ***/
IF @object_type = '182'
BEGIN
	DECLARE   @OBOT_nvcBoeType			CHAR(1)			-- �Ա�(I), ����(O) ����
	
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		SELECT   TOP 1 @OBOT_nvcBoeType=BOT1.BoeType
		  FROM OBOT
		 INNER JOIN BOT1 ON OBOT.AbsEntry=BOT1.AbsEntry
		 WHERE OBOT.AbsEntry = @list_of_cols_val_tab_del
		
		IF @OBOT_nvcBoeType = 'I'		-- �Ա�
		BEGIN
			-- ���� ����� ó�� �Ұ�
			IF (	SELECT COUNT(*)
					  FROM (
						SELECT ORCT.U_PWC_BPLId
						  FROM OBOE WITH (NOLOCK)
						 INNER JOIN ORCT WITH (NOLOCK) ON OBOE.PmntNum=ORCT.DocEntry
						 WHERE OBOE.BoeKey IN (SELECT BOEAbs FROM BOT1 WHERE AbsEntry = @list_of_cols_val_tab_del)  
						 GROUP BY ORCT.U_PWC_BPLId
					     ) T1
				) > 1			
			BEGIN
				SET @error = -1
				SET @error_message = N'���� ������� ��츸 ���� ó���� �����մϴ�.'
				RETURN
			END
		END
		ELSE IF @OBOT_nvcBoeType = 'O'	-- ����
		BEGIN
			-- ���� ����� ó�� �Ұ�
			IF (	SELECT COUNT(*)
					  FROM (
						SELECT OVPM.U_PWC_BPLId
						  FROM OBOE WITH (NOLOCK)
						 INNER JOIN OVPM WITH (NOLOCK) ON OBOE.PmntNum=OVPM.DocEntry
						 WHERE OBOE.BoeKey IN (SELECT BOEAbs FROM BOT1 WHERE AbsEntry = @list_of_cols_val_tab_del)  
						 GROUP BY OVPM.U_PWC_BPLId
					     ) T1
				) > 1			
			BEGIN
				SET @error = -1
				SET @error_message = N'���� ������� ��츸 ���� ó���� �����մϴ�.'
				RETURN
			END
		END
		
	END
END


/*** [@PWC_TRODBT] : ���Ա� ������ ��� ***/
IF @object_type = 'PWC_UDO_TRODBT'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		DECLARE   @nvc_TRODBT_DebtCode	NVARCHAR(15)
				, @nvc_TRODBT_CardCode	NVARCHAR(15)
				, @dat_TRODBT_DebtDate	DATETIME
				, @dat_TRODBT_DuexDate	DATETIME
				, @int_TRODBT_BPLId		INT
				, @ins_TRODBT_DebtType	NVARCHAR(2)
				, @nvc_TRODBT_DebtAcct	NVARCHAR(15)				

		SELECT    @nvc_TRODBT_DebtCode = ISNULL(Code, '')
				, @nvc_TRODBT_CardCode = ISNULL(U_CardCode, '')
				, @dat_TRODBT_DebtDate = ISNULL(U_DebtDate, '19000101')
				, @dat_TRODBT_DuexDate = ISNULL(U_DuexDate, '19000101')
				, @int_TRODBT_BPLId = ISNULL(U_BPLId, -99)
				, @ins_TRODBT_DebtType = ISNULL(U_DebtType, '')
				, @nvc_TRODBT_DebtAcct = ISNULL(U_DebtAcct, '')				
		  FROM [@PWC_TRODBT]
		 WHERE Code=@list_of_cols_val_tab_del
		
		IF @nvc_TRODBT_DebtCode = ''
		BEGIN
			SET @error = -1
			SET @error_message = N'���Ա� ������ȣ�� ���������� �ڵ� ä������ �ʾҽ��ϴ�. �����ڿ��� ������ �ּ���.'
			RETURN
		END
		
		IF @nvc_TRODBT_CardCode = ''
		BEGIN
			SET @error = -1
			SET @error_message = N'���޾�ü �ڵ带 �Է��� �ּ���.'
			RETURN
		END
		
		IF @dat_TRODBT_DebtDate = '19000101'
		BEGIN
			SET @error = -1
			SET @error_message = N'�������� �Է��� �ּ���.'
			RETURN
		END
		
		IF @dat_TRODBT_DuexDate = '19000101'
		BEGIN
			SET @error = -1
			SET @error_message = N'�������� �Է��� �ּ���.'
			RETURN
		END
		
		IF @int_TRODBT_BPLId = -99
		BEGIN
			SET @error = -1
			SET @error_message = N'������� ������ �ּ���.'
			RETURN
		END
		
		IF @int_TRODBT_BPLId = -99
		BEGIN
			SET @error = -1
			SET @error_message = N'������� ������ �ּ���.'
			RETURN
		END
		
		IF @ins_TRODBT_DebtType = ''
		BEGIN
			SET @error = -1
			SET @error_message = N'���԰����� �Է��� �ּ���.'
			RETURN
		END
						
		IF EXISTS(SELECT U_DuexDate FROM [@PWC_TRDBT1] WHERE Code=@list_of_cols_val_tab_del AND U_DuexDate IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'[��ȯ������ ��] - ��ȯ�������� �Է����ּ���.'
			RETURN
		END
		
		IF EXISTS(SELECT U_DuexDate FROM [@PWC_TRDBT1] WHERE Code=@list_of_cols_val_tab_del GROUP BY U_DuexDate HAVING COUNT(U_DuexDate) > 1)
		BEGIN
			SET @error = -1
			SET @error_message = N'[��ȯ������ ��] - �ߺ��Ǵ� ��ȯ�������� �����մϴ�.'
			RETURN
		END
		
		IF EXISTS(SELECT U_FromDate FROM [@PWC_TRDBT2] WHERE Code=@list_of_cols_val_tab_del AND U_FromDate IS NULL OR U_ToxxDate IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'[���������� ��] - ������ ���� �Ⱓ�� �Է����ּ���.'
			RETURN
		END
		
		IF EXISTS(	SELECT LineId
					  FROM (
						 SELECT (SELECT S1.LineId FROM [@PWC_TRDBT2] S1 WHERE S1.Code=@list_of_cols_val_tab_del AND S1.LineId<>DBT2.LineId AND S1.U_FromDate<=DBT2.U_ToxxDate AND S1.U_ToxxDate>=DBT2.U_FromDate) AS LineId
						   FROM [@PWC_TRDBT2] DBT2
						  WHERE DBT2.Code=@list_of_cols_val_tab_del
				           ) P1
					 WHERE LineId IS NOT NULL  )
		BEGIN
			SET @error = -1
			SET @error_message = N'[���������� ��] - �ߺ��Ǵ� ������ ���� �Ⱓ�� �����մϴ�.'
			RETURN
		END
	END
	
	/* ���� ��� */
	IF @transaction_type = 'D'
	BEGIN
		IF EXISTS(SELECT TOP 1 TransId FROM JDT1 WITH (NOLOCK) WHERE U_PWC_DebtCode=@list_of_cols_val_tab_del)
		BEGIN
			SET @error = -1
			SET @error_message = N'�ش� ���Ա� ������ �����͸� �̹� �ٸ� ��ǥ���� ��� ���̹Ƿ� ������ �� �����ϴ�.'
			RETURN
		END
	END
END


END