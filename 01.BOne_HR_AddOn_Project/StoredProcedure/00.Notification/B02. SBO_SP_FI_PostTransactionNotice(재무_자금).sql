IF OBJECT_ID('SBO_SP_FI_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_FI_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_FI_PostTransactionNotice
 ��      �� : �繫����, �ڱݰ��� PostTransactionNotice
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_FI_PostTransactionNotice] 

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
		DELETE [@PWC_TRDBT3] WHERE Code IN (1, 2)
		
		RETURN		-- �۾� ����� RETURN�� ����Ͽ� SP EXIT (�ϳ��� ������Ʈ�� ���Ͽ� ���� ������ �ۼ��� ��)
	END
END
****************************************** Sample ********************************************************/
--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------
DECLARE	  @m_intTransId		INT		-- �а� �ŷ���ȣ
		, @m_intBPLId		INT		-- �����
		

/*** OJDT : �а� ***/
IF @object_type = '30'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* ���Ա� ������ ��ǥ ���� ���
		DECLARE @tbl_OJDT_ODBT_1 TABLE
		(
			DebtCode	NVARCHAR(15)
		)
		
		-- 1. ���� ��� ���Ա� ���� ��ȣ ����Ʈ ����
		INSERT INTO @tbl_OJDT_ODBT_1 (DebtCode)
		SELECT DebtCode
		  FROM (
			SELECT U_PWC_DebtCode AS DebtCode
			  FROM JDT1
			 WHERE TransId = @list_of_cols_val_tab_del		   
			   AND U_PWC_DebtCode <> ''
			 UNION ALL
			SELECT ODBT.Code		-- �а� - ���Ա� ������ȣ �Է� �� ���� ��쿡 ���� ó��
			  FROM [@PWC_TRODBT] ODBT WITH (NOLOCK)
			 INNER JOIN [@PWC_TRDBT3] DBT3 WITH (NOLOCK) ON ODBT.Code=DBT3.Code
			 WHERE DBT3.U_TransIdx = @list_of_cols_val_tab_del
			 ) A
		 GROUP BY DebtCode
		
		IF (SELECT COUNT(*) FROM @tbl_OJDT_ODBT_1) > 0
		BEGIN
			-- 2. ���Ա� ������ ������ȣ ����
			DELETE [@PWC_TRDBT3]
			 WHERE Code IN (SELECT DebtCode FROM @tbl_OJDT_ODBT_1)
			
			-- 3. ���Ա� ������ - ���� ��ǥ ���
			INSERT INTO [@PWC_TRDBT3] (	
					  Code, LineId, Object, U_TransIdx, U_JdtLineId					-- 5
					, U_RefDate, U_CardCode, U_CardName, U_Account, U_AcctName		-- 10
					, U_Debit, U_Credit, U_FCDebit, U_FCCredit)
			SELECT    JDT1.U_PWC_DebtCode
					, ROW_NUMBER() OVER (PARTITION BY JDT1.U_PWC_DebtCode ORDER BY JDT1.RefDate, JDT1.TransId, JDT1.Line_ID)
					, 'PWC_UDO_TRODBT'
					, JDT1.TransId
					, JDT1.Line_ID + 1 AS LineId
					, JDT1.RefDate
					, CASE WHEN OCRD.CardName IS NULL THEN '' ELSE JDT1.ShortName END
					, ISNULL(OCRD.CardName, '')
					, JDT1.Account
					, OACT.AcctName
					, ISNULL(Debit, 0) AS Debit
					, ISNULL(Credit, 0) AS Credit
					, ISNULL(FCDebit, 0) AS Debit										
					, ISNULL(FCCredit, 0) AS Credit
			  FROM JDT1
			 INNER JOIN OACT WITH (NOLOCK) ON JDT1.Account=OACT.AcctCode
			 INNER JOIN [@PWC_TRODBT] ODBT WITH (NOLOCK) ON JDT1.U_PWC_DebtCode=ODBT.Code
			  LEFT JOIN OCRD OCRD WITH (NOLOCK) ON JDT1.ShortName=OCRD.CardCode
			 WHERE JDT1.U_PWC_DebtCode IN (SELECT DebtCode FROM @tbl_OJDT_ODBT_1)			 
		END
	END
	
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN				
		--* ǰ�� �ۼ���, ������ �� ���� �۾�
		IF (SELECT ISNULL(U_PWC_RptNumb, '') FROM OJDT WHERE TransId=@list_of_cols_val_tab_del) <> ''
		BEGIN
			EXEC PWC_SP_SetExpRptInfo @object_type, @list_of_cols_val_tab_del
		END
	END
		
	RETURN	-- ����	
END


/*** ORCT : �ڱݰ��� - �Ա� ***/
IF @object_type = '24'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM ORCT
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END
	
	/* ��� ��� */
	IF @transaction_type = 'C'
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=OJDT.TransId
				, @m_intBPLId=ORCT.U_PWC_BPLId
		  FROM ORCT INNER JOIN OJDT ON ORCT.TransId=OJDT.StornoToTr
		 WHERE ORCT.DocEntry = @list_of_cols_val_tab_del
		  
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END
		
	RETURN	-- ����
END


/*** ODPS : �ڱݰ��� - ���� ***/
IF @object_type = '25'
BEGIN
	/* �߰�, ����, ��� ��� */
	IF @transaction_type IN ('A', 'U', 'C')
	BEGIN		
		--* ���� ó�� ������ ������ ��� RETURN (������������ ó��)
		IF (SELECT TOP 1 DeposType FROM ODPS WHERE DeposId = @list_of_cols_val_tab_del) = 'B'	
		BEGIN
			RETURN
		END
		
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransAbs
				, @m_intBPLId=U_PWC_BPLId
		  FROM ODPS
		 WHERE DeposId = @list_of_cols_val_tab_del

		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
		
	RETURN	-- ����
END


/*** OVPM : �ڱݰ��� - ���� ***/
IF @object_type = '46'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM OVPM
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END
	
	/* ��� ��� */
	IF @transaction_type = 'C'
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=OJDT.TransId
				, @m_intBPLId=OVPM.U_PWC_BPLId
		  FROM OVPM INNER JOIN OJDT ON OVPM.TransId=OJDT.StornoToTr
		 WHERE OVPM.DocEntry = @list_of_cols_val_tab_del
		  
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END
		
	RETURN	-- ����
END

/*** OBOT : �ڱݰ��� - �������� ***/
IF @object_type = '182'
BEGIN
	DECLARE   @OBOT_nvcBoeType			CHAR(1)			-- �Ա�(I), ����(O) ����
			, @OBOT_nvcBoeKey			INT				-- ���� Key
			, @OBOT_nvcTransactionRoot	NVARCHAR(10)	-- �������� ��� (FROM -> TO)
	
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		
		SELECT    TOP 1 @OBOT_nvcBoeKey = BOT1.BOEAbs
				, @m_intTransId = OBOT.TransId
				, @OBOT_nvcTransactionRoot = OBOT.StatusFrom + OBOT.StatusTo
				, @OBOT_nvcBoeType=BOT1.BoeType
		  FROM OBOT
		 INNER JOIN BOT1 ON OBOT.AbsEntry=BOT1.AbsEntry
		 WHERE OBOT.AbsEntry = @list_of_cols_val_tab_del
		
		IF @OBOT_nvcBoeType = 'I'		-- �Ա� (���� �����ÿ��� �Աݿ��� ����� ó��)
		BEGIN			
			-- �Ա� - ����� ������ �����´�.
			SELECT @m_intBPLId=ORCT.U_PWC_BPLId
			  FROM OBOE WITH (NOLOCK)
			 INNER JOIN ORCT WITH (NOLOCK) ON OBOE.PmntNum=ORCT.DocEntry
			 WHERE OBOE.BoeKey = @OBOT_nvcBoeKey			   
			
			IF @OBOT_nvcTransactionRoot IN ('GD', 'DG')	-- ���� -> ����, ���� -> ������ ��� ���ݿ� ����� ���� ó��
			BEGIN			
				 -- ���� ����� ����				 				
				UPDATE ODPS
				   SET U_PWC_BPLId=@m_intBPLId
				 WHERE TransAbs = @m_intTransId
			END
		END
		ELSE IF @OBOT_nvcBoeType = 'O'	-- ���� (���� �����ÿ��� ���޿��� ����� ó��)
		BEGIN
			-- ���� - ����� ������ �����´�.
			SELECT @m_intBPLId=OVPM.U_PWC_BPLId
			  FROM OBOE WITH (NOLOCK)
			 INNER JOIN OVPM WITH (NOLOCK) ON OBOE.PmntNum=OVPM.DocEntry
			 WHERE OBOE.BoeKey = @OBOT_nvcBoeKey			   
		END

		--* �а� - ����� ����
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END	
	
	RETURN	-- ����
END


/*** [@PWC_TRODBT] : ���Ա� ������ ��� ***/
IF @object_type = 'PWC_UDO_TRODBT'
BEGIN
	DECLARE	  @ODBT_nvcOCRDLIsTotlColt	NVARCHAR(1)
			, @ODBT_nvcCardCode			NVARCHAR(15)
			, @ODBT_nvcIsTotlColt		NVARCHAR(1)
	
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		--* BP & ���Ա� ������ ���� �㺸 ���� ����
		SELECT    @ODBT_nvcCardCode = U_CardCode
				, @ODBT_nvcIsTotlColt = ISNULL(U_IsTotlColt, 'N')
		  FROM [@PWC_TRODBT]
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		SELECT @ODBT_nvcOCRDLIsTotlColt = ISNULL(U_PWC_LIsTotlColt, 'N')
		  FROM OCRD WITH (NOLOCK)
		 WHERE CardCode = @ODBT_nvcCardCode
		
		IF @ODBT_nvcIsTotlColt <> @ODBT_nvcOCRDLIsTotlColt
		BEGIN
			UPDATE [@PWC_TRODBT]
			   SET U_IsTotlColt = @ODBT_nvcOCRDLIsTotlColt
			 WHERE DocEntry = @list_of_cols_val_tab_del
		END
	END
	
	RETURN -- ����
END

END