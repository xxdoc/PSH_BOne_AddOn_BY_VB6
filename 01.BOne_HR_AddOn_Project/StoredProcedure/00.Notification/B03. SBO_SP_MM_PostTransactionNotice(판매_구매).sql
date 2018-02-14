IF OBJECT_ID('SBO_SP_MM_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_MM_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_MM_PostTransactionNotice
 ��      �� : �ǸŰ���, ���Ű��� PostTransactionNotice
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_MM_PostTransactionNotice] 

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

/*** ODLN : �ǸŰ��� - ��ǰ ***/
IF @object_type = '15'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ODLN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
				
	END
		
	RETURN	-- ����
END


/*** ORDN : �ǸŰ��� - ��ǰ ***/
IF @object_type = '16'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ORDN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** ODPI : �ǸŰ��� - A/R���ݿ�û ***/
IF @object_type = '203'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ODPI
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** OINV : �ǸŰ��� - A/R���� ***/
IF @object_type = '13'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM OINV
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** ORIN : �ǸŰ��� - A/R�뺯�޸� ***/
IF @object_type = '14'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ORIN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** OPDN : ���Ű��� - �԰�PO ***/
IF @object_type = '20'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM OPDN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** ORPD : ���Ű��� - ��ǰ ***/
IF @object_type = '21'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ORPD
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** ODPO : ���Ű��� - A/P���ݿ�û ***/
IF @object_type = '204'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ODPO
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** OPCH : ���Ű��� - A/P���� ***/
IF @object_type = '18'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM OPCH
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN		
		--* ǰ�� �ۼ���, ������ ���� -> �а� ����		
		IF (SELECT ISNULL(U_PWC_RptNumb, '') FROM OPCH WHERE DocEntry=@list_of_cols_val_tab_del) <> ''
		BEGIN
			EXEC PWC_SP_SetExpRptInfo @object_type, @list_of_cols_val_tab_del
		END
	END
	
	RETURN	-- ����
END


/*** ORPC : ���Ű��� - A/P�뺯�޸� ***/
IF @object_type = '19'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ORPC
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN		
		--* ǰ�� �ۼ���, ������ ���� -> �а� ����
		IF (SELECT ISNULL(U_PWC_RptNumb, '') FROM ORPC WHERE DocEntry=@list_of_cols_val_tab_del) <> ''
		BEGIN
			EXEC PWC_SP_SetExpRptInfo @object_type, @list_of_cols_val_tab_del
		END
	END
	
	RETURN	-- ����
END


/*** OIPF : ���Ű��� - ���Ժδ��� ***/
IF @object_type = '69'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=JdtNum
				, @m_intBPLId=U_PWC_BPLId
		  FROM OIPF
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END



END