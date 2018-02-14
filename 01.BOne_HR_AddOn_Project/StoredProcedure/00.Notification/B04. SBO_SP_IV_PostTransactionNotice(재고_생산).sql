IF OBJECT_ID('SBO_SP_IV_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_IV_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_IV_PostTransactionNotice
 ��      �� : ������, ������� PostTransactionNotice
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_IV_PostTransactionNotice] 

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


/*** OIGN : ������ - �԰� ***/
IF @object_type = '59'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM OIGN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** OIGE : ������ - ��� ***/
IF @object_type = '60'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM OIGE
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** OWTQ : ������ - ��� ���� ��û ***/
IF @object_type = '1250000001'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �԰� ������� NULL�� ��� ��� ������� �԰� ��������� ����
		UPDATE OWTR
		   SET U_PWC_BPLId2 = U_PWC_BPLId
		 WHERE DocEntry = @list_of_cols_val_tab_del
		   AND U_PWC_BPLId2 IS NULL
	END
	
	RETURN	-- ����
END


/*** OWTR : ������ - ������� ***/
IF @object_type = '67'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		DECLARE @intBPLId2	INT	-- �԰� �����
				
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
				, @intBPLId2=U_PWC_BPLId2
		  FROM OWTR
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
		
		-- �԰� ������� NULL�� ��� ��� ������� �԰� ��������� ����
		IF @intBPLId2 IS NULL
		BEGIN
			SET @intBPLId2 = @m_intBPLId
			
			UPDATE OWTR
			   SET U_PWC_BPLId2 = @intBPLId2
			 WHERE DocEntry = @list_of_cols_val_tab_del
		END
		
		-- ��� ������ �԰� ������� �ٸ� ���
		IF @m_intBPLId <> @intBPLId2 
		BEGIN
			UPDATE JDT1
			   SET U_PWC_BpliCode = @intBPLId2
			 WHERE TransId = @m_intTransId
			   AND ISNULL(Debit, 0) <> 0
		END
	END
	
	RETURN	-- ����
END


/*** OMRV : ������ - ������� ***/
IF @object_type = '162'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* �а� - ����� ����
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM OMRV
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- ����
END


/*** OWOR : ������� - ������� ***/
IF @object_type = '202'
BEGIN
	/* ���� ��� */
	IF @transaction_type = 'U'
	BEGIN
		DECLARE   @OWOR_chrCurrIsCmpltW	CHAR(1)
				, @OWOR_chrWillIsCmpltW	CHAR(1)
		
		--* ���� �Ϸ� ���� ó��
		SELECT    @OWOR_chrCurrIsCmpltW = ISNULL(U_IsCmpltW, 'N')
				, @OWOR_chrWillIsCmpltW = CASE WHEN ISNULL(CmpltQty, 0)+ISNULL(RjctQty, 0) >= PlannedQty THEN 'Y'
										       ELSE 'N'
								           END
		  FROM OWOR
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		IF @OWOR_chrCurrIsCmpltW <> @OWOR_chrWillIsCmpltW
		BEGIN
			UPDATE OWOR
			   SET U_IsCmpltW = @OWOR_chrWillIsCmpltW
			 WHERE DocEntry = @list_of_cols_val_tab_del
		END
	END
	
	RETURN	-- ����
END


/*** MDC_MM_CPD103 : ������� - �۾��ҷ�/Loss ��� ***/
IF @object_type = 'MDC_MM_CPD103'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN		
		--* ��� ǰ����� -> ����ҷ� �� �ν� ��Ͽ� ���� (��Ҹ� ����)
		UPDATE CPD103L 
		   SET    U_Price=OINM.CalcPrice
				, U_LineTotl=ABS(OINM.TransValue)
		  FROM [@MDC_MM_CPD103H] CPD103H
		 INNER JOIN [@MDC_MM_CPD103L] CPD103L ON CPD103H.DocEntry=CPD103L.DocEntry
		 INNER JOIN OINM WITH (NOLOCK) ON OINM.TransType='60' AND CPD103H.U_OutEntry=OINM.CreatedBy AND (CPD103L.LineId-1)=OINM.DocLineNum
		 WHERE CPD103H.DocEntry = @list_of_cols_val_tab_del
		   AND ISNULL(CPD103L.U_ItemCode, '') <> ''		 
	END
	
	RETURN	-- ����
END


/*** MDC_MM_CPD105 : ������� - �۾� �ܷ� �԰� ��� ***/
IF @object_type = 'MDC_MM_CPD105'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN		
		--* �԰� �߻�(����) ǰ����� -> �۾� �ܷ� �԰� (�԰� ��� ǰ������� ǥ���ϱ� ����)
		UPDATE CPD105L
		   SET    U_Price=OINM.CalcPrice
				, U_LineTotl=OINM.TransValue
		  FROM [@MDC_MM_CPD105H] CPD105H
		 INNER JOIN [@MDC_MM_CPD105L] CPD105L ON CPD105H.DocEntry=CPD105L.DocEntry
		 INNER JOIN OINM WITH (NOLOCK) ON OINM.TransType='59' AND CPD105H.U_InEntry=OINM.CreatedBy AND (CPD105L.LineId-1)=OINM.DocLineNum
		 WHERE CPD105H.DocEntry = @list_of_cols_val_tab_del
		   AND ISNULL(CPD105L.U_ItemCode, '') <> ''
	END
	
	RETURN	-- ����
END


/*** MDC_MM_CPD106 : ������� - ���� ���� ���� ���[PVC, �˹̴�] ***/
IF @object_type = 'MDC_MM_CPD106'
BEGIN	
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN		
		--* ���κ� ���� ���� ���� ó�� ���ο� ���� ���� �Ϸ� ���� �� ����
		IF NOT EXISTS ( 
			SELECT OCPD.DocEntry
			  FROM [@MDC_MM_CPD106H] AS OCPD
			 INNER JOIN [@MDC_MM_CPD106L] AS CPD1 ON OCPD.DocEntry=CPD1.DocEntry
			 WHERE OCPD.DocEntry = @list_of_cols_val_tab_del 
			   AND ISNULL(OCPD.U_IsCmplt1, 'N') = 'N'
			   AND CPD1.U_WREntry IS NULL
		)
		BEGIN
			UPDATE [@MDC_MM_CPD106H]
			   SET U_IsCmplt1 = 'Y'
			 WHERE DocEntry = @list_of_cols_val_tab_del
		END
	END
	
	/* ��� ��� */
	IF @transaction_type = 'C'
	BEGIN
		--* ���� ���� ���� �۾� ���� ���� ����(�߰��ÿ��� �ҽ����� ó��)
		UPDATE [@MDC_MM_CPD206H]
		   SET    U_OWORYN = 'N'
				, U_OWORDate = NULL
		 WHERE DocEntry IN (SELECT U_MCEntry FROM [@MDC_MM_CPD106L] WHERE DocEntry = @list_of_cols_val_tab_del AND ISNULL(U_WREntry, -1) <> -1)
	END
	
	RETURN	-- ����
END


END