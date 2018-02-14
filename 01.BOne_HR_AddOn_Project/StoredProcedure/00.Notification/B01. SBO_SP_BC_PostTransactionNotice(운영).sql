IF OBJECT_ID('SBO_SP_BC_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_BC_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_BC_PostTransactionNotice
 ��      �� : ����� PostTransactionNotice
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_BC_PostTransactionNotice] 

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
DECLARE	@m_intDocEntry	INT

/*** OBPL : ����� ***/
IF @object_type = '247'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		--* �������̺�� ����� ����
		SELECT @m_intDocEntry = AutoKey
		  FROM ONNM WITH (NOLOCK)
		 WHERE ObjectCode = 'PWC_UDO_BCOBPL'
		
		INSERT INTO [@PWC_BCOBPL](Code, Name, DocEntry, [Object], LogInst
								, UserSign, CreateDate, CreateTime, DataSource)
		SELECT    BPLId
				, BPLName
				, @m_intDocEntry
				, 'PWC_UDO_BCOBPL'
				, 0
				, UserSign2
				, CONVERT(NVARCHAR(8), GETDATE(), 112)
				, REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
				, 'I'
		  FROM OBPL
		 WHERE BPLId = @list_of_cols_val_tab_del
		
		UPDATE ONNM
		   SET AutoKey=@m_intDocEntry+1
		 WHERE ObjectCode = 'PWC_UDO_BCOBPL'
		
		RETURN		-- �۾� ����� RETURN�� ����Ͽ� SP EXIT (�ϳ��� ������Ʈ�� ���Ͽ� ���� ������ �ۼ��� ��)
	END
	
	/* ���� ��� */
	IF @transaction_type = 'U'
	BEGIN
	
		--* �������̺�� ����� ����
		UPDATE BCOBPL SET Name=BPLName, UpdateDate=CONVERT(NVARCHAR(8), GETDATE(), 112), UpdateTime=REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
		  FROM OBPL
		 INNER JOIN [@PWC_BCOBPL] BCOBPL ON OBPL.BPLId=BCOBPL.Code
		 WHERE OBPL.BPLId = @list_of_cols_val_tab_del
		
		RETURN
	END
	
	/* ���� ��� */
	IF @transaction_type = 'D'
	BEGIN
	
		--* �������̺�� ����� ����
		DELETE [@PWC_BCOBPL]
		 WHERE Code = @list_of_cols_val_tab_del
		
		RETURN
	END
END

/*** OCRG : ����Ͻ���Ʈ�� �׷� ***/
IF @object_type = '10'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		--* BP �ڵ� ä�� ��Ģ ���̺� BP �׷� �߰�
		SELECT @m_intDocEntry = AutoKey
		  FROM ONNM WITH (NOLOCK)
		 WHERE ObjectCode = 'PWC_UDO_BPOCDM'

		INSERT INTO [@PWC_BPOCDM] (Code, Name, DocEntry, [Object], UserSign
							, CreateDate, CreateTime, DataSource)
		SELECT    GroupType + CONVERT(NVARCHAR(6), GroupCode)
				, GroupName
				, @m_intDocEntry
				, 'PWC_UDO_BPOCDM'
				, UserSign
				, CONVERT(NVARCHAR(8), GETDATE(), 112)
				, REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
				, 'I'
		  FROM OCRG
		 WHERE GroupCode = @list_of_cols_val_tab_del
		 
		UPDATE ONNM
		   SET AutoKey=@m_intDocEntry+1
		 WHERE ObjectCode = 'PWC_UDO_BPOCDM'
		
		RETURN
	END
	
	/* ���� ��� */
	IF @transaction_type = 'U'
	BEGIN
		--* BP �ڵ� ä�� ��Ģ ���̺� BP �׷� ����
		UPDATE [@PWC_BPOCDM]
		   SET Name=(SELECT TOP 1 GroupName FROM OCRG WHERE GroupCode = @list_of_cols_val_tab_del)
		 WHERE RIGHT(Code, LEN(Code) - 1) = @list_of_cols_val_tab_del
		
		RETURN
	END
	
	/* ���� ��� */
	IF @transaction_type = 'D'
	BEGIN
		--* BP �ڵ� ä�� ��Ģ ���̺� BP �׷� ����
		DELETE [@PWC_BPOCDM]
		 WHERE RIGHT(Code, LEN(Code) - 1) = @list_of_cols_val_tab_del
		
		RETURN
	END
END

/*** OALC : ���Ժδ��� ���� ***/
IF @object_type = '48'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		--* �������̺�� ���Ժδ��� ���� ����
		SELECT @m_intDocEntry = AutoKey
		  FROM ONNM WITH (NOLOCK)
		 WHERE ObjectCode = 'PWC_UDO_BCOALC'
		
		INSERT INTO [@PWC_BCOALC](Code, Name, DocEntry, [Object], LogInst
								, UserSign, CreateDate, CreateTime, DataSource)
		SELECT    AlcCode
				, AlcName
				, @m_intDocEntry
				, 'PWC_UDO_BCOALC'
				, 0
				, UserSign
				, CONVERT(NVARCHAR(8), GETDATE(), 112)
				, REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
				, DataSource
		  FROM OALC
		 WHERE AlcCode = @list_of_cols_val_tab_del
		
		UPDATE ONNM
		   SET AutoKey=@m_intDocEntry+1
		 WHERE ObjectCode = 'PWC_UDO_BCOALC'
		
		RETURN		-- �۾� ����� RETURN�� ����Ͽ� SP EXIT (�ϳ��� ������Ʈ�� ���Ͽ� ���� ������ �ۼ��� ��)
	END
	
	/* ���� ��� */
	IF @transaction_type = 'U'
	BEGIN
	
		--* �������̺�� ���Ժδ��� ���� ����
		UPDATE BCOALC SET Name=AlcName, UpdateDate=CONVERT(NVARCHAR(8), GETDATE(), 112), UpdateTime=REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
		  FROM OALC
		 INNER JOIN [@PWC_BCOALC] BCOALC ON OALC.AlcCode=BCOALC.Code
		 WHERE OALC.AlcCode = @list_of_cols_val_tab_del
		
		RETURN
	END
	
	/* ���� ��� */
	IF @transaction_type = 'D'
	BEGIN
	
		--* �������̺�� ���Ժδ��� ���� ����
		DELETE [@PWC_BCOALC]
		 WHERE Code = @list_of_cols_val_tab_del
		
		RETURN
	END
END

/*** OPYM : ���� ��� ***/
IF @object_type = '147'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		--* �������̺�� ���� ��� ����
		SELECT @m_intDocEntry = AutoKey
		  FROM ONNM WITH (NOLOCK)
		 WHERE ObjectCode = 'PWC_UDO_BCOPYM'
		
		INSERT INTO [@PWC_BCOPYM](Code, Name, DocEntry, [Object], LogInst
								, UserSign, CreateDate, CreateTime, DataSource, U_Active)
		SELECT    PayMethCod
				, Descript
				, @m_intDocEntry
				, 'PWC_UDO_BCOPYM'
				, 0
				, UserSign
				, CONVERT(NVARCHAR(8), GETDATE(), 112)
				, REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
				, DataSource
				, Active
		  FROM OPYM
		 WHERE PayMethCod = @list_of_cols_val_tab_del
		
		UPDATE ONNM
		   SET AutoKey=@m_intDocEntry+1
		 WHERE ObjectCode = 'PWC_UDO_BCOPYM'
		
		RETURN		-- �۾� ����� RETURN�� ����Ͽ� SP EXIT (�ϳ��� ������Ʈ�� ���Ͽ� ���� ������ �ۼ��� ��)
	END
	
	/* ���� ��� */
	IF @transaction_type = 'U'
	BEGIN
	
		--* �������̺�� ���� ��� ����
		UPDATE BCOPYM SET Name=Descript, U_Active=Active, UpdateDate=CONVERT(NVARCHAR(8), GETDATE(), 112), UpdateTime=REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
		  FROM OPYM
		 INNER JOIN [@PWC_BCOPYM] BCOPYM ON OPYM.PayMethCod=BCOPYM.Code
		 WHERE OPYM.PayMethCod = @list_of_cols_val_tab_del
		
		RETURN
	END
	
	/* ���� ��� */
	IF @transaction_type = 'D'
	BEGIN
	
		--* �������̺�� ���� ��� ����
		DELETE [@PWC_BCOPYM]
		 WHERE Code = @list_of_cols_val_tab_del
		
		RETURN
	END
END

END