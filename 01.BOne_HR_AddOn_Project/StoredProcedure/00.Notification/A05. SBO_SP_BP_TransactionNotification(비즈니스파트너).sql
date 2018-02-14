IF OBJECT_ID('SBO_SP_BP_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_BP_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_BP_TransactionNotification
 ��      �� : ����Ͻ���Ʈ�ʰ��� TransactionNotification
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_BP_TransactionNotification] 

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
			SET @error = -1									-- Error Code ����
			SET @error_message = N'�ڵ带 �Է��� �ּ���.'		-- Error Message ����
			RETURN											-- Error �߻��� RETURN�� ����Ͽ� SP EXIT
		END
	END
END
****************************************** Sample ********************************************************/
--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------

/*** OCRD : ����Ͻ���Ʈ�� ������ ������ ***/
IF @object_type = '2'
BEGIN
	DECLARE   @OCRD_nvcGroupCode	NVARCHAR(6)		-- BP �׷� �ڵ�
			, @OCRD_chrCDMIsAuto	CHAR(1)			-- BP �ڵ� ���� - �ڵ�ä������
			, @OCRD_insCDMCodeLeng	SMALLINT		-- BP �ڵ� ���� - ä�� �ѱ���			
			
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* BP �ڵ� ���� ���� ���� �˻�
		SELECT @OCRD_nvcGroupCode=GroupCode
		  FROM OCRD
		 WHERE CardCode = @list_of_cols_val_tab_del
		
		SELECT    @OCRD_chrCDMIsAuto=U_IsAutoNo
				, @OCRD_insCDMCodeLeng=ISNULL(U_CodeLeng, 0)				
		  FROM [@PWC_BPOCDM] WITH (NOLOCK)
		 WHERE RIGHT(CODE, LEN(CODE) - 1) = @OCRD_nvcGroupCode
		
		-- BP �׷��� �ڵ� ���� ���̺� ���������� ��ϵ��� ���� ��� (PostTran...���� ó���Ǵ� �κ�)
		IF @OCRD_chrCDMIsAuto IS NULL
		BEGIN
			SET @error = -1
			SET @error_message = N'�ش� ����Ͻ� ��Ʈ�� �׷쿡 ���� ����Ͻ� ��Ʈ�� �ڵ� ���� ��Ģ�� �������� �ʾҽ��ϴ�. �����ڿ��� ������ �ּ���.'
			RETURN
		END
		
		-- �ڵ� ä���� ����ϸ鼭 ä�� �ѱ��̰� ���ǵ��� ���� ���
		IF @OCRD_chrCDMIsAuto = 'Y' AND @OCRD_insCDMCodeLeng <= 0
		BEGIN
			SET @error = -1
			SET @error_message = N'�ش� ����Ͻ� ��Ʈ�� �׷쿡 ���� ����Ͻ� ��Ʈ�� �ڵ� ���� ��Ģ�� ���������� ���ǵ��� �ʾҽ��ϴ�. [�⺻���Ŀ��� ����]'
			RETURN
		END
		
		-- �ڵ� ���� üũ
		IF @OCRD_chrCDMIsAuto = 'Y' AND LEN(@list_of_cols_val_tab_del) <> @OCRD_insCDMCodeLeng
		BEGIN
			SET @error = -1
			SET @error_message = N'�ش� ����Ͻ� ��Ʈ�� �׷��� �ڵ� ���̴� ' + CONVERT(NVARCHAR(6), @OCRD_insCDMCodeLeng) + '�Դϴ�.'
			RETURN
		END
	END
END


/*** [@PWC_BPOCDM] : ����Ͻ���Ʈ�� �ڵ� ���� ***/
IF @object_type = 'PWC_UDO_BPOCDM'
BEGIN
	/* ���� ��� */
	IF @transaction_type = 'U'
	BEGIN
		--* BP �ڵ� ���� ���� �˻�
		IF EXISTS(	SELECT Code
					  FROM [@PWC_BPOCDM]
					 WHERE CODE = @list_of_cols_val_tab_del
					   AND U_IsAutoNo = 'Y'
					   AND ISNULL(U_CodeLeng, 0) <= 0	)
		BEGIN
			SET @error = -1
			SET @error_message = N'�ڵ� ä���� ����ϴ� ��� ä�� �ѱ��̴� �ʼ��Դϴ�. [' + @list_of_cols_val_tab_del + ']'
			RETURN
		END
		
		IF EXISTS(	SELECT Code
					  FROM [@PWC_BPOCDM]
					 WHERE CODE = @list_of_cols_val_tab_del
					   AND U_IsAutoNo = 'N'
					   AND U_CodeLeng IS NOT NULL	)
		BEGIN
			SET @error = -1
			SET @error_message = N'�ڵ� ä���� ������� �ʴ� ��� ä�� �ѱ��̴� �Է��� �� �����ϴ�. [' + @list_of_cols_val_tab_del + ']'
			RETURN
		END
	END
END


END