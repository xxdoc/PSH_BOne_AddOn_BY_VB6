IF OBJECT_ID('SBO_SP_BC_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_BC_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_BC_TransactionNotification
 ��      �� : ����� TransactionNotification
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_BC_TransactionNotification] 

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
/*** PWC_UDO_BCOCSY : �ý��۰����ڵ���(�繫) ***/
IF @object_type = 'PWC_UDO_BCOCSY'
BEGIN
	IF @transaction_type IN ('A', 'U')
	BEGIN
		IF EXISTS(SELECT U_CODE FROM [@PWC_BCCSY1] WHERE CODE=@list_of_cols_val_tab_del AND U_CODE IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'�ڵ带 �Է��� �ּ���.'
			RETURN
		END
		
		IF (SELECT COUNT(*) FROM [@PWC_BCCSY1] WHERE CODE=@list_of_cols_val_tab_del) = 0
		BEGIN
			SET @error = -1
			SET @error_message = N'[' + @list_of_cols_val_tab_del + N'] - �ý��� ���� �ڵ带 �ּ� �� ���̶� �Է��� �ּ���.'
			RETURN
		END
	
		IF EXISTS(SELECT U_CODE FROM [@PWC_BCCSY1] WHERE CODE=@list_of_cols_val_tab_del GROUP BY U_CODE HAVING COUNT(U_CODE) > 1)
		BEGIN
			SET @error = -1
			SET @error_message = N'[' + @list_of_cols_val_tab_del + N'] - �ý��� ���� �ڵ� �� �ߺ��Ǵ� �ڵ� ���� �����մϴ�.'
			RETURN
		END
	END
END

END