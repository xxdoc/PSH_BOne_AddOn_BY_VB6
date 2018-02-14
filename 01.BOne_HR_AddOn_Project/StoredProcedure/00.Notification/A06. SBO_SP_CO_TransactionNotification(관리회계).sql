IF OBJECT_ID('SBO_SP_CO_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_CO_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_CO_TransactionNotification
 ��      �� : ���� ȸ�� TransactionNotification
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_CO_TransactionNotification] 

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



END