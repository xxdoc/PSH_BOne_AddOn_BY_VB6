IF OBJECT_ID('SBO_SP_CO_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_CO_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_CO_PostTransactionNotice
 ��      �� : ���� ȸ�� PostTransactionNotice
 ��  ��  �� : 
 ��      �� : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_CO_PostTransactionNotice] 

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


END