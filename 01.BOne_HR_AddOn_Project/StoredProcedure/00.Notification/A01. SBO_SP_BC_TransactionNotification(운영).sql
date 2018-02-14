IF OBJECT_ID('SBO_SP_BC_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_BC_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_BC_TransactionNotification
 설      명 : 운영관리 TransactionNotification
 작  성  자 : 
 일      시 : 
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
			SET @error = -1										-- Error Code 설정
			SET @error_message = N'코드를 입력해 주세요.'		-- Error Message 설정
			RETURN												-- Error 발생시 RETURN을 사용하여 SP EXIT
		END
	END
END
****************************************** Sample ********************************************************/
--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------
/*** PWC_UDO_BCOCSY : 시스템공통코드등록(재무) ***/
IF @object_type = 'PWC_UDO_BCOCSY'
BEGIN
	IF @transaction_type IN ('A', 'U')
	BEGIN
		IF EXISTS(SELECT U_CODE FROM [@PWC_BCCSY1] WHERE CODE=@list_of_cols_val_tab_del AND U_CODE IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'코드를 입력해 주세요.'
			RETURN
		END
		
		IF (SELECT COUNT(*) FROM [@PWC_BCCSY1] WHERE CODE=@list_of_cols_val_tab_del) = 0
		BEGIN
			SET @error = -1
			SET @error_message = N'[' + @list_of_cols_val_tab_del + N'] - 시스템 공통 코드를 최소 한 건이라도 입력해 주세요.'
			RETURN
		END
	
		IF EXISTS(SELECT U_CODE FROM [@PWC_BCCSY1] WHERE CODE=@list_of_cols_val_tab_del GROUP BY U_CODE HAVING COUNT(U_CODE) > 1)
		BEGIN
			SET @error = -1
			SET @error_message = N'[' + @list_of_cols_val_tab_del + N'] - 시스템 공통 코드 중 중복되는 코드 값이 존재합니다.'
			RETURN
		END
	END
END

END