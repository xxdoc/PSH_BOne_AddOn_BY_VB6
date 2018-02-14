IF OBJECT_ID('SBO_SP_BP_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_BP_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_BP_PostTransactionNotice
 설      명 : 비즈니스파트너관리 PostTransactionNotice
 작  성  자 : 
 일      시 : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_BP_PostTransactionNotice] 

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
		
		RETURN		-- 작업 종료시 RETURN을 사용하여 SP EXIT (하나의 오브젝트에 대하여 종료 시점에 작성할 것)
	END
END
****************************************** Sample ********************************************************/
--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------

/*** OCRD : 비즈니스파트너 마스터 데이터 ***/
IF @object_type = '2'
BEGIN
	DECLARE @OCRD_nvcLIsTotlColt	NVARCHAR(1)
	
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* BP & 차입금 마스터 연동
		SELECT @OCRD_nvcLIsTotlColt = ISNULL(U_PWC_LIsTotlColt, 'N')
		  FROM OCRD
		 WHERE CardCode = @list_of_cols_val_tab_del
		
		UPDATE [@PWC_TRODBT]
		   SET U_IsTotlColt = @OCRD_nvcLIsTotlColt
		 WHERE U_CardCode = @list_of_cols_val_tab_del
			
		RETURN
	END
END

END