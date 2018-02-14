IF OBJECT_ID('SBO_SP_BP_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_BP_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_BP_TransactionNotification
 설      명 : 비즈니스파트너관리 TransactionNotification
 작  성  자 : 
 일      시 : 
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
			SET @error = -1									-- Error Code 설정
			SET @error_message = N'코드를 입력해 주세요.'		-- Error Message 설정
			RETURN											-- Error 발생시 RETURN을 사용하여 SP EXIT
		END
	END
END
****************************************** Sample ********************************************************/
--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------

/*** OCRD : 비즈니스파트너 마스터 데이터 ***/
IF @object_type = '2'
BEGIN
	DECLARE   @OCRD_nvcGroupCode	NVARCHAR(6)		-- BP 그룹 코드
			, @OCRD_chrCDMIsAuto	CHAR(1)			-- BP 코드 관리 - 자동채번여부
			, @OCRD_insCDMCodeLeng	SMALLINT		-- BP 코드 관리 - 채번 총길이			
			
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* BP 코드 관리 정의 여부 검사
		SELECT @OCRD_nvcGroupCode=GroupCode
		  FROM OCRD
		 WHERE CardCode = @list_of_cols_val_tab_del
		
		SELECT    @OCRD_chrCDMIsAuto=U_IsAutoNo
				, @OCRD_insCDMCodeLeng=ISNULL(U_CodeLeng, 0)				
		  FROM [@PWC_BPOCDM] WITH (NOLOCK)
		 WHERE RIGHT(CODE, LEN(CODE) - 1) = @OCRD_nvcGroupCode
		
		-- BP 그룹이 코드 관리 테이블에 정상적으로 등록되지 않은 경우 (PostTran...에서 처리되는 부분)
		IF @OCRD_chrCDMIsAuto IS NULL
		BEGIN
			SET @error = -1
			SET @error_message = N'해당 비즈니스 파트너 그룹에 대한 비즈니스 파트너 코드 관리 규칙이 존재하지 않았습니다. 관리자에게 문의해 주세요.'
			RETURN
		END
		
		-- 자동 채번을 사용하면서 채번 총길이가 정의되지 않은 경우
		IF @OCRD_chrCDMIsAuto = 'Y' AND @OCRD_insCDMCodeLeng <= 0
		BEGIN
			SET @error = -1
			SET @error_message = N'해당 비즈니스 파트너 그룹에 대한 비즈니스 파트너 코드 관리 규칙이 정상적으로 정의되지 않았습니다. [기본서식에서 정의]'
			RETURN
		END
		
		-- 코드 길이 체크
		IF @OCRD_chrCDMIsAuto = 'Y' AND LEN(@list_of_cols_val_tab_del) <> @OCRD_insCDMCodeLeng
		BEGIN
			SET @error = -1
			SET @error_message = N'해당 비즈니스 파트너 그룹의 코드 길이는 ' + CONVERT(NVARCHAR(6), @OCRD_insCDMCodeLeng) + '입니다.'
			RETURN
		END
	END
END


/*** [@PWC_BPOCDM] : 비즈니스파트너 코드 관리 ***/
IF @object_type = 'PWC_UDO_BPOCDM'
BEGIN
	/* 갱신 모드 */
	IF @transaction_type = 'U'
	BEGIN
		--* BP 코드 관리 기준 검사
		IF EXISTS(	SELECT Code
					  FROM [@PWC_BPOCDM]
					 WHERE CODE = @list_of_cols_val_tab_del
					   AND U_IsAutoNo = 'Y'
					   AND ISNULL(U_CodeLeng, 0) <= 0	)
		BEGIN
			SET @error = -1
			SET @error_message = N'자동 채번을 사용하는 경우 채번 총길이는 필수입니다. [' + @list_of_cols_val_tab_del + ']'
			RETURN
		END
		
		IF EXISTS(	SELECT Code
					  FROM [@PWC_BPOCDM]
					 WHERE CODE = @list_of_cols_val_tab_del
					   AND U_IsAutoNo = 'N'
					   AND U_CodeLeng IS NOT NULL	)
		BEGIN
			SET @error = -1
			SET @error_message = N'자동 채번을 사용하지 않는 경우 채번 총길이는 입력할 수 없습니다. [' + @list_of_cols_val_tab_del + ']'
			RETURN
		END
	END
END


END