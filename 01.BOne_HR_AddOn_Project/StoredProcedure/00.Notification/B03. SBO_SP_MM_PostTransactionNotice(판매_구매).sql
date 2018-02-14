IF OBJECT_ID('SBO_SP_MM_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_MM_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_MM_PostTransactionNotice
 설      명 : 판매관리, 구매관리 PostTransactionNotice
 작  성  자 : 
 일      시 : 
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
		
		RETURN		-- 작업 종료시 RETURN을 사용하여 SP EXIT (하나의 오브젝트에 대하여 종료 시점에 작성할 것)
	END
END
****************************************** Sample ********************************************************/
--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------
DECLARE	  @m_intTransId		INT		-- 분개 거래번호
		, @m_intBPLId		INT		-- 사업장

/*** ODLN : 판매관리 - 납품 ***/
IF @object_type = '15'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ODLN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
				
	END
		
	RETURN	-- 종료
END


/*** ORDN : 판매관리 - 반품 ***/
IF @object_type = '16'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ORDN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** ODPI : 판매관리 - A/R선금요청 ***/
IF @object_type = '203'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ODPI
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** OINV : 판매관리 - A/R송장 ***/
IF @object_type = '13'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM OINV
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** ORIN : 판매관리 - A/R대변메모 ***/
IF @object_type = '14'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ORIN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** OPDN : 구매관리 - 입고PO ***/
IF @object_type = '20'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM OPDN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** ORPD : 구매관리 - 반품 ***/
IF @object_type = '21'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ORPD
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** ODPO : 구매관리 - A/P선금요청 ***/
IF @object_type = '204'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ODPO
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** OPCH : 구매관리 - A/P송장 ***/
IF @object_type = '18'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM OPCH
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN		
		--* 품의 작성자, 승인자 정보 -> 분개 갱신		
		IF (SELECT ISNULL(U_PWC_RptNumb, '') FROM OPCH WHERE DocEntry=@list_of_cols_val_tab_del) <> ''
		BEGIN
			EXEC PWC_SP_SetExpRptInfo @object_type, @list_of_cols_val_tab_del
		END
	END
	
	RETURN	-- 종료
END


/*** ORPC : 구매관리 - A/P대변메모 ***/
IF @object_type = '19'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=BPLId
		  FROM ORPC
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN		
		--* 품의 작성자, 승인자 정보 -> 분개 갱신
		IF (SELECT ISNULL(U_PWC_RptNumb, '') FROM ORPC WHERE DocEntry=@list_of_cols_val_tab_del) <> ''
		BEGIN
			EXEC PWC_SP_SetExpRptInfo @object_type, @list_of_cols_val_tab_del
		END
	END
	
	RETURN	-- 종료
END


/*** OIPF : 구매관리 - 수입부대비용 ***/
IF @object_type = '69'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=JdtNum
				, @m_intBPLId=U_PWC_BPLId
		  FROM OIPF
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END



END