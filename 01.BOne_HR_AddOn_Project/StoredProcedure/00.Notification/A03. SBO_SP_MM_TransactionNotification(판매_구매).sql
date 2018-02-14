IF OBJECT_ID('SBO_SP_MM_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_MM_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_MM_TransactionNotification
 설      명 : 판매관리, 구매관리 TransactionNotification
 작  성  자 : 
 일      시 : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_MM_TransactionNotification] 

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

/*** OIPF : 구매관리 - 수입부대비용 ***/
IF @object_type = '69'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 사업장 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OIPF WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END
	END
END

/*** OPCH : 구매관리 - A/P예약송장 ***/
IF @object_type = '18'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* AP예약송장일경우 BL번호 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OPCH WHERE DocEntry=@list_of_cols_val_tab_del AND U_POTYPE = '2' AND UPDINVNT = 'O' AND U_BL_NO IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'BL번호가 입력되지 않았습니다..'
			RETURN
		END
		--* AP예약송장일경우결제조건 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OPCH WHERE DocEntry=@list_of_cols_val_tab_del AND U_POTYPE = '2' AND UPDINVNT = 'O' AND U_SETTLE IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'결제조건이 입력되지 않았습니다..'
			RETURN
		END		
	END
	
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		DECLARE   @OPCH_nvcAcctCode	NVARCHAR(15)	-- 필수 체크용 AcctCode (시공비 유형, 현장 경비 유형)
		
		--* A/P 송장 - 서비스 - 시공비 유형, 현장 경비 유형 필수 여부 체크
		IF (SELECT DocType FROM OPCH WHERE DocEntry=@list_of_cols_val_tab_del) = 'S'
		BEGIN
			-- 시공비 유형 필수 누락 계정 추출
			SET @OPCH_nvcAcctCode = ''
			SELECT TOP 1 @OPCH_nvcAcctCode = PCH1.AcctCode
			  FROM PCH1
			 INNER JOIN OACT ON PCH1.AcctCode=OACT.AcctCode
			 WHERE PCH1.DocEntry=@list_of_cols_val_tab_del
			   AND OACT.U_PWC_MdInsExpTp = 'Y'
			   AND ISNULL(PCH1.U_PWC_PjInsExpTp, '') = ''
			   			
			IF @OPCH_nvcAcctCode <> ''
			BEGIN
				SET @error = -1
				SET @error_message = N'해당 계정의 경우 시공비 유형은 필수입니다. [G/L 계정 : ' + @OPCH_nvcAcctCode + ']'
				RETURN
			END
			
			-- 현장 경비 유형 필수 누락 계정 추출
			SET @OPCH_nvcAcctCode = ''
			SELECT TOP 1 @OPCH_nvcAcctCode = PCH1.AcctCode
			  FROM PCH1
			 INNER JOIN OACT ON PCH1.AcctCode=OACT.AcctCode
			 WHERE PCH1.DocEntry=@list_of_cols_val_tab_del
			   AND OACT.U_PWC_MdExpTp = 'Y'
			   AND ISNULL(PCH1.U_PWC_PjExpTp, '') = ''
			
			IF @OPCH_nvcAcctCode <> ''
			BEGIN
				SET @error = -1
				SET @error_message = N'해당 계정의 경우 현장 경비 유형은 필수입니다. [G/L 계정 : ' + @OPCH_nvcAcctCode + ']'
				RETURN
			END
		END
	END
END

/*** ORPC : 구매관리 - A/P대변메모 ***/
IF @object_type = '19'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		DECLARE   @ORPC_nvcAcctCode	NVARCHAR(15)	-- 필수 체크용 AcctCode (시공비 유형, 현장 경비 유형)

		--* A/P 대변메모 - 서비스 - 시공비 유형, 현장 경비 유형 필수 여부 체크
		IF (SELECT DocType FROM ORPC WHERE DocEntry=@list_of_cols_val_tab_del) = 'S'
		BEGIN
			-- 시공비 유형 필수 누락 계정 추출
			SET @ORPC_nvcAcctCode = ''
			SELECT TOP 1 @ORPC_nvcAcctCode = RPC1.AcctCode
			  FROM RPC1
			 INNER JOIN OACT ON RPC1.AcctCode=OACT.AcctCode
			 WHERE RPC1.DocEntry=@list_of_cols_val_tab_del
			   AND OACT.U_PWC_MdInsExpTp = 'Y'
			   AND ISNULL(RPC1.U_PWC_PjInsExpTp, '') = ''
			   			
			IF @ORPC_nvcAcctCode <> ''
			BEGIN
				SET @error = -1
				SET @error_message = N'해당 계정의 경우 시공비 유형은 필수입니다. [G/L 계정 : ' + @ORPC_nvcAcctCode + ']'
				RETURN
			END
			
			-- 현장 경비 유형 필수 누락 계정 추출
			SET @ORPC_nvcAcctCode = ''
			SELECT TOP 1 @ORPC_nvcAcctCode = RPC1.AcctCode
			  FROM RPC1
			 INNER JOIN OACT ON RPC1.AcctCode=OACT.AcctCode
			 WHERE RPC1.DocEntry=@list_of_cols_val_tab_del
			   AND OACT.U_PWC_MdExpTp = 'Y'
			   AND ISNULL(RPC1.U_PWC_PjExpTp, '') = ''
			
			IF @ORPC_nvcAcctCode <> ''
			BEGIN
				SET @error = -1
				SET @error_message = N'해당 계정의 경우 현장 경비 유형은 필수입니다. [G/L 계정 : ' + @ORPC_nvcAcctCode + ']'
				RETURN
			END
		END
	END
END

/*** OINV : 판매관리 - A/R송장 ***/
IF @object_type = '13'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* A/R송장일경우 BL번호 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OINV WHERE DocEntry=@list_of_cols_val_tab_del AND U_SOTYPE = '2' AND UPDINVNT = 'I' AND U_BL_NO IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'BL번호가 입력되지 않았습니다..'
			RETURN
		END
		--* A/R송장일경우 결제조건 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OINV WHERE DocEntry=@list_of_cols_val_tab_del AND U_SOTYPE = '2' AND UPDINVNT = 'I' AND U_SETTLE IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'결제조건이 입력되지 않았습니다..'
			RETURN
		END
	END
END

/*** ORDR : 판매관리 - 판매오더 ***/
IF @object_type = '17'
BEGIN
	--판매오더 복제시 진행수량(U_INGQTY)이 복제 되는경우를 방지하기 위함
	IF @transaction_type IN ('A')
	BEGIN
		UPDATE [RDR1] SET U_INGQTY = 0 WHERE DocEntry = @list_of_cols_val_tab_del
	END 
END

/*** OPOR : 구매관리 - 구매오더 ***/
IF @object_type = '22'
BEGIN
	--구매오더 복제시 진행수량(U_INGQTY)이 복제 되는경우를 방지하기 위함
	IF @transaction_type IN ('A')
	BEGIN
		UPDATE [POR1] SET U_INGQTY = 0 WHERE DocEntry = @list_of_cols_val_tab_del
	END 
END

/*** ODLN : 판매관리 - 납품 ***/
IF @object_type = '15'
BEGIN
	--납품 운송비가 있는경우 운송-부서는 필수로 체크하기위함.
	IF @transaction_type IN ('A','U')
	BEGIN
		IF EXISTS(SELECT DocEntry FROM ODLN WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'운송-부서가 입력되지 않았습니다..'
			RETURN
		END
		
	END 
END

/*** ORDN : 판매관리 - 반품 ***/
IF @object_type = '16'
BEGIN
	--납품 운송비가 있는경우 운송-부서는 필수로 체크하기위함.
	IF @transaction_type IN ('A','U')
	BEGIN
		IF EXISTS(SELECT DocEntry FROM ORDN WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'운송-부서가 입력되지 않았습니다..'
			RETURN
		END
	END 
END




--------------------------------------------------------------------------------------------------------------------------------
--	원천문서의 월과 불일치 데이터 체크
--------------------------------------------------------------------------------------------------------------------------------
IF @object_type = '18'			--	AP송장(OPCH0),입고PO(OPDN)
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPCH T0 
				INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN PDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN OPDN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='20' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @error = -1
		SET @ERROR_MESSAGE = N'입고PO의 전기월과 A/P송장의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		RETURN
	END
	--ELSE IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY
	--			INNER JOIN POR1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
	--			INNER JOIN OPOR T3 ON T2.DOCENTRY=T3.DOCENTRY
	--			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='22' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	--BEGIN
	--	SET @ERROR = -1
	--	SET @ERROR_MESSAGE = N'구매오더의 전기월과 A/P송장의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
	--	RETURN
	--END
END
ELSE IF @object_type = '19'			--	A/P대변메모
BEGIN
	--IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY
	--			INNER JOIN PCH1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
	--			INNER JOIN OPCH T3 ON T2.DOCENTRY=T3.DOCENTRY
	--			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='18' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	--BEGIN
	--	SET @ERROR = -1
	--	SET @ERROR_MESSAGE = N'A/P송장의 전기월과 A/P대변메모의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
	--	RETURN
	--END
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN RPD1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ORPD T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='21' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'입고반품의 전기월과 A/P대변메모의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		RETURN
	END
END
ELSE IF @object_type = '20'			--	입고po
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN RPD1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ORPD T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='21' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'입고반품의 전기월과 입고PO의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		RETURN
	END
	ELSE IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN PCH1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN OPCH T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY= @list_of_cols_val_tab_del 
				AND T1.BASETYPE = '18' 
				AND T0.DOCDATE < T3.DOCDATE
				AND T3.UPDINVNT = 'O')
	BEGIN
		SET @error = -1
		SET @ERROR_MESSAGE = N'입고PO의 전기월이 A/P예약송장의 전기월 보다 작습니다.[관리자에게 문의바람]'
		RETURN
	END
END
ELSE IF @object_type = '21'			--	입고반품
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPD T0 INNER JOIN RPD1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN PDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN OPDN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='20' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'입고PO의 전기월과 입고반품의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		RETURN
	END
END
ELSE IF @object_type = '13'			--	A/R송장
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OINV T0 INNER JOIN INV1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN DLN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ODLN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='15' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'납품의 전기월과 A/R송장의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		RETURN
	END
END
ELSE IF @object_type = '14'			--	A/P대변메모
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN INV1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN OINV T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='13' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'A/R송장의 전기월과 A/R대변메모의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		RETURN
	END
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN RDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ORDN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='16' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'반품의 전기월과 A/R대변메모의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		RETURN
	END
END
ELSE IF @object_type = '15'			--	입고po
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ODLN T0 INNER JOIN DLN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN RDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ORDN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='16' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'반품의 전기월과 납품의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		RETURN
	END
END
ELSE IF @object_type = '16'			--	입고반품
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORDN T0 INNER JOIN RDN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN DLN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ODLN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='15' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'납품의 전기월과 반품의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		RETURN
	END
END


END