IF OBJECT_ID('SBO_SP_IV_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_IV_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_IV_TransactionNotification
 설      명 : 재고관리 TransactionNotification
 작  성  자 : 
 일      시 : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_IV_TransactionNotification] 

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
DECLARE @VALUE01 NVARCHAR(MAX)

/*** OITM : 재고관리 - 품목마스터 ***/
IF @object_type = '4'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 품목원가 필수 체크
		IF EXISTS(SELECT ItemCode FROM OITM 
					WHERE ItemCode= @list_of_cols_val_tab_del
					AND EvalSystem = 'S'
					AND ISNULL(AvgPrice,0) = 0)
		BEGIN
			SET @error = -1
			SET @error_message = N'제품/반제품 품목에 품목원가 0입니다.'
			RETURN
		END
	END
END


/*** OIGN : 재고관리 - 입고 ***/
IF @object_type = '59'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 사업장 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OIGN WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END
		
		IF EXISTS(SELECT DocEntry FROM OIGN WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'운송-부서가 입력되지 않았습니다..'
			RETURN
		END
		--* 단가 0원 체크
		IF EXISTS(SELECT DocEntry FROM IGN1 WHERE DocEntry = @list_of_cols_val_tab_del AND ISNULL(Price,0) = 0 )
		BEGIN
			SET @error = -1
			SET @error_message = N'단가가 0원은 입력할 수 없습니다. '
			RETURN
		END
	END
	
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		--* 생산 입고 - 생산 오더보다 수량이 많은 경우 오류 메시지 출력
		IF EXISTS(
			SELECT DocEntry
			  FROM OWOR
			 WHERE DocEntry IN (SELECT BaseEntry FROM IGN1 WHERE DocEntry=@list_of_cols_val_tab_del AND BaseType='202')
			   AND PlannedQty < ISNULL(CmpltQty, 0)+ISNULL(RjctQty, 0)
		)
		BEGIN
			SET @error = -1
			SET @error_message = N'생산 입고 수량은 생산 오더의 계획 수량보다 클 수 없습니다.'
			RETURN
		END
		
		--* 수동 구성품 생산 출고가 일어나지 않은 경우 생산 입고가 되지 않도록 처리
		--IF EXISTS(
		--	SELECT OWOR.DocEntry
		--	  FROM OWOR
		--	 INNER JOIN WOR1 ON OWOR.DocEntry=WOR1.DocEntry
		--	 WHERE OWOR.DocEntry IN (SELECT BaseEntry FROM IGN1 WHERE DocEntry=@list_of_cols_val_tab_del AND BaseType='202')
		--	   AND WOR1.IssueType = 'M'
		--	   AND (ISNULL(OWOR.CmpltQty, 0) + ISNULL(OWOR.RjctQty, 0)) * WOR1.BaseQty > WOR1.IssuedQty
		--)
		--BEGIN
		--	SET @error = -1
		--	SET @error_message = N'수동 구성품의 경우 생산 출고 후 생산 입고가 이루어져야 됩니다.'
		--	RETURN
		--END
	END
END


/*** OIGE : 재고관리 - 출고 ***/
IF @object_type = '60'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 사업장 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OIGE WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END
		
		IF EXISTS(SELECT DocEntry FROM OIGE WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'운송-부서가 입력되지 않았습니다..'
			RETURN
		END
	END
END


/*** OWTQ : 재고관리 - 재고 이전 요청 ***/
IF @object_type = '1250000001'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 사업장 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OWTQ WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END
		
		IF EXISTS(SELECT DocEntry FROM OWTQ WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'운송-부서가 입력되지 않았습니다..'
			RETURN
		END
	END
END


/*** OWTR : 재고관리 - 재고 이전 ***/
IF @object_type = '67'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 사업장 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OWTR WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END		
		
		IF EXISTS(SELECT DocEntry FROM OWTR WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'운송-부서가 입력되지 않았습니다..'
			RETURN
		END
	END	
END


/*** OMRV : 재고관리 - 재고재평가 ***/
IF @object_type = '162'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 사업장 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OMRV WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END
	END
END


/*** OWOR : 생산관리 - 생산 오더 ***/
IF @object_type = '202'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A','U')
	BEGIN
		--* 생산 품목 CostCenter 필수 체크
		IF EXISTS(SELECT DocEntry FROM OWOR WHERE DocEntry=@list_of_cols_val_tab_del AND ISNULL(OcrCode, '')='')
		BEGIN
			SET @error = -1
			SET @error_message = N'생산 품목 배부 규칙은 필수입니다.'
			RETURN
		END
		
		--* 구성 품목 CostCenter 필수 체크
		IF EXISTS(SELECT DocEntry FROM WOR1 WHERE DocEntry=@list_of_cols_val_tab_del AND ISNULL(OcrCode, '')='')
		BEGIN
			SET @error = -1
			SET @error_message = N'구성 품목 Cost Center는 필수입니다.'
			RETURN
		END

		--* 생산 모품목 품목원가 필수 체크
		IF EXISTS(SELECT T0.DocEntry FROM OWOR T0 INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode
					WHERE T0.DocEntry= @list_of_cols_val_tab_del
					AND T1.EvalSystem = 'S'
					AND ISNULL(T1.AvgPrice,0) = 0)
		BEGIN
			SET @error = -1
			SET @error_message = N'생산 모품목에 품목원가 0입니다.'
			RETURN
		END
		--* 생산 자품목 품목원가 필수 체크
		ELSE IF EXISTS(SELECT T0.DocEntry FROM WOR1 T0 INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode
					WHERE T0.DocEntry=@list_of_cols_val_tab_del
					AND T1.EvalSystem = 'S'
					AND ISNULL(T1.AvgPrice,0) = 0)
		BEGIN
			SET @error = -1
			SET @error_message = N'생산 자품목에 품목원가 0입니다.'
			RETURN
		END
	END
END

IF @object_type = 'MDC_MM_CSM001'
BEGIN
	IF @transaction_type IN ('A', 'U')
	BEGIN
		SELECT @VALUE01 = LEN(@list_of_cols_val_tab_del)
		IF @VALUE01 > 4
		BEGIN
			SET @error = -1										-- Error Code 설정
			SET @error_message = N'코드는 4자리로 입력하세요.'	-- Error Message 설정
			RETURN												-- Error 발생시 RETURN을 사용하여 SP EXIT
		END
	END
END


/*** MDC_MM_CPD103 : 생산관리 - 작업불량 및 로스 등록 ***/
IF @object_type = 'MDC_MM_CPD103'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		IF EXISTS(SELECT DocEntry FROM [@MDC_MM_CPD103H] WHERE DocEntry = @list_of_cols_val_tab_del AND U_OutEntry IS NULL)
		BEGIN
			SET @error = -1										
			SET @error_message = N'등록에 의한 출고 처리가 정상적으로 이루어지지 않았습니다.'
			RETURN
		END
		
		IF (SELECT COUNT(DocEntry) FROM [@MDC_MM_CPD103L] WHERE DocEntry = @list_of_cols_val_tab_del AND ISNULL(U_ItemCode, '') <> '') = 0 
		BEGIN
			SET @error = -1										
			SET @error_message = N'작업 불량 및 로스 품목 정보가 존재하지 않습니다.'
			RETURN
		END
	END
	
	/* 취소 모드 */
	IF @transaction_type = 'C'
	BEGIN
		IF EXISTS(SELECT DocEntry FROM [@MDC_MM_CPD103H] WHERE DocEntry = @list_of_cols_val_tab_del AND U_InEntry IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'취소에 의한 입고 처리가 정상적으로 이루어지지 않았습니다.'
			RETURN
		END
	END	
END


/*** MDC_MM_CPD105 : 생산관리 - 작업 잔량 입고 등록 ***/
IF @object_type = 'MDC_MM_CPD105'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		IF EXISTS(SELECT DocEntry FROM [@MDC_MM_CPD105H] WHERE DocEntry = @list_of_cols_val_tab_del AND U_InEntry IS NULL)
		BEGIN
			SET @error = -1										
			SET @error_message = N'등록에 의한 입고 처리가 정상적으로 이루어지지 않았습니다.'
			RETURN
		END
		
		IF (SELECT COUNT(DocEntry) FROM [@MDC_MM_CPD105L] WHERE DocEntry = @list_of_cols_val_tab_del AND ISNULL(U_ItemCode, '') <> '') = 0 
		BEGIN
			SET @error = -1										
			SET @error_message = N'작업 잔량 입고 품목 정보가 존재하지 않습니다.'
			RETURN
		END
	END
	
	/* 취소 모드 */
	IF @transaction_type = 'C'
	BEGIN
		IF EXISTS(SELECT DocEntry FROM [@MDC_MM_CPD105H] WHERE DocEntry = @list_of_cols_val_tab_del AND U_OutEntry IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'취소에 의한 출고 처리가 정상적으로 이루어지지 않았습니다.'
			RETURN
		END
	END
END


/*** MDC_MM_CPD211 : 생산관리 - PMS 생산 의뢰 접수_산업 ***/
IF @object_type = 'MDC_MM_CPD211'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		--* PMS Data 변경 여부 체크		 
		IF EXISTS(
			SELECT SAPOPRR.U_PrjSeqn
			  FROM (
				SELECT    PRR1.U_PrjSeqn
						, PRR1.U_ItemCode			
						, CONVERT(NVARCHAR(8), U_DueDate, 112) AS DueDate
						, PRR1.U_ReceQty
				  FROM [@MDC_MM_CPD211H] OPRR
				 INNER JOIN [@MDC_MM_CPD211L] PRR1 ON OPRR.DocEntry=PRR1.DocEntry
				 WHERE OPRR.DocEntry = @list_of_cols_val_tab_del
			 ) SAPOPRR
			 INNER JOIN (
				SELECT    PMSRQI.PJT_SEQ
						, PMSPJI.ITEM_CD
						, PMSRQI.TARGET_DT
						, PMSRQI.REQ_QTY
				  FROM [EAGON_PMS].[dbo].[TPMS_PJT_ITEM_I] PMSPJI 
				 INNER JOIN [EAGON_PMS].[dbo].[TPMS_PREQ_ITEM] PMSRQI ON PMSPJI.PJT_CD=PMSRQI.PJT_CD AND PMSPJI.PJT_SEQ=PMSRQI.PJT_SEQ
				 WHERE PMSPJI.PJT_CD = (SELECT U_PrjCode COLLATE Korean_Wansung_CI_AS FROM [@MDC_MM_CPD211H] WHERE DocEntry = @list_of_cols_val_tab_del)
				   AND PMSRQI.PREQ_NUM = (SELECT U_PreqNum COLLATE Korean_Wansung_CI_AS FROM [@MDC_MM_CPD211H] WHERE DocEntry = @list_of_cols_val_tab_del)
			 ) PMSREQI ON SAPOPRR.U_PrjSeqn=PMSREQI.PJT_SEQ -- 하나의 생산의뢰 번호에 대하여 PJT_SEQ는 중복 될 수 없으므로 PJT_SEQ만 조인
			 WHERE SAPOPRR.U_ItemCode <> PMSREQI.ITEM_CD COLLATE Korean_Wansung_Unicode_CI_AS
				OR SAPOPRR.DueDate <> PMSREQI.TARGET_DT COLLATE Korean_Wansung_Unicode_CI_AS
				OR SAPOPRR.U_ReceQty <> PMSREQI.REQ_QTY
		)
		BEGIN
			SET @error = -1
			SET @error_message = N'PMS에서 다른 사용자에 의하여 생산 의뢰 정보가 변경되었습니다.'
			RETURN
		END
	END
END

IF @object_type = 'MDC_MM_CPD104'
BEGIN
	IF @transaction_type IN ('A', 'U')
	BEGIN
		IF EXISTS(SELECT U_OcrCode FROM [@MDC_MM_CPD104L] 
				WHERE DocEntry = @list_of_cols_val_tab_del
				GROUP BY U_OcrCode HAVING COUNT(U_OcrCode) > 1)
		BEGIN
			SET @error = -1										-- Error Code 설정
			SET @error_message = N'중복된 작업구분이 있습니다.'	-- Error Message 설정
			RETURN												-- Error 발생시 RETURN을 사용하여 SP EXIT
		END
	END
END


/*** MDC_MM_CPD106 : 생산관리 - 외주 생산 지시 등록[PVC, 알미늄] ***/
IF @object_type = 'MDC_MM_CPD106'
BEGIN
	/* 취소 모드 */
	IF @transaction_type = 'C'
	BEGIN
		IF EXISTS(	SELECT DocEntry
					  FROM OWOR WITH (NOLOCK)
					 WHERE DocEntry IN (SELECT U_WREntry FROM [@MDC_MM_CPD106L] WHERE DocEntry = @list_of_cols_val_tab_del)
					   AND [Status] <> 'C'
		)
		BEGIN
			SET @error = -1
			SET @error_message = N'해당 문서의 생산 오더가 모두 취소 처리 되지 않은 경우 취소를 할 수 없습니다.'
			RETURN
		END
	END
END


END