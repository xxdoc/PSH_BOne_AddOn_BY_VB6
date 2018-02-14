IF OBJECT_ID('SBO_SP_FI_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_FI_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_FI_PostTransactionNotice
 설      명 : 재무관리, 자금관리 PostTransactionNotice
 작  성  자 : 
 일      시 : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_FI_PostTransactionNotice] 

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
		

/*** OJDT : 분개 ***/
IF @object_type = '30'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 차입금 마스터 전표 내역 등록
		DECLARE @tbl_OJDT_ODBT_1 TABLE
		(
			DebtCode	NVARCHAR(15)
		)
		
		-- 1. 적용 대상 차입금 관리 번호 리스트 생성
		INSERT INTO @tbl_OJDT_ODBT_1 (DebtCode)
		SELECT DebtCode
		  FROM (
			SELECT U_PWC_DebtCode AS DebtCode
			  FROM JDT1
			 WHERE TransId = @list_of_cols_val_tab_del		   
			   AND U_PWC_DebtCode <> ''
			 UNION ALL
			SELECT ODBT.Code		-- 분개 - 차입금 관리번호 입력 후 지운 경우에 대한 처리
			  FROM [@PWC_TRODBT] ODBT WITH (NOLOCK)
			 INNER JOIN [@PWC_TRDBT3] DBT3 WITH (NOLOCK) ON ODBT.Code=DBT3.Code
			 WHERE DBT3.U_TransIdx = @list_of_cols_val_tab_del
			 ) A
		 GROUP BY DebtCode
		
		IF (SELECT COUNT(*) FROM @tbl_OJDT_ODBT_1) > 0
		BEGIN
			-- 2. 차입금 마스터 관리번호 삭제
			DELETE [@PWC_TRDBT3]
			 WHERE Code IN (SELECT DebtCode FROM @tbl_OJDT_ODBT_1)
			
			-- 3. 차입금 마스터 - 관련 전표 등록
			INSERT INTO [@PWC_TRDBT3] (	
					  Code, LineId, Object, U_TransIdx, U_JdtLineId					-- 5
					, U_RefDate, U_CardCode, U_CardName, U_Account, U_AcctName		-- 10
					, U_Debit, U_Credit, U_FCDebit, U_FCCredit)
			SELECT    JDT1.U_PWC_DebtCode
					, ROW_NUMBER() OVER (PARTITION BY JDT1.U_PWC_DebtCode ORDER BY JDT1.RefDate, JDT1.TransId, JDT1.Line_ID)
					, 'PWC_UDO_TRODBT'
					, JDT1.TransId
					, JDT1.Line_ID + 1 AS LineId
					, JDT1.RefDate
					, CASE WHEN OCRD.CardName IS NULL THEN '' ELSE JDT1.ShortName END
					, ISNULL(OCRD.CardName, '')
					, JDT1.Account
					, OACT.AcctName
					, ISNULL(Debit, 0) AS Debit
					, ISNULL(Credit, 0) AS Credit
					, ISNULL(FCDebit, 0) AS Debit										
					, ISNULL(FCCredit, 0) AS Credit
			  FROM JDT1
			 INNER JOIN OACT WITH (NOLOCK) ON JDT1.Account=OACT.AcctCode
			 INNER JOIN [@PWC_TRODBT] ODBT WITH (NOLOCK) ON JDT1.U_PWC_DebtCode=ODBT.Code
			  LEFT JOIN OCRD OCRD WITH (NOLOCK) ON JDT1.ShortName=OCRD.CardCode
			 WHERE JDT1.U_PWC_DebtCode IN (SELECT DebtCode FROM @tbl_OJDT_ODBT_1)			 
		END
	END
	
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN				
		--* 품의 작성자, 승인자 명 갱신 작업
		IF (SELECT ISNULL(U_PWC_RptNumb, '') FROM OJDT WHERE TransId=@list_of_cols_val_tab_del) <> ''
		BEGIN
			EXEC PWC_SP_SetExpRptInfo @object_type, @list_of_cols_val_tab_del
		END
	END
		
	RETURN	-- 종료	
END


/*** ORCT : 자금관리 - 입금 ***/
IF @object_type = '24'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM ORCT
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END
	
	/* 취소 모드 */
	IF @transaction_type = 'C'
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=OJDT.TransId
				, @m_intBPLId=ORCT.U_PWC_BPLId
		  FROM ORCT INNER JOIN OJDT ON ORCT.TransId=OJDT.StornoToTr
		 WHERE ORCT.DocEntry = @list_of_cols_val_tab_del
		  
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END
		
	RETURN	-- 종료
END


/*** ODPS : 자금관리 - 예금 ***/
IF @object_type = '25'
BEGIN
	/* 추가, 갱신, 취소 모드 */
	IF @transaction_type IN ('A', 'U', 'C')
	BEGIN		
		--* 예금 처리 유형이 어음인 경우 RETURN (어음관리에서 처리)
		IF (SELECT TOP 1 DeposType FROM ODPS WHERE DeposId = @list_of_cols_val_tab_del) = 'B'	
		BEGIN
			RETURN
		END
		
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransAbs
				, @m_intBPLId=U_PWC_BPLId
		  FROM ODPS
		 WHERE DeposId = @list_of_cols_val_tab_del

		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
		
	RETURN	-- 종료
END


/*** OVPM : 자금관리 - 지급 ***/
IF @object_type = '46'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM OVPM
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END
	
	/* 취소 모드 */
	IF @transaction_type = 'C'
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=OJDT.TransId
				, @m_intBPLId=OVPM.U_PWC_BPLId
		  FROM OVPM INNER JOIN OJDT ON OVPM.TransId=OJDT.StornoToTr
		 WHERE OVPM.DocEntry = @list_of_cols_val_tab_del
		  
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END
		
	RETURN	-- 종료
END

/*** OBOT : 자금관리 - 어음관리 ***/
IF @object_type = '182'
BEGIN
	DECLARE   @OBOT_nvcBoeType			CHAR(1)			-- 입금(I), 지급(O) 구분
			, @OBOT_nvcBoeKey			INT				-- 어음 Key
			, @OBOT_nvcTransactionRoot	NVARCHAR(10)	-- 어음관리 경로 (FROM -> TO)
	
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		
		SELECT    TOP 1 @OBOT_nvcBoeKey = BOT1.BOEAbs
				, @m_intTransId = OBOT.TransId
				, @OBOT_nvcTransactionRoot = OBOT.StatusFrom + OBOT.StatusTo
				, @OBOT_nvcBoeType=BOT1.BoeType
		  FROM OBOT
		 INNER JOIN BOT1 ON OBOT.AbsEntry=BOT1.AbsEntry
		 WHERE OBOT.AbsEntry = @list_of_cols_val_tab_del
		
		IF @OBOT_nvcBoeType = 'I'		-- 입금 (어음 생성시에는 입금에서 사업장 처리)
		BEGIN			
			-- 입금 - 사업장 정보를 가져온다.
			SELECT @m_intBPLId=ORCT.U_PWC_BPLId
			  FROM OBOE WITH (NOLOCK)
			 INNER JOIN ORCT WITH (NOLOCK) ON OBOE.PmntNum=ORCT.DocEntry
			 WHERE OBOE.BoeKey = @OBOT_nvcBoeKey			   
			
			IF @OBOT_nvcTransactionRoot IN ('GD', 'DG')	-- 생성 -> 예금, 예금 -> 생성의 경우 예금에 사업장 갱신 처리
			BEGIN			
				 -- 예금 사업장 갱신				 				
				UPDATE ODPS
				   SET U_PWC_BPLId=@m_intBPLId
				 WHERE TransAbs = @m_intTransId
			END
		END
		ELSE IF @OBOT_nvcBoeType = 'O'	-- 지급 (어음 생성시에는 지급에서 사업장 처리)
		BEGIN
			-- 지급 - 사업장 정보를 가져온다.
			SELECT @m_intBPLId=OVPM.U_PWC_BPLId
			  FROM OBOE WITH (NOLOCK)
			 INNER JOIN OVPM WITH (NOLOCK) ON OBOE.PmntNum=OVPM.DocEntry
			 WHERE OBOE.BoeKey = @OBOT_nvcBoeKey			   
		END

		--* 분개 - 사업장 연동
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId

	END	
	
	RETURN	-- 종료
END


/*** [@PWC_TRODBT] : 차입금 마스터 등록 ***/
IF @object_type = 'PWC_UDO_TRODBT'
BEGIN
	DECLARE	  @ODBT_nvcOCRDLIsTotlColt	NVARCHAR(1)
			, @ODBT_nvcCardCode			NVARCHAR(15)
			, @ODBT_nvcIsTotlColt		NVARCHAR(1)
	
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		--* BP & 차입금 마스터 포괄 담보 여부 연동
		SELECT    @ODBT_nvcCardCode = U_CardCode
				, @ODBT_nvcIsTotlColt = ISNULL(U_IsTotlColt, 'N')
		  FROM [@PWC_TRODBT]
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		SELECT @ODBT_nvcOCRDLIsTotlColt = ISNULL(U_PWC_LIsTotlColt, 'N')
		  FROM OCRD WITH (NOLOCK)
		 WHERE CardCode = @ODBT_nvcCardCode
		
		IF @ODBT_nvcIsTotlColt <> @ODBT_nvcOCRDLIsTotlColt
		BEGIN
			UPDATE [@PWC_TRODBT]
			   SET U_IsTotlColt = @ODBT_nvcOCRDLIsTotlColt
			 WHERE DocEntry = @list_of_cols_val_tab_del
		END
	END
	
	RETURN -- 종료
END

END