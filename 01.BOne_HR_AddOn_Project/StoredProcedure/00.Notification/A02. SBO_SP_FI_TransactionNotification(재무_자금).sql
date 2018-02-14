IF OBJECT_ID('SBO_SP_FI_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_FI_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_FI_TransactionNotification
 설      명 : 재무관리, 자금관리 TransactionNotification
 작  성  자 : 
 일      시 : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_FI_TransactionNotification] 

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
/*** OJDT : 분개 ***/
IF @object_type = '30'
BEGIN

	/* 다른 종류의 출처(입금, 지급의 경우)로 생성된 경우에는 해당 IF문 위 쪽에 작업 진행 */
	IF NOT EXISTS(SELECT TransId FROM OJDT WHERE TransId=@list_of_cols_val_tab_del AND TransType='30')
	BEGIN
		RETURN	-- 분개에서 작성된 경우에만 하위 유효성 검사가 구동 되도록 설정(현재까지 발견된 입금, 지급의 경우에는 해당 구문이 구동되므로 이와 같이 처리)
	END
	
	DECLARE @OJDT_nvcAcctCode	NVARCHAR(15)

	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 차입금 관리번호 존재 여부 검사
		IF EXISTS(	SELECT JDT1.U_PWC_DebtCode
					  FROM JDT1 JDT1
		              LEFT JOIN [@PWC_TRODBT] ODBT WITH (NOLOCK) ON JDT1.U_PWC_DebtCode=ODBT.Code
		             WHERE JDT1.TransId=@list_of_cols_val_tab_del
		               AND JDT1.U_PWC_DebtCode IS NOT NULL
		               AND ODBT.Code IS NULL	)
		BEGIN
			SET @error = -1
			SET @error_message = N'존재하지 않는 차입금 관리번호를 입력하였습니다.'
			RETURN
		END
		
		--* 사업장 필수 입력 체크
		IF EXISTS(
			SELECT OJDT.TransId
			  FROM OJDT
			 INNER JOIN JDT1 ON OJDT.TransId=JDT1.TransId
			 WHERE OJDT.TransId=@list_of_cols_val_tab_del
			   AND OJDT.U_BA_TCODE IS NULL					-- 고정자산 발생 건의 경우 사업장 유효 검사 제외
			   AND JDT1.U_PWC_BpliCode IS NULL
		)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END
		
		--* BP 필수 누락 계정 추출
		SET @OJDT_nvcAcctCode = ''
		SELECT TOP 1 @OJDT_nvcAcctCode = JDT1.Account
		  FROM JDT1
		 INNER JOIN OACT WITH (NOLOCK) ON JDT1.Account=OACT.AcctCode
		 WHERE JDT1.TransId=@list_of_cols_val_tab_del
		   AND OACT.U_PWC_MdBpCd = 'Y'
		   AND ISNULL(JDT1.U_PWC_CardCode, '') = ''
		   			
		IF @OJDT_nvcAcctCode <> ''
		BEGIN
			SET @error = -1
			SET @error_message = N'해당 계정의 경우 B/P코드는 필수입니다. [G/L 계정 : ' + @OJDT_nvcAcctCode + ']'
			RETURN
		END
	END
	
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		--* 차입금 관리번호 비활성 건 입력 여부 검사
		IF EXISTS(	SELECT ODBT.Code
		              FROM JDT1 JDT1
		             INNER JOIN [@PWC_TRODBT] ODBT WITH (NOLOCK) ON JDT1.U_PWC_DebtCode=ODBT.Code 
		             WHERE JDT1.TransId=@list_of_cols_val_tab_del
		               AND ODBT.U_Active='N'	)
		BEGIN
			SET @error = -1
			SET @error_message = N'비활성 처리된 차입금 관리번호는 입력할 수 없습니다.'
			RETURN
		END
	END
	
	/* 갱신 모드 */
	IF @transaction_type = 'U'
	BEGIN
		--* 차입금 관리번호 비활성 건 수정 여부 검사
		IF EXISTS(	SELECT DBT3.Code
					  FROM [@PWC_TRODBT] ODBT WITH (NOLOCK)
					 INNER JOIN [@PWC_TRDBT3] DBT3 WITH (NOLOCK) ON ODBT.Code=DBT3.Code
					 INNER JOIN JDT1 JDT1 ON DBT3.U_TransIdx=JDT1.TransId AND DBT3.U_JdtLineId=JDT1.Line_Id
					 WHERE JDT1.TransId = @list_of_cols_val_tab_del
					   AND ODBT.U_Active = 'N'
					   AND ODBT.Code <> ISNULL(JDT1.U_PWC_DebtCode, '')	)
		BEGIN
			SET @error = -1
			SET @error_message = N'비활성 처리된 차입금 관리번호는 수정할 수 없습니다.'
			RETURN
		END	
	END
END


/*** ORCT : 자금관리 - 입금 ***/
IF @object_type = '24'
BEGIN
	/* 추가, 갱신, 취소 모드 */
	IF @transaction_type IN ('A', 'U', 'C')
	BEGIN
		--* 사업장 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM ORCT WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END
	END
END


/*** ODPS : 자금관리 - 예금 ***/
IF @object_type = '25'
BEGIN

	/* 추가, 갱신, 취소 모드 */
	IF @transaction_type IN ('A', 'U', 'C')
	BEGIN
		--* 예금 처리 유형이 어음인 경우 유효성 검사 X (어음관리에서 처리)
		IF (SELECT TOP 1 DeposType FROM ODPS WHERE DeposId = @list_of_cols_val_tab_del) = 'B'	
		BEGIN
			RETURN
		END
		
		--* 사업장 필수 입력 체크
		IF EXISTS(SELECT DeposId FROM ODPS WHERE DeposId=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END		
	END
END


/*** OVPM : 자금관리 - 지급 ***/
IF @object_type = '46'
BEGIN
	/* 추가, 갱신, 취소 모드 */
	IF @transaction_type IN ('A', 'U', 'C')
	BEGIN
		--* 사업장 필수 입력 체크
		IF EXISTS(SELECT DocEntry FROM OVPM WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 입력해 주세요.'
			RETURN
		END
	END
END


/*** OBOT : 자금관리 - 어음관리 ***/
IF @object_type = '182'
BEGIN
	DECLARE   @OBOT_nvcBoeType			CHAR(1)			-- 입금(I), 지급(O) 구분
	
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		SELECT   TOP 1 @OBOT_nvcBoeType=BOT1.BoeType
		  FROM OBOT
		 INNER JOIN BOT1 ON OBOT.AbsEntry=BOT1.AbsEntry
		 WHERE OBOT.AbsEntry = @list_of_cols_val_tab_del
		
		IF @OBOT_nvcBoeType = 'I'		-- 입금
		BEGIN
			-- 다중 사업장 처리 불가
			IF (	SELECT COUNT(*)
					  FROM (
						SELECT ORCT.U_PWC_BPLId
						  FROM OBOE WITH (NOLOCK)
						 INNER JOIN ORCT WITH (NOLOCK) ON OBOE.PmntNum=ORCT.DocEntry
						 WHERE OBOE.BoeKey IN (SELECT BOEAbs FROM BOT1 WHERE AbsEntry = @list_of_cols_val_tab_del)  
						 GROUP BY ORCT.U_PWC_BPLId
					     ) T1
				) > 1			
			BEGIN
				SET @error = -1
				SET @error_message = N'동일 사업장의 경우만 다중 처리가 가능합니다.'
				RETURN
			END
		END
		ELSE IF @OBOT_nvcBoeType = 'O'	-- 지급
		BEGIN
			-- 다중 사업장 처리 불가
			IF (	SELECT COUNT(*)
					  FROM (
						SELECT OVPM.U_PWC_BPLId
						  FROM OBOE WITH (NOLOCK)
						 INNER JOIN OVPM WITH (NOLOCK) ON OBOE.PmntNum=OVPM.DocEntry
						 WHERE OBOE.BoeKey IN (SELECT BOEAbs FROM BOT1 WHERE AbsEntry = @list_of_cols_val_tab_del)  
						 GROUP BY OVPM.U_PWC_BPLId
					     ) T1
				) > 1			
			BEGIN
				SET @error = -1
				SET @error_message = N'동일 사업장의 경우만 다중 처리가 가능합니다.'
				RETURN
			END
		END
		
	END
END


/*** [@PWC_TRODBT] : 차입금 마스터 등록 ***/
IF @object_type = 'PWC_UDO_TRODBT'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		DECLARE   @nvc_TRODBT_DebtCode	NVARCHAR(15)
				, @nvc_TRODBT_CardCode	NVARCHAR(15)
				, @dat_TRODBT_DebtDate	DATETIME
				, @dat_TRODBT_DuexDate	DATETIME
				, @int_TRODBT_BPLId		INT
				, @ins_TRODBT_DebtType	NVARCHAR(2)
				, @nvc_TRODBT_DebtAcct	NVARCHAR(15)				

		SELECT    @nvc_TRODBT_DebtCode = ISNULL(Code, '')
				, @nvc_TRODBT_CardCode = ISNULL(U_CardCode, '')
				, @dat_TRODBT_DebtDate = ISNULL(U_DebtDate, '19000101')
				, @dat_TRODBT_DuexDate = ISNULL(U_DuexDate, '19000101')
				, @int_TRODBT_BPLId = ISNULL(U_BPLId, -99)
				, @ins_TRODBT_DebtType = ISNULL(U_DebtType, '')
				, @nvc_TRODBT_DebtAcct = ISNULL(U_DebtAcct, '')				
		  FROM [@PWC_TRODBT]
		 WHERE Code=@list_of_cols_val_tab_del
		
		IF @nvc_TRODBT_DebtCode = ''
		BEGIN
			SET @error = -1
			SET @error_message = N'차입금 관리번호가 정상적으로 자동 채번되지 않았습니다. 관리자에게 문의해 주세요.'
			RETURN
		END
		
		IF @nvc_TRODBT_CardCode = ''
		BEGIN
			SET @error = -1
			SET @error_message = N'공급업체 코드를 입력해 주세요.'
			RETURN
		END
		
		IF @dat_TRODBT_DebtDate = '19000101'
		BEGIN
			SET @error = -1
			SET @error_message = N'차입일을 입력해 주세요.'
			RETURN
		END
		
		IF @dat_TRODBT_DuexDate = '19000101'
		BEGIN
			SET @error = -1
			SET @error_message = N'만기일을 입력해 주세요.'
			RETURN
		END
		
		IF @int_TRODBT_BPLId = -99
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 선택해 주세요.'
			RETURN
		END
		
		IF @int_TRODBT_BPLId = -99
		BEGIN
			SET @error = -1
			SET @error_message = N'사업장을 선택해 주세요.'
			RETURN
		END
		
		IF @ins_TRODBT_DebtType = ''
		BEGIN
			SET @error = -1
			SET @error_message = N'차입과목을 입력해 주세요.'
			RETURN
		END
						
		IF EXISTS(SELECT U_DuexDate FROM [@PWC_TRDBT1] WHERE Code=@list_of_cols_val_tab_del AND U_DuexDate IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'[상환스케쥴 탭] - 상환예정일을 입력해주세요.'
			RETURN
		END
		
		IF EXISTS(SELECT U_DuexDate FROM [@PWC_TRDBT1] WHERE Code=@list_of_cols_val_tab_del GROUP BY U_DuexDate HAVING COUNT(U_DuexDate) > 1)
		BEGIN
			SET @error = -1
			SET @error_message = N'[상환스케쥴 탭] - 중복되는 상환예정일이 존재합니다.'
			RETURN
		END
		
		IF EXISTS(SELECT U_FromDate FROM [@PWC_TRDBT2] WHERE Code=@list_of_cols_val_tab_del AND U_FromDate IS NULL OR U_ToxxDate IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'[이자율변동 탭] - 이자율 변동 기간을 입력해주세요.'
			RETURN
		END
		
		IF EXISTS(	SELECT LineId
					  FROM (
						 SELECT (SELECT S1.LineId FROM [@PWC_TRDBT2] S1 WHERE S1.Code=@list_of_cols_val_tab_del AND S1.LineId<>DBT2.LineId AND S1.U_FromDate<=DBT2.U_ToxxDate AND S1.U_ToxxDate>=DBT2.U_FromDate) AS LineId
						   FROM [@PWC_TRDBT2] DBT2
						  WHERE DBT2.Code=@list_of_cols_val_tab_del
				           ) P1
					 WHERE LineId IS NOT NULL  )
		BEGIN
			SET @error = -1
			SET @error_message = N'[이자율변동 탭] - 중복되는 이자율 변동 기간이 존재합니다.'
			RETURN
		END
	END
	
	/* 제거 모드 */
	IF @transaction_type = 'D'
	BEGIN
		IF EXISTS(SELECT TOP 1 TransId FROM JDT1 WITH (NOLOCK) WHERE U_PWC_DebtCode=@list_of_cols_val_tab_del)
		BEGIN
			SET @error = -1
			SET @error_message = N'해당 차입금 마스터 데이터를 이미 다른 전표에서 사용 중이므로 제거할 수 없습니다.'
			RETURN
		END
	END
END


END