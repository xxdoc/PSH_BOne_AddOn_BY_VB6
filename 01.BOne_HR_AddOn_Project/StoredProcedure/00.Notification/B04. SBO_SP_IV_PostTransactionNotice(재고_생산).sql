IF OBJECT_ID('SBO_SP_IV_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_IV_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_IV_PostTransactionNotice
 설      명 : 재고관리, 생산관리 PostTransactionNotice
 작  성  자 : 
 일      시 : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_IV_PostTransactionNotice] 

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


/*** OIGN : 재고관리 - 입고 ***/
IF @object_type = '59'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM OIGN
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** OIGE : 재고관리 - 출고 ***/
IF @object_type = '60'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM OIGE
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** OWTQ : 재고관리 - 재고 이전 요청 ***/
IF @object_type = '1250000001'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 입고 사업장이 NULL인 경우 출고 사업장을 입고 사업장으로 갱신
		UPDATE OWTR
		   SET U_PWC_BPLId2 = U_PWC_BPLId
		 WHERE DocEntry = @list_of_cols_val_tab_del
		   AND U_PWC_BPLId2 IS NULL
	END
	
	RETURN	-- 종료
END


/*** OWTR : 재고관리 - 재고이전 ***/
IF @object_type = '67'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		DECLARE @intBPLId2	INT	-- 입고 사업장
				
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
				, @intBPLId2=U_PWC_BPLId2
		  FROM OWTR
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
		
		-- 입고 사업장이 NULL인 경우 출고 사업장을 입고 사업장으로 갱신
		IF @intBPLId2 IS NULL
		BEGIN
			SET @intBPLId2 = @m_intBPLId
			
			UPDATE OWTR
			   SET U_PWC_BPLId2 = @intBPLId2
			 WHERE DocEntry = @list_of_cols_val_tab_del
		END
		
		-- 출고 사업장과 입고 사업장이 다른 경우
		IF @m_intBPLId <> @intBPLId2 
		BEGIN
			UPDATE JDT1
			   SET U_PWC_BpliCode = @intBPLId2
			 WHERE TransId = @m_intTransId
			   AND ISNULL(Debit, 0) <> 0
		END
	END
	
	RETURN	-- 종료
END


/*** OMRV : 재고관리 - 재고재평가 ***/
IF @object_type = '162'
BEGIN
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* 분개 - 사업장 연동
		SELECT    @m_intTransId=TransId
				, @m_intBPLId=U_PWC_BPLId
		  FROM OMRV
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		EXEC PWC_SP_SetBPLIdToJournalEntry @m_intTransId, @m_intBPLId
	END
	
	RETURN	-- 종료
END


/*** OWOR : 생산관리 - 생산오더 ***/
IF @object_type = '202'
BEGIN
	/* 갱신 모드 */
	IF @transaction_type = 'U'
	BEGIN
		DECLARE   @OWOR_chrCurrIsCmpltW	CHAR(1)
				, @OWOR_chrWillIsCmpltW	CHAR(1)
		
		--* 생산 완료 여부 처리
		SELECT    @OWOR_chrCurrIsCmpltW = ISNULL(U_IsCmpltW, 'N')
				, @OWOR_chrWillIsCmpltW = CASE WHEN ISNULL(CmpltQty, 0)+ISNULL(RjctQty, 0) >= PlannedQty THEN 'Y'
										       ELSE 'N'
								           END
		  FROM OWOR
		 WHERE DocEntry = @list_of_cols_val_tab_del
		
		IF @OWOR_chrCurrIsCmpltW <> @OWOR_chrWillIsCmpltW
		BEGIN
			UPDATE OWOR
			   SET U_IsCmpltW = @OWOR_chrWillIsCmpltW
			 WHERE DocEntry = @list_of_cols_val_tab_del
		END
	END
	
	RETURN	-- 종료
END


/*** MDC_MM_CPD103 : 생산관리 - 작업불량/Loss 등록 ***/
IF @object_type = 'MDC_MM_CPD103'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN		
		--* 출고 품목원가 -> 생산불량 및 로스 등록에 갱신 (취소를 위함)
		UPDATE CPD103L 
		   SET    U_Price=OINM.CalcPrice
				, U_LineTotl=ABS(OINM.TransValue)
		  FROM [@MDC_MM_CPD103H] CPD103H
		 INNER JOIN [@MDC_MM_CPD103L] CPD103L ON CPD103H.DocEntry=CPD103L.DocEntry
		 INNER JOIN OINM WITH (NOLOCK) ON OINM.TransType='60' AND CPD103H.U_OutEntry=OINM.CreatedBy AND (CPD103L.LineId-1)=OINM.DocLineNum
		 WHERE CPD103H.DocEntry = @list_of_cols_val_tab_del
		   AND ISNULL(CPD103L.U_ItemCode, '') <> ''		 
	END
	
	RETURN	-- 종료
END


/*** MDC_MM_CPD105 : 생산관리 - 작업 잔량 입고 등록 ***/
IF @object_type = 'MDC_MM_CPD105'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN		
		--* 입고 발생(현재) 품목원가 -> 작업 잔량 입고 (입고 당시 품목원가를 표현하기 위함)
		UPDATE CPD105L
		   SET    U_Price=OINM.CalcPrice
				, U_LineTotl=OINM.TransValue
		  FROM [@MDC_MM_CPD105H] CPD105H
		 INNER JOIN [@MDC_MM_CPD105L] CPD105L ON CPD105H.DocEntry=CPD105L.DocEntry
		 INNER JOIN OINM WITH (NOLOCK) ON OINM.TransType='59' AND CPD105H.U_InEntry=OINM.CreatedBy AND (CPD105L.LineId-1)=OINM.DocLineNum
		 WHERE CPD105H.DocEntry = @list_of_cols_val_tab_del
		   AND ISNULL(CPD105L.U_ItemCode, '') <> ''
	END
	
	RETURN	-- 종료
END


/*** MDC_MM_CPD106 : 생산관리 - 외주 생산 지시 등록[PVC, 알미늄] ***/
IF @object_type = 'MDC_MM_CPD106'
BEGIN	
	/* 추가, 갱신 모드 */
	IF @transaction_type IN ('A', 'U')
	BEGIN		
		--* 라인별 생산 오더 정상 처리 여부에 따른 생산 완료 여부 값 설정
		IF NOT EXISTS ( 
			SELECT OCPD.DocEntry
			  FROM [@MDC_MM_CPD106H] AS OCPD
			 INNER JOIN [@MDC_MM_CPD106L] AS CPD1 ON OCPD.DocEntry=CPD1.DocEntry
			 WHERE OCPD.DocEntry = @list_of_cols_val_tab_del 
			   AND ISNULL(OCPD.U_IsCmplt1, 'N') = 'N'
			   AND CPD1.U_WREntry IS NULL
		)
		BEGIN
			UPDATE [@MDC_MM_CPD106H]
			   SET U_IsCmplt1 = 'Y'
			 WHERE DocEntry = @list_of_cols_val_tab_del
		END
	END
	
	/* 취소 모드 */
	IF @transaction_type = 'C'
	BEGIN
		--* 자재 산출 정보 작업 지시 여부 갱신(추가시에는 소스에서 처리)
		UPDATE [@MDC_MM_CPD206H]
		   SET    U_OWORYN = 'N'
				, U_OWORDate = NULL
		 WHERE DocEntry IN (SELECT U_MCEntry FROM [@MDC_MM_CPD106L] WHERE DocEntry = @list_of_cols_val_tab_del AND ISNULL(U_WREntry, -1) <> -1)
	END
	
	RETURN	-- 종료
END


END