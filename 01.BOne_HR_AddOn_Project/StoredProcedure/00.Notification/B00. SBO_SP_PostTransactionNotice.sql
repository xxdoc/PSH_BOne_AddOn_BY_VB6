/****** Object:  StoredProcedure [dbo].[SBO_SP_PostTransactionNotice]    Script Date: 06/01/2012 13:09:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[SBO_SP_PostTransactionNotice]

@object_type nvarchar(20), 				-- SBO Object Type
@transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
@num_of_cols_in_key int,
@list_of_key_cols_tab_del nvarchar(255),
@list_of_cols_val_tab_del nvarchar(255)

AS

begin

-- Return values
declare @error  int				-- Result (0 for no error)
declare @error_message nvarchar (200) 		-- Error string to be displayed
select @error = 0
select @error_message = N'Ok'


--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------
/*** 운영관리 ***/
-- 247	: 설정 - 재무관리 - 사업장
-- 10	: 설정 - 비즈니스파트너관리 - 고객, 공급업체 그룹
-- 48	: 설정 - 구매 - 수입부대비용
-- 147	: 설정 - 자금관리 - 지급 방법
IF @object_type IN ('247', '10', '48', '147')
BEGIN
	EXECUTE SBO_SP_BC_PostTransactionNotice	  @object_type, @transaction_type, @num_of_cols_in_key
											, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
											, @error output, @error_message output	
	GOTO Point_Exit	
END

/*** 재무관리, 자금관리 ***/
-- 30	: 재무관리 - 분개
-- 24	: 자금관리 - 입금
-- 25	: 자금관리 - 예금
-- 46	: 자금관리 - 지급
-- 182	: 자금관리 - 어음관리
-- PWC_UDO_TRODBT : 자금관리 - 차입금 마스터
IF @object_type IN ('30', '24', '25', '46', '182', 'PWC_UDO_TRODBT')
BEGIN
	EXECUTE SBO_SP_FI_PostTransactionNotice	  @object_type, @transaction_type, @num_of_cols_in_key
											, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
											, @error output, @error_message output
	GOTO Point_Exit
END

/*** 판매관리, 구매관리 ***/
-- 15	: 판매관리 - 납품
-- 16	: 판매관리 - 반품
-- 203	: 판매관리 - A/R선금요청
-- 13	: 판매관리 - A/R송장
-- 14	: 판매관리 - A/R대변메모
-- 20	: 구매관리 - 입고PO
-- 21	: 구매관리 - 반품
-- 204	: 구매관리 - A/P선금요청
-- 18	: 구매관리 - A/P송장
-- 19	: 구매관리 - A/P대변메모
-- 69	: 구매관리 - 수입부대비용
IF @object_type IN ('15', '16', '203', '13', '14', '20', '21', '204', '18', '19', '69')
BEGIN
	EXECUTE SBO_SP_MM_PostTransactionNotice	  @object_type, @transaction_type, @num_of_cols_in_key
											, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
											, @error output, @error_message output
	GOTO Point_Exit
END

/*** 재고관리, 생산관리 ***/
-- 59	: 재고관리 - 입고
-- 60	: 재고관리 - 출고
-- 1250000001	: 재고관리 - 재고이전 요청
-- 67	: 재고관리 - 재고이전
-- 162	: 재고관리 - 재고재평가
-- 202	: 생산관리 - 생산오더
-- MDC_MM_CPD103: 생산관리 - 작업불량/Loss 등록
-- MDC_MM_CPD105: 생산관리 - 작업 잔량 입고 등록
-- MDC_MM_CPD106: 생산관리 - 외주 생산 지시 등록[PVC, 알미늄]
IF @object_type IN ('59', '60', '67', '162', '202', '1250000001', 'MDC_MM_CPD103', 'MDC_MM_CPD105', 'MDC_MM_CPD106')
BEGIN
	EXECUTE SBO_SP_IV_PostTransactionNotice	  @object_type, @transaction_type, @num_of_cols_in_key
											, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
											, @error output, @error_message output
	GOTO Point_Exit
END

/*** 비즈니스파트너관리 ***/
IF @object_type IN ('-999')
BEGIN
	EXECUTE SBO_SP_BP_PostTransactionNotice	  @object_type, @transaction_type, @num_of_cols_in_key
											, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
											, @error output, @error_message output
	GOTO Point_Exit
END

/*** 관리 회계 ***/
IF @object_type IN ('PWC_UDO_COOCCG', 'PWC_UDO_COOPRG')
BEGIN
	EXECUTE SBO_SP_CO_PostTransactionNotice	  @object_type, @transaction_type, @num_of_cols_in_key
											, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
											, @error output, @error_message output
	GOTO Point_Exit
END

Point_Exit:
-- Select the return values
select @error, @error_message

end