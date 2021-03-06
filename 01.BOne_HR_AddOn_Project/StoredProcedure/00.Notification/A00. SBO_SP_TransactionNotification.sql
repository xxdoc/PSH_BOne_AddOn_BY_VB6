/****** Object:  StoredProcedure [dbo].[SBO_SP_TransactionNotification]    Script Date: 06/01/2012 11:26:24 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[SBO_SP_TransactionNotification] 

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
--INSERT INTO ObjHistory
--SELECT    (SELECT ISNULL(MAX(Seq), 0) + 1 FROM ObjHistory)
--		, @object_type
--		, @transaction_type
--		, @list_of_cols_val_tab_del
		
/*** 운영관리 ***/
-- PWC_UDO_BCOCSY : 시스템공통코드등록(재무)
IF @object_type IN ('PWC_UDO_BCOCSY')
BEGIN
	EXECUTE SBO_SP_BC_TransactionNotification	  @object_type, @transaction_type, @num_of_cols_in_key
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
-- PWC_UDO_TRODBT : 자금관리 - 차입금 마스터 등록
IF @object_type IN ('30', '24', '25', '46', '182', 'PWC_UDO_TRODBT')
BEGIN
	EXECUTE SBO_SP_FI_TransactionNotification	  @object_type, @transaction_type, @num_of_cols_in_key
												, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
												, @error output, @error_message output
	GOTO Point_Exit
END

/*** 판매관리, 구매관리 ***/
-- 69	: 구매관리 - 수입부대비용
-- 18   : 구매관리 - AP예약송장
-- 19   : 구매관리 - AP대변메모
-- 13   : 판매관리 - AR송장
-- 17   : 판매관리 - 판매오더
-- 22   : 구매관리 - 구매오더
-- 15   : 판매관리 - 납품
-- 16   : 판매관리 - 반품
--IF @object_type IN ('69','18','19','13','17','22')		-- 월별불일치 체크 미포함
IF @object_type IN ('69','18','19','20','21','13','14','15','16','17','22')		-- 월별불일치 체크 오브젝트 포함
BEGIN
	EXECUTE SBO_SP_MM_TransactionNotification	  @object_type, @transaction_type, @num_of_cols_in_key
												, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
												, @error output, @error_message output
	GOTO Point_Exit
END

/*** 재고관리, 생산관리 ***/
-- 59	: 재고관리 - 생산입고
-- 60	: 재고관리 - 생산출고
-- 1250000001 : 재고관리 - 재고 이전 요청
-- 67	: 재고관리 - 재고 이전
-- 162	: 재고관리 - 재고재평가
-- MDC_MM_CSM001 : 색상등록 4자리 입력 체크
-- MDC_MM_CPD103 : 생산관리 - 작업불량 및 로스 등록
-- MDC_MM_CPD105 : 생산관리 - 작업 잔량 입고 등록
-- MDC_MM_CPD211 : 생산관리 - PMS 생산 의뢰 접수_산업
-- MDC_MM_CPD104 : 생산일일공수 공정구분 중복 체크
-- MDC_MM_CPD106 : 생산관리 - 외주 생산 지시 등록[PVC, 알미늄]
-- 4	: 품목마스터 - 표준품목(품목원가 체크)
IF @object_type IN ('59', '60', '1250000001', '67', '162', '202', '4','MDC_MM_CSM001', 'MDC_MM_CPD103', 'MDC_MM_CPD105', 'MDC_MM_CPD211', 'MDC_MM_CPD104', 'MDC_MM_CPD106')
BEGIN
	EXECUTE SBO_SP_IV_TransactionNotification	  @object_type, @transaction_type, @num_of_cols_in_key
												, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
												, @error output, @error_message output
	GOTO Point_Exit
END

/*** 비즈니스파트너관리 ***/
-- 2	: 비즈니스 파트너 마스터 데이터
-- PWC_UDO_BPOCDM : 비즈니스 파트너 코드 관리 (채번규칙)
IF @object_type IN ('2', 'PWC_UDO_BPOCDM')
BEGIN
	EXECUTE SBO_SP_BP_TransactionNotification	  @object_type, @transaction_type, @num_of_cols_in_key
												, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
												, @error output, @error_message output
	GOTO Point_Exit
END

/*** 관리 회계 ***/
-- PWC_UDO_COOCEG : 원가 요소 그룹 등록
-- PWC_UDO_COOCCG : 코스트/손익 센터 그룹 등록
-- PWC_UDO_COOPRG : 프로젝트 그룹 등록
-- PWC_UDO_COOOCR : 코스트/손익 센터 배부율 등록
-- PWC_UDO_COOPCR : 프로젝트 배부율 등록
-- PWC_UDO_COOCRA : 원가 조정 계정
IF @object_type IN ('PWC_UDO_COOCEG', 'PWC_UDO_COOCCG', 'PWC_UDO_COOPRG', 'PWC_UDO_COOOCR', 'PWC_UDO_COOPCR', 'PWC_UDO_COOCRA')
BEGIN
	EXECUTE SBO_SP_CO_TransactionNotification	  @object_type, @transaction_type, @num_of_cols_in_key
												, @list_of_key_cols_tab_del, @list_of_cols_val_tab_del
												, @error output, @error_message output
	GOTO Point_Exit
END

Point_Exit:
 --Select the return values
select @error, '[NT] ' + @error_message

end