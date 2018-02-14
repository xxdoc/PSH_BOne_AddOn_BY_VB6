IF OBJECT_ID('SBO_SP_BC_PostTransactionNotice') IS NOT NULL
   DROP PROCEDURE SBO_SP_BC_PostTransactionNotice
GO
/********************************************************************************************************************************************************                                     
 프로시져명 : SBO_SP_BC_PostTransactionNotice
 설      명 : 운영관리 PostTransactionNotice
 작  성  자 : 
 일      시 : 
**********************************************************************************************************************************************************/ 
CREATE proc [dbo].[SBO_SP_BC_PostTransactionNotice] 

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
DECLARE	@m_intDocEntry	INT

/*** OBPL : 사업장 ***/
IF @object_type = '247'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		--* 연결테이블용 사업장 복제
		SELECT @m_intDocEntry = AutoKey
		  FROM ONNM WITH (NOLOCK)
		 WHERE ObjectCode = 'PWC_UDO_BCOBPL'
		
		INSERT INTO [@PWC_BCOBPL](Code, Name, DocEntry, [Object], LogInst
								, UserSign, CreateDate, CreateTime, DataSource)
		SELECT    BPLId
				, BPLName
				, @m_intDocEntry
				, 'PWC_UDO_BCOBPL'
				, 0
				, UserSign2
				, CONVERT(NVARCHAR(8), GETDATE(), 112)
				, REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
				, 'I'
		  FROM OBPL
		 WHERE BPLId = @list_of_cols_val_tab_del
		
		UPDATE ONNM
		   SET AutoKey=@m_intDocEntry+1
		 WHERE ObjectCode = 'PWC_UDO_BCOBPL'
		
		RETURN		-- 작업 종료시 RETURN을 사용하여 SP EXIT (하나의 오브젝트에 대하여 종료 시점에 작성할 것)
	END
	
	/* 갱신 모드 */
	IF @transaction_type = 'U'
	BEGIN
	
		--* 연결테이블용 사업장 갱신
		UPDATE BCOBPL SET Name=BPLName, UpdateDate=CONVERT(NVARCHAR(8), GETDATE(), 112), UpdateTime=REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
		  FROM OBPL
		 INNER JOIN [@PWC_BCOBPL] BCOBPL ON OBPL.BPLId=BCOBPL.Code
		 WHERE OBPL.BPLId = @list_of_cols_val_tab_del
		
		RETURN
	END
	
	/* 제거 모드 */
	IF @transaction_type = 'D'
	BEGIN
	
		--* 연결테이블용 사업장 제거
		DELETE [@PWC_BCOBPL]
		 WHERE Code = @list_of_cols_val_tab_del
		
		RETURN
	END
END

/*** OCRG : 비즈니스파트너 그룹 ***/
IF @object_type = '10'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		--* BP 자동 채번 규칙 테이블에 BP 그룹 추가
		SELECT @m_intDocEntry = AutoKey
		  FROM ONNM WITH (NOLOCK)
		 WHERE ObjectCode = 'PWC_UDO_BPOCDM'

		INSERT INTO [@PWC_BPOCDM] (Code, Name, DocEntry, [Object], UserSign
							, CreateDate, CreateTime, DataSource)
		SELECT    GroupType + CONVERT(NVARCHAR(6), GroupCode)
				, GroupName
				, @m_intDocEntry
				, 'PWC_UDO_BPOCDM'
				, UserSign
				, CONVERT(NVARCHAR(8), GETDATE(), 112)
				, REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
				, 'I'
		  FROM OCRG
		 WHERE GroupCode = @list_of_cols_val_tab_del
		 
		UPDATE ONNM
		   SET AutoKey=@m_intDocEntry+1
		 WHERE ObjectCode = 'PWC_UDO_BPOCDM'
		
		RETURN
	END
	
	/* 갱신 모드 */
	IF @transaction_type = 'U'
	BEGIN
		--* BP 자동 채번 규칙 테이블에 BP 그룹 갱신
		UPDATE [@PWC_BPOCDM]
		   SET Name=(SELECT TOP 1 GroupName FROM OCRG WHERE GroupCode = @list_of_cols_val_tab_del)
		 WHERE RIGHT(Code, LEN(Code) - 1) = @list_of_cols_val_tab_del
		
		RETURN
	END
	
	/* 제거 모드 */
	IF @transaction_type = 'D'
	BEGIN
		--* BP 자동 채번 규칙 테이블에 BP 그룹 제거
		DELETE [@PWC_BPOCDM]
		 WHERE RIGHT(Code, LEN(Code) - 1) = @list_of_cols_val_tab_del
		
		RETURN
	END
END

/*** OALC : 수입부대비용 유형 ***/
IF @object_type = '48'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		--* 연결테이블용 수입부대비용 유형 복제
		SELECT @m_intDocEntry = AutoKey
		  FROM ONNM WITH (NOLOCK)
		 WHERE ObjectCode = 'PWC_UDO_BCOALC'
		
		INSERT INTO [@PWC_BCOALC](Code, Name, DocEntry, [Object], LogInst
								, UserSign, CreateDate, CreateTime, DataSource)
		SELECT    AlcCode
				, AlcName
				, @m_intDocEntry
				, 'PWC_UDO_BCOALC'
				, 0
				, UserSign
				, CONVERT(NVARCHAR(8), GETDATE(), 112)
				, REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
				, DataSource
		  FROM OALC
		 WHERE AlcCode = @list_of_cols_val_tab_del
		
		UPDATE ONNM
		   SET AutoKey=@m_intDocEntry+1
		 WHERE ObjectCode = 'PWC_UDO_BCOALC'
		
		RETURN		-- 작업 종료시 RETURN을 사용하여 SP EXIT (하나의 오브젝트에 대하여 종료 시점에 작성할 것)
	END
	
	/* 갱신 모드 */
	IF @transaction_type = 'U'
	BEGIN
	
		--* 연결테이블용 수입부대비용 유형 갱신
		UPDATE BCOALC SET Name=AlcName, UpdateDate=CONVERT(NVARCHAR(8), GETDATE(), 112), UpdateTime=REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
		  FROM OALC
		 INNER JOIN [@PWC_BCOALC] BCOALC ON OALC.AlcCode=BCOALC.Code
		 WHERE OALC.AlcCode = @list_of_cols_val_tab_del
		
		RETURN
	END
	
	/* 제거 모드 */
	IF @transaction_type = 'D'
	BEGIN
	
		--* 연결테이블용 수입부대비용 유형 제거
		DELETE [@PWC_BCOALC]
		 WHERE Code = @list_of_cols_val_tab_del
		
		RETURN
	END
END

/*** OPYM : 지급 방법 ***/
IF @object_type = '147'
BEGIN
	/* 추가 모드 */
	IF @transaction_type = 'A'
	BEGIN
		--* 연결테이블용 지급 방법 복제
		SELECT @m_intDocEntry = AutoKey
		  FROM ONNM WITH (NOLOCK)
		 WHERE ObjectCode = 'PWC_UDO_BCOPYM'
		
		INSERT INTO [@PWC_BCOPYM](Code, Name, DocEntry, [Object], LogInst
								, UserSign, CreateDate, CreateTime, DataSource, U_Active)
		SELECT    PayMethCod
				, Descript
				, @m_intDocEntry
				, 'PWC_UDO_BCOPYM'
				, 0
				, UserSign
				, CONVERT(NVARCHAR(8), GETDATE(), 112)
				, REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
				, DataSource
				, Active
		  FROM OPYM
		 WHERE PayMethCod = @list_of_cols_val_tab_del
		
		UPDATE ONNM
		   SET AutoKey=@m_intDocEntry+1
		 WHERE ObjectCode = 'PWC_UDO_BCOPYM'
		
		RETURN		-- 작업 종료시 RETURN을 사용하여 SP EXIT (하나의 오브젝트에 대하여 종료 시점에 작성할 것)
	END
	
	/* 갱신 모드 */
	IF @transaction_type = 'U'
	BEGIN
	
		--* 연결테이블용 지급 방법 갱신
		UPDATE BCOPYM SET Name=Descript, U_Active=Active, UpdateDate=CONVERT(NVARCHAR(8), GETDATE(), 112), UpdateTime=REPLACE(CONVERT(NVARCHAR(5), GETDATE(), 14), ':', '')
		  FROM OPYM
		 INNER JOIN [@PWC_BCOPYM] BCOPYM ON OPYM.PayMethCod=BCOPYM.Code
		 WHERE OPYM.PayMethCod = @list_of_cols_val_tab_del
		
		RETURN
	END
	
	/* 제거 모드 */
	IF @transaction_type = 'D'
	BEGIN
	
		--* 연결테이블용 지급 방법 제거
		DELETE [@PWC_BCOPYM]
		 WHERE Code = @list_of_cols_val_tab_del
		
		RETURN
	END
END

END