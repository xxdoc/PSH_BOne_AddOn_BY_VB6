
/****** Object:  StoredProcedure [dbo].[SBO_SP_TransactionNotification]    Script Date: 02/23/2011 13:08:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[SBO_SP_TransactionNotification] 

@object_type nvarchar(20), 				-- SBO Object Type
@transaction_type nchar(1),				-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
@num_of_cols_in_key int,
@list_of_key_cols_tab_del nvarchar(255),
@list_of_cols_val_tab_del nvarchar(255)

AS
BEGIN

	 --Return values
	DECLARE @ERROR  INT				-- Result (0 for no error)
	DECLARE @ERROR_MESSAGE NVARCHAR(200) 		-- Error string to be displayed
	SELECT @ERROR = 0
	SELECT @ERROR_MESSAGE = N'Ok'

	DECLARE	@R_VALUE01		NVARCHAR(100)
	DECLARE	@R_VALUE02		NVARCHAR(100)
	DECLARE	@R_VALUE03		NVARCHAR(100)
	DECLARE	@RETURNVALUE	INT
	DECLARE	@RETURNSTRING	NVARCHAR(100)

	DECLARE	@V_BASEREF		NVARCHAR(20), @V_BASELINE	INT, @V_ENDQTY	NUMERIC(19,6), @V_BASETYPE	NVARCHAR(20),
			@V_ORDTYPE		NVARCHAR(20),
			@V_TEMPCHR1		NVARCHAR(20), @V_TEMPCHR2 NVARCHAR(20), @V_TEMPCHR3 NVARCHAR(20), 
			@V_TEMPDEC1 INT, @V_TEMPDEC2 INT

	SET @RETURNVALUE = 0
	SET @RETURNSTRING = N'Ok'

	DECLARE @R_mCHECK		CHAR(1)
	--원천문서 월 데이터 불일치 체크 여부
	SELECT TOP 1 @R_mCHECK=U_MINOR FROM [@ZSY001L] WHERE CODE='KBP998'

--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------

	/*STD_AddOn 2010.06.22 허향행 추가*/
	EXEC [DBO].[MDC_STDADDON_SP_TransactionNotification] 
	@object_type,@transaction_type,
	@num_of_cols_in_key,
	@list_of_key_cols_tab_del,
	@list_of_cols_val_tab_del,
	@error OUTPUT,
	@error_message OUTPUT

------------------------------------------------------------------------------------------------------------------------------

	-- 판매오더일때
	IF @object_type = '17' AND @transaction_type IN ('A') 
	BEGIN
		SET @R_VALUE01 = ''
		
		--	영업사원 필수 체크
		IF EXISTS(SELECT DOCENTRY FROM ORDR WHERE DOCENTRY=@list_of_cols_val_tab_del
					AND ISNULL(U_SalesCode,'')='')
		BEGIN
			SET @ERROR = 1
			SET @ERROR_MESSAGE = '[NT] 영업사원이 입력되지 않았습니다.'
		END
		--  간납처코드 필수 체크
		IF EXISTS(SELECT DOCENTRY FROM ORDR WHERE DOCENTRY=@list_of_cols_val_tab_del
					AND ISNULL(NUMATCARD,'')='')
		BEGIN
			SET @ERROR = 1
			SET @ERROR_MESSAGE = '[NT] 간납처코드가 입력되지 않았습니다.'
		END
	END
	
	--	AP송장/AP대변메모/입고PO/반품/납품/반품/AR송장/AR대변메모 단가 제로 체크
	IF @object_type IN ('18','19','20','21','15','16','13','14') AND @transaction_type IN ('A') 
	BEGIN
		--	구매오더 저장시 단가 제로는 불가
		--	신규 디비시 적용 요망.
		SET @R_VALUE01 = ''

		--	간납(파트너)코드 필수 체크
		IF @ERROR=0 AND @object_type = '13'			--	A/R송장
		BEGIN
			IF EXISTS(SELECT DOCENTRY FROM OINV WHERE DOCENTRY=@list_of_cols_val_tab_del
						AND ISNULL(NUMATCARD,'')='')
			BEGIN
				-- 간납처코드 입력
				SELECT @V_BASEREF = CARDCODE, @V_ORDTYPE = NUMATCARD, @V_TEMPCHR1=DOCENTRY FROM OINV
				WHERE DOCENTRY=@list_of_cols_val_tab_del
				
				IF ISNULL(@V_TEMPCHR1,'') <> '' AND ISNULL(@V_ORDTYPE,'') = ''
				BEGIN
					UPDATE OINV SET NUMATCARD=@V_BASEREF WHERE DOCENTRY=@V_TEMPCHR1
				END
			END
		END
		ELSE IF @ERROR=0 AND @object_type = '14'	--	A/R대변메모
		BEGIN
			IF EXISTS(SELECT DOCENTRY FROM ORIN WHERE DOCENTRY=@list_of_cols_val_tab_del
						AND ISNULL(NUMATCARD,'')='')
			BEGIN
				-- 간납처코드 입력
				SELECT @V_BASEREF = CARDCODE, @V_ORDTYPE = NUMATCARD, @V_TEMPCHR1=DOCENTRY FROM ORIN
				WHERE DOCENTRY=@list_of_cols_val_tab_del
				
				IF ISNULL(@V_TEMPCHR1,'') <> '' AND ISNULL(@V_ORDTYPE,'') = ''
				BEGIN
					UPDATE ORIN SET NUMATCARD=@V_BASEREF WHERE DOCENTRY=@V_TEMPCHR1
				END
			END
		END
		
		----	원천문서의 월과 불일치 데이터 체크
		--IF @ERROR=0 AND @R_mCHECK='Y' AND @object_type = '18'			--	AP송장(OPCH0),입고PO(OPDN)
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPCH T0 
		--				INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN PDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OPDN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='20' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] 입고PO의 전기월과 A/P송장의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--	ELSE IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN POR1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OPOR T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='22' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] 구매오더의 전기월과 A/P송장의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '19'			--	A/P대변메모
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN PCH1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OPCH T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='18' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] A/P송장의 전기월과 A/P대변메모의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN RPD1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ORPD T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='21' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] 입고반품의 전기월과 A/P대변메모의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '20'			--	입고po
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN RPD1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ORPD T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='21' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] 입고반품의 전기월과 입고PO의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '21'			--	입고반품
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPD T0 INNER JOIN RPD1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN PDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OPDN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='20' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] 입고PO의 전기월과 입고반품의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '13'			--	A/R송장
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OINV T0 INNER JOIN INV1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN DLN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ODLN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='15' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] 납품의 전기월과 A/R송장의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '14'			--	A/P대변메모
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN INV1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OINV T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='13' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] A/R송장의 전기월과 A/R대변메모의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN RDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ORDN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='16' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] 반품의 전기월과 A/R대변메모의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '15'			--	입고po
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ODLN T0 INNER JOIN DLN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN RDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ORDN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='16' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] 반품의 전기월과 납품의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '16'			--	입고반품
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORDN T0 INNER JOIN RDN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN DLN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ODLN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='15' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] 납품의 전기월과 반품의 전기월이 일치하지 않습니다.[관리자에게 문의바람]'
		--	END
		--END
		
		-- A/R송장 체크
		IF @object_type = '13' AND @ERROR = 0 
		BEGIN
			SELECT TOP 1 @V_TEMPCHR1=TRANSID, @V_TEMPCHR2=NUMATCARD, @V_TEMPCHR3=CARDCODE, @V_TEMPDEC1=BPLID, @V_TEMPDEC2 = Transid
			FROM OINV T0 WHERE T0.DOCENTRY=@list_of_cols_val_tab_del

			--SELECT BASEREF1 FROM JDT1 WHERE T0.DOCENTRY=@list_of_cols_val_tab_del

			--	분개에 업데이트(간납코드(파트너 코드) = U_EARDCODE)
			UPDATE OJDT SET U_BPLID=@V_TEMPDEC1 WHERE TRANSID=@V_TEMPCHR1
			UPDATE JDT1 SET U_EARDCODE=@V_TEMPCHR2, U_VATBP=@V_TEMPCHR3, U_EARDNAME = (SELECT TOP 1 CARDNAME FROM OCRD 
																  WHERE CARDCODE = @V_TEMPCHR2)
			WHERE TRANSID=@V_TEMPCHR1
			
			--관리계정코드를 가지고 관리계정명을 입력
			UPDATE A SET U_AcctName = B.AcctName
			FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
			WHERE TRANSID = @V_TEMPCHR1
		END
		--  AR대변메모 체크
		IF @object_type = '14' AND @ERROR = 0 
		BEGIN
			SELECT TOP 1 @V_BASETYPE=T1.BASETYPE, @V_BASEREF=T1.BASEENTRY,
							@V_TEMPCHR1=T0.TRANSID, @V_TEMPCHR2=T0.NUMATCARD, @V_TEMPCHR3=CARDCODE, @V_TEMPDEC1=T0.BPLID
			FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY 
			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del 

			--	분개에 업데이트
			UPDATE OJDT SET U_BPLID=@V_TEMPDEC1 WHERE TRANSID=@V_TEMPCHR1
			UPDATE JDT1 SET U_EARDCODE=@V_TEMPCHR2, U_VATBP=@V_TEMPCHR3, U_EARDNAME=(SELECT TOP 1 CARDNAME FROM OCRD 
																WHERE CARDCODE = @V_TEMPCHR1) 
			WHERE TRANSID=@V_TEMPCHR1
			
			--관리계정코드를 가지고 관리계정명을 입력
			UPDATE A SET U_AcctName = B.AcctName
			FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
			WHERE TRANSID = @V_TEMPCHR1
		END
		--  AP송장 체크
		IF @object_type = '18' AND @ERROR = 0 
		BEGIN
			SELECT TOP 1 @V_BASETYPE=T1.BASETYPE, @V_BASEREF=T1.BASEENTRY,
							@V_TEMPCHR1=T0.TRANSID, @V_TEMPCHR2=T0.NUMATCARD, @V_TEMPCHR3=CARDCODE, @V_TEMPDEC1=T0.BPLID
			FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY 
			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del 

			--	분개에 업데이트
			UPDATE OJDT SET U_BPLID=@V_TEMPDEC1 WHERE TRANSID=@V_TEMPCHR1
			UPDATE JDT1 SET U_EARDCODE=@V_TEMPCHR2, U_VATBP=@V_TEMPCHR3
			WHERE TRANSID=@V_TEMPCHR1
			
			--관리계정코드를 가지고 관리계정명을 입력
			UPDATE A SET U_AcctName = B.AcctName
			FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
			WHERE TRANSID = @V_TEMPCHR1
		END
		--  AP대변메모 체크
		IF @object_type = '19' AND @ERROR = 0 
		BEGIN
			SELECT TOP 1 @V_BASETYPE=T1.BASETYPE, @V_BASEREF=T1.BASEENTRY,
							@V_TEMPCHR1=T0.TRANSID, @V_TEMPCHR2=T0.NUMATCARD, @V_TEMPCHR3=CARDCODE, @V_TEMPDEC1=T0.BPLID
			FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY 
			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del 

			--	분개에 업데이트
			UPDATE OJDT SET U_BPLID=@V_TEMPDEC1 WHERE TRANSID=@V_TEMPCHR1
			UPDATE JDT1 SET U_EARDCODE=@V_TEMPCHR2, U_VATBP=@V_TEMPCHR3
			WHERE TRANSID=@V_TEMPCHR1
			
			--관리계정코드를 가지고 관리계정명을 입력
			UPDATE A SET U_AcctName = B.AcctName
			FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
			WHERE TRANSID = @V_TEMPCHR1
		END
	END

	---- 기타출고시 원재료일 경우 단가와 품목원가가 틀릴경우 입력방지
	--IF @object_type = '60' AND @transaction_type IN ('A') 
	--	BEGIN
	--		SET @R_VALUE01 = ''

	--		SELECT @R_VALUE01 = T0.ITEMCODE FROM [dbo].[IGE1] T0 INNER JOIN OITM T1 ON T0.ITEMCODE = T1.ITEMCODE
	--		WHERE T0.PRICE <> 0
	--			AND T0.BASETYPE = -1
	--			AND T0.U_LINETYPE IN ('02','03','04','05')
	--			AND T0.DocEntry = @list_of_cols_val_tab_del 

	--		IF @R_VALUE01 <> ''
	--		BEGIN
	--			SET @ERROR = 1
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]타계정유형은 단가가 제로이여야 합니다.'
	--		END
			
	--		SET @R_VALUE01 = ''

	--		SELECT @R_VALUE01 = T0.ITEMCODE FROM [dbo].[IGE1] T0 INNER JOIN OITM T1 ON T0.ITEMCODE = T1.ITEMCODE
	--		WHERE T0.PRICE = 0
	--			AND T0.BASETYPE = -1
	--			AND T0.U_LINETYPE IN ('02','03','04','05')
	--			AND T0.DocEntry = @list_of_cols_val_tab_del 

	--		IF @R_VALUE01 <> ''
	--		BEGIN
	--			SET @ERROR = 1
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]해당 타계정유형은 단가 입력은 필수입니다.'
	--		END
	--	END

	----	타계정 출고일 경우 계정설정확인
	--IF @object_type = '60' AND @transaction_type IN ('A') 
	--BEGIN
	--	SET @R_VALUE01 = ''
		
	--	SELECT @R_VALUE01 = A.ITEMCODE
	--	FROM IGE1 A INNER JOIN OIGE B ON A.DOCENTRY = B.DOCENTRY
	--	WHERE A.BASETYPE = -1 
	--		AND ISNULL(A.U_LINETYPE,'') = ''
	--		AND A.DOCENTRY = @list_of_cols_val_tab_del

	--	IF @R_VALUE01 <>  ''
	--		BEGIN
	--			SET @ERROR = 1
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]입출고유형이 설정이 되어 있지 않습니다.'
	--		END
	--	ELSE
	--		BEGIN
			
	--		SET @R_VALUE01 = ''
			
	--		SELECT @R_VALUE01 = A.ITEMCODE FROM IGE1 A INNER JOIN OIGE B ON A.DOCENTRY = B.DOCENTRY 
	--		INNER JOIN OITM C ON A.ITEMCODE = C.ITEMCODE
	--		INNER JOIN [@ZSY001L] D ON Convert(NvarChar(3),C.ItmsGrpCod) = left(D.U_CdName,3)
	--									AND A.U_LINETYPE = left(D.U_RelCd,2)
	--									AND D.CODE = 'KBP008'
	--		WHERE A.BASETYPE = -1 
	--			AND A.ACCTCODE <> ISNULL(D.U_MINOR,'') 
	--			AND A.DOCENTRY = @list_of_cols_val_tab_del
				
	--		IF @R_VALUE01 <> ''
	--		BEGIN
	--			SET @ERROR = 1
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]출고유형에 따른 계정설정이 맞지 않습니다.'
	--		END
	--	END
	--	--IF @R_VALUE01 = ''  OR @R_VALUE01 IS NULL
	--	--BEGIN
	--	--	--	7번대 계정일 경우 차원1 필수 입력
	--	--	SELECT TOP 1 @R_VALUE01 = A.ITEMCODE FROM IGE1 A INNER JOIN OIGE B ON A.DOCENTRY=B.DOCENTRY
	--	--	WHERE A.DOCENTRY = @list_of_cols_val_tab_del AND A.BASETYPE = -1 
	--	--		AND LEFT(A.ACCTCODE, 1) = '7' AND ISNULL(A.OCRCODE, '') = ''
	--	--		AND ISNULL(A.U_LINETYPE,'') <> '' 

	--	--	IF @R_VALUE01 <> ''
	--	--	BEGIN
	--	--		SET @ERROR = 1
	--	--		SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]비용계정일 경우 차원1은 필수 항목입니다.'
	--	--	END
	--	--END
	--END

	----	타계정 입고일 경우 계정설정확인
	--IF @object_type = '59' AND @transaction_type IN ('A') 
	--BEGIN
	--	SET @R_VALUE01 = ''
		
	--	SELECT @R_VALUE01 = A.ITEMCODE
	--	FROM IGN1 A INNER JOIN OIGN B ON A.DOCENTRY = B.DOCENTRY
	--	WHERE A.BASETYPE = -1 
	--		AND ISNULL(A.U_LINETYPE,'') = ''
	--		AND A.DOCENTRY = @list_of_cols_val_tab_del

	--	IF @R_VALUE01 <>  ''
	--		BEGIN
	--			SET @ERROR = 1
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]입출고유형이 설정이 되어 있지 않습니다.'
	--		END
	--	ELSE
	--		BEGIN
			
	--		SET @R_VALUE01 = ''
			
	--		SELECT @R_VALUE01 = A.ITEMCODE FROM IGN1 A INNER JOIN OIGN B ON A.DOCENTRY = B.DOCENTRY 
	--		INNER JOIN OITM C ON A.ITEMCODE = C.ITEMCODE
	--		INNER JOIN [@ZSY001L] D ON Convert(NvarChar(3),C.ItmsGrpCod) = left(D.U_CdName,3)
	--									AND A.U_LINETYPE = left(D.U_RelCd,2)
	--									AND D.CODE = 'KBP008'
	--		WHERE A.BASETYPE = -1 
	--			AND A.ACCTCODE <> ISNULL(D.U_MINOR,'') 
	--			AND A.DOCENTRY = @list_of_cols_val_tab_del
				
	--		IF @R_VALUE01 <> ''
	--		BEGIN
	--			SET @ERROR = 1
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]출고유형에 따른 계정설정이 맞지 않습니다.'
	--		END
	--	END
	--	--IF @R_VALUE01 = ''  OR @R_VALUE01 IS NULL
	--	--BEGIN
	--	--	--	7번대 계정일 경우 차원1 필수 입력
	--	--	SELECT TOP 1 @R_VALUE01 = A.ITEMCODE FROM IGN1 A INNER JOIN OIGN B ON A.DOCENTRY=B.DOCENTRY
	--	--	WHERE A.DOCENTRY = @list_of_cols_val_tab_del AND A.BASETYPE = -1 
	--	--		AND LEFT(A.ACCTCODE, 1) = '7' AND ISNULL(A.OCRCODE, '') = ''
	--	--		AND ISNULL(A.U_LINETYPE,'') <> '' 

	--	--	IF @R_VALUE01 <> ''
	--	--	BEGIN
	--	--		SET @ERROR = 1
	--	--		SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]비용계정일 경우 차원1은 필수 항목입니다.'
	--	--	END
	--	--END
	--END

	--	지급시 분개에 관리번호 추가
	IF @OBJECT_TYPE IN ('46') AND @transaction_type IN ('A') 
	BEGIN
		SET	@R_VALUE01	= ''
		SET	@R_VALUE02	= ''
		SET	@R_VALUE03	= ''
		--지급에서 분개 번호 추출
		SELECT @R_VALUE01 = U_CTRNUM FROM OVPM
		WHERE DOCENTRY = @list_of_cols_val_tab_del
		
		--분개에 관리번호가 몇번까지 갔는지 체크
		--SELECT CONVERT(INT,RIGHT(MAX(U_MNum),3)) + 1  FROM OJDT WHERE TransId = 52798
		SELECT @R_VALUE02 = CONVERT(INT,RIGHT(MAX(U_MNum),3)) FROM OJDT 
		WHERE RefDate = CONVERT(VARCHAR(8),GETDATE(),112)
		
		SET @R_VALUE02 = RIGHT(CONVERT(VARCHAR(8),GETDATE(),112),6) + [DBO].[USER_NumZero] (@R_VALUE02 + 1,3)
		
		----분개에 관리번호 수정
		UPDATE OJDT SET U_MNum = @R_VALUE02
		WHERE TransId IN (SELECT TransId FROM OVPM WHERE U_CTRNUM = @R_VALUE01)
	END
	
	--	입금시 원천문서인 A/R송장, A/R대변메모의 파트너 체크
	IF @OBJECT_TYPE IN ('24') AND @transaction_type IN ('A') 
	BEGIN
		SET	@R_VALUE01	= ''
		SET	@R_VALUE02	= ''
		SET @V_BASETYPE = ''

		-- 입금에 간납처코드 체크
		SELECT @R_VALUE01=U_EARDCODE, @R_VALUE02 = U_PAYTYPE FROM ORCT WHERE DOCENTRY = @list_of_cols_val_tab_del
		BEGIN
			IF ISNULL(@R_VALUE01,'') = ''
			BEGIN
				--B1에서 직접 제품을 입금처리시 간납처코드를 필수 입력하도록 해야됨.
				SET @ERROR = 1
				SET @ERROR_MESSAGE = '[NT]입금문서 작성시 간납처코드를 입력 하셔야 합니다.'
			END
			IF ISNULL(@R_VALUE02,'') = ''
			BEGIN
				--B1에서 직접 제품을 입금처리시 간납처코드를 필수 입력하도록 해야됨.
				SET @ERROR = 1
				SET @ERROR_MESSAGE = '[NT]입금문서 작성시 수금유형을 선택하셔야 합니다.'
			END
		END
		
		
		SELECT * FROM ONNM WHERE OBJECTCODE = '17'
		
		SET	@R_VALUE01	= ''
		SET @V_BASETYPE = ''
		
		IF EXISTS (SELECT COUNT(*)
					FROM (
						SELECT T.NUMATCARD
						FROM (
							SELECT B.NUMATCARD 
							FROM RCT2 A
							INNER JOIN OINV B ON A.DOCENTRY=B.DOCENTRY AND A.INVTYPE=B.OBJTYPE
							WHERE A.INVTYPE IN ('13') AND A.DOCNUM=@list_of_cols_val_tab_del
							GROUP BY B.NUMATCARD
							UNION ALL
							SELECT B.NUMATCARD
							FROM RCT2 A
							INNER JOIN ORIN B ON A.DOCENTRY=B.DOCENTRY AND A.INVTYPE=B.OBJTYPE
							WHERE A.INVTYPE IN ('14') AND A.DOCNUM=@list_of_cols_val_tab_del
							GROUP BY B.NUMATCARD
	--						UNION ALL
	--						SELECT B.U_EARDCODE
	--						FROM RCT2 A
	--						INNER JOIN JDT1 B ON A.DOCENTRY=B.TRANSID AND A.INVTYPE=B.OBJTYPE
	--						WHERE A.INVTYPE IN ('30') AND A.DOCNUM=@list_of_cols_val_tab_del
	--						GROUP BY B.U_EARDCODE
						) T GROUP BY T.NUMATCARD
					) T1 HAVING COUNT(*)>1)
		BEGIN
			SET @ERROR = 1
			SET @ERROR_MESSAGE = '[NT]입금문서 작성시 각 간납처별로 입금을 하셔야 합니다.'	
		END
		ELSE
		BEGIN
			--	분개에 간납처 업데이트
			SELECT TOP 1 @V_BASETYPE=T.NUMATCARD, @V_BASEREF=T.TRANSID
			FROM (
				SELECT B.NUMATCARD, AA.TRANSID
				FROM RCT2 A INNER JOIN ORCT AA ON A.DOCNUM=AA.DOCENTRY
				INNER JOIN OINV B ON A.DOCENTRY=B.DOCENTRY AND A.INVTYPE=B.OBJTYPE
				WHERE A.INVTYPE IN ('13') AND A.DOCNUM=@list_of_cols_val_tab_del
				AND AA.CARDCODE=B.CARDCODE
				GROUP BY B.NUMATCARD, AA.TRANSID
				UNION ALL
				SELECT B.NUMATCARD, AA.TRANSID
				FROM RCT2 A INNER JOIN ORCT AA ON A.DOCNUM=AA.DOCENTRY
				INNER JOIN ORIN B ON A.DOCENTRY=B.DOCENTRY AND A.INVTYPE=B.OBJTYPE
				WHERE A.INVTYPE IN ('14') AND A.DOCNUM=@list_of_cols_val_tab_del
				AND AA.CARDCODE=B.CARDCODE
				GROUP BY B.NUMATCARD, AA.TRANSID
				UNION ALL
				SELECT B.U_EARDCODE, AA.TRANSID
				FROM RCT2 A INNER JOIN ORCT AA ON A.DOCNUM=AA.DOCENTRY
				INNER JOIN JDT1 B ON A.DOCENTRY=B.TRANSID AND A.INVTYPE=B.OBJTYPE
				WHERE A.INVTYPE IN ('30') AND A.DOCNUM=@list_of_cols_val_tab_del
				AND AA.CARDCODE=B.SHORTNAME
				GROUP BY B.U_EARDCODE, AA.TRANSID
			) T GROUP BY T.NUMATCARD, T.TRANSID
			
			-- 분개에 간납처코드
			UPDATE JDT1 SET U_EARDCODE=@V_BASETYPE WHERE TRANSID=@V_BASEREF
		END
		----	어음번호 체크(10 OR 20)
		--IF EXISTS(SELECT B.REFNUM FROM ORCT A INNER JOIN OBOE B ON A.BOEABS=B.BOEKEY
		--			WHERE A.DOCENTRY=@list_of_cols_val_tab_del
		--			AND A.BOESUM>0
		--			AND NOT ISNULL(LEN(B.REFNUM),0) IN (10,20))
		--BEGIN
		--	SET @ERROR = 1
		--	SET @ERROR_MESSAGE = '[NT]어음번호의 자릿수는 10 or 20 이여야 합니다.'
		--END
	END

	--	어음 예금시 원천문서인 A/R송장, A/R대변메모의 간납처 체크
	IF @OBJECT_TYPE IN ('182') AND @transaction_type IN ('A') 
	BEGIN
		SET	@R_VALUE01	= ''
		SET @V_BASETYPE = ''

		DECLARE S60051 SCROLL CURSOR FOR

		SELECT D.U_EARDCODE, A.TRANSID, B.BOEABS
		FROM OBOT A
		INNER JOIN BOT1 B ON A.ABSENTRY=B.ABSENTRY 
		INNER JOIN ORCT C ON B.BOEABS=C.BOEABS
		INNER JOIN JDT1 D ON C.TRANSID=D.TRANSID
		WHERE B.BOETYPE='I'
		AND ISNULL(D.U_EARDCODE,'')<>''
		AND A.ABSENTRY=@list_of_cols_val_tab_del
		GROUP BY D.U_EARDCODE, A.TRANSID, B.BOEABS
		
		OPEN S60051
		FETCH NEXT FROM S60051 INTO	@V_BASETYPE, @V_BASEREF, @V_BASELINE

		WHILE (@@FETCH_STATUS = 0)
		BEGIN
			--	분개에 간납처 업데이트
			UPDATE JDT1 SET U_EARDCODE=@V_BASETYPE 
			WHERE TRANSID=@V_BASEREF AND REF1=@V_BASELINE AND TRANSTYPE='182' AND CREATEDBY=@list_of_cols_val_tab_del	

			FETCH NEXT FROM S60051 INTO	@V_BASETYPE, @V_BASEREF, @V_BASELINE
		END
		CLOSE S60051
		DEALLOCATE S60051
	END
	--	비즈니스파트너 체크(이름, 세금코드)
	IF @OBJECT_TYPE IN ('2') AND @transaction_type IN ('A') 
	BEGIN
		SET @V_BASEREF	=''
		SET @V_BASETYPE	=''
		SET @V_ORDTYPE	=''
		SET @V_TEMPCHR1	=''
		SET @R_VALUE01	=''

		SELECT	@R_VALUE01=ISNULL(CARDNAME,''),
				@V_BASEREF=ISNULL(ECVATGROUP,''),
				@V_TEMPCHR1=CARDTYPE
		FROM OCRD 
		WHERE CARDCODE=@list_of_cols_val_tab_del
			
		IF @V_BASEREF=''
		BEGIN
			--SET @ERROR = 1
			SET @ERROR_MESSAGE = '[NT]회계텝의 세금코드를 입력해 주세요.'
		END
		ELSE IF @R_VALUE01=''
		BEGIN
			SET @ERROR = 1
			SET @ERROR_MESSAGE = '[NT]거래처명을 입력해 주세요.'
		END
	END
	IF @OBJECT_TYPE IN ('30') AND @transaction_type IN ('A','U') 
	BEGIN
		SET @V_BASEREF	=''
		SET @V_BASETYPE	=''
		SET @V_ORDTYPE	=''
		SET @V_TEMPCHR1	=''
		SET @R_VALUE01	=''
		SET @R_VALUE02	=''

		SELECT	@V_BASEREF=ISNULL(PROFITCODE,''),
				@R_VALUE01=Transid
		FROM JDT1 
		WHERE Transid=@list_of_cols_val_tab_del
		AND left(Account ,2) in ('55','62','63')
		
		IF @R_VALUE01<>''
		BEGIN
			IF @V_BASEREF=''
				BEGIN
					SET @ERROR = 1
					SET @ERROR_MESSAGE = '[NT]분개-코스트센터를 입력해 주세요.'
				END
			ELSE
				BEGIN
					SELECT	@V_BASEREF=ISNULL(U_PFCODE,''),
							@R_VALUE01=Transid
					FROM JDT1 
					WHERE Transid=@list_of_cols_val_tab_del
					AND Account IN ('550101000','550121000','550081000','550082000','550083000')
					
					IF @R_VALUE01<>''
					BEGIN
						IF @V_BASEREF=''
						BEGIN
							SET @ERROR = 1
							SET @ERROR_MESSAGE = '[NT]분개-손익거래처를 입력해 주세요.'
						END
					END
				END
		END
		
		--관리계정코드를 가지고 관리계정명을 입력
		UPDATE A SET U_AcctName = B.AcctName
		FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
		WHERE TRANSID = @list_of_cols_val_tab_del
		
		--BP거래처가 9로 시작하는 거래처에 법인카드 번호를 입력
		UPDATE A SET A.REF3LINE = B.CNTCTPRSN
		FROM JDT1 A INNER JOIN OCRD B ON A.ShortName = B.CARDCODE AND LEFT(B.CARDCODE,1) = '9'
		WHERE Transid=@list_of_cols_val_tab_del
	END
END 

SELECT @ERROR, @ERROR_MESSAGE 
