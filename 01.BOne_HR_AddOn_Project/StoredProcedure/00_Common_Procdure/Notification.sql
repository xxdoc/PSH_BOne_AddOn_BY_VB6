
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
	--��õ���� �� ������ ����ġ üũ ����
	SELECT TOP 1 @R_mCHECK=U_MINOR FROM [@ZSY001L] WHERE CODE='KBP998'

--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------

	/*STD_AddOn 2010.06.22 ������ �߰�*/
	EXEC [DBO].[MDC_STDADDON_SP_TransactionNotification] 
	@object_type,@transaction_type,
	@num_of_cols_in_key,
	@list_of_key_cols_tab_del,
	@list_of_cols_val_tab_del,
	@error OUTPUT,
	@error_message OUTPUT

------------------------------------------------------------------------------------------------------------------------------

	-- �Ǹſ����϶�
	IF @object_type = '17' AND @transaction_type IN ('A') 
	BEGIN
		SET @R_VALUE01 = ''
		
		--	������� �ʼ� üũ
		IF EXISTS(SELECT DOCENTRY FROM ORDR WHERE DOCENTRY=@list_of_cols_val_tab_del
					AND ISNULL(U_SalesCode,'')='')
		BEGIN
			SET @ERROR = 1
			SET @ERROR_MESSAGE = '[NT] ��������� �Էµ��� �ʾҽ��ϴ�.'
		END
		--  ����ó�ڵ� �ʼ� üũ
		IF EXISTS(SELECT DOCENTRY FROM ORDR WHERE DOCENTRY=@list_of_cols_val_tab_del
					AND ISNULL(NUMATCARD,'')='')
		BEGIN
			SET @ERROR = 1
			SET @ERROR_MESSAGE = '[NT] ����ó�ڵ尡 �Էµ��� �ʾҽ��ϴ�.'
		END
	END
	
	--	AP����/AP�뺯�޸�/�԰�PO/��ǰ/��ǰ/��ǰ/AR����/AR�뺯�޸� �ܰ� ���� üũ
	IF @object_type IN ('18','19','20','21','15','16','13','14') AND @transaction_type IN ('A') 
	BEGIN
		--	���ſ��� ����� �ܰ� ���δ� �Ұ�
		--	�ű� ���� ���� ���.
		SET @R_VALUE01 = ''

		--	����(��Ʈ��)�ڵ� �ʼ� üũ
		IF @ERROR=0 AND @object_type = '13'			--	A/R����
		BEGIN
			IF EXISTS(SELECT DOCENTRY FROM OINV WHERE DOCENTRY=@list_of_cols_val_tab_del
						AND ISNULL(NUMATCARD,'')='')
			BEGIN
				-- ����ó�ڵ� �Է�
				SELECT @V_BASEREF = CARDCODE, @V_ORDTYPE = NUMATCARD, @V_TEMPCHR1=DOCENTRY FROM OINV
				WHERE DOCENTRY=@list_of_cols_val_tab_del
				
				IF ISNULL(@V_TEMPCHR1,'') <> '' AND ISNULL(@V_ORDTYPE,'') = ''
				BEGIN
					UPDATE OINV SET NUMATCARD=@V_BASEREF WHERE DOCENTRY=@V_TEMPCHR1
				END
			END
		END
		ELSE IF @ERROR=0 AND @object_type = '14'	--	A/R�뺯�޸�
		BEGIN
			IF EXISTS(SELECT DOCENTRY FROM ORIN WHERE DOCENTRY=@list_of_cols_val_tab_del
						AND ISNULL(NUMATCARD,'')='')
			BEGIN
				-- ����ó�ڵ� �Է�
				SELECT @V_BASEREF = CARDCODE, @V_ORDTYPE = NUMATCARD, @V_TEMPCHR1=DOCENTRY FROM ORIN
				WHERE DOCENTRY=@list_of_cols_val_tab_del
				
				IF ISNULL(@V_TEMPCHR1,'') <> '' AND ISNULL(@V_ORDTYPE,'') = ''
				BEGIN
					UPDATE ORIN SET NUMATCARD=@V_BASEREF WHERE DOCENTRY=@V_TEMPCHR1
				END
			END
		END
		
		----	��õ������ ���� ����ġ ������ üũ
		--IF @ERROR=0 AND @R_mCHECK='Y' AND @object_type = '18'			--	AP����(OPCH0),�԰�PO(OPDN)
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPCH T0 
		--				INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN PDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OPDN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='20' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] �԰�PO�� ������� A/P������ ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--	ELSE IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN POR1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OPOR T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='22' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] ���ſ����� ������� A/P������ ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '19'			--	A/P�뺯�޸�
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN PCH1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OPCH T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='18' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] A/P������ ������� A/P�뺯�޸��� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN RPD1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ORPD T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='21' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] �԰��ǰ�� ������� A/P�뺯�޸��� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '20'			--	�԰�po
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN RPD1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ORPD T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='21' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] �԰��ǰ�� ������� �԰�PO�� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '21'			--	�԰��ǰ
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPD T0 INNER JOIN RPD1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN PDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OPDN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='20' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] �԰�PO�� ������� �԰��ǰ�� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '13'			--	A/R����
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OINV T0 INNER JOIN INV1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN DLN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ODLN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='15' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] ��ǰ�� ������� A/R������ ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '14'			--	A/P�뺯�޸�
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN INV1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN OINV T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='13' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] A/R������ ������� A/R�뺯�޸��� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN RDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ORDN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='16' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] ��ǰ�� ������� A/R�뺯�޸��� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '15'			--	�԰�po
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ODLN T0 INNER JOIN DLN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN RDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ORDN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='16' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] ��ǰ�� ������� ��ǰ�� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--END
		--ELSE IF @ERROR=0 AND @R_mCHECK='Y' AND @R_VALUE01='' AND @object_type = '16'			--	�԰��ǰ
		--BEGIN
		--	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORDN T0 INNER JOIN RDN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
		--				INNER JOIN DLN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
		--				INNER JOIN ODLN T3 ON T2.DOCENTRY=T3.DOCENTRY
		--				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='15' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
		--	BEGIN
		--		SET @ERROR = 1
		--		SET @ERROR_MESSAGE = '[NT] ��ǰ�� ������� ��ǰ�� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		--	END
		--END
		
		-- A/R���� üũ
		IF @object_type = '13' AND @ERROR = 0 
		BEGIN
			SELECT TOP 1 @V_TEMPCHR1=TRANSID, @V_TEMPCHR2=NUMATCARD, @V_TEMPCHR3=CARDCODE, @V_TEMPDEC1=BPLID, @V_TEMPDEC2 = Transid
			FROM OINV T0 WHERE T0.DOCENTRY=@list_of_cols_val_tab_del

			--SELECT BASEREF1 FROM JDT1 WHERE T0.DOCENTRY=@list_of_cols_val_tab_del

			--	�а��� ������Ʈ(�����ڵ�(��Ʈ�� �ڵ�) = U_EARDCODE)
			UPDATE OJDT SET U_BPLID=@V_TEMPDEC1 WHERE TRANSID=@V_TEMPCHR1
			UPDATE JDT1 SET U_EARDCODE=@V_TEMPCHR2, U_VATBP=@V_TEMPCHR3, U_EARDNAME = (SELECT TOP 1 CARDNAME FROM OCRD 
																  WHERE CARDCODE = @V_TEMPCHR2)
			WHERE TRANSID=@V_TEMPCHR1
			
			--���������ڵ带 ������ ������������ �Է�
			UPDATE A SET U_AcctName = B.AcctName
			FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
			WHERE TRANSID = @V_TEMPCHR1
		END
		--  AR�뺯�޸� üũ
		IF @object_type = '14' AND @ERROR = 0 
		BEGIN
			SELECT TOP 1 @V_BASETYPE=T1.BASETYPE, @V_BASEREF=T1.BASEENTRY,
							@V_TEMPCHR1=T0.TRANSID, @V_TEMPCHR2=T0.NUMATCARD, @V_TEMPCHR3=CARDCODE, @V_TEMPDEC1=T0.BPLID
			FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY 
			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del 

			--	�а��� ������Ʈ
			UPDATE OJDT SET U_BPLID=@V_TEMPDEC1 WHERE TRANSID=@V_TEMPCHR1
			UPDATE JDT1 SET U_EARDCODE=@V_TEMPCHR2, U_VATBP=@V_TEMPCHR3, U_EARDNAME=(SELECT TOP 1 CARDNAME FROM OCRD 
																WHERE CARDCODE = @V_TEMPCHR1) 
			WHERE TRANSID=@V_TEMPCHR1
			
			--���������ڵ带 ������ ������������ �Է�
			UPDATE A SET U_AcctName = B.AcctName
			FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
			WHERE TRANSID = @V_TEMPCHR1
		END
		--  AP���� üũ
		IF @object_type = '18' AND @ERROR = 0 
		BEGIN
			SELECT TOP 1 @V_BASETYPE=T1.BASETYPE, @V_BASEREF=T1.BASEENTRY,
							@V_TEMPCHR1=T0.TRANSID, @V_TEMPCHR2=T0.NUMATCARD, @V_TEMPCHR3=CARDCODE, @V_TEMPDEC1=T0.BPLID
			FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY 
			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del 

			--	�а��� ������Ʈ
			UPDATE OJDT SET U_BPLID=@V_TEMPDEC1 WHERE TRANSID=@V_TEMPCHR1
			UPDATE JDT1 SET U_EARDCODE=@V_TEMPCHR2, U_VATBP=@V_TEMPCHR3
			WHERE TRANSID=@V_TEMPCHR1
			
			--���������ڵ带 ������ ������������ �Է�
			UPDATE A SET U_AcctName = B.AcctName
			FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
			WHERE TRANSID = @V_TEMPCHR1
		END
		--  AP�뺯�޸� üũ
		IF @object_type = '19' AND @ERROR = 0 
		BEGIN
			SELECT TOP 1 @V_BASETYPE=T1.BASETYPE, @V_BASEREF=T1.BASEENTRY,
							@V_TEMPCHR1=T0.TRANSID, @V_TEMPCHR2=T0.NUMATCARD, @V_TEMPCHR3=CARDCODE, @V_TEMPDEC1=T0.BPLID
			FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY 
			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del 

			--	�а��� ������Ʈ
			UPDATE OJDT SET U_BPLID=@V_TEMPDEC1 WHERE TRANSID=@V_TEMPCHR1
			UPDATE JDT1 SET U_EARDCODE=@V_TEMPCHR2, U_VATBP=@V_TEMPCHR3
			WHERE TRANSID=@V_TEMPCHR1
			
			--���������ڵ带 ������ ������������ �Է�
			UPDATE A SET U_AcctName = B.AcctName
			FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
			WHERE TRANSID = @V_TEMPCHR1
		END
	END

	---- ��Ÿ���� ������� ��� �ܰ��� ǰ������� Ʋ����� �Է¹���
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
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]Ÿ���������� �ܰ��� �����̿��� �մϴ�.'
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
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]�ش� Ÿ���������� �ܰ� �Է��� �ʼ��Դϴ�.'
	--		END
	--	END

	----	Ÿ���� ����� ��� ��������Ȯ��
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
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]����������� ������ �Ǿ� ���� �ʽ��ϴ�.'
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
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]��������� ���� ���������� ���� �ʽ��ϴ�.'
	--		END
	--	END
	--	--IF @R_VALUE01 = ''  OR @R_VALUE01 IS NULL
	--	--BEGIN
	--	--	--	7���� ������ ��� ����1 �ʼ� �Է�
	--	--	SELECT TOP 1 @R_VALUE01 = A.ITEMCODE FROM IGE1 A INNER JOIN OIGE B ON A.DOCENTRY=B.DOCENTRY
	--	--	WHERE A.DOCENTRY = @list_of_cols_val_tab_del AND A.BASETYPE = -1 
	--	--		AND LEFT(A.ACCTCODE, 1) = '7' AND ISNULL(A.OCRCODE, '') = ''
	--	--		AND ISNULL(A.U_LINETYPE,'') <> '' 

	--	--	IF @R_VALUE01 <> ''
	--	--	BEGIN
	--	--		SET @ERROR = 1
	--	--		SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]�������� ��� ����1�� �ʼ� �׸��Դϴ�.'
	--	--	END
	--	--END
	--END

	----	Ÿ���� �԰��� ��� ��������Ȯ��
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
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]����������� ������ �Ǿ� ���� �ʽ��ϴ�.'
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
	--			SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]��������� ���� ���������� ���� �ʽ��ϴ�.'
	--		END
	--	END
	--	--IF @R_VALUE01 = ''  OR @R_VALUE01 IS NULL
	--	--BEGIN
	--	--	--	7���� ������ ��� ����1 �ʼ� �Է�
	--	--	SELECT TOP 1 @R_VALUE01 = A.ITEMCODE FROM IGN1 A INNER JOIN OIGN B ON A.DOCENTRY=B.DOCENTRY
	--	--	WHERE A.DOCENTRY = @list_of_cols_val_tab_del AND A.BASETYPE = -1 
	--	--		AND LEFT(A.ACCTCODE, 1) = '7' AND ISNULL(A.OCRCODE, '') = ''
	--	--		AND ISNULL(A.U_LINETYPE,'') <> '' 

	--	--	IF @R_VALUE01 <> ''
	--	--	BEGIN
	--	--		SET @ERROR = 1
	--	--		SET @ERROR_MESSAGE = @R_VALUE01 + '[NT]�������� ��� ����1�� �ʼ� �׸��Դϴ�.'
	--	--	END
	--	--END
	--END

	--	���޽� �а��� ������ȣ �߰�
	IF @OBJECT_TYPE IN ('46') AND @transaction_type IN ('A') 
	BEGIN
		SET	@R_VALUE01	= ''
		SET	@R_VALUE02	= ''
		SET	@R_VALUE03	= ''
		--���޿��� �а� ��ȣ ����
		SELECT @R_VALUE01 = U_CTRNUM FROM OVPM
		WHERE DOCENTRY = @list_of_cols_val_tab_del
		
		--�а��� ������ȣ�� ������� ������ üũ
		--SELECT CONVERT(INT,RIGHT(MAX(U_MNum),3)) + 1  FROM OJDT WHERE TransId = 52798
		SELECT @R_VALUE02 = CONVERT(INT,RIGHT(MAX(U_MNum),3)) FROM OJDT 
		WHERE RefDate = CONVERT(VARCHAR(8),GETDATE(),112)
		
		SET @R_VALUE02 = RIGHT(CONVERT(VARCHAR(8),GETDATE(),112),6) + [DBO].[USER_NumZero] (@R_VALUE02 + 1,3)
		
		----�а��� ������ȣ ����
		UPDATE OJDT SET U_MNum = @R_VALUE02
		WHERE TransId IN (SELECT TransId FROM OVPM WHERE U_CTRNUM = @R_VALUE01)
	END
	
	--	�Աݽ� ��õ������ A/R����, A/R�뺯�޸��� ��Ʈ�� üũ
	IF @OBJECT_TYPE IN ('24') AND @transaction_type IN ('A') 
	BEGIN
		SET	@R_VALUE01	= ''
		SET	@R_VALUE02	= ''
		SET @V_BASETYPE = ''

		-- �Աݿ� ����ó�ڵ� üũ
		SELECT @R_VALUE01=U_EARDCODE, @R_VALUE02 = U_PAYTYPE FROM ORCT WHERE DOCENTRY = @list_of_cols_val_tab_del
		BEGIN
			IF ISNULL(@R_VALUE01,'') = ''
			BEGIN
				--B1���� ���� ��ǰ�� �Ա�ó���� ����ó�ڵ带 �ʼ� �Է��ϵ��� �ؾߵ�.
				SET @ERROR = 1
				SET @ERROR_MESSAGE = '[NT]�Աݹ��� �ۼ��� ����ó�ڵ带 �Է� �ϼž� �մϴ�.'
			END
			IF ISNULL(@R_VALUE02,'') = ''
			BEGIN
				--B1���� ���� ��ǰ�� �Ա�ó���� ����ó�ڵ带 �ʼ� �Է��ϵ��� �ؾߵ�.
				SET @ERROR = 1
				SET @ERROR_MESSAGE = '[NT]�Աݹ��� �ۼ��� ���������� �����ϼž� �մϴ�.'
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
			SET @ERROR_MESSAGE = '[NT]�Աݹ��� �ۼ��� �� ����ó���� �Ա��� �ϼž� �մϴ�.'	
		END
		ELSE
		BEGIN
			--	�а��� ����ó ������Ʈ
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
			
			-- �а��� ����ó�ڵ�
			UPDATE JDT1 SET U_EARDCODE=@V_BASETYPE WHERE TRANSID=@V_BASEREF
		END
		----	������ȣ üũ(10 OR 20)
		--IF EXISTS(SELECT B.REFNUM FROM ORCT A INNER JOIN OBOE B ON A.BOEABS=B.BOEKEY
		--			WHERE A.DOCENTRY=@list_of_cols_val_tab_del
		--			AND A.BOESUM>0
		--			AND NOT ISNULL(LEN(B.REFNUM),0) IN (10,20))
		--BEGIN
		--	SET @ERROR = 1
		--	SET @ERROR_MESSAGE = '[NT]������ȣ�� �ڸ����� 10 or 20 �̿��� �մϴ�.'
		--END
	END

	--	���� ���ݽ� ��õ������ A/R����, A/R�뺯�޸��� ����ó üũ
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
			--	�а��� ����ó ������Ʈ
			UPDATE JDT1 SET U_EARDCODE=@V_BASETYPE 
			WHERE TRANSID=@V_BASEREF AND REF1=@V_BASELINE AND TRANSTYPE='182' AND CREATEDBY=@list_of_cols_val_tab_del	

			FETCH NEXT FROM S60051 INTO	@V_BASETYPE, @V_BASEREF, @V_BASELINE
		END
		CLOSE S60051
		DEALLOCATE S60051
	END
	--	����Ͻ���Ʈ�� üũ(�̸�, �����ڵ�)
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
			SET @ERROR_MESSAGE = '[NT]ȸ������ �����ڵ带 �Է��� �ּ���.'
		END
		ELSE IF @R_VALUE01=''
		BEGIN
			SET @ERROR = 1
			SET @ERROR_MESSAGE = '[NT]�ŷ�ó���� �Է��� �ּ���.'
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
					SET @ERROR_MESSAGE = '[NT]�а�-�ڽ�Ʈ���͸� �Է��� �ּ���.'
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
							SET @ERROR_MESSAGE = '[NT]�а�-���Ͱŷ�ó�� �Է��� �ּ���.'
						END
					END
				END
		END
		
		--���������ڵ带 ������ ������������ �Է�
		UPDATE A SET U_AcctName = B.AcctName
		FROM JDT1 A INNER JOIN OACT B ON A.Account = B.AcctCode
		WHERE TRANSID = @list_of_cols_val_tab_del
		
		--BP�ŷ�ó�� 9�� �����ϴ� �ŷ�ó�� ����ī�� ��ȣ�� �Է�
		UPDATE A SET A.REF3LINE = B.CNTCTPRSN
		FROM JDT1 A INNER JOIN OCRD B ON A.ShortName = B.CARDCODE AND LEFT(B.CARDCODE,1) = '9'
		WHERE Transid=@list_of_cols_val_tab_del
	END
END 

SELECT @ERROR, @ERROR_MESSAGE 
