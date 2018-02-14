IF OBJECT_ID('SBO_SP_MM_TransactionNotification') IS NOT NULL
   DROP PROCEDURE SBO_SP_MM_TransactionNotification
GO
/********************************************************************************************************************************************************                                     
 ���ν����� : SBO_SP_MM_TransactionNotification
 ��      �� : �ǸŰ���, ���Ű��� TransactionNotification
 ��  ��  �� : 
 ��      �� : 
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
			SET @error = -1										-- Error Code ����
			SET @error_message = N'�ڵ带 �Է��� �ּ���.'		-- Error Message ����
			RETURN												-- Error �߻��� RETURN�� ����Ͽ� SP EXIT
		END
	END
END
****************************************** Sample ********************************************************/
--------------------------------------------------------------------------------------------------------------------------------
--	ADD	YOUR	CODE	HERE
--------------------------------------------------------------------------------------------------------------------------------

/*** OIPF : ���Ű��� - ���Ժδ��� ***/
IF @object_type = '69'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* ����� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OIPF WHERE DocEntry=@list_of_cols_val_tab_del AND U_PWC_BPLId IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'������� �Է��� �ּ���.'
			RETURN
		END
	END
END

/*** OPCH : ���Ű��� - A/P������� ***/
IF @object_type = '18'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* AP��������ϰ�� BL��ȣ �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OPCH WHERE DocEntry=@list_of_cols_val_tab_del AND U_POTYPE = '2' AND UPDINVNT = 'O' AND U_BL_NO IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'BL��ȣ�� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END
		--* AP��������ϰ��������� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OPCH WHERE DocEntry=@list_of_cols_val_tab_del AND U_POTYPE = '2' AND UPDINVNT = 'O' AND U_SETTLE IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'���������� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END		
	END
	
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		DECLARE   @OPCH_nvcAcctCode	NVARCHAR(15)	-- �ʼ� üũ�� AcctCode (�ð��� ����, ���� ��� ����)
		
		--* A/P ���� - ���� - �ð��� ����, ���� ��� ���� �ʼ� ���� üũ
		IF (SELECT DocType FROM OPCH WHERE DocEntry=@list_of_cols_val_tab_del) = 'S'
		BEGIN
			-- �ð��� ���� �ʼ� ���� ���� ����
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
				SET @error_message = N'�ش� ������ ��� �ð��� ������ �ʼ��Դϴ�. [G/L ���� : ' + @OPCH_nvcAcctCode + ']'
				RETURN
			END
			
			-- ���� ��� ���� �ʼ� ���� ���� ����
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
				SET @error_message = N'�ش� ������ ��� ���� ��� ������ �ʼ��Դϴ�. [G/L ���� : ' + @OPCH_nvcAcctCode + ']'
				RETURN
			END
		END
	END
END

/*** ORPC : ���Ű��� - A/P�뺯�޸� ***/
IF @object_type = '19'
BEGIN
	/* �߰� ��� */
	IF @transaction_type = 'A'
	BEGIN
		DECLARE   @ORPC_nvcAcctCode	NVARCHAR(15)	-- �ʼ� üũ�� AcctCode (�ð��� ����, ���� ��� ����)

		--* A/P �뺯�޸� - ���� - �ð��� ����, ���� ��� ���� �ʼ� ���� üũ
		IF (SELECT DocType FROM ORPC WHERE DocEntry=@list_of_cols_val_tab_del) = 'S'
		BEGIN
			-- �ð��� ���� �ʼ� ���� ���� ����
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
				SET @error_message = N'�ش� ������ ��� �ð��� ������ �ʼ��Դϴ�. [G/L ���� : ' + @ORPC_nvcAcctCode + ']'
				RETURN
			END
			
			-- ���� ��� ���� �ʼ� ���� ���� ����
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
				SET @error_message = N'�ش� ������ ��� ���� ��� ������ �ʼ��Դϴ�. [G/L ���� : ' + @ORPC_nvcAcctCode + ']'
				RETURN
			END
		END
	END
END

/*** OINV : �ǸŰ��� - A/R���� ***/
IF @object_type = '13'
BEGIN
	/* �߰�, ���� ��� */
	IF @transaction_type IN ('A', 'U')
	BEGIN
		--* A/R�����ϰ�� BL��ȣ �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OINV WHERE DocEntry=@list_of_cols_val_tab_del AND U_SOTYPE = '2' AND UPDINVNT = 'I' AND U_BL_NO IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'BL��ȣ�� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END
		--* A/R�����ϰ�� �������� �ʼ� �Է� üũ
		IF EXISTS(SELECT DocEntry FROM OINV WHERE DocEntry=@list_of_cols_val_tab_del AND U_SOTYPE = '2' AND UPDINVNT = 'I' AND U_SETTLE IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'���������� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END
	END
END

/*** ORDR : �ǸŰ��� - �Ǹſ��� ***/
IF @object_type = '17'
BEGIN
	--�Ǹſ��� ������ �������(U_INGQTY)�� ���� �Ǵ°�츦 �����ϱ� ����
	IF @transaction_type IN ('A')
	BEGIN
		UPDATE [RDR1] SET U_INGQTY = 0 WHERE DocEntry = @list_of_cols_val_tab_del
	END 
END

/*** OPOR : ���Ű��� - ���ſ��� ***/
IF @object_type = '22'
BEGIN
	--���ſ��� ������ �������(U_INGQTY)�� ���� �Ǵ°�츦 �����ϱ� ����
	IF @transaction_type IN ('A')
	BEGIN
		UPDATE [POR1] SET U_INGQTY = 0 WHERE DocEntry = @list_of_cols_val_tab_del
	END 
END

/*** ODLN : �ǸŰ��� - ��ǰ ***/
IF @object_type = '15'
BEGIN
	--��ǰ ��ۺ� �ִ°�� ���-�μ��� �ʼ��� üũ�ϱ�����.
	IF @transaction_type IN ('A','U')
	BEGIN
		IF EXISTS(SELECT DocEntry FROM ODLN WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'���-�μ��� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END
		
	END 
END

/*** ORDN : �ǸŰ��� - ��ǰ ***/
IF @object_type = '16'
BEGIN
	--��ǰ ��ۺ� �ִ°�� ���-�μ��� �ʼ��� üũ�ϱ�����.
	IF @transaction_type IN ('A','U')
	BEGIN
		IF EXISTS(SELECT DocEntry FROM ORDN WHERE DocEntry = @list_of_cols_val_tab_del AND U_Transcost > 0 AND U_CarDept IS NULL)
		BEGIN
			SET @error = -1
			SET @error_message = N'���-�μ��� �Էµ��� �ʾҽ��ϴ�..'
			RETURN
		END
	END 
END




--------------------------------------------------------------------------------------------------------------------------------
--	��õ������ ���� ����ġ ������ üũ
--------------------------------------------------------------------------------------------------------------------------------
IF @object_type = '18'			--	AP����(OPCH0),�԰�PO(OPDN)
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPCH T0 
				INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN PDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN OPDN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='20' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @error = -1
		SET @ERROR_MESSAGE = N'�԰�PO�� ������� A/P������ ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		RETURN
	END
	--ELSE IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY
	--			INNER JOIN POR1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
	--			INNER JOIN OPOR T3 ON T2.DOCENTRY=T3.DOCENTRY
	--			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='22' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	--BEGIN
	--	SET @ERROR = -1
	--	SET @ERROR_MESSAGE = N'���ſ����� ������� A/P������ ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
	--	RETURN
	--END
END
ELSE IF @object_type = '19'			--	A/P�뺯�޸�
BEGIN
	--IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY
	--			INNER JOIN PCH1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
	--			INNER JOIN OPCH T3 ON T2.DOCENTRY=T3.DOCENTRY
	--			WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='18' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	--BEGIN
	--	SET @ERROR = -1
	--	SET @ERROR_MESSAGE = N'A/P������ ������� A/P�뺯�޸��� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
	--	RETURN
	--END
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN RPD1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ORPD T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='21' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'�԰��ǰ�� ������� A/P�뺯�޸��� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		RETURN
	END
END
ELSE IF @object_type = '20'			--	�԰�po
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN RPD1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ORPD T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='21' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'�԰��ǰ�� ������� �԰�PO�� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
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
		SET @ERROR_MESSAGE = N'�԰�PO�� ������� A/P��������� ����� ���� �۽��ϴ�.[�����ڿ��� ���ǹٶ�]'
		RETURN
	END
END
ELSE IF @object_type = '21'			--	�԰��ǰ
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORPD T0 INNER JOIN RPD1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN PDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN OPDN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='20' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'�԰�PO�� ������� �԰��ǰ�� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		RETURN
	END
END
ELSE IF @object_type = '13'			--	A/R����
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM OINV T0 INNER JOIN INV1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN DLN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ODLN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='15' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'��ǰ�� ������� A/R������ ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		RETURN
	END
END
ELSE IF @object_type = '14'			--	A/P�뺯�޸�
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN INV1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN OINV T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='13' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'A/R������ ������� A/R�뺯�޸��� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		RETURN
	END
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN RDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ORDN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='16' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'��ǰ�� ������� A/R�뺯�޸��� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		RETURN
	END
END
ELSE IF @object_type = '15'			--	�԰�po
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ODLN T0 INNER JOIN DLN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN RDN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ORDN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='16' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'��ǰ�� ������� ��ǰ�� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		RETURN
	END
END
ELSE IF @object_type = '16'			--	�԰��ǰ
BEGIN
	IF EXISTS(SELECT TOP 1 T0.DOCENTRY FROM ORDN T0 INNER JOIN RDN1 T1 ON T0.DOCENTRY=T1.DOCENTRY
				INNER JOIN DLN1 T2 ON T1.BASETYPE=T2.OBJTYPE AND T1.BASEENTRY=T2.DOCENTRY AND T1.BASELINE=T2.LINENUM
				INNER JOIN ODLN T3 ON T2.DOCENTRY=T3.DOCENTRY
				WHERE T0.DOCENTRY=@list_of_cols_val_tab_del AND T1.BASETYPE='15' AND MONTH(T0.DOCDATE)<>MONTH(T3.DOCDATE))
	BEGIN
		SET @ERROR = -1
		SET @ERROR_MESSAGE = N'��ǰ�� ������� ��ǰ�� ������� ��ġ���� �ʽ��ϴ�.[�����ڿ��� ���ǹٶ�]'
		RETURN
	END
END


END