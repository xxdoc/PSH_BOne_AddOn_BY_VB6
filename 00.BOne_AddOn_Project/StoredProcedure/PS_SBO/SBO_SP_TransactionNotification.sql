--USE [PSHDB_TEST]
GO
/****** Object:  StoredProcedure [dbo].[SBO_SP_TransactionNotification]    Script Date: 09/15/2010 09:09:42 ******/
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
-----------마케팅문서관리------------------------------------------------------------------
IF @object_type IN('23','17','15','16','13','22','20','21','18','19') AND @transaction_type IN('A','U')
BEGIN
	IF @object_type = '23'
		UPDATE [QUT1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '23'
		UPDATE [OQUT] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '17'
		UPDATE [RDR1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '17'
		UPDATE [ORDR] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '15'
		UPDATE [DLN1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '15'
		UPDATE [ODLN] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '16'
		UPDATE [RDN1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '16'
		UPDATE [ORDN] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '13'
		UPDATE [RIN1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '13'
		UPDATE [ORIN] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '22'
		UPDATE [POR1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '22'
		UPDATE [OPOR] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '20'
		UPDATE [PDN1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '20'
		UPDATE [OPDN] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '21'
		UPDATE [RPD1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '21'
		UPDATE [ORPD] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '18'
		UPDATE [PCH1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '18'
		UPDATE [OPCH] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '19'
		UPDATE [RPC1] SET U_LineNum = LineNum WHERE DocEntry = @list_of_cols_val_tab_del
	IF @object_type = '19'
		UPDATE [ORPC] SET DocCur = (CASE WHEN CurSource IN('S','L') THEN 'KRW' ELSE DocCur END) WHERE DocEntry = @list_of_cols_val_tab_del
END
-----------사용자문서관리------------------------------------------------------------------
IF @object_type IN('PS_SD030','PS_SD040','PS_PP030','PS_PP040','PS_PP070','PS_PP080') AND @transaction_type IN('A','U')
BEGIN
	IF @object_type = 'PS_SD030'
	BEGIN
		UPDATE [@PS_SD030L] SET U_LineId = LineId WHERE DocEntry = @list_of_cols_val_tab_del	
	END
	ELSE IF @object_type = 'PS_SD040'
	BEGIN
		UPDATE [@PS_SD040L] SET U_LineId = LineId WHERE DocEntry = @list_of_cols_val_tab_del
	END
	ELSE IF @object_type = 'PS_PP030'
	BEGIN
		UPDATE [@PS_PP030L] SET U_LineId = LineId WHERE DocEntry = @list_of_cols_val_tab_del
		UPDATE [@PS_PP030M] SET U_LineId = LineId WHERE DocEntry = @list_of_cols_val_tab_del
	END
	ELSE IF @object_type = 'PS_PP040'
	BEGIN
		UPDATE [@PS_PP040L] SET U_LineId = LineId WHERE DocEntry = @list_of_cols_val_tab_del
		UPDATE [@PS_PP040M] SET U_LineId = LineId WHERE DocEntry = @list_of_cols_val_tab_del
		UPDATE [@PS_PP040N] SET U_LineId = LineId WHERE DocEntry = @list_of_cols_val_tab_del
	END
	ELSE IF @object_type = 'PS_PP070'
	BEGIN
		UPDATE [@PS_PP070L] SET U_LineId = LineId WHERE DocEntry = @list_of_cols_val_tab_del
	END
	ELSE IF @object_type = 'PS_PP080'
	BEGIN
		UPDATE [@PS_PP080L] SET U_LineId = LineId WHERE DocEntry = @list_of_cols_val_tab_del
	END
	
END
-----------사용자로그이력관리--------------------------------------------------------------
DECLARE @usercode NVARCHAR(8)
DECLARE @CurrentLoginDate DATETIME
DECLARE @CurrentLoginTime NVARCHAR(8)
DECLARE @LastLoginDate DATETIME
DECLARE @LastLoginTime NVARCHAR (8)
DECLARE @Code NVARCHAR (8) 
DECLARE @Name NVARCHAR (30)
IF @object_type = '12' and @transaction_type = 'U' --사용자는 로긴할때마다 최종로긴데이터를 업뎃한다. 그것을 이용한다.
BEGIN
	SET @usercode = (SELECT user_code FROM [OUSR] WHERE INTERNAL_K = @list_of_cols_val_tab_del) 
	SET @CurrentLoginDate = GETDATE()							--현재로긴일자
	SET @CurrentLoginTime = CONVERT(VARCHAR(8), GETDATE(), 108) --현재로긴시간
	--그냥최종로긴시간의 max값 가져오면 잘못된 max값을가져옴 오늘날짜의 max값을가져와야함
	--SET @LastLoginTime = (SELECT MAX(U_LoginTime) FROM [@USER_DATA] WHERE U_UserCode = @usercode and Convert(varchar(8),U_LoginDate,112)=Convert(varchar(8),@CurrentLoginDate,112))	--최종로긴시간
	--최종로긴시간의 MAX값은 U_LoginDate 값에서 형변환해서 구해온다.
	SET @LastLoginTime = (SELECT CONVERT(VARCHAR(100),MAX(U_LoginDate),108) FROM [@USER_DATA] WHERE U_UserCode = @usercode)	--최종로긴시간
	SET @LastLoginDate=(SELECT MAX(U_LoginDate) FROM [@USER_DATA] WHERE U_UserCode = @usercode)		--최종로긴일자
	--Code와 Name는 그냥 채번하는식으로 입력
	IF (Select Max(Convert(int,Code)) from [@USER_DATA]) IS NULL
	BEGIN
		SET @Code = 1
		SET @Name = 1
	END
	ELSE
	BEGIN
		SET @Code = (Select Max(Convert(int,Code)) from [@USER_DATA]) + 1 --Code값을int로 변환시키지않으면 10이후부터 채번안됨주의
		SET @Name = (Select Max(Convert(int,Name)) from [@USER_DATA]) + 1 
	END 

	IF DATEDIFF(s, @LastLoginTime, @CurrentLoginTime) > 60  OR @LastLoginTime IS NULL 
		--오늘의 마지막로긴시간과 1분정도차이로 저장
	BEGIN 
		INSERT INTO [@USER_DATA] (Code, Name, U_UserCode, U_LoginDate, U_LoginTime,U_SPID) 
		VALUES (@Code, @Name, @usercode, @CurrentLoginDate, @CurrentLoginTime,@@SPID)
	END
		--어제의 로긴시간과 오늘의 첫로긴시간과 비교하면 -발생되는데 이때도 저장되어야하기에 아래조건추가
	ELSE IF DATEDIFF(day, @LastLoginDate, @CurrentLoginDate) > 0   --0보다크면 지날날이다. 오늘과어제 임을구분
			AND DATEDIFF(s, @LastLoginTime, @CurrentLoginTime) < 0 --어제마지막시간 과 오늘 어제보다빠른시간과비교시 -차이발생 에도 저장가능
	BEGIN 
		INSERT INTO [@USER_DATA] (Code, Name, U_UserCode, U_LoginDate, U_LoginTime,U_SPID) 
		VALUES (@Code, @Name, @usercode, @CurrentLoginDate, @CurrentLoginTime,@@SPID)
	END
END
-------------------------------------------------------------------------------------------
select @error, @error_message

end