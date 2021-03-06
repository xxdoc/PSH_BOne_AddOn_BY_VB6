USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[SBO_SP_PostTransactionNotice]    Script Date: 11/29/2010 17:59:09 ******/
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

IF @object_type IN ('13','18','24','46') AND @transaction_type IN('A','U') BEGIN
    UPDATE a SET U_AcctName = b.AcctName
      FROM JDT1 a
      JOIN OACT b on b.AcctCode = a.Account
     WHERE a.TransType = @object_type
       AND a.CreatedBy = @list_of_cols_val_tab_del
END

IF @object_type = '21' AND @transaction_type IN('A','U') BEGIN
    EXECUTE PS_Z_RETU_GR @list_of_cols_val_tab_del, @error OUTPUT , @error_message OUTPUT
END

IF @object_type = '16' AND @transaction_type IN('A','U') BEGIN
    EXECUTE PS_Z_RETU_GI @list_of_cols_val_tab_del, @error OUTPUT , @error_message OUTPUT
END

--------------------------------------------------------------------------------------------------------------------------------

-- Select the return values
select @error, @error_message

end