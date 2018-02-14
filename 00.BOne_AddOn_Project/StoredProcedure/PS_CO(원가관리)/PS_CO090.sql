USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_CO090]    Script Date: 11/16/2010 12:27:52 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Lee Byong Gak
-- Create date: 2010.11.01
-- Description:	통계주요지표 값 입력
-- =============================================
create PROCEDURE [dbo].[PS_CO090]
--ALTER PROCEDURE [dbo].[PS_CO090]
    @iFrDate    AS DateTime,
    @iToDate    AS DateTime
AS
BEGIN


select U_ATCode, U_ATName, U_Total, U_Unit, U_DataCnt 
  From [@PS_CO090L] AS A INNER JOIN [@PS_CO090H] AS B
        ON A.Code = B.Code
 WHERE  B.U_ClsPrd BETWEEN @iFrDate AND @iToDate

end

exec [PS_CO900] '20101001','20101031'