USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_CO080_01]    Script Date: 11/04/2010 19:58:24 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO

/****************************************************************************************************************/
/*  Module         : CO																							*/
/*  Description    : 코스트센터비용집계 [PS_CO080]                                                                     */
/*  Create Date    : 2010.11.04                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_CO080_01]
--CREATE PROCEDURE [dbo].[PS_CO080_01]
 (
 
	@FRefDate   	As DateTime,             --Date타입
    @TRefDate       As DateTime
)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON

    -- Insert statements for procedure here
    
Select C.PrcCode, d.OcrName, B.U_CECode, B.U_CEName, B.U_Class, ((A.Debit - A.Credit) * C.PrcAmount / C.OcrTotal) as Bit 
  From JDT1 as A inner join [@PS_CO010L] as B 
       on A.Account = B.U_CECode 
          inner join OCR1 as C 
       on A.ProfitCode = C.OcrCode 
          inner join OOCR as D 
       on C.PrcCode = D.OcrCode 
 Where B.U_Category = '10' 
   AND A.RefDate between @FRefDate And @TRefDate 
   
   union all 
   
Select C.PrcCode, d.OcrName, B.U_CECode, B.U_CEName, B.U_Class, ((A.Debit - A.Credit) * C.PrcAmount / C.OcrTotal) as Bit 
  From JDT1 as A inner join [@PS_CO010L] as B 
       on A.Account = B.U_CECode 
          inner join MDR1 as C 
       on A.ProfitCode = C.OcrCode 
          inner join OOCR as D 
       on C.PrcCode = D.OcrCode 
Where B.U_Category = '10' 
  AND A.RefDate between @FRefDate And @TRefDate     
END

--exec [PS_CO080_01] '20101001','20101031'