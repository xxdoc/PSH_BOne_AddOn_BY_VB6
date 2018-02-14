USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM015_01]    Script Date: 02/21/2011 19:13:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : 통합거래처 품의서 대상조회                                                                 */
/*  Create Date    : 2011.02.21                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--ALTER PROC [dbo].[PS_MM015_01]
Create  PROC [dbo].[PS_MM015_01]
(
	@BPLId			As Char(1)
)
AS

SET NOCOUNT ON

Select G.BEDAT,		--'품의일자',
	   G.EBELN,		--'풍산품의번호', 
	   G.Lifnr,		--  '사업자번호',
	   (select o.CardCode From OCRD o Where o.VatRegNum = G.Lifnr and o.CardType = 'S') As CardCode, --'거래처코드' ,
       (select o.CardName From OCRD o Where o.VatRegNum = G.Lifnr  and o.CardType = 'S') As CardName, --'거래처명',
	   G.Cnt -- '건수'
From (
SELECT a.U_E_BEDAT As BEDAT,
	   a.U_E_EBELN As EBELN,
       a.U_E_Lifnr As Lifnr,
	   Cnt = Count(a.U_E_EBELN)
FROM [@PS_MM010H] H, [@PS_MM010L] a
WHERE H.DocEntry = a.DocEntry
  and Isnull(a.U_E_EBELN,'') <> ''
  and H.Canceled = 'N'
  And Not Exists (select L.DocEntry from [@PS_MM030H] H, [@PS_MM030L] L Where H.DocEntry = L.DocEntry and H.Canceled = 'N' and a.DocEntry = L.U_PQDocNum and a.U_LineNum = L.U_PQLinNum)
  And H.U_BPLId = @BPLId
  Group by a.U_E_BEDAT, a.U_E_EBELN, a.U_E_Lifnr
 ) G
 Order By G.EBELN
 
 --EXEC [PS_MM015_01] '1'