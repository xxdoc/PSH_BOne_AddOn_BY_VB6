USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM015_02]    Script Date: 02/21/2011 19:13:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : 통합거래처 품의서 대상 세부내역                                                                 */
/*  Create Date    : 2011.02.21                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--ALTER PROC [dbo].[PS_MM015_02]
Create  PROC [dbo].[PS_MM015_02]
(
	@EBELN			As NvarChar(10)
)
AS

SET NOCOUNT ON
Select a.DocNum,
	   b.U_LineNum As LineNum,
	   c.CardCode,
	   a.U_BPLId As BPLId,
	   a.U_Purchase As Purchase,
	   a.U_PQType As PQType,
	   b.U_ItemCode As ItemCode,
	   b.U_Qty As PQty, --청구수량
	   b.U_Weight As Weight, --청구중량,
	   U_E_MENGE As MENGE, --'품의수량',
	   b.U_E_NETWR AS NETWR, --'금액',
	   a.U_CntcCode As CntcCode,  --사번
	   b.U_E_BEDAT As BEDAT, --품의일자 
	   b.U_E_EINDT As EINDT --납품일자
From [@PS_MM010H] a Inner Join [@PS_MM010L] b On a.DocEntry = b.DocEntry Inner Join [OCRD] c On b.U_E_LIFNR = c.VatRegNum
    And Isnull(b.U_E_LIFNR,'') <> ''
  Where a.U_PQType = '20'
    And c.CardType = 'S'
    And b.U_E_EBELN = @EBELN
    and Not Exists (select * from [@PS_MM030H] H, [@PS_MM030L] L Where H.DocEntry = L.DocEntry and H.Canceled = 'N' and b.DocEntry = L.U_PQDocNum and b.U_LineNum = L.U_PQLinNum)
  Order by a.U_BPLId, c.CardCode, a.DocNum, b.U_LineNum

--EXEC [PS_MM015_02] '4500180992'



