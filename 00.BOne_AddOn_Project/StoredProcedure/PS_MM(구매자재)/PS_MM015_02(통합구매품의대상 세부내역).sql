USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM015_02]    Script Date: 02/21/2011 19:13:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : ���հŷ�ó ǰ�Ǽ� ��� ���γ���                                                                 */
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
	   b.U_Qty As PQty, --û������
	   b.U_Weight As Weight, --û���߷�,
	   U_E_MENGE As MENGE, --'ǰ�Ǽ���',
	   b.U_E_NETWR AS NETWR, --'�ݾ�',
	   a.U_CntcCode As CntcCode,  --���
	   b.U_E_BEDAT As BEDAT, --ǰ������ 
	   b.U_E_EINDT As EINDT --��ǰ����
From [@PS_MM010H] a Inner Join [@PS_MM010L] b On a.DocEntry = b.DocEntry Inner Join [OCRD] c On b.U_E_LIFNR = c.VatRegNum
    And Isnull(b.U_E_LIFNR,'') <> ''
  Where a.U_PQType = '20'
    And c.CardType = 'S'
    And b.U_E_EBELN = @EBELN
    and Not Exists (select * from [@PS_MM030H] H, [@PS_MM030L] L Where H.DocEntry = L.DocEntry and H.Canceled = 'N' and b.DocEntry = L.U_PQDocNum and b.U_LineNum = L.U_PQLinNum)
  Order by a.U_BPLId, c.CardCode, a.DocNum, b.U_LineNum

--EXEC [PS_MM015_02] '4500180992'



