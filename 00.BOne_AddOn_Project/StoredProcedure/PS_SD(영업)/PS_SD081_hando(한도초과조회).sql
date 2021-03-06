USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_SD081_hando]    Script Date: 03/17/2011 09:27:24 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****************************************************************************************************************/
/*  Module         : SD																							*/
/*  Description    : 거래처별 여신한도 조회 [PS_SD081]                                                          */
/*  Create Date    : 2011.03.17                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : N.G.Y																			*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--ALTER PROC [dbo].[PS_SD081_hando]
Create  PROC [dbo].[PS_SD081_hando]
(	
	@BPLId			As Nvarchar(10),
	@CardCode		As Nvarchar(20),
	@DocDate		As Nvarchar(8)
)
AS
SET NOCOUNT ON
--BEGIN ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
--Declare @Balance As Numeric(19,6)
--Select Sum(Debit - Credit) from JDT1 where ShortName = '12494' And Account = '11104010'

select G.CardCode,
		G.CardName,
		G.U_CreditP, --여신담보
		G.U_MiSuP, -- 미수
		G.U_arfamt, -- 채권
		G.U_budo, -- 부도
		G.U_MiSuP + G.U_arfamt + G.U_budo As misutot, --채권계
		G.U_CreditP - G.U_MiSuP + G.U_arfamt + G.U_budo As U_Balance, 
		G.U_OutPreP + G.U_RequestP As reqamt, --출고요청금액 + 출고승인금액
		(G.U_MiSuP + G.U_arfamt + G.U_budo + G.U_OutPreP + G.U_RequestP) - G.U_CreditP As overamt
		
		
from (
Select O.CardCode,
	   O.CardName,
	   O.CreditLine As U_CreditP, --담보, 여신
	   Isnull((select SUM(b.Debit) - sum(b.Credit)
		from  [ZMDC_JDT1] a inner join [JDT1] b on a.TransId=b.TransId and a.Line_Id=b.Line_Id
							inner join [OACT] c on a.AcctCode=c.AcctCode
		where RefDate <= @DocDate
		and   a.AcctCode = '11104010'
		and b.ShortName = O.CardCode ),0) + O.DNotesBal As U_MiSuP, --미수금
		
		Isnull((select SUM(b.Debit) - sum(b.Credit)
		from  [ZMDC_JDT1] a inner join [JDT1] b on a.TransId=b.TransId and a.Line_Id=b.Line_Id
							inner join [OACT] c on a.AcctCode=c.AcctCode
		where RefDate <= @DocDate
		and   a.AcctCode = '11104060'
		and b.ShortName = O.CardCode  ),0) As U_arfamt, --채권
		
		isnull((select SUM(b.Debit) - sum(b.Credit)
		from  [ZMDC_JDT1] a inner join [JDT1] b on a.TransId=b.TransId and a.Line_Id=b.Line_Id
							inner join [OACT] c on a.AcctCode=c.AcctCode
		where RefDate <= @DocDate
		and   a.AcctCode = '11104070'
		and b.ShortName = O.CardCode  ),0) As U_budo, --부도어음
		   
	   Round((SD030.LinTotal * 1.1),0) As U_OutPreP, --출고요청금액
	   Isnull(( Select	Sum(b.U_RequestP)
			From [@PS_SD080H] a 
			     Inner Join [@PS_SD080L] b On a.DocEntry = b.DocEntry
			Where b.U_CardCode = O.CardCode	
			  And a.Status = 'O'
			  And	a.U_OkYN = 'Y'
			  And	a.U_DocDate = SD030.DocDate ),0) As U_RequestP --한도초과신청금액
from (
select SD030H.U_CardCode As CardCode,
	   SD030H.U_CardName As CardName,
	   SD030H.U_DocDate As DocDate,
	   LinTotal = Sum(SD030L.U_LinTotal)
from [@PS_SD030H] SD030H,
	 [@PS_SD030L] SD030L
where SD030H.DocEntry = SD030L.DocEntry
  and SD030H.U_DocDate = @DocDate 
  and SD030H.U_CardCode like @CardCode + '%'--'11911'
  And SD030H.U_BPLId = @BPLId
  Group by SD030H.U_CardCode, SD030H.U_CardName, SD030H.U_DocDate
  ) SD030,
   [OCRD] O 
   where SD030.CardCode = O.CardCode
	) G
 Where (G.U_MiSuP + G.U_arfamt + G.U_budo + G.U_OutPreP + G.U_RequestP ) > (G.U_CreditP)

-- EXEC [PS_SD081_hando] '4','%','20110316'

