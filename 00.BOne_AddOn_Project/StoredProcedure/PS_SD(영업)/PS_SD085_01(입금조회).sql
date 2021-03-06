USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_SD085_01]    Script Date: 03/24/2011 08:21:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****************************************************************************************************************/
/*  Module         : SD																							*/
/*  Description    : 입금조회 [PS_SD085_01]			                                                            */
/*  Create Date    : 2011.03.24                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : N.G.Y																			*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_SD085_01]
--Create  PROC [dbo].[PS_SD085_01]
(	
	@BPLId			As Nvarchar(10),
	@DocDateFr		As Nvarchar(8),
	@DocDateTo		As Nvarchar(8),
	@CardCode		As Nvarchar(10)
)
AS
SET NOCOUNT ON

Select t1.RefDate As DocDate,
	   t1.TransId,
	   t2.Ref1,
	   t3.CardCode,
	   t3.CardName,
	   t2.LineMemo,
	   t2.Account,
	   Amt = Charindex(t2.Account,'11104010') * Credit,
	   Refamt = Charindex(t2.Account,'11104060') * Debit,
	   RefNum = Isnull((select b.RefNum
				  from ORCT a,
					   OBOE b
				where a.DocEntry = b.PmntNum
				  and a.TransId = t2.TransId
				  and b.BoeAcct = t2.Account),'')
  from OJDT t1,
 	   JDT1 t2,
		OCRD t3
	     where t1.TransId = t2.TransId
	      And t2.ShortName = t3.CardCode
	      And t2.Account in ('11104010', '11104060')
	      And t1.RefDate between @DocDateFr and @DocDateTo
	      And t1.U_BPLId = @BPLId
	      And t1.TransType in ('24')
	      And t3.CardCode like @CardCode + '%'

-- EXEC [PS_SD085_01] '4','20110101','20110131'

