USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_test_01]    Script Date: 03/23/2011 19:56:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 휘팅 스크랩 발생량 계산   																	*/
/*  Create Date    : 2011.01.10                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/


ALTER PROC [dbo].[PS_PP040_Scrap]
--Create  PROC [dbo].[PS_PP040_Scrap]
(
	@DocDatef  NvarChar(30),
	@DocDatet  NvarChar(30)
)
AS


Create Table #Z_test ( DocEntry nvarchar(10),
					   LineId nvarchar(10),
					   ItemCode NVarchar(Max) collate Korean_Wansung_Unicode_CI_AS ,
					   OrdNum Nvarchar(100),
					   CpCode Nvarchar(Max) collate Korean_Wansung_Unicode_CI_AS,
					   CpName Nvarchar(100),
					   Sequence Nvarchar(10),
					   pp030hno Nvarchar(10),
					   Qty Numeric(12,0),
					   Wgt Numeric(12,3),
					   CpUnwt Numeric(10,5),
					   BCpCode Nvarchar(Max) collate Korean_Wansung_Unicode_CI_AS,
					   BCpUnWt Numeric(10,5),
					   Scrap   Numeric(12,3) )
					   
Create Table #Z_test_1 (ItemCode NVARCHAR(max) collate Korean_Wansung_Unicode_CI_AS,
						 ItemName NvarChar(100),
						 OrdNum	NVARCHAR(max) collate Korean_Wansung_Unicode_CI_AS,
						 DocDate date,
						 CpCode NVARCHAR(10)  collate Korean_Wansung_Unicode_CI_AS,
						 CpName NvarChar(100),
						 yqty Numeric(12,2) )
						
begin

Insert Into #Z_test ( DocEntry,  LineId, ItemCode, OrdNum, CpCode, CpName, pp030hno, Sequence, Qty, Wgt )
select L.DocEntry, L.LineId, L.U_ItemCode, L.U_OrdNum, L.U_CpCode, L.U_CpName, L.U_PP030HNo, L.U_Sequence, L.U_YQty, L.U_YWeight
 from [@PS_PP040H] H,
	  [@PS_PP040L] L
where H.DocEntry = L.DocEntry
  and H.U_DocDate Between @DocDatef and @DocDatet
  and H.U_OrdGbn = '101'
  
 Update #Z_test
    set BCPCode = M.U_CpCode
   From [@PS_PP030H] H,
		[@PS_PP030M] M
 Where H.DocEntry = M.DocEntry
  And M.DocEntry = #Z_test.pp030hno
  And M.U_Sequence = (#Z_test.Sequence - 1)
  
--이전중량 첫공정 투입단중
 Update #Z_test
    set BCpUnWt = o.U_useMkg
   From [@PS_PP030H] H,
		[@PS_PP030M] M,
		OITM o
 Where H.DocEntry = M.DocEntry
  And H.U_ItemCode = o.ItemCode
  And M.DocEntry = #Z_test.pp030hno
  And M.U_Sequence = 1
  
---이전공정단중
Update #Z_test
   set BCpUnWt = H.U_CpUnWt
  From [@PS_PP004H] H
 Where #Z_test.ItemCode = H.U_ItemCode
   and #Z_test.BCpCode = H.U_CpCode

Update #Z_test
   set CpUnWt = H.U_CpUnWt
  From [@PS_PP004H] H
 Where #Z_test.ItemCode = H.U_ItemCode
   and #Z_test.CpCode = H.U_CpCode
   
 --바렐과 포장은제품단중으로
   
 Update #Z_test
    set CpUnWt = m.U_UnWeight
   from OITM m
  where #Z_test.ItemCode = m.ItemCode
    and Cpcode in ('CP30112', 'CP30114')

--//실동 테이블 Scrap Reset
Update [@PS_PP040L] Set U_ScrapWt = 0
from [@PS_PP040H] H
where H.DocEntry = [@PS_PP040L].DocEntry
  and H.U_DocDate Between @DocDatef and @DocDatet
  and H.U_OrdGbn = '101'


Update [@ps_pp040L]
  set U_ScrapWt = round((([@ps_pp040L].U_Pqty * b.BCpUnWt) - ([@ps_pp040L].U_Pqty * b.CpUnWt)) /1000,0)
  from [@ps_pp040H] a,
	   #Z_test b
 where a.DocEntry = [@ps_pp040L].DocEntry
   and [@ps_pp040L].DocEntry = b.DocEntry
   and [@ps_pp040L].LineId = b.LineId
   and a.U_OrdGbn = '101'
   and a.U_OrdType in ('10','30','50','60')

Update #Z_test
   set Scrap = Case When #Z_test.CpCode = 'CP30112' Or #Z_test.CpCode = 'CP30114' Then 0 
				Else (Case When #Z_test.Sequence = 1 Then Round(#Z_test.QTy * Isnull(M.U_UseMKg,0) / 1000,0)
						Else Round(#Z_test.Qty * Isnull(#Z_test.BCpUnWt,0) / 1000,0)
						End) - Round(#Z_test.Qty * Isnull(#Z_test.CpUnWt,0) / 1000,0)
				End
  From OITM M
 Where #Z_test.ItemCode = M.ItemCode



select A.DocEntry,
	   A.LineId,
	   A.ItemCode,
	   M.ItemName,
	   A.OrdNum,
	   A.Cpcode,
	   A.CpName,
	   A.Sequence,
	   A.BCpCode,
	   A.Qty,
	   A.Scrap
  from #Z_test A,
  	   OITM M
 Where A.ItemCode = M.ItemCode
   
 --) G
 --Where G.CpCode <> 'CP30112' And G.CpCode <> 'CP30114'
end

--EXEC [dbo].[PS_PP040_Scrap] '20110201', '20110228'