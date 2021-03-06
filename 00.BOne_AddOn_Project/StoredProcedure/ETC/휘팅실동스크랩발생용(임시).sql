USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_test_01]    Script Date: 03/11/2011 13:14:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 제공품 내역서 집계       																	*/
/*  Create Date    : 2011.01.10                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/


ALTER PROC [dbo].[PS_test_01]
--Create  PROC [dbo].[PS_test_01]
(
	@DocDatef  NvarChar(30),
	@DocDatet  NvarChar(30)
)
AS

--Create Table #Z_test ( CpCode Nvarchar(20),
--						CpName NvarChar(100),
--                        ItemCode NVARCHAR(max) collate Korean_Wansung_Unicode_CI_AS,
--                        OrdNum	 NVARCHAR(max) collate Korean_Wansung_Unicode_CI_AS,	                        
--                        ItemName NVARCHAR(60),	        
--				        CNT  Numeric(19,0),
--				        DocDate NvarChar(8),
--				        sequence NVARCHAR,
--						yqty Numeric(19,2),  -- 중량
--						BPLId NVARCHAR(1)
--					  )

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
					   BCpUnWt Numeric(10,5) )
					   
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
  and convert(char(8),H.U_DocDate,112) like '2011%'
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

Update [@ps_pp040L]
  set U_ScrapWt = round((([@ps_pp040L].U_Pqty * b.BCpUnWt) - ([@ps_pp040L].U_Pqty * b.CpUnWt)) /1000,0)
  from [@ps_pp040H] a,
	   #Z_test b
 where a.DocEntry = [@ps_pp040L].DocEntry
   and [@ps_pp040L].DocEntry = b.DocEntry
   and [@ps_pp040L].LineId = b.LineId
   and a.U_OrdGbn = '101'
   and a.U_OrdType in ('10','30','50','60')

select A.ItemCode,
	   M.ItemName,
	   A.OrdNum,
	   A.Cpcode,
	   A.CpName,
	   A.Sequence,
	   A.BCpCode,
	   A.Qty,
	   Wgt = Round(A.Wgt,0),
	   wt  = Round(A.Qty * Isnull(A.CpUnWt,0) / 1000,0),
	   bwgt = (Case When Sequence = 1 Then Round(A.QTy * Isnull(M.U_UseMKg,0) / 1000,0)
				Else Round(A.Qty * Isnull(A.BCpUnWt,0) / 1000,0)
			  End ),
	   Scrap = Case When A.CpCode = 'CP30112' Or A.CpCode = 'CP30114' Then 0 
				Else (Case When Sequence = 1 Then Round(A.QTy * Isnull(M.U_UseMKg,0) / 1000,0)
						Else Round(A.Qty * Isnull(A.BCpUnWt,0) / 1000,0)
						End) - Round(A.Wgt,0)
				End,
	   Scrap1 = Case When A.CpCode = 'CP30112' Or A.CpCode = 'CP30114' Then 0 
				Else (Case When Sequence = 1 Then Round(A.QTy * Isnull(M.U_UseMKg,0) / 1000,0)
						Else Round(A.Qty * Isnull(A.BCpUnWt,0) / 1000,0)
						End) - Round(A.Qty * Isnull(A.CpUnWt,0) / 1000,0)
				End
	    from #Z_test A,
			 OITM M
 Where A.ItemCode = M.ItemCode
   and A.ItemCode = '101080026'
 --) G
 --Where G.CpCode <> 'CP30112' And G.CpCode <> 'CP30114'
end

--EXEC [dbo].[PS_test_01] '20110101', '20110131'

--EXEC [dbo].[PS_PP685_01] '2', '20101201', '104010006', 'CP50101'

--EXEC [PS_PP685_01] '2','20101231','104010041','%'






----select * from [@PS_PP040L]

----select * from [@PS_PP030L]
