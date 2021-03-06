USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP361_001]    Script Date: 05/26/2011 11:05:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 생산관리																				    */
/*  Description    : 작번별 생산진행현황-메인작번현황															*/
/*  ALTER  Date    : 2011.03.18																					*/
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_PP361_001]
ALTER     PROC [dbo].[PS_PP361_001]
(
	@SYYYYMM	as Nvarchar(7),   --작번등록년월 시작
    @EYYYYMM    as Nvarchar(7),   --작번등록년월 종료
    @BPLId		as Nvarchar(5),   --사업장
    @ItemGB		as Nvarchar(10),  --품목구분
    @Section	as Nvarchar(10),  --구분
    @ItemName   as Nvarchar(200), --품명
    @Size       as Nvarchar(200), --규격
    @OrdNum     as Nvarchar(200)  --작번
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

select G.OrdNum as OrdNum,
       G.CardName as CardName,
       G.ItemCode as ItemCode,
       G.ItemName as ItemName,
       G.Size as Size,
       G.SalUnitMsr as Unit,
       MAX(isnull(G.Quantity,0)) as Quantity,
       Sum(isnull(G.LinTotal,0)) as LinTotal,
       SUM(isnull(G.SulGe,0)) as SulGe,
       Sum(isnull(G.Jache,0)) as WorkCost,
       Sum(isnull(G.WajuGa,0)) as WajuGa,
       Sum(isnull(G.WajuJe,0)) as Wajuje,
       Sum(G.Total) as Total,
       G.Sugum  as Sugum,
       G.DocDate as DocDate,
       G.DocDueDate as DocDueDate,
       Max(G.Yalru) as Yalru,
       Max(G.YesNo) as  YesNo
  from 
(
select	Convert(Char(30),a.U_OrdNum) as OrdNum,
		d.CardName	 as CardName,
		a.U_ItemCode as ItemCode,
		Convert(char(50),a.U_ItemName) as ItemName,
		Convert(char(50),c.U_Size)     as Size,
		c.SalUnitMsr    as SalUnitMsr,
		--a.U_JakUnit  as '단위',
		e.Quantity	as Quantity,
		isnull(h.Lintotal,0) as LinTotal,
		isnull(j.WorkCost,0) as SulGe,
		isnull(g.WorkCost,0) as Jache,
		--isnull(g.WorkCost,0)) * WorkCost,
		Max((isnull(l.Wajuga,0))) as WajuGa,
		(isnull(m.Wajuje,0)) as WajuJe,
		(isnull(h.Lintotal,0) + isnull(j.WorkCost,0) + isnull(g.WorkCost,0) + 
		(isnull(l.Wajuga,0)) + (isnull(m.Wajuje,0))) AS Total,		
		d.DocDate	 as DocDate,
		Max(e.LineTotal) as Sugum,
		--e.LineTotal+d.VatSum as Sugum,
		d.DocDueDate as DocDueDate,
	    Max(f.DocDate) as Yalru,
		Case When Max(isnull(f.DocDate,'')) = ''  Then 'B'
			 When Max(isnull(f.DocDate,'')) <> '' Then 'C' end as YesNo
		--,a.U_OrdNum,f.U_OrdNum
from [@PS_PP030H] a inner join [@PS_PP030L] b on a.DocEntry=b.DocEntry
					left  join [OITM] c on a.U_ItemCode=c.ItemCode
					left  join [ORDR] d on a.U_SjNum=d.DocEntry
					left  join [RDR1] e on d.DocEntry=e.DocEntry and a.U_SjLine=e.LineNum
					left  join [@PS_PP020H] o on o.DocEntry = a.U_BaseNum	
					left  join (--**완료일자 
								select b.U_PP030HNo, b.U_PP030MNo, b.U_OrdNum,max(a.U_DocDate) DocDate, SUM(b.U_Cost01) as Cost01
								 from [@PS_PP080H] a inner join [@PS_PP080L] b on a.Docentry=b.DocEntry
								where b.U_OrdSub1 = '00'
								 group by b.U_PP030HNo, b.U_PP030MNo, b.U_OrdNum
								) f on a.DocEntry = f.U_PP030HNo and a.U_OrdNum = f.U_OrdNum
					left  join ( --**자체가공비
								select	b.U_PP030HNo, b.U_OrdNum, Sum(b.U_WorkTime * c.U_Price) as WorkCost
								from [@PS_PP040H] a inner join [@PS_PP040L] b on a.DocEntry=b.DocEntry 
													left  join [@PS_PP001L] c on b.U_CpCode=c.U_CpCode --공정단가
								where a.Canceled<>'Y'
								  and c.U_CpCode <> 'CP21301'
								  Group by b.U_PP030HNo, b.U_OrdNum
								) g on a.DocEntry=g.U_PP030HNo and a.U_OrdNum=g.U_OrdNum	
					left join  ( --**자재비
                    	        Select	a.U_OrdNum as OrdNum,
										a.U_PP030HNo as PP030HNo,
										Sum(e.U_LinTotal) as LinTotal
								  From	[@PS_MM005H] a
										Left Join (Select b.DocEntry, b.U_LineNum, b.U_CGNo 
													 From [@PS_MM010H] a Inner Join [@PS_MM010L] b On a.DocEntry = b.DocEntry
													Where a.Status = 'O') b On a.U_CgNum = b.U_CGNo
										Left Join (Select a.U_DocDate, a.U_DueDate, a.U_CardCode, a.U_CardName, b.DocEntry, b.U_LineNum, b.U_PQDocNum,
														  b.U_PQLinNum, b.U_Qty, b.U_Weight, b.U_Price, b.U_LinTotal
													 From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry
													Where a.Status = 'O') c On IsNull(c.U_PQDocNum, 0) = b.DocEntry And IsNull(c.U_PQLinNum, 0) = b.U_LineNum
										Left Join (Select a.U_DocDate, b.U_PODocNum, b.U_POLinNum, b.DocEntry, b.U_LineNum, b.U_Qty, b.U_Weight, b.U_LinTotal
											 From [@PS_MM050H] a Inner Join [@PS_MM050L] b On a.DocEntry = b.DocEntry
											Where a.Status = 'O') d On IsNull(d.U_PODocNum, 0) = c.DocEntry And IsNull(d.U_POLinNum, 0) = c.U_LineNum
										Left Join (Select a.U_DocDate, Left(b.U_GADocLin, CharIndex('-', b.U_GADocLin) - 1) As GADocNum, 
											  			  Right(b.U_GADocLin, Len(b.U_GADocLin) - CharIndex('-', b.U_GADocLin)) As GALinNum,
														  b.U_Qty, b.U_Weight, b.U_LinTotal, a.U_Purchase
													 From [@PS_MM070H] a Inner Join [@PS_MM070L] b On a.DocEntry = b.DocEntry
													Where a.Status = 'O') e On IsNull(e.GADocNum, 0) = d.DocEntry And IsNull(e.GALinNum, 0) = d.U_LineNum
										Left Join [OITM] z ON z.ItemCode = Case When a.U_BPLId = '2' And a.U_OrdType = '40' then Left(a.U_ItemCode, CharIndex('-', a.U_ItemCode) - 1) --b.U_ItemCode
												        Else a.U_ItemCode End
										Left Join [@PSH_ITMBSORT] y On y.Code = z.U_ItmBSort
										Left Join [@PS_PP030H] x On x.DocEntry = a.U_PP030HNo	
								  Where e.U_Purchase in ('10','20') 		
								 Group by a.U_OrdNum, a.U_PP030HNo ) h on h.PP030HNo = a.DocEntry
	
					--left join  ( --**자재비
     --                          select a.U_OrdNum, a.U_PP030HNo, a.U_PP030LNo,
     --                                 Left(b.U_GADocLin, CharIndex('-', b.U_GADocLin) - 1) As GADocNum, 
					--				  Right(b.U_GADocLin, Len(b.U_GADocLin) - CharIndex('-', b.U_GADocLin)) As GALinNum, 
					--				  Max(B.U_LinTotal) as Lintotal
     --                            from [@PS_MM005H] a inner join [@PS_MM070L] b on a.U_ItemCode = b.U_ItemCode
					--								 inner join [@PS_MM070H] c on b.DocEntry = c.DocEntry and c.Status = 'O'
					--								 inner join [@PS_MM050H] d on Left(b.U_GADocLin, CharIndex('-', b.U_GADocLin) - 1) = d.DocEntry and d.Status = 'O'
					--								 inner join [@PS_MM050L] e on d.DocEntry = e.DocEntry and Right(b.U_GADocLin, Len(b.U_GADocLin) - CharIndex('-', b.U_GADocLin)) = b.U_LineNum                                    
					--			 Where left(a.U_OrdNum,11) = @OrdNum 
					--			  Group by a.U_OrdNum, b.U_GADocLin, a.U_PP030HNo, a.U_PP030LNo ) h 
					--			                             on left(h.U_OrdNum,11) = a.U_OrdNum and SUBSTRING(h.U_OrdNum,13,2) = a.U_OrdSub1             
					--											and RIGHT(h.U_OrdNum,3) = a.U_OrdSub2         
					--											and h.U_PP030HNo = b.DocEntry and h.U_PP030LNo = b.LineId
					left  join ( --**설계비
								select	b.U_PP030HNo, b.U_PP030MNo, b.U_OrdNum,--b.U_CpCode,
								sum(convert(numeric(19,6),b.U_PQty)) * c.U_Price as WorkCost --도면적용매수 + 신규도면매수
								from [@PS_PP040H] a inner join [@PS_PP040L] b on a.DocEntry=b.DocEntry
													left  join [@PS_PP001L] c on b.U_CpCode=c.U_CpCode --공정단가
								where a.Canceled<>'Y'
								  and a.U_OrdType = '70'
								group by b.U_PP030HNo, b.U_PP030MNo, b.U_OrdNum, c.U_Price
								) j on a.DocEntry=j.U_PP030HNo and a.U_OrdNum=j.U_OrdNum		
                    left join  ( --**외주가공비
								Select  a.U_PP030HNo,  x.U_OrdNum,  Sum(Isnull(e.U_LinTotal,0)) As Wajuga						--검수금액
								  From	[@PS_MM005H] a
										Left Join (Select b.DocEntry, b.U_LineNum, b.U_CGNo 
													 From [@PS_MM010H] a Inner Join [@PS_MM010L] b On a.DocEntry = b.DocEntry
													Where a.Status = 'O') b On a.U_CgNum = b.U_CGNo															  --b.견적
										Left Join (Select a.U_DocDate, a.U_DueDate, a.U_CardCode, a.U_CardName, b.DocEntry, b.U_LineNum, b.U_PQDocNum,
														  b.U_PQLinNum, b.U_Qty, b.U_Weight, b.U_Price, b.U_LinTotal
													 From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry
													Where a.Status = 'O') c On IsNull(c.U_PQDocNum, 0) = b.DocEntry And IsNull(c.U_PQLinNum, 0) = b.U_LineNum  --c.품의
										Left Join (Select max(a.U_DocDate) As DocDate, b.U_PODocNum, b.U_POLinNum, Sum(b.U_Qty) As U_Qty, Sum(b.U_Weight) As U_Weight, Sum(b.U_LinTotal) As U_LinTotal
													 From [@PS_MM050H] a Inner Join [@PS_MM050L] b On a.DocEntry = b.DocEntry
													Where a.Status = 'O' Group by b.U_PODocNum, b.U_POLinNum ) d On IsNull(d.U_PODocNum, 0) = c.DocEntry And IsNull(d.U_POLinNum, 0) = c.U_LineNum	--d.입고
										Left Join (Select max(a.U_DocDate) As DocDate, c.U_PODocNum, c.U_POLinNum, Sum(b.U_Qty) As U_Qty, Sum(b.U_Weight) As U_Weight, Sum(b.U_LinTotal) As U_LinTotal
													 From [@PS_MM070H] a Inner Join [@PS_MM070L] b On a.DocEntry = b.DocEntry And a.Status = 'O'
																		 Inner Join [@PS_MM050L] c On Left(b.U_GADocLin, CharIndex('-', b.U_GADocLin) - 1) = c.DocEntry And Right(b.U_GADocLin, Len(b.U_GADocLin) - CharIndex('-', b.U_GADocLin)) = c.U_LineNum
																		 Inner join [@PS_MM050H] d On c.DocEntry = d.DocEntry And d.Status = 'O'
																		 
													Group by c.U_PODocNum, c.U_POLinNum ) e On e.U_PODocNum =  c.DocEntry And e.U_POLinNum = c.U_LineNum		--e.검수
										Left Join [OITM] z ON z.ItemCode = Case When a.U_BPLId = '2' And a.U_OrdType = '40' then Left(a.U_ItemCode, CharIndex('-', a.U_ItemCode) - 1) --b.U_ItemCode
																				Else a.U_ItemCode End
										Left Join [@PSH_ITMBSORT] y On y.Code = z.U_ItmBSort
										Left Join [@PS_PP030H] x On x.DocEntry = a.U_PP030HNo
								 Where IsNull(a.U_OrdType, '') Like '30'
								   --And	a.U_DocDate Between '20100101' And '20110525' 
								   And	a.U_Status = 'O'
								   --And  ISNULL(x.U_OrdNum,'') Like 'CM201011060' + '%'
							Group by a.U_PP030HNo, x.U_OrdNum ) l on l.U_PP030HNo = a.DocEntry
                    left join  ( --**외주제작비
								Select  a.U_PP030HNo,  x.U_OrdNum,  Sum(Isnull(e.U_LinTotal,0)) As Wajuje						--검수금액
								  From	[@PS_MM005H] a
										Left Join (Select b.DocEntry, b.U_LineNum, b.U_CGNo 
													 From [@PS_MM010H] a Inner Join [@PS_MM010L] b On a.DocEntry = b.DocEntry
													Where a.Status = 'O') b On a.U_CgNum = b.U_CGNo															  --b.견적
										Left Join (Select a.U_DocDate, a.U_DueDate, a.U_CardCode, a.U_CardName, b.DocEntry, b.U_LineNum, b.U_PQDocNum,
														  b.U_PQLinNum, b.U_Qty, b.U_Weight, b.U_Price, b.U_LinTotal
													 From [@PS_MM030H] a Inner Join [@PS_MM030L] b On a.DocEntry = b.DocEntry
													Where a.Status = 'O') c On IsNull(c.U_PQDocNum, 0) = b.DocEntry And IsNull(c.U_PQLinNum, 0) = b.U_LineNum  --c.품의
										Left Join (Select max(a.U_DocDate) As DocDate, b.U_PODocNum, b.U_POLinNum, Sum(b.U_Qty) As U_Qty, Sum(b.U_Weight) As U_Weight, Sum(b.U_LinTotal) As U_LinTotal
													 From [@PS_MM050H] a Inner Join [@PS_MM050L] b On a.DocEntry = b.DocEntry
													Where a.Status = 'O' Group by b.U_PODocNum, b.U_POLinNum ) d On IsNull(d.U_PODocNum, 0) = c.DocEntry And IsNull(d.U_POLinNum, 0) = c.U_LineNum	--d.입고
										Left Join (Select max(a.U_DocDate) As DocDate, c.U_PODocNum, c.U_POLinNum, Sum(b.U_Qty) As U_Qty, Sum(b.U_Weight) As U_Weight, Sum(b.U_LinTotal) As U_LinTotal
													 From [@PS_MM070H] a Inner Join [@PS_MM070L] b On a.DocEntry = b.DocEntry And a.Status = 'O'
																		 Inner Join [@PS_MM050L] c On Left(b.U_GADocLin, CharIndex('-', b.U_GADocLin) - 1) = c.DocEntry And Right(b.U_GADocLin, Len(b.U_GADocLin) - CharIndex('-', b.U_GADocLin)) = c.U_LineNum
																		 Inner join [@PS_MM050H] d On c.DocEntry = d.DocEntry And d.Status = 'O'
																		 
													Group by c.U_PODocNum, c.U_POLinNum ) e On e.U_PODocNum =  c.DocEntry And e.U_POLinNum = c.U_LineNum		--e.검수
										Left Join [OITM] z ON z.ItemCode = Case When a.U_BPLId = '2' And a.U_OrdType = '40' then Left(a.U_ItemCode, CharIndex('-', a.U_ItemCode) - 1) --b.U_ItemCode
																				Else a.U_ItemCode End
										Left Join [@PSH_ITMBSORT] y On y.Code = z.U_ItmBSort
										Left Join [@PS_PP030H] x On x.DocEntry = a.U_PP030HNo
								 Where IsNull(a.U_OrdType, '') Like '40'
								   --And	a.U_DocDate Between '20100101' And '20110525' 
								   And	a.U_Status = 'O'
								   --And  ISNULL(x.U_OrdNum,'') Like 'CM201011060' + '%'
							Group by a.U_PP030HNo, x.U_OrdNum ) m on m.U_PP030HNo = a.DocEntry					            
where a.U_BPLId = '2' --동래공장
  and   a.U_OrdGbn in ('105','106') --기계공구,몰드
  and   Substring(a.U_ItemCode,3,4) + '-' + Substring(a.U_ItemCode,7,2) between @SYYYYMM and @EYYYYMM
  and   left(a.U_ItemCode,1) like @BPLId
  and  (@ItemName='' or a.U_ItemName like ('%' + @ItemName + '%'))
  and  (@Size='' or c.U_Size like ('%' + @Size + '%'))
  and  (@OrdNum='' or a.U_OrdNum like ('%' + @OrdNum + '%'))

Group by a.U_OrdNum, d.CardName, a.U_ItemCode, a.U_ItemName, c.U_Size, c.SalUnitMsr,d.DocDate, g.WorkCost,
         e.Quantity, d.DocDueDate, h.Lintotal, j.WorkCost, m.Wajuje, l.Wajuga ) G
         
where G.YesNo like @Section 

Group by G.OrdNum, G.CardName, G.ItemCode, G.ItemName, G.Size, G.SalUnitMsr, G.DocDate, G.Sugum, G.DocDueDate

Order by left(G.OrdNum,1) 

----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_PP361_001] '2011-01','2011-01','','A','','','',''
--EXEC [PS_PP361_001] '2011-01','2011-01','2','A','','','',''


--EXEC PS_PP361_001 '2011-02','2011-02','%','%', '%','%', '%', '%'

--EXEC PS_PP361_001 '2011-02','2011-02','%','%', '%','%', '%', 'CM201011060'