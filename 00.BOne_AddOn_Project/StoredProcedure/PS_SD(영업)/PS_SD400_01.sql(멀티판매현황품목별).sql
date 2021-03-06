USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_SD400_01]    Script Date: 03/11/2011 20:57:19 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : SD																							*/
/*  Description    : 판매관리 > 판매현황(품목별) [PS_SD400_01]							                    */
/*  Create Date    : 2010.12.02                                                                                 */
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/

--Create PROC [dbo].[PS_SD400_01]
ALTER PROC [dbo].[PS_SD400_01]
(	
	@BPLId		Nvarchar(10),
	@DocDateFr	Date,
	@DocDateTo	Date,
	@BeginDay	date
)	
AS	 
SELECT
		G.ItemCode AS ItemCode,
		G.ItemName AS ItemName,
		Convert(Nvarchar(100),G.Size) AS Size,
		SUM(ISNULL(G.MTalQty, 0)) AS MTalQty,
		SUM(ISNULL(G.YTalQty, 0)) AS YTalQty,
		SUM(ISNULL(G.MTalTot, 0)) AS MTalTot,
		SUM(ISNULL(G.YTalTot, 0)) AS YTalTot,
		SUM(ISNULL(G.MNTalQty, 0)) AS MNTalQty,
		SUM(ISNULL(G.YNTalQty, 0)) AS YNTalQty,
		SUM(ISNULL(G.MNTalTot, 0)) AS MNTalTot,
		SUM(ISNULL(G.YNTalTot, 0)) AS YNTalTot,
		SUM(ISNULL(G.MSumQty, 0)) AS MSumQty,
		SUM(ISNULL(G.YSumQty, 0)) AS YSumQty,
		SUM(ISNULL(G.MSumTot, 0)) AS MSumTot,
		SUM(ISNULL(G.YSumTot, 0)) AS YSumTot
	FROM
(
select c.ItemCode AS ItemCode,
	   c.FrgnName AS ItemName,
	   c.U_size AS Size, 
	   Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '10') * b.Quantity) AS MTalQty,-- 월간 탈지 중량
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '10') * b.Quantity) AS YTalQty,-- 년간	탈지 중량
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '10') * b.LineTotal) AS MTalTot,-- 월간 탈지 금액
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '10') * b.LineTotal) AS YTalTot,-- 년간 탈지 금액
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '20') * b.Quantity) AS MNTalQty,-- 월간 비탈지 중량
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '20') * b.Quantity) AS YNTalQty,-- 년간 비탈지 중량	
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '20') * b.LineTotal) AS MNTalTot,-- 월간 비탈지 금액
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '20') * b.LineTotal) AS YNTalTot,-- 년간 비탈지 금액
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * b.Quantity) AS MSumQty,-- 월간 중량 합계
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * b.Quantity) AS YSumQty,-- 년간 중량 합계
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * b.LineTotal) AS MSumTot,-- 월간 금액 합계
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * b.LineTotal) AS YSumTot-- 년간 금액 합계
	   
from OINV a 
	 inner join INV1 b ON a.DocEntry = b.DocEntry
	 inner join OITM c on b.ItemCode = c.ItemCode
	 inner Join DLN1 d on b.BaseEntry = d.DocEntry And b.BaseLine = d.LineNum
	 LEFT  JOIN [IBT1] IBT1 ON IBT1.BaseType = 15 And d.DocEntry = IBT1.BaseNum And d.LineNum = IBT1.BaseLinNum
	 LEFT  JOIN [@PS_PP030H] PP030 ON IBT1.BatchNum = PP030.U_OrdNum
where (a.BPLId = @BPLId Or @BPLId = '' )
  and c.U_ItmBsort = '104'
  and a.DocDate between Convert(char(4),@DocDateFr,112) + '0101' and @DocDateTo
group by c.ItemCode, c.FrgnName, c.U_Size

Union all

select isnull(c.ItemCode, '') AS ItemCode,
	   isnull(C.ItemName, '') AS ItemName,
	   isnull(c.U_size, '') AS Size, 
	   Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '10') * b.Quantity) AS MTalQty,-- 월간 탈지 중량
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '10') * b.Quantity) AS YTalQty,-- 년간	탈지 중량
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '10') * b.LineTotal) AS MTalTot,-- 월간 탈지 금액
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '10') * b.LineTotal) AS YTalTot,-- 년간 탈지 금액
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '20') * b.Quantity) AS MNTalQty,-- 월간 비탈지 중량
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '20') * b.Quantity) AS YNTalQty,-- 년간 비탈지 중량	
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '20') * b.LineTotal) AS MNTalTot,-- 월간 비탈지 금액
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * charindex(PP030.U_Mulgbn1, '20') * b.LineTotal) AS YNTalTot,-- 년간 비탈지 금액
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * b.Quantity) AS MSumQty,-- 월간 중량 합계
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * b.Quantity) AS YSumQty,-- 년간 중량 합계
		Sum(CHARINDEX(CONVERT(char(6),a.DocDate,112), CONVERT(char(6),@DocDateTo,112)) * b.LineTotal) AS MSumTot,-- 월간 금액 합계
		Sum(CHARINDEX(CONVERT(char(4),a.DocDate,112), CONVERT(char(4),@DocDateTo,112)) * b.LineTotal) AS YSumTot-- 년간 금액 합계
from ORIN a
	 inner join RIN1 b ON a.DocEntry = b.DocEntry 
	 inner join OITM c on b.ItemCode = c.ItemCode
	 inner Join DLN1 d on b.BaseEntry = d.DocEntry And b.BaseLine = d.LineNum
	 LEFT  JOIN [IBT1] IBT1 ON IBT1.BaseType = 15 And d.DocEntry = IBT1.BaseNum And d.LineNum = IBT1.BaseLinNum
	 LEFT  JOIN [@PS_PP030H] PP030 ON IBT1.BatchNum = PP030.U_OrdNum
where (a.BPLId = @BPLId Or @BPLId = '' )
  and c.U_ItmBsort = '104'
  and a.DocDate between Convert(char(4),@DocDateFr,112) + '0101' and @DocDateTo
  group by isnull(c.ItemCode, ''), isnull(c.ItemName, ''), isnull(c.U_Size, '')
  ) G
 Group by G.ItemCode, G.ItemName, Convert(Nvarchar(100),G.Size)

-- EXEC PS_SD400_01 '1', '20110101', '20110228', '2011.01.01'
-- EXEC PS_SD400_01 '2', '20101101', '20101208',  '2010.01.01'