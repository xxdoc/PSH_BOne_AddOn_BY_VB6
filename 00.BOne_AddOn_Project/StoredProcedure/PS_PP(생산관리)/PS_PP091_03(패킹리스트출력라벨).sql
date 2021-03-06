USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP091_03]    Script Date: 03/08/2011 10:13:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 패킹리스트조회 및 출력 [PS_PP091]                                                          */
/*  Create Date    : 2010.11.26                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Ryu Yung Jo																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP091_03]
--Create  PROC [dbo].[PS_PP091_03]
AS

SET NOCOUNT ON
--BEGIN ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
Select	Convert(Nvarchar(20), e.U_CardName) As CardName
		,a.U_InDate As PackDate
		,Convert(Nvarchar(20), b.U_ItemCode) As ItemCode
		,Convert(Nvarchar(100), z.FrgnName) As ItemName
		,Convert(Nvarchar(100), z.U_Size) As Size		
		,Convert(Nvarchar(20), a.U_PackNo) As PackNo
		,d.BPLName
		,Convert(Nvarchar(20),Case When d.BPLId = '1' and left(b.U_LotNo,6) < '201103' Then Left(e.U_LotNo, 7) + '-' + 'S' + z.U_CallSize + (select RIGHT(DS.lotno,7) from Z_DSMDFRY DS Where DS.custlotno = L.U_BatchNum)
		     When d.BPLId = '1' and left(b.U_LotNo,6) >= '201103' Then Left(e.U_LotNo, 7) + '-' + 'S' + z.U_CallSize + RIGHT(b.U_LotNo, 7)
			 Else Left(e.U_LotNo, 7) + '-' + 'V' + z.U_CallSize + RIGHT(b.U_LotNo, 7)
		End) As LotNo
		--,Convert(Nvarchar(20), Left(e.U_LotNo, 7) + '-' + (CASE When d.BPLId = '1' Then 'S' When d.BPLId = '2' Then 'V' End) + z.U_CallSize + RIGHT(b.U_LotNo, 7)) As LotNo
		,b.U_Weight As Weight
		,Case When IsNull(f.U_MulGbn1, '') = '10' Then 'D/G' When IsNull(f.U_MulGbn1, '') = '20' Then 'UN D/G' Else '' End As Status
		,g.SumQty As SumQty
		,g.NetWt As NetWt		
  From	[@PS_PP090H] a 
		Inner Join [@PS_PP090L] b On a.DocEntry = b.DocEntry
		Inner Join [Z_PS_PP091] c On a.U_PackNo = c.PackNo
		Inner Join [OBPL] d On d.BPLId = a.U_BPLId
		Inner Join [@PS_QM020H] e On e.U_ItemCode = b.U_ItemCode And e.U_OrdNum = b.U_LotNo
		Inner Join [@PS_PP030H] f On f.U_OrdNum = b.U_LotNo
		Inner Join [@PS_PP030L] L On f.DocEntry = L.DocEntry
		Inner Join [OITM] z On z.ItemCode = b.U_ItemCode
		Inner Join (Select	a.DocEntry, Sum(b.U_Qty) As SumQty, Sum(b.u_Weight) As NetWt
					  From	[@PS_PP090H] a Inner Join [@PS_PP090L] b On a.DocEntry = b.DocEntry
					Group by a.DocEntry) g On g.DocEntry = a.DocEntry
--Group by a.U_PackNo, d.BPLName, a.U_InDate, b.U_ItemCode, b.U_ItemName, b.U_LotNo, b.U_Weight, e.U_CardName, z.FrgnName, z.U_Size,
--		 Convert(Nvarchar(20), Left(e.U_LotNo, 7) + '-' + (CASE When d.BPLId = '1' Then 'S' When d.BPLId = '2' Then 'V' End) + z.U_CallSize + RIGHT(b.U_LotNo, 7)),
--		 b.U_LineNum, f.U_MulGbn1
Order by PackNo, b.U_LineNum
--THE END //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-- EXEC [PS_PP091_03] 
-- select * from [@PS_PP090H]  select * from [@PS_PP090L]
-- select * from [@PS_QM020H]

-- update [@ps_mm005h] set  U_Status = 'O', U_CntcCode = '1', U_DeptCode = '3'









