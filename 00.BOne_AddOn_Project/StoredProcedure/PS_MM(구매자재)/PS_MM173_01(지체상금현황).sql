USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM173_01]    Script Date: 05/05/2011 09:11:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : ���Ű���																				    */
/*  Description    : ��ü��� ��� ��Ȳ																			*/
/*  ALTER  Date    : 2011.05.05																					*/
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_MM173_01]
ALTER     PROC [dbo].[PS_MM173_01]
(
    @BPLId		as NvarChar(10),  --�����
    @DocDatef    as Nvarchar(8),  -- ����F
    @DocDatet    as Nvarchar(8)   -- ����T
 )
AS

Begin

Create Table #MM172_Temp
( DocEntry	NvarChar(20) collate Korean_Wansung_Unicode_CI_AS,
  LineNum	NvarChar(10) collate Korean_Wansung_Unicode_CI_AS,
  ItemCode  NvarChar(20),
  ItemName  NvarChar(100),
  Qty		Numeric(12,3),
  Amt		Numeric(12,0),
  CardCode  Nvarchar(10),
  CardName  Nvarchar(50),
  inDate	Date,	--�԰���
  ShipDate  Date, -- ������
  MM030NO	Nvarchar(10), --ǰ�ǹ�����ȣ
  MM050NO	Nvarchar(10), --�԰�����ȣ
  MM070NO	Nvarchar(10),-- �˼�������ȣ
  MM070LNO  Nvarchar(10),-- �˼����ι�ȣ
  ItemGpCd  Nvarchar(10),-- ���ű���
  Status	Nvarchar(20)-- ��ü���ó�� ����
 )
 
 Insert Into #MM172_Temp( DocEntry, LineNum, ItemCode, ItemName, Qty, Amt ) 
  Select t.DocEntry, t.LineNum, t.ItemCode, t.Dscription, Qty = Sum(t.Qty), Amt = Sum(t.Amt)
  From (
  
  SELECT b.DocEntry, b.LineNum, b.ItemCode, b.Dscription, Case When a.DocType = 'I' Then b.Quantity Else b.U_SWeight End As Qty , b.LineTotal As Amt
    FROM OPDN a Inner Join PDN1 b On a.DocEntry = b.DocEntry
   Where a.DocDate between @DocDatef and @DocDatet
     And a.BPLId = @BPLId
  Union all  
   
  SELECT b.BaseEntry As DocEntry, b.BaseLine As LineNum, b.ItemCode, b.Dscription, Case When a.DocType = 'I' Then b.Quantity * -1 Else b.U_SWeight * -1 End As Qty , b.LineTotal * -1 As Amt
    FROM ORPD a Inner Join RPD1 b On a.DocEntry = b.DocEntry
   Where a.DocDate between @DocDatef and @DocDatet
     And a.BPLId = @BPLId
   ) t
  Group by t.DocEntry, t.LineNum, t.ItemCode, t.Dscription
  Having Sum(t.Amt) > 0

--//�ŷ�ó, ������ Update
Update #MM172_Temp
   set CardCode = a.CardCode,
	   CardName = a.CardName,
	   ShipDate = d.DocDueDate
  From OPDN a,
	   PDN1 b,
	   POR1 c,
	   OPOR d
 Where a.DocEntry = b.DocEntry
   And b.DocEntry = #MM172_Temp.DocEntry
   And b.LineNum = #MM172_Temp.LineNum
   And b.BaseEntry = c.DocEntry
   And b.BaseLine = c.LineNum
   And c.DocEntry = d.DocEntry
   
--//�԰����� Update
Update #MM172_Temp
   set InDate = c.U_DocDate,
	   MM030NO = d.U_PODocNum,
	   MM050NO = c.DocEntry,
	   MM070NO = a.DocEntry,
	   MM070LNO = b.LineId,
	   ItemGpCd = b.U_ItemGpCd
  from [@ps_mm070H] a,
	   [@ps_mm070L] b,
	   [@ps_mm050H] c,
	   [@ps_mm050L] d
 where a.DocEntry = b.DocEntry
   And c.DocEntry = d.DocEntry
   And b.U_GaDocLin = Convert(Nvarchar(10),d.DocEntry) + '-' + Convert(Nvarchar(10),d.LineId)
   and #MM172_Temp.DocEntry = a.U_GRDocNum
   and #MM172_Temp.LineNum = b.VisOrder
  
  Update #MM172_Temp
	 Set Status = a.U_RepayYN
	From [@PS_MM170L] a
   Where #MM172_Temp.DocEntry = a.U_GRDocNum
     and #MM172_Temp.LineNum = a.U_GRLinNum
  
Select DocEntry,
	   LineNum,
	   MM030NO,
	   MM050NO,
	   MM070NO,
	   MM070LNO,
	   ItemGpCd = (Case When ItemGpCd = '10' Then '�����ǰ��'
						When ItemGpCd = '20' Then '�����ǰ��'
						When ItemGpCd = '30' Then '������ǰ��'
						When ItemGpCd = '40' Then '��������ǰ��'
						When ItemGpCd = '50' Then '��ǰǰ��'
						When ItemGpCd = '60' Then '�����ڻ�ǰ��' End),
	   CardCode,
	   CardName,
	   InDate,
	   ShipDate,
	   Isnull(ItemCode,'') ItemCode,
	   ItemName,
	   Qty,
	   Amt,
	   datediff (dd, ShipDate, InDate) il,
       round(datediff (dd, shipDate, InDate) * 0.0015 * amt,0) jamt,
       Status = Case When Isnull(Status,'') = '' Then '��ó��'
					 When Isnull(Status,'') = 'Y' Then '�ΰ�ó��'
					 When Isnull(Status,'') = 'N' Then '����ó��' End
  from #MM172_Temp
Where ( datediff (dd, ShipDate, InDate)> 20
   and   datediff (dd, ShipDate, InDate) * 0.0015 * amt >= 10000 )
   or  ( datediff (dd, ShipDate, InDate) > 10
   and   datediff (dd, ShipDate, InDate) <= 20
   and   datediff (dd, ShipDate, InDate) * 0.0015 * amt >= 100000 )

--select * from [@ps_mm070L] Where DocEntry = 2672

End
--EXEC [PS_MM173_01] '1', '20110401', '20110430'