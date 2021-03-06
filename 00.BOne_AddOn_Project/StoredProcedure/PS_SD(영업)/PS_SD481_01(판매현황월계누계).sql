USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_SD481_01]    Script Date: 04/23/2011 11:56:55 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : SD																							*/
/*  Description    : 판매관리 > 수주처별 판매현황 (집계표) [PS_SD481_01]										                */
/*  Create Date    : 2011.04.23                                                                                */
/*  Modified Date  :																							*/
/*  Creator        : N.G.Y																				*/
/*  Company        : Poongsan Holdings																			*/

--CREATE PROC [dbo].[PS_SD481_01]
ALTER PROC [dbo].[PS_SD481_01]
(	
	@BPLId		Nvarchar(10),
	@DocDateFr	Date--,	
	--@DocDateTo	Date--,
	--@SaleGbn	Nvarchar(30),
	--@ItmBsort	Nvarchar(30)
)
AS

-- AR/CR table definition

create table #Z_Invoice ( 
	CardCode nvarchar(10) collate Korean_Wansung_Unicode_CI_AS,
	ym		 char(6),
	Code nvarchar(10) collate Korean_Wansung_Unicode_CI_AS,
	Name nvarchar(30) collate Korean_Wansung_Unicode_CI_AS,
	MCode nvarchar(10) collate Korean_Wansung_Unicode_CI_AS,
	MName nvarchar(30) collate Korean_Wansung_Unicode_CI_AS,
	vatgroup nvarchar(5) collate Korean_Wansung_Unicode_CI_AS,
	Quantity decimal(10,0),
	Weight decimal(15,3), 
	LineTotal decimal(15,0)
 )
						  			  
BEGIN

If @BPLId = '3'
begin

	insert into #Z_Invoice ( CardCode, ym, Code, Name, MCode, MName, vatgroup, Quantity, Weight,  LineTotal )
	select a.CardCode,
		   Convert(char(6),a.DocDate,112),
		   d.Code, 
		   d.Name,
		   e.U_Code as MCode,
		   e.U_CodeName as MName, 
		   b.VatGroup,	
		   ISNULL(b.Quantity,0) as Quantity, 
		   Case When c.U_UnWeight <= 0 Then b.Quantity
		   Else 	round((b.Quantity * c.U_UnWeight)/1000,0) END AS Weight,
		   ISNULL(b.LineTotal,0) as LineTotal
	from OINV a 
		 inner join INV1 b ON a.DocEntry = b.DocEntry
		 inner join OITM c on b.ItemCode = c.ItemCode
		 inner join [@PSH_ITMBSORT] d on c.U_ItmBsort = d.Code
		 inner join [@PSH_ITMMSORT] e on c.U_itmmsort = e.U_Code
	where (a.BPLId In ('3', '5'))
	  and a.DocDate between left(Convert(char(8),@DocDateFr,112),4) + '0101' and @DocDateFr
	  --and a.DocDate between @DocDateFr and @DocDateTo
--	  and (d.Code = @ItmBsort Or @ItmBsort = '')
	  --and d.code between '101' and '104'
		
		-- DR
	insert into #Z_Invoice ( CardCode, ym, Code, Name, MCode, MName, vatgroup, Quantity, Weight,  LineTotal )
	select  a.CardCode,
		   Convert(char(6),a.DocDate,112),
			d.Code, 
			d.Name,
			e.U_Code as MCode,
			e.U_CodeName as MName, 
			b.VatGroup, 
			ISNULL(b.Quantity,0) * (-1) as Quantity, 
		   Case When c.U_UnWeight <= 0 Then b.Quantity * -1 Else round((b.Quantity * c.U_UnWeight) * (-1)/1000,0) END AS Weight, 
			ISNULL(b.LineTotal,0) * (-1) as LineTotal
	from ORIN a
		 inner join RIN1 b ON a.DocEntry = b.DocEntry 
		 inner join OITM c on b.ItemCode = c.ItemCode
		 inner join [@PSH_ITMBSORT] d on c.U_ItmBsort = d.Code
		 inner join [@PSH_ITMMSORT] e on c.U_itmmsort = e.U_Code
		 
	where (a.BPLId In ('3', '5'))
	  and a.DocDate between left(Convert(char(8),@DocDateFr,112),4) + '0101' and @DocDateFr
	  --and a.DocDate between @DocDateFr and @DocDateTo
--	  and (d.Code = @ItmBsort Or @ItmBsort = '')
	  --and d.code between '101' and '104' 
 End

Else
Begin
	-- AR
	insert into #Z_Invoice ( CardCode, ym, Code, Name, MCode, MName, vatgroup, Quantity, Weight,  LineTotal )
	select a.CardCode,
		   Convert(char(6),a.DocDate,112),
		   d.Code, 
		   d.Name,
		   e.U_Code as MCode,
		   e.U_CodeName as MName, 
		   b.VatGroup,	
		   ISNULL(b.Quantity,0) as Quantity, 
		   Case When c.U_UnWeight <= 0 Then b.Quantity
		   Else 	round((b.Quantity * c.U_UnWeight)/1000,0) END AS Weight,
		   ISNULL(b.LineTotal,0) as LineTotal
	from OINV a 
		 inner join INV1 b ON a.DocEntry = b.DocEntry
		 inner join OITM c on b.ItemCode = c.ItemCode
		 inner join [@PSH_ITMBSORT] d on c.U_ItmBsort = d.Code
		 inner join [@PSH_ITMMSORT] e on c.U_itmmsort = e.U_Code
	where (a.BPLId = @BPLId Or @BPLId = '')
	  and a.DocDate between left(Convert(char(8),@DocDateFr,112),4) + '0101' and @DocDateFr
	  --and a.DocDate between @DocDateFr and @DocDateTo
	  --and (d.Code = @ItmBsort Or @ItmBsort = '')
	  --and d.code between '101' and '104'
		
		-- DR
	insert into #Z_Invoice ( CardCode, ym, Code, Name, MCode, MName, vatgroup, Quantity, Weight,  LineTotal )
	select a.CardCode,
		   Convert(char(6),a.DocDate,112),
			d.Code, 
			d.Name,
			e.U_Code as MCode,
			e.U_CodeName as MName, 
			b.VatGroup, 
			ISNULL(b.Quantity,0) * (-1) as Quantity, 
		   Case When c.U_UnWeight <= 0 Then b.Quantity * -1 Else round((b.Quantity * c.U_UnWeight) * (-1)/1000,0) END AS Weight, 
			ISNULL(b.LineTotal,0) * (-1) as LineTotal
	from ORIN a
		 inner join RIN1 b ON a.DocEntry = b.DocEntry 
		 inner join OITM c on b.ItemCode = c.ItemCode
		 inner join [@PSH_ITMBSORT] d on c.U_ItmBsort = d.Code
		 inner join [@PSH_ITMMSORT] e on c.U_itmmsort = e.U_Code
		 
	where (a.BPLId = @BPLId Or @BPLId = '' )
	  and a.DocDate between left(Convert(char(8),@DocDateFr,112),4) + '0101' and @DocDateFr
	  --and a.DocDate between @DocDateFr and @DocDateTo
	  --and (d.Code = @ItmBsort Or @ItmBsort = '')
	  --and d.code between '101' and '104' 
End
	
	-- 판매구분 전체
	--If @SaleGbn = ''
	--BEGIN	   
		Select t.CardName,
			   t.Code As '대분류코드',
			   t.name As '대분류명',
			   t.mcode AS '중분류코드',
			   t.mname As '중분류명',	
			   t.mquantity as '월판매수량',
			   t.mweight as '월판매중량',
			   t.mlinetotal as '월판매금액',	
			   t.yquantity as '판매누계수량',
			   t.yweight as '판매누계중량',
			   t.ylinetotal as '판매누계금액'
		From (
		select Case When b.U_CdNaming In ('A','C','D','P','I') Then b.CardName
			 When b.U_CdNaming = 'U' Then '수출'
			 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '사외'
		End As CardName,
		Case When b.U_CdNaming In ('A','C','D','P','I') Then '1'
			 When b.U_CdNaming = 'U' Then '2'
			 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '3'
		End As cnt ,
		convert(nvarchar(20),a.code) As code,
		convert(nvarchar(20), a.name)As name, 
		convert(nvarchar(50),a.mcode)As mcode,
		convert(nvarchar(50), a.mname) As mname,
		Sum(isnull(Charindex(ym, Convert(char(6),@DocDateFr,112)) * a.Quantity,0)) as mquantity,
		Sum(isnull(Charindex(ym, Convert(char(6),@DocDateFr,112)) * a.Weight,0)) as mweight,
		Sum(isnull(Charindex(ym, Convert(char(6),@DocDateFr,112)) * a.LineTotal,0)) as mlinetotal,
		Sum(isnull(a.Quantity,0)) as yquantity,
		Sum(isnull(a.Weight,0)) as yweight,
		Sum(isnull(a.LineTotal,0)) as ylinetotal
		from #Z_Invoice a Inner Join OCRD b On a.CardCode = b.CardCode
		--group by Code, name, MCode, mname
		group by Case When b.U_CdNaming In ('A','C','D','P','I') Then b.CardName
					  When b.U_CdNaming = 'U' Then '수출'
			 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '사외' End,
		Case When b.U_CdNaming In ('A','C','D','P','I') Then '1'
			 When b.U_CdNaming = 'U' Then '2'
			 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '3'
		End,
				 a.Code, a.name, a.MCode, a.mname
		) t
	Order by t.cnt, t.CardName, t.Code, t.name, t.mcode, t.mname
	--END
	
	---- 판매구분 내수
	--Else if @SaleGbn = 'N'
	--BEGIN
	--	select Case When b.U_CdNaming In ('A','C','D','P','I') Then b.CardName
	--		 When b.U_CdNaming = 'U' Then '수출'
	--		 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '사외'
	--	End,
	--	convert(nvarchar(20),a.code) code, convert(nvarchar(20), a.name) name, 
	--		   convert(nvarchar(50),a.mcode) mcode, convert(nvarchar(50), a.mname) mname,
	--		   Sum(isnull(a.Quantity,0)) as quantity, Sum(isnull(a.Weight,0)) as weight, Sum(isnull(a.LineTotal,0)) as linetotal
	--	from #Z_Invoice a Inner Join OCRD b On a.CardCode = b.CardCode
	--	where a.vatgroup in ('A0', 'A1', 'A2', 'A5', 'A6', 'A7', 'A8', 'A9')
	--	group by Case When b.U_CdNaming In ('A','C','D','P','I') Then b.CardName
	--				  When b.U_CdNaming = 'U' Then '수출'
	--		 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '사외' End,
	--			 a.Code, a.name, a.MCode, a.mname
	--END

	---- 판매구분 수출
	--Else if @SaleGbn = 'Y'
	--BEGIN
	--	select Case When b.U_CdNaming In ('A','C','D','P','I') Then b.CardName
	--		 When b.U_CdNaming = 'U' Then '수출'
	--		 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '사외'
	--	End,
	--	convert(nvarchar(20),a.code) code, convert(nvarchar(20), a.name) name, 
	--		   convert(nvarchar(50),a.mcode) mcode, convert(nvarchar(50), a.mname) mname,
	--		   Sum(isnull(a.Quantity,0)) as quantity, Sum(isnull(a.Weight,0)) as weight, Sum(isnull(a.LineTotal,0)) as linetotal
	--	from #Z_Invoice a Inner Join OCRD b On a.CardCode = b.CardCode
	--	where vatgroup = 'A4'
	--	group by Case When b.U_CdNaming In ('A','C','D','P','I') Then b.CardName
	--				  When b.U_CdNaming = 'U' Then '수출'
	--		 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '사외' End,
	--			 a.Code, a.name, a.MCode, a.mname
	--END

	---- 판매구분 로칼
	--Else if @SaleGbn = 'L'
	--BEGIN
	--	select Case When b.U_CdNaming In ('A','C','D','P','I') Then b.CardName
	--		 When b.U_CdNaming = 'U' Then '수출'
	--		 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '사외'
	--	End,
	--	convert(nvarchar(20),a.code) code, convert(nvarchar(20), a.name) name, 
	--		   convert(nvarchar(50),a.mcode) mcode, convert(nvarchar(50), a.mname) mname,
	--		   Sum(isnull(a.Quantity,0)) as quantity, Sum(isnull(a.Weight,0)) as weight, Sum(isnull(a.LineTotal,0)) as linetotal
	--	from #Z_Invoice a Inner Join OCRD b On a.CardCode = b.CardCode
	--	where vatgroup = 'A3'
	--	group by Case When b.U_CdNaming In ('A','C','D','P','I') Then b.CardName
	--				  When b.U_CdNaming = 'U' Then '수출'
	--		 When b.U_CdNaming = 'H' Or Isnull(b.U_CdNaming,'') = '' Then '사외' End,
	--			 a.Code, a.name, a.MCode, a.mname
	--END

	
	--exec [PS_SD481_01] '3', '20110331'
END	