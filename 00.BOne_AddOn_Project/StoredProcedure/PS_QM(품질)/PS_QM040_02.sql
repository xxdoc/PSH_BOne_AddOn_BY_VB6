SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 품질관리																				    */
/*  Description    : 검사성적서																				*/
/*  ALTER  Date    : 2010.11.11																				*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_QM040_02]
ALTER     PROC [dbo].[PS_QM040_02]
(
  @PackingNo		as nvarchar(20),
  @ItemCode			as nvarchar(20),
  @CardCode			as nvarchar(20)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
--//Header Data
select	a.U_PackNo,count(a.U_PackNo) as CountCoil,sum(b.U_Weight) as SumWeight
into	#QM040_01
From	[@PS_PP090H] a	inner join [@PS_PP090L] b on a.DocEntry = b.DocEntry
Where	a.U_PackNo = @PackingNo
group by a.U_PackNo
--select	* from #QM040_01
------------------------------------------------------------------------------
------------------------------------------------------------------------------
--//1.Mechanical Property & Chemical Composition 기계적 및 화학적 성질********
Select	distinct
		convert(nvarchar(100),a.U_PackNo)   PackNo,
		convert(nvarchar(100),b.U_ItemCode) ItemCode,
		convert(nvarchar(100),c.FrgnName)   FrgnName,
		convert(nvarchar(100),c.U_Size)	    Size,
		convert(nvarchar(100),d.U_CardCode) CardCode,
		convert(nvarchar(100),d.U_CardName) CardName,
		
		f.CountCoil,
		f.SumWeight,
		
		convert(nvarchar(100),d.U_LotNo)	 LotNo,
		
		--[@PS_QM010H]:검사사양서
		e.U_TS_min	 TS_min,
		e.U_TS_max	 TS_max,
		e.U_Elong    Elong,
		e.U_Hard_min Hard_min,
		e.U_Hard_max Hard_max,
		e.U_C_Fe_min C_Fe_min,
		e.U_C_Fe_max C_Fe_max,
		e.U_C_P_min  C_P_min,
		e.U_C_P_max  C_P_max,
		e.U_Chem_Pb  Chem_Pb,
		
		--[@PS_QM020H]:검사성적서
		d.U_M_TS_min	M_TS_min,
		d.U_M_TS_max	M_TS_max,
		d.U_M_El_min	M_El_min,
		d.U_M_El_max	M_El_max,
		d.U_M_Hd_min	M_Hd_min,
		d.U_M_Hd_max	M_Hd_max,
		d.U_M_Ch_Fe		M_Ch_Fe,
		d.U_M_Ch_P		M_Ch_P		
		
from	[@PS_PP090H] a	inner join [@PS_PP090L] b on a.DocEntry = b.DocEntry
						left  join [OITM] c on b.U_ItemCode=c.ItemCode
						left  join [@PS_QM020H] d on b.U_LotNo = d.U_OrdNum
						left  join [@PS_QM010H] e on d.U_ItemCode = e.U_ItemCode and d.U_CardCode = e.U_CardCode
						left  join #QM040_01 f on a.U_PackNo = f.U_PackNo
where	a.U_PackNo = @PackingNo
order by convert(nvarchar(100),a.U_PackNo)

------------------------------------------------------------------------------
------------------------------------------------------------------------------
--EXEC [PS_QM040_02] '20101111002', '109217','10004'
--EXEC [PS_QM040_02] '20101111001', '104010001','10003'

