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
--CREATE  PROC [dbo].[PS_QM040_03]
ALTER     PROC [dbo].[PS_QM040_03]
(
  @PackingNo		as nvarchar(20),
  @ItemCode			as nvarchar(20),
  @CardCode			as nvarchar(20)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------
--//2.Dimension & Surface 치수 및 표면상태
Select	convert(nvarchar(100),a.U_PackNo)	PackNo,
		convert(nvarchar(100),b.U_LotNo)	FG_LotNo1,
		'S'+convert(nvarchar(100),c.U_CallSize)+right(convert(nvarchar(100),b.U_LotNo),7) as FG_LotNo2,
		convert(nvarchar(100),b.U_ItemCode)	ItemCode,
		convert(nvarchar(100),c.FrgnName)	FrgnName,
		convert(nvarchar(100),c.U_Size)		Size,
		convert(nvarchar(100),d.U_CardCode)	CardCode,
		convert(nvarchar(100),d.U_CardName)	CardName,
		
		b.U_Weight Weight,
		
		case when isnull(f.U_MulGbn1,'')='10' then 'D/G'
			 when isnull(f.U_MulGbn1,'')='20' then 'UN D/G' end as 'Appearance',
				
		--[@PS_QM010H]:검사사양서
		e.U_A_Pl		A_Pl,
		e.U_A_Mi		A_Mi,
		e.U_A_Val		A_Val,
		e.U_B_Pl		B_Pl,
		e.U_B_Mi		B_Mi,
		e.U_B_Val		B_Val,
		e.U_C_Pl		C_Pl,
		e.U_C_Mi		C_Mi,
		e.U_C_Val		C_Val,
		e.U_D_Pl		D_Pl,
		e.U_D_Mi		D_Mi,
		e.U_D_Val		D_Val,
		e.U_EE1_Val		EE1_Val,
		e.U_F_Pl		F_Pl,
		e.U_F_Mi		F_Mi,
		e.U_F_Val		F_Val,
		e.U_F1_Pl		F1_Pl,
		e.U_F1_Mi		F1_Mi,
		e.U_F1_Val		F1_Val,
		e.U_G_Pl		G_Pl,
		e.U_G_Mi		G_Mi,
		e.U_G_Val		G_Val,
		e.U_EdgeBurr	EdgeBurr,
		e.U_Camber_C	Camber_C,
		e.U_Camber_M	Camber_M,
		e.U_Surface		Surface,
		e.U_CrossBow	CrossBow,
		e.U_CR_Val		CR_Val,
		e.U_CR_Pl		CR_Pl,
		e.U_CR_Mi		CR_Mi,
		
		--[@PS_QM020H]:검사성적서
		d.U_M_A_min		M_A_min,
		d.U_M_A_max		M_A_max,
		d.U_M_B_min		M_B_min,
		d.U_M_B_max		M_B_max,
		d.U_M_C			M_C,
		d.U_M_D			M_D,
		d.U_M_EE1		M_EE1,
		d.U_M_F			M_F,
		d.U_M_F1		M_F1,
		d.U_M_G			M_G,
		d.U_M_EdgeBu	M_EdgeBu,
		d.U_M_Camber	M_Camber,
		d.U_M_SR_min	M_SR_min,
		d.U_M_SR_max	M_SR_max,
		d.U_M_CrossB	M_CrossB,
		d.U_M_CR_min	M_CR_min,
		d.U_M_CR_max	M_CR_max	
				
				
From	[@PS_PP090H] a	inner join [@PS_PP090L] b on a.DocEntry = b.DocEntry
						left  join [OITM] c on b.U_ItemCode=c.ItemCode
						left  join [@PS_QM020H] d on b.U_LotNo = d.U_OrdNum
						left  join [@PS_QM010H] e on d.U_ItemCode=e.U_ItemCode and d.U_CardCode=e.U_CardCode
						left  join [@PS_PP030H] f on b.U_LotNo = f.U_OrdNum
Where	a.U_PackNo = @PackingNo
order by a.U_PackNo

----------------------------------------------------------------------------------------------------------------------------------------
--EXEC [PS_QM040_03] '20101111002', '109217','10004'
--EXEC [PS_QM040_03] '20101111001', '104010001','10003'