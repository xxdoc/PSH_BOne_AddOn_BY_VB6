SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 前龙包府																				    */
/*  Description    : 八荤己利辑																				*/
/*  ALTER  Date    : 2010.11.09																					*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_QM040_01]
ALTER     PROC [dbo].[PS_QM040_01]
(
  @YYYYMM		as nvarchar(7)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-----------------------------------------------------------------------------------------------------------------------------------------

Select	a.U_PackNo,
		b.U_ItemCode,
		b.U_ItemName,
		c.U_CardCode,
		c.U_CardName
		
From	[@PS_PP090H] a inner Join [@PS_PP090L] b on a.DocEntry = b.DocEntry
					   left  join [@PS_QM020H] c on b.U_LotNo = c.U_OrdNum
Where	convert(char(7),a.U_InDate,20) = @YYYYMM
Group by a.U_PackNo,b.U_ItemCode,b.U_ItemName,c.U_CardCode,c.U_CardName
order by a.U_PackNo


----------------------------------------------------------------------------------------------------------------------------------------
--EXEC PS_QM040_01 '2010-11'