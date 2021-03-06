USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM030_04]    Script Date: 02/20/2011 20:27:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : MM030                                                                                      */
/*  Create Date    : 2011.02.20                                                                                 */
/*  Modified Date  : 품의작성시 원자재의 작번 Select																							*/
/*  Creator        : Noh Geun Yong																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_MM030_04]
--Create  PROC [dbo].[PS_MM030_04]
(
	@DocEtnry			As Nvarchar(10),
	@LineId		As Nvarchar(10)
)
AS

SET NOCOUNT ON
--BEGIN ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////					
--Create Table #Temp01 (
--	 BPLId		Nvarchar(1) collate Korean_Wansung_Unicode_CI_AS
--	,ItemCode	Nvarchar(20) collate Korean_Wansung_Unicode_CI_AS
--	,CardName	Nvarchar(100) 
--	,DocDate	DateTime
--	,Price		Numeric(19,6)
--)

select PP030H.U_OrdNum,
	   PP030H.U_OrdSub1,
	   PP030H.U_OrdSub2
  from [@PS_MM010L] MM010L,
       [@PS_MM005H] MM005H,
       [@PS_PP030H] PP030H
Where MM010L.U_CGNo = MM005H.DocEntry
  AND MM005H.U_PP030HNo = PP030H.DocEntry
  AND MM010L.DocEntry = @DocEtnry
  AND MM010L.U_LineNum = @LineId

--EXEC [PS_MM030_04] '272', '1'

--THE END //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
-- EXEC [PS_MM030_02] '1749'
--  select * from oitm
--     select * from [@PS_MM005h]
--   select * from [OPRC]
-- select U_ItmMSort From [OITM]








