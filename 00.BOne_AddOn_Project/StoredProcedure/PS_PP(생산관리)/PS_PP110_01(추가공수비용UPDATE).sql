USE [PSHDB_NOH]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP110_01]    Script Date: 04/22/2011 08:09:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 추가공수 및 금액 조회 [PS_PP110]                                                                        */
/*  Create Date    : 2011.04.22                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP110_01]
--Create  PROC [dbo].[PS_PP110_01]
(
	@YYMM		As Char(6)
)
AS

SET NOCOUNT ON
BEGIN --////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
select t.YYMM,
	   t.POEntry,
	   t.POLine,
	   t1.U_OrdNum As OrdNum,
	   t1.U_OrdSub1 As OrdSub1,
	   t1.U_OrdSub2 As OrdSub2,
	   t2.FrgnName As ItemName,
	   t2.U_Size As Size,
	   t.CpCode,
	   t.CpName,
	   t.Inval,
	   t.ReqVal
 from Z_PS_CO130B t Inner Join [@PS_PP030H] t1 On t.POEntry = t1.DocEntry
							Inner Join [OITM] t2 On t1.U_ItemCode = t2.ItemCode
 Where t.YYMM = @YYMM
 
END --//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

-- EXEC [PS_PP110_01] '201103'
-- select * from [@PS_MM070L]
-- update [@PS_MM030H] set U_POStatus = 'N' where docentry = 1
--    select *, U_Weight from rdr1

--   select * from IGE1











