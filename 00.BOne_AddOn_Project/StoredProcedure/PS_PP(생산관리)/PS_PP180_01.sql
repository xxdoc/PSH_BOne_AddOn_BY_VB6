USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP180_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 본사 휘팅 출하 대기 자료																				*/
/*  Create Date    : 2010.11.22                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP180_01]
--Create PROC [dbo].[PS_PP180_01]
(
	@MovDocNo Nvarchar(8)
)
AS

BEGIN

--기준날짜보다 이후 날짜 또는 Null 값을 가져온다.

SELECT A.DocEntry   AS DocEntry, 
	   CONVERT(NVARCHAR(20),B.U_ItemCode) AS ItemCode, 
	   CONVERT(NVARCHAR(50),B.U_ItemName) AS ItemName, 
	   CONVERT(NVARCHAR(10),B.U_OrdNum) AS OrdNum,
	   CardCode = (select CardCode FROM ORDR WHERE DocNum = C.U_SjNum),
	   CdarName = (select CardName FROM ORDR WHERE DocNum = C.U_SjNum),
	   B.U_SelQty  AS SelQty,
	   B.U_SelWt   AS SelWt
  FROM [@PS_PP070H] AS A INNER JOIN [@PS_PP070L] AS B
       ON A.DocEntry  = B.DocEntry
  INNER JOIN [@PS_PP030H] AS C
       ON A.DocEntry = C.DocEntry
  WHERE ISNULL(B.U_MovDocNo,'') = '' 
     OR LEFT(B.U_MovDocNo,8) > @MovDocNo 

End


--EXEC [dbo].[PS_PP180_01] '20101230'

  --where not exists (select * from [@PS_PP075H] D, [@PS_PP075L]  E
		--			  where D.DocEntry = E.DocEntry
		--			    AND  B.DocEntry = left(E.U_PP070No,1)
		--				and B.LineId = RIGHT(E.U_PP070No,1)
		--				and D.CreateDate >= '20101116')