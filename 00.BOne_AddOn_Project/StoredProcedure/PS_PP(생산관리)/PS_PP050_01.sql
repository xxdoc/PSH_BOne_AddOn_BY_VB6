USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP050_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : 공정별 진행현황																	*/
/*  Create Date    : 2010.11.29                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP050_01]
--Create PROC [dbo].[PS_PP050_01]
(
	@OrdNum NVARCHAR(30)
)
AS

BEGIN

SELECT  A.DocEntry,
        CONVERT(NVARCHAR(20),A.U_ItemCode)	AS	ItemCode,		--제품코드
		CONVERT(NVARCHAR(60),A.U_ItemName)	AS	ItemName,		--제품명
		CONVERT(NVARCHAR(10),A.U_OrdNum)		AS	OrdNum,			--관리번호
		BatchNum = (SELECT MAX(CONVERT(NVARCHAR(30),G.U_BatchNum)) FROM [@PS_PP030H] AS F INNER JOIN [@PS_PP030L] AS G
								ON F.DocEntry = G.DocEntry
					 WHERE F.U_OrdNum = A.U_OrdNum),	--거래처No
		PackNo = (SELECT MAX(CONVERT(NVARCHAR(30),I.U_PackNo)) FROM [@PS_PP090H]AS H INNER JOIN [@PS_PP090L] AS I
								ON H.DocEntry = I.DocEntry
				   WHERE I.U_ItemCode = A.U_ItemCode),	--패킹No
		Unit = (SELECT CONVERT(NVARCHAR(10),U_Unit2) FROM OITM WHERE ItemCode = A.U_ItemCode),    --단위
		Size	= (SELECT CONVERT(NVARCHAR(20),U_Size) FROM OITM WHERE ItemCode = A.U_ItemCode),	--규격
		CardName = (SELECT CONVERT(NVARCHAR(60),U_CardName) FROM [@PS_QM020H] WHERE U_OrdNum = A.U_OrdNum),   --납품처						--납품처	
		D.U_FailName	AS	FailName,		--불량사유	
		CONVERT(NVARCHAR(10),A.U_Sequence)	AS	Sequence,		--공정순위
		CONVERT(NVARCHAR(20),A.U_CpCode)	AS	CpCode,				--공정번호
		CONVERT(NVARCHAR(60),A.U_CpName)	AS	CpName,
		A.U_BQty	AS	BQty,				--인수량
		A.U_YQty	AS	YQty,				--인계량
		CONVERT(NVARCHAR(10),B.U_WorkName)	AS	WorkName,
		CONVERT(NUMERIC(19,2),A.U_WorkTime)	AS	WorkTime,		--공수
		C.U_DocDate	AS	DocDate,			--작업일자
		A.U_ScrapWt	AS	ScrapWt,			--스크랩중량
		A.U_NQty	AS	NQty,				--불량
		CONVERT(NVARCHAR(30),C.U_MoldCode)	AS MoldCode,		--금형번호
		CONVERT(NVARCHAR(50),C.U_UseMCode)	AS UseMCode			--기계장치번호 
FROM [@PS_PP040L] AS A INNER JOIN [@PS_PP040M] AS B
		ON A.DocEntry = B.DocEntry
	INNER JOIN [@PS_PP040H] AS C
		ON A.DocEntry = C.DocEntry
	INNER JOIN  [@PS_PP040N] AS D
		ON A.DocEntry = D.DocEntry AND
		   A.LineId   = D.LineId
WHERE A.U_OrdNum = @OrdNum	    
	
order by A.DocEntry

End	

--EXEC [dbo].[PS_PP050_01] '20101110001'