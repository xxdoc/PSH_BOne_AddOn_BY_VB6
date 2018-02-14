USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP150_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/******************************************************************************************************************/
/*  Module         : PP								    														*/
/*  Description    : �����Ȳ(����)  															*/
/*  Create Date    : 2010.12.01                                                                                   */
/*  Modified Date  :										       													*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/******************************************************************************************************************/
--ALTER PROC [dbo].[PS_PP150_01]
Create PROC [dbo].[PS_PP150_01]
(
	@DocDateFr  Date,
	@DocDateTo  Date
)
AS

BEGIN

SELECT	(CONVERT(NVARCHAR(20),A.U_OrdNum) + '-' + CONVERT(NVARCHAR(10),A.U_OrdSub1)) AS OrdNum, 
		CONVERT(CHAR(60),A.U_ItemCode) AS ItemCode, 
		CONVERT(NVARCHAR(60),A.U_ItemName) AS ItemName,
		Selwt = (SELECT SUM(U_Selwt) FROM [@PS_PP030H] WHERE U_ItemCode = A.U_ItemCode),		--���ü���
		SUM(0) AS JISURANG,																		--�����߷�
		CONVERT(NVARCHAR(20),B.U_CpCode) AS CpCode,
        CpName = (SELECT CONVERT(NVARCHAR(50),MAX(U_CpBName)) FROM [@PS_PP030M] WHERE U_CpBCode = B.U_CpCode),
		B.U_Sequence AS Sequence,
		SUM(IsNull(Case When A.U_OrdType <> '30' Then IsNull(B.U_PQty, 0) End, 0)) As PQty,     --�������
		SUM(IsNull(Case When A.U_OrdType = '30' Then IsNull(B.U_PQty, 0) End, 0)) As OutQty,	--�۾�Ÿ��30�� ��� �������
		SUM(B.U_NQty) AS NQty,																	--�ҷ�����
	    SUM(B.U_PQty) AS GONGSENG,																--�������귮
	    A.U_DocDate	AS DocDate,
	    SUM(0) AS JEGONG,
	    SUM(0) AS SEOULDE,
	    SUM(0) AS SEOULJE
  FROM [@PS_PP040H] AS A INNER JOIN [@PS_PP040L] AS B
			ON A.DocEntry = B.DocEntry
 WHERE A.U_DocDate BETWEEN @DocDateFr AND @DocDateTo
			
GROUP BY  A.U_OrdNum, A.U_OrdSub1, A.U_ItemCode, A.U_ItemName, B.U_CpCode, B.U_Sequence, A.U_DocDate

Order BY  A.U_OrdNum, A.U_OrdSub1, A.U_ItemCode

End


--EXEC [dbo].[PS_PP150_01] '20101101','20101130'