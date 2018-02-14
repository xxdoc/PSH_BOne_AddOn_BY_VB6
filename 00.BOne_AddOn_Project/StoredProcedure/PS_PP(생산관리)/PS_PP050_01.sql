USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP050_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : ������ ������Ȳ																	*/
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
        CONVERT(NVARCHAR(20),A.U_ItemCode)	AS	ItemCode,		--��ǰ�ڵ�
		CONVERT(NVARCHAR(60),A.U_ItemName)	AS	ItemName,		--��ǰ��
		CONVERT(NVARCHAR(10),A.U_OrdNum)		AS	OrdNum,			--������ȣ
		BatchNum = (SELECT MAX(CONVERT(NVARCHAR(30),G.U_BatchNum)) FROM [@PS_PP030H] AS F INNER JOIN [@PS_PP030L] AS G
								ON F.DocEntry = G.DocEntry
					 WHERE F.U_OrdNum = A.U_OrdNum),	--�ŷ�óNo
		PackNo = (SELECT MAX(CONVERT(NVARCHAR(30),I.U_PackNo)) FROM [@PS_PP090H]AS H INNER JOIN [@PS_PP090L] AS I
								ON H.DocEntry = I.DocEntry
				   WHERE I.U_ItemCode = A.U_ItemCode),	--��ŷNo
		Unit = (SELECT CONVERT(NVARCHAR(10),U_Unit2) FROM OITM WHERE ItemCode = A.U_ItemCode),    --����
		Size	= (SELECT CONVERT(NVARCHAR(20),U_Size) FROM OITM WHERE ItemCode = A.U_ItemCode),	--�԰�
		CardName = (SELECT CONVERT(NVARCHAR(60),U_CardName) FROM [@PS_QM020H] WHERE U_OrdNum = A.U_OrdNum),   --��ǰó						--��ǰó	
		D.U_FailName	AS	FailName,		--�ҷ�����	
		CONVERT(NVARCHAR(10),A.U_Sequence)	AS	Sequence,		--��������
		CONVERT(NVARCHAR(20),A.U_CpCode)	AS	CpCode,				--������ȣ
		CONVERT(NVARCHAR(60),A.U_CpName)	AS	CpName,
		A.U_BQty	AS	BQty,				--�μ���
		A.U_YQty	AS	YQty,				--�ΰ跮
		CONVERT(NVARCHAR(10),B.U_WorkName)	AS	WorkName,
		CONVERT(NUMERIC(19,2),A.U_WorkTime)	AS	WorkTime,		--����
		C.U_DocDate	AS	DocDate,			--�۾�����
		A.U_ScrapWt	AS	ScrapWt,			--��ũ���߷�
		A.U_NQty	AS	NQty,				--�ҷ�
		CONVERT(NVARCHAR(30),C.U_MoldCode)	AS MoldCode,		--������ȣ
		CONVERT(NVARCHAR(50),C.U_UseMCode)	AS UseMCode			--�����ġ��ȣ 
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