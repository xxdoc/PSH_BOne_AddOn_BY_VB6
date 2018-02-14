USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP170_01]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : ���� ���ǰ ������ ��Ȳ                													*/
/*  Create Date    : 2010.11.26                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_PP170_01]
--Create PROC [dbo].[PS_PP170_01]
(
	@DocDate Date
)
AS

BEGIN

	SELECT	CONVERT(NVARCHAR(30),A.U_ItemCode)  AS ItemCode,
			CONVERT(NVARCHAR(60),A.U_ItemName)  AS ItemName,
			CONVERT(NVARCHAR(20),A.U_MovDocNo)  AS MovDocNo,							--���Ϲ�ȣ
			CONVERT(NVARCHAR(20),A.U_PorNum)	  AS PorNum,							--�����ȣ
		    Size = (SELECT CONVERT(NVARCHAR(20),U_Size) FROM OITM WHERE ItemCode = A.U_ItemCode),			--�԰�
		    CallSize = (SELECT CONVERT(NVARCHAR(20),U_CallSize) FROM OITM WHERE ItemCode = A.U_ItemCode),	--ȣĪ�԰�
			ItemType = (SELECT MAX(Name) FROM [@PSH_SHAPE]AS A INNER JOIN OITM								--����ǥ��
							ON Code = U_ItemType),
			Mark = (SELECT MAX(NAME) FROM [@PSH_MARK] AS G INNER JOIN OITM AS H 
						ON G.CODE = H.U_Mark),
			A.U_PkQty AS PkQty,															--�̵�����
			A.U_PkWt  AS PkWt,															--�̵��߷�
			A.U_NPkQty AS NPkQty,
			isnull(A.U_NPkWt,0)  AS NPkWt,
			(ISNULL(A.U_PkQty,0) - ISNULL(A.U_PkQty,0)) AS JANSU,
			(ISNULL(A.U_PkWt,0)  - ISNULL(A.U_NPkWt,0)) AS JANJUNG
     FROM [@PS_PP077H] AS A
	WHERE isnull(A.U_NPkQty,0) > 0
	  AND A.Canceled <> 'Y'
	  AND A.U_InDate > @DocDate
	  
END	  

--EXEC PS_PP170_01 '20101110'