USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP035_04]    Script Date: 11/09/2010 16:08:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : PP																							*/
/*  Description    : �����۾����ü�																				*/
/*  Create Date    : 2010.11.10                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Lee Byong Gak																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--EXEC [PS_PP035_04]
ALTER PROC [dbo].[PS_PP035_04]
--Create PROC [dbo].[PS_PP035_04]
(
	@Seq	As Nvarchar(1)
)
AS
BEGIN
IF OBJECT_ID('Temp_LBG11') IS NULL
	BEGIN
		CREATE TABLE [Temp_LBG11]
		( 
		DocEntry  numeric
		)
	END	
If @Seq = 'M'
	
	SELECT 
	       CONVERT(NVARCHAR(20),(A.U_OrdNum + A.U_OrdSub1 + A.U_OrdSub2)) AS JAKBUN,                                       --�۹�
		   CONVERT(NUMERIC(19,0),A.DocEntry) AS DocEntry,
		   CONVERT(NVARCHAR(30),B.FrgnName) AS FrgnName,
		   CONVERT(NVARCHAR(50),B.U_SIZE) AS SIZE,
		   CONVERT(NUMERIC(19,6),FLOOR(A.U_SelWt)) AS SELWT, 
		   CONVERT(NVARCHAR(8),A.U_DueDate,112) AS DATE,
		   CONVERT(NVARCHAR(50),C.CardName) AS CardName,
		   FLOOR((Select D.Quantity from RDR1 D WHERE A.U_SjNum = D.DocEntry And A.U_SjLine = D.LineNum)) AS Quantity,     --���ּ���
		   FLOOR((Select D.TotalSumSy from RDR1 D WHERE A.U_SjNum = D.DocEntry And A.U_SjLine = D.LineNum)) AS TotalSumSy, --���ֱݾ�
		   Convert(char(8),A.U_DocDate,112) AS DocDate,
		   FLOOR(A.U_SelWt) AS Selwt,                                                                                      --�������ü���
		   CONVERT(CHAR(8),A.U_DueDate,112) AS DueDate,
		   CONVERT(nvarchar(100),A.U_Comments) AS Comments                                                                 --Ư�����
		   --E.LineId AS LineId,
		   --CONVERT(NVARCHAR(20),E.U_CpBName) AS CpBName,
		   --CONVERT(NVARCHAR(20),E.U_CpName) AS CpName,
		   --CONVERT(nvarchar(10),E.U_Unit) As Unit,
		   --E.U_StdHour AS StdHour,                                                                    
		   --FLOOR((U_StdHour * (select U_Price From [@PS_PP001L] where U_CpCode = E.U_CpCode))) AS AMT,                     --ǥ�رݾ�
		   --CONVERT(CHAR(8),E.U_ReDate,112) AS ReDate,
	  FROM [@PS_PP030H] AS A INNER JOIN OITM AS B
		   ON A.U_ItemCode = B.ITEMCODE
		   inner join ORDR AS C
		   on A.U_SjNum = C.DocEntry
		   --inner join [@PS_PP030M] AS E
		   --ON A.DocEntry = E.DocEntry
	 WHERE A.DocEntry in (select DocEntry FROM Temp_LBG11)
	 
If @Seq = 'S'
	SELECT CONVERT(NUMERIC(19,0),A.DocEntry) AS DocEntry,
	       A.LineId AS LineId,
		   CONVERT(NVARCHAR(50),B.U_SIZE) AS SIZE,
		   CONVERT(NVARCHAR(150),ItemCode + '/' + ItemName) AS Name,		   
	       --CONVERT(NVARCHAR(100),A.U_ItemCode)  AS ItemCode,            --�����ڵ�
	       --CONVERT(NVARCHAR(50),(A.U_ItemName)) AS ItemName,            --�����ڵ��
	       CONVERT(NVARCHAR(50),A.U_CntcCode)    AS CntcCode,             --û����
           CONVERT(NVARCHAR(10),B.BuyUnitMsr)    AS BuyUnitMsr,           --����
           CONVERT(FLOAT,A.U_Weight)             AS Weight,               --����
           CONVERT(NVARCHAR(10),A.U_CntcName)    AS CntcName,
           CONVERT(NVARCHAR(10),A.U_ProcType)    AS ProcType,             --����(���޹��) 
     CASE WHEN U_ProcType = '10'
          THEN  'û��'
          WHEN U_ProcType = '20'
          THEN  '����'
          WHEN U_ProcType = '30'
          THEN '���'
     END AS ProcType     
	  FROM [@PS_PP030L] AS A INNER JOIN OITM AS B
		   ON A.U_ItemCode = B.ITEMCODE
		      INNER JOIN [@PS_PP030H] AS C
		   ON A.DocEntry = C.DocEntry
	 WHERE A.DocEntry in (select DocEntry FROM Temp_LBG11)	

If @Seq = 'E'
    SELECT A.DocEntry,
           E.LineId AS LineId,
		   CONVERT(NVARCHAR(20),E.U_CpBName) AS CpBName,
		   CONVERT(NVARCHAR(20),E.U_CpName) AS CpName,
		   CONVERT(nvarchar(10),E.U_Unit) As Unit,
		   E.U_StdHour AS StdHour,                                                                    
		   FLOOR((U_StdHour * (select U_Price From [@PS_PP001L] where U_CpCode = E.U_CpCode))) AS AMT,                     --ǥ�رݾ�
		   CONVERT(CHAR(8),E.U_ReDate,112) AS ReDate
	  FROM [@PS_PP030H] AS A INNER JOIN [@PS_PP030M] AS E
		   ON A.DocEntry = E.DocEntry
	 WHERE A.DocEntry in (select DocEntry FROM Temp_LBG11)

END

--  EXEC [PS_PP035_04] 'S'
--  EXEC [PS_PP035_04] 'M'
--  EXEC [PS_PP035_04] 'E'    
