USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM060_01]    Script Date: 11/04/2010 12:52:43 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : 작업지시등록 > 라인 테이블 대/중분류 코드 UPDATE[PS_MM060]                                 */
/*  Create Date    : 2010.10.02                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_MM060_01]
(
@U_ItmMcode AS Nvarchar(20)
)
AS
BEGIN
 
INSERT  INTO [@PS_MM060M](Code, LineId, Object, LogInst, U_LineNum, U_LogNo, U_ItmBcode, A.U_ItmBname,
		      A.U_ItmMcode, U_ItmMname, U_Spec1, U_Price1, U_Price2, U_inDate)

	  Select A.Code, 
			 A.LineId + ISNULL(MAX(B.LineId), 0) AS LineId,
			 A.Object, 
			 A.LogInst, 
			 A.U_LineNum, 
			 ISNULL(B.U_LogNo, 0) + 1 As U_LogNo, 
			 A.U_ItmBcode, 
			 A.U_ItmBname,
			 A.U_ItmMcode,
			 A.U_ItmMname, 
			 A.U_Spec1, 
			 A.U_Price1, 
			 A.U_Price2, 
			 A.U_inDate
		From [@PS_MM060L] AS A Left Join
			 [@PS_MM060M] AS B
		  ON A.Code = B.Code  
	   WHERE A.U_ItmMcode = @U_ItmMcode  AND 
			 B.U_LogNo = (SELECT MAX(B.U_LogNo) 
							FROM [@PS_MM060L] AS A Left Join
								 [@PS_MM060M] AS B ON A.Code = B.Code  
						   WHERE A.U_ItmMcode = @U_ItmMcode
						Group By A.U_ItmMCode)
	Group By A.Code, A.LineId, A.Object, A.LogInst, A.U_LineNum, A.U_ItmBcode, A.U_ItmBname,
			 A.U_ItmMcode, A.U_ItmMname, A.U_Spec1, A.U_Price1, A.U_Price2, A.U_inDate, B.U_LogNo		
			   
 END
