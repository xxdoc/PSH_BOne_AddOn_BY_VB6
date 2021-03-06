USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_MM171_01]    Script Date: 11/04/2010 12:54:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****************************************************************************************************************/
/*  Module         : MM																							*/
/*  Description    : 지체상금등록 > 지체상금조회 INSERT[PS_MM171]			                                */
/*  Create Date    : 2010.09.10                                                                                 */
/*  Modified Date  :																							*/
/*  Creator        : Kim Dong sub																				*/
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
ALTER PROC [dbo].[PS_MM171_01]
(
	@CardCode NVARCHAR(50),
	@BPLId    NVARCHAR(50)
)
AS
BEGIN
 SELECT A.DocNum     AS GRDocNum,					--입고번호
        A.LineNum    AS GRLinNum,					--입고라인       
	    A.CardCode   AS CardCode,					--거래처코드
	    A.CardName   AS CardName,					--거래처명
	    A.ItemCode   AS ItemCode,					--아이템코드
	    A.Dscription AS ItemName,					--아이템명	
	    SUM(ISNULL(A.LineTotal, 0))  AS LinTotal,	--대상금액
	    A.DocDate    AS ImDate,						--입고일
	    A.DocDueDate AS DueDate,					--납기일
	    A.LateDay    AS LateDay,					--지체일수
	    CASE WHEN SUM(ISNULL(A.LateDay * 0.0015 * A.LineTotal, 0)) > 1000 THEN
	  			  SUM(ISNULL(A.LateDay * 0.0015 * A.LineTotal, 0))  
	   		     END AS RepayP,						--지체금액
	    A.U_ordTyp   AS Doctype,			     	--문서구분
	    A.FullName   AS CntcName,					--입고담당
	    A.U_BDocNUm  AS PODocNum,					--품의번호
	    A.BuyUnitMsr AS Unit,						--단위
	    A.U_Size     AS Size,						--규격
	    SUM(ISNULL(A.Quantity, 0)) AS Qty,			--수량
	    A.U_Unweight AS Unweight,					--단중	     				
	    A.U_ItmBsort AS ItmBsort,					--대분류
	    A.U_ItmMsort As ItmMsort,					--중분류
	    A.U_ItemType As ItemType,					--형태타입
		A.U_Quality  As Quality,					--질별
		A.U_Mark     AS Mark,						--인증기호
		A.U_CallSize As CallSize,					--호칭규격
		A.U_ObasUnit As ObasUnit					--매입기준단위	 
  FROM   
      (
		SELECT A.DocNum,     B.LineNum,   A.CardCode, A.CardName,   B.ItemCode,   
			   B.Dscription, B.LineTotal, A.DocDate,  A.DocDueDate, 
			   CONVERT(INT, A.DocDate) - CONVERT(INT, A.DocDueDate) AS LateDay, 
			   A.U_ordTyp,   D.lastName + D.firstName AS FullName,  E.U_BDocNum,
			   F.BuyUnitMsr,     F.U_Size,	  B.Quantity, F.U_UnWeight,
			   F.U_ItmBsort, F.U_ItmMsort, F.U_ItemType, F.U_Quality,
			   F.U_Mark, F.U_CallSize, F.U_ObasUnit 
			   --, G.Name, H.U_CodeName, I.Name, J.Name, K.Name, L.Name
		  FROM [OPDN]       A							   INNER JOIN		    -- 입고PO 헤더
			   [PDN1]       B ON A.DocEntry = B.DocEntry   LEFT JOIN			-- 입고PO 라인
			   [@PS_MM170L] C ON B.DocEntry = C.U_GRDocNum
							 AND B.LineNum  = C.U_GRLinNum LEFT JOIN			-- 재고이동 라인
			                     
			   [OHEM]       D ON A.CntctCode = D.empID     LEFT JOIN			-- 사원 마스터
			   [@PS_MM050L]	E ON B.DocEntry = E.U_PODocNum 	 					-- 품의 NO
							 AND B.LineNum  = E.U_POLinNum LEFT JOIN			                   								 
			   [OITM]       F ON B.ItemCode = F.ItemCode   --LEFT JOIN			-- 아이템 마스터
			   
		WHERE  NOT EXISTS ( SELECT 'X'
							 FROM [@PS_MM170L] 
							WHERE U_GRDocNum = B.DocEntry
							  AND U_GRLInNum = B.LineNum )
		  AND (@CardCode = '' OR A.CardCode = @CardCode)
		  AND (@BPLId = '' OR A.BPLId = @BPLId)                    
	  ) A	  
 WHERE A.LateDay BETWEEN '10' AND '20'
GROUP BY A.DocNum,   
         A.LineNum,   
	     A.CardCode,   
	     A.CardName,  
	     A.ItemCode,  
	     A.Dscription,	    
	     A.DocDate,
	     A.DocDueDate,
	     A.LateDay, 
	     A.U_ordTyp, 
	     A.FullName, 
	     A.U_BDocNum,
	     A.BuyUnitMsr,
	     A.U_Size,	    
	     A.U_Unweight,
	     A.U_ItmBsort, 
	     A.U_ItmMsort, 
	     A.U_ItemType,
		 A.U_Quality, 
		 A.U_Mark,    
		 A.U_CallSize,
		 A.U_ObasUnit 
 HAVING  SUM(ISNULL(A.LateDay * 0.0015 * A.LineTotal, 0)) > 1000
 
 UNION ALL
 
 SELECT A.DocNum     AS GRDocNum,					--입고번호
        A.LineNum    AS GRLinNum,					--입고라인       
	    A.CardCode   AS CardCode,					--거래처코드
	    A.CardName   AS CardName,					--거래처명
	    A.ItemCode   AS ItemCode,					--아이템코드
	    A.Dscription AS ItemName,					--아이템명	
	    SUM(ISNULL(A.LineTotal, 0))  AS LinTotal,	--대상금액
	    A.DocDate    AS ImDate,						--입고일
	    A.DocDueDate AS DueDate,					--납기일
	    A.LateDay    AS LateDay,					--지체일수
	    CASE WHEN SUM(ISNULL(A.LateDay * 0.0015 * A.LineTotal, 0)) > 100 THEN
	  			  SUM(ISNULL(A.LateDay * 0.0015 * A.LineTotal, 0))  
	   		     END AS RepayP,						--지체금액
	    A.U_ordTyp   AS Doctype,			     	--문서구분
	    A.FullName   AS CntcName,					--입고담당
	    A.U_BDocNUm  AS PODocNum,					--품의번호
	    A.BuyUnitMsr AS Unit,						--단위
	    A.U_Size     AS Size,						--규격
	    SUM(ISNULL(A.Quantity, 0)) AS Qty,			--수량
	    A.U_Unweight AS Unweight,					--단중
	    A.U_ItmBsort AS ItmBsort,					--대분류
	    A.U_ItmMsort As ItmMsort,					--중분류
	    A.U_ItemType As ItemType,					--형태타입
		A.U_Quality  As Quality,					--질별
		A.U_Mark     AS Mark,						--인증기호
		A.U_CallSize As CallSize,					--호칭규격
		A.U_ObasUnit As ObasUnit					--매입기준단위	   				
  FROM 
  
      (
		SELECT A.DocNum,     B.LineNum,   A.CardCode, A.CardName,   B.ItemCode,   
			   B.Dscription, B.LineTotal, A.DocDate,  A.DocDueDate, 
			   CONVERT(INT, A.DocDate) - CONVERT(INT, A.DocDueDate) AS LateDay, 
			   A.U_ordTyp,   D.lastName + D.firstName AS FullName,  E.U_BDocNum,
			   F.BuyUnitMsr,     F.U_Size,	  B.Quantity, F.U_UnWeight, 
			   F.U_ItmBsort, F.U_ItmMsort, F.U_ItemType, F.U_Quality,
			   F.U_Mark, F.U_CallSize, F.U_ObasUnit
		  FROM [OPDN]       A							   INNER JOIN		    -- 입고PO 헤더
			   [PDN1]       B ON A.DocEntry = B.DocEntry   LEFT JOIN			-- 입고PO 라인
			   [@PS_MM170L] C ON B.DocEntry = C.U_GRDocNum
							 AND B.LineNum  = C.U_GRLinNum LEFT JOIN			-- 재고이동 라인
			                     
			   [OHEM]       D ON A.CntctCode = D.empID     LEFT JOIN			-- 사원 마스터
			   [@PS_MM050L]	E ON B.DocEntry = E.U_PODocNum 	 					-- 품의 NO
							 AND B.LineNum  = E.U_POLinNum LEFT JOIN			                   								 
			   [OITM]       F ON B.ItemCode = F.ItemCode						-- 아이템 마스터
			   
		WHERE  NOT EXISTS ( SELECT 'X'
							 FROM [@PS_MM170L] 
							WHERE U_GRDocNum = B.DocEntry
							  AND U_GRLInNum = B.LineNum )
		  AND (@CardCode = '' OR A.CardCode = @CardCode)
		  AND (@BPLId = '' OR A.BPLId = @BPLId)                    
	  ) A	
	    
 WHERE A.LateDay > '20'
GROUP BY A.DocNum,   
         A.LineNum,   
	     A.CardCode,   
	     A.CardName,  
	     A.ItemCode,  
	     A.Dscription,	    
	     A.DocDate,
	     A.DocDueDate,
	     A.LateDay, 
	     A.U_ordTyp, 
	     A.FullName, 
	     A.U_BDocNum,
	     A.BuyUnitMsr,
	     A.U_Size,	    
	     A.U_Unweight,
	     A.U_ItmBsort, 
	     A.U_ItmMsort, 
	     A.U_ItemType,
		 A.U_Quality, 
		 A.U_Mark,    
		 A.U_CallSize,
		 A.U_ObasUnit    
 HAVING  SUM(ISNULL(A.LateDay * 0.0015 * A.LineTotal, 0)) > 100

END  
	
-- EXEC [PS_MM171_01] '', ''	