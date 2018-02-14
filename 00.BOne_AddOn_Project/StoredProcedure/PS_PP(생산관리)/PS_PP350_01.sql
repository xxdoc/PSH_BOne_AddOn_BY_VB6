SELECT	A.U_ItemCode, 
		A.U_ItemName, 
		B.U_CpCode, 
		B.U_CpName,
		C.U_WorkCode,
		C.U_WorkName,
		(B.U_YQty + B.U_NQty) AS INSU,
		B.U_YQty,
		B.U_NQty,
		B.U_WorkTime
  FROM [@PS_PP040H] AS A INNER JOIN [@PS_PP040L] AS B
        ON A.DocEntry = B.DocEntry
  INNER JOIN [@PS_PP040M] AS C
        ON B.DocEntry = C.DocEntry
        
        
 SELECT	CONVERT(CHAR(20),A.U_ItemCode)		AS ItemCode, 
		CONVERT(CHAR(60),A.U_ItemName)		AS ItemName, 
		A.U_DocDate			AS DocDate,
		Mark = (SELECT U_Mark FROM OITM WHERE ItemCode = A.U_ItemCode),
		CONVERT(CHAR(20),B.U_CpCode)			AS CpCode, 
		CONVERT(CHAR(60),B.U_CpName)			AS CpName,
		CONVERT(CHAR(10),C.U_WorkCode)		AS WorkCode,
		CONVERT(CHAR(10),C.U_WorkName)		AS WorkName,
		(B.U_YQty + B.U_NQty) AS INSU,
		CONVERT(NUMERIC(19,6),B.U_YQty)			AS YQty,
		CONVERT(NUMERIC(19,6),B.U_NQty)			AS NQty,
		CONVERT(NUMERIC(19,6),B.U_WorkTime)		AS WorkTime
  FROM [@PS_PP040H] AS A INNER JOIN [@PS_PP040L] AS B
        ON A.DocEntry = B.DocEntry
	INNER JOIN [@PS_PP040M] AS C
        ON B.DocEntry = C.DocEntry
        ) G
    INNER JOIN [@PSH_MARK] AS D
		ON G.Mark = D.Code
  WHERE G.DocDate BETWEEN @DocDateFr AND @DocDateTo

ORDER BY G.ItemCode, G.CpCode
        