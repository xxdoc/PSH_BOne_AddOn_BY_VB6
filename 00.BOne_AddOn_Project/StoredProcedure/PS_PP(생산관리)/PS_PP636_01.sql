SELECT B.SWW AS SWW,
       CASE C.ItmsGrpCod WHEN '104' THEN '¿øÀç·á' END,
       B.ItemCode	AS ItemCode,
       B.Dscription AS Dscription,
       B.DocDate    AS DocDate
       
  FROM OIGN AS A INNER JOIN IGN1 AS B
			ON A.DocEntry = B.DocEntry
		 INNER JOIN OITM AS C
			ON B.ItemCode = C.ItemCode
 WHERE not exists ( select * 
                  from IGE1 AS G
				 where	G.ItemCode = B.ItemCode )
   AND C.ItmsGrpCod = '104'		
	
Group by B.SWW, B.ItemCode, C.ItmsGrpCod, B.Dscription, B.DocDate
	
Order by B.ItemCode, B.DocDate


SELECT * FROM OITM
WHERE U_ItmBsort = '104'