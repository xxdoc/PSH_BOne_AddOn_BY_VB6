SELECT  A.DocEntry,
        A.U_ItemCode,
		A.U_ItemName,
		A.U_CardName,
		A.U_LotNo,
		CallSize = ( SELECT U_CallSize FROM OITM WHERE ItemCode = A.U_ItemCode AND U_ItmBsort = '104' )
		--OrdNum = ( SELECT U_OrdNum FROM OITM WHERE ItemCode = A.U_ItemCode AND U_ItmBsort = '104' )
  FROM [@PS_QM020H] AS A
  WHERE NOT EXISTS ( SELECT * FROM [@PS_PP090L] G INNER JOIN [@PS_PP090H] AS H
								ON G.DocEntry = H.DocEntry
					  WHERE A.U_LotNo = G.U_LotNo )
					    --AND B.U_BPLId = H.U_BPLId )
