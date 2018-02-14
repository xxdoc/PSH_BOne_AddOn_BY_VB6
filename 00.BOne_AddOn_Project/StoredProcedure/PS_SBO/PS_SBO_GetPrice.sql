IF OBJECT_ID('PS_SBO_GETPRICE') IS NOT NULL
BEGIN
	DROP PROC PS_SBO_GETPRICE
END
GO
CREATE PROC PS_SBO_GETPRICE
(
	@CardCode NVARCHAR(100),
	@ItemCode NVARCHAR(100)
)
AS
BEGIN
	IF(SELECT U_ItmBsort FROM [OITM] WHERE ItemCode = @ItemCode) = '101' --휘팅의경우
	BEGIN
		IF --고객의 대분류와 소분류가 일치하는것이 존재할경우
			(SELECT COUNT(*)
			FROM [@PS_MM065H] T1 LEFT JOIN [@PS_MM065L] T2 ON T1.DocEntry = T2.DocEntry
			WHERE 
				T1.U_CardCode = @CardCode
				AND T2.U_ItmBcode = (SELECT U_ItmBsort FROM OITM WHERE ItemCode=@ItemCode) 
				AND T2.U_ItmMcode = (SELECT U_ItmMsort FROM OITM WHERE ItemCode=@ItemCode)) = 1 
		BEGIN
			SELECT 
				CASE WHEN X1.DcGbn='1' 
				THEN ROUND( PRICE *  ( (100-CustDC)/100 * ( (100-CashDC)/100 ) ) ,0)
				WHEN X1.DcGbn='2' 
				THEN(CASE WHEN CashDC=0 THEN ROUND((PRICE + CustDC),0) WHEN CashDC>0 THEN ROUND( ( (100+CustDC)/100 * ( (100+CashDC)/100 ) ) ,0)END)
				END AS S_PRICE  
			FROM 
				(SELECT	
					Case When T1.U_ItmMcode='10101' then  (CASE WHEN T0.U_TikAngle=90 THEN  T1.U_price1 ELSE T1.U_Price2 END) --엘보우의 경우
					WHEN T1.U_ItmMcode='10107' then  (CASE WHEN T0.U_TikAngle=10 THEN  T1.U_price1 ELSE T1.U_Price2 END) --후랜지의 경우
					else T1.U_Price1 end as PRICE
					,T2.U_DcGbn AS DcGbn	-- 1:%  2:원 
					,T2.U_CustDC AS CustDc	--할인적용율
					,T2.U_CashDC AS CashDC	--현금할인액
				FROM 
					OITM T0
					INNER JOIN [@PS_MM060L] T1 
					ON T0.U_ItmBsort=T1.U_ItmBcode AND T0.U_ItmMsort=T1.U_ItmMcode
					--규격이 현재규격보다작은것들중에서 제일큰것
					AND T1.U_Spec1=
					(SELECT 
						MAX(CONVERT(FLOAT,Tx.U_Spec1)) 
					FROM 
						[@PS_MM060L] Tx 
					WHERE 
						Tx.U_ItmBcode= T0.U_ItmBsort AND Tx.U_ItmMcode=T0.U_ItmMsort and Convert(float,U_Spec1) <= Convert(float,T0.U_Spec1)
					)
					INNER JOIN [@PS_MM065L] T2 ON T1.U_ItmBcode=T2.U_ItmBcode AND T1.U_ItmMcode=T2.U_ItmMcode
					LEFT JOIN [@PS_MM065H] T3 ON T3.DocEntry = T2.DocEntry
				WHERE 
					T0.ItemCode = @ItemCode --현재품목
					AND T3.U_CardCode = @CardCode
					AND T2.U_CustDC > 0 --할인적용율이 0이상일때만 가져옴
					AND T2.U_ItmBcode = (SELECT U_ItmBsort FROM OITM WHERE ItemCode=@ItemCode) 
					AND T2.U_ItmMcode = (SELECT U_ItmMsort FROM OITM WHERE ItemCode=@ItemCode)
				) X1
		END
		ELSE IF --고객의 소분류에 공통이 존재
			(SELECT COUNT(*)
			FROM [@PS_MM065H] T1 LEFT JOIN [@PS_MM065L] T2 ON T1.DocEntry = T2.DocEntry
			WHERE 
				T1.U_CardCode = @CardCode --AND T2.U_CardName='일흥산업(주)'
				AND T2.U_ItmBcode = (SELECT U_ItmBsort FROM OITM WHERE ItemCode=@ItemCode) 
				AND RIGHT(T2.U_ItmMcode,2) = '00') = 1 
		BEGIN
			SELECT 
				CASE WHEN X1.DcGbn='1' 
				THEN ROUND( PRICE *  ( (100-CustDC)/100 * ( (100-CashDC)/100 ) ) ,0)
				WHEN X1.DcGbn='2' 
				THEN(CASE WHEN CashDC=0 THEN ROUND((PRICE + CustDC),0) WHEN CashDC>0 THEN ROUND( ( (100+CustDC)/100 * ( (100+CashDC)/100 ) ) ,0)END)
				END AS S_PRICE  
			FROM 
				(SELECT	
					Case When T1.U_ItmMcode='10101' then  (CASE WHEN T0.U_TikAngle=90 THEN  T1.U_price1 ELSE T1.U_Price2 END) --엘보우의 경우
					WHEN T1.U_ItmMcode='10107' then  (CASE WHEN T0.U_TikAngle=10 THEN  T1.U_price1 ELSE T1.U_Price2 END) --후랜지의 경우
					else T1.U_Price1 end as PRICE
					,T2.U_DcGbn AS DcGbn	-- 1:%  2:원 
					,T2.U_CustDC AS CustDc	--할인적용율
					,T2.U_CashDC AS CashDC	--현금할인액
				FROM 
					OITM T0
					INNER JOIN [@PS_MM060L] T1 
					ON T0.U_ItmBsort=T1.U_ItmBcode AND T0.U_ItmMsort=T1.U_ItmMcode
					--규격이 현재규격보다작은것들중에서 제일큰것
					AND T1.U_Spec1=
					(SELECT 
						MAX(CONVERT(FLOAT,Tx.U_Spec1)) 
					FROM 
						[@PS_MM060L] Tx 
					WHERE 
						Tx.U_ItmBcode= T0.U_ItmBsort AND Tx.U_ItmMcode=T0.U_ItmMsort and Convert(float,U_Spec1) <= Convert(float,T0.U_Spec1)
					)
					INNER JOIN [@PS_MM065L] T2 ON T0.U_ItmBsort=T2.U_ItmBcode AND T2.U_ItmMcode = (CASE WHEN RIGHT(T2.U_ItmMcode,2) = '00' THEN T2.U_ItmMcode ELSE T0.U_ItmMsort END)
					LEFT JOIN [@PS_MM065H] T3 ON T3.DocEntry = T2.DocEntry
				WHERE 
					T0.ItemCode = @ItemCode --현재품목
					AND T3.U_CardCode = @CardCode
					AND T2.U_CustDC > 0 --할인적용율이 0이상일때만 가져옴
					AND T2.U_ItmBcode = (SELECT U_ItmBsort FROM OITM WHERE ItemCode=@ItemCode) 
					AND RIGHT(T2.U_ItmMcode,2) = '00'
				) X1
		END
		ELSE --고객에 해당대분류,소분류의 할인이 존재하지 않음
		BEGIN
			SELECT 0
		END
	END
	ELSE
	BEGIN
		SELECT 0
	END
END