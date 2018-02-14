--PH_PY301_자녀 조회


SELECT		T0.U_FamNam AS [성명],
				LEFT(T0.U_FamPer, 6) + '-' + RIGHT(T0.U_FamPer, 8) AS [주민등록번호],
				T1.U_CodeNm AS [관계],
				CASE
					WHEN RIGHT(LEFT(T0.U_FamPer, 7), 1) = '1' OR RIGHT(LEFT(T0.U_FamPer, 7), 1) = '3' OR RIGHT(LEFT(T0.U_FamPer, 7), 1) = '5' THEN '남'
					WHEN RIGHT(LEFT(T0.U_FamPer, 7), 1) = '2' OR RIGHT(LEFT(T0.U_FamPer, 7), 1) = '4' OR RIGHT(LEFT(T0.U_FamPer, 7), 1) = '6' THEN '여'
				END AS [성별]
FROM			[@PH_PY001D] AS T0
				LEFT JOIN
				[@PS_HR200L] AS T1
					ON T0.U_FamGun = T1.U_Code
					AND T1.Code = 'P121'
					AND T1.U_UseYN = 'Y'
WHERE		T0.Code = $[@PH_PY301A.U_CntcCode.0]
				AND T0.U_FamGun = '03' --본인과의 관계가 "자"인 정보만 조회
				
