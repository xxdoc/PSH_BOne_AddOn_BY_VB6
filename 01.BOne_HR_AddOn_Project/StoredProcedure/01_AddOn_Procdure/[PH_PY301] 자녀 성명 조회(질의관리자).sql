--PH_PY301_�ڳ� ��ȸ


SELECT		T0.U_FamNam AS [����],
				LEFT(T0.U_FamPer, 6) + '-' + RIGHT(T0.U_FamPer, 8) AS [�ֹε�Ϲ�ȣ],
				T1.U_CodeNm AS [����],
				CASE
					WHEN RIGHT(LEFT(T0.U_FamPer, 7), 1) = '1' OR RIGHT(LEFT(T0.U_FamPer, 7), 1) = '3' OR RIGHT(LEFT(T0.U_FamPer, 7), 1) = '5' THEN '��'
					WHEN RIGHT(LEFT(T0.U_FamPer, 7), 1) = '2' OR RIGHT(LEFT(T0.U_FamPer, 7), 1) = '4' OR RIGHT(LEFT(T0.U_FamPer, 7), 1) = '6' THEN '��'
				END AS [����]
FROM			[@PH_PY001D] AS T0
				LEFT JOIN
				[@PS_HR200L] AS T1
					ON T0.U_FamGun = T1.U_Code
					AND T1.Code = 'P121'
					AND T1.U_UseYN = 'Y'
WHERE		T0.Code = $[@PH_PY301A.U_CntcCode.0]
				AND T0.U_FamGun = '03' --���ΰ��� ���谡 "��"�� ������ ��ȸ
				
