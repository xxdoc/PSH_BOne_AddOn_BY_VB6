DECLARE @MDC_DOCENTRY		INT
DECLARE @MDC_DATE			VARCHAR(8)
DECLARE @MDC_TIME			VARCHAR(4)

DELETE [@PH_SY001H] --WHERE CODE IN('CSY001','CSY002','CSY003')
DELETE [@PH_SY001L] --WHERE CODE IN('CSY001','CSY002','CSY003')

SELECT @MDC_DATE = CONVERT(VARCHAR(8),GETDATE(),112)
SELECT @MDC_TIME = SUBSTRING(CONVERT(VARCHAR(8),GETDATE(),108),1,2)+SUBSTRING(CONVERT(VARCHAR(8),GETDATE(),108),4,2)


--select * from [@PH_SY001L]

----------------------------------------------------------------------------------------------------------------------------------------------------------------
SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
/* 헤더 */
INSERT INTO [@PH_SY001H] VALUES(N'CSY001',	N'Company Info', @MDC_DOCENTRY,	'N',			N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'Company Connection Info')
/* 라인 */
INSERT INTO [@PH_SY001L] VALUES(N'CSY001',	1,	N'PH_SY001'	,NULL,	'1',	N'PATH',		N'DB_Info_YN,PATH,Screen,Report',					N'N',N'C:\Users\LAJOLLA\Desktop\VB버전템플릿(공통코드추가)',N'HR_Screen',N'HR_Report',N'')
INSERT INTO [@PH_SY001L] VALUES(N'CSY001',	2,	N'PH_SY001'	,NULL,	'2',	N'ODBC',		N'ODBC_YN,ODBC_NAME,ODBC_DBNAME,ID,PW',				N'Y',N'MDCERP',N'PSH_HR',N'sa',N'password1!')
INSERT INTO [@PH_SY001L] VALUES(N'CSY001',	3,	N'PH_SY001'	,NULL,	'3',	N'NETWORK',		N'NETWORK_YN,NETWORK_DRIVE,NETWORK_PATH,ID,PW',		N'N',N'Q:',N'\\191.1.1.220\B1_SHR\PathINI',N'administrator',N'password1!')
----------------------------------------------------------------------------------------------------------------------------------------------------------------
SELECT * FROM [@PH_SY001H] 
SELECT * FROM [@PH_SY001L] 
----UPDATE [@CSY001L] SET U_Value01 = 'D:\Moring_Project\EAGON\02_AddOnSource\PathINI'
--WHERE Code = 'CSY001' and LineId = '1'


--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'CSY002',	N'MES InterFace Info', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'MES Connection Info')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'CSY002',	1,	N'PH_SY001'	,NULL,	'1',	N'MES_IF_YN',	N'MES I/F YN',					N'N',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'CSY002',	2,	N'PH_SY001'	,NULL,	'2',	N'MES_IP',		N'MES I/F Server IP',			N'192.168.100.42',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'CSY002',	3,	N'PH_SY001'	,NULL,	'3',	N'MES_DOMAIN',	N'MES I/F Server Name',			N'MES',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'CSY002',	4,	N'PH_SY001'	,NULL,	'4',	N'MES_DBNAME',	N'MES I/F DB Name',				N'INTERFACE',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'CSY002',	5,	N'PH_SY001'	,NULL,	'5',	N'MES_SQL_ID',	N'MES I/F SQL Connect ID/PW',	N'mesadmin',N'mes!@34',N'',N'',N'')
----------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'CSY003',	N'G/W InterFace Info', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'G/W Connection Info')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'CSY003',	1,	N'PH_SY001'	,NULL,	'1',	N'GW_IF_YN',	N'GW I/F YN',					N'N',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'CSY003',	2,	N'PH_SY001'	,NULL,	'2',	N'GW_IP',		N'GW I/F Server IP',			N'192.168.100.35',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'CSY003',	3,	N'PH_SY001'	,NULL,	'3',	N'GW_DOMAIN',	N'GW I/F Server Name',			N'GW_DB',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'CSY003',	4,	N'PH_SY001'	,NULL,	'4',	N'GW_DBNAME',	N'GW I/F DB Name',				N'INTERFACE',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'CSY003',	5,	N'PH_SY001'	,NULL,	'5',	N'GW_SQL_ID',	N'GW I/F SQL Connect ID',		N'CSYgw',N'!@12sql',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'EB001',	N'사업구분', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'사업구분')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'EB001',	1,	N'PH_SY001'	,NULL,	'1',	N'01',	N'시스템',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'EB001',	2,	N'PH_SY001'	,NULL,	'2',	N'02',	N'커튼월',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'EB001',	3,	N'PH_SY001'	,NULL,	'3',	N'03',	N'마루',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'EB001',	4,	N'PH_SY001'	,NULL,	'4',	N'04',	N'합판',				N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'EB002',	N'품목구분', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'품목구분')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'EB002',	1,	N'PH_SY001'	,NULL,	'1',	N'01',	N'제품',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'EB002',	2,	N'PH_SY001'	,NULL,	'2',	N'02',	N'원자재',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'EB002',	3,	N'PH_SY001'	,NULL,	'3',	N'03',	N'반제품',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'EB002',	4,	N'PH_SY001'	,NULL,	'4',	N'04',	N'상품',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'EB002',	5,	N'PH_SY001'	,NULL,	'5',	N'05',	N'부자재',				N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'MM003',	N'자재구분', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'자재구분')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'MM003',	1,	N'PH_SY001'	,NULL,	'1',	N'P',	N'PVC',					N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM003',	2,	N'PH_SY001'	,NULL,	'2',	N'5',	N'목재',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM003',	3,	N'PH_SY001'	,NULL,	'3',	N'6',	N'알미늄',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM003',	4,	N'PH_SY001'	,NULL,	'4',	N'7',	N'부자재-수입',			N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM003',	5,	N'PH_SY001'	,NULL,	'5',	N'8',	N'부자재-공중',			N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM003',	6,	N'PH_SY001'	,NULL,	'6',	N'9',	N'부자재-국내',			N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM003',	7,	N'PH_SY001'	,NULL,	'7',	N'R',	N'보강제',				N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'MM004',	N'자재요청진행상태', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'자재요청진행상태')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	1,	N'PH_SY001'	,NULL,	'1',	N'M00',	N'자재의뢰안함',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	2,	N'PH_SY001'	,NULL,	'2',	N'M01',	N'자재의뢰요청중',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	3,	N'PH_SY001'	,NULL,	'3',	N'M02',	N'구매의뢰중',					N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	4,	N'PH_SY001'	,NULL,	'4',	N'M03',	N'구매발주완료',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	5,	N'PH_SY001'	,NULL,	'5',	N'T00',	N'타계정의뢰안함',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	6,	N'PH_SY001'	,NULL,	'6',	N'T01',	N'타계정구매의뢰중',			N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	7,	N'PH_SY001'	,NULL,	'7',	N'T02',	N'타계정구매발주완료',			N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	8,	N'PH_SY001'	,NULL,	'7',	N'P00',	N'PMS의뢰안함',					N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	9,	N'PH_SY001'	,NULL,	'7',	N'P01',	N'PMS구매의뢰중',				N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'MM004',	10,	N'PH_SY001'	,NULL,	'7',	N'P02',	N'PMS구매발주완료',				N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------






----수입관련 공통코드 TC001
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC001',	N'수입 원천문서 유형', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'자재구분')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	1,	N'PH_SY001'	,NULL,	'1',	'22', N'구매오더', N'PO'		,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	2,	N'PH_SY001'	,NULL,	'2',	'18', N'A/P 예약송장', N'PO'	,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	3,	N'PH_SY001'	,NULL,	'3',	'POBL', N'수입 BL', N'PO'		,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	4,	N'PH_SY001'	,NULL,	'4',	'POBLT', N'수입 통관', N'PO'	,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	5,	N'PH_SY001'	,NULL,	'5',	'POLC', N'수입 LC', N'PO'		,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	6,	N'PH_SY001'	,NULL,	'6',	'POLCA', N'수입 LC Amend', N'PO',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	7,	N'PH_SY001'	,NULL,	'7',	'POTL', N'수입 결재조건', N'PO'	,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	8,	N'PH_SY001'	,NULL,	'8',	'17', N'판매오더', N'SO'		,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	9,	N'PH_SY001'	,NULL,	'9',	'SOTL', N'수출 결재조건',N'SO'	,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	10,	N'PH_SY001'	,NULL,	'10',	'SOPK', N'수출 Packing', N'SO'	,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	11,	N'PH_SY001'	,NULL,	'11',	'SOIV', N'수출 Invoice', N'SO'	,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	12,	N'PH_SY001'	,NULL,	'12',	'SOBL', N'수출 BL', N'SO'		,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	13,	N'PH_SY001'	,NULL,	'13',	'SOBLT', N'수출 통관', N'SO'	,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	14,	N'PH_SY001'	,NULL,	'14',	'SOLC', N'수출 LC', N'SO'		,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	15,	N'PH_SY001'	,NULL,	'15',	'SOLCA', N'수출 LC Amend', N'SO',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	16,	N'PH_SY001'	,NULL,	'16',	'POPO', N'수입 입고 PO', N'PO'	,N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC001',	17,	N'PH_SY001'	,NULL,	'17',	'15', N'수출 납품', N'SO'		,N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC002',	N'결재조건', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'결재조건')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC002',	1,	N'PH_SY001'	,NULL,	'1',	'TT', N'T/T'					,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC002',	2,	N'PH_SY001'	,NULL,	'2',	'LC', N'L/C'					,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC003',	N'인도조건', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'인도조건')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC003',	1,	N'PH_SY001'	,NULL,	'1',	'CIF', N'Cost,Insurance, Frig'	,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC003',	2,	N'PH_SY001'	,NULL,	'2',	'DDU', N'Delivered Duty Unpai'	,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC003',	3,	N'PH_SY001'	,NULL,	'3',	'FOB', N'Free on Board'			,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC004',	N'결재추가정보', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'결재추가정보')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC004',	1,	N'PH_SY001'	,NULL,	'1',	'INA', N'IN ADVANCE'			,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC004',	2,	N'PH_SY001'	,NULL,	'2',	'FBL', N'FROM B/L DATE'			,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC004',	3,	N'PH_SY001'	,NULL,	'3',	'ATS', N'At Sight'				,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC004',	4,	N'PH_SY001'	,NULL,	'4',	'USC', N'Usance'				,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC005',	N'포장방법', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'포장방법')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC005',	1,	N'PH_SY001'	,NULL,	'1',	'1', N'Wood'					,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC005',	2,	N'PH_SY001'	,NULL,	'2',	'2', N'PLASTIC'					,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC005',	3,	N'PH_SY001'	,NULL,	'3',	'3', N'PAPER'					,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC006',	N'운송방법', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'운송방법')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC006',	1,	N'PH_SY001'	,NULL,	'1',	'SEA', N'SEA'					,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC006',	2,	N'PH_SY001'	,NULL,	'2',	'AIR', N'AIR'					,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC007',	N'검사방법', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'검사방법')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC007',	1,	N'PH_SY001'	,NULL,	'1',	'1', N'Quality fat #1'			,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC008',	N'B/L구분', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'B/L구분')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC008',	1,	N'PH_SY001'	,NULL,	'1',	'BL', N'선화증권'				,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC008',	2,	N'PH_SY001'	,NULL,	'2',	'AWB', N'항공화물운송장'		,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC009',	N'CAGO 유형', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'CAGO 유형')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC009',	1,	N'PH_SY001'	,NULL,	'1',	'1', N'Full Container'			,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC010',	N'통관계획', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'통관계획')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC010',	1,	N'PH_SY001'	,NULL,	'1',	'F', N'도착후부두직반출'		,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC011',	N'신고구분', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'신고구분')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC011',	1,	N'PH_SY001'	,NULL,	'1',	'B', N'일반서류신고'			,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC012',	N'거래구분', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'거래구분')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC012',	1,	N'PH_SY001'	,NULL,	'1',	'1', N'일반형태수입'			,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC013',	N'징수형태', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'징수형태')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC013',	1,	N'PH_SY001'	,NULL,	'1',	'11', N'신고납부 수리전납부'	,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC014',	N'수입종류', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'수입종류')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC014',	1,	N'PH_SY001'	,NULL,	'1',	'K', N'일반수입'				,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC015',	N'전표상태', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'전표상태')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC015',	1,	N'PH_SY001'	,NULL,	'1',	'A', N'활성'					,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC015',	2,	N'PH_SY001'	,NULL,	'2',	'C', N'취소'					,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC015',	3,	N'PH_SY001'	,NULL,	'3',	'L', N'종료'					,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC016',	N'구매유형', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'구매유형')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC016',	1,	N'PH_SY001'	,NULL,	'1',	'1', N'내수'					,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC016',	2,	N'PH_SY001'	,NULL,	'2',	'2', N'수입'					,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC017',	N'PALLET 규격', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'PALLET 규격')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC017',	1,	N'PH_SY001'	,NULL,	'1',	'1', N'1.10*1.10*1.40[m]'		,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC017',	2,	N'PH_SY001'	,NULL,	'2',	'2', N'1.10*1.10*1.50[m]'		,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC017',	3,	N'PH_SY001'	,NULL,	'3',	'3', N'1.10*1.13*1.40[m]'		,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC017',	4,	N'PH_SY001'	,NULL,	'4',	'4', N'1.10*1.10*0.50[m]'		,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC017',	5,	N'PH_SY001'	,NULL,	'5',	'5', N'1.10*1.30*0.75[m]'		,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC017',	6,	N'PH_SY001'	,NULL,	'6',	'6', N'1.10*1.30*1.05[m]'		,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC017',	7,	N'PH_SY001'	,NULL,	'7',	'7', N'1.13*1.14*1.50[m]'		,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
--/* 헤더 */
--INSERT INTO [@PH_SY001H] VALUES(N'TC018',	N'운임지불방식', @MDC_DOCENTRY,	'N',	N'PH_SY001',	0,	1,	'N',	@MDC_DATE,	@MDC_TIME,	0,	0,	'I',	N'운임지불방식')
--/* 라인 */
--INSERT INTO [@PH_SY001L] VALUES(N'TC018',	1,	N'PH_SY001'	,NULL,	'1',	'01', N'Collect'				,N'',N'',N'',N'',N'')
--INSERT INTO [@PH_SY001L] VALUES(N'TC018',	2,	N'PH_SY001'	,NULL,	'2',	'02', N'Prepaid'				,N'',N'',N'',N'',N'')
------------------------------------------------------------------------------------------------------------------------------------------------------------------





--AUTOKEY갱신
SELECT @MDC_DOCENTRY = ISNULL(MAX(DOCENTRY), 0) + 1 FROM [@PH_SY001H]
UPDATE ONNM SET AUTOKEY = @MDC_DOCENTRY
WHERE OBJECTCODE = 'CSY001'


--SELECT * FROM [@PH_SY001L] WHERE Code IN('CSY001','CSY002','CSY003')
