VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY521"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY521.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 근로소득전산매체수록
'//  FormType       : 2000060521
'//  Create Date    : 2006.01.24
'//  Modified Date  :
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private sRecordset      As SAPbobsCOM.Recordset
Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     '클래스에서 선택한 마지막 아이템 Uid값

Private oJsnYear        As String
Private JSNGBN          As String
Private STRMON          As String
Private ENDMON          As String
Private MSTBRK          As String
Private DPTSTR          As String
Private DPTEND          As String
Private MSTCOD          As String
Private oFilePath       As String

Private FILNAM      As String * 30    '파  일  명
Private MaxRow      As Long
Private BUSCNT      As Integer  '/ B레코드일련번호
Private BUSTOT      As Integer  '/ B레코드총갯수

Private NEWCNT      As Integer
Private OLDCNT      As Integer
Private C_MSTCOD    As String
Private C_CLTCOD    As String
Private E_BUYCNT    As String
Private C_BUYCNT    As String
Private CLTCOD      As String

Private Type Arecord
        RECGBN As String * 1     '레코드구분
        DTAGBN As String * 2     '자료  구분
        TAXCOD As String * 3     '세  무  서
        PRTDAT As String * 8     '제출  일자
        
        RPTGBN As String * 1     '제  출  자 (1;세무대리인, 2;법인, 3;개인)
        TAXAGE As String * 6     '세무대리인
        HOMTID As String * 20    '홈텍스ID
        PGMCOD As String * 4     '세무프로그램코드
        BUSNBR As String * 10    '사업자번호
        SANGHO As String * 40    '상      호
        DAMDPT As String * 30    '담당부서명
        DAMNAM As String * 30    '담당 성명
        DAMTEL As String * 15    '담당전화번호
        
        BUSCNT As String * 5     'B Record수(신고의무자수)
        HANCOD As String * 3     '한글코드종류
        FILLER As String * 1082  '공      란
End Type
Dim Arec As Arecord

Private Type Brecord
        RECGBN As String * 1     '레코드구분
        DTAGBN As String * 2     '자료  구분
        TAXCOD As String * 3     '세  무  서
        BUSCNT As String * 6     '일련  번호
        
        BUSNBR As String * 10    '사업자번호
        SANGHO As String * 40    '상      호
        COMPRT As String * 30    '대  표  자
        PERNBR As String * 13    '주민  번호
        
        NEWCNT As String * 7     'C Record수
        OLDCNT As String * 7     'D Record수
        INCOME As String * 14    '소득금액총액
        GULGAB As String * 13    '결정소득세
        GULCOM As String * 13    '    법인세(공란)
        GULJUM As String * 13    '    주민세
        GULNON As String * 13    '    농특세
        GULTOT As String * 13    '    총  액
        RNGCOD As String * 1     '제출대상기간
        FILLER As String * 1061  '공      란
End Type
Dim Brec As Brecord

'/ 근로 주(현) 근무처 레코드 /
Private Type Crecord
        RECGBN As String * 1     '레코드구분
        DTAGBN As String * 2     '자료  구분
        TAXCOD As String * 3     '세  무  서
        SQNNBR As String * 6     '일련  번호
        
        BUSNBR As String * 10    '사업자번호
        
        JONCNT As String * 2     '종전근무처수
        DWEGBN As String * 1     '거주자구분
        RGNCOD As String * 2     '거주지국코드
        RGNTAX As String * 1     '외국인단일세율적용
        MSTNAM As String * 30    '성      명
        INTGBN As String * 1     '내외국인구분
        PERNBR As String * 13    '주민  번호
        GUKCOD As String * 2     '국적코드                                  (2010년 추가)
        HUSMAN As String * 1     '세대주여부                                (2010년 추가)
        JSNGBN As String * 1     '연말정산구분                              (2010년 추가)
        
        BUSNB1 As String * 10    '주(현) 사업자번호                         (2010년 추가)
        SANGHO As String * 40    '주(현) 근무처명                           (2010년 추가)
        STRINT As String * 8     '귀속년도시작=>근무기간시작연월일
        ENDINT As String * 8     '귀속년도종료=>근무기간종료연월일
        STRGAM As String * 8     '감면기간시작
        ENDGAM As String * 8     '감면기간종료101
        
        PAYAMT As String * 11    '급여  총액
        BNSAMT As String * 11    '상여  총액
        INJBNS As String * 11    '인정  상여
        JUSBNS As String * 11    '주식매수선택권행사이익                    (2007년추가)
        URIBNS As String * 11    '우리사주조합인출금                        (2009년추가)
        FILD01 As String * 22    '공란(2009년추가)
        TOTAMT As String * 11    '급상여총액
        
        BIGG01 As String * 10    '비 과 세(G01:학자금)
        BIGH01 As String * 10    '비 과 세(H01:무보수위원수당)
        BIGH05 As String * 10    '비 과 세(H05:경호,승선수당)
        BIGH06 As String * 10    '비 과 세(H06:유아,초중등)
        BIGH07 As String * 10    '비 과 세(H07:고등교육법)
        BIGH08 As String * 10    '비 과 세(H08:특별법)
        BIGH09 As String * 10    '비 과 세(H09:연구기관등)
        BIGH10 As String * 10    '비 과 세(H10:기업연구소)
        BIGH11 As String * 10    '비 과 세(H11:취재수당)
        BIGH12 As String * 10    '비 과 세(H12:벽지수당)
        BIGH13 As String * 10    '비 과 세(H13:재해관련급여)
        BIGI01 As String * 10    '비 과 세(I01:외국정부등근로자)
        BIGK01 As String * 10    '비 과 세(K01:외국주둔군인등)
        BIGM01 As String * 10    '비 과 세(M01:국외근로100만원)
        BIGM02 As String * 10    '비 과 세(M02:국외근로150만원)
        BIGM03 As String * 10    '비 과 세(M03:국외근로)
        BIGO01 As String * 10    '비 과 세(O01:야간근로수당)
        BIGQ01 As String * 10    '비 과 세(Q01:출산보육수당)
        BIGR10 As String * 10    '비 과 세(R10:근로장학금)                  (2011년 신설)
        BIGS01 As String * 10    '비 과 세(S01:주식매수선택권)
        BIGT01 As String * 10    '비 과 세(T01:외국인기술자)
'        BIGX01 As String * 10    '비 과 세(X01:외국인근로자)                (2010년 폐지)
'        BIGY01 As String * 10    '비 과 세(Y01:우리사주조합배정)            (2010년 폐지)
        BIGY02 As String * 10    '비 과 세(Y02:우리사주조합인출금50%)
        BIGY03 As String * 10    '비 과 세(Y03:우리사주조합인출금75%)
        'BIGY20 As String * 10    '비 과 세(Y20:주택자금보조금)              (2010년 폐지)
        BIGY21 As String * 10    '비 과 세(Y21:장기미취업자 중소기업 취업)  (2010년 추가)
        BIGZ01 As String * 10    '비 과 세(Z01:해저광물자원개발)
        BIGY22 As String * 10    '비 과 세(그밖의비과세)=명확화:(Y22:지정비과세)
        BIGTOT As String * 10    '비과세 합계
        BIGTO1 As String * 10    '감면소득 계
        
        INCOME As String * 11    '근로소득수입금액90
        PILGNL As String * 10    '근로소득공제
        GNLOSD As String * 11    '근로소득금액
        
        INJBAS As String * 8     '본인  공제
        INJBWO As String * 8     '배우자공제
        BUYNSU As String * 2     '부양가족인원
        INJBYN As String * 8     '부양가족공제
        
        GYNGLO As String * 2     '경로우대인원
        INJGYN As String * 8     '경로우대공제
        JANGAE As String * 2     '장애자인원
        INJJAE As String * 8     '장애자공제
        INJBNY As String * 8     '부녀자공제
        BUYN06 As String * 2     '자녀양육공제인원
        INJSON As String * 8     '자녀양육공제금액
        CHLSAN As String * 2     '출산입양자공제인원
        INJCHL As String * 8     '출산입양자공제금액
        FILD02 As String * 10    '공란
        
        DAGYSU As String * 2     '다자녀추가공제인원                        (2007년추가)
        INJADD As String * 8     '소수추가공제->2007년부터 다자녀추가공제로변경
        
        KUKGON As String * 10    '국민연금보험료공제
        YUNGON As String * 10    '기타연금보험료공제(공무원연금)            (2010년 항목분리)
        YUNGO1 As String * 10    '기타연금보험료공제(군인연금)              (2010년 항목분리)
        YUNGO2 As String * 10    '기타연금보험료공제(사립학교교직원연금)    (2010년 항목분리)
        YUNGO3 As String * 10    '기타연금보험료공제(별정우체국연금)        (2010년 항목분리)
        GITRE2 As String * 10    '퇴직연금소득공제(과학기술인공제)          (2010년 항목분리)
        GITRET As String * 10    '퇴직연금소득공제(근로자퇴직급여보장법)    (2010년 항목분리)
        
        PILMBH As String * 10    '보험료_건강보험료                         (2010년 항목분리)
        PILGBH As String * 10    '보험료_고용보험료                         (2010년 항목분리)
        PILBHM As String * 10    '보험료_보장성보험                         (2010년 항목분리)
        PILJHM As String * 10    '보험료_장애인전용                         (2010년 항목분리)
        PILMED As String * 10    '의료비공제
        PILSCH As String * 8     '교육비공제
        PILHUS As String * 8     '주택자금_주택임대차차입금 원리금상환공제금액-대출기관(구:주택자금공제)
        PILHU2 As String * 8     '주택자금_주택임대차차입금 원리금상환공제금액-거주자(구:주택자금공제) (2011년)
        PILWOL As String * 8     '주택자금_월세액                           (2010년 추가)
        PILJHE As String * 8     '장기주택저당차입금이자상환공제금액-15년미만(2011년 년수별세분화됨)
        PILJH2 As String * 8     '장기주택저당차입금이자상환공제금액-29년미만(2011년)
        PILJH3 As String * 8     '장기주택저당차입금이자상환공제금액-30년이상(2011년)
        PILGBU As String * 10    '기부금공제
        PILFLD As String * 20    '공란
        PILTOT As String * 10    '계
        PILGON As String * 8     '표준공제
        CHAGAM As String * 11    '차감소득금액
        
        GITGYN As String * 8     '개인연금저축소득공제
        GITYUN As String * 8     '연금소득공제
        GITSGI As String * 10    '소기업공제부금소득공제
        GITHUS As String * 10    '주택마련저축소득공제_청약저축             (2010년 항목분리)
        GITHU1 As String * 10    '주택마련저축소득공제_주택청약종합저축     (2010년 항목분리)
        GITHU2 As String * 10    '주택마련저축소득공제_장기주택마련저축     (2010년 항목분리)
        GITHU3 As String * 10    '주택마련저축소득공제_근로자주택마련저축   (2010년 항목분리)
        GITINV As String * 10    '투자조합공제
        GITCAD As String * 8     '신용카드공제
'        GITUSG As String * 1     '우리사주조합소득공제(음수1 양수0)         (2010년 삭제)
        GITUSJ As String * 10    '우리사주조합출연금
        GITJFD As String * 10    '장기주식형저축소득공제
        GITGYU As String * 10    '고용유지중소기업근로자소득공제(2009년추가)
        FILD03 As String * 10    '공란(2009년추가)
'        GITTOG As String * 1     '기타소득공제계(기호 음수1, 양수0)         (2010년 삭제)
        GITTOT As String * 10    '기타소득공제계
        TAXSTD As String * 11    '종합소득과세표준
        SANTAX As String * 10    '산출  세액
        
        GAMSOD As String * 10    '소득세  법
        GAMJOS As String * 10    '조  감  법
        GAMJYK As String * 10    '조 세 조약(2011년추가)
        GAMFLD As String * 10    '공      란
        GAMTOT As String * 10    '세액감면계
        
        TAXGNL As String * 10    '근로  소득
        NABSEE As String * 10    '납세  조합
        TAXBRO As String * 10    '주택  차입
        TAXGBU As String * 10    '기부정치자금
        TAXFRG As String * 10    '외국  납부
        FILD04 As String * 10    '공란
        TAXTOT As String * 10    '세액공제계
        
        GULGAB As String * 10    '결정소득세
        GULJUM As String * 10    '    주민세
        GULNON As String * 10    '    특별세
       ' GULTOT As String * 10    '    세액계(2011년삭제)
        
        NANGAB As String * 10    '현기납부소득세
        NANJUM As String * 10    '        주민세
        NANNON As String * 10    '        특별세
     '   NANTOT As String * 10    '        세액계(2011년 삭제)
        
        CHAGAG As String * 1     '현)차감소득세(기호 음수1, 양수0)
        CHAGAB As String * 10    '현)    소득세
        CHAJUG As String * 1     '현)    주민세(기호 음수1, 양수0)
        CHAJUM As String * 10    '       주민세
        CHANOG As String * 1     '현)    특별세(기호 음수1, 양수0)
        CHANON As String * 10    '       특별세
'        CHATOG As String * 1     '현)    세액계(기호 음수1, 양수0)(2011년삭제)
'        CHATOT As String * 10    '       세액계(2011년삭제)
        FILLER As String * 5     '공      란
End Type
Dim Crec As Crecord

Private Type D_Record
        RECGBN As String * 1    '레코드구분
        DTAGBN As String * 2    '자료  구분
        TAXCOD As String * 3    '세  무  서
        SQNNBR As String * 6    '일련  번호
        
        BUSNBR As String * 10   '사업자번호
        FILD01 As String * 50   '공      란
        
        PERNBR As String * 13   '주민  번호
        
        TAXJOH As String * 1    '납세조합구분                               (2010년 추가)
        JONNAM As String * 40   '근무처  명
        JONNBR As String * 10   '사업자번호
        STRINT As String * 8    '근무기간시작연월일                         (2009년 추가)
        ENDINT As String * 8    '근무기간종료연월일                         (2009년 추가)
        STRGAM As String * 8    '감면기간시작                               (2009년 추가)
        ENDGAM As String * 8    '감면기간종료                               (2009년 추가)
                
        PAYAMT As String * 11   '급여  총액
        BNSAMT As String * 11   '상여  총액
        INJBNS As String * 11   '인정  상여
        JUSBNS As String * 11   '주식매수선택권행사이익                     (2007년 추가)
        URIBNS As String * 11   '우리사주조합인출금                         (2009년추가)
        FILD02 As String * 22   '공란                                       (2009년추가)
        TOTAMT As String * 11   '급상여총액
        
        BIGG01 As String * 10    '비 과 세(G01:학자금)
        BIGH01 As String * 10    '비 과 세(H01:무보수위원수당)
        BIGH05 As String * 10    '비 과 세(H05:경호,승선수당)
        BIGH06 As String * 10    '비 과 세(H06:유아,초중등)
        BIGH07 As String * 10    '비 과 세(H07:고등교육법)
        BIGH08 As String * 10    '비 과 세(H08:특별법)
        BIGH09 As String * 10    '비 과 세(H09:연구기관등)
        BIGH10 As String * 10    '비 과 세(H10:기업연구소)
        BIGH11 As String * 10    '비 과 세(H11:취재수당)
        BIGH12 As String * 10    '비 과 세(H12:벽지수당)
        BIGH13 As String * 10    '비 과 세(H13:재해관련급여)
        BIGI01 As String * 10    '비 과 세(I01:외국정부등근로자)
        BIGK01 As String * 10    '비 과 세(K01:외국주둔군인등)
        BIGM01 As String * 10    '비 과 세(M01:국외근로100만원)
        BIGM02 As String * 10    '비 과 세(M02:국외근로150만원)
        BIGM03 As String * 10    '비 과 세(M02:국외근로)
        BIGO01 As String * 10    '비 과 세(O01:야간근로수당)
        BIGQ01 As String * 10    '비 과 세(Q01:출산보육수당)
        BIGR10 As String * 10    '비 과 세(R10:근로장학금)                   (2011년추가)
        BIGS01 As String * 10    '비 과 세(S01:주식매수선택권)
        BIGT01 As String * 10    '비 과 세(T01:외국인기술자)
'        BIGX01 As String * 10    '비 과 세(X01:외국인근로자)                (2010년 폐지)
'        BIGY01 As String * 10    '비 과 세(Y01:우리사주조합배정)            (2010년 폐지)
        BIGY02 As String * 10    '비 과 세(Y02:우리사주조합인출금50%)
        BIGY03 As String * 10    '비 과 세(Y03:우리사주조합인출금75%)
        'BIGY20 As String * 10    '비 과 세(Y20:주택자금보조금)              (2010년 폐지)
        BIGY21 As String * 10    '비 과 세(Y21:장기미취업자 중소기업취업)   (2010년 추가)
        BIGZ01 As String * 10    '비 과 세(Z01:해저광물자원개발)
        BIGY22 As String * 10    '비 과 세(그밖의비과세->지정비과세)
        BIGTOT As String * 10    '비과세합계
        BIGTO1 As String * 10    '감면소득 계
        
        NANGAB As String * 10    '현기납부소득세
        NANJUM As String * 10    '        주민세
        NANNON As String * 10    '        특별세
        'NANTOT As String * 10    '       세액계 (2011년 폐지)
        
        JONCNT As String * 2    '종전근무처일련번호
        FILLER As String * 692  '공      란
End Type
Dim Drec As D_Record

'/ 부양가족공제자명세 레코드 /
Private Type E_Record
        RECGBN As String * 1    '레코드구분
        DTAGBN As String * 2    '자료  구분
        TAXCOD As String * 3    '세  무  서
        SQNNBR As String * 6    '일련  번호
        
        BUSNBR As String * 10   '사업자번호
        
        PERNBR As String * 13   '주민  번호
        
        CHKCOD(1 To 5) As String * 1    '관계코드
        CHKINT(1 To 5) As String * 1    '내외국인구분
        CHKNAM(1 To 5) As String * 20   '성명
        CHKPER(1 To 5) As String * 13   '주민등록번호
        CHKBAS(1 To 5) As String * 1    '기본공제
        CHKJAN(1 To 5) As String * 1    '장애자공제
        CHKBY6(1 To 5) As String * 1    '자녀양육
        CHKBUY(1 To 5) As String * 1    '부녀자             (2007년추가)
        CHKJEL(1 To 5) As String * 1    '경로우대           (2007년추가)
        'CHKDAG(1 To 5) As String * 1    '다자녀(2007년추가) ->2009년제외
        CHKCHS(1 To 5) As String * 1    '출산입양           (2008년추가)
        BOHAM1(1 To 5) As String * 10   '보험료
        MEDAM1(1 To 5) As String * 10   '의료비
        EDCAM1(1 To 5) As String * 10   '교육비
        CADAM1(1 To 5) As String * 10   '신용카드
        CSHCA1(1 To 5) As String * 10   '직불카드           (2010년 추가)
        CSHAM1(1 To 5) As String * 10   '현금영수증
        GBUAM1(1 To 5) As String * 10   '기부금(2009년추가)
        BOHAM2(1 To 5) As String * 10   '보험료
        MEDAM2(1 To 5) As String * 10   '의료비
        EDCAM2(1 To 5) As String * 10   '교육비
        CADAM2(1 To 5) As String * 10   '신용카드
        CSHCA2(1 To 5) As String * 10   '직불카드           (2010년 추가)
        GBUAM2(1 To 5) As String * 10   '기부금
        FAMNBR As String * 2        '부양가족일련번호
        FILLER As String * 368   '공      란
End Type
Dim Erec As E_Record

Private Type F_Record
        RECGBN As String * 1    '레코드구분
        DTAGBN As String * 2    '자료  구분
        TAXCOD As String * 3    '세  무  서
        SQNNBR As String * 6    '일련  번호
        
        BUSNBR As String * 10   '사업자번호
        
        PERNBR As String * 13   '주민  번호
        
        SAVGBN(1 To 15) As String * 2   '소득공제구분
        SAVCOD(1 To 15) As String * 3   '금융기관코드
        SAVNAM(1 To 15) As String * 30  '금융기관상호
        SAVNUM(1 To 15) As String * 20  '계좌번호
        STYEAR(1 To 15) As String * 2   '납입연차
        SAVAMT(1 To 15) As String * 10  '불입금액
        SARAMT(1 To 15) As String * 10  '공제금액
        
        FILLER As String * 70
End Type
Dim Frec As F_Record

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY521.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY521_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "ZPY521"
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    
    oForm.Freeze True
    CreateItems
    oForm.Freeze False
    
    oForm.EnableMenu ("1281"), False '/ 찾기
    oForm.EnableMenu ("1282"), True  '/ 추가
    oForm.EnableMenu ("1284"), False '/ 취소
    oForm.EnableMenu ("1293"), False '/ 행삭제
    oForm.Update
    oForm.Visible = True

    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
    Sbo_Application.StatusBar.SetText "Form_Load Error:" & Err.Description, bmt_Short, smt_Error
    If (oForm Is Nothing) = False Then
        oForm.Freeze False
        Set oForm = Nothing
    End If
End Sub
'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
'et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_ITEM_PRESSED
            If pval.BeforeAction Then
                If pval.ItemUid = "CBtn1" Then   '/ ChooseBtn사원리스트
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUid = "1" And oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If File_Create = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        BubbleEvent = False
                        oForm.Mode = fm_OK_MODE
                    End If
                ElseIf pval.ItemUid = "Btn1" Then
                   oFilePath = ZP_Form.vbGetBrowseDirectory(ZP_Form)
                   oForm.DataSources.UserDataSources("Path").ValueEx = oFilePath
                   BubbleEvent = False
                   Exit Sub
               End If
            Else
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True And _
              (pval.ItemUid = "JsnYear" Or _
              pval.ItemUid = "MSTCOD") Then
               FlushToItemValue pval.ItemUid
            End If
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
             If pval.BeforeAction = True And pval.ItemUid <> "1000001" And pval.ItemUid <> "2" Then
                If Last_Item = "JsnYear" Then     '/정산년도
                    If Trim$(oForm.Items(Last_Item).Specific.Value) <> "" Then
                        If MDC_SetMod.ChkYearMonth(Trim$(CStr((oForm.Items(Last_Item).Specific.Value))) & "01") = False Then
                            oForm.Items(Last_Item).Update
                            Sbo_Application.StatusBar.SetText "정산 연도를 확인하여 주십시오.", bmt_Short, smt_Error
                            BubbleEvent = False
                        End If
                    End If
                ElseIf Last_Item = "MSTCOD" Then
                    If Trim$(oForm.Items(Last_Item).Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items(Last_Item).Specific.String) & "'", "") = True Then
                        oForm.Items(Last_Item).Update
                        Sbo_Application.StatusBar.SetText "사원 번호를 확인하여 주십시오.", bmt_Short, smt_Error
                        BubbleEvent = False
                    End If
                End If
            End If
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
             If pval.BeforeAction = True And pval.ItemUid = "JsnYear" And pval.CharPressed = 9 Then
                If Len(Trim$(oForm.Items(pval.ItemUid).Specific.String)) < 4 Then
                    oForm.Items(pval.ItemUid).Specific.Value = Format$(oForm.Items(pval.ItemUid).Specific.Value, "2000")
                End If
                If MDC_SetMod.ChkYearMonth(Trim$(CStr((oForm.Items(pval.ItemUid).Specific.Value))) & "01") = False Then
                    Sbo_Application.StatusBar.SetText "정산 연도를 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And pval.ItemUid = "MSTCOD" And pval.CharPressed = 9 Then
                If Trim$(oForm.Items("MSTCOD").Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'", "") = True Then
                    Sbo_Application.StatusBar.SetText "사원 번호를 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
            End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUid
                End If
            Else
                Last_Item = pval.ItemUid
            End If
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '컬렉션에서 삭제및 모든 메모리 제거
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            If pval.BeforeAction = False Then
               RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oMat1 = Nothing
            End If
    End Select
    
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    
    If pval.BeforeAction = True Then
        Exit Sub
    End If
    
    Select Case pval.MenuUID
        Case "1287" '/ 복제
        Case "1281", "1282":
            oForm.Items("JsnYear").CLICK ct_Regular
        Case "1288" To "1291":
        Case "1293"
    End Select
    Exit Sub
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    Dim i       As Long
    Dim sQry    As String
    Dim oCombo  As SAPbouiCOM.ComboBox
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormDataEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select

    End If
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
    
Raise_FormDataEvent_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo1         As SAPbouiCOM.ComboBox
    Dim oCombo2         As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oEdit           As SAPbouiCOM.EditText
    Dim oColumn         As SAPbouiCOM.Column
    Dim sQry            As String

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
        
    Call oForm.DataSources.UserDataSources.Add("JsnYear", dt_SHORT_TEXT, 4)     '/ 생성년도
    Call oForm.DataSources.UserDataSources.Add("JSNGBN", dt_SHORT_TEXT, 10)     '/ 생성구분
    Call oForm.DataSources.UserDataSources.Add("JSNTYP", dt_SHORT_TEXT, 10)     '/ 생성구분
    Call oForm.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 10)      '/ 지점
    Call oForm.DataSources.UserDataSources.Add("DptStr", dt_SHORT_TEXT, 10)     '/ 부서코드
    Call oForm.DataSources.UserDataSources.Add("DptEnd", dt_SHORT_TEXT, 10)
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 8)
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 30)
    Call oForm.DataSources.UserDataSources.Add("EmpID", dt_SHORT_TEXT, 10)
    Call oForm.DataSources.UserDataSources.Add("PRTDAT", dt_DATE, 10)
    Call oForm.DataSources.UserDataSources.Add("SMonth", dt_SHORT_TEXT, 2)
    Call oForm.DataSources.UserDataSources.Add("EMonth", dt_SHORT_TEXT, 2)
    Call oForm.DataSources.UserDataSources.Add("Path", dt_SHORT_TEXT)
    
    Set oEdit = oForm.Items("JsnYear").Specific
    oEdit.DataBind.SetBound True, "", "JsnYear"
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    Set oEdit = oForm.Items("EmpID").Specific
    oEdit.DataBind.SetBound True, "", "EmpID"
    Set oEdit = oForm.Items("Path").Specific
    oEdit.DataBind.SetBound True, "", "Path"
    Set oEdit = oForm.Items("PRTDAT").Specific
    oEdit.DataBind.SetBound True, "", "PRTDAT"
    Set oEdit = oForm.Items("SMonth").Specific
    oEdit.DataBind.SetBound True, "", "SMonth"
    Set oEdit = oForm.Items("EMonth").Specific
    oEdit.DataBind.SetBound True, "", "EMonth"
    
    '// 생성구분
    Set oCombo1 = oForm.Items("JSNGBN").Specific

    oCombo1.ValidValues.Add "1", "연말정산(재직자)"
    oCombo1.ValidValues.Add "2", "중도정산(퇴사자)"
    oCombo1.ValidValues.Add "%", "전체"
    
    oCombo1.Select "%", psk_ByValue  '/ 전체
    
    '// 생성구분
    Set oCombo1 = oForm.Items("JSNTYP").Specific
    oCombo1.ValidValues.Add "1", "연간(01.01~12.31)지급분"
    oCombo1.ValidValues.Add "2", "폐업에 의한 수시 제출분"
    oCombo1.ValidValues.Add "3", "수시 분할제출분"
    Call oCombo1.Select(0, psk_Index)   '/ 전체
    
    '// 신고의무자
    Set oCombo1 = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    sQry = " SELECT T0.U_WCHCLT, MAX(T1.NAME)  FROM [@PH_PY005A] T0 INNER JOIN [@PH_PY005A] T1 ON T0.U_WCHCLT = T1.CODE GROUP BY T0.U_WCHCLT  ORDER BY T0.U_WCHCLT"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
    Loop
    If oCombo1.ValidValues.Count > 0 Then
        Call oCombo1.Select(0, psk_Index)  '/ 전체
    End If
    
    '// 사업장
    Set oCombo1 = oForm.Items("BPLId").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    oCombo1.ValidValues.Add "%", "모두"
    Do Until oRecordSet.EOF
        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
    Loop
    If oCombo1.ValidValues.Count > 0 Then
        Call oCombo1.Select(0, psk_Index)  '/ 전체
    End If
    '// 부서
    Set oCombo1 = oForm.Items("DptStr").Specific
    Set oCombo2 = oForm.Items("DptEnd").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    oCombo1.ValidValues.Add "-1", "모두"
    oCombo2.ValidValues.Add "-1", "모두"
    Do Until oRecordSet.EOF
        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oCombo2.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
    Loop
    If oCombo1.ValidValues.Count > 0 Then
        Call oCombo1.Select(0, psk_Index)
        Call oCombo2.Select(0, psk_Index)
    End If
    Set oMat1 = oForm.Items("Mat1").Specific
    
    Call oForm.DataSources.UserDataSources.Add("Col0", dt_SHORT_TEXT)
    Call oForm.DataSources.UserDataSources.Add("Col1", dt_SHORT_TEXT)
    
    Set oColumn = oMat1.Columns("Col0")
    oColumn.DataBind.SetBound True, "", "Col0"
    
    Set oColumn = oMat1.Columns("Col1")
    oColumn.DataBind.SetBound True, "", "Col1"
    
    If Trim$(ZPAY_GBL_JSNYER) = "" Or Trim$(ZPAY_GBL_JSNYER) = "*" Then
        oForm.DataSources.UserDataSources("JsnYear").ValueEx = Format$(Now, "YYYY")
    Else
        oForm.DataSources.UserDataSources("JsnYear").ValueEx = ZPAY_GBL_JSNYER
    End If
    If Trim$(ZPAY_GBL_JSNYER) <> "" Then
        oForm.DataSources.UserDataSources("SMonth").ValueEx = "01"
        oForm.DataSources.UserDataSources("EMonth").ValueEx = "12"
    End If
    
    
    Set oEdit = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Function File_Create() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oStr            As String
    Dim sQry            As String
    
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    ErrNum = 0
    '/ Question
    If Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    oMat1.Clear
    MaxRow = 0
    
    BUSCNT = 0   '/ B레크드 일련번호
    BUSTOT = 0
    JSNGBN = oForm.Items("JSNGBN").Specific.Selected.Value
    STRMON = Format$(oForm.DataSources.UserDataSources("SMonth").ValueEx, "00")
    ENDMON = Format$(oForm.DataSources.UserDataSources("EMonth").ValueEx, "00")
    oJsnYear = oForm.Items("JsnYear").Specific.Value
    '/ 파일경로설정
    If oFilePath = "" Then oFilePath = "C:\EOSDATA"
    oFilePath = IIf(Right$(oFilePath, 1) = "\", oFilePath, oFilePath & "\")
    oStr = CreateFolder(Trim$(oFilePath))
    If Trim$(oStr) <> "" Then
        ErrNum = 5
        GoTo Error_Message
    End If
    
    '/ 근로 제출자(대리인) 레코드
    If File_Create_ARecord = False Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    Close #1
    Open FILNAM For Output As #1
        '/ A레코드: 근로 원천징수의무자별 집계 레코드
        Print #1, MDC_SetMod.sStr(Arec.RECGBN) & MDC_SetMod.sStr(Arec.DTAGBN) & MDC_SetMod.sStr(Arec.TAXCOD) & MDC_SetMod.sStr(Arec.PRTDAT) & _
                  MDC_SetMod.sStr(Arec.RPTGBN) & MDC_SetMod.sStr(Arec.TAXAGE) & MDC_SetMod.sStr(Arec.HOMTID) & MDC_SetMod.sStr(Arec.PGMCOD) & _
                  MDC_SetMod.sStr(Arec.BUSNBR) & MDC_SetMod.sStr(Arec.SANGHO) & MDC_SetMod.sStr(Arec.DAMDPT) & MDC_SetMod.sStr(Arec.DAMNAM) & _
                  MDC_SetMod.sStr(Arec.DAMTEL) & MDC_SetMod.sStr(Arec.BUSCNT) & MDC_SetMod.sStr(Arec.HANCOD) & _
                  MDC_SetMod.sStr(Arec.FILLER)
                  
                  
        Call Matrix_AddRow("제출자 레코드 생성 완료!", True)
        
        '/ B레코드: 근로 집계 레코드 /***********************************************/
        sQry = "SELECT Code, U_TAXCODE, U_BUSNUM, U_CLTNAME, U_COMPRT, U_PERNUM FROM [@PH_PY005A] "
        sQry = sQry & " WHERE  U_WCHCLT = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "' ORDER BY Code"
        sRecordset.DoQuery sQry
        Do Until sRecordset.EOF
            '/ B레코드: 근로 집계 레코드 /***********************************************/
            Brec.TAXCOD = sRecordset.Fields("U_TAXCODE").Value
            Brec.BUSNBR = Replace(sRecordset.Fields("U_BUSNUM").Value, "-", "")
            Brec.SANGHO = Trim$(sRecordset.Fields("U_CLTNAME").Value)
            Brec.COMPRT = Trim$(sRecordset.Fields("U_COMPRT").Value)
            Brec.PERNBR = Replace(Trim$(sRecordset.Fields("U_PERNUM").Value), "-", "")
            
            CLTCOD = sRecordset.Fields(0).Value
            '// 2010.03.02 최동권 수정
            '// 자사정보만 있고, 소득정보가 없어서 C레코드 없이 B레코드만 생성되는 경우가 있으므로
            '// 소득자료가 없으면 B레코드가 생성되지 않도록 함(Nefs)
            Select Case File_Create_BRecord
                Case 0
                    Call Matrix_AddRow(CLTCOD & "- 징수의무자의 집계 레코드 생성 완료!", True)
                    '/ C레코드: 근로 주(현)근무처 레코드 /***********************************************/
                    NEWCNT = 0
                    OLDCNT = 0
                    If File_Create_CRecord = False Then
                        ErrNum = 4
                        GoTo Error_Message
                    End If
                    Call Matrix_AddRow(CLTCOD & "- 징수의무자의 데이터 레코드" & NEWCNT & "건 생성 완료!", True)
                Case 1
                    ErrNum = 3
                    GoTo Error_Message
                Case 2  '// 해당자사에 소득자료가 없으면 B,C,D,E레코드 생성을 건너뜀
            
            End Select
            '/
            sRecordset.MoveNext
        Loop
    Close #1
    oForm.DataSources.UserDataSources("Path").Value = FILNAM
    Sbo_Application.StatusBar.SetText "전산매체수록이 정상적으로 완료되었습니다.", bmt_Short, smt_Success
    File_Create = True
    Set sRecordset = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set sRecordset = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "취소하였습니다.", bmt_Short, smt_Success
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "A레코드(근로 제출자 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "B레코드(근로 원천징수의무자별 집계 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "C레코드(근로 주(현)근무처 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "CreateFolder Error : " & oStr, bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "File_Create 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    File_Create = False
End Function
Private Function File_Create_ARecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim PRTDAT          As String
    Dim BUSNUM          As String
    Dim CheckA          As String
    
    CheckA = False  '/체크필요유무
    ErrNum = 0
    PRTDAT = Mid$(oForm.Items("PRTDAT").Specific.String, 1, 4) & Mid$(oForm.Items("PRTDAT").Specific.String, 6, 2) & Mid$(oForm.Items("PRTDAT").Specific.String, 9, 2)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    '/ 신고의무자수
    '// 자사정보만 있고, 소득정보가 없어서 C레코드 없이 B레코드만 생성되는 경우가 있으므로
    '// 소득자료가 없으면 B레코드가 생성되지 않도록 함
    sQry = "SELECT COUNT(CODE) from [@PH_PY005A] T0 " & _
           "WHERE U_WCHCLT = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "' " & _
           "AND Code IN (SELECT U_CLTCOD FROM [@ZPY504H] WHERE U_JSNYER = '" & oJsnYear & "')"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        BUSTOT = oRecordSet.Fields(0).Value
    End If
    If Val(BUSTOT) = 0 Then
        ErrNum = 3
        GoTo Error_Message
    End If
    '/ 업체정보
    sQry = "SELECT * FROM [@PH_PY005A] WHERE Code = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "'"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
     '/ 파일명
        BUSNUM = Replace(oRecordSet.Fields("U_BUSNUM").Value, "-", "")
        If Len(Trim$(BUSNUM)) <> 10 Then
            ErrNum = 2
            GoTo Error_Message
        End If
        FILNAM = oFilePath & "C" & Mid$(BUSNUM, 1, 7) & "." & Mid$(BUSNUM, 8, 3)
       '/ A Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        Arec.RECGBN = "A"                          '
        Arec.DTAGBN = "20"
        Arec.TAXCOD = oRecordSet.Fields("U_TAXCODE").Value
        Arec.PRTDAT = PRTDAT
        Arec.RPTGBN = oRecordSet.Fields("U_TaxDGbn").Value
        Arec.TAXAGE = oRecordSet.Fields("U_TaxDCode").Value
        Arec.HOMTID = Trim$(oRecordSet.Fields("U_HOMETID").Value)
        Arec.PGMCOD = "9000"
        Arec.BUSNBR = Replace(oRecordSet.Fields("U_TAXDBUS").Value, "-", "")
        Arec.SANGHO = Trim$(oRecordSet.Fields("U_TAXDNAM").Value)
        Arec.DAMDPT = Trim$(oRecordSet.Fields("U_CHGDPT").Value)
        Arec.DAMNAM = Trim$(oRecordSet.Fields("U_CHGNAME").Value)
        Arec.DAMTEL = Trim$(oRecordSet.Fields("U_CHGTEL").Value)
        Arec.BUSCNT = Format$(BUSTOT, String$(Len(Arec.BUSCNT), "0"))  '/ 원천징수의무자수
        Arec.HANCOD = "101"
        Arec.FILLER = Space$(Len(Arec.FILLER))
     
     '/ 필수입력 체크
        If Trim$(Arec.TAXCOD) = "" Then
            Call Matrix_AddRow("A레코드:세무서코드가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.RPTGBN) = "" Then
            Call Matrix_AddRow("A레코드:제출자구분가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.BUSNBR) = "" Then
            Call Matrix_AddRow("A레코드:제출자사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.SANGHO) = "" Then
            Call Matrix_AddRow("A레코드:제출자상호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.DAMDPT) = "" Then
            Call Matrix_AddRow("A레코드:담당자부서가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.DAMNAM) = "" Then
            Call Matrix_AddRow("A레코드:담당자성명이 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.DAMTEL) = "" Then
            Call Matrix_AddRow("A레코드:담당자전화번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Arec.BUSCNT = 0 Then
            Call Matrix_AddRow("A레코드:신고내역이 존재하는 B레코드가 없습니다. 확인하여 주십시오.", True, True)
            CheckA = True
        End If
    End If
    
    If CheckA = False Then
        File_Create_ARecord = True
    Else
        File_Create_ARecord = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
   If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "귀속년도의 자사정보가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "자사정보등록의 사업자번호가 올바르지 않습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "자사정보등록의 신고의무사업장이 존재하지 않습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("A레코드오류: " & Err.Description, False, True)
    End If
    File_Create_ARecord = False
End Function

'//------------------------------------------------------------
'// 반환값 리스트
'// 0 : 정상적으로 레코드 생성
'// 1 : 에러
'// 2 : B ~ E 레코드 생성 안함
'//------------------------------------------------------------
Private Function File_Create_BRecord() As Integer
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim CheckB          As String
    CheckB = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 집계정보
    sQry = "EXEC ZPY521 'B', " & "'" & oJsnYear & "', '" & JSNGBN & "', '" & STRMON & "', '" & ENDMON & "', " & _
                                 "'" & CLTCOD & "', '" & MSTBRK & "', '" & DPTSTR & "', '" & DPTEND & "','" & MSTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    ElseIf oRecordSet.Fields("SUM_NEWCNT").Value = 0 Then
        ErrNum = 2
        GoTo Error_Message
    Else
     '/ B Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        BUSCNT = BUSCNT + 1
        
        Brec.RECGBN = "B"
        Brec.DTAGBN = "20"
        Brec.BUSCNT = Format$(BUSCNT, String$(Len(Brec.BUSCNT), "0"))   '/ 원천징수의무자수 일련번호
        
        Brec.NEWCNT = Format$(oRecordSet.Fields("SUM_NEWCNT").Value, String$(Len(Brec.NEWCNT), "0"))       'C Record수
        Brec.OLDCNT = Format$(oRecordSet.Fields("SUM_OLDCNT").Value, String$(Len(Brec.OLDCNT), "0"))       'D Record수
        Brec.INCOME = Format$(oRecordSet.Fields("SUM_INCOME").Value, String$(Len(Brec.INCOME), "0"))
        Brec.GULGAB = Format$(oRecordSet.Fields("SUM_GULGAB").Value, String$(Len(Brec.GULGAB), "0"))
        Brec.GULCOM = Format$(0, String$(Len(Brec.GULCOM), "0"))
        Brec.GULJUM = Format$(oRecordSet.Fields("SUM_GULJUM").Value, String$(Len(Brec.GULJUM), "0"))
        Brec.GULNON = Format$(oRecordSet.Fields("SUM_GULNON").Value, String$(Len(Brec.GULNON), "0"))
        Brec.GULTOT = Format$(oRecordSet.Fields("SUM_GULGAB").Value + oRecordSet.Fields("SUM_GULJUM").Value + oRecordSet.Fields("SUM_GULNON").Value, String$(Len(Brec.GULTOT), "0"))
        Brec.RNGCOD = oForm.Items("JSNTYP").Specific.Selected.Value
        
        Brec.FILLER = Space$(Len(Brec.FILLER))
 
        Print #1, MDC_SetMod.sStr(Brec.RECGBN) & MDC_SetMod.sStr(Brec.DTAGBN) & MDC_SetMod.sStr(Brec.TAXCOD) & MDC_SetMod.sStr(Brec.BUSCNT) & _
                  MDC_SetMod.sStr(Brec.BUSNBR) & MDC_SetMod.sStr(Brec.SANGHO) & MDC_SetMod.sStr(Brec.COMPRT) & MDC_SetMod.sStr(Brec.PERNBR) & _
                  MDC_SetMod.sStr(Brec.NEWCNT) & MDC_SetMod.sStr(Brec.OLDCNT) & MDC_SetMod.sStr(Brec.INCOME) & MDC_SetMod.sStr(Brec.GULGAB) & _
                  MDC_SetMod.sStr(Brec.GULCOM) & MDC_SetMod.sStr(Brec.GULJUM) & MDC_SetMod.sStr(Brec.GULNON) & MDC_SetMod.sStr(Brec.GULTOT) & _
                  MDC_SetMod.sStr(Brec.RNGCOD) & MDC_SetMod.sStr(Brec.FILLER)
                  
     '/ 필수입력 체크
        If Trim$(Brec.BUSNBR) = "" Then
            Call Matrix_AddRow("B레코드:사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckB = True
        End If
        If Trim$(Brec.COMPRT) = "" Then
            Call Matrix_AddRow("B레코드:대표자명이 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckB = True
        End If
        If Trim$(Brec.PERNBR) = "" Then
            Call Matrix_AddRow("B레코드:법인(주민)등록번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckB = True
        End If
    End If
    
    If CheckB = False Then
        File_Create_BRecord = 0
    Else
        File_Create_BRecord = 1
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "집계레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
        File_Create_BRecord = 1
    ElseIf ErrNum = 2 Then
        File_Create_BRecord = 2
    Else
        Call Matrix_AddRow("B레코드오류: " & Err.Description, False)
        File_Create_BRecord = 1
    End If
    
End Function
Private Function File_Create_CRecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckC      As String
    Dim OLDBIG      As Double
    Dim PILTOT      As Double
    
    CheckC = False  '/체크필요유무
    ErrNum = 0
    C_MSTCOD = ""
    C_CLTCOD = ""
    C_BUYCNT = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

'/ 사원 정보 조회
    
    sQry = "EXEC ZPY521 'C', '" & oJsnYear & "', '" & JSNGBN & "', '" & STRMON & "', '" & ENDMON & "', " & _
                            "'" & CLTCOD & "', '" & MSTBRK & "', '" & DPTSTR & "', '" & DPTEND & "','" & MSTCOD & "'"
 
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo error_Message
    End If
    Do Until oRecordSet.EOF
        NEWCNT = NEWCNT + 1
        C_MSTCOD = oRecordSet.Fields("U_MSTCOD").Value
        C_CLTCOD = oRecordSet.Fields("U_CLTCOD").Value
'/ 근로소득 주(현) 근무처 레코드 /
        '/ C Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        Crec.RECGBN = "C"
        Crec.DTAGBN = "20"
        Crec.TAXCOD = Brec.TAXCOD
        Crec.SQNNBR = Format$(NEWCNT, String$(Len(Crec.SQNNBR), "0"))  '/ 일련번호
        Crec.BUSNBR = Brec.BUSNBR
        Crec.JONCNT = Format$(oRecordSet.Fields("SUM_OLDCNT").Value, String$(Len(Crec.JONCNT), "0"))   '/ 종전근무처수
        Crec.DWEGBN = Format$(oRecordSet.Fields("U_DWEGBN").Value, String$(Len(Crec.DWEGBN), "0"))
        Crec.PERNBR = Replace(oRecordSet.Fields("U_PERNBR").Value, "-", "")
        Crec.INTGBN = Format$(oRecordSet.Fields("U_INTGBN").Value, String$(Len(Crec.INTGBN), "0"))
        If Crec.DWEGBN = "1" Then
            Crec.RGNCOD = ""
        Else
            Crec.RGNCOD = oRecordSet.Fields("U_DWECOD").Value  '/ 거주지국
        End If
        If Crec.INTGBN = "1" Then
            Crec.GUKCOD = ""
        Else
            Crec.GUKCOD = oRecordSet.Fields("U_GUKCOD").Value  '/ 국적코드
        End If
        Crec.HUSMAN = oRecordSet.Fields("U_HUSMAN").Value   '/ 세대주여부
        Crec.JSNGBN = oRecordSet.Fields("U_JSNGBN").Value   '/ 정산구분
        Crec.RGNTAX = oRecordSet.Fields("U_FRGTAX").Value   '/ 외국인단일세율
        Crec.BUSNB1 = Replace(oRecordSet.Fields("U_BUSNUM").Value, "-", "") '/ 주(현)사업자번호
        Crec.SANGHO = oRecordSet.Fields("U_CLTNAM").Value   '/ 주(현) 근무처명

        Crec.STRINT = Format$(Mid$(Replace(oRecordSet.Fields("U_STRINT").Value, "-", ""), 1, 8), String$(Len(Crec.STRINT), "0"))
        Crec.ENDINT = Format$(Mid$(Replace(oRecordSet.Fields("U_ENDINT").Value, "-", ""), 1, 8), String$(Len(Crec.ENDINT), "0"))
        Crec.MSTNAM = oRecordSet.Fields("U_MSTNAM").Value
        Crec.STRGAM = Format$(0, String$(Len(Crec.STRGAM), "0"))
        Crec.ENDGAM = Format$(0, String$(Len(Crec.ENDGAM), "0"))
        
        Crec.PAYAMT = Format$(oRecordSet.Fields("U_PAYAMT").Value, String$(Len(Crec.PAYAMT), "0"))
        Crec.BNSAMT = Format$(oRecordSet.Fields("U_BNSAMT").Value, String$(Len(Crec.BNSAMT), "0"))
        Crec.INJBNS = Format$(oRecordSet.Fields("U_INBAMT").Value, String$(Len(Crec.INJBNS), "0"))
        Crec.JUSBNS = Format$(oRecordSet.Fields("U_JUSAMT").Value, String$(Len(Crec.JUSBNS), "0"))  '/2007
        Crec.URIBNS = Format$(oRecordSet.Fields("U_URIAMT").Value, String$(Len(Crec.URIBNS), "0"))  '/2009
        Crec.FILD01 = Format$(0, String$(Len(Crec.FILD01), "0"))
        Crec.TOTAMT = Format$(oRecordSet.Fields("U_TOTAMT").Value, String$(Len(Crec.TOTAMT), "0"))
        
        OLDBIG = Val(oRecordSet.Fields("U_BIGWA1").Value) + Val(oRecordSet.Fields("U_BIGWA3").Value) + Val(oRecordSet.Fields("U_BIGWA5").Value) _
                + Val(oRecordSet.Fields("U_BIGWA6").Value) + Val(oRecordSet.Fields("U_BIGWU3").Value)
                              
        Crec.BIGG01 = Format$(oRecordSet.Fields("U_BTXG01").Value, String$(Len(Crec.BIGG01), "0"))
        Crec.BIGH01 = Format$(oRecordSet.Fields("U_BTXH01").Value, String$(Len(Crec.BIGH01), "0"))
        Crec.BIGH05 = Format$(oRecordSet.Fields("U_BTXH05").Value, String$(Len(Crec.BIGH05), "0"))
        Crec.BIGH06 = Format$(oRecordSet.Fields("U_BTXH06").Value, String$(Len(Crec.BIGH06), "0"))
        Crec.BIGH07 = Format$(oRecordSet.Fields("U_BTXH07").Value, String$(Len(Crec.BIGH07), "0"))
        Crec.BIGH08 = Format$(oRecordSet.Fields("U_BTXH08").Value, String$(Len(Crec.BIGH08), "0"))
        Crec.BIGH09 = Format$(oRecordSet.Fields("U_BTXH09").Value, String$(Len(Crec.BIGH09), "0"))
        Crec.BIGH10 = Format$(oRecordSet.Fields("U_BTXH10").Value, String$(Len(Crec.BIGH10), "0"))
        Crec.BIGH11 = Format$(oRecordSet.Fields("U_BTXH11").Value, String$(Len(Crec.BIGH11), "0"))
        Crec.BIGH12 = Format$(oRecordSet.Fields("U_BTXH12").Value, String$(Len(Crec.BIGH12), "0"))
        Crec.BIGH13 = Format$(oRecordSet.Fields("U_BTXH13").Value, String$(Len(Crec.BIGH13), "0"))
        Crec.BIGI01 = Format$(oRecordSet.Fields("U_BTXI01").Value, String$(Len(Crec.BIGI01), "0"))
        Crec.BIGK01 = Format$(oRecordSet.Fields("U_BTXK01").Value, String$(Len(Crec.BIGK01), "0"))
        Crec.BIGM01 = Format$(oRecordSet.Fields("U_BTXM01").Value, String$(Len(Crec.BIGM01), "0"))
        Crec.BIGM02 = Format$(oRecordSet.Fields("U_BTXM02").Value, String$(Len(Crec.BIGM02), "0"))
        Crec.BIGM03 = Format$(oRecordSet.Fields("U_BTXM03").Value, String$(Len(Crec.BIGM03), "0"))
        Crec.BIGO01 = Format$(oRecordSet.Fields("U_BTXO01").Value, String$(Len(Crec.BIGO01), "0"))
        Crec.BIGQ01 = Format$(oRecordSet.Fields("U_BTXQ01").Value, String$(Len(Crec.BIGQ01), "0"))
        Crec.BIGR10 = Format$(oRecordSet.Fields("U_BTXY01").Value, String$(Len(Crec.BIGR10), "0"))  '/ 2011 Y01->R10으로 안쓰는필드대체
        Crec.BIGS01 = Format$(oRecordSet.Fields("U_BTXS01").Value, String$(Len(Crec.BIGS01), "0"))
        Crec.BIGT01 = Format$(oRecordSet.Fields("U_BTXT01").Value, String$(Len(Crec.BIGT01), "0"))
'        Crec.BIGX01 = Format$(oRecordSet.Fields("U_BTXX01").Value, String$(Len(Crec.BIGX01), "0"))
        Crec.BIGY02 = Format$(oRecordSet.Fields("U_BTXY02").Value, String$(Len(Crec.BIGY02), "0"))
        Crec.BIGY03 = Format$(oRecordSet.Fields("U_BTXY03").Value, String$(Len(Crec.BIGY03), "0"))
        Crec.BIGY21 = Format$(oRecordSet.Fields("U_BTXY21").Value, String$(Len(Crec.BIGY21), "0"))
        Crec.BIGZ01 = Format$(oRecordSet.Fields("U_BTXZ01").Value, String$(Len(Crec.BIGZ01), "0"))
        Crec.BIGY22 = Format$(oRecordSet.Fields("U_BTXY20").Value, String$(Len(Crec.BIGY22), "0"))  '/ 그밖의비과
        Crec.BIGTOT = Format$(oRecordSet.Fields("U_BTXTOT").Value, String$(Len(Crec.BIGTOT), "0"))
        Crec.BIGTO1 = Format$(oRecordSet.Fields("U_BTXTO1").Value, String$(Len(Crec.BIGTO1), "0"))
               
        Crec.INCOME = Format$(oRecordSet.Fields("U_INCOME").Value, String$(Len(Crec.INCOME), "0"))
        Crec.PILGNL = Format$(oRecordSet.Fields("U_PILGNL").Value, String$(Len(Crec.PILGNL), "0"))
        Crec.GNLOSD = Format$(oRecordSet.Fields("U_GNLOSD").Value, String$(Len(Crec.GNLOSD), "0"))
        Crec.INJBAS = Format$(oRecordSet.Fields("U_INJBAS").Value, String$(Len(Crec.INJBAS), "0"))
        Crec.INJBWO = Format$(oRecordSet.Fields("U_INJBWO").Value, String$(Len(Crec.INJBWO), "0"))
        Crec.BUYNSU = Format$(oRecordSet.Fields("U_BUYNSU").Value, String$(Len(Crec.BUYNSU), "0"))
        Crec.INJBYN = Format$(oRecordSet.Fields("U_INJBYN").Value, String$(Len(Crec.INJBYN), "0"))
        Crec.GYNGLO = Format$(oRecordSet.Fields("U_GYNGLO").Value, String$(Len(Crec.GYNGLO), "0"))
        Crec.INJGYN = Format$(oRecordSet.Fields("U_INJGYN").Value, String$(Len(Crec.INJGYN), "0"))
        Crec.JANGAE = Format$(oRecordSet.Fields("U_JANGAE").Value, String$(Len(Crec.JANGAE), "0"))
        Crec.INJJAE = Format$(oRecordSet.Fields("U_INJJAE").Value, String$(Len(Crec.INJJAE), "0"))
        Crec.INJBNY = Format$(oRecordSet.Fields("U_INJBNJ").Value, String$(Len(Crec.INJBNY), "0"))
        Crec.BUYN06 = Format$(oRecordSet.Fields("U_BUYN06").Value, String$(Len(Crec.BUYN06), "0"))
        Crec.INJSON = Format$(oRecordSet.Fields("U_INJSON").Value, String$(Len(Crec.INJSON), "0"))
        Crec.CHLSAN = Format$(oRecordSet.Fields("U_CHLSAN").Value, String$(Len(Crec.CHLSAN), "0"))
        Crec.INJCHL = Format$(oRecordSet.Fields("U_INJCHL").Value, String$(Len(Crec.INJCHL), "0"))
        Crec.FILD02 = Format$(0, String$(Len(Crec.FILD02), "0"))
        Crec.DAGYSU = Format$(oRecordSet.Fields("U_DAGYSU").Value, String$(Len(Crec.DAGYSU), "0"))  '/ 다자녀수(2007)
        Crec.INJADD = Format$(oRecordSet.Fields("U_INJADD").Value, String$(Len(Crec.INJADD), "0"))  '/ 다자녀공제금액
        Crec.KUKGON = Format$(oRecordSet.Fields("U_KUKGON").Value, String$(Len(Crec.KUKGON), "0"))  '/ 국민연금보험료공제
        Crec.YUNGON = Format$(oRecordSet.Fields("U_YUNGON").Value, String$(Len(Crec.YUNGON), "0"))  '/ 연금보험료공제(공무원)
        Crec.YUNGO1 = Format$(oRecordSet.Fields("U_YUNGO1").Value, String$(Len(Crec.YUNGO1), "0"))  '/               (군인)
        Crec.YUNGO2 = Format$(oRecordSet.Fields("U_YUNGO2").Value, String$(Len(Crec.YUNGO2), "0"))  '/               (사립)
        Crec.YUNGO3 = Format$(oRecordSet.Fields("U_YUNGO3").Value, String$(Len(Crec.YUNGO3), "0"))  '/               (별정우체국)
        Crec.GITRE2 = Format$(oRecordSet.Fields("U_GITRE2").Value, String$(Len(Crec.GITRE2), "0"))  '/ 퇴직연금(과학기술인)
        Crec.GITRET = Format$(oRecordSet.Fields("U_GITRET").Value, String$(Len(Crec.GITRET), "0"))  '/         (근로자퇴직급여보장법)
        
        Crec.PILMBH = Format$(oRecordSet.Fields("U_PILMBH").Value, String$(Len(Crec.PILMBH), "0"))
        Crec.PILGBH = Format$(oRecordSet.Fields("U_PILGBH").Value, String$(Len(Crec.PILGBH), "0"))
        Crec.PILBHM = Format$(oRecordSet.Fields("U_PILBHM").Value, String$(Len(Crec.PILBHM), "0"))
        Crec.PILJHM = Format$(oRecordSet.Fields("U_PILJHM").Value, String$(Len(Crec.PILJHM), "0"))
        Crec.PILMED = Format$(oRecordSet.Fields("U_PILMED").Value, String$(Len(Crec.PILMED), "0"))
        Crec.PILSCH = Format$(oRecordSet.Fields("U_PILSCH").Value, String$(Len(Crec.PILSCH), "0"))
        Crec.PILHUS = Format$(oRecordSet.Fields("U_PILHUS").Value, String$(Len(Crec.PILHUS), "0"))   '/ 주택임대차차입금원리상환공제금액
        Crec.PILHU2 = Format$(oRecordSet.Fields("U_PILHU2").Value, String$(Len(Crec.PILHU2), "0"))   '/ 주택임대차차입금원리상환공제금액-거주자(분리) 2011
        Crec.PILWOL = Format$(oRecordSet.Fields("U_PILWOL").Value, String$(Len(Crec.PILWOL), "0"))   '/ 월세액
        Crec.PILJHE = Format$(oRecordSet.Fields("U_PILJHE").Value, String$(Len(Crec.PILJHE), "0"))   '/ 장기주택저당차입금이자상환공제(2008)
        Crec.PILJH2 = Format$(oRecordSet.Fields("U_PILJH2").Value, String$(Len(Crec.PILJH2), "0"))   '/ 장기주택저당차입금이자상환공제(2011 분리)
        Crec.PILJH3 = Format$(oRecordSet.Fields("U_PILJH3").Value, String$(Len(Crec.PILJH3), "0"))   '/ 장기주택저당차입금이자상환공제(2011 분리)
        Crec.PILGBU = Format$(oRecordSet.Fields("U_PILGBU").Value, String$(Len(Crec.PILGBU), "0"))
        'Crec.PILHUN = Format$(oRecordSet.Fields("U_PILHUN").Value, String$(Len(Crec.PILHUN), "0"))  '/ 2009년 혼인장례 제거
        Crec.PILFLD = Format$(0, String$(Len(Crec.PILFLD), "0"))
        If oRecordSet.Fields("U_PILTOT").Value > 0 Then
            Crec.PILTOT = Format$(oRecordSet.Fields("U_PILTOT").Value, String$(Len(Crec.PILTOT), "0"))
        Else
            PILTOT = Val(oRecordSet.Fields("U_PILMBH").Value) + Val(oRecordSet.Fields("U_PILGBH").Value) + _
                     Val(oRecordSet.Fields("U_PILBHM").Value) + Val(oRecordSet.Fields("U_PILJHM").Value) + _
                     Val(oRecordSet.Fields("U_PILMED").Value) + Val(oRecordSet.Fields("U_PILSCH").Value) + _
                     Val(oRecordSet.Fields("U_PILHUS").Value) + Val(oRecordSet.Fields("U_PILHU2").Value) + _
                     Val(oRecordSet.Fields("U_PILWOL").Value) + Val(oRecordSet.Fields("U_PILJHE").Value) + _
                     Val(oRecordSet.Fields("U_PILJH2").Value) + Val(oRecordSet.Fields("U_PILJH3").Value) + Val(oRecordSet.Fields("U_PILGBU").Value)
            Crec.PILTOT = Format$(PILTOT, String$(Len(Crec.PILTOT), "0"))
        End If
        Crec.PILGON = Format$(oRecordSet.Fields("U_PILGON").Value, String$(Len(Crec.PILGON), "0"))
        Crec.CHAGAM = Format$(oRecordSet.Fields("U_CHAGAM").Value, String$(Len(Crec.CHAGAM), "0"))
        Crec.GITGYN = Format$(oRecordSet.Fields("U_GITGYN").Value, String$(Len(Crec.GITGYN), "0"))
        Crec.GITYUN = Format$(oRecordSet.Fields("U_GITYUN").Value, String$(Len(Crec.GITYUN), "0"))
        Crec.GITSGI = Format$(oRecordSet.Fields("U_GITSGI").Value, String$(Len(Crec.GITSGI), "0"))
        Crec.GITHUS = Format$(oRecordSet.Fields("U_GITHUS").Value, String$(Len(Crec.GITHUS), "0"))    '/주택마련저축(청약)
        Crec.GITHU1 = Format$(oRecordSet.Fields("U_GITHU1").Value, String$(Len(Crec.GITHU1), "0"))    '/            (종합저축)
        Crec.GITHU2 = Format$(oRecordSet.Fields("U_GITHU2").Value, String$(Len(Crec.GITHU2), "0"))    '/            (장기주택)
        Crec.GITHU3 = Format$(oRecordSet.Fields("U_GITHU3").Value, String$(Len(Crec.GITHU3), "0"))    '/            (근로자)
        Crec.GITINV = Format$(oRecordSet.Fields("U_GITINV").Value, String$(Len(Crec.GITINV), "0"))
        Crec.GITCAD = Format$(oRecordSet.Fields("U_GITCAD").Value, String$(Len(Crec.GITCAD), "0"))
'        If Val(oRecordSet.Fields("U_GITUSJ").Value) < 0 Then
'            Crec.GITUSG = 1 '우리사주조합소득공제(음수1 양수0)
'        Else
'            Crec.GITUSG = 0
'        End If
        Crec.GITUSJ = Format$(oRecordSet.Fields("U_GITUSJ").Value, String$(Len(Crec.GITUSJ), "0"))
        Crec.GITJFD = Format$(oRecordSet.Fields("U_GITJFD").Value, String$(Len(Crec.GITJFD), "0"))
        Crec.GITGYU = Format$(oRecordSet.Fields("U_GITGYU").Value, String$(Len(Crec.GITJFD), "0"))
        Crec.FILD03 = Format$(0, String$(Len(Crec.FILD03), "0"))
'        If Val(oRecordSet.Fields("U_GITTOT").Value) < 0 Then
'            Crec.GITTOG = 1
'        Else
'            Crec.GITTOG = 0 '기타소득공제계(기호 음수1, 양수0)
'        End If
        Crec.GITTOT = Format$(oRecordSet.Fields("U_GITTOT").Value, String$(Len(Crec.GITTOT), "0"))
        Crec.TAXSTD = Format$(oRecordSet.Fields("U_TAXSTD").Value, String$(Len(Crec.TAXSTD), "0"))
        Crec.SANTAX = Format$(oRecordSet.Fields("U_SANTAX").Value, String$(Len(Crec.SANTAX), "0"))
        Crec.GAMSOD = Format$(oRecordSet.Fields("U_GAMSOD").Value, String$(Len(Crec.GAMSOD), "0"))
        Crec.GAMJOS = Format$(oRecordSet.Fields("U_GAMJOS").Value, String$(Len(Crec.GAMJOS), "0"))
        Crec.GAMJYK = Format$(oRecordSet.Fields("U_GAMJYK").Value, String$(Len(Crec.GAMJYK), "0"))
        Crec.GAMFLD = String$(Len(Crec.GAMFLD), "0")
        Crec.GAMTOT = Format$(oRecordSet.Fields("U_GAMTOT").Value, String$(Len(Crec.GAMTOT), "0"))
        Crec.TAXGNL = Format$(oRecordSet.Fields("U_TAXGNL").Value, String$(Len(Crec.TAXGNL), "0"))
        Crec.NABSEE = Format$(oRecordSet.Fields("U_TAXNAB").Value, String$(Len(Crec.NABSEE), "0"))
        Crec.TAXBRO = Format$(oRecordSet.Fields("U_TAXBRO").Value, String$(Len(Crec.TAXBRO), "0"))
        Crec.TAXGBU = Format$(oRecordSet.Fields("U_TAXGBU").Value, String$(Len(Crec.TAXGBU), "0"))
        Crec.TAXFRG = Format$(oRecordSet.Fields("U_TAXFRG").Value, String$(Len(Crec.TAXFRG), "0"))
        Crec.FILD04 = Format$(0, String$(Len(Crec.FILD04), "0"))
        Crec.TAXTOT = Format$(oRecordSet.Fields("U_TAXTOT").Value, String$(Len(Crec.TAXTOT), "0"))
        Crec.GULGAB = Format$(oRecordSet.Fields("U_GULGAB").Value, String$(Len(Crec.GULGAB), "0"))
        Crec.GULJUM = Format$(oRecordSet.Fields("U_GULJUM").Value, String$(Len(Crec.GULJUM), "0"))
        Crec.GULNON = Format$(oRecordSet.Fields("U_GULNON").Value, String$(Len(Crec.GULNON), "0"))
        'Crec.GULTOT = Format$(oRecordSet.Fields("U_GULGAB").Value + oRecordSet.Fields("U_GULJUM").Value + oRecordSet.Fields("U_GULNON").Value, String$(Len(Crec.GULTOT), "0"))
        Crec.NANGAB = Format$(oRecordSet.Fields("U_NANGAB").Value, String$(Len(Crec.NANGAB), "0"))
        Crec.NANJUM = Format$(oRecordSet.Fields("U_NANJUM").Value, String$(Len(Crec.NANJUM), "0"))
        Crec.NANNON = Format$(oRecordSet.Fields("U_NANNON").Value, String$(Len(Crec.NANNON), "0"))
        'Crec.NANTOT = Format$(oRecordSet.Fields("U_NANGAB").Value + oRecordSet.Fields("U_NANJUM").Value + oRecordSet.Fields("U_NANNON").Value, String$(Len(Crec.NANTOT), "0"))
        '/ 차감징수세액(2009년추가)
        If Val(oRecordSet.Fields("U_CHAGAB").Value) < 0 Then
            Crec.CHAGAG = 1
            Crec.CHAGAB = Format$(-1 * oRecordSet.Fields("U_CHAGAB").Value, String$(Len(Crec.CHAGAB), "0"))   '기현)차감소득세
        Else
            Crec.CHAGAG = 0 '현)차감소득세(기호 음수1, 양수0)
            Crec.CHAGAB = Format$(oRecordSet.Fields("U_CHAGAB").Value, String$(Len(Crec.CHAGAB), "0"))   '기현)차감소득세
        End If
        If Val(oRecordSet.Fields("U_CHAJUM").Value) < 0 Then
            Crec.CHAJUG = 1
            Crec.CHAJUM = Format$(-1 * oRecordSet.Fields("U_CHAJUM").Value, String$(Len(Crec.CHAJUM), "0"))   '기현)차감소득세
        Else
            Crec.CHAJUG = 0 '현)차감주민세(기호 음수1, 양수0)
            Crec.CHAJUM = Format$(oRecordSet.Fields("U_CHAJUM").Value, String$(Len(Crec.CHAJUM), "0"))   '기현)차감소득세
        End If
        If Val(oRecordSet.Fields("U_CHANON").Value) < 0 Then
            Crec.CHANOG = 1
            Crec.CHANON = Format$(-1 * oRecordSet.Fields("U_CHANON").Value, String$(Len(Crec.CHANON), "0"))   '기현)차감소득세
        Else
            Crec.CHANOG = 0 '현)차감주민세(기호 음수1, 양수0)
            Crec.CHANON = Format$(oRecordSet.Fields("U_CHANON").Value, String$(Len(Crec.CHANON), "0"))   '기현)차감소득세
        End If
        
        If (Val(oRecordSet.Fields("U_CHAGAB").Value) + Val(oRecordSet.Fields("U_CHAJUM").Value) + Val(oRecordSet.Fields("U_CHANON").Value)) < 0 Then
            'Crec.CHATOG = 1
            'Crec.CHATOT = Format$(-1 * (Val(oRecordSet.Fields("U_CHAGAB").Value) + Val(oRecordSet.Fields("U_CHAJUM").Value) + Val(oRecordSet.Fields("U_CHANON").Value)), String$(Len(Crec.CHATOT), "0"))   '기현)차감소득세
        Else
            'Crec.CHATOG = 0 '현)차감세액계(기호 음수1, 양수0)
            'Crec.CHATOT = Format$((Val(oRecordSet.Fields("U_CHAGAB").Value) + Val(oRecordSet.Fields("U_CHAJUM").Value) + Val(oRecordSet.Fields("U_CHANON").Value)), String$(Len(Crec.CHATOT), "0"))   '기현)차감소득세
        End If
        
        Crec.FILLER = Space$(Len(Crec.FILLER))
        
      
        Print #1, MDC_SetMod.sStr(Crec.RECGBN) & MDC_SetMod.sStr(Crec.DTAGBN) & MDC_SetMod.sStr(Crec.TAXCOD) & MDC_SetMod.sStr(Crec.SQNNBR) & MDC_SetMod.sStr(Crec.BUSNBR) & MDC_SetMod.sStr(Crec.JONCNT) & MDC_SetMod.sStr(Crec.DWEGBN) & _
                MDC_SetMod.sStr(Crec.RGNCOD) & MDC_SetMod.sStr(Crec.RGNTAX) & MDC_SetMod.sStr(Crec.MSTNAM) & MDC_SetMod.sStr(Crec.INTGBN) & MDC_SetMod.sStr(Crec.PERNBR) & MDC_SetMod.sStr(Crec.GUKCOD) & MDC_SetMod.sStr(Crec.HUSMAN) & MDC_SetMod.sStr(Crec.JSNGBN) & MDC_SetMod.sStr(Crec.BUSNB1) & _
                MDC_SetMod.sStr(Crec.SANGHO) & MDC_SetMod.sStr(Crec.STRINT) & MDC_SetMod.sStr(Crec.ENDINT) & MDC_SetMod.sStr(Crec.STRGAM) & MDC_SetMod.sStr(Crec.ENDGAM) & MDC_SetMod.sStr(Crec.PAYAMT) & _
                MDC_SetMod.sStr(Crec.BNSAMT) & MDC_SetMod.sStr(Crec.INJBNS) & MDC_SetMod.sStr(Crec.JUSBNS) & MDC_SetMod.sStr(Crec.URIBNS) & MDC_SetMod.sStr(Crec.FILD01) & MDC_SetMod.sStr(Crec.TOTAMT) & _
                MDC_SetMod.sStr(Crec.BIGG01) & MDC_SetMod.sStr(Crec.BIGH01) & MDC_SetMod.sStr(Crec.BIGH05) & MDC_SetMod.sStr(Crec.BIGH06) & MDC_SetMod.sStr(Crec.BIGH07) & MDC_SetMod.sStr(Crec.BIGH08) & MDC_SetMod.sStr(Crec.BIGH09) & _
                MDC_SetMod.sStr(Crec.BIGH10) & MDC_SetMod.sStr(Crec.BIGH11) & MDC_SetMod.sStr(Crec.BIGH12) & MDC_SetMod.sStr(Crec.BIGH13) & MDC_SetMod.sStr(Crec.BIGI01) & MDC_SetMod.sStr(Crec.BIGK01) & MDC_SetMod.sStr(Crec.BIGM01) & _
                MDC_SetMod.sStr(Crec.BIGM02) & MDC_SetMod.sStr(Crec.BIGM03) & MDC_SetMod.sStr(Crec.BIGO01) & MDC_SetMod.sStr(Crec.BIGQ01) & MDC_SetMod.sStr(Crec.BIGR10) & MDC_SetMod.sStr(Crec.BIGS01) & MDC_SetMod.sStr(Crec.BIGT01) & MDC_SetMod.sStr(Crec.BIGY02) & _
                MDC_SetMod.sStr(Crec.BIGY03) & MDC_SetMod.sStr(Crec.BIGY21) & MDC_SetMod.sStr(Crec.BIGZ01) & MDC_SetMod.sStr(Crec.BIGY22) & MDC_SetMod.sStr(Crec.BIGTOT) & MDC_SetMod.sStr(Crec.BIGTO1) & _
                MDC_SetMod.sStr(Crec.INCOME) & MDC_SetMod.sStr(Crec.PILGNL) & MDC_SetMod.sStr(Crec.GNLOSD) & _
                MDC_SetMod.sStr(Crec.INJBAS) & MDC_SetMod.sStr(Crec.INJBWO) & MDC_SetMod.sStr(Crec.BUYNSU) & MDC_SetMod.sStr(Crec.INJBYN) & MDC_SetMod.sStr(Crec.GYNGLO) & MDC_SetMod.sStr(Crec.INJGYN) & MDC_SetMod.sStr(Crec.JANGAE) & _
                MDC_SetMod.sStr(Crec.INJJAE) & MDC_SetMod.sStr(Crec.INJBNY) & MDC_SetMod.sStr(Crec.BUYN06) & MDC_SetMod.sStr(Crec.INJSON) & MDC_SetMod.sStr(Crec.CHLSAN) & MDC_SetMod.sStr(Crec.INJCHL) & MDC_SetMod.sStr(Crec.FILD02) & _
                MDC_SetMod.sStr(Crec.DAGYSU) & MDC_SetMod.sStr(Crec.INJADD) & MDC_SetMod.sStr(Crec.KUKGON) & MDC_SetMod.sStr(Crec.YUNGON) & MDC_SetMod.sStr(Crec.YUNGO1) & MDC_SetMod.sStr(Crec.YUNGO2) & MDC_SetMod.sStr(Crec.YUNGO3) & _
                MDC_SetMod.sStr(Crec.GITRE2) & MDC_SetMod.sStr(Crec.GITRET) & MDC_SetMod.sStr(Crec.PILMBH) & MDC_SetMod.sStr(Crec.PILGBH) & MDC_SetMod.sStr(Crec.PILBHM) & MDC_SetMod.sStr(Crec.PILJHM) & MDC_SetMod.sStr(Crec.PILMED) & _
                MDC_SetMod.sStr(Crec.PILSCH) & MDC_SetMod.sStr(Crec.PILHUS) & MDC_SetMod.sStr(Crec.PILHU2) & MDC_SetMod.sStr(Crec.PILWOL) & MDC_SetMod.sStr(Crec.PILJHE) & MDC_SetMod.sStr(Crec.PILJH2) & MDC_SetMod.sStr(Crec.PILJH3) & MDC_SetMod.sStr(Crec.PILGBU) & MDC_SetMod.sStr(Crec.PILFLD) & MDC_SetMod.sStr(Crec.PILTOT) & _
                MDC_SetMod.sStr(Crec.PILGON) & MDC_SetMod.sStr(Crec.CHAGAM) & MDC_SetMod.sStr(Crec.GITGYN) & MDC_SetMod.sStr(Crec.GITYUN) & MDC_SetMod.sStr(Crec.GITSGI) & MDC_SetMod.sStr(Crec.GITHUS) & MDC_SetMod.sStr(Crec.GITHU1) & _
                MDC_SetMod.sStr(Crec.GITHU2) & MDC_SetMod.sStr(Crec.GITHU3) & MDC_SetMod.sStr(Crec.GITINV) & MDC_SetMod.sStr(Crec.GITCAD) & MDC_SetMod.sStr(Crec.GITUSJ) & MDC_SetMod.sStr(Crec.GITJFD) & MDC_SetMod.sStr(Crec.GITGYU) & _
                MDC_SetMod.sStr(Crec.FILD03) & MDC_SetMod.sStr(Crec.GITTOT) & MDC_SetMod.sStr(Crec.TAXSTD) & MDC_SetMod.sStr(Crec.SANTAX) & MDC_SetMod.sStr(Crec.GAMSOD) & MDC_SetMod.sStr(Crec.GAMJOS) & MDC_SetMod.sStr(Crec.GAMJYK) & _
                MDC_SetMod.sStr(Crec.GAMFLD) & MDC_SetMod.sStr(Crec.GAMTOT) & MDC_SetMod.sStr(Crec.TAXGNL) & MDC_SetMod.sStr(Crec.NABSEE) & MDC_SetMod.sStr(Crec.TAXBRO) & MDC_SetMod.sStr(Crec.TAXGBU) & MDC_SetMod.sStr(Crec.TAXFRG) & _
                MDC_SetMod.sStr(Crec.FILD04) & MDC_SetMod.sStr(Crec.TAXTOT) & MDC_SetMod.sStr(Crec.GULGAB) & MDC_SetMod.sStr(Crec.GULJUM) & MDC_SetMod.sStr(Crec.GULNON) & MDC_SetMod.sStr(Crec.NANGAB) & _
                MDC_SetMod.sStr(Crec.NANJUM) & MDC_SetMod.sStr(Crec.NANNON) & MDC_SetMod.sStr(Crec.CHAGAG) & MDC_SetMod.sStr(Crec.CHAGAB) & MDC_SetMod.sStr(Crec.CHAJUG) & MDC_SetMod.sStr(Crec.CHAJUM) & _
                MDC_SetMod.sStr(Crec.CHANOG) & MDC_SetMod.sStr(Crec.CHANON) & MDC_SetMod.sStr(Crec.FILLER)

           
              
        '/ 기본공제수
        If Val(Crec.INJBWO) > 0 Then
            C_BUYCNT = 1
        Else
            C_BUYCNT = 0
        End If
        C_BUYCNT = C_BUYCNT + Val(Crec.BUYNSU)
        Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " 생성 완료.", True)
     '/ 필수입력 체크
        If Crec.DWEGBN = "0" Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "-거주자구분이 누락되었습니다. 확인하여 주십시오.", False, True)
            CheckC = True
        End If
        If Crec.DWEGBN <> "1" And Trim$(Crec.RGNCOD) = "" Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "-비거주자일경우 거주지국코드는 필수입니다. 확인하여 주십시오.", False, True)
            CheckC = True
        End If
        If Crec.INTGBN <> "1" And Trim$(Crec.GUKCOD) = "" Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "-외국인일경우 국적코드는 필수입니다. 확인하여 주십시오.", False, True)
            CheckC = True
        End If
        If Crec.INTGBN = "1" And Len(Trim$(Crec.PERNBR)) <> 13 Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "-주민등록번호를 확인하여 주십시오.", False, True)
            CheckC = True
        End If
        If OLDBIG <> 0 Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "-비과세금액 세부코드이관작업을 하셔야합니다.", False, True)
            CheckC = True
        End If
        
        '/ D레코드: 종전근무처 레코드
        If Val(Crec.JONCNT) > 0 Then
            If File_Create_DRecord = False Then
                ErrNum = 2
                GoTo Error_Message
            End If
        End If
        '/ E레코드: 부양가족 레코드
        '/ 외국인은 단일세율을 적용할 경우 수록 안함
        If Crec.INTGBN = "1" Or (Crec.INTGBN = "9" And Crec.RGNTAX = "2") Then
            If File_Create_ERecord = False Then
                ErrNum = 3
                GoTo Error_Message
            End If
        End If
        If Val(Crec.GITRET) <> 0 Or Val(Crec.GITRE2) <> 0 Or Val(Crec.GITGYN) <> 0 Or Val(Crec.GITYUN) <> 0 Or _
           Val(Crec.GITHUS) <> 0 Or Val(Crec.GITHU1) <> 0 Or Val(Crec.GITHU2) <> 0 Or Val(Crec.GITHU3) <> 0 Or _
           Val(Crec.GITJFD) <> 0 Then
            If File_Create_FRecord = False Then
                ErrNum = 4
                GoTo Error_Message
            End If
        End If
        oRecordSet.MoveNext
    Loop
        
        
    If CheckC = False Then
        File_Create_CRecord = True
    Else
        File_Create_CRecord = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "주(현)근무처 레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "D레코드(종전근무처 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "E레코드(부양가족 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "F레코드(연금.저축 레코드) 생성 실패.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("C레코드오류: " & Err.Description, False)
    End If
    File_Create_CRecord = False
End Function

Private Function File_Create_DRecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckD      As String
    Dim JONCNT      As Integer
    
    CheckD = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 종전근무지정보
    sQry = "SELECT      T0.* "
    sQry = sQry & " FROM [@ZPY502L] T0 INNER JOIN [@ZPY502H] T1 ON T0.DocEntry = T1.DocEntry"
    sQry = sQry & " WHERE    T1.U_JSNYER = '" & oJsnYear & "'"
    sQry = sQry & " AND      T1.U_MSTCOD = '" & C_MSTCOD & "'"
    sQry = sQry & " AND      T1.U_CLTCOD = '" & C_CLTCOD & "'"
    sQry = sQry & " ORDER BY T0.LineID"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    JONCNT = 0
    Do Until oRecordSet.EOF
        '/ D Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
      OLDCNT = OLDCNT + 1   '/ 전사원 종전근무지수
      JONCNT = JONCNT + 1   '/ 해당사원의 종전근무지일련번호
      Drec.RECGBN = "D"
      Drec.DTAGBN = "20"
      Drec.TAXCOD = Arec.TAXCOD
      Drec.SQNNBR = Format$(NEWCNT, String$(Len(Drec.SQNNBR), "0"))
      Drec.BUSNBR = Brec.BUSNBR
      'Drec.FILD01 = Format$(0, String$(Len(Drec.FILD01), "0"))
      Drec.FILD01 = Space$(Len(Drec.FILD01))
      Drec.PERNBR = Crec.PERNBR
      
      Drec.TAXJOH = "2"     '// 납세조합 구분
      Drec.JONNAM = Trim$(oRecordSet.Fields("U_JONNAM").Value)
      Drec.JONNBR = Replace(oRecordSet.Fields("U_JONNBR").Value, "-", "")
      Drec.STRINT = Replace(oRecordSet.Fields("U_JONSTR").Value, "-", "")
      Drec.ENDINT = Replace(oRecordSet.Fields("U_JONEND").Value, "-", "")
      Drec.STRGAM = Format$(Replace(oRecordSet.Fields("U_JONGFR").Value, "-", ""), String$(Len(Drec.STRGAM), "0"))
      Drec.ENDGAM = Format$(Replace(oRecordSet.Fields("U_JONGTO").Value, "-", ""), String$(Len(Drec.STRGAM), "0"))
      
      Drec.PAYAMT = Format$(oRecordSet.Fields("U_JONPAY").Value, String$(Len(Drec.PAYAMT), "0"))
      Drec.BNSAMT = Format$(oRecordSet.Fields("U_JONBNS").Value, String$(Len(Drec.BNSAMT), "0"))
      Drec.INJBNS = Format$(oRecordSet.Fields("U_INJBNS").Value, String$(Len(Drec.INJBNS), "0"))
      Drec.JUSBNS = Format$(oRecordSet.Fields("U_JONJUS").Value, String$(Len(Drec.JUSBNS), "0"))
      Drec.URIBNS = Format$(oRecordSet.Fields("U_URIBNS").Value, String$(Len(Drec.URIBNS), "0"))
      Drec.FILD02 = Format$(0, String$(Len(Drec.FILD02), "0"))
      Drec.TOTAMT = Format$(oRecordSet.Fields("U_JONPAY").Value + oRecordSet.Fields("U_JONBNS").Value + oRecordSet.Fields("U_INJBNS").Value + oRecordSet.Fields("U_JONJUS").Value + oRecordSet.Fields("U_URIBNS").Value, String$(Len(Drec.TOTAMT), "0"))
      
      Drec.BIGG01 = Format$(oRecordSet.Fields("U_JBTG01").Value, String$(Len(Drec.BIGG01), "0"))
      Drec.BIGH01 = Format$(oRecordSet.Fields("U_JBTH01").Value, String$(Len(Drec.BIGH01), "0"))
      Drec.BIGH05 = Format$(oRecordSet.Fields("U_JBTH05").Value, String$(Len(Drec.BIGH05), "0"))
      Drec.BIGH06 = Format$(oRecordSet.Fields("U_JBTH06").Value, String$(Len(Drec.BIGH06), "0"))
      Drec.BIGH07 = Format$(oRecordSet.Fields("U_JBTH07").Value, String$(Len(Drec.BIGH07), "0"))
      Drec.BIGH08 = Format$(oRecordSet.Fields("U_JBTH08").Value, String$(Len(Drec.BIGH08), "0"))
      Drec.BIGH09 = Format$(oRecordSet.Fields("U_JBTH09").Value, String$(Len(Drec.BIGH09), "0"))
      Drec.BIGH10 = Format$(oRecordSet.Fields("U_JBTH10").Value, String$(Len(Drec.BIGH10), "0"))
      Drec.BIGH11 = Format$(oRecordSet.Fields("U_JBTH11").Value, String$(Len(Drec.BIGH11), "0"))
      Drec.BIGH12 = Format$(oRecordSet.Fields("U_JBTH12").Value, String$(Len(Drec.BIGH12), "0"))
      Drec.BIGH13 = Format$(oRecordSet.Fields("U_JBTH13").Value, String$(Len(Drec.BIGH13), "0"))
      Drec.BIGI01 = Format$(oRecordSet.Fields("U_JBTI01").Value, String$(Len(Drec.BIGI01), "0"))
      Drec.BIGK01 = Format$(oRecordSet.Fields("U_JBTK01").Value, String$(Len(Drec.BIGK01), "0"))
      Drec.BIGM01 = Format$(oRecordSet.Fields("U_JBTM01").Value, String$(Len(Drec.BIGM01), "0"))
      Drec.BIGM02 = Format$(oRecordSet.Fields("U_JBTM02").Value, String$(Len(Drec.BIGM02), "0"))
      Drec.BIGM03 = Format$(oRecordSet.Fields("U_JBTM03").Value, String$(Len(Drec.BIGM03), "0"))
      Drec.BIGO01 = Format$(oRecordSet.Fields("U_JBTO01").Value, String$(Len(Drec.BIGO01), "0"))
      Drec.BIGQ01 = Format$(oRecordSet.Fields("U_JBTQ01").Value, String$(Len(Drec.BIGQ01), "0"))
      Drec.BIGR10 = Format$(oRecordSet.Fields("U_JBTY01").Value, String$(Len(Drec.BIGR10), "0"))
      Drec.BIGS01 = Format$(oRecordSet.Fields("U_JBTS01").Value, String$(Len(Drec.BIGS01), "0"))
      Drec.BIGT01 = Format$(oRecordSet.Fields("U_JBTT01").Value, String$(Len(Drec.BIGT01), "0"))
'      Drec.BIGX01 = Format$(oRecordSet.Fields("U_JBTX01").Value, String$(Len(Drec.BIGX01), "0"))
      Drec.BIGY02 = Format$(oRecordSet.Fields("U_JBTY02").Value, String$(Len(Drec.BIGY02), "0"))
      Drec.BIGY03 = Format$(oRecordSet.Fields("U_JBTY03").Value, String$(Len(Drec.BIGY03), "0"))
      Drec.BIGY21 = Format$(oRecordSet.Fields("U_JBTY21").Value, String$(Len(Drec.BIGY21), "0"))
      Drec.BIGZ01 = Format$(oRecordSet.Fields("U_JBTZ01").Value, String$(Len(Drec.BIGZ01), "0"))
      Drec.BIGY22 = Format$(oRecordSet.Fields("U_JBTY20").Value, String$(Len(Drec.BIGY22), "0"))
      Drec.BIGTOT = Format$(oRecordSet.Fields("U_JBTTOT").Value - oRecordSet.Fields("U_JBTT01").Value - oRecordSet.Fields("U_JBTZ01").Value, String$(Len(Drec.BIGTOT), "0"))
      Drec.BIGTO1 = Format$(oRecordSet.Fields("U_JBTT01").Value + oRecordSet.Fields("U_JBTZ01").Value, String$(Len(Drec.BIGTO1), "0"))
      
      Drec.NANGAB = Format$(oRecordSet.Fields("U_JONGAB").Value, String$(Len(Drec.NANGAB), "0"))
      Drec.NANJUM = Format$(oRecordSet.Fields("U_JONJUM").Value, String$(Len(Drec.NANJUM), "0"))
      Drec.NANNON = Format$(oRecordSet.Fields("U_JONNON").Value, String$(Len(Drec.NANNON), "0"))
'      Drec.NANTOT = Format$(oRecordSet.Fields("U_JONGAB").Value + oRecordSet.Fields("U_JONJUM").Value + oRecordSet.Fields("U_JONNON").Value, String$(Len(Drec.NANTOT), "0"))
  
      Drec.JONCNT = Format$(JONCNT, String$(Len(Drec.JONCNT), "0"))
      Drec.FILLER = Space$(Len(Drec.FILLER))
      Print #1, MDC_SetMod.sStr(Drec.RECGBN) & MDC_SetMod.sStr(Drec.DTAGBN) & MDC_SetMod.sStr(Drec.TAXCOD) & _
                MDC_SetMod.sStr(Drec.SQNNBR) & MDC_SetMod.sStr(Drec.BUSNBR) & MDC_SetMod.sStr(Drec.FILD01) & _
                MDC_SetMod.sStr(Drec.PERNBR) & MDC_SetMod.sStr(Drec.TAXJOH) & MDC_SetMod.sStr(Drec.JONNAM) & _
                MDC_SetMod.sStr(Drec.JONNBR) & MDC_SetMod.sStr(Drec.STRINT) & MDC_SetMod.sStr(Drec.ENDINT) & _
                MDC_SetMod.sStr(Drec.STRGAM) & MDC_SetMod.sStr(Drec.ENDGAM) & MDC_SetMod.sStr(Drec.PAYAMT) & _
                MDC_SetMod.sStr(Drec.BNSAMT) & MDC_SetMod.sStr(Drec.INJBNS) & MDC_SetMod.sStr(Drec.JUSBNS) & _
                MDC_SetMod.sStr(Drec.URIBNS) & MDC_SetMod.sStr(Drec.FILD02) & MDC_SetMod.sStr(Drec.TOTAMT) & _
                MDC_SetMod.sStr(Drec.BIGG01) & MDC_SetMod.sStr(Drec.BIGH01) & MDC_SetMod.sStr(Drec.BIGH05) & _
                MDC_SetMod.sStr(Drec.BIGH06) & MDC_SetMod.sStr(Drec.BIGH07) & MDC_SetMod.sStr(Drec.BIGH08) & _
                MDC_SetMod.sStr(Drec.BIGH09) & MDC_SetMod.sStr(Drec.BIGH10) & MDC_SetMod.sStr(Drec.BIGH11) & _
                MDC_SetMod.sStr(Drec.BIGH12) & MDC_SetMod.sStr(Drec.BIGH13) & MDC_SetMod.sStr(Drec.BIGI01) & _
                MDC_SetMod.sStr(Drec.BIGK01) & MDC_SetMod.sStr(Drec.BIGM01) & MDC_SetMod.sStr(Drec.BIGM02) & _
                MDC_SetMod.sStr(Drec.BIGM03) & MDC_SetMod.sStr(Drec.BIGO01) & MDC_SetMod.sStr(Drec.BIGQ01) & MDC_SetMod.sStr(Drec.BIGR10) & _
                MDC_SetMod.sStr(Drec.BIGS01) & MDC_SetMod.sStr(Drec.BIGT01) & MDC_SetMod.sStr(Drec.BIGY02) & _
                MDC_SetMod.sStr(Drec.BIGY03) & MDC_SetMod.sStr(Drec.BIGY21) & MDC_SetMod.sStr(Drec.BIGZ01) & _
                MDC_SetMod.sStr(Drec.BIGY22) & MDC_SetMod.sStr(Drec.BIGTOT) & MDC_SetMod.sStr(Drec.BIGTO1) & _
                MDC_SetMod.sStr(Drec.NANGAB) & MDC_SetMod.sStr(Drec.NANJUM) & MDC_SetMod.sStr(Drec.NANNON) & _
                MDC_SetMod.sStr(Drec.JONCNT) & MDC_SetMod.sStr(Drec.FILLER)
        
     '/ 필수입력 체크
        If Trim$(Drec.JONNAM) = "" Then
            Call Matrix_AddRow("D레코드:종전근무처 법인명(상호)가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckD = True
        End If
        If Trim$(Drec.JONNBR) = "" Then
            Call Matrix_AddRow("D레코드:종전근무처 사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckD = True
        End If
        
        oRecordSet.MoveNext
    Loop
    
    Call Matrix_AddRow("D레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " 종전근무지 생성 완료.", True)
    
    If CheckD = False Then
        File_Create_DRecord = True
    Else
        File_Create_DRecord = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "종전근무지레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("D레코드오류: " & Err.Description, False)
    End If
    File_Create_DRecord = False
End Function

Private Function File_Create_ERecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckE      As String
    Dim BUYCNT      As Integer
    Dim FAMCNT      As Integer
    Dim i           As Integer
    
    Dim TOTBOH      As Double
    Dim TOTMED      As Double
    Dim TOTEDC      As Double
    Dim TOTCAD      As Double
    Dim TOTCA1      As Double
    Dim TOTCSH      As Double
    Dim TOTGBU      As Double
    Dim TOTAMT      As Double
    CheckE = False  '/체크필요유무
    ErrNum = 0
    E_BUYCNT = 0
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 부양가족정보
    sQry = " Exec RPY504_2  '" & oJsnYear & "', '01', '12', '3', N'" & C_CLTCOD & "', N'%',  N'%','" & C_MSTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        '/ E Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        Erec.RECGBN = "E"
        Erec.DTAGBN = "20"
        Erec.TAXCOD = Arec.TAXCOD
        Erec.SQNNBR = Format$(NEWCNT, String$(Len(Erec.SQNNBR), "0"))
        Erec.BUSNBR = Brec.BUSNBR
        Erec.PERNBR = Crec.PERNBR
        Erec.FILLER = Space$(Len(Erec.FILLER))
         
        TOTBOH = 0: TOTMED = 0: TOTEDC = 0: TOTCAD = 0: TOTCSH = 0: TOTGBU = 0
         
        BUYCNT = 0
        FAMCNT = 1
        Do Until oRecordSet.EOF
            TOTAMT = 0
            TOTAMT = Val(oRecordSet.Fields("U_BOHAMT1").Value) + Val(oRecordSet.Fields("U_BOHAMT2").Value) + _
                     Val(oRecordSet.Fields("U_EDCAMT1").Value) + Val(oRecordSet.Fields("U_EDCAMT2").Value) + _
                     Val(oRecordSet.Fields("U_MEDAMT1").Value) + Val(oRecordSet.Fields("U_MEDAMT2").Value) + _
                     Val(oRecordSet.Fields("U_CADAMT1").Value) + Val(oRecordSet.Fields("U_CADAMT2").Value) + _
                     Val(oRecordSet.Fields("U_CSHCAD1").Value) + Val(oRecordSet.Fields("U_CSHCAD2").Value) + _
                     Val(oRecordSet.Fields("U_CSHAMT1").Value) + _
                     Val(oRecordSet.Fields("U_GBUAMT1").Value) + Val(oRecordSet.Fields("U_GBUAMT2").Value)
            If (Val(oRecordSet.Fields("U_CHKBAS").Value) = 1 Or Val(oRecordSet.Fields("U_CHKCHL").Value) = 1) Or (TOTAMT > 0) Then
                BUYCNT = BUYCNT + 1   '/ 해당사원의 부양가족일련번호
                If Val(oRecordSet.Fields("U_CHKBAS").Value) = 1 Then
                    E_BUYCNT = E_BUYCNT + 1
                End If
                '/초기화
                If BUYCNT = 1 Then
                    For i = 1 To 5
                        Erec.CHKCOD(i) = Space$(Len(Erec.CHKCOD(i)))
                        Erec.CHKINT(i) = Space$(Len(Erec.CHKINT(i)))
                        Erec.CHKNAM(i) = Space$(Len(Erec.CHKNAM(i)))
                        Erec.CHKPER(i) = Space$(Len(Erec.CHKPER(i)))
                        Erec.CHKBAS(i) = Space$(Len(Erec.CHKBAS(i)))
                        Erec.CHKJAN(i) = Space$(Len(Erec.CHKJAN(i)))
                        Erec.CHKBY6(i) = Space$(Len(Erec.CHKBY6(i)))
                        Erec.CHKBUY(i) = Space$(Len(Erec.CHKBUY(i)))
                        Erec.CHKJEL(i) = Space$(Len(Erec.CHKJEL(i)))
                       ' Erec.CHKDAG(i) = Space$(Len(Erec.CHKDAG(i)))
                        Erec.CHKCHS(i) = Space$(Len(Erec.CHKCHS(i)))
                        
                        Erec.BOHAM1(i) = Format$(0, String$(Len(Erec.BOHAM1(i)), "0"))
                        Erec.MEDAM1(i) = Format$(0, String$(Len(Erec.MEDAM1(i)), "0"))
                        Erec.EDCAM1(i) = Format$(0, String$(Len(Erec.EDCAM1(i)), "0"))
                        Erec.CADAM1(i) = Format$(0, String$(Len(Erec.CADAM1(i)), "0"))
                        Erec.CSHCA1(i) = Format$(0, String$(Len(Erec.CSHCA1(i)), "0"))
                        Erec.CSHAM1(i) = Format$(0, String$(Len(Erec.CSHAM1(i)), "0"))
                        Erec.GBUAM1(i) = Format$(0, String$(Len(Erec.GBUAM1(i)), "0"))
                        Erec.BOHAM2(i) = Format$(0, String$(Len(Erec.BOHAM2(i)), "0"))
                        Erec.MEDAM2(i) = Format$(0, String$(Len(Erec.MEDAM2(i)), "0"))
                        Erec.EDCAM2(i) = Format$(0, String$(Len(Erec.EDCAM2(i)), "0"))
                        Erec.CADAM2(i) = Format$(0, String$(Len(Erec.CADAM2(i)), "0"))
                        Erec.CSHCA2(i) = Format$(0, String$(Len(Erec.CSHCA2(i)), "0"))
                        Erec.GBUAM2(i) = Format$(0, String$(Len(Erec.GBUAM2(i)), "0"))
                    Next i
                End If
                Erec.CHKCOD(BUYCNT) = Trim$(oRecordSet.Fields("U_CHKCOD").Value)
                Erec.CHKINT(BUYCNT) = oRecordSet.Fields("U_CHKINT").Value  '/ 내외국인
                Erec.CHKNAM(BUYCNT) = Trim$(oRecordSet.Fields("U_FAMNAM").Value)
                Erec.CHKPER(BUYCNT) = Replace(oRecordSet.Fields("U_FAMPER").Value, "-", "")
                Erec.CHKBAS(BUYCNT) = IIf(oRecordSet.Fields("U_CHKBAS").Value = "0", "", oRecordSet.Fields("U_CHKBAS").Value)
                Erec.CHKJAN(BUYCNT) = IIf(oRecordSet.Fields("U_CHKJAN").Value = "0", "", oRecordSet.Fields("U_CHKJAN").Value)
                Erec.CHKBY6(BUYCNT) = IIf(oRecordSet.Fields("U_CHKCHL").Value = "0", "", oRecordSet.Fields("U_CHKCHL").Value)
                Erec.CHKBUY(BUYCNT) = IIf(oRecordSet.Fields("U_CHKBUY").Value = "0", "", oRecordSet.Fields("U_CHKBUY").Value)
                Erec.CHKJEL(BUYCNT) = IIf(oRecordSet.Fields("U_CHKJEL").Value = "0", "", oRecordSet.Fields("U_CHKJEL").Value)
                'Erec.CHKDAG(BUYCNT) = IIf(oRecordSet.Fields("U_CHKDAJ").Value = "0", "", oRecordSet.Fields("U_CHKDAJ").Value)
                Erec.CHKCHS(BUYCNT) = IIf(oRecordSet.Fields("U_CHKCHS").Value = "0", "", oRecordSet.Fields("U_CHKCHS").Value)
                
                Erec.BOHAM1(BUYCNT) = Format$(oRecordSet.Fields("U_BOHAMT1").Value, String$(Len(Erec.BOHAM1(BUYCNT)), "0"))
                Erec.MEDAM1(BUYCNT) = Format$(oRecordSet.Fields("U_MEDAMT1").Value, String$(Len(Erec.MEDAM1(BUYCNT)), "0"))
                Erec.EDCAM1(BUYCNT) = Format$(oRecordSet.Fields("U_EDCAMT1").Value, String$(Len(Erec.EDCAM1(BUYCNT)), "0"))
                Erec.CADAM1(BUYCNT) = Format$(oRecordSet.Fields("U_CADAMT1").Value, String$(Len(Erec.CADAM1(BUYCNT)), "0"))
                Erec.CSHCA1(BUYCNT) = Format$(oRecordSet.Fields("U_CSHCAD1").Value, String$(Len(Erec.CSHCA1(BUYCNT)), "0"))
                Erec.CSHAM1(BUYCNT) = Format$(oRecordSet.Fields("U_CSHAMT1").Value, String$(Len(Erec.CSHAM1(BUYCNT)), "0"))
                Erec.GBUAM1(BUYCNT) = Format$(oRecordSet.Fields("U_GBUAMT1").Value, String$(Len(Erec.GBUAM1(BUYCNT)), "0"))
                Erec.BOHAM2(BUYCNT) = Format$(oRecordSet.Fields("U_BOHAMT2").Value, String$(Len(Erec.BOHAM2(BUYCNT)), "0"))
                Erec.MEDAM2(BUYCNT) = Format$(oRecordSet.Fields("U_MEDAMT2").Value, String$(Len(Erec.MEDAM2(BUYCNT)), "0"))
                Erec.EDCAM2(BUYCNT) = Format$(oRecordSet.Fields("U_EDCAMT2").Value, String$(Len(Erec.EDCAM2(BUYCNT)), "0"))
                Erec.CADAM2(BUYCNT) = Format$(oRecordSet.Fields("U_CADAMT2").Value, String$(Len(Erec.CADAM2(BUYCNT)), "0"))
                Erec.CSHCA2(BUYCNT) = Format$(oRecordSet.Fields("U_CSHCAD2").Value, String$(Len(Erec.CSHCA2(BUYCNT)), "0"))
                Erec.GBUAM2(BUYCNT) = Format$(oRecordSet.Fields("U_GBUAMT2").Value, String$(Len(Erec.GBUAM2(BUYCNT)), "0"))
                '/ 부양가족명세서상 총계
                TOTBOH = TOTBOH + Val(oRecordSet.Fields("U_BOHAMT1").Value) + Val(oRecordSet.Fields("U_BOHAMT2").Value)
                TOTMED = TOTMED + Val(oRecordSet.Fields("U_MEDAMT1").Value) + Val(oRecordSet.Fields("U_MEDAMT2").Value)
                TOTEDC = TOTEDC + Val(oRecordSet.Fields("U_EDCAMT1").Value) + Val(oRecordSet.Fields("U_EDCAMT2").Value)
                TOTCAD = TOTCAD + Val(oRecordSet.Fields("U_CADAMT1").Value) + Val(oRecordSet.Fields("U_CADAMT2").Value)
                TOTCA1 = TOTCA1 + Val(oRecordSet.Fields("U_CSHCAD1").Value) + Val(oRecordSet.Fields("U_CSHCAD2").Value)
                TOTCSH = TOTCSH + Val(oRecordSet.Fields("U_CSHAMT1").Value)
                TOTGBU = TOTGBU + Val(oRecordSet.Fields("U_GBUAMT1").Value) + Val(oRecordSet.Fields("U_GBUAMT2").Value)
                
                'E_BUYCNT = E_BUYCNT + Val(Erec.CHKBAS(BUYCNT))
                '/ 필수입력 체크
                If Trim$(Erec.CHKPER(BUYCNT)) = "" Then
                    Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & Trim$(Erec.CHKNAM(BUYCNT)) & "주민등록번호가 누락되었습니다. 입력하여 주십시오.", True, True)
                    CheckE = True
                End If
                If Trim$(Erec.CHKNAM(BUYCNT)) = "" Then
                    Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & BUYCNT & ")번째 가족의 성명이 누락되었습니다. 입력하여 주십시오.", True, True)
                    CheckE = True
                End If
                If Trim$(Erec.CHKINT(BUYCNT)) = "" Then
                    Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & Trim$(Erec.CHKNAM(BUYCNT)) & "의 내외국인구분이 누락되었습니다. 입력하여 주십시오.", True, True)
                    CheckE = True
                End If
                If Trim$(Erec.CHKCOD(BUYCNT)) = "" Then
                    Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & Trim$(Erec.CHKNAM(BUYCNT)) & "의 관계코드가 누락되었습니다. 입력하여 주십시오.", True, True)
                    CheckE = True
                End If
                '/ 마지막줄이면 다음줄로
                If BUYCNT = 5 Then
                    Erec.FAMNBR = Format$(FAMCNT, String$(Len(Erec.FAMNBR), "0"))
                    '/ E레코드삽입
                    Print #1, MDC_SetMod.sStr(Erec.RECGBN) & MDC_SetMod.sStr(Erec.DTAGBN) & MDC_SetMod.sStr(Erec.TAXCOD) & MDC_SetMod.sStr(Erec.SQNNBR) & MDC_SetMod.sStr(Erec.BUSNBR) & MDC_SetMod.sStr(Erec.PERNBR) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(1)) & MDC_SetMod.sStr(Erec.CHKINT(1)) & MDC_SetMod.sStr(Erec.CHKNAM(1)) & MDC_SetMod.sStr(Erec.CHKPER(1)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(1)) & MDC_SetMod.sStr(Erec.CHKJAN(1)) & MDC_SetMod.sStr(Erec.CHKBY6(1)) & MDC_SetMod.sStr(Erec.CHKBUY(1)) & MDC_SetMod.sStr(Erec.CHKJEL(1)) & MDC_SetMod.sStr(Erec.CHKCHS(1)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(1)) & MDC_SetMod.sStr(Erec.MEDAM1(1)) & MDC_SetMod.sStr(Erec.EDCAM1(1)) & MDC_SetMod.sStr(Erec.CADAM1(1)) & MDC_SetMod.sStr(Erec.CSHCA1(1)) & MDC_SetMod.sStr(Erec.CSHAM1(1)) & MDC_SetMod.sStr(Erec.GBUAM1(1)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(1)) & MDC_SetMod.sStr(Erec.MEDAM2(1)) & MDC_SetMod.sStr(Erec.EDCAM2(1)) & MDC_SetMod.sStr(Erec.CADAM2(1)) & MDC_SetMod.sStr(Erec.CSHCA2(1)) & MDC_SetMod.sStr(Erec.GBUAM2(1)) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(2)) & MDC_SetMod.sStr(Erec.CHKINT(2)) & MDC_SetMod.sStr(Erec.CHKNAM(2)) & MDC_SetMod.sStr(Erec.CHKPER(2)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(2)) & MDC_SetMod.sStr(Erec.CHKJAN(2)) & MDC_SetMod.sStr(Erec.CHKBY6(2)) & MDC_SetMod.sStr(Erec.CHKBUY(2)) & MDC_SetMod.sStr(Erec.CHKJEL(2)) & MDC_SetMod.sStr(Erec.CHKCHS(2)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(2)) & MDC_SetMod.sStr(Erec.MEDAM1(2)) & MDC_SetMod.sStr(Erec.EDCAM1(2)) & MDC_SetMod.sStr(Erec.CADAM1(2)) & MDC_SetMod.sStr(Erec.CSHCA1(2)) & MDC_SetMod.sStr(Erec.CSHAM1(2)) & MDC_SetMod.sStr(Erec.GBUAM1(2)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(2)) & MDC_SetMod.sStr(Erec.MEDAM2(2)) & MDC_SetMod.sStr(Erec.EDCAM2(2)) & MDC_SetMod.sStr(Erec.CADAM2(2)) & MDC_SetMod.sStr(Erec.CSHCA2(2)) & MDC_SetMod.sStr(Erec.GBUAM2(2)) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(3)) & MDC_SetMod.sStr(Erec.CHKINT(3)) & MDC_SetMod.sStr(Erec.CHKNAM(3)) & MDC_SetMod.sStr(Erec.CHKPER(3)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(3)) & MDC_SetMod.sStr(Erec.CHKJAN(3)) & MDC_SetMod.sStr(Erec.CHKBY6(3)) & MDC_SetMod.sStr(Erec.CHKBUY(3)) & MDC_SetMod.sStr(Erec.CHKJEL(3)) & MDC_SetMod.sStr(Erec.CHKCHS(3)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(3)) & MDC_SetMod.sStr(Erec.MEDAM1(3)) & MDC_SetMod.sStr(Erec.EDCAM1(3)) & MDC_SetMod.sStr(Erec.CADAM1(3)) & MDC_SetMod.sStr(Erec.CSHCA1(3)) & MDC_SetMod.sStr(Erec.CSHAM1(3)) & MDC_SetMod.sStr(Erec.GBUAM1(3)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(3)) & MDC_SetMod.sStr(Erec.MEDAM2(3)) & MDC_SetMod.sStr(Erec.EDCAM2(3)) & MDC_SetMod.sStr(Erec.CADAM2(3)) & MDC_SetMod.sStr(Erec.CSHCA2(3)) & MDC_SetMod.sStr(Erec.GBUAM2(3)) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(4)) & MDC_SetMod.sStr(Erec.CHKINT(4)) & MDC_SetMod.sStr(Erec.CHKNAM(4)) & MDC_SetMod.sStr(Erec.CHKPER(4)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(4)) & MDC_SetMod.sStr(Erec.CHKJAN(4)) & MDC_SetMod.sStr(Erec.CHKBY6(4)) & MDC_SetMod.sStr(Erec.CHKBUY(4)) & MDC_SetMod.sStr(Erec.CHKJEL(4)) & MDC_SetMod.sStr(Erec.CHKCHS(4)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(4)) & MDC_SetMod.sStr(Erec.MEDAM1(4)) & MDC_SetMod.sStr(Erec.EDCAM1(4)) & MDC_SetMod.sStr(Erec.CADAM1(4)) & MDC_SetMod.sStr(Erec.CSHCA1(4)) & MDC_SetMod.sStr(Erec.CSHAM1(4)) & MDC_SetMod.sStr(Erec.GBUAM1(4)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(4)) & MDC_SetMod.sStr(Erec.MEDAM2(4)) & MDC_SetMod.sStr(Erec.EDCAM2(4)) & MDC_SetMod.sStr(Erec.CADAM2(4)) & MDC_SetMod.sStr(Erec.CSHCA2(4)) & MDC_SetMod.sStr(Erec.GBUAM2(4)) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(5)) & MDC_SetMod.sStr(Erec.CHKINT(5)) & MDC_SetMod.sStr(Erec.CHKNAM(5)) & MDC_SetMod.sStr(Erec.CHKPER(5)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(5)) & MDC_SetMod.sStr(Erec.CHKJAN(5)) & MDC_SetMod.sStr(Erec.CHKBY6(5)) & MDC_SetMod.sStr(Erec.CHKBUY(5)) & MDC_SetMod.sStr(Erec.CHKJEL(5)) & MDC_SetMod.sStr(Erec.CHKCHS(5)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(5)) & MDC_SetMod.sStr(Erec.MEDAM1(5)) & MDC_SetMod.sStr(Erec.EDCAM1(5)) & MDC_SetMod.sStr(Erec.CADAM1(5)) & MDC_SetMod.sStr(Erec.CSHCA1(5)) & MDC_SetMod.sStr(Erec.CSHAM1(5)) & MDC_SetMod.sStr(Erec.GBUAM1(5)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(5)) & MDC_SetMod.sStr(Erec.MEDAM2(5)) & MDC_SetMod.sStr(Erec.EDCAM2(5)) & MDC_SetMod.sStr(Erec.CADAM2(5)) & MDC_SetMod.sStr(Erec.CSHCA2(5)) & MDC_SetMod.sStr(Erec.GBUAM2(5)) & _
                                MDC_SetMod.sStr(Erec.FAMNBR) & MDC_SetMod.sStr(Erec.FILLER)
                    '/ 다음줄넘김
                    BUYCNT = 0
                    FAMCNT = FAMCNT + 1
                End If
            End If
            oRecordSet.MoveNext
        Loop
        If BUYCNT > 0 Then
            Erec.FAMNBR = Format$(FAMCNT, String$(Len(Erec.FAMNBR), "0"))
            '/ E레코드삽입
                    Print #1, MDC_SetMod.sStr(Erec.RECGBN) & MDC_SetMod.sStr(Erec.DTAGBN) & MDC_SetMod.sStr(Erec.TAXCOD) & MDC_SetMod.sStr(Erec.SQNNBR) & MDC_SetMod.sStr(Erec.BUSNBR) & MDC_SetMod.sStr(Erec.PERNBR) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(1)) & MDC_SetMod.sStr(Erec.CHKINT(1)) & MDC_SetMod.sStr(Erec.CHKNAM(1)) & MDC_SetMod.sStr(Erec.CHKPER(1)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(1)) & MDC_SetMod.sStr(Erec.CHKJAN(1)) & MDC_SetMod.sStr(Erec.CHKBY6(1)) & MDC_SetMod.sStr(Erec.CHKBUY(1)) & MDC_SetMod.sStr(Erec.CHKJEL(1)) & MDC_SetMod.sStr(Erec.CHKCHS(1)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(1)) & MDC_SetMod.sStr(Erec.MEDAM1(1)) & MDC_SetMod.sStr(Erec.EDCAM1(1)) & MDC_SetMod.sStr(Erec.CADAM1(1)) & MDC_SetMod.sStr(Erec.CSHCA1(1)) & MDC_SetMod.sStr(Erec.CSHAM1(1)) & MDC_SetMod.sStr(Erec.GBUAM1(1)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(1)) & MDC_SetMod.sStr(Erec.MEDAM2(1)) & MDC_SetMod.sStr(Erec.EDCAM2(1)) & MDC_SetMod.sStr(Erec.CADAM2(1)) & MDC_SetMod.sStr(Erec.CSHCA2(1)) & MDC_SetMod.sStr(Erec.GBUAM2(1)) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(2)) & MDC_SetMod.sStr(Erec.CHKINT(2)) & MDC_SetMod.sStr(Erec.CHKNAM(2)) & MDC_SetMod.sStr(Erec.CHKPER(2)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(2)) & MDC_SetMod.sStr(Erec.CHKJAN(2)) & MDC_SetMod.sStr(Erec.CHKBY6(2)) & MDC_SetMod.sStr(Erec.CHKBUY(2)) & MDC_SetMod.sStr(Erec.CHKJEL(2)) & MDC_SetMod.sStr(Erec.CHKCHS(2)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(2)) & MDC_SetMod.sStr(Erec.MEDAM1(2)) & MDC_SetMod.sStr(Erec.EDCAM1(2)) & MDC_SetMod.sStr(Erec.CADAM1(2)) & MDC_SetMod.sStr(Erec.CSHCA1(2)) & MDC_SetMod.sStr(Erec.CSHAM1(2)) & MDC_SetMod.sStr(Erec.GBUAM1(2)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(2)) & MDC_SetMod.sStr(Erec.MEDAM2(2)) & MDC_SetMod.sStr(Erec.EDCAM2(2)) & MDC_SetMod.sStr(Erec.CADAM2(2)) & MDC_SetMod.sStr(Erec.CSHCA2(2)) & MDC_SetMod.sStr(Erec.GBUAM2(2)) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(3)) & MDC_SetMod.sStr(Erec.CHKINT(3)) & MDC_SetMod.sStr(Erec.CHKNAM(3)) & MDC_SetMod.sStr(Erec.CHKPER(3)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(3)) & MDC_SetMod.sStr(Erec.CHKJAN(3)) & MDC_SetMod.sStr(Erec.CHKBY6(3)) & MDC_SetMod.sStr(Erec.CHKBUY(3)) & MDC_SetMod.sStr(Erec.CHKJEL(3)) & MDC_SetMod.sStr(Erec.CHKCHS(3)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(3)) & MDC_SetMod.sStr(Erec.MEDAM1(3)) & MDC_SetMod.sStr(Erec.EDCAM1(3)) & MDC_SetMod.sStr(Erec.CADAM1(3)) & MDC_SetMod.sStr(Erec.CSHCA1(3)) & MDC_SetMod.sStr(Erec.CSHAM1(3)) & MDC_SetMod.sStr(Erec.GBUAM1(3)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(3)) & MDC_SetMod.sStr(Erec.MEDAM2(3)) & MDC_SetMod.sStr(Erec.EDCAM2(3)) & MDC_SetMod.sStr(Erec.CADAM2(3)) & MDC_SetMod.sStr(Erec.CSHCA2(3)) & MDC_SetMod.sStr(Erec.GBUAM2(3)) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(4)) & MDC_SetMod.sStr(Erec.CHKINT(4)) & MDC_SetMod.sStr(Erec.CHKNAM(4)) & MDC_SetMod.sStr(Erec.CHKPER(4)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(4)) & MDC_SetMod.sStr(Erec.CHKJAN(4)) & MDC_SetMod.sStr(Erec.CHKBY6(4)) & MDC_SetMod.sStr(Erec.CHKBUY(4)) & MDC_SetMod.sStr(Erec.CHKJEL(4)) & MDC_SetMod.sStr(Erec.CHKCHS(4)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(4)) & MDC_SetMod.sStr(Erec.MEDAM1(4)) & MDC_SetMod.sStr(Erec.EDCAM1(4)) & MDC_SetMod.sStr(Erec.CADAM1(4)) & MDC_SetMod.sStr(Erec.CSHCA1(4)) & MDC_SetMod.sStr(Erec.CSHAM1(4)) & MDC_SetMod.sStr(Erec.GBUAM1(4)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(4)) & MDC_SetMod.sStr(Erec.MEDAM2(4)) & MDC_SetMod.sStr(Erec.EDCAM2(4)) & MDC_SetMod.sStr(Erec.CADAM2(4)) & MDC_SetMod.sStr(Erec.CSHCA2(4)) & MDC_SetMod.sStr(Erec.GBUAM2(4)) & _
                                MDC_SetMod.sStr(Erec.CHKCOD(5)) & MDC_SetMod.sStr(Erec.CHKINT(5)) & MDC_SetMod.sStr(Erec.CHKNAM(5)) & MDC_SetMod.sStr(Erec.CHKPER(5)) & _
                                MDC_SetMod.sStr(Erec.CHKBAS(5)) & MDC_SetMod.sStr(Erec.CHKJAN(5)) & MDC_SetMod.sStr(Erec.CHKBY6(5)) & MDC_SetMod.sStr(Erec.CHKBUY(5)) & MDC_SetMod.sStr(Erec.CHKJEL(5)) & MDC_SetMod.sStr(Erec.CHKCHS(5)) & _
                                MDC_SetMod.sStr(Erec.BOHAM1(5)) & MDC_SetMod.sStr(Erec.MEDAM1(5)) & MDC_SetMod.sStr(Erec.EDCAM1(5)) & MDC_SetMod.sStr(Erec.CADAM1(5)) & MDC_SetMod.sStr(Erec.CSHCA1(5)) & MDC_SetMod.sStr(Erec.CSHAM1(5)) & MDC_SetMod.sStr(Erec.GBUAM1(5)) & _
                                MDC_SetMod.sStr(Erec.BOHAM2(5)) & MDC_SetMod.sStr(Erec.MEDAM2(5)) & MDC_SetMod.sStr(Erec.EDCAM2(5)) & MDC_SetMod.sStr(Erec.CADAM2(5)) & MDC_SetMod.sStr(Erec.CSHCA2(5)) & MDC_SetMod.sStr(Erec.GBUAM2(5)) & _
                                MDC_SetMod.sStr(Erec.FAMNBR) & MDC_SetMod.sStr(Erec.FILLER)
            Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " 부양가족레코드 생성 완료.", True)
        End If
    End If
    '/ 기본부양가족공제 기본공제가족수와 가족명세서의 기본공제가족수가 다를경우
    If (E_BUYCNT - 1) <> C_BUYCNT Then
        Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " 기본공제받은 부양가족이 누락되었습니다.", True, True)
        CheckE = False
    End If
    '/ 기본부양가족공제의 금액과 소득공제항목등록한 총공제금액이 다를경우
    sQry = "SELECT  SUM(ISNULL(U_BOHAMT,0)+ISNULL(U_JGABOA,0) + ISNULL(U_BOHAL1,0) + ISNULL(U_BOHAL2,0)) AS BOHAMT,"
    sQry = sQry & " SUM(ISNULL(U_JGAMED,0)+ISNULL(U_GENMED,0)) AS MEDAMT,"
    sQry = sQry & " SUM(ISNULL(U_BONSCH,0)+ISNULL(U_JGASCH,0)+ISNULL(U_JICSCH,0)) AS EDCAMT,"
    sQry = sQry & " SUM(ISNULL(U_CADSAV,0)+ISNULL(U_GIRSAV,0)) AS CADAMT,"
    sQry = sQry & " SUM(ISNULL(U_CA1SAV,0)) AS CA1SAV,"
    sQry = sQry & " SUM(ISNULL(U_CSHSAV,0)) AS CSHAMT,"
    sQry = sQry & " SUM(IsNull(U_LAWGBU, 0) + IsNull(U_POCGBU, 0) + IsNull(U_SP1GBU, 0) + IsNull(U_SP2GBU, 0) + IsNull(U_USJGBU, 0) + IsNull(U_JIJGBU, 0)++ IsNull(U_JNGGBU, 0)) As GBUAMT"
    sQry = sQry & " FROM [@ZPY501H] WHERE U_JSNYER = '" & oJsnYear & "' AND U_MSTCOD = '" & C_MSTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        '/1. 보험료총액이 다를경우
        If TOTBOH < Val(oRecordSet.Fields("BOHAMT").Value) Then
            Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " [소득공제항목]이 부양가족명세의 보험료보다 보험료총계가 큽니다. 확인하십시오.", True, True)
        End If
        '/2. 의료비총액이 다를경우
        If TOTMED <> Val(oRecordSet.Fields("MEDAMT").Value) Then
            Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " [소득공제항목]과 부양가족명세의 의료비총계가 다릅니다. 확인하십시오.", True, True)
        End If
        '/3. 교육비총액이 다를경우
        If TOTEDC < Val(oRecordSet.Fields("EDCAMT").Value) Then
            Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " [소득공제항목]이 부양가족명세의 교육비보다 교육비총계가 큽니다. 확인하십시오.", True, True)
        End If
        '/4. 신용카드총액이 다를경우
        If TOTCAD <> Val(oRecordSet.Fields("CADAMT").Value) Then
            Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " [소득공제항목]과 부양가족명세의 신용카드등총계가 다릅니다. 확인하십시오.", True, True)
        End If
        '/5. 직불카드총액이 다를경우
        If TOTCA1 <> Val(oRecordSet.Fields("CA1SAV").Value) Then
            Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " [소득공제항목]과 부양가족명세의 직불카드총계가 다릅니다. 확인하십시오.", True, True)
        End If
        '/6. 현금영수증총액이 다를경우
        If TOTCSH <> Val(oRecordSet.Fields("CSHAMT").Value) Then
            Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " [소득공제항목]과 부양가족명세의 현금영수증총계가 다릅니다. 확인하십시오.", True, True)
        End If
        '/7. 기부금총액이 다를경우
        If TOTGBU < Val(oRecordSet.Fields("GBUAMT").Value) Then
            Call Matrix_AddRow("E레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " [소득공제항목]이 부양가족명세의 기부금총계보다 큽니다. 확인하십시오.", True, True)
        End If
    End If
    If CheckE = False Then
        File_Create_ERecord = True
    Else
        File_Create_ERecord = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "부양가족레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("E레코드오류: " & Err.Source & " " & Err.Description, False)
    End If
    File_Create_ERecord = False
End Function

Private Function File_Create_FRecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckF      As String
    Dim SAVCNT      As Integer
    Dim iRow        As Long
    
    Dim RETSAV      As Double:  Dim RETSA1      As Double:  Dim GYNSAV      As Double:  Dim YUNSAV      As Double
    Dim HUSAMT      As Double:  Dim HU1AMT      As Double:  Dim HU2AMT      As Double:  Dim HU3AMT      As Double
    Dim JFDAM1      As Double:  Dim JFDAM2      As Double:  Dim JFDAM3      As Double
    
    CheckF = False  '/체크필요유무
    ErrNum = 0
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 부양가족정보
    sQry = "EXEC ZPY521 'F', '" & oJsnYear & "', '" & JSNGBN & "', '" & STRMON & "', '" & ENDMON & "', " & _
                            "'" & C_CLTCOD & "', '" & MSTBRK & "', '" & DPTSTR & "', '" & DPTEND & "','" & C_MSTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        '/ F Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        Frec.RECGBN = "F"
        Frec.DTAGBN = "20"
        Frec.TAXCOD = Arec.TAXCOD
        Frec.SQNNBR = Format$(NEWCNT, String$(Len(Frec.SQNNBR), "0"))
        Frec.BUSNBR = Brec.BUSNBR
        Frec.PERNBR = Crec.PERNBR
        Frec.FILLER = Space$(Len(Frec.FILLER))
         
        SAVCNT = 0
        Do Until oRecordSet.EOF

            SAVCNT = SAVCNT + 1   '/ 해당사원의 부양가족일련번호
            '/초기화
            If SAVCNT = 1 Then
                For iRow = 1 To 15
                    Frec.SAVGBN(iRow) = Space$(Len(Frec.SAVGBN(iRow)))
                    Frec.SAVCOD(iRow) = Space$(Len(Frec.SAVCOD(iRow)))
                    Frec.SAVNAM(iRow) = Space$(Len(Frec.SAVNAM(iRow)))
                    Frec.SAVNUM(iRow) = Space$(Len(Frec.SAVNUM(iRow)))
                    
                    Frec.STYEAR(iRow) = Format$(0, String$(Len(Frec.STYEAR(iRow)), "0"))
                    Frec.SAVAMT(iRow) = Format$(0, String$(Len(Frec.SAVAMT(iRow)), "0"))
                    Frec.SARAMT(iRow) = Format$(0, String$(Len(Frec.SARAMT(iRow)), "0"))

                Next iRow
            End If
            
            Frec.SAVGBN(SAVCNT) = Trim$(oRecordSet.Fields("U_SAVGBN").Value)
            Frec.SAVCOD(SAVCNT) = Trim$(oRecordSet.Fields("U_SAVCOD").Value)
            Frec.SAVNAM(SAVCNT) = Trim$(oRecordSet.Fields("U_SAVNAM").Value)
            Frec.SAVNUM(SAVCNT) = Trim$(oRecordSet.Fields("U_SAVNUM").Value)
            
            Frec.STYEAR(SAVCNT) = Format$(oRecordSet.Fields("U_STYEAR").Value, String$(Len(Frec.STYEAR(SAVCNT)), "0"))
            Frec.SAVAMT(SAVCNT) = Format$(oRecordSet.Fields("U_SAVAMT").Value, String$(Len(Frec.SAVAMT(SAVCNT)), "0"))
            Frec.SARAMT(SAVCNT) = Format$(oRecordSet.Fields("U_SARAMT").Value, String$(Len(Frec.SARAMT(SAVCNT)), "0"))

            '/ 부양가족명세서상 총계
            Select Case Trim$(oRecordSet.Fields("U_SAVGBN").Value)
            
            Case "11":      RETSAV = RETSAV + oRecordSet.Fields("U_SAVAMT").Value
            Case "12":      RETSA1 = RETSA1 + oRecordSet.Fields("U_SAVAMT").Value
            Case "21":      GYNSAV = GYNSAV + oRecordSet.Fields("U_SAVAMT").Value
            Case "22":      YUNSAV = YUNSAV + oRecordSet.Fields("U_SAVAMT").Value
            Case "31":      HUSAMT = HUSAMT + oRecordSet.Fields("U_SAVAMT").Value
            Case "32":      HU1AMT = HU1AMT + oRecordSet.Fields("U_SAVAMT").Value
            Case "33":      HU2AMT = HU2AMT + oRecordSet.Fields("U_SAVAMT").Value
            Case "34":      HU3AMT = HU3AMT + oRecordSet.Fields("U_SAVAMT").Value
            Case "41":
                Select Case oRecordSet.Fields("U_STYEAR").Value
                Case "01":  JFDAM1 = JFDAM1 + oRecordSet.Fields("U_SAVAMT").Value
                Case "02":  JFDAM2 = JFDAM2 + oRecordSet.Fields("U_SAVAMT").Value
                Case "03":  JFDAM3 = JFDAM3 + oRecordSet.Fields("U_SAVAMT").Value
                End Select
            End Select
            
            '/ 필수입력 체크
            If Trim$(Frec.SAVGBN(SAVCNT)) = "41" And Trim$(Frec.STYEAR(SAVCNT)) = "00" Then
                Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & Trim$(Frec.SAVNAM(SAVCNT)) & " 납입년차가 누락되었습니다. 입력하여 주십시오.", True, True)
                CheckF = True
            End If

            '/ 마지막줄이면 다음줄로
            If SAVCNT = 15 Then
                '/ E레코드삽입
                Print #1, MDC_SetMod.sStr(Frec.RECGBN) & MDC_SetMod.sStr(Frec.DTAGBN) & MDC_SetMod.sStr(Frec.TAXCOD) & MDC_SetMod.sStr(Frec.SQNNBR) & MDC_SetMod.sStr(Frec.BUSNBR) & MDC_SetMod.sStr(Frec.PERNBR) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(1)) & MDC_SetMod.sStr(Frec.SAVCOD(1)) & MDC_SetMod.sStr(Frec.SAVNAM(1)) & MDC_SetMod.sStr(Frec.SAVNUM(1)) & MDC_SetMod.sStr(Frec.STYEAR(1)) & MDC_SetMod.sStr(Frec.SAVAMT(1)) & MDC_SetMod.sStr(Frec.SARAMT(1)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(2)) & MDC_SetMod.sStr(Frec.SAVCOD(2)) & MDC_SetMod.sStr(Frec.SAVNAM(2)) & MDC_SetMod.sStr(Frec.SAVNUM(2)) & MDC_SetMod.sStr(Frec.STYEAR(2)) & MDC_SetMod.sStr(Frec.SAVAMT(2)) & MDC_SetMod.sStr(Frec.SARAMT(2)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(3)) & MDC_SetMod.sStr(Frec.SAVCOD(3)) & MDC_SetMod.sStr(Frec.SAVNAM(3)) & MDC_SetMod.sStr(Frec.SAVNUM(3)) & MDC_SetMod.sStr(Frec.STYEAR(3)) & MDC_SetMod.sStr(Frec.SAVAMT(3)) & MDC_SetMod.sStr(Frec.SARAMT(3)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(4)) & MDC_SetMod.sStr(Frec.SAVCOD(4)) & MDC_SetMod.sStr(Frec.SAVNAM(4)) & MDC_SetMod.sStr(Frec.SAVNUM(4)) & MDC_SetMod.sStr(Frec.STYEAR(4)) & MDC_SetMod.sStr(Frec.SAVAMT(4)) & MDC_SetMod.sStr(Frec.SARAMT(4)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(5)) & MDC_SetMod.sStr(Frec.SAVCOD(5)) & MDC_SetMod.sStr(Frec.SAVNAM(5)) & MDC_SetMod.sStr(Frec.SAVNUM(5)) & MDC_SetMod.sStr(Frec.STYEAR(5)) & MDC_SetMod.sStr(Frec.SAVAMT(5)) & MDC_SetMod.sStr(Frec.SARAMT(5)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(6)) & MDC_SetMod.sStr(Frec.SAVCOD(6)) & MDC_SetMod.sStr(Frec.SAVNAM(6)) & MDC_SetMod.sStr(Frec.SAVNUM(6)) & MDC_SetMod.sStr(Frec.STYEAR(6)) & MDC_SetMod.sStr(Frec.SAVAMT(6)) & MDC_SetMod.sStr(Frec.SARAMT(6)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(7)) & MDC_SetMod.sStr(Frec.SAVCOD(7)) & MDC_SetMod.sStr(Frec.SAVNAM(7)) & MDC_SetMod.sStr(Frec.SAVNUM(7)) & MDC_SetMod.sStr(Frec.STYEAR(7)) & MDC_SetMod.sStr(Frec.SAVAMT(7)) & MDC_SetMod.sStr(Frec.SARAMT(7)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(8)) & MDC_SetMod.sStr(Frec.SAVCOD(8)) & MDC_SetMod.sStr(Frec.SAVNAM(8)) & MDC_SetMod.sStr(Frec.SAVNUM(8)) & MDC_SetMod.sStr(Frec.STYEAR(8)) & MDC_SetMod.sStr(Frec.SAVAMT(8)) & MDC_SetMod.sStr(Frec.SARAMT(8)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(9)) & MDC_SetMod.sStr(Frec.SAVCOD(9)) & MDC_SetMod.sStr(Frec.SAVNAM(9)) & MDC_SetMod.sStr(Frec.SAVNUM(9)) & MDC_SetMod.sStr(Frec.STYEAR(9)) & MDC_SetMod.sStr(Frec.SAVAMT(9)) & MDC_SetMod.sStr(Frec.SARAMT(9)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(10)) & MDC_SetMod.sStr(Frec.SAVCOD(10)) & MDC_SetMod.sStr(Frec.SAVNAM(10)) & MDC_SetMod.sStr(Frec.SAVNUM(10)) & MDC_SetMod.sStr(Frec.STYEAR(10)) & MDC_SetMod.sStr(Frec.SAVAMT(10)) & MDC_SetMod.sStr(Frec.SARAMT(10)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(11)) & MDC_SetMod.sStr(Frec.SAVCOD(11)) & MDC_SetMod.sStr(Frec.SAVNAM(11)) & MDC_SetMod.sStr(Frec.SAVNUM(11)) & MDC_SetMod.sStr(Frec.STYEAR(11)) & MDC_SetMod.sStr(Frec.SAVAMT(11)) & MDC_SetMod.sStr(Frec.SARAMT(11)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(12)) & MDC_SetMod.sStr(Frec.SAVCOD(12)) & MDC_SetMod.sStr(Frec.SAVNAM(12)) & MDC_SetMod.sStr(Frec.SAVNUM(12)) & MDC_SetMod.sStr(Frec.STYEAR(12)) & MDC_SetMod.sStr(Frec.SAVAMT(12)) & MDC_SetMod.sStr(Frec.SARAMT(12)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(13)) & MDC_SetMod.sStr(Frec.SAVCOD(13)) & MDC_SetMod.sStr(Frec.SAVNAM(13)) & MDC_SetMod.sStr(Frec.SAVNUM(13)) & MDC_SetMod.sStr(Frec.STYEAR(13)) & MDC_SetMod.sStr(Frec.SAVAMT(13)) & MDC_SetMod.sStr(Frec.SARAMT(13)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(14)) & MDC_SetMod.sStr(Frec.SAVCOD(14)) & MDC_SetMod.sStr(Frec.SAVNAM(14)) & MDC_SetMod.sStr(Frec.SAVNUM(14)) & MDC_SetMod.sStr(Frec.STYEAR(14)) & MDC_SetMod.sStr(Frec.SAVAMT(14)) & MDC_SetMod.sStr(Frec.SARAMT(14)) & _
                        MDC_SetMod.sStr(Frec.SAVGBN(15)) & MDC_SetMod.sStr(Frec.SAVCOD(15)) & MDC_SetMod.sStr(Frec.SAVNAM(15)) & MDC_SetMod.sStr(Frec.SAVNUM(15)) & MDC_SetMod.sStr(Frec.STYEAR(15)) & MDC_SetMod.sStr(Frec.SAVAMT(15)) & MDC_SetMod.sStr(Frec.SARAMT(15)) & _
                        MDC_SetMod.sStr(Frec.FILLER)
                '/ 다음줄넘김
                SAVCNT = 0
            End If

            oRecordSet.MoveNext
        Loop
        If SAVCNT > 0 Then
            '/ E레코드삽입
            Print #1, MDC_SetMod.sStr(Frec.RECGBN) & MDC_SetMod.sStr(Frec.DTAGBN) & MDC_SetMod.sStr(Frec.TAXCOD) & MDC_SetMod.sStr(Frec.SQNNBR) & MDC_SetMod.sStr(Frec.BUSNBR) & MDC_SetMod.sStr(Frec.PERNBR) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(1)) & MDC_SetMod.sStr(Frec.SAVCOD(1)) & MDC_SetMod.sStr(Frec.SAVNAM(1)) & MDC_SetMod.sStr(Frec.SAVNUM(1)) & MDC_SetMod.sStr(Frec.STYEAR(1)) & MDC_SetMod.sStr(Frec.SAVAMT(1)) & MDC_SetMod.sStr(Frec.SARAMT(1)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(2)) & MDC_SetMod.sStr(Frec.SAVCOD(2)) & MDC_SetMod.sStr(Frec.SAVNAM(2)) & MDC_SetMod.sStr(Frec.SAVNUM(2)) & MDC_SetMod.sStr(Frec.STYEAR(2)) & MDC_SetMod.sStr(Frec.SAVAMT(2)) & MDC_SetMod.sStr(Frec.SARAMT(2)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(3)) & MDC_SetMod.sStr(Frec.SAVCOD(3)) & MDC_SetMod.sStr(Frec.SAVNAM(3)) & MDC_SetMod.sStr(Frec.SAVNUM(3)) & MDC_SetMod.sStr(Frec.STYEAR(3)) & MDC_SetMod.sStr(Frec.SAVAMT(3)) & MDC_SetMod.sStr(Frec.SARAMT(3)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(4)) & MDC_SetMod.sStr(Frec.SAVCOD(4)) & MDC_SetMod.sStr(Frec.SAVNAM(4)) & MDC_SetMod.sStr(Frec.SAVNUM(4)) & MDC_SetMod.sStr(Frec.STYEAR(4)) & MDC_SetMod.sStr(Frec.SAVAMT(4)) & MDC_SetMod.sStr(Frec.SARAMT(4)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(5)) & MDC_SetMod.sStr(Frec.SAVCOD(5)) & MDC_SetMod.sStr(Frec.SAVNAM(5)) & MDC_SetMod.sStr(Frec.SAVNUM(5)) & MDC_SetMod.sStr(Frec.STYEAR(5)) & MDC_SetMod.sStr(Frec.SAVAMT(5)) & MDC_SetMod.sStr(Frec.SARAMT(5)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(6)) & MDC_SetMod.sStr(Frec.SAVCOD(6)) & MDC_SetMod.sStr(Frec.SAVNAM(6)) & MDC_SetMod.sStr(Frec.SAVNUM(6)) & MDC_SetMod.sStr(Frec.STYEAR(6)) & MDC_SetMod.sStr(Frec.SAVAMT(6)) & MDC_SetMod.sStr(Frec.SARAMT(6)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(7)) & MDC_SetMod.sStr(Frec.SAVCOD(7)) & MDC_SetMod.sStr(Frec.SAVNAM(7)) & MDC_SetMod.sStr(Frec.SAVNUM(7)) & MDC_SetMod.sStr(Frec.STYEAR(7)) & MDC_SetMod.sStr(Frec.SAVAMT(7)) & MDC_SetMod.sStr(Frec.SARAMT(7)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(8)) & MDC_SetMod.sStr(Frec.SAVCOD(8)) & MDC_SetMod.sStr(Frec.SAVNAM(8)) & MDC_SetMod.sStr(Frec.SAVNUM(8)) & MDC_SetMod.sStr(Frec.STYEAR(8)) & MDC_SetMod.sStr(Frec.SAVAMT(8)) & MDC_SetMod.sStr(Frec.SARAMT(8)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(9)) & MDC_SetMod.sStr(Frec.SAVCOD(9)) & MDC_SetMod.sStr(Frec.SAVNAM(9)) & MDC_SetMod.sStr(Frec.SAVNUM(9)) & MDC_SetMod.sStr(Frec.STYEAR(9)) & MDC_SetMod.sStr(Frec.SAVAMT(9)) & MDC_SetMod.sStr(Frec.SARAMT(9)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(10)) & MDC_SetMod.sStr(Frec.SAVCOD(10)) & MDC_SetMod.sStr(Frec.SAVNAM(10)) & MDC_SetMod.sStr(Frec.SAVNUM(10)) & MDC_SetMod.sStr(Frec.STYEAR(10)) & MDC_SetMod.sStr(Frec.SAVAMT(10)) & MDC_SetMod.sStr(Frec.SARAMT(10)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(11)) & MDC_SetMod.sStr(Frec.SAVCOD(11)) & MDC_SetMod.sStr(Frec.SAVNAM(11)) & MDC_SetMod.sStr(Frec.SAVNUM(11)) & MDC_SetMod.sStr(Frec.STYEAR(11)) & MDC_SetMod.sStr(Frec.SAVAMT(11)) & MDC_SetMod.sStr(Frec.SARAMT(11)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(12)) & MDC_SetMod.sStr(Frec.SAVCOD(12)) & MDC_SetMod.sStr(Frec.SAVNAM(12)) & MDC_SetMod.sStr(Frec.SAVNUM(12)) & MDC_SetMod.sStr(Frec.STYEAR(12)) & MDC_SetMod.sStr(Frec.SAVAMT(12)) & MDC_SetMod.sStr(Frec.SARAMT(12)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(13)) & MDC_SetMod.sStr(Frec.SAVCOD(13)) & MDC_SetMod.sStr(Frec.SAVNAM(13)) & MDC_SetMod.sStr(Frec.SAVNUM(13)) & MDC_SetMod.sStr(Frec.STYEAR(13)) & MDC_SetMod.sStr(Frec.SAVAMT(13)) & MDC_SetMod.sStr(Frec.SARAMT(13)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(14)) & MDC_SetMod.sStr(Frec.SAVCOD(14)) & MDC_SetMod.sStr(Frec.SAVNAM(14)) & MDC_SetMod.sStr(Frec.SAVNUM(14)) & MDC_SetMod.sStr(Frec.STYEAR(14)) & MDC_SetMod.sStr(Frec.SAVAMT(14)) & MDC_SetMod.sStr(Frec.SARAMT(14)) & _
                    MDC_SetMod.sStr(Frec.SAVGBN(15)) & MDC_SetMod.sStr(Frec.SAVCOD(15)) & MDC_SetMod.sStr(Frec.SAVNAM(15)) & MDC_SetMod.sStr(Frec.SAVNUM(15)) & MDC_SetMod.sStr(Frec.STYEAR(15)) & MDC_SetMod.sStr(Frec.SAVAMT(15)) & MDC_SetMod.sStr(Frec.SARAMT(15)) & _
                    MDC_SetMod.sStr(Frec.FILLER)
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " 연금저축명세 생성 완료.", True)
        End If
    End If

    '/ 기본부양가족공제의 금액과 소득공제항목등록한 총공제금액이 다를경우
    sQry = "SELECT  SUM(ISNULL(U_RETSAV,0)) AS RETSAV,"
    sQry = sQry & " SUM(ISNULL(U_RETSA1,0)) AS RETSA1,"
    sQry = sQry & " SUM(ISNULL(U_GYNSAV,0)) AS GYNSAV,"
    sQry = sQry & " SUM(ISNULL(U_YUNSAV,0)) AS YUNSAV,"
    sQry = sQry & " SUM(ISNULL(U_HUSAMT,0)) AS HUSAMT,"
    sQry = sQry & " SUM(ISNULL(U_HU1AMT,0)) AS HU1AMT,"
    sQry = sQry & " SUM(ISNULL(U_HU2AMT,0)) AS HU2AMT,"
    sQry = sQry & " SUM(ISNULL(U_HU3AMT,0)) AS HU3AMT,"
    sQry = sQry & " SUM(ISNULL(U_JFDAM1,0)) AS JFDAM1,"
    sQry = sQry & " SUM(ISNULL(U_JFDAM2,0)) AS JFDAM2,"
    sQry = sQry & " SUM(ISNULL(U_JFDAM3,0)) AS JFDAM3 "
    sQry = sQry & " FROM [@ZPY501H] "
    sQry = sQry & " WHERE U_JSNYER = '" & oJsnYear & "'"
    sQry = sQry & " AND U_MSTCOD = '" & C_MSTCOD & "'"
    sQry = sQry & " AND U_CLTCOD = '" & C_CLTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        '/1. 퇴직연금(근로자 퇴직급여보장법)
        If RETSAV < Val(oRecordSet.Fields("RETSAV").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "퇴직연금(근로자 퇴직급여보장법) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/2. 퇴직연금(과학기술인공제)
        If RETSA1 < Val(oRecordSet.Fields("RETSA1").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "퇴직연금(과학기술인공제) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/3. 연금저축(개인연금저축)
        If GYNSAV < Val(oRecordSet.Fields("GYNSAV").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "연금저축(개인연금저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/4. 연금저축(연금저축)
        If YUNSAV < Val(oRecordSet.Fields("YUNSAV").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "연금저축(연금저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/5. 주택마련저축(청약저축)
        If HUSAMT < Val(oRecordSet.Fields("HUSAMT").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "주택마련저축(청약저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/6. 주택마련저축(주택청약종합저축)
        If HU1AMT < Val(oRecordSet.Fields("HU1AMT").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "주택마련저축(주택청약종합저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/7. 주택마련저축(장기주택마련저축)
        If HU2AMT < Val(oRecordSet.Fields("HU2AMT").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "주택마련저축(장기주택마련저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/8. 주택마련저축(근로자주택마련저축)
        If HU3AMT < Val(oRecordSet.Fields("HU3AMT").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "주택마련저축(근로자주택마련저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/9. 주택마련저축(장기주식형-1년차)
        If JFDAM1 < Val(oRecordSet.Fields("JFDAM1").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "주택마련저축(장기주식형-1년차) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/10. 주택마련저축(장기주식형-2년차)
        If JFDAM2 < Val(oRecordSet.Fields("JFDAM2").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "주택마련저축(장기주식형-2년차) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        '/11. 주택마련저축(장기주식형-3년차)
        If JFDAM3 < Val(oRecordSet.Fields("JFDAM3").Value) Then
            Call Matrix_AddRow("F레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "주택마련저축(장기주식형-3년차) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", True, True)
        End If
        
    End If
    If CheckF = False Then
        File_Create_FRecord = True
    Else
        File_Create_FRecord = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    Call Matrix_AddRow("F레코드오류: " & Err.Source & " " & Err.Description, False)
    File_Create_FRecord = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim MstInfo         As ZPAY_g_EmpID
        
    Select Case oUID
      Case "JsnYear"
            oJsnYear = oForm.Items(oUID).Specific.String
            If Trim$(oJsnYear) = "" Then
                ZPAY_GBL_JSNYER = oJsnYear
            End If
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oForm.Items(oUID).Specific.String = ""
                oForm.DataSources.UserDataSources("MSTNAM").ValueEx = ""
                oForm.DataSources.UserDataSources("EmpID").ValueEx = ""
            Else
                oForm.Items(oUID).Specific.String = UCase$(oForm.Items(oUID).Specific.String)
                MstInfo = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oForm.DataSources.UserDataSources("MSTNAM").ValueEx = MstInfo.MSTNAM
                oForm.DataSources.UserDataSources("EmpID").ValueEx = MstInfo.EmpID
            End If
            oForm.Items("MSTNAM").Update
            oForm.Items("EmpID").Update
    End Select
    oForm.Items(oUID).Update
End Sub

Private Function HeaderSpaceLineDel() As Boolean
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'저장할 데이터의 유효성을 점검한다
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
On Error GoTo HeaderSpaceLineDel
    Dim ErrNum      As Integer
    
    ErrNum = 0
    '/ 필수Check
    If MDC_SetMod.ChkYearMonth(oForm.Items("JsnYear").Specific.String & "01") = False Then     '/ 정산년도
        ErrNum = 1
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("JSNTYP").Specific.Selected Is Nothing Then
        ErrNum = 2
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("BPLId").Specific.Selected Is Nothing Then
        ErrNum = 3
        GoTo HeaderSpaceLineDel
    ElseIf Trim$(oForm.Items("PRTDAT").Specific.String) = "" Then
        ErrNum = 4
        GoTo HeaderSpaceLineDel
    End If
    MSTBRK = oForm.Items("BPLId").Specific.Selected.Value
    DPTSTR = oForm.Items("DptStr").Specific.Selected.Value
    If DPTSTR = "-1" Then DPTSTR = "00000001"   '20120209 "00000001" 변경
    DPTEND = oForm.Items("DptEnd").Specific.Selected.Value
    If DPTEND = "-1" Then DPTEND = "ZZZZZZZZ"
    MSTCOD = oForm.Items("MSTCOD").Specific.Value
    If Trim$(MSTCOD) = "" Then MSTCOD = "%"
    
    HeaderSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
HeaderSpaceLineDel:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "귀속년도를 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "기간코드는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "지점은 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "제출일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    
    HeaderSpaceLineDel = False
End Function

Private Sub Matrix_AddRow(ByVal MatrixMsg As String, Optional Insert_YN As Boolean, Optional MatrixErr As Boolean)
    If MatrixErr = True Then
        oForm.DataSources.UserDataSources("Col0").Value = "??"
    Else
        oForm.DataSources.UserDataSources("Col0").Value = ""
    End If
    oForm.DataSources.UserDataSources("Col1").Value = MatrixMsg
    If Insert_YN = True Then
        oMat1.AddRow
        MaxRow = MaxRow + 1
    End If
    oMat1.SetLineData MaxRow
End Sub

