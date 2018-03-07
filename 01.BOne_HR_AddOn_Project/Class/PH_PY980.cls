VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY980"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : PH_PY980.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 근로소득지급명세서자료 전산매체수록
'//  FormType       :
'//  Create Date    : 2014.01.17
'//  Modified Date  : 2015.01.21
'//  Creator        : NGY
'//  Modifier       :
'//  Copyright  (c) Poongsan Holdings
'//****************************************************************************

Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private sRecordset      As SAPbobsCOM.Recordset
Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     '클래스에서 선택한 마지막 아이템 Uid값

Private CLTCOD          As String
Private yyyy            As String
Private HtaxID          As String
Private TeamName        As String
Private Dname           As String
Private Dtel            As String
Private DocDate         As String
Private oFilePath       As String

Private FILNAM      As String * 30    '파  일  명
Private MaxRow      As Long
Private BUSCNT      As Integer  '/ B레코드일련번호
Private BUSTOT      As Integer  '/ B레코드총갯수

Private NEWCNT      As Integer
Private OLDCNT      As Integer
Private C_SAUP      As String
Private C_YYYY      As String
Private C_SABUN     As String
Private E_BUYCNT    As String
Private C_BUYCNT    As String

'2013년기준 1400 BYTE
'2014년기준 1520 BYTE
'2014년기준 1580 BYTE  re
'2015년귀속 1610 BYTE
'2016년귀속 1620 BYTE
'2017년귀속 1620 BYTE

'/ A 제출자 레코드
Private Type A_record
        A001   As String * 1     '레코드구분
        A002   As String * 2     '자료구분
        A003   As String * 3     '세무서
        A004   As String * 8     '제출일자
        A005   As String * 1     '제출자구분 (1;세무대리인, 2;법인, 3;개인)
        A006   As String * 6     '세무대리인
        A007   As String * 20    '홈텍스ID
        A008   As String * 4     '세무프로그램코드
        A009   As String * 10    '사업자번호
        A010   As String * 40    '법인명(상호)
        A011   As String * 30    '담당자부서
        A012   As String * 30    '담당자성명
        A013   As String * 15    '담당자전화번호
        A014   As String * 5     '신고의무자수
        A015   As String * 3     '한글코드종류
        A016   As String * 1442  '공란
End Type
Dim A_rec As A_record

'/ B 원천징수의무자별 집계 레코드
Private Type B_record
        B001   As String * 1     '레코드구분
        B002   As String * 2     '자료구분
        B003   As String * 3     '세무서
        B004   As String * 6     '일련번호
        B005   As String * 10    '사업자번호
        B006   As String * 40    '법인명(상호)
        B007   As String * 30    '대표자
        B008   As String * 13    '주민(법인)번호
        B009   As String * 7     '주(현)근무처(C레코드)수
        B010   As String * 7     '종(전)근무처(D레코드)수
        B011   As String * 14    '총급여총계
        B012   As String * 13    '결정세액(소득세)총계
        B013   As String * 13    '결정세액(지방소득세)총계
        B014   As String * 13    '결정세액(농특세)총계
        B015   As String * 13    '결정세액총계
        B016   As String * 1     '제출대상기간
        B017   As String * 1434  '공란
        
End Type
Dim B_rec As B_record

'/ C 주(현)근무지 레코드
Private Type C_record
        C001   As String * 1     '레코드구분
        C002   As String * 2     '자료구분
        C003   As String * 3     '세무서
        C004   As String * 6     '일련번호
        C005   As String * 10    '사업자번호
        C006   As String * 2     '종(전)근무처수
        C007   As String * 1     '거주자구분코드
        C008   As String * 2     '거주지국코드
        C009   As String * 1     '외국인단일세율적용
        C009A  As String * 1     '외국법인소속파견근로자여부 1,여 2,부
        C010   As String * 30    '성명
        C011   As String * 1     '내.외국인구분
        C012   As String * 13    '주민등록번호
        C013   As String * 2     '국적코드
        C014   As String * 1     '세대주여부
        C015   As String * 1     '연말정산구분
        '/근무처별소득명세_주(현)근무처
        C016   As String * 10    '주현근무처-사업자번호
        C017   As String * 40    '주현근무처-근무처명
        C018   As String * 8     '근무기간 시작연월일
        C019   As String * 8     '근무기간 종료연월일
        C020   As String * 8     '감면기간 시작연월일
        C021   As String * 8     '감면기간 종료연월일
        C022   As String * 11    '급여총액
        C023   As String * 11    '상여총액
        C024   As String * 11    '인정상여
        C025   As String * 11    '주식매수선택권행사이익
        C026   As String * 11    '우리사주조합인출금
        C027   As String * 11    '임원퇴직소득금액한도초과액
        C028   As String * 21    '공란
        C029   As String * 11    '계
        '/주(현)근무처 비과세 및 감면소득
        C030   As String * 10    '비과세(G01:학자금)
        C031   As String * 10    '비과세(H01:무보수위원수당)
        C032   As String * 10    '비과세(H05:경호,승선수당)
        C033   As String * 10    '비과세(H06:유아,초중등)
        C034   As String * 10    '비과세(H07:고등교육법)
        C035   As String * 10    '비과세(H08:특별법)
        C036   As String * 10    '비과세(H09:연구기관등)
        C037   As String * 10    '비과세(H10:기업부설연구소)
        C038   As String * 10    '비과세(H14:보육교사근무환경개선비)
        C039   As String * 10    '비과세(H15:사립유치원수석교사.교사의인건비)
        C040   As String * 10    '비과세(H11:취재수당)
        C041   As String * 10    '비과세(H12:벽지수당)
        C042   As String * 10    '비과세(H13:재해관련급여)
        C043   As String * 10    '비과세(H16:정부공공기관지방이전기관종사자이주수당)
        C044   As String * 10    '비과세(I01:외국정부등근로자)
        C045   As String * 10    '비과세(K01:외국주둔군인등)
        C046   As String * 10    '비과세(M01:국외근로100만원)
        C047   As String * 10    '비과세(M02:국외근로300만원)
        C048   As String * 10    '비과세(M03:국외근로)
        C049   As String * 10    '비과세(O01:야간근로수당)
        C050   As String * 10    '비과세(Q01:출산보육수당)
        C051   As String * 10    '비과세(R10:근로장학금)
        C052   As String * 10    '비과세(R11:직무발명보상금)
        C053   As String * 10    '비과세(S01:주식매수선택권)
        C054   As String * 10    '비과세(T01:외국인기술자)
        C055   As String * 10    '비과세(Y02:우리사주조합인출금50%)
        C056   As String * 10    '비과세(Y03:우리사주조합인출금75%)
        C056A  As String * 10    '비과세(Y03:우리사주조합인출금100%)
        C057   As String * 10    '비과세(Y22:전공의수련보조수당)
        C058   As String * 10    '비과세(T10:중소기업취업청년소득세감면100%)
        C058A  As String * 10    '비과세(T10:중소기업취업청년소득세감면50%)
        C058B  As String * 10    '비과세(T10:중소기업취업청년소득세감면70%)
        C059   As String * 10    '비과세(T20:조세조약상교직자감면)
        C060   As String * 10    '비과세 계
        C061   As String * 10    '감면소득 계
       'C062   As String * 10    '2016삭제
        '/정산명세
        C063   As String * 11    '총급여
        C064   As String * 10    '근로소득공제
        C065   As String * 11    '근로소득금액
        '/기본공제
        C066   As String * 8     '본인공제금액
        C067   As String * 8     '배우자공제금액
        C068   As String * 2     '부양가족공제인원
        C069   As String * 8     '부양가족공제금액
        '/추가공제
        C070   As String * 2     '경로우대공제인원
        C071   As String * 8     '경로우대공제금액
        C072   As String * 2     '장애자공제인원
        C073   As String * 8     '장애자공제금액
        C074   As String * 8     '부녀자공제금액
        C075   As String * 10    '한부모공제금액
        '/연금보험료공제
        C076   As String * 10    '국민연금보험료공제
        C077   As String * 10    '공적연금보험료공제(공무원연금)
        C078   As String * 10    '공적연금보험료공제(군인연금)
        C079   As String * 10    '공적연금보험료공제(사립학교교직원연금)
        C080   As String * 10    '공적연금보험료공제(별정우체국연금)
        '/특별소득공제
        C081   As String * 10    '보험료_건강보험료
        C082   As String * 10    '보험료_고용보험료
        C083   As String * 8     '주택자금_주택임대차차입금 원리금상환공제금액-대출기관
        C084   As String * 8     '주택자금_주택임대차차입금 원리금상환공제금액-거주자
        C085   As String * 8     '2011 장기주택저당차입금이자상환공제금액-15년미만
        C086   As String * 8     '2011 장기주택저당차입금이자상환공제금액-29년미만
        C087   As String * 8     '2011 장기주택저당차입금이자상환공제금액-30년이상
        C088   As String * 8     '2012 이후차입분,15년이상-고정금리비거치상환대출
        C089   As String * 8     '2012 이후차입분,15년이상-기타대출
        C090   As String * 8     '2015 이후차입분,15년이상-고정금리이면서비거치상환대출
        C091   As String * 8     '2015 이후차입분,15년이상-고정금리이거나비거치상환대출
        C092   As String * 8     '2015 이후차입분,15년이상-기타대출
        C093   As String * 8     '2015 이후차입분,10~15년-고정금리이거나비거치상환대출
        
        C094   As String * 11    '기부금(이월분)
       'C095   As String * 20    '공란  2016삭제
        C096   As String * 11    '특별소득공제계
        C097   As String * 11    '차감소득금액
        '/그밖의소득공제
        C098   As String * 8     '개인연금저축소득공제
        C099   As String * 10    '소기업공제부금소득공제
        C100   As String * 10    '주택마련저축소득공제_청약저축
        C101   As String * 10    '주택마련저축소득공제_주택청약종합저축
        C102   As String * 10    '주택마련저축소득공제_근로자주택마련저축
        C103   As String * 10    '투자조합출자등소득공제
        C104   As String * 8     '신용카드등소득공제
        C105   As String * 10    '우리사주조합출연금
       'C106   As String * 11    '우리사주조합기부금(2014년이전기부금)'0'  2016삭제
        C107   As String * 10    '고용유지중소기업근로자소득공제
       'C108   As String * 10    '목돈안드는전세이자상환액공제 2017삭제
        C109   As String * 10    '장기집합투자증권저축
        C110   As String * 11    '그밖의소득공제계
        C111   As String * 11    '소득공제종합한도초과액
        C112   As String * 11    '종합소득과세표준
        C113   As String * 10    '산출세액
        '/세액감면
        C114   As String * 10    '소득세법
        C115   As String * 10    '조특법
        C116   As String * 10    '조특법제30조
        C117   As String * 10    '조세조약
       'C118   As String * 10    '공란  2016삭제
        C119   As String * 10    '세액감면계
        '/세액공제
        C120   As String * 10    '근로소득세액공제
        C121   As String * 2     '자녀세액공제인원
        C122   As String * 10    '자녀세액공제
        C123   As String * 2     '6세이하세액공제인원
        C124   As String * 10    '6세이하세액공제
        C125   As String * 2     '출산.입양세액공제인원
        C126   As String * 10    '출산.입양세액공제
        C127   As String * 10    '연금계좌_과학기술인공제_공제대상금액
        C128   As String * 10    '연금계좌_과학기술인공제_세액공제액
        C129   As String * 10    '연금계좌_근로자퇴직급여보장법에따른 퇴직급여_공제대상금액
        C130   As String * 10    '연금계좌_근로자퇴직급여보장법에따른 퇴직급여_세액공제액
        C131   As String * 10    '연금계좌_연금저축_공제대상금액
        C132   As String * 10    '연금계좌_연금저축_세액공제액
        C133   As String * 10    '특별세액공제_보장성보험료_공제대상금액
        C134   As String * 10    '특별세액공제_보장성보험료_세액공제액
        C135   As String * 10    '특별세액공제_장애인전용보험료_공제대상금액
        C136   As String * 10    '특별세액공제_장애인전용보험료_세액공제액
        C137   As String * 10    '특별세액공제_의료비_공제대상금액
        C138   As String * 10    '특별세액공제_의료비_세액공제액
        C139   As String * 10    '특별세액공제_교육비_공제대상금액
        C140   As String * 10    '특별세액공제_교육비_세액공제액
        C141   As String * 10    '특별세액공제_기부금_정치자금_10만원이하_공제대상금액
        C142   As String * 10    '특별세액공제_기부금_정치자금_10만원이하_세액공제액
        C143   As String * 11    '특별세액공제_기부금_정치자금_10만원초과_공제대상금액
        C144   As String * 10    '특별세액공제_기부금_정치자금_10만원초과_세액공제액
        C145   As String * 11    '특별세액공제_기부금_법정기부금_공제대상금액
        C146   As String * 10    '특별세액공제_기부금_법정기부금_세액공제액
        C147   As String * 11    '특별세액공제_기부금_우리사주조합기부금_공제대상금액
        C148   As String * 10    '특별세액공제_기부금_우리사주조합기부금_세액공제액
        C149   As String * 11    '특별세액공제_기부금_지정기부금_공제대상금액(종교단체외)
        C150   As String * 10    '특별세액공제_기부금_지정기부금_세액공제액(종교단체외)
        C150A  As String * 11    '특별세액공제_기부금_지정기부금_공제대상금액(종교단체)
        C150B  As String * 10    '특별세액공제_기부금_지정기부금_세액공제액(종교단체)
        C151   As String * 10    '특별세액공제계
        C152   As String * 10    '표준세액공제
        C153   As String * 10    '납세조합공제
        C154   As String * 10    '주택차입금
        C155   As String * 10    '외국납부
        C156   As String * 10    '월세세액공제_공제대상금액
        C157   As String * 8     '월세세액공제_세액공제액
        C158   As String * 10    '세액공제계
        '/결정세액
        C159   As String * 10    '소득세
        C160   As String * 10    '지방소득세
        C161   As String * 10    '농특세
        '/기납부세액_주(현)근무지
        C162   As String * 10    '소득세
        C163   As String * 10    '지방소득세
        C164   As String * 10    '농특세
        '/납부특례세액
        C165   As String * 10    '소득세
        C166   As String * 10    '지방소득세
        C167   As String * 10    '농특세
        '/차감징수세액
        C168_1 As String * 1     '소득세(기호 음수1, 양수0)
        C168_2 As String * 10    '소득세
        C169_1 As String * 1     '지방소득세(기호 음수1, 양수0)
        C169_2 As String * 10    '지방소득세
        C170_1 As String * 1     '농특세(기호 음수1, 양수0)
        C170_2 As String * 10    '농특세
        C171   As String * 1     '사업자단위과세자여부 '2'부
        C172   As String * 4     '종사업장일련번호 '공란'
        C173   As String * 20    '공란
End Type
Dim C_rec As C_record

'/ D 종(전)근무지 레코드
Private Type D_Record
        D001   As String * 1    '레코드구분
        D002   As String * 2    '자료구분
        D003   As String * 3    '세무서
        D004   As String * 6    '일련번호
        D005   As String * 10   '사업자등록번호
        D006   As String * 50   '공란
        D007   As String * 13   '소득자주민번호
        D008   As String * 1    '납세조합구분
        D009   As String * 40   '법인명(상호)
        D010   As String * 10   '사업자등록번호
        D011   As String * 8    '근무기간시작연월일
        D012   As String * 8    '근무기간종료연월일
        D013   As String * 8    '감면기간시작
        D014   As String * 8    '감면기간종료
        D015   As String * 11   '급여총액
        D016   As String * 11   '상여총액
        D017   As String * 11   '인정상여
        D018   As String * 11   '주식매수선택권행사이익
        D019   As String * 11   '우리사주조합인출금
        D020   As String * 11   '임원퇴직소득금액한도초과액
        D021   As String * 22   '공란
        D022   As String * 11   '계
        D023   As String * 10   '비과세(G01:학자금)
        D024   As String * 10   '비과세(H01:무보수위원수당)
        D025   As String * 10   '비과세(H05:경호,승선수당)
        D026   As String * 10   '비과세(H06:유아,초중등)
        D027   As String * 10   '비과세(H07:고등교육법)
        D028   As String * 10   '비과세(H08:특별법)
        D029   As String * 10   '비과세(H09:연구기관등)
        D030   As String * 10   '비과세(H10:기업부설연구소)
        D031   As String * 10   '비과세(H14:보육교사근무환경개선비)
        D032   As String * 10   '비과세(H15:사립유치원수석교사.교사의인건비)
        D033   As String * 10   '비과세(H11:취재수당)
        D034   As String * 10   '비과세(H12:벽지수당)
        D035   As String * 10   '비과세(H13:재해관련급여)
        D036   As String * 10   '비과세(H16:정부공공기관지방이전기관종사자이주수당)
        D037   As String * 10   '비과세(I01:외국정부등근로자)
        D038   As String * 10   '비과세(K01:외국주둔군인등)
        D039   As String * 10   '비과세(M01:국외근로100만원)
        D040   As String * 10   '비과세(M02:국외근로300만원)
        D041   As String * 10   '비과세(M03:국외근로)
        D042   As String * 10   '비과세(O01:야간근로수당)
        D043   As String * 10   '비과세(Q01:출산보육수당)
        D044   As String * 10   '비과세(R10:근로장학금)
        D045   As String * 10   '비과세(R11:직무발명보상금)
        D046   As String * 10   '비과세(S01:주식매수선택권)
        D047   As String * 10   '비과세(T01:외국인기술자)
        D048   As String * 10   '비과세(Y02:우리사주조합인출금50%)
        D049   As String * 10   '비과세(Y03:우리사주조합인출금75%)
        D049A  As String * 10   '비과세(Y03:우리사주조합인출금100%)
        D050   As String * 10   '비과세(Y22:전공의수련보조수당)
        D051   As String * 10   '비과세(T10:중소기업취업청년소득세감면100%)
        D051A  As String * 10   '비과세(T10:중소기업취업청년소득세감면50%)
        D051B  As String * 10   '비과세(T10:중소기업취업청년소득세감면70%)
        D052   As String * 10   '비과세(T20:조세조약상교직자감면)
        D053   As String * 10   '비과세 계
        D054   As String * 10   '감면소득 계
        D055   As String * 10   '공란 '0'기재
        D056   As String * 10   '소득세
        D057   As String * 10   '지방소득세
        D058   As String * 10   '농특세
        D059   As String * 2    '종(전)근무처일련번호
        D060   As String * 961  '공란
End Type
Dim D_rec As D_Record

'/ E 소득공제명세 레코드
Private Type E_Record
        E001   As String * 1    '레코드구분
        E002   As String * 2    '자료구분
        E003   As String * 3    '세무서
        E004   As String * 6    '일련번호
        E005   As String * 10   '사업자등록번호
        E006   As String * 13   '주민등록번호
        
        ' ARRY 5   E007 - E146
        E007(1 To 5) As String * 1    '관계코드
        E008(1 To 5) As String * 1    '내외국인구분
        E009(1 To 5) As String * 20   '성명
        E010(1 To 5) As String * 13   '주민등록번호
        E011(1 To 5) As String * 1    '기본공제
        E012(1 To 5) As String * 1    '장애자공제
        E013(1 To 5) As String * 1    '부녀자공제
        E014(1 To 5) As String * 1    '경로우대
        E015(1 To 5) As String * 1    '한부모공제
        E016(1 To 5) As String * 1    '출산.입양공제
        E017(1 To 5) As String * 1    '6세이하공제
        E018(1 To 5) As String * 1    '교육비공제 1,2,3,4,5
        E019(1 To 5) As String * 10   '국세청-보험료_건강,고용
        E020(1 To 5) As String * 10   '국세청-보험료_보장성
        E021(1 To 5) As String * 10   '국세청-보험료_장애인전용보장성
        E022(1 To 5) As String * 10   '국세청-의료비
        E023(1 To 5) As String * 10   '국세청-교육비
        E024(1 To 5) As String * 10   '국세청-신용카드
        E025(1 To 5) As String * 10   '국세청-직불카드
        E026(1 To 5) As String * 10   '국세청-현금영수증
        E027(1 To 5) As String * 10   '국세청-전통시장사용액
        E028(1 To 5) As String * 10   '국세청-대중교통이용액
        E029(1 To 5) As String * 13   '국세청-기부금
        E030(1 To 5) As String * 10   '기타-보험료_건강,고용
        E031(1 To 5) As String * 10   '기타-보험료_보장성
        E032(1 To 5) As String * 10   '기타-보험료_장애인전용보장성
        E033(1 To 5) As String * 10   '기타-의료비
        E034(1 To 5) As String * 10   '기타-교육비
        E035(1 To 5) As String * 10   '기타-신용카드
        E036(1 To 5) As String * 10   '기타-직불카드
        E037(1 To 5) As String * 10   '기타-전통시장사용액
        E038(1 To 5) As String * 10   '기타-대중교통이용액
        E039(1 To 5) As String * 13   '기부금
        
        E172   As String * 2          '부양가족레코드일련번호
        E173   As String * 288        '공란
End Type
Dim E_rec As E_Record

'/ F 연금.저축 등 소득.세액 공제명세 레코드
Private Type F_Record
        F001   As String * 1    '레코드구분
        F002   As String * 2    '자료구분
        F003   As String * 3    '세무서
        F004   As String * 6    '일련번호
        F005   As String * 10   '사업자번호
        F006   As String * 13   '주민번호
        ' ARRY 15   F007 - F096
        F007(1 To 15) As String * 2   '소득공제구분
        F008(1 To 15) As String * 3   '금융기관코드
        F009(1 To 15) As String * 50  '금융기관상호
        F010(1 To 15) As String * 20  '계좌번호
        F011(1 To 15) As String * 10  '납입금액
        F012(1 To 15) As String * 10  '소득세액공제금액
        
        F097   As String * 160  '공란
End Type
Dim F_rec As F_Record

'/ G 월세.거주자간 주택임차차임금 원리금 상환액 소득공제명세 레코드
Private Type G_Record
        G001   As String * 1    '레코드구분
        G002   As String * 2    '자료구분
        G003   As String * 3    '세무서
        G004   As String * 6    '일련번호
        G005   As String * 10   '사업자번호
        G006   As String * 13   '주민번호
        '/1
        G007   As String * 20   '임대인성명1
        G008   As String * 13   '주민등록번호
        G009   As String * 1    '주택유형
        G010   As String * 5    '주택계약면적
        G011   As String * 100  '임대차계약서상주소지
        G012   As String * 8    '임대차계약기간시작
        G013   As String * 8    '임대차계약기간종료
        G014   As String * 10   '연간월세액
        G015   As String * 10   '세액공제금액
        G016   As String * 20   '대주성명
        G017   As String * 13   '대주주민등록번호
        G018   As String * 8    '금전소비대차 계약기간시작
        G019   As String * 8    '금전소비대차 계약기간종료
        G020   As String * 4    '차입금이자율
        G021   As String * 10   '원리금산환액계
        G022   As String * 10   '원금
        G023   As String * 10   '이자
        G024   As String * 10   '공제금액
        G025   As String * 20   '임대인성명
        G026   As String * 13   '주민등록번호
        G027   As String * 1    '주택유형
        G028   As String * 5    '주택계약면적
        G029   As String * 100  '임대차계약서상주소지
        G030   As String * 8    '임대차계약기간시작
        G031   As String * 8    '임대차계약기간종료
        G032   As String * 10   '전세보증금
        '/2
        G033   As String * 20   '임대인성명2
        G034   As String * 13   '주민등록번호
        G035   As String * 1    '주택유형
        G036   As String * 5    '주택계약면적
        G037   As String * 100  '임대차계약서상주소지
        G038   As String * 8    '임대차계약기간시작
        G039   As String * 8    '임대차계약기간종료
        G040   As String * 10   '연간월세액
        G041   As String * 10   '세액공제금액
        G042   As String * 20   '대주성명
        G043   As String * 13   '대주주민등록번호
        G044   As String * 8    '금전소비대차 계약기간시작
        G045   As String * 8    '금전소비대차 계약기간종료
        G046   As String * 4    '차입금이자율
        G047   As String * 10   '원리금산환액계
        G048   As String * 10   '원금
        G049   As String * 10   '이자
        G050   As String * 10   '공제금액
        G051   As String * 20   '임대인성명
        G052   As String * 13   '주민등록번호
        G053   As String * 1    '주택유형
        G054   As String * 5    '주택계약면적
        G055   As String * 100  '임대차계약서상주소지
        G056   As String * 8    '임대차계약기간시작
        G057   As String * 8    '임대차계약기간종료
        G058   As String * 10   '전세보증금
        '/3
        G059   As String * 20   '임대인성명3
        G060   As String * 13   '주민등록번호
        G061   As String * 1    '주택유형
        G062   As String * 5    '주택계약면적
        G063   As String * 100  '임대차계약서상주소지
        G064   As String * 8    '임대차계약기간시작
        G065   As String * 8    '임대차계약기간종료
        G066   As String * 10   '연간월세액
        G067   As String * 10   '세액공제금액
        G068   As String * 20   '대주성명
        G069   As String * 13   '대주주민등록번호
        G070   As String * 8    '금전소비대차 계약기간시작
        G071   As String * 8    '금전소비대차 계약기간종료
        G072   As String * 4    '차입금이자율
        G073   As String * 10   '원리금산환액계
        G074   As String * 10   '원금
        G075   As String * 10   '이자
        G076   As String * 10   '공제금액
        G077   As String * 20   '임대인성명
        G078   As String * 13   '주민등록번호
        G079   As String * 1    '주택유형
        G080   As String * 5    '주택계약면적
        G081   As String * 100  '임대차계약서상주소지
        G082   As String * 8    '임대차계약기간시작
        G083   As String * 8    '임대차계약기간종료
        G084   As String * 10   '전세보증금
        
        G085   As String * 2    '일련번호
        G086   As String * 284  '공란
End Type
Dim G_rec As G_Record

'/ H 기부조정명세 레코드
Private Type H_record
        H001   As String * 1     '레코드구분
        H002   As String * 2     '자료구분
        H003   As String * 3     '세무서
        H004   As String * 6     '일련번호
        H005   As String * 10    '사업자번호
        H006   As String * 13    '소득자주민등록번호
        H007   As String * 1     '내,외국인코드
        H008   As String * 30    '성명
        H009   As String * 2     '유형코드
        H010   As String * 4     '기부년도
        H011   As String * 13    '기부금액
        H012   As String * 13    '전년까지공제된금액
        H013   As String * 13    '공제대상금액
        H014   As String * 13    '해당년도공제금액 필요경비 '0'  2016
        H015   As String * 13    '해당년도공제금액세액(소득)공제금액  2016
        H016   As String * 13    '해당년도에공제받지못한금액_소멸금액
        H017   As String * 13    '해당년도에공제받지못한금액_이월금액
        H018   As String * 5     '기부조정명세일련번호
        H019   As String * 1452  '공란
End Type
Dim H_rec As H_record

'/ I 해당년도 기부명세 레코드
Private Type I_Record
        I001   As String * 1    '레코드구분
        I002   As String * 2    '자료구분
        I003   As String * 3    '세무서
        I004   As String * 6    '일련번호
        I005   As String * 10   '사업자등록번호
        I006   As String * 13   '주민등록번호
        I007   As String * 2    '유형코드
        I008   As String * 13   '기부처-사업자등록번호
        I009   As String * 30   '기부처-법인명(상호)
        I010   As String * 1    '관계
        I011   As String * 1    '내,외국인코드
        I012   As String * 20   '성명
        I013   As String * 13   '주민등록번호
        I014   As String * 5    '건수
        I015   As String * 13   '기부금합계금액
        I016   As String * 13   '공제대상기부금액
        I017   As String * 13   '기부장려금신청금액
        I018   As String * 13   '기타
        I019   As String * 5    '해당연도기부명세일련번호
        I020   As String * 1443 '공란
End Type
Dim I_rec As I_Record


'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY980.srf")
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
    
    oFormUniqueID = "PH_PY980_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "PH_PY980"
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
'                If pval.ItemUID = "CBtn1" Then   '/ ChooseBtn사원리스트
'                    oForm.Items("MSTCOD").CLICK ct_Regular
'                    Sbo_Application.ActivateMenuItem ("7425")
'                    BubbleEvent = False
'                Else

                 If pval.ItemUID = "Btn01" Then
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
                
                 End If
            Else
            End If
            
        Case et_COMBO_SELECT:
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID
                           '//사업장이 바뀌면
                           Case "CLTCOD"
                                sQry = "SELECT U_HomeTId, U_ChgDpt, U_ChgName, U_ChgTel  FROM [@PH_PY005A] WHERE U_CLTCode = '" & Trim(oForm.Items("CLTCOD").Specific.VALUE) & "'"
                                oRecordSet.DoQuery sQry
                                oForm.Items("HtaxID").Specific.String = Trim(oRecordSet.Fields("U_HomeTId").VALUE)
                                oForm.Items("TeamName").Specific.String = Trim(oRecordSet.Fields("U_ChgDpt").VALUE)
                                oForm.Items("Dname").Specific.String = Trim(oRecordSet.Fields("U_ChgName").VALUE)
                                oForm.Items("Dtel").Specific.String = Trim(oRecordSet.Fields("U_ChgTel").VALUE)
                           
                    End Select
                End If
            End If
            
            
            
            
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
    
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
    
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
        
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
    
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
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    oForm.Items("CLTCOD").DisplayDesc = True
    '// 접속자에 따른 권한별 사업장 콤보박스세팅
    Call CLTCOD_Select(oForm, "CLTCOD")
    
    oForm.Items("YYYY").Specific.String = Format(Now, "YYYY") - 1    '년도 기본년도에서 - 1
    
    Call oForm.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)     '제출일자
    oForm.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Function File_Create() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oStr            As String
    Dim sQry            As String
    
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    '화면변수를 전역변수로 MOVE
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = Trim(oForm.Items("YYYY").Specific.VALUE)
    HtaxID = Trim(oForm.Items("HtaxID").Specific.VALUE)
    TeamName = Trim(oForm.Items("TeamName").Specific.VALUE)
    Dname = Trim(oForm.Items("Dname").Specific.VALUE)
    Dtel = Trim(oForm.Items("Dtel").Specific.VALUE)
    DocDate = Trim(oForm.Items("DocDate").Specific.VALUE)
    
    ErrNum = 0
    
    '/ Question
    If Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
'    Sbo_Application.StatusBar.SetText "전산매체수록중..............", bmt_Short, smt_Success
    
    '/ A RECORD 처리
    If File_Create_A_record = False Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    '/ B RECORD 처리
    If File_Create_B_record = False Then
        ErrNum = 3
        GoTo Error_Message
    End If
    
    '/ C RECORD 처리  D.E.F.G 처리
    If File_Create_C_record = False Then
        ErrNum = 4
        GoTo Error_Message
    End If
    
    Close #1
    
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
    Else
        Sbo_Application.StatusBar.SetText "File_Create 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    File_Create = False
End Function
Private Function File_Create_A_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim PRTDAT          As String
    Dim saup            As String
    Dim CheckA          As String
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CheckA = False  '/체크필요유무
    ErrNum = 0
    
    '/ A_RECORE QUERY
    sQry = "EXEC PH_PY980_A '" & CLTCOD & "', '" & HtaxID & "', '" & TeamName & "', '" & Dname & "', '" & Dtel & "', '" & DocDate & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        ' PATH및 파일이름 만들기
        saup = oRecordSet.Fields("A009").VALUE  '사업자번호
        oFilePath = "C:\BANK\C" + Mid$(saup, 1, 7) + "." + Mid$(saup, 8, 3)
        
        
        'A RECORD MOVE
        
        A_rec.A001 = oRecordSet.Fields("A001").VALUE
        A_rec.A002 = oRecordSet.Fields("A002").VALUE
        A_rec.A003 = oRecordSet.Fields("A003").VALUE
        A_rec.A004 = oRecordSet.Fields("A004").VALUE
        A_rec.A005 = oRecordSet.Fields("A005").VALUE
        A_rec.A006 = oRecordSet.Fields("A006").VALUE
        A_rec.A007 = oRecordSet.Fields("A007").VALUE
        A_rec.A008 = oRecordSet.Fields("A008").VALUE
        A_rec.A009 = oRecordSet.Fields("A009").VALUE
        A_rec.A010 = oRecordSet.Fields("A010").VALUE
        A_rec.A011 = oRecordSet.Fields("A011").VALUE
        A_rec.A012 = oRecordSet.Fields("A012").VALUE
        A_rec.A013 = oRecordSet.Fields("A013").VALUE
        
        A_rec.A014 = Format$(oRecordSet.Fields("A014").VALUE, String$(Len(A_rec.A014), "0"))
        A_rec.A015 = oRecordSet.Fields("A015").VALUE
        A_rec.A016 = oRecordSet.Fields("A016").VALUE
        
        Close #1
        Open oFilePath For Output As #1
        Print #1, MDC_SetMod.sStr(A_rec.A001) & MDC_SetMod.sStr(A_rec.A002) & MDC_SetMod.sStr(A_rec.A003) & MDC_SetMod.sStr(A_rec.A004) & _
                  MDC_SetMod.sStr(A_rec.A005) & MDC_SetMod.sStr(A_rec.A006) & MDC_SetMod.sStr(A_rec.A007) & MDC_SetMod.sStr(A_rec.A008) & _
                  MDC_SetMod.sStr(A_rec.A009) & MDC_SetMod.sStr(A_rec.A010) & MDC_SetMod.sStr(A_rec.A011) & MDC_SetMod.sStr(A_rec.A012) & _
                  MDC_SetMod.sStr(A_rec.A013) & MDC_SetMod.sStr(A_rec.A014) & MDC_SetMod.sStr(A_rec.A015) & MDC_SetMod.sStr(A_rec.A016)
                  
    End If
    
    If CheckA = False Then
        File_Create_A_record = True
    Else
        File_Create_A_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "귀속년도의 자사정보(A RECORD)가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("A레코드오류: " & Err.Description, False, True)
    End If
    
    File_Create_A_record = False
    
End Function

Private Function File_Create_B_record() As Integer
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim CheckB          As String
    
    CheckB = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ B_RECORE QUERY
    sQry = "EXEC PH_PY980_B '" & CLTCOD & "', '" & yyyy & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        'B RECORD MOVE
        
        B_rec.B001 = oRecordSet.Fields("B001").VALUE
        B_rec.B002 = oRecordSet.Fields("B002").VALUE
        B_rec.B003 = oRecordSet.Fields("B003").VALUE
        B_rec.B004 = oRecordSet.Fields("B004").VALUE
        B_rec.B005 = oRecordSet.Fields("B005").VALUE
        B_rec.B006 = oRecordSet.Fields("B006").VALUE
        B_rec.B007 = oRecordSet.Fields("B007").VALUE
        B_rec.B008 = oRecordSet.Fields("B008").VALUE
        B_rec.B009 = Format$(oRecordSet.Fields("B009").VALUE, String$(Len(B_rec.B009), "0"))
        B_rec.B010 = Format$(oRecordSet.Fields("B010").VALUE, String$(Len(B_rec.B010), "0"))
        B_rec.B011 = Format$(oRecordSet.Fields("B011").VALUE, String$(Len(B_rec.B011), "0"))
        B_rec.B012 = Format$(oRecordSet.Fields("B012").VALUE, String$(Len(B_rec.B012), "0"))
        B_rec.B013 = Format$(oRecordSet.Fields("B013").VALUE, String$(Len(B_rec.B013), "0"))
        B_rec.B014 = Format$(oRecordSet.Fields("B014").VALUE, String$(Len(B_rec.B014), "0"))
        B_rec.B015 = Format$(oRecordSet.Fields("B015").VALUE, String$(Len(B_rec.B015), "0"))
        B_rec.B016 = oRecordSet.Fields("B016").VALUE
        B_rec.B017 = oRecordSet.Fields("B017").VALUE
        
'        Close #1
'        Open oFilePath For Output As #1
        Print #1, MDC_SetMod.sStr(B_rec.B001) & MDC_SetMod.sStr(B_rec.B002) & MDC_SetMod.sStr(B_rec.B003) & MDC_SetMod.sStr(B_rec.B004) & _
                  MDC_SetMod.sStr(B_rec.B005) & MDC_SetMod.sStr(B_rec.B006) & MDC_SetMod.sStr(B_rec.B007) & MDC_SetMod.sStr(B_rec.B008) & _
                  MDC_SetMod.sStr(B_rec.B009) & MDC_SetMod.sStr(B_rec.B010) & MDC_SetMod.sStr(B_rec.B011) & MDC_SetMod.sStr(B_rec.B012) & _
                  MDC_SetMod.sStr(B_rec.B013) & MDC_SetMod.sStr(B_rec.B014) & MDC_SetMod.sStr(B_rec.B015) & MDC_SetMod.sStr(B_rec.B016) & _
                  MDC_SetMod.sStr(B_rec.B017)
    End If
    
    If CheckB = False Then
        File_Create_B_record = True
    Else
        File_Create_B_record = False
    End If
    
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "B레코드가 존재하지 않습니다.", bmt_Short, smt_Error
        File_Create_B_record = 1
    Else
        Call Matrix_AddRow("B레코드오류: " & Err.Description, False)
        File_Create_B_record = 1
    End If
    
End Function

Private Function File_Create_C_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckC      As String
    Dim OLDBIG      As Double
    Dim PILTOT      As Double
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    CheckC = False  '/체크필요유무
    ErrNum = 0
    
    '/ C_RECORE QUERY
    sQry = "EXEC PH_PY980_C '" & CLTCOD & "', '" & yyyy & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("작성시작!", oRecordSet.RecordCount, False)
    
    NEWCNT = 0
    Do Until oRecordSet.EOF
    
        NEWCNT = NEWCNT + 1     '/ 일련번호
        
        C_SAUP = oRecordSet.Fields("saup").VALUE
        C_YYYY = oRecordSet.Fields("yyyy").VALUE
        C_SABUN = oRecordSet.Fields("sabun").VALUE
        
        'C RECORD MOVE
        
        C_rec.C001 = oRecordSet.Fields("C001").VALUE
        C_rec.C002 = oRecordSet.Fields("C002").VALUE
        C_rec.C003 = oRecordSet.Fields("C003").VALUE
        C_rec.C004 = Format$(NEWCNT, String$(Len(C_rec.C004), "0"))  '/ 일련번호
        C_rec.C005 = oRecordSet.Fields("C005").VALUE
        C_rec.C006 = Format$(oRecordSet.Fields("C006").VALUE, String$(Len(C_rec.C006), "0"))
        C_rec.C007 = oRecordSet.Fields("C007").VALUE
        C_rec.C008 = oRecordSet.Fields("C008").VALUE
        C_rec.C009 = oRecordSet.Fields("C009").VALUE
        C_rec.C009A = oRecordSet.Fields("C009A").VALUE  '2016
        C_rec.C010 = oRecordSet.Fields("C010").VALUE
        C_rec.C011 = oRecordSet.Fields("C011").VALUE
        C_rec.C012 = oRecordSet.Fields("C012").VALUE
        C_rec.C013 = oRecordSet.Fields("C013").VALUE
        C_rec.C014 = oRecordSet.Fields("C014").VALUE
        C_rec.C015 = oRecordSet.Fields("C015").VALUE
        C_rec.C016 = oRecordSet.Fields("C016").VALUE
        C_rec.C017 = oRecordSet.Fields("C017").VALUE
        C_rec.C018 = oRecordSet.Fields("C018").VALUE
        C_rec.C019 = oRecordSet.Fields("C019").VALUE
        C_rec.C020 = oRecordSet.Fields("C020").VALUE
        C_rec.C021 = oRecordSet.Fields("C021").VALUE
        C_rec.C022 = Format$(oRecordSet.Fields("C022").VALUE, String$(Len(C_rec.C022), "0"))
        C_rec.C023 = Format$(oRecordSet.Fields("C023").VALUE, String$(Len(C_rec.C023), "0"))
        C_rec.C024 = Format$(oRecordSet.Fields("C024").VALUE, String$(Len(C_rec.C024), "0"))
        C_rec.C025 = Format$(oRecordSet.Fields("C025").VALUE, String$(Len(C_rec.C025), "0"))
        C_rec.C026 = Format$(oRecordSet.Fields("C026").VALUE, String$(Len(C_rec.C026), "0"))
        C_rec.C027 = Format$(oRecordSet.Fields("C027").VALUE, String$(Len(C_rec.C027), "0"))
        C_rec.C028 = Format$(oRecordSet.Fields("C028").VALUE, String$(Len(C_rec.C028), "0"))
        C_rec.C029 = Format$(oRecordSet.Fields("C029").VALUE, String$(Len(C_rec.C029), "0"))
        C_rec.C030 = Format$(oRecordSet.Fields("C030").VALUE, String$(Len(C_rec.C030), "0"))
        C_rec.C031 = Format$(oRecordSet.Fields("C031").VALUE, String$(Len(C_rec.C031), "0"))
        C_rec.C032 = Format$(oRecordSet.Fields("C032").VALUE, String$(Len(C_rec.C032), "0"))
        C_rec.C033 = Format$(oRecordSet.Fields("C033").VALUE, String$(Len(C_rec.C033), "0"))
        C_rec.C034 = Format$(oRecordSet.Fields("C034").VALUE, String$(Len(C_rec.C034), "0"))
        C_rec.C035 = Format$(oRecordSet.Fields("C035").VALUE, String$(Len(C_rec.C035), "0"))
        C_rec.C036 = Format$(oRecordSet.Fields("C036").VALUE, String$(Len(C_rec.C036), "0"))
        C_rec.C037 = Format$(oRecordSet.Fields("C037").VALUE, String$(Len(C_rec.C037), "0"))
        C_rec.C038 = Format$(oRecordSet.Fields("C038").VALUE, String$(Len(C_rec.C038), "0"))
        C_rec.C039 = Format$(oRecordSet.Fields("C039").VALUE, String$(Len(C_rec.C039), "0"))
        C_rec.C040 = Format$(oRecordSet.Fields("C040").VALUE, String$(Len(C_rec.C040), "0"))
        C_rec.C041 = Format$(oRecordSet.Fields("C041").VALUE, String$(Len(C_rec.C041), "0"))
        C_rec.C042 = Format$(oRecordSet.Fields("C042").VALUE, String$(Len(C_rec.C042), "0"))
        C_rec.C043 = Format$(oRecordSet.Fields("C043").VALUE, String$(Len(C_rec.C043), "0"))
        C_rec.C044 = Format$(oRecordSet.Fields("C044").VALUE, String$(Len(C_rec.C044), "0"))
        C_rec.C045 = Format$(oRecordSet.Fields("C045").VALUE, String$(Len(C_rec.C045), "0"))
        C_rec.C046 = Format$(oRecordSet.Fields("C046").VALUE, String$(Len(C_rec.C046), "0"))
        C_rec.C047 = Format$(oRecordSet.Fields("C047").VALUE, String$(Len(C_rec.C047), "0"))
        C_rec.C048 = Format$(oRecordSet.Fields("C048").VALUE, String$(Len(C_rec.C048), "0"))
        C_rec.C049 = Format$(oRecordSet.Fields("C049").VALUE, String$(Len(C_rec.C049), "0"))
        C_rec.C050 = Format$(oRecordSet.Fields("C050").VALUE, String$(Len(C_rec.C050), "0"))
        C_rec.C051 = Format$(oRecordSet.Fields("C051").VALUE, String$(Len(C_rec.C051), "0"))
        C_rec.C052 = Format$(oRecordSet.Fields("C052").VALUE, String$(Len(C_rec.C052), "0"))
        C_rec.C053 = Format$(oRecordSet.Fields("C053").VALUE, String$(Len(C_rec.C053), "0"))
        C_rec.C054 = Format$(oRecordSet.Fields("C054").VALUE, String$(Len(C_rec.C054), "0"))
        C_rec.C055 = Format$(oRecordSet.Fields("C055").VALUE, String$(Len(C_rec.C055), "0"))
        C_rec.C056 = Format$(oRecordSet.Fields("C056").VALUE, String$(Len(C_rec.C056), "0"))
        C_rec.C056A = Format$(oRecordSet.Fields("C056A").VALUE, String$(Len(C_rec.C056A), "0"))
        C_rec.C057 = Format$(oRecordSet.Fields("C057").VALUE, String$(Len(C_rec.C057), "0"))
        C_rec.C058 = Format$(oRecordSet.Fields("C058").VALUE, String$(Len(C_rec.C058), "0"))
        C_rec.C058A = Format$(oRecordSet.Fields("C058A").VALUE, String$(Len(C_rec.C058A), "0"))
        C_rec.C058B = Format$(oRecordSet.Fields("C058B").VALUE, String$(Len(C_rec.C058B), "0"))
        C_rec.C059 = Format$(oRecordSet.Fields("C059").VALUE, String$(Len(C_rec.C059), "0"))
        C_rec.C060 = Format$(oRecordSet.Fields("C060").VALUE, String$(Len(C_rec.C060), "0"))
        C_rec.C061 = Format$(oRecordSet.Fields("C061").VALUE, String$(Len(C_rec.C061), "0"))
        'C_rec.C062 = Format$(oRecordSet.Fields("C062").VALUE, String$(Len(C_rec.C062), "0"))  2016
        C_rec.C063 = Format$(oRecordSet.Fields("C063").VALUE, String$(Len(C_rec.C063), "0"))
        C_rec.C064 = Format$(oRecordSet.Fields("C064").VALUE, String$(Len(C_rec.C064), "0"))
        C_rec.C065 = Format$(oRecordSet.Fields("C065").VALUE, String$(Len(C_rec.C065), "0"))
        C_rec.C066 = Format$(oRecordSet.Fields("C066").VALUE, String$(Len(C_rec.C066), "0"))
        C_rec.C067 = Format$(oRecordSet.Fields("C067").VALUE, String$(Len(C_rec.C067), "0"))
        C_rec.C068 = Format$(oRecordSet.Fields("C068").VALUE, String$(Len(C_rec.C068), "0"))
        C_rec.C069 = Format$(oRecordSet.Fields("C069").VALUE, String$(Len(C_rec.C069), "0"))
        C_rec.C070 = Format$(oRecordSet.Fields("C070").VALUE, String$(Len(C_rec.C070), "0"))
        C_rec.C071 = Format$(oRecordSet.Fields("C071").VALUE, String$(Len(C_rec.C071), "0"))
        C_rec.C072 = Format$(oRecordSet.Fields("C072").VALUE, String$(Len(C_rec.C072), "0"))
        C_rec.C073 = Format$(oRecordSet.Fields("C073").VALUE, String$(Len(C_rec.C073), "0"))
        C_rec.C074 = Format$(oRecordSet.Fields("C074").VALUE, String$(Len(C_rec.C074), "0"))
        C_rec.C075 = Format$(oRecordSet.Fields("C075").VALUE, String$(Len(C_rec.C075), "0"))
        C_rec.C076 = Format$(oRecordSet.Fields("C076").VALUE, String$(Len(C_rec.C076), "0"))
        C_rec.C077 = Format$(oRecordSet.Fields("C077").VALUE, String$(Len(C_rec.C077), "0"))
        C_rec.C078 = Format$(oRecordSet.Fields("C078").VALUE, String$(Len(C_rec.C078), "0"))
        C_rec.C079 = Format$(oRecordSet.Fields("C079").VALUE, String$(Len(C_rec.C079), "0"))
        C_rec.C080 = Format$(oRecordSet.Fields("C080").VALUE, String$(Len(C_rec.C080), "0"))
        C_rec.C081 = Format$(oRecordSet.Fields("C081").VALUE, String$(Len(C_rec.C081), "0"))
        C_rec.C082 = Format$(oRecordSet.Fields("C082").VALUE, String$(Len(C_rec.C082), "0"))
        C_rec.C083 = Format$(oRecordSet.Fields("C083").VALUE, String$(Len(C_rec.C083), "0"))
        C_rec.C084 = Format$(oRecordSet.Fields("C084").VALUE, String$(Len(C_rec.C084), "0"))
        C_rec.C085 = Format$(oRecordSet.Fields("C085").VALUE, String$(Len(C_rec.C085), "0"))
        C_rec.C086 = Format$(oRecordSet.Fields("C086").VALUE, String$(Len(C_rec.C086), "0"))
        C_rec.C087 = Format$(oRecordSet.Fields("C087").VALUE, String$(Len(C_rec.C087), "0"))
        C_rec.C088 = Format$(oRecordSet.Fields("C088").VALUE, String$(Len(C_rec.C088), "0"))
        C_rec.C089 = Format$(oRecordSet.Fields("C089").VALUE, String$(Len(C_rec.C089), "0"))
        C_rec.C090 = Format$(oRecordSet.Fields("C090").VALUE, String$(Len(C_rec.C090), "0"))
        C_rec.C091 = Format$(oRecordSet.Fields("C091").VALUE, String$(Len(C_rec.C091), "0"))
        C_rec.C092 = Format$(oRecordSet.Fields("C092").VALUE, String$(Len(C_rec.C092), "0"))
        C_rec.C093 = Format$(oRecordSet.Fields("C093").VALUE, String$(Len(C_rec.C093), "0"))
        C_rec.C094 = Format$(oRecordSet.Fields("C094").VALUE, String$(Len(C_rec.C094), "0"))
        'C_rec.C095 = Format$(oRecordSet.Fields("C095").VALUE, String$(Len(C_rec.C095), "0")) 2016
        C_rec.C096 = Format$(oRecordSet.Fields("C096").VALUE, String$(Len(C_rec.C096), "0"))
        C_rec.C097 = Format$(oRecordSet.Fields("C097").VALUE, String$(Len(C_rec.C097), "0"))
        C_rec.C098 = Format$(oRecordSet.Fields("C098").VALUE, String$(Len(C_rec.C098), "0"))
        C_rec.C099 = Format$(oRecordSet.Fields("C099").VALUE, String$(Len(C_rec.C099), "0"))
        C_rec.C100 = Format$(oRecordSet.Fields("C100").VALUE, String$(Len(C_rec.C100), "0"))
        C_rec.C101 = Format$(oRecordSet.Fields("C101").VALUE, String$(Len(C_rec.C101), "0"))
        C_rec.C102 = Format$(oRecordSet.Fields("C102").VALUE, String$(Len(C_rec.C102), "0"))
        C_rec.C103 = Format$(oRecordSet.Fields("C103").VALUE, String$(Len(C_rec.C103), "0"))
        C_rec.C104 = Format$(oRecordSet.Fields("C104").VALUE, String$(Len(C_rec.C104), "0"))
        C_rec.C105 = Format$(oRecordSet.Fields("C105").VALUE, String$(Len(C_rec.C105), "0"))
        'C_rec.C106 = Format$(oRecordSet.Fields("C106").VALUE, String$(Len(C_rec.C106), "0")) 2016
        C_rec.C107 = Format$(oRecordSet.Fields("C107").VALUE, String$(Len(C_rec.C107), "0"))
        'C_rec.C108 = Format$(oRecordSet.Fields("C108").VALUE, String$(Len(C_rec.C108), "0")) 2017
        C_rec.C109 = Format$(oRecordSet.Fields("C109").VALUE, String$(Len(C_rec.C109), "0"))
        C_rec.C110 = Format$(oRecordSet.Fields("C110").VALUE, String$(Len(C_rec.C110), "0"))
        C_rec.C111 = Format$(oRecordSet.Fields("C111").VALUE, String$(Len(C_rec.C111), "0"))
        C_rec.C112 = Format$(oRecordSet.Fields("C112").VALUE, String$(Len(C_rec.C112), "0"))
        C_rec.C113 = Format$(oRecordSet.Fields("C113").VALUE, String$(Len(C_rec.C113), "0"))
        C_rec.C114 = Format$(oRecordSet.Fields("C114").VALUE, String$(Len(C_rec.C114), "0"))
        C_rec.C115 = Format$(oRecordSet.Fields("C115").VALUE, String$(Len(C_rec.C115), "0"))
        C_rec.C116 = Format$(oRecordSet.Fields("C116").VALUE, String$(Len(C_rec.C116), "0"))
        C_rec.C117 = Format$(oRecordSet.Fields("C117").VALUE, String$(Len(C_rec.C117), "0"))
        'C_rec.C118 = Format$(oRecordSet.Fields("C118").VALUE, String$(Len(C_rec.C118), "0")) 2016
        C_rec.C119 = Format$(oRecordSet.Fields("C119").VALUE, String$(Len(C_rec.C119), "0"))
        C_rec.C120 = Format$(oRecordSet.Fields("C120").VALUE, String$(Len(C_rec.C120), "0"))
        C_rec.C121 = Format$(oRecordSet.Fields("C121").VALUE, String$(Len(C_rec.C121), "0"))
        C_rec.C122 = Format$(oRecordSet.Fields("C122").VALUE, String$(Len(C_rec.C122), "0"))
        C_rec.C123 = Format$(oRecordSet.Fields("C123").VALUE, String$(Len(C_rec.C123), "0"))
        C_rec.C124 = Format$(oRecordSet.Fields("C124").VALUE, String$(Len(C_rec.C124), "0"))
        C_rec.C125 = Format$(oRecordSet.Fields("C125").VALUE, String$(Len(C_rec.C125), "0"))
        C_rec.C126 = Format$(oRecordSet.Fields("C126").VALUE, String$(Len(C_rec.C126), "0"))
        C_rec.C127 = Format$(oRecordSet.Fields("C127").VALUE, String$(Len(C_rec.C127), "0"))
        C_rec.C128 = Format$(oRecordSet.Fields("C128").VALUE, String$(Len(C_rec.C128), "0"))
        C_rec.C129 = Format$(oRecordSet.Fields("C129").VALUE, String$(Len(C_rec.C129), "0"))
        C_rec.C130 = Format$(oRecordSet.Fields("C130").VALUE, String$(Len(C_rec.C130), "0"))
        C_rec.C131 = Format$(oRecordSet.Fields("C131").VALUE, String$(Len(C_rec.C131), "0"))
        C_rec.C132 = Format$(oRecordSet.Fields("C132").VALUE, String$(Len(C_rec.C132), "0"))
        C_rec.C133 = Format$(oRecordSet.Fields("C133").VALUE, String$(Len(C_rec.C133), "0"))
        C_rec.C134 = Format$(oRecordSet.Fields("C134").VALUE, String$(Len(C_rec.C134), "0"))
        C_rec.C135 = Format$(oRecordSet.Fields("C135").VALUE, String$(Len(C_rec.C135), "0"))
        C_rec.C136 = Format$(oRecordSet.Fields("C136").VALUE, String$(Len(C_rec.C136), "0"))
        C_rec.C137 = Format$(oRecordSet.Fields("C137").VALUE, String$(Len(C_rec.C137), "0"))
        C_rec.C138 = Format$(oRecordSet.Fields("C138").VALUE, String$(Len(C_rec.C138), "0"))
        C_rec.C139 = Format$(oRecordSet.Fields("C139").VALUE, String$(Len(C_rec.C139), "0"))
        C_rec.C140 = Format$(oRecordSet.Fields("C140").VALUE, String$(Len(C_rec.C140), "0"))
        C_rec.C141 = Format$(oRecordSet.Fields("C141").VALUE, String$(Len(C_rec.C141), "0"))
        C_rec.C142 = Format$(oRecordSet.Fields("C142").VALUE, String$(Len(C_rec.C142), "0"))
        C_rec.C143 = Format$(oRecordSet.Fields("C143").VALUE, String$(Len(C_rec.C143), "0"))
        C_rec.C144 = Format$(oRecordSet.Fields("C144").VALUE, String$(Len(C_rec.C144), "0"))
        C_rec.C145 = Format$(oRecordSet.Fields("C145").VALUE, String$(Len(C_rec.C145), "0"))
        C_rec.C146 = Format$(oRecordSet.Fields("C146").VALUE, String$(Len(C_rec.C146), "0"))
        C_rec.C147 = Format$(oRecordSet.Fields("C147").VALUE, String$(Len(C_rec.C147), "0"))
        C_rec.C148 = Format$(oRecordSet.Fields("C148").VALUE, String$(Len(C_rec.C148), "0"))
        C_rec.C149 = Format$(oRecordSet.Fields("C149").VALUE, String$(Len(C_rec.C149), "0"))
        C_rec.C150 = Format$(oRecordSet.Fields("C150").VALUE, String$(Len(C_rec.C150), "0"))
        C_rec.C150A = Format$(oRecordSet.Fields("C150A").VALUE, String$(Len(C_rec.C150A), "0"))   '2016
        C_rec.C150B = Format$(oRecordSet.Fields("C150B").VALUE, String$(Len(C_rec.C150B), "0"))   '2016
        C_rec.C151 = Format$(oRecordSet.Fields("C151").VALUE, String$(Len(C_rec.C151), "0"))
        C_rec.C152 = Format$(oRecordSet.Fields("C152").VALUE, String$(Len(C_rec.C152), "0"))
        C_rec.C153 = Format$(oRecordSet.Fields("C153").VALUE, String$(Len(C_rec.C153), "0"))
        C_rec.C154 = Format$(oRecordSet.Fields("C154").VALUE, String$(Len(C_rec.C154), "0"))
        C_rec.C155 = Format$(oRecordSet.Fields("C155").VALUE, String$(Len(C_rec.C155), "0"))
        C_rec.C156 = Format$(oRecordSet.Fields("C156").VALUE, String$(Len(C_rec.C156), "0"))
        C_rec.C157 = Format$(oRecordSet.Fields("C157").VALUE, String$(Len(C_rec.C157), "0"))
        C_rec.C158 = Format$(oRecordSet.Fields("C158").VALUE, String$(Len(C_rec.C158), "0"))
        C_rec.C159 = Format$(oRecordSet.Fields("C159").VALUE, String$(Len(C_rec.C159), "0"))
        C_rec.C160 = Format$(oRecordSet.Fields("C160").VALUE, String$(Len(C_rec.C160), "0"))
        C_rec.C161 = Format$(oRecordSet.Fields("C161").VALUE, String$(Len(C_rec.C161), "0"))
        C_rec.C162 = Format$(oRecordSet.Fields("C162").VALUE, String$(Len(C_rec.C162), "0"))
        C_rec.C163 = Format$(oRecordSet.Fields("C163").VALUE, String$(Len(C_rec.C163), "0"))
        C_rec.C164 = Format$(oRecordSet.Fields("C164").VALUE, String$(Len(C_rec.C164), "0"))
        C_rec.C165 = Format$(oRecordSet.Fields("C165").VALUE, String$(Len(C_rec.C165), "0"))
        C_rec.C166 = Format$(oRecordSet.Fields("C166").VALUE, String$(Len(C_rec.C166), "0"))
        C_rec.C167 = Format$(oRecordSet.Fields("C167").VALUE, String$(Len(C_rec.C167), "0"))
        
        C_rec.C168_1 = Format$(oRecordSet.Fields("C168_1").VALUE, String$(Len(C_rec.C168_1), "0"))
        C_rec.C168_2 = Format$(oRecordSet.Fields("C168_2").VALUE, String$(Len(C_rec.C168_2), "0"))
        C_rec.C169_1 = Format$(oRecordSet.Fields("C169_1").VALUE, String$(Len(C_rec.C169_1), "0"))
        C_rec.C169_2 = Format$(oRecordSet.Fields("C169_2").VALUE, String$(Len(C_rec.C169_2), "0"))
        C_rec.C170_1 = Format$(oRecordSet.Fields("C170_1").VALUE, String$(Len(C_rec.C170_1), "0"))
        C_rec.C170_2 = Format$(oRecordSet.Fields("C170_2").VALUE, String$(Len(C_rec.C170_2), "0"))
        C_rec.C171 = oRecordSet.Fields("C171").VALUE
        C_rec.C172 = oRecordSet.Fields("C172").VALUE
        C_rec.C173 = oRecordSet.Fields("C173").VALUE
        
        '예제
        'C_rec.PERNBR = Replace(oRecordSet.Fields("U_PERNBR").VALUE, "-", "")
        
        'OLDBIG = Val(oRecordSet.Fields("U_BIGWA1").VALUE) + Val(oRecordSet.Fields("U_BIGWA3").VALUE) + Val(oRecordSet.Fields("U_BIGWA5").VALUE) _
        '        + Val(oRecordSet.Fields("U_BIGWA6").VALUE) + Val(oRecordSet.Fields("U_BIGWU3").VALUE)
                              
        'C_rec.FILD02 = Format$(0, String$(Len(C_rec.FILD02), "0"))
        'C_rec.GAMFLD = String$(Len(C_rec.GAMFLD), "0")
        'C_rec.FILLER = Space$(Len(C_rec.FILLER))
        'C_rec.C022 = Format$(oRecordSet.Fields("C022").VALUE, , String$(Len(C_rec.C022), "0"))
        
      
        Print #1, MDC_SetMod.sStr(C_rec.C001) & MDC_SetMod.sStr(C_rec.C002) & MDC_SetMod.sStr(C_rec.C003) & MDC_SetMod.sStr(C_rec.C004) & MDC_SetMod.sStr(C_rec.C005) & MDC_SetMod.sStr(C_rec.C006) & MDC_SetMod.sStr(C_rec.C007) & MDC_SetMod.sStr(C_rec.C008) & MDC_SetMod.sStr(C_rec.C009) & MDC_SetMod.sStr(C_rec.C009A) & _
                  MDC_SetMod.sStr(C_rec.C010) & MDC_SetMod.sStr(C_rec.C011) & MDC_SetMod.sStr(C_rec.C012) & MDC_SetMod.sStr(C_rec.C013) & MDC_SetMod.sStr(C_rec.C014) & MDC_SetMod.sStr(C_rec.C015) & MDC_SetMod.sStr(C_rec.C016) & MDC_SetMod.sStr(C_rec.C017) & MDC_SetMod.sStr(C_rec.C018) & MDC_SetMod.sStr(C_rec.C019) & _
                  MDC_SetMod.sStr(C_rec.C020) & MDC_SetMod.sStr(C_rec.C021) & MDC_SetMod.sStr(C_rec.C022) & MDC_SetMod.sStr(C_rec.C023) & MDC_SetMod.sStr(C_rec.C024) & MDC_SetMod.sStr(C_rec.C025) & MDC_SetMod.sStr(C_rec.C026) & MDC_SetMod.sStr(C_rec.C027) & MDC_SetMod.sStr(C_rec.C028) & MDC_SetMod.sStr(C_rec.C029) & _
                  MDC_SetMod.sStr(C_rec.C030) & MDC_SetMod.sStr(C_rec.C031) & MDC_SetMod.sStr(C_rec.C032) & MDC_SetMod.sStr(C_rec.C033) & MDC_SetMod.sStr(C_rec.C034) & MDC_SetMod.sStr(C_rec.C035) & MDC_SetMod.sStr(C_rec.C036) & MDC_SetMod.sStr(C_rec.C037) & MDC_SetMod.sStr(C_rec.C038) & MDC_SetMod.sStr(C_rec.C039) & _
                  MDC_SetMod.sStr(C_rec.C040) & MDC_SetMod.sStr(C_rec.C041) & MDC_SetMod.sStr(C_rec.C042) & MDC_SetMod.sStr(C_rec.C043) & MDC_SetMod.sStr(C_rec.C044) & MDC_SetMod.sStr(C_rec.C045) & MDC_SetMod.sStr(C_rec.C046) & MDC_SetMod.sStr(C_rec.C047) & MDC_SetMod.sStr(C_rec.C048) & MDC_SetMod.sStr(C_rec.C049) & _
                  MDC_SetMod.sStr(C_rec.C050) & MDC_SetMod.sStr(C_rec.C051) & MDC_SetMod.sStr(C_rec.C052) & MDC_SetMod.sStr(C_rec.C053) & MDC_SetMod.sStr(C_rec.C054) & MDC_SetMod.sStr(C_rec.C055) & MDC_SetMod.sStr(C_rec.C056) & MDC_SetMod.sStr(C_rec.C056A) & MDC_SetMod.sStr(C_rec.C057) & MDC_SetMod.sStr(C_rec.C058) & _
                  MDC_SetMod.sStr(C_rec.C058A) & MDC_SetMod.sStr(C_rec.C058B) & MDC_SetMod.sStr(C_rec.C059) & MDC_SetMod.sStr(C_rec.C060) & MDC_SetMod.sStr(C_rec.C061) & MDC_SetMod.sStr(C_rec.C063) & MDC_SetMod.sStr(C_rec.C064) & MDC_SetMod.sStr(C_rec.C065) & MDC_SetMod.sStr(C_rec.C066) & MDC_SetMod.sStr(C_rec.C067) & _
                  MDC_SetMod.sStr(C_rec.C068) & MDC_SetMod.sStr(C_rec.C069) & MDC_SetMod.sStr(C_rec.C070) & MDC_SetMod.sStr(C_rec.C071) & MDC_SetMod.sStr(C_rec.C072) & MDC_SetMod.sStr(C_rec.C073) & MDC_SetMod.sStr(C_rec.C074) & MDC_SetMod.sStr(C_rec.C075) & MDC_SetMod.sStr(C_rec.C076) & MDC_SetMod.sStr(C_rec.C077) & _
                  MDC_SetMod.sStr(C_rec.C078) & MDC_SetMod.sStr(C_rec.C079) & MDC_SetMod.sStr(C_rec.C080) & MDC_SetMod.sStr(C_rec.C081) & MDC_SetMod.sStr(C_rec.C082) & MDC_SetMod.sStr(C_rec.C083) & MDC_SetMod.sStr(C_rec.C084) & MDC_SetMod.sStr(C_rec.C085) & MDC_SetMod.sStr(C_rec.C086) & MDC_SetMod.sStr(C_rec.C087) & _
                  MDC_SetMod.sStr(C_rec.C088) & MDC_SetMod.sStr(C_rec.C089) & MDC_SetMod.sStr(C_rec.C090) & MDC_SetMod.sStr(C_rec.C091) & MDC_SetMod.sStr(C_rec.C092) & MDC_SetMod.sStr(C_rec.C093) & MDC_SetMod.sStr(C_rec.C094) & MDC_SetMod.sStr(C_rec.C096) & MDC_SetMod.sStr(C_rec.C097) & MDC_SetMod.sStr(C_rec.C098) & _
                  MDC_SetMod.sStr(C_rec.C099) & MDC_SetMod.sStr(C_rec.C100) & MDC_SetMod.sStr(C_rec.C101) & MDC_SetMod.sStr(C_rec.C102) & MDC_SetMod.sStr(C_rec.C103) & MDC_SetMod.sStr(C_rec.C104) & MDC_SetMod.sStr(C_rec.C105) & MDC_SetMod.sStr(C_rec.C107) & MDC_SetMod.sStr(C_rec.C109) & MDC_SetMod.sStr(C_rec.C110) & _
                  MDC_SetMod.sStr(C_rec.C111) & MDC_SetMod.sStr(C_rec.C112) & MDC_SetMod.sStr(C_rec.C113) & MDC_SetMod.sStr(C_rec.C114) & MDC_SetMod.sStr(C_rec.C115) & MDC_SetMod.sStr(C_rec.C116) & MDC_SetMod.sStr(C_rec.C117) & MDC_SetMod.sStr(C_rec.C119) & MDC_SetMod.sStr(C_rec.C120) & MDC_SetMod.sStr(C_rec.C121) & _
                  MDC_SetMod.sStr(C_rec.C122) & MDC_SetMod.sStr(C_rec.C123) & MDC_SetMod.sStr(C_rec.C124) & MDC_SetMod.sStr(C_rec.C125) & MDC_SetMod.sStr(C_rec.C126) & MDC_SetMod.sStr(C_rec.C127) & MDC_SetMod.sStr(C_rec.C128) & MDC_SetMod.sStr(C_rec.C129) & MDC_SetMod.sStr(C_rec.C130) & MDC_SetMod.sStr(C_rec.C131) & _
                  MDC_SetMod.sStr(C_rec.C132) & MDC_SetMod.sStr(C_rec.C133) & MDC_SetMod.sStr(C_rec.C134) & MDC_SetMod.sStr(C_rec.C135) & MDC_SetMod.sStr(C_rec.C136) & MDC_SetMod.sStr(C_rec.C137) & MDC_SetMod.sStr(C_rec.C138) & MDC_SetMod.sStr(C_rec.C139) & MDC_SetMod.sStr(C_rec.C140) & MDC_SetMod.sStr(C_rec.C141) & _
                  MDC_SetMod.sStr(C_rec.C142) & MDC_SetMod.sStr(C_rec.C143) & MDC_SetMod.sStr(C_rec.C144) & MDC_SetMod.sStr(C_rec.C145) & MDC_SetMod.sStr(C_rec.C146) & MDC_SetMod.sStr(C_rec.C147) & MDC_SetMod.sStr(C_rec.C148) & MDC_SetMod.sStr(C_rec.C149) & MDC_SetMod.sStr(C_rec.C150) & MDC_SetMod.sStr(C_rec.C150A) & _
                  MDC_SetMod.sStr(C_rec.C150B) & MDC_SetMod.sStr(C_rec.C151) & MDC_SetMod.sStr(C_rec.C152) & MDC_SetMod.sStr(C_rec.C153) & MDC_SetMod.sStr(C_rec.C154) & MDC_SetMod.sStr(C_rec.C155) & MDC_SetMod.sStr(C_rec.C156) & MDC_SetMod.sStr(C_rec.C157) & MDC_SetMod.sStr(C_rec.C158) & MDC_SetMod.sStr(C_rec.C159) & _
                  MDC_SetMod.sStr(C_rec.C160) & MDC_SetMod.sStr(C_rec.C161) & MDC_SetMod.sStr(C_rec.C162) & MDC_SetMod.sStr(C_rec.C163) & MDC_SetMod.sStr(C_rec.C164) & MDC_SetMod.sStr(C_rec.C165) & MDC_SetMod.sStr(C_rec.C166) & MDC_SetMod.sStr(C_rec.C167) & MDC_SetMod.sStr(C_rec.C168_1) & MDC_SetMod.sStr(C_rec.C168_2) & _
                  MDC_SetMod.sStr(C_rec.C169_1) & MDC_SetMod.sStr(C_rec.C169_2) & MDC_SetMod.sStr(C_rec.C170_1) & MDC_SetMod.sStr(C_rec.C170_2) & MDC_SetMod.sStr(C_rec.C171) & MDC_SetMod.sStr(C_rec.C172) & MDC_SetMod.sStr(C_rec.C173)
                   
        
        '/ D레코드: 종전근무처 레코드
        If Val(C_rec.C006) > 0 Then
            If File_Create_D_record = False Then
                ErrNum = 2
                GoTo Error_Message
            End If
        End If
        
        '/ E레코드: 부양가족 레코드
        If File_Create_E_record = False Then
           ErrNum = 3
           GoTo Error_Message
        End If
        
        '/ F레코드
        If File_Create_F_record = False Then
           ErrNum = 4
           GoTo Error_Message
        End If
        
        '/ G레코드
        If File_Create_G_record = False Then
           ErrNum = 5
           GoTo Error_Message
        End If
        
        '/ H레코드
        If File_Create_H_record = False Then
           ErrNum = 6
           GoTo Error_Message
        End If
        
        '/ I레코드
        If File_Create_I_record = False Then
           ErrNum = 7
           GoTo Error_Message
        End If
        
        
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 작성중........!"
        
        
        oRecordSet.MoveNext
    Loop
        
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
        
    If CheckC = False Then
        File_Create_C_record = True
    Else
        File_Create_C_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "C레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "D레코드(종전근무처 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "E레코드(부양가족 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "F레코드(연금.저축 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "G레코드(월세액.주택자료 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "H레코드(기부금조정명세 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "I레코드(해당연도 기부금명세 레코드) 생성 실패.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("C레코드오류: " & Err.Description, False)
    End If
    File_Create_C_record = False
End Function

Private Function File_Create_D_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckD      As String
    Dim JONCNT      As Integer
    
    CheckD = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ D_RECORE QUERY
    sQry = "EXEC PH_PY980_D '" & C_SAUP & "', '" & C_YYYY & "', '" & C_SABUN & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    Do Until oRecordSet.EOF
    
        'D RECORD MOVE
        
        D_rec.D001 = oRecordSet.Fields("D001").VALUE
        D_rec.D002 = oRecordSet.Fields("D002").VALUE
        D_rec.D003 = oRecordSet.Fields("D003").VALUE
        D_rec.D004 = Format$(C_rec.C004, String$(Len(D_rec.D004), "0"))  '/ C레코드의 일련번호
        D_rec.D005 = oRecordSet.Fields("D005").VALUE
        D_rec.D006 = oRecordSet.Fields("D006").VALUE
        D_rec.D007 = oRecordSet.Fields("D007").VALUE
        D_rec.D008 = oRecordSet.Fields("D008").VALUE
        D_rec.D009 = oRecordSet.Fields("D009").VALUE
        D_rec.D010 = oRecordSet.Fields("D010").VALUE
        D_rec.D011 = oRecordSet.Fields("D011").VALUE
        D_rec.D012 = oRecordSet.Fields("D012").VALUE
        D_rec.D013 = oRecordSet.Fields("D013").VALUE
        D_rec.D014 = oRecordSet.Fields("D014").VALUE
        D_rec.D015 = Format$(oRecordSet.Fields("D015").VALUE, String$(Len(D_rec.D015), "0"))
        D_rec.D016 = Format$(oRecordSet.Fields("D016").VALUE, String$(Len(D_rec.D016), "0"))
        D_rec.D017 = Format$(oRecordSet.Fields("D017").VALUE, String$(Len(D_rec.D017), "0"))
        D_rec.D018 = Format$(oRecordSet.Fields("D018").VALUE, String$(Len(D_rec.D018), "0"))
        D_rec.D019 = Format$(oRecordSet.Fields("D019").VALUE, String$(Len(D_rec.D019), "0"))
        D_rec.D020 = Format$(oRecordSet.Fields("D020").VALUE, String$(Len(D_rec.D020), "0"))
        D_rec.D021 = Format$(oRecordSet.Fields("D021").VALUE, String$(Len(D_rec.D021), "0"))
        D_rec.D022 = Format$(oRecordSet.Fields("D022").VALUE, String$(Len(D_rec.D022), "0"))
        D_rec.D023 = Format$(oRecordSet.Fields("D023").VALUE, String$(Len(D_rec.D023), "0"))
        D_rec.D024 = Format$(oRecordSet.Fields("D024").VALUE, String$(Len(D_rec.D024), "0"))
        D_rec.D025 = Format$(oRecordSet.Fields("D025").VALUE, String$(Len(D_rec.D025), "0"))
        D_rec.D026 = Format$(oRecordSet.Fields("D026").VALUE, String$(Len(D_rec.D026), "0"))
        D_rec.D027 = Format$(oRecordSet.Fields("D027").VALUE, String$(Len(D_rec.D027), "0"))
        D_rec.D028 = Format$(oRecordSet.Fields("D028").VALUE, String$(Len(D_rec.D028), "0"))
        D_rec.D029 = Format$(oRecordSet.Fields("D029").VALUE, String$(Len(D_rec.D029), "0"))
        D_rec.D030 = Format$(oRecordSet.Fields("D030").VALUE, String$(Len(D_rec.D030), "0"))
        D_rec.D031 = Format$(oRecordSet.Fields("D031").VALUE, String$(Len(D_rec.D031), "0"))
        D_rec.D032 = Format$(oRecordSet.Fields("D032").VALUE, String$(Len(D_rec.D032), "0"))
        D_rec.D033 = Format$(oRecordSet.Fields("D033").VALUE, String$(Len(D_rec.D033), "0"))
        D_rec.D034 = Format$(oRecordSet.Fields("D034").VALUE, String$(Len(D_rec.D034), "0"))
        D_rec.D035 = Format$(oRecordSet.Fields("D035").VALUE, String$(Len(D_rec.D035), "0"))
        D_rec.D036 = Format$(oRecordSet.Fields("D036").VALUE, String$(Len(D_rec.D036), "0"))
        D_rec.D037 = Format$(oRecordSet.Fields("D037").VALUE, String$(Len(D_rec.D037), "0"))
        D_rec.D038 = Format$(oRecordSet.Fields("D038").VALUE, String$(Len(D_rec.D038), "0"))
        D_rec.D039 = Format$(oRecordSet.Fields("D039").VALUE, String$(Len(D_rec.D039), "0"))
        D_rec.D040 = Format$(oRecordSet.Fields("D040").VALUE, String$(Len(D_rec.D040), "0"))
        D_rec.D041 = Format$(oRecordSet.Fields("D041").VALUE, String$(Len(D_rec.D041), "0"))
        D_rec.D042 = Format$(oRecordSet.Fields("D042").VALUE, String$(Len(D_rec.D042), "0"))
        D_rec.D043 = Format$(oRecordSet.Fields("D043").VALUE, String$(Len(D_rec.D043), "0"))
        D_rec.D044 = Format$(oRecordSet.Fields("D044").VALUE, String$(Len(D_rec.D044), "0"))
        D_rec.D045 = Format$(oRecordSet.Fields("D045").VALUE, String$(Len(D_rec.D045), "0"))
        D_rec.D046 = Format$(oRecordSet.Fields("D046").VALUE, String$(Len(D_rec.D046), "0"))
        D_rec.D047 = Format$(oRecordSet.Fields("D047").VALUE, String$(Len(D_rec.D047), "0"))
        D_rec.D048 = Format$(oRecordSet.Fields("D048").VALUE, String$(Len(D_rec.D048), "0"))
        D_rec.D049 = Format$(oRecordSet.Fields("D049").VALUE, String$(Len(D_rec.D049), "0"))
        D_rec.D049A = Format$(oRecordSet.Fields("D049A").VALUE, String$(Len(D_rec.D049A), "0"))
        D_rec.D050 = Format$(oRecordSet.Fields("D050").VALUE, String$(Len(D_rec.D050), "0"))
        D_rec.D051 = Format$(oRecordSet.Fields("D051").VALUE, String$(Len(D_rec.D051), "0"))
        D_rec.D051A = Format$(oRecordSet.Fields("D051A").VALUE, String$(Len(D_rec.D051A), "0"))  '2016
        D_rec.D051B = Format$(oRecordSet.Fields("D051B").VALUE, String$(Len(D_rec.D051B), "0"))  '2016
        D_rec.D052 = Format$(oRecordSet.Fields("D052").VALUE, String$(Len(D_rec.D052), "0"))
        D_rec.D053 = Format$(oRecordSet.Fields("D053").VALUE, String$(Len(D_rec.D053), "0"))
        D_rec.D054 = Format$(oRecordSet.Fields("D054").VALUE, String$(Len(D_rec.D054), "0"))
        D_rec.D055 = Format$(oRecordSet.Fields("D055").VALUE, String$(Len(D_rec.D055), "0"))
        D_rec.D056 = Format$(oRecordSet.Fields("D056").VALUE, String$(Len(D_rec.D056), "0"))
        D_rec.D057 = Format$(oRecordSet.Fields("D057").VALUE, String$(Len(D_rec.D057), "0"))
        D_rec.D058 = Format$(oRecordSet.Fields("D058").VALUE, String$(Len(D_rec.D058), "0"))
        D_rec.D059 = Format$(oRecordSet.Fields("D059").VALUE, String$(Len(D_rec.D059), "0"))
        D_rec.D060 = oRecordSet.Fields("D060").VALUE
        
        Print #1, MDC_SetMod.sStr(D_rec.D001) & MDC_SetMod.sStr(D_rec.D002) & MDC_SetMod.sStr(D_rec.D003) & MDC_SetMod.sStr(D_rec.D004) & MDC_SetMod.sStr(D_rec.D005) & MDC_SetMod.sStr(D_rec.D006) & MDC_SetMod.sStr(D_rec.D007) & MDC_SetMod.sStr(D_rec.D008) & MDC_SetMod.sStr(D_rec.D009) & MDC_SetMod.sStr(D_rec.D010) & _
                  MDC_SetMod.sStr(D_rec.D011) & MDC_SetMod.sStr(D_rec.D012) & MDC_SetMod.sStr(D_rec.D013) & MDC_SetMod.sStr(D_rec.D014) & MDC_SetMod.sStr(D_rec.D015) & MDC_SetMod.sStr(D_rec.D016) & MDC_SetMod.sStr(D_rec.D017) & MDC_SetMod.sStr(D_rec.D018) & MDC_SetMod.sStr(D_rec.D019) & MDC_SetMod.sStr(D_rec.D020) & _
                  MDC_SetMod.sStr(D_rec.D021) & MDC_SetMod.sStr(D_rec.D022) & MDC_SetMod.sStr(D_rec.D023) & MDC_SetMod.sStr(D_rec.D024) & MDC_SetMod.sStr(D_rec.D025) & MDC_SetMod.sStr(D_rec.D026) & MDC_SetMod.sStr(D_rec.D027) & MDC_SetMod.sStr(D_rec.D028) & MDC_SetMod.sStr(D_rec.D029) & MDC_SetMod.sStr(D_rec.D030) & _
                  MDC_SetMod.sStr(D_rec.D031) & MDC_SetMod.sStr(D_rec.D032) & MDC_SetMod.sStr(D_rec.D033) & MDC_SetMod.sStr(D_rec.D034) & MDC_SetMod.sStr(D_rec.D035) & MDC_SetMod.sStr(D_rec.D036) & MDC_SetMod.sStr(D_rec.D037) & MDC_SetMod.sStr(D_rec.D038) & MDC_SetMod.sStr(D_rec.D039) & MDC_SetMod.sStr(D_rec.D040) & _
                  MDC_SetMod.sStr(D_rec.D041) & MDC_SetMod.sStr(D_rec.D042) & MDC_SetMod.sStr(D_rec.D043) & MDC_SetMod.sStr(D_rec.D044) & MDC_SetMod.sStr(D_rec.D045) & MDC_SetMod.sStr(D_rec.D046) & MDC_SetMod.sStr(D_rec.D047) & MDC_SetMod.sStr(D_rec.D048) & MDC_SetMod.sStr(D_rec.D049) & MDC_SetMod.sStr(D_rec.D049A) & _
                  MDC_SetMod.sStr(D_rec.D050) & MDC_SetMod.sStr(D_rec.D051) & MDC_SetMod.sStr(D_rec.D051A) & MDC_SetMod.sStr(D_rec.D051B) & MDC_SetMod.sStr(D_rec.D052) & MDC_SetMod.sStr(D_rec.D053) & MDC_SetMod.sStr(D_rec.D054) & MDC_SetMod.sStr(D_rec.D055) & MDC_SetMod.sStr(D_rec.D056) & MDC_SetMod.sStr(D_rec.D057) & _
                  MDC_SetMod.sStr(D_rec.D058) & MDC_SetMod.sStr(D_rec.D059) & MDC_SetMod.sStr(D_rec.D060)

        oRecordSet.MoveNext
    Loop
    
    If CheckD = False Then
        File_Create_D_record = True
    Else
        File_Create_D_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "종전근무지D레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("D레코드오류: " & Err.Description, False)
    End If
    File_Create_D_record = False
End Function

Private Function File_Create_E_record() As Boolean
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
    
    '/ E_RECORE QUERY
    sQry = "EXEC PH_PY980_E '" & C_SAUP & "', '" & C_YYYY & "', '" & C_SABUN & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount > 0 Then
    
        BUYCNT = 0  '가족수
        FAMCNT = 1  'E레코드일련번호
        
        E_rec.E001 = oRecordSet.Fields("E001").VALUE
        E_rec.E002 = oRecordSet.Fields("E002").VALUE
        E_rec.E003 = oRecordSet.Fields("E003").VALUE
        E_rec.E004 = Format$(C_rec.C004, String$(Len(E_rec.E004), "0"))  '/ C레코드의 일련번호
        E_rec.E005 = oRecordSet.Fields("E005").VALUE
        E_rec.E006 = oRecordSet.Fields("E006").VALUE
         
        'TOTBOH = 0: TOTMED = 0: TOTEDC = 0: TOTCAD = 0: TOTCSH = 0: TOTGBU = 0
        
        Do Until oRecordSet.EOF
            BUYCNT = BUYCNT + 1   '/ 해당사원의 부양가족일련번호
            '/초기화
            If BUYCNT = 1 Then
                For i = 1 To 5
                    E_rec.E007(i) = Space$(Len(E_rec.E007(i)))
                    E_rec.E008(i) = Space$(Len(E_rec.E008(i)))
                    E_rec.E009(i) = Space$(Len(E_rec.E009(i)))
                    E_rec.E010(i) = Space$(Len(E_rec.E010(i)))
                    E_rec.E011(i) = Space$(Len(E_rec.E011(i)))
                    E_rec.E012(i) = Space$(Len(E_rec.E012(i)))
                    E_rec.E013(i) = Space$(Len(E_rec.E013(i)))
                    E_rec.E014(i) = Space$(Len(E_rec.E014(i)))
                    E_rec.E015(i) = Space$(Len(E_rec.E015(i)))
                    E_rec.E016(i) = Space$(Len(E_rec.E016(i)))
                    E_rec.E017(i) = Space$(Len(E_rec.E017(i)))
                    E_rec.E018(i) = Space$(Len(E_rec.E018(i)))
                    
                    E_rec.E019(i) = Format$(0, String$(Len(E_rec.E019(i)), "0"))
                    E_rec.E020(i) = Format$(0, String$(Len(E_rec.E020(i)), "0"))
                    E_rec.E021(i) = Format$(0, String$(Len(E_rec.E021(i)), "0"))
                    E_rec.E022(i) = Format$(0, String$(Len(E_rec.E022(i)), "0"))
                    E_rec.E023(i) = Format$(0, String$(Len(E_rec.E023(i)), "0"))
                    E_rec.E024(i) = Format$(0, String$(Len(E_rec.E024(i)), "0"))
                    E_rec.E025(i) = Format$(0, String$(Len(E_rec.E025(i)), "0"))
                    E_rec.E026(i) = Format$(0, String$(Len(E_rec.E026(i)), "0"))
                    E_rec.E027(i) = Format$(0, String$(Len(E_rec.E027(i)), "0"))
                    E_rec.E028(i) = Format$(0, String$(Len(E_rec.E028(i)), "0"))
                    E_rec.E029(i) = Format$(0, String$(Len(E_rec.E029(i)), "0"))
                    E_rec.E030(i) = Format$(0, String$(Len(E_rec.E030(i)), "0"))
                    E_rec.E031(i) = Format$(0, String$(Len(E_rec.E031(i)), "0"))
                    E_rec.E032(i) = Format$(0, String$(Len(E_rec.E032(i)), "0"))
                    E_rec.E033(i) = Format$(0, String$(Len(E_rec.E033(i)), "0"))
                    E_rec.E034(i) = Format$(0, String$(Len(E_rec.E034(i)), "0"))
                    E_rec.E035(i) = Format$(0, String$(Len(E_rec.E035(i)), "0"))
                    E_rec.E036(i) = Format$(0, String$(Len(E_rec.E036(i)), "0"))
                    E_rec.E037(i) = Format$(0, String$(Len(E_rec.E037(i)), "0"))
                    E_rec.E038(i) = Format$(0, String$(Len(E_rec.E038(i)), "0"))
                    E_rec.E039(i) = Format$(0, String$(Len(E_rec.E039(i)), "0"))
                Next i
            End If
            
            E_rec.E007(BUYCNT) = oRecordSet.Fields("E007").VALUE
            E_rec.E008(BUYCNT) = oRecordSet.Fields("E008").VALUE
            E_rec.E009(BUYCNT) = oRecordSet.Fields("E009").VALUE
            E_rec.E010(BUYCNT) = oRecordSet.Fields("E010").VALUE
            E_rec.E011(BUYCNT) = oRecordSet.Fields("E011").VALUE
            E_rec.E012(BUYCNT) = oRecordSet.Fields("E012").VALUE
            E_rec.E013(BUYCNT) = oRecordSet.Fields("E013").VALUE
            E_rec.E014(BUYCNT) = oRecordSet.Fields("E014").VALUE
            E_rec.E015(BUYCNT) = oRecordSet.Fields("E015").VALUE
            E_rec.E016(BUYCNT) = oRecordSet.Fields("E016").VALUE
            E_rec.E017(BUYCNT) = oRecordSet.Fields("E017").VALUE
            E_rec.E018(BUYCNT) = oRecordSet.Fields("E018").VALUE
            E_rec.E019(BUYCNT) = Format$(oRecordSet.Fields("E019").VALUE, String$(Len(E_rec.E019(BUYCNT)), "0"))
            E_rec.E020(BUYCNT) = Format$(oRecordSet.Fields("E020").VALUE, String$(Len(E_rec.E020(BUYCNT)), "0"))
            E_rec.E021(BUYCNT) = Format$(oRecordSet.Fields("E021").VALUE, String$(Len(E_rec.E021(BUYCNT)), "0"))
            E_rec.E022(BUYCNT) = Format$(oRecordSet.Fields("E022").VALUE, String$(Len(E_rec.E022(BUYCNT)), "0"))
            E_rec.E023(BUYCNT) = Format$(oRecordSet.Fields("E023").VALUE, String$(Len(E_rec.E023(BUYCNT)), "0"))
            E_rec.E024(BUYCNT) = Format$(oRecordSet.Fields("E024").VALUE, String$(Len(E_rec.E024(BUYCNT)), "0"))
            E_rec.E025(BUYCNT) = Format$(oRecordSet.Fields("E025").VALUE, String$(Len(E_rec.E025(BUYCNT)), "0"))
            E_rec.E026(BUYCNT) = Format$(oRecordSet.Fields("E026").VALUE, String$(Len(E_rec.E026(BUYCNT)), "0"))
            E_rec.E027(BUYCNT) = Format$(oRecordSet.Fields("E027").VALUE, String$(Len(E_rec.E027(BUYCNT)), "0"))
            E_rec.E028(BUYCNT) = Format$(oRecordSet.Fields("E028").VALUE, String$(Len(E_rec.E028(BUYCNT)), "0"))
            E_rec.E029(BUYCNT) = Format$(oRecordSet.Fields("E029").VALUE, String$(Len(E_rec.E029(BUYCNT)), "0"))
            E_rec.E030(BUYCNT) = Format$(oRecordSet.Fields("E030").VALUE, String$(Len(E_rec.E030(BUYCNT)), "0"))
            E_rec.E031(BUYCNT) = Format$(oRecordSet.Fields("E031").VALUE, String$(Len(E_rec.E031(BUYCNT)), "0"))
            E_rec.E032(BUYCNT) = Format$(oRecordSet.Fields("E032").VALUE, String$(Len(E_rec.E032(BUYCNT)), "0"))
            E_rec.E033(BUYCNT) = Format$(oRecordSet.Fields("E033").VALUE, String$(Len(E_rec.E033(BUYCNT)), "0"))
            E_rec.E034(BUYCNT) = Format$(oRecordSet.Fields("E034").VALUE, String$(Len(E_rec.E034(BUYCNT)), "0"))
            E_rec.E035(BUYCNT) = Format$(oRecordSet.Fields("E035").VALUE, String$(Len(E_rec.E035(BUYCNT)), "0"))
            E_rec.E036(BUYCNT) = Format$(oRecordSet.Fields("E036").VALUE, String$(Len(E_rec.E036(BUYCNT)), "0"))
            E_rec.E037(BUYCNT) = Format$(oRecordSet.Fields("E037").VALUE, String$(Len(E_rec.E037(BUYCNT)), "0"))
            E_rec.E038(BUYCNT) = Format$(oRecordSet.Fields("E038").VALUE, String$(Len(E_rec.E038(BUYCNT)), "0"))
            E_rec.E039(BUYCNT) = Format$(oRecordSet.Fields("E039").VALUE, String$(Len(E_rec.E039(BUYCNT)), "0"))
            
            
            oRecordSet.MoveNext
            
            'If BUYCNT = 5 Then    '5개면 인쇄
            If BUYCNT = 5 Or oRecordSet.EOF Then   '5개나 끝이면 인쇄
                E_rec.E172 = Format$(FAMCNT, String$(Len(E_rec.E172), "0"))
                E_rec.E173 = Space$(Len(E_rec.E173))
                '/ E레코드삽입
                Print #1, MDC_SetMod.sStr(E_rec.E001) & MDC_SetMod.sStr(E_rec.E002) & MDC_SetMod.sStr(E_rec.E003) & MDC_SetMod.sStr(E_rec.E004) & MDC_SetMod.sStr(E_rec.E005) & MDC_SetMod.sStr(E_rec.E006) & _
                          MDC_SetMod.sStr(E_rec.E007(1)) & MDC_SetMod.sStr(E_rec.E008(1)) & MDC_SetMod.sStr(E_rec.E009(1)) & MDC_SetMod.sStr(E_rec.E010(1)) & MDC_SetMod.sStr(E_rec.E011(1)) & MDC_SetMod.sStr(E_rec.E012(1)) & MDC_SetMod.sStr(E_rec.E013(1)) & MDC_SetMod.sStr(E_rec.E014(1)) & MDC_SetMod.sStr(E_rec.E015(1)) & MDC_SetMod.sStr(E_rec.E016(1)) & _
                          MDC_SetMod.sStr(E_rec.E017(1)) & MDC_SetMod.sStr(E_rec.E018(1)) & MDC_SetMod.sStr(E_rec.E019(1)) & MDC_SetMod.sStr(E_rec.E020(1)) & MDC_SetMod.sStr(E_rec.E021(1)) & MDC_SetMod.sStr(E_rec.E022(1)) & MDC_SetMod.sStr(E_rec.E023(1)) & MDC_SetMod.sStr(E_rec.E024(1)) & MDC_SetMod.sStr(E_rec.E025(1)) & MDC_SetMod.sStr(E_rec.E026(1)) & _
                          MDC_SetMod.sStr(E_rec.E027(1)) & MDC_SetMod.sStr(E_rec.E028(1)) & MDC_SetMod.sStr(E_rec.E029(1)) & MDC_SetMod.sStr(E_rec.E030(1)) & MDC_SetMod.sStr(E_rec.E031(1)) & MDC_SetMod.sStr(E_rec.E032(1)) & MDC_SetMod.sStr(E_rec.E033(1)) & MDC_SetMod.sStr(E_rec.E034(1)) & MDC_SetMod.sStr(E_rec.E035(1)) & MDC_SetMod.sStr(E_rec.E036(1)) & _
                          MDC_SetMod.sStr(E_rec.E037(1)) & MDC_SetMod.sStr(E_rec.E038(1)) & MDC_SetMod.sStr(E_rec.E039(1)) & MDC_SetMod.sStr(E_rec.E007(2)) & MDC_SetMod.sStr(E_rec.E008(2)) & MDC_SetMod.sStr(E_rec.E009(2)) & MDC_SetMod.sStr(E_rec.E010(2)) & MDC_SetMod.sStr(E_rec.E011(2)) & MDC_SetMod.sStr(E_rec.E012(2)) & MDC_SetMod.sStr(E_rec.E013(2)) & _
                          MDC_SetMod.sStr(E_rec.E014(2)) & MDC_SetMod.sStr(E_rec.E015(2)) & MDC_SetMod.sStr(E_rec.E016(2)) & MDC_SetMod.sStr(E_rec.E017(2)) & MDC_SetMod.sStr(E_rec.E018(2)) & MDC_SetMod.sStr(E_rec.E019(2)) & MDC_SetMod.sStr(E_rec.E020(2)) & MDC_SetMod.sStr(E_rec.E021(2)) & MDC_SetMod.sStr(E_rec.E022(2)) & MDC_SetMod.sStr(E_rec.E023(2)) & _
                          MDC_SetMod.sStr(E_rec.E024(2)) & MDC_SetMod.sStr(E_rec.E025(2)) & MDC_SetMod.sStr(E_rec.E026(2)) & MDC_SetMod.sStr(E_rec.E027(2)) & MDC_SetMod.sStr(E_rec.E028(2)) & MDC_SetMod.sStr(E_rec.E029(2)) & MDC_SetMod.sStr(E_rec.E030(2)) & MDC_SetMod.sStr(E_rec.E031(2)) & MDC_SetMod.sStr(E_rec.E032(2)) & MDC_SetMod.sStr(E_rec.E033(2)) & _
                          MDC_SetMod.sStr(E_rec.E034(2)) & MDC_SetMod.sStr(E_rec.E035(2)) & MDC_SetMod.sStr(E_rec.E036(2)) & MDC_SetMod.sStr(E_rec.E037(2)) & MDC_SetMod.sStr(E_rec.E038(2)) & MDC_SetMod.sStr(E_rec.E039(2)) & MDC_SetMod.sStr(E_rec.E007(3)) & MDC_SetMod.sStr(E_rec.E008(3)) & MDC_SetMod.sStr(E_rec.E009(3)) & MDC_SetMod.sStr(E_rec.E010(3)) & _
                          MDC_SetMod.sStr(E_rec.E011(3)) & MDC_SetMod.sStr(E_rec.E012(3)) & MDC_SetMod.sStr(E_rec.E013(3)) & MDC_SetMod.sStr(E_rec.E014(3)) & MDC_SetMod.sStr(E_rec.E015(3)) & MDC_SetMod.sStr(E_rec.E016(3)) & MDC_SetMod.sStr(E_rec.E017(3)) & MDC_SetMod.sStr(E_rec.E018(3)) & MDC_SetMod.sStr(E_rec.E019(3)) & MDC_SetMod.sStr(E_rec.E020(3)) & _
                          MDC_SetMod.sStr(E_rec.E021(3)) & MDC_SetMod.sStr(E_rec.E022(3)) & MDC_SetMod.sStr(E_rec.E023(3)) & MDC_SetMod.sStr(E_rec.E024(3)) & MDC_SetMod.sStr(E_rec.E025(3)) & MDC_SetMod.sStr(E_rec.E026(3)) & MDC_SetMod.sStr(E_rec.E027(3)) & MDC_SetMod.sStr(E_rec.E028(3)) & MDC_SetMod.sStr(E_rec.E029(3)) & MDC_SetMod.sStr(E_rec.E030(3)) & _
                          MDC_SetMod.sStr(E_rec.E031(3)) & MDC_SetMod.sStr(E_rec.E032(3)) & MDC_SetMod.sStr(E_rec.E033(3)) & MDC_SetMod.sStr(E_rec.E034(3)) & MDC_SetMod.sStr(E_rec.E035(3)) & MDC_SetMod.sStr(E_rec.E036(3)) & MDC_SetMod.sStr(E_rec.E037(3)) & MDC_SetMod.sStr(E_rec.E038(3)) & MDC_SetMod.sStr(E_rec.E039(3)) & MDC_SetMod.sStr(E_rec.E007(4)) & _
                          MDC_SetMod.sStr(E_rec.E008(4)) & MDC_SetMod.sStr(E_rec.E009(4)) & MDC_SetMod.sStr(E_rec.E010(4)) & MDC_SetMod.sStr(E_rec.E011(4)) & MDC_SetMod.sStr(E_rec.E012(4)) & MDC_SetMod.sStr(E_rec.E013(4)) & MDC_SetMod.sStr(E_rec.E014(4)) & MDC_SetMod.sStr(E_rec.E015(4)) & MDC_SetMod.sStr(E_rec.E016(4)) & MDC_SetMod.sStr(E_rec.E017(4)) & _
                          MDC_SetMod.sStr(E_rec.E018(4)) & MDC_SetMod.sStr(E_rec.E019(4)) & MDC_SetMod.sStr(E_rec.E020(4)) & MDC_SetMod.sStr(E_rec.E021(4)) & MDC_SetMod.sStr(E_rec.E022(4)) & MDC_SetMod.sStr(E_rec.E023(4)) & MDC_SetMod.sStr(E_rec.E024(4)) & MDC_SetMod.sStr(E_rec.E025(4)) & MDC_SetMod.sStr(E_rec.E026(4)) & MDC_SetMod.sStr(E_rec.E027(4)) & _
                          MDC_SetMod.sStr(E_rec.E028(4)) & MDC_SetMod.sStr(E_rec.E029(4)) & MDC_SetMod.sStr(E_rec.E030(4)) & MDC_SetMod.sStr(E_rec.E031(4)) & MDC_SetMod.sStr(E_rec.E032(4)) & MDC_SetMod.sStr(E_rec.E033(4)) & MDC_SetMod.sStr(E_rec.E034(4)) & MDC_SetMod.sStr(E_rec.E035(4)) & MDC_SetMod.sStr(E_rec.E036(4)) & MDC_SetMod.sStr(E_rec.E037(4)) & _
                          MDC_SetMod.sStr(E_rec.E038(4)) & MDC_SetMod.sStr(E_rec.E039(4)) & MDC_SetMod.sStr(E_rec.E007(5)) & MDC_SetMod.sStr(E_rec.E008(5)) & MDC_SetMod.sStr(E_rec.E009(5)) & MDC_SetMod.sStr(E_rec.E010(5)) & MDC_SetMod.sStr(E_rec.E011(5)) & MDC_SetMod.sStr(E_rec.E012(5)) & MDC_SetMod.sStr(E_rec.E013(5)) & MDC_SetMod.sStr(E_rec.E014(5)) & _
                          MDC_SetMod.sStr(E_rec.E015(5)) & MDC_SetMod.sStr(E_rec.E016(5)) & MDC_SetMod.sStr(E_rec.E017(5)) & MDC_SetMod.sStr(E_rec.E018(5)) & MDC_SetMod.sStr(E_rec.E019(5)) & MDC_SetMod.sStr(E_rec.E020(5)) & MDC_SetMod.sStr(E_rec.E021(5)) & MDC_SetMod.sStr(E_rec.E022(5)) & MDC_SetMod.sStr(E_rec.E023(5)) & MDC_SetMod.sStr(E_rec.E024(5)) & _
                          MDC_SetMod.sStr(E_rec.E025(5)) & MDC_SetMod.sStr(E_rec.E026(5)) & MDC_SetMod.sStr(E_rec.E027(5)) & MDC_SetMod.sStr(E_rec.E028(5)) & MDC_SetMod.sStr(E_rec.E029(5)) & MDC_SetMod.sStr(E_rec.E030(5)) & MDC_SetMod.sStr(E_rec.E031(5)) & MDC_SetMod.sStr(E_rec.E032(5)) & MDC_SetMod.sStr(E_rec.E033(5)) & MDC_SetMod.sStr(E_rec.E034(5)) & _
                          MDC_SetMod.sStr(E_rec.E035(5)) & MDC_SetMod.sStr(E_rec.E036(5)) & MDC_SetMod.sStr(E_rec.E037(5)) & MDC_SetMod.sStr(E_rec.E038(5)) & MDC_SetMod.sStr(E_rec.E039(5)) & MDC_SetMod.sStr(E_rec.E172) & MDC_SetMod.sStr(E_rec.E173)
                '/ 다음줄넘김
                BUYCNT = 0
                FAMCNT = FAMCNT + 1
            End If
            
            
            
        Loop
'        If BUYCNT > 0 Then   '5개를 다채우지 못하고 종료시 WRITE
'            E_rec.E147 = Format$(FAMCNT, String$(Len(E_rec.E147), "0"))
'            E_rec.E148 = Space$(Len(E_rec.E148))
'            '/ E레코드삽입
'            Print #1, MDC_SetMod.sStr(E_rec.E001) & MDC_SetMod.sStr(E_rec.E002) & MDC_SetMod.sStr(E_rec.E003) & MDC_SetMod.sStr(E_rec.E004) & MDC_SetMod.sStr(E_rec.E005) & MDC_SetMod.sStr(E_rec.E006) & _
'                      MDC_SetMod.sStr(E_rec.E007(1)) & MDC_SetMod.sStr(E_rec.E008(1)) & MDC_SetMod.sStr(E_rec.E009(1)) & MDC_SetMod.sStr(E_rec.E010(1)) & MDC_SetMod.sStr(E_rec.E011(1)) & MDC_SetMod.sStr(E_rec.E012(1)) & MDC_SetMod.sStr(E_rec.E013(1)) & MDC_SetMod.sStr(E_rec.E014(1)) & MDC_SetMod.sStr(E_rec.E015(1)) & MDC_SetMod.sStr(E_rec.E016(1)) & _
'                      MDC_SetMod.sStr(E_rec.E017(1)) & MDC_SetMod.sStr(E_rec.E018(1)) & MDC_SetMod.sStr(E_rec.E019(1)) & MDC_SetMod.sStr(E_rec.E020(1)) & MDC_SetMod.sStr(E_rec.E021(1)) & MDC_SetMod.sStr(E_rec.E022(1)) & MDC_SetMod.sStr(E_rec.E023(1)) & MDC_SetMod.sStr(E_rec.E024(1)) & MDC_SetMod.sStr(E_rec.E025(1)) & MDC_SetMod.sStr(E_rec.E026(1)) & _
'                      MDC_SetMod.sStr(E_rec.E027(1)) & MDC_SetMod.sStr(E_rec.E028(1)) & MDC_SetMod.sStr(E_rec.E029(1)) & MDC_SetMod.sStr(E_rec.E030(1)) & MDC_SetMod.sStr(E_rec.E031(1)) & MDC_SetMod.sStr(E_rec.E032(1)) & MDC_SetMod.sStr(E_rec.E033(1)) & MDC_SetMod.sStr(E_rec.E034(1)) & _
'                      MDC_SetMod.sStr(E_rec.E007(2)) & MDC_SetMod.sStr(E_rec.E008(2)) & MDC_SetMod.sStr(E_rec.E009(2)) & MDC_SetMod.sStr(E_rec.E010(2)) & MDC_SetMod.sStr(E_rec.E011(2)) & MDC_SetMod.sStr(E_rec.E012(2)) & MDC_SetMod.sStr(E_rec.E013(2)) & MDC_SetMod.sStr(E_rec.E014(2)) & MDC_SetMod.sStr(E_rec.E015(2)) & MDC_SetMod.sStr(E_rec.E016(2)) & _
'                      MDC_SetMod.sStr(E_rec.E017(2)) & MDC_SetMod.sStr(E_rec.E018(2)) & MDC_SetMod.sStr(E_rec.E019(2)) & MDC_SetMod.sStr(E_rec.E020(2)) & MDC_SetMod.sStr(E_rec.E021(2)) & MDC_SetMod.sStr(E_rec.E022(2)) & MDC_SetMod.sStr(E_rec.E023(2)) & MDC_SetMod.sStr(E_rec.E024(2)) & MDC_SetMod.sStr(E_rec.E025(2)) & MDC_SetMod.sStr(E_rec.E026(2)) & _
'                      MDC_SetMod.sStr(E_rec.E027(2)) & MDC_SetMod.sStr(E_rec.E028(2)) & MDC_SetMod.sStr(E_rec.E029(2)) & MDC_SetMod.sStr(E_rec.E030(2)) & MDC_SetMod.sStr(E_rec.E031(2)) & MDC_SetMod.sStr(E_rec.E032(2)) & MDC_SetMod.sStr(E_rec.E033(2)) & MDC_SetMod.sStr(E_rec.E034(2)) & _
'                      MDC_SetMod.sStr(E_rec.E007(3)) & MDC_SetMod.sStr(E_rec.E008(3)) & MDC_SetMod.sStr(E_rec.E009(3)) & MDC_SetMod.sStr(E_rec.E010(3)) & MDC_SetMod.sStr(E_rec.E011(3)) & MDC_SetMod.sStr(E_rec.E012(3)) & MDC_SetMod.sStr(E_rec.E013(3)) & MDC_SetMod.sStr(E_rec.E014(3)) & MDC_SetMod.sStr(E_rec.E015(3)) & MDC_SetMod.sStr(E_rec.E016(3)) & _
'                      MDC_SetMod.sStr(E_rec.E017(3)) & MDC_SetMod.sStr(E_rec.E018(3)) & MDC_SetMod.sStr(E_rec.E019(3)) & MDC_SetMod.sStr(E_rec.E020(3)) & MDC_SetMod.sStr(E_rec.E021(3)) & MDC_SetMod.sStr(E_rec.E022(3)) & MDC_SetMod.sStr(E_rec.E023(3)) & MDC_SetMod.sStr(E_rec.E024(3)) & MDC_SetMod.sStr(E_rec.E025(3)) & MDC_SetMod.sStr(E_rec.E026(3)) & _
'                      MDC_SetMod.sStr(E_rec.E027(3)) & MDC_SetMod.sStr(E_rec.E028(3)) & MDC_SetMod.sStr(E_rec.E029(3)) & MDC_SetMod.sStr(E_rec.E030(3)) & MDC_SetMod.sStr(E_rec.E031(3)) & MDC_SetMod.sStr(E_rec.E032(3)) & MDC_SetMod.sStr(E_rec.E033(3)) & MDC_SetMod.sStr(E_rec.E034(3)) & _
'                      MDC_SetMod.sStr(E_rec.E007(4)) & MDC_SetMod.sStr(E_rec.E008(4)) & MDC_SetMod.sStr(E_rec.E009(4)) & MDC_SetMod.sStr(E_rec.E010(4)) & MDC_SetMod.sStr(E_rec.E011(4)) & MDC_SetMod.sStr(E_rec.E012(4)) & MDC_SetMod.sStr(E_rec.E013(4)) & MDC_SetMod.sStr(E_rec.E014(4)) & MDC_SetMod.sStr(E_rec.E015(4)) & MDC_SetMod.sStr(E_rec.E016(4)) & _
'                      MDC_SetMod.sStr(E_rec.E017(4)) & MDC_SetMod.sStr(E_rec.E018(4)) & MDC_SetMod.sStr(E_rec.E019(4)) & MDC_SetMod.sStr(E_rec.E020(4)) & MDC_SetMod.sStr(E_rec.E021(4)) & MDC_SetMod.sStr(E_rec.E022(4)) & MDC_SetMod.sStr(E_rec.E023(4)) & MDC_SetMod.sStr(E_rec.E024(4)) & MDC_SetMod.sStr(E_rec.E025(4)) & MDC_SetMod.sStr(E_rec.E026(4)) & _
'                      MDC_SetMod.sStr(E_rec.E027(4)) & MDC_SetMod.sStr(E_rec.E028(4)) & MDC_SetMod.sStr(E_rec.E029(4)) & MDC_SetMod.sStr(E_rec.E030(4)) & MDC_SetMod.sStr(E_rec.E031(4)) & MDC_SetMod.sStr(E_rec.E032(4)) & MDC_SetMod.sStr(E_rec.E033(4)) & MDC_SetMod.sStr(E_rec.E034(4)) & _
'                      MDC_SetMod.sStr(E_rec.E007(5)) & MDC_SetMod.sStr(E_rec.E008(5)) & MDC_SetMod.sStr(E_rec.E009(5)) & MDC_SetMod.sStr(E_rec.E010(5)) & MDC_SetMod.sStr(E_rec.E011(5)) & MDC_SetMod.sStr(E_rec.E012(5)) & MDC_SetMod.sStr(E_rec.E013(5)) & MDC_SetMod.sStr(E_rec.E014(5)) & MDC_SetMod.sStr(E_rec.E015(5)) & MDC_SetMod.sStr(E_rec.E016(5)) & _
'                      MDC_SetMod.sStr(E_rec.E017(5)) & MDC_SetMod.sStr(E_rec.E018(5)) & MDC_SetMod.sStr(E_rec.E019(5)) & MDC_SetMod.sStr(E_rec.E020(5)) & MDC_SetMod.sStr(E_rec.E021(5)) & MDC_SetMod.sStr(E_rec.E022(5)) & MDC_SetMod.sStr(E_rec.E023(5)) & MDC_SetMod.sStr(E_rec.E024(5)) & MDC_SetMod.sStr(E_rec.E025(5)) & MDC_SetMod.sStr(E_rec.E026(5)) & _
'                      MDC_SetMod.sStr(E_rec.E027(5)) & MDC_SetMod.sStr(E_rec.E028(5)) & MDC_SetMod.sStr(E_rec.E029(5)) & MDC_SetMod.sStr(E_rec.E030(5)) & MDC_SetMod.sStr(E_rec.E031(5)) & MDC_SetMod.sStr(E_rec.E032(5)) & MDC_SetMod.sStr(E_rec.E033(5)) & MDC_SetMod.sStr(E_rec.E034(5)) & _
'                      MDC_SetMod.sStr(E_rec.E147) & MDC_SetMod.sStr(E_rec.E148)
'        End If

    Else
    
      ErrNum = 1
    
    End If
    
    If CheckE = False Then
        File_Create_E_record = True
    Else
        File_Create_E_record = False
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
    File_Create_E_record = False
End Function

Private Function File_Create_F_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckF      As String
    Dim SAVCNT      As Integer
    Dim iRow        As Long
    
    CheckF = False  '/체크필요유무
    ErrNum = 0
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ F_RECORE QUERY
    sQry = "EXEC PH_PY980_F '" & C_SAUP & "', '" & C_YYYY & "', '" & C_SABUN & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount > 0 Then
        
        F_rec.F001 = oRecordSet.Fields("F001").VALUE
        F_rec.F002 = oRecordSet.Fields("F002").VALUE
        F_rec.F003 = oRecordSet.Fields("F003").VALUE
        F_rec.F004 = Format$(C_rec.C004, String$(Len(F_rec.F004), "0"))  '/ C레코드의 일련번호
        F_rec.F005 = oRecordSet.Fields("F005").VALUE
        F_rec.F006 = oRecordSet.Fields("F006").VALUE
         
        SAVCNT = 0
        Do Until oRecordSet.EOF

            SAVCNT = SAVCNT + 1   '/ 레코드번호
            '/초기화
            If SAVCNT = 1 Then   '첫번째때 초기화
                For iRow = 1 To 15
                    F_rec.F007(iRow) = Space$(Len(F_rec.F007(iRow)))
                    F_rec.F008(iRow) = Space$(Len(F_rec.F008(iRow)))
                    F_rec.F009(iRow) = Space$(Len(F_rec.F009(iRow)))
                    F_rec.F010(iRow) = Space$(Len(F_rec.F010(iRow)))
                    
                    F_rec.F011(iRow) = Format$(0, String$(Len(F_rec.F011(iRow)), "0"))
                    F_rec.F012(iRow) = Format$(0, String$(Len(F_rec.F012(iRow)), "0"))
                Next iRow
            End If
            
            F_rec.F007(SAVCNT) = oRecordSet.Fields("F007").VALUE
            F_rec.F008(SAVCNT) = oRecordSet.Fields("F008").VALUE
            F_rec.F009(SAVCNT) = oRecordSet.Fields("F009").VALUE
            F_rec.F010(SAVCNT) = oRecordSet.Fields("F010").VALUE
            F_rec.F011(SAVCNT) = Format$(oRecordSet.Fields("F011").VALUE, String$(Len(F_rec.F011(SAVCNT)), "0"))
            F_rec.F012(SAVCNT) = Format$(oRecordSet.Fields("F012").VALUE, String$(Len(F_rec.F012(SAVCNT)), "0"))

            oRecordSet.MoveNext
            
            
            If SAVCNT = 15 Or oRecordSet.EOF Then '15개나 끝이면 인쇄
                F_rec.F097 = Space$(Len(F_rec.F097))
                '/ F레코드삽입
                Print #1, MDC_SetMod.sStr(F_rec.F001) & MDC_SetMod.sStr(F_rec.F002) & MDC_SetMod.sStr(F_rec.F003) & MDC_SetMod.sStr(F_rec.F004) & MDC_SetMod.sStr(F_rec.F005) & MDC_SetMod.sStr(F_rec.F006) & _
                          MDC_SetMod.sStr(F_rec.F007(1)) & MDC_SetMod.sStr(F_rec.F008(1)) & MDC_SetMod.sStr(F_rec.F009(1)) & MDC_SetMod.sStr(F_rec.F010(1)) & MDC_SetMod.sStr(F_rec.F011(1)) & MDC_SetMod.sStr(F_rec.F012(1)) & _
                          MDC_SetMod.sStr(F_rec.F007(2)) & MDC_SetMod.sStr(F_rec.F008(2)) & MDC_SetMod.sStr(F_rec.F009(2)) & MDC_SetMod.sStr(F_rec.F010(2)) & MDC_SetMod.sStr(F_rec.F011(2)) & MDC_SetMod.sStr(F_rec.F012(2)) & _
                          MDC_SetMod.sStr(F_rec.F007(3)) & MDC_SetMod.sStr(F_rec.F008(3)) & MDC_SetMod.sStr(F_rec.F009(3)) & MDC_SetMod.sStr(F_rec.F010(3)) & MDC_SetMod.sStr(F_rec.F011(3)) & MDC_SetMod.sStr(F_rec.F012(3)) & _
                          MDC_SetMod.sStr(F_rec.F007(4)) & MDC_SetMod.sStr(F_rec.F008(4)) & MDC_SetMod.sStr(F_rec.F009(4)) & MDC_SetMod.sStr(F_rec.F010(4)) & MDC_SetMod.sStr(F_rec.F011(4)) & MDC_SetMod.sStr(F_rec.F012(4)) & _
                          MDC_SetMod.sStr(F_rec.F007(5)) & MDC_SetMod.sStr(F_rec.F008(5)) & MDC_SetMod.sStr(F_rec.F009(5)) & MDC_SetMod.sStr(F_rec.F010(5)) & MDC_SetMod.sStr(F_rec.F011(5)) & MDC_SetMod.sStr(F_rec.F012(5)) & _
                          MDC_SetMod.sStr(F_rec.F007(6)) & MDC_SetMod.sStr(F_rec.F008(6)) & MDC_SetMod.sStr(F_rec.F009(6)) & MDC_SetMod.sStr(F_rec.F010(6)) & MDC_SetMod.sStr(F_rec.F011(6)) & MDC_SetMod.sStr(F_rec.F012(6)) & _
                          MDC_SetMod.sStr(F_rec.F007(7)) & MDC_SetMod.sStr(F_rec.F008(7)) & MDC_SetMod.sStr(F_rec.F009(7)) & MDC_SetMod.sStr(F_rec.F010(7)) & MDC_SetMod.sStr(F_rec.F011(7)) & MDC_SetMod.sStr(F_rec.F012(7)) & _
                          MDC_SetMod.sStr(F_rec.F007(8)) & MDC_SetMod.sStr(F_rec.F008(8)) & MDC_SetMod.sStr(F_rec.F009(8)) & MDC_SetMod.sStr(F_rec.F010(8)) & MDC_SetMod.sStr(F_rec.F011(8)) & MDC_SetMod.sStr(F_rec.F012(8)) & _
                          MDC_SetMod.sStr(F_rec.F007(9)) & MDC_SetMod.sStr(F_rec.F008(9)) & MDC_SetMod.sStr(F_rec.F009(9)) & MDC_SetMod.sStr(F_rec.F010(9)) & MDC_SetMod.sStr(F_rec.F011(9)) & MDC_SetMod.sStr(F_rec.F012(9)) & _
                          MDC_SetMod.sStr(F_rec.F007(10)) & MDC_SetMod.sStr(F_rec.F008(10)) & MDC_SetMod.sStr(F_rec.F009(10)) & MDC_SetMod.sStr(F_rec.F010(10)) & MDC_SetMod.sStr(F_rec.F011(10)) & MDC_SetMod.sStr(F_rec.F012(10)) & _
                          MDC_SetMod.sStr(F_rec.F007(11)) & MDC_SetMod.sStr(F_rec.F008(11)) & MDC_SetMod.sStr(F_rec.F009(11)) & MDC_SetMod.sStr(F_rec.F010(11)) & MDC_SetMod.sStr(F_rec.F011(11)) & MDC_SetMod.sStr(F_rec.F012(11)) & _
                          MDC_SetMod.sStr(F_rec.F007(12)) & MDC_SetMod.sStr(F_rec.F008(12)) & MDC_SetMod.sStr(F_rec.F009(12)) & MDC_SetMod.sStr(F_rec.F010(12)) & MDC_SetMod.sStr(F_rec.F011(12)) & MDC_SetMod.sStr(F_rec.F012(12)) & _
                          MDC_SetMod.sStr(F_rec.F007(13)) & MDC_SetMod.sStr(F_rec.F008(13)) & MDC_SetMod.sStr(F_rec.F009(13)) & MDC_SetMod.sStr(F_rec.F010(13)) & MDC_SetMod.sStr(F_rec.F011(13)) & MDC_SetMod.sStr(F_rec.F012(13)) & _
                          MDC_SetMod.sStr(F_rec.F007(14)) & MDC_SetMod.sStr(F_rec.F008(14)) & MDC_SetMod.sStr(F_rec.F009(14)) & MDC_SetMod.sStr(F_rec.F010(14)) & MDC_SetMod.sStr(F_rec.F011(14)) & MDC_SetMod.sStr(F_rec.F012(14)) & _
                          MDC_SetMod.sStr(F_rec.F007(15)) & MDC_SetMod.sStr(F_rec.F008(15)) & MDC_SetMod.sStr(F_rec.F009(15)) & MDC_SetMod.sStr(F_rec.F010(15)) & MDC_SetMod.sStr(F_rec.F011(15)) & MDC_SetMod.sStr(F_rec.F012(15)) & _
                          MDC_SetMod.sStr(F_rec.F097)
                SAVCNT = 0
            End If
        Loop
    End If

        
    If CheckF = False Then
        File_Create_F_record = True
    Else
        File_Create_F_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    Call Matrix_AddRow("F레코드오류: " & Err.Source & " " & Err.Description, False)
    File_Create_F_record = False
End Function

Private Function File_Create_G_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckG      As String
    
    CheckG = False  '/체크필요유무
    ErrNum = 0
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ G_RECORE QUERY
    sQry = "EXEC PH_PY980_G '" & C_SAUP & "', '" & C_YYYY & "', '" & C_SABUN & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount > 0 Then
        
        G_rec.G001 = oRecordSet.Fields("G001").VALUE
        G_rec.G002 = oRecordSet.Fields("G002").VALUE
        G_rec.G003 = oRecordSet.Fields("G003").VALUE
        G_rec.G004 = Format$(C_rec.C004, String$(Len(F_rec.F004), "0"))  '/ C레코드의 일련번호
        G_rec.G005 = oRecordSet.Fields("G005").VALUE
        G_rec.G006 = oRecordSet.Fields("G006").VALUE
        
        G_rec.G007 = oRecordSet.Fields("G007").VALUE
        G_rec.G008 = oRecordSet.Fields("G008").VALUE
        G_rec.G009 = oRecordSet.Fields("G009").VALUE
        G_rec.G010 = Format$(oRecordSet.Fields("G010").VALUE, String$(Len(G_rec.G010), "0"))
        G_rec.G011 = oRecordSet.Fields("G011").VALUE
        G_rec.G012 = Format$(oRecordSet.Fields("G012").VALUE, String$(Len(G_rec.G012), "0"))
        G_rec.G013 = Format$(oRecordSet.Fields("G013").VALUE, String$(Len(G_rec.G013), "0"))
        G_rec.G014 = Format$(oRecordSet.Fields("G014").VALUE, String$(Len(G_rec.G014), "0"))
        G_rec.G015 = Format$(oRecordSet.Fields("G015").VALUE, String$(Len(G_rec.G015), "0"))
        G_rec.G016 = oRecordSet.Fields("G016").VALUE
        G_rec.G017 = oRecordSet.Fields("G017").VALUE
        G_rec.G018 = Format$(oRecordSet.Fields("G018").VALUE, String$(Len(G_rec.G018), "0"))
        G_rec.G019 = Format$(oRecordSet.Fields("G019").VALUE, String$(Len(G_rec.G019), "0"))
        G_rec.G020 = Format$(oRecordSet.Fields("G020").VALUE, String$(Len(G_rec.G020), "0"))
        G_rec.G021 = Format$(oRecordSet.Fields("G021").VALUE, String$(Len(G_rec.G021), "0"))
        G_rec.G022 = Format$(oRecordSet.Fields("G022").VALUE, String$(Len(G_rec.G022), "0"))
        G_rec.G023 = Format$(oRecordSet.Fields("G023").VALUE, String$(Len(G_rec.G023), "0"))
        G_rec.G024 = Format$(oRecordSet.Fields("G024").VALUE, String$(Len(G_rec.G024), "0"))
        G_rec.G025 = oRecordSet.Fields("G025").VALUE
        G_rec.G026 = oRecordSet.Fields("G026").VALUE
        G_rec.G027 = oRecordSet.Fields("G027").VALUE
        G_rec.G028 = Format$(oRecordSet.Fields("G028").VALUE, String$(Len(G_rec.G028), "0"))
        G_rec.G029 = oRecordSet.Fields("G029").VALUE
        G_rec.G030 = Format$(oRecordSet.Fields("G030").VALUE, String$(Len(G_rec.G030), "0"))
        G_rec.G031 = Format$(oRecordSet.Fields("G031").VALUE, String$(Len(G_rec.G031), "0"))
        G_rec.G032 = Format$(oRecordSet.Fields("G032").VALUE, String$(Len(G_rec.G032), "0"))
        
        G_rec.G033 = oRecordSet.Fields("G033").VALUE
        G_rec.G034 = oRecordSet.Fields("G034").VALUE
        G_rec.G035 = oRecordSet.Fields("G035").VALUE
        G_rec.G036 = Format$(oRecordSet.Fields("G036").VALUE, String$(Len(G_rec.G036), "0"))
        G_rec.G037 = oRecordSet.Fields("G037").VALUE
        G_rec.G038 = Format$(oRecordSet.Fields("G038").VALUE, String$(Len(G_rec.G038), "0"))
        G_rec.G039 = Format$(oRecordSet.Fields("G039").VALUE, String$(Len(G_rec.G039), "0"))
        G_rec.G040 = Format$(oRecordSet.Fields("G040").VALUE, String$(Len(G_rec.G040), "0"))
        G_rec.G041 = Format$(oRecordSet.Fields("G041").VALUE, String$(Len(G_rec.G041), "0"))
        G_rec.G042 = oRecordSet.Fields("G042").VALUE
        G_rec.G043 = oRecordSet.Fields("G043").VALUE
        G_rec.G044 = Format$(oRecordSet.Fields("G044").VALUE, String$(Len(G_rec.G044), "0"))
        G_rec.G045 = Format$(oRecordSet.Fields("G045").VALUE, String$(Len(G_rec.G045), "0"))
        G_rec.G046 = Format$(oRecordSet.Fields("G046").VALUE, String$(Len(G_rec.G046), "0"))
        G_rec.G047 = Format$(oRecordSet.Fields("G047").VALUE, String$(Len(G_rec.G047), "0"))
        G_rec.G048 = Format$(oRecordSet.Fields("G048").VALUE, String$(Len(G_rec.G048), "0"))
        G_rec.G049 = Format$(oRecordSet.Fields("G049").VALUE, String$(Len(G_rec.G049), "0"))
        G_rec.G050 = Format$(oRecordSet.Fields("G050").VALUE, String$(Len(G_rec.G050), "0"))
        G_rec.G051 = oRecordSet.Fields("G051").VALUE
        G_rec.G052 = oRecordSet.Fields("G052").VALUE
        G_rec.G053 = oRecordSet.Fields("G053").VALUE
        G_rec.G054 = Format$(oRecordSet.Fields("G054").VALUE, String$(Len(G_rec.G054), "0"))
        G_rec.G055 = oRecordSet.Fields("G055").VALUE
        G_rec.G056 = Format$(oRecordSet.Fields("G056").VALUE, String$(Len(G_rec.G056), "0"))
        G_rec.G057 = Format$(oRecordSet.Fields("G057").VALUE, String$(Len(G_rec.G057), "0"))
        G_rec.G058 = Format$(oRecordSet.Fields("G058").VALUE, String$(Len(G_rec.G058), "0"))
        
        G_rec.G059 = oRecordSet.Fields("G059").VALUE
        G_rec.G060 = oRecordSet.Fields("G060").VALUE
        G_rec.G061 = oRecordSet.Fields("G061").VALUE
        G_rec.G062 = Format$(oRecordSet.Fields("G062").VALUE, String$(Len(G_rec.G062), "0"))
        G_rec.G063 = oRecordSet.Fields("G063").VALUE
        G_rec.G064 = Format$(oRecordSet.Fields("G064").VALUE, String$(Len(G_rec.G064), "0"))
        G_rec.G065 = Format$(oRecordSet.Fields("G065").VALUE, String$(Len(G_rec.G065), "0"))
        G_rec.G066 = Format$(oRecordSet.Fields("G066").VALUE, String$(Len(G_rec.G066), "0"))
        G_rec.G067 = Format$(oRecordSet.Fields("G067").VALUE, String$(Len(G_rec.G067), "0"))
        G_rec.G068 = oRecordSet.Fields("G068").VALUE
        G_rec.G069 = oRecordSet.Fields("G069").VALUE
        G_rec.G070 = Format$(oRecordSet.Fields("G070").VALUE, String$(Len(G_rec.G070), "0"))
        G_rec.G071 = Format$(oRecordSet.Fields("G071").VALUE, String$(Len(G_rec.G071), "0"))
        G_rec.G072 = Format$(oRecordSet.Fields("G072").VALUE, String$(Len(G_rec.G072), "0"))
        G_rec.G073 = Format$(oRecordSet.Fields("G073").VALUE, String$(Len(G_rec.G073), "0"))
        G_rec.G074 = Format$(oRecordSet.Fields("G074").VALUE, String$(Len(G_rec.G074), "0"))
        G_rec.G075 = Format$(oRecordSet.Fields("G075").VALUE, String$(Len(G_rec.G075), "0"))
        G_rec.G076 = Format$(oRecordSet.Fields("G076").VALUE, String$(Len(G_rec.G076), "0"))
        G_rec.G077 = oRecordSet.Fields("G077").VALUE
        G_rec.G078 = oRecordSet.Fields("G078").VALUE
        G_rec.G079 = oRecordSet.Fields("G079").VALUE
        G_rec.G080 = Format$(oRecordSet.Fields("G080").VALUE, String$(Len(G_rec.G080), "0"))
        G_rec.G081 = oRecordSet.Fields("G081").VALUE
        G_rec.G082 = Format$(oRecordSet.Fields("G082").VALUE, String$(Len(G_rec.G082), "0"))
        G_rec.G083 = Format$(oRecordSet.Fields("G083").VALUE, String$(Len(G_rec.G083), "0"))
        G_rec.G084 = Format$(oRecordSet.Fields("G084").VALUE, String$(Len(G_rec.G084), "0"))
        G_rec.G085 = oRecordSet.Fields("G085").VALUE
        G_rec.G086 = oRecordSet.Fields("G086").VALUE
        
         
        Print #1, MDC_SetMod.sStr(G_rec.G001) & MDC_SetMod.sStr(G_rec.G002) & MDC_SetMod.sStr(G_rec.G003) & MDC_SetMod.sStr(G_rec.G004) & MDC_SetMod.sStr(G_rec.G005) & MDC_SetMod.sStr(G_rec.G006) & MDC_SetMod.sStr(G_rec.G007) & MDC_SetMod.sStr(G_rec.G008) & MDC_SetMod.sStr(G_rec.G009) & MDC_SetMod.sStr(G_rec.G010) & _
                  MDC_SetMod.sStr(G_rec.G011) & MDC_SetMod.sStr(G_rec.G012) & MDC_SetMod.sStr(G_rec.G013) & MDC_SetMod.sStr(G_rec.G014) & MDC_SetMod.sStr(G_rec.G015) & MDC_SetMod.sStr(G_rec.G016) & MDC_SetMod.sStr(G_rec.G017) & MDC_SetMod.sStr(G_rec.G018) & MDC_SetMod.sStr(G_rec.G019) & MDC_SetMod.sStr(G_rec.G020) & _
                  MDC_SetMod.sStr(G_rec.G021) & MDC_SetMod.sStr(G_rec.G022) & MDC_SetMod.sStr(G_rec.G023) & MDC_SetMod.sStr(G_rec.G024) & MDC_SetMod.sStr(G_rec.G025) & MDC_SetMod.sStr(G_rec.G026) & MDC_SetMod.sStr(G_rec.G027) & MDC_SetMod.sStr(G_rec.G028) & MDC_SetMod.sStr(G_rec.G029) & MDC_SetMod.sStr(G_rec.G030) & _
                  MDC_SetMod.sStr(G_rec.G031) & MDC_SetMod.sStr(G_rec.G032) & MDC_SetMod.sStr(G_rec.G033) & MDC_SetMod.sStr(G_rec.G034) & MDC_SetMod.sStr(G_rec.G035) & MDC_SetMod.sStr(G_rec.G036) & MDC_SetMod.sStr(G_rec.G037) & MDC_SetMod.sStr(G_rec.G038) & MDC_SetMod.sStr(G_rec.G039) & MDC_SetMod.sStr(G_rec.G040) & _
                  MDC_SetMod.sStr(G_rec.G041) & MDC_SetMod.sStr(G_rec.G042) & MDC_SetMod.sStr(G_rec.G043) & MDC_SetMod.sStr(G_rec.G044) & MDC_SetMod.sStr(G_rec.G045) & MDC_SetMod.sStr(G_rec.G046) & MDC_SetMod.sStr(G_rec.G047) & MDC_SetMod.sStr(G_rec.G048) & MDC_SetMod.sStr(G_rec.G049) & MDC_SetMod.sStr(G_rec.G050) & _
                  MDC_SetMod.sStr(G_rec.G051) & MDC_SetMod.sStr(G_rec.G052) & MDC_SetMod.sStr(G_rec.G053) & MDC_SetMod.sStr(G_rec.G054) & MDC_SetMod.sStr(G_rec.G055) & MDC_SetMod.sStr(G_rec.G056) & MDC_SetMod.sStr(G_rec.G057) & MDC_SetMod.sStr(G_rec.G058) & MDC_SetMod.sStr(G_rec.G059) & MDC_SetMod.sStr(G_rec.G060) & _
                  MDC_SetMod.sStr(G_rec.G061) & MDC_SetMod.sStr(G_rec.G062) & MDC_SetMod.sStr(G_rec.G063) & MDC_SetMod.sStr(G_rec.G064) & MDC_SetMod.sStr(G_rec.G065) & MDC_SetMod.sStr(G_rec.G066) & MDC_SetMod.sStr(G_rec.G067) & MDC_SetMod.sStr(G_rec.G068) & MDC_SetMod.sStr(G_rec.G069) & MDC_SetMod.sStr(G_rec.G070) & _
                  MDC_SetMod.sStr(G_rec.G071) & MDC_SetMod.sStr(G_rec.G072) & MDC_SetMod.sStr(G_rec.G073) & MDC_SetMod.sStr(G_rec.G074) & MDC_SetMod.sStr(G_rec.G075) & MDC_SetMod.sStr(G_rec.G076) & MDC_SetMod.sStr(G_rec.G077) & MDC_SetMod.sStr(G_rec.G078) & MDC_SetMod.sStr(G_rec.G079) & MDC_SetMod.sStr(G_rec.G080) & _
                  MDC_SetMod.sStr(G_rec.G081) & MDC_SetMod.sStr(G_rec.G082) & MDC_SetMod.sStr(G_rec.G083) & MDC_SetMod.sStr(G_rec.G084) & MDC_SetMod.sStr(G_rec.G085) & MDC_SetMod.sStr(G_rec.G086)
    End If

        
    If CheckG = False Then
        File_Create_G_record = True
    Else
        File_Create_G_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    Call Matrix_AddRow("G레코드오류: " & Err.Source & " " & Err.Description, False)
    File_Create_G_record = False
End Function

Private Function File_Create_H_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckH      As String
    Dim JONCNT      As Integer
    Dim HCount      As Integer
    
    CheckH = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ H_RECORE QUERY
    sQry = "EXEC PH_PY980_H '" & C_SAUP & "', '" & C_YYYY & "', '" & C_SABUN & "'"
    oRecordSet.DoQuery sQry
    
    HCount = 0
    
    If oRecordSet.RecordCount > 0 Then
       Do Until oRecordSet.EOF
    
        'H RECORD MOVE
        
        HCount = HCount + 1
        
        H_rec.H001 = oRecordSet.Fields("H001").VALUE
        H_rec.H002 = oRecordSet.Fields("H002").VALUE
        H_rec.H003 = oRecordSet.Fields("H003").VALUE
        H_rec.H004 = Format$(C_rec.C004, String$(Len(H_rec.H004), "0"))  '/ 일련번호
        H_rec.H005 = oRecordSet.Fields("H005").VALUE
        H_rec.H006 = oRecordSet.Fields("H006").VALUE
        H_rec.H007 = oRecordSet.Fields("H007").VALUE
        H_rec.H008 = oRecordSet.Fields("H008").VALUE
        H_rec.H009 = oRecordSet.Fields("H009").VALUE
        H_rec.H010 = oRecordSet.Fields("H010").VALUE
        
        H_rec.H011 = Format$(oRecordSet.Fields("H011").VALUE, String$(Len(H_rec.H011), "0"))
        H_rec.H012 = Format$(oRecordSet.Fields("H012").VALUE, String$(Len(H_rec.H012), "0"))
        H_rec.H013 = Format$(oRecordSet.Fields("H013").VALUE, String$(Len(H_rec.H013), "0"))
        H_rec.H014 = Format$(oRecordSet.Fields("H014").VALUE, String$(Len(H_rec.H014), "0"))
        H_rec.H015 = Format$(oRecordSet.Fields("H015").VALUE, String$(Len(H_rec.H015), "0"))
        H_rec.H016 = Format$(oRecordSet.Fields("H016").VALUE, String$(Len(H_rec.H016), "0"))
        H_rec.H017 = Format$(oRecordSet.Fields("H017").VALUE, String$(Len(H_rec.H017), "0"))
        H_rec.H018 = Format$(HCount, String$(Len(H_rec.H018), "0"))  '/ 일련번호
        H_rec.H019 = oRecordSet.Fields("H019").VALUE
        
        Print #1, MDC_SetMod.sStr(H_rec.H001) & MDC_SetMod.sStr(H_rec.H002) & MDC_SetMod.sStr(H_rec.H003) & MDC_SetMod.sStr(H_rec.H004) & MDC_SetMod.sStr(H_rec.H005) & MDC_SetMod.sStr(H_rec.H006) & MDC_SetMod.sStr(H_rec.H007) & MDC_SetMod.sStr(H_rec.H008) & MDC_SetMod.sStr(H_rec.H009) & MDC_SetMod.sStr(H_rec.H010) & _
                  MDC_SetMod.sStr(H_rec.H011) & MDC_SetMod.sStr(H_rec.H012) & MDC_SetMod.sStr(H_rec.H013) & MDC_SetMod.sStr(H_rec.H014) & MDC_SetMod.sStr(H_rec.H015) & MDC_SetMod.sStr(H_rec.H016) & MDC_SetMod.sStr(H_rec.H017) & MDC_SetMod.sStr(H_rec.H018) & MDC_SetMod.sStr(H_rec.H019)
        oRecordSet.MoveNext
      Loop
      
    End If
    
    If CheckH = False Then
        File_Create_H_record = True
    Else
        File_Create_H_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Call Matrix_AddRow("H레코드오류: " & Err.Description, False)
    File_Create_H_record = False
End Function

Private Function File_Create_I_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckI      As String
    Dim JONCNT      As Integer
    Dim ICount      As Integer
    
    CheckI = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ I_RECORE QUERY
    sQry = "EXEC PH_PY980_I '" & C_SAUP & "', '" & C_YYYY & "', '" & C_SABUN & "'"
    oRecordSet.DoQuery sQry
    
    ICount = 0
    
    If oRecordSet.RecordCount > 0 Then
       Do Until oRecordSet.EOF
    
        'I RECORD MOVE
        
        ICount = ICount + 1
        
        I_rec.I001 = oRecordSet.Fields("I001").VALUE
        I_rec.I002 = oRecordSet.Fields("I002").VALUE
        I_rec.I003 = oRecordSet.Fields("I003").VALUE
        I_rec.I004 = Format$(C_rec.C004, String$(Len(I_rec.I004), "0"))  '/ C레코드의 일련번호
        I_rec.I005 = oRecordSet.Fields("I005").VALUE
        I_rec.I006 = oRecordSet.Fields("I006").VALUE
        I_rec.I007 = oRecordSet.Fields("I007").VALUE
        I_rec.I008 = oRecordSet.Fields("I008").VALUE
        I_rec.I009 = oRecordSet.Fields("I009").VALUE
        I_rec.I010 = oRecordSet.Fields("I010").VALUE
        I_rec.I011 = oRecordSet.Fields("I011").VALUE
        I_rec.I012 = oRecordSet.Fields("I012").VALUE
        I_rec.I013 = oRecordSet.Fields("I013").VALUE
        I_rec.I014 = Format$(oRecordSet.Fields("I014").VALUE, String$(Len(I_rec.I014), "0"))
        I_rec.I015 = Format$(oRecordSet.Fields("I015").VALUE, String$(Len(I_rec.I015), "0"))
        I_rec.I016 = Format$(oRecordSet.Fields("I016").VALUE, String$(Len(I_rec.I016), "0"))
        I_rec.I017 = Format$(oRecordSet.Fields("I017").VALUE, String$(Len(I_rec.I017), "0"))
        I_rec.I018 = Format$(oRecordSet.Fields("I018").VALUE, String$(Len(I_rec.I018), "0"))
        I_rec.I019 = Format$(ICount, String$(Len(I_rec.I019), "0"))  '/일련번호
        I_rec.I020 = oRecordSet.Fields("I020").VALUE
        
        Print #1, MDC_SetMod.sStr(I_rec.I001) & MDC_SetMod.sStr(I_rec.I002) & MDC_SetMod.sStr(I_rec.I003) & MDC_SetMod.sStr(I_rec.I004) & MDC_SetMod.sStr(I_rec.I005) & MDC_SetMod.sStr(I_rec.I006) & MDC_SetMod.sStr(I_rec.I007) & MDC_SetMod.sStr(I_rec.I008) & MDC_SetMod.sStr(I_rec.I009) & MDC_SetMod.sStr(I_rec.I010) & _
                  MDC_SetMod.sStr(I_rec.I011) & MDC_SetMod.sStr(I_rec.I012) & MDC_SetMod.sStr(I_rec.I013) & MDC_SetMod.sStr(I_rec.I014) & MDC_SetMod.sStr(I_rec.I015) & MDC_SetMod.sStr(I_rec.I016) & MDC_SetMod.sStr(I_rec.I017) & MDC_SetMod.sStr(I_rec.I018) & MDC_SetMod.sStr(I_rec.I019) & MDC_SetMod.sStr(I_rec.I020)
     
        oRecordSet.MoveNext
       Loop
    End If
    If CheckI = False Then
        File_Create_I_record = True
    Else
        File_Create_I_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Call Matrix_AddRow("I레코드오류: " & Err.Description, False)
    File_Create_I_record = False
End Function

Private Sub Matrix_AddRow(ByVal MatrixMsg As String, Optional Insert_YN As Boolean, Optional MatrixErr As Boolean)
    If MatrixErr = True Then
        oForm.DataSources.UserDataSources("Col0").VALUE = "??"
    Else
        oForm.DataSources.UserDataSources("Col0").VALUE = ""
    End If
    oForm.DataSources.UserDataSources("Col1").VALUE = MatrixMsg
    If Insert_YN = True Then
        oMat1.AddRow
        MaxRow = MaxRow + 1
    End If
    oMat1.SetLineData MaxRow
End Sub
'화면변수 CHECK
Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel
    Dim ErrNum      As Integer
    
    ErrNum = 0
    '/ 필수Check
    If oForm.Items("HtaxID").Specific.VALUE = "" Then
        ErrNum = 1
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("TeamName").Specific.VALUE = "" Then
        ErrNum = 2
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("Dname").Specific.VALUE = "" Then
        ErrNum = 3
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("Dtel").Specific.VALUE = "" Then
        ErrNum = 4
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("DocDate").Specific.VALUE = "" Then
        ErrNum = 5
        GoTo HeaderSpaceLineDel
    End If
    
    HeaderSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
HeaderSpaceLineDel:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "홈텍스ID(5자리이상)를 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "담당자부서는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "담당자성명은 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "담당자전화번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "제출일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    
    HeaderSpaceLineDel = False
End Function


