Attribute VB_Name = "MDC_SetFilter"
Option Explicit

Public Function Execute()
    Dim oFilters As SAPbouiCOM.EventFilters
    Dim oFilter  As SAPbouiCOM.EventFilter

    Set oFilters = New SAPbouiCOM.EventFilters
'
    Call ITEM_PRESSED(oFilter, oFilters)                '1
    Call KEY_DOWN(oFilter, oFilters)                    '2
    Call GOT_FOCUS(oFilter, oFilters)                   '3
    Call LOST_FOCUS(oFilter, oFilters)                  '4
    Call COMBO_SELECT(oFilter, oFilters)                '5
    Call CLICK(oFilter, oFilters)                       '6
    Call DOUBLE_CLICK(oFilter, oFilters)                '7
    Call MATRIX_LINK_PRESSED(oFilter, oFilters)         '8
'    Call MATRIX_COLLAPSE_PRESSED(oFilter, oFilters)     '9
    Call VALIDATE(oFilter, oFilters)                    '10
    Call MATRIX_LOAD(oFilter, oFilters)                 '11
'    Call DATASOURCE_LOAD(oFilter, oFilters)             '12
    Call Form_Load(oFilter, oFilters)                   '16
    Call FORM_UNLOAD(oFilter, oFilters)                 '17
'    Call FORM_ACTIVATE(oFilter, oFilters)               '18
'    Call FORM_DEACTIVATE(oFilter, oFilters)             '19
'    Call FORM_CLOSE(oFilter, oFilters)                  '20
    Call Form_Resize(oFilter, oFilters)                 '21
'    Call FORM_KEY_DOWN(oFilter, oFilters)               '22
'    Call FORM_MENU_HILIGHT(oFilter, oFilters)           '23
'    Call PRINT(oFilter, oFilters)                       '24
'    Call PRINT_DATA(oFilter, oFilters)                  '25
    Call CHOOSE_FROM_LIST(oFilter, oFilters)            '27
    Call RIGHT_CLICK(oFilter, oFilters)                 '28
    Call MENU_CLICK(oFilter, oFilters)                  '32
    Call FORM_DATA_ADD(oFilter, oFilters)               '33
    Call FORM_DATA_UPDATE(oFilter, oFilters)            '34
'    Call FORM_DATA_DELETE(oFilter, oFilters)            '35
    Call FORM_DATA_LOAD(oFilter, oFilters)              '36

    '// Setting the application with the EventFilters object
    Sbo_Application.SetFilter oFilters
    
    Set oFilter = Nothing
    Set oFilters = Nothing
    
End Function

Private Sub ITEM_PRESSED(ByRef oFilter As SAPbouiCOM.EventFilter, _
                         ByRef oFilters As SAPbouiCOM.EventFilters)  '1
    Set oFilter = oFilters.Add(et_ITEM_PRESSED)
    
    
    '//System Form Type
    
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    oFilter.AddEx "PH_PY000"            '사용자권한관리
    
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터 등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    oFilter.AddEx "PH_PY004"            '근무조편성등록
    oFilter.AddEx "PH_PY005"            '사업장정보등록
    oFilter.AddEx "PH_PY006"            '승호작업등록
    oFilter.AddEx "PH_PY007"            '유류단가등록
    oFilter.AddEx "PH_PY008"            '일근태등록
    oFilter.AddEx "PH_PY009"            '기찰자료UPLOAD
    oFilter.AddEx "PH_PY010"            '일일근태처리
    oFilter.AddEx "PH_PY011"            '전문직 호칭 일괄 변경(2013.07.05 송명규 추가)
    oFilter.AddEx "PH_PY012"            '출장등록
    oFilter.AddEx "PH_PY013"            '위해일수계산
    oFilter.AddEx "PH_PY014"            '위해일수수정
    oFilter.AddEx "PH_PY015"            '연차적치등록
    oFilter.AddEx "PH_PY016"            '기본업무등록
    oFilter.AddEx "PH_PY017"            '월근태집계
    oFilter.AddEx "PH_PY018"            '휴일근무체크(연봉제)
    oFilter.AddEx "PH_PY019"            '반변경등록
    oFilter.AddEx "PH_PY020"            '일근태 업무변경등록
    oFilter.AddEx "PH_PY021"            '사원비상연락처관리
    
    oFilter.AddEx "PH_PY201"            '정년임박자 휴가경비 등록
    oFilter.AddEx "PH_PY202"            '정년임박자 휴가경비 조회
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회
    
    '//인사 - 리포트
    oFilter.AddEx "PH_PY501"            '여권발급현황
    oFilter.AddEx "PH_PY505"            '입사자대장
    oFilter.AddEx "PH_PY510"            '사원명부
    oFilter.AddEx "PH_PY515"            '재직자사원명부
    oFilter.AddEx "PH_PY520"            '퇴직및퇴직예정자대장
    oFilter.AddEx "PH_PY525"            '학력별인원현황
    oFilter.AddEx "PH_PY530"            '연령별인원현황
    oFilter.AddEx "PH_PY535"            '근속년수별인원현황
    oFilter.AddEx "PH_PY540"            '인원현황(대외용)
    oFilter.AddEx "PH_PY545"            '인원현황(대내용)
    oFilter.AddEx "PH_PY550"            '전체인원현황
    oFilter.AddEx "PH_PY555"            '일일근무자현황
    oFilter.AddEx "PH_PY560"            '일출근현황
    oFilter.AddEx "PH_PY565"            '연장근무자현황
    oFilter.AddEx "PH_PY570"            '연장/휴일근무자현황
    oFilter.AddEx "PH_PY575"            '근태기찰현황
    oFilter.AddEx "PH_PY580"            '개인별근태월보
    oFilter.AddEx "PH_PY585"            '일일출근기록부
    oFilter.AddEx "PH_PY590"            '기간별근태집계표
    oFilter.AddEx "PH_PY595"            '근속년수현황
    oFilter.AddEx "PH_PY600"            '일자별연장근무현황
    oFilter.AddEx "PH_PY605"            '근속보전휴가발생및사용내역
    oFilter.AddEx "PH_PY610"            '근태구분별사용내역
    oFilter.AddEx "PH_PY615"            '당직근무현황
    oFilter.AddEx "PH_PY620"            '연봉제휴일근무자현황
    oFilter.AddEx "PH_PY635"            '여행,교육자현황
    oFilter.AddEx "PH_PY640"            '국민연금퇴직전환금현황
    oFilter.AddEx "PH_PY645"            '자격수당지급현황
    oFilter.AddEx "PH_PY650"            '노동조합간부현황
    oFilter.AddEx "PH_PY655"            '보훈대상자현황
    oFilter.AddEx "PH_PY660"            '장애근로자현황
    oFilter.AddEx "PH_PY665"            '사원자녀현황
    oFilter.AddEx "PH_PY670"            '개인별차량현황
    oFilter.AddEx "PH_PY675"            '근무편성현황
    oFilter.AddEx "PH_PY676"            '근태시간내역조회
    oFilter.AddEx "PH_PY677"            '일일근태이상자조회
    oFilter.AddEx "PH_PY679"            '개인별 근태집계 조회
    oFilter.AddEx "PH_PY680"            '상벌현황
    oFilter.AddEx "PH_PY685"            '포상가급현황
    oFilter.AddEx "PH_PY690"            '생일자현황
    oFilter.AddEx "PH_PY695"            '인사기록카드
    oFilter.AddEx "PH_PY705"            '교통비지급근태확인
    oFilter.AddEx "PH_PY860"            '호봉표조회
    oFilter.AddEx "PH_PY503"            '승진대상자명부
    oFilter.AddEx "PH_PY678"            '당직근무자 일괄 등록
    oFilter.AddEx "PH_PY507"            '휴직자현황
    oFilter.AddEx "PH_PY681"            '비근무일수현황
    oFilter.AddEx "PH_PY935"            '정기승호현황
    oFilter.AddEx "PH_PY551"            '평균인원조회
    oFilter.AddEx "PH_PY508"            '재직증명 등록 및 발급
    oFilter.AddEx "PH_PY522"            '임금피크대상자현황
    oFilter.AddEx "PH_PY523"            '임금피크대상자월별차수현황
    oFilter.AddEx "PH_PY524"            '퇴직금 중간 정산내역
    oFilter.AddEx "PH_PY683"            '교대근무인정현황
    oFilter.AddEx "PH_PYA65"            '년차현황 (집계)
    oFilter.AddEx "PH_PY583"            '개인별 근태집계 조회
    
    '//급여관리
    oFilter.AddEx "PH_PY100"            '기준세액설정
    oFilter.AddEx "PH_PY101"            '보험률등록
    oFilter.AddEx "PH_PY102"            '수당항목설정
    oFilter.AddEx "PH_PY103"            '공제항목설정
    oFilter.AddEx "PH_PY104"            '고정수당공제금액일괄등록
    oFilter.AddEx "PH_PY105"            '호봉표등록
    oFilter.AddEx "PH_PY106"            '수당계산식설정
    oFilter.AddEx "PH_PY107"            '급상여기준일설정
    oFilter.AddEx "PH_PY108"            '상여율지급설정
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    oFilter.AddEx "PH_PY109_1"            '급상여변동자료 항목수정
    oFilter.AddEx "PH_PY110"            '개인상여율등록
    oFilter.AddEx "PH_PY111"            '급상여계산
    oFilter.AddEx "PH_PY112"            '급상여자료관리
    oFilter.AddEx "PH_PY113"            '급상여분개자료생성
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY115"            '퇴직금계산
    oFilter.AddEx "PH_PY116"            '퇴직금분개자료생성
    oFilter.AddEx "PH_PY117"            '급상여마감작업
    oFilter.AddEx "PH_PY118"            '급상여Email발송
    oFilter.AddEx "PH_PY119"            '급상여은행파일생성
    oFilter.AddEx "PH_PY120"            '급상여소급집계처리
    oFilter.AddEx "PH_PY121"            '평가가급액 등록
    oFilter.AddEx "PH_PY122"            '급상여출력 개인부서설정등록
    oFilter.AddEx "PH_PY123"            '가압류등록
    oFilter.AddEx "PH_PY125"            '퇴직연금 설정
    oFilter.AddEx "PH_PY127"            '//개인별 4대보험 보수월액 및 정산금액입력
    oFilter.AddEx "PH_PY130"            '팀별 성과급차등 등급등록
    oFilter.AddEx "PH_PY131"            '성과급차등 계수등록
    oFilter.AddEx "PH_PY132"            '성과급차 개인별 계산
    oFilter.AddEx "PH_PY133"            '연봉제 횟차 관리
    oFilter.AddEx "PH_PY134"            '소득세/주민세 조정관리
    oFilter.AddEx "PH_PY129"            '개인별퇴직연금(DC형) 계산
    '//급여관리 - 리포트
    oFilter.AddEx "PH_PY625"            '세탁자명부
    oFilter.AddEx "PH_PY630"            '사원별노조비공제현황
    oFilter.AddEx "PH_PY700"            '급여지급대장
    oFilter.AddEx "PH_PY710"            '상여지급대장
    oFilter.AddEx "PH_PY715"            '급여부서별집계대장
    oFilter.AddEx "PH_PY720"            '상여부서별집계대장
    oFilter.AddEx "PH_PY725"            '급여직급별집계대장
    oFilter.AddEx "PH_PY740"            '상여직급별집계대장
    oFilter.AddEx "PH_PY730"            '급여봉투출력
    oFilter.AddEx "PH_PY735"            '상여봉투출력
    oFilter.AddEx "PH_PY745"            '연간지급현황
    oFilter.AddEx "PH_PY750"            '근로소득징수현황
    oFilter.AddEx "PH_PY755"            '동호회가입현황
    oFilter.AddEx "PH_PY760"            '평균임금및퇴직금산출내역서
    oFilter.AddEx "PH_PY765"            '급여증감내역서
    oFilter.AddEx "PH_PY770"            '퇴직소득원천징수영수증출력
    oFilter.AddEx "PH_PY775"            '개인별년차현황
    oFilter.AddEx "PH_PY776"            '잔여년차현황
    oFilter.AddEx "PH_PY780"            '월고용보험내역
    oFilter.AddEx "PH_PY785"            '월국민연금내역
    oFilter.AddEx "PH_PY790"            '월건강보험내역
    oFilter.AddEx "PH_PY795"            '연간부서별급여내역
    oFilter.AddEx "PH_PY800"            '인건비지급자료
    oFilter.AddEx "PH_PY805"            '급여수당변동내역
    oFilter.AddEx "PH_PY810"            '직급별통상임금내역
    oFilter.AddEx "PH_PY815"            '평균임금내역
    oFilter.AddEx "PH_PY820"            '통상임금내역
    oFilter.AddEx "PH_PY825"            '전문직O/T현황
    oFilter.AddEx "PH_PY830"            '부서별인건비현황 (기획)
    oFilter.AddEx "PH_PY835"            '직급별O/T및수당현황
    oFilter.AddEx "PH_PY840"            '풍산전자공시자료
    oFilter.AddEx "PH_PY845"            '기간별급여지급내역
    oFilter.AddEx "PH_PY850"            '소급분지급명세서
    oFilter.AddEx "PH_PY855"            '개인별임금지급대장
    oFilter.AddEx "PH_PY865"            '고용보험현황 (계산용)
    oFilter.AddEx "PH_PY870"            '담당별월O/T및수당현황
    oFilter.AddEx "PH_PY875"            '직급별수당집계대장
    oFilter.AddEx "PH_PY716"            '기간별급여부서별집계대장
    oFilter.AddEx "PH_PY721"            '기간별상여부서별집계대장
    oFilter.AddEx "PH_PY717"            '기간별급여반별집계대장
    oFilter.AddEx "PH_PY718"            '생산완료금액대비O/T현황
    oFilter.AddEx "PH_PY701"            '급여지급대장 (노조용)
    
    oFilter.AddEx "PH_PYA10"            '급여지급대장(부서)
    oFilter.AddEx "PH_PYA20"            '급여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA30"            '상여지급대장(부서)
    oFilter.AddEx "PH_PYA40"            '상여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA50"            'DC전환자부담금지급내역
    oFilter.AddEx "PH_PYA75"            '교통비외수당지급대장
    
    '//정산관리
    
    oFilter.AddEx "PH_PY401"            '전근무지등록
    oFilter.AddEx "PH_PY402"            '정산기초자료 등록
    oFilter.AddEx "PH_PY405"            '의료비등록
    oFilter.AddEx "PH_PY407"            '기부금등록
    oFilter.AddEx "PH_PY409"            '기부금조정명세등록
    oFilter.AddEx "PH_PY411"            '연금.저축등소득공제등록
    oFilter.AddEx "PH_PY413"            '월세액.주택임차차입금자료 등록
    oFilter.AddEx "PH_PY415"            '정산계산
    oFilter.AddEx "PH_PY417"            '정산 은행파일생성
    oFilter.AddEx "PH_PY980"            '신고_근로소득지급명세서자료작성
    oFilter.AddEx "PH_PY985"            '신고_의료비지급명세서자료작성
    oFilter.AddEx "PH_PY990"            '신고_기부금명세서자료작성
    oFilter.AddEx "PH_PY995"            '신고_퇴직소득지급명세서자료작성
    oFilter.AddEx "PH_PY419"            '표준세액적용대상자등록
    
    oFilter.AddEx "PH_PY910"            '소득공제신고서출력
    oFilter.AddEx "PH_PY915"            '근로소득원천징수부출력
    oFilter.AddEx "PH_PY920"            '원천징수영수증출력
    oFilter.AddEx "PH_PY925"            '기부금명세서출력
    oFilter.AddEx "PH_PY930"            '정산징수및환급대장
    oFilter.AddEx "PH_PY931"            '표준세액적용대상자조회
    oFilter.AddEx "PH_PY932"            '전근무지등록현황
    oFilter.AddEx "PH_PY933"            '보수총액신고기초자료
    oFilter.AddEx "PH_PYA55"            '정산징수및환급대장(집계)
    oFilter.AddEx "PH_PYA70"            '소득세원천징수세액조정신청서출력
    
    
    oFilter.AddEx "ZPY341"              '월별 정산자료 생성
    oFilter.AddEx "ZPY343"              '월별 자료 관리
    oFilter.AddEx "ZPY421"              '퇴직소득전산매체수록
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    oFilter.AddEx "ZPY502"              '종(전) 근무지 등록
    oFilter.AddEx "ZPY503"              '정산세액계산
    oFilter.AddEx "ZPY504"              '정산결과조회
    oFilter.AddEx "ZPY505"              '기부금명세등록
    oFilter.AddEx "ZPY506"              '의료비명세등록
    oFilter.AddEx "ZPY507"              '정산결과조회(전체)
    oFilter.AddEx "ZPY508"              '연금저축 소득공제 명세 등록
    oFilter.AddEx "ZPY509"              '정산자료 마감작업
    oFilter.AddEx "ZPY510"              '종전근무지 일괄생성
    oFilter.AddEx "ZPY521"              '근로소득전산매체수록
    oFilter.AddEx "ZPY522"              '의료비 기부금 전산매체수록
    
    oFilter.AddEx "RPY401"              '퇴직원천징수 영수증
    oFilter.AddEx "RPY501"              '월별자료현황
    oFilter.AddEx "RPY502"              '종전근무지현황
    oFilter.AddEx "RPY503"              '근로소득 원천징수부
    oFilter.AddEx "RPY504"              '근로소득 원천영수증
    oFilter.AddEx "RPY505"              '소득자료집계표
    oFilter.AddEx "RPY506"              '정산징수환급대장
    oFilter.AddEx "RPY508"              '연말정산집계표
    oFilter.AddEx "RPY509"              '갑근세신고검토표
    oFilter.AddEx "RPY510"              '비과세근로소득명세서
    oFilter.AddEx "RPY511"              '기부금명세서
    
    
    '//기타관리
    oFilter.AddEx "PH_PY301"            '학자금신청등록
    oFilter.AddEx "PH_PY302"            '학자금지급완료처리
    oFilter.AddEx "PH_PY303"            '학자금은행파일생성
    oFilter.AddEx "PH_PY305"            '학자금신청서
    oFilter.AddEx "PH_PY306"            '학자금신청내역(개인별)
    oFilter.AddEx "PH_PY307"            '학자금신청내역(분기별)
    oFilter.AddEx "PH_PY309"            '대부금등록
    oFilter.AddEx "PH_PY310"            '대부금개별상환
    oFilter.AddEx "PH_PY311"            '통근버스운행등록
    oFilter.AddEx "PH_PY312"            '버스요금 개인별등록
    oFilter.AddEx "PH_PY313"            '대부금계산
    oFilter.AddEx "PH_PY314"            '대부금계산 내역 조회(급여변동자료용)
    oFilter.AddEx "PH_PY030"            '공용등록
    oFilter.AddEx "PH_PY031"            '출장등록
    oFilter.AddEx "PH_PY032"            '사용외출등록
    oFilter.AddEx "PH_PY315"            '개인별대부금잔액현황
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)
    
End Sub

Private Sub KEY_DOWN(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '2
    Set oFilter = oFilters.Add(et_KEY_DOWN)
    
    
    '//System Form Type
    '//인사관리
    '//급여관리
    '//급여관리-리포트
    oFilter.AddEx "PH_PY718"            '생산완료금액대비O/T현황
    oFilter.AddEx "PH_PY701"            '급여지급대장 (노조용)
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//인사관리
    oFilter.AddEx "PH_PY005"            '사업장정보등록
    oFilter.AddEx "PH_PY008"            '일근태등록
    oFilter.AddEx "PH_PY011"            '전문직 호칭 일괄 변경(2013.07.05 송명규 추가)
    oFilter.AddEx "PH_PY012"            '출장등록
    oFilter.AddEx "PH_PY014"            '위해일수수정
    oFilter.AddEx "PH_PY015"            '연차적치등록
    oFilter.AddEx "PH_PY018"            '휴일근무체크(연봉제)
    oFilter.AddEx "PH_PY508"            '재직증명 등록 및 발급(2015.05.18 송명규 추가)
    oFilter.AddEx "PH_PY021"            '사원비상연락처관리
    
    oFilter.AddEx "PH_PY201"            '정년임박자 휴가경비 등록
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회
    oFilter.AddEx "PH_PY695"            '인사기록카드
    
    '근태관리 - 리포트
    oFilter.AddEx "PH_PY580"            '개인근태월보
    oFilter.AddEx "PH_PY575"            '근태기찰현황
    oFilter.AddEx "PH_PY681"            '비근무일수현황

    '//급여관리
    oFilter.AddEx "PH_PY102"            '수당항목설정
    oFilter.AddEx "PH_PY104"            '고정수당공제금액일괄등록
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    oFilter.AddEx "PH_PY109_1"            '급상여변동자료 항목수정
    oFilter.AddEx "PH_PY110"            '개인별상여율등록
    oFilter.AddEx "PH_PY111"            '급상여계산
    oFilter.AddEx "PH_PY113"            '급상여분개자료생성
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY115"            '퇴직금계산
    oFilter.AddEx "PH_PY116"            '퇴직금분개자료생성
    oFilter.AddEx "PH_PY121"            '평가가급액 등록
    oFilter.AddEx "PH_PY676"            '근태시간내역조회
    oFilter.AddEx "PH_PY677"            '일일근태이상자조회
    oFilter.AddEx "PH_PY700"            '급여지급대장
    oFilter.AddEx "PH_PY710"            '상여지급대장
    oFilter.AddEx "PH_PY715"            '급여지급대장(부서집계)
    oFilter.AddEx "PH_PY720"            '상여지급대장(부서집계)
    oFilter.AddEx "PH_PY122"            '급상여출력 개인부서설정등록
    oFilter.AddEx "PH_PY123"            '가압류등록
    oFilter.AddEx "PH_PY678"            '당직근무자 일괄 등록
    
    
    '//정산관리
    oFilter.AddEx "PH_PY401"            '전근무지등록
    oFilter.AddEx "PH_PY402"            '정산기초자료 등록
    oFilter.AddEx "PH_PY405"            '의료비등록
    oFilter.AddEx "PH_PY407"            '기부금등록
    oFilter.AddEx "PH_PY409"            '기부금조정명세등록
    oFilter.AddEx "PH_PY411"            '연금.저축등소득공제등록
    oFilter.AddEx "PH_PY413"            '월세액.주택임차차입금자료 등록
    oFilter.AddEx "PH_PY415"            '정산계산
    oFilter.AddEx "PH_PY980"            '신고_근로소득지급명세서자료작성
    oFilter.AddEx "PH_PY985"            '신고_의료비지급명세서자료작성
    oFilter.AddEx "PH_PY990"            '신고_기부금명세서자료작성
    oFilter.AddEx "PH_PY995"            '신고_퇴직소득지급명세서자료작성
    oFilter.AddEx "PH_PY419"            '표준세액적용대상자등록
    
    oFilter.AddEx "PH_PY910"            '소득공제신고서출력
    oFilter.AddEx "PH_PY915"            '근로소득원천징수부출력
    oFilter.AddEx "PH_PY920"            '원천징수영수증출력
    oFilter.AddEx "PH_PY925"            '기부금명세서출력
    oFilter.AddEx "PH_PY930"            '정산징수및환급대장
    oFilter.AddEx "PH_PY931"            '표준세액적용대상자조회
    oFilter.AddEx "PH_PY932"            '전근무지등록현황
    oFilter.AddEx "PH_PY933"            '보수총액신고기초자료
    oFilter.AddEx "PH_PYA55"            '정산징수및환급대장(집계)
    oFilter.AddEx "PH_PYA70"            '소득세원천징수세액조정신청서출력
    
    
    oFilter.AddEx "ZPY341"              '월별 정산자료 생성
    oFilter.AddEx "ZPY343"              '월별 자료 관리
    oFilter.AddEx "ZPY421"              '퇴직소득전산매체수록
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    oFilter.AddEx "ZPY502"              '종(전) 근무지 등록
    oFilter.AddEx "ZPY503"              '정산세액계산
    oFilter.AddEx "ZPY504"              '정산결과조회
    oFilter.AddEx "ZPY505"              '기부금명세등록
    oFilter.AddEx "ZPY506"              '의료비명세등록
    oFilter.AddEx "ZPY508"              '연금저축 소득공제 명세 등록
    oFilter.AddEx "ZPY509"              '정산자료 마감작업
    oFilter.AddEx "ZPY510"              '종전근무지 일괄생성
    oFilter.AddEx "ZPY521"              '근로소득전산매체수록
    oFilter.AddEx "ZPY522"              '의료비 기부금 전산매체수록
    
    oFilter.AddEx "RPY401"              '퇴직원천징수 연수증
    oFilter.AddEx "RPY501"              '월별자료현황
    oFilter.AddEx "RPY502"              '종전근무지현황
    oFilter.AddEx "RPY503"              '근로소득 원천징수부
    oFilter.AddEx "RPY504"              '근로소득 원천영수증
    oFilter.AddEx "RPY505"              '소득자료집계표
    oFilter.AddEx "RPY506"              '정산징수환급대장
    oFilter.AddEx "RPY508"              '연말정산집계표
    oFilter.AddEx "RPY509"              '갑근세신고검토표
    oFilter.AddEx "RPY510"              '비과세근로소득명세서
    oFilter.AddEx "RPY511"              '기부금명세서
    
    '//기타관리
    oFilter.AddEx "PH_PY301"            '학자금신청등록
    oFilter.AddEx "PH_PY305"            '학자금신청서
    oFilter.AddEx "PH_PY306"            '학자금신청내역(개인별)
    oFilter.AddEx "PH_PY309"            '대부금등록
    oFilter.AddEx "PH_PY310"            '대부금개별상환
    oFilter.AddEx "PH_PY313"            '대부금계산
    oFilter.AddEx "PH_PY314"            '대부금계산 내역 조회(급여변동자료용)
    oFilter.AddEx "PH_PY030"            '공용등록
    oFilter.AddEx "PH_PY031"            '출장등록
    oFilter.AddEx "PH_PY032"            '사용외출등록
    oFilter.AddEx "PH_PY315"            '개인별대부금잔액현황
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)

End Sub

Private Sub GOT_FOCUS(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '3
    Set oFilter = oFilters.Add(et_GOT_FOCUS)
    
    
    '//System Form Type
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    oFilter.AddEx "PH_PY000"            '사용자권한관리
    
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터 등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    oFilter.AddEx "PH_PY004"            '근무조편성등록
    oFilter.AddEx "PH_PY005"            '사업장정보등록
    oFilter.AddEx "PH_PY006"            '승호작업등록
    oFilter.AddEx "PH_PY007"            '유류단가등록
    oFilter.AddEx "PH_PY008"            '일근태등록
    oFilter.AddEx "PH_PY009"            '기찰자료UPLOAD
    oFilter.AddEx "PH_PY011"            '전문직 호칭 일괄 변경(2013.07.05 송명규 추가)
    oFilter.AddEx "PH_PY013"            '위해일수계산
    oFilter.AddEx "PH_PY014"            '위해일수수정
    oFilter.AddEx "PH_PY015"            '연차적치등록
    oFilter.AddEx "PH_PY016"            '기본업무등록
    oFilter.AddEx "PH_PY017"            '월근태집계
    oFilter.AddEx "PH_PY018"            '휴일근무체크(연봉제)
    oFilter.AddEx "PH_PY019"            '반변경등록
    oFilter.AddEx "PH_PY020"            '일근태 업무변경등록
    
    oFilter.AddEx "PH_PY201"            '정년임박자 휴가경비 등록
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회
    
    '//인사 - 리포트
    oFilter.AddEx "PH_PY501"            '여권발급현황
    oFilter.AddEx "PH_PY505"            '입사자대장
    oFilter.AddEx "PH_PY510"            '사원명부
    oFilter.AddEx "PH_PY515"            '재직자사원명부
    oFilter.AddEx "PH_PY520"            '퇴직및퇴직예정자대장
    oFilter.AddEx "PH_PY525"            '학력별인원현황
    oFilter.AddEx "PH_PY530"            '연령별인원현황
    oFilter.AddEx "PH_PY535"            '근속년수별인원현황
    oFilter.AddEx "PH_PY540"            '인원현황(대외용)
    oFilter.AddEx "PH_PY545"            '인원현황(대내용)
    oFilter.AddEx "PH_PY550"            '전체인원현황
    oFilter.AddEx "PH_PY555"            '일일근무자현황
    oFilter.AddEx "PH_PY560"            '일출근현황
    oFilter.AddEx "PH_PY565"            '연장근무자현황
    oFilter.AddEx "PH_PY570"            '연장/휴일근무자현황
    oFilter.AddEx "PH_PY575"            '근태기찰현황
    oFilter.AddEx "PH_PY580"            '개인별근태월보
    oFilter.AddEx "PH_PY585"            '일일출근기록부
    oFilter.AddEx "PH_PY590"            '기간별근태집계표
    oFilter.AddEx "PH_PY595"            '근속년수현황
    oFilter.AddEx "PH_PY600"            '일자별연장근무현황
    oFilter.AddEx "PH_PY605"            '근속보전휴가발생및사용내역
    oFilter.AddEx "PH_PY610"            '근태구분별사용내역
    oFilter.AddEx "PH_PY615"            '당직근무현황
    oFilter.AddEx "PH_PY620"            '연봉제휴일근무자현황
    oFilter.AddEx "PH_PY635"            '여행,교육자현황
    oFilter.AddEx "PH_PY640"            '국민연금퇴직전환금현황
    oFilter.AddEx "PH_PY645"            '자격수당지급현황
    oFilter.AddEx "PH_PY650"            '노동조합간부현황
    oFilter.AddEx "PH_PY655"            '보훈대상자현황
    oFilter.AddEx "PH_PY660"            '장애근로자현황
    oFilter.AddEx "PH_PY665"            '사원자녀현황
    oFilter.AddEx "PH_PY670"            '개인별차량현황
    oFilter.AddEx "PH_PY675"            '근무편성현황
    oFilter.AddEx "PH_PY679"            '개인별 근태집계 조회
    oFilter.AddEx "PH_PY680"            '상벌현황
    oFilter.AddEx "PH_PY685"            '포상가급현황
    oFilter.AddEx "PH_PY690"            '생일자현황
    oFilter.AddEx "PH_PY695"            '인사기록카드
    oFilter.AddEx "PH_PY705"            '교통비지급근태확인
    oFilter.AddEx "PH_PY860"            '호봉표조회
    oFilter.AddEx "PH_PY503"            '승진대상자명부
    oFilter.AddEx "PH_PY678"            '당직근무자 일괄 등록
    oFilter.AddEx "PH_PY507"            '휴직자현황
    oFilter.AddEx "PH_PY681"            '비근무일수현황
    oFilter.AddEx "PH_PY935"            '정기승호현황
    oFilter.AddEx "PH_PY551"            '평균인원조회
    oFilter.AddEx "PH_PY508"            '재직증명 등록 및 발급
    oFilter.AddEx "PH_PY522"            '임금피크대상자현황
    oFilter.AddEx "PH_PY523"            '임금피크대상자월별차수현황
    oFilter.AddEx "PH_PY524"            '퇴직금 중간 정산내역
    oFilter.AddEx "PH_PY683"            '교대근무인정현황
    oFilter.AddEx "PH_PYA65"            '년차현황 (집계)
    oFilter.AddEx "PH_PY583"            '개인별 근태집계 조회
    
    '//급여관리
    oFilter.AddEx "PH_PY100"            '기준세액설정
    oFilter.AddEx "PH_PY101"            '보험률등록
    oFilter.AddEx "PH_PY102"            '수당항목설정
    oFilter.AddEx "PH_PY103"            '공제항목설정
    oFilter.AddEx "PH_PY104"            '고정수당공제금액일괄등록
    oFilter.AddEx "PH_PY105"            '호봉표등록
    oFilter.AddEx "PH_PY106"            '수당계산식설정
    oFilter.AddEx "PH_PY107"            '급상여기준일설정
    oFilter.AddEx "PH_PY108"            '상여율지급설정
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    oFilter.AddEx "PH_PY109_1"            '급상여변동자료 항목수정
    oFilter.AddEx "PH_PY110"            '개인상여율등록
    oFilter.AddEx "PH_PY111"            '급상여계산
    oFilter.AddEx "PH_PY112"            '급상여자료관리
    oFilter.AddEx "PH_PY113"            '급상여분개자료생성
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY115"            '퇴직금계산
    oFilter.AddEx "PH_PY116"            '퇴직금분개자료생성
    oFilter.AddEx "PH_PY117"            '급상여마감작업
    oFilter.AddEx "PH_PY118"            '급상여Email발송
    oFilter.AddEx "PH_PY120"            '급상여소급집계처리
    oFilter.AddEx "PH_PY121"            '평가가급액 등록
    oFilter.AddEx "PH_PY122"            '급상여출력 개인부서설정등록
    oFilter.AddEx "PH_PY123"            '가압류등록
    oFilter.AddEx "PH_PY125"            '퇴직연금 설정
    oFilter.AddEx "PH_PY127"            '//개인별 4대보험 보수월액 및 정산금액입력
    oFilter.AddEx "PH_PY129"            '개인별퇴직연금(DC형) 계산
    
    '//급여관리 - 리포트
    oFilter.AddEx "PH_PY625"            '세탁자명부
    oFilter.AddEx "PH_PY630"            '사원별노조비공제현황
    oFilter.AddEx "PH_PY700"            '급여지급대장
    oFilter.AddEx "PH_PY710"            '상여지급대장
    oFilter.AddEx "PH_PY715"            '급여부서별집계대장
    oFilter.AddEx "PH_PY720"            '상여부서별집계대장
    oFilter.AddEx "PH_PY725"            '급여직급별집계대장
    oFilter.AddEx "PH_PY740"            '상여직급별집계대장
    oFilter.AddEx "PH_PY730"            '급여봉투출력
    oFilter.AddEx "PH_PY735"            '상여봉투출력
    oFilter.AddEx "PH_PY745"            '연간지급현황
    oFilter.AddEx "PH_PY750"            '근로소득징수현황
    oFilter.AddEx "PH_PY755"            '동호회가입현황
    oFilter.AddEx "PH_PY760"            '평균임금및퇴직금산출내역서
    oFilter.AddEx "PH_PY765"            '급여증감내역서
    oFilter.AddEx "PH_PY770"            '퇴직소득원천징수영수증출력
    oFilter.AddEx "PH_PY775"            '개인별년차현황
    oFilter.AddEx "PH_PY776"            '잔여년차현황
    oFilter.AddEx "PH_PY780"            '월고용보험내역
    oFilter.AddEx "PH_PY785"            '월국민연금내역
    oFilter.AddEx "PH_PY790"            '월건강보험내역
    oFilter.AddEx "PH_PY795"            '연간부서별급여내역
    oFilter.AddEx "PH_PY800"            '인건비지급자료
    oFilter.AddEx "PH_PY805"            '급여수당변동내역
    oFilter.AddEx "PH_PY810"            '직급별통상임금내역
    oFilter.AddEx "PH_PY815"            '평균임금내역
    oFilter.AddEx "PH_PY820"            '통상임금내역
    oFilter.AddEx "PH_PY825"            '전문직O/T현황
    oFilter.AddEx "PH_PY830"            '부서별인건비현황 (기획)
    oFilter.AddEx "PH_PY835"            '직급별O/T및수당현황
    oFilter.AddEx "PH_PY840"            '풍산전자공시자료
    oFilter.AddEx "PH_PY845"            '기간별급여지급내역
    oFilter.AddEx "PH_PY850"            '소급분지급명세서
    oFilter.AddEx "PH_PY855"            '개인별임금지급대장
    oFilter.AddEx "PH_PY865"            '고용보험현황 (계산용)
    oFilter.AddEx "PH_PY870"            '담당별월O/T및수당현황
    oFilter.AddEx "PH_PY875"            '직급별수당집계대장
    oFilter.AddEx "PH_PY716"            '기간별급여부서별집계대장
    oFilter.AddEx "PH_PY721"            '기간별상여부서별집계대장
    oFilter.AddEx "PH_PY717"            '기간별급여반별집계대장
    oFilter.AddEx "PH_PY718"            '생산완료금액대비O/T현황
    oFilter.AddEx "PH_PY701"            '급여지급대장 (노조용)
    
    oFilter.AddEx "PH_PYA10"            '급여지급대장(부서)
    oFilter.AddEx "PH_PYA20"            '급여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA30"            '상여지급대장(부서)
    oFilter.AddEx "PH_PYA40"            '상여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA50"            'DC전환자부담금지급내역
    oFilter.AddEx "PH_PYA75"            '교통비외수당지급대장
    
    '//정산관리
    oFilter.AddEx "PH_PY401"            '전근무지등록
    oFilter.AddEx "PH_PY402"            '정산기초자료 등록
    oFilter.AddEx "PH_PY405"            '의료비등록
    oFilter.AddEx "PH_PY407"            '기부금등록
    oFilter.AddEx "PH_PY409"            '기부금조정명세등록
    oFilter.AddEx "PH_PY411"            '연금.저축등소득공제등록
    oFilter.AddEx "PH_PY413"            '월세액.주택임차차입금자료 등록
    oFilter.AddEx "PH_PY415"            '정산계산
    oFilter.AddEx "PH_PY980"            '신고_근로소득지급명세서자료작성
    oFilter.AddEx "PH_PY985"            '신고_의료비지급명세서자료작성
    oFilter.AddEx "PH_PY990"            '신고_기부금명세서자료작성
    oFilter.AddEx "PH_PY995"            '신고_퇴직소득지급명세서자료작성
    oFilter.AddEx "PH_PY419"            '표준세액적용대상자등록
    
    oFilter.AddEx "PH_PY910"            '소득공제신고서출력
    oFilter.AddEx "PH_PY915"            '근로소득원천징수부출력
    oFilter.AddEx "PH_PY920"            '원천징수영수증출력
    oFilter.AddEx "PH_PY925"            '기부금명세서출력
    oFilter.AddEx "PH_PY930"            '정산징수및환급대장
    oFilter.AddEx "PH_PY931"            '표준세액적용대상자조회
    oFilter.AddEx "PH_PY932"            '전근무지등록현황
    oFilter.AddEx "PH_PY933"            '보수총액신고기초자료
    oFilter.AddEx "PH_PYA55"            '정산징수및환급대장(집계)
    oFilter.AddEx "PH_PYA70"            '소득세원천징수세액조정신청서출력
    
    
    oFilter.AddEx "ZPY341"              '월별 정산자료 생성
    oFilter.AddEx "ZPY343"              '월별 자료 관리
    oFilter.AddEx "ZPY421"              '퇴직소득전산매체수록
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    oFilter.AddEx "ZPY502"              '종(전) 근무지 등록
    oFilter.AddEx "ZPY503"              '정산세액계산
    oFilter.AddEx "ZPY504"              '정산결과조회
    oFilter.AddEx "ZPY505"              '기부금명세등록
    oFilter.AddEx "ZPY506"              '의료비명세등록
    oFilter.AddEx "ZPY508"              '연금저축 소득공제 명세 등록
    oFilter.AddEx "ZPY509"              '정산자료 마감작업
    oFilter.AddEx "ZPY521"              '근로소득전산매체수록
    oFilter.AddEx "ZPY522"              '의료비 기부금 전산매체수록
    
    '//기타관리
    oFilter.AddEx "PH_PY311"            '통근버스운행등록
    oFilter.AddEx "PH_PY312"            '버스요금 개인별등록
    oFilter.AddEx "PH_PY309"            '대부금등록
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)
    
End Sub

Private Sub LOST_FOCUS(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '4
    Set oFilter = oFilters.Add(et_LOST_FOCUS)

    
    '//System Form Type
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    '//AddOn Form Type
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
End Sub

Private Sub COMBO_SELECT(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '5
    Set oFilter = oFilters.Add(et_COMBO_SELECT)

    
    '//System Form Type
   
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터 등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    oFilter.AddEx "PH_PY004"            '근무조편성등록
    oFilter.AddEx "PH_PY005"            '사업장정보등록
    oFilter.AddEx "PH_PY006"            '승호작업등록
    oFilter.AddEx "PH_PY007"            '유류단가등록
    oFilter.AddEx "PH_PY008"            '일근태등록
    oFilter.AddEx "PH_PY009"            '기찰자료UPLOAD
    oFilter.AddEx "PH_PY012"            '출장등록
    oFilter.AddEx "PH_PY013"            '위해일수계산
    oFilter.AddEx "PH_PY014"            '위해일수수정
    oFilter.AddEx "PH_PY016"            '기본업무등록
    oFilter.AddEx "PH_PY017"            '월근태집계
    oFilter.AddEx "PH_PY018"            '휴일근무자체크(연봉제)
    oFilter.AddEx "PH_PY019"            '반변경등록
    oFilter.AddEx "PH_PY020"            '일근태 업무변경등록
    oFilter.AddEx "PH_PY021"            '사원비상연락처관리
    
    oFilter.AddEx "PH_PY201"            '정년임박자 휴가경비 등록
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회
    
    '//인사 - 리포트
    oFilter.AddEx "PH_PY501"            '여권발급현황
    oFilter.AddEx "PH_PY505"            '입사자대장
    oFilter.AddEx "PH_PY510"            '사원명부
    oFilter.AddEx "PH_PY515"            '재직자사원명부
    oFilter.AddEx "PH_PY520"            '퇴직및퇴직예정자대장
    oFilter.AddEx "PH_PY525"            '학력별인원현황
    oFilter.AddEx "PH_PY530"            '연령별인원현황
    oFilter.AddEx "PH_PY535"            '근속년수별인원현황
    oFilter.AddEx "PH_PY540"            '인원현황(대외용)
    oFilter.AddEx "PH_PY545"            '인원현황(대내용)
    oFilter.AddEx "PH_PY550"            '전체인원현황
    oFilter.AddEx "PH_PY555"            '일일근무자현황
    oFilter.AddEx "PH_PY560"            '일출근현황
    oFilter.AddEx "PH_PY565"            '연장근무자현황
    oFilter.AddEx "PH_PY570"            '연장/휴일근무자현황
    oFilter.AddEx "PH_PY575"            '근태기찰현황
    oFilter.AddEx "PH_PY580"            '개인별근태월보
    oFilter.AddEx "PH_PY585"            '일일출근기록부
    oFilter.AddEx "PH_PY590"            '기간별근태집계표
    oFilter.AddEx "PH_PY595"            '근속년수현황
    oFilter.AddEx "PH_PY600"            '일자별연장근무현황
    oFilter.AddEx "PH_PY605"            '근속보전휴가발생및사용내역
    oFilter.AddEx "PH_PY610"            '근태구분별사용내역
    oFilter.AddEx "PH_PY615"            '당직근무현황
    oFilter.AddEx "PH_PY620"            '연봉제휴일근무자현황
    oFilter.AddEx "PH_PY635"            '여행,교육자현황
    oFilter.AddEx "PH_PY640"            '국민연금퇴직전환금현황
    oFilter.AddEx "PH_PY645"            '자격수당지급현황
    oFilter.AddEx "PH_PY650"            '노동조합간부현황
    oFilter.AddEx "PH_PY655"            '보훈대상자현황
    oFilter.AddEx "PH_PY660"            '장애근로자현황
    oFilter.AddEx "PH_PY665"            '사원자녀현황
    oFilter.AddEx "PH_PY670"            '개인별차량현황
    oFilter.AddEx "PH_PY675"            '근무편성현황
    oFilter.AddEx "PH_PY676"            '근태시간내역조회
    oFilter.AddEx "PH_PY677"            '일일근태이상자조회
    oFilter.AddEx "PH_PY679"            '개인별 근태집계 조회
    oFilter.AddEx "PH_PY680"            '상벌현황
    oFilter.AddEx "PH_PY685"            '포상가급현황
    oFilter.AddEx "PH_PY690"            '생일자현황
    oFilter.AddEx "PH_PY695"            '인사기록카드
    oFilter.AddEx "PH_PY705"            '교통비지급근태확인
    oFilter.AddEx "PH_PY860"            '호봉표조회
    oFilter.AddEx "PH_PY503"            '승진대상자명부
    oFilter.AddEx "PH_PY678"            '당직근무자 일괄 등록
    oFilter.AddEx "PH_PY507"            '휴직자현황
    oFilter.AddEx "PH_PY681"            '비근무일수현황
    oFilter.AddEx "PH_PY935"            '정기승호현황
    oFilter.AddEx "PH_PY551"            '평균인원조회
    oFilter.AddEx "PH_PY508"            '재직증명 등록 및 발급
    oFilter.AddEx "PH_PY522"            '임금피크대상자현황
    oFilter.AddEx "PH_PY523"            '임금피크대상자월별차수현황
    oFilter.AddEx "PH_PY524"            '퇴직금 중간 정산내역
    oFilter.AddEx "PH_PY683"            '교대근무인정현황
    oFilter.AddEx "PH_PYA65"            '년차현황 (집계)
    oFilter.AddEx "PH_PY583"            '개인별 근태집계 조회
    
    '//급여관리
    oFilter.AddEx "PH_PY100"            '기준세액설정
    oFilter.AddEx "PH_PY101"            '보험률등록
    oFilter.AddEx "PH_PY102"            '수당항목설정
    oFilter.AddEx "PH_PY103"            '공제항목설정
    oFilter.AddEx "PH_PY104"            '고정수당공제금액일괄등록
    oFilter.AddEx "PH_PY105"            '호봉표등록
    oFilter.AddEx "PH_PY106"            '수당계산식설정
    oFilter.AddEx "PH_PY107"            '급상여기준일설정
    oFilter.AddEx "PH_PY108"            '상여율지급설정
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    oFilter.AddEx "PH_PY109_1"          '급상여변동자료 항목수정
    oFilter.AddEx "PH_PY110"            '개인상여율등록
    oFilter.AddEx "PH_PY111"            '급상여계산
    oFilter.AddEx "PH_PY112"            '급상여자료관리
    oFilter.AddEx "PH_PY113"            '급상여분개자료생성
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY115"            '퇴직금계산
    oFilter.AddEx "PH_PY116"            '퇴직금분개자료생성
    oFilter.AddEx "PH_PY117"            '급상여마감작업
    oFilter.AddEx "PH_PY118"            '급상여Email발송
    oFilter.AddEx "PH_PY119"            '급상여은행파일생성
    oFilter.AddEx "PH_PY120"            '급상여소급집계처리
    oFilter.AddEx "PH_PY121"            '평가가급액 등록
    oFilter.AddEx "PH_PY122"            '급상여출력 개인부서설정등록
    oFilter.AddEx "PH_PY123"            '가압류등록
    oFilter.AddEx "PH_PY127"            '//개인별 4대보험 보수월액 및 정산금액입력
    oFilter.AddEx "PH_PY130"            '팀별 성과급차등 등급등록
    oFilter.AddEx "PH_PY131"            '성과급차등 계수등록
    oFilter.AddEx "PH_PY132"            '성과급차 개인별 계산
    oFilter.AddEx "PH_PY133"            '연봉제 횟차 관리
    oFilter.AddEx "PH_PY134"            '소득세/주민세 조정관리
    oFilter.AddEx "PH_PY129"            '개인별퇴직연금(DC형) 계산
    
    '//급여관리 - 리포트
    oFilter.AddEx "PH_PY625"            '세탁자명부
    oFilter.AddEx "PH_PY630"            '사원별노조비공제현황
    oFilter.AddEx "PH_PY700"            '급여지급대장
    oFilter.AddEx "PH_PY710"            '상여지급대장
    oFilter.AddEx "PH_PY715"            '급여부서별집계대장
    oFilter.AddEx "PH_PY720"            '상여부서별집계대장
    oFilter.AddEx "PH_PY725"            '급여직급별집계대장
    oFilter.AddEx "PH_PY740"            '상여직급별집계대장
    oFilter.AddEx "PH_PY730"            '급여봉투출력
    oFilter.AddEx "PH_PY735"            '상여봉투출력
    oFilter.AddEx "PH_PY745"            '연간지급현황
    oFilter.AddEx "PH_PY750"            '근로소득징수현황
    oFilter.AddEx "PH_PY755"            '동호회가입현황
    oFilter.AddEx "PH_PY760"            '평균임금및퇴직금산출내역서
    oFilter.AddEx "PH_PY765"            '급여증감내역서
    oFilter.AddEx "PH_PY770"            '퇴직소득원천징수영수증출력
    oFilter.AddEx "PH_PY775"            '개인별년차현황
    oFilter.AddEx "PH_PY776"            '잔여년차현황
    oFilter.AddEx "PH_PY780"            '월고용보험내역
    oFilter.AddEx "PH_PY785"            '월국민연금내역
    oFilter.AddEx "PH_PY790"            '월건강보험내역
    oFilter.AddEx "PH_PY795"            '연간부서별급여내역
    oFilter.AddEx "PH_PY800"            '인건비지급자료
    oFilter.AddEx "PH_PY805"            '급여수당변동내역
    oFilter.AddEx "PH_PY810"            '직급별통상임금내역
    oFilter.AddEx "PH_PY815"            '평균임금내역
    oFilter.AddEx "PH_PY820"            '통상임금내역
    oFilter.AddEx "PH_PY825"            '전문직O/T현황
    oFilter.AddEx "PH_PY830"            '부서별인건비현황 (기획)
    oFilter.AddEx "PH_PY835"            '직급별O/T및수당현황
    oFilter.AddEx "PH_PY840"            '풍산전자공시자료
    oFilter.AddEx "PH_PY845"            '기간별급여지급내역
    oFilter.AddEx "PH_PY850"            '소급분지급명세서
    oFilter.AddEx "PH_PY855"            '개인별임금지급대장
    oFilter.AddEx "PH_PY865"            '고용보험현황 (계산용)
    oFilter.AddEx "PH_PY870"            '담당별월O/T및수당현황
    oFilter.AddEx "PH_PY875"            '직급별수당집계대장
    oFilter.AddEx "PH_PY716"            '기간별급여부서별집계대장
    oFilter.AddEx "PH_PY721"            '기간별상여부서별집계대장
    oFilter.AddEx "PH_PY717"            '기간별급여반별집계대장
    oFilter.AddEx "PH_PY718"            '생산완료금액대비O/T현황
    oFilter.AddEx "PH_PY701"            '급여지급대장 (노조용)
    
    oFilter.AddEx "PH_PYA10"            '급여지급대장(부서)
    oFilter.AddEx "PH_PYA20"            '급여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA30"            '상여지급대장(부서)
    oFilter.AddEx "PH_PYA40"            '상여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA50"            'DC전환자부담금지급내역
    oFilter.AddEx "PH_PYA75"            '교통비외수당지급대장
    
    '//정산관리
    oFilter.AddEx "PH_PY402"            '정산기초자료등록
    oFilter.AddEx "PH_PY405"            '의료비등록
    oFilter.AddEx "PH_PY407"            '기부금등록
    oFilter.AddEx "PH_PY409"            '기부금조정명세등록
    oFilter.AddEx "PH_PY411"            '연금.저축등소득공제등록
    oFilter.AddEx "PH_PY413"            '월세액.주택임차차입금자료 등록
    
    oFilter.AddEx "PH_PY910"            '소득공제신고서출력
    oFilter.AddEx "PH_PY915"            '근로소득원천징수부출력
    oFilter.AddEx "PH_PY920"            '원천징수영수증출력
    oFilter.AddEx "PH_PY925"            '기부금명세서출력
    oFilter.AddEx "PH_PY930"            '정산징수및환급대장
    oFilter.AddEx "PH_PY931"            '표준세액적용대상자조회
    oFilter.AddEx "PH_PY932"            '전근무지등록현황
    oFilter.AddEx "PH_PY933"            '보수총액신고기초자료
    oFilter.AddEx "PH_PYA55"            '정산징수및환급대장(집계)
    oFilter.AddEx "PH_PYA70"            '소득세원천징수세액조정신청서출력
    
    
    oFilter.AddEx "PH_PY980"            '근로소득지급명세서_전산매체자료작성
    oFilter.AddEx "PH_PY985"            '의료비지급명세서_전산매체자료작성
    oFilter.AddEx "PH_PY990"            '기부금지급명세서_전산매체자료작성
    oFilter.AddEx "PH_PY995"            '퇴직소득지급명세서_전산매체자료작성
    
    oFilter.AddEx "ZPY341"              '월별 정산자료 생성
    oFilter.AddEx "ZPY343"              '월별 자료 관리
    oFilter.AddEx "ZPY421"              '퇴직소득전산매체수록
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    oFilter.AddEx "ZPY503"              '정산세액계산
    
    '//기타관리
    oFilter.AddEx "PH_PY301"            '학자금신청등록
    oFilter.AddEx "PH_PY307"            '학자금신청내역(분기별)
    oFilter.AddEx "PH_PY030"            '공용등록
    oFilter.AddEx "PH_PY031"            '출장등록
    oFilter.AddEx "PH_PY032"            '사용외출등록
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)
    
End Sub

Private Sub CLICK(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '6
   Set oFilter = oFilters.Add(et_CLICK)
   
    
    '//System Form Type
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    oFilter.AddEx "PH_PY000"            '사용자권한관리
    
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터 등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    oFilter.AddEx "PH_PY004"            '근무조편성등록
    oFilter.AddEx "PH_PY005"            '사업장정보등록
    oFilter.AddEx "PH_PY006"            '승호작업등록
    oFilter.AddEx "PH_PY007"            '유류단가등록
    oFilter.AddEx "PH_PY008"            '일근태등록
    oFilter.AddEx "PH_PY009"            '기찰자료UPLOAD
    oFilter.AddEx "PH_PY011"            '전문직 호칭 일괄 변경(2013.07.05 송명규 추가)
    oFilter.AddEx "PH_PY013"            '위해일수계산
    oFilter.AddEx "PH_PY014"            '위해일수수정
    oFilter.AddEx "PH_PY015"            '연차적치등록
    oFilter.AddEx "PH_PY016"            '기본업무등록
    oFilter.AddEx "PH_PY017"            '월근태집계
    oFilter.AddEx "PH_PY018"            '휴일근무체크(연봉제)
    oFilter.AddEx "PH_PY019"            '반변경등록
    oFilter.AddEx "PH_PY020"            '일근태 업무변경등록
    oFilter.AddEx "PH_PY021"            '사원비상연락처관리
    
    
    
    oFilter.AddEx "PH_PY202"            '정년임박자 휴가경비 등록 조회
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회
    
    '//인사 - 리포트
    oFilter.AddEx "PH_PY501"            '여권발급현황
    oFilter.AddEx "PH_PY505"            '입사자대장
    oFilter.AddEx "PH_PY510"            '사원명부
    oFilter.AddEx "PH_PY515"            '재직자사원명부
    oFilter.AddEx "PH_PY520"            '퇴직및퇴직예정자대장
    oFilter.AddEx "PH_PY525"            '학력별인원현황
    oFilter.AddEx "PH_PY530"            '연령별인원현황
    oFilter.AddEx "PH_PY535"            '근속년수별인원현황
    oFilter.AddEx "PH_PY540"            '인원현황(대외용)
    oFilter.AddEx "PH_PY545"            '인원현황(대내용)
    oFilter.AddEx "PH_PY550"            '전체인원현황
    oFilter.AddEx "PH_PY555"            '일일근무자현황
    oFilter.AddEx "PH_PY560"            '일출근현황
    oFilter.AddEx "PH_PY565"            '연장근무자현황
    oFilter.AddEx "PH_PY570"            '연장/휴일근무자현황
    oFilter.AddEx "PH_PY575"            '근태기찰현황
    oFilter.AddEx "PH_PY580"            '개인별근태월보
    oFilter.AddEx "PH_PY585"            '일일출근기록부
    oFilter.AddEx "PH_PY590"            '기간별근태집계표
    oFilter.AddEx "PH_PY595"            '근속년수현황
    oFilter.AddEx "PH_PY600"            '일자별연장근무현황
    oFilter.AddEx "PH_PY605"            '근속보전휴가발생및사용내역
    oFilter.AddEx "PH_PY610"            '근태구분별사용내역
    oFilter.AddEx "PH_PY615"            '당직근무현황
    oFilter.AddEx "PH_PY620"            '연봉제휴일근무자현황
    oFilter.AddEx "PH_PY635"            '여행,교육자현황
    oFilter.AddEx "PH_PY640"            '국민연금퇴직전환금현황
    oFilter.AddEx "PH_PY645"            '자격수당지급현황
    oFilter.AddEx "PH_PY650"            '노동조합간부현황
    oFilter.AddEx "PH_PY655"            '보훈대상자현황
    oFilter.AddEx "PH_PY660"            '장애근로자현황
    oFilter.AddEx "PH_PY665"            '사원자녀현황
    oFilter.AddEx "PH_PY670"            '개인별차량현황
    oFilter.AddEx "PH_PY675"            '근무편성현황
    oFilter.AddEx "PH_PY676"            '근태시간내역조회
    oFilter.AddEx "PH_PY677"            '일일근태이상자조회
    oFilter.AddEx "PH_PY679"            '개인별 근태집계 조회
    oFilter.AddEx "PH_PY680"            '상벌현황
    oFilter.AddEx "PH_PY685"            '포상가급현황
    oFilter.AddEx "PH_PY690"            '생일자현황
    oFilter.AddEx "PH_PY695"            '인사기록카드
    oFilter.AddEx "PH_PY705"            '교통비지급근태확인
    oFilter.AddEx "PH_PY860"            '호봉표조회
    oFilter.AddEx "PH_PY503"            '승진대상자명부
    oFilter.AddEx "PH_PY678"            '당직근무자 일괄 등록
    oFilter.AddEx "PH_PY507"            '휴직자현황
    oFilter.AddEx "PH_PY681"            '비근무일수현황
    oFilter.AddEx "PH_PY935"            '정기승호현황
    oFilter.AddEx "PH_PY551"            '평균인원조회
    oFilter.AddEx "PH_PY508"            '재직증명 등록 및 발급
    oFilter.AddEx "PH_PY522"            '임금피크대상자현황
    oFilter.AddEx "PH_PY523"            '임금피크대상자월별차수현황
    oFilter.AddEx "PH_PY524"            '퇴직금 중간 정산내역
    oFilter.AddEx "PH_PY683"            '교대근무인정현황
    oFilter.AddEx "PH_PYA65"            '년차현황 (집계)
    oFilter.AddEx "PH_PY583"            '개인별 근태집계 조회
    
    '//급여관리
    oFilter.AddEx "PH_PY100"            '기준세액설정
    oFilter.AddEx "PH_PY101"            '보험률등록
    oFilter.AddEx "PH_PY102"            '수당항목설정
    oFilter.AddEx "PH_PY103"            '공제항목설정
    oFilter.AddEx "PH_PY104"            '고정수당공제금액일괄등록
    oFilter.AddEx "PH_PY105"            '호봉표등록
    oFilter.AddEx "PH_PY106"            '수당계산식설정
    oFilter.AddEx "PH_PY107"            '급상여기준일설정
    oFilter.AddEx "PH_PY108"            '상여율지급설정
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    oFilter.AddEx "PH_PY109_1"          '급상여변동자료 항목수정
    oFilter.AddEx "PH_PY110"            '개인상여율등록
    oFilter.AddEx "PH_PY111"            '급상여계산
    oFilter.AddEx "PH_PY112"            '급상여자료관리
    oFilter.AddEx "PH_PY113"            '급상여분개자료생성
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY115"            '퇴직금계산
    oFilter.AddEx "PH_PY116"            '퇴직금분개자료생성
    oFilter.AddEx "PH_PY117"            '급상여마감작업
    oFilter.AddEx "PH_PY118"            '급상여Email발송
    oFilter.AddEx "PH_PY119"            '급상여은행파일생성
    oFilter.AddEx "PH_PY120"            '급상여소급집계처리
    oFilter.AddEx "PH_PY121"            '평가가급액 등록
    oFilter.AddEx "PH_PY122"            '급상여출력 개인부서설정등록
    oFilter.AddEx "PH_PY123"            '가압류등록
    oFilter.AddEx "PH_PY125"            '퇴직연금 설정
    oFilter.AddEx "PH_PY127"            '//개인별 4대보험 보수월액 및 정산금액입력
    oFilter.AddEx "PH_PY130"            '팀별 성과급차등 등급등록
    oFilter.AddEx "PH_PY131"            '성과급차등 계수등록
    oFilter.AddEx "PH_PY132"            '성과급차 개인별 계산
    oFilter.AddEx "PH_PY133"            '연봉제 횟차 관리
    oFilter.AddEx "PH_PY134"            '소득세/주민세 조정관리
    oFilter.AddEx "PH_PY129"            '개인별퇴직연금(DC형) 계산
    
    
    
    '//급여관리 - 리포트
    oFilter.AddEx "PH_PY625"            '세탁자명부
    oFilter.AddEx "PH_PY630"            '사원별노조비공제현황
    oFilter.AddEx "PH_PY700"            '급여지급대장
    oFilter.AddEx "PH_PY710"            '상여지급대장
    oFilter.AddEx "PH_PY715"            '급여부서별집계대장
    oFilter.AddEx "PH_PY720"            '상여부서별집계대장
    oFilter.AddEx "PH_PY725"            '급여직급별집계대장
    oFilter.AddEx "PH_PY740"            '상여직급별집계대장
    oFilter.AddEx "PH_PY730"            '급여봉투출력
    oFilter.AddEx "PH_PY735"            '상여봉투출력
    oFilter.AddEx "PH_PY745"            '연간지급현황
    oFilter.AddEx "PH_PY750"            '근로소득징수현황
    oFilter.AddEx "PH_PY755"            '동호회가입현황
    oFilter.AddEx "PH_PY760"            '평균임금및퇴직금산출내역서
    oFilter.AddEx "PH_PY765"            '급여증감내역서
    oFilter.AddEx "PH_PY770"            '퇴직소득원천징수영수증출력
    oFilter.AddEx "PH_PY775"            '개인별년차현황
    oFilter.AddEx "PH_PY776"            '잔여년차현황
    oFilter.AddEx "PH_PY780"            '월고용보험내역
    oFilter.AddEx "PH_PY785"            '월국민연금내역
    oFilter.AddEx "PH_PY790"            '월건강보험내역
    oFilter.AddEx "PH_PY795"            '연간부서별급여내역
    oFilter.AddEx "PH_PY800"            '인건비지급자료
    oFilter.AddEx "PH_PY805"            '급여수당변동내역
    oFilter.AddEx "PH_PY810"            '직급별통상임금내역
    oFilter.AddEx "PH_PY815"            '평균임금내역
    oFilter.AddEx "PH_PY820"            '통상임금내역
    oFilter.AddEx "PH_PY825"            '전문직O/T현황
    oFilter.AddEx "PH_PY830"            '부서별인건비현황 (기획)
    oFilter.AddEx "PH_PY835"            '직급별O/T및수당현황
    oFilter.AddEx "PH_PY840"            '풍산전자공시자료
    oFilter.AddEx "PH_PY845"            '기간별급여지급내역
    oFilter.AddEx "PH_PY850"            '소급분지급명세서
    oFilter.AddEx "PH_PY855"            '개인별임금지급대장
    oFilter.AddEx "PH_PY865"            '고용보험현황 (계산용)
    oFilter.AddEx "PH_PY870"            '담당별월O/T및수당현황
    oFilter.AddEx "PH_PY875"            '직급별수당집계대장
    oFilter.AddEx "PH_PY716"            '기간별급여부서별집계대장
    oFilter.AddEx "PH_PY721"            '기간별상여부서별집계대장
    oFilter.AddEx "PH_PY717"            '기간별급여반별집계대장
    oFilter.AddEx "PH_PY718"            '생산완료금액대비O/T현황
    oFilter.AddEx "PH_PY701"            '급여지급대장 (노조용)
    
    oFilter.AddEx "PH_PYA10"            '급여지급대장(부서)
    oFilter.AddEx "PH_PYA20"            '급여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA30"            '상여지급대장(부서)
    oFilter.AddEx "PH_PYA40"            '상여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA50"            'DC전환자부담금지급내역
    oFilter.AddEx "PH_PYA75"            '교통비외수당지급대장
    
    
    '//정산관리
    oFilter.AddEx "PH_PY401"            '전근무지등록
    oFilter.AddEx "PH_PY402"            '정산기초자료 등록
    oFilter.AddEx "PH_PY405"            '의료비등록
    oFilter.AddEx "PH_PY407"            '기부금등록
    oFilter.AddEx "PH_PY409"            '기부금조정명세등록
    oFilter.AddEx "PH_PY411"            '연금.저축등소득공제등록
    oFilter.AddEx "PH_PY413"            '월세액.주택임차차입금자료 등록
    oFilter.AddEx "PH_PY980"            '신고_근로소득지급명세서자료작성
    oFilter.AddEx "PH_PY985"            '신고_의료비지급명세서자료작성
    oFilter.AddEx "PH_PY990"            '신고_기부금명세서자료작성
    oFilter.AddEx "PH_PY995"            '신고_퇴직소득지급명세서자료작성
    oFilter.AddEx "PH_PY419"            '표준세액적용대상자등록
    
    oFilter.AddEx "PH_PY910"            '소득공제신고서출력
    oFilter.AddEx "PH_PY915"            '근로소득원천징수부출력
    oFilter.AddEx "PH_PY920"            '원천징수영수증출력
    oFilter.AddEx "PH_PY925"            '기부금명세서출력
    oFilter.AddEx "PH_PY930"            '정산징수및환급대장
    oFilter.AddEx "PH_PY931"            '표준세액적용대상자조회
    oFilter.AddEx "PH_PY932"            '전근무지등록현황
    oFilter.AddEx "PH_PY933"            '보수총액신고기초자료
    oFilter.AddEx "PH_PYA55"            '정산징수및환급대장(집계)
    oFilter.AddEx "PH_PYA70"            '소득세원천징수세액조정신청서출력
    
    
    oFilter.AddEx "ZPY341"              '월별 정산자료 생성
    oFilter.AddEx "ZPY421"              '퇴직소득전산매체수록
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    oFilter.AddEx "ZPY502"              '종(전) 근무지 등록
    oFilter.AddEx "ZPY503"              '정산세액계산
    oFilter.AddEx "ZPY504"              '정산결과조회
    oFilter.AddEx "ZPY505"              '기부금명세등록
    oFilter.AddEx "ZPY506"              '의료비명세등록
    oFilter.AddEx "ZPY508"              '연금저축 소득공제 명세 등록
    oFilter.AddEx "ZPY509"              '정산자료 마감작업
    oFilter.AddEx "ZPY521"              '근로소득전산매체수록
    oFilter.AddEx "ZPY522"              '의료비 기부금 전산매체수록
    
    '//기타관리
    oFilter.AddEx "PH_PY307"            '학자금신청내역(분기별)
    oFilter.AddEx "PH_PY309"            '대부금등록
    oFilter.AddEx "PH_PY311"            '통근버스운행등록
    oFilter.AddEx "PH_PY312"            '버스요금 개인별등록
    
    oFilter.AddEx "PH_PY030"            '공용등록
    oFilter.AddEx "PH_PY031"            '출장등록
    oFilter.AddEx "PH_PY032"            '사용외출등록
    oFilter.AddEx "PH_PY315"            '개인별대부금잔액현황
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)
    
End Sub

Private Sub DOUBLE_CLICK(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '7
    Set oFilter = oFilters.Add(et_DOUBLE_CLICK)

    
    '//System Form Type
    
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    
    '//AddOn Form Type
    '//운영관리
    oFilter.AddEx "PH_PY000"            '사용자권한관리
    '//인사관리
    '//급여관리
    oFilter.AddEx "PH_PY104"            '고정수당공제금액 일괄등록
    oFilter.AddEx "PH_PY118"            '급상여Email발송
    '//정산관리
    oFilter.AddEx "PH_PY402"              '정산기초자료등록
    '//기타관리
    
End Sub

Private Sub MATRIX_LINK_PRESSED(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '8
    Set oFilter = oFilters.Add(et_MATRIX_LINK_PRESSED)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    oFilter.AddEx "ZPY507"              '정산결과조회(전체)
    
    '//기타관리
End Sub

Private Sub MATRIX_COLLAPSE_PRESSED(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '9
    Set oFilter = oFilters.Add(et_MATRIX_COLLAPSE_PRESSED)

End Sub

Private Sub VALIDATE(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '10
    Set oFilter = oFilters.Add(et_VALIDATE)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    oFilter.AddEx "PH_PY000"            '사용자권한관리
    
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터 등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    oFilter.AddEx "PH_PY005"            '사업장정보등록
    oFilter.AddEx "PH_PY007"            '유류단가등록
    oFilter.AddEx "PH_PY008"            '일근태등록
    oFilter.AddEx "PH_PY012"            '출장등록
    oFilter.AddEx "PH_PY013"            '위해일수계산
    oFilter.AddEx "PH_PY014"            '위해일수수정
    oFilter.AddEx "PH_PY015"            '연차적치등록
    oFilter.AddEx "PH_PY016"            '기본업무등록
    oFilter.AddEx "PH_PY017"            '월근태집계
    oFilter.AddEx "PH_PY018"            '휴일근무체크(연봉제)
    oFilter.AddEx "PH_PY019"            '반변경등록
    oFilter.AddEx "PH_PY020"            '일근태 업무변경등록
    oFilter.AddEx "PH_PY021"            '사원비상연락처관리

    oFilter.AddEx "PH_PY201"            '정년임박자 휴가경비 등록.
    oFilter.AddEx "PH_PY202"            '정년임박자 휴가경비 조회.
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회

   '//인사 - 리포트
    oFilter.AddEx "PH_PY501"            '여권발급현황
    oFilter.AddEx "PH_PY505"            '입사자대장
    oFilter.AddEx "PH_PY510"            '사원명부
    oFilter.AddEx "PH_PY515"            '재직자사원명부
    oFilter.AddEx "PH_PY520"            '퇴직및퇴직예정자대장
    oFilter.AddEx "PH_PY525"            '학력별인원현황
    oFilter.AddEx "PH_PY530"            '연령별인원현황
    oFilter.AddEx "PH_PY535"            '근속년수별인원현황
    oFilter.AddEx "PH_PY540"            '인원현황(대외용)
    oFilter.AddEx "PH_PY545"            '인원현황(대내용)
    oFilter.AddEx "PH_PY550"            '전체인원현황
    oFilter.AddEx "PH_PY555"            '일일근무자현황
    oFilter.AddEx "PH_PY560"            '일출근현황
    oFilter.AddEx "PH_PY565"            '연장근무자현황
    oFilter.AddEx "PH_PY570"            '연장/휴일근무자현황
    oFilter.AddEx "PH_PY575"            '근태기찰현황
    oFilter.AddEx "PH_PY580"            '개인별근태월보
    oFilter.AddEx "PH_PY585"            '일일출근기록부
    oFilter.AddEx "PH_PY590"            '기간별근태집계표
    oFilter.AddEx "PH_PY595"            '근속년수현황
    oFilter.AddEx "PH_PY600"            '일자별연장근무현황
    oFilter.AddEx "PH_PY605"            '근속보전휴가발생및사용내역
    oFilter.AddEx "PH_PY610"            '근태구분별사용내역
    oFilter.AddEx "PH_PY615"            '당직근무현황
    oFilter.AddEx "PH_PY620"            '연봉제휴일근무자현황
    oFilter.AddEx "PH_PY635"            '여행,교육자현황
    oFilter.AddEx "PH_PY640"            '국민연금퇴직전환금현황
    oFilter.AddEx "PH_PY645"            '자격수당지급현황
    oFilter.AddEx "PH_PY650"            '노동조합간부현황
    oFilter.AddEx "PH_PY655"            '보훈대상자현황
    oFilter.AddEx "PH_PY660"            '장애근로자현황
    oFilter.AddEx "PH_PY665"            '사원자녀현황
    oFilter.AddEx "PH_PY670"            '개인별차량현황
    oFilter.AddEx "PH_PY675"            '근무편성현황
    oFilter.AddEx "PH_PY676"            '근태시간내역조회
    oFilter.AddEx "PH_PY677"            '일일근태이상자조회
    oFilter.AddEx "PH_PY679"            '개인별 근태집계 조회
    oFilter.AddEx "PH_PY680"            '상벌현황
    oFilter.AddEx "PH_PY685"            '포상가급현황
    oFilter.AddEx "PH_PY690"            '생일자현황
    oFilter.AddEx "PH_PY695"            '인사기록카드
    oFilter.AddEx "PH_PY705"            '교통비지급근태확인
    oFilter.AddEx "PH_PY860"            '호봉표조회
    oFilter.AddEx "PH_PY503"            '승진대상자명부
    oFilter.AddEx "PH_PY678"            '당직근무자 일괄 등록
    oFilter.AddEx "PH_PY507"            '휴직자현황
    oFilter.AddEx "PH_PY681"            '비근무일수현황
    oFilter.AddEx "PH_PY935"            '정기승호현황
    oFilter.AddEx "PH_PY551"            '평균인원조회
    oFilter.AddEx "PH_PY508"            '재직증명 등록 및 발급
    oFilter.AddEx "PH_PY522"            '임금피크대상자현황
    oFilter.AddEx "PH_PY523"            '임금피크대상자월별차수현황
    oFilter.AddEx "PH_PY524"            '퇴직금 중간 정산내역
    oFilter.AddEx "PH_PY683"            '교대근무인정현황
    oFilter.AddEx "PH_PYA65"            '년차현황 (집계)
    oFilter.AddEx "PH_PY583"            '개인별 근태집계 조회
    
    '//급여관리
    oFilter.AddEx "PH_PY100"            '기준세액설정
    oFilter.AddEx "PH_PY101"            '보험률등록
    oFilter.AddEx "PH_PY102"            '수당항목설정
    oFilter.AddEx "PH_PY103"            '공제항목설정
    oFilter.AddEx "PH_PY104"            '고정수당공제금액일괄등록
    oFilter.AddEx "PH_PY105"            '호봉표등록
    oFilter.AddEx "PH_PY106"            '수당계산식설정
    oFilter.AddEx "PH_PY107"            '급상여기준일설정
    oFilter.AddEx "PH_PY108"            '상여율지급설정
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    oFilter.AddEx "PH_PY109_1"          '급상여변동자료 항목수정
    oFilter.AddEx "PH_PY110"            '개인상여율등록
    oFilter.AddEx "PH_PY111"            '급상여계산
    oFilter.AddEx "PH_PY112"            '급상여자료관리
    oFilter.AddEx "PH_PY113"            '급상여분개자료생성
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY115"            '퇴직금계산
    oFilter.AddEx "PH_PY116"            '퇴직금분개자료생성
    oFilter.AddEx "PH_PY117"            '급상여마감작업
    oFilter.AddEx "PH_PY118"            '급상여Email발송
    oFilter.AddEx "PH_PY120"            '급상여소급집계처리
    oFilter.AddEx "PH_PY121"            '평가가급액 등록
    oFilter.AddEx "PH_PY122"            '급상여출력 개인부서설정등록
    oFilter.AddEx "PH_PY123"            '가압류등록
    oFilter.AddEx "PH_PY130"            '팀별 성과급차등 등급등록
    oFilter.AddEx "PH_PY131"            '성과급차등 계수등록
    oFilter.AddEx "PH_PY132"            '성과급차 개인별 계산
    oFilter.AddEx "PH_PY133"            '연봉제 횟차 관리
    oFilter.AddEx "PH_PY134"            '소득세/주민세 조정관리
    oFilter.AddEx "PH_PY129"            '개인별퇴직연금(DC형) 계산
    
    '//급여관리 - 리포트
    oFilter.AddEx "PH_PY625"            '세탁자명부
    oFilter.AddEx "PH_PY630"            '사원별노조비공제현황
    oFilter.AddEx "PH_PY700"            '급여지급대장
    oFilter.AddEx "PH_PY710"            '상여지급대장
    oFilter.AddEx "PH_PY715"            '급여부서별집계대장
    oFilter.AddEx "PH_PY720"            '상여부서별집계대장
    oFilter.AddEx "PH_PY725"            '급여직급별집계대장
    oFilter.AddEx "PH_PY740"            '상여직급별집계대장
    oFilter.AddEx "PH_PY730"            '급여봉투출력
    oFilter.AddEx "PH_PY735"            '상여봉투출력
    oFilter.AddEx "PH_PY745"            '연간지급현황
    oFilter.AddEx "PH_PY750"            '근로소득징수현황
    oFilter.AddEx "PH_PY755"            '동호회가입현황
    oFilter.AddEx "PH_PY760"            '평균임금및퇴직금산출내역서
    oFilter.AddEx "PH_PY765"            '급여증감내역서
    oFilter.AddEx "PH_PY770"            '퇴직소득원천징수영수증출력
    oFilter.AddEx "PH_PY775"            '개인별년차현황
    oFilter.AddEx "PH_PY776"            '잔여년차현황
    oFilter.AddEx "PH_PY780"            '월고용보험내역
    oFilter.AddEx "PH_PY785"            '월국민연금내역
    oFilter.AddEx "PH_PY790"            '월건강보험내역
    oFilter.AddEx "PH_PY795"            '연간부서별급여내역
    oFilter.AddEx "PH_PY800"            '인건비지급자료
    oFilter.AddEx "PH_PY805"            '급여수당변동내역
    oFilter.AddEx "PH_PY810"            '직급별통상임금내역
    oFilter.AddEx "PH_PY815"            '평균임금내역
    oFilter.AddEx "PH_PY820"            '통상임금내역
    oFilter.AddEx "PH_PY825"            '전문직O/T현황
    oFilter.AddEx "PH_PY830"            '부서별인건비현황 (기획)
    oFilter.AddEx "PH_PY835"            '직급별O/T및수당현황
    oFilter.AddEx "PH_PY840"            '풍산전자공시자료
    oFilter.AddEx "PH_PY845"            '기간별급여지급내역
    oFilter.AddEx "PH_PY850"            '소급분지급명세서
    oFilter.AddEx "PH_PY855"            '개인별임금지급대장
    oFilter.AddEx "PH_PY865"            '고용보험현황 (계산용)
    oFilter.AddEx "PH_PY870"            '담당별월O/T및수당현황
    oFilter.AddEx "PH_PY875"            '직급별수당집계대장
    oFilter.AddEx "PH_PY716"            '기간별급여부서별집계대장
    oFilter.AddEx "PH_PY721"            '기간별상여부서별집계대장
    oFilter.AddEx "PH_PY717"            '기간별급여반별집계대장
    oFilter.AddEx "PH_PY718"            '생산완료금액대비O/T현황
    oFilter.AddEx "PH_PY701"            '급여지급대장 (노조용)
    
    oFilter.AddEx "PH_PYA10"            '급여지급대장(부서)
    oFilter.AddEx "PH_PYA20"            '급여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA30"            '상여지급대장(부서)
    oFilter.AddEx "PH_PYA40"            '상여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA50"            'DC전환자부담금지급내역
    oFilter.AddEx "PH_PYA75"            '교통비외수당지급대장
    
    '//정산관리
    oFilter.AddEx "PH_PY401"            '전근무지등록
    oFilter.AddEx "PH_PY402"            '정산기초자료 등록
    oFilter.AddEx "PH_PY405"            '의료비등록
    oFilter.AddEx "PH_PY407"            '기부금등록
    oFilter.AddEx "PH_PY409"            '기부금조정명세등록
    oFilter.AddEx "PH_PY411"            '연금.저축등소득공제등록
    oFilter.AddEx "PH_PY413"            '월세액.주택임차차입금자료 등록
    oFilter.AddEx "PH_PY415"            '정산계산
    oFilter.AddEx "PH_PY417"            '정산 은행파일생성
    oFilter.AddEx "PH_PY980"            '신고_근로소득지급명세서자료작성
    oFilter.AddEx "PH_PY985"            '신고_의료비지급명세서자료작성
    oFilter.AddEx "PH_PY990"            '신고_기부금명세서자료작성
    oFilter.AddEx "PH_PY995"            '신고_퇴직소득지급명세서자료작성
    oFilter.AddEx "PH_PY419"            '표준세액적용대상자등록
    
    oFilter.AddEx "PH_PY910"            '소득공제신고서출력
    oFilter.AddEx "PH_PY915"            '근로소득원천징수부출력
    oFilter.AddEx "PH_PY920"            '원천징수영수증출력
    oFilter.AddEx "PH_PY925"            '기부금명세서출력
    oFilter.AddEx "PH_PY930"            '정산징수및환급대장
    oFilter.AddEx "PH_PY931"            '표준세액적용대상자조회
    oFilter.AddEx "PH_PY932"            '전근무지등록현황
    oFilter.AddEx "PH_PY933"            '보수총액신고기초자료
    oFilter.AddEx "PH_PYA55"            '정산징수및환급대장(집계)
    oFilter.AddEx "PH_PYA70"            '소득세원천징수세액조정신청서출력
    
    
    oFilter.AddEx "ZPY341"              '월별 정산자료 생성
    oFilter.AddEx "ZPY343"              '월별 자료 관리
    oFilter.AddEx "ZPY421"              '퇴직소득전산매체수록
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    oFilter.AddEx "ZPY502"              '종(전) 근무지 등록
    oFilter.AddEx "ZPY503"              '정산세액계산
    oFilter.AddEx "ZPY504"              '정산결과조회
    oFilter.AddEx "ZPY505"              '기부금명세등록
    oFilter.AddEx "ZPY506"              '의료비명세등록
    oFilter.AddEx "ZPY507"              '정산결과조회(전체)
    oFilter.AddEx "ZPY508"              '연금저축 소득공제 명세 등록
    oFilter.AddEx "ZPY509"              '정산자료 마감작업
    oFilter.AddEx "ZPY510"              '종전근무지 일괄생성
    oFilter.AddEx "ZPY521"              '근로소득전산매체수록
    oFilter.AddEx "ZPY522"              '의료비 기부금 전산매체수록
    
    '//기타관리
    oFilter.AddEx "PH_PY301"            '학자금신청등록
    oFilter.AddEx "PH_PY305"            '학자금신청서
    oFilter.AddEx "PH_PY306"            '학자금신청내역(개인별)
    oFilter.AddEx "PH_PY309"            '대부금등록
    oFilter.AddEx "PH_PY310"            '대부금개별상환
    oFilter.AddEx "PH_PY311"            '통근버스운행등록
    oFilter.AddEx "PH_PY313"            '대부금계산
    oFilter.AddEx "PH_PY314"            '대부금계산 내역 조회(급여변동자료용)
    oFilter.AddEx "PH_PY030"            '공용등록
    oFilter.AddEx "PH_PY031"            '출장등록
    oFilter.AddEx "PH_PY032"            '사용외출등록
    oFilter.AddEx "PH_PY315"            '개인별대부금잔액현황
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)
    
End Sub

Private Sub MATRIX_LOAD(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '11
    Set oFilter = oFilters.Add(et_MATRIX_LOAD)
    
    
    '//System Form Type
    
    '//운영관리
    '//판매관리
    '//구매관리
    '//재고관리
    '//생산관리
    
    
    '//AddOn Form Type
    '//운영관리
    oFilter.AddEx "PH_PY000"            '사용자권한관리
    
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터 등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    oFilter.AddEx "PH_PY018"            '휴일근무체크(연봉제)

    oFilter.AddEx "PH_PY201"            '정년임박자 휴가경비 등록
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회

    '//급여관리
    oFilter.AddEx "PH_PY100"            '기준세액설정
    oFilter.AddEx "PH_PY101"            '보험률등록
    oFilter.AddEx "PH_PY102"            '수당항목설정
    oFilter.AddEx "PH_PY103"            '공제항목설정
    oFilter.AddEx "PH_PY105"            '호봉등록표
    oFilter.AddEx "PH_PY106"            '수당계산식설정
    oFilter.AddEx "PH_PY107"            '급상여기준일설정
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    oFilter.AddEx "PH_PY109_1"          '급상여변동자료 항목수정
    oFilter.AddEx "PH_PY110"            '개인상여율등록
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY121"            '평가가급액 등록
    oFilter.AddEx "PH_PY122"            '급상여출력 개인부서설정등록
    oFilter.AddEx "PH_PY130"            '팀별 성과급차등 등급등록
    oFilter.AddEx "PH_PY131"            '성과급차등 계수등록
    oFilter.AddEx "PH_PY132"            '성과급차 개인별 계산
    oFilter.AddEx "PH_PY133"             '연봉제 횟차 관리
    oFilter.AddEx "PH_PY134"            '소득세/주민세 조정관리
    oFilter.AddEx "PH_PY129"            '개인별퇴직연금(DC형) 계산
    
    '//정산관리
    oFilter.AddEx "ZPY343"              '월별 자료 관리
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    oFilter.AddEx "ZPY502"              '종(전) 근무지 등록
    oFilter.AddEx "ZPY505"              '기부금명세등록
    oFilter.AddEx "ZPY506"              '의료비명세등록
    oFilter.AddEx "ZPY508"              '연금저축 소득공제 명세 등록
    oFilter.AddEx "ZPY509"              '정산자료 마감작업
    
    
    '//기타관리
    oFilter.AddEx "PH_PY301"            '학자금신청등록
    oFilter.AddEx "PH_PY309"            '대부금등록
    oFilter.AddEx "PH_PY310"            '대부금개별상환
    oFilter.AddEx "PH_PY311"            '통근버스운행등록
    oFilter.AddEx "PH_PY313"            '대부금계산
    oFilter.AddEx "PH_PY012"            '출장등록
    oFilter.AddEx "PH_PY315"            '개인별대부금잔액현황
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)
    
End Sub

Private Sub DATASOURCE_LOAD(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '12
    Set oFilter = oFilters.Add(et_DATASOURCE_LOAD)
    
    
    '//System Form Type
    
    '//운영관리
    '//판매관리
    '//구매관리
    '//재고관리
    '//생산관리

    
    '//AddOn Form Type
    
    '//운영관리
    '//판매관리
    '//구매관리
    '//재고관리
    '//생산관리
    
End Sub

Private Sub Form_Load(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '16
    Set oFilter = oFilters.Add(et_FORM_LOAD)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    oFilter.AddEx "PH_PY000"            '사용자권한관리
    
'    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터 등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    oFilter.AddEx "PH_PY004"            '근무조편성등록
    oFilter.AddEx "PH_PY005"            '사업장정보등록
    oFilter.AddEx "PH_PY006"            '승호작업등록
    oFilter.AddEx "PH_PY007"            '유류단가등록
    oFilter.AddEx "PH_PY008"            '일근태등록
    oFilter.AddEx "PH_PY009"            '기찰자료UPLOAD
    oFilter.AddEx "PH_PY010"            '일일근태처리
    oFilter.AddEx "PH_PY011"            '전문직 호칭 일괄 변경(2013.07.05 송명규 추가)
    oFilter.AddEx "PH_PY013"            '위해일수계산
    oFilter.AddEx "PH_PY014"            '위해일수수정
    oFilter.AddEx "PH_PY015"            '연차적치등록
    oFilter.AddEx "PH_PY016"            '기본업무등록
    oFilter.AddEx "PH_PY017"            '월근태집계
    oFilter.AddEx "PH_PY018"            '휴일근무체크(연봉제)
    oFilter.AddEx "PH_PY019"            '반변경등록
    oFilter.AddEx "PH_PY020"            '일근태 업무변경등록
    oFilter.AddEx "PH_PY021"            '사원비상연락처관리
    oFilter.AddEx "PH_PY201"            '정년임박자 휴가경비 등록
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회
    
    '//인사 - 리포트
    oFilter.AddEx "PH_PY501"            '여권발급현황
    oFilter.AddEx "PH_PY505"            '입사자대장
    oFilter.AddEx "PH_PY510"            '사원명부
    oFilter.AddEx "PH_PY515"            '재직자사원명부
    oFilter.AddEx "PH_PY520"            '퇴직및퇴직예정자대장
    oFilter.AddEx "PH_PY525"            '학력별인원현황
    oFilter.AddEx "PH_PY530"            '연령별인원현황
    oFilter.AddEx "PH_PY535"            '근속년수별인원현황
    oFilter.AddEx "PH_PY540"            '인원현황(대외용)
    oFilter.AddEx "PH_PY545"            '인원현황(대내용)
    oFilter.AddEx "PH_PY550"            '전체인원현황
    oFilter.AddEx "PH_PY555"            '일일근무자현황
    oFilter.AddEx "PH_PY560"            '일출근현황
    oFilter.AddEx "PH_PY565"            '연장근무자현황
    oFilter.AddEx "PH_PY570"            '연장/휴일근무자현황
    oFilter.AddEx "PH_PY575"            '근태기찰현황
    oFilter.AddEx "PH_PY580"            '개인별근태월보
    oFilter.AddEx "PH_PY585"            '일일출근기록부
    oFilter.AddEx "PH_PY590"            '기간별근태집계표
    oFilter.AddEx "PH_PY595"            '근속년수현황
    oFilter.AddEx "PH_PY600"            '일자별연장근무현황
    oFilter.AddEx "PH_PY605"            '근속보전휴가발생및사용내역
    oFilter.AddEx "PH_PY610"            '근태구분별사용내역
    oFilter.AddEx "PH_PY615"            '당직근무현황
    oFilter.AddEx "PH_PY620"            '연봉제휴일근무자현황
    oFilter.AddEx "PH_PY635"            '여행,교육자현황
    oFilter.AddEx "PH_PY640"            '국민연금퇴직전환금현황
    oFilter.AddEx "PH_PY645"            '자격수당지급현황
    oFilter.AddEx "PH_PY650"            '노동조합간부현황
    oFilter.AddEx "PH_PY655"            '보훈대상자현황
    oFilter.AddEx "PH_PY660"            '장애근로자현황
    oFilter.AddEx "PH_PY665"            '사원자녀현황
    oFilter.AddEx "PH_PY670"            '개인별차량현황
    oFilter.AddEx "PH_PY675"            '근무편성현황
    oFilter.AddEx "PH_PY676"            '근태시간내역조회
    oFilter.AddEx "PH_PY677"            '일일근태이상자조회
    oFilter.AddEx "PH_PY679"            '개인별 근태집계 조회
    oFilter.AddEx "PH_PY680"            '상벌현황
    oFilter.AddEx "PH_PY685"            '포상가급현황
    oFilter.AddEx "PH_PY690"            '생일자현황
    oFilter.AddEx "PH_PY695"            '인사기록카드
    oFilter.AddEx "PH_PY705"            '교통비지급근태확인
    oFilter.AddEx "PH_PY860"            '호봉표조회
    oFilter.AddEx "PH_PY503"            '승진대상자명부
    oFilter.AddEx "PH_PY678"            '당직근무자 일괄 등록
    oFilter.AddEx "PH_PY507"            '휴직자현황
    oFilter.AddEx "PH_PY681"            '비근무일수현황
    oFilter.AddEx "PH_PY935"            '정기승호현황
    oFilter.AddEx "PH_PY551"            '평균인원조회
    oFilter.AddEx "PH_PY508"            '재직증명 등록 및 발급
    oFilter.AddEx "PH_PY522"            '임금피크대상자현황
    oFilter.AddEx "PH_PY523"            '임금피크대상자월별차수현황
    oFilter.AddEx "PH_PY524"            '퇴직금 중간 정산 내역
    oFilter.AddEx "PH_PY683"            '교대근무인정현황
    oFilter.AddEx "PH_PYA65"            '년차현황 (집계)
    oFilter.AddEx "PH_PY583"            '개인별 근태집계 조회
    
    '//급여관리
    oFilter.AddEx "PH_PY100"            '기준세액설정
    oFilter.AddEx "PH_PY101"            '보험률등록
    oFilter.AddEx "PH_PY102"            '수당항목설정
    oFilter.AddEx "PH_PY103"            '공제항목설정
    oFilter.AddEx "PH_PY104"            '고정수당공제금액일괄등록
    oFilter.AddEx "PH_PY105"            '호봉표등록
    oFilter.AddEx "PH_PY106"            '수당계산식설정
    oFilter.AddEx "PH_PY107"            '급상여기준일설정
    oFilter.AddEx "PH_PY108"            '상여율지급설정
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    oFilter.AddEx "PH_PY109_1"          '급상여변동자료 항목수정
    oFilter.AddEx "PH_PY110"            '개인상여율등록
    oFilter.AddEx "PH_PY111"            '급상여계산
    oFilter.AddEx "PH_PY112"            '급상여자료관리
    oFilter.AddEx "PH_PY113"            '급상여분개자료생성
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY115"            '퇴직금계산
    oFilter.AddEx "PH_PY116"            '퇴직금분개자료생성
    oFilter.AddEx "PH_PY117"            '급상여마감작업
    oFilter.AddEx "PH_PY118"            '급상여Email발송
    oFilter.AddEx "PH_PY119"            '급상여은행파일생성
    oFilter.AddEx "PH_PY120"            '급상여소급집계처리
    oFilter.AddEx "PH_PY121"            '평가가급액 등록
    oFilter.AddEx "PH_PY122"            '급상여출력 개인부서설정등록
    oFilter.AddEx "PH_PY123"            '가압류등록
    oFilter.AddEx "PH_PY125"            '퇴직연금 설정
    oFilter.AddEx "PH_PY127"            '//개인별 4대보험 보수월액 및 정산금액입력
    oFilter.AddEx "PH_PY130"            '팀별 성과급차등 등급등록
    oFilter.AddEx "PH_PY131"            '성과급차등 계수등록
    oFilter.AddEx "PH_PY132"            '성과급차 개인별 계산
    oFilter.AddEx "PH_PY133"            '연봉제 횟차 관리
    oFilter.AddEx "PH_PY134"            '소득세/주민세 조정관리
    oFilter.AddEx "PH_PY129"            '개인별퇴직연금(DC형) 계산
    
    '//급여관리 - 리포트
    oFilter.AddEx "PH_PY625"            '세탁자명부
    oFilter.AddEx "PH_PY630"            '사원별노조비공제현황
    oFilter.AddEx "PH_PY700"            '급여지급대장
    oFilter.AddEx "PH_PY710"            '상여지급대장
    oFilter.AddEx "PH_PY715"            '급여부서별집계대장
    oFilter.AddEx "PH_PY720"            '상여부서별집계대장
    oFilter.AddEx "PH_PY725"            '급여직급별집계대장
    oFilter.AddEx "PH_PY740"            '상여직급별집계대장
    oFilter.AddEx "PH_PY730"            '급여봉투출력
    oFilter.AddEx "PH_PY735"            '상여봉투출력
    oFilter.AddEx "PH_PY745"            '연간지급현황
    oFilter.AddEx "PH_PY750"            '근로소득징수현황
    oFilter.AddEx "PH_PY755"            '동호회가입현황
    oFilter.AddEx "PH_PY760"            '평균임금및퇴직금산출내역서
    oFilter.AddEx "PH_PY765"            '급여증감내역서
    oFilter.AddEx "PH_PY770"            '퇴직소득원천징수영수증출력
    oFilter.AddEx "PH_PY775"            '개인별년차현황
    oFilter.AddEx "PH_PY776"            '잔여년차현황
    oFilter.AddEx "PH_PY780"            '월고용보험내역
    oFilter.AddEx "PH_PY785"            '월국민연금내역
    oFilter.AddEx "PH_PY790"            '월건강보험내역
    oFilter.AddEx "PH_PY795"            '연간부서별급여내역
    oFilter.AddEx "PH_PY800"            '인건비지급자료
    oFilter.AddEx "PH_PY805"            '급여수당변동내역
    oFilter.AddEx "PH_PY810"            '직급별통상임금내역
    oFilter.AddEx "PH_PY815"            '평균임금내역
    oFilter.AddEx "PH_PY820"            '통상임금내역
    oFilter.AddEx "PH_PY825"            '전문직O/T현황
    oFilter.AddEx "PH_PY830"            '부서별인건비현황 (기획)
    oFilter.AddEx "PH_PY835"            '직급별O/T및수당현황
    oFilter.AddEx "PH_PY840"            '풍산전자공시자료
    oFilter.AddEx "PH_PY845"            '기간별급여지급내역
    oFilter.AddEx "PH_PY850"            '소급분지급명세서
    oFilter.AddEx "PH_PY855"            '개인별임금지급대장
    oFilter.AddEx "PH_PY865"            '고용보험현황 (계산용)
    oFilter.AddEx "PH_PY870"            '담당별월O/T및수당현황
    oFilter.AddEx "PH_PY875"            '직급별수당집계대장
    oFilter.AddEx "PH_PY716"            '기간별급여부서별집계대장
    oFilter.AddEx "PH_PY721"            '기간별상여부서별집계대장
    oFilter.AddEx "PH_PY717"            '기간별급여반별집계대장
    oFilter.AddEx "PH_PY718"            '생산완료금액대비O/T현황
    oFilter.AddEx "PH_PY701"            '급여지급대장 (노조용)
    
    oFilter.AddEx "PH_PYA10"            '급여지급대장(부서)
    oFilter.AddEx "PH_PYA20"            '급여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA30"            '상여지급대장(부서)
    oFilter.AddEx "PH_PYA40"            '상여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA50"            'DC전환자부담금지급내역
    oFilter.AddEx "PH_PYA75"            '교통비외수당지급대장
    
    '//정산관리
    oFilter.AddEx "PH_PY401"            '전근무지등록
    oFilter.AddEx "PH_PY402"            '정산기초자료 등록
    oFilter.AddEx "PH_PY405"            '의료비등록
    oFilter.AddEx "PH_PY407"            '기부금등록
    oFilter.AddEx "PH_PY409"            '기부금조정명세등록
    oFilter.AddEx "PH_PY411"            '연금.저축등소득공제등록
    oFilter.AddEx "PH_PY413"            '월세액.주택임차차입금자료 등록
    oFilter.AddEx "PH_PY415"            '정산계산
    oFilter.AddEx "PH_PY417"            '정산 은행파일생성
    oFilter.AddEx "PH_PY980"            '신고_근로소득지급명세서자료작성
    oFilter.AddEx "PH_PY985"            '신고_의료비지급명세서자료작성
    oFilter.AddEx "PH_PY990"            '신고_기부금명세서자료작성
    oFilter.AddEx "PH_PY995"            '신고_퇴직소득지급명세서자료작성
    oFilter.AddEx "PH_PY419"            '표준세액적용대상자등록
    
    oFilter.AddEx "PH_PY910"            '소득공제신고서출력
    oFilter.AddEx "PH_PY915"            '근로소득원천징수부출력
    oFilter.AddEx "PH_PY920"            '원천징수영수증출력
    oFilter.AddEx "PH_PY925"            '기부금명세서출력
    oFilter.AddEx "PH_PY930"            '정산징수및환급대장
    oFilter.AddEx "PH_PY931"            '표준세액적용대상자조회
    oFilter.AddEx "PH_PY932"            '전근무지등록현황
    oFilter.AddEx "PH_PY933"            '보수총액신고기초자료
    oFilter.AddEx "PH_PYA55"            '정산징수및환급대장(집계)
    oFilter.AddEx "PH_PYA70"            '소득세원천징수세액조정신청서출력
    
    
    oFilter.AddEx "ZPY341"              '월별 정산자료 생성
    oFilter.AddEx "ZPY343"              '월별 자료 관리
    oFilter.AddEx "ZPY421"              '퇴직소득전산매체수록
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    oFilter.AddEx "ZPY502"              '종(전) 근무지 등록
    oFilter.AddEx "ZPY503"              '정산세액계산
    oFilter.AddEx "ZPY504"              '정산결과조회
    oFilter.AddEx "ZPY505"              '기부금명세등록
    oFilter.AddEx "ZPY506"              '의료비명세등록
    oFilter.AddEx "ZPY507"              '정산결과조회(전체)
    oFilter.AddEx "ZPY508"              '연금저축 소득공제 명세 등록
    oFilter.AddEx "ZPY509"              '정산자료 마감작업
    oFilter.AddEx "ZPY510"              '종전근무지 일괄생성
    oFilter.AddEx "ZPY521"              '근로소득전산매체수록
    oFilter.AddEx "ZPY522"              '의료비 기부금 전산매체수록

    oFilter.AddEx "RPY401"              '퇴직원천징수 영수증
    oFilter.AddEx "RPY501"              '월별자료현황
    oFilter.AddEx "RPY502"              '종전근무지현황
    oFilter.AddEx "RPY503"              '근로소득 원천징수부
    oFilter.AddEx "RPY504"              '근로소득 원천영수증
    oFilter.AddEx "RPY505"              '소득자료집계표
    oFilter.AddEx "RPY506"              '정산징수환급대장
    oFilter.AddEx "RPY508"              '연말정산집계표
    oFilter.AddEx "RPY509"              '갑근세신고검토표
    oFilter.AddEx "RPY510"              '비과세근로소득명세서
    oFilter.AddEx "RPY511"              '기부금명세서
    '//기타관리
    oFilter.AddEx "PH_PY301"            '학자금신청등록
    oFilter.AddEx "PH_PY302"            '학자금지급완료처리
    oFilter.AddEx "PH_PY303"            '학자금은행파일생성
    oFilter.AddEx "PH_PY305"            '학자금신청서
    oFilter.AddEx "PH_PY306"            '학자금신청내역(개인별)
    oFilter.AddEx "PH_PY307"            '학자금신청내역(분기별)
    oFilter.AddEx "PH_PY309"            '대부금등록
    oFilter.AddEx "PH_PY310"            '대부금개별상환
    oFilter.AddEx "PH_PY311"            '통근버스운행등록
    oFilter.AddEx "PH_PY312"            '버스요금 개인별등록
    oFilter.AddEx "PH_PY313"            '대부금계산
    oFilter.AddEx "PH_PY030"            '공용등록
    oFilter.AddEx "PH_PY031"            '출장등록
    oFilter.AddEx "PH_PY032"            '사용외출등록
    oFilter.AddEx "PH_PY315"            '개인별대부금잔액현황
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)
    
End Sub

Private Sub FORM_UNLOAD(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '17
    Set oFilter = oFilters.Add(et_FORM_UNLOAD)

    
    '//System Form Type
    '//운영관리
    '//인사관리

    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    oFilter.AddEx "PH_PY000"            '사용자권한관리
    
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터 등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    oFilter.AddEx "PH_PY004"            '근무조편성등록
    oFilter.AddEx "PH_PY005"            '사업장정보등록
    oFilter.AddEx "PH_PY006"            '승호작업등록
    oFilter.AddEx "PH_PY007"            '유류단가등록
    oFilter.AddEx "PH_PY008"            '일근태등록
    oFilter.AddEx "PH_PY011"            '전문직 호칭 일괄 변경(2013.07.05 송명규 추가)
    oFilter.AddEx "PH_PY013"            '위해일수계산
    oFilter.AddEx "PH_PY014"            '위해일수수정
    oFilter.AddEx "PH_PY015"            '연차적치등록
    oFilter.AddEx "PH_PY016"            '기본업무등록
    oFilter.AddEx "PH_PY017"            '월근태집계
    oFilter.AddEx "PH_PY018"            '휴일근무체크(연봉제)
    oFilter.AddEx "PH_PY019"            '반변경등록
    oFilter.AddEx "PH_PY020"            '일근태 업무변경등록
    oFilter.AddEx "PH_PY021"            '사원비상연락처관리
    oFilter.AddEx "PH_PY201"            '정년임박자 휴가경비 등록
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회
    
    '//인사 - 리포트
    oFilter.AddEx "PH_PY501"            '여권발급현황
    oFilter.AddEx "PH_PY505"            '입사자대장
    oFilter.AddEx "PH_PY510"            '사원명부
    oFilter.AddEx "PH_PY515"            '재직자사원명부
    oFilter.AddEx "PH_PY520"            '퇴직및퇴직예정자대장
    oFilter.AddEx "PH_PY525"            '학력별인원현황
    oFilter.AddEx "PH_PY530"            '연령별인원현황
    oFilter.AddEx "PH_PY535"            '근속년수별인원현황
    oFilter.AddEx "PH_PY540"            '인원현황(대외용)
    oFilter.AddEx "PH_PY545"            '인원현황(대내용)
    oFilter.AddEx "PH_PY550"            '전체인원현황
    oFilter.AddEx "PH_PY555"            '일일근무자현황
    oFilter.AddEx "PH_PY560"            '일출근현황
    oFilter.AddEx "PH_PY565"            '연장근무자현황
    oFilter.AddEx "PH_PY570"            '연장/휴일근무자현황
    oFilter.AddEx "PH_PY575"            '근태기찰현황
    oFilter.AddEx "PH_PY580"            '개인별근태월보
    oFilter.AddEx "PH_PY585"            '일일출근기록부
    oFilter.AddEx "PH_PY590"            '기간별근태집계표
    oFilter.AddEx "PH_PY595"            '근속년수현황
    oFilter.AddEx "PH_PY600"            '일자별연장근무현황
    oFilter.AddEx "PH_PY605"            '근속보전휴가발생및사용내역
    oFilter.AddEx "PH_PY610"            '근태구분별사용내역
    oFilter.AddEx "PH_PY615"            '당직근무현황
    oFilter.AddEx "PH_PY620"            '연봉제휴일근무자현황
    oFilter.AddEx "PH_PY635"            '여행,교육자현황
    oFilter.AddEx "PH_PY640"            '국민연금퇴직전환금현황
    oFilter.AddEx "PH_PY645"            '자격수당지급현황
    oFilter.AddEx "PH_PY650"            '노동조합간부현황
    oFilter.AddEx "PH_PY655"            '보훈대상자현황
    oFilter.AddEx "PH_PY660"            '장애근로자현황
    oFilter.AddEx "PH_PY665"            '사원자녀현황
    oFilter.AddEx "PH_PY670"            '개인별차량현황
    oFilter.AddEx "PH_PY675"            '근무편성현황
    oFilter.AddEx "PH_PY676"            '근태시간내역조회
    oFilter.AddEx "PH_PY677"            '일일근태이상자조회
    oFilter.AddEx "PH_PY679"            '개인별 근태집계 조회
    oFilter.AddEx "PH_PY680"            '상벌현황
    oFilter.AddEx "PH_PY685"            '포상가급현황
    oFilter.AddEx "PH_PY690"            '생일자현황
    oFilter.AddEx "PH_PY695"            '인사기록카드
    oFilter.AddEx "PH_PY705"            '교통비지급근태확인
    oFilter.AddEx "PH_PY860"            '호봉표조회
    oFilter.AddEx "PH_PY503"            '승진대상자명부
    oFilter.AddEx "PH_PY678"            '당직근무자 일괄 등록
    oFilter.AddEx "PH_PY507"            '휴직자현황
    oFilter.AddEx "PH_PY681"            '비근무일수현황
    oFilter.AddEx "PH_PY935"            '정기승호현황
    oFilter.AddEx "PH_PY551"            '평균인원조회
    oFilter.AddEx "PH_PY508"            '재직증명 등록 및 발급
    oFilter.AddEx "PH_PY522"            '임금피크대상자현황
    oFilter.AddEx "PH_PY523"            '임금피크대상자월별차수현황
    oFilter.AddEx "PH_PY524"            '퇴직금 중간 정산내역
    oFilter.AddEx "PH_PY683"            '교대근무인정현황
    oFilter.AddEx "PH_PYA65"            '년차현황 (집계)
    oFilter.AddEx "PH_PY583"            '개인별 근태집계 조회
    
    '//급여관리
    oFilter.AddEx "PH_PY100"            '기준세액설정
    oFilter.AddEx "PH_PY101"            '보험률등록
    oFilter.AddEx "PH_PY102"            '수당항목설정
    oFilter.AddEx "PH_PY103"            '공제항목설정
    oFilter.AddEx "PH_PY104"            '고정수당공제금액일괄등록
    oFilter.AddEx "PH_PY105"            '호봉표등록
    oFilter.AddEx "PH_PY106"            '수당계산식설정
    oFilter.AddEx "PH_PY107"            '급상여기준일설정
    oFilter.AddEx "PH_PY108"            '상여율지급설정
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    oFilter.AddEx "PH_PY109_1"          '급상여변동자료 항목수정
    oFilter.AddEx "PH_PY110"            '개인상여율등록
    oFilter.AddEx "PH_PY111"            '급상여계산
    oFilter.AddEx "PH_PY112"            '급상여자료관리
    oFilter.AddEx "PH_PY113"            '급상여분개자료생성
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY115"            '퇴직금계산
    oFilter.AddEx "PH_PY116"            '퇴직금분개자료생성
    oFilter.AddEx "PH_PY117"            '급상여마감작업
    oFilter.AddEx "PH_PY118"            '급상여Email발송
    oFilter.AddEx "PH_PY119"            '급상여은행파일생성
    oFilter.AddEx "PH_PY120"            '급상여소급집계처리
    oFilter.AddEx "PH_PY121"            '평가가급액 등록
    oFilter.AddEx "PH_PY122"            '급상여출력 개인부서설정등록
    oFilter.AddEx "PH_PY123"            '가압류등록
    oFilter.AddEx "PH_PY125"            '퇴직연금 설정
    oFilter.AddEx "PH_PY127"            '//개인별 4대보험 보수월액 및 정산금액입력
    oFilter.AddEx "PH_PY130"            '팀별 성과급차등 등급등록
    oFilter.AddEx "PH_PY131"            '성과급차등 계수등록
    oFilter.AddEx "PH_PY132"            '성과급차 개인별 계산
    oFilter.AddEx "PH_PY133"            '연봉제 횟차 관리
    oFilter.AddEx "PH_PY134"            '소득세/주민세 조정관리
    oFilter.AddEx "PH_PY129"            '개인별퇴직연금(DC형) 계산
    
    '//급여관리 - 리포트
    oFilter.AddEx "PH_PY625"            '세탁자명부
    oFilter.AddEx "PH_PY630"            '사원별노조비공제현황
    oFilter.AddEx "PH_PY700"            '급여지급대장
    oFilter.AddEx "PH_PY710"            '상여지급대장
    oFilter.AddEx "PH_PY715"            '급여부서별집계대장
    oFilter.AddEx "PH_PY720"            '상여부서별집계대장
    oFilter.AddEx "PH_PY725"            '급여직급별집계대장
    oFilter.AddEx "PH_PY740"            '상여직급별집계대장
    oFilter.AddEx "PH_PY730"            '급여봉투출력
    oFilter.AddEx "PH_PY735"            '상여봉투출력
    oFilter.AddEx "PH_PY745"            '연간지급현황
    oFilter.AddEx "PH_PY750"            '근로소득징수현황
    oFilter.AddEx "PH_PY755"            '동호회가입현황
    oFilter.AddEx "PH_PY760"            '평균임금및퇴직금산출내역서
    oFilter.AddEx "PH_PY765"            '급여증감내역서
    oFilter.AddEx "PH_PY770"            '퇴직소득원천징수영수증출력
    oFilter.AddEx "PH_PY775"            '개인별년차현황
    oFilter.AddEx "PH_PY776"            '잔여년차현황
    oFilter.AddEx "PH_PY780"            '월고용보험내역
    oFilter.AddEx "PH_PY785"            '월국민연금내역
    oFilter.AddEx "PH_PY790"            '월건강보험내역
    oFilter.AddEx "PH_PY795"            '연간부서별급여내역
    oFilter.AddEx "PH_PY800"            '인건비지급자료
    oFilter.AddEx "PH_PY805"            '급여수당변동내역
    oFilter.AddEx "PH_PY810"            '직급별통상임금내역
    oFilter.AddEx "PH_PY815"            '평균임금내역
    oFilter.AddEx "PH_PY820"            '통상임금내역
    oFilter.AddEx "PH_PY825"            '전문직O/T현황
    oFilter.AddEx "PH_PY830"            '부서별인건비현황 (기획)
    oFilter.AddEx "PH_PY835"            '직급별O/T및수당현황
    oFilter.AddEx "PH_PY840"            '풍산전자공시자료
    oFilter.AddEx "PH_PY845"            '기간별급여지급내역
    oFilter.AddEx "PH_PY850"            '소급분지급명세서
    oFilter.AddEx "PH_PY855"            '개인별임금지급대장
    oFilter.AddEx "PH_PY865"            '고용보험현황 (계산용)
    oFilter.AddEx "PH_PY870"            '담당별월O/T및수당현황
    oFilter.AddEx "PH_PY875"            '직급별수당집계대장
    oFilter.AddEx "PH_PY716"            '기간별급여부서별집계대장
    oFilter.AddEx "PH_PY721"            '기간별상여부서별집계대장
    oFilter.AddEx "PH_PY717"            '기간별급여반별집계대장
    oFilter.AddEx "PH_PY718"            '생산완료금액대비O/T현황
    oFilter.AddEx "PH_PY701"            '급여지급대장 (노조용)
    
    oFilter.AddEx "PH_PYA10"            '급여지급대장(부서)
    oFilter.AddEx "PH_PYA20"            '급여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA30"            '상여지급대장(부서)
    oFilter.AddEx "PH_PYA40"            '상여부서별집계대장(부서)
    oFilter.AddEx "PH_PYA50"            'DC전환자부담금지급내역
    oFilter.AddEx "PH_PYA75"            '교통비외수당지급대장
    
    '//정산관리
    oFilter.AddEx "PH_PY401"            '전근무지등록
    oFilter.AddEx "PH_PY402"            '정산기초자료 등록
    oFilter.AddEx "PH_PY405"            '의료비등록
    oFilter.AddEx "PH_PY407"            '기부금등록
    oFilter.AddEx "PH_PY409"            '기부금조정명세등록
    oFilter.AddEx "PH_PY411"            '연금.저축등소득공제등록
    oFilter.AddEx "PH_PY413"            '월세액.주택임차차입금자료 등록
    oFilter.AddEx "PH_PY415"            '정산계산
    oFilter.AddEx "PH_PY417"            '정산 은행파일생성
    oFilter.AddEx "PH_PY980"            '신고_근로소득지급명세서자료작성
    oFilter.AddEx "PH_PY985"            '신고_의료비지급명세서자료작성
    oFilter.AddEx "PH_PY990"            '신고_기부금명세서자료작성
    oFilter.AddEx "PH_PY995"            '신고_퇴직소득지급명세서자료작성
    oFilter.AddEx "PH_PY419"            '표준세액적용대상자등록
    
    oFilter.AddEx "PH_PY910"            '소득공제신고서출력
    oFilter.AddEx "PH_PY915"            '근로소득원천징수부출력
    oFilter.AddEx "PH_PY920"            '원천징수영수증출력
    oFilter.AddEx "PH_PY925"            '기부금명세서출력
    oFilter.AddEx "PH_PY930"            '정산징수및환급대장
    oFilter.AddEx "PH_PY931"            '표준세액적용대상자조회
    oFilter.AddEx "PH_PY932"            '전근무지등록현황
    oFilter.AddEx "PH_PY933"            '보수총액신고기초자료
    oFilter.AddEx "PH_PYA55"            '정산징수및환급대장(집계)
    oFilter.AddEx "PH_PYA70"            '소득세원천징수세액조정신청서출력
    
    oFilter.AddEx "ZPY341"              '월별 정산자료 생성
    oFilter.AddEx "ZPY343"              '월별 자료 관리
    oFilter.AddEx "ZPY421"              '퇴직소득전산매체수록
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    oFilter.AddEx "ZPY502"              '종(전) 근무지 등록
    oFilter.AddEx "ZPY503"              '정산세액계산
    oFilter.AddEx "ZPY504"              '정산결과조회
    oFilter.AddEx "ZPY505"              '기부금명세등록
    oFilter.AddEx "ZPY506"              '의료비명세등록
    oFilter.AddEx "ZPY507"              '정산결과조회(전체)
    oFilter.AddEx "ZPY508"              '연금저축 소득공제 명세 등록
    oFilter.AddEx "ZPY509"              '정산자료 마감작업
    oFilter.AddEx "ZPY510"              '종전근무지 일괄생성
    oFilter.AddEx "ZPY521"              '근로소득전산매체수록
    oFilter.AddEx "ZPY522"              '의료비 기부금 전산매체수록
    
    oFilter.AddEx "RPY401"              '퇴직원천징수 연수증
    oFilter.AddEx "RPY501"              '월별자료현황
    oFilter.AddEx "RPY502"              '종전근무지현황
    oFilter.AddEx "RPY503"              '근로소득 원천징수부
    oFilter.AddEx "RPY504"              '근로소득 원천영수증
    oFilter.AddEx "RPY505"              '소득자료집계표
    oFilter.AddEx "RPY506"              '정산징수환급대장
    oFilter.AddEx "RPY508"              '연말정산집계표
    oFilter.AddEx "RPY509"              '갑근세신고검토표
    oFilter.AddEx "RPY510"              '비과세근로소득명세서
    oFilter.AddEx "RPY511"              '기부금명세서
    
    '//기타관리
    oFilter.AddEx "PH_PY301"            '학자금신청등록
    oFilter.AddEx "PH_PY302"            '학자금지급완료처리
    oFilter.AddEx "PH_PY303"            '학자금은행파일생성
    oFilter.AddEx "PH_PY305"            '학자금신청서
    oFilter.AddEx "PH_PY306"            '학자금신청내역(개인별)
    oFilter.AddEx "PH_PY307"            '학자금신청내역(분기별)
    oFilter.AddEx "PH_PY311"            '통근버스운행등록
    oFilter.AddEx "PH_PY312"            '버스요금 개인별등록
    oFilter.AddEx "PH_PY313"            '대부금계산
    oFilter.AddEx "PH_PY030"            '공용등록
    oFilter.AddEx "PH_PY031"            '출장등록
    oFilter.AddEx "PH_PY032"            '사용외출등록
    oFilter.AddEx "PH_PY315"            '개인별대부금잔액현황
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)
    
End Sub

Private Sub FORM_ACTIVATE(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '18
    Set oFilter = oFilters.Add(et_FORM_ACTIVATE)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
End Sub

Private Sub FORM_DEACTIVATE(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '19
    Set oFilter = oFilters.Add(et_FORM_DEACTIVATE)
    
    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
End Sub
Private Sub FORM_CLOSE(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '20
    Set oFilter = oFilters.Add(et_FORM_CLOSE)

End Sub

Private Sub Form_Resize(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '21
    Set oFilter = oFilters.Add(et_FORM_RESIZE)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    oFilter.AddEx "PH_PY007"            '유류단가등록
    oFilter.AddEx "PH_PY508"            '재직증명 등록 및 발급
    oFilter.AddEx "PH_PY021"            '사원비상연락처관리
    oFilter.AddEx "PH_PY201"            '정년임박자 휴가경비 등록
    oFilter.AddEx "PH_PY203"            '교육실적등록
    oFilter.AddEx "PH_PY204"            '교육계획등록
    oFilter.AddEx "PH_PY205"            '교육계획VS실적조회
    
    '//급여관리
    oFilter.AddEx "PH_PY100"            '기준세액설정
    oFilter.AddEx "PH_PY101"            '보험률등록
    oFilter.AddEx "PH_PY106"            '수당계산식설정
    oFilter.AddEx "PH_PY114"            '퇴직금기준설정
    oFilter.AddEx "PH_PY130"            '팀별 성과급차등 등급등록
    oFilter.AddEx "PH_PY131"            '성과급차등 계수등록
    oFilter.AddEx "PH_PY132"            '성과급차 개인별 계산
    oFilter.AddEx "PH_PY133"            '연봉제 횟차 관리
    oFilter.AddEx "PH_PY134"            '소득세/주민세 조정관리
    oFilter.AddEx "PH_PY129"            '개인별퇴직연금(DC형) 계산
    
    '//정산관리
    oFilter.AddEx "ZPY501"              '소득공제항목 등록
    
    '//기타관리
    oFilter.AddEx "PH_PY301"            '학자금신청등록
    oFilter.AddEx "PH_PY302"            '학자금지급완료처리
    oFilter.AddEx "PH_PY305"            '학자금신청서
    oFilter.AddEx "PH_PY306"            '학자금신청내역(개인별)
    oFilter.AddEx "PH_PY307"            '학자금신청내역(분기별)
    oFilter.AddEx "PH_PY032"            '사용외출등록
    oFilter.AddEx "PH_PY034"            '공용분개처리
    oFilter.AddEx "PH_PYA60"            '학자금신청내역(집계)
    
    '//근태관리
    oFilter.AddEx "PH_PY677"            '근태기찰이상자 수정
    
End Sub

Private Sub FORM_KEY_DOWN(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '22
    Set oFilter = oFilters.Add(et_FORM_KEY_DOWN)
    
    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리

End Sub
Private Sub FORM_MENU_HILIGHT(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '23
    Set oFilter = oFilters.Add(et_FORM_MENU_HILIGHT)
    
    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
End Sub

Private Sub vPRINT(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '24
    Set oFilter = oFilters.Add(et_PRINT)
    
    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리

    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
End Sub

Private Sub PRINT_DATA(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '25
    Set oFilter = oFilters.Add(et_PRINT_DATA)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리

    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
End Sub

Private Sub CHOOSE_FROM_LIST(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '27
    Set oFilter = oFilters.Add(et_CHOOSE_FROM_LIST)
    
    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터등록
    oFilter.AddEx "PH_PY005"            '사업장정보등록
    '//급여관리
    oFilter.AddEx "PH_PY103"            '공제항목설정

    
    '//정산관리
    '//기타관리
End Sub

Private Sub RIGHT_CLICK(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '28
    Set oFilter = oFilters.Add(et_RIGHT_CLICK)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리

    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY003"            '근태월력설정
    '//급여관리
    oFilter.AddEx "PH_PY109"            '급상여변동자료등록
    
    
    '//정산관리
    '//기타관리
    

End Sub

Private Sub MENU_CLICK(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '32
    Set oFilter = oFilters.Add(et_MENU_CLICK)
    
    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
End Sub

Private Sub FORM_DATA_ADD(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '33
    Set oFilter = oFilters.Add(et_FORM_DATA_ADD)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리

    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
        
End Sub

Private Sub FORM_DATA_UPDATE(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '34
    Set oFilter = oFilters.Add(et_FORM_DATA_UPDATE)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
End Sub

Private Sub FORM_DATA_DELETE(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '35
    Set oFilter = oFilters.Add(et_FORM_DATA_DELETE)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리

    
    '//AddOn Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
End Sub

Private Sub FORM_DATA_LOAD(ByRef oFilter As SAPbouiCOM.EventFilter, ByRef oFilters As SAPbouiCOM.EventFilters)  '36
    Set oFilter = oFilters.Add(et_FORM_DATA_LOAD)

    
    '//System Form Type
    '//운영관리
    '//인사관리
    '//급여관리
    
    '//정산관리
    '//기타관리
    
    '//AddOn Form Type
    '//운영관리
    oFilter.AddEx "PH_PY000"            '사용자권한관리
    
    '//인사관리
    oFilter.AddEx "PH_PY001"            '인사마스터등록
    oFilter.AddEx "PH_PY002"            '근태시간구분 등록
    oFilter.AddEx "PH_PY105"            '호봉표등록
    '//급여관리
    oFilter.AddEx "PH_PY112"            '급상여자료관리
    '//정산관리
    '//기타관리
        
End Sub




