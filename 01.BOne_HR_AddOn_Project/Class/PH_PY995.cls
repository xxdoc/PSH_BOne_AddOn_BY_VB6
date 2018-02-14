VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY995"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : PH_PY995.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 퇴직소득지급명세서자료 전산매체수록
'//  FormType       :
'//  Create Date    : 2014.02.04
'//  Modified Date  : 2015.02.10
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
'2014년기준 1090 BYTE
'2016년기준 1110 BYTE

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
        A016   As String * 932   '공란
End Type
Dim A_rec As A_record


Private Type B_record
        B001   As String * 1     '레코드구분
        B002   As String * 2     '자료구분
        B003   As String * 3     '세무서
        B004   As String * 6     '일련번호
        B005   As String * 10    '사업자번호
        B006   As String * 40    '법인명(상호)
        B007   As String * 30    '대표자(성명)
        B008   As String * 13    '주민(법인)등록번호
        B009   As String * 1     '제출대상기간코드
        B010   As String * 7     '퇴직소득자(C레코드)수
        B011   As String * 7     '공란
        B012   As String * 14    '정산-과세대상퇴직금여합계
        B013_1 As String * 1     '부호
        B013_2 As String * 13    '신고대상소득세합계
        B014   As String * 13    '이연퇴직소득세액합계
        B015_1 As String * 1     '부호
        B015_2 As String * 13    '차감원천징수-소득세액합계
        B016_1 As String * 1     '부호
        B016_2 As String * 13    '차감원천징수-지방소득세액합계
        B017_1 As String * 1     '부호
        B017_2 As String * 13    '차감원천징수-농어촌특별세액합계
        B018_1 As String * 1     '부호
        B018_2 As String * 13    '차감원천징수세액-계합계
        B019   As String * 893  '공란
End Type
Dim B_rec As B_record


Private Type C_record
        C001   As String * 1     '레코드구분
        C002   As String * 2     '자료구분
        C003   As String * 3     '세무서
        C004   As String * 6     '일련번호
        C005   As String * 10    '사업자번호
        C006   As String * 1     '징수의무자구분 1.사업장 3.공적연금사업자
        '/소득자
        C007   As String * 1     '거주구분 1.거주자 2.비거주자
        C008   As String * 1     '내.외국인구분 1.내국인 9.외국인
        C009   As String * 2     '거주지국코드
        C010   As String * 30    '성명
        C011   As String * 13    '주민등록번호
        C012   As String * 1     '임원여부  1.여 2.부
        C013   As String * 8     '확정급여형퇴직연금제도가입일
        C014   As String * 11    '2011.12.31 퇴직금
        C015   As String * 8     '귀속연도시작연월일
        C016   As String * 8     '귀속연도종료연월일
        C017   As String * 1     '퇴직사유
        '/퇴직급여현황-중간지급등
        C018   As String * 40    '근무처명
        C019   As String * 10    '근무처사업자등록번호
        C020   As String * 11    '퇴직급여
        C021   As String * 11    '비과세퇴직급여
        C022   As String * 11    '과세대상퇴직급여
        '/퇴직급여현황-최종분
        C023   As String * 40    '근무처명
        C024   As String * 10    '근무처사업자등록번호
        C025   As String * 11    '퇴직급여
        C026   As String * 11    '비과세퇴직급여
        C027   As String * 11    '과세대상퇴직급여
        '/퇴직급여현황-정산
        C028   As String * 11    '퇴직급여
        C029   As String * 11    '비과세퇴직급여
        C030   As String * 11    '과세대상퇴직급여
        '/근속연수-중간지급등
        C031   As String * 8     '입사일
        C032   As String * 8     '기산일
        C033   As String * 8     '퇴사일
        C034   As String * 8     '지급일
        C035   As String * 4     '근속월수
        C036   As String * 4     '제외월수
        C037   As String * 4     '가산월수
        C038   As String * 4     '중복월수
        C039   As String * 4     '근속연수
        '/근속연수-최종
        C040   As String * 8     '입사일
        C041   As String * 8     '기산일
        C042   As String * 8     '퇴사일
        C043   As String * 8     '지급일
        C044   As String * 4     '근속월수
        C045   As String * 4     '제외월수
        C046   As String * 4     '가산월수
        C047   As String * 4     '중복월수
        C048   As String * 4     '근속연수
        '/근속연수-정산
        C049   As String * 8     '입사일
        C050   As String * 8     '기산일
        C051   As String * 8     '퇴사일
        C052   As String * 8     '지급일
        C053   As String * 4     '근속월수
        C054   As String * 4     '제외월수
        C055   As String * 4     '가산월수
        C056   As String * 4     '중복월수
        C057   As String * 4     '근속연수
        '/근속연수-안분-2012.12.31이전
        C058   As String * 8     '입사일
        C059   As String * 8     '기산일
        C060   As String * 8     '퇴사일
        C061   As String * 8     '지급일
        C062   As String * 4     '근속월수
        C063   As String * 4     '제외월수
        C064   As String * 4     '가산월수
        C065   As String * 4     '중복월수
        C066   As String * 4     '근속연수
        '/근속연수-안분-2013.01.01이후
        C067   As String * 8     '입사일
        C068   As String * 8     '기산일
        C069   As String * 8     '퇴사일
        C070   As String * 8     '지급일
        C071   As String * 4     '근속월수
        C072   As String * 4     '제외월수
        C073   As String * 4     '가산월수
        C074   As String * 4     '중복월수
        C075   As String * 4     '근속연수
        '/개정규정에따른계산방법-과세표준계산 2016
        C076   As String * 11    '퇴직소득
        C077   As String * 11    '근속연수공제
        C078   As String * 11    '환산급여
        C079   As String * 11    '환산급여별공제
        C080   As String * 11    '퇴직소득과세표준
        '/개정규정에따른계산방법-세액계산 2016
        C081   As String * 11    '환산산출세액
        C082   As String * 11    '산출세액
        '/종전규정에 따른 계산 방법 - 과세표준계산
        C083   As String * 11    '퇴직소득
        C084   As String * 11    '퇴직소득정률공제
        C085   As String * 11    '근속연수공제
        C086   As String * 11    '퇴직소득과세표준
        '/종전규정에 따른 계산 방법 - 세액계산 - 2012.12.31 이전
        C087   As String * 11    '과세표준안분
        C088   As String * 11    '연평균과세표준
        C089   As String * 11    '환산과세표준
        C090   As String * 11    '환산산출세액
        C091   As String * 11    '연평균산출세액
        C092   As String * 11    '산출세액
        '/종전규정에 따른 계산 방법 - 세액계산 - 2013.1.1 이후
        C093   As String * 11    '과세표준안분
        C094   As String * 11    '연평균과세표준
        C095   As String * 11    '환산과세표준
        C096   As String * 11    '환산산출세액
        C097   As String * 11    '연평균산출세액
        C098   As String * 11    '산출세액
        '/종전규정에 따른 계산 방법 - 세액계산 - 합계
        C099   As String * 11    '과세표준안분
        C100   As String * 11    '연평균과세표준
        C101   As String * 11    '환산과세표준
        C102   As String * 11    '환산산출세액
        C103   As String * 11    '연평균산출세액
        C104   As String * 11    '산출세액
        '/퇴직소득세액계산
        C105   As String * 4     '퇴직일이속하는과세연도
        C106   As String * 11    '퇴직소득세산출세액
        C107   As String * 11    '기납부(또는기과세이연)세액
        C108_1 As String * 1     '부호
        C108_2 As String * 11    '신고대상세액
        '/이연퇴직소득세액계산
        C109_1 As String * 1     '부호
        C109_2 As String * 11    '신고대상세액
        C110   As String * 11    '계좌입금금액_합계
        C111   As String * 11    '퇴직급여
        C112   As String * 11    '이연퇴직소득세
        '/납부명세-신고대상세액
        C113_1 As String * 1     '부호
        C113_2 As String * 11    '소득세
        C114_1 As String * 1     '부호
        C114_2 As String * 11    '지방소득세
        C115_1 As String * 1     '부호
        C115_2 As String * 11    '농어총특별세
        C116_1 As String * 1     '부호
        C116_2 As String * 11    '계
        '/납부명세-이연퇴직소득세
        C117   As String * 11    '소득세
        C118   As String * 11    '지방소득세
        C119   As String * 11    '농어촌특별세
        C120   As String * 11    '계
        '/납부명세-차감원천징수세액
        C121_1 As String * 1     '부호
        C121_2 As String * 11    '소득세
        C122_1 As String * 1     '부호
        C122_2 As String * 11    '지방소득세
        C123_1 As String * 1     '부호
        C123_2 As String * 11    '농어총특별세
        C124_1 As String * 1     '부호
        C124_2 As String * 11    '계
        C125   As String * 2     '공란
End Type

Dim C_rec As C_record

'2014년 D레코드 추가

Private Type D_Record
        D001   As String * 1     '레코드구분
        D002   As String * 2     '자료구분
        D003   As String * 3     '세무서
        D004   As String * 6     '일련번호
        '/원천징수의무자
        D005   As String * 10    '사업자번호
        D006   As String * 50    '공란
        '/소득자
        D007   As String * 13    '소득자주민등록번호
        '/연금계좌입금명세
        D008   As String * 2     '연금계좌_일련번호
        D009   As String * 30    '연금계좌_취급자
        D010   As String * 10    '연금계좌_사업자등록번호
        D011   As String * 20    '연금계좌_계좌번호
        D012   As String * 8     '연금계좌_입금일
        D013   As String * 11    '연금계좌_계좌입금금액
        D014   As String * 944   '공란
End Type
Dim D_rec As D_Record

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY995.srf")
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
    
    oFormUniqueID = "PH_PY995_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "PH_PY995"
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
                 If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                 If pval.ItemUID = "Btn01" Then
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
    
    '/ C RECORD 처리
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
        Sbo_Application.StatusBar.SetText "A레코드(제출자 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "B레코드(원천징수의무자별 집계 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "C레코드(퇴직소득자 레코드) 생성 실패.", bmt_Short, smt_Error
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
    sQry = "EXEC PH_PY995_A '" & CLTCOD & "', '" & HtaxID & "', '" & TeamName & "', '" & Dname & "', '" & Dtel & "', '" & DocDate & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        ' PATH및 파일이름 만들기
        saup = oRecordSet.Fields("A009").VALUE  '사업자번호
        oFilePath = "C:\BANK\EA" + Mid$(saup, 1, 7) + "." + Mid$(saup, 8, 3)
        
        
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
    sQry = "EXEC PH_PY995_B '" & CLTCOD & "', '" & yyyy & "'"
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
        B_rec.B009 = oRecordSet.Fields("B009").VALUE
        B_rec.B010 = Format$(oRecordSet.Fields("B010").VALUE, String$(Len(B_rec.B010), "0"))
        B_rec.B011 = oRecordSet.Fields("B011").VALUE
        B_rec.B012 = Format$(oRecordSet.Fields("B012").VALUE, String$(Len(B_rec.B012), "0"))
        B_rec.B013_1 = oRecordSet.Fields("B013_1").VALUE
        B_rec.B013_2 = Format$(oRecordSet.Fields("B013_2").VALUE, String$(Len(B_rec.B013_2), "0"))
        B_rec.B014 = Format$(oRecordSet.Fields("B014").VALUE, String$(Len(B_rec.B014), "0"))
        B_rec.B015_1 = oRecordSet.Fields("B015_1").VALUE
        B_rec.B015_2 = Format$(oRecordSet.Fields("B015_2").VALUE, String$(Len(B_rec.B015_2), "0"))
        B_rec.B016_1 = oRecordSet.Fields("B016_1").VALUE
        B_rec.B016_2 = Format$(oRecordSet.Fields("B016_2").VALUE, String$(Len(B_rec.B016_2), "0"))
        B_rec.B017_1 = oRecordSet.Fields("B017_1").VALUE
        B_rec.B017_2 = Format$(oRecordSet.Fields("B017_2").VALUE, String$(Len(B_rec.B017_2), "0"))
        B_rec.B018_1 = oRecordSet.Fields("B018_1").VALUE
        B_rec.B018_2 = Format$(oRecordSet.Fields("B018_2").VALUE, String$(Len(B_rec.B018_2), "0"))
        B_rec.B019 = oRecordSet.Fields("B019").VALUE
        
        Print #1, MDC_SetMod.sStr(B_rec.B001) & MDC_SetMod.sStr(B_rec.B002) & MDC_SetMod.sStr(B_rec.B003) & MDC_SetMod.sStr(B_rec.B004) & _
                  MDC_SetMod.sStr(B_rec.B005) & MDC_SetMod.sStr(B_rec.B006) & MDC_SetMod.sStr(B_rec.B007) & MDC_SetMod.sStr(B_rec.B008) & _
                  MDC_SetMod.sStr(B_rec.B009) & MDC_SetMod.sStr(B_rec.B010) & MDC_SetMod.sStr(B_rec.B011) & MDC_SetMod.sStr(B_rec.B012) & _
                  MDC_SetMod.sStr(B_rec.B013_1) & MDC_SetMod.sStr(B_rec.B013_2) & MDC_SetMod.sStr(B_rec.B014) & MDC_SetMod.sStr(B_rec.B015_1) & _
                  MDC_SetMod.sStr(B_rec.B015_2) & MDC_SetMod.sStr(B_rec.B016_1) & MDC_SetMod.sStr(B_rec.B016_2) & _
                  MDC_SetMod.sStr(B_rec.B017_1) & MDC_SetMod.sStr(B_rec.B017_2) & MDC_SetMod.sStr(B_rec.B018_1) & MDC_SetMod.sStr(B_rec.B018_2) & MDC_SetMod.sStr(B_rec.B019)
                  
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
    Dim PSABUN      As String
    Dim OLDBIG      As Double
    Dim PILTOT      As Double
    Dim SCount      As Integer
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    CheckC = False  '/체크필요유무
    ErrNum = 0
    
    '/ C_RECORE QUERY
    sQry = "EXEC PH_PY995_C '" & CLTCOD & "', '" & yyyy & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("작성시작!", oRecordSet.RecordCount, False)
    
    NEWCNT = 0
    
    Do Until oRecordSet.EOF
        
        'C RECORD MOVE
        
        NEWCNT = NEWCNT + 1     '/ 일련번호
        
        '--D RECORD 호출용----------
        C_SAUP = oRecordSet.Fields("saup").VALUE
        C_YYYY = oRecordSet.Fields("yyyy").VALUE
        C_SABUN = oRecordSet.Fields("sabun").VALUE
        '--------------
        
        C_rec.C001 = oRecordSet.Fields("C001").VALUE
        C_rec.C002 = oRecordSet.Fields("C002").VALUE
        C_rec.C003 = oRecordSet.Fields("C003").VALUE
        C_rec.C004 = Format$(NEWCNT, String$(Len(C_rec.C004), "0"))  '/ 일련번호
        C_rec.C005 = oRecordSet.Fields("C005").VALUE
        C_rec.C006 = oRecordSet.Fields("C006").VALUE
        C_rec.C007 = oRecordSet.Fields("C007").VALUE
        C_rec.C008 = oRecordSet.Fields("C008").VALUE
        C_rec.C009 = oRecordSet.Fields("C009").VALUE
        C_rec.C010 = oRecordSet.Fields("C010").VALUE
        C_rec.C011 = oRecordSet.Fields("C011").VALUE
        C_rec.C012 = oRecordSet.Fields("C012").VALUE
        C_rec.C013 = Format$(oRecordSet.Fields("C013").VALUE, String$(Len(C_rec.C013), "0"))
        C_rec.C014 = Format$(oRecordSet.Fields("C014").VALUE, String$(Len(C_rec.C014), "0"))
        C_rec.C015 = Format$(oRecordSet.Fields("C015").VALUE, String$(Len(C_rec.C015), "0"))
        C_rec.C016 = Format$(oRecordSet.Fields("C016").VALUE, String$(Len(C_rec.C016), "0"))
        C_rec.C017 = oRecordSet.Fields("C017").VALUE
        C_rec.C018 = oRecordSet.Fields("C018").VALUE
        C_rec.C019 = oRecordSet.Fields("C019").VALUE
        C_rec.C020 = Format$(oRecordSet.Fields("C020").VALUE, String$(Len(C_rec.C020), "0"))
        C_rec.C021 = Format$(oRecordSet.Fields("C021").VALUE, String$(Len(C_rec.C021), "0"))
        C_rec.C022 = Format$(oRecordSet.Fields("C022").VALUE, String$(Len(C_rec.C022), "0"))
        C_rec.C023 = oRecordSet.Fields("C023").VALUE
        C_rec.C024 = oRecordSet.Fields("C024").VALUE
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
        C_rec.C057 = Format$(oRecordSet.Fields("C057").VALUE, String$(Len(C_rec.C057), "0"))
        C_rec.C058 = Format$(oRecordSet.Fields("C058").VALUE, String$(Len(C_rec.C058), "0"))
        C_rec.C059 = Format$(oRecordSet.Fields("C059").VALUE, String$(Len(C_rec.C059), "0"))
        C_rec.C060 = Format$(oRecordSet.Fields("C060").VALUE, String$(Len(C_rec.C060), "0"))
        C_rec.C061 = Format$(oRecordSet.Fields("C061").VALUE, String$(Len(C_rec.C061), "0"))
        C_rec.C062 = Format$(oRecordSet.Fields("C062").VALUE, String$(Len(C_rec.C062), "0"))
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
        C_rec.C095 = Format$(oRecordSet.Fields("C095").VALUE, String$(Len(C_rec.C095), "0"))
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
        C_rec.C106 = Format$(oRecordSet.Fields("C106").VALUE, String$(Len(C_rec.C106), "0"))
        C_rec.C107 = Format$(oRecordSet.Fields("C107").VALUE, String$(Len(C_rec.C107), "0"))
        C_rec.C108_1 = Format$(oRecordSet.Fields("C108_1").VALUE, String$(Len(C_rec.C108_1), "0"))
        C_rec.C108_2 = Format$(oRecordSet.Fields("C108_2").VALUE, String$(Len(C_rec.C108_2), "0"))
        C_rec.C109_1 = Format$(oRecordSet.Fields("C109_1").VALUE, String$(Len(C_rec.C109_1), "0"))
        C_rec.C109_2 = Format$(oRecordSet.Fields("C109_2").VALUE, String$(Len(C_rec.C109_2), "0"))
        C_rec.C110 = Format$(oRecordSet.Fields("C110").VALUE, String$(Len(C_rec.C110), "0"))
        C_rec.C111 = Format$(oRecordSet.Fields("C111").VALUE, String$(Len(C_rec.C111), "0"))
        C_rec.C112 = Format$(oRecordSet.Fields("C112").VALUE, String$(Len(C_rec.C112), "0"))
        C_rec.C113_1 = Format$(oRecordSet.Fields("C113_1").VALUE, String$(Len(C_rec.C113_1), "0"))
        C_rec.C113_2 = Format$(oRecordSet.Fields("C113_2").VALUE, String$(Len(C_rec.C113_2), "0"))
        C_rec.C114_1 = Format$(oRecordSet.Fields("C114_1").VALUE, String$(Len(C_rec.C114_1), "0"))
        C_rec.C114_2 = Format$(oRecordSet.Fields("C114_2").VALUE, String$(Len(C_rec.C114_2), "0"))
        C_rec.C115_1 = Format$(oRecordSet.Fields("C115_1").VALUE, String$(Len(C_rec.C115_1), "0"))
        C_rec.C115_2 = Format$(oRecordSet.Fields("C115_2").VALUE, String$(Len(C_rec.C115_2), "0"))
        C_rec.C116_1 = Format$(oRecordSet.Fields("C116_1").VALUE, String$(Len(C_rec.C116_1), "0"))
        C_rec.C116_2 = Format$(oRecordSet.Fields("C116_2").VALUE, String$(Len(C_rec.C116_2), "0"))
        C_rec.C117 = Format$(oRecordSet.Fields("C117").VALUE, String$(Len(C_rec.C117), "0"))
        C_rec.C118 = Format$(oRecordSet.Fields("C118").VALUE, String$(Len(C_rec.C118), "0"))
        C_rec.C119 = Format$(oRecordSet.Fields("C119").VALUE, String$(Len(C_rec.C119), "0"))
        C_rec.C120 = Format$(oRecordSet.Fields("C120").VALUE, String$(Len(C_rec.C120), "0"))
        C_rec.C121_1 = Format$(oRecordSet.Fields("C121_1").VALUE, String$(Len(C_rec.C121_1), "0"))
        C_rec.C121_2 = Format$(oRecordSet.Fields("C121_2").VALUE, String$(Len(C_rec.C121_2), "0"))
        C_rec.C122_1 = Format$(oRecordSet.Fields("C122_1").VALUE, String$(Len(C_rec.C122_1), "0"))
        C_rec.C122_2 = Format$(oRecordSet.Fields("C122_2").VALUE, String$(Len(C_rec.C122_2), "0"))
        C_rec.C123_1 = Format$(oRecordSet.Fields("C123_1").VALUE, String$(Len(C_rec.C123_1), "0"))
        C_rec.C123_2 = Format$(oRecordSet.Fields("C123_2").VALUE, String$(Len(C_rec.C123_2), "0"))
        C_rec.C124_1 = Format$(oRecordSet.Fields("C124_1").VALUE, String$(Len(C_rec.C124_1), "0"))
        C_rec.C124_2 = Format$(oRecordSet.Fields("C124_2").VALUE, String$(Len(C_rec.C124_2), "0"))
        C_rec.C125 = oRecordSet.Fields("C125").VALUE
        
        '예제
        'C_rec.PERNBR = Replace(oRecordSet.Fields("U_PERNBR").VALUE, "-", "")
        
        'OLDBIG = Val(oRecordSet.Fields("U_BIGWA1").VALUE) + Val(oRecordSet.Fields("U_BIGWA3").VALUE) + Val(oRecordSet.Fields("U_BIGWA5").VALUE) _
        '        + Val(oRecordSet.Fields("U_BIGWA6").VALUE) + Val(oRecordSet.Fields("U_BIGWU3").VALUE)
                              
        'C_rec.FILD02 = Format$(0, String$(Len(C_rec.FILD02), "0"))
        'C_rec.GAMFLD = String$(Len(C_rec.GAMFLD), "0")
        'C_rec.FILLER = Space$(Len(C_rec.FILLER))
        'C_rec.C022 = Format$(oRecordSet.Fields("C022").VALUE, , String$(Len(C_rec.C022), "0"))
        
      
        Print #1, MDC_SetMod.sStr(C_rec.C001) & MDC_SetMod.sStr(C_rec.C002) & MDC_SetMod.sStr(C_rec.C003) & MDC_SetMod.sStr(C_rec.C004) & MDC_SetMod.sStr(C_rec.C005) & MDC_SetMod.sStr(C_rec.C006) & MDC_SetMod.sStr(C_rec.C007) & MDC_SetMod.sStr(C_rec.C008) & MDC_SetMod.sStr(C_rec.C009) & MDC_SetMod.sStr(C_rec.C010) & _
                  MDC_SetMod.sStr(C_rec.C011) & MDC_SetMod.sStr(C_rec.C012) & MDC_SetMod.sStr(C_rec.C013) & MDC_SetMod.sStr(C_rec.C014) & MDC_SetMod.sStr(C_rec.C015) & MDC_SetMod.sStr(C_rec.C016) & MDC_SetMod.sStr(C_rec.C017) & MDC_SetMod.sStr(C_rec.C018) & MDC_SetMod.sStr(C_rec.C019) & MDC_SetMod.sStr(C_rec.C020) & _
                  MDC_SetMod.sStr(C_rec.C021) & MDC_SetMod.sStr(C_rec.C022) & MDC_SetMod.sStr(C_rec.C023) & MDC_SetMod.sStr(C_rec.C024) & MDC_SetMod.sStr(C_rec.C025) & MDC_SetMod.sStr(C_rec.C026) & MDC_SetMod.sStr(C_rec.C027) & MDC_SetMod.sStr(C_rec.C028) & MDC_SetMod.sStr(C_rec.C029) & MDC_SetMod.sStr(C_rec.C030) & _
                  MDC_SetMod.sStr(C_rec.C031) & MDC_SetMod.sStr(C_rec.C032) & MDC_SetMod.sStr(C_rec.C033) & MDC_SetMod.sStr(C_rec.C034) & MDC_SetMod.sStr(C_rec.C035) & MDC_SetMod.sStr(C_rec.C036) & MDC_SetMod.sStr(C_rec.C037) & MDC_SetMod.sStr(C_rec.C038) & MDC_SetMod.sStr(C_rec.C039) & MDC_SetMod.sStr(C_rec.C040) & _
                  MDC_SetMod.sStr(C_rec.C041) & MDC_SetMod.sStr(C_rec.C042) & MDC_SetMod.sStr(C_rec.C043) & MDC_SetMod.sStr(C_rec.C044) & MDC_SetMod.sStr(C_rec.C045) & MDC_SetMod.sStr(C_rec.C046) & MDC_SetMod.sStr(C_rec.C047) & MDC_SetMod.sStr(C_rec.C048) & MDC_SetMod.sStr(C_rec.C049) & MDC_SetMod.sStr(C_rec.C050) & _
                  MDC_SetMod.sStr(C_rec.C051) & MDC_SetMod.sStr(C_rec.C052) & MDC_SetMod.sStr(C_rec.C053) & MDC_SetMod.sStr(C_rec.C054) & MDC_SetMod.sStr(C_rec.C055) & MDC_SetMod.sStr(C_rec.C056) & MDC_SetMod.sStr(C_rec.C057) & MDC_SetMod.sStr(C_rec.C058) & MDC_SetMod.sStr(C_rec.C059) & MDC_SetMod.sStr(C_rec.C060) & _
                  MDC_SetMod.sStr(C_rec.C061) & MDC_SetMod.sStr(C_rec.C062) & MDC_SetMod.sStr(C_rec.C063) & MDC_SetMod.sStr(C_rec.C064) & MDC_SetMod.sStr(C_rec.C065) & MDC_SetMod.sStr(C_rec.C066) & MDC_SetMod.sStr(C_rec.C067) & MDC_SetMod.sStr(C_rec.C068) & MDC_SetMod.sStr(C_rec.C069) & MDC_SetMod.sStr(C_rec.C070) & _
                  MDC_SetMod.sStr(C_rec.C071) & MDC_SetMod.sStr(C_rec.C072) & MDC_SetMod.sStr(C_rec.C073) & MDC_SetMod.sStr(C_rec.C074) & MDC_SetMod.sStr(C_rec.C075) & MDC_SetMod.sStr(C_rec.C076) & MDC_SetMod.sStr(C_rec.C077) & MDC_SetMod.sStr(C_rec.C078) & MDC_SetMod.sStr(C_rec.C079) & MDC_SetMod.sStr(C_rec.C080) & _
                  MDC_SetMod.sStr(C_rec.C081) & MDC_SetMod.sStr(C_rec.C082) & MDC_SetMod.sStr(C_rec.C083) & MDC_SetMod.sStr(C_rec.C084) & MDC_SetMod.sStr(C_rec.C085) & MDC_SetMod.sStr(C_rec.C086) & MDC_SetMod.sStr(C_rec.C087) & MDC_SetMod.sStr(C_rec.C088) & MDC_SetMod.sStr(C_rec.C089) & MDC_SetMod.sStr(C_rec.C090) & _
                  MDC_SetMod.sStr(C_rec.C091) & MDC_SetMod.sStr(C_rec.C092) & MDC_SetMod.sStr(C_rec.C093) & MDC_SetMod.sStr(C_rec.C094) & MDC_SetMod.sStr(C_rec.C095) & MDC_SetMod.sStr(C_rec.C096) & MDC_SetMod.sStr(C_rec.C097) & MDC_SetMod.sStr(C_rec.C098) & MDC_SetMod.sStr(C_rec.C099) & MDC_SetMod.sStr(C_rec.C100) & _
                  MDC_SetMod.sStr(C_rec.C101) & MDC_SetMod.sStr(C_rec.C102) & MDC_SetMod.sStr(C_rec.C103) & MDC_SetMod.sStr(C_rec.C104) & MDC_SetMod.sStr(C_rec.C105) & MDC_SetMod.sStr(C_rec.C106) & MDC_SetMod.sStr(C_rec.C107) & MDC_SetMod.sStr(C_rec.C108_1) & MDC_SetMod.sStr(C_rec.C108_2) & MDC_SetMod.sStr(C_rec.C109_1) & MDC_SetMod.sStr(C_rec.C109_2) & MDC_SetMod.sStr(C_rec.C110) & _
                  MDC_SetMod.sStr(C_rec.C111) & MDC_SetMod.sStr(C_rec.C112) & MDC_SetMod.sStr(C_rec.C113_1) & MDC_SetMod.sStr(C_rec.C113_2) & MDC_SetMod.sStr(C_rec.C114_1) & MDC_SetMod.sStr(C_rec.C114_2) & MDC_SetMod.sStr(C_rec.C115_1) & MDC_SetMod.sStr(C_rec.C115_2) & MDC_SetMod.sStr(C_rec.C116_1) & MDC_SetMod.sStr(C_rec.C116_2) & MDC_SetMod.sStr(C_rec.C117) & _
                  MDC_SetMod.sStr(C_rec.C118) & MDC_SetMod.sStr(C_rec.C119) & MDC_SetMod.sStr(C_rec.C120) & MDC_SetMod.sStr(C_rec.C121_1) & MDC_SetMod.sStr(C_rec.C121_2) & MDC_SetMod.sStr(C_rec.C122_1) & MDC_SetMod.sStr(C_rec.C122_2) & MDC_SetMod.sStr(C_rec.C123_1) & MDC_SetMod.sStr(C_rec.C123_2) & MDC_SetMod.sStr(C_rec.C124_1) & MDC_SetMod.sStr(C_rec.C124_2) & MDC_SetMod.sStr(C_rec.C125)
                  
        
        '/ D레코드: 연금계좌입금명세 레코드 수록
        If Val(C_rec.C110) > 0 Then     '/계좌입금금액_합계  2016부터있음
            If File_Create_D_record = False Then
                ErrNum = 2
                GoTo Error_Message
            End If
        End If
        
        
        oRecordSet.MoveNext
        
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 작성중........!"
        
        
    Loop
        
        
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
        Sbo_Application.StatusBar.SetText "C레코드가 존재하지 않습니다.", bmt_Short, smt_Error
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
    sQry = "EXEC PH_PY995_D '" & C_SAUP & "', '" & C_YYYY & "', '" & C_SABUN & "'"
    
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
        D_rec.D012 = Format$(oRecordSet.Fields("D012").VALUE, String$(Len(D_rec.D012), "0"))
        D_rec.D013 = Format$(oRecordSet.Fields("D013").VALUE, String$(Len(D_rec.D013), "0"))
        D_rec.D014 = oRecordSet.Fields("D014").VALUE
        
        Print #1, MDC_SetMod.sStr(D_rec.D001) & MDC_SetMod.sStr(D_rec.D002) & MDC_SetMod.sStr(D_rec.D003) & MDC_SetMod.sStr(D_rec.D004) & MDC_SetMod.sStr(D_rec.D005) & MDC_SetMod.sStr(D_rec.D006) & MDC_SetMod.sStr(D_rec.D007) & MDC_SetMod.sStr(D_rec.D008) & MDC_SetMod.sStr(D_rec.D009) & MDC_SetMod.sStr(D_rec.D010) & _
                  MDC_SetMod.sStr(D_rec.D011) & MDC_SetMod.sStr(D_rec.D012) & MDC_SetMod.sStr(D_rec.D013) & MDC_SetMod.sStr(D_rec.D014)
                  
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
        Sbo_Application.StatusBar.SetText "D레코드가 존재하지 않습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("D레코드오류: " & Err.Description, False)
    End If
    File_Create_D_record = False
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






