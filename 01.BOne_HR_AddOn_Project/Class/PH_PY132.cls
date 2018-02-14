VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY132"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY132.cls
'//  Module         : 급여관리 > 성과급 차등 개인별 계산
'//  Desc           : 성과급 차등 개인별 계산
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY132A   As SAPbouiCOM.DBDataSource
Private oDS_PH_PY132B   As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

'//------------------------------------------------------
Private G_TotCnt    As Integer    '/ 대상인원수
Private G_PayCnt    As Integer    '/ 계산인원수
Private G_ChkCnt    As Integer    '/ 잠금제외수


Private oYM         As String       '/ 귀속연월
Private oJOBTYP     As String       '/ 지급종류
Private oJOBGBN     As String       '/ 지급구분
Private oJIGBIL     As String       '/ 지급일자
Private oJOBTRG     As String       '/ 지급대상구분
Private oCLTCOD     As String       '/ 사업장
Private oSTRDPT     As String       '/ 부서시작코드
Private oENDDPT     As String       '/ 부서종료코드
Private oMSTCOD     As String       '/ 사원번호
Private oJSNCHK     As String       '/ 연말정산포함
Private oYCHCHK     As String       '/ 연차지급포함

Private oBNSCAL     As String       '/ 상여계산방법
Private oBNSRAT     As String       '/ 상여율
Private oSTRTAX     As String       '/ 세액대상기간시작(급여)
Private oENDTAX     As String       '/ 세액대상기간종료(급여)
Private oSTRBNS     As String       '/ 세액대상기간시작(상여)
Private oENDBNS     As String       '/ 세액대상기간종료(상여)
Private oGNEDAT     As String       '/ 상여계산기준일
Private oEXPDAT     As String       '/ 퇴사자제외일
Private oRETCHK     As String       '/ 상여퇴직임금에포함

Private StrDate     As String
Private EndDate     As String
Private MaxRow      As Long
Private JSNYER      As Integer

Private G06_CHK     As Boolean  '/ 소득세정산
Private G07_CHK     As Boolean  '/ 주민세정산
Private G08_CHK     As Boolean  '/ 건강보험정산
Private G90_CHK     As Boolean  '/ 국민연금정산
Private G91_CHK     As Boolean  '/ 농특세정산
Private G92_CHK     As Boolean  '/ 고용보험정산
Private G04_BNSUSE  As String
Private PAY_001     As String
Private PAY_007     As String

Private Type WG03TILR
    CSUCOD(1 To 24) As String
    CSUNAM(1 To 24) As String
    MPYGBN(1 To 24) As String   '/ 월정급여
    CSUKUM(1 To 24) As Double   '/ 수당한도금액
    GWATYP(1 To 24) As String   '/ 과세구분
    GBHGBN(1 To 24) As String   '/ 고용보험여부
    ROUNDT(1 To 24) As String   '/ 사사오입구분(끝전처리)
    RODLEN(1 To 24) As Integer  '/ 끝전처리자릿수
    GONSIL(1 To 24) As String   '/ 급여수식
    BNSUSE(1 To 24) As String   '/ 상여항목
    BTXCOD(1 To 24) As String   '/ 비과세코드
End Type
Dim WK_C As WG03TILR

Private Type WG04TILR
    GONCOD(1 To 18) As String
    GONNAM(1 To 18) As String
    BNSUSE(1 To 18) As String   '/ 상여
    GONSIL(1 To 18) As String   '/ 계산식
    ROUNDT(1 To 18) As String
    RODLEN(1 To 18) As Integer
End Type
Dim WK_G As WG04TILR

Private U_CSUCOD    As String
Private U_GONCOD    As String
Private WK_TIMAMT   As Double        '시    급
Private WK_DAYAMT   As Double        '일    급
Private WK_STDAMT   As Double        '월    급
Private WK_BNSAMT   As Double        '상여기본
Private WK_APPBNS   As Double        '적용상여금
Private WK_AVRAMT   As Double        '급여기본등록의 평균임금

Private Type WG33PAYR
    DocNum      As Integer  '문서번호
    U_MSTCOD    As String * 8
    U_MSTNAM    As String * 50
    U_EmpID     As String * 10
    U_MSTBRK    As String * 10
    U_CLTCOD    As String * 8
    U_MSTDPT    As String * 10       '부서
    U_MSTSTP    As String * 10       '직책
    U_CLTNAM    As String * 20
    U_BRKNAM    As String * 20
    U_DPTNAM    As String * 20       '부서
    U_STPNAM    As String * 20       '직책
    U_PAYTYP    As String * 10
    U_JOBTRG    As String * 8        '급여지급일구분
    U_MEDAMT    As Double
    U_KUKAMT    As Double
    U_GBHAMT    As Double
    U_PERNBR    As String
    U_JIGCOD    As String * 10
    U_HOBONG    As String * 10
    U_STDAMT    As Double            '기 본 급
    U_BASAMT    As Double            '통상일급
    U_DAYAMT    As Double            '기본일급
    U_INPDAT    As String * 8        '입사일자
    U_INEDAT    As String * 8        '수습만료일
    U_OUTDAT    As String * 8        '퇴사일자
    U_TAXCNT    As Integer
    U_CHLCNT    As Integer
    U_NJCGBN    As String
    U_MEDFRG    As Double
    U_CSUAMT(1 To 24) As Double
    U_GWASEE    As Double
    U_BTAX01    As Double
    U_BTAX02    As Double
    U_BTAX03    As Double
    U_BTAX04    As Double
    U_BTAX05    As Double
    U_BTAX06    As Double
    U_BTAX07    As Double
    
    U_BTXG01    As Double
    U_BTXH01    As Double
    U_BTXH05    As Double
    U_BTXH06    As Double
    U_BTXH07    As Double
    U_BTXH08    As Double
    U_BTXH09    As Double
    U_BTXH10    As Double
    U_BTXH11    As Double
    U_BTXH12    As Double
    U_BTXH13    As Double
    U_BTXI01    As Double
    U_BTXK01    As Double
    U_BTXM01    As Double
    U_BTXM02    As Double
    U_BTXM03    As Double
    U_BTXO01    As Double
    U_BTXQ01    As Double
    U_BTXR10    As Double
    U_BTXS01    As Double
    U_BTXT01    As Double
    U_BTXY01    As Double
    U_BTXY02    As Double
    U_BTXY03    As Double
    U_BTXY21    As Double
    U_BTXZ01    As Double
    U_BTXY22    As Double
    
    U_BTXX01    As Double
    U_BTXY20    As Double
    
    U_BTXTOT    As Double
    
    U_TOTPAY    As Double
    U_GONAMT(1 To 18) As Double
    U_TOTGON    As Double
    U_SILJIG    As Double
    '// 상여금
    U_AVRPAY    As Double
    U_NABTAX    As Double
    U_BNSRAT    As Double
    U_APPRAT    As Double
    U_GNSYER    As Integer
    U_GNSMON    As Integer
    U_TAXTRM    As Integer
    U_BONUSS    As Double
End Type
Dim WG03 As WG33PAYR

Private Type WG01CODR
    TB1AMT(1 To 10) As Double
    TB1GON(1 To 10) As Double
    TB1RAT(1 To 10) As Double
    TB1KUM(1 To 10) As Double
End Type
Dim WG01 As WG01CODR
Private TB1_BT3COD As String  '/ 사용하는 국외비과세코드
Private TB1_BT5COD As String  '/ 사용하는 연구비과세코드


Private X01_Val         As Double   '/X01
Private X02_Val         As Double   '/X02
Private X03_Val         As Double   '/X03
Private X04_Val         As Double   '/X04

Private X10_Val         As Integer
Private X11_Val         As Integer
Private X12_Val         As Integer
Private X13_Val         As Integer
Private X14_Val         As String
Private X15_Val         As Double
Private X16_Val         As Integer
Private X17_Val         As Double
Private X18_Val         As Integer
Private X19_Val         As Integer
Private X20_Val         As Integer

Private REMARK1         As String
Private REMARK2         As String
Private REMARK3         As String
Private TermCHK         As Boolean
Private sRecordset      As SAPbobsCOM.Recordset

Private tDocEntry   As Long         '//저장전 문서번호 저장
Private CalcYN      As Boolean      '//계산 여부


Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY132.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY132_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY132"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
    
    
    oForm.Freeze True
    Call PH_PY132_CreateItems
    Call PH_PY132_EnableMenus
    Call PH_PY132_SetDocument(oFromDocEntry01)
'    Call PH_PY132_FormResize
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing
    Exit Sub
    
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY132_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    Dim iCol%
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY132_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------------------
    '// 데이터셋정의
    '//----------------------------------------------------------------------------------------------
    Set oDS_PH_PY132A = oForm.DataSources.DBDataSources("@PH_PY132A")   '//헤더
    Set oDS_PH_PY132B = oForm.DataSources.DBDataSources("@PH_PY132B")   '//라인
    
    Set oMat1 = oForm.Items("Mat1").Specific       '

    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
'/ 사업장
'    Set oCombo = oForm.Items("CLTCOD").Specific
'    oForm.Items("CLTCOD").DisplayDesc = True

    
'/ 귀속연월
    oDS_PH_PY132A.setValue "U_YM", 0, Format$(Now, "YYYYMM")
    
'/ 지급구분
    Set oCombo = oForm.Items("JOBGBN").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JOBGBN").DisplayDesc = True
    oCombo.Select "1", psk_ByValue
    
    
'    Call oForm.DataSources.UserDataSources.Add("Col0", dt_SHORT_TEXT)
'    Call oForm.DataSources.UserDataSources.Add("Col1", dt_SHORT_TEXT)
'    Set oColumn = oMat1.Columns("Col0")
'    oColumn.DataBind.SetBound True, "", "Col0"
'    Set oColumn = oMat1.Columns("Col1")
'    oColumn.DataBind.SetBound True, "", "Col1"

    
'// 상여
    For iCol = 1 To 8
        Set oCombo = oForm.Items("AP" & iCol & "GBN").Specific
        oCombo.ValidValues.Add "1", "개월 이상"
        oCombo.ValidValues.Add "2", "개월 미만"
        oCombo.ValidValues.Add "3", "일수 이상"
        oCombo.ValidValues.Add "4", "일수 미만"
        If oCombo.ValidValues.Count > 0 Then
           oDS_PH_PY132A.setValue "U_AP" & iCol & "GBN", 0, "1"
        End If
        oForm.Items("AP" & iCol & "GBN").DisplayDesc = True
    Next iCol

    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY132_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY132_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY132_EnableMenus()
    
    On Error GoTo PH_PY132_EnableMenus_Error

    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1287", True)     '//복제
    
    Exit Sub
    
PH_PY132_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY132_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY132_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY132_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY132_FormItemEnabled
        Call PH_PY132_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY132_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY132_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY132_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY132_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY132_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        
        oForm.ActiveItem = "YM"
        oForm.Items("Code").Enabled = False
        oForm.Items("Btn1").Visible = True
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
        
        '/ 귀속연월
        oDS_PH_PY132A.setValue "U_YM", 0, Format$(Now, "YYYYMM")
        '/ 지급구분
        Set oCombo = oForm.Items("JOBGBN").Specific
        oCombo.Select "1", psk_ByValue
        
    
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("Code").Enabled = True
        oForm.Items("Btn1").Visible = False
        oForm.ActiveItem = "Code"
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
'        oForm.Mode = fm_UPDATE_MODE
        oForm.ActiveItem = "YM"
        oForm.Items("Code").Enabled = False
        oForm.Items("Btn1").Visible = True
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY132_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY132_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim JIGBIL      As String
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY132_DataValidCheck = False Then  '/ 유효성 검사
                            BubbleEvent = False
                        End If
                    End If
                    tDocEntry = oForm.Items("Code").Specific.VALUE
                ElseIf pval.ItemUID = "Btn1" Then       '/ 급(상)여계산
                    CalcYN = True
                    tDocEntry = oForm.Items("Code").Specific.VALUE
                    If PH_PY132_DataValidCheck = True And Pay_Calc = True Then  '/ 유효성 검사
                        If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                            oForm.Items("1").CLICK ct_Regular
                        ElseIf oForm.Mode = fm_OK_MODE Then
                            Sbo_Application.StatusBar.SetText "성과급 차등계산이 진행중입니다.", bmt_Short, smt_Warning
                            If oForm.Items("JOBGBN").Specific.VALUE <> "2" Then
                                '//정상계산
                                oRecordSet.DoQuery "EXEC PH_PY132 '" & tDocEntry & "'"
                            Else
                                '//소급계산
                                oRecordSet.DoQuery "EXEC PH_PY132_SOGUB '" & tDocEntry & "'"
                            End If
                            Sbo_Application.StatusBar.SetText "성과급 차등계산이 완료되었습니다.", bmt_Short, smt_Success
                        End If
                    Else
                        BubbleEvent = False
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If pval.ActionSuccess = True Then
                        If CalcYN = True Then
                            CalcYN = False
                            If oForm.Mode = fm_ADD_MODE Or fm_UPDATE_MODE Then
'                                Sbo_Application.StatusBar.SetText "성과급 차등계산이 진행중입니다.", bmt_Short, smt_Warning
'                                If oForm.Items("JOBGBN").Specific.VALUE <> "2" Then
'                                '//정상계산
'                                    oRecordSet.DoQuery "EXEC PH_PY132 '" & tDocEntry & "'"
'                                Else
'                                    '//소급계산
'                                    oRecordSet.DoQuery "EXEC PH_PY132_SOGUB '" & tDocEntry & "'"
'                                End If
                            
                                Sbo_Application.StatusBar.SetText "성과급 차등계산이 완료되었습니다.", bmt_Short, smt_Success
                                oForm.Mode = fm_FIND_MODE
                                PH_PY132_FormItemEnabled
                                oForm.Items("1").CLICK ct_Regular
                            End If
                        ElseIf CalcYN = False Then
                            If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                                oForm.Mode = fm_FIND_MODE
                                Call PH_PY132_FormItemEnabled
                                oForm.Items("1").CLICK ct_Regular
                            Else
                                Call PH_PY132_FormItemEnabled
                                
                            End If
                        
                        End If
                    End If
                End If
                
                
            End If
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True And pval.ItemUID = "YM" And pval.CharPressed = 9 And pval.FormMode = fm_ADD_MODE Then
               If Trim$(oForm.Items(pval.ItemUID).Specific.String) = "" Then
                   Sbo_Application.StatusBar.SetText "귀속연월은 필수입니다. 확인하여 주십시오.", bmt_Short, smt_Error
                   BubbleEvent = False
               End If
            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat1"
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUID
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Case Else
                oLastItemUID = pval.ItemUID
                oLastColUID = ""
                oLastColRow = 0
            End Select
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                '//지급구분
                If pval.ItemUID = "JOBGBN" Then
                    If Not oForm.Items("JOBGBN").Specific.Selected Is Nothing Then
                        oYM = Trim$(oDS_PH_PY132A.GetValue("U_YM", 0))
                        JIGBIL = Format(Now, "yyyymmdd")
                        oForm.Items("JIGBIL").Specific.VALUE = JIGBIL
                        oJOBGBN = Trim$(oDS_PH_PY132A.GetValue("U_JOBGBN", 0))
                    Else
                        oJOBGBN = ""
                    End If
                    If oJOBGBN <> "" Then Call Display_BonussRate
                    
                End If
                '//급여지급대상일
                
                
            End If
'----------------------------------------------------------
        Case et_CLICK: '//6

'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7

'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                '//귀속년월
                If pval.ItemUID = "YM" Then
                    If Trim$(oForm.Items("YM").Specific.VALUE) = "" Then
                        oDS_PH_PY132A.setValue "U_YM", 0, Format$(Now, "YYYYMM")
                    Else
                        oDS_PH_PY132A.setValue "U_YM", 0, oForm.Items("YM").Specific.String
                    End If
                    oYM = Trim$(oDS_PH_PY132A.GetValue("U_YM", 0))
                    
                    If oYM <> "" Then Call Display_BonussRate
                End If
                
            End If
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY132_FormItemEnabled
                Call PH_PY132_AddMatrixRow
            End If
'----------------------------------------------------------
        Case et_DATASOURCE_LOAD '//12
           
'----------------------------------------------------------
        Case et_FORM_LOAD: '//16
            
'----------------------------------------------------------
        Case et_FORM_UNLOAD: '//17
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '컬렉션에서 삭제및 모든 메모리 제거
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_PH_PY132A = Nothing
'
'                Set oMat1 = Nothing
            End If
'----------------------------------------------------------
        Case et_FORM_ACTIVATE: '//18
            
'----------------------------------------------------------
        Case et_FORM_DEACTIVATE: '//19
            
'----------------------------------------------------------
        Case et_FORM_CLOSE '//20
            
'----------------------------------------------------------
        Case et_FORM_RESIZE '//21
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.BeforeAction = False Then
'
'            End If
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.Before_Action = False Then
'
'            End If
'----------------------------------------------------------
        Case et_PICKER_CLICKED '//37
            
'----------------------------------------------------------
        Case et_GRID_SORT '//38
            
'----------------------------------------------------------
        Case et_Drag '//39
            
    End Select
    
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    oForm.Freeze (False)
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i As Long
    On Error GoTo Raise_FormMenuEvent_Error
    Call oForm.Freeze(True)
    
    If (pval.BeforeAction = True) Then
        Select Case pval.MenuUID
            Case "1283":
                If Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1284":
            Case "1286":
            Case "1293":
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
                'Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY132A", "DocEntry")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY132_FormItemEnabled
                Call PH_PY132_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY132_FormItemEnabled
                Call PH_PY132_AddMatrixRow
'                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY132_FormItemEnabled
                Call PH_PY132_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY132_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY132B.Size - 1)
'                        If oDS_PH_PY132B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY132B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY132B.Size
'                        Call oDS_PH_PY132B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
'                End If
'                Call PH_PY132_AddMatrixRow
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_FormDataEvent_Error
    
    If (BusinessObjectInfo.BeforeAction = True) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select
    End If
    Exit Sub
    
Raise_FormDataEvent_Error:
    
        Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_RightClickEvent_Error
    
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
    End If
    Select Case pval.ItemUID
    Case "Mat1"
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUID
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Case Else
        oLastItemUID = pval.ItemUID
        oLastColUID = ""
        oLastColRow = 0
    End Select
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY132_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY132_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 용]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount

    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY132B.GetValue("U_CLTCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY132B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY132B.InsertRecord (oRow)
            End If
            oDS_PH_PY132B.Offset = oRow
            oDS_PH_PY132B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY132B.setValue "U_CLTCOD", oRow, ""
'            oDS_PH_PY132B.setValue "U_TeamCode", oRow, ""
'            oDS_PH_PY132B.setValue "U_TeamName", oRow, ""
            
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY132B.Offset = oRow - 1
            oDS_PH_PY132B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY132B.setValue "U_CLTCOD", oRow - 1, ""
'            oDS_PH_PY132B.setValue "U_TeamCode", oRow - 1, ""
'            oDS_PH_PY132B.setValue "U_FILD03", oRow - 1, 0
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY132B.Offset = oRow
        oDS_PH_PY132B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY132B.setValue "U_CLTCOD", oRow, ""
'        oDS_PH_PY132B.setValue "U_TeamCode", oRow, ""
'        oDS_PH_PY132B.setValue "U_FILD03", oRow, 0
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY132_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY132_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY132_FormClear()
On Error GoTo PH_PY132_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY132'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY132_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY132_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY132_DataValidCheck() As Boolean
On Error GoTo PH_PY132_DataValidCheck_Error
    Dim ErrNum          As Integer
    Dim sQry    As String
    Dim tCode   As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    ErrNum = 0
    
    PH_PY132_DataValidCheck = False
    
    '/
    oYM = Trim$(oDS_PH_PY132A.GetValue("U_YM", 0))
    oJOBGBN = Trim$(oDS_PH_PY132A.GetValue("U_JOBGBN", 0))
    oBNSRAT = oDS_PH_PY132A.GetValue("U_bBNSRAT", 0)
    oJIGBIL = oDS_PH_PY132A.GetValue("U_JIGBIL", 0)
    
    '/ Check
    Select Case True
      Case MDC_SetMod.ChkYearMonth(oYM) = False
           Sbo_Application.StatusBar.SetText "귀속 연월을 확인하여 주십시오.", bmt_Short, smt_Error
           Exit Function
      Case oJOBGBN = ""
           Sbo_Application.StatusBar.SetText "지급 구분을 선택하여 주십시오.", bmt_Short, smt_Error
           Exit Function
      Case oJIGBIL = ""
           Sbo_Application.StatusBar.SetText "지급일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
           Exit Function
    End Select

'/ 상여계산시 체크
    If oJOBTYP <> "1" Then
        Select Case True
        Case Val(oBNSRAT) = 0
           Sbo_Application.StatusBar.SetText "상여율은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
           Exit Function
        End Select
    End If
    
    
    tCode = Trim(oDS_PH_PY132A.GetValue("U_YM", 0)) + Trim(oDS_PH_PY132A.GetValue("U_JOBGBN", 0)) + _
            Trim(oDS_PH_PY132A.GetValue("U_Number", 0))
            
    '//코드생성
    '//코드 중복 체크
    If oForm.Mode = fm_ADD_MODE Then
        sQry = "SELECT CODE FROM [@PH_PY132A] WHERE CODE = '" & tCode & "'"
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount > 0 Then
            Sbo_Application.SetStatusBarMessage "코드가 존재합니다", bmt_Short, True
            PH_PY132_DataValidCheck = False
            Exit Function
        Else
            Call oDS_PH_PY132A.setValue("Code", 0, tCode)
            Call oDS_PH_PY132A.setValue("Name", 0, tCode)
        End If
    End If
    
    
    PH_PY132_DataValidCheck = True
    Exit Function
    
PH_PY132_DataValidCheck_Error:
    PH_PY132_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY132_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY132_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY132_Validate_Error
    PH_PY132_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY132A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY132_Validate = False
        GoTo PH_PY132_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY132_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY132_Validate_Error:
    PH_PY132_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY132_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Function Pay_Calc() As Boolean
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim Temp1           As String
    Dim Temp2           As String
    Dim i               As Long
    
    Dim YM As String
    Dim JOBGBN As String
    Dim Number As String
    Dim JIGBIL As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Pay_Calc = False
    
    
    YM = Trim(oDS_PH_PY132A.GetValue("U_YM", 0))
    JOBGBN = Trim(oDS_PH_PY132A.GetValue("U_JOBGBN", 0))
    Number = Trim(oDS_PH_PY132A.GetValue("U_Number", 0))
    JIGBIL = Trim(oDS_PH_PY132A.GetValue("U_JIGBIL", 0))
    
    sQry = "Select Count(*) From [@PH_PY132A] "
    sQry = sQry + " Where U_YM = '" & YM & "'"
    sQry = sQry + " AND U_Number = '" & Number & "'"
    sQry = sQry + " AND U_JIGBIL = '" & JIGBIL & "'"
    
    'sQry = "Select U_CLTCOD, U_YM, U_JOBTYP, U_JOBGBN, U_JOBTRG From [@PH_PY112A] Where U_CLTCOD = '" & Trim(oDS_PH_PY132A.GetValue("U_CLTCOD", 0)) & "' AND U_YM = '" & Trim(oDS_PH_PY132A.GetValue("U_YM", 0)) & "' AND U_JOBTYP = '" & Trim(oDS_PH_PY132A.GetValue("U_JOBTYP", 0)) & "' "
    oRecordSet.DoQuery sQry
        
    If oRecordSet.Fields(0).VALUE > 0 Then
        If Sbo_Application.MessageBox("기존에 급여계산 결과가 있습니다. 계속 진행하시겠습니까?", 2, "Yes", "No") = 2 Then
            Pay_Calc = False
            Set oRecordSet = Nothing
            Exit Function
        End If
        Pay_Calc = True
    Else
        Pay_Calc = True
    End If
    Set oRecordSet = Nothing
'    Temp1 = Trim(oDS_PH_PY132A.GetValue("U_CLTCOD", 0)) & Trim(oDS_PH_PY132A.GetValue("U_YM", 0)) & Trim(oDS_PH_PY132A.GetValue("U_JOBTYP", 0)) & _
'            Trim(oDS_PH_PY132A.GetValue("U_JOBGBN", 0)) & Trim(oDS_PH_PY132A.GetValue("U_JOBTRG", 0))
'
'    If oRecordSet.RecordCount > 0 Then
'        Do Until oRecordSet.EOF
'            Temp2 = ""
'            For i = 0 To oRecordSet.Fields.Count - 1
'                Temp2 = Temp2 & Trim(oRecordSet.Fields(i).VALUE)
'            Next
'
'            If Temp1 = Temp2 Then
'                GoTo Continue
'            End If
'            oRecordSet.MoveNext
'        Loop
'    End If
'    GoTo Continue2
'
'Continue:
'    If Sbo_Application.MessageBox("기존에 급여계산 결과가 있습니다. 계속 진행하시겠습니까?", 2, "Yes", "No") = 2 Then
'        Pay_Calc = False
'        Exit Function
'    End If
'
'
'Continue2:
'    Pay_Calc = True
'    Set oRecordSet = Nothing
'

Exit Function

Error_Message:
    Set oRecordSet = Nothing
    Pay_Calc = False
End Function

Private Sub Display_BonussRate()
On Error GoTo Error_Message
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim oRecordset2 As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim ErrNum      As Integer
    Dim iCol        As Integer
    Dim iBNSMON     As Long
    
    ErrNum = 0
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    Set oRecordset2 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim JOBGBN As String
    
    '//소급일때 상여와 동일하게
    If oJOBGBN = "2" Then
        JOBGBN = "1"
    Else
        JOBGBN = oJOBGBN
    End If
    
    oCLTCOD = "1"
    
    sQry = " SELECT U_BNSCAL, U_BNSMON, U_BNSRAT, U_AP1MON, U_AP2MON, U_AP3MON, U_AP4MON, U_AP5MON,"
    sQry = sQry & " U_AP6MON, U_AP7MON, U_AP8MON, U_AP1RAT, U_AP2RAT, U_AP3RAT, U_AP4RAT, U_AP5RAT, "
    sQry = sQry & " U_AP6RAT, U_AP7RAT, U_AP8RAT, U_AP1AMT, U_AP2AMT, U_AP3AMT, U_AP4AMT, U_AP5AMT,"
    sQry = sQry & " U_AP6AMT, U_AP7AMT, U_AP8AMT, U_AP1GBN, U_AP2GBN, U_AP3GBN, U_AP4GBN, U_AP5GBN,"
    sQry = sQry & " U_AP6GBN, U_AP7GBN, U_AP8GBN  FROM [@PH_PY108A] "
    sQry = sQry & " WHERE U_CLTCOD = '" & oCLTCOD & "'  AND U_JOBGBN = '" & JOBGBN & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Or oJOBTYP = "1" Then
        oDS_PH_PY132A.setValue "U_bBNSRAT", 0, 0
        oDS_PH_PY132A.setValue "U_bMONTH1", 0, 0
        oDS_PH_PY132A.setValue "U_bMONTH2", 0, 0
        oDS_PH_PY132A.setValue "U_bMONTH3", 0, 0
        oDS_PH_PY132A.setValue "U_bMONTH4", 0, 0
        oDS_PH_PY132A.setValue "U_bMONTH5", 0, 0
        oDS_PH_PY132A.setValue "U_bMONTH6", 0, 0
        oDS_PH_PY132A.setValue "U_bMONTH7", 0, 0
        oDS_PH_PY132A.setValue "U_bMONTH8", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPRAT1", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPRAT2", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPRAT3", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPRAT4", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPRAT5", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPRAT6", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPRAT7", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPRAT8", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPAMT1", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPAMT2", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPAMT3", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPAMT4", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPAMT5", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPAMT6", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPAMT7", 0, 0
        oDS_PH_PY132A.setValue "U_bAPPAMT8", 0, 0
        For iCol = 1 To 8
            oDS_PH_PY132A.setValue "U_AP" & iCol & "GBN", 0, "1"
        Next iCol
        
        '// 2010.04.05 최동권 추가
        oYM = Trim$(oDS_PH_PY132A.GetValue("U_YM", 0))
        
        
    Else
        oDS_PH_PY132A.setValue "U_bBNSCAL", 0, oRecordSet.Fields(0).VALUE
        oDS_PH_PY132A.setValue "U_bBNSRAT", 0, oRecordSet.Fields(2).VALUE
        oDS_PH_PY132A.setValue "U_bMONTH1", 0, oRecordSet.Fields(3).VALUE
        oDS_PH_PY132A.setValue "U_bMONTH2", 0, oRecordSet.Fields(4).VALUE
        oDS_PH_PY132A.setValue "U_bMONTH3", 0, oRecordSet.Fields(5).VALUE
        oDS_PH_PY132A.setValue "U_bMONTH4", 0, oRecordSet.Fields(6).VALUE
        oDS_PH_PY132A.setValue "U_bMONTH5", 0, oRecordSet.Fields(7).VALUE
        oDS_PH_PY132A.setValue "U_bMONTH6", 0, oRecordSet.Fields(8).VALUE
        oDS_PH_PY132A.setValue "U_bMONTH7", 0, oRecordSet.Fields(9).VALUE
        oDS_PH_PY132A.setValue "U_bMONTH8", 0, oRecordSet.Fields(10).VALUE
        oDS_PH_PY132A.setValue "U_bAPPRAT1", 0, oRecordSet.Fields(11).VALUE
        oDS_PH_PY132A.setValue "U_bAPPRAT2", 0, oRecordSet.Fields(12).VALUE
        oDS_PH_PY132A.setValue "U_bAPPRAT3", 0, oRecordSet.Fields(13).VALUE
        oDS_PH_PY132A.setValue "U_bAPPRAT4", 0, oRecordSet.Fields(14).VALUE
        oDS_PH_PY132A.setValue "U_bAPPRAT5", 0, oRecordSet.Fields(15).VALUE
        oDS_PH_PY132A.setValue "U_bAPPRAT6", 0, oRecordSet.Fields(16).VALUE
        oDS_PH_PY132A.setValue "U_bAPPRAT7", 0, oRecordSet.Fields(17).VALUE
        oDS_PH_PY132A.setValue "U_bAPPRAT8", 0, oRecordSet.Fields(18).VALUE
        oDS_PH_PY132A.setValue "U_bAPPAMT1", 0, oRecordSet.Fields(19).VALUE
        oDS_PH_PY132A.setValue "U_bAPPAMT2", 0, oRecordSet.Fields(20).VALUE
        oDS_PH_PY132A.setValue "U_bAPPAMT3", 0, oRecordSet.Fields(21).VALUE
        oDS_PH_PY132A.setValue "U_bAPPAMT4", 0, oRecordSet.Fields(22).VALUE
        oDS_PH_PY132A.setValue "U_bAPPAMT5", 0, oRecordSet.Fields(23).VALUE
        oDS_PH_PY132A.setValue "U_bAPPAMT6", 0, oRecordSet.Fields(24).VALUE
        oDS_PH_PY132A.setValue "U_bAPPAMT7", 0, oRecordSet.Fields(25).VALUE
        oDS_PH_PY132A.setValue "U_bAPPAMT8", 0, oRecordSet.Fields(26).VALUE
        For iCol = 1 To 8
'            oForm.DataSources.UserDataSources("AP" & iCol & "GBN").ValueEx = oRecordSet.Fields(26 + iCol).Value
            oDS_PH_PY132A.setValue "U_AP" & iCol & "GBN", 0, oRecordSet.Fields(26 + iCol).VALUE
        Next iCol
        
        '// 2010.04.05 최동권 추가
        oYM = Trim$(oDS_PH_PY132A.GetValue("U_YM", 0))
        If oYM <> "" Then
        
            iBNSMON = oRecordSet.Fields(1).VALUE * -1
            If iBNSMON <> 0 Then iBNSMON = iBNSMON + 1
            oRecordset2.DoQuery ("SELECT CONVERT(VARCHAR(6),DATEADD(MM, " & CStr(iBNSMON) & ", '" & oYM & "01'),112) FROM OADM")
            
            
        End If
    End If
    
    oForm.Update

    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
End Sub
