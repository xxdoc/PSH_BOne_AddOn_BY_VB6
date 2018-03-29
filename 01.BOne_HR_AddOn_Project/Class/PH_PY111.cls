VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY111.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 급상여 계산
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Private oDS_PH_PY111A   As SAPbouiCOM.DBDataSource


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
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY111.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY111_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY111"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "DocEntry"
    
    
    oForm.Freeze True
    Call PH_PY111_CreateItems
    Call PH_PY111_EnableMenus
    Call PH_PY111_SetDocument(oFromDocEntry01)
'    Call PH_PY111_FormResize
    
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

Private Function PH_PY111_CreateItems() As Boolean

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
    
    On Error GoTo PH_PY111_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------------------
    '// 데이터셋정의
    '//----------------------------------------------------------------------------------------------
    Set oDS_PH_PY111A = oForm.DataSources.DBDataSources("@PH_PY111A")   '//헤더
    
    
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
'/ 사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
    oForm.Items("CLTCOD").DisplayDesc = True

    
'/ 귀속연월
    oDS_PH_PY111A.setValue "U_YM", 0, Format$(Now, "YYYYMM")
    oDS_PH_PY111A.setValue "U_CHGCHK", 0, "N"
    oDS_PH_PY111A.setValue "U_TAXCHK", 0, "Y"  '//소득세계산
'/ 지급종류
    Set oCombo = oForm.Items("JOBTYP").Specific
    oCombo.ValidValues.Add "1", "급여"
    oCombo.ValidValues.Add "2", "상여"
    oForm.Items("JOBTYP").DisplayDesc = True
    oCombo.Select "1", psk_ByValue
    
'/ 지급구분
    Set oCombo = oForm.Items("JOBGBN").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JOBGBN").DisplayDesc = True
    oCombo.Select "1", psk_ByValue
    
'/ 급여대상자구분
    Set oCombo = oForm.Items("JOBTRG").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "모두"
    oForm.Items("JOBTRG").DisplayDesc = True
    oCombo.Select "1", psk_ByValue

'/ 직원구분
    Set oCombo = oForm.Items("JIGTYP").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P126' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "모두"
    oForm.Items("JIGTYP").DisplayDesc = True
    oCombo.Select "%", psk_ByValue
    
'/ 부서
    Set oCombo = oForm.Items("STRDPT").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='1' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "모두"
    oForm.Items("STRDPT").DisplayDesc = True
    oCombo.Select "%", psk_ByValue
    
    Set oCombo = oForm.Items("ENDDPT").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='1' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "모두"
    oForm.Items("ENDDPT").DisplayDesc = True
    
    oCombo.Select "%", psk_ByValue
    
'    Call oForm.DataSources.UserDataSources.Add("Col0", dt_SHORT_TEXT)
'    Call oForm.DataSources.UserDataSources.Add("Col1", dt_SHORT_TEXT)
'    Set oColumn = oMat1.Columns("Col0")
'    oColumn.DataBind.SetBound True, "", "Col0"
'    Set oColumn = oMat1.Columns("Col1")
'    oColumn.DataBind.SetBound True, "", "Col1"

  '/ Check 버튼
    Set oCheck = oForm.Items("JSNCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    Set oCheck = oForm.Items("YCHCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    Set oCheck = oForm.Items("RETCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    oCheck.Checked = True
    Set oCheck = oForm.Items("TAXCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    oCheck.Checked = True
    Set oCheck = oForm.Items("GBHCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    oCheck.Checked = True
    Set oCheck = oForm.Items("CHGCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    
'// 상여
    For iCol = 1 To 8
        Set oCombo = oForm.Items("AP" & iCol & "GBN").Specific
        oCombo.ValidValues.Add "1", "개월 이상"
        oCombo.ValidValues.Add "2", "개월 미만"
        oCombo.ValidValues.Add "3", "일수 이상"
        oCombo.ValidValues.Add "4", "일수 미만"
        If oCombo.ValidValues.Count > 0 Then
           oDS_PH_PY111A.setValue "U_AP" & iCol & "GBN", 0, "1"
        End If
        oForm.Items("AP" & iCol & "GBN").DisplayDesc = True
    Next iCol

'// 상여계산방법
    Set oCombo = oForm.Items("bBNSCAL").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P215' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("bBNSCAL").DisplayDesc = True

'/ 91 기초세액 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    For iCol = 1 To 10
        WG01.TB1AMT(iCol) = 0:  WG01.TB1GON(iCol) = 0:  WG01.TB1RAT(iCol) = 0:  WG01.TB1KUM(iCol) = 0
    Next iCol
'/
    sQry = "SELECT U_CODNBR, U_CODAMT, U_CODGON, U_CODRAT, U_CODKUM "
    sQry = sQry & " FROM [@PH_PY100B] WHERE CODE = (SELECT Top 1 Code FROM [@PH_PY100A] WHERE Code <='" & Format$(Now, "YYYY") & "' ORDER BY Code Desc) "
    sQry = sQry & " AND U_CODNBR BETWEEN '0001' AND '0010' ORDER BY CODE, U_CODNBR DESC"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
       WG01.TB1AMT(Val(oRecordSet.Fields("U_CODNBR").VALUE)) = Val(oRecordSet.Fields("U_CODAMT").VALUE)
       WG01.TB1GON(Val(oRecordSet.Fields("U_CODNBR").VALUE)) = Val(oRecordSet.Fields("U_CODGON").VALUE)
       WG01.TB1RAT(Val(oRecordSet.Fields("U_CODNBR").VALUE)) = Val(oRecordSet.Fields("U_CODRAT").VALUE)
       WG01.TB1KUM(Val(oRecordSet.Fields("U_CODNBR").VALUE)) = Val(oRecordSet.Fields("U_CODKUM").VALUE)
       oRecordSet.MoveNext
    Loop
        
    CalcYN = False

    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY111_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY111_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY111_EnableMenus()
    
    On Error GoTo PH_PY111_EnableMenus_Error

    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1287", True)     '//복제
    
    Exit Sub
    
PH_PY111_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY111_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY111_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY111_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY111_FormItemEnabled
        Call PH_PY111_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY111_FormItemEnabled
        oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY111_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY111_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY111_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY111_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        Call PH_PY111_FormClear
        oForm.ActiveItem = "CLTCOD"
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("Btn1").Visible = True
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        oCLTCOD = Trim$(oDS_PH_PY111A.GetValue("U_CLTCOD", 0))
        '/ 귀속연월
            oDS_PH_PY111A.setValue "U_YM", 0, Format$(Now, "YYYYMM")
            oDS_PH_PY111A.setValue "U_CHGCHK", 0, "N"
        '/ 지급종류
            Set oCombo = oForm.Items("JOBTYP").Specific
            oCombo.Select "1", psk_ByValue
        '/ 지급구분
            Set oCombo = oForm.Items("JOBGBN").Specific
            oCombo.Select "1", psk_ByValue
        '/ 급여대상자구분
            Set oCombo = oForm.Items("JOBTRG").Specific
            oCombo.Select "1", psk_ByValue
    
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("DocEntry").Enabled = True
        oForm.Items("Btn1").Visible = False
        oForm.ActiveItem = "DocEntry"
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        oCLTCOD = Trim$(oDS_PH_PY111A.GetValue("U_CLTCOD", 0))
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
'        oForm.Mode = fm_UPDATE_MODE
        oForm.ActiveItem = "CLTCOD"
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("Btn1").Visible = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        oCLTCOD = Trim$(oDS_PH_PY111A.GetValue("U_CLTCOD", 0))
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY111_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY111_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    tDocEntry = oForm.Items("DocEntry").Specific.VALUE
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY111_DataValidCheck = False Then  '/ 유효성 검사
                            BubbleEvent = False
                        End If
                    End If
                ElseIf pval.ItemUID = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then   '/ ChooseBtn사원리스트
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUID = "Btn1" Then       '/ 급(상)여계산
                    CalcYN = True
                    tDocEntry = oForm.Items("DocEntry").Specific.VALUE
                    If PH_PY111_DataValidCheck = True And Pay_Calc = True Then  '/ 유효성 검사
                        If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                            oForm.Items("1").CLICK ct_Regular
                        ElseIf oForm.Mode = fm_OK_MODE Then
                            Sbo_Application.StatusBar.SetText "급상여계산이 진행중입니다.", bmt_Short, smt_Warning
                            If oForm.Items("JOBGBN").Specific.VALUE <> "2" Then
                                '//정상계산
                                oRecordSet.DoQuery "EXEC PH_PY111 '" & tDocEntry & "'"
                            Else
                                '//소급계산
                                oRecordSet.DoQuery "EXEC PH_PY111_SOGUB '" & tDocEntry & "'"
                            End If
                            Sbo_Application.StatusBar.SetText "급상여계산이 완료되었습니다.", bmt_Short, smt_Success
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
                                Sbo_Application.StatusBar.SetText "급상여계산이 진행중입니다.", bmt_Short, smt_Warning
                                If oForm.Items("JOBGBN").Specific.VALUE <> "2" Then
                                '//정상계산
                                    oRecordSet.DoQuery "EXEC PH_PY111 '" & tDocEntry & "'"
                                Else
                                    '//소급계산
                                    oRecordSet.DoQuery "EXEC PH_PY111_SOGUB '" & tDocEntry & "'"
                                End If
                            
                                Sbo_Application.StatusBar.SetText "급상여계산이 완료되었습니다.", bmt_Short, smt_Success
                                oForm.Mode = fm_FIND_MODE
                                PH_PY111_FormItemEnabled
                                oForm.Items("DocEntry").Specific.VALUE = tDocEntry
                                oForm.Items("1").CLICK ct_Regular
                            End If
                        ElseIf CalcYN = False Then
                            If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                                oForm.Mode = fm_FIND_MODE
                                Call PH_PY111_FormItemEnabled
                                oForm.Items("DocEntry").Specific.VALUE = tDocEntry
                                oForm.Items("1").CLICK ct_Regular
                            Else
                                Call PH_PY111_FormItemEnabled
                                
                            End If
                        
                        End If
                    End If
                End If
                
                If pval.ItemUID = "CHGCHK" Then     '//수당/공제 계산식 반영제외
                    If Trim$(oDS_PH_PY111A.GetValue("U_CHGCHK", 0)) = "Y" Then
                        oDS_PH_PY111A.setValue "U_TAXCHK", 0, "N"          '//소득세 계산함
                        oDS_PH_PY111A.setValue "U_GBHCHK", 0, "N"           '//고용보험 계산함
                    End If
                End If
                If pval.ItemUID = "TAZCHK" Or pval.ItemUID = "GBHCHK" Then
                    If Trim$(oDS_PH_PY111A.GetValue("U_TAXCHK", 0)) = "Y" Or Trim$(oDS_PH_PY111A.GetValue("U_GBHCHK", 0)) = "Y" Then
                        oDS_PH_PY111A.setValue "U_CHGCHK", 0, "N"
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
            ElseIf pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 Then
               If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oForm.Items(pval.ItemUID).Specific.String + "'", "") = True Then
                   oForm.Items("MSTCOD").CLICK ct_Regular
                   Sbo_Application.ActivateMenuItem ("7425")
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
                '// 사업장
                If pval.ItemUID = "CLTCOD" Then
                    If Not oForm.Items("CLTCOD").Specific.Selected Is Nothing Then
                        oCLTCOD = Trim$(oDS_PH_PY111A.GetValue("U_CLTCOD", 0))
                        If oCLTCOD <> "" Then Display_BonussRate
                    Else
                        oJOBTYP = ""
                    End If
                End If
                '// 지급종류
                If pval.ItemUID = "JOBTYP" Then
                    If Not oForm.Items("JOBTYP").Specific.Selected Is Nothing Then
                        oJOBTYP = Trim$(oForm.Items("JOBTYP").Specific.Selected.VALUE)
                        If oJOBTYP <> "" Then Display_BonussRate
                    Else
                        oJOBTYP = ""
                    End If
                    If oYM <= "201012" Then
                        oDS_PH_PY111A.setValue "U_GBHCHK", 0, "Y"
                    ElseIf (oJOBTYP = "1" Or oJOBTYP = "3") And oJOBGBN = "1" Then
                        oDS_PH_PY111A.setValue "U_GBHCHK", 0, "Y"
                        
                    Else
                        oDS_PH_PY111A.setValue "U_GBHCHK", 0, "N"
                    End If
                    oForm.Items("GBHCHK").Update
                End If
                '//지급구분
                If pval.ItemUID = "JOBGBN" Then
                    If Not oForm.Items("JOBGBN").Specific.Selected Is Nothing Then
                        oJOBGBN = Trim$(oDS_PH_PY111A.GetValue("U_JOBGBN", 0))
                    Else
                        oJOBGBN = ""
                    End If
                    If oJOBGBN <> "" Then Call Display_BonussRate
                    If oYM <= "201012" Then
                        oDS_PH_PY111A.setValue "U_GBHCHK", 0, "Y"
                    ElseIf (oJOBTYP = "1" Or oJOBTYP = "3") And oJOBGBN = "1" Then
                        oDS_PH_PY111A.setValue "U_GBHCHK", 0, "Y"
                    Else
                        oDS_PH_PY111A.setValue "U_GBHCHK", 0, "N"
                    End If
                    oForm.Items("GBHCHK").Update
                End If
                '//급여지급대상일
                
                If pval.ItemUID = "JOBTRG" Then
                    If Not oForm.Items("JOBTRG").Specific.Selected Is Nothing Then
                        oJOBTRG = Trim$(oDS_PH_PY111A.GetValue("U_JOBTRG", 0))
                        If Trim$(oDS_PH_PY111A.GetValue("U_YM", 0)) <> "" Then
                            oYM = Trim$(oDS_PH_PY111A.GetValue("U_YM", 0))
                            oRecordSet.DoQuery " SELECT DBO.Func_PAYTerm( '" & Format$(oYM, "0000-00") & "-01" & "', '" & IIf(oJOBTRG = "%", "1", oJOBTRG) & "')"
                            If oRecordSet.RecordCount > 0 Then
                                JIGBIL = Mid$(oRecordSet.Fields(0).VALUE, 19, 8)
                                oGNEDAT = Mid$(oRecordSet.Fields(0).VALUE, 10, 8)
                                oEXPDAT = Mid$(oRecordSet.Fields(0).VALUE, 28, 8)
                                oDS_PH_PY111A.setValue "U_JIGBIL", 0, JIGBIL
                                oDS_PH_PY111A.setValue "U_bGNEDAT", 0, oGNEDAT
                                oDS_PH_PY111A.setValue "U_bEXPDAT", 0, oEXPDAT
                               
                                '// 2010.04.05 최동권 수정(상여지급 주기로 자동 계산)
            '                    oDS_PH_PY111A.setValue "U_bPAYSTR", 0,  Mid$(oRecordSet.Fields(0).Value, 1, 6)
            '                    oDS_PH_PY111A.setValue "U_bPAYEND", 0,  Mid$(oRecordSet.Fields(0).Value, 10, 6)
            '                    oDS_PH_PY111A.setValue "U_bBNSSTR", 0,  Mid$(oRecordSet.Fields(0).Value, 1, 6)
            '                    oDS_PH_PY111A.setValue "U_bBNSEND", 0,  Mid$(oRecordSet.Fields(0).Value, 10, 6)
                                
                                oForm.Items("JIGBIL").Update
                                oForm.Items("bGNEDAT").Update
                                oForm.Items("bEXPDAT").Update
            '                    oForm.Items("bPAYSTR").Update
            '                    oForm.Items("bPAYEND").Update
            '                    oForm.Items("bBNSSTR").Update
            '                    oForm.Items("bBNSEND").Update
                            End If
                            Set oRecordSet = Nothing
                            
                        End If
                    Else
                        oJOBTRG = ""
                    End If
                End If
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
                        oDS_PH_PY111A.setValue "U_YM", 0, Format$(Now, "YYYYMM")
                    Else
                        oDS_PH_PY111A.setValue "U_YM", 0, oForm.Items("YM").Specific.String
                    End If
                    oYM = Trim$(oDS_PH_PY111A.GetValue("U_YM", 0))
                    If oJOBTRG <> "" Then
                        oRecordSet.DoQuery " SELECT DBO.Func_PAYTerm( '" & Format$(oYM, "0000-00") & "-01" & "', '" & IIf(oJOBTRG = "%", "1", oJOBTRG) & "')"
                        If oRecordSet.RecordCount > 0 Then
                            JIGBIL = Mid$(oRecordSet.Fields(0).VALUE, 19, 8)
                            oGNEDAT = Mid$(oRecordSet.Fields(0).VALUE, 10, 8)
                            oEXPDAT = Mid$(oRecordSet.Fields(0).VALUE, 28, 8)
                            oDS_PH_PY111A.setValue "U_bGNEDAT", 0, oGNEDAT
                            oDS_PH_PY111A.setValue "U_bEXPDAT", 0, oEXPDAT
                            oDS_PH_PY111A.setValue "U_JIGBIL", 0, JIGBIL
                            oForm.Items("JIGBIL").Update
                        End If
                        Set oRecordSet = Nothing
                    End If
                    If oYM <> "" Then Call Display_BonussRate
                End If
                '//사번
                If pval.ItemUID = "MSTCOD" Then
                    If Trim$(oForm.Items("MSTCOD").Specific.VALUE) <> "" Then
                        oDS_PH_PY111A.setValue "U_MSTCOD", 0, oForm.Items("MSTCOD").Specific.String
                        oDS_PH_PY111A.setValue "U_MSTNAM", 0, MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.String & "'", "")
                    Else
                        oDS_PH_PY111A.setValue "U_MSTNAM", 0, ""
                    End If
                    oForm.Items("MSTNAM").Update
                End If
'                '//세액 대상기간(상여)
'                If pval.ItemUID = "bBNSSTR" Then
'                    If Trim$(oForm.Items("bBNSSTR").Specific.Value) = "" Then
'                        oDS_PH_PY111A.setValue "U_bBNSSTR", 0, ""
'                        oDS_PH_PY111A.setValue "U_bPAYSTR", 0, ""
'                    Else
'                        oDS_PH_PY111A.setValue "U_bBNSSTR", 0, oForm.Items("bBNSSTR").Specific.Value
'                        oDS_PH_PY111A.setValue "U_bPAYSTR", 0, oForm.Items("bBNSSTR").Specific.Value
'                    End If
'                    oForm.Items("bPAYSTR").Update
'                End If
'                '//세액 대상기간(상여)
'                If pval.ItemUID = "bBNSEND" Then
'                    If Trim$(oForm.Items("bBNSEND").Specific.Value) = "" Then
'                        oDS_PH_PY111A.setValue "U_bBNSEND", 0, ""
'                        oDS_PH_PY111A.setValue "U_bPAYEND", 0, ""
'                    Else
'                        oDS_PH_PY111A.setValue "U_bBNSEND", 0, oForm.Items("bBNSEND").Specific.Value
'                        oDS_PH_PY111A.setValue "U_bPAYEND", 0, oForm.Items("bBNSEND").Specific.Value
'                    End If
'                    oForm.Items("bPAYEND").Update
'                End If
            End If
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11

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
                Set oDS_PH_PY111A = Nothing
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
'                If oForm.Items("DocEntry").Specific.VALUE = 0 Then
'                    Sbo_Application.MessageBox ("퇴직급여 계산용입니다. 삭제할 수 없습니다.")
'                    BubbleEvent = False
'                    Exit Sub
'                End If
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY111A", "DocEntry")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY111_FormItemEnabled
                Call PH_PY111_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY111_FormItemEnabled
                Call PH_PY111_AddMatrixRow
'                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY111_FormItemEnabled
                Call PH_PY111_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY111_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY111B.Size - 1)
'                        If oDS_PH_PY111B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY111B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY111B.Size
'                        Call oDS_PH_PY111B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
'                End If
'                Call PH_PY111_AddMatrixRow
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

Sub PH_PY111_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY111_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 용]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY111B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY111B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY111B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY111B.Offset = oRow
'            oDS_PH_PY111B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY111B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY111B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY111B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY111B.Offset = oRow - 1
'            oDS_PH_PY111B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY111B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY111B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY111B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY111B.Offset = oRow
'        oDS_PH_PY111B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY111B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY111B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY111B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY111_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY111_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY111_FormClear()
On Error GoTo PH_PY111_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY111'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY111_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY111_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY111_DataValidCheck() As Boolean
On Error GoTo PH_PY111_DataValidCheck_Error
    Dim ErrNum          As Integer
    ErrNum = 0
    
    PH_PY111_DataValidCheck = False
    
    '/
    oYM = Trim$(oDS_PH_PY111A.GetValue("U_YM", 0))
    oJOBTYP = Trim$(oDS_PH_PY111A.GetValue("U_JOBTYP", 0))
    oJOBGBN = Trim$(oDS_PH_PY111A.GetValue("U_JOBGBN", 0))
    oCLTCOD = Trim$(oDS_PH_PY111A.GetValue("U_CLTCOD", 0))
    
    oSTRDPT = Trim$(oDS_PH_PY111A.GetValue("U_STRDPT", 0))
    oENDDPT = Trim$(oDS_PH_PY111A.GetValue("U_ENDDPT", 0))
    If oENDDPT = "%" Then oENDDPT = "ZZZZZZZZ"
    oJIGBIL = oDS_PH_PY111A.GetValue("U_JIGBIL", 0)
    oMSTCOD = Trim$(oDS_PH_PY111A.GetValue("U_MSTCOD", 0))
    oJSNCHK = oDS_PH_PY111A.GetValue("U_JSNCHK", 0)
    oYCHCHK = oDS_PH_PY111A.GetValue("U_YCHCHK", 0)
'/ 상여관련
    oBNSCAL = oDS_PH_PY111A.GetValue("U_bBNSCAL", 0)
    oBNSRAT = oDS_PH_PY111A.GetValue("U_bBNSRAT", 0)
    oSTRTAX = Trim$(oDS_PH_PY111A.GetValue("U_bPAYSTR", 0))
    oENDTAX = Trim$(oDS_PH_PY111A.GetValue("U_bPAYEND", 0))
    oSTRBNS = Trim$(oDS_PH_PY111A.GetValue("U_bBNSSTR", 0))
    oENDBNS = Trim$(oDS_PH_PY111A.GetValue("U_bBNSEND", 0))
    
    oRETCHK = oDS_PH_PY111A.GetValue("U_RETCHK", 0)
    oGNEDAT = oDS_PH_PY111A.GetValue("U_bGNEDAT", 0)
 '/ Check
    Select Case True
      Case MDC_SetMod.ChkYearMonth(oYM) = False
           Sbo_Application.StatusBar.SetText "귀속 연월을 확인하여 주십시오.", bmt_Short, smt_Error
           Exit Function
      Case oJOBTYP = ""
           Sbo_Application.StatusBar.SetText "지급 종류를 선택하여 주십시오.", bmt_Short, smt_Error
           Exit Function
      Case oJOBGBN = ""
           Sbo_Application.StatusBar.SetText "지급 구분을 선택하여 주십시오.", bmt_Short, smt_Error
           Exit Function
      Case oJIGBIL = ""
           Sbo_Application.StatusBar.SetText "지급일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
           Exit Function
      Case oDS_PH_PY111A.GetValue("U_JOBTRG", 0) = ""
           Sbo_Application.StatusBar.SetText "지급 대상구분은 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
           Exit Function
    End Select
    oJOBTRG = Trim$(oForm.Items("JOBTRG").Specific.Selected.VALUE)

''/ 작업연월 잠김여부 체크
'    If MDC_SetMod.Get_PayLockInfo(oYM, oJOBTYP, oJOBGBN, oJOBTRG) = True Then
'        Sbo_Application.StatusBar.SetText "귀속연월은 마감처리되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error
'        Exit Function
'    End If
'/ 상여계산시 체크
    If oJOBTYP <> "1" Then
        Select Case True
        Case Val(oBNSRAT) = 0
           Sbo_Application.StatusBar.SetText "상여율은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
           Exit Function
        Case Len(oSTRTAX) <> 6, Len(oENDTAX) <> 6
           Sbo_Application.StatusBar.SetText "세액대상기간은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
           Exit Function
        Case oSTRTAX > oENDTAX
           Sbo_Application.StatusBar.SetText "세액대상기간이 올바르지 않습니다. 확인하여 주십시오.", bmt_Short, smt_Error
           Exit Function
        Case oDS_PH_PY111A.GetValue("U_bGNEDAT", 0) = ""
           Sbo_Application.StatusBar.SetText "상여상여계산기준일은 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
           Exit Function
        Case oDS_PH_PY111A.GetValue("U_bEXPDAT", 0) = ""
           Sbo_Application.StatusBar.SetText "상여지급제한일자는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
           Exit Function
        End Select
    End If
    
    
    If oJOBTYP = "1" Then
        If oJOBGBN = "3" Then
        '//잔여급여계산시 근태기준일 입력
            If oDS_PH_PY111A.GetValue("U_GTDateFr", 0) = "" Then
                Sbo_Application.StatusBar.SetText "잔여급여 계산시 근태기준일(시작)은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
                Exit Function
            End If
            If oDS_PH_PY111A.GetValue("U_GTDateTo", 0) = "" Then
                Sbo_Application.StatusBar.SetText "잔여급여 계산시 근태기준일(종료)은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
                Exit Function
            End If
        End If
    End If
    PH_PY111_DataValidCheck = True
    Exit Function
    
PH_PY111_DataValidCheck_Error:
    PH_PY111_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY111_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY111_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY111_Validate_Error
    PH_PY111_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY111A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY111_Validate = False
        GoTo PH_PY111_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY111_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY111_Validate_Error:
    PH_PY111_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY111_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PH_PY111_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    
    Dim BPLID           As String
    Dim ItmBsort        As String
    Dim DocDate         As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY111_Print_Report01_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY111_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[S142] 발주서"
    ReportName = "S142_1.rpt"
    sQry = "EXEC PH_PY111_1 '" & oForm.Items("8").Specific.VALUE & "'"
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '/ Procedure 실행"
    sQry = "EXEC [PS_PP820_01] '" & BPLID & "', '" & ItmBsort & "', '" & DocDate & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
            Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
        End If
    Else
        Sbo_Application.SetStatusBarMessage "조회된 데이터가 없습니다.", bmt_Short, True
    End If
    
    Exit Sub
    
PH_PY111_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY111_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function Pay_Calc() As Boolean
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim Temp1           As String
    Dim Temp2           As String
    Dim i               As Long
    
    Dim CLTCOD As String
    Dim YM As String
    Dim JOBTYP As String
    Dim JOBGBN As String
    Dim JOBTRG As String
    Dim JIGBIL As String
    Dim MSTCOD As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Pay_Calc = False
    
    
    CLTCOD = Trim(oDS_PH_PY111A.GetValue("U_CLTCOD", 0))
    YM = Trim(oDS_PH_PY111A.GetValue("U_YM", 0))
    JOBTYP = Trim(oDS_PH_PY111A.GetValue("U_JOBTYP", 0))
    JOBGBN = Trim(oDS_PH_PY111A.GetValue("U_JOBGBN", 0))
    JOBTRG = Trim(oDS_PH_PY111A.GetValue("U_JOBTRG", 0))
    JIGBIL = Trim(oDS_PH_PY111A.GetValue("U_JIGBIL", 0))
    MSTCOD = Trim(oDS_PH_PY111A.GetValue("U_MSTCOD", 0))
    
    sQry = "        Select  Count(*) "
    sQry = sQry & " From [@PH_PY112A] "
    sQry = sQry & " Where   U_CLTCOD = '" & CLTCOD & "' "
    sQry = sQry & "         AND U_YM = '" & YM & "'"
    sQry = sQry & "         AND U_JOBTYP = '" & JOBTYP & "'"
    sQry = sQry & "         AND U_JOBGBN = '" & JOBGBN & "'"
    sQry = sQry & "         AND U_JOBTRG = '" & JOBTRG & "'"
    sQry = sQry & "         AND U_JIGBIL = '" & JIGBIL & "'"
    sQry = sQry & "         AND U_MSTCOD LIKE '%" & MSTCOD & "%'"
    
    'sQry = "Select U_CLTCOD, U_YM, U_JOBTYP, U_JOBGBN, U_JOBTRG From [@PH_PY112A] Where U_CLTCOD = '" & Trim(oDS_PH_PY111A.GetValue("U_CLTCOD", 0)) & "' AND U_YM = '" & Trim(oDS_PH_PY111A.GetValue("U_YM", 0)) & "' AND U_JOBTYP = '" & Trim(oDS_PH_PY111A.GetValue("U_JOBTYP", 0)) & "' "
    oRecordSet.DoQuery sQry
        
    If Val(oRecordSet.Fields(0).VALUE) > 0 Then
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
    
'    Temp1 = Trim(oDS_PH_PY111A.GetValue("U_CLTCOD", 0)) & Trim(oDS_PH_PY111A.GetValue("U_YM", 0)) & Trim(oDS_PH_PY111A.GetValue("U_JOBTYP", 0)) & _
'            Trim(oDS_PH_PY111A.GetValue("U_JOBGBN", 0)) & Trim(oDS_PH_PY111A.GetValue("U_JOBTRG", 0))
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
    
    sQry = " SELECT U_BNSCAL, U_BNSMON, U_BNSRAT, U_AP1MON, U_AP2MON, U_AP3MON, U_AP4MON, U_AP5MON,"
    sQry = sQry & " U_AP6MON, U_AP7MON, U_AP8MON, U_AP1RAT, U_AP2RAT, U_AP3RAT, U_AP4RAT, U_AP5RAT, "
    sQry = sQry & " U_AP6RAT, U_AP7RAT, U_AP8RAT, U_AP1AMT, U_AP2AMT, U_AP3AMT, U_AP4AMT, U_AP5AMT,"
    sQry = sQry & " U_AP6AMT, U_AP7AMT, U_AP8AMT, U_AP1GBN, U_AP2GBN, U_AP3GBN, U_AP4GBN, U_AP5GBN,"
    sQry = sQry & " U_AP6GBN, U_AP7GBN, U_AP8GBN  FROM [@PH_PY108A] "
    sQry = sQry & " WHERE U_CLTCOD = '" & oCLTCOD & "'  AND U_JOBGBN = '" & JOBGBN & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Or oJOBTYP = "1" Then
        oDS_PH_PY111A.setValue "U_bBNSCAL", 0, "1"
        oDS_PH_PY111A.setValue "U_bBNSRAT", 0, 0
        oDS_PH_PY111A.setValue "U_bMONTH1", 0, 0
        oDS_PH_PY111A.setValue "U_bMONTH2", 0, 0
        oDS_PH_PY111A.setValue "U_bMONTH3", 0, 0
        oDS_PH_PY111A.setValue "U_bMONTH4", 0, 0
        oDS_PH_PY111A.setValue "U_bMONTH5", 0, 0
        oDS_PH_PY111A.setValue "U_bMONTH6", 0, 0
        oDS_PH_PY111A.setValue "U_bMONTH7", 0, 0
        oDS_PH_PY111A.setValue "U_bMONTH8", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPRAT1", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPRAT2", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPRAT3", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPRAT4", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPRAT5", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPRAT6", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPRAT7", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPRAT8", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPAMT1", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPAMT2", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPAMT3", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPAMT4", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPAMT5", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPAMT6", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPAMT7", 0, 0
        oDS_PH_PY111A.setValue "U_bAPPAMT8", 0, 0
        For iCol = 1 To 8
            oDS_PH_PY111A.setValue "U_AP" & iCol & "GBN", 0, "1"
        Next iCol
        
        '// 2010.04.05 최동권 추가
        oYM = Trim$(oDS_PH_PY111A.GetValue("U_YM", 0))
        
        oDS_PH_PY111A.setValue "U_bPAYSTR", 0, oYM
        oDS_PH_PY111A.setValue "U_bPAYEND", 0, oYM
        oDS_PH_PY111A.setValue "U_bBNSSTR", 0, oYM
        oDS_PH_PY111A.setValue "U_bBNSEND", 0, oYM
    
        oForm.Items("bPAYSTR").Update
        oForm.Items("bPAYEND").Update
        oForm.Items("bBNSSTR").Update
        oForm.Items("bBNSEND").Update
    Else
        oDS_PH_PY111A.setValue "U_bBNSCAL", 0, oRecordSet.Fields(0).VALUE
        oDS_PH_PY111A.setValue "U_bBNSRAT", 0, oRecordSet.Fields(2).VALUE
        oDS_PH_PY111A.setValue "U_bMONTH1", 0, oRecordSet.Fields(3).VALUE
        oDS_PH_PY111A.setValue "U_bMONTH2", 0, oRecordSet.Fields(4).VALUE
        oDS_PH_PY111A.setValue "U_bMONTH3", 0, oRecordSet.Fields(5).VALUE
        oDS_PH_PY111A.setValue "U_bMONTH4", 0, oRecordSet.Fields(6).VALUE
        oDS_PH_PY111A.setValue "U_bMONTH5", 0, oRecordSet.Fields(7).VALUE
        oDS_PH_PY111A.setValue "U_bMONTH6", 0, oRecordSet.Fields(8).VALUE
        oDS_PH_PY111A.setValue "U_bMONTH7", 0, oRecordSet.Fields(9).VALUE
        oDS_PH_PY111A.setValue "U_bMONTH8", 0, oRecordSet.Fields(10).VALUE
        oDS_PH_PY111A.setValue "U_bAPPRAT1", 0, oRecordSet.Fields(11).VALUE
        oDS_PH_PY111A.setValue "U_bAPPRAT2", 0, oRecordSet.Fields(12).VALUE
        oDS_PH_PY111A.setValue "U_bAPPRAT3", 0, oRecordSet.Fields(13).VALUE
        oDS_PH_PY111A.setValue "U_bAPPRAT4", 0, oRecordSet.Fields(14).VALUE
        oDS_PH_PY111A.setValue "U_bAPPRAT5", 0, oRecordSet.Fields(15).VALUE
        oDS_PH_PY111A.setValue "U_bAPPRAT6", 0, oRecordSet.Fields(16).VALUE
        oDS_PH_PY111A.setValue "U_bAPPRAT7", 0, oRecordSet.Fields(17).VALUE
        oDS_PH_PY111A.setValue "U_bAPPRAT8", 0, oRecordSet.Fields(18).VALUE
        oDS_PH_PY111A.setValue "U_bAPPAMT1", 0, oRecordSet.Fields(19).VALUE
        oDS_PH_PY111A.setValue "U_bAPPAMT2", 0, oRecordSet.Fields(20).VALUE
        oDS_PH_PY111A.setValue "U_bAPPAMT3", 0, oRecordSet.Fields(21).VALUE
        oDS_PH_PY111A.setValue "U_bAPPAMT4", 0, oRecordSet.Fields(22).VALUE
        oDS_PH_PY111A.setValue "U_bAPPAMT5", 0, oRecordSet.Fields(23).VALUE
        oDS_PH_PY111A.setValue "U_bAPPAMT6", 0, oRecordSet.Fields(24).VALUE
        oDS_PH_PY111A.setValue "U_bAPPAMT7", 0, oRecordSet.Fields(25).VALUE
        oDS_PH_PY111A.setValue "U_bAPPAMT8", 0, oRecordSet.Fields(26).VALUE
        For iCol = 1 To 8
'            oForm.DataSources.UserDataSources("AP" & iCol & "GBN").ValueEx = oRecordSet.Fields(26 + iCol).Value
            oDS_PH_PY111A.setValue "U_AP" & iCol & "GBN", 0, oRecordSet.Fields(26 + iCol).VALUE
        Next iCol
        
        '// 2010.04.05 최동권 추가
        oYM = Trim$(oDS_PH_PY111A.GetValue("U_YM", 0))
        If oYM <> "" Then
        
            iBNSMON = oRecordSet.Fields(1).VALUE * -1
            If iBNSMON <> 0 Then iBNSMON = iBNSMON + 1
            oRecordset2.DoQuery ("SELECT CONVERT(VARCHAR(6),DATEADD(MM, " & CStr(iBNSMON) & ", '" & oYM & "01'),112) FROM OADM")
            
            oDS_PH_PY111A.setValue "U_bPAYSTR", 0, oRecordset2.Fields(0).VALUE
            oDS_PH_PY111A.setValue "U_bPAYEND", 0, oYM
            oDS_PH_PY111A.setValue "U_bBNSSTR", 0, oRecordset2.Fields(0).VALUE
            oDS_PH_PY111A.setValue "U_bBNSEND", 0, oYM
        
            oForm.Items("bPAYSTR").Update
            oForm.Items("bPAYEND").Update
            oForm.Items("bBNSSTR").Update
            oForm.Items("bBNSEND").Update
        End If
    End If
    
    oForm.Update

    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Create_Tiltle = "Display_BonussRate :" & Space$(10) & Err.Description
End Sub

Private Function Create_Tiltle() As Variant
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
    Dim CSUCNT  As Integer, GONCNT  As Integer
    Dim iCol    As Integer
'/
    ErrNum = 0
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'/ 1.항목명 초기화
    CSUCNT = 0:    GONCNT = 0
    For iCol = 1 To 24
        WK_C.CSUCOD(iCol) = ""
        WK_C.CSUNAM(iCol) = "---"
        WK_C.CSUKUM(iCol) = 0
        WK_C.MPYGBN(iCol) = "N"
        WK_C.GWATYP(iCol) = ""
        WK_C.GBHGBN(iCol) = ""
        WK_C.ROUNDT(iCol) = "R"
        WK_C.RODLEN(iCol) = 1
        WK_C.GONSIL(iCol) = ""
        WK_C.BNSUSE(iCol) = "N"
        WK_C.BTXCOD(iCol) = ""
    Next iCol
    For iCol = 1 To 18
        WK_G.GONCOD(iCol) = ""
        WK_G.GONNAM(iCol) = "---"
        WK_G.BNSUSE(iCol) = "N"
        WK_G.GONSIL(iCol) = ""
        WK_G.ROUNDT(iCol) = "R"
        WK_G.RODLEN(iCol) = 1
    Next iCol
'/ 2.수당항목 가져오기
    sQry = "Exec PH_PY102  '" & oCLTCOD & "' , '" & oYM & "', '', '', '', ''"
    oRecordSet.DoQuery sQry
    CSUCNT = oRecordSet.RecordCount
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1    '/ 수당항목관리를 확인하세요.
        GoTo Error_Message
    Else
        iCol = 1
        Do Until oRecordSet.EOF
            WK_C.CSUCOD(iCol) = oRecordSet.Fields("U_CSUCOD").VALUE
            WK_C.CSUNAM(iCol) = oRecordSet.Fields("U_CSUNAM").VALUE
            WK_C.CSUKUM(iCol) = oRecordSet.Fields("U_KUMAMT").VALUE
            WK_C.MPYGBN(iCol) = oRecordSet.Fields("U_MONPAY").VALUE
            WK_C.GWATYP(iCol) = oRecordSet.Fields("U_GWATYP").VALUE
            WK_C.GBHGBN(iCol) = oRecordSet.Fields("U_GBHGBN").VALUE
            WK_C.ROUNDT(iCol) = oRecordSet.Fields("U_ROUNDT").VALUE
            WK_C.RODLEN(iCol) = CInt(oRecordSet.Fields("U_LENGTH").VALUE)
            WK_C.BNSUSE(iCol) = oRecordSet.Fields("U_BNSUSE").VALUE
            WK_C.BTXCOD(iCol) = oRecordSet.Fields("U_BTXCOD").VALUE
            U_CSUCOD = oRecordSet.Fields("Code").VALUE
            iCol = iCol + 1
            oRecordSet.MoveNext
        Loop
    End If
'/ 3.공제항목 가져오기
    sQry = "Exec PH_PY103  '" & oCLTCOD & "' , '" & oYM & "',  '', ''"
    oRecordSet.DoQuery sQry
    GONCNT = oRecordSet.RecordCount
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 2    '/ 공제항목관리를 확인하세요.
        GoTo Error_Message
    Else
        iCol = 1
        G06_CHK = False   '/ 소득세정산
        G07_CHK = False   '/ 주민세정산
        G08_CHK = False   '/ 건강보험정산
        G90_CHK = False   '/ 국민연금정산
        G91_CHK = False   '/ 농특세정산
        Do Until oRecordSet.EOF
            WK_G.GONCOD(iCol) = oRecordSet.Fields("U_CSUCOD").VALUE
            WK_G.GONNAM(iCol) = oRecordSet.Fields("U_CSUNAM").VALUE
            WK_G.BNSUSE(iCol) = oRecordSet.Fields("U_BNSUSE").VALUE     '/ 상여
            WK_G.GONSIL(iCol) = oRecordSet.Fields("U_SILCUN").VALUE     '/ 급여계산식
            WK_G.ROUNDT(iCol) = oRecordSet.Fields("U_ROUNDT").VALUE     '/ 끝전처리
            WK_G.RODLEN(iCol) = CInt(oRecordSet.Fields("U_LENGTH").VALUE)
            If Trim$(WK_G.GONCOD(iCol)) = "G06" Then G06_CHK = True
            If Trim$(WK_G.GONCOD(iCol)) = "G07" Then G07_CHK = True
            If Trim$(WK_G.GONCOD(iCol)) = "G08" Then G08_CHK = True
            If Trim$(WK_G.GONCOD(iCol)) = "G90" Then G90_CHK = True
            If Trim$(WK_G.GONCOD(iCol)) = "G91" Then G91_CHK = True
            If Trim$(WK_G.GONCOD(iCol)) = "G04" Then G04_BNSUSE = oRecordSet.Fields("U_BNSUSE").VALUE
            U_GONCOD = oRecordSet.Fields("Code").VALUE
            iCol = iCol + 1
            oRecordSet.MoveNext
        Loop
    End If
'/ 4.수당타이틀저장
    sQry = "SELECT CODE FROM [@PH_PY111T] WHERE U_YM = '" & oYM & "'"
    sQry = sQry & " AND U_JOBTYP = '" & oJOBTYP & "' AND U_JOBGBN = '" & oJOBGBN & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        '/ 1.1) Insert
        sQry = "INSERT INTO [@PH_PY111T]  (Code, Name, U_YM, U_JOBTYP, U_JOBGBN,"
        sQry = sQry & " U_CSUD01, U_CSUD02, U_CSUD03, U_CSUD04, U_CSUD05, U_CSUD06, U_CSUD07, U_CSUD08,"
        sQry = sQry & " U_CSUD09, U_CSUD10, U_CSUD11, U_CSUD12, U_CSUD13, U_CSUD14, U_CSUD15, U_CSUD16,"
        sQry = sQry & " U_CSUD17, U_CSUD18, U_CSUD19, U_CSUD20, U_CSUD21, U_CSUD22, U_CSUD23, U_CSUD24,"
        sQry = sQry & " U_GONG01, U_GONG02, U_GONG03, U_GONG04, U_GONG05, U_GONG06, U_GONG07, U_GONG08,"
        sQry = sQry & " U_GONG09, U_GONG10, U_GONG11, U_GONG12, U_GONG13, U_GONG14, U_GONG15, U_GONG16,"
        sQry = sQry & " U_GONG17, U_GONG18,"
        sQry = sQry & " U_CSUC01, U_CSUC02, U_CSUC03, U_CSUC04, U_CSUC05, U_CSUC06, U_CSUC07, U_CSUC08,"
        sQry = sQry & " U_CSUC09, U_CSUC10, U_CSUC11, U_CSUC12, U_CSUC13, U_CSUC14, U_CSUC15, U_CSUC16,"
        sQry = sQry & " U_CSUC17, U_CSUC18, U_CSUC19, U_CSUC20, U_CSUC21, U_CSUC22, U_CSUC23, U_CSUC24,"
        sQry = sQry & " U_GONC01, U_GONC02, U_GONC03, U_GONC04, U_GONC05, U_GONC06, U_GONC07, U_GONC08,"
        sQry = sQry & " U_GONC09, U_GONC10, U_GONC11, U_GONC12, U_GONC13, U_GONC14, U_GONC15, U_GONC16,"
        sQry = sQry & " U_GONC17, U_GONC18) VALUES ( "
        sQry = sQry & "  '" & oYM & oJOBTYP & oJOBGBN & "'"
        sQry = sQry & ", '" & oYM & oJOBTYP & oJOBGBN & "'"
        sQry = sQry & ", '" & oYM & " ', '" & oJOBTYP & "', '" & oJOBGBN & "'"
        For iCol = 1 To 24
        sQry = sQry & ", N'" & WK_C.CSUNAM(iCol) & "'"
        Next iCol
        For iCol = 1 To 18
        sQry = sQry & ", N'" & WK_G.GONNAM(iCol) & "'"
        Next iCol
        For iCol = 1 To 24
        sQry = sQry & ", N'" & WK_C.CSUCOD(iCol) & "'"
        Next iCol
        For iCol = 1 To 18
        sQry = sQry & ", N'" & WK_G.GONCOD(iCol) & "'"
        Next iCol
        sQry = sQry & ")"
        oRecordSet.DoQuery sQry
    Else
        '/ 1.2) Update
        sQry = "UPDATE [@PH_PY111T] SET U_JOBGBN = '" & oJOBGBN & "'"
        For iCol = 1 To 24
            sQry = sQry & ", U_CSUD" & Format$(iCol, "00") & " = N'" & WK_C.CSUNAM(iCol) & "'"
        Next iCol
        For iCol = 1 To 18
            sQry = sQry & ", U_GONG" & Format$(iCol, "00") & " = N'" & WK_G.GONNAM(iCol) & "'"
        Next iCol
        For iCol = 1 To 24
            sQry = sQry & ", U_CSUC" & Format$(iCol, "00") & " = N'" & WK_C.CSUCOD(iCol) & "'"
        Next iCol
        For iCol = 1 To 18
            sQry = sQry & ", U_GONC" & Format$(iCol, "00") & " = N'" & WK_G.GONCOD(iCol) & "'"
        Next iCol
        sQry = sQry & " WHERE   U_YM = '" & oYM & "'"
        sQry = sQry & " AND     U_JOBTYP = '" & oJOBTYP & "'"
        sQry = sQry & " AND     U_JOBGBN = '" & oJOBGBN & "'"
        oRecordSet.DoQuery sQry
    End If

    Set oRecordSet = Nothing
    Create_Tiltle = ""
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Create_Tiltle = "(PH_PY102)수당항목자료가 없습니다. 확인하여 주십시오. "
    ElseIf ErrNum = 2 Then
        Create_Tiltle = "(PH_PY103)공제항목자료가 없습니다. 확인하여 주십시오. "
    Else
        Create_Tiltle = "Create_Title :" & Space$(10) & Err.Description
    End If
End Function

Private Sub Execution_Process()
On Error GoTo Error_Message
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim oStrE           As String
    Dim MSTCOD          As String
    Dim oRow            As Long
    Dim Result          As String
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Dim V_StatusCnt&, oProValue&, TOTCNT&  'progbar
'    oMat1.Clear
    MaxRow = 0
    ErrNum = 0
    REMARK1 = ""
    REMARK2 = ""
    REMARK3 = ""
'/1. 수당타이틀구함
    oStrE = Create_Tiltle    '/ 수당타이틀 생성
    If oStrE <> "" Then
        ErrNum = 1
        GoTo Error_Message
    End If
 '/ Question
    If Sbo_Application.MessageBox("급여 계산을 실행하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
        ErrNum = 2
        GoTo Error_Message
    End If
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
'/ 2.급여대상기간 구함.
    If oJOBTRG = "%" Then
       sQry = " SELECT T0.Code   FROM [@PH_PY107B] T0"
       sQry = sQry & " WHERE T0.CODE = (SELECT TOP 1 CODE FROM [@PH_PY107A] WHERE CODE <= '" & oYM & "' ORDER BY CODE DESC)"
       sQry = sQry & " GROUP BY T0.Code, T0.U_STRMON, T0.U_STRDAY, T0.U_JIGMON, T0.U_JIGDAY"
       sRecordset.DoQuery sQry
       If sRecordset.RecordCount > 1 Then   '// 지급일자나 급여계산일자가 다를경우 전체로 할수없슴.
            ErrNum = 5
            GoTo Error_Message
       Else
            Set oCombo = oForm.Items("JOBTRG").Specific  '//
            If oCombo.ValidValues.Count > 1 Then
                oJOBTRG = Trim$(oCombo.ValidValues.Item(0).VALUE)
            End If
            Set oCombo = Nothing
       End If
    End If
    
    '/ 급여계산기간
    sQry = "SELECT DBO.Func_PAYTerm( '" & Format$(oYM, "0000-00") & "-01" & "', '" & oJOBTRG & "') "
    sRecordset.DoQuery sQry
    If sRecordset.RecordCount = 0 Then
        ErrNum = 3
        GoTo Error_Message
    Else
        If Trim$(sRecordset.Fields(0).VALUE) = "" Then
            ErrNum = 3
            GoTo Error_Message
        End If
        StrDate = Mid$(sRecordset.Fields(0).VALUE, 1, 8)   '/ 급여계산시작일
        EndDate = Mid$(sRecordset.Fields(0).VALUE, 10, 8)  '/ 급여계산종료일
    End If
    oJOBTRG = Trim$(oForm.Items("JOBTRG").Specific.Selected.VALUE)
    
    '/ 소득세 계산방법
    PAY_001 = MDC_SetMod.Get_ReData("U_CODGBN", "CODE", "[@ZPY304H]", "'PAY001'", "")
    '/ 상여금 평균임금 가져오는 방법
    PAY_007 = MDC_SetMod.Get_ReData("U_CODGBN", "CODE", "[@ZPY304H]", "'PAY007'", "")
'/ 급여계산 처리 루틴
    If Mid$(oYM, 5, 2) = "12" Then JSNYER = Val(Left$(oYM, 4)) Else JSNYER = Val(Left$(oYM, 4)) - 1
    oMSTCOD = oForm.Items("MSTCOD").Specific.String
'/ 급여입퇴사자범위아닌사람 제거
    sQry = "DELETE FROM  [@PH_PY111A]"
    sQry = sQry & " FROM [@PH_PY111A] T0  INNER JOIN [OHEM] T1 ON T0.U_MSTCOD = T1.U_MSTCOD"
    sQry = sQry & " INNER JOIN [@PH_PY001A] T2 ON T0.U_MSTCOD = T2.Code"
    sQry = sQry & " WHERE   T0.U_YM = '" & oYM & "'"
    sQry = sQry & " AND     T0.U_JOBTYP = '" & oJOBTYP & "'"
    sQry = sQry & " AND     T0.U_JOBGBN = '" & oJOBGBN & "'"
    sQry = sQry & " AND     (T2.U_JOBTRG LIKE '" & oJOBTRG & "' OR ISNULL(T2.U_JOBTRG, 0) = '0')"
    sQry = sQry & " AND     T0.U_ENDCHK <> 'Y'"  '잠금아닌것만
    sQry = sQry & " AND     (CONVERT(CHAR(8), T1.StartDate,112) > '" & EndDate & "'"
    sQry = sQry & "            OR (ISNULL(CONVERT(CHAR(8),T1.TermDate,112),'') <> ''"
    sQry = sQry & "                 AND CONVERT(CHAR(8), T1.TermDate, 112) < '" & StrDate & "'))"  '/ 당월퇴사자 제외
    If Trim$(oDS_PH_PY111A.GetValue("U_JOBTRG", 0)) <> "%" Then
        sQry = sQry & "   AND ISNULL(T2.U_JOBTRG, 0) = '" & oJOBTRG & "'"  '급여지급일
    End If
    If Trim$(oDS_PH_PY111A.GetValue("U_CLTCOD", 0)) <> "%" Then
        sQry = sQry & "   AND ISNULL(T0.U_CLTCOD, 0) = '" & oCLTCOD & "'"  '자사코드
    End If
    sQry = sQry & " AND     T0.U_MSTDPT BETWEEN '" & oSTRDPT & "' AND '" & oENDDPT & "'"
    If oMSTCOD <> "%" And oMSTCOD <> "" Then
        sQry = sQry & " AND     T0.U_MSTCOD LIKE '" & oMSTCOD & "'"  '사원번호
    End If
    sRecordset.DoQuery sQry
    
'/3. 인사마스터 조회-대상자 조회
    sQry = " SELECT  T0.U_MSTCOD, (T0.LastName + T0.FirstName) AS U_MSTNAM, T0.EmpID, T0.U_CLTCOD, T0.Branch AS U_MSTBRK, "
    sQry = sQry & " T3.U_MSTDPT, T4.U_MSTSTP,  ISNULL(T6.U_CodeNm,'') AS U_CLTNAM, T2.Name AS U_BRKNAM, T3.Name AS U_DPTNAM, T4.Name AS U_STPNAM, T1.U_PAYTYP,"
    sQry = sQry & " T1.U_JIGCOD, T1.U_HOBONG, T1.U_STDAMT, T1.U_BNSAMT, CONVERT(CHAR(8), T0.StartDate,112) AS U_INPDAT, "
    sQry = sQry & "  CONVERT(CHAR(8), T0.TermDate, 112) AS U_OUTDAT, ISNULL(T1.U_GBHSEL, 'N') AS U_GBHSEL, T1.U_KUKGRD, T1.U_MEDAMT, ISNULL(T1.U_KUKAMT,0) AS U_KUKAMT, ISNULL(T1.U_GBHAMT,0) AS U_GBHAMT, "
    sQry = sQry & " (1+ ISNULL(T1.U_BAEWOO,0)+ISNULL(T1.U_BUYNSU,0)) AS U_TAXCNT, ISNULL(T1.U_DAGYSU,0) AS U_CHLCNT,T1.U_TAXSEL, ISNULL(T1.U_FRGSEL, 'N') AS U_FRGSEL, ISNULL(T1.U_KUKOVR,'N') AS U_KUKOVR, "
    sQry = sQry & " ISNULL(T1.U_BX1SEL, 'N') AS U_BX1SEL , ISNULL(T1.U_BNSSEL, 'N') AS U_BNSSEL, T0.U_INPRAT, T0.U_GNTEXE, T0.U_INPGBN, T0.U_NOJGBN, ISNULL(T1.U_AVRAMT,0) AS U_AVRAMT,"
    sQry = sQry & " CONVERT(CHAR(8), T0.U_INEDAT, 112) AS U_INEDAT, CONVERT(CHAR(8), T0.U_GRPDAT, 112) AS U_GRPDAT, ISNULL(T1.U_NJCGBN,'N') AS U_NJCGBN,ISNULL(T1.U_MEDFRG, '0') AS U_MEDFRG, "
    sQry = sQry & " ISNULL(T5.U_CHAGAB, 0) AS U_CHAGAB, ISNULL(T5.U_CHANON,0) AS U_CHANON, ISNULL(T5.U_CHAJUM,0) AS U_CHAJUM, ISNULL(REPLACE(T0.GovID,'-',''),'') AS U_PERNBR, T1.U_JOBTRG "
    sQry = sQry & " FROM [OHEM] T0 INNER JOIN [@PH_PY001A] T1 ON T0.U_MSTCOD = T1.Code"
    sQry = sQry & "                    LEFT JOIN [OUBR] T2 ON T0.Branch = T2.Code"
    sQry = sQry & "                    LEFT JOIN [OUDP] T3 ON T0.Dept = T3.Code"
    sQry = sQry & "                    LEFT JOIN [OHPS] T4 ON T0.position = T4.posID"
    sQry = sQry & "                    LEFT JOIN [@ZPY504H] T5 ON T0.U_MSTCOD = T5.U_MSTCOD AND ISNULL(T5.U_JSNYER, '') = '" & Trim$(JSNYER) & "' AND ISNULL(T5.U_JSNGBN,'') = '1' AND T0.U_CLTCOD = T5.U_CLTCOD" '/ 연말정산
    sQry = sQry & "                    LEFT JOIN [@PS_HR200L] T6 ON T0.U_CLTCOD = T6.U_CODE AND T6.CODE = 'P144'"
    sQry = sQry & " WHERE CONVERT(CHAR(8), T0.StartDate,112) <= '" & EndDate & "'"
    sQry = sQry & "   AND (ISNULL(CONVERT(CHAR(8),T0.TermDate,112),'') ='" & Space$(0) & "'"
    sQry = sQry & "    OR CONVERT(CHAR(8), T0.TermDate, 112) >= '" & StrDate & "')"   '/ 당월퇴사자 제외
    sQry = sQry & "   AND ISNULL(T1.U_JOBTRG, 0) <> '0'"                    '/ 급여지급제외자"
    If oJOBTRG <> "%" Then
        sQry = sQry & "   AND ISNULL(T1.U_JOBTRG, 0) = '" & oJOBTRG & "'"  '급여지급일
    End If
    If Trim$(oDS_PH_PY111A.GetValue("U_CLTCOD", 0)) <> "%" Then
        sQry = sQry & "   AND ISNULL(T0.U_CLTCOD, 0) = '" & oCLTCOD & "'"  '자사코드
    End If
    sQry = sQry & "   AND T3.U_MSTDPT BETWEEN '" & oSTRDPT & "' AND '" & oENDDPT & "'"
    If oMSTCOD <> "%" And oMSTCOD <> "" Then
        sQry = sQry & "   AND ISNULL(T0.U_MSTCOD, '') = '" & oMSTCOD & "'"  '사원번호
    End If
        sQry = sQry & "   "  '정산년도
    sQry = sQry & " ORDER BY T0.U_CLTCOD, T0.Branch, T3.U_MSTDPT, T4.U_MSTSTP, T0.U_MSTCOD"
    sRecordset.DoQuery sQry
    If sRecordset.RecordCount = 0 Then
        ErrNum = 4
        GoTo Error_Message
    End If
'//
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    
    Set oProgBar = Sbo_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 100, False)
'// Process /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    sRecordset.MoveLast
    G_TotCnt = sRecordset.RecordCount  '/ 총급여처리대상인원수
    G_PayCnt = 0
    G_ChkCnt = 0
    oRow = 0
 '/ ProgressBar 메모리 할당
    TOTCNT = G_TotCnt
    V_StatusCnt = Round(TOTCNT / 100, 0)
    oProValue = 1
    sRecordset.MoveFirst
    Do Until sRecordset.EOF
       MSTCOD = sRecordset.Fields("U_MSTCOD").VALUE
       oRow = oRow + 1
       '/ 초기화
        Call PAY_Initial
       '/ 급여계산
        Call Basis_Setting

       If oJOBTYP <> "2" And oDS_PH_PY111A.GetValue("U_CHGCHK", 0) = "N" Then
'            Select Case Trim$(MDC_COMpanyGubun)
'            'Case "CL", "HS", "SO"
'            Case "HS"
'                Call PayRoll_Process1
'            Case Else    '/ 로직사용
                Result = PayRoll_Process
                If Trim$(Result) <> "" Then
                    oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & sRecordset.Fields("U_MSTNAM").VALUE & Space$(1) & Result
'                    oMat1.AddRow
'                    ErrNum = 9
'                    GoTo error_Message
                End If
'            End Select
       End If
       '/ 상여계산
       If oJOBTYP <> "1" And sRecordset.Fields("U_BNSSEL").VALUE = "Y" And oDS_PH_PY111A.GetValue("U_CHGCHK", 0) = "N" Then   '/ 상여지급여부
            Result = BonusRoll_Process
            If Trim$(Result) <> "" Then
                    oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & sRecordset.Fields("U_MSTNAM").VALUE & Space$(1) & Result
'                    oMat1.AddRow
'                ErrNum = 8
'                GoTo error_Message
            End If
       End If
       '/ 변동수당/공제항목 추가
       Call ChangeItem_Process
       '/ 세액계산
'       Select Case Trim$(MDC_COMpanyGubun)
'       Case "CL", "HS", "SO"
'            Call PAY_Tax1
'       Case Else
            Call PAY_Tax
'       End Select
       '/ 연말정산징수환급해주세용
       If Trim(oDS_PH_PY111A.GetValue("U_JSNCHK", 0)) = "Y" Then
            Call JeongSan_Process
       End If
       '/ 급상여자료 저장
       If WG03.U_TOTPAY <> 0 Or WG03.U_TOTGON <> 0 Then
            oForm.DataSources.UserDataSources("Col0").VALUE = oRow
            Result = PH_PY111_Save
            If Trim$(Result) = "Y" Then
                oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & sRecordset.Fields("U_MSTNAM").VALUE & " 급여계산 완료."
            Else
                oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & sRecordset.Fields("U_MSTNAM").VALUE & Space$(1) & Result
            End If
       Else
            oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & sRecordset.Fields("U_MSTNAM").VALUE & " 급여총지급액이 0입니다."
       End If
'        oMat1.AddRow
        
        '// 상태보여주기
        If (TOTCNT > 100 And oRow = oProValue * V_StatusCnt) Or TOTCNT <= 100 Then
            oProgBar.Text = oRow & "/ " & TOTCNT & " 건 처리중...!"
            oProValue = oProValue + 1
            oProgBar.VALUE = oProValue
        End If
        
        
       sRecordset.MoveNext
    Loop
 '/ END
    oProgBar.Stop
    sQry = "총: " & G_TotCnt & "( 처리 :" & G_PayCnt & " 잠김 :" & G_ChkCnt & ")건이 처리되었습니다."
    If Trim$(REMARK1) <> "" Then sQry = sQry & vbCrLf & "고용보험 나이제외:" & REMARK1
    If Trim$(REMARK2) <> "" Then sQry = sQry & vbCrLf & "국민연금 나이제외:" & REMARK2
    If Trim$(REMARK3) <> "" Then sQry = sQry & vbCrLf & "실지급액이 마이너스 사원:" & REMARK3
    If TermCHK = True Then sQry = sQry & vbCrLf & "상여세액기간에 급여자료가 없습니다.세액이 적게 나올 수 있습니다."
    
    Sbo_Application.MessageBox (sQry)
    Sbo_Application.StatusBar.SetText "총:" & G_TotCnt & "( 처리 :" & G_PayCnt & " 잠김 :" & G_ChkCnt & ")건이 처리되었습니다.", bmt_Short, smt_Success
 '/ ProgressBar 중지/메모리 반환
    Set sRecordset = Nothing
    Set oProgBar = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set sRecordset = Nothing
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText oStrE, bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "급여계산작업이 취소되었습니다.", bmt_Short, smt_Success
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "(ZPY110)급여대상기준일을 먼저 등록하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "지급 대상자가 없습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "지급일 또는 급여기간이 다를 경우 지급 대상자 구분을 모두로 선택할수 없습니다. 확인해주세요", bmt_Short, smt_Success
    ElseIf ErrNum = 8 Then
        Sbo_Application.StatusBar.SetText "Bonus Error: 사번:" & MSTCOD & Result, bmt_Short, smt_Error
    ElseIf ErrNum = 9 Then
        Sbo_Application.StatusBar.SetText "PayRoll Error: 사번:" & MSTCOD & Result, bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Execution_Process Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub

Private Sub PAY_Initial()
     Dim iCol    As Integer
    Dim GNSGBN  As String
    
    GNSGBN = ""
    WG03.DocNum = 0      '문서번호
    WG03.U_BASAMT = 0
    WG03.U_DAYAMT = 0
    WG03.U_MSTCOD = Trim$(sRecordset.Fields("U_MSTCOD").VALUE)
    WG03.U_MSTNAM = Trim$(sRecordset.Fields("U_MSTNAM").VALUE)
    WG03.U_EmpID = sRecordset.Fields("EmpID").VALUE
    WG03.U_CLTCOD = sRecordset.Fields("U_CLTCOD").VALUE
    WG03.U_MSTBRK = sRecordset.Fields("U_MSTBRK").VALUE
    WG03.U_MSTDPT = sRecordset.Fields("U_MSTDPT").VALUE
    WG03.U_MSTSTP = sRecordset.Fields("U_MSTSTP").VALUE
    WG03.U_CLTNAM = sRecordset.Fields("U_CLTNAM").VALUE
    WG03.U_BRKNAM = sRecordset.Fields("U_BRKNAM").VALUE
    WG03.U_DPTNAM = sRecordset.Fields("U_DPTNAM").VALUE
    WG03.U_STPNAM = sRecordset.Fields("U_STPNAM").VALUE
    WG03.U_PAYTYP = MDC_SetMod.Get_ReData("U_RelCd", "U_Minor", "[@PS_HR200L]", "'" & Trim$(sRecordset.Fields("U_PAYTYP").VALUE) & "'", " AND Code='P132'")
    WG03.U_JOBTRG = sRecordset.Fields("U_JOBTRG").VALUE
    WG03.U_PERNBR = sRecordset.Fields("U_PERNBR").VALUE
    WG03.U_JIGCOD = sRecordset.Fields("U_JIGCOD").VALUE
    WG03.U_HOBONG = sRecordset.Fields("U_HOBONG").VALUE
    WG03.U_STDAMT = sRecordset.Fields("U_STDAMT").VALUE
    WG03.U_MEDAMT = sRecordset.Fields("U_MEDAMT").VALUE
    WG03.U_KUKAMT = sRecordset.Fields("U_KUKAMT").VALUE
    WG03.U_GBHAMT = sRecordset.Fields("U_GBHAMT").VALUE
    
    '/ 입사일자(근속일기준: 1-그룹입사일기준, 2-입사일기준)
    GNSGBN = MDC_SetMod.Get_ReData("U_GNSGBN", "U_YM", "[@PH_PY106A]", "'" & Trim$(U_CSUCOD) & "'", " AND U_PAYTYP='" & Trim$(WG03.U_PAYTYP) & "'")
    If Trim$(GNSGBN) = "1" Then '/ 그룹입사일기준
        WG03.U_INPDAT = sRecordset.Fields("U_GRPDAT").VALUE
    Else
        WG03.U_INPDAT = sRecordset.Fields("U_INPDAT").VALUE
    End If
    WG03.U_INEDAT = sRecordset.Fields("U_INEDAT").VALUE
    WG03.U_OUTDAT = sRecordset.Fields("U_OUTDAT").VALUE
    WG03.U_TAXCNT = sRecordset.Fields("U_TAXCNT").VALUE
    WG03.U_CHLCNT = sRecordset.Fields("U_CHLCNT").VALUE
    WG03.U_NJCGBN = sRecordset.Fields("U_NJCGBN").VALUE
    WG03.U_MEDFRG = Val(sRecordset.Fields("U_MEDFRG").VALUE)
    '// 수당항목
    For iCol = 1 To 24
        WG03.U_CSUAMT(iCol) = 0
    Next iCol
    WG03.U_GWASEE = 0
    WG03.U_BTAX01 = 0
    WG03.U_BTAX02 = 0
    WG03.U_BTAX03 = 0
    WG03.U_BTAX04 = 0
    WG03.U_BTAX05 = 0
    WG03.U_BTAX06 = 0
    WG03.U_BTAX07 = 0
    
    WG03.U_BTXG01 = 0
    WG03.U_BTXH01 = 0
    WG03.U_BTXH05 = 0
    WG03.U_BTXH06 = 0
    WG03.U_BTXH07 = 0
    WG03.U_BTXH08 = 0
    WG03.U_BTXH09 = 0
    WG03.U_BTXH10 = 0
    WG03.U_BTXH11 = 0
    WG03.U_BTXH12 = 0
    WG03.U_BTXH13 = 0
    WG03.U_BTXI01 = 0
    WG03.U_BTXK01 = 0
    WG03.U_BTXM01 = 0
    WG03.U_BTXM02 = 0
    WG03.U_BTXM03 = 0
    WG03.U_BTXO01 = 0
    WG03.U_BTXQ01 = 0
    WG03.U_BTXR10 = 0
    WG03.U_BTXS01 = 0
    WG03.U_BTXT01 = 0
    WG03.U_BTXX01 = 0
    WG03.U_BTXY01 = 0
    WG03.U_BTXY02 = 0
    WG03.U_BTXY03 = 0
    WG03.U_BTXY20 = 0
    WG03.U_BTXY21 = 0
    WG03.U_BTXY22 = 0
    WG03.U_BTXZ01 = 0
    WG03.U_BTXTOT = 0
    
    WG03.U_TOTPAY = 0
    '// 공제항목
    For iCol = 1 To 18
        WG03.U_GONAMT(iCol) = 0
    Next iCol

    WG03.U_TOTGON = 0
    WG03.U_SILJIG = 0
    '// 상여금
    WG03.U_AVRPAY = 0
    WG03.U_NABTAX = 0
    WG03.U_BNSRAT = 0
    WG03.U_APPRAT = 0
    WG03.U_GNSYER = 0
    WG03.U_GNSMON = 0
    WG03.U_TAXTRM = 0
    WG03.U_BONUSS = 0
    '// 2011.10.05 상여금 평균임금 급여기본등록에서 가져오기
    WK_AVRAMT = sRecordset.Fields("U_AVRAMT").VALUE
    '// 2009년추가 비과세코드가져오기
    TB1_BT3COD = MDC_SetMod.Get_ReData("MAX(U_BTXCOD)", "U_GWATYP", "[@ZPY107L]", "'5'", " AND Left(Code,4) <='" & Left$(oYM, 4) & "' GROUP BY CODE ORDER BY CODE DESC")
    If Trim$(TB1_BT3COD) = "" Then TB1_BT3COD = "M01"
    
    TB1_BT5COD = MDC_SetMod.Get_ReData("MAX(U_BTXCOD)", "U_GWATYP", "[@ZPY107L]", "'6'", " AND Left(Code,4) <='" & Left$(oYM, 4) & "' GROUP BY CODE ORDER BY CODE DESC")
    If Trim$(TB1_BT5COD) = "" Then TB1_BT5COD = "H10"
    
End Sub

Private Function PayRoll_Process() As String
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim pRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
    Dim iCol            As Integer, kCol As Long
    Dim SuSilStr        As String
    Dim Tmp_CSUAMT     As Double
    
    ErrNum = 0
'// 각종수당들은 해당이름으로 변수만들어서 저장, 계산식 만드는것은 나름 저장그래야 루프문안씀
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    Set pRecordset = oCompany.GetBusinessObject(BoRecordset)

'// 1.1. 공식가져오기-수당
    sQry = " SELECT T0.U_LINSEQ, T0.U_CSUCOD, T0.U_SILCUN, T0.U_SILCOD"
    sQry = sQry & " FROM [@PH_PY106B] T0 INNER JOIN [@PH_PY106A] T1 ON T0.DocEntry = T1.DocEntry"
    sQry = sQry & " WHERE   T1.U_YM = '" & Trim$(U_CSUCOD) & "'"
    sQry = sQry & " AND     T1.U_PAYTYP = '" & Trim$(WG03.U_PAYTYP) & "'"
    sQry = sQry & " ORDER BY T0.U_LINSEQ, T0.U_CSUCOD"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    Do Until oRecordSet.EOF
        iCol = Val(oRecordSet.Fields("U_LINSEQ").VALUE)
        If iCol > 0 And Trim$(oRecordSet.Fields("U_CSUCOD").VALUE) <> "" Then
            WK_C.GONSIL(iCol) = oRecordSet.Fields("U_SILCUN").VALUE  '/ 급여계산식
        Else
            If Left$(oRecordSet.Fields("U_CSUCOD").VALUE, 1) = "X" Then
                SuSilStr = oRecordSet.Fields("U_SILCUN").VALUE   '/ 급여계산식
                '/ 2.2. 공식안 시스템제공코드값으로 변경
                SuSilStr = Change_GOSIL(SuSilStr)
                Select Case Trim$(oRecordSet.Fields("U_CSUCOD").VALUE)
                Case "X01"       '/ 기본일급
                     X01_Val = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                Case "X02"       '/ 통상일급
                     X02_Val = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                Case "X03"       '/ 기본시급
                     X03_Val = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                Case "X04"       '/ 통상시급
                     X04_Val = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                End Select
            End If
        End If
        oRecordSet.MoveNext
    Loop
    WG03.U_DAYAMT = X01_Val
    WG03.U_BASAMT = X02_Val

'// 2. 공식가지고 계산하기-수당
    For iCol = 1 To 24
        If Trim$(WK_C.CSUCOD(iCol)) <> "" Then '/ 수당항목이 있는것만
            If Trim$(WK_C.GONSIL(iCol)) <> "" Then  '/ 공식이있으면 공식계산
               If oJOBTYP = "1" Or oJOBTYP = "3" Then    '/ 상여만아니면
                '/ 2.1. 공식-계산결과값가져오는거면..
                    SuSilStr = WK_C.GONSIL(iCol)   '/ 계산된값일 경우
                    For kCol = 1 To 24
                        SuSilStr = Replace(SuSilStr, "#" & Format$(kCol, "00"), WG03.U_CSUAMT(kCol))
                    Next kCol
                '/ 2.2. 공식안 시스템제공코드값으로 변경
                    SuSilStr = Change_GOSIL(SuSilStr)
                '/ 2.3. 공식계산하기
                   Tmp_CSUAMT = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                '/ 2.4. 정답가져오기
                    WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(Tmp_CSUAMT, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                    If Trim$(WK_C.CSUCOD(iCol)) = "A04" And (oJOBTYP = "1" Or sRecordset.Fields("U_BNSSEL").VALUE <> "Y") Then
                        WG03.U_CSUAMT(iCol) = 0
                    End If
                End If
            End If
        End If
    Next iCol

    Set oRecordSet = Nothing
    Set pRecordset = Nothing
    PayRoll_Process = ""
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Set pRecordset = Nothing
    If ErrNum = 1 Then
        PayRoll_Process = "수당계산식이 없습니다. 확인하여 주십시오. (적용연월:" & U_CSUCOD & "급여형태: " & WG03.U_PAYTYP & ")"
    Else
        PayRoll_Process = Err.Number & Space$(10) & Err.Description
    End If
End Function

Private Sub Basis_Setting()
    Dim STRDAT          As String
    Dim ENDDAT          As String
'/
    X01_Val = 0: X02_Val = 0: X03_Val = 0: X04_Val = 0
    X10_Val = 0: X12_Val = 0: X13_Val = 0: X14_Val = ""
    X15_Val = 0: X16_Val = 0: X17_Val = 0: X18_Val = 0
    X19_Val = 0: X20_Val = 0
'// X01,X02,X03,X04 <---수식초기에 셋팅


'// X10:종료일수(해당작업기간의 월총일수 가져오기)
    X10_Val = DateDiff("d", Format$(StrDate, "0000-00-00"), Format$(EndDate, "0000-00-00")) + 1 '/ 한샘, 케어라인
    

'// X11, X12, X13: 근속년수/근속월수/근속일수 가져오기
    If Trim$(WG03.U_OUTDAT) <> "" And WG03.U_OUTDAT <= EndDate Then
        Call MDC_SetMod.Term2(WG03.U_INPDAT, WG03.U_OUTDAT)
    Else
        Call MDC_SetMod.Term2(WG03.U_INPDAT, EndDate)
    End If
    X11_Val = ZPAY_GBL_GNSYER
    X12_Val = ZPAY_GBL_GNSMON
    X13_Val = ZPAY_GBL_GNSDAY

'// X14: 당월입퇴사유무
    If (WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate) Or _
       (WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate) Then
        X14_Val = "1"
    Else
        X14_Val = "0"
    End If

'// X15:총지급액 WG03.TOTPAY

'// X16: 입퇴사일수
    X16_Val = X10_Val
    If (WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate) Or _
       (WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate) Then
        Select Case True
         Case (WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate) And _
              (WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate)  '/ 당월입퇴사자
                Call MDC_SetMod.Term2(WG03.U_INPDAT, WG03.U_OUTDAT)
         Case WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate    '/ 당월입사자
              Call MDC_SetMod.Term2(WG03.U_INPDAT, EndDate)
         Case WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate    '/ 당월퇴사자
              Call MDC_SetMod.Term2(StrDate, WG03.U_OUTDAT)
        End Select
        X16_Val = (ZPAY_GBL_GNSMON * X10_Val + ZPAY_GBL_GNSDAY)
    End If
    
'// X18:수습일 (당월수습기간에 해당하는지 여부)
    X18_Val = 0
    If (WG03.U_INEDAT >= StrDate Or (WG03.U_INEDAT <= EndDate And StrDate < WG03.U_INEDAT)) Then
        If StrDate <= WG03.U_INPDAT Then STRDAT = WG03.U_INPDAT Else STRDAT = StrDate
        If EndDate <= WG03.U_INEDAT Then ENDDAT = EndDate Else ENDDAT = WG03.U_INEDAT
        If Trim$(WG03.U_OUTDAT) <> "" Then
            If WG03.U_OUTDAT <= EndDate And WG03.U_OUTDAT <= WG03.U_INEDAT Then
                ENDDAT = WG03.U_OUTDAT
            End If
        End If
        Call MDC_SetMod.Term2(STRDAT, ENDDAT)
        X18_Val = (ZPAY_GBL_GNSMON * X10_Val + ZPAY_GBL_GNSDAY)
    End If

'// X17:수습율
    X17_Val = Val(sRecordset.Fields("U_INPRAT").VALUE) / 100   '/ 수습율
    If Val(X17_Val) = 0 Then X17_Val = 1   '/수습율입력안하면 기본급 100%

'// X19:근속일수(급여종료일기준)
    If Trim$(WG03.U_OUTDAT) <> "" And WG03.U_OUTDAT <= EndDate Then
        X19_Val = DateDiff("d", Format$(WG03.U_INPDAT, "0000-00-00"), Format$(WG03.U_OUTDAT, "0000-00-00")) + 1
    Else
        X19_Val = DateDiff("d", Format$(WG03.U_INPDAT, "0000-00-00"), Format$(EndDate, "0000-00-00")) + 1
    End If

'// X20:근속일수(상여종료일기준) --2009.10.22 수정함 퇴사일반영.
    If Trim$(WG03.U_OUTDAT) <> "" And WG03.U_OUTDAT <= oGNEDAT Then
        X20_Val = DateDiff("d", Format$(WG03.U_INPDAT, "0000-00-00"), Format$(WG03.U_OUTDAT, "0000-00-00")) + 1
    Else
        X20_Val = DateDiff("d", Format$(WG03.U_INPDAT, "0000-00-00"), Format$(oGNEDAT, "0000-00-00")) + 1
    End If
End Sub

Private Function Change_GOSIL(xSuSil$) As String
  '/ 공식안에 시스템제공코드있으면 해당 값으로 변경
    xSuSil = Replace(xSuSil, "X01", X01_Val)  '/ 기본일급
    xSuSil = Replace(xSuSil, "X02", X02_Val)  '/ 통상일급
    xSuSil = Replace(xSuSil, "X03", X03_Val)  '/ 기본시급
    xSuSil = Replace(xSuSil, "X04", X04_Val)  '/ 통상시급
    xSuSil = Replace(xSuSil, "X10", X10_Val)      '/ 월총일수
    xSuSil = Replace(xSuSil, "X11", X11_Val)      '/ 근속년수
    xSuSil = Replace(xSuSil, "X12", X12_Val)      '/ 근속월수
    xSuSil = Replace(xSuSil, "X13", X13_Val)      '/ 근속일수
    xSuSil = Replace(xSuSil, "X14", X14_Val)    '/ 당월입퇴사자유무
    xSuSil = Replace(xSuSil, "X16", X16_Val)      '/ 당월입퇴사근무기준일
    xSuSil = Replace(xSuSil, "X17", X17_Val)      '/ 수습율
    xSuSil = Replace(xSuSil, "X18", X18_Val)      '/ 수습일수
    xSuSil = Replace(xSuSil, "X19", X19_Val)      '/ 근속일수(급여종료일기준)
    xSuSil = Replace(xSuSil, "X20", X20_Val)      '/ 근속일수(상여종료일기준)
    xSuSil = Replace(xSuSil, "X21", "''" & oYM & "''")      '/ 귀속연월
    xSuSil = Replace(xSuSil, "X22", "''" & oJOBTYP & "''")      '/ 지급종류
    xSuSil = Replace(xSuSil, "X23", "''" & oJOBGBN & "''")      '/ 지급구분
    
    Change_GOSIL = xSuSil
End Function

Private Sub PayRoll_Process1()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim pRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim iCol            As Integer
    Dim INPRAT          As Double
    Dim ENDDAY          As Double
    Dim INPCheck        As String
    Dim Imsi_DAYAMT     As Double
    Dim Imsi_TIMAMT     As Double
'// 각종수당들은 해당이름으로 변수만들어서 저장, 계산식 만드는것은 나름 저장그래야 루프문안씀
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    Set pRecordset = oCompany.GetBusinessObject(BoRecordset)
'/2. 기본급여사항가져오기.
    '/ 고정수당
    sQry = " SELECT U_FILD01, U_FILD02, U_FILD03 FROM [@ZPY127L] WHERE CODE  ='" & Trim$(WG03.U_MSTCOD) & "' "
    sQry = sQry & "  AND ISNULL(U_FILD01, '') <> '' " 'AND ISNULL(U_FILD03, 0) <> 0"
    sQry = sQry & " ORDER BY CODE, LineID"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        For iCol = 1 To 24
            If oRecordSet.Fields("U_FILD01").VALUE = WK_C.CSUCOD(iCol) Then
                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(oRecordSet.Fields("U_FILD03").VALUE, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
            ElseIf WK_C.CSUCOD(iCol) = "A01" Then
                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(sRecordset.Fields("U_STDAMT").VALUE, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
            End If
        Next iCol
        oRecordSet.MoveNext
    Loop
'    '/ 고정공제
'    sQry = " SELECT U_FILD01, U_FILD02, U_FILD03 FROM [@ZPY127M] WHERE CODE  ='" & Trim$(WG03.U_MSTCOD) & "' "
'    sQry = sQry & "  AND ISNULL(U_FILD01, '') <> '' AND ISNULL(U_FILD03, 0) <> 0"
'    sQry = sQry & " ORDER BY CODE, LineID"
'    oRecordSet.DoQuery sQry
'    Do Until oRecordSet.EOF
'        For iCol = 1 To 18
'            If oRecordSet.Fields("U_FILD01").Value = WK_G.GONCOD(iCol) Then
'                WG03.U_GONAMT(iCol) = oRecordSet.Fields("U_FILD03").Value
'            End If
'        Next iCol
'        oRecordSet.MoveNext
'    Loop
    '/ 일할계산 방법 구함(말일기준, 30일기준)
    sQry = "SELECT TOP 1 U_INPDAY FROM [@PH_PY107A] WHERE CODE <= '" & oYM & "' ORDER BY CODE DESC"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        INPCheck = oRecordSet.Fields(0).VALUE
        If INPCheck = 2 Then
            ENDDAY = DateDiff("d", Format$(StrDate, "0000-00-00"), Format$(EndDate, "0000-00-00")) + 1 '/ 한샘, 케어라인
        Else
            ENDDAY = 30
        End If
    Else
        ENDDAY = 30
    End If

'/ 2.2) 수습사원 급여 수습율로 구함.
    If Val(sRecordset.Fields("U_INPRAT").VALUE) > 0 Then
        INPRAT = Val(sRecordset.Fields("U_INPRAT").VALUE) / 100   '/ 율
        For iCol = 1 To 24
            WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) * INPRAT, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        Next iCol
    End If
'/ 근태자료관련 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    Dim WG23_CHLDAY%, WG23_HGNDAY%, WG23_GULDAY%, WG23_WCHDAY%, WG23_WCHHGA%
    Dim WG23_JCHDAY%, WG23_YCHHGA%, WG23_SNHDAY%, WG23_SNHHGA%
    Dim WG23_JIGCNT%, WG23_JOTCNT%
    Dim WG23_GNTTIM#, WG23_UGBTIM#, WG23_JUPTIM#, WG23_NHTTIM#, WG23_HGNTIM#, WG23_HYUJUP#, WG23_HYUNHT#
    Dim WG23_JIGTIM#, WG23_JOTTIM#, WG23_OICTIM#
    
    sQry = " SELECT T0.*"
    sQry = sQry & " FROM [@ZPY230L] T0 INNER JOIN [@ZPY230H] T1 ON T0.DocEntry = T1.DocEntry"
    sQry = sQry & " WHERE   T1.U_GNTYMM = '" & oYM & "'"
    sQry = sQry & " AND     T0.U_MSTCOD = '" & Trim$(WG03.U_MSTCOD) & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        WG23_CHLDAY = 0: WG23_HGNDAY = 0: WG23_GULDAY = 0
        WG23_WCHDAY = 0: WG23_WCHHGA = 0: WG23_JCHDAY = 0
        WG23_YCHHGA = 0: WG23_SNHDAY = 0: WG23_SNHHGA = 0
        WG23_JIGCNT = 0: WG23_JOTCNT = 0: WG23_GNTTIM = 0
        WG23_UGBTIM = 0: WG23_JUPTIM = 0: WG23_NHTTIM = 0
        WG23_HGNTIM = 0: WG23_HYUJUP = 0: WG23_HYUNHT = 0
        WG23_JIGTIM = 0: WG23_JOTTIM = 0: WG23_OICTIM = 0
    Else
        WG23_CHLDAY = oRecordSet.Fields("U_CHLDAY").VALUE
        WG23_HGNDAY = oRecordSet.Fields("U_HGNDAY").VALUE
        WG23_GULDAY = oRecordSet.Fields("U_GULDAY").VALUE
        WG23_WCHDAY = oRecordSet.Fields("U_WCHDAY").VALUE
        WG23_WCHHGA = oRecordSet.Fields("U_WCHHGA").VALUE
        WG23_JCHDAY = oRecordSet.Fields("U_JCHDAY").VALUE
        WG23_YCHHGA = oRecordSet.Fields("U_YCHHGA").VALUE
        WG23_SNHDAY = oRecordSet.Fields("U_SNHDAY").VALUE
        WG23_SNHHGA = oRecordSet.Fields("U_SNHHGA").VALUE
        WG23_JIGCNT = oRecordSet.Fields("U_JIGCNT").VALUE
        WG23_JOTCNT = oRecordSet.Fields("U_JOTCNT").VALUE
        WG23_GNTTIM = oRecordSet.Fields("U_GNTTIM").VALUE
        WG23_UGBTIM = oRecordSet.Fields("U_UGBTIM").VALUE
        WG23_JUPTIM = oRecordSet.Fields("U_JUPTIM").VALUE
        WG23_NHTTIM = oRecordSet.Fields("U_NHTTIM").VALUE
        WG23_HGNTIM = oRecordSet.Fields("U_HGNTIM").VALUE
        WG23_HYUJUP = oRecordSet.Fields("U_HYUJUP").VALUE
        WG23_HYUNHT = oRecordSet.Fields("U_HYUNHT").VALUE
        WG23_JIGTIM = oRecordSet.Fields("U_JIGTIM").VALUE
        WG23_JOTTIM = oRecordSet.Fields("U_JOTTIM").VALUE
        WG23_OICTIM = oRecordSet.Fields("U_OICTIM").VALUE
        
   End If
'/ 2.3)연봉직,월급직의 당월입퇴사자 일할계산
    '// 보통 입퇴사일기준으로 기본급, 고정수당 일할계산함.
     If Trim$(WG03.U_PAYTYP) = "1" Or Trim$(WG03.U_PAYTYP) = "2" Then
'        Select Case MDC_COMpanyGubun
        'Case "CL" '// 케어라인은 관리직도 입퇴사자 입퇴사일기준으로 일할계산하지 않고 월근태를 사용함.
'        Case Else
                 Call MDC_SetMod.Term2(StrDate, EndDate)
                 If (WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate) Or _
                    (WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate) Then
                     Select Case True
                      Case (WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate) And (WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate)
                       '/ 당월입퇴사자 기본급 /
                         Call MDC_SetMod.Term2(WG03.U_INPDAT, WG03.U_OUTDAT)
                      Case WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate
                      '/ 당월입사자 기본급 /
                           Call MDC_SetMod.Term2(WG03.U_INPDAT, EndDate)
                      Case WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate
                      '/ 당월퇴사자 기본급 /
                           Call MDC_SetMod.Term2(StrDate, WG03.U_OUTDAT)
                     End Select
                     For iCol = 1 To 24
                         WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * (ZPAY_GBL_GNSMON * ENDDAY + ZPAY_GBL_GNSDAY), WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                     Next iCol
                 End If
'        End Select
    End If
'/ 3) 일급직, 시급직 근태 일할 계산
        Select Case Trim$(WG03.U_PAYTYP)
        Case "1"        '/ 1.연봉직
            WK_DAYAMT = sRecordset.Fields("U_STDAMT").VALUE / 365
            WK_TIMAMT = WK_DAYAMT / 8
            WG03.U_STDAMT = MDC_SetMod.IInt(WK_TIMAMT, 1)
        Case "2"        '/ 2.월급직
            WK_DAYAMT = sRecordset.Fields("U_STDAMT").VALUE / 30
            WK_TIMAMT = WK_DAYAMT / 8                  '한샘
            WG03.U_STDAMT = MDC_SetMod.IInt(WK_TIMAMT, 1)
        Case "3"        '/ 3.일급직
            WK_DAYAMT = sRecordset.Fields("U_STDAMT").VALUE
            WK_TIMAMT = WK_DAYAMT / 8
            WG03.U_STDAMT = WK_DAYAMT
        Case "4"        '/ 4.시급직
            WK_TIMAMT = sRecordset.Fields("U_STDAMT").VALUE
            WK_DAYAMT = WK_TIMAMT * 8
            WG03.U_STDAMT = WK_TIMAMT
        End Select
        Imsi_DAYAMT = 0
        Imsi_TIMAMT = 0
        WG03.U_DAYAMT = WK_DAYAMT

'   '/ 근태관리
    '/ 근태계산대상인 체크인자만.
        If Trim$(sRecordset.Fields("U_GNTEXE").VALUE) = "Y" Then
            For iCol = 1 To 24
                '/ 생산직만
                If Trim$(sRecordset.Fields("U_BX1SEL").VALUE) = "Y" Then
                    If Trim$(WK_C.CSUCOD(iCol)) = "A01" And (Trim$(WG03.U_PAYTYP) = "3" Or Trim$(WG03.U_PAYTYP) = "4") Then
                        WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt((WG23_GNTTIM) * WK_TIMAMT, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                    End If
'                    If MDC_COMpanyGubun = "CL" Then
'                        '/ (2007-05-15케어라인)생산직만-장려수당 15일이상 근무시 100%지급, 15일이하이면 지급제외.
'                        If Trim$(WK_C.CSUCOD(iCol)) = "E02" And WG23_CHLDAY < 15 Then
'                            WG03.U_CSUAMT(iCol) = 0
'                        End If
'                        '/ 케어라인: 일할계산(관리직,생산직모두)(결근,병가있으면 자가계발비, 식대보조,교통비 일할계산)
'                        If WG23_GULDAY > 0 Or ((WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate) _
'                                                 Or (WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate)) Then
'                            If Trim$(WK_C.CSUCOD(iCol)) = "E03" Then '/ 자가계발비
'                                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
'                            ElseIf Trim$(WK_C.CSUCOD(iCol)) = "B01" Then '/ 식대보조
'                                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
'                            ElseIf Trim$(WK_C.CSUCOD(iCol)) = "B02" Then '/ 교통비
'                                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
'                            End If
'                        End If
'                    ElseIf MDC_COMpanyGubun = "SO" Then
'                        '/ 씬터온: 생산장려 결근이 있으면 생산장려지급안함.
'                        If WG23_GULDAY > 0 Or ((WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate) _
'                                                 Or (WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate)) Then
'                            If Trim$(WK_C.CSUCOD(iCol)) = "E02" Then '/ 생산장려
'                                WG03.U_CSUAMT(iCol) = 0
'                            End If
'                        End If
'
'                    End If
                Else
'                    If MDC_COMpanyGubun = "CL" Then
'                        If Trim$(WG03.U_PAYTYP) = "1" Or Trim$(WG03.U_PAYTYP) = "2" Then
'                            Select Case Trim$(WK_C.CSUCOD(iCol))
'                            Case "E01" '/ (2007-05-15케어라인)관리직만-정근수당 20일이상 100%지급 20일 이하이면 일할계산
'                                 If Trim$(WK_C.CSUCOD(iCol)) = "E01" And WG23_CHLDAY < 20 Then
'                                    WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
'                                 End If
'                            Case Else
'                                 WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
'                            End Select
'                        End If
'                    End If
                End If
            '// 근태관련 수당들 계산
                Select Case Trim$(WK_C.CSUCOD(iCol))
                Case "D02" '/ 케어라인 : 연장수당, 한샘 : 야간수당, 씬터온:잔업수당
                    WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WK_TIMAMT * WG23_JUPTIM * 1.5, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                Case "D04" '/ 특근수당
                    WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WK_TIMAMT * WG23_HGNTIM * 1.5, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                Case "D06"  '/ 특근연장
                    WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WK_TIMAMT * WG23_HYUJUP * 2, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                Case "D03" '/ 야근수당, 한샘: 심야수당, 씬터온:야간수당
                    '잔업수당(관리직은고정급, 생산직은 시급계산)
                    WG03.U_CSUAMT(iCol) = MDC_SetMod.IInt(WK_TIMAMT * WG23_NHTTIM * 2, 1)   '/ 한샘
'                   WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(Imsi_TIMAMT * WG23_NHTTIM * 0.5, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                Case "C05" '/ 주휴수당 (씬터온주휴있음)
                        If (Trim$(WG03.U_PAYTYP) = "3" Or Trim$(WG03.U_PAYTYP) = "4") Then     '/ 한샘 반장(월급직)은 주휴수당안나감
                           WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WK_DAYAMT * WG23_JCHDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                        End If
                Case "D01" '/ 일신: 유급수당, 기본연장  한샘 : 잔업수당 =연장시간
'                        Select Case MDC_COMpanyGubun
'                        Case "CL"
'                                '/ 추가연장: 5.30 별정직 기본연장이 고정급으로 이미들어가있으면 월근태에서 연장계산 제외
'                                If WG03.U_CSUAMT(iCol) = 0 Then
'                                    WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WK_TIMAMT * WG23_UGBTIM * 1.25, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
'                                End If
'                        Case "HS"
'                                If Trim$(sRecordset.Fields("U_BX1SEL").Value) = "Y" Then  '/ 생산직만계산
'                                   WG03.U_CSUAMT(iCol) = MDC_SetMod.IInt(WK_TIMAMT * WG23_JUPTIM * 1.5, 1)           '/ 한샘
'                                End If
'                        End Select
                Case "C02" '/ 월차수당(일신)씬터온월차없음
                    If (WG23_WCHDAY - WG23_WCHHGA) > 0 Then
'                        If MDC_COMpanyGubun = "IW" Then
'                            WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(sRecordset.Fields("U_STDAMT").Value / 30 * (WG23_WCHDAY - WG23_WCHHGA), WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
'                        Else
                            WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WK_DAYAMT * (WG23_WCHDAY - WG23_WCHHGA), WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
'                        End If
                    End If
                Case "C03" '/ 보건수당(일신)씬터온보건없음
                        If (WG23_SNHDAY - WG23_SNHHGA) > 0 Then
'                            If MDC_COMpanyGubun = "IW" Then
'                                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(sRecordset.Fields("U_STDAMT").Value / 30, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol)) * (WG23_SNHDAY - WG23_SNHHGA)
'                            Else
                                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(sRecordset.Fields("U_STDAMT").VALUE / 30, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol)) * (WG23_SNHDAY - WG23_SNHHGA)
'                            End If
                        End If
                Case "A01"
'                    '// 케어라인 (2007.05.21 주차수당없이 기본수당에 포함됨.)
'                    If MDC_COMpanyGubun = "CL" Then
'                        If (Trim$(WG03.U_PAYTYP) = "3" Or Trim$(WG03.U_PAYTYP) = "4") Then     '/ 한샘 반장(월급직)은 주휴수당안나감
'                           WG03.U_CSUAMT(iCol) = WG03.U_CSUAMT(iCol) + (MDC_SetMod.RInt(WK_DAYAMT * WG23_JCHDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol)))
'                        End If
'                    End If
                End Select
                
            Next iCol
        End If
        
    Set oRecordSet = Nothing
    Set pRecordset = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Set pRecordset = Nothing
    Sbo_Application.StatusBar.SetText "PayRoll_Process1 Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Function BonusRoll_Process() As String
On Error GoTo Error_Message
    Dim ErrNum              As Integer
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim iCol                As Long:    Dim kCol        As Long
    Dim U_INPGBN            As String
    
    Dim SuSilStr            As String
    Dim Tmp_CSUAMT          As Double
    Dim WK_GNMDAY           As Integer
    Dim S_GNMMON(1 To 8)    As Integer
    Dim BNS_RODLEN          As Integer: Dim BNS_ROUNDT  As String
    Dim WK_ENDDAT           As String:  Dim WK_LMTDAT   As String
    Dim BnsResult           As String

    
    ErrNum = 0
    BNS_ROUNDT = "F":   BNS_RODLEN = 10
    WK_ENDDAT = Trim$(oDS_PH_PY111A.GetValue("U_bGNEDAT", 0))
    WK_LMTDAT = Trim$(oDS_PH_PY111A.GetValue("U_bEXPDAT", 0))
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'/1. 기본급여사항가져오기
    WK_STDAMT = sRecordset.Fields("U_BNSAMT").VALUE   '/ 기본 상여금
    
    '// 1.1. 공식가져오기-수당
    sQry = " SELECT T0.U_LINSEQ, T0.U_CSUCOD, T0.U_SILCUN, T0.U_SILCOD, T1.U_BNSLEN, T1.U_BNSRND "
    sQry = sQry & " FROM [@PH_PY106B] T0 INNER JOIN [@PH_PY106A] T1 ON T0.DocEntry = T1.DocEntry "
    sQry = sQry & " WHERE   T1.U_YM = '" & Trim$(U_CSUCOD) & "'"
    sQry = sQry & " AND     T1.U_PAYTYP = '" & Trim$(WG03.U_PAYTYP) & "'"
    sQry = sQry & " ORDER BY CAST(T0.U_LINSEQ AS INT), T0.U_CSUCOD "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    '// 2010.04.05 최동권 추가
    '// 수당 계산식 설정에서 상여자리수와 끝전처리 세팅값을 가져옴
    If IsNull(oRecordSet.Fields("U_BNSLEN").VALUE) = False Then
        BNS_RODLEN = Val(oRecordSet.Fields("U_BNSLEN").VALUE)
    End If
    If IsNull(oRecordSet.Fields("U_BNSRND").VALUE) = False Then
        BNS_ROUNDT = oRecordSet.Fields("U_BNSRND").VALUE
    End If
    
    Do Until oRecordSet.EOF
        iCol = Val(oRecordSet.Fields("U_LINSEQ").VALUE)
        If iCol > 0 Then
            If oJOBTYP = "2" Then
                WK_C.GONSIL(iCol) = oRecordSet.Fields("U_SILCOD").VALUE   '// 상여계산식 읽어오기
            Else
                WK_C.GONSIL(iCol) = oRecordSet.Fields("U_SILCUN").VALUE   '// 급여계산식 읽어오기
            End If
        Else
            If Left$(oRecordSet.Fields("U_CSUCOD").VALUE, 1) = "X" Then
                SuSilStr = oRecordSet.Fields("U_SILCUN").VALUE
                '/ 2.2. 공식안 시스템제공코드값으로 변경
                SuSilStr = Change_GOSIL(SuSilStr)
                Select Case Trim$(oRecordSet.Fields("U_CSUCOD").VALUE)
                Case "X01"       '/ 기본일급
                     X01_Val = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                Case "X02"       '/ 기본시급
                     X02_Val = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                Case "X03"       '/ 통상일급
                     X03_Val = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                Case "X04"       '/ 통상시급
                     X04_Val = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                End Select
            End If
        End If
        oRecordSet.MoveNext
    Loop
    WG03.U_DAYAMT = X01_Val
    WG03.U_BASAMT = X02_Val
    
'// 2. 공식가지고 계산하기-수당
    SuSilStr = ""
    WK_BNSAMT = 0
    For iCol = 1 To 24
        If Trim$(WK_C.CSUCOD(iCol)) <> "" Then '/ 수당항목이 있는것만
            If Trim$(WK_C.BNSUSE(iCol)) = "Y" Then
                If Trim$(WK_C.GONSIL(iCol)) <> "" Then  '/ 공식이있으면 공식계산
                '/ 2.1. 공식-계산결과값가져오는거면..
                    SuSilStr = WK_C.GONSIL(iCol)   '/ 계산된값일 경우
                    For kCol = 1 To 24
                        SuSilStr = Replace(SuSilStr, "#" & Format$(kCol, "00"), WG03.U_CSUAMT(kCol))
                    Next kCol
                '/ 2.2. 공식안 시스템제공코드값으로 변경
                    SuSilStr = Change_GOSIL(SuSilStr)
                '/ 2.3. 공식계산하기
                    Tmp_CSUAMT = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                '/ 2.4. 정답가져오기

                     '/2010.04.05 최동권 수정 (상여끝전 처리값을 수당계산식 설정에서 가져옴)
'                    If Trim$(WK_C.CSUCOD(iCol)) = "A04" Or Trim$(WK_C.CSUCOD(iCol)) = "A01" Then
'                        BNS_RODLEN = WK_C.RODLEN(iCol)
'                        BNS_ROUNDT = WK_C.ROUNDT(iCol)
'                    End If
                   If oJOBTYP = "3" And Trim$(WK_C.CSUCOD(iCol)) = "A04" Then  '// 급상여:상여금수당만
                      WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(Tmp_CSUAMT, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                      WK_BNSAMT = WG03.U_CSUAMT(iCol)
                   ElseIf oJOBTYP = "2" Then                               '// 상여:상여모두
                      WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(Tmp_CSUAMT, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
                      If Trim$(WK_C.CSUCOD(iCol)) = "A01" And WG03.U_CSUAMT(iCol) = 0 Then WG03.U_CSUAMT(iCol) = WK_STDAMT
                      WK_BNSAMT = WK_BNSAMT + WG03.U_CSUAMT(iCol)
                   End If
                End If
            End If
        End If
    Next iCol
    
    '// 2010.04.05 최동권 추가
    '// 기간급여 평균으로 상여 계산(TNC 이슈사항)
    If oBNSCAL = "2" Or oBNSCAL = "3" Then
        BnsResult = BonusRoll_Process2
        If BnsResult <> "" Then
            ErrNum = 2
            GoTo Error_Message
        End If
    End If
  
  
'        SuSilStr = ""
'        If oRecordSet.RecordCount > 0 Then
'            SuSilStr = oRecordSet.Fields("U_SILCUN").Value   '/ 계산된값일 경우
'            '/ 2.2. 공식안 시스템제공코드값으로 변경
'            SuSilStr = Change_GOSIL(SuSilStr)   '/ 월총일수
'            If Trim$(SuSilStr) <> "" Then
'                WK_BNSAMT = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
'            Else
'                WK_BNSAMT = 0
'            End If
'        Else
'            WK_BNSAMT = sRecordSet.Fields("U_BNSAMT").Value
'        End If
'    End If
'/3.평균임금관련(급여세액기간)/~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'    'sQry = "SELECT SUM(CASE WHEN U_JOBTYP ='2' THEN 0 ELSE 1 END) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,"
'    '200907.22 수정 같은월 급여 2개일경우 2로 잡힘.
'    sQry = "SELECT COUNT(DISTINCT U_YM) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,"
'    sQry = sQry & "   SUM(U_GONG05) AS TOTGBH"
'    sQry = sQry & " FROM  [@PH_PY111A] WHERE U_MSTCOD  = '" & Trim$(WG03.U_MSTCOD) & "'"
'    sQry = sQry & "                    AND U_YM >= '" & Mid$(oSTRTAX, 1, 6) & "'"
'    If Trim$(oENDTAX) < oYM Then  '// 세액종료연월이 귀속연월보다 적을경우
'        sQry = sQry & "                AND U_YM <= '" & Mid$(oENDTAX, 1, 6) & "'"
'    Else                                   '// 세액종료연월보다 크거나 같으면 해당 계산중인 귀속연월,지급종류,지급구분은 제외
'        sQry = sQry & "                AND (U_YM <= '" & Mid$(oENDTAX, 1, 6) & "'"
'        sQry = sQry & "                AND NOT (U_YM = '" & oYM & "' AND (U_JOBTYP = '" & oJOBTYP & "' AND U_JOBGBN = '" & oJOBGBN & "')))"
'    End If
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount <> 0 Then
'       WG03.U_AVRPAY = IIf(IsNull(oRecordSet.Fields("TOTPAY").Value), 0, oRecordSet.Fields("TOTPAY").Value)
'       WG03.U_NABTAX = IIf(IsNull(oRecordSet.Fields("TOTGAB").Value), 0, oRecordSet.Fields("TOTGAB").Value)
''       WG03.U_TAXTRM = IIf(IsNull(oRecordSet.Fields("TOTTRM").Value), 0, oRecordSet.Fields("TOTTRM").Value)
'    End If
'    '/ 상여나가는 달 세액기간
'    TermCHK = False
'    If (oJOBTYP = "2" Or oJOBTYP = "3") Then
'        sQry = "SELECT COUNT(DISTINCT U_YM) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,"
'        sQry = sQry & "   SUM(U_GONG05) AS TOTGBH"
'        sQry = sQry & " FROM  [@PH_PY111A] WHERE U_MSTCOD  = '" & Trim$(WG03.U_MSTCOD) & "'"
'        sQry = sQry & "                      AND U_YM >= '" & Mid$(oSTRTAX, 1, 6) & "'"
'        sQry = sQry & "                      AND (U_YM < '" & Mid$(oENDTAX, 1, 6) & "'"
'        sQry = sQry & "                      AND U_JOBTYP <> '2'"
'        sQry = sQry & "                      OR (U_YM = '" & Mid$(oENDTAX, 1, 6) & "' AND (U_JOBTYP = '1' OR U_JOBTYP='3')))"
'        oRecordSet.DoQuery sQry
'        If oRecordSet.RecordCount <> 0 Then
'           WG03.U_TAXTRM = IIf(IsNull(oRecordSet.Fields("TOTTRM").Value), 0, oRecordSet.Fields("TOTTRM").Value)
'        End If
'        If oJOBTYP = "2" And WG03.U_TAXTRM = 0 Then
'            TermCHK = True  '/ 상여만일경우 세액기간이 0일경우 확인메세지출력
'        End If
'    End If

'//*********************************************************************************************************************/
'/ 2010.02.23  함미경 수정 (관련업체:대우IS등 )
'/ 수정 상여가 먼저나가는 업체 세액계산시 급여가져오는 세액기간과 상여가져오는 세액기간 별도 구분
'/ 위 세액기간 2개월동안 급여1개월 상여2개월분일경우 평균임금/1개월로 계산 세액높게 책정, 생성할 상여달 급여미지급일경우 급여세액기간 전전월~전월,상여 전월~당월
'/ 급여계산기준으로 개월수 계산하도록 변경함
'//*********************************************************************************************************************/
    '/ 상여기간계산
    sQry = "SELECT COUNT(DISTINCT U_YM) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,"
    sQry = sQry & "   SUM(U_GONG05) AS TOTGBH"
    sQry = sQry & " FROM  [@PH_PY111A] WHERE U_MSTCOD  = '" & Trim$(WG03.U_MSTCOD) & "'"
    sQry = sQry & "                    AND U_YM >= '" & Mid$(oSTRBNS, 1, 6) & "'"
    If Trim$(oENDBNS) < oYM Then  '// 세액종료연월이 귀속연월보다 적을경우
        sQry = sQry & "                AND U_YM <= '" & Mid$(oENDBNS, 1, 6) & "'"
    Else                                   '// 세액종료연월보다 크거나 같으면 해당 계산중인 귀속연월,지급종류,지급구분은 제외
        sQry = sQry & "                AND (U_YM <= '" & Mid$(oENDBNS, 1, 6) & "'"
        sQry = sQry & "                AND NOT (U_YM = '" & oYM & "' AND (U_JOBTYP = '" & oJOBTYP & "' AND U_JOBGBN = '" & oJOBGBN & "')))"
    End If
    sQry = sQry & "                    AND U_JOBTYP = '2'"  '/ 상여만
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount <> 0 Then
       WG03.U_AVRPAY = IIf(IsNull(oRecordSet.Fields("TOTPAY").VALUE), 0, oRecordSet.Fields("TOTPAY").VALUE)
       WG03.U_NABTAX = IIf(IsNull(oRecordSet.Fields("TOTGAB").VALUE), 0, oRecordSet.Fields("TOTGAB").VALUE)
'       WG03.U_TAXTRM = IIf(IsNull(oRecordSet.Fields("TOTTRM").Value), 0, oRecordSet.Fields("TOTTRM").Value)
    End If
    '/ 상여나가는 달 세액기간
    TermCHK = False
    If (oJOBTYP = "2" Or oJOBTYP = "3") Then
        sQry = "SELECT COUNT(DISTINCT U_YM) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,"
        sQry = sQry & "   SUM(U_GONG05) AS TOTGBH"
        sQry = sQry & " FROM  [@PH_PY111A] WHERE U_MSTCOD  = '" & Trim$(WG03.U_MSTCOD) & "'"
        sQry = sQry & "                      AND U_YM >= '" & Mid$(oSTRTAX, 1, 6) & "'"
        If Trim$(oENDTAX) < oYM Then  '// 세액종료연월이 귀속연월보다 적을경우
            sQry = sQry & "                AND U_YM <= '" & Mid$(oENDTAX, 1, 6) & "'"
        Else                                   '// 세액종료연월보다 크거나 같으면 해당 계산중인 귀속연월,지급종류,지급구분은 제외
            sQry = sQry & "                AND (U_YM <= '" & Mid$(oENDTAX, 1, 6) & "'"
            sQry = sQry & "                AND NOT (U_YM = '" & oYM & "' AND (U_JOBTYP = '" & oJOBTYP & "' AND U_JOBGBN = '" & oJOBGBN & "')))"
        End If
        sQry = sQry & "                    AND U_JOBTYP <> '2'"  '/ 상여만
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount <> 0 Then
            WG03.U_AVRPAY = WG03.U_AVRPAY + IIf(IsNull(oRecordSet.Fields("TOTPAY").VALUE), 0, oRecordSet.Fields("TOTPAY").VALUE)
            WG03.U_NABTAX = WG03.U_NABTAX + IIf(IsNull(oRecordSet.Fields("TOTGAB").VALUE), 0, oRecordSet.Fields("TOTGAB").VALUE)
            WG03.U_TAXTRM = IIf(IsNull(oRecordSet.Fields("TOTTRM").VALUE), 0, oRecordSet.Fields("TOTTRM").VALUE)
        End If
        If oJOBTYP = "2" And WG03.U_TAXTRM = 0 Then
            TermCHK = True  '/ 상여만일경우 세액기간이 0일경우 확인메세지출력
        End If
    End If
 
'/4. 근속년월일 계산 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
   S_GNMMON(1) = 0: S_GNMMON(2) = 0: S_GNMMON(3) = 0: S_GNMMON(4) = 0: S_GNMMON(5) = 0
   S_GNMMON(6) = 0: S_GNMMON(7) = 0: S_GNMMON(8) = 0
   WK_GNMDAY = 0
   Call MDC_SetMod.Term2(WG03.U_INPDAT, WK_ENDDAT)
   WG03.U_GNSYER = ZPAY_GBL_GNSYER
   WG03.U_GNSMON = ZPAY_GBL_GNSMON

   WK_GNMDAY = DateDiff("d", Format$(WG03.U_INPDAT, "0000-00-00"), Format$(WK_ENDDAT, "0000-00-00")) + 1
   '/ 적용1
   Select Case Trim$(oDS_PH_PY111A.GetValue("U_AP1GBN", 0))
   Case "3", "4":   S_GNMMON(1) = WK_GNMDAY         '/ 일수
   Case Else:       S_GNMMON(1) = ZPAY_GBL_GNMMON   '/ 개월수
   End Select
   '/ 적용2
   Select Case Trim$(oDS_PH_PY111A.GetValue("U_AP2GBN", 0))
   Case "3", "4":   S_GNMMON(2) = WK_GNMDAY         '/ 일수
   Case Else:       S_GNMMON(2) = ZPAY_GBL_GNMMON   '/ 개월수
   End Select
   '/ 적용3
   Select Case Trim$(oDS_PH_PY111A.GetValue("U_AP3GBN", 0))
   Case "3", "4":   S_GNMMON(3) = WK_GNMDAY         '/ 일수
   Case Else:       S_GNMMON(3) = ZPAY_GBL_GNMMON   '/ 개월수
   End Select
   '/ 적용4
   Select Case Trim$(oDS_PH_PY111A.GetValue("U_AP4GBN", 0))
   Case "3", "4":   S_GNMMON(4) = WK_GNMDAY         '/ 일수
   Case Else:       S_GNMMON(4) = ZPAY_GBL_GNMMON   '/ 개월수
   End Select
   '/ 적용5
   Select Case Trim$(oDS_PH_PY111A.GetValue("U_AP5GBN", 0))
   Case "3", "4":   S_GNMMON(5) = WK_GNMDAY         '/ 일수
   Case Else:       S_GNMMON(5) = ZPAY_GBL_GNMMON   '/ 개월수
   End Select
   '/ 적용6
   Select Case Trim$(oDS_PH_PY111A.GetValue("U_AP6GBN", 0))
   Case "3", "4":   S_GNMMON(6) = WK_GNMDAY         '/ 일수
   Case Else:       S_GNMMON(6) = ZPAY_GBL_GNMMON   '/ 개월수
   End Select
   '/ 적용7
   Select Case Trim$(oDS_PH_PY111A.GetValue("U_AP7GBN", 0))
   Case "3", "4":   S_GNMMON(7) = WK_GNMDAY         '/ 일수
   Case Else:       S_GNMMON(7) = ZPAY_GBL_GNMMON   '/ 개월수
   End Select
   '/ 적용8
   Select Case Trim$(oDS_PH_PY111A.GetValue("U_AP8GBN", 0))
   Case "3", "4":   S_GNMMON(8) = WK_GNMDAY         '/ 일수
   Case Else:       S_GNMMON(8) = ZPAY_GBL_GNMMON   '/ 개월수
   End Select
'/5. 상여금계산  /
   
    If oJOBTYP = "2" Then
        WG03.U_TAXTRM = IIf(WG03.U_TAXTRM < 1, 1, WG03.U_TAXTRM)
    End If
    WK_APPBNS = 0
    WG03.U_BNSRAT = Trim(oDS_PH_PY111A.GetValue("U_bBNSRAT", 0))
    
    '/*******입사구분 경력으로 들어왔음 상관없이 개월수 상관없이 적용율 100% 지급함 (신입인경우만 상여율 적용혀)
    U_INPGBN = MDC_SetMod.Get_ReData("U_RelCd", "U_Minor", "[@PS_HR200L]", "'" & sRecordset.Fields("U_INPGBN").VALUE & "'", " AND Code =N'P133'")
    If Trim$(U_INPGBN) = "3" Then  '// 3-경력자이면 100적용율
        WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT1", 0))  '/ 경력직이면 맨 상위적용율 적용되도록.
        WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                        (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0))
        WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0))
    Else
        Select Case True
        '/ 적용1: 이상
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP1GBN", 0)) <> "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP1GBN", 0)) <> "4") And _
              S_GNMMON(1) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH1", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT1", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0))
        '/ 적용1: 미만
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP1GBN", 0)) = "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP1GBN", 0)) = "4") And _
              S_GNMMON(1) < Val(oDS_PH_PY111A.GetValue("U_bMONTH1", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT1", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0))
        '/ 적용2: 이상
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP2GBN", 0)) <> "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP2GBN", 0)) <> "4") And _
              S_GNMMON(2) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH2", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT2", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT2", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT2", 0))
        '/ 적용2: 미만
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP2GBN", 0)) = "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP2GBN", 0)) = "4") And _
              S_GNMMON(2) < Val(oDS_PH_PY111A.GetValue("U_bMONTH2", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT2", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT2", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT2", 0))
        '/ 적용3: 이상
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP3GBN", 0)) <> "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP3GBN", 0)) <> "4") And _
              S_GNMMON(3) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH3", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT3", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT3", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT3", 0))
        '/ 적용3: 미만
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP3GBN", 0)) = "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP3GBN", 0)) = "4") And _
              S_GNMMON(3) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH3", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT3", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT3", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT3", 0))
        '/ 적용4: 이상
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP4GBN", 0)) <> "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP4GBN", 0)) <> "4") And _
              S_GNMMON(4) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH4", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT4", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT4", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT4", 0))
        '/ 적용4: 미만
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP4GBN", 0)) = "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP4GBN", 0)) = "4") And _
              S_GNMMON(4) < Val(oDS_PH_PY111A.GetValue("U_bMONTH4", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT4", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT4", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT4", 0))
        '/ 적용5: 이상
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP5GBN", 0)) <> "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP5GBN", 0)) <> "4") And _
              S_GNMMON(5) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH5", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT5", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT5", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT5", 0))
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP5GBN", 0)) = "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP5GBN", 0)) = "4") And _
              S_GNMMON(5) < Val(oDS_PH_PY111A.GetValue("U_bMONTH5", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT5", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT5", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT5", 0))
        '/ 적용6: 이상
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP6GBN", 0)) <> "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP6GBN", 0)) <> "4") And _
              S_GNMMON(6) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH6", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT6", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT6", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT6", 0))
        '/ 적용6: 미만
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP6GBN", 0)) = "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP6GBN", 0)) = "4") And _
              S_GNMMON(6) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH6", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT6", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT6", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT6", 0))
        '/ 적용7: 이상
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP7GBN", 0)) <> "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP7GBN", 0)) <> "4") And _
              S_GNMMON(7) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH7", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT7", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT7", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT7", 0))
        '/ 적용7: 미만
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP7GBN", 0)) = "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP7GBN", 0)) = "4") And _
              S_GNMMON(7) < Val(oDS_PH_PY111A.GetValue("U_bMONTH7", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT7", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT7", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT7", 0))
        '/ 적용8: 이상
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP8GBN", 0)) <> "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP8GBN", 0)) <> "4") And _
              S_GNMMON(8) >= Val(oDS_PH_PY111A.GetValue("U_bMONTH8", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT8", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT8", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT8", 0))
        Case (Trim$(oDS_PH_PY111A.GetValue("U_AP8GBN", 0)) = "2" Or _
              Trim$(oDS_PH_PY111A.GetValue("U_AP8GBN", 0)) = "4") And _
              S_GNMMON(8) < Val(oDS_PH_PY111A.GetValue("U_bMONTH8", 0))
            WG03.U_APPRAT = Val(oDS_PH_PY111A.GetValue("U_bAPPRAT8", 0))
            WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * _
                            (WG03.U_BNSRAT / 100) + Val(oDS_PH_PY111A.GetValue("U_bAPPAMT8", 0))
            WK_APPBNS = Val(oDS_PH_PY111A.GetValue("U_bAPPAMT8", 0))
        End Select
    End If
    
    '// 2010.04.07 최동권 추가
    '// 개인별 상여율 적용(TNC 이슈사항)
    BnsResult = BonusRate_Process
    If BnsResult <> "" Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    'WG03.U_BONUSS = MDC_SetMod.IInt(WG03.U_BONUSS, 10)
    'MDC_SetMod.RInt(Tmp_CSUAMT, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
    If BNS_ROUNDT = "" Then BNS_ROUNDT = "F"
    If BNS_RODLEN = 0 Then BNS_RODLEN = 10
    WG03.U_BONUSS = MDC_SetMod.RInt(WG03.U_BONUSS, BNS_RODLEN, BNS_ROUNDT)
    '/ 당월퇴사자처리(당월만근시 지급, 중도퇴사자 제외)
    If Trim$(WG03.U_OUTDAT) <> "" Then
        If WK_LMTDAT > WG03.U_OUTDAT Then
            WG03.U_BONUSS = 0
        End If
    End If
    
    'If oJOBTYP = "2" Then    '/ 상여만이면
    '    WG03.U_CSUAMT(1) = WG03.U_BONUSS
    'Else
    '    WG03.U_TOTPAY = WG03.U_TOTPAY + WG03.U_BONUSS
    'End If
    
    Set oRecordSet = Nothing
    BonusRoll_Process = ""
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        BonusRoll_Process = "수당계산식에 일치하는 자료가 없습니다."
    ElseIf ErrNum = 2 Then
        BonusRoll_Process = BnsResult
    Else
        BonusRoll_Process = Err.Number & Space$(10) & Err.Description
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : BonusRoll_Process2
' Author    : Choi Dong Kwon
' Date      : 2010-04-06
' Purpose   : 계산식이 아닌 방법으로 상여계산시 사용
'---------------------------------------------------------------------------------------
Private Function BonusRoll_Process2() As String
On Error GoTo Error_Message
    Dim iCol        As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If oBNSCAL = "2" Or oBNSCAL = "3" Then
        '// Procedure Result
        '// U_TOTPAY : 급여 총액
        '// U_PAYCNT : 급여 기간 월수
        '// U_PAYTRM : 급여 실지급 월수
        '// U_AVRPAY : 평균 급여
        sQry = "EXEC PH_PY111_BONUSS '" & Trim$(WG03.U_MSTCOD) & "', '" & oSTRTAX & "', '" & oENDTAX & "', '" & _
                                        oSTRBNS & "', '" & oENDBNS & "', '" & oBNSCAL & "'"
        oRecordSet.DoQuery sQry
        WK_BNSAMT = oRecordSet.Fields("U_AVRPAY").VALUE
        WG03.U_BONUSS = oRecordSet.Fields("U_AVRPAY").VALUE
        If oRecordSet.Fields("U_AVRPAY").VALUE <= 0 Then
            WK_BNSAMT = 0
            WG03.U_BONUSS = 0
        End If
        '// 수당항목 중에 상여금(A04)이 있을 경우 상여금 금액에 넣을 것
        For iCol = 1 To 24
            If WK_C.CSUCOD(iCol) = "A04" Then
                WG03.U_CSUAMT(iCol) = oRecordSet.Fields("U_AVRPAY").VALUE
            End If
        Next iCol
    End If

    Set oRecordSet = Nothing
    BonusRoll_Process2 = ""
    
    Exit Function
Error_Message:
    Set oRecordSet = Nothing
    BonusRoll_Process2 = Err.Description

End Function

'---------------------------------------------------------------------------------------
' Procedure : BonusRate_Process
' Author    : Choi Dong Kwon
' Date      : 2010-04-06
' Purpose   : 개인별 상여율 계산
'---------------------------------------------------------------------------------------
Private Function BonusRate_Process() As String
On Error GoTo Error_Message
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '// 개인별 상여율이 등록되어 있는지 체크
    sQry = "       SELECT T1.U_BNSRAT, T1.U_BNSAMT "
    sQry = sQry & "FROM   [@ZPY315H] T0 "
    sQry = sQry & "       INNER JOIN [@ZPY315L] T1 ON T0.DocEntry = T1.DocEntry "
    sQry = sQry & "WHERE  T0.U_YM = '" & oYM & "' "
    sQry = sQry & "AND    T0.U_JOBTYP = '" & oJOBTYP & "' "
    sQry = sQry & "AND    T0.U_JOBGBN = '" & oJOBGBN & "' "
    sQry = sQry & "AND    T1.U_MSTCOD = '" & Trim$(WG03.U_MSTCOD) & "' "
    sQry = sQry & "AND   (T1.U_BNSRAT <> 0 OR T1.U_BNSAMT <> 0) "
    
    oRecordSet.DoQuery sQry
    
    '// 개인별 상여율이 등록되어 있지 않은 경우 Skip
    If oRecordSet.RecordCount > 0 Then
'        WG03.U_BNSRAT = Val(oRecordSet.Fields("U_BNSRAT").Value)
        WG03.U_APPRAT = Val(oRecordSet.Fields("U_BNSRAT").VALUE)
        WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT + Val(oRecordSet.Fields("U_BNSAMT").VALUE)
        WK_APPBNS = WK_BNSAMT
    End If
    BonusRate_Process = ""
    
    Set oRecordSet = Nothing
    Exit Function
Error_Message:
    Set oRecordSet = Nothing
    BonusRate_Process = Err.Description
End Function

Private Sub ChangeItem_Process()
On Error GoTo Error_Message

    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim iCol        As Integer

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    '/ 변동사항 조회
    sQry = " SELECT  T0.U_CSUTYP, T0.U_CSUCOD, T0.U_CSUAMT"
    sQry = sQry & " FROM [@ZPY302L] T0 INNER JOIN [@ZPY302H] T1 ON T0.DocEntry = T1.DocEntry"
    sQry = sQry & " WHERE T1.U_YM = '" & oYM & "'"
    sQry = sQry & " AND T1.U_JOBTYP = '" & oJOBTYP & "'"
    sQry = sQry & " AND T1.U_JOBGBN = '" & oJOBGBN & "'"
'   sQry = sQry & " AND T1.U_JOBTRG = '" & oJOBTRG & "'"
    sQry = sQry & " AND T0.U_MSTCOD = '" & Trim$(WG03.U_MSTCOD) & "'"
    sQry = sQry & " ORDER BY T0.U_CSUTYP, T0.U_CSUCOD"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        If oRecordSet.Fields("U_CSUTYP").VALUE = "1" Then     '/ 수당항목
            For iCol = 1 To 24
                If oRecordSet.Fields("U_CSUCOD").VALUE = WK_C.CSUCOD(iCol) Then
                    WG03.U_CSUAMT(iCol) = WG03.U_CSUAMT(iCol) + oRecordSet.Fields("U_CSUAMT").VALUE
                End If
            Next iCol
        ElseIf oRecordSet.Fields("U_CSUTYP").VALUE = "2" Then '/ 공제항목
            For iCol = 1 To 18
                If oRecordSet.Fields("U_CSUCOD").VALUE = WK_G.GONCOD(iCol) Then
                    WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + oRecordSet.Fields("U_CSUAMT").VALUE
                End If
            Next iCol
        End If
        oRecordSet.MoveNext
    Loop
    '/ 건강보험 정산 조회
    sQry = " SELECT  SUM(ISNULL(T0.U_CHAMED,0)) AS U_MEDJSN"
    sQry = sQry & " FROM [@ZPY311L] T0 INNER JOIN [@ZPY311H] T1 ON T0.DocEntry = T1.DocEntry"
    sQry = sQry & " WHERE T1.U_YM = '" & oYM & "'"
    sQry = sQry & " AND T1.U_JOBTYP = '" & oJOBTYP & "'"
    sQry = sQry & " AND T1.U_JOBGBN = '" & oJOBGBN & "'"
    sQry = sQry & " AND T1.U_JOBTRG = '" & oJOBTRG & "'"
    sQry = sQry & " AND T0.U_MSTCOD = '" & Trim$(WG03.U_MSTCOD) & "'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        For iCol = 1 To 18
            If (G08_CHK = False And WK_G.GONCOD(iCol) = "G04") Or (G08_CHK = True And WK_G.GONCOD(iCol) = "G08") Then
                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + oRecordSet.Fields("U_MEDJSN").VALUE
                WG03.U_TOTGON = WG03.U_TOTGON + oRecordSet.Fields("U_MEDJSN").VALUE
            End If
        Next iCol
        oRecordSet.MoveNext
    Loop
    
'/ 급상여변동자료(기간) -변동금액으로 대체/~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    sQry = "SELECT  T0.U_CSUTYP, T0.U_CSUCOD, T0.U_CSUAMT, T0.U_MSTCOD"
    sQry = sQry & " FROM [@ZPY312L] T0 INNER JOIN [@ZPY312H] T1 ON T0.DocEntry = T1.DocEntry"
    sQry = sQry & " WHERE T1.U_JOBTYP = '" & oJOBTYP & "'"
    sQry = sQry & " AND   T1.U_JOBGBN = '" & oJOBGBN & "'"
    sQry = sQry & " AND   T0.U_CSUTYP = '1'"    '수당항목만
    sQry = sQry & " AND   T0.U_MSTCOD = '" & Trim$(WG03.U_MSTCOD) & "'"
    sQry = sQry & " AND   " & "'" & oYM & "' BETWEEN T0.U_STRYMM AND T0.U_ENDYMM"
    sQry = sQry & " ORDER BY T0.U_MSTCOD, T1.U_YM DESC, T0.U_STRYMM DESC"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        If oJOBTYP <> "3" Then
            WG03.U_BONUSS = 0
        End If
        Do Until oRecordSet.EOF
            For iCol = 1 To 24
                If oRecordSet.Fields("U_CSUCOD").VALUE = WK_C.CSUCOD(iCol) Then
                    WG03.U_CSUAMT(iCol) = oRecordSet.Fields("U_CSUAMT").VALUE
                    If oJOBTYP = "3" And Trim$(WK_C.CSUCOD(iCol)) = "A04" Then
                        WG03.U_BONUSS = oRecordSet.Fields("U_CSUAMT").VALUE
                        WG03.U_APPRAT = "100"
                       ' WG03.U_AVRPAY = 0    '/ 2011.10.05 일성신약 변동자료기간별로 상여데이터올려도 평균임금반영하여 계산되도록
                       ' WG03.U_NABTAX = 1
                        WG03.U_BNSRAT = "100"
                        WK_APPBNS = WG03.U_BONUSS
                    ElseIf oJOBTYP = "2" And Trim$(WK_C.BNSUSE(iCol)) = "Y" Then
                        WG03.U_BONUSS = WG03.U_BONUSS + oRecordSet.Fields("U_CSUAMT").VALUE
                        WG03.U_APPRAT = "100"
                       ' WG03.U_AVRPAY = 0    '/ 2011.10.05 일성신약 변동자료기간별로 상여데이터올려도 평균임금반영하여 계산되도록
                       ' WG03.U_NABTAX = 1
                        WG03.U_BNSRAT = "100"
                        WK_APPBNS = WG03.U_BONUSS
                    End If
                End If
            Next iCol
            oRecordSet.MoveNext
        Loop
    End If
    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "PH_PY111_Save Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub JeongSan_Process()
    Dim iCol             As Integer
    
    '/ 소득세,주민세항목
    For iCol = 1 To 18
        'If (G06_CHK = False And WK_G.GONCOD(iCol) = "G01") Or (G06_CHK = True And WK_G.GONCOD(iCol) = "G06") Then
        If (G06_CHK = True And WK_G.GONCOD(iCol) = "G06") Then
            WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + Val(sRecordset.Fields("U_CHAGAB").VALUE)
        'ElseIf (G07_CHK = True And WK_G.GONCOD(iCol) = "G07") Or (G07_CHK = False And WK_G.GONCOD(iCol) = "G02") Then
        ElseIf (G07_CHK = True And WK_G.GONCOD(iCol) = "G07") Then
            WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + Val(sRecordset.Fields("U_CHAJUM").VALUE)
        ElseIf (G91_CHK = True And WK_G.GONCOD(iCol) = "G91") Then  '/ 농특세
            WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + Val(sRecordset.Fields("U_CHANON").VALUE)
        End If
    Next iCol
    '/ 총공제금액
    WG03.U_TOTGON = WG03.U_TOTGON + Val(sRecordset.Fields("U_CHAGAB").VALUE) + Val(sRecordset.Fields("U_CHAJUM").VALUE) + Val(sRecordset.Fields("U_CHANON").VALUE)
    '/ 실지급액
    WG03.U_SILJIG = MDC_SetMod.IInt(WG03.U_TOTPAY - WG03.U_TOTGON, 1)
    
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Sbo_Application.StatusBar.SetText "JeongSan_Process Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub PAY_Tax()
On Error GoTo Error_Message
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim iCol        As Integer
    Dim GABGUN As Double, JUMINN As Double, GBHAMT As Double, MEDAMT As Double, KUKAMT As Double
    Dim INCOME As Double, GWASEE As Double, BTAX01 As Double, BTAX02 As Double, BTAX03 As Double, BTAX04 As Double, BTAX05 As Double, BTAX06 As Double, BTAX07 As Double
    Dim MONPAY As Double, BT1AMT As Double, FODCSU As Double, CARCSU As Double, FRNCSU As Double, CHLCSU As Double, YGUCSU As Double, GITBTX As Double, GITBTX_N As Double
    Dim BNSCSU As Double, BT1KUM As Double, TOTGBH As Double, GBHRAT As Double
    Dim SuSilStr    As String, Tmp_CSUAMT   As Double
    Dim KUKEND      As String, KUKSTR       As String, WK_MEDAMT    As Double, GBHEND   As String, GBHSTR   As String
    Dim U_NJYYMM    As String, U_NJYRAT     As Double, U_NJCRAT     As Double, NJCMED   As Double, MEDCHK   As String
    Dim kCol        As Long, WK_GBHAMT      As Double
    Dim WK_YBTXAM(1 To 7) As Double
    Dim WK_BNSCHK   As Boolean
    
    Dim Dw_Ilsu As Integer
   
    GABGUN = 0: JUMINN = 0: GBHAMT = 0: MEDAMT = 0: KUKAMT = 0
    INCOME = 0: GWASEE = 0: BTAX01 = 0: BTAX02 = 0: BTAX03 = 0: BTAX04 = 0: BTAX05 = 0: BTAX06 = 0: BTAX07 = 0:
    MONPAY = 0: BT1AMT = 0: FODCSU = 0: CARCSU = 0: TOTGBH = 0: CHLCSU = 0: YGUCSU = 0
    BNSCSU = 0: GITBTX = 0: GITBTX_N = 0
    U_NJYYMM = 0: U_NJYRAT = 0: U_NJCRAT = 0: NJCMED = 0
    
    Dw_Ilsu = 0

'/
    WK_BNSCHK = False
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

'/ 급여총액 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'/ 월정체크, 한도액, 생산비과대상자,부양가족공제,고용보험계산자
    For iCol = 1 To 24
        If oJOBTYP = "3" And WK_C.CSUCOD(iCol) = "A04" Then
            If oDS_PH_PY111A.GetValue("U_CHGCHK", 0) = "N" Then
                WG03.U_CSUAMT(iCol) = WG03.U_BONUSS
                BNSCSU = BNSCSU + Val(WG03.U_CSUAMT(iCol))
            Else
                WG03.U_BONUSS = WG03.U_CSUAMT(iCol)
            End If
        End If
        If Trim$(WK_C.GWATYP(iCol)) = "" Then WK_C.GWATYP(iCol) = "1"
        Select Case WK_C.GWATYP(iCol)
        Case "1"   '/ 과세
            If WK_C.CSUKUM(iCol) > 0 And WG03.U_CSUAMT(iCol) > WK_C.CSUKUM(iCol) Then
               WG03.U_CSUAMT(iCol) = Val(WK_C.CSUKUM(iCol))
            End If
        Case "2"   '/ 식대보조
            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
                FODCSU = 0
            Else
                FODCSU = Val(WG03.U_CSUAMT(iCol))
            End If
            If WK_C.CSUKUM(iCol) > 0 And FODCSU > WK_C.CSUKUM(iCol) Then
                BTAX02 = BTAX02 + WK_C.CSUKUM(iCol)
            Else
                BTAX02 = BTAX02 + FODCSU
            End If
        Case "3"   '/ 차량보조
            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
                CARCSU = 0
            Else
                CARCSU = Val(WG03.U_CSUAMT(iCol))
            End If
            If WK_C.CSUKUM(iCol) > 0 And CARCSU > WK_C.CSUKUM(iCol) Then
                BTAX02 = BTAX02 + WK_C.CSUKUM(iCol)
            Else
                BTAX02 = BTAX02 + CARCSU
            End If
        Case "4"   '/ 생산비과
            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
                BT1AMT = BT1AMT + 0
            Else
                BT1AMT = BT1AMT + Val(WG03.U_CSUAMT(iCol))
            End If
            
            If BT1KUM < WK_C.CSUKUM(iCol) Then BT1KUM = WK_C.CSUKUM(iCol)    '비과세4-마지막줄의 한도액만
        Case "5"   '/ 국외비과세
            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
                FRNCSU = FRNCSU + 0
            Else
                FRNCSU = FRNCSU + Val(WG03.U_CSUAMT(iCol))
            End If
            If WK_C.CSUKUM(iCol) > 0 And FRNCSU > WK_C.CSUKUM(iCol) Then
                Select Case UCase$(WK_C.BTXCOD(iCol))
                Case "M01":       WG03.U_BTXM01 = WG03.U_BTXM01 + WK_C.CSUKUM(iCol)
                Case "M02":       WG03.U_BTXM02 = WG03.U_BTXM02 + WK_C.CSUKUM(iCol)
                Case "M03":       WG03.U_BTXM03 = WG03.U_BTXM03 + WK_C.CSUKUM(iCol)
                Case Else:        BTAX03 = BTAX03 + WK_C.CSUKUM(iCol)
                End Select
            Else
                Select Case UCase$(WK_C.BTXCOD(iCol))
                Case "M01":       WG03.U_BTXM01 = WG03.U_BTXM01 + FRNCSU
                Case "M02":       WG03.U_BTXM02 = WG03.U_BTXM02 + FRNCSU
                Case "M03":       WG03.U_BTXM03 = WG03.U_BTXM03 + FRNCSU
                Case Else:        BTAX03 = BTAX03 + FRNCSU
                End Select
            End If
        Case "6"   '/ 연구수당
            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
                YGUCSU = YGUCSU + 0
            Else
                YGUCSU = YGUCSU + Val(WG03.U_CSUAMT(iCol))
            End If
            If WK_C.CSUKUM(iCol) > 0 And YGUCSU > WK_C.CSUKUM(iCol) Then
                Select Case UCase$(WK_C.BTXCOD(iCol))
                Case "H06":       WG03.U_BTXH06 = WG03.U_BTXH06 + WK_C.CSUKUM(iCol)
                Case "H07":       WG03.U_BTXH07 = WG03.U_BTXH07 + WK_C.CSUKUM(iCol)
                Case "H08":       WG03.U_BTXH08 = WG03.U_BTXH08 + WK_C.CSUKUM(iCol)
                Case "H09":       WG03.U_BTXH09 = WG03.U_BTXH09 + WK_C.CSUKUM(iCol)
                Case "H10":       WG03.U_BTXH10 = WG03.U_BTXH10 + WK_C.CSUKUM(iCol)
                Case Else:        BTAX05 = BTAX05 + WK_C.CSUKUM(iCol)
                End Select
            Else
                Select Case UCase$(WK_C.BTXCOD(iCol))
                Case "H06":       WG03.U_BTXH06 = WG03.U_BTXH06 + YGUCSU
                Case "H07":       WG03.U_BTXH07 = WG03.U_BTXH07 + YGUCSU
                Case "H08":       WG03.U_BTXH08 = WG03.U_BTXH08 + YGUCSU
                Case "H09":       WG03.U_BTXH09 = WG03.U_BTXH09 + YGUCSU
                Case "H10":       WG03.U_BTXH10 = WG03.U_BTXH10 + YGUCSU
                Case Else:        BTAX05 = BTAX05 + YGUCSU
                End Select
            End If
        Case "7"   '/ 보육수당
            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
                CHLCSU = CHLCSU + 0
            Else
                CHLCSU = CHLCSU + Val(WG03.U_CSUAMT(iCol))
            End If
            If WK_C.CSUKUM(iCol) > 0 And CHLCSU > WK_C.CSUKUM(iCol) Then
                WG03.U_BTXQ01 = WG03.U_BTXQ01 + WK_C.CSUKUM(iCol)
            Else
                WG03.U_BTXQ01 = WG03.U_BTXQ01 + CHLCSU
            End If
        Case "8"   '/ 비과세-기타제출
            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
                GITBTX = 0
            Else
                GITBTX = Val(WG03.U_CSUAMT(iCol))  '/ 누적한도체크아님 각 수당별한도체크.
            End If
            If WK_C.CSUKUM(iCol) > 0 And GITBTX > WK_C.CSUKUM(iCol) Then
                Select Case UCase$(WK_C.BTXCOD(iCol))
                Case "G01":     WG03.U_BTXG01 = WG03.U_BTXG01 + WK_C.CSUKUM(iCol)
                Case "H01":     WG03.U_BTXH01 = WG03.U_BTXH01 + WK_C.CSUKUM(iCol)
                Case "H05":     WG03.U_BTXH05 = WG03.U_BTXH05 + WK_C.CSUKUM(iCol)
                Case "H11":     WG03.U_BTXH11 = WG03.U_BTXH11 + WK_C.CSUKUM(iCol)
                Case "H12":     WG03.U_BTXH12 = WG03.U_BTXH12 + WK_C.CSUKUM(iCol)
                Case "H13":     WG03.U_BTXH13 = WG03.U_BTXH13 + WK_C.CSUKUM(iCol)
                Case "I01":     WG03.U_BTXI01 = WG03.U_BTXI01 + WK_C.CSUKUM(iCol)
                Case "K01":     WG03.U_BTXK01 = WG03.U_BTXK01 + WK_C.CSUKUM(iCol)
                Case "R10":     WG03.U_BTXR10 = WG03.U_BTXR10 + WK_C.CSUKUM(iCol)
                Case "S01":     WG03.U_BTXS01 = WG03.U_BTXS01 + WK_C.CSUKUM(iCol)
                Case "T01":     WG03.U_BTXT01 = WG03.U_BTXT01 + WK_C.CSUKUM(iCol)
                Case "X01":     WG03.U_BTXX01 = WG03.U_BTXX01 + WK_C.CSUKUM(iCol)
                Case "Y01":     WG03.U_BTXY01 = WG03.U_BTXY01 + WK_C.CSUKUM(iCol)
                Case "Y02":     WG03.U_BTXY02 = WG03.U_BTXY02 + WK_C.CSUKUM(iCol)
                Case "Y03":     WG03.U_BTXY03 = WG03.U_BTXY03 + WK_C.CSUKUM(iCol)
                Case "Y20":     WG03.U_BTXY20 = WG03.U_BTXY20 + WK_C.CSUKUM(iCol)
                Case "Y21":     WG03.U_BTXY21 = WG03.U_BTXY21 + WK_C.CSUKUM(iCol)
                Case "Y22":     WG03.U_BTXY22 = WG03.U_BTXY22 + WK_C.CSUKUM(iCol)
                Case "Z01":     WG03.U_BTXZ01 = WG03.U_BTXZ01 + WK_C.CSUKUM(iCol)
                Case Else:      BTAX04 = BTAX04 + WK_C.CSUKUM(iCol)
                End Select
            Else
                Select Case UCase$(WK_C.BTXCOD(iCol))
                Case "G01":     WG03.U_BTXG01 = WG03.U_BTXG01 + GITBTX
                Case "H01":     WG03.U_BTXH01 = WG03.U_BTXH01 + GITBTX
                Case "H05":     WG03.U_BTXH05 = WG03.U_BTXH05 + GITBTX
                Case "H11":     WG03.U_BTXH11 = WG03.U_BTXH11 + GITBTX
                Case "H12":     WG03.U_BTXH12 = WG03.U_BTXH12 + GITBTX
                Case "H13":     WG03.U_BTXH13 = WG03.U_BTXH13 + GITBTX
                Case "I01":     WG03.U_BTXI01 = WG03.U_BTXI01 + GITBTX
                Case "K01":     WG03.U_BTXK01 = WG03.U_BTXK01 + GITBTX
                Case "R10":     WG03.U_BTXR10 = WG03.U_BTXR10 + GITBTX
                Case "S01":     WG03.U_BTXS01 = WG03.U_BTXS01 + GITBTX
                Case "T01":     WG03.U_BTXT01 = WG03.U_BTXT01 + GITBTX
                Case "X01":     WG03.U_BTXX01 = WG03.U_BTXX01 + GITBTX
                Case "Y01":     WG03.U_BTXY01 = WG03.U_BTXY01 + GITBTX
                Case "Y02":     WG03.U_BTXY02 = WG03.U_BTXY02 + GITBTX
                Case "Y03":     WG03.U_BTXY03 = WG03.U_BTXY03 + GITBTX
                Case "Y20":     WG03.U_BTXY20 = WG03.U_BTXY20 + GITBTX
                Case "Y21":     WG03.U_BTXY21 = WG03.U_BTXY21 + GITBTX
                Case "Y22":     WG03.U_BTXY22 = WG03.U_BTXY22 + GITBTX
                Case "Z01":     WG03.U_BTXZ01 = WG03.U_BTXZ01 + GITBTX
                Case Else:      BTAX04 = BTAX04 + GITBTX
                End Select
            End If
        Case Else   '/ 비과세-기타-미제출
            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
                GITBTX_N = 0
            Else
                GITBTX_N = Val(WG03.U_CSUAMT(iCol))    '/ 누적한도체크아님 각 수당별한도체크.
            End If
            If WK_C.CSUKUM(iCol) > 0 And GITBTX_N > WK_C.CSUKUM(iCol) Then
                BTAX07 = BTAX07 + WK_C.CSUKUM(iCol)
            Else
                BTAX07 = BTAX07 + GITBTX_N
            End If
        End Select
    
        If oYM >= "201001" Then '// 2010년부터 폐지되는 비과세
            WG03.U_BTXX01 = 0
            WG03.U_BTXY01 = 0
            WG03.U_BTXY02 = 0
            WG03.U_BTXY20 = 0
        End If
        
        '/ 상여만일경우 U_BONUSS=상여금 U_TOTPAY=상여금+기타소득 아닐경우 상여금만
        If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" And oDS_PH_PY111A.GetValue("U_CHGCHK", 0) = "N" Then
            If WK_BNSCHK = True Then    '/ 상여를 적용금액이나 변동금액을 적용하였을경우는 기본급또는 상여금(A04)에 나머지상여금액이 0
                WG03.U_CSUAMT(iCol) = 0
            Else
                '/ 상여 상여금에 포함 수당은 총지급액에 합산하지않아야함.
                If WK_APPBNS > 0 And WK_C.CSUCOD(iCol) = "A01" Then
                       WG03.U_CSUAMT(iCol) = WG03.U_BONUSS
                       WK_APPBNS = 0
                       WK_BNSCHK = True
                ElseIf WK_APPBNS > 0 And WK_C.CSUCOD(iCol) = "A04" Then '/ 기본급이 상여포함구분이 아닐경우 상여금란에 표시
                       WG03.U_CSUAMT(iCol) = WG03.U_BONUSS
                       WK_APPBNS = 0
                       WK_BNSCHK = True
                End If
            End If
        Else
            WG03.U_TOTPAY = WG03.U_TOTPAY + Val(WG03.U_CSUAMT(iCol))
            '/ 월정급여
            If WK_C.MPYGBN(iCol) = "Y" Then
                MONPAY = MONPAY + Val(WG03.U_CSUAMT(iCol))
            End If
            '/ 고용보험대상금액
            If WK_C.GBHGBN(iCol) = "Y" Then
                TOTGBH = TOTGBH + Val(WG03.U_CSUAMT(iCol))
            End If
        End If
    Next iCol
    If oJOBTYP = "2" Then '/ 상여일경우 상여금금액 더 적용
        WG03.U_TOTPAY = WG03.U_TOTPAY + WG03.U_BONUSS
        MONPAY = MONPAY + WG03.U_BONUSS
        TOTGBH = TOTGBH + WG03.U_BONUSS
    End If
    If oJOBTYP <> "1" And oDS_PH_PY111A.GetValue("U_CHGCHK", 0) = "Y" Then
        Select Case oJOBTYP
        Case "1"
        Case "2"
            WG03.U_BONUSS = WG03.U_TOTPAY
            WG03.U_APPRAT = "100"
            WG03.U_BNSRAT = "100"
        Case "3"
            WG03.U_APPRAT = "100"
            WG03.U_BNSRAT = "100"
        End Select
    End If

    WG03.U_TOTPAY = Format$(WG03.U_TOTPAY, "##########0")
   If WG03.U_TOTPAY = 0 Then Exit Sub
' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 '/ 공제내역 계산/
' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/ 비과세산출 /*******************************************************************************/
    '/ 비과세누적액 산출
        WK_YBTXAM(1) = 0: WK_YBTXAM(2) = 0: WK_YBTXAM(3) = 0: WK_YBTXAM(4) = 0: WK_YBTXAM(5) = 0: WK_YBTXAM(6) = 0:: WK_YBTXAM(7) = 0
        sQry = "EXEC PH_PY111_BTX '" & oYM & "','" & oJOBTYP & "', '" & oJOBGBN & "', '" & Trim$(WG03.U_MSTCOD) & "'"
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount > 0 Then
            WK_YBTXAM(1) = Val(oRecordSet.Fields(0).VALUE)   '/ 연누적액(생산직비과세)
            WK_YBTXAM(2) = Val(oRecordSet.Fields(1).VALUE)   '/ 월누적(식대,차량보조)
            WK_YBTXAM(3) = Val(oRecordSet.Fields(2).VALUE)   '/ 월누적(국외비과세)
            WK_YBTXAM(4) = Val(oRecordSet.Fields(3).VALUE)   '/ 월누적(기타비과세(제출)
            WK_YBTXAM(5) = Val(oRecordSet.Fields(4).VALUE)   '/ 월누적(연구비과세)
            WK_YBTXAM(6) = Val(oRecordSet.Fields(5).VALUE)   '/ 월누적(출산보육비과세)
            WK_YBTXAM(7) = Val(oRecordSet.Fields(6).VALUE)   '/ 월누적-기타비과세(미제출)
        End If
        
    '/ 비과세3-국외비과세대상
    If sRecordset.Fields("U_FRGSEL").VALUE = "Y" And oJOBTYP <> "2" Then  '/ 비과세대상자만
        Select Case Trim$(TB1_BT3COD)
        Case "M02"
            If WG03.U_TOTPAY < WG01.TB1AMT(7) Then   '/ 월한도액(백오십만원)
               WG03.U_BTXM02 = WG03.U_TOTPAY
            Else
               WG03.U_BTXM02 = WG01.TB1AMT(7)
            End If
            If (WG03.U_BTXM02 + WK_YBTXAM(3)) > WG01.TB1AMT(7) Then WG03.U_BTXM02 = WG01.TB1AMT(7) - (WG03.U_BTXM02 + WK_YBTXAM(3))
            If WG03.U_BTXM02 < 0 Then WG03.U_BTXM02 = 0
        Case "M03"
            If WG03.U_TOTPAY < WG03.U_BTXM03 Then   '/ 한도제한제외
               WG03.U_BTXM03 = WG03.U_TOTPAY
            End If
            If WG03.U_BTXM03 < 0 Then WG03.U_BTXM03 = 0
        Case Else '/ M01
            If WG03.U_TOTPAY < WG01.TB1AMT(4) Then   '/ 월한도액(백만원)
               WG03.U_BTXM01 = WG03.U_TOTPAY
            Else
               WG03.U_BTXM01 = WG01.TB1AMT(4)
            End If
            If (WG03.U_BTXM01 + WK_YBTXAM(3)) > WG01.TB1AMT(4) Then WG03.U_BTXM01 = WG01.TB1AMT(4) - (WG03.U_BTXM01 + WK_YBTXAM(3))
            If WG03.U_BTXM01 < 0 Then WG03.U_BTXM01 = 0
        End Select
    End If
    GWASEE = WG03.U_TOTPAY - WG03.U_BTXM01 - WG03.U_BTXM02 - WG03.U_BTXM03
    
    '/ 비과세1-생산직 비과세
    If sRecordset.Fields("U_BX1SEL").VALUE = "Y" Then '/ 비과세대상자만
       If (MONPAY - BT1AMT) <= 1000000 Then
            WG03.U_BTXO01 = BT1AMT  '비과세 해당자
       Else
            WG03.U_BTXO01 = 0
       End If
       If BT1KUM <> 0 And WG03.U_BTXO01 > BT1KUM Then WG03.U_BTXO01 = BT1KUM  '/ 한도액구함
       '/ 연간 누적 비과세1 체크(상여계산시는 제외)
       If oJOBTYP <> "2" Then
          '/ (비과세1+연누적비과세)> 연240만원 THEN 비과세1 = 연240만원-(연누적비과세+비과세1)
          If (WG03.U_BTXO01 + WK_YBTXAM(1)) > WG01.TB1KUM(1) Then
             WG03.U_BTXO01 = WG01.TB1KUM(1) - WK_YBTXAM(1)
          End If
          If WG03.U_BTXO01 < 0 Then BTAX01 = 0
       End If
    End If

    If oYM < "200901" Then  '/ 2009년이전자료이면 이전비과세란으로
        '/ 생산직비과세
        BTAX01 = WG03.U_BTXO01
        '/ 국외비과세
        BTAX03 = WG03.U_BTXM01 + WG03.U_BTXM02 + WG03.U_BTXM03
        '/ 연구개발
        BTAX05 = WG03.U_BTXH06 + WG03.U_BTXH07 + WG03.U_BTXH08 + WG03.U_BTXH09 + WG03.U_BTXH10
        '/ 출산보육
        BTAX06 = WG03.U_BTXQ01
        '/ 기타지급조서제출분
        If oYM < "201001" Then
            BTAX04 = WG03.U_BTXG01 + WG03.U_BTXH01 + WG03.U_BTXH05 + WG03.U_BTXH11 + WG03.U_BTXH12 + WG03.U_BTXH13 + WG03.U_BTXI01 + WG03.U_BTXK01 + _
                     WG03.U_BTXS01 + WG03.U_BTXT01 + WG03.U_BTXX01 + WG03.U_BTXY01 + WG03.U_BTXY02 + WG03.U_BTXY03 + WG03.U_BTXY20 + WG03.U_BTXZ01
        Else
            BTAX04 = WG03.U_BTXG01 + WG03.U_BTXH01 + WG03.U_BTXH05 + WG03.U_BTXH11 + WG03.U_BTXH12 + WG03.U_BTXH13 + WG03.U_BTXI01 + WG03.U_BTXK01 + _
                     WG03.U_BTXS01 + WG03.U_BTXT01 + WG03.U_BTXY02 + WG03.U_BTXY03 + WG03.U_BTXY21 + WG03.U_BTXZ01
        End If
        WG03.U_BTXG01 = 0:    WG03.U_BTXH01 = 0:    WG03.U_BTXH05 = 0:    WG03.U_BTXH06 = 0:    WG03.U_BTXH07 = 0:    WG03.U_BTXH08 = 0
        WG03.U_BTXH09 = 0:    WG03.U_BTXH10 = 0:    WG03.U_BTXH11 = 0:    WG03.U_BTXH12 = 0:    WG03.U_BTXH13 = 0:    WG03.U_BTXI01 = 0
        WG03.U_BTXK01 = 0:    WG03.U_BTXM01 = 0:    WG03.U_BTXM02 = 0:    WG03.U_BTXM03 = 0:    WG03.U_BTXO01 = 0:    WG03.U_BTXQ01 = 0
        WG03.U_BTXR10 = 0:    WG03.U_BTXS01 = 0:    WG03.U_BTXT01 = 0:    WG03.U_BTXX01 = 0:    WG03.U_BTXY01 = 0:    WG03.U_BTXY02 = 0:
        WG03.U_BTXY03 = 0:    WG03.U_BTXY20 = 0:    WG03.U_BTXY21 = 0:    WG03.U_BTXY22 = 0: WG03.U_BTXZ01 = 0
    End If
    
    '/ 1.2) 과세대상급여 ( 총급여액-비과세소득 )
    If GWASEE < BTAX01 Then BTAX01 = GWASEE
    GWASEE = GWASEE - BTAX01
    '/ 비과세2-식대차량보조
    If GWASEE < BTAX02 Then BTAX02 = GWASEE
    GWASEE = GWASEE - BTAX02
    '/ 비과세4
    If GWASEE < BTAX04 Then BTAX04 = GWASEE
    GWASEE = GWASEE - BTAX04
    '/ 비과세5
    If GWASEE < BTAX05 Then BTAX05 = GWASEE
    GWASEE = GWASEE - BTAX05
    '/ 비과세6
    If GWASEE < BTAX06 Then BTAX06 = GWASEE
    GWASEE = GWASEE - BTAX06
    '/ 비과세7
    If GWASEE < BTAX07 Then BTAX07 = GWASEE
    GWASEE = GWASEE - BTAX07
    '/ 비과세 (G01~Z01)
    If GWASEE < WG03.U_BTXG01 Then WG03.U_BTXG01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXG01
    If GWASEE < WG03.U_BTXH01 Then WG03.U_BTXH01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH01
    If GWASEE < WG03.U_BTXH05 Then WG03.U_BTXH05 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH05
    If GWASEE < WG03.U_BTXH06 Then WG03.U_BTXH06 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH06
    If GWASEE < WG03.U_BTXH07 Then WG03.U_BTXH07 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH07
    If GWASEE < WG03.U_BTXH08 Then WG03.U_BTXH08 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH08
    If GWASEE < WG03.U_BTXH09 Then WG03.U_BTXH09 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH09
    If GWASEE < WG03.U_BTXH10 Then WG03.U_BTXH10 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH10
    If GWASEE < WG03.U_BTXH11 Then WG03.U_BTXH11 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH11
    If GWASEE < WG03.U_BTXH12 Then WG03.U_BTXH12 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH12
    If GWASEE < WG03.U_BTXH13 Then WG03.U_BTXH13 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXH13
    If GWASEE < WG03.U_BTXI01 Then WG03.U_BTXI01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXI01
    If GWASEE < WG03.U_BTXK01 Then WG03.U_BTXK01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXK01
    If GWASEE < WG03.U_BTXO01 Then WG03.U_BTXO01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXO01
    If GWASEE < WG03.U_BTXQ01 Then WG03.U_BTXQ01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXQ01
    If GWASEE < WG03.U_BTXR10 Then WG03.U_BTXR10 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXR10
    If GWASEE < WG03.U_BTXS01 Then WG03.U_BTXS01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXS01
    If GWASEE < WG03.U_BTXT01 Then WG03.U_BTXT01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXT01
    If GWASEE < WG03.U_BTXX01 Then WG03.U_BTXX01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXX01
    If GWASEE < WG03.U_BTXY01 Then WG03.U_BTXY01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXY01
    If GWASEE < WG03.U_BTXY02 Then WG03.U_BTXY02 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXY02
    If GWASEE < WG03.U_BTXY03 Then WG03.U_BTXY03 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXY03
    If GWASEE < WG03.U_BTXY20 Then WG03.U_BTXY20 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXY20
    If GWASEE < WG03.U_BTXY21 Then WG03.U_BTXY21 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXY21
    If GWASEE < WG03.U_BTXY22 Then WG03.U_BTXY22 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXY22
    If GWASEE < WG03.U_BTXZ01 Then WG03.U_BTXZ01 = GWASEE
    GWASEE = GWASEE - WG03.U_BTXZ01
    
'/ 1.2) 과세대상급여 ( 총급여액-비과세소득 )
    WG03.U_GWASEE = GWASEE
    WG03.U_BTAX01 = BTAX01
    WG03.U_BTAX02 = BTAX02
    WG03.U_BTAX03 = BTAX03
    WG03.U_BTAX04 = BTAX04
    WG03.U_BTAX05 = BTAX05
    WG03.U_BTAX06 = BTAX06
    WG03.U_BTAX07 = BTAX07
    
 '/ 비과세합계
    If oYM < "201001" Then
        WG03.U_BTXTOT = WG03.U_BTAX01 + WG03.U_BTAX02 + WG03.U_BTAX03 + WG03.U_BTAX04 + WG03.U_BTAX05 + WG03.U_BTAX06 + WG03.U_BTAX07
        WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXG01 + WG03.U_BTXH01 + WG03.U_BTXH05 + WG03.U_BTXH06 + WG03.U_BTXH07 + WG03.U_BTXH08 + WG03.U_BTXH09
        WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXH10 + WG03.U_BTXH11 + WG03.U_BTXH12 + WG03.U_BTXH13 + WG03.U_BTXI01 + WG03.U_BTXK01 + WG03.U_BTXM01
        WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXM02 + WG03.U_BTXM03 + WG03.U_BTXO01 + WG03.U_BTXQ01 + WG03.U_BTXS01 + WG03.U_BTXT01 + WG03.U_BTXX01
        WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXY01 + WG03.U_BTXY02 + WG03.U_BTXY03 + WG03.U_BTXY20 + WG03.U_BTXZ01
    Else
        WG03.U_BTXTOT = WG03.U_BTAX01 + WG03.U_BTAX02 + WG03.U_BTAX03 + WG03.U_BTAX04 + WG03.U_BTAX05 + WG03.U_BTAX06 + WG03.U_BTAX07
        WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXG01 + WG03.U_BTXH01 + WG03.U_BTXH05 + WG03.U_BTXH06 + WG03.U_BTXH07 + WG03.U_BTXH08 + WG03.U_BTXH09
        WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXH10 + WG03.U_BTXH11 + WG03.U_BTXH12 + WG03.U_BTXH13 + WG03.U_BTXI01 + WG03.U_BTXK01 + WG03.U_BTXM01
        WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXM02 + WG03.U_BTXM03 + WG03.U_BTXO01 + WG03.U_BTXQ01 + WG03.U_BTXR10 + WG03.U_BTXS01 + WG03.U_BTXT01
        WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXY01 + WG03.U_BTXY02 + WG03.U_BTXY03 + WG03.U_BTXY21 + WG03.U_BTXY22 + WG03.U_BTXZ01
    End If

'/ 세액계산 /*******************************************************************************/
    '/ 세액계산안함으로 하면 소득세,주민세 계산안함.
    If oDS_PH_PY111A.GetValue("U_TAXCHK", 0) = "N" Then
        '/ 총지급액, 과세대상금액, 공제인원수
        GABGUN = 0
        JUMINN = 0
   
    Else
        '/ 1.3) 세액계산
        If sRecordset.Fields("U_TAXSEL").VALUE = "Y" Then
            INCOME = GWASEE '- IIf(sRecordSet.Fields("U_BX1SEL").Value = "Y", BNSCSU, 0)
            '/ (2010.01.19) 급상여이면 세액기간에+1하던걸 세액기간이 없을때만 +1로변경함
            'If oJOBTYP = "3" Then WG03.U_TAXTRM = WG03.U_TAXTRM + 1
            If oJOBTYP = "3" And WG03.U_TAXTRM = 0 Then WG03.U_TAXTRM = WG03.U_TAXTRM + 1
            If oJOBTYP <> "1" Then  '/ 상여처리
             '/ (2010.01.19) 변동자료만 적용버튼추가되면서 세액기간없을경우 1로 셋팅하도록함.
                If WG03.U_TAXTRM = 0 Then WG03.U_TAXTRM = 1
                '// (2011.10.05) 상여금 급여기본등록의 평균임금참조로 설정시
                If Trim$(PAY_007) = "2" Then
                    WG03.U_AVRPAY = WK_AVRAMT * WG03.U_TAXTRM
                End If
               INCOME = MDC_SetMod.IInt((GWASEE + WG03.U_AVRPAY) / WG03.U_TAXTRM, 1)
            End If
            
         '/ 총지급액, 과세대상금액, 공제인원수
            If Trim$(PAY_001) = "4" Then
                MDC_SetMod.Get_GabGunSe_Table GABGUN, JUMINN, INCOME, WG03.U_TAXCNT, WG03.U_CHLCNT, oYM, INCOME, PAY_001
            Else
                MDC_SetMod.Get_GabGunSe GABGUN, JUMINN, INCOME, WG03.U_TAXCNT, WG03.U_CHLCNT, oYM, INCOME, PAY_001
            End If
        Else
            GABGUN = 0
            JUMINN = 0
        End If
    End If
    '/****************************************************************************************************************************/
    '/ 나이체크
    '/ 고용보험(만18세이상~만65미만가입가능) 64세이상월부터 고용보험제외(200805월기준 1944.5.25일생일경우 5월분 고용보험까지납부,6월분부터 공제제외)
    '/ 국민연금(만18세이상~만60미만가입가능) 60세이상월부터 국민연금제외(200805월기준 1948.5.25일생일경우 5월분 국민연금까지납부,6월분부터 공제제외)
    '/****************************************************************************************************************************/
     KUKSTR = "": KUKEND = ""
     GBHSTR = "": GBHEND = ""
     If Trim$(WG03.U_PERNBR) <> "" Then
         Select Case Mid$(Trim$(WG03.U_PERNBR), 7, 1)
         Case "1", "2", "5", "6"
            KUKSTR = Format$("19" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01", "0000-00-00")
            KUKEND = Format$(DateAdd("yyyy", 60, KUKSTR), "yyyymm")
            GBHSTR = Format$("19" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01", "0000-00-00")
            GBHEND = Format$(DateAdd("yyyy", 64, GBHSTR), "yyyymm")
         Case "9", "0"
            KUKSTR = Format$("18" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01", "0000-00-00")
            KUKEND = Format$(DateAdd("yyyy", 60, KUKSTR), "yyyymm")
            GBHSTR = Format$("18" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01", "0000-00-00")
            GBHEND = Format$(DateAdd("yyyy", 64, GBHSTR), "yyyymm")
         Case "3", "4", "7", "8"
            KUKSTR = "20" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01"
            KUKEND = Format$(DateAdd("y", 60, KUKSTR), "yyyymm")
            GBHSTR = "20" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01"
            GBHEND = Format$(DateAdd("y", 64, GBHSTR), "yyyymm")
         Case Else
            KUKSTR = ""
            KUKEND = ""
            GBHSTR = ""
            GBHEND = ""
         End Select
    End If
  
    '/ 1.4) 고용보험
    '// 고용보험 계산안함으로 하면 제외(단, 체크되어 있으면 무조건 계산)
    '2012.07.11 대우(동우공영:DWGY)의 요청으로 고용보험료의 산출방식이 "고용보험 보수월액기준"이면서 중도입사자일때 일할 계산 처리 추가
    '고용보험 보수월액의 일할 계산(중도입사자) = 고용보험 보수월액 / 월총일수 * 근무일수
    '2012.07.13 대우(동우공영:DWGY)의 요청으로 고용보험료의 산출방식이 "고용보험 보수월액기준"이면서 중도입사자일때 30일로 계산 수정요청
    '고용보험 보수월액의 일할 계산(중도입사자) = 고용보험 보수월액 / 30 * 근무일수
    If oDS_PH_PY111A.GetValue("U_GBHCHK", 0) = "N" Then
            GBHAMT = 0
    Else
        '// 고용보험율 등록 데이터 확인
        sQry = "SELECT TOP 1 U_EMPRAT1, ISNULL(U_EXECHK,'1') AS U_EXECHK FROM [@ZPY118H] "
        sQry = sQry & " WHERE CODE <= '" & oYM & "' ORDER BY CODE DESC"
        oRecordSet.DoQuery sQry
        
        If oRecordSet.RecordCount > 0 Then
            GBHRAT = oRecordSet.Fields("U_EMPRAT1").VALUE       '// 보험율
            Select Case oRecordSet.Fields("U_EXECHK").VALUE     '// 보험료 계산대상 소득금액
                Case "2":   WK_GBHAMT = TOTGBH
                Case "3":   WK_GBHAMT = WG03.U_GWASEE
               '2012.07.11 대우(동우공영:DWGY) 요청
               'Case Else:  WK_GBHAMT = WG03.U_GBHAMT
                Case Else
                    WK_GBHAMT = WG03.U_GBHAMT
            End Select
        Else    '// 고용보험율 등록 데이터가 없으면 내부 기본값으로 계산
            If oYM < "201104" Then
                GBHRAT = 0.45
                If oYM < "201101" Then
                    WK_GBHAMT = TOTGBH          '// 2011년 이전은 보수총액
                Else
                    WK_GBHAMT = WG03.U_GBHAMT   '// 2011년 이후는 보수월액
                End If
            Else
                GBHRAT = 0.55
                WK_GBHAMT = WG03.U_GBHAMT
            End If
        End If

        If sRecordset.Fields("U_GBHSEL").VALUE = "Y" Then
            GBHAMT = MDC_SetMod.IInt(WK_GBHAMT * (GBHRAT / 100), 10)
            
            '/ 고용보험제외(~만65미만가입가능) 64세이상월부터 고용보험제외(200805월기준 1943.5.25일생일경우 4월분 고용보험까지납부,5월분부터 공제제외)
             If Trim$(WG03.U_PERNBR) <> "" Then
                If Trim$(GBHEND) <> "" And Trim$(GBHEND) <= oYM Then
                    GBHAMT = 0
                    REMARK1 = REMARK1 & Space$(1) & WG03.U_MSTCOD & WG03.U_MSTNAM
                End If
             End If
        End If
    End If

    '// 건강보험과 국민연금은 급여+정기, 급상여+정기일때만 자동계산
    '// 2010.08.18 최동권 수정. 일성신약은 상여일 때도 건강/장기요양 보험료 계산(과세금액 기준)
    If (oJOBTYP = "1" And oJOBGBN = "1") Or (oJOBTYP = "3" And oJOBGBN = "1") Or G04_BNSUSE = "Y" Then
        WK_MEDAMT = 0
        '/ 1.5) 의료보험 /'/ 당월 1일이후 입사자 건강보험료 제외
        If Val(WG03.U_MEDAMT) <> 0 Then
            sQry = "SELECT TOP 1 U_EMPRAT, U_FROM, U_TO, U_NJYYMM, U_NJYRAT, U_NJCRAT, ISNULL(U_EXECHK,1) AS U_EXECHK FROM [@ZPY103H] "
            sQry = sQry & " WHERE CODE <= '" & oYM & "' ORDER BY CODE DESC"
            oRecordSet.DoQuery sQry
            If oRecordSet.RecordCount > 0 Then
                U_NJYYMM = oRecordSet.Fields("U_NJYYMM").VALUE  '/ 장기요양보험적용월
                U_NJYRAT = oRecordSet.Fields("U_NJYRAT").VALUE  '/ 장기요양보험적용율
                U_NJCRAT = oRecordSet.Fields("U_NJCRAT").VALUE  '/ 장기요양보험경감보험요율
                MEDCHK = oRecordSet.Fields("U_EXECHK").VALUE    '/ 건강보험산출방식구분(1:보수월액기준,2:월과세금액기준,3:지급총액기준)
                Select Case Trim$(MEDCHK)
                Case "2":   WK_MEDAMT = WG03.U_GWASEE           '/ 과세금액기준
                Case "3":   WK_MEDAMT = WG03.U_TOTPAY           '/ 총지급액기준
                Case Else:  WK_MEDAMT = Val(WG03.U_MEDAMT)
                End Select

                If Val(WK_MEDAMT) < oRecordSet.Fields("U_FROM").VALUE Then
                    MEDAMT = oRecordSet.Fields("U_FROM").VALUE
                ElseIf oRecordSet.Fields("U_TO").VALUE > 0 And Val(WK_MEDAMT) > oRecordSet.Fields("U_TO").VALUE Then
                    MEDAMT = oRecordSet.Fields("U_TO").VALUE
                Else
                    MEDAMT = Val(WK_MEDAMT)
                End If
                MEDAMT = MDC_SetMod.IInt(MEDAMT * Val(oRecordSet.Fields("U_EMPRAT").VALUE) / 100, 10)
                If Val(WG03.U_MEDFRG) <> 0 Then '/ 해외파견경감대상자일경우
                    MEDAMT = MEDAMT - (MDC_SetMod.IInt(MEDAMT * (WG03.U_MEDFRG / 100) + 9, 10)) '/ 경감보험료(10원미만절상)
                End If
            End If
            If oYM = Mid$(Replace(WG03.U_INPDAT, "-", ""), 1, 6) Then
               If oYM & "01" <> Mid$(WG03.U_INPDAT, 1, 8) Then
                    MEDAMT = 0                                 '/ 당월 1일이후 입사자 건강보험료 제외
               End If
            End If
        End If
        '/ 1.5.5) 노인장기요양보험료율
        If MEDAMT <> 0 Then
            sQry = "SELECT TOP 1 U_NJYYMM, U_NJYRAT, U_NJCRAT FROM [@ZPY103H] "
            sQry = sQry & " WHERE U_NJYYMM <= '" & oYM & "' ORDER BY U_NJYYMM DESC"
            oRecordSet.DoQuery sQry
            If oRecordSet.RecordCount > 0 Then
                U_NJYYMM = oRecordSet.Fields("U_NJYYMM").VALUE
                U_NJYRAT = oRecordSet.Fields("U_NJYRAT").VALUE
                U_NJCRAT = oRecordSet.Fields("U_NJCRAT").VALUE
            Else
                U_NJYYMM = ""
                U_NJYRAT = 0
                U_NJCRAT = 0
            End If
        End If
    End If

    If (oJOBTYP = "1" And oJOBGBN = "1") Or (oJOBTYP = "3" And oJOBGBN = "1") Then
        '/ 1.6) 국민연금 /
        If oYM <= "200803" Then
            sQry = "SELECT TOP 1 ISNULL(T0.U_EMPAMT, 0) FROM [@ZPY102L] T0 WHERE T0.Code <= '" & oYM & "' "
            sQry = sQry & " AND T0.U_CODNBR ='" & sRecordset.Fields("U_KUKGRD").VALUE & "'"
            sQry = sQry & " ORDER BY CODE DESC"
            oRecordSet.DoQuery sQry
            If oRecordSet.RecordCount > 0 Then
                KUKAMT = Val(oRecordSet.Fields(0).VALUE)
            End If
        Else
            '/ 국민연금보수월액이랑 등급없으면 연금공제 제외
            If Trim$(sRecordset.Fields("U_KUKGRD").VALUE) = "" And Val(WG03.U_KUKAMT) = 0 Then
                KUKAMT = 0
            Else
                sQry = "SELECT TOP 1 U_EMPRAT, U_FROM, U_TO FROM [@ZPY102H] "
                sQry = sQry & " WHERE CODE <= '" & oYM & "' ORDER BY CODE DESC"
                oRecordSet.DoQuery sQry
                If oRecordSet.RecordCount > 0 Then
                    If Val(WG03.U_KUKAMT) = 0 Then
                        KUKAMT = 0
                    ElseIf Val(WG03.U_KUKAMT) < oRecordSet.Fields("U_FROM").VALUE Then
                        KUKAMT = oRecordSet.Fields("U_FROM").VALUE
                    ElseIf oRecordSet.Fields("U_TO").VALUE > 0 And Val(WG03.U_KUKAMT) > oRecordSet.Fields("U_TO").VALUE Then
                        KUKAMT = oRecordSet.Fields("U_TO").VALUE
                    Else
                        KUKAMT = Val(WG03.U_KUKAMT)
                    End If
                    KUKAMT = MDC_SetMod.IInt(KUKAMT * Val(oRecordSet.Fields("U_EMPRAT").VALUE) / 100, 10)
                End If
            End If
        End If
        '/ 국민연금(만18세이상~만60미만가입가능) 60세이상월부터 국민연금제외(200805월기준 1948.5.25일생일경우 5월분 국민연금까지납부,6월분부터 공제제외)
        '/ 예외:급여기본등록의 국민연금 연장에 체크되어있으신분은 나이제한공제제외대상에서 빠짐(국민연금공제)
         If Trim$(WG03.U_PERNBR) <> "" Then
            If Trim$(KUKEND) <> "" And Trim$(KUKEND) < oYM And sRecordset.Fields("U_KUKOVR").VALUE = "N" Then
                KUKAMT = 0
                REMARK2 = REMARK2 & Space$(1) & WG03.U_MSTCOD & WG03.U_MSTNAM
            End If
         End If
        If oYM = Mid$(Replace(WG03.U_INPDAT, "-", ""), 1, 6) Then
           If oYM & "01" <> Mid$(WG03.U_INPDAT, 1, 8) Then
                KUKAMT = 0                                 '/ 당월 1일이후 입사자 국민연금 제외
           End If
        End If
   
    End If
  
    '// 소득세, 주민세, 고용보험은 계산여부를 별도로 선택가능하므로
    '// 수당/공제 계산식 반영제외와 무관하게 처리함
    For iCol = 1 To 18
        Select Case WK_G.GONCOD(iCol)
        Case "G01" '/ 소득세
            If oJOBTYP = "1" Then
                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + MDC_SetMod.IInt(GABGUN, 10)
            Else
                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + (MDC_SetMod.IInt((GABGUN * WG03.U_TAXTRM) - WG03.U_NABTAX, 10))
            End If
            If WG03.U_GONAMT(iCol) < 0 Then WG03.U_GONAMT(iCol) = 0
        Case "G02" '/ 주민세
            If oJOBTYP = "1" Then
                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + MDC_SetMod.IInt(JUMINN, 10)
            Else
                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + (MDC_SetMod.IInt((MDC_SetMod.IInt((GABGUN * WG03.U_TAXTRM) - WG03.U_NABTAX, 10)) * 0.1, 10))
            End If
            If WG03.U_GONAMT(iCol) < 0 Then WG03.U_GONAMT(iCol) = 0
        Case "G05"  '/ 고용보험
            WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + GBHAMT
        Case "G04"  '/ 건강보험
            If G04_BNSUSE = "Y" And oDS_PH_PY111A.GetValue("U_CHGCHK", 0) = "Y" Then
                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + MEDAMT
            End If
        Case "G10"  '/ 장기요양보험
            If G04_BNSUSE = "Y" And oDS_PH_PY111A.GetValue("U_CHGCHK", 0) = "Y" Then
                If U_NJYYMM <= oYM Then
                    NJCMED = MDC_SetMod.IInt(MEDAMT * (U_NJYRAT / 100), 10) '/ (10원미만 절사)
                    If Trim$(WG03.U_NJCGBN) = "Y" Then '/ 경감대상자일경우
                        NJCMED = NJCMED - (MDC_SetMod.IInt(NJCMED * (U_NJCRAT / 100) + 9, 10)) '/ 경감보험료(10원미만절상)
                    End If
                    WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + NJCMED       '/ 장기요양보험료
                End If
            End If
        End Select
    Next iCol
  
'// 2. 공식가지고 계산하기-고정공제
    If oDS_PH_PY111A.GetValue("U_CHGCHK", 0) = "N" Then
        For iCol = 1 To 18
            Select Case WK_G.GONCOD(iCol)
            Case "G03" '/ 국민연금
                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + KUKAMT
            Case "G04" '/ 건강보험
                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + MEDAMT
            Case "G10" '/ 노인장기요양보험
                If U_NJYYMM <= oYM Then
                    NJCMED = MDC_SetMod.IInt(MEDAMT * (U_NJYRAT / 100), 10) '/ (10원미만 절사)
                    If Trim$(WG03.U_NJCGBN) = "Y" Then '/ 경감대상자일경우
                        NJCMED = NJCMED - (MDC_SetMod.IInt(NJCMED * (U_NJCRAT / 100) + 9, 10)) '/ 경감보험료(10원미만절상)
                    End If
                    WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + NJCMED       '/ 장기요양보험료
                End If
            Case Else
                If oJOBTYP = "2" And WK_G.BNSUSE(iCol) = "N" Then  '/ 상여만일경우 공제체크안된것 제외
                    WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + 0
                Else
                    If Trim$(WK_G.GONCOD(iCol)) <> "" And Trim$(WK_G.GONSIL(iCol)) <> "" Then    '/ 공식이있으면 공식계산
                    '/ 2.1. 공식-계산결과값가져오는거면..
                        SuSilStr = Change_GOSIL(WK_G.GONSIL(iCol))   '/ 계산된값일 경우
                        For kCol = 1 To 18
                            SuSilStr = Replace(SuSilStr, "#" & Format$(kCol, "00"), WG03.U_GONAMT(kCol))    '/ 계산된 공제값
                        Next kCol
                        For kCol = 1 To 24
                            SuSilStr = Replace(SuSilStr, "#C" & Format$(kCol, "00"), WG03.U_CSUAMT(kCol))   '/ 계산된 수당값
                        Next kCol
                        SuSilStr = Replace(SuSilStr, "X15", WG03.U_TOTPAY)  '/ X15:지급총액
                    '/ 2.2. 공식계산하기
                        Tmp_CSUAMT = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
                    '/ 2.3. 정답가져오기
                        WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + MDC_SetMod.RInt(Tmp_CSUAMT, WK_G.RODLEN(iCol), WK_G.ROUNDT(iCol))
                    End If
                End If
            End Select
        Next iCol
    End If
'/ 급상여변동자료(기간) -변동금액으로 대체/~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    sQry = "SELECT  T0.U_CSUTYP, T0.U_CSUCOD, T0.U_CSUAMT, T0.U_MSTCOD"
    sQry = sQry & " FROM [@ZPY312L] T0 INNER JOIN [@ZPY312H] T1 ON T0.DocEntry = T1.DocEntry"
    sQry = sQry & " WHERE T1.U_JOBTYP = '" & oJOBTYP & "'"
    sQry = sQry & " AND   T1.U_JOBGBN = '" & oJOBGBN & "'"
    sQry = sQry & " AND   T0.U_CSUTYP = '2'"    '공제항목만
    sQry = sQry & " AND   T0.U_MSTCOD = '" & Trim$(WG03.U_MSTCOD) & "'"
    sQry = sQry & " AND   " & "'" & oYM & "' BETWEEN T0.U_STRYMM AND T0.U_ENDYMM"
    sQry = sQry & " ORDER BY T0.U_MSTCOD, T1.U_YM DESC, T0.U_STRYMM DESC"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        For iCol = 1 To 18
            If oRecordSet.Fields("U_CSUCOD").VALUE = WK_G.GONCOD(iCol) Then
                WG03.U_GONAMT(iCol) = oRecordSet.Fields("U_CSUAMT").VALUE
            End If
        Next iCol
        oRecordSet.MoveNext
    Loop

'/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'/2. 총공제액
'/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    For iCol = 1 To 18
        WG03.U_TOTGON = WG03.U_TOTGON + Val(WG03.U_GONAMT(iCol))
    Next iCol
'/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'/3. 실지급액
'/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    WG03.U_SILJIG = MDC_SetMod.IInt(WG03.U_TOTPAY - WG03.U_TOTGON, 1)
   'If WG03.U_SILJIG < 1000 Then WG03.U_SILJIG = 0
    
    If WG03.U_SILJIG < 0 Then  '// 실지급액이 마이너스인 사원 확인요청
        REMARK3 = REMARK3 & Space$(1) & WG03.U_MSTCOD & WG03.U_MSTNAM
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "PH_PY111A_TAX Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Function PH_PY111_Save() As String
On Error GoTo Error_Message
    Dim oRecordSet    As SAPbobsCOM.Recordset
    Dim sQry          As String
    Dim DocNum        As Long
    Dim iCol          As Integer
    Dim S_TYPNAM      As String
    Dim S_GBNNAM      As String
    Dim ErrNum        As Integer
    Dim DfSeries      As Long
    
    DfSeries = Get_Series_No("UTD")  '/ 2009.01월에 추가
          
    S_TYPNAM = oForm.Items("JOBTYP").Specific.Selected.Description
    S_GBNNAM = oForm.Items("JOBGBN").Specific.Selected.Description
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'/ 1) Search
    sQry = "SELECT ISNULL(U_ENDCHK, 'N') AS ENDCHK FROM [@PH_PY111A]"
    sQry = sQry & " WHERE U_JOBGBN =  '" & oJOBGBN & "'"
    sQry = sQry & " AND U_JOBTYP = '" & oJOBTYP & "'"
    sQry = sQry & " AND U_YM = '" & oYM & "'"
    sQry = sQry & " AND U_MSTCOD = '" & Trim$(WG03.U_MSTCOD) & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
'/ 2) Insert
      ' 2.1)Autokey
        DocNum = MDC_SetMod.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY111'", "")

        If DocNum <= 0 Then DocNum = 1
      ' 2.2)삽입
        sQry = "INSERT INTO [@PH_PY111A] (DocEntry, DocNum, Period, Instance,  Series, Handwrtten, Canceled, Object, UserSign, Transfered, Status, CreateDate, CreateTime, DataSource,"
        sQry = sQry & " U_YM, U_JOBTYP, U_JOBGBN, U_MSTCOD, U_RETCHK,U_ENDCHK, U_CSUCOD, U_GONCOD,"
        sQry = sQry & " U_JIGBIL, U_MSTNAM, U_EmpID, U_CLTCOD, U_MSTBRK,  U_MSTDPT, U_MSTSTP, U_CLTNAM, U_BRKNAM, U_DPTNAM, U_STPNAM,"
        sQry = sQry & "  U_PAYTYP, U_JOBTRG, U_JIGCOD, U_HOBONG, U_STDAMT, U_INPDAT, U_OUTDAT, U_TAXCNT, U_BUYN20, U_DAYAMT, U_BASAMT,"
        sQry = sQry & "  U_CSUD01, U_CSUD02, U_CSUD03, U_CSUD04, U_CSUD05, U_CSUD06, U_CSUD07, U_CSUD08,"
        sQry = sQry & "  U_CSUD09, U_CSUD10, U_CSUD11, U_CSUD12, U_CSUD13, U_CSUD14, U_CSUD15, U_CSUD16,"
        sQry = sQry & "  U_CSUD17, U_CSUD18, U_CSUD19, U_CSUD20, U_CSUD21, U_CSUD22, U_CSUD23, U_CSUD24,"
        sQry = sQry & "  U_GWASEE, U_BTAX01, U_BTAX02, U_BTAX03, U_BTAX04, U_BTAX05, U_BTAX06, U_BTAX07,"
        sQry = sQry & "  U_BTXG01, U_BTXH01, U_BTXH05, U_BTXH06, U_BTXH07, U_BTXH08, U_BTXH09, U_BTXH10,"
        sQry = sQry & "  U_BTXH11, U_BTXH12, U_BTXH13, U_BTXI01, U_BTXK01, U_BTXM01, U_BTXM02, U_BTXM03,"
        sQry = sQry & "  U_BTXO01, U_BTXQ01, U_BTXS01, U_BTXT01, U_BTXX01, U_BTXY01, U_BTXY02, U_BTXY03,"
        sQry = sQry & "  U_BTXY20, U_BTXZ01, U_BTXTOT,"
        sQry = sQry & "  U_TOTPAY, U_TOTGON, U_SILJIG,"
        sQry = sQry & "  U_GONG01, U_GONG02, U_GONG03, U_GONG04, U_GONG05, U_GONG06, U_GONG07, U_GONG08,"
        sQry = sQry & "  U_GONG09, U_GONG10, U_GONG11, U_GONG12, U_GONG13, U_GONG14, U_GONG15, U_GONG16, U_GONG17, U_GONG18,   "
        sQry = sQry & "  U_AVRPAY, U_NABTAX, U_BNSRAT, U_APPRAT, U_GNSYER,"
        sQry = sQry & "  U_GNSMON , U_TAXTRM, U_BONUSS, U_TYPNAM, U_GBNNAM "
        sQry = sQry & " ) VALUES(" & DocNum & ", " & DocNum & "," & "15, 0, '" & DfSeries & "', 'N','N', 'PH_PY111', " & oCompany.UserSignature & ", 'N', 'O', '" & Format$(Now, "yyyy-mm-dd") & "', '" & Format$(Now, "hhmm") & "', 'I',"
        sQry = sQry & "'" & oYM & "', '" & oJOBTYP & "','" & oJOBGBN & "', '" & Trim$(WG03.U_MSTCOD) & "', '" & Trim$(oDS_PH_PY111A.GetValue("U_RETCHK", 0)) & "', 'N', '" & U_CSUCOD & "', '" & U_GONCOD & "',"
        sQry = sQry & "'" & Trim$(oJIGBIL) & "', N'" & Trim$(WG03.U_MSTNAM) & "', '" & Trim$(WG03.U_EmpID) & "', '" & Trim$(WG03.U_CLTCOD) & "', '" & Trim$(WG03.U_MSTBRK) & "', '" & Trim$(WG03.U_MSTDPT) & "', '" & Trim$(WG03.U_MSTSTP) & "', N'" & Trim$(WG03.U_CLTNAM) & "', N'" & Trim$(WG03.U_BRKNAM) & "', N'" & Trim$(WG03.U_DPTNAM) & "' , N'" & Trim$(WG03.U_STPNAM) & "',"
        sQry = sQry & "'" & Trim$(WG03.U_PAYTYP) & "', '" & Trim$(WG03.U_JOBTRG) & "', '" & Trim$(WG03.U_JIGCOD) & "', '" & Trim$(WG03.U_HOBONG) & "', '" & Trim$(WG03.U_STDAMT) & "', '" & Format$(WG03.U_INPDAT, "0000-00-00") & "' , "
        sQry = sQry & " CASE WHEN '" & Trim$(WG03.U_OUTDAT) & "'='' THEN NULL ELSE '" & Format$(WG03.U_OUTDAT, "0000-00-00") & "' END , '" & Val(WG03.U_TAXCNT) & "','" & Val(WG03.U_CHLCNT) & "','" & Val(WG03.U_DAYAMT) & "', '" & Val(WG03.U_BASAMT) & "',"
        sQry = sQry & "'" & Val(WG03.U_CSUAMT(1)) & "', '" & Val(WG03.U_CSUAMT(2)) & "', '" & Val(WG03.U_CSUAMT(3)) & "', '" & Val(WG03.U_CSUAMT(4)) & "', '" & Val(WG03.U_CSUAMT(5)) & "' , '" & Val(WG03.U_CSUAMT(6)) & "', '" & Val(WG03.U_CSUAMT(7)) & "', '" & Val(WG03.U_CSUAMT(8)) & "',"
        sQry = sQry & "'" & Val(WG03.U_CSUAMT(9)) & "', '" & Val(WG03.U_CSUAMT(10)) & "', '" & Val(WG03.U_CSUAMT(11)) & "', '" & Val(WG03.U_CSUAMT(12)) & "', '" & Val(WG03.U_CSUAMT(13)) & "' , '" & Val(WG03.U_CSUAMT(14)) & "', '" & Val(WG03.U_CSUAMT(15)) & "', '" & Val(WG03.U_CSUAMT(16)) & "',"
        sQry = sQry & "'" & Val(WG03.U_CSUAMT(17)) & "', '" & Val(WG03.U_CSUAMT(18)) & "', '" & Val(WG03.U_CSUAMT(19)) & "', '" & Val(WG03.U_CSUAMT(20)) & "', '" & Val(WG03.U_CSUAMT(21)) & "' , '" & Val(WG03.U_CSUAMT(22)) & "', '" & Val(WG03.U_CSUAMT(23)) & "', '" & Val(WG03.U_CSUAMT(24)) & "',"
        sQry = sQry & "'" & Val(WG03.U_GWASEE) & "', '" & Val(WG03.U_BTAX01) & "', '" & Val(WG03.U_BTAX02) & "', '" & Val(WG03.U_BTAX03) & "', '" & Val(WG03.U_BTAX04) & "', '" & Val(WG03.U_BTAX05) & "', '" & Val(WG03.U_BTAX06) & "', '" & Val(WG03.U_BTAX07) & "', "
        sQry = sQry & "'" & Val(WG03.U_BTXG01) & "', '" & Val(WG03.U_BTXH01) & "', '" & Val(WG03.U_BTXH05) & "', '" & Val(WG03.U_BTXH06) & "', '" & Val(WG03.U_BTXH07) & "', "
        sQry = sQry & "'" & Val(WG03.U_BTXH08) & "', '" & Val(WG03.U_BTXH09) & "', '" & Val(WG03.U_BTXH10) & "', '" & Val(WG03.U_BTXH11) & "', '" & Val(WG03.U_BTXH12) & "', "
        sQry = sQry & "'" & Val(WG03.U_BTXH13) & "', '" & Val(WG03.U_BTXI01) & "', '" & Val(WG03.U_BTXK01) & "', '" & Val(WG03.U_BTXM01) & "', '" & Val(WG03.U_BTXM02) & "', "
        sQry = sQry & "'" & Val(WG03.U_BTXM03) & "', '" & Val(WG03.U_BTXO01) & "', '" & Val(WG03.U_BTXQ01) & "', '" & Val(WG03.U_BTXS01) & "', '" & Val(WG03.U_BTXT01) & "', "
        sQry = sQry & "'" & Val(WG03.U_BTXX01) & "', '" & Val(WG03.U_BTXY01) + Val(WG03.U_BTXR10) & "', '" & Val(WG03.U_BTXY02) & "', '" & Val(WG03.U_BTXY03) & "', '" & Val(WG03.U_BTXY20) + Val(WG03.U_BTXY22) & "',"
        sQry = sQry & "'" & Val(WG03.U_BTXZ01) & "', '" & Val(WG03.U_BTXTOT) & "', "
        sQry = sQry & "'" & Val(WG03.U_TOTPAY) & "' , '" & Val(WG03.U_TOTGON) & "', '" & Val(WG03.U_SILJIG) & "', "
        sQry = sQry & "'" & Val(WG03.U_GONAMT(1)) & "', '" & Val(WG03.U_GONAMT(2)) & "', '" & Val(WG03.U_GONAMT(3)) & "', '" & Val(WG03.U_GONAMT(4)) & "', '" & Val(WG03.U_GONAMT(5)) & "' , '" & Val(WG03.U_GONAMT(6)) & "', '" & Val(WG03.U_GONAMT(7)) & "', '" & Val(WG03.U_GONAMT(8)) & "',"
        sQry = sQry & "'" & Val(WG03.U_GONAMT(9)) & "', '" & Val(WG03.U_GONAMT(10)) & "', '" & Val(WG03.U_GONAMT(11)) & "', '" & Val(WG03.U_GONAMT(12)) & "','" & Val(WG03.U_GONAMT(13)) & "','" & Val(WG03.U_GONAMT(14)) & "', '" & Val(WG03.U_GONAMT(15)) & "', '" & Val(WG03.U_GONAMT(16)) & "',"
        sQry = sQry & "'" & Val(WG03.U_GONAMT(17)) & "', '" & Val(WG03.U_GONAMT(18)) & "', "
        sQry = sQry & "'" & Val(WG03.U_AVRPAY) & "', '" & Val(WG03.U_NABTAX) & "' , '" & Val(WG03.U_BNSRAT) & "', '" & Val(WG03.U_APPRAT) & "', '" & Val(WG03.U_GNSYER) & "',"
        sQry = sQry & "'" & Val(WG03.U_GNSMON) & "', '" & Val(WG03.U_TAXTRM) & "', '" & Val(WG03.U_BONUSS) & "', N'" & S_TYPNAM & "', N'" & S_GBNNAM & "'"
        sQry = sQry & "  )"
        oRecordSet.DoQuery sQry
      ' 2.3)Autokey증가
        DocNum = DocNum + 1
        sQry = "UPDATE ONNM SET AutoKey = " & DocNum & " WHERE  ObjectCode='PH_PY111'"
        oRecordSet.DoQuery sQry
    Else
'/ 3) Update
        If Trim$(oRecordSet.Fields("ENDCHK").VALUE) = "Y" Then
        ' 3.1)잠긴 자료
            ErrNum = 1
            G_ChkCnt = G_ChkCnt + 1
            GoTo Error_Message
        End If
        ' 3.2)갱신
        sQry = " UPDATE [@PH_PY111A]         "
        sQry = sQry & " SET UpdateDate = '" & Format$(Now, "yyyy-mm-dd") & "'"
        sQry = sQry & " ,   UpdateTime = '" & Format$(Now, "hhmm") & "'"
        sQry = sQry & " ,   U_MSTNAM = N'" & Trim$(WG03.U_MSTNAM) & "'"
        sQry = sQry & " ,   UserSign = " & oCompany.UserSignature
        sQry = sQry & " ,   U_EmpID = '" & Trim$(WG03.U_EmpID) & "'"
        sQry = sQry & " ,   U_ENDCHK = 'N'"  '잠금여부
        sQry = sQry & " ,   U_CSUCOD = '" & U_CSUCOD & "'"
        sQry = sQry & " ,   U_GONCOD = '" & U_GONCOD & "'"
        sQry = sQry & " ,   U_RETCHK = '" & oDS_PH_PY111A.GetValue("U_RETCHK", 0) & "'" '퇴직금포함
        sQry = sQry & " ,   U_JIGBIL = '" & Trim$(oJIGBIL) & "'"
        sQry = sQry & " ,   U_CLTCOD = '" & Trim$(WG03.U_CLTCOD) & "'"
        sQry = sQry & " ,   U_MSTBRK = '" & Trim$(WG03.U_MSTBRK) & "'"
        sQry = sQry & " ,   U_MSTDPT = '" & Trim$(WG03.U_MSTDPT) & "'"
        sQry = sQry & " ,   U_MSTSTP = '" & Trim$(WG03.U_MSTSTP) & "'"
        sQry = sQry & " ,   U_CLTNAM = N'" & Trim$(WG03.U_CLTNAM) & "'"
        sQry = sQry & " ,   U_BRKNAM = N'" & Trim$(WG03.U_BRKNAM) & "'"
        sQry = sQry & " ,   U_DPTNAM = N'" & Trim$(WG03.U_DPTNAM) & "'"
        sQry = sQry & " ,   U_STPNAM = N'" & Trim$(WG03.U_STPNAM) & "'"
        sQry = sQry & " ,   U_PAYTYP = '" & Trim$(WG03.U_PAYTYP) & "'"
        sQry = sQry & " ,   U_JOBTRG = '" & Trim$(WG03.U_JOBTRG) & "'"
        sQry = sQry & " ,   U_JIGCOD = '" & Trim$(WG03.U_JIGCOD) & "'"
        sQry = sQry & " ,   U_HOBONG = '" & Trim$(WG03.U_HOBONG) & "'"
        sQry = sQry & " ,   U_STDAMT = " & WG03.U_STDAMT
        sQry = sQry & " ,   U_INPDAT = CASE WHEN '" & Trim$(WG03.U_INPDAT) & "'='' THEN NULL ELSE '" & Format$(WG03.U_INPDAT, "0000-00-00") & "' END "
        sQry = sQry & " ,   U_OUTDAT = CASE WHEN '" & Trim$(WG03.U_OUTDAT) & "'='' THEN NULL ELSE '" & Format$(WG03.U_OUTDAT, "0000-00-00") & "' END "
        sQry = sQry & " ,   U_TAXCNT = '" & WG03.U_TAXCNT & "'"
        sQry = sQry & " ,   U_BUYN20 = '" & WG03.U_CHLCNT & "'"
        sQry = sQry & " ,   U_DAYAMT = '" & WG03.U_DAYAMT & "'"
        sQry = sQry & " ,   U_BASAMT = '" & WG03.U_BASAMT & "'"
        For iCol = 1 To 24
        sQry = sQry & " ,   U_CSUD" & Format$(iCol, "00") & " = " & WG03.U_CSUAMT(iCol)
        Next iCol
        sQry = sQry & " ,   U_GWASEE = '" & WG03.U_GWASEE & "'"
        sQry = sQry & " ,   U_BTAX01 = '" & WG03.U_BTAX01 & "'"
        sQry = sQry & " ,   U_BTAX02 = '" & WG03.U_BTAX02 & "'"
        sQry = sQry & " ,   U_BTAX03 = '" & WG03.U_BTAX03 & "'"
        sQry = sQry & " ,   U_BTAX04 = '" & WG03.U_BTAX04 & "'"
        sQry = sQry & " ,   U_BTAX05 = '" & WG03.U_BTAX05 & "'"
        sQry = sQry & " ,   U_BTAX06 = '" & WG03.U_BTAX06 & "'"
        sQry = sQry & " ,   U_BTAX07 = '" & WG03.U_BTAX07 & "'"
        sQry = sQry & " ,   U_BTXG01 = '" & WG03.U_BTXG01 & "'"
        sQry = sQry & " ,   U_BTXH01 = '" & WG03.U_BTXH01 & "'"
        sQry = sQry & " ,   U_BTXH05 = '" & WG03.U_BTXH05 & "'"
        sQry = sQry & " ,   U_BTXH06 = '" & WG03.U_BTXH06 & "'"
        sQry = sQry & " ,   U_BTXH07 = '" & WG03.U_BTXH07 & "'"
        sQry = sQry & " ,   U_BTXH08 = '" & WG03.U_BTXH08 & "'"
        sQry = sQry & " ,   U_BTXH09 = '" & WG03.U_BTXH09 & "'"
        sQry = sQry & " ,   U_BTXH10 = '" & WG03.U_BTXH10 & "'"
        sQry = sQry & " ,   U_BTXH11 = '" & WG03.U_BTXH11 & "'"
        sQry = sQry & " ,   U_BTXH12 = '" & WG03.U_BTXH12 & "'"
        sQry = sQry & " ,   U_BTXH13 = '" & WG03.U_BTXH13 & "'"
        sQry = sQry & " ,   U_BTXI01 = '" & WG03.U_BTXI01 & "'"
        sQry = sQry & " ,   U_BTXK01 = '" & WG03.U_BTXK01 & "'"
        sQry = sQry & " ,   U_BTXM01 = '" & WG03.U_BTXM01 & "'"
        sQry = sQry & " ,   U_BTXM02 = '" & WG03.U_BTXM02 & "'"
        sQry = sQry & " ,   U_BTXM03 = '" & WG03.U_BTXM03 & "'"
        sQry = sQry & " ,   U_BTXO01 = '" & WG03.U_BTXO01 & "'"
        sQry = sQry & " ,   U_BTXQ01 = '" & WG03.U_BTXQ01 & "'"
        sQry = sQry & " ,   U_BTXS01 = '" & WG03.U_BTXS01 & "'"
        sQry = sQry & " ,   U_BTXT01 = '" & WG03.U_BTXT01 & "'"
        sQry = sQry & " ,   U_BTXX01 = '" & WG03.U_BTXX01 & "'"
        sQry = sQry & " ,   U_BTXY01 = '" & WG03.U_BTXY01 + WG03.U_BTXR10 & "'"
        sQry = sQry & " ,   U_BTXY02 = '" & WG03.U_BTXY02 & "'"
        sQry = sQry & " ,   U_BTXY03 = '" & WG03.U_BTXY03 & "'"
        sQry = sQry & " ,   U_BTXY20 = '" & WG03.U_BTXY20 + WG03.U_BTXY22 & "'"
        sQry = sQry & " ,   U_BTXZ01 = '" & WG03.U_BTXZ01 & "'"
        sQry = sQry & " ,   U_BTXTOT = '" & WG03.U_BTXTOT & "'"
        sQry = sQry & " ,   U_TOTPAY = '" & WG03.U_TOTPAY & "'"
        For iCol = 1 To 18
        sQry = sQry & " ,   U_GONG" & Format$(iCol, "00") & " =  '" & WG03.U_GONAMT(iCol) & "'"
        Next iCol
        sQry = sQry & " ,   U_TOTGON = '" & WG03.U_TOTGON & "'"
        sQry = sQry & " ,   U_SILJIG = '" & WG03.U_SILJIG & "'"
        sQry = sQry & " ,   U_AVRPAY = '" & WG03.U_AVRPAY & "'"
        sQry = sQry & " ,   U_NABTAX = '" & WG03.U_NABTAX & "'"
        sQry = sQry & " ,   U_BNSRAT = '" & WG03.U_BNSRAT & "'"
        sQry = sQry & " ,   U_APPRAT = '" & WG03.U_APPRAT & "'"
        sQry = sQry & " ,   U_GNSYER = '" & WG03.U_GNSYER & "'"
        sQry = sQry & " ,   U_GNSMON = '" & WG03.U_GNSMON & "'"
        sQry = sQry & " ,   U_TAXTRM = '" & WG03.U_TAXTRM & "'"
        sQry = sQry & " ,   U_BONUSS = '" & WG03.U_BONUSS & "'"
        sQry = sQry & " ,   U_TYPNAM = N'" & S_TYPNAM & "'"
        sQry = sQry & " ,   U_GBNNAM = N'" & S_GBNNAM & "'"
        sQry = sQry & " WHERE U_YM = '" & oYM & "'"
        sQry = sQry & " AND   U_JOBTYP = '" & oJOBTYP & "'"
        sQry = sQry & " AND   U_JOBGBN = '" & oJOBGBN & "'"
        sQry = sQry & " AND   U_MSTCOD = '" & Trim$(WG03.U_MSTCOD) & "'"
        oRecordSet.DoQuery sQry

    End If
    G_PayCnt = G_PayCnt + 1
    PH_PY111_Save = "Y"
    Set oRecordSet = Nothing
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        PH_PY111_Save = " 잠겨져 있습니다."
    Else
        PH_PY111_Save = Err.Description
    End If
End Function



'Private Sub PAY_Tax1()
'On Error GoTo error_Message
'    Dim oRecordSet    As SAPbobsCOM.Recordset
'    Dim sQry          As String
'    Dim iCol          As Integer
'    Dim BnsUsed       As Boolean
'    Dim GABGUN As Double, JUMINN As Double, GBHAMT As Double, MEDAMT As Double, KUKAMT As Double
'    Dim INCOME As Double, GWASEE As Double, BTAX01 As Double, BTAX02 As Double, BTAX03 As Double, BTAX04 As Double, BTAX05 As Double, BTAX06 As Double, BTAX07 As Double
'    Dim MONPAY As Double, BT1AMT As Double, FODCSU As Double, CARCSU As Double, FRNCSU As Double, CHLCSU As Double, YGUCSU As Double, GITBTX As Double, GITBTX_N As Double
'    Dim BNSCSU As Double, BT1KUM As Double, TOTGBH As Double, TOTMED As Double
'    Dim KUKEND       As String, KUKSTR       As String
'    Dim U_NJYYMM As String, U_NJYRAT As Double, U_NJCRAT As Double, NJCMED As Double
'    Dim BASAMT As Double
'    Dim WK_YBTXAM(1 To 7) As Double
'    GABGUN = 0: JUMINN = 0: GBHAMT = 0: MEDAMT = 0: KUKAMT = 0: NJCMED = 0
'    INCOME = 0: GWASEE = 0: BTAX01 = 0: BTAX02 = 0: BTAX03 = 0: BTAX04 = 0: BTAX05 = 0: BTAX06 = 0: BTAX07 = 0:
'    MONPAY = 0: BT1AMT = 0: FODCSU = 0: CARCSU = 0: TOTGBH = 0: CHLCSU = 0: YGUCSU = 0: TOTMED = 0
'    BNSCSU = 0: GITBTX = 0: GITBTX_N = 0
'    BASAMT = 0  '/ 통상월급
'    U_NJYYMM = 0: U_NJYRAT = 0: U_NJCRAT = 0
'    BnsUsed = False
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
' '/ 급여총액 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
''/ 월정체크, 한도액, 생산비과대상자,부양가족공제,고용보험계산자
'    For iCol = 1 To 24
'        If WK_C.CSUCOD(iCol) = "A04" Then
'            If ods_PH_PY111A.getvalue("U_CHGCHK",0) =  "N" Then
'                WG03.U_CSUAMT(iCol) = WG03.U_BONUSS
'                BNSCSU = BNSCSU + Val(WG03.U_CSUAMT(iCol))
'                BnsUsed = True
'            Else
'                WG03.U_BONUSS = WG03.U_CSUAMT(iCol)
'            End If
'        End If
'        If Trim$(WK_C.GWATYP(iCol)) = "" Then WK_C.GWATYP(iCol) = "1"
'        Select Case WK_C.GWATYP(iCol)
'        Case "1"   '/ 과세
'            If WK_C.CSUKUM(iCol) > 0 And WG03.U_CSUAMT(iCol) > WK_C.CSUKUM(iCol) Then
'               WG03.U_CSUAMT(iCol) = Val(WK_C.CSUKUM(iCol))
'            End If
'        Case "2"   '/ 식대보조
'            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
'                FODCSU = 0
'            Else
'                FODCSU = Val(WG03.U_CSUAMT(iCol))
'            End If
'
'            If WK_C.CSUKUM(iCol) > 0 And FODCSU > WK_C.CSUKUM(iCol) Then
'                BTAX02 = BTAX02 + WK_C.CSUKUM(iCol)
'            Else
'                BTAX02 = BTAX02 + FODCSU
'            End If
'        Case "3"   '/ 차량보조
'            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
'                CARCSU = 0
'            Else
'                CARCSU = Val(WG03.U_CSUAMT(iCol))
'            End If
'
'            If WK_C.CSUKUM(iCol) > 0 And CARCSU > WK_C.CSUKUM(iCol) Then
'                BTAX02 = BTAX02 + WK_C.CSUKUM(iCol)
'            Else
'                BTAX02 = BTAX02 + CARCSU
'            End If
'        Case "4"   '/ 생산비과
'            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
'                BT1AMT = BT1AMT + 0
'            Else
'                BT1AMT = BT1AMT + Val(WG03.U_CSUAMT(iCol))
'            End If
'            If BT1KUM < WK_C.CSUKUM(iCol) Then BT1KUM = WK_C.CSUKUM(iCol)    '비과세4-마지막줄의 한도액만
'        Case "5"   '/ 국외비과세
'            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
'                FRNCSU = FRNCSU + 0
'            Else
'                FRNCSU = FRNCSU + Val(WG03.U_CSUAMT(iCol))
'            End If
'
'            If WK_C.CSUKUM(iCol) > 0 And FRNCSU > WK_C.CSUKUM(iCol) Then
'                BTAX03 = BTAX03 + WK_C.CSUKUM(iCol)
'            Else
'                BTAX03 = BTAX03 + FRNCSU
'            End If
'        Case "6"   '/ 연구수당
'            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
'                YGUCSU = YGUCSU + 0
'            Else
'                YGUCSU = YGUCSU + Val(WG03.U_CSUAMT(iCol))
'            End If
'            If WK_C.CSUKUM(iCol) > 0 And YGUCSU > WK_C.CSUKUM(iCol) Then
'                BTAX05 = BTAX05 + WK_C.CSUKUM(iCol)
'            Else
'                BTAX05 = BTAX05 + YGUCSU
'            End If
'        Case "7"   '/ 보육수당
'            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
'                CHLCSU = CHLCSU + 0
'            Else
'                CHLCSU = CHLCSU + Val(WG03.U_CSUAMT(iCol))
'            End If
'            If WK_C.CSUKUM(iCol) > 0 And CHLCSU > WK_C.CSUKUM(iCol) Then
'                BTAX06 = BTAX06 + WK_C.CSUKUM(iCol)
'            Else
'                BTAX06 = BTAX06 + CHLCSU
'            End If
'        Case "8"
'            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
'                GITBTX = 0
'            Else
'                GITBTX = Val(WG03.U_CSUAMT(iCol))  '/ 누적한도체크아님 각 수당별한도체크.
'            End If
'
'            If WK_C.CSUKUM(iCol) > 0 And GITBTX > WK_C.CSUKUM(iCol) Then
'                BTAX04 = BTAX04 + WK_C.CSUKUM(iCol)
'            Else
'                BTAX04 = BTAX04 + GITBTX
'            End If
'        Case Else   '/ 비과세-기타-미제출
'            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
'                GITBTX_N = 0
'            Else
'                GITBTX_N = Val(WG03.U_CSUAMT(iCol))    '/ 누적한도체크아님 각 수당별한도체크.
'            End If
'            If WK_C.CSUKUM(iCol) > 0 And GITBTX_N > WK_C.CSUKUM(iCol) Then
'                BTAX07 = BTAX07 + WK_C.CSUKUM(iCol)
'            Else
'                BTAX07 = BTAX07 + GITBTX_N
'            End If
'        End Select
'        If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" And ods_PH_PY111A.getvalue("U_CHGCHK",0) =  "N" Then
'            '/ 상여 상여금에 포함 수당은 총지급액에 합산하지않아야함.
'        Else
'            WG03.U_TOTPAY = WG03.U_TOTPAY + Val(WG03.U_CSUAMT(iCol))
'            '/ 월정급여
'            If WK_C.MPYGBN(iCol) = "Y" Then
'                MONPAY = MONPAY + Val(WG03.U_CSUAMT(iCol))
'            End If
'            '/ 고용보험대상금액
'            If WK_C.GBHGBN(iCol) = "Y" Then
'                TOTGBH = TOTGBH + Val(WG03.U_CSUAMT(iCol))
'            End If
'        End If
'       '/ 통상임금
'        Select Case Trim$(WK_C.CSUCOD(iCol))
'        Case "A01", "E06", "E02", "E11", "E04", "C05" '/ (e06직책수당 + e02생산장려 + e11현장수당 + e04기능수당+C05주휴수당)/30일
'               BASAMT = BASAMT + WG03.U_CSUAMT(iCol)
'        End Select
'    Next iCol
'    If oJOBTYP = "2" Then
'        WG03.U_TOTPAY = WG03.U_TOTPAY + WG03.U_BONUSS
'        MONPAY = MONPAY + WG03.U_BONUSS
'        TOTGBH = TOTGBH + WG03.U_BONUSS
'    End If
'    If oJOBTYP <> "1" And ods_PH_PY111A.getvalue("U_CHGCHK",0) =  "Y" Then
'        Select Case oJOBTYP
'        Case "1"
'        Case "2"
'            WG03.U_BONUSS = WG03.U_TOTPAY
'            WG03.U_APPRAT = "100"
'            WG03.U_BNSRAT = "100"
'        Case "3"
'            WG03.U_APPRAT = "100"
'            WG03.U_BNSRAT = "100"
'        End Select
'    End If
'
'    '/ 상여총액(급여대장에 상여항목포함안될경우)
''    If oJOBTYP = "2" And BnsUsed = False Then  '/ 상여만이면
''        WG03.U_CSUAMT(1) = WG03.U_BONUSS
''        BNSCSU = BNSCSU + Val(WG03.U_CSUAMT(1))
''        WG03.U_TOTPAY = WG03.U_TOTPAY + WG03.U_BONUSS
''        '/ 고용보험대상금액
''        If WK_C.GBHGBN(1) = "Y" Then
''            TOTGBH = TOTGBH + WG03.U_BONUSS
''        End If
''        '/ 월정급여
''        If WK_C.MPYGBN(1) = "Y" Then
''            MONPAY = MONPAY + WG03.U_BONUSS
''        End If
''
''    End If
'    WG03.U_TOTPAY = Format$(WG03.U_TOTPAY, "##########0")
'   If WG03.U_TOTPAY = 0 Then Exit Sub
'' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' '/ 공제내역 계산/
'' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''/ 세액계산 /*******************************************************************************/
''/ 비과세산출 /*******************************************************************************/
'    '/ 비과세누적액 산출
'        WK_YBTXAM(1) = 0: WK_YBTXAM(2) = 0: WK_YBTXAM(3) = 0: WK_YBTXAM(4) = 0: WK_YBTXAM(5) = 0: WK_YBTXAM(6) = 0:: WK_YBTXAM(7) = 0
'        sQry = "EXEC PH_PY111_BTX '" & oYM & "','" & oJOBTYP & "', '" & oJOBGBN & "', '" & Trim$(WG03.U_MSTCOD) & "'"
'        oRecordSet.DoQuery sQry
'        If oRecordSet.RecordCount > 0 Then
'            WK_YBTXAM(1) = Val(oRecordSet.Fields(0).Value)   '/ 연누적액
'            WK_YBTXAM(2) = Val(oRecordSet.Fields(1).Value)   '/ 월누적
'            WK_YBTXAM(3) = Val(oRecordSet.Fields(2).Value)   '/ 월누적
'            WK_YBTXAM(4) = Val(oRecordSet.Fields(3).Value)   '/ 월누적
'            WK_YBTXAM(5) = Val(oRecordSet.Fields(4).Value)   '/ 월누적
'            WK_YBTXAM(6) = Val(oRecordSet.Fields(5).Value)   '/ 월누적
'            WK_YBTXAM(7) = Val(oRecordSet.Fields(6).Value)   '/ 월누적-기타비과세(미제출)
'        End If
'
'    '/ 1.1) 생산직 비과세
'    If sRecordset.Fields("U_BX1SEL").Value = "Y" Then             '/ 비과세대상자만
'       If (MONPAY - BT1AMT) <= 1000000 Then
'            BTAX01 = BT1AMT
'       Else
'            BTAX01 = 0
'       End If
'       If BT1KUM <> 0 And BTAX01 > BT1KUM Then BTAX01 = BT1KUM  '/ 한도액구함
'    '/ 연간 누적 비과세1 체크(상여계산시는 제외)
'       If oJOBTYP <> "2" Then
'          '/ (비과세1+연누적비과세)> 연240만원 THEN 비과세1 = 연240만원-(연누적비과세+비과세1)
'          If (BTAX01 + WK_YBTXAM(1)) > WG01.TB1KUM(1) Then
'             BTAX01 = WG01.TB1KUM(1) - WK_YBTXAM(1)
'          End If
'
'          If BTAX01 < 0 Then BTAX01 = 0
'       End If
'    End If
'
'    '/ 1.2) 과세대상급여 ( 총급여액-비과세소득 )
'    GWASEE = WG03.U_TOTPAY - BTAX01 - BTAX02 - BTAX03 - BTAX04 - BTAX05 - BTAX06
'    WG03.U_GWASEE = GWASEE
'    WG03.U_BTAX01 = BTAX01
'    WG03.U_BTAX02 = BTAX02
'    WG03.U_BTAX03 = BTAX03
'    WG03.U_BTAX04 = BTAX04
'    WG03.U_BTAX05 = BTAX05
'    WG03.U_BTAX06 = BTAX06
'    WG03.U_BTAX07 = BTAX07
'    '/ 1.3) 세액계산
'    '/ 세액계산안함으로 하면 소득세,주민세 계산안함.
'    If ods_PH_PY111A.getvalue("U_TAXCHK",0) =  "N" Then
'        '/ 총지급액, 과세대상금액, 공제인원수
'        GABGUN = 0
'        JUMINN = 0
'    Else
'        If sRecordset.Fields("U_TAXSEL").Value = "Y" Then
'            INCOME = GWASEE '- IIf(sRecordSet.Fields("U_BX1SEL").Value = "Y", BNSCSU, 0)
'            If oJOBTYP = "3" Then WG03.U_TAXTRM = WG03.U_TAXTRM + 1
'            If oJOBTYP <> "1" Then  '/ 상여처리
'               INCOME = MDC_SetMod.IInt((GWASEE + WG03.U_AVRPAY) / WG03.U_TAXTRM, 1)
'            End If
'         '/ 총지급액, 과세대상금액, 공제인원수
'            MDC_SetMod.Get_GabGunSe GABGUN, JUMINN, INCOME, WG03.U_TAXCNT, WG03.U_CHLCNT, oYM, INCOME, PAY_001
'        Else
'            GABGUN = 0
'            JUMINN = 0
'        End If
'    End If
'    '/ 1.4) 고용보험
'    '// 고용보험 계산안함으로 하면 제외
'    If ods_PH_PY111A.getvalue("U_GBHCHK",0) =  "N" Then
'            GBHAMT = 0
'    Else
'        If sRecordset.Fields("U_GBHSEL").Value = "Y" Then
'            GBHAMT = MDC_SetMod.IInt(TOTGBH * (0.45 / 100), 10)
'        End If
'    End If
'    '// 건강보험과 국민연금은 급여+정기, 급상여+정기일때만 자동계산
'    If (oJOBTYP = "1" And oJOBGBN = "1") Or (oJOBTYP = "3" And oJOBGBN = "1") Then
'        '/ 1.5) 의료보험 /'/ 당월 1일이후 입사자 건강보험료 제외
'        If Val(WG03.U_MEDAMT) <> 0 Then
'            sQry = "SELECT TOP 1 U_EMPRAT, U_FROM, U_TO, U_NJYYMM, U_NJYRAT, U_NJCRAT FROM [@ZPY103H] "
'            sQry = sQry & " WHERE CODE <= '" & oYM & "' ORDER BY CODE DESC"
'            oRecordSet.DoQuery sQry
'            If oRecordSet.RecordCount > 0 Then
'                U_NJYYMM = oRecordSet.Fields("U_NJYYMM").Value
'                U_NJYRAT = oRecordSet.Fields("U_NJYRAT").Value
'                U_NJCRAT = oRecordSet.Fields("U_NJCRAT").Value
'                If Val(WG03.U_MEDAMT) < oRecordSet.Fields("U_FROM").Value Then
'                    MEDAMT = oRecordSet.Fields("U_FROM").Value
'                ElseIf oRecordSet.Fields("U_TO").Value > 0 And Val(WG03.U_MEDAMT) > oRecordSet.Fields("U_TO").Value Then
'                    MEDAMT = oRecordSet.Fields("U_TO").Value
'                Else
'                    MEDAMT = Val(WG03.U_MEDAMT)
'                End If
'                MEDAMT = MDC_SetMod.IInt(MEDAMT * Val(oRecordSet.Fields("U_EMPRAT").Value) / 100, 10)
'            End If
'            If oYM = Mid$(Replace(WG03.U_INPDAT, "-", ""), 1, 6) Then
'               If oYM & "01" <> Mid$(WG03.U_INPDAT, 1, 8) Then
'                    MEDAMT = 0                                 '/ 당월 1일이후 입사자 건강보험료 제외
'               End If
'            End If
'        End If
'        '/ 1.5.5) 노인장기요양보험료율
'        If MEDAMT <> 0 Then
'            sQry = "SELECT TOP 1 U_NJYYMM, U_NJYRAT, U_NJCRAT FROM [@ZPY103H] "
'            sQry = sQry & " WHERE U_NJYYMM <= '" & oYM & "' ORDER BY U_NJYYMM DESC"
'            oRecordSet.DoQuery sQry
'            If oRecordSet.RecordCount > 0 Then
'                U_NJYYMM = oRecordSet.Fields("U_NJYYMM").Value
'                U_NJYRAT = oRecordSet.Fields("U_NJYRAT").Value
'                U_NJCRAT = oRecordSet.Fields("U_NJCRAT").Value
'            Else
'                U_NJYYMM = ""
'                U_NJYRAT = 0
'                U_NJCRAT = 0
'            End If
'        End If
'        '/ 1.6) 국민연금 /
'        If oYM <= "200803" Or Val(WG03.U_KUKAMT) = 0 Then
'            sQry = "SELECT TOP 1 ISNULL(T0.U_EMPAMT, 0) FROM [@ZPY102L] T0 WHERE T0.Code <= '" & oYM & "' "
'            sQry = sQry & " AND T0.U_CODNBR ='" & sRecordset.Fields("U_KUKGRD").Value & "'"
'            sQry = sQry & " ORDER BY CODE DESC"
'            oRecordSet.DoQuery sQry
'            If oRecordSet.RecordCount > 0 Then
'                KUKAMT = Val(oRecordSet.Fields(0).Value)
'            End If
'        Else
'            If Trim$(sRecordset.Fields("U_KUKGRD").Value) = "" And Val(WG03.U_KUKAMT) = 0 Then
'                KUKAMT = 0
'            Else
'                sQry = "SELECT TOP 1 U_EMPRAT, U_FROM, U_TO FROM [@ZPY102H] "
'                sQry = sQry & " WHERE CODE <= '" & oYM & "' ORDER BY CODE DESC"
'                oRecordSet.DoQuery sQry
'                If oRecordSet.RecordCount > 0 Then
'                    If Val(WG03.U_KUKAMT) = 0 Then
'                        KUKAMT = 0
'                    ElseIf Val(WG03.U_KUKAMT) < oRecordSet.Fields("U_FROM").Value Then
'                        KUKAMT = oRecordSet.Fields("U_FROM").Value
'                    ElseIf oRecordSet.Fields("U_TO").Value > 0 And Val(WG03.U_KUKAMT) > oRecordSet.Fields("U_TO").Value Then
'                        KUKAMT = oRecordSet.Fields("U_TO").Value
'                    Else
'                        KUKAMT = Val(WG03.U_KUKAMT)
'                    End If
'                    KUKAMT = MDC_SetMod.IInt(KUKAMT * Val(oRecordSet.Fields("U_EMPRAT").Value) / 100, 10)
'                End If
'            End If
'        End If
'        '/ 국민연금(만18세이상~만60미만가입가능) 60세이상월부터 국민연금제외(200805월기준 1948.5.25일생일경우 5월분 국민연금까지납부,6월분부터 공제제외)
'         If Trim$(WG03.U_PERNBR) <> "" Then
'             Select Case Mid$(Trim$(WG03.U_PERNBR), 7, 1)
'             Case "1", "2", "5", "6"
'                KUKSTR = Format$("19" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01", "0000-00-00")
'                KUKEND = Format$(DateAdd("yyyy", 60, KUKSTR), "yyyymm")
'             Case "3", "4", "7", "8"
'                KUKSTR = "20" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01"
'                KUKEND = Format$(DateAdd("y", 60, KUKSTR), "yyyymm")
'             Case Else
'                KUKSTR = ""
'                KUKEND = ""
'             End Select
'            If Trim$(KUKEND) <> "" And Trim$(KUKEND) < oYM Then
'                KUKAMT = 0
'            End If
'         End If
''///////////////////////////////////////////////////
'    End If
'    If WG03.U_TAXTRM = 0 Then WG03.U_TAXTRM = 1
''// 2. 공식가지고 계산하기-고정공제
'    If ods_PH_PY111A.getvalue("U_CHGCHK",0) =  "N" Then
'        For iCol = 1 To 18
'            Select Case WK_G.GONCOD(iCol)
'            Case "G01" '/ 소득세
'                If oJOBTYP = "1" Then
'                    WG03.U_GONAMT(iCol) = MDC_SetMod.IInt(GABGUN, 10)
'                Else
'                    WG03.U_GONAMT(iCol) = MDC_SetMod.IInt((GABGUN * WG03.U_TAXTRM) - WG03.U_NABTAX, 10)
'                End If
'                If WG03.U_GONAMT(iCol) < 0 Then WG03.U_GONAMT(iCol) = 0
'            Case "G02" '/ 주민세
'                If oJOBTYP = "1" Then
'                    WG03.U_GONAMT(iCol) = MDC_SetMod.IInt(JUMINN, 10)
'                Else
'                    WG03.U_GONAMT(iCol) = MDC_SetMod.IInt((MDC_SetMod.IInt((GABGUN * WG03.U_TAXTRM) - WG03.U_NABTAX, 10)) * 0.1, 10)
'                End If
'                If WG03.U_GONAMT(iCol) < 0 Then WG03.U_GONAMT(iCol) = 0
'            Case "G05" '/ 고용보험
'                WG03.U_GONAMT(iCol) = GBHAMT
'            Case "G03" '/ 국민연금
'                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + KUKAMT
'            Case "G04" '/ 건강보험
'                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + MEDAMT
'                TOTMED = WG03.U_GONAMT(iCol)
'            Case "G09" '/ 노조회비
'                If oJOBTYP <> "2" Then    '/ 상여만 지급일경우 건강보험, 국민연금은 제외함.
'                    If Trim$(sRecordset.Fields("U_NOJGBN").Value) = "Y" Then
'                       ' WG03.U_GONAMT(iCol) = MDC_SetMod.RInt(WG03.U_TOTPAY * 0.015, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol)) '2008.05월급여까지
'
'                    '/ '2008.07.15지급분(6월귀속)부터적용
'                         WG03.U_GONAMT(iCol) = MDC_SetMod.RInt(BASAMT * 0.02, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
'                    Else
'                        WG03.U_GONAMT(iCol) = 0
'                    End If
'                End If
'            Case "G10" '/ 노인장기요양보험
'                If U_NJYYMM <= oYM Then
'                    NJCMED = MDC_SetMod.IInt(MEDAMT * (U_NJYRAT / 100), 10) '/ (10원미만 절사)
'                    If Trim$(WG03.U_NJCGBN) = "Y" Then '/ 경감대상자일경우
'                        NJCMED = NJCMED - (MDC_SetMod.IInt(NJCMED * (U_NJCRAT / 100) + 9, 10)) '/ 경감보험료(10원미만절상)
'                    End If
'                    WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + NJCMED       '/ 장기요양보험료
'                End If
'            End Select
'        Next iCol
'    End If
''/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
''/2. 총공제액
''/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'    For iCol = 1 To 18
'        WG03.U_TOTGON = WG03.U_TOTGON + Val(WG03.U_GONAMT(iCol))
'    Next iCol
''/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
''/3. 실지급액
''/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'    WG03.U_SILJIG = MDC_SetMod.IInt(WG03.U_TOTPAY - WG03.U_TOTGON, 1)
'   'If WG03.U_SILJIG < 1000 Then WG03.U_SILJIG = 0
'
'    Set oRecordSet = Nothing
'    Exit Sub
''/////////////////////////////////////////////////////////////////////////////////////////////////
'error_Message:
'    Set oRecordSet = Nothing
'    Sbo_Application.StatusBar.SetText "PH_PY111_Save Error :" & Space$(10) & err.Description, bmt_Short, smt_Error
'End Sub

Private Function Get_ReAmt(oJobDate$, oMstCode$, oSusil$) As Double
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '반환컬럼,조건 컬럼,테이블,조건값,앤드절
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
 
    sQry = "Exec PH_PY106  '" & Trim(oJobDate) & "', '" & Trim(oMstCode$) & "', '" & oSusil & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        Get_ReAmt = 0
    Else
        Get_ReAmt = Val(oRecordSet.Fields(0).VALUE)
    End If
    
    Set oRecordSet = Nothing
End Function


