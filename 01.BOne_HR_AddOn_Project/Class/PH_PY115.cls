VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY115"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY115.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 퇴직금계산
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Private oDS_PH_PY115A   As SAPbouiCOM.DBDataSource
Private oDS_PH_PY115B   As SAPbouiCOM.DBDataSource
Public oMat1            As SAPbouiCOM.Matrix

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Private MsterChk        As Boolean


Private Type WG01CODR
    TB1AMT(1 To 100) As Double
    TB1GON(1 To 100) As Double
    TB1RAT(1 To 100) As Double
    TB1KUM(1 To 100) As Double
End Type
Dim WG01 As WG01CODR

Private FIXTYP      As String
Private ROUNDT      As String
Private RODLEN      As Integer
Private RETCHK      As String
Private RATUSE      As String
Private GNSGBN      As String
Private RETYCH      As String
Private oOLDCHK     As String
    
Private oCode           As String
Private WK_SILSIL(1 To 9) As String
Private WK_ROUNDT(1 To 9) As String
Private WK_LENGTH(1 To 9) As Integer

Private oCLTOPT(1 To 3) As String

Private Type ZPY401_H
    SHWJIG As Double
    SHWTIL As Double
    SHWSUR As Double
    SHWGON As Double
    SHWGWA As Double
    SHWYAG As Double
    SHWYAS As Double
End Type
Dim HS As ZPY401_H

Private Type ZPY401_T
    SPCGON As Double
    SODGON As Double
    RETGON As Double
    TAXSTD As Double
    YTXSTD As Double
    YSANTX As Double
    SANTAX As Double
    RTXGON As Double
    FRNGON As Double
    GULGAB As Double
    WSGGON As Double   '환산급여공제
End Type
Dim WK401 As ZPY401_T

Private SHWJIG(1 To 4) As Double
Private SHWIYO(1 To 4) As Double
Private SHWGSU(1 To 4) As Double
Private SHWTIL(1 To 2) As Double
Private SHWSUR(1 To 2) As Double
Private SHWGON(1 To 2) As Double
Private SHWGWA(1 To 2) As Double
Private SHWYAG(1 To 2) As Double
Private SHWYAS(1 To 2) As Double
Private SHWSAN(1 To 2) As Double

Private BubJung_YN As String
Private WT_SHWTOT As Double  '/ 세액환산명세총계


Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY115.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY115_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY115"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
'    oForm.DataBrowser.BrowseBy = "DocNum"
        
    oForm.Freeze True
    Call PH_PY115_CreateItems
    Call PH_PY115_EnableMenus
    Call PH_PY115_SetDocument(oFromDocEntry01)
'    Call PH_PY115_FormResize
    
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

Private Function PH_PY115_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    Dim oFolder     As SAPbouiCOM.Folder
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY115_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// 데이터셋정의
    '//----------------------------------------------------------------------------------------------

    Set oDS_PH_PY115A = oForm.DataSources.DBDataSources("@PH_PY115A")   '//헤더
    Set oDS_PH_PY115B = oForm.DataSources.DBDataSources("@PH_PY115B")   '//라인(급여내역)

    Set oMat1 = oForm.Items("Mat1").Specific            '//@PH_PY115B
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    
    '/ Check 버튼
    Set oCheck = oForm.Items("Check1").Specific     '/ 잠금
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    Set oCheck = oForm.Items("SPCGBN").Specific     '/ 법정퇴직주현근무지세액공제여부
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    Set oCheck = oForm.Items("JS1SPC").Specific     '/ 법정퇴직종전근무지퇴직세액공제여부
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    Set oCheck = oForm.Items("BUBCHK").Specific     '/ 서식구분
    oCheck.ValOff = "N": oCheck.ValOn = "Y"

    Set oCombo = oForm.Items("JSNGBN").Specific
    oCombo.ValidValues.Add "1", "퇴직정산"
    oCombo.ValidValues.Add "2", "중도정산"
    
    Set oCombo = oForm.Items("RETRES").Specific
    oCombo.ValidValues.Add "1", "정년퇴직"
    oCombo.ValidValues.Add "2", "정리해고"
    oCombo.ValidValues.Add "3", "자발적퇴직"
    oCombo.ValidValues.Add "4", "임원퇴직"
    oCombo.ValidValues.Add "5", "중도정산"
    oCombo.ValidValues.Add "6", "기    타"
    
'//----------------------------------------------------------------------------------------------
''/ 사업장
'    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'    oForm.Items("CLTCOD").DisplayDesc = True
'    oCombo.ValidValues.Add "%", "모두"
'    '// 접속자에 따른 사업장 선택
'    oForm.DataSources.UserDataSources("CLTCOD").ValueEx = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
'
  '/ Folder
'    Call oForm.DataSources.UserDataSources.Add("FolderD", dt_SHORT_TEXT, 1)
'    For i = 1 To 5
'        Set oFolder = oForm.Items("Folder" & i).Specific
'        'oForm.Items("Folder" & i).AffectsFormMode = False
'        oFolder.DataBind.SetBound True, "", "FolderD"
'        If i = 1 Then
'            oFolder.Select
'        Else
'            oFolder.GroupWith ("Folder" & i - 1)
'        End If
'        oForm.Items("Folder" & i).Enabled = True
'    Next i
    
    oForm.Items("Folder1").AffectsFormMode = False
    Set oFolder = oForm.Items("Folder1").Specific
    oFolder.Select
    oForm.Items("Folder2").AffectsFormMode = False
    oForm.Items("Folder3").AffectsFormMode = False
    oForm.Items("Folder4").AffectsFormMode = False
    oForm.Items("Folder5").AffectsFormMode = False
    oForm.Items("Folder6").AffectsFormMode = False
    oForm.Items("Folder1").Enabled = True
    oForm.Items("Folder2").Enabled = True
    oForm.Items("Folder3").Enabled = True
    oForm.Items("Folder4").Enabled = True
    oForm.Items("Folder5").Enabled = True
    oForm.Items("Folder6").Enabled = True
    
    oForm.PaneLevel = 1
    
 '// 1.1. 공식가져오기-수당
    Call BaseSetting
    
    oMat1.Columns("CSUCOD").Editable = False
    oMat1.Columns("CSUNAM").Editable = False
    
    Set oFolder = Nothing
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY115_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY115_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY115_EnableMenus()
    
    On Error GoTo PH_PY115_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY115_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY115_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY115_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY115_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY115_FormItemEnabled
        Call PH_PY115_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY115_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY115_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY115_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY115_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY115_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        Call PH_PY115_FormClear
'        Call PayTitle
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then

        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then

        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY115_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY115_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim j           As Long
    Dim PAYTOT(3)   As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED:
        If pval.BeforeAction Then
            If pval.ItemUID = "1" Then
                'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                '//추가및 업데이트시에
                'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If MDC_SetMod.Value_ChkYn("[@PH_PY115A]", "U_MSTCOD", "'" & oForm.Items("MSTCOD").Specific.String & "'", " AND U_ENDRET = '" & oForm.Items("ENDRET").Specific.String & "'") = False Then
                            Sbo_Application.StatusBar.SetText "이미 저장되어져 있는 헤더의 내용과 일치합니다", bmt_Short, smt_Error
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If PH_PY115_DataValidCheck = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        If Val(oDS_PH_PY115A.GetValue("U_RETPAY", 0)) <= 0 Then
                            BubbleEvent = False
                            Sbo_Application.StatusBar.SetText "퇴직급여의 값이 0입니다. 확인하여 주십시오.", bmt_Short, smt_Error
                            Exit Sub
                        Else
                            If MsterChk = True Then
                                Call MasterUpdate
                                
                            End If
                        End If
                    End If
                End If
            ElseIf pval.ItemUID = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then   '/ ChooseBtn사원리스트
                oForm.Items("MSTCOD").CLICK ct_Regular
                Sbo_Application.ActivateMenuItem ("7425")
                BubbleEvent = False
            ElseIf pval.ItemUID = "Btn3" And oForm.Mode <> fm_FIND_MODE Then    '/ 세액 계산
                If Compute_TAX = True And Compute_TAX_3 = True And oForm.Mode = fm_OK_MODE Then
                    oForm.Mode = fm_UPDATE_MODE
                End If
            ElseIf pval.ItemUID = "Btn2" And oForm.Mode <> fm_FIND_MODE Then    '/ 퇴직금 계산
                If PH_PY115_DataValidCheck(pval.ItemUID) = False Then
                    BubbleEvent = False
                Else
                    If Compute_PAY = True And oForm.Mode = fm_OK_MODE Then
                        oForm.Mode = fm_UPDATE_MODE
                    End If
                End If
            ElseIf pval.ItemUID = "Btn1" And oForm.Mode <> fm_FIND_MODE Then    '/ 퇴직금 생성
                If PH_PY115_DataValidCheck = False Then
                    BubbleEvent = False
                Else
                    If Create_Data = True And oForm.Mode = fm_OK_MODE Then
                        oForm.Mode = fm_UPDATE_MODE
                    End If
                End If
            ElseIf pval.ItemUID = "LBtn1" Then
'                oCompany.UserName
                If oDS_PH_PY115A.GetValue("U_MSTCOD", 0) Then
                
                End If
                Dim oTmpObject As Variant
                Set oTmpObject = New PH_PY001
                If Not oTmpObject Is Nothing Then
                    oTmpObject.LoadForm ("")
                    
'                    If Sbo_Application.Forms.ActiveForm.Type = "PH_PY001" Then
                        Sbo_Application.Forms.ActiveForm.Mode = fm_FIND_MODE
                        Sbo_Application.Forms.ActiveForm.Freeze True
                        Sbo_Application.Forms.ActiveForm.Items("Code").Specific.VALUE = oDS_PH_PY115A.GetValue("U_MSTCOD", 0)
                        Sbo_Application.Forms.ActiveForm.Items("1").CLICK ct_Regular
                        Sbo_Application.Forms.ActiveForm.Freeze False
                        Set oTmpObject = Nothing
                        BubbleEvent = False
                        Exit Sub
'                    End If
                End If
            End If
        Else
            If pval.ItemUID = "1" And pval.ActionSuccess = True And (oForm.Mode = fm_OK_MODE Or oForm.Mode = fm_ADD_MODE) Then
                Call PH_PY115_FormItemEnabled
            ElseIf pval.ItemUID = "Folder1" Then
                oForm.PaneLevel = 1
            ElseIf pval.ItemUID = "Folder2" Then
                oForm.PaneLevel = 2
            ElseIf pval.ItemUID = "Folder3" Then
                oForm.PaneLevel = 3
            ElseIf pval.ItemUID = "Folder4" Then
                oForm.PaneLevel = 4
            ElseIf pval.ItemUID = "Folder5" Then
                oForm.PaneLevel = 5
            ElseIf pval.ItemUID = "Folder6" Then
                oForm.PaneLevel = 6
            End If
        End If
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '//추가모드에서 코드이벤트가 코드에서 일어 났을때
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
         If pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE Then
            If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(pval.ItemUID).Specific.String & "'", "") = True Then
                oForm.Items("MSTCOD").CLICK ct_Regular
                Sbo_Application.ActivateMenuItem ("7425")
                BubbleEvent = False
            End If
         ElseIf pval.BeforeAction = True And pval.ItemUID = "J01NBR" And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE And Trim$(oForm.Items("J01NAM").Specific.String) <> "" Then
            If Trim$(oForm.Items(pval.ItemUID).Specific.String) = "" Then
                Call Sbo_Application.StatusBar.SetText("종전 사업자번호를 입력하여 주십시오.", bmt_Short, smt_Error)
                BubbleEvent = False
            Else
                '/ 사업자번호 체크
                If Len(oForm.Items(pval.ItemUID).Specific.String) <= 12 Then
                    If MDC_SetMod.TaxNoCheck(oForm.Items(pval.ItemUID).Specific.String) = False Then
                        Call Sbo_Application.StatusBar.SetText("사업자번호가 틀립니다. 확인하세요.", bmt_Short, smt_Error)
                        BubbleEvent = False
                    End If
                End If
            End If
         End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat1", "Grid1"
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
'            Call oForm.Freeze(True)
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.BeforeAction = False Then
'                If pval.ItemChanged = True Then
'
'                End If
'            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" And oForm.Mode <> fm_FIND_MODE Then
            If oLastItemUID = "MSTCOD" Then
                If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(oLastItemUID).Specific.String & "'", "") = True _
                And oForm.Items(oLastItemUID).Specific.String <> "" And oLastItemUID <> pval.ItemUID Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7

'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
            If pval.BeforeAction = False And pval.ItemChanged = True Then
'            If pval.FormMode <> fm_FIND_MODE And _
'               (pval.ItemUID = "GNSYER" Or pval.ItemUID = "GNSMON" Or pval.ItemUID = "GNSDAY" Or pval.ItemUID = "TGSDAY" Or _
'                pval.ItemUID = "GNMYER" Or pval.ItemUID = "GNMMON" Or pval.ItemUID = "EXPMON" Or pval.ItemUID = "BNST04" Or _
'                pval.ItemUID = "BNST05" Or pval.ItemUID = "BNST06" Or pval.ItemUID = "BNST07" Or pval.ItemUID = "BNST08" Or _
'                pval.ItemUID = "BNST09" Or pval.ItemUID = "BNST10" Or pval.ItemUID = "BNST11" Or pval.ItemUID = "BNST12" Or _
'                pval.ItemUID = "YONCSU" Or pval.ItemUID = "SUDAMT" Or _
'                pval.ItemUID = "GITAA1" Or pval.ItemUID = "TGSDAY" Or pval.ItemUID = "ADDRAT") Then
'            then
            If pval.ItemUID = "MSTCOD" Then
               FlushToItemValue pval.ItemUID
            ElseIf pval.FormMode <> fm_FIND_MODE And _
               (pval.ItemUID = "PAYSD1" Or pval.ItemUID = "PAYSD2" Or pval.ItemUID = "PAYSD3" Or pval.ItemUID = "PAYSD4" Or _
                pval.ItemUID = "PAYED1" Or pval.ItemUID = "PAYED2" Or pval.ItemUID = "PAYED3" Or pval.ItemUID = "PAYED4" Or _
                pval.ItemUID = "JINDAT" Or pval.ItemUID = "JOTDAT" Or pval.ItemUID = "JIGBIL" Or _
                pval.ItemUID = "JRET01" Or pval.ItemUID = "JADD01" Or pval.ItemUID = "JYIL03" Or pval.ItemUID = "JYIL04") Then
                Call FlushToItemValue(pval.ItemUID)
            '// 근속일수 구함
            ElseIf pval.FormMode = fm_ADD_MODE And (pval.ItemUID = "ENDRET") Then
                Call FlushToItemValue(pval.ItemUID)
            '// 근무일수 구함
            ElseIf pval.FormMode = fm_ADD_MODE And (pval.ItemUID = "GNSYER" Or pval.ItemUID = "GNSMON" Or pval.ItemUID = "GNSDAY") Then
                Call FlushToItemValue(pval.ItemUID)
            '// 퇴직금 산출내역 변경시 금액 자동계산
            ElseIf pval.FormMode <> fm_FIND_MODE And _
               (pval.ItemUID = "PAYT01" Or pval.ItemUID = "PAYT02" Or pval.ItemUID = "PAYT03" Or pval.ItemUID = "PAYT04" Or _
                pval.ItemUID = "BNST01" Or pval.ItemUID = "BNST02" Or pval.ItemUID = "BNST03" Or pval.ItemUID = "BNST04" Or _
                pval.ItemUID = "BNST05" Or pval.ItemUID = "BNST06" Or pval.ItemUID = "BNST07" Or pval.ItemUID = "BNST08" Or _
                pval.ItemUID = "BNST09" Or pval.ItemUID = "BNST10" Or pval.ItemUID = "BNST11" Or pval.ItemUID = "BNST12" Or _
                pval.ItemUID = "YONCSU" Or pval.ItemUID = "SUDAMT" Or pval.ItemUID = "GITAA2" Or pval.ItemUID = "ExPay2" Or _
                pval.ItemUID = "GITAA1" Or pval.ItemUID = "TGSDAY" Or pval.ItemUID = "ADDRAT") Then

                    If Compute_PAY = True Then
                        If oForm.Mode = fm_OK_MODE Then oForm.Mode = fm_UPDATE_MODE
                    End If
            '//*** 퇴직연금명세 ***//
            ElseIf pval.FormMode <> fm_FIND_MODE And _
               (pval.ItemUID = "JYNTOT" Or pval.ItemUID = "JYNWON" Or pval.ItemUID = "JYNBUL" Or pval.ItemUID = "JYNGON" Or _
                pval.ItemUID = "MYNTOT" Or pval.ItemUID = "MYNWON" Or pval.ItemUID = "MYNBUL" Or pval.ItemUID = "MYNGON") Then
                Call Display_ilSiGum
            
            '//퇴직이연계좌
'            ElseIf pval.FormMode <> fm_FIND_MODE And pval.ItemUID = "MYGACC" Then
'                '//퇴직이연 계좌번호를 넣으면 퇴직소득금액 자동입력
'                If Trim(oDS_PH_PY115A.GetValue("U_MYGACC", 0)) <> "" Then
'                    oDS_PH_PY115A.setValue "U_IPGM01", 0, oDS_PH_PY115A.GetValue("U_TJKPAY", 0)
'                Else
'                    oDS_PH_PY115A.setValue "U_IPGM01", 0, 0
'                End If
                
'            ElseIf pval.FormMode <> fm_FIND_MODE And pval.ItemUID = "IPGM01" Then
'                oDS_PH_PY115A.setValue "U_GWASEY", 0, Round(oDS_PH_PY115A.GetValue("U_SANTAX1", 0) * oDS_PH_PY115A.GetValue("U_IPGM01", 0) / oDS_PH_PY115A.GetValue("U_TJKPAY", 0), 0)
            '//*** 세액환산명세 ***//
            ElseIf pval.FormMode <> fm_FIND_MODE And _
               (pval.ItemUID = "SH1JIG" Or pval.ItemUID = "SH3JIG") Then '/ 세액환산-법정
                Call Display_TaxWhanSan(1, oDS_PH_PY115A.GetValue("U_BUBCHK", 0))
            ElseIf pval.FormMode <> fm_FIND_MODE And _
               (pval.ItemUID = "SH2JIG" Or pval.ItemUID = "SH4JIG") Then       '/ 세액환산-법정외
                Call Display_TaxWhanSan(2, oDS_PH_PY115A.GetValue("U_BUBCHK", 0))
            '//*** 공제항목명세 ***//
            ElseIf pval.FormMode <> fm_FIND_MODE And _
               (pval.ItemUID = "GULJUM" Or pval.ItemUID = "JONGAB" Or pval.ItemUID = "JONJUM" Or pval.ItemUID = "GITJA1" Or _
                pval.ItemUID = "JSNGAB" Or pval.ItemUID = "JSNJUM" Or pval.ItemUID = "MEDJSN" Or pval.ItemUID = "GITJA2") Then
                Call Compute_TAX_3
                
'            '// 세액산출내역 변경시 금액 자동계산
'            ElseIf pval.FormMode <> fm_FIND_MODE And _
'               (pval.ItemUID = "GITJA1" Or pval.ItemUID = "MEDJSN" Or pval.ItemUID = "JRET01" Or pval.ItemUID = "JSUD01" Or _
'                pval.ItemUID = "JBHM01" Or pval.ItemUID = "JONGAB" Or pval.ItemUID = "JONJUM" Or pval.ItemUID = "JSNGAB" Or _
'                pval.ItemUID = "JSNJUM" Or pval.ItemUID = "TAXGON" Or pval.ItemUID = "TJKPAY" Or _
'                pval.ItemUID = "GULGAB" Or pval.ItemUID = "JONGAB" Or pval.ItemUID = "JONJUM" Or pval.ItemUID = "BHMAMT") Then
'                If PH_PY115_DATAVALIDCHECK = False Then
'                    BubbleEvent = False
'                Else
'                    If Compute_TAX("ETC") = True Then
'                        If oForm.Mode = fm_OK_MODE Then oForm.Mode = fm_UPDATE_MODE
'                    End If
'                End If
            '//급여내역 변경시 자동 금액계산
            ElseIf pval.FormMode <> fm_FIND_MODE And _
                (pval.ItemUID = "Mat1" And Left(pval.ColUID, 6) = "CSUAMT") Then
                oMat1.FlushToDataSource
                For i = 0 To UBound(PAYTOT)
                    PAYTOT(i) = 0
                Next i
                
                For i = 0 To oMat1.VisualRowCount - 1
                    For j = 0 To 3
                        PAYTOT(j) = PAYTOT(j) + Val(Trim(oDS_PH_PY115B.GetValue("U_CSUAMT" & j + 1, i)))
                    Next j
                Next i
                For i = 0 To UBound(PAYTOT)
                    oDS_PH_PY115A.setValue "U_PAYT0" & i + 1, 0, PAYTOT(i)
                Next i
                
                oForm.Items("PAYT01").Update
                oForm.Items("PAYT02").Update
                oForm.Items("PAYT03").Update
                oForm.Items("PAYT04").Update
            End If
        End If
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
'            If pval.BeforeAction = True Then
'            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'
'                Call PH_PY115_FormItemEnabled
'                Call PH_PY115_AddMatrixRow
'
'            End If
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
                Set oDS_PH_PY115A = Nothing
                Set oMat1 = Nothing
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
    Dim oMast       As ZPAY_g_EmpID
    Dim sQry        As String
    Dim Code      As String
    Dim oRecordSet  As SAPbobsCOM.Recordset

Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    On Error GoTo Raise_FormMenuEvent_Error
    Call oForm.Freeze(True)
    
    If (pval.BeforeAction = True) Then
        Select Case pval.MenuUID
            Case "1283":
                Code = oForm.Items("Code").Specific.String
                sQry = "Select ENDRET = Convert(Char(8),DateAdd(dd, 1,U_ENDRET),112) From [@PH_PY115A] Where Code = '" & Code & "'"
                oRecordSet.DoQuery sQry
                
                oMast = MDC_SetMod.Get_EmpID_InFo(oForm.Items("MSTCOD").Specific.String)
                
                If oRecordSet.Fields(0).VALUE > Trim$(oMast.RETDAT) Then
                    If Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") = 2 Then
                        BubbleEvent = False
                        Call oForm.Freeze(False)
                        Exit Sub
                        
                    End If
                Else
                    Sbo_Application.MessageBox ("퇴직(중간)정산 확정자료는 삭제할 수 없습니다.")
                    BubbleEvent = False
                    Call oForm.Freeze(False)
                    Exit Sub
                End If
            Case "1284":
            Case "1286":
            Case "1293":
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY115_FormItemEnabled
                Call PH_PY115_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY115_FormItemEnabled
                Call PH_PY115_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY115_FormItemEnabled
                Call PH_PY115_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY115_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY115B.Size - 1)
'                        If oDS_PH_PY115B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY115B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY115B.Size
'                        Call oDS_PH_PY115B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
'                End If
'                Call PH_PY115_AddMatrixRow
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

Sub PH_PY115_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY115_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 용]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY115B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY115B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY115B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY115B.Offset = oRow
'            oDS_PH_PY115B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY115B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY115B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY115B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY115B.Offset = oRow - 1
'            oDS_PH_PY115B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY115B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY115B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY115B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY115B.Offset = oRow
'        oDS_PH_PY115B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY115B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY115B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY115B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY115_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY115_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY115_FormClear()
On Error GoTo PH_PY115_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY115'", "")
    If DocEntry = 0 Then
        oForm.Items("Code").Specific.VALUE = "00000001"
    Else
        oForm.Items("Code").Specific.VALUE = Right(Replace("0000000" & Str(DocEntry), " ", ""), 8)
    End If
    Exit Sub
PH_PY115_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY115_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY115_DataValidCheck(Optional oUID As String) As Boolean
On Error GoTo PH_PY115_DataValidCheck_Error
    PH_PY115_DataValidCheck = False
    Dim ErrNum      As Integer
    Dim TOTAMT      As Double
    
    
    ErrNum = 0
    '/ 헤더부분 체크
    Select Case True
    Case Trim$(oDS_PH_PY115A.GetValue("U_Check", 0)) = "Y" And Trim$(oOLDCHK) = "Y"
        Sbo_Application.StatusBar.SetText "잠금 자료입니다. 확인하여 주십시오.", bmt_Short, smt_Error
        Exit Function
    Case oDS_PH_PY115A.GetValue("U_MSTCOD", 0) = ""
        Sbo_Application.StatusBar.SetText "사원번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        Exit Function
    Case oDS_PH_PY115A.GetValue("U_STRRET", 0) = ""
        Sbo_Application.StatusBar.SetText "정산시작일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        Exit Function
    Case oDS_PH_PY115A.GetValue("U_ENDRET", 0) = ""
        Sbo_Application.StatusBar.SetText "정산종료일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        Exit Function
    Case oDS_PH_PY115A.GetValue("U_ENDRET", 0) < oDS_PH_PY115A.GetValue("U_STRRET", 0)
        Sbo_Application.StatusBar.SetText "정산종료일자가 정산시작일자보다 작습니다. 확인하여 주십시오.", bmt_Short, smt_Error
        Exit Function
    Case oDS_PH_PY115A.GetValue("U_JIGBIL", 0) = ""
        Sbo_Application.StatusBar.SetText "지급일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        Exit Function
    Case oDS_PH_PY115A.GetValue("U_SINYMM", 0) = ""
        Sbo_Application.StatusBar.SetText "신고연월은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        Exit Function
    Case Val(oDS_PH_PY115A.GetValue("U_SANTAX", 0)) < Val(oDS_PH_PY115A.GetValue("U_TAXGON", 0))
        Sbo_Application.StatusBar.SetText "세액(외국납부)공제는 산출세액을 초과할수 없습니다.", bmt_Short, smt_Error
        Exit Function
    Case oDS_PH_PY115A.GetValue("U_STRINT", 0) = "", oDS_PH_PY115A.GetValue("U_ENDINT", 0) = ""
        Sbo_Application.StatusBar.SetText "귀속기간은 필수입니다. 입력하시기 바랍니다.", bmt_Short, smt_Error
        Exit Function
    Case oDS_PH_PY115A.GetValue("U_RETRES", 0) = ""
        Sbo_Application.StatusBar.SetText "퇴직사유는 필수입니다. 선택하시기 바랍니다.", bmt_Short, smt_Error
        Exit Function
    Case Trim$(oDS_PH_PY115A.GetValue("U_ENDINT", 0)) >= "20091231" And _
        (Trim$(oDS_PH_PY115A.GetValue("U_SPCGBN", 0)) = "Y" Or Trim$(oDS_PH_PY115A.GetValue("U_JS1SPC", 0)) = "Y")
        Sbo_Application.StatusBar.SetText "세액공제는 2009년에 한해서만 가능 합니다.", bmt_Short, smt_Error
        Exit Function
    '// SOCGBN - 퇴직세액공제여부, JS1SPC - 종전세액공제여부
        
    End Select
    TOTAMT = 0
    TOTAMT = TOTAMT + Val(oDS_PH_PY115A.GetValue("U_TJKPAY", 0))        '//퇴직급여
    TOTAMT = TOTAMT + Val(oDS_PH_PY115A.GetValue("U_SUDAMT", 0))        '//명예퇴직수당(50%)
    TOTAMT = TOTAMT + Val(oDS_PH_PY115A.GetValue("U_BHMAMT", 0))        '//단체퇴직보험
    TOTAMT = TOTAMT + Val(oDS_PH_PY115A.GetValue("U_YILPA1", 0))        '//법정퇴직연금일시금
    TOTAMT = TOTAMT + Val(oDS_PH_PY115A.GetValue("U_YILPA2", 0))        '//법정외퇴직연금일시금
    
    TOTAMT = TOTAMT + Val(oDS_PH_PY115A.GetValue("U_JTOT01", 0))        '//종전퇴직계1
    TOTAMT = TOTAMT + Val(oDS_PH_PY115A.GetValue("U_JTOT02", 0))        '//종전퇴직계2
    
    '//RETPAY - 총퇴직금, SH1JIG - 환산-지급예상액(현)
    If oUID <> "Btn2" Then '// 퇴직금계산이 아닐때
        If TOTAMT <> Val(oDS_PH_PY115A.GetValue("U_RETPAY", 0)) And Val(oDS_PH_PY115A.GetValue("U_SH1JIG", 0)) = 0 Then
            Sbo_Application.StatusBar.SetText "퇴직금이 변경되었습니다. 세액계산이 필요합니다.", bmt_Short, smt_Error
            Exit Function
        End If
    End If
    
    '//국민연금퇴직전환금 체크
    
    PH_PY115_DataValidCheck = True
    
    Exit Function
    
PH_PY115_DataValidCheck_Error:

    PH_PY115_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY115_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY115_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY115_Validate_Error
    PH_PY115_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY115A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY115_Validate = False
        GoTo PH_PY115_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY115_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY115_Validate_Error:
    PH_PY115_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY115_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PH_PY115_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    
    Dim BPLID           As String
    Dim ItmBsort        As String
    Dim DocDate         As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY115_Print_Report01_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY115_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[S142] 발주서"
    ReportName = "S142_1.rpt"
    sQry = "EXEC PH_PY115_1 '" & oForm.Items("8").Specific.VALUE & "'"
    
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
    
PH_PY115_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY115_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub BaseSetting()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim i%
    Dim ErrNum          As Integer
    Dim JSNYER          As String
    
    ErrNum = 0
    FIXTYP = 0:    ROUNDT = 0:    RODLEN = 0:    RETCHK = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    'JSNYER = Left$(oDS_PH_PY115A.GetValue("U_ENDRET", 0), 4)  '/ 정산종료일
    '//2009.10.26 중도정산할경우 2009년에 2005년까지 정산해달라고 하는경우등등 지급하기로 한 약정일의 귀속년도기준으로 봄.
    '//해서 귀속기간년도로 변경함.
    JSNYER = Left$(oDS_PH_PY115A.GetValue("U_ENDINT", 0), 4)  '/ 귀속기간종료
    If Trim$(JSNYER) = "" Then JSNYER = Format$(Now, "YYYY")
    
    '/ 91 기초세액 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    For i = 1 To 100
        WG01.TB1AMT(i) = 0:  WG01.TB1GON(i) = 0:  WG01.TB1RAT(i) = 0:  WG01.TB1KUM(i) = 0
    Next i
    '/
    sQry = "SELECT U_CODNBR, U_CODAMT, U_CODGON, U_CODRAT, U_CODKUM "
    'sQry = sQry & " FROM [@PH_PY100B] WHERE Code = '" & Format$(Now, "YYYY") & "' "
    sQry = sQry & " FROM [@PH_PY100B] WHERE Code = (SELECT Top 1 Code FROM [@PH_PY100A] WHERE Code <='" & Trim$(JSNYER) & "' ORDER BY Code Desc) "
    sQry = sQry & " ORDER BY Code, Convert(Integer,U_CODNBR) DESC"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
       WG01.TB1AMT(Val(oRecordSet.Fields("U_CODNBR").VALUE)) = Val(oRecordSet.Fields("U_CODAMT").VALUE)
       WG01.TB1GON(Val(oRecordSet.Fields("U_CODNBR").VALUE)) = Val(oRecordSet.Fields("U_CODGON").VALUE)
       WG01.TB1RAT(Val(oRecordSet.Fields("U_CODNBR").VALUE)) = Val(oRecordSet.Fields("U_CODRAT").VALUE)
       WG01.TB1KUM(Val(oRecordSet.Fields("U_CODNBR").VALUE)) = Val(oRecordSet.Fields("U_CODKUM").VALUE)
       oRecordSet.MoveNext
    Loop
'
' '// 업체 환경 설정내용 가져오기
'    oCLTOPT(1) = "0": oCLTOPT(2) = "": oCLTOPT(3) = ""
'    sQry = "SELECT Code, U_CODGBN, U_CODST1, U_CODST2 FROM [@ZPY304H] WHERE Code IN ('GIT101','GIT201') "
'    oRecordSet.DoQuery sQry
'    Do Until oRecordSet.EOF
'        Select Case Trim$(oRecordSet.Fields("Code").Value)
'        Case "GIT101": oCLTOPT(1) = oRecordSet.Fields("U_CODGBN").Value
'        Case "GIT201": oCLTOPT(2) = oRecordSet.Fields("U_CODST1").Value
'                       oCLTOPT(3) = oRecordSet.Fields("U_CODST2").Value
'        End Select
'       oRecordSet.MoveNext
'    Loop
        
    '/ 타이틀
    sQry = " SELECT TOP 1 * FROM [@PH_PY114A] T1 ORDER BY T1.Code DESC"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        oCode = oRecordSet.Fields("Code").VALUE
        FIXTYP = oRecordSet.Fields("U_FIXTYP").VALUE            '/평균임금산정방법
        ROUNDT = oRecordSet.Fields("U_ROUNDT").VALUE            '/일할끝전처리
        RODLEN = CInt(Val(oRecordSet.Fields("U_LENGTH").VALUE)) '/일할단위
        RETCHK = oRecordSet.Fields("U_RETCHK").VALUE            '/비정기상여포함여부
        RATUSE = oRecordSet.Fields("U_RATUSE").VALUE            '/임원누진율
        GNSGBN = oRecordSet.Fields("U_GNSGBN").VALUE            '/근속년수기준
        RETYCH = oRecordSet.Fields("U_RETYCH").VALUE            '/퇴사월연차수당처리
    End If
    '/ 라인정보
    i = 0
    sQry = " SELECT  T0.U_CSUCOD, T0.U_SILCOD, T0.U_ROUNDT, T0.U_LENGTH FROM [@PH_PY114B] T0"
    sQry = sQry & " WHERE T0.Code = '" & oCode & "' ORDER BY T0.Code, T0.U_CSUCOD"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        i = i + 1
        WK_SILSIL(i) = oRecordSet.Fields("U_SILCOD").VALUE      '/코딩용계산식
        WK_ROUNDT(i) = oRecordSet.Fields("U_ROUNDT").VALUE      '/끝전처리
        WK_LENGTH(i) = CInt(Val(oRecordSet.Fields("U_LENGTH").VALUE))   '/단위
        oRecordSet.MoveNext
    Loop

    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    oForm.Freeze False
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "퇴직금기준설정에 대한 데이터가 없습니다. 입력하여주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "퇴직급여총액이 0입니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "BaseSetting 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub

Private Function Compute_GNMYER() As Boolean
On Error GoTo Error_Message
    Dim GNMMON(1 To 4)     As Integer
    Dim DUPMON(1 To 4)     As Integer
    Dim JMMMON As Integer
    Dim GN1YER As Integer
    Dim GN2YER As Integer
    
'// 법정퇴직 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    '/ 1) 주(현)-법정 근무월수
    GNMMON(1) = Val(oDS_PH_PY115A.GetValue("U_GNSYER", 0)) * 12 + Val(oDS_PH_PY115A.GetValue("U_GNSMON", 0))    '/(근속년수 * 12 + 근속월수)
    If Val(oDS_PH_PY115A.GetValue("U_GNSDAY", 0)) > 0 Then  '/근속일수
        GNMMON(1) = GNMMON(1) + 1
    End If
    DUPMON(1) = Val(oDS_PH_PY115A.GetValue("U_EXPMON", 0))   '/ 제외일수
    GN1YER = (GNMMON(1) - DUPMON(1))
        
    '/ 2)종(전)-근무월수
    GNMMON(2) = Val(oDS_PH_PY115A.GetValue("U_JONYER", 0)) * 12 + Val(oDS_PH_PY115A.GetValue("U_JONMON", 0))    '/(종전년수 * 12 + 종전월수)
    DUPMON(2) = Val(oDS_PH_PY115A.GetValue("U_JEXMON", 0))  '/종전제외월수
    
    '/ 3)중복월수-주(현)과 종(전)일자비교
    JMMMON = Val(oDS_PH_PY115A.GetValue("U_JMMMON", 0))
    
    '/ 4) 세법상 근속년수-법정퇴직
    If GNMMON(2) > 0 Then
        GN1YER = (GNMMON(1) - DUPMON(1) + GNMMON(2) - DUPMON(2) - JMMMON)
    End If
    If (GN1YER Mod 12) = 0 Then
        GN1YER = GN1YER / 12
    Else
        GN1YER = MDC_SetMod.IInt(GN1YER / 12, 1) + 1
    End If

'// 법정외 퇴직 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    '/ 1) 주(현)-법정외 근무월수
    GNMMON(3) = Val(oDS_PH_PY115A.GetValue("U_GN2MON", 0))
    DUPMON(3) = Val(oDS_PH_PY115A.GetValue("U_EX2MON", 0))
    '/ 2)종(전)-근무월수
    
    
    '/ 4) 세법상 근속년수-법정외퇴직
    GN2YER = GNMMON(3) - DUPMON(3)
    If (GN2YER Mod 12) = 0 Then
        GN2YER = GN2YER / 12
    Else
        GN2YER = MDC_SetMod.IInt(GN2YER / 12, 1) + 1
    End If
    
 '/ Display
    oDS_PH_PY115A.setValue "U_GNMMON", 0, GNMMON(1)       '/ 주(현) 법정 근속월수
    oDS_PH_PY115A.setValue "U_GNMDAY", 0, GNMMON(2)       '/ 종(전) 법정 근속월수(필드대체해서사용하는중)

    oDS_PH_PY115A.setValue "U_GNMYER", 0, GN1YER          '/ 세법상의 주현 법정 근속년수
    oDS_PH_PY115A.setValue "U_GN2YER", 0, GN2YER          '/ 세법상의 주현 법정외 근속년수
   
   oForm.Items("GNMMON").Update
   oForm.Items("GNMDAY").Update
   oForm.Items("GNMYER").Update
    
    Compute_GNMYER = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    oForm.Freeze False
    
    Sbo_Application.StatusBar.SetText "Compute_GNMYER 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    Compute_GNMYER = False
End Function
'---------------------------------------------------------------------------------------
' Procedure : Compute_GNSDAY
' 작성자    : 함미경
' Date      :
' Purpose   : 퇴직금 계산 루틴
' 수정자    : 함미경
' 수정일자  : 2009-03-09
' 수정사항  : 1) 프로시저 분리 및 이름변경(Execution -> Compute_GNSDAY)
'             2) 근속연월일/총근무일수구함.
'---------------------------------------------------------------------------------------
Private Function Compute_GNSDAY() As Boolean
On Error GoTo Error_Message
    Dim STRDAT      As String * 8
    Dim ENDDAT      As String * 8
    Dim INPDAT      As String * 8
    Dim OUTDAT      As String * 8

    Dim WK_GNSYER As Integer:   Dim WK_GNSMON As Integer:   Dim WK_GNSDAY As Integer:   Dim WK_TGSDAY As Integer
    Dim WK_GN1YER As Integer:   Dim WK_GN1MON As Integer: Dim WK_EXPMON As Integer
    Dim WK_GN2YER As Integer:   Dim WK_GN2MON As Integer: Dim WK_EX2MON As Integer
    
    '/1. 귀속연도
    STRDAT = oDS_PH_PY115A.GetValue("U_STRRET", 0)    '/ 정산시작일
    ENDDAT = oDS_PH_PY115A.GetValue("U_ENDRET", 0)    '/ 정산종료일
    INPDAT = oDS_PH_PY115A.GetValue("U_INPDAT", 0)    '/ 입사일
    OUTDAT = oDS_PH_PY115A.GetValue("U_OUTDAT", 0)    '/ 퇴사일
    
    If Trim$(STRDAT) = "" And Trim$(ENDDAT) = "" Then
        Exit Function
    End If
    
    '/1. 지급조서 귀속기간 재설정
    If Trim$(oDS_PH_PY115A.GetValue("U_STRINT", 0)) = "" Then   '/귀속시작일
        If Mid$(STRDAT, 1, 4) = Mid$(ENDDAT, 1, 4) Then
            oDS_PH_PY115A.setValue "U_STRINT", 0, STRDAT
            If Mid$(oDS_PH_PY115A.GetValue("U_JOTDAT", 0), 1, 4) = Mid$(ENDDAT, 1, 4) Then  '/종전퇴사일
                oDS_PH_PY115A.setValue "U_STRINT", 0, Mid$(ENDDAT, 1, 4) & "0101"
            End If
        Else
            oDS_PH_PY115A.setValue "U_STRINT", 0, Mid$(ENDDAT, 1, 4) & "0101"
        End If
    End If
    If Trim$(oDS_PH_PY115A.GetValue("U_ENDINT", 0)) = "" Then
        oDS_PH_PY115A.setValue "U_ENDINT", 0, ENDDAT
    End If
    '/ 1. 법정퇴직금 /*******************************************************************/
    Call MDC_SetMod.Term2(STRDAT, ENDDAT)
    
    '/ 1.1) 근속연월일
    WK_GNSYER = Val(ZPAY_GBL_GNSYER)
    WK_GNSMON = Val(ZPAY_GBL_GNSMON)
    WK_GNSDAY = Val(ZPAY_GBL_GNSDAY)
    WK_TGSDAY = MDC_SetMod.TermDay(STRDAT, ENDDAT)
    If InStr(1, WK_SILSIL(7), "X07") > 0 Then '/ 평균임금계산시 X07:근속년 360으로 구할경우
        WK_TGSDAY = (WK_GNSYER * 360) + (WK_GNSMON * 30) + WK_GNSDAY + 1
    End If
    
    '/ 1.2) 법정퇴직금 근속월
    WK_GN1MON = WK_GNSYER * 12 + WK_GNSMON
    If (WK_GNSDAY) > 0 Then
        WK_GN1MON = WK_GN1MON + 1
    End If
    '/1.2) 법정퇴직금 제외월
    WK_EXPMON = Val(oDS_PH_PY115A.GetValue("U_EXPMON", 0))
    
    '/1.3) 법정퇴직금 근속년
    WK_GN1YER = WK_GN1MON - WK_EXPMON
    If (WK_GN1YER Mod 12) = 0 Then
        WK_GN1YER = MDC_SetMod.IInt((WK_GN1MON - WK_EXPMON) / 12, 1)
    Else
        WK_GN1YER = MDC_SetMod.IInt((WK_GN1MON - WK_EXPMON) / 12 + 1, 1)
    End If
    '/2. 법정이외퇴직금 /*******************************************************************/
    If Trim$(ENDDAT) <> "" Then
        If oDS_PH_PY115A.GetValue("U_JINDAT", 0) = INPDAT Then
            Call MDC_SetMod.Term2(STRDAT, ENDDAT)
        Else
            Call MDC_SetMod.Term2(INPDAT, ENDDAT)
        End If
        WK_GN2MON = Val(ZPAY_GBL_GNSYER) * 12 + Val(ZPAY_GBL_GNSMON)
        If Val(ZPAY_GBL_GNSDAY) > 0 Then WK_GN2MON = WK_GN2MON + 1
        '/1.2) 법정퇴직금 제외월
        WK_EX2MON = Val(oDS_PH_PY115A.GetValue("U_EXPMON", 0))
        
        WK_GN2YER = MDC_SetMod.IInt((WK_GN2MON - WK_EX2MON) / 12, 1)
        If (WK_GN2YER Mod 12) = 0 Then
            WK_GN2YER = MDC_SetMod.IInt((WK_GN2MON - WK_EX2MON) / 12, 1)
        Else
            WK_GN2YER = MDC_SetMod.IInt((WK_GN2MON - WK_EX2MON) / 12 + 1, 1)
        End If
        
    End If

'/ Display
    oDS_PH_PY115A.setValue "U_GNSYER", 0, WK_GNSYER '/ 근속년
    oDS_PH_PY115A.setValue "U_GNSMON", 0, WK_GNSMON '/ 근속월
    oDS_PH_PY115A.setValue "U_GNSDAY", 0, WK_GNSDAY '/ 근속일
    oDS_PH_PY115A.setValue "U_TGSDAY", 0, WK_TGSDAY '/ 총근속일
  '/
    oDS_PH_PY115A.setValue "U_GNMMON", 0, WK_GN1MON '/ 법정월수
    oDS_PH_PY115A.setValue "U_GNMYER", 0, WK_GN1YER '/ 세법상의근속년수
    
    oDS_PH_PY115A.setValue "U_GN2MON", 0, WK_GN2MON '/ 법정외월수
    oDS_PH_PY115A.setValue "U_GN2YER", 0, WK_GN2YER '/ 법정외년수
    
    
    If oDS_PH_PY115A.GetValue("U_STRRET", 0) <= "20121231" Then
    
        '/2012년12월31 이전 근속월수
        STRDAT = oDS_PH_PY115A.GetValue("U_DSDAT1", 0)    '/ 정산시작일
        ENDDAT = oDS_PH_PY115A.GetValue("U_DEDAT1", 0)    '/ 정산종료일
        
        Call MDC_SetMod.Term2(STRDAT, ENDDAT)
        
        WK_GNSYER = Val(ZPAY_GBL_GNSYER)
        WK_GNSMON = Val(ZPAY_GBL_GNSMON)
        WK_GNSDAY = Val(ZPAY_GBL_GNSDAY)
        
        WK_GNSMON = WK_GNSYER * 12 + WK_GNSMON
        
        If WK_GNSDAY > 0 Then
            WK_GNSMON = WK_GNSMON + 1
        End If
        
        oDS_PH_PY115A.setValue "U_GNMON1", 0, WK_GNSMON '/ 근속월
        
        If (WK_GNSMON Mod 12) = 0 Then
            WK_GN2YER = MDC_SetMod.IInt(WK_GNSMON / 12, 1)
        Else
            WK_GN2YER = MDC_SetMod.IInt(WK_GNSMON / 12 + 1, 1)
        End If
        
        oDS_PH_PY115A.setValue "U_GNYER1", 0, WK_GN2YER '/ 근속년
    Else
        oDS_PH_PY115A.setValue "U_GNMON1", 0, 0 '/ 근속월
        oDS_PH_PY115A.setValue "U_GNYER1", 0, 0 '/ 근속년
    End If
    
    '/2012년12월31 이후 근속월수
    STRDAT = oDS_PH_PY115A.GetValue("U_DSDAT2", 0)    '/ 정산시작일
    ENDDAT = oDS_PH_PY115A.GetValue("U_DEDAT2", 0)    '/ 정산종료일
    
    Call MDC_SetMod.Term2(STRDAT, ENDDAT)
    
    WK_GNSYER = Val(ZPAY_GBL_GNSYER)
    WK_GNSMON = Val(ZPAY_GBL_GNSMON)
    WK_GNSDAY = Val(ZPAY_GBL_GNSDAY)
    
    WK_GNSMON = WK_GNSYER * 12 + WK_GNSMON
    
    If WK_GNSDAY > 0 Then
        WK_GNSMON = WK_GNSMON + 1
    End If
    
    oDS_PH_PY115A.setValue "U_GNMON2", 0, WK_GNSMON '/ 근속월
    
'    If (WK_GNSMON Mod 12) = 0 Then
'        WK_GN2YER = MDC_SetMod.IInt(WK_GNSMON / 12, 1)
'    Else
'        WK_GN2YER = MDC_SetMod.IInt(WK_GNSMON / 12 + 1, 1)
'    End If
    
    WK_GN2YER = oDS_PH_PY115A.GetValue("U_GNMYER", 0) - oDS_PH_PY115A.GetValue("U_GNYER1", 0)
    
    oDS_PH_PY115A.setValue "U_GNYER2", 0, WK_GN2YER '/ 근속년
    
    Compute_GNSDAY = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    oForm.Freeze False
    
    Sbo_Application.StatusBar.SetText "Compute_GNSDAY 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    Compute_GNSDAY = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : Compute_PAY
' 작성자    : 함미경
' Date      :
' Purpose   : 퇴직금 계산 루틴
' 수정자    : 최동권
' 수정일자  : 2009-02-05
' 수정사항  : 1) 프로시저 분리 및 이름변경(Execution -> Compute_PAY)
'             2) 퇴직금 + 세액 계산 -> 퇴직금만 계산
'---------------------------------------------------------------------------------------

Private Function Compute_PAY() As Boolean
On Error GoTo Error_Message
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim MSTCOD      As String
    Dim YEAR        As String

    Dim STRDAT      As String * 8
    Dim ENDDAT      As String * 8
    Dim INPDAT      As String * 8
    Dim OUTDAT      As String * 8
    Dim TMPDAT      As String * 8
    
    Dim ErrNum      As Integer
    Dim iRow        As Long:
    Dim AVRPAY      As Double:  Dim AVRBNS      As Double:  Dim AVRYON      As Double
    Dim AVRDAY      As Double:  Dim MONPAY      As Double
    Dim ExPay2      As Double:  Dim AvExp2
    Dim TOTPAY      As Double:  Dim TOTBNS      As Double:  Dim TOTYON      As Double
    
    Dim WK_GNSYER As Integer:   Dim WK_GNSMON As Integer:   Dim WK_GNSDAY As Integer:   Dim WK_TGSDAY As Integer
    Dim WK_JONYER As Integer:   Dim WK_JONMON As Integer:   Dim WK_JMMMON As Integer:   Dim WK_PAYTYP As String
    Dim WK_GNMYER As Integer:   Dim WK_GNMMON As Integer:   Dim WK_JM2MON As Integer
    Dim WK_YERTJK As Double:    Dim WK_MONTJK As Double:    Dim WK_DAYTJK As Double:    Dim WK_TJKPAY As Double
    Dim WK_ADDRAT As Double:    Dim WK_GITAA2 As Double:    Dim WK_SUDAMT As Double
    Dim X06_YN As Integer
    Dim WK_GNSDAY_M As Integer
    
    Dim SuSilStr  As String:    Dim WK_AMOUNT As Double
    ErrNum = 0

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
  '/ 세액 및 계산식 다시 불러오기
    Call BaseSetting
    
    '/ 귀속연도
    STRDAT = oDS_PH_PY115A.GetValue("U_STRRET", 0)  '/정산시작일
    ENDDAT = oDS_PH_PY115A.GetValue("U_ENDRET", 0)  '/정산종료일
    INPDAT = oDS_PH_PY115A.GetValue("U_INPDAT", 0)  '/입사일자
    OUTDAT = oDS_PH_PY115A.GetValue("U_OUTDAT", 0)  '/퇴사일자
    
    '/ 누진율 적용시
    WK_ADDRAT = Val(oDS_PH_PY115A.GetValue("U_ADDRAT", 0))
    If WK_ADDRAT = 0 Then WK_ADDRAT = 1
    '/ 급여형태~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WK_PAYTYP = MDC_SetMod.Get_ReData("U_PAYTYP", "Code", "[@PH_PY001A]", "'" & Trim$(oDS_PH_PY115A.GetValue("U_MSTCOD", 0)) & "'")
    
    '/ 주현 근속년월일 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WK_GNSYER = Val(oDS_PH_PY115A.GetValue("U_GNSYER", 0))    '/ 주(현) 근속년
    WK_GNSMON = Val(oDS_PH_PY115A.GetValue("U_GNSMON", 0))    '/ 주(현) 근속월
    WK_GNSDAY = Val(oDS_PH_PY115A.GetValue("U_GNSDAY", 0))    '/ 주(현) 근속일
    WK_TGSDAY = Val(oDS_PH_PY115A.GetValue("U_TGSDAY", 0))    '/ 주(현) 총계산일수
    
    '/ 2011.10.04 경보제약 일평균=년평균임금*X09(근속1년미만나머지 일수)/365 추가요청건
    If WK_GNSMON > 0 Then
        TMPDAT = Format(DateAdd("yyyy", WK_GNSYER, Format(STRDAT, "0000-00-00")), "YYYYMMDD")
        WK_GNSDAY_M = MDC_SetMod.TermDay(TMPDAT, ENDDAT)
    Else
        WK_GNSDAY_M = WK_GNSDAY
    End If
    '/ 종전 근속연월일 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WK_JONYER = Val(oDS_PH_PY115A.GetValue("U_JONYER", 0))    '/ 종(전) 근속년
    WK_JONMON = Val(oDS_PH_PY115A.GetValue("U_JONMON", 0))    '/ 종(전) 근속월
    WK_JMMMON = Val(oDS_PH_PY115A.GetValue("U_JMMMON", 0))    '/ 종(전) 중복월수
    WK_JM2MON = Val(oDS_PH_PY115A.GetValue("U_JM2MON", 0))    '/ 종(전) 중복월수2
    '***************************************************************************
    ' 근무년수가 1년 이상이 되어야 하므로 1미만은 퇴직금계산에서 제외된다
    '***************************************************************************
    '대경전장은 1년미만자도 퇴직금나간다네..
    If WK_GNSYER = 0 And WK_JONYER = 0 Then
'        ErrNum = 1
'        GoTo error_Message
    End If
        
    '/ 근무년월일 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    If WK_GNSDAY > 0 Or (WK_GNSMON + WK_JONMON - WK_JMMMON) > 0 Then
       WK_GNMYER = WK_GNSYER + WK_JONYER + 1
    Else
       WK_GNMYER = WK_GNSYER + WK_JONYER
    End If
    If WK_GNSDAY > 0 Then
       WK_GNMMON = WK_GNSYER * 12 + WK_GNSMON + 1
    Else
       WK_GNMMON = WK_GNSYER * 12 + WK_GNSMON
    End If
    
    X06_YN = 0
    '/ 해당월의 마지막 일자
    '/ 평균임금
    AVRPAY = 0:    AVRBNS = 0:    AVRYON = 0:    AVRDAY = 0
    MONPAY = 0:    TOTPAY = 0:    TOTBNS = 0:    TOTYON = 0
    '/ 최근3개월간 급여계산일수
    AVRDAY = AVRDAY + Val(oDS_PH_PY115A.GetValue("U_PAYDY1", 0))
    AVRDAY = AVRDAY + Val(oDS_PH_PY115A.GetValue("U_PAYDY2", 0))
    AVRDAY = AVRDAY + Val(oDS_PH_PY115A.GetValue("U_PAYDY3", 0))
    AVRDAY = AVRDAY + Val(oDS_PH_PY115A.GetValue("U_PAYDY4", 0))
    '/ 최근3개월간 급여
    TOTPAY = TOTPAY + Val(oDS_PH_PY115A.GetValue("U_PAYT01", 0))
    TOTPAY = TOTPAY + Val(oDS_PH_PY115A.GetValue("U_PAYT02", 0))
    TOTPAY = TOTPAY + Val(oDS_PH_PY115A.GetValue("U_PAYT03", 0))
    TOTPAY = TOTPAY + Val(oDS_PH_PY115A.GetValue("U_PAYT04", 0))
    '/ 최근1년간 상여
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST01", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST02", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST03", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST04", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST05", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST06", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST07", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST08", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST09", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST10", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST11", 0))
    TOTBNS = TOTBNS + Val(oDS_PH_PY115A.GetValue("U_BNST12", 0))
    
    TOTYON = Val(oDS_PH_PY115A.GetValue("U_YONCSU", 0))
    
    '//평균체력단련비
    ExPay2 = Val(oDS_PH_PY115A.GetValue("U_ExPay2", 0))
    
    '//임원일경우 누진 근속년 적용
    Dim GNSYER As Integer
    Dim GNSMON As Integer
    Dim GNSDAY As Integer
    Dim ADDRAT As Integer
    Dim NUJINYN As String
    
    sQry = "Exec PH_PY115_03 '" & Trim$(oDS_PH_PY115A.GetValue("U_MSTCOD", 0)) & "','" & Trim$(oDS_PH_PY115A.GetValue("U_ENDRET", 0)) & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        GNSYER = oRecordSet.Fields("YY").VALUE
        GNSMON = oRecordSet.Fields("MM").VALUE
        GNSDAY = 0
        NUJINYN = oRecordSet.Fields("NUJINYN").VALUE
        
    Else
        NUJINYN = "N"
    End If
        
    If NUJINYN = "Y" Then
        ADDRAT = 1
    End If
    
'/////////////////////////////////////////////////////////////////
'// 퇴직급여 계산하기
    For iRow = 1 To 9
        If Trim$(WK_SILSIL(iRow)) <> "" Then '/ 공식이 있으면 공식계산
            SuSilStr = WK_SILSIL(iRow)
            SuSilStr = Replace(SuSilStr, "P01", TOTPAY)     '/ 총급여
            SuSilStr = Replace(SuSilStr, "P02", TOTBNS)     '/ 총상여
            SuSilStr = Replace(SuSilStr, "P03", TOTYON)     '/ 총연차
            SuSilStr = Replace(SuSilStr, "A01", AVRPAY)     '/ 평균급여
            SuSilStr = Replace(SuSilStr, "A02", AVRBNS)     '/ 평균상여
            SuSilStr = Replace(SuSilStr, "A03", AVRYON)     '/ 평균연차
            SuSilStr = Replace(SuSilStr, "P04", ExPay2)     '/ 총체력단련비
            SuSilStr = Replace(SuSilStr, "A04", AvExp2)     '/ 평균체력단련비
            SuSilStr = Replace(SuSilStr, "A05", Format$(MONPAY, "###.0"))  '/ 평균임금
            SuSilStr = Replace(SuSilStr, "A06", WK_YERTJK)  '/ 연퇴직금
            SuSilStr = Replace(SuSilStr, "A07", WK_MONTJK)  '/ 월퇴직금
            SuSilStr = Replace(SuSilStr, "A08", WK_DAYTJK)  '/ 일퇴직금
            SuSilStr = Replace(SuSilStr, "A09", WK_TJKPAY)  '/ 퇴직금계
            SuSilStr = Replace(SuSilStr, "X01", AVRDAY)     '/ 급여기간일수
            SuSilStr = Replace(SuSilStr, "X02", WK_TGSDAY)  '/ 총근속일수
            If NUJINYN = "N" Then
                SuSilStr = Replace(SuSilStr, "X03", WK_GNSYER)  '/ 근속년수
                SuSilStr = Replace(SuSilStr, "X04", WK_GNSMON)  '/ 근속월수
                SuSilStr = Replace(SuSilStr, "X05", WK_GNSDAY)  '/ 근속일수
            Else
                SuSilStr = Replace(SuSilStr, "X03", GNSYER)  '/ 근속년수
                SuSilStr = Replace(SuSilStr, "X04", GNSMON)  '/ 근속월수
                SuSilStr = Replace(SuSilStr, "X05", GNSDAY)  '/ 근속일수
            End If
            X06_YN = InStr(SuSilStr, "X06")
            If NUJINYN = "N" Then
                SuSilStr = Replace(SuSilStr, "X06", WK_ADDRAT)  '/ 누진개월(율)
            Else
                SuSilStr = Replace(SuSilStr, "X06", ADDRAT)  '/ 누진개월(율)
            End If
            
            
            SuSilStr = Replace(SuSilStr, "X07", WK_TGSDAY)  '/ 근속년수계산(2)
            SuSilStr = Replace(SuSilStr, "X08", WK_PAYTYP)  '/ 급여형태
            SuSilStr = Replace(SuSilStr, "X09", WK_GNSDAY_M) '/ 근속일수(1년미만 일수)
            
            Select Case iRow
            Case 5:  oDS_PH_PY115A.setValue "U_AVRSIL", 0, SuSilStr
            Case 6:  oDS_PH_PY115A.setValue "U_YERSIL", 0, SuSilStr
            Case 7:  oDS_PH_PY115A.setValue "U_MONSIL", 0, SuSilStr
            Case 8:  oDS_PH_PY115A.setValue "U_DAYSIL", 0, SuSilStr
            End Select
            '/ 금액산출
            WK_AMOUNT = 0
            sQry = "SELECT 1.0 * " & SuSilStr
            oRecordSet.DoQuery sQry
            If oRecordSet.RecordCount <> 0 Then
                WK_AMOUNT = MDC_SetMod.RInt(oRecordSet.Fields(0).VALUE, WK_LENGTH(iRow), WK_ROUNDT(iRow))
            Else
                WK_AMOUNT = 0
            End If
            Select Case iRow
            Case 1: AVRPAY = WK_AMOUNT
            Case 2: AVRBNS = WK_AMOUNT
            Case 3: AVRYON = WK_AMOUNT
            Case 4: AvExp2 = WK_AMOUNT
            Case 5: MONPAY = WK_AMOUNT
            Case 6: WK_YERTJK = WK_AMOUNT
            Case 7: WK_MONTJK = WK_AMOUNT
            Case 8: WK_DAYTJK = WK_AMOUNT
            Case 9: '// 퇴직금 = 임원누진율사용과 누진적용계산식이 없을경우 누진율자동으로 아니면 수식그대로 사용.
                    If Trim$(RATUSE) = "Y" And Val(X06_YN) = 0 Then
                        If NUJINYN = "N" Then
                            WK_TJKPAY = WK_AMOUNT * WK_ADDRAT '2009.06.12 포레스코 계산식에누진적용율조건별사용으로 중복누진율적용됨.
                        Else
                            WK_TJKPAY = WK_AMOUNT * ADDRAT
                        End If
                    Else
                        WK_TJKPAY = WK_AMOUNT
                    End If
            
            End Select
            
        End If
    Next iRow
    '/ 기타수당2=미도래연차수당등..
    WK_GITAA2 = Val(oDS_PH_PY115A.GetValue("U_GITAA2", 0))
    WK_SUDAMT = Val(oDS_PH_PY115A.GetValue("U_SUDAMT", 0))
    WK_TJKPAY = WK_TJKPAY + WK_GITAA2
    '/ 기타수당1=퇴직금계산에는 제외하고 퇴직금산정서에 포함만 하는 금액임.
    

    '***************************************************************************
    ' 결과값 UPDATE
    '***************************************************************************
    oForm.Freeze True
    oDS_PH_PY115A.setValue "U_GNMYER", 0, WK_GNMYER
    oDS_PH_PY115A.setValue "U_GNMMON", 0, WK_GNMMON
    oDS_PH_PY115A.setValue "U_ADDRAT", 0, WK_ADDRAT
    oDS_PH_PY115A.setValue "U_AVRPAY", 0, AVRPAY      '/ 평균급여
    oDS_PH_PY115A.setValue "U_AVRBNS", 0, AVRBNS      '/ 평균상여
    oDS_PH_PY115A.setValue "U_AVRYON", 0, AVRYON      '/ 평균연차
    oDS_PH_PY115A.setValue "U_AvExp2", 0, AvExp2      '/ 평균체력단련비
    oDS_PH_PY115A.setValue "U_MONPAY", 0, MONPAY      '/ 평균임금
    
    oDS_PH_PY115A.setValue "U_YERTJK", 0, WK_YERTJK   '/ 년퇴직금
    oDS_PH_PY115A.setValue "U_MONTJK", 0, WK_MONTJK   '/ 월퇴직금
    oDS_PH_PY115A.setValue "U_DAYTJK", 0, WK_DAYTJK   '/ 일퇴직금
    oDS_PH_PY115A.setValue "U_TJKPAY", 0, WK_TJKPAY   '/ 퇴직금
    
    YEAR = Left(Trim(Str(ENDDAT)), 4)
    MSTCOD = oDS_PH_PY115A.GetValue("U_MSTCOD", 0)
    sQry = "Select b.U_Amt From [@PH_PY125A] a inner Join [@PH_PY125B] b On a.Code = b.Code And a.U_YEAR = '" & YEAR & "' And b.U_MSTCOD = '" & MSTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        oDS_PH_PY115A.setValue "U_TJKPAY01", 0, oRecordSet.Fields(0).VALUE   '/ 퇴직연금액
        oDS_PH_PY115A.setValue "U_TJKPAY02", 0, WK_TJKPAY - oRecordSet.Fields(0).VALUE   '/ 퇴직연금액
    End If
        
    
    oForm.Freeze False
    oForm.Update

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Compute_PAY = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    oForm.Freeze False
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "근속년수가 1년미만이므로, 퇴직금계산이 불가합니다!", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "퇴직급여총액이 0입니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Compute_PAY 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    Compute_PAY = False
End Function

Private Function Compute_TAX() As Boolean
On Error GoTo Error_Message
    Dim ErrNum  As Integer
    
    Dim STRDAT  As String * 8
    Dim JINDAT  As String * 8
    Dim JOTDAT  As String * 8
    Dim INPDAT  As String * 8
    Dim ST2RET  As String * 8

    Dim WT_RTXGBN_H As String
    Dim WT_RTXGBN_J As String
    Dim WT_RTXYER_1 As Integer
    Dim WT_RTXYER_2 As Integer
    
    '/ 근속년수
    Dim WT_GNMYER As Integer
    Dim WT_JONYER As Integer
    Dim WT_JMMMON As Integer
    Dim WT_JMMYER As Integer
    Dim WT_JDOYER As Integer
    
    '/ 퇴직급여
    Dim WT_BHMAMT As Double:
    Dim WT_TJKPAY As Double: Dim WT_SUDAMT As Double: Dim WT_YILPA1 As Double: Dim WT_YILPA2 As Double
    Dim WT_RETPAY As Double: Dim WT_RETPA1 As Double: Dim WT_RETPA2 As Double
    Dim WT_JRET01 As Double: Dim WT_JSUD01 As Double: Dim WT_JYIL01 As Double: Dim WT_JYIL03 As Double: Dim WT_JYIL04 As Double
    Dim WT_JTOT01 As Double: Dim WT_JTOT02 As Double: Dim WT_BTXAMT As Double
    
    '/ 세액계산
    Dim WT_SPCGON As Double:   Dim WT_SODGON As Double: Dim WT_RETGON As Double
    Dim WT_TAXSTD As Double:   Dim WT_YTXSTD As Double: Dim WT_YSANTX As Double
    Dim WT_SANTAX As Double:   Dim WT_RTXGON As Double: Dim WT_FRNGON As Double: Dim WT_FRNGO1 As Double: Dim WT_FRNGO2 As Double
    Dim WT_GULGAB As Double:   Dim WT_GULJUM As Double:
    
    '/ 세액환산정보
    Dim D_SHWJIG(1 To 4) As Double   '/ 세액환산지급예상액
    Dim D_SHWYAS(1 To 4) As Double   '/ 세액환산-환산연평균산출세액
    Dim D_SHWSUR(1 To 4) As Double   '/ 세액환산-수령액
    
    Dim D_RTXGON(1 To 4) As Double   '/ 퇴직소득세액공제
    Dim D_RTXKUM(1 To 4) As Double   '/ 퇴직소득세액공제한도
    Dim D_RTXSAN(1 To 4) As Double   '/ 산출세액
    Dim D_GULGAB(1 To 2) As Double
    Dim i As Integer
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
'/ Clear
    ErrNum = 0
    For i = 1 To 4
        D_SHWJIG(i) = 0
        D_SHWYAS(i) = 0
        D_SHWSUR(i) = 0
        
        D_RTXGON(i) = 0
        D_RTXKUM(i) = 0
        D_RTXSAN(i) = 0
    Next i

    '/ 귀속연도
    STRDAT = oDS_PH_PY115A.GetValue("U_STRRET", 0)    '/ 주현근무지 정산시작일
    INPDAT = oDS_PH_PY115A.GetValue("U_INPDAT", 0)    '/ 주현근무지 입사일자
    JINDAT = oDS_PH_PY115A.GetValue("U_JINDAT", 0)    '/ 종전근무지 입사일자
    JOTDAT = oDS_PH_PY115A.GetValue("U_JOTDAT", 0)    '/ 종전근무지 퇴사일자
    ST2RET = oDS_PH_PY115A.GetValue("U_ST2RET", 0)    '/ 종전근무지 중도정산일
     '/ 퇴직세액공제여부
     WT_RTXGBN_H = oDS_PH_PY115A.GetValue("U_SPCGBN", 0)   '/ 2009년주현퇴직소득세액공제여부(Y/N)
     WT_RTXGBN_J = oDS_PH_PY115A.GetValue("U_JS1SPC", 0)   '/ 2009년 종전퇴직소득세액공제여부(Y/N)
'///************************************************************************************************************/
'/// 1.근속년월일 계산 /
'///************************************************************************************************************/
     Call Compute_GNMYER
     '/ 세법상 근무년수 주(현)  :18년
     WT_GNMYER = Val(oDS_PH_PY115A.GetValue("U_GNMMON", 0)) - Val(oDS_PH_PY115A.GetValue("U_EXPMON", 0))
     If (WT_GNMYER Mod 12) = 0 Then
        WT_GNMYER = WT_GNMYER / 12
     Else
        WT_GNMYER = MDC_SetMod.IInt(WT_GNMYER / 12, 1) + 1
     End If
     '/ 세법상 근무년수 종(전)  :2년
     WT_JONYER = Val(oDS_PH_PY115A.GetValue("U_GNMDAY", 0)) - Val(oDS_PH_PY115A.GetValue("U_JEXMON", 0))
     If (WT_JONYER Mod 12) = 0 Then
        WT_JONYER = WT_JONYER / 12
     Else
        WT_JONYER = MDC_SetMod.IInt(WT_JONYER / 12, 1) + 1
     End If
     '/ 세법상 중복월수
     WT_JMMMON = Val(oDS_PH_PY115A.GetValue("U_JMMMON", 0))   '/ 세법상 중복월수         :25년
     '/ 세법상 중복년수
     WT_JMMYER = WT_JMMMON / 12    '/ 세법상 중복년수
     If (WT_JMMMON Mod 12) = 0 Then
        WT_JMMYER = WT_JMMMON / 12
     Else
        WT_JMMYER = MDC_SetMod.IInt(WT_JMMMON / 12, 1) + 1
     End If
     '/ 중도정산 년도
     If INPDAT <> STRDAT Then
        Call MDC_SetMod.Term2(INPDAT, STRDAT)
        WT_JDOYER = Val(ZPAY_GBL_GNSYER) * 12 + Val(ZPAY_GBL_GNSMON)
        If Val(ZPAY_GBL_GNSDAY) > 0 Then WT_JDOYER = WT_JDOYER + 1
        If (WT_JDOYER Mod 12) = 0 Then
           WT_JDOYER = WT_JDOYER / 12
        Else
           WT_JDOYER = MDC_SetMod.IInt(WT_JDOYER / 12, 1) + 1
        End If
     End If
    oDS_PH_PY115A.setValue "U_GN3YER", 0, WT_JDOYER
    '// 종전근무지에 종전입사일자와 주현의 입사일자가 동일할경우 법정외근속월수도 주현과 동일하게 변경함.
    If Left(JOTDAT, 4) = Left(STRDAT, 4) Then
        oDS_PH_PY115A.setValue "U_GN2MON", 0, oDS_PH_PY115A.GetValue("U_GNMMON", 0)
    End If
  
 '/ 세액 및 계산식 다시 불러오기(정산종료연도를 바꾸거나 네비게이션사용시 폼로드년도 다를수있슴)
    Call BaseSetting

'///************************************************************************************************************/
'/// 2. 퇴직금 지급내역 //
'///************************************************************************************************************/
   '/ 현근무지 소득(법정)
    WT_TJKPAY = Val(oDS_PH_PY115A.GetValue("U_TJKPAY", 0))    '/ 주현(법정) 퇴직급여
    WT_BHMAMT = Val(oDS_PH_PY115A.GetValue("U_BHMAMT", 0))    '/ 주현(법정) 단체보험료
    WT_YILPA1 = Val(oDS_PH_PY115A.GetValue("U_YILPA1", 0))    '/ 주현(법정) 퇴직연금일시금
    '/ 현근무지 소득(법정외)
    WT_SUDAMT = Val(oDS_PH_PY115A.GetValue("U_SUDAMT", 0))    '/ 주현(법정외) 명예퇴직수당
    WT_YILPA2 = Val(oDS_PH_PY115A.GetValue("U_YILPA2", 0))    '/ 주현(법정외) 퇴직연금일시금
    
    '/ 전근무지 소득(법정)
    WT_JRET01 = Val(oDS_PH_PY115A.GetValue("U_JRET01", 0))    '/ 종전(법정) 퇴직급여
    WT_JYIL03 = Val(oDS_PH_PY115A.GetValue("U_JYIL03", 0))    '/ 종전(법정) 퇴직연금일시금
    WT_JTOT01 = Val(oDS_PH_PY115A.GetValue("U_JTOT01", 0))    '/ 종전(법정) 계
    
    '/ 전근무지 소득 (법정외)
    WT_JSUD01 = Val(oDS_PH_PY115A.GetValue("U_JADD01", 0))    '/ 종전(법정외) 명예퇴직수당
    WT_JYIL04 = Val(oDS_PH_PY115A.GetValue("U_JYIL04", 0))    '/ 종전(법정외) 퇴직연금일시금
    WT_JTOT02 = Val(oDS_PH_PY115A.GetValue("U_JTOT02", 0))    '/ 종전(법정외) 계
    
    
    WT_RETPA1 = WT_TJKPAY + WT_BHMAMT + WT_YILPA1 + WT_JTOT01   '/ 법정퇴직금  계
    WT_RETPA2 = WT_SUDAMT + WT_YILPA2 + WT_JTOT02                         '/ 법정외퇴직금계
    '/ 퇴직급여액[퇴직급여+명예퇴직수당+단체퇴직수당+종전퇴직금](21) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WT_RETPAY = WT_RETPA1 + WT_RETPA2
    
    '/ 비과세계 (주현+종전)
    WT_BTXAMT = Val(oDS_PH_PY115A.GetValue("U_BTXPAY", 0)) + Val(oDS_PH_PY115A.GetValue("U_BTXP01", 0))

'///************************************************************************************************************/
'/// 2. 체크 옵션
'///************************************************************************************************************/
    '/ 명예퇴직금처리여부(조건: 퇴직정산년도이전 년도에 중도정산한 사원의 퇴직정산시 법정외퇴직금이 존재할 경우
    '/ 베루(중도정산임에도 불구하고, 명예퇴직수당을 지급
    'If Left$(INPDAT, 4) < Left$(STRDAT, 4) And (WT_SUDAMT + WT_JSUD01) > 0 And oDS_PH_PY115A.GetValue("U_JSNGBN", 0) = "1" Then
    If Left$(INPDAT, 4) < Left$(STRDAT, 4) And (WT_SUDAMT + WT_JSUD01) > 0 Then
        BubJung_YN = "Y"
    Else
        BubJung_YN = "N"
    End If
'/ 주현근무지의 입사일자와 정산시작일이 다를경우: 중도정산을 한경우
    If (INPDAT <> STRDAT And Left(JOTDAT, 4) = Left(STRDAT, 4)) Or (Trim(ST2RET) = "" And Trim(JOTDAT) <> "") Then
        BubJung_YN = "N"
    End If

    oDS_PH_PY115A.setValue "U_BUBCHK", 0, BubJung_YN
    oForm.Items("BUBCHK").Update
    
    If (WT_RETPA1 + WT_SUDAMT + WT_BTXAMT) = 0 Then
        Sbo_Application.StatusBar.SetText "주(현) 퇴직급여가 0 입니다. 퇴직금 계산을 먼저 하시기 바랍니다. ", bmt_Short, smt_Error
        Exit Function
    End If
    
            
'///************************************************************************************************************/
'/// 3. 퇴직연금내역 //
     '/ 퇴직연금에 가입한 근로자가 퇴직연금일시금을 지급받는 경우에 한하여 입력함.
'///************************************************************************************************************/
    
    '/ 1. 퇴직연금명세-퇴직연금일시금 계산(주현, 종전)
    Call Display_ilSiGum

'///************************************************************************************************************/
'/// 4.세액환산명세 //
    '/ 퇴직연금을 일시금과 연금으로 나누어 수령하는 경우
    '/ 퇴직연금을 수령하던 자가 중도해지등으로 일시금으로 지급받는 경우
    '/ 퇴직소득을 과세이연한 경우 등에 한하여 입력함.
'///************************************************************************************************************/
        '// 세액환산명세 입력받은 값
        WT_SHWTOT = 0
        For i = 1 To 4
            SHWJIG(i) = Val(oDS_PH_PY115A.GetValue("U_SH" & i & "JIG", 0))    '// 퇴직연금일시금지급예상액
            SHWIYO(i) = Val(oDS_PH_PY115A.GetValue("U_SH" & i & "IYO", 0))    '// 과세이연금액
            SHWGSU(i) = Val(oDS_PH_PY115A.GetValue("U_SH" & i & "GSU", 0))    '// 기수령한퇴직급여액
            WT_SHWTOT = WT_SHWTOT + SHWJIG(i) + SHWIYO(i) + SHWGSU(i)
        Next i
        WT_FRNGO1 = Val(oDS_PH_PY115A.GetValue("U_JS1TAX", 0))  '// 외국납부세액(법정)
        WT_FRNGO2 = Val(oDS_PH_PY115A.GetValue("U_JS2TAX", 0))  '// 외국납부세액(법정외)
        WT_FRNGON = WT_FRNGO1 + WT_FRNGO2
'        '//수령가능금액(법정)=
'Private SHWTIL(1 To 2) As Double
'Private SHWSUR(1 To 2) As Double
'Private SHWGON(1 To 2) As Double
'Private SHWGWA(1 To 2) As Double
'Private SHWYAG(1 To 2) As Double
'Private SHWYAS(1 To 2) As Double
'Private SHWSAN(1 To 2) As Double
        '// 세액환산명세가 있을경우
        If WT_SHWTOT > 0 Then
             '/ 2.1) 세액환산명세-주(현)
             Call Display_TaxWhanSan(3, BubJung_YN)
             D_SHWJIG(3) = HS.SHWJIG
             D_SHWYAS(3) = HS.SHWYAS
             D_SHWSUR(3) = HS.SHWSUR
        
             '/ 2.2) 세액환산명세-종(전)
             Call Display_TaxWhanSan(4, BubJung_YN)
             D_SHWJIG(4) = HS.SHWJIG
             D_SHWYAS(4) = HS.SHWYAS
             D_SHWSUR(4) = HS.SHWSUR
         
             '/ 2.3) 세액환산명세-법정퇴직
             Call Display_TaxWhanSan(1, BubJung_YN)
             D_SHWJIG(1) = HS.SHWJIG
             D_SHWYAS(1) = HS.SHWYAS
             D_SHWSUR(1) = HS.SHWSUR
             
             '/ 2.4) 세액환산명세-법정퇴직이외
             Call Display_TaxWhanSan(2, BubJung_YN)
             D_SHWJIG(2) = HS.SHWJIG
             D_SHWYAS(2) = HS.SHWYAS
             D_SHWSUR(2) = HS.SHWSUR
            '***************************************************************************
            ' 결과값 UPDATE
            '***************************************************************************
            
            oDS_PH_PY115A.setValue "U_SH2GON", 0, WK401.RETGON   '/ 퇴직소득공제
            oDS_PH_PY115A.setValue "U_SH2GWA", 0, WK401.TAXSTD   '/ 퇴직소득과세표준
            oDS_PH_PY115A.setValue "U_SH2YAG", 0, WK401.YTXSTD   '/ 연평균과세표준
            oDS_PH_PY115A.setValue "U_SH2YAS", 0, WK401.YSANTX   '/ 연평균산출세액
            oDS_PH_PY115A.setValue "U_JS2SAN", 0, WK401.SANTAX   '/ 산출세액
            oDS_PH_PY115A.setValue "U_JS2GON", 0, WK401.RTXGON   '/ 퇴직소득세액공제
        
        End If

'///************************************************************************************************************/
    ' 퇴직소득세액  계산
'///************************************************************************************************************/
    '/ 근무년수(법정)
        WT_RTXYER_1 = IIf(WT_GNMYER < WT_JONYER, WT_JONYER, WT_GNMYER)
    '/ 근무년수(법정외)
        WT_RTXYER_2 = Val(oDS_PH_PY115A.GetValue("U_GN2YER", 0))
    '/ 주현근무지의 입사일자와 정산시작일이 다를경우: 중도정산을 한경우
        If BubJung_YN = "N" Then
            WT_RETPA1 = WT_RETPA1 + WT_RETPA2       '// 동일년도중도정산은 종전으로 처리함.모두 법정퇴직금처리함.
            WT_RETPA2 = 0
        End If
    '//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    '// 법정퇴직금
    '// 종전이나 주현에서 중도정산을 하지 않을 경우는 모두 법정퇴직금
    '//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        Call Compute_TAX_2(WT_RETPA1, WT_RTXYER_1, 0, D_SHWJIG(1), D_SHWYAS(1), D_SHWSUR(1), "BB")
        
        oDS_PH_PY115A.setValue "U_JS1RET", 0, WT_RETPA1       '/ 퇴직급여총액
        oDS_PH_PY115A.setValue "U_JS1GON", 0, WK401.RETGON    '/ 퇴직소득공제
        oDS_PH_PY115A.setValue "U_JS1GRATE", 0, WK401.SPCGON    '/ 퇴직소득정율공제         '/2013/12/25 노근용추가
        oDS_PH_PY115A.setValue "U_JS1GYEAR", 0, WK401.SODGON    '/ 퇴직소득근속년수공제     '/2013/12/25 노근용추가
        oDS_PH_PY115A.setValue "U_JS1STD", 0, WK401.TAXSTD    '/ 퇴직소득과세표준
        
        
        
        If Left$(oDS_PH_PY115A.GetValue("U_ENDINT", 0), 4) >= "2016" Then
        
            '2016년 이후 계산
            oDS_PH_PY115A.setValue "U_JS2GYEAR", 0, WK401.SODGON    '/ 퇴직소득근속년수공제     '/2016/1/6 노근용추가
            '/ 환산급여     '/2016/1/6 노근용추가
            oDS_PH_PY115A.setValue "U_JS2AMT", 0, (WT_RETPA1 - WK401.SODGON) / Val(oDS_PH_PY115A.GetValue("U_GNMYER", 0)) * 12
                
            '/2016년 환산급여별 공제
            oDS_PH_PY115A.setValue "U_JS2GAMT", 0, WK401.WSGGON
            '//2016년 퇴직소득과세표준
            oDS_PH_PY115A.setValue "U_JS3STD", 0, oDS_PH_PY115A.GetValue("U_JS2AMT", 0) - oDS_PH_PY115A.GetValue("U_JS2GAMT", 0)
                
            
            Dim FD_JS3STD As Double
            '//2016년 이후 환산산출세액
            FD_JS3STD = Val(oDS_PH_PY115A.GetValue("U_JS3STD", 0))
            For i = 51 To 59
                If FD_JS3STD <= WG01.TB1AMT(i) Then
                   FD_JS3STD = MDC_SetMod.IInt((FD_JS3STD - WG01.TB1GON(i)) * (WG01.TB1RAT(i) / 100) + WG01.TB1KUM(i), 1)
                   Exit For
                End If
            Next i
            
            
            oDS_PH_PY115A.setValue "U_JS2HYST3", 0, FD_JS3STD    '/ 환산산출세액(2016/01/01 이후)
            
            '/ 산출세액(2016/01/01 이후)
            oDS_PH_PY115A.setValue "U_JS2SAN03", 0, oDS_PH_PY115A.GetValue("U_JS2HYST3", 0) / 12 * Val(oDS_PH_PY115A.GetValue("U_GNMYER", 0))
            
            '//퇴직과세연도(2016년 이후)
            oDS_PH_PY115A.setValue "U_TAXYEAR", 0, Left(Trim(Str(oDS_PH_PY115A.GetValue("U_ENDRET", 0))), 4)
            
        End If
            
            
        
        '/ 퇴직소득과세표준 안분 (2012/12/31이전, 이후) '/2013/12/25 노근용추가
        oDS_PH_PY115A.setValue "U_JS1STD01", 0, IInt(WK401.TAXSTD * Val(oDS_PH_PY115A.GetValue("U_GNYER1", 0)) / (Val(oDS_PH_PY115A.GetValue("U_GNYER1", 0)) + Val(oDS_PH_PY115A.GetValue("U_GNYER2", 0))), 1)
        oDS_PH_PY115A.setValue "U_JS1STD02", 0, WK401.TAXSTD - Val(oDS_PH_PY115A.GetValue("U_JS1STD01", 0))
        
        oDS_PH_PY115A.setValue "U_JS1YSD", 0, WK401.YTXSTD    '/ 연평균과세표준
        
        '/ 연평균과세표준 안분 (2012/12/31이전, 이후)  '/2013/12/25 노근용추가
        If oDS_PH_PY115A.GetValue("U_GNYER1", 0) > 0 Then
            oDS_PH_PY115A.setValue "U_JS1YSD01", 0, IInt(Val(oDS_PH_PY115A.GetValue("U_JS1STD01", 0)) / Val(oDS_PH_PY115A.GetValue("U_GNYER1", 0)), 1)
        Else
            oDS_PH_PY115A.setValue "U_JS1YSD01", 0, 0
        End If
        
        oDS_PH_PY115A.setValue "U_JS1YSD02", 0, IInt(Val(oDS_PH_PY115A.GetValue("U_JS1STD02", 0)) / Val(oDS_PH_PY115A.GetValue("U_GNYER2", 0)), 1)
        
        '/ 환산과세표준(2013/01/01 이후)
        oDS_PH_PY115A.setValue "U_JS1HYTD2", 0, Val(oDS_PH_PY115A.GetValue("U_JS1YSD02", 0)) * 5
        
        
'

        Dim FD_i        As Integer
        Dim FD_JSHYTD   As Double
        
        FD_JSHYTD = Val(oDS_PH_PY115A.GetValue("U_JS1HYTD2", 0))
        For FD_i = 51 To 59
            If FD_JSHYTD <= WG01.TB1AMT(FD_i) Then
               FD_JSHYTD = MDC_SetMod.IInt((FD_JSHYTD - WG01.TB1GON(FD_i)) * (WG01.TB1RAT(FD_i) / 100) + WG01.TB1KUM(FD_i), 1)
               Exit For
            End If
        Next FD_i
        
        
        oDS_PH_PY115A.setValue "U_JS1HYST2", 0, FD_JSHYTD    '/ 환산산출세액(2013/01/01 이후)
        
        
        
        
        
        
        
        
        '//2013/01/01 이후 근속이 있으면
        If oDS_PH_PY115A.GetValue("U_GNMON1", 0) > 0 Then
            oDS_PH_PY115A.setValue "U_JS1YST01", 0, WK401.YSANTX    '/ 연평균산출세액(2013/01/01 이전)
        Else
            oDS_PH_PY115A.setValue "U_JS1YST01", 0, 0    '/ 연평균산출세액(2013/01/01 이전)
        End If
        
        oDS_PH_PY115A.setValue "U_JS1YST02", 0, Int(FD_JSHYTD / 5)   '/ 연평균산출세액(2013/01/01 이후)
        'oDS_PH_PY115A.setValue "U_JS1YST02", 0, Round(FD_JSHYTD / 5, 0)   '/ 연평균산출세액(2013/01/01 이후)
        
        oDS_PH_PY115A.setValue "U_JS1SAN01", 0, Val(oDS_PH_PY115A.GetValue("U_JS1YST01", 0)) * Val(oDS_PH_PY115A.GetValue("U_GNYER1", 0))
        oDS_PH_PY115A.setValue "U_JS1SAN02", 0, Val(oDS_PH_PY115A.GetValue("U_JS1YST02", 0)) * Val(oDS_PH_PY115A.GetValue("U_GNYER2", 0))
        
        oDS_PH_PY115A.setValue "U_JS1YST", 0, Val(oDS_PH_PY115A.GetValue("U_JS1YST01", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS1YST02", 0))
         oDS_PH_PY115A.setValue "U_JS1SAN", 0, Val(oDS_PH_PY115A.GetValue("U_JS1SAN01", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS1SAN02", 0))
'        oDS_PH_PY115A.setValue "U_JS1YST", 0, WK401.YSANTX    '/ 연평균산출세액
'        oDS_PH_PY115A.setValue "U_JS1SAN", 0, WK401.SANTAX    '/ 산출세액
        oDS_PH_PY115A.setValue "U_JS1STX", 0, WK401.RTXGON    '/ 퇴직소득세액공제
        
        D_GULGAB(1) = MDC_SetMod.IInt(Val(oDS_PH_PY115A.GetValue("U_JS1SAN", 0)) - WT_FRNGO1 - WK401.RTXGON, 1)

    '//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    '// 법정외퇴직금
    '//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    
        
        
        Call Compute_TAX_2(WT_RETPA2, WT_RTXYER_2, 0, D_SHWJIG(2), D_SHWYAS(2), D_SHWSUR(2), "BE")
        
        oDS_PH_PY115A.setValue "U_JS2RET", 0, WT_RETPA2       '/ 퇴직급여총액
        oDS_PH_PY115A.setValue "U_JS2GON", 0, WK401.RETGON    '/ 퇴직소득공제
        oDS_PH_PY115A.setValue "U_JS2STD", 0, WK401.TAXSTD    '/ 퇴직소득과세표준
        oDS_PH_PY115A.setValue "U_JS2YSD", 0, WK401.YTXSTD    '/ 연평균과세표준
        oDS_PH_PY115A.setValue "U_JS2YST", 0, WK401.YSANTX    '/ 연평균산출세액
        oDS_PH_PY115A.setValue "U_JS2SAN", 0, WK401.SANTAX    '/ 산출세액
        oDS_PH_PY115A.setValue "U_JS2STX", 0, WK401.RTXGON    '/ 퇴직소득세액공제
            
'        D_GULGAB(2) = MDC_SetMod.IInt(WK401.SANTAX - WT_FRNGO2 - WK401.RTXGON, 10)
        D_GULGAB(2) = MDC_SetMod.IInt(WK401.SANTAX - WT_FRNGO2 - WK401.RTXGON, 1)
     '/ 3.퇴직급여
         WT_RETPAY = Val(oDS_PH_PY115A.GetValue("U_JS1RET", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS2RET", 0))
     '/ 4. 퇴직소득공제(35 = ①특별공제 + ②소득공제) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
         WT_RETGON = Val(oDS_PH_PY115A.GetValue("U_JS1GON", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS2GON", 0))
     '/ 5. 퇴직소득과세표준(36 = 34 - 35) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
         WT_TAXSTD = Val(oDS_PH_PY115A.GetValue("U_JS1STD", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS2STD", 0))
     '/ 6. 연평균과세표준(37 = 36 / 근속년수) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
         WT_YTXSTD = Val(oDS_PH_PY115A.GetValue("U_JS1YSD", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS2YSD", 0))
     '/ 7. 연평균산출세액(38) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
         WT_YSANTX = Val(oDS_PH_PY115A.GetValue("U_JS1YST", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS2YST", 0))
     '/ 8. 산출세액(39 = 38 * 근속년수) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
         WT_SANTAX = Val(oDS_PH_PY115A.GetValue("U_JS1SAN", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS2SAN", 0))
     '/ 9. 퇴직세액공제 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
         WT_RTXGON = Val(oDS_PH_PY115A.GetValue("U_JS1STX", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS2STX", 0))
     '/ 9. 퇴직세액공제 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
         WT_FRNGON = Val(oDS_PH_PY115A.GetValue("U_JS1TAX", 0)) + Val(oDS_PH_PY115A.GetValue("U_JS2TAX", 0))
     '/ 결정세액 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
         WT_GULGAB = D_GULGAB(1) + D_GULGAB(2)
     '/ 결정주민세 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WT_GULJUM = MDC_SetMod.IInt((WT_GULGAB * 0.1), 1)
        'WT_GULJUM = MDC_SetMod.IInt((WT_GULGAB * 0.1), 10)
     
    '***************************************************************************
    ' 결과값 UPDATE
    '***************************************************************************
        oDS_PH_PY115A.setValue "U_RETPAY", 0, WT_RETPAY   '/ 퇴직급여총액
        oDS_PH_PY115A.setValue "U_RETGON", 0, WT_RETGON   '/ 퇴직소득공제
        oDS_PH_PY115A.setValue "U_TAXSTD", 0, WT_TAXSTD   '/ 퇴직소득과세표준
        oDS_PH_PY115A.setValue "U_YTXSTD", 0, WT_YTXSTD   '/ 연평균과세표준
        oDS_PH_PY115A.setValue "U_YSANTX", 0, WT_YSANTX   '/ 연평균산출세액
        oDS_PH_PY115A.setValue "U_SANTAX", 0, WT_SANTAX   '/ 산출세액
        oDS_PH_PY115A.setValue "U_SPCGON", 0, WT_RTXGON   '/ 퇴직소득세액공제
        oDS_PH_PY115A.setValue "U_TAXGON", 0, WT_FRNGON   '/ 외국납부
        
        If Left$(oDS_PH_PY115A.GetValue("U_ENDINT", 0), 4) < "2016" Then
            oDS_PH_PY115A.setValue "U_GULGAB", 0, WT_GULGAB   '/ 소득세
            oDS_PH_PY115A.setValue "U_GULJUM", 0, WT_GULJUM   '/ 주민세
            
            If Trim(oDS_PH_PY115A.GetValue("U_MYGACC", 0)) <> "" And oDS_PH_PY115A.GetValue("U_IPGM01", 0) <> 0 Then
                oDS_PH_PY115A.setValue "U_GWASEY", 0, Round(oDS_PH_PY115A.GetValue("U_SANTAX", 0) * oDS_PH_PY115A.GetValue("U_IPGM01", 0) / oDS_PH_PY115A.GetValue("U_TJKPAY", 0), 0)
                oDS_PH_PY115A.setValue "U_SEYGAB", 0, oDS_PH_PY115A.GetValue("U_GWASEY", 0)
                oDS_PH_PY115A.setValue "U_SEYJUM", 0, MDC_SetMod.IInt((oDS_PH_PY115A.GetValue("U_SEYGAB", 0) * 0.1), 1)   '/ 주민세
            End If
        Else
        
        
            '2016년 이후 특례적용 산출세액
            
            Select Case Trim(oDS_PH_PY115A.GetValue("U_TAXYEAR", 0))
                Case "2016"
                    oDS_PH_PY115A.setValue "U_SPSANTX", 0, (oDS_PH_PY115A.GetValue("U_SANTAX", 0) * 0.8) + (oDS_PH_PY115A.GetValue("U_JS2SAN03", 0) * 0.2)
                Case "2017"
                    oDS_PH_PY115A.setValue "U_SPSANTX", 0, (oDS_PH_PY115A.GetValue("U_SANTAX", 0) * 0.6) + (oDS_PH_PY115A.GetValue("U_JS2SAN03", 0) * 0.4)
                Case "2018"
                    oDS_PH_PY115A.setValue "U_SPSANTX", 0, (oDS_PH_PY115A.GetValue("U_SANTAX", 0) * 0.4) + (oDS_PH_PY115A.GetValue("U_JS2SAN03", 0) * 0.6)
                Case "2019"
                    oDS_PH_PY115A.setValue "U_SPSANTX", 0, (oDS_PH_PY115A.GetValue("U_SANTAX", 0) * 0.2) + (oDS_PH_PY115A.GetValue("U_JS2SAN03", 0) * 0.8)
                Case Else
                    oDS_PH_PY115A.setValue "U_SPSANTX", 0, oDS_PH_PY115A.GetValue("U_JS2SAN03", 0)
            
            End Select
            
            '//신고대상 세액
            oDS_PH_PY115A.setValue "U_SANTAX1", 0, oDS_PH_PY115A.GetValue("U_SPSANTX", 0)
            
            '//과세이연계좌 입력시 이연퇴직 소득세계산
            
            If Trim(oDS_PH_PY115A.GetValue("U_MYGACC", 0)) <> "" And oDS_PH_PY115A.GetValue("U_IPGM01", 0) <> 0 Then
                oDS_PH_PY115A.setValue "U_GWASEY", 0, Round(oDS_PH_PY115A.GetValue("U_SANTAX1", 0) * oDS_PH_PY115A.GetValue("U_IPGM01", 0) / oDS_PH_PY115A.GetValue("U_TJKPAY", 0), 0)
                oDS_PH_PY115A.setValue "U_SEYGAB", 0, oDS_PH_PY115A.GetValue("U_GWASEY", 0)
                oDS_PH_PY115A.setValue "U_SEYJUM", 0, MDC_SetMod.IInt((oDS_PH_PY115A.GetValue("U_SEYGAB", 0) * 0.1), 1)   '/ 주민세
            End If
            
            
            oDS_PH_PY115A.setValue "U_GULGAB", 0, oDS_PH_PY115A.GetValue("U_SANTAX1", 0)   '/ 소득세
            oDS_PH_PY115A.setValue "U_GULJUM", 0, MDC_SetMod.IInt((oDS_PH_PY115A.GetValue("U_SANTAX1", 0) * 0.1), 1)   '/ 주민세
        End If
    
    oForm.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Compute_TAX = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    oForm.Freeze False
    Sbo_Application.StatusBar.SetText "Compute_TAX 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    Compute_TAX = False
End Function

Private Function Compute_TAX_2(P_RETAMT As Double, P_GNMYER As Integer, P_DUPYER As Integer, P_SHWJIG As Double, P_SHWYAS As Double, P_SHWSUR As Double, Optional HS_TYPE As String, Optional SODAMT As Double) As Double
'***************************************************************************
' 퇴직소득세액  계산
'/ P_RETAMT : 대상퇴직금
'/ P_GNMYER : 세법상 근속년수
'/ P_DUPYER : 중복년수
'/ P_SHWJIG : 세액환산명세-퇴직연금지급예상액
'/ P_SHWYAS : 세액환산명세-환산연평균산출세액
'/ P_SHWSUR : 세액환산명세-수령가능퇴직급여액
'***************************************************************************
    Dim FD_i        As Integer
    Dim FD_RTXGBN_H As String
    Dim FD_RTXGBN_J As String
    Dim FD_YTXSTD As Double
    Dim FD_YSANTX As Double
    
 '/ Clear
    WK401.SPCGON = 0:    WK401.SODGON = 0:    WK401.RETGON = 0
    WK401.TAXSTD = 0:    WK401.YSANTX = 0:    WK401.YTXSTD = 0
    WK401.SANTAX = 0:    WK401.RTXGON = 0:    WK401.FRNGON = 0
    WK401.GULGAB = 0
    WK401.WSGGON = 0 '환산급여별 공제
    FD_YTXSTD = 0
    FD_YSANTX = 0
    If P_RETAMT <= 0 Then
        Compute_TAX_2 = 0
        Exit Function
    End If
    
 '/ 세액공제여부(Y/N)
    FD_RTXGBN_H = oDS_PH_PY115A.GetValue("U_SPCGBN", 0) '/ 주현근무지
    FD_RTXGBN_J = oDS_PH_PY115A.GetValue("U_JS1SPC", 0) '/ 종전근무지
    If Trim$(FD_RTXGBN_H) = "" Then FD_RTXGBN_H = "N"
    If Trim$(FD_RTXGBN_J) = "" Then FD_RTXGBN_J = "N"
  
    '/ 퇴직소득공제 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    '① 특별공제: 퇴직급여공제(일반퇴직금+명예퇴직수당(SUDAMT)+단체퇴직수당(BHMAMT))
        '(퇴직급여+단체퇴직보험료+2002년귀속분부터의 명예퇴직수당) * 50%
        '(2001년귀속분까지 명예퇴직수당) * 75%
        WK401.SPCGON = MDC_SetMod.IInt(P_RETAMT * (WG01.TB1RAT(91) / 100), 1)
        
        
    '② 퇴직소득공제: 근속연수에 따른 공제
        '근속연수에 따라 다음의 초과누진방법으로 공제
        '(근속연수: 5년이하          -           30만원*근속연수
        '           5년초과~10년이하 -  150만원+ 50만원*(근속연수- 5년)
        '          10년초과~20년이상 -  400만원+ 80만원*(근속연수-10년)
        '          20년초과          - 1200만원+120만원*(근속연수-20))
         If HS_TYPE = "BB" Then  '/ 법정퇴직금일때는 퇴직소득공제, 법정외퇴직금은 중복공제 안됨.
            For FD_i = 93 To 96
                If (P_GNMYER - P_DUPYER) <= WG01.TB1AMT(FD_i) Then
                   WK401.SODGON = WG01.TB1KUM(FD_i) + WG01.TB1GON(FD_i) * ((P_GNMYER - P_DUPYER) - WG01.TB1AMT(FD_i - 1))
                   Exit For
                End If
            Next FD_i
          End If
        WK401.SODGON = MDC_SetMod.IInt(WK401.SODGON, 1)

        
        If Left$(oDS_PH_PY115A.GetValue("U_ENDINT", 0), 4) >= "2016" Then
            '환산급여별 공제 2016년1월 1일 부터 시행
            If HS_TYPE = "BB" Then
                For FD_i = 81 To 85
                    If (oDS_PH_PY115A.GetValue("U_JS2AMT", 0)) <= WG01.TB1AMT(FD_i) Then
                       WK401.WSGGON = WG01.TB1KUM(FD_i) + ((oDS_PH_PY115A.GetValue("U_JS2AMT", 0) - WG01.TB1GON(FD_i)) * WG01.TB1RAT(FD_i) / 100)
                       Exit For
                    End If
                Next FD_i
              End If
        End If
        'WK401.WSGGON
        
        '//예외처리) 법정외 퇴직
        
        If HS_TYPE = "2" Then
           '/ 법정퇴직년도공제만큼 차감
            WK401.SODGON = WK401.SODGON - SODAMT
           
        ElseIf HS_TYPE = "3" Then
            Compute_TAX_2 = 0
            Exit Function
        End If
        
        
        '/ (차감)
        If WK401.SODGON > (P_RETAMT - WK401.SPCGON) Then
            WK401.SODGON = P_RETAMT - WK401.SPCGON
        End If
    '/ 퇴직소득공제(22 = ①특별공제 + ②소득공제) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WK401.RETGON = WK401.SPCGON + WK401.SODGON
    '/ 퇴직소득과세표준(23 = 21 - 22) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WK401.TAXSTD = P_RETAMT - WK401.RETGON
    '/ 연평균과세표준(24 = 23 / 근속년수) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WK401.YTXSTD = MDC_SetMod.IInt(WK401.TAXSTD / P_GNMYER, 1)
        
    '/ 연평균산출세액(25) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        If BubJung_YN = "Y" And WT_SHWTOT <> 0 Then
            '/ 전체 연평균산출세액(법정+법정외)
            If HS_TYPE = "2" Then
                FD_YTXSTD = Val(oDS_PH_PY115A.GetValue("U_SH1YAG", 0)) + WK401.YTXSTD
            Else
                FD_YTXSTD = Val(oDS_PH_PY115A.GetValue("U_SH1YAG", 0)) + Val(oDS_PH_PY115A.GetValue("U_SH2YAG", 0))
            End If
            '(소득세 기본세율 적용)
            For FD_i = 51 To 59
                If FD_YTXSTD <= WG01.TB1AMT(FD_i) Then
                    '/2013/12/25 노근용수정
                    FD_YSANTX = MDC_SetMod.IInt((FD_YTXSTD - WG01.TB1GON(FD_i)) * (WG01.TB1RAT(FD_i) / 100) + WG01.TB1KUM(FD_i), 1)
'                   FD_YSANTX = MDC_SetMod.IInt(FD_YTXSTD * (WG01.TB1RAT(FD_i) / 100) - WG01.TB1KUM(FD_i), 1)
                   Exit For
                End If
            Next FD_i
            If FD_YSANTX < 0 Then FD_YSANTX = 0
            oDS_PH_PY115A.setValue "U_YSANTX", 0, FD_YSANTX   '/ 전체산출세액먼저 셋팅
            '/ 연평균산출세액=현재연평균과세표준/법정+법정외연평균과세표준* 총산출세액
            If FD_YTXSTD <> 0 Then
                WK401.YSANTX = Round(WK401.YTXSTD / FD_YTXSTD * FD_YSANTX, 0)
'                If oDS_PH_PY115A.GetValue("U_OUTDAT", 0) >= "20120726" Then WK401.YSANTX = 0
            End If
        Else
            '(소득세 기본세율 적용)
            For FD_i = 51 To 59
                If WK401.YTXSTD <= WG01.TB1AMT(FD_i) Then
                    '/2013/12/25 노근용수정
                   WK401.YSANTX = Round((WK401.YTXSTD - WG01.TB1GON(FD_i)) * (WG01.TB1RAT(FD_i) / 100) + WG01.TB1KUM(FD_i), 0)
'                   WK401.YSANTX = Round(WK401.YTXSTD * (WG01.TB1RAT(FD_i) / 100) - WG01.TB1KUM(FD_i), 0)
'                   If oDS_PH_PY115A.GetValue("U_OUTDAT", 0) >= "20120726" Then WK401.YSANTX = 0
                   Exit For
                End If
            Next FD_i
        
        
        End If
    
        WK401.YSANTX = MDC_SetMod.IInt(WK401.YSANTX, 1)
        If WK401.YSANTX < 0 Then WK401.YSANTX = 0
    '/ 산출세액 (26 = 25 * 근속년수)  /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        If P_SHWJIG > 0 Then '/ 환산지급예상액이 있을경우
            If P_SHWSUR <> 0 Then
               WK401.SANTAX = MDC_SetMod.IInt(P_SHWYAS * P_GNMYER * P_RETAMT / P_SHWSUR, 1)
            End If
        Else
            WK401.SANTAX = MDC_SetMod.IInt(WK401.YSANTX * P_GNMYER, 1)
        End If
        
  '/ 퇴직세액공제(2009년 한시적:근무년수* 24만원한도) =  WK401.TAXGON
       If Trim$(FD_RTXGBN_H) = "Y" Then
            If P_SHWJIG > 0 Then '/ 환산지급예상액이 있을경우
                If WK401.SANTAX > 0 Then
                    WK401.RTXGON = MDC_SetMod.IInt(WK401.SANTAX * WK401.SANTAX / WK401.SANTAX * 0.3, 1)
                Else
                    WK401.RTXGON = 0
                End If
            Else
                WK401.RTXGON = MDC_SetMod.IInt(WK401.SANTAX * 0.3, 1)
            End If
                        
            If (P_GNMYER * 240000) < WK401.RTXGON Then
                WK401.RTXGON = (P_GNMYER * 240000)
            End If
       Else
            WK401.RTXGON = 0
       End If
        
  
    '/ 외국납부세액공제
       WK401.FRNGON = Val(oDS_PH_PY115A.GetValue("U_TAXGON", 0))   '/ 외국납부세액
    
    '/ 결정세액(26) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WK401.GULGAB = MDC_SetMod.IInt(WK401.SANTAX - WK401.FRNGON - WK401.RTXGON, 1)

        Compute_TAX_2 = WK401.GULGAB
        
End Function
'---------------------------------------------------------------------------------------
' Procedure : Compute_TAX
' 작성자    : 함미경
' Date      :
' Purpose   : 퇴직 세액 계산 루틴
' 수정자    : 최동권
' 수정일자  : 2009-02-05
' 수정사항  : 1) 프로시저 분리 및 이름변경(Execution -> Compute_TAX)
'             2) 퇴직금 + 세액 계산 -> 세액만 계산
' 중도정산년도와 정산년도종년년도가 다르며, 퇴직명예수당등 법정외금액이 있을경우
'---------------------------------------------------------------------------------------
'***************************************************************************
' 퇴직소득세액  계산
'***************************************************************************
Private Function Compute_TAX_3()
On Error GoTo Error_Message
    Dim WK_TOTGON As Double:   Dim WK_SILJIG As Double: Dim WK_CHAGAB As Double: Dim WK_CHAJUM As Double
    Dim WK_RETPAY As Double
    
    '/ 퇴직급여계
    WK_RETPAY = Val(oDS_PH_PY115A.GetValue("U_RETPAY", 0))
    '/ 차감세액 : 결정세액 - 종전납부세액
    WK_CHAGAB = Val(oDS_PH_PY115A.GetValue("U_GULGAB", 0)) - Val(oDS_PH_PY115A.GetValue("U_JONGAB", 0)) - Val(oDS_PH_PY115A.GetValue("U_SEYGAB", 0))
    WK_CHAJUM = Val(oDS_PH_PY115A.GetValue("U_GULJUM", 0)) - Val(oDS_PH_PY115A.GetValue("U_JONJUM", 0)) - Val(oDS_PH_PY115A.GetValue("U_SEYJUM", 0))
    WK_CHAGAB = MDC_SetMod.IInt(WK_CHAGAB, 10)
    WK_CHAJUM = MDC_SetMod.IInt(WK_CHAJUM, 10)
   
   '/ 공제총액(갑근세 + 특별세 + 주민세 + 퇴직전환 + 기타공제+ 건강보험정산 + 외국납부세액) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WK_TOTGON = WK_CHAGAB + WK_CHAJUM
    WK_TOTGON = WK_TOTGON + Val(oDS_PH_PY115A.GetValue("U_GITJA1", 0))    '/ 국민연금전환금
    WK_TOTGON = WK_TOTGON + Val(oDS_PH_PY115A.GetValue("U_MEDJSN", 0))    '/ 건강보험정산
    WK_TOTGON = WK_TOTGON + Val(oDS_PH_PY115A.GetValue("U_JSNGAB", 0))    '/ 중도정산소득세
    WK_TOTGON = WK_TOTGON + Val(oDS_PH_PY115A.GetValue("U_JSNJUM", 0))    '/ 중도정산주민세
    WK_TOTGON = WK_TOTGON + Val(oDS_PH_PY115A.GetValue("U_GITJA2", 0))    '/ 기타공제
   
    
    '/ 실지급액(퇴직급여액 - 공제총액) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WK_SILJIG = MDC_SetMod.IInt((WK_RETPAY + Val(oDS_PH_PY115A.GetValue("U_GITAA1", 0)) - WK_TOTGON), 1)
 
    '***************************************************************************
    ' 결과값 UPDATE
    '***************************************************************************
    oDS_PH_PY115A.setValue "U_CHAGAB", 0, WK_CHAGAB   '/ 차감소득세
    oDS_PH_PY115A.setValue "U_CHAJUM", 0, WK_CHAJUM   '/ 차감주민세
    oDS_PH_PY115A.setValue "U_TOTGON", 0, WK_TOTGON   '/ 공제총계
    oDS_PH_PY115A.setValue "U_SILJIG", 0, WK_SILJIG   '/ 실지급액
    
    Compute_TAX_3 = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    oForm.Freeze False
    Sbo_Application.StatusBar.SetText "Compute_TAX3 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    Compute_TAX_3 = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : Create_Data
' Author    : 함미경
' Date      :
' Purpose   : 퇴직금 생성
'---------------------------------------------------------------------------------------
Private Function Create_Data() As Boolean
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
   
    Dim ErrNum      As Integer
    
    Dim i               As Long
    Dim j               As Long
    Dim PAYSTD(1 To 4)  As String * 8
    Dim PAYEND(1 To 4)  As String * 8
    Dim PAYDAY(1 To 4)  As String * 2
    Dim PAYAMT(1 To 4)  As Double
    Dim BNSYMM(1 To 12) As String * 6
    Dim BNSAMT(1 To 12) As Double
    Dim YONCSU          As Double
    Dim CLTCOD          As String
    Dim MSTCOD          As String
    Dim RETDAT          As String
    Dim KUKJUN          As Double
    Dim HUGA            As Double
    Dim JSNGBN          As String
    Dim Div             As String   '퇴충계산시: 1, 퇴직금계산시 :2
    
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'/ Question
   If Sbo_Application.MessageBox("퇴직금 생성을 하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
      Exit Function
   End If
   
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    '/ Clear
    For i = 1 To 4
        PAYSTD(i) = "": PAYEND(i) = "": PAYDAY(i) = ""
        PAYAMT(i) = 0
    Next i
    For i = 1 To 12
        BNSYMM(i) = "": BNSAMT(i) = 0
    Next i
    YONCSU = 0
''/ 자료 조회 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    MSTCOD = oForm.Items("MSTCOD").Specific.String
    CLTCOD = MDC_SetMod.Get_ReData("U_CLTCOD", "Code", "[@PH_PY001A]", "'" & MSTCOD & "'")
    RETDAT = Format$(oDS_PH_PY115A.GetValue("U_ENDRET", 0), "0000-00-00")
    If Not oForm.Items("JSNGBN").Specific.Selected Is Nothing Then
        JSNGBN = oForm.Items("JSNGBN").Specific.Selected.VALUE
    End If
    
    Div = "2"  '퇴직금계산
    
    sQry = "Exec PH_PY115_01 '" & Trim$(MSTCOD) & "' , '" & Trim$(RETDAT) & "', '" & FIXTYP & "', '" & _
                               ROUNDT & "', '" & RODLEN & "', '" & RETCHK & "','','','" & RETYCH & "', '" & Div & "' "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount <= 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        If Val(oRecordSet.Fields("PAYT04").VALUE) > 0 Then
            PAYSTD(1) = Replace(oRecordSet.Fields("PAYSD1").VALUE, "-", "")
            PAYEND(1) = Replace(oRecordSet.Fields("PAYED1").VALUE, "-", "")
            PAYDAY(1) = oRecordSet.Fields("PAYDY1").VALUE
            PAYAMT(1) = oRecordSet.Fields("PAYT01").VALUE
            
            PAYSTD(2) = Replace(oRecordSet.Fields("PAYSD2").VALUE, "-", "")
            PAYEND(2) = Replace(oRecordSet.Fields("PAYED2").VALUE, "-", "")
            PAYDAY(2) = oRecordSet.Fields("PAYDY2").VALUE
            PAYAMT(2) = oRecordSet.Fields("PAYT02").VALUE
            
            PAYSTD(3) = Replace(oRecordSet.Fields("PAYSD3").VALUE, "-", "")
            PAYEND(3) = Replace(oRecordSet.Fields("PAYED3").VALUE, "-", "")
            PAYDAY(3) = oRecordSet.Fields("PAYDY3").VALUE
            PAYAMT(3) = oRecordSet.Fields("PAYT03").VALUE
            
            PAYSTD(4) = Replace(oRecordSet.Fields("PAYSD4").VALUE, "-", "")
            PAYEND(4) = Replace(oRecordSet.Fields("PAYED4").VALUE, "-", "")
            PAYDAY(4) = oRecordSet.Fields("PAYDY4").VALUE
            PAYAMT(4) = oRecordSet.Fields("PAYT04").VALUE
        Else
            PAYSTD(1) = Replace(oRecordSet.Fields("PAYSD1").VALUE, "-", "")
            PAYEND(1) = Replace(oRecordSet.Fields("PAYED1").VALUE, "-", "")
            PAYDAY(1) = oRecordSet.Fields("PAYDY1").VALUE
            PAYAMT(1) = oRecordSet.Fields("PAYT01").VALUE
            
            PAYSTD(2) = Replace(oRecordSet.Fields("PAYSD2").VALUE, "-", "")
            PAYEND(2) = Replace(oRecordSet.Fields("PAYED2").VALUE, "-", "")
            PAYDAY(2) = oRecordSet.Fields("PAYDY2").VALUE
            PAYAMT(2) = oRecordSet.Fields("PAYT02").VALUE
            
            PAYSTD(3) = Replace(oRecordSet.Fields("PAYSD3").VALUE, "-", "")
            PAYEND(3) = Replace(oRecordSet.Fields("PAYED3").VALUE, "-", "")
            PAYDAY(3) = oRecordSet.Fields("PAYDY3").VALUE
            PAYAMT(3) = oRecordSet.Fields("PAYT03").VALUE
        End If
        BNSYMM(1) = oRecordSet.Fields("BNSD01").VALUE
        BNSAMT(1) = oRecordSet.Fields("BNSA01").VALUE
        BNSYMM(2) = oRecordSet.Fields("BNSD02").VALUE
        BNSAMT(2) = oRecordSet.Fields("BNSA02").VALUE
        BNSYMM(3) = oRecordSet.Fields("BNSD03").VALUE
        BNSAMT(3) = oRecordSet.Fields("BNSA03").VALUE
        BNSYMM(4) = oRecordSet.Fields("BNSD04").VALUE
        BNSAMT(4) = oRecordSet.Fields("BNSA04").VALUE
        BNSYMM(5) = oRecordSet.Fields("BNSD05").VALUE
        BNSAMT(5) = oRecordSet.Fields("BNSA05").VALUE
        BNSYMM(6) = oRecordSet.Fields("BNSD06").VALUE
        BNSAMT(6) = oRecordSet.Fields("BNSA06").VALUE
        BNSYMM(7) = oRecordSet.Fields("BNSD07").VALUE
        BNSAMT(7) = oRecordSet.Fields("BNSA07").VALUE
        BNSYMM(8) = oRecordSet.Fields("BNSD08").VALUE
        BNSAMT(8) = oRecordSet.Fields("BNSA08").VALUE
        BNSYMM(9) = oRecordSet.Fields("BNSD09").VALUE
        BNSAMT(9) = oRecordSet.Fields("BNSA09").VALUE
        BNSYMM(10) = oRecordSet.Fields("BNSD10").VALUE
        BNSAMT(10) = oRecordSet.Fields("BNSA10").VALUE
        BNSYMM(11) = oRecordSet.Fields("BNSD11").VALUE
        BNSAMT(11) = oRecordSet.Fields("BNSA11").VALUE
        BNSYMM(12) = oRecordSet.Fields("BNSD12").VALUE
        BNSAMT(12) = oRecordSet.Fields("BNSA12").VALUE
    '/ 연차수당
        YONCSU = oRecordSet.Fields("YONCSU").VALUE
    '/ 퇴직전환금
        KUKJUN = oRecordSet.Fields("KUKJUN").VALUE
    '/ 휴가비
        HUGA = oRecordSet.Fields("HUGA").VALUE
        
    '/ 퇴직금세액공제(2009년 한시적)
        If Left$(RETDAT, 4) = "2009" And JSNGBN = "1" Then
            oDS_PH_PY115A.setValue "U_SPCGBN", 0, "Y"
        Else
            oDS_PH_PY115A.setValue "U_SPCGBN", 0, "N"
        End If
        
    End If
    '***************************************************************************
    ' 결과값 UPDATE
    '***************************************************************************
    '/ 급여자료
    oDS_PH_PY115A.setValue "U_PAYSD1", 0, PAYSTD(1)
    oDS_PH_PY115A.setValue "U_PAYED1", 0, PAYEND(1)
    oDS_PH_PY115A.setValue "U_PAYDY1", 0, PAYDAY(1)
    oDS_PH_PY115A.setValue "U_PAYT01", 0, PAYAMT(1)
    
    oDS_PH_PY115A.setValue "U_PAYSD2", 0, PAYSTD(2)
    oDS_PH_PY115A.setValue "U_PAYED2", 0, PAYEND(2)
    oDS_PH_PY115A.setValue "U_PAYDY2", 0, PAYDAY(2)
    oDS_PH_PY115A.setValue "U_PAYT02", 0, PAYAMT(2)

    oDS_PH_PY115A.setValue "U_PAYSD3", 0, PAYSTD(3)
    oDS_PH_PY115A.setValue "U_PAYED3", 0, PAYEND(3)
    oDS_PH_PY115A.setValue "U_PAYDY3", 0, PAYDAY(3)
    oDS_PH_PY115A.setValue "U_PAYT03", 0, PAYAMT(3)

    oDS_PH_PY115A.setValue "U_PAYSD4", 0, PAYSTD(4)
    oDS_PH_PY115A.setValue "U_PAYED4", 0, PAYEND(4)
    oDS_PH_PY115A.setValue "U_PAYDY4", 0, PAYDAY(4)
    oDS_PH_PY115A.setValue "U_PAYT04", 0, PAYAMT(4)
    
    '/ 상여자료
    oDS_PH_PY115A.setValue "U_BNSD01", 0, BNSYMM(1)
    oDS_PH_PY115A.setValue "U_BNST01", 0, BNSAMT(1)
    oDS_PH_PY115A.setValue "U_BNSD02", 0, BNSYMM(2)
    oDS_PH_PY115A.setValue "U_BNST02", 0, BNSAMT(2)
    oDS_PH_PY115A.setValue "U_BNSD03", 0, BNSYMM(3)
    oDS_PH_PY115A.setValue "U_BNST03", 0, BNSAMT(3)
    oDS_PH_PY115A.setValue "U_BNSD04", 0, BNSYMM(4)
    oDS_PH_PY115A.setValue "U_BNST04", 0, BNSAMT(4)
    oDS_PH_PY115A.setValue "U_BNSD05", 0, BNSYMM(5)
    oDS_PH_PY115A.setValue "U_BNST05", 0, BNSAMT(5)
    oDS_PH_PY115A.setValue "U_BNSD06", 0, BNSYMM(6)
    oDS_PH_PY115A.setValue "U_BNST06", 0, BNSAMT(6)
    oDS_PH_PY115A.setValue "U_BNSD07", 0, BNSYMM(7)
    oDS_PH_PY115A.setValue "U_BNST07", 0, BNSAMT(7)
    oDS_PH_PY115A.setValue "U_BNSD08", 0, BNSYMM(8)
    oDS_PH_PY115A.setValue "U_BNST08", 0, BNSAMT(8)
    oDS_PH_PY115A.setValue "U_BNSD09", 0, BNSYMM(9)
    oDS_PH_PY115A.setValue "U_BNST09", 0, BNSAMT(9)
    oDS_PH_PY115A.setValue "U_BNSD10", 0, BNSYMM(10)
    oDS_PH_PY115A.setValue "U_BNST10", 0, BNSAMT(10)
    oDS_PH_PY115A.setValue "U_BNSD11", 0, BNSYMM(11)
    oDS_PH_PY115A.setValue "U_BNST11", 0, BNSAMT(11)
    oDS_PH_PY115A.setValue "U_BNSD12", 0, BNSYMM(12)
    oDS_PH_PY115A.setValue "U_BNST12", 0, BNSAMT(12)
'/ 연차수당
    oDS_PH_PY115A.setValue "U_YONCSU", 0, YONCSU

'/ 퇴직금 1-퇴직정산일경우 중도정산소득세 가져오기를 선택한 경우
    If oDS_PH_PY115A.GetValue("U_JSNGBN", 0) = "1" Then
        Select Case Trim$(oCLTOPT(1))
        Case "1", "3"
            Call Create_JsnData
        End Select
    End If
'/ 퇴직전환금
   oDS_PH_PY115A.setValue "U_GITJA1", 0, KUKJUN
   
   oDS_PH_PY115A.setValue "U_ExPay2", 0, HUGA

'/ 급여내역 조회
    
    sQry = "Exec PH_PY115_02 '" & MSTCOD & "','" & RETDAT & "'"
    oRecordSet.DoQuery sQry

    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oDS_PH_PY115B.InsertRecord (i)
            oDS_PH_PY115B.Offset = j
            oDS_PH_PY115B.setValue "U_LineNum", i, i + 1
            oDS_PH_PY115B.setValue "U_CSUCOD", i, oRecordSet.Fields(0).VALUE
            oDS_PH_PY115B.setValue "U_CSUNAM", i, oRecordSet.Fields(1).VALUE
            oDS_PH_PY115B.setValue "U_CSUAMT1", i, oRecordSet.Fields(2).VALUE
            oDS_PH_PY115B.setValue "U_CSUAMT2", i, oRecordSet.Fields(3).VALUE
            oDS_PH_PY115B.setValue "U_CSUAMT3", i, oRecordSet.Fields(4).VALUE
            oDS_PH_PY115B.setValue "U_CSUAMT4", i, oRecordSet.Fields(5).VALUE
            oRecordSet.MoveNext
        Next
    End If
    
    oMat1.Columns("CSUAMT1").RightJustified = True
    oMat1.Columns("CSUAMT2").RightJustified = True
    oMat1.Columns("CSUAMT3").RightJustified = True
    oMat1.Columns("CSUAMT4").RightJustified = True
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
        
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
        
    oForm.Update
    Set oRecordSet = Nothing
    Create_Data = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "퇴직금 생성 자료가 없습니다!", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "이전3개월 급여자료가 없습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Creat_Data 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    Create_Data = False
End Function

'Private Function Execution1(Optional IMSI As String) As Boolean
'On Error GoTo error_Message
'    Dim LSTDAT      As String
'    Dim STRDAT      As String * 8
'    Dim ENDDAT      As String * 8
'
'    Dim ErrNum      As Integer
'    Dim Chk_Data    As String
'    Dim i           As Long
'    Dim k           As Long
'    Dim AVRPAY      As Double
'    Dim AVRBNS      As Double
'    Dim AVRYON      As Double
'    Dim AVRDAY      As Double
'    Dim MONPAY      As Double
'    Dim WK_GNSYER As Integer: Dim WK_GNSMON As Integer: Dim WK_GNSDAY As Integer
'    Dim WK_JONYER As Integer: Dim WK_JONMON As Integer: Dim WK_JMMMON As Integer
'    Dim WK_GNMYER As Integer: Dim WK_GNMMON As Integer: Dim WK_GNMDAY As Integer
'    Dim WK_YERTJK As Double: Dim WK_MONTJK As Double: Dim WK_DAYTJK As Double: Dim WK_TJKPAY As Double
'    Dim WK_RETPAY As Double:
'    Dim WK_JTOT01 As Double: Dim WK_JRET01 As Double: Dim WK_JSUD01 As Double: Dim WK_JBHM01 As Double
'
'    ErrNum = 0
'
'    '/ 귀속연도 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'    STRDAT = oDS_PH_PY115A.GetValue("U_STRRET", 0)
'    ENDDAT = oDS_PH_PY115A.GetValue("U_ENDRET", 0)
'
'    If Trim$(oDS_PH_PY115A.GetValue("U_STRINT", 0)) = "" Then
'        If Mid$(STRDAT, 1, 4) = Mid$(ENDDAT, 1, 4) Then
'            oDS_PH_PY115A.SetValue "U_STRINT", 0, STRDAT
'            If Mid$(oDS_PH_PY115A.GetValue("U_JOTDAT", 0), 1, 4) = Mid$(ENDDAT, 1, 4) Then
'                oDS_PH_PY115A.SetValue "U_STRINT", 0, Mid$(ENDDAT, 1, 4) & "0101"
'            End If
'        Else
'            oDS_PH_PY115A.SetValue "U_STRINT", 0, Mid$(ENDDAT, 1, 4) & "0101"
'        End If
'    End If
'    If Trim$(oDS_PH_PY115A.GetValue("U_ENDINT", 0)) = "" Then
'        oDS_PH_PY115A.SetValue "U_ENDINT", 0, ENDDAT
'    End If
'    '/ 주현 근속년월일 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'    Call MDC_SetMod.Term2(STRDAT, ENDDAT)
'    WK_GNSYER = Val(ZPAY_GBL_GNSYER)
'    WK_GNSMON = Val(ZPAY_GBL_GNSMON)
'    WK_GNSDAY = Val(ZPAY_GBL_GNSDAY)
'    STRDAT = "20021007"
'    ENDDAT = "20060920"
'    ZPAY_GBL_GNMDAY = TermDay(STRDAT, ENDDAT) + 1
'    '/ 종전 근속연월일 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'    WK_JONYER = Val(oDS_PH_PY115A.GetValue("U_JONYER", 0))
'    WK_JONMON = Val(oDS_PH_PY115A.GetValue("U_JONMON", 0))
'    WK_JMMMON = Val(oDS_PH_PY115A.GetValue("U_JMMMON", 0))
'    '/ (WK_GNMDAY: 종전근무월수로 사용)
'    WK_GNMDAY = WK_JONYER * 12 + WK_JONMON
'    '***************************************************************************
'    ' 근무년수가 1년 이상이 되어야 하므로 1미만은 퇴직금계산에서 제외된다
'    '***************************************************************************
'    '대경전장은 1년미만자도 퇴직금나간다네..
'    If WK_GNSYER = 0 And WK_JONYER Then
''        ErrNum = 1
''        GoTo error_Message
'    End If
'
'    '/ 근무년월일 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'    If WK_GNSDAY > 0 Or (WK_GNSMON + WK_JONMON - WK_JMMMON) > 0 Then
'       WK_GNMYER = WK_GNSYER + WK_JONYER + 1
'    Else
'       WK_GNMYER = WK_GNSYER + WK_JONYER
'    End If
'    If WK_GNSDAY > 0 Then
'       WK_GNMMON = WK_GNSYER * 12 + WK_GNSMON + 1
'    Else
'       WK_GNMMON = WK_GNSYER * 12 + WK_GNSMON
'    End If
'
'    '/ 해당월의 마지막 일자
'    LSTDAT = Format$(DateAdd("m", 1, Format$(Mid$(ENDDAT, 1, 6) & "01", "0000-00-00")) - 1, "yyyymmdd")
'    '/ 평균임금
'    AVRPAY = 0
'    AVRBNS = 0
'    AVRYON = 0
'    AVRDAY = 0
'    MONPAY = 0
'
'    AVRDAY = AVRDAY + Val(oDS_PH_PY115A.GetValue("U_PAYDY1", 0))
'    AVRDAY = AVRDAY + Val(oDS_PH_PY115A.GetValue("U_PAYDY2", 0))
'    AVRDAY = AVRDAY + Val(oDS_PH_PY115A.GetValue("U_PAYDY3", 0))
'    AVRDAY = AVRDAY + Val(oDS_PH_PY115A.GetValue("U_PAYDY4", 0))
'    '/ 최근3개월간 급여
'    AVRPAY = AVRPAY + Val(oDS_PH_PY115A.GetValue("U_PAYT01", 0))
'    AVRPAY = AVRPAY + Val(oDS_PH_PY115A.GetValue("U_PAYT02", 0))
'    AVRPAY = AVRPAY + Val(oDS_PH_PY115A.GetValue("U_PAYT03", 0))
'    AVRPAY = AVRPAY + Val(oDS_PH_PY115A.GetValue("U_PAYT04", 0))
'    '/ 최근1년간 상여
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST01", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST02", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST03", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST04", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST05", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST06", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST07", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST08", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST09", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST10", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST11", 0))
'    AVRBNS = AVRBNS + Val(oDS_PH_PY115A.GetValue("U_BNST12", 0))
'
'    AVRYON = Val(oDS_PH_PY115A.GetValue("U_YONCSU", 0))
'    '/ 평균
'    MONPAY = Format$((AVRBNS + AVRYON) / 12, "#########0")
''/////////////////////////////////////////////////////////////////
''    AVRPAY = Format$((AVRPAY / 3), "#########0")
''    AVRBNS = Format$((AVRBNS / 12), "#########0")
''    AVRYON = Format$((AVRYON / 12), "#########0")
''/ 케어라인만 적용(070418)
'    AVRPAY = MDC_SetMod.IInt((AVRPAY / 90) + 5, 10)
'    MONPAY = Format$(MONPAY + AVRPAY, "########0")
''////////////////////////////////////////////////////////////////
''/ 퇴직급여(퇴직금) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'
''    WK_YERTJK = Format$((MONPAY * WK_GNSYER), "########0")
''    WK_MONTJK = Format$((MONPAY * WK_GNSMON / 12), "########0")
''    WK_DAYTJK = Format$((MONPAY * WK_GNSDAY / 365), "########0")
''    WK_TJKPAY = Format$((WK_YERTJK + WK_MONTJK + WK_DAYTJK), "########0")
''/ 케어라인만 적용(070418)
'    WK_YERTJK = 0
'    WK_MONTJK = MDC_SetMod.IInt(MONPAY * ZPAY_GBL_GNMDAY / 365 * 30 + 9, 10)
''    WK_DAYTJK = 0
'    WK_TJKPAY = Format$((WK_YERTJK + WK_MONTJK + WK_DAYTJK), "########0")
'
'
''/ 퇴직급여액[퇴직급여+명예퇴직수당+단체퇴직수당+종전퇴직금](21) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'    WK_RETPAY = WK_TJKPAY + Val(oDS_PH_PY115A.GetValue("U_SUDAMT", 0)) + Val(oDS_PH_PY115A.GetValue("U_BHMAMT", 0))
''/ 종전근무지
'    WK_JRET01 = Val(oDS_PH_PY115A.GetValue("U_JRET01", 0))
'    WK_JSUD01 = Val(oDS_PH_PY115A.GetValue("U_JSUD01", 0))
'    WK_JBHM01 = Val(oDS_PH_PY115A.GetValue("U_JBHM01", 0))
'    WK_JTOT01 = WK_JRET01 + WK_JSUD01 + WK_JBHM01
''/
'    If WK_RETPAY = 0 Then
'        ErrNum = 2
'        GoTo error_Message
'    End If
''***************************************************************************
'' 퇴직소득세액  계산
''***************************************************************************
'    Dim WK_SPCGON As Double:    Dim WK_SODGON As Double:    Dim WK_RETGON As Double
'    Dim WK_TAXSTD As Double:    Dim WK_YTXSTD As Double:    Dim WK_YSANTX As Double
'    Dim WK_SANTAX As Double:    Dim WK_TAXGON As Double:    Dim WK_GULGAB As Double
'    Dim WK_GULJUM As Double:    Dim WK_CHAGAB As Double:    Dim WK_CHAJUM As Double
'    Dim WK_TOTGON As Double:    Dim WK_SILJIG As Double:    Dim WK_KUKJUN As Double
'    '/ 퇴직소득공제 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'    '① 특별공제: 퇴직급여공제(일반퇴직금+명예퇴직수당(SUDAMT)+단체퇴직수당(BHMAMT))
'        '(퇴직급여+단체퇴직보험료+2002년귀속분부터의 명예퇴직수당) * 50%
'        '(2001년귀속분까지 명예퇴직수당) * 75%
'        WK_SPCGON = ((WK_TJKPAY + WK_JTOT01 + Val(oDS_PH_PY115A.GetValue("U_BHMAMT", 0)) + Val(oDS_PH_PY115A.GetValue("U_SUDAMT", 0))) * (WG01.TB1RAT(91) / 100)) + _
'                    (Val(oDS_PH_PY115A.GetValue("U_SUD2AMT", 0)) * 0.75)
'
'
'        WK_SPCGON = MDC_SetMod.IInt(WK_SPCGON, 1)
'    '② 퇴직소득공제: 근속연수에 따른 공제
'        '근속연수에 따라 다음의 초과누진방법으로 공제
'        '(근속연수: 5년이하          -           30만원*근속연수
'        '           5년초과~10년이하 -  150만원+ 50만원*(근속연수- 5년)
'        '          10년초과~20년이상 -  400만원+ 80만원*(근속연수-10년)
'        '          20년초과          - 1200만원+120만원*(근속연수-20))
'          For i = 93 To 96
'              If WK_GNMYER <= WG01.TB1AMT(i) Then
'                 WK_SODGON = WG01.TB1KUM(i) + WG01.TB1GON(i) * (WK_GNMYER - WG01.TB1AMT(i - 1))
'                 Exit For
'              End If
'          Next i
'
''        If WK_GNMYER <= 5 Then
''            WK_SODGON = 300000 * WK_GNMYER
''        ElseIf WK_GNMYER <= 10 Then
''            WK_SODGON = 1500000 + 500000 * (WK_GNMYER - 5)
''        ElseIf WK_GNMYER <= 20 Then
''            WK_SODGON = 4000000 + 800000 * (WK_GNMYER - 10)
''        ElseIf WK_GNMYER <= 99 Then
''            WK_SODGON = 12000000 + 1200000 * (WK_GNMYER - 20)
''        End If
'        WK_SODGON = MDC_SetMod.IInt(WK_SODGON, 1)
'        '/ (차감)
'        If WK_SODGON > (WK_RETPAY + WK_JTOT01 - WK_SPCGON) Then
'            WK_SODGON = WK_RETPAY + WK_JTOT01 - WK_SPCGON
'        End If
'    '/ 퇴직소득공제(22 = ①특별공제 + ②소득공제) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        WK_RETGON = WK_SPCGON + WK_SODGON
'    '/ 퇴직소득과세표준(23 = 21 - 22) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        WK_TAXSTD = WK_RETPAY + WK_JTOT01 - WK_RETGON
'    '/ 연평균과세표준(24 = 23 / 근속년수) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        WK_YTXSTD = MDC_SetMod.IInt(WK_TAXSTD / WK_GNMYER, 1)
'    '/ 연평균산출세액(25) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        '(소득세 기본세율 적용)
'        For i = 51 To 59
'            If WK_YTXSTD <= WG01.TB1AMT(i) Then
'               WK_YSANTX = WK_YTXSTD * (WG01.TB1RAT(i) / 100) - WG01.TB1KUM(i)
'               Exit For
'            End If
'        Next i
'
''        If WK_YTXSTD <= 10000000 Then
''            WK_YSANTX = WK_YTXSTD * 0.08
''        ElseIf WK_YTXSTD <= 40000000 Then
''            WK_YSANTX = WK_YTXSTD * 0.17 - 900000
''        ElseIf WK_YTXSTD <= 80000000 Then
''            WK_YSANTX = WK_YTXSTD * 0.26 - 4500000
''        Else
''            WK_YSANTX = WK_YTXSTD * 0.35 - 11700000
''        End If
'        WK_YSANTX = MDC_SetMod.IInt(WK_YSANTX, 1)
'        If WK_YSANTX < 0 Then WK_YSANTX = 0
'    '/ 산출세액(26 = 25 * 근속년수) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        WK_SANTAX = MDC_SetMod.IInt(WK_YSANTX * WK_GNMYER, 1)
'    '/ 퇴직소득세액공제(27) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        '2005년부터 퇴직소득세액공제 없어짐.
'        WK_TAXGON = 0
'    '/ 결정세액(26) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        WK_GULGAB = MDC_SetMod.IInt(WK_SANTAX - WK_TAXGON, 10)
'        '/ 대경전장만 기존에 일자계산이 틀린게 있다네 입력가능하게 변경..쯧쯧
'         If IMSI = "GULGAB" Then
'             WK_GULGAB = Val(oDS_PH_PY115A.GetValue("U_GULGAB", 0))
'         End If
'        WK_GULJUM = MDC_SetMod.IInt((WK_GULGAB * 0.1), 10)
'    '/ 차감세액(27) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        WK_CHAGAB = MDC_SetMod.IInt(WK_GULGAB - Val(oDS_PH_PY115A.GetValue("U_JONGAB", 0)), 10)
'        WK_CHAJUM = MDC_SetMod.IInt(WK_GULJUM - Val(oDS_PH_PY115A.GetValue("U_JONJUM", 0)), 10)
'    '/ 퇴직전환 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        WK_KUKJUN = 0
'    '/ 공제총액(갑근세 + 특별세 + 주민세 + 퇴직전환 + 기타공제+ 건강보험정산) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        WK_TOTGON = MDC_SetMod.IInt(WK_CHAGAB + WK_CHAJUM + WK_KUKJUN + _
'                    oDS_PH_PY115A.GetValue("U_GITJA1", 0) + Val(oDS_PH_PY115A.GetValue("U_MEDJSN", 0)), 10)
'    '/ 실지급액(퇴직급여액 - 공제총액) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'        WK_SILJIG = MDC_SetMod.IInt((WK_RETPAY + Val(oDS_PH_PY115A.GetValue("U_GITAA1", 0)) - WK_TOTGON), 1)
'
'    '***************************************************************************
'    ' 결과값 UPDATE
'    '***************************************************************************
'    oForm.Freeze True
'    oDS_PH_PY115A.SetValue "U_JTOT01", 0, WK_JTOT01
'    oDS_PH_PY115A.SetValue "U_GNSYER", 0, WK_GNSYER
'    oDS_PH_PY115A.SetValue "U_GNSMON", 0, WK_GNSMON
'    oDS_PH_PY115A.SetValue "U_GNSDAY", 0, WK_GNSDAY
'    oDS_PH_PY115A.SetValue "U_GNMYER", 0, WK_GNMYER
'    oDS_PH_PY115A.SetValue "U_GNMMON", 0, WK_GNMMON
'    oDS_PH_PY115A.SetValue "U_GNMDAY", 0, WK_GNMDAY
'
'    oDS_PH_PY115A.SetValue "U_AVRPAY", 0, AVRPAY
'    oDS_PH_PY115A.SetValue "U_AVRBNS", 0, AVRBNS
'    oDS_PH_PY115A.SetValue "U_AVRYON", 0, AVRYON
'    oDS_PH_PY115A.SetValue "U_MONPAY", 0, MONPAY
'
'    oDS_PH_PY115A.SetValue "U_YERTJK", 0, WK_YERTJK
'    oDS_PH_PY115A.SetValue "U_MONTJK", 0, WK_MONTJK
'    oDS_PH_PY115A.SetValue "U_DAYTJK", 0, WK_DAYTJK
'    oDS_PH_PY115A.SetValue "U_TJKPAY", 0, WK_TJKPAY
'    oDS_PH_PY115A.SetValue "U_RETPAY", 0, WK_RETPAY
'    oDS_PH_PY115A.SetValue "U_RETGON", 0, WK_RETGON
'    oDS_PH_PY115A.SetValue "U_TAXSTD", 0, WK_TAXSTD
'    oDS_PH_PY115A.SetValue "U_YTXSTD", 0, WK_YTXSTD
'    oDS_PH_PY115A.SetValue "U_YSANTX", 0, WK_YSANTX
'    oDS_PH_PY115A.SetValue "U_SANTAX", 0, WK_SANTAX
'    oDS_PH_PY115A.SetValue "U_TAXGON", 0, WK_TAXGON
'    oDS_PH_PY115A.SetValue "U_GULGAB", 0, WK_GULGAB
'    oDS_PH_PY115A.SetValue "U_GULJUM", 0, WK_GULJUM
'    oDS_PH_PY115A.SetValue "U_CHAGAB", 0, WK_CHAGAB
'    oDS_PH_PY115A.SetValue "U_CHAJUM", 0, WK_CHAJUM
'    oDS_PH_PY115A.SetValue "U_TOTGON", 0, WK_TOTGON
'    oDS_PH_PY115A.SetValue "U_SILJIG", 0, WK_SILJIG
'    oForm.Freeze False
'    oForm.Update
'
'    Execution1 = True
'    Exit Function
''/////////////////////////////////////////////////////////////////////////////////////////////////
'error_Message:
'    oForm.Freeze False
'    If ErrNum = 1 Then
'        Sbo_Application.StatusBar.SetText "근속년수가 1년미만이므로, 퇴직금계산이 불가합니다!", bmt_Short, smt_Error
'    ElseIf ErrNum = 2 Then
'        Sbo_Application.StatusBar.SetText "퇴직급여총액이 0입니다. 확인하여 주십시오.", bmt_Short, smt_Error
'    Else
'        Sbo_Application.StatusBar.SetText "Execution1 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
'    End If
'    Execution1 = False
'End Function



Private Sub MasterUpdate()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    
    Sbo_Application.MessageBox ("계산확정 후 중간정산은 인사마스터의 기산일을, 퇴직은 퇴직일자 와 재직구분을 꼭 수정바랍니다.")
'/ Question
    '//계산후 인사마스터 기산일에 Update 로직
    '// 김영호 과장요청으로 막음, 계산후 직접 인사마스터에 기산일을 수정하기로함
    '// 기산일 체크는 사원입력시 체크함. '2015/05/07 N.G.Y
                                
'   If Sbo_Application.MessageBox("퇴직금 중도정산일자를 사원마스터에 반영하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
'      Exit Sub
'   End If
'
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'    sQry = "UPDATE [@PH_PY001A] SET U_RETDAT = '" & Val(oDS_PH_PY115A.GetValue("U_ENDRET", 0)) & "'"
'    sQry = sQry & "   WHERE Code = N'" & oForm.Items("MSTCOD").Specific.VALUE & "'"
'    oRecordSet.DoQuery sQry
    '/ 메세지
    
    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "MasterUpdate 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    
End Sub

Private Sub Compute_APPRAT()
On Error GoTo Error_Message
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String

     Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'/ 임원 누진율
    sQry = " SELECT  ISNULL(T0.U_ADDRAT,1) FROM [@PH_PY114C] T0 INNER JOIN [OHPS] T1 ON T0.U_MSTSTP = T1.posID"
    sQry = sQry & " INNER JOIN [@PH_PY001A] T2 ON T2.U_Position = T1.POSID"
    sQry = sQry & " WHERE T0.Code = '" & oCode & "' AND T2.Code = '" & Trim$(oDS_PH_PY115A.GetValue("U_MSTCOD", 0)) & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        oDS_PH_PY115A.setValue "U_ADDRAT", 0, Val(oRecordSet.Fields(0).VALUE)
    Else
        oDS_PH_PY115A.setValue "U_ADDRAT", 0, 1
    End If
    oForm.Items("ADDRAT").Update

    Set oRecordSet = Nothing
Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    
    Sbo_Application.StatusBar.SetText "Compute_APPRAT 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub Create_JsnData()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
    Dim WK_JSNYER As String
    Dim WK_TOTGON As Double:   Dim WK_SILJIG As Double
    Dim WK_JSNGAB As Double:   Dim WK_JSNJUM As Double
    ErrNum = 0
    
    WK_JSNYER = Left$(oDS_PH_PY115A.GetValue("U_ENDRET", 0), 4)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
''/ 자료 조회 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    sQry = " SELECT ISNULL(U_CHAGAB,0) AS U_CHAGAB, ISNULL(U_CHAJUM,0) AS U_CHAJUM "
    sQry = sQry & " FROM [@ZPY504H] WHERE U_JSNYER = '" & Trim$(WK_JSNYER) & "' AND   U_JSNGBN = '2'"
    sQry = sQry & " AND U_MSTCOD = '" & Trim$(oDS_PH_PY115A.GetValue("U_MSTCOD", 0)) & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        WK_JSNGAB = oRecordSet.Fields(0).VALUE
        WK_JSNJUM = oRecordSet.Fields(1).VALUE
    End If
    '***************************************************************************
    ' 결과값 UPDATE
    '***************************************************************************
    '/ 공제총액(갑근세 + 주민세 + 퇴직전환 + 건강보험정산+정산소득세+정산주민세) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WK_TOTGON = Val(oDS_PH_PY115A.GetValue("U_CHAGAB", 0)) + Val(oDS_PH_PY115A.GetValue("U_CHAJUM", 0)) + _
                    Val(oDS_PH_PY115A.GetValue("U_GITJA1", 0)) + Val(oDS_PH_PY115A.GetValue("U_MEDJSN", 0)) + _
                    WK_JSNGAB + WK_JSNJUM
    '/ 실지급액(퇴직급여액 - 공제총액) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WK_SILJIG = Val(oDS_PH_PY115A.GetValue("U_RETPAY", 0)) + Val(oDS_PH_PY115A.GetValue("U_GITAA1", 0)) - WK_TOTGON
    
    oDS_PH_PY115A.setValue "U_JSNGAB", 0, WK_JSNGAB
    oDS_PH_PY115A.setValue "U_JSNJUM", 0, WK_JSNJUM
    oDS_PH_PY115A.setValue "U_TOTGON", 0, WK_TOTGON
    oDS_PH_PY115A.setValue "U_SILJIG", 0, WK_SILJIG
    
    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "Create_JsnData 실행중 오류가 발생하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub Display_ilSiGum()
'/********************************************************************************************************/
'/ 퇴직연금 명세 계산
'/********************************************************************************************************/
    Dim WK_MYNTOT As Double, WK_MYNBUL As Double, WK_MYNGON As Double, WK_MYNWON As Double, WK_YILPAY As Double
    Dim WK_JYNTOT As Double, WK_JYNBUL As Double, WK_JYNGON As Double, WK_JYNWON As Double, WK_JYIL01 As Double
    
    '/ 주현 퇴직연금 명세
    WK_MYNTOT = Val(oDS_PH_PY115A.GetValue("U_MYNTOT", 0))    '/ 총수령액
    WK_MYNWON = Val(oDS_PH_PY115A.GetValue("U_MYNWON", 0))    '/ 원리금합계액
    
    WK_MYNBUL = Val(oDS_PH_PY115A.GetValue("U_MYNBUL", 0))    '/ 소득자불입액
    WK_MYNGON = Val(oDS_PH_PY115A.GetValue("U_MYNGON", 0))    '/ 퇴직연금소득공제액
    
    '/ 종전 퇴직연금 명세
    WK_JYNTOT = Val(oDS_PH_PY115A.GetValue("U_JYNTOT", 0))   '/ 총수령액
    WK_JYNWON = Val(oDS_PH_PY115A.GetValue("U_JYNWON", 0))   '/ 원리금합계액
    
    WK_JYNBUL = Val(oDS_PH_PY115A.GetValue("U_JYNBUL", 0))   '/ 소득자불입액
    WK_JYNGON = Val(oDS_PH_PY115A.GetValue("U_JYNGON", 0))   '/ 퇴직연금소득공제액
    
    '/ 퇴직연금일시금 계산-주현
    If WK_MYNWON = 0 Then WK_MYNWON = 1
    If WK_JYNWON = 0 Then WK_JYNWON = 1
    WK_YILPAY = MDC_SetMod.IInt(WK_MYNTOT * (1 - (WK_MYNBUL - WK_MYNGON) / WK_MYNWON), 1)
    WK_JYIL01 = MDC_SetMod.IInt(WK_JYNTOT * (1 - (WK_JYNBUL - WK_JYNGON) / WK_JYNWON), 1)
    
 '/ 총수령액이 없으면 일시금도 0
    If WK_JYNTOT = 0 Then WK_JYIL01 = 0
    If WK_MYNTOT = 0 Then WK_YILPAY = 0
    
 '/ Display
    oDS_PH_PY115A.setValue "U_JYIL01", 0, WK_JYIL01
    oForm.Items("JYIL01").Update
    
    oDS_PH_PY115A.setValue "U_YILPAY", 0, WK_YILPAY
    oForm.Items("YILPAY").Update
 
End Sub

Private Sub Display_TaxWhanSan(HS_TYPE As Integer, RPTGBN As String)
'/********************************************************************************************************/
'/ 세액환산 명세 계산
'/ WS_TYPE : 1-법정퇴직금세액환산명세, 2-법정외퇴직금세액환산명세,
'/ WS_TYPE : 3-주현퇴직금세액환산명세, 4-종전퇴직세액환산명세
'/********************************************************************************************************/
    Dim i As Integer
    
    Dim FD_YUNTOT As Double:    Dim FD_YUNBUL As Double:    Dim FD_YUNGON As Double
    Dim FD_YUNWON As Double:    Dim FD_RETAMT As Double
    Dim FD_GNMYER As Integer
    Dim FD_GNMMON As Integer
    
    
    
    '/ Clear
    HS.SHWJIG = 0:    HS.SHWTIL = 0:    HS.SHWSUR = 0:    HS.SHWGON = 0:
    HS.SHWGWA = 0:    HS.SHWYAG = 0:    HS.SHWYAS = 0
    
    '/ 퇴직연금자료
    Select Case HS_TYPE
    Case 1
        '/ 지급예상액
        SHWJIG(1) = Val(oDS_PH_PY115A.GetValue("U_SH1JIG", 0))
        SHWJIG(2) = Val(oDS_PH_PY115A.GetValue("U_SH3JIG", 0))
        HS.SHWJIG = Val(oDS_PH_PY115A.GetValue("U_SH1JIG", 0)) + Val(oDS_PH_PY115A.GetValue("U_SH3JIG", 0))  '/ 법정지급예상액
        
        FD_GNMMON = Val(oDS_PH_PY115A.GetValue("U_GNMMON", 0)) - Val(oDS_PH_PY115A.GetValue("U_EXPMON", 0))
        FD_YUNTOT = Val(oDS_PH_PY115A.GetValue("U_MYNTOT", 0)) + Val(oDS_PH_PY115A.GetValue("U_JYNTOT", 0)) '/ 퇴직연금=총수령액
        FD_YUNBUL = Val(oDS_PH_PY115A.GetValue("U_MYNBUL", 0)) + Val(oDS_PH_PY115A.GetValue("U_JYNBUL", 0)) '/ 퇴직연금=불입액
        FD_YUNGON = Val(oDS_PH_PY115A.GetValue("U_MYNGON", 0)) + Val(oDS_PH_PY115A.GetValue("U_JYNGON", 0)) '/ 퇴직연금=공제액
        FD_YUNWON = Val(oDS_PH_PY115A.GetValue("U_MYNWON", 0)) + Val(oDS_PH_PY115A.GetValue("U_JYNWON", 0)) '/ 퇴직연금=원리금합계액
        '/ 법정퇴직급여=주현+종전
        If Trim$(RPTGBN) = "Y" Then
            FD_RETAMT = Val(oDS_PH_PY115A.GetValue("U_TJKPAY", 0)) + Val(oDS_PH_PY115A.GetValue("U_JRET01", 0))
            If Trim$(oDS_PH_PY115A.GetValue("U_ST2RET", 0)) <> "" Then FD_RETAMT = FD_RETAMT + Val(oDS_PH_PY115A.GetValue("U_JSUD01", 0))
        Else
            'FD_RETAMT = Val(oDS_PH_PY115A.GetValue("U_TJKPAY", 0)) + Val(oDS_PH_PY115A.GetValue("U_SUDAMT", 0)) + _
                        Val(oDS_PH_PY115A.GetValue("U_JRET01", 0)) + Val(oDS_PH_PY115A.GetValue("U_JSUD01", 0))
            FD_RETAMT = Val(oDS_PH_PY115A.GetValue("U_TJKPAY", 0)) + Val(oDS_PH_PY115A.GetValue("U_SUDAMT", 0)) + _
                        Val(oDS_PH_PY115A.GetValue("U_JRET01", 0)) + Val(oDS_PH_PY115A.GetValue("U_JADD01", 0)) + _
                        Val(oDS_PH_PY115A.GetValue("U_SH1IYO", 0)) + Val(oDS_PH_PY115A.GetValue("U_SH2IYO", 0)) + _
                        Val(oDS_PH_PY115A.GetValue("U_SH3IYO", 0)) + Val(oDS_PH_PY115A.GetValue("U_SH4IYO", 0)) + _
                        Val(oDS_PH_PY115A.GetValue("U_SH1GSU", 0)) + Val(oDS_PH_PY115A.GetValue("U_SH2GSU", 0)) + _
                        Val(oDS_PH_PY115A.GetValue("U_SH3GSU", 0)) + Val(oDS_PH_PY115A.GetValue("U_SH4GSU", 0)) + _
                        Val(oDS_PH_PY115A.GetValue("U_SH2TIL", 0))
        End If
    Case 2
        '/ 지급예상액
        SHWJIG(1) = Val(oDS_PH_PY115A.GetValue("U_SH2JIG", 0))
        SHWJIG(2) = Val(oDS_PH_PY115A.GetValue("U_SH4JIG", 0))
        HS.SHWJIG = SHWJIG(1) + SHWJIG(2)  '/ 지급예상액
        
        FD_YUNBUL = Val(oDS_PH_PY115A.GetValue("U_MYNBUL", 0)) + Val(oDS_PH_PY115A.GetValue("U_JYNBUL", 0)) '/ 퇴직연금=불입액
        FD_YUNGON = Val(oDS_PH_PY115A.GetValue("U_MYNGON", 0)) + Val(oDS_PH_PY115A.GetValue("U_JYNGON", 0)) '/ 퇴직연금=공제액
        FD_YUNWON = Val(oDS_PH_PY115A.GetValue("U_MYNWON", 0)) + Val(oDS_PH_PY115A.GetValue("U_JYNWON", 0)) '/ 퇴직연금=원리금합계액
        
        FD_GNMMON = Val(oDS_PH_PY115A.GetValue("U_GNMMON", 0)) - Val(oDS_PH_PY115A.GetValue("U_EXPMON", 0))
        FD_YUNTOT = Val(oDS_PH_PY115A.GetValue("U_MYNTOT", 0)) + Val(oDS_PH_PY115A.GetValue("U_JYNTOT", 0)) '/ 퇴직연금=총수령액
        
       '/ 법정외퇴직급여=주현+종전
        If HS.SHWJIG > 0 Then
            FD_RETAMT = Val(oDS_PH_PY115A.GetValue("U_SUDAMT", 0))
            If Trim$(oDS_PH_PY115A.GetValue("U_ST2RET", 0)) = "" Then FD_RETAMT = FD_RETAMT + Val(oDS_PH_PY115A.GetValue("U_JSUD01", 0))
        Else
            FD_RETAMT = 0
        End If

    Case 3  '/ 주현
        '/ 지급예상액
        HS.SHWJIG = Val(oDS_PH_PY115A.GetValue("U_SH1JIG", 0))
        FD_GNMMON = Val(oDS_PH_PY115A.GetValue("U_GNMMON", 0)) - Val(oDS_PH_PY115A.GetValue("U_EXPMON", 0))
        FD_YUNTOT = Val(oDS_PH_PY115A.GetValue("U_MYNTOT", 0)) '/ 퇴직연금=총수령액
        FD_YUNBUL = Val(oDS_PH_PY115A.GetValue("U_MYNBUL", 0)) '/ 퇴직연금=불입액
        FD_YUNGON = Val(oDS_PH_PY115A.GetValue("U_MYNGON", 0)) '/ 퇴직연금=공제액
        FD_YUNWON = Val(oDS_PH_PY115A.GetValue("U_MYNWON", 0)) '/ 퇴직연금=원리금합계액
        '/
        FD_RETAMT = Val(oDS_PH_PY115A.GetValue("U_TJKPAY", 0)) + Val(oDS_PH_PY115A.GetValue("U_SUDAMT", 0))
    Case 4
        '/ 지급예상액
        HS.SHWJIG = Val(oDS_PH_PY115A.GetValue("U_SH2JIG", 0))
        
        FD_GNMMON = Val(oDS_PH_PY115A.GetValue("U_GNMDAY", 0)) - Val(oDS_PH_PY115A.GetValue("U_JEXMON", 0))
        FD_YUNTOT = Val(oDS_PH_PY115A.GetValue("U_JYNTOT", 0))  '/ 퇴직연금=총수령액
        FD_YUNBUL = Val(oDS_PH_PY115A.GetValue("U_JYNBUL", 0)) '/ 퇴직연금=불입액
        FD_YUNGON = Val(oDS_PH_PY115A.GetValue("U_JYNGON", 0))  '/ 퇴직연금=공제액
        FD_YUNWON = Val(oDS_PH_PY115A.GetValue("U_JYNWON", 0))  '/ 퇴직연금=원리금합계액
        '/
        FD_RETAMT = Val(oDS_PH_PY115A.GetValue("U_JRET01", 0)) + Val(oDS_PH_PY115A.GetValue("U_JSUD01", 0))
    End Select
'/ 세법상의 근속년수
    If (FD_GNMMON Mod 12) = 0 Then
        FD_GNMYER = FD_GNMMON / 12
    Else
        FD_GNMYER = MDC_SetMod.IInt(FD_GNMMON / 12, 1) + 1
    End If


'/*******************************************************************************************************************************/
'/ 세액환산계산
'/ ; 세액환산명세,퇴직연금일시금지급예상액, 이연금액등이 있을경우 계산됨.
'/ ******************************************************************************************************************************/
    If HS.SHWJIG > 0 Then
        If Trim$(RPTGBN) = "Y" Then    '/ 명예퇴직리포트
            Select Case Trim$(HS_TYPE)
            Case "1"    '// 법정
                    '/ 22.총일시금(WK_SHWTIL) : 불입액이 있을 경우(지급예상액 * (1-(불입액-소득공제액)/총수령액)),없을 경우(지급예상액과 동일)
                    If FD_YUNWON <> 0 Then
                        HS.SHWTIL = MDC_SetMod.IInt(HS.SHWJIG * (1 - (FD_YUNBUL - FD_YUNGON) / FD_YUNWON), 1)
                    Else    '// 법정외
                        HS.SHWTIL = SHWJIG(1)
                    End If
            Case "2"
                    '/ 22.총일시금(WK_SHWTIL) :
                    HS.SHWTIL = SHWJIG(2)
            End Select
        Else
            '/ 22.총일시금(WK_SHWTIL) : 불입액이 있을 경우(지급예상액 * (1-(불입액-소득공제액)/총수령액)),없을 경우(지급예상액과 동일)
            If FD_YUNTOT <> 0 Then  '/ 퇴직연금불입액>0
                HS.SHWTIL = MDC_SetMod.IInt(HS.SHWJIG * (1 - (FD_YUNBUL - FD_YUNGON) / FD_YUNWON), 1)
            Else
                HS.SHWTIL = HS.SHWJIG
            End If
        End If
        '/ 23.수령가능퇴직급여액(WK_SHWSUR)= 총일시금+주현퇴직금+주현명예수당+종전퇴직금+종전명예수당
        HS.SHWSUR = HS.SHWTIL + FD_RETAMT
        
        '/ 24.환산퇴직소득공제=(45%공제액+근속년수별공제)
        Call Compute_TAX_2(HS.SHWSUR, FD_GNMYER, 0, HS.SHWJIG, HS.SHWTIL, HS.SHWSUR)
        HS.SHWGON = WK401.RETGON
        
    '/ 25.환산퇴직소득과세표준(HS.SHWGWA)
        HS.SHWGWA = WK401.TAXSTD
        
    '/ 26.환산연평균과세표준(HS.SHWYAG)
        HS.SHWYAG = WK401.YTXSTD
        If HS.SHWYAG < 0 Then HS.SHWYAG = 0
        
    '/ 27.환산연평균산출세액(HS.SHWYAS)
        HS.SHWYAS = WK401.YSANTX
    
    Else
        HS.SHWTIL = 0
        HS.SHWSUR = 0
        HS.SHWGON = 0
        HS.SHWGWA = 0
        HS.SHWYAG = 0
        HS.SHWYAS = 0
    End If

'/  Display '/*********************************************************************************************************************/
    Select Case HS_TYPE
    Case 1
        i = HS_TYPE
        oDS_PH_PY115A.setValue "U_SH" & i & "TIL", 0, HS.SHWTIL
        oDS_PH_PY115A.setValue "U_SH" & i & "SUR", 0, HS.SHWSUR
        oDS_PH_PY115A.setValue "U_SH" & i & "GON", 0, HS.SHWGON
        oDS_PH_PY115A.setValue "U_SH" & i & "GWA", 0, HS.SHWGWA
        oDS_PH_PY115A.setValue "U_SH" & i & "YAG", 0, HS.SHWYAG
        oDS_PH_PY115A.setValue "U_SH" & i & "YAS", 0, HS.SHWYAS
        oForm.Items("SH" & i & "TIL").Update
        oForm.Items("SH" & i & "SUR").Update
        oForm.Items("SH" & i & "GON").Update
        oForm.Items("SH" & i & "GWA").Update
        oForm.Items("SH" & i & "YAG").Update
        oForm.Items("SH" & i & "YAS").Update
    Case 2
        i = HS_TYPE
        oDS_PH_PY115A.setValue "U_SH" & i & "TIL", 0, HS.SHWTIL
        oForm.Items("SH" & i & "TIL").Update
    End Select
'/  ******************************************************************************************************************************/
    
End Sub



Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim oMast       As ZPAY_g_EmpID
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim STRDAT      As String
    Dim ENDDAT      As String
    
    Dim WK_JONYER   As Integer: Dim WK_JONMON   As Integer
    Dim WK_GNMDAY   As Integer: Dim WK_DUPMON   As Integer
    
    Dim JRET01      As Double
    Dim JYIL03      As Double
    Dim JTOT01      As Double
    Dim JADD01      As Double
    Dim JYIL04      As Double
    Dim JTOT02      As Double
    
    Set oCombo = oForm.Items("JSNGBN").Specific
    oForm.Freeze True
    Select Case oUID
        Case "MSTCOD"     '/ 사원번호
            If oForm.Items(oUID).Specific.String = "" Then
                oDS_PH_PY115A.setValue "U_MSTCOD", 0, ""
                oDS_PH_PY115A.setValue "U_MSTNAM", 0, ""
                oDS_PH_PY115A.setValue "U_EmpID", 0, ""
                oDS_PH_PY115A.setValue "U_INPDAT", 0, ""
                oDS_PH_PY115A.setValue "U_OUTDAT", 0, ""
                oDS_PH_PY115A.setValue "U_STRRET", 0, ""
                oDS_PH_PY115A.setValue "U_ENDRET", 0, ""
                
                oDS_PH_PY115A.setValue "U_DSDAT1", 0, ""
                oDS_PH_PY115A.setValue "U_DEDAT1", 0, ""
                oDS_PH_PY115A.setValue "U_DSDAT2", 0, ""
                oDS_PH_PY115A.setValue "U_DEDAT2", 0, ""
            Else
                oDS_PH_PY115A.setValue "U_MSTCOD", 0, UCase$(oForm.Items(oUID).Specific.String)
                oMast = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oDS_PH_PY115A.setValue "U_MSTNAM", 0, oMast.MSTNAM
                oDS_PH_PY115A.setValue "U_EmpID", 0, oMast.EmpID
                
                oDS_PH_PY115A.setValue "U_INPDAT", 0, IIf(GNSGBN = "1", oMast.GrpDat, oMast.StartDate)
                oDS_PH_PY115A.setValue "U_OUTDAT", 0, oMast.TermDate
                
                If oForm.Mode = fm_ADD_MODE Then
                    '/ 정산시작일자
                    
                    If Trim$(oMast.RETDAT) <> Trim$(oMast.ENDRET) Then
                        Sbo_Application.MessageBox ("중간정산일과 인사마스터의 기산일을 확인바랍니다.")
                    End If
                    If Trim$(oMast.RETDAT) = "" Or Left(Trim$(oMast.RETDAT), 4) = "1900" Then
                        oDS_PH_PY115A.setValue "U_STRRET", 0, IIf(GNSGBN = "1", oMast.GrpDat, oMast.StartDate)
                        
                        '/2012년12월31일 이전 기산일, 퇴사일
                        If IIf(GNSGBN = "1", oMast.GrpDat, oMast.StartDate) <= "20121231" Then
                            oDS_PH_PY115A.setValue "U_DSDAT1", 0, IIf(GNSGBN = "1", oMast.GrpDat, oMast.StartDate)
                            oDS_PH_PY115A.setValue "U_DEDAT1", 0, Format$("2012-12-31", "yyyymmdd")
                        End If
                    Else
                        STRDAT = DateAdd("d", 0, Format$(oMast.RETDAT, "0000-00-00"))
                        oDS_PH_PY115A.setValue "U_STRRET", 0, Replace(STRDAT, "-", "")
                        
                        If STRDAT <= "20121231" Then
                            '/2012년12월31일 이전 기산일, 퇴사일
                            oDS_PH_PY115A.setValue "U_DSDAT1", 0, Replace(STRDAT, "-", "")
                            oDS_PH_PY115A.setValue "U_DEDAT1", 0, Format$("2012-12-31", "yyyymmdd")
                        End If
                    End If
                    
                    
                    
                    
                    
                    
                   '/ 정산종료일자
                    If Trim$(oMast.TermDate) = "" Then
                        oDS_PH_PY115A.setValue "U_ENDRET", 0, Format$(Now, "yyyymmdd")
                        
                        '/2012년12월31일 이전 기산일, 퇴사일
                        
                        If oDS_PH_PY115A.GetValue("U_STRRET", 0) >= "20130101" Then
                            oDS_PH_PY115A.setValue "U_DSDAT2", 0, oDS_PH_PY115A.GetValue("U_STRRET", 0)
                        Else
                            oDS_PH_PY115A.setValue "U_DSDAT2", 0, Format$("2013-01-01", "yyyymmdd")
                        End If
                        oDS_PH_PY115A.setValue "U_DEDAT2", 0, Format$(Now, "yyyymmdd")
                        
                        Call oCombo.Select("1", psk_Index)
                        oDS_PH_PY115A.setValue "U_JSNNAM", 0, oCombo.Selected.Description
                        oForm.Items("ENDRET").Enabled = True
                        oForm.Items("ENDRET").CLICK ct_Regular
                        MsterChk = True
                    Else
                        oDS_PH_PY115A.setValue "U_ENDRET", 0, oMast.TermDate
                        
                        If oDS_PH_PY115A.GetValue("U_STRRET", 0) >= "20130101" Then
                            oDS_PH_PY115A.setValue "U_DSDAT2", 0, oDS_PH_PY115A.GetValue("U_STRRET", 0)
                        Else
                            '/2012년12월31일 이전 기산일, 퇴사일
                            oDS_PH_PY115A.setValue "U_DSDAT2", 0, Format$("2013-01-01", "yyyymmdd")
                        End If
                        oDS_PH_PY115A.setValue "U_DEDAT2", 0, oMast.TermDate
                    
                        
                        Call oCombo.Select("1", psk_ByValue)
                        oDS_PH_PY115A.setValue "U_JSNNAM", 0, oCombo.Selected.Description
                        oForm.Items("ENDRET").Enabled = False
                        MsterChk = False
                    End If
                    
                    
                    '/ 임원누진율
                    'If Trim$(RATUSE) = "Y" Then  '// 임원누진개월(율) 사용
                        Call Compute_APPRAT
                    'End If
                End If
            End If
            oForm.Items("MSTNAM").Update
            oForm.Items("EmpID").Update
            oForm.Items("INPDAT").Update
            oForm.Items("OUTDAT").Update
            oForm.Items("STRRET").Update
            oForm.Items("ENDRET").Update
            Call Compute_GNSDAY
            Call Compute_GNMYER
        Case "ENDRET"
            MsterChk = True
            
            'oDS_PH_PY115A.setValue "U_DSDAT2", 0, Format$("2013-01-01", "yyyymmdd")
            If oDS_PH_PY115A.GetValue("U_STRRET", 0) >= "20130101" Then
                oDS_PH_PY115A.setValue "U_DSDAT2", 0, oDS_PH_PY115A.GetValue("U_STRRET", 0)
            Else
                oDS_PH_PY115A.setValue "U_DSDAT2", 0, Format$("2013-01-01", "yyyymmdd")
            End If
                        
            oDS_PH_PY115A.setValue "U_DEDAT2", 0, oDS_PH_PY115A.GetValue("U_ENDRET", 0)
            
            Call Compute_GNSDAY
        Case "GNSYER", "GNSMON", "GNSDAY"
            Call Compute_GNMYER
        '// 평균임금 계산기간
        Case "PAYSD1", "PAYED1"
            If Trim$(oDS_PH_PY115A.GetValue("U_PAYED1", 0)) = "" Then
                oDS_PH_PY115A.setValue "U_PAYED1", 0, Mid$(oDS_PH_PY115A.GetValue("U_PAYSD1", 0), 1, 6) & MDC_SetMod.Lday(oDS_PH_PY115A.GetValue("U_PAYSD1", 0))
                oForm.Items("PAYED1").Update
            End If
            oDS_PH_PY115A.setValue "U_PAYDY1", 0, TermDay(oDS_PH_PY115A.GetValue("U_PAYSD1", 0), oDS_PH_PY115A.GetValue("U_PAYED1", 0))
            oForm.Items("PAYDY1").Update
        Case "PAYSD2", "PAYED2"
            If Trim$(oDS_PH_PY115A.GetValue("U_PAYED2", 0)) = "" Then
                oDS_PH_PY115A.setValue "U_PAYED2", 0, Mid$(oDS_PH_PY115A.GetValue("U_PAYSD2", 0), 1, 6) & MDC_SetMod.Lday(oDS_PH_PY115A.GetValue("U_PAYSD2", 0))
                oForm.Items("PAYED2").Update
            End If
            oDS_PH_PY115A.setValue "U_PAYDY2", 0, TermDay(oDS_PH_PY115A.GetValue("U_PAYSD2", 0), oDS_PH_PY115A.GetValue("U_PAYED2", 0))
            oForm.Items("PAYDY2").Update
        Case "PAYSD3", "PAYED3"
            If Trim$(oDS_PH_PY115A.GetValue("U_PAYED3", 0)) = "" Then
                oDS_PH_PY115A.setValue "U_PAYED3", 0, Mid$(oDS_PH_PY115A.GetValue("U_PAYSD3", 0), 1, 6) & MDC_SetMod.Lday(oDS_PH_PY115A.GetValue("U_PAYSD3", 0))
                oForm.Items("PAYED3").Update
            End If
            oDS_PH_PY115A.setValue "U_PAYDY3", 0, TermDay(oDS_PH_PY115A.GetValue("U_PAYSD3", 0), oDS_PH_PY115A.GetValue("U_PAYED3", 0))
            oForm.Items("PAYDY3").Update
        Case "PAYSD4", "PAYED4"
            If Trim$(oDS_PH_PY115A.GetValue("U_PAYED2", 0)) = "" Then
                oDS_PH_PY115A.setValue "U_PAYED4", 0, Mid$(oDS_PH_PY115A.GetValue("U_PAYSD4", 0), 1, 6) & MDC_SetMod.Lday(oDS_PH_PY115A.GetValue("U_PAYSD4", 0))
                oForm.Items("PAYED4").Update
            End If
            oDS_PH_PY115A.setValue "U_PAYDY4", 0, TermDay(oDS_PH_PY115A.GetValue("U_PAYSD4", 0), oDS_PH_PY115A.GetValue("U_PAYED4", 0))
            oForm.Items("PAYDY4").Update
        '// 종전 근무지 입사일, 퇴사일
        Case "JOTDAT", "JINDAT"
            STRDAT = oDS_PH_PY115A.GetValue("U_JINDAT", 0)
            ENDDAT = oDS_PH_PY115A.GetValue("U_JOTDAT", 0)
            If Trim$(STRDAT) <> "" And Trim$(ENDDAT) <> "" Then
                Call MDC_SetMod.Term2(STRDAT, ENDDAT)
                '/ 근무년월일 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                WK_JONYER = Val(ZPAY_GBL_GNSYER)
                WK_JONMON = Val(ZPAY_GBL_GNSMON)
                If Val(ZPAY_GBL_GNSDAY) > 0 Then
                    WK_JONMON = WK_JONMON + 1
                End If
                WK_GNMDAY = (WK_JONYER * 12) + WK_JONMON
                '/ 중복월수
                If STRDAT < oDS_PH_PY115A.GetValue("U_STRRET", 0) Then
                   STRDAT = oDS_PH_PY115A.GetValue("U_STRRET", 0)
                End If
                If ENDDAT > oDS_PH_PY115A.GetValue("U_ENDRET", 0) Then
                    ENDDAT = oDS_PH_PY115A.GetValue("U_ENDRET", 0)
                End If
                Call MDC_SetMod.Term2(STRDAT, ENDDAT)
                WK_DUPMON = Val(ZPAY_GBL_GNSYER) * 12
                If Val(ZPAY_GBL_GNSDAY) > 0 Then
                    WK_DUPMON = WK_DUPMON + Val(ZPAY_GBL_GNSMON) + 1
                Else
                    WK_DUPMON = WK_DUPMON + Val(ZPAY_GBL_GNSMON)
                End If
                
                oDS_PH_PY115A.setValue "U_JONYER", 0, WK_JONYER   '/ 종전근속년수
                oDS_PH_PY115A.setValue "U_JONMON", 0, WK_JONMON   '/ 종전근속월수
                oDS_PH_PY115A.setValue "U_GNMDAY", 0, WK_GNMDAY   '/ 종전세법상근무년수
                '/ 중복월수:
                oDS_PH_PY115A.setValue "U_JMMMON", 0, WK_DUPMON
            End If
        Case "JRET01", "JYIL03"
            JRET01 = oDS_PH_PY115A.GetValue("U_JRET01", 0)
            JYIL03 = oDS_PH_PY115A.GetValue("U_JYIL03", 0)
            JTOT01 = JRET01 + JYIL03
            oDS_PH_PY115A.setValue "U_JTOT01", 0, JTOT01
            oForm.Items("JTOT01").Update
        Case "JADD01", "JYIL04"
            JADD01 = oDS_PH_PY115A.GetValue("U_JADD01", 0)
            JYIL04 = oDS_PH_PY115A.GetValue("U_JYIL04", 0)
            JTOT02 = JADD01 + JYIL04
            oDS_PH_PY115A.setValue "U_JTOT02", 0, JTOT02
            oForm.Items("JTOT02").Update
        Case "JIGBIL"
            STRDAT = oDS_PH_PY115A.GetValue("U_JIGBIL", 0)
            '/ 신고월 설정하기
            If Trim$(STRDAT) <> "" Then
                Select Case Mid$(STRDAT, 7, 2)
                Case Is <= 10: oDS_PH_PY115A.setValue "U_SINYMM", 0, Left$(STRDAT, 6)
                Case Else:     oDS_PH_PY115A.setValue "U_SINYMM", 0, Format$(DateAdd("M", 1, Format$(STRDAT, "0000-00-00")), "YYYYMM")
                End Select
                oForm.Items("SINYMM").Update
            End If
    End Select
    oForm.Update
    oForm.Freeze False
End Sub

Private Sub PayTitle()

Dim i           As Long
Dim sQry        As String
Dim MSTCOD      As String
Dim oRecordSet  As SAPbobsCOM.Recordset

Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    MSTCOD = oDS_PH_PY115A.GetValue("U_MSTCOD", 0)
    '/급여내역 - 항목 설정
    sQry = "SELECT T0.U_CSUCOD, T0.U_CSUNAM"
    sQry = sQry & " FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE T1.U_CLTCOD =  (SELECT U_CLTCOD FROM [@PH_PY001A] WHERE Code = '" & MSTCOD & "')"
'        sQry = sQry & " WHERE T1.U_YM = '" & YM & "' OR (T1.U_YM <> '" & YM & "' AND T1.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY102A]))"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oDS_PH_PY115B.InsertRecord (i)
            oDS_PH_PY115B.Offset = i
            oDS_PH_PY115B.setValue "U_LineNum", i, i + 1
            oDS_PH_PY115B.setValue "U_CSUCOD", i, oRecordSet.Fields(0).VALUE
            oDS_PH_PY115B.setValue "U_CSUNAM", i, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next
    End If

    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
End Sub


