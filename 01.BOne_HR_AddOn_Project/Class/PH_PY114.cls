VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY114.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 퇴직금기준설정
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix
Public oMat2 As SAPbouiCOM.Matrix

Private oDS_PH_PY114A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY114B As SAPbouiCOM.DBDataSource
Private oDS_PH_PY114C As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Private CSUCOD(1 To 9) As String
Private CSUNAM(1 To 9) As String
Private SILCOD(1 To 27) As String
Private SILCUN(1 To 27) As String

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY114.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY114_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY114"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY114_CreateItems
    Call PH_PY114_EnableMenus
    Call PH_PY114_SetDocument(oFromDocEntry01)
'    Call PH_PY114_FormResize
    
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

Private Function PH_PY114_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    Dim oOption     As SAPbouiCOM.OptionBtn
    Dim oFolder     As SAPbouiCOM.Folder
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY114_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    

    Set oDS_PH_PY114A = oForm.DataSources.DBDataSources("@PH_PY114A")   '//헤더
    Set oDS_PH_PY114B = oForm.DataSources.DBDataSources("@PH_PY114B")   '//라인
    Set oDS_PH_PY114C = oForm.DataSources.DBDataSources("@PH_PY114C")   '//라인

    Set oMat1 = oForm.Items("Mat1").Specific
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    
    Set oMat2 = oForm.Items("Mat2").Specific
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns

    oForm.PaneLevel = 1

    '// UserDataSources
    With oForm.DataSources.UserDataSources
        .Add "MEDNAM", dt_LONG_TEXT
        .Add "KUKNAM", dt_LONG_TEXT
        .Add "JUNNAM", dt_LONG_TEXT
        .Add "BH1NAM", dt_LONG_TEXT
        .Add "BH2NAM", dt_LONG_TEXT
        .Add "CSUNAM", dt_LONG_TEXT
        .Add "GONNAM", dt_LONG_TEXT
    End With
    
    Set oEdit = oForm.Items("MEDNAM").Specific
    oEdit.DataBind.SetBound True, "", "MEDNAM"
    Set oEdit = oForm.Items("KUKNAM").Specific
    oEdit.DataBind.SetBound True, "", "KUKNAM"
    Set oEdit = oForm.Items("JUNNAM").Specific
    oEdit.DataBind.SetBound True, "", "JUNNAM"
    Set oEdit = oForm.Items("BH1NAM").Specific
    oEdit.DataBind.SetBound True, "", "BH1NAM"
    Set oEdit = oForm.Items("BH2NAM").Specific
    oEdit.DataBind.SetBound True, "", "BH2NAM"
    Set oEdit = oForm.Items("CSUNAM").Specific
    oEdit.DataBind.SetBound True, "", "CSUNAM"
    Set oEdit = oForm.Items("GONNAM").Specific
    oEdit.DataBind.SetBound True, "", "GONNAM"

    
    '// 1.평균임금산정방법
    Set oCombo = oForm.Items("FIXTYP").Specific
    oCombo.ValidValues.Add "M1", "월단위(3개월/월중도퇴사자 4개월)"
    oCombo.ValidValues.Add "M2", "월단위(3개월/월말퇴사:~당월, 월중퇴사:~전월)"
    oCombo.ValidValues.Add "M3", "월단위(3개월기준급여)"
    oCombo.ValidValues.Add "D1", "일단위(퇴사월말일+2개월말일:07.11.13~08.02.11)"
    oCombo.ValidValues.Add "D2", "일단위(퇴사일-이전90일까지: 07.11.14~08.02.11)"
    oCombo.ValidValues.Add "D3", "일단위(퇴사일-2개월퇴사전일:07.11.12~08.02.11)"
'    If oCombo.ValidValues.Count > 0 Then
'        oCombo.Select 0, psk_Index
'    End If
    '// 2.일할 끝전처리
    Set oCombo = oForm.Items("ROUNDT").Specific
    oCombo.ValidValues.Add "R", "반올림"
    oCombo.ValidValues.Add "F", "절사"
    oCombo.ValidValues.Add "C", "절상"

    Set oColumn = oMat1.Columns("ROUNDT")
    oColumn.ValidValues.Add "R", "반올림"
    oColumn.ValidValues.Add "F", "절사"
    oColumn.ValidValues.Add "C", "절상"
    '// 5.단위
    Set oCombo = oForm.Items("LENGTH").Specific
    oCombo.ValidValues.Add "1", "  원"
    oCombo.ValidValues.Add "10", "십원"
    oCombo.ValidValues.Add "100", "백원"
    oCombo.ValidValues.Add "1000", "천원"
    
    Set oColumn = oMat1.Columns("LENGTH")
    oColumn.ValidValues.Add "1", "  원"
    oColumn.ValidValues.Add "10", "십원"
    oColumn.ValidValues.Add "100", "백원"
    oColumn.ValidValues.Add "1000", "천원"
    
     '// 4.옵션버튼(비정기상여포함여부)
    Set oOption = oForm.Items("Opt11").Specific
    oOption.GroupWith ("Opt12")
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_RETCHK"
    
    Set oOption = oForm.Items("Opt12").Specific
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_RETCHK"
    oOption.GroupWith ("Opt11")
    
    Set oOption = oForm.Items("Opt11").Specific
    oOption.Selected = True
     '// 5.옵션버튼(1년미만자 포함여부)
    Set oOption = oForm.Items("Opt21").Specific
    oOption.GroupWith ("Opt22")
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_1YYCHK"
    
    Set oOption = oForm.Items("Opt22").Specific
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_1YYCHK"
    oOption.GroupWith ("Opt21")
    
    Set oOption = oForm.Items("Opt22").Specific
    oOption.Selected = True
     '// 6.옵션버튼(근속년수 기준)
    Set oOption = oForm.Items("Opt31").Specific
    oOption.GroupWith ("Opt32")
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_GNSGBN"
    
    Set oOption = oForm.Items("Opt32").Specific
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_GNSGBN"
    oOption.GroupWith ("Opt31")
    
    Set oOption = oForm.Items("Opt31").Specific
    oOption.Selected = True
    
    '// 7.옵션버튼(분개생성 방법)
    Set oOption = oForm.Items("Opt41").Specific
    oOption.GroupWith ("Opt42")
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_BILCHK"
    
    Set oOption = oForm.Items("Opt42").Specific
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_BILCHK"
    oOption.GroupWith ("Opt41")
    
    Set oOption = oForm.Items("Opt41").Specific
    oOption.Selected = True
    
    '// 8.옵션버튼(퇴사월 연차수당 처리)
    Set oOption = oForm.Items("Opt51").Specific
    oOption.GroupWith ("Opt52")
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_RETYCH"
    
    Set oOption = oForm.Items("Opt52").Specific
    oOption.DataBind.SetBound True, "@PH_PY114A", "U_RETYCH"
    oOption.GroupWith ("Opt51")
    
    Set oOption = oForm.Items("Opt51").Specific
    oOption.Selected = True
    
    '/ 초기값 셋팅
    For i = 1 To 9
        CSUCOD(i) = "A0" & i
    Next i
    CSUNAM(1) = "평균급여"
    CSUNAM(2) = "평균상여"
    CSUNAM(3) = "평균연차"
    CSUNAM(4) = "평균체력단련비"
    CSUNAM(5) = "평균임금"
    CSUNAM(6) = "연퇴직금"
    CSUNAM(7) = "월퇴직금"
    CSUNAM(8) = "일퇴직금"
    CSUNAM(9) = "퇴직금계"
    
    
    '/ M1 월단위계산수식
    SILCOD(1) = "P01":                          SILCUN(1) = "총급여"
    SILCOD(2) = "P02  / 12 * 3":                SILCUN(2) = "총상여 / 12 * 3"
    SILCOD(3) = "P03  / 12 * 3":                SILCUN(3) = "총연차 / 12 * 3"
    SILCOD(4) = "P04  / 12 * 3":                SILCUN(4) = "총체력단련비 / 12 * 3"
    SILCOD(5) = "( A01 + A02 + A03 + A04 ) / X01":    SILCUN(5) = "(평균급여+평균상여+평균연차+평균체력단련비) / 급여기간일수"
    SILCOD(6) = "":                             SILCUN(6) = ""
    SILCOD(7) = "":                             SILCUN(7) = ""
    SILCOD(8) = "A05 * 30 * X02 / 365":         SILCUN(8) = "평균임금*30*총근속일수/365"
    SILCOD(9) = "A06 + A07 + A08":              SILCUN(9) = "연퇴직금+월퇴직금+일퇴직금"
    '/ D1 일단위계산수식
    SILCOD(10) = "P01":                         SILCUN(10) = "총급여"
    SILCOD(11) = "P02  / 12 * 3":               SILCUN(11) = "총상여 / 12 * 3"
    SILCOD(12) = "P03  / 12 * 3":               SILCUN(12) = "총연차 / 12 * 3"
    SILCOD(13) = "P04  / 12 * 3":               SILCUN(13) = "총체력단련비 / 12 * 3"
    SILCOD(14) = "( A01 + A02 + A03 + A04 ) / X01":   SILCUN(14) = "(평균급여+평균상여+평균연차+평균체력단련비) / 급여기간일수수"
    SILCOD(15) = "":                            SILCUN(15) = ""
    SILCOD(16) = "":                            SILCUN(16) = ""
    SILCOD(17) = "A05 * 30 * X02 / 365":        SILCUN(17) = "평균임금*30*총근속일수/365"
    SILCOD(18) = "A06 + A07 + A08":             SILCUN(18) = "연퇴직금+월퇴직금+일퇴직금"
    '/ D2 일단위계산수식
    SILCOD(19) = "P01":                         SILCUN(19) = "총급여"
    SILCOD(20) = "P02  / 12 * 3":               SILCUN(20) = "총상여 / 12 * 3"
    SILCOD(21) = "P03  / 12 * 3":               SILCUN(21) = "총연차 / 12 * 3"
    SILCOD(22) = "P04  / 12 * 3":               SILCUN(22) = "총체력단련비 / 12 * 3"
    SILCOD(23) = "( A01 + A02 + A03 + A04 ) / X01":   SILCUN(23) = "(평균급여+평균상여+평균연차+평균체력단련비) / 급여기간일수"
    SILCOD(24) = "":                            SILCUN(24) = ""
    SILCOD(25) = "":                            SILCUN(25) = ""
    SILCOD(26) = "A05 * 30 * X02 / 365":        SILCUN(26) = "평균임금*30*총근속일수/365"
    SILCOD(27) = "A06 + A07 + A08":             SILCUN(27) = "연퇴직금+월퇴직금+일퇴직금"
    
'/ Folder /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    oForm.DataSources.UserDataSources.Add "FolderDS1", dt_SHORT_TEXT, 1
    For i = 1 To 3
        Set oFolder = oForm.Items("Folder" & i).Specific
        oFolder.DataBind.SetBound True, "", "FolderDS1"
        If i = 1 Then
            oFolder.Select
        Else
            oFolder.GroupWith ("Folder" & i - 1)
        End If
    Next i
      
    oForm.Items("Folder1").AffectsFormMode = False
    oForm.Items("Folder2").AffectsFormMode = False
    oForm.Items("Folder3").AffectsFormMode = False
    oForm.Items("Folder1").Enabled = True
    oForm.Items("Folder2").Enabled = True
    oForm.Items("Folder3").Enabled = True

    
    '// 직위
    Set oColumn = oMat2.Columns("MSTSTP")
    sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
    Loop
 
 '/ 귀속연월
    oForm.Items("Code").Specific.Value = Format$(Now, "YYYY")
    
     '// 누진계산사용유무
    Set oCombo = oForm.Items("RATUSE").Specific
    oCombo.ValidValues.Add "N", "미사용"
    oCombo.ValidValues.Add "Y", "사용"

    
    Set oFolder = Nothing
    Set oOption = Nothing
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY114_CreateItems_Error:
    Set oFolder = Nothing
    Set oOption = Nothing
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY114_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY114_EnableMenus()
    
    On Error GoTo PH_PY114_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY114_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY114_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY114_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY114_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY114_FormItemEnabled
        Call PH_PY114_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY114_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY114_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY114_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY114_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY114_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("Code").Enabled = True
        oForm.Items("Opt11").Specific.Selected = True
        oForm.Items("Opt22").Specific.Selected = True
        oForm.Items("Opt31").Specific.Selected = True
        oForm.Items("Opt42").Specific.Selected = True
        oForm.Items("Opt51").Specific.Selected = True
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("Code").Enabled = True
        oForm.Items("Opt11").Specific.Selected = True
        oForm.Items("Opt22").Specific.Selected = True
        oForm.Items("Opt31").Specific.Selected = True
        oForm.Items("Opt42").Specific.Selected = True
        oForm.Items("Opt51").Specific.Selected = True
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("Code").Enabled = False
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY114_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY114_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY114_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            If PH_PY114_DataValidCheck = False Then
                                BubbleEvent = False
                            End If
                        End If
                    End If
                ElseIf pval.ItemUID = "Folder1" Then
                    oForm.PaneLevel = 1
                ElseIf pval.ItemUID = "Folder2" Then
                    oForm.PaneLevel = 2
                ElseIf pval.ItemUID = "Folder3" Then
                    oForm.PaneLevel = 3
                End If
            Else
                If pval.ItemUID = "1" And pval.ActionSuccess = True And oForm.Mode = fm_ADD_MODE Then
'                    oForm.Mode = fm_OK_MODE
'                    Call Sbo_Application.ActivateMenuItem("1282")
                End If
            End If
        
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE Then
            
                If pval.ItemUID = "JOBDAT" Then
                   If oMat2.RowCount > 0 Then
                       oMat2.Columns("Col2").Cells(oMat1.VisualRowCount).CLICK ct_Regular
                       BubbleEvent = False
                   End If
                End If
            End If
            If pval.BeforeAction = True And pval.CharPressed = 9 Then
                If pval.ItemUID = "MEDACC" Then
                    If MDC_SetMod.Value_ChkYn("OACT", "FORMATCODE", "'" & oForm.Items("MEDACC").Specific.Value & "'") = True Or _
                        oForm.Items("MEDACC").Specific.String = "" Then
                        oForm.Items("MEDACC").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                ElseIf pval.ItemUID = "KUKACC" Then
                    If MDC_SetMod.Value_ChkYn("OACT", "FORMATCODE", "'" & oForm.Items("KUKACC").Specific.Value & "'") = True Or _
                        oForm.Items("KUKACC").Specific.String = "" Then
                        oForm.Items("KUKACC").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                ElseIf pval.ItemUID = "JUNACC" Then
                    If MDC_SetMod.Value_ChkYn("OACT", "FORMATCODE", "'" & oForm.Items("JUNACC").Specific.Value & "'") = True Or _
                        oForm.Items("JUNACC").Specific.String = "" Then
                        oForm.Items("JUNACC").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                ElseIf pval.ItemUID = "BH1ACC" Then
                    If MDC_SetMod.Value_ChkYn("OACT", "FORMATCODE", "'" & oForm.Items("BH1ACC").Specific.Value & "'") = True Or _
                        oForm.Items("BH1ACC").Specific.String = "" Then
                        oForm.Items("BH1ACC").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                ElseIf pval.ItemUID = "BH2ACC" Then
                    If MDC_SetMod.Value_ChkYn("OACT", "FORMATCODE", "'" & oForm.Items("BH2ACC").Specific.Value & "'") = True Or _
                        oForm.Items("BH2ACC").Specific.String = "" Then
                        oForm.Items("BH2ACC").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                ElseIf pval.ItemUID = "CSUACC" Then
                    If MDC_SetMod.Value_ChkYn("OACT", "FORMATCODE", "'" & oForm.Items("CSUACC").Specific.Value & "'") = True Or _
                        oForm.Items("CSUACC").Specific.String = "" Then
                        oForm.Items("CSUACC").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                ElseIf pval.ItemUID = "GONACC" Then
                    If MDC_SetMod.Value_ChkYn("OACT", "FORMATCODE", "'" & oForm.Items("GONACC").Specific.Value & "'") = True Or _
                        oForm.Items("GONACC").Specific.String = "" Then
                        oForm.Items("GONACC").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
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
            If pval.ItemChanged = True And pval.ItemUID = "FIXTYP" Then
                If pval.BeforeAction = False Then
                    oForm.Items("FIXTYP").Update
                    Call Display_PH_PY114B
                    If Left$(oDS_PH_PY114A.GetValue("U_FIXTYP", 0), 1) = "M" Then
                       oDS_PH_PY114A.setValue "U_ROUNDT", 0, ""
                       oDS_PH_PY114A.setValue "U_LENGTH", 0, ""
                       oForm.Items("ROUNDT").Enabled = False
                       oForm.Items("LENGTH").Enabled = False
                       oForm.Items("ROUNDT").Update
                       oForm.Items("LENGTH").Update
                    Else
                       oForm.Items("ROUNDT").Enabled = True
                       oForm.Items("LENGTH").Enabled = True
                    End If
                End If
'            ElseIf pval.ItemChanged = True And pval.ItemUID = "Mat2" And pval.ColUID = "Col1" Then
'                If pval.BeforeAction = False Then
'                    FlushToItemValue pval.ColUID, pval.Row
'                End If
            End If
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat1", "Grid1"
                    If pval.Row > 0 Then
                        '//Call oMat1.SelectRow(pval.Row, True, False)

                        oLastItemUID = pval.ItemUID
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Case Else
                    oLastItemUID = pval.ItemUID
                    oLastColUID = ""
                    oLastColRow = 0
                End Select
            ElseIf pval.BeforeAction = False Then

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
                If pval.ItemUID = "Code" Then

                ElseIf pval.ItemUID = "Mat2" _
                    And (pval.ColUID = "Col1" Or pval.ColUID = "Col2") Then
                    oMat2.FlushToDataSource
                    oDS_PH_PY114C.Offset = pval.Row - 1
                    oMat2.SetLineData pval.Row
        
                    If pval.Row = oMat2.RowCount And Trim$(oDS_PH_PY114C.GetValue("U_MSTSTP", pval.Row - 1)) <> "" Then
                        Call PH_PY114_AddMatrixRow
                        oMat2.Columns("Col2").Cells(pval.Row).CLICK ct_Regular
                    End If
                ElseIf pval.ItemUID = "MEDACC" Or pval.ItemUID = "KUKACC" Or pval.ItemUID = "JUNACC" _
                    Or pval.ItemUID = "BH1ACC" Or pval.ItemUID = "BH2ACC" Or pval.ItemUID = "CSUACC" _
                    Or pval.ItemUID = "GONACC" Then
                    
                    If oDS_PH_PY114A.GetValue("U_" & pval.ItemUID, 0) = "" Then
                        oForm.DataSources.UserDataSources(pval.ItemUID).ValueEx = ""
                    Else
                        oForm.DataSources.UserDataSources(pval.ItemUID).ValueEx = _
                        MDC_SetMod.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" & oDS_PH_PY114A.GetValue("U_" & pval.ItemUID, 0) & "'")
                    End If
                End If
            End If
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = False And pval.ItemUID = "Mat2" And oForm.Mode <> fm_FIND_MODE Then
'                Call PH_PY114_FormItemEnabled
                oMat1.LoadFromDataSource
                oMat2.LoadFromDataSource
                Call PH_PY114_AddMatrixRow
                
'                If oMat2.VisualRowCount >= 1 Then
'                    If Trim$(oDS_PH_PY114C.GetValue("U_MSTSTP", oMat2.VisualRowCount - 1)) <> "" Then
'                        Call PH_PY114_AddMatrixRow
'                    End If
'                Else
'                    If oDS_PH_PY114C.Size = 1 Then
'                        Call PH_PY114_AddMatrixRow
'                    End If
'                End If
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
                Set oDS_PH_PY114A = Nothing
                Set oDS_PH_PY114B = Nothing
                Set oDS_PH_PY114C = Nothing

                Set oMat1 = Nothing
                Set oMat2 = Nothing
                
            End If
'----------------------------------------------------------
        Case et_FORM_ACTIVATE: '//18
            
'----------------------------------------------------------
        Case et_FORM_DEACTIVATE: '//19
            
'----------------------------------------------------------
        Case et_FORM_CLOSE '//20
            
'----------------------------------------------------------
        Case et_FORM_RESIZE '//21
            If pval.BeforeAction = True Then

            ElseIf pval.BeforeAction = False Then
                oForm.Items.Item("Rec2").Width = oForm.Items.Item("Mat1").Width + 20
                oForm.Items.Item("Rec2").Height = oForm.Items.Item("Mat1").Height + 10
                
                oForm.Items.Item("Rec1").Width = oForm.Items.Item("Rec2").Width - oForm.Items.Item("Rec1").Left + 20
                
                oForm.Items.Item("Rec3").Top = oForm.Items.Item("Rec2").Top + oForm.Items.Item("Rec2").Height + 22
                oForm.Items.Item("Rec3").Width = oForm.Items.Item("Rec2").Width
                
            End If
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
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY114_FormItemEnabled
                Call PH_PY114_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY114_FormItemEnabled
                Call PH_PY114_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY114_FormItemEnabled
                Call PH_PY114_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY114_FormItemEnabled
            Case "1293" '// 행삭제
                '// [MAT1 용]
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource

                    While (i <= oDS_PH_PY114B.Size - 1)
                        If oDS_PH_PY114B.GetValue("U_CSUCOD", i) = "" Then
                            oDS_PH_PY114B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend

                    For i = 0 To oDS_PH_PY114B.Size
                        Call oDS_PH_PY114B.setValue("U_LineNum", i, i + 1)
                    Next i

                    oMat1.LoadFromDataSource
                End If
                '// [MAT2 용]
                 If oMat2.RowCount <> oMat2.VisualRowCount Then
                    oMat2.FlushToDataSource

                    While (i <= oDS_PH_PY114B.Size - 1)
                        If oDS_PH_PY114B.GetValue("U_MSTSTP", i) = "" Then
                            oDS_PH_PY114B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend

                    For i = 0 To oDS_PH_PY114B.Size
                        Call oDS_PH_PY114B.setValue("U_LineNum", i, i + 1)
                    Next i

                    oMat2.LoadFromDataSource
                End If
                Call PH_PY114_AddMatrixRow
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

Sub PH_PY114_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY114_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 용]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY114B.GetValue("U_CSUCOD", oRow - 1)) <> "" Then
'            If oDS_PH_PY114B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY114B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY114B.Offset = oRow
'            oDS_PH_PY114B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY114B.setValue "U_CSUCOD", oRow, ""
'            oDS_PH_PY114B.setValue "U_CSUNAM", oRow, ""
'            oDS_PH_PY114B.setValue "U_SILCUN", oRow, ""
'            oDS_PH_PY114B.setValue "U_ROUNDT", oRow, ""
'            oDS_PH_PY114B.setValue "U_LENGTH", oRow, ""
'            oDS_PH_PY114B.setValue "U_SILCOD", oRow, ""
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY114B.Offset = oRow - 1
'            oDS_PH_PY114B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY114B.setValue "U_CSUCOD", oRow - 1, ""
'            oDS_PH_PY114B.setValue "U_CSUNAM", oRow - 1, ""
'            oDS_PH_PY114B.setValue "U_SILCUN", oRow - 1, ""
'            oDS_PH_PY114B.setValue "U_ROUNDT", oRow - 1, ""
'            oDS_PH_PY114B.setValue "U_LENGTH", oRow - 1, ""
'            oDS_PH_PY114B.setValue "U_SILCOD", oRow - 1, ""
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY114B.Offset = oRow
'        oDS_PH_PY114B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY114B.setValue "U_CSUCOD", oRow, ""
'        oDS_PH_PY114B.setValue "U_CSUNAM", oRow, ""
'        oDS_PH_PY114B.setValue "U_SILCUN", oRow, ""
'        oDS_PH_PY114B.setValue "U_ROUNDT", oRow, ""
'        oDS_PH_PY114B.setValue "U_LENGTH", oRow, ""
'        oDS_PH_PY114B.setValue "U_SILCOD", oRow, ""
'        oMat1.LoadFromDataSource
'    End If

    '//[Mat2 용]
    oMat2.FlushToDataSource
    oRow = oMat2.VisualRowCount

    If oMat2.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY114C.GetValue("U_MSTSTP", oRow - 1)) <> "" Then
            If oDS_PH_PY114C.Size <= oMat2.VisualRowCount Then
                oDS_PH_PY114C.InsertRecord (oRow)
            End If
            oDS_PH_PY114C.Offset = oRow
            oDS_PH_PY114C.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY114C.setValue "U_MSTSTP", oRow, ""
            oDS_PH_PY114C.setValue "U_ADDRAT", oRow, 0
            oMat2.LoadFromDataSource
        Else
            oDS_PH_PY114C.Offset = oRow - 1
            oDS_PH_PY114C.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY114C.setValue "U_MSTSTP", oRow - 1, ""
            oDS_PH_PY114C.setValue "U_ADDRAT", oRow - 1, 0
            oMat2.LoadFromDataSource
        End If
    ElseIf oMat2.VisualRowCount = 0 Then
        oDS_PH_PY114C.Offset = oRow
        oDS_PH_PY114C.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY114C.setValue "U_MSTSTP", oRow, ""
        oDS_PH_PY114C.setValue "U_ADDRAT", oRow, 0
        oMat2.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY114_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY114_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY114_FormClear()
On Error GoTo PH_PY114_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY114'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY114_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY114_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY114_DataValidCheck() As Boolean
On Error GoTo PH_PY114_DataValidCheck_Error
    PH_PY114_DataValidCheck = False
    Dim i               As Long
    Dim k               As Long
    Dim sQry            As String
    Dim DocNum          As String
    Dim Chk_Data        As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '// 헤더체크
    Select Case True
    Case Trim$(oDS_PH_PY114A.GetValue("Code", 0)) = ""
        PH_PY114_DataValidCheck = False
        Sbo_Application.StatusBar.SetText "코드는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Success
        Exit Function
    Case oDS_PH_PY114A.GetValue("U_FIXTYP", 0) = ""
        PH_PY114_DataValidCheck = False
        Sbo_Application.StatusBar.SetText "평균임금 산정방법은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Success
        Exit Function
    Case Left$(oDS_PH_PY114A.GetValue("U_FIXTYP", 0), 1) = "D" And Trim$(oDS_PH_PY114A.GetValue("U_ROUNDT", 0)) = ""
        PH_PY114_DataValidCheck = False
        Sbo_Application.StatusBar.SetText "일할 끝전처리방법을 선택하여 주십시오.", bmt_Short, smt_Error
        Exit Function
    Case Left$(oDS_PH_PY114A.GetValue("U_FIXTYP", 0), 1) = "D" And Trim$(oDS_PH_PY114A.GetValue("U_LENGTH", 0)) = ""
        PH_PY114_DataValidCheck = False
        Sbo_Application.StatusBar.SetText "일할 단위를 선택하여 주십시오.", bmt_Short, smt_Error
        Exit Function
    End Select
    
    DocNum = Exist_YN(oDS_PH_PY114A.GetValue("Code", 0))
    If Trim$(DocNum) <> "" And Trim$(oDS_PH_PY114A.GetValue("Code", 0)) <> Trim$(DocNum) Then
        Sbo_Application.StatusBar.SetText "기존의 데이터가 있습니다. 확인하여 주십시오.", bmt_Short, smt_Error
        PH_PY114_DataValidCheck = False
        Exit Function
    End If
    
    
    '// 라인체크
    
    oMat1.FlushToDataSource
    oMat2.FlushToDataSource
        
    If oMat1.RowCount = 1 Then
        Sbo_Application.StatusBar.SetText "입력할 데이터가 없습니다. 입력하여 주십시오.", bmt_Short, smt_Error
        PH_PY114_DataValidCheck = False
        Exit Function
    End If
    
    For i = 0 To oMat1.VisualRowCount - 2
        oDS_PH_PY114B.Offset = i
        If Trim$(oDS_PH_PY114B.GetValue("U_CSUCOD", i)) = "" Then
            Sbo_Application.StatusBar.SetText "코드는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
            oMat1.Columns("Col4").Cells(i + 1).CLICK ct_Regular
            PH_PY114_DataValidCheck = False
            Exit Function
        Else
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '중복체크작업
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            Chk_Data = Trim$(oDS_PH_PY114B.GetValue("U_CSUCOD", i))
            For k = i + 1 To oMat1.VisualRowCount - 2
                oDS_PH_PY114B.Offset = k
                If Trim$(Chk_Data) = Trim$(oDS_PH_PY114B.GetValue("U_CSUCOD", k)) Then
                    Sbo_Application.StatusBar.SetText "내용이 중복입력되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error
                    oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
                    PH_PY114_DataValidCheck = False
                    Exit Function
                End If
            Next k
        End If
    Next i
'    If Trim$(oDS_PH_PY114B.GetValue("U_MSTSTP", oDS_PH_PY114B.Size - 1)) = "" Then
'        oDS_PH_PY114B.RemoveRecord oDS_PH_PY114B.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
'    End If
    
'// Mat2
    If oMat2.RowCount > 1 Then
        If oDS_PH_PY114A.GetValue("U_RATUSE", 0) = "Y" Then
            For i = 0 To oMat2.VisualRowCount - 2
                oDS_PH_PY114C.Offset = i
                If Trim$(oDS_PH_PY114C.GetValue("U_MSTSTP", i)) = "" Then
                    Sbo_Application.StatusBar.SetText "직위는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
                    oMat2.Columns("Col2").Cells(i + 1).CLICK ct_Regular
                    PH_PY114_DataValidCheck = False
                    Exit Function
                Else
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    '중복체크작업
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    Chk_Data = Trim$(oDS_PH_PY114C.GetValue("U_MSTSTP", i))
                    For k = i + 1 To oMat2.VisualRowCount - 2
                        oDS_PH_PY114C.Offset = k
                        If Trim$(Chk_Data) = Trim$(oDS_PH_PY114C.GetValue("U_MSTSTP", k)) Then
                            Sbo_Application.StatusBar.SetText "임원누진 내용이 중복입력되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error
                            oMat2.Columns("Col2").Cells(i + 1).CLICK ct_Regular
                            PH_PY114_DataValidCheck = False
                            Exit Function
                        End If
                    Next k
                End If
            Next i
        End If
        If Trim$(oDS_PH_PY114C.GetValue("U_MSTSTP", oDS_PH_PY114C.Size - 1)) = "" Then
            oDS_PH_PY114C.RemoveRecord oDS_PH_PY114C.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
        End If
    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat1.LoadFromDataSource
    oMat2.LoadFromDataSource
    
    PH_PY114_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY114_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY114_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY114_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY114_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY114_Validate_Error
    PH_PY114_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY114A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY114_Validate = False
        GoTo PH_PY114_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY114_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY114_Validate_Error:
    PH_PY114_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY114_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PH_PY114_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    
    Dim BPLID           As String
    Dim ItmBsort        As String
    Dim DocDate         As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY114_Print_Report01_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY114_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[S142] 발주서"
    ReportName = "S142_1.rpt"
    sQry = "EXEC PH_PY114_1 '" & oForm.Items("8").Specific.Value & "'"
    
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
    
PH_PY114_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY114_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Account_Name()
    
    Dim MEDACC$, KUKACC$, JUNACC$, BH1ACC$, BH2ACC$, CSUACC$, GONACC$
    
    MEDACC = oDS_PH_PY114A.GetValue("U_MEDACC", 0)
    KUKACC = oDS_PH_PY114A.GetValue("U_KUKACC", 0)
    JUNACC = oDS_PH_PY114A.GetValue("U_JUNACC", 0)
    BH1ACC = oDS_PH_PY114A.GetValue("U_BH1ACC", 0)
    BH2ACC = oDS_PH_PY114A.GetValue("U_BH2ACC", 0)
    CSUACC = oDS_PH_PY114A.GetValue("U_CSUACC", 0)
    GONACC = oDS_PH_PY114A.GetValue("U_GONACC", 0)
    
    If MEDACC = "" Then
        oForm.DataSources.UserDataSources("MEDNAM").ValueEx = ""
    Else
        oForm.DataSources.UserDataSources("MEDNAM").ValueEx = _
        MDC_SetMod.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" & MEDACC & "'")
    End If
    
    If KUKACC = "" Then
        oForm.DataSources.UserDataSources("KUKNAM").ValueEx = ""
    Else
        oForm.DataSources.UserDataSources("KUKNAM").ValueEx = _
        MDC_SetMod.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" & KUKACC & "'")
    End If
    
    If JUNACC = "" Then
        oForm.DataSources.UserDataSources("JUNNAM").ValueEx = ""
    Else
        oForm.DataSources.UserDataSources("JUNNAM").ValueEx = _
        MDC_SetMod.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" & JUNACC & "'")
    End If
    
    If BH1ACC = "" Then
        oForm.DataSources.UserDataSources("BH1NAM").ValueEx = ""
    Else
        oForm.DataSources.UserDataSources("BH1NAM").ValueEx = _
        MDC_SetMod.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" & BH1ACC & "'")
    End If
    
    If BH2ACC = "" Then
        oForm.DataSources.UserDataSources("BH2NAM").ValueEx = ""
    Else
        oForm.DataSources.UserDataSources("BH2NAM").ValueEx = _
        MDC_SetMod.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" & BH2ACC & "'")
    End If
    
    If CSUACC = "" Then
        oForm.DataSources.UserDataSources("CSUNAM").ValueEx = ""
    Else
        oForm.DataSources.UserDataSources("CSUNAM").ValueEx = _
        MDC_SetMod.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" & CSUACC & "'")
    End If
    
    If GONACC = "" Then
        oForm.DataSources.UserDataSources("GONNAM").ValueEx = ""
    Else
        oForm.DataSources.UserDataSources("GONNAM").ValueEx = _
        MDC_SetMod.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" & GONACC & "'")
    End If
    
End Sub

Private Sub Display_PH_PY114B()
On Error GoTo Error_Message
    Dim FIXTYP          As String
    Dim i               As Integer
    Dim cnt             As Integer
'/ Matrix2 초기화
    FIXTYP = oForm.Items("FIXTYP").Specific.Selected.Value
 '/ Question
    If Sbo_Application.MessageBox("기존 퇴직금 계산식을 초기 셋팅용 계산식으로 변경하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
        Exit Sub
    End If
    
    cnt = oDS_PH_PY114B.Size
    If cnt > 0 Then
        For i = 1 To cnt - 1
            oDS_PH_PY114B.RemoveRecord oDS_PH_PY114B.Size - 1
        Next i
    Else
        oMat1.LoadFromDataSource
    End If
    Select Case Trim$(FIXTYP)
    Case "M1": cnt = 1
    Case "D1": cnt = 10
    Case Else: cnt = 19
    End Select
    For i = 0 To 7
        If i + 1 > oDS_PH_PY114B.Size Then
            oDS_PH_PY114B.InsertRecord (i)
        End If
        oDS_PH_PY114B.Offset = i
        oDS_PH_PY114B.setValue "U_LineNum", i, i + 1
        oDS_PH_PY114B.setValue "U_CSUCOD", i, CSUCOD(i + 1)
        oDS_PH_PY114B.setValue "U_CSUNAM", i, CSUNAM(i + 1)
        oDS_PH_PY114B.setValue "U_SILCOD", i, SILCOD(i + cnt)
        oDS_PH_PY114B.setValue "U_SILCUN", i, SILCUN(i + cnt)
        oDS_PH_PY114B.setValue "U_ROUNDT", i, "R"
        oDS_PH_PY114B.setValue "U_LENGTH", i, "1"
    Next i
    oMat1.LoadFromDataSource
    
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Sbo_Application.StatusBar.SetText "Display_ZPY403 Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub



Private Function Exist_YN(STDDAT$) As String
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'저장할 데이터의 기존데이터가 있는지 확인한다
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT Top 1 T1.CODE FROM [@PH_PY114A] T1 where code = '" & Trim$(STDDAT) & "'"
    oRecordSet.DoQuery sQry
    
    Do Until oRecordSet.EOF
        Exist_YN = oRecordSet(0).Value
        oRecordSet.MoveNext
    Loop
    
    If Trim$(Exist_YN) = "" Then
        Exist_YN = ""
        Exit Function
    End If
    
    Set oRecordSet = Nothing
End Function
