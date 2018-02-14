VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY005.cls
'//  Module         : 인사관리 > 인사관리
'//  Desc           : 사업장정보등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Private oDS_PH_PY005A As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY005.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY005_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY005"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY005_CreateItems
    Call PH_PY005_EnableMenus
    Call PH_PY005_SetDocument(oFromDocEntry01)
'    Call PH_PY005_FormResize
    
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

Private Function PH_PY005_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY005_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY005A = oForm.DataSources.DBDataSources("@PH_PY005A")

'    Set oEdit = oForm.Items("TaxDGbn").Specific
'    oEdit.DataBind
    
    '//------------------------------------------------------------------------------------------------
    '//일반
    '//------------------------------------------------------------------------------------------------
    
    '//제출인구분
    Set oCombo = oForm.Items("TaxDGbn").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P006' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
'    oCombo.Select 0, psk_Index
    oForm.Items("TaxDGbn").DisplayDesc = True

    '//자동체번 생성규칙
'    Set oCombo = oForm.Items("AutoChk").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
''    oCombo.Select 0, psk_Index
'    oForm.Items("AutoChk").DisplayDesc = True
    
    '// 납세자구분
    Set oCombo = oForm.Items("BUSTYP").Specific
    oCombo.ValidValues.Add "1", "개인"
    oCombo.ValidValues.Add "8", "법인"
'    oCombo.Select 0, psk_Index
    oForm.Items("BUSTYP").DisplayDesc = True
    
    '// 사업자단위과세
    Set oCombo = oForm.Items("SAUPJA").Specific
    oCombo.ValidValues.Add "N", "N"
    oCombo.ValidValues.Add "Y", "Y"
'    oCombo.Select 0, psk_Index
    oForm.Items("SAUPJA").DisplayDesc = True
    
    '// 원천신고구분
    Set oCombo = oForm.Items("SINTYP").Specific
    oCombo.ValidValues.Add "1", "매월"
    oCombo.ValidValues.Add "2", "반기"
'    oCombo.Select 0, psk_Index
    oForm.Items("SINTYP").DisplayDesc = True

    '// 일괄납부여부
    Set oCombo = oForm.Items("ILGTYP").Specific
    oCombo.ValidValues.Add "1", "부"
    oCombo.ValidValues.Add "2", "여"
    oCombo.Select 0, psk_Index
    oForm.Items("ILGTYP").DisplayDesc = True
    
    '//자동체번 생성규칙
    Set oCombo = oForm.Items("WCHCLT").Specific
    sQry = "SELECT BPLId, BPLName FROM [OBPL]"
    Call SetReDataCombo(oForm, sQry, oCombo)
    
    oForm.Items("WCHCLT").DisplayDesc = True

    '//자동체번 생성규칙
    Set oCombo = oForm.Items("SUPCLT").Specific
    sQry = "SELECT BPLId, BPLName FROM [OBPL]"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("SUPCLT").DisplayDesc = True

    '//자동체번 생성규칙
    Set oCombo = oForm.Items("JUMCLT").Specific
    sQry = "SELECT BPLId, BPLName FROM [OBPL]"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JUMCLT").DisplayDesc = True
    
    '//------------------------------------------------------------------------------------------------
    '//세부정보
    '//------------------------------------------------------------------------------------------------
    '//사원번호구성체계
    Set oCombo = oForm.Items("AutoChk").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P005' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("AutoChk").DisplayDesc = True
    
    '// 이행상황신고서집계방법
    Set oCombo = oForm.Items("WCHTYP").Specific
    oCombo.ValidValues.Add "1", "귀속연월"
    oCombo.ValidValues.Add "2", "지급연월"
    oCombo.ValidValues.Add "3", "귀속연월 OR 지급연월"
    oCombo.ValidValues.Add "4", "귀속연월 AND 지급연월"
    oCombo.ValidValues.Add "5", "신고연월"
    oForm.Items("WCHTYP").DisplayDesc = True

    '// 사원번호 자릿수
    Set oCombo = oForm.Items("EmpTLen").Specific
    For i = 1 To 8
        oCombo.ValidValues.Add i, i & " Length"
    Next i
    '// 결재란수(기본 6개)
    Set oCombo = oForm.Items("EmpType").Specific
    For i = 1 To 6
        oCombo.ValidValues.Add i, i
    Next i
    
    '/ Check 버튼
    Set oCheck = oForm.Items("govIDChk").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
        
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY005_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY005_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY005_EnableMenus()
    
    On Error GoTo PH_PY005_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY005_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY005_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY005_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY005_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY005_FormItemEnabled

    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY005_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY005_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY005_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY005_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY005_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then

        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
'        '// 접속자에 따른 권한별 사업장 콤보박스세팅
'        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY005_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY005_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUid = "1" Then
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    '//추가및 업데이트시에
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY005_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        
                    End If
                ElseIf pval.ItemUid = "CBtn1" Then
                    oForm.Items("TaxCode").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUid = "CBtn2" Then
                    oForm.Items("EmpID").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUid = "CBtn6" Then
                    oForm.Items("BNKCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUid = "FLD01" Then
                    oForm.PaneLevel = 1
                ElseIf pval.ItemUid = "FLD02" Then
                    oForm.PaneLevel = 2
                ElseIf pval.ItemUid = "INSBtn1" Then
                    Call Picture_Save
                End If

            ElseIf pval.BeforeAction = False Then
                If pval.ActionSuccess = True Then
                    If pval.ItemUid = "1" Then
                        If oForm.Mode = fm_ADD_MODE Then
                            If pval.ActionSuccess = True Then
                                Call PH_PY005_FormItemEnabled
        
                            End If
                        ElseIf oForm.Mode = fm_UPDATE_MODE Then
                            If pval.ActionSuccess = True Then
                                Call PH_PY005_FormItemEnabled
        
                            End If
                        ElseIf oForm.Mode = fm_OK_MODE Then
                            If pval.ActionSuccess = True Then
                                Call PH_PY005_FormItemEnabled
                            End If
                        End If
                    End If
                End If
                
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True Then
                If pval.ItemUid = "Code" And pval.CharPressed = 9 And pval.FormMode = fm_ADD_MODE Then
                    If MDC_SetMod.Value_ChkYn("[@PH_PY005A]", "Code", "'" + oForm.Items(pval.ItemUid).Specific.String + "'", "") = False Then
                        Sbo_Application.StatusBar.SetText "동일한 자사코드가 존재합니다. 확인하여 주십시오.", bmt_Short, smt_Error
                        BubbleEvent = False
                    End If
                End If
             ElseIf pval.BeforeAction = True Then
                If pval.ItemUid = "TaxCode" And pval.CharPressed = 9 Then
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001]", "U_Code", "'" + oForm.Items(pval.ItemUid).Specific.String + "'", _
                    " AND Code='P007'") = True Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
             End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUid
            Case "Mat1"
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUid
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Case Else
                oLastItemUID = pval.ItemUid
                oLastColUID = ""
                oLastColRow = 0
            End Select
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUid
                    Case "TaxDGbn"   '/ 대리인구분
                        If oForm.Items(pval.ItemUid).Specific.Selected Is Nothing Then
                           oDS_PH_PY005A.setValue "U_TaxDGbn", 0, ""
                           oDS_PH_PY005A.setValue "U_TaxDGnm", 0, ""
                        Else
                            oDS_PH_PY005A.setValue "U_TaxDGbn", 0, oForm.Items(pval.ItemUid).Specific.Selected.Value
                            oDS_PH_PY005A.setValue "U_TaxDGnm", 0, oForm.Items(pval.ItemUid).Specific.Selected.Description
                            If oDS_PH_PY005A.GetValue("U_TaxDGbn", 0) = "2" Then '법인일경우
                                oDS_PH_PY005A.setValue "U_TaxDNam", 0, oDS_PH_PY005A.GetValue("Code", 0)
                                oDS_PH_PY005A.setValue "U_TaxDBus", 0, oDS_PH_PY005A.GetValue("U_BusNum", 0)
                                oForm.Items("TaxDNam").Update
                                oForm.Items("TaxDBus").Update
                            End If
                            oForm.Update
                        End If
                    Case "AutoChk"
                        If Not oForm.Items(pval.ItemUid).Specific.Selected Is Nothing Then
                            oDS_PH_PY005A.setValue "U_AutoChk", 0, oForm.Items(pval.ItemUid).Specific.Selected.Value
                            '/ 사원번호 자릿수 자동 셋팅
                            Select Case Trim$(oDS_PH_PY005A.GetValue("U_AutoChk", 0))
                            Case "1", "2", "3"
                                oDS_PH_PY005A.setValue "U_EmpTLen", 0, 8
                            Case "4", "5", "6"
                                oDS_PH_PY005A.setValue "U_EmpTLen", 0, 6
                            Case "7"
                                oDS_PH_PY005A.setValue "U_EmpTLen", 0, 5
                            End Select
                            oForm.Items("EmpTLen").Update
                        End If
                    Case "EmpType"
                        Call Sign_Enabled
                    End Select
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True And pval.ItemUid <> "1000001" And pval.ItemUid <> "2" Then
                If oLastItemUID = "TaxCode" Then          '//세무서코드
'                    If MDC_SetMod.Value_ChkYn("[@ZPY001L]", "U_Minor", "'" + oForm.Items(oLastItemUID).Specific.String + "'", _
'                    " AND Code='P007'") = True _
'                      And oForm.Items(oLastItemUID).Specific.String <> "" And oLastItemUID <> pval.ItemUID Then
'                        Sbo_Application.ActivateMenuItem ("7425")
'                        BubbleEvent = False
'                    End If
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
            Call oForm.Freeze(True)
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                Select Case pval.ItemUid
                Case "Code"
                    If oForm.Items(pval.ItemUid).Specific.String = "" Then
                        oDS_PH_PY005A.setValue "U_CLTCode", 0, ""
                        oDS_PH_PY005A.setValue "Code", 0, ""
                    Else
                        oDS_PH_PY005A.setValue "U_CLTCode", 0, UCase$(oForm.Items(pval.ItemUid).Specific.String)
                        oDS_PH_PY005A.setValue "Code", 0, UCase$(oForm.Items(pval.ItemUid).Specific.String)
                    End If
                Case "Name"
                    If oForm.Items(pval.ItemUid).Specific.String = "" Then
                        oDS_PH_PY005A.setValue "Name", 0, ""
                        oDS_PH_PY005A.setValue "U_CLTName", 0, ""
                    Else
                        oDS_PH_PY005A.setValue "Name", 0, UCase$(oForm.Items(pval.ItemUid).Specific.String)
                        oDS_PH_PY005A.setValue "U_CLTName", 0, UCase$(oForm.Items(pval.ItemUid).Specific.String)
                    End If
                    oForm.Items("CLTName").Update
                Case "EmpID"
                    If oForm.Items(pval.ItemUid).Specific.String = "" Then
                        oDS_PH_PY005A.setValue "U_EmpID", 0, ""
                        oDS_PH_PY005A.setValue "U_ComPrt", 0, ""
                    Else
                        oDS_PH_PY005A.setValue "U_EmpID", 0, oForm.Items(pval.ItemUid).Specific.String
                        oDS_PH_PY005A.setValue "U_ComPrt", 0, MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUid).Specific.Value & "'", "")
                    End If
                    oForm.Items("ComPrt").Update
                Case "TaxCode"
                    If oForm.Items(pval.ItemUid).Specific.String = "" Then
                        oDS_PH_PY005A.setValue "U_TaxCode", 0, ""
                        oDS_PH_PY005A.setValue "U_TaxName", 0, ""
                        oDS_PH_PY005A.setValue "U_TaxAcct", 0, ""
                    Else
                        oDS_PH_PY005A.setValue "U_TaxCode", 0, oForm.Items(pval.ItemUid).Specific.String
                        oDS_PH_PY005A.setValue "U_TaxName", 0, MDC_SetMod.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm.Items(pval.ItemUid).Specific.String & "'", "And Code='P007'")
                        oDS_PH_PY005A.setValue "U_TaxAcct", 0, MDC_SetMod.Get_ReData("U_Char1", "U_Code", "[@PS_HR200L]", "'" & oForm.Items(pval.ItemUid).Specific.String & "'", "And Code='P007'")
                    End If
                    oForm.Items("TaxName").Update
                    oForm.Items("TaxAcct").Update
                End Select
            End If
                
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11

'----------------------------------------------------------
        Case et_DATASOURCE_LOAD '//12
           
'----------------------------------------------------------
        Case et_FORM_LOAD: '//16
            
'----------------------------------------------------------
        Case et_FORM_UNLOAD: '//17
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_PH_PY005A = Nothing

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

            End If
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
            If pval.BeforeAction = True Then
            
            ElseIf pval.Before_Action = False Then
'                If pval.ItemUID = "Code" Then
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY005A", "Code")
'                End If
            End If

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
                Call PH_PY005_FormItemEnabled

            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY005_FormItemEnabled
            Case "1282":    '//문서추가
                Call PH_PY005_FormItemEnabled
                oDS_PH_PY005A.setValue "U_WCHTYP", 0, "5"
                oDS_PH_PY005A.setValue "U_SAUPJA", 0, "N"
                oForm.Items("WCHTYP").Update
                oForm.Items("FLD01").CLICK ct_Regular
                
            Case "1288", "1289", "1290", "1291":
                Call PH_PY005_FormItemEnabled
            Case "1293" '// 행삭제
                
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
    Select Case pval.ItemUid
    Case "Mat1"
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUid
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Case Else
        oLastItemUID = pval.ItemUid
        oLastColUID = ""
        oLastColRow = 0
    End Select
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY005_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY005_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY005B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY005B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY005B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY005B.Offset = oRow
'            oDS_PH_PY005B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY005B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY005B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY005B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY005B.Offset = oRow - 1
'            oDS_PH_PY005B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY005B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY005B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY005B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY005B.Offset = oRow
'        oDS_PH_PY005B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY005B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY005B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY005B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
'
    Call oForm.Freeze(False)
    Exit Sub
PH_PY005_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY005_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY005_FormClear()
On Error GoTo PH_PY005_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY005'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY005_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY005_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY005_DataValidCheck() As Boolean
On Error GoTo PH_PY005_DataValidCheck_Error
    PH_PY005_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    '//----------------------------------------------------------------------------------
    '//기본사항 탭
    '//----------------------------------------------------------------------------------
    PH_PY005_DataValidCheck = False
    oForm.PaneLevel = 1
    If oForm.Mode = fm_ADD_MODE Then
        If MDC_SetMod.Value_ChkYn("[@PH_PY005A]", "Code", "'" & oForm.Items("CLTCode").Specific.String & "'") = False Then
            Call Sbo_Application.StatusBar.SetText("이미 저장되어져 있는 코드가 존재합니다", bmt_Short, smt_Error)
            oForm.Items("CLTCode").CLICK ct_Regular
            Exit Function
        End If
    End If
    If Trim$(oDS_PH_PY005A.GetValue("U_BusNum", 0)) = "" Then
        Call Sbo_Application.StatusBar.SetText("사업자번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error)
        oForm.Items("BusNum").CLICK ct_Regular
        Exit Function
    ElseIf MDC_SetMod.Value_ChkYn("[@PH_PY005A]", "U_BusNum", _
           "'" & Trim$(oDS_PH_PY005A.GetValue("U_BusNum", 0)) & "'", _
           " AND Code <> '" & oForm.Items("Code").Specific.String & "'") = False Then
        Call Sbo_Application.StatusBar.SetText("사업자번호가 중복되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error)
        oForm.Items("BusNum").CLICK ct_Regular
        Exit Function
    End If
    If Trim$(oDS_PH_PY005A.GetValue("U_BUSTYP", 0)) = "" Then
        Call Sbo_Application.StatusBar.SetText("납세자구분은 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error)
        oForm.Items("BUSTYP").CLICK ct_Regular
        Exit Function
    End If
    If Trim$(oDS_PH_PY005A.GetValue("U_SINTYP", 0)) = "" Then
        Call Sbo_Application.StatusBar.SetText("원천신고구분은 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error)
        oForm.Items("SINTYP").CLICK ct_Regular
        Exit Function
    End If
    
    
'   If oForm.Items("IMGCHK").Specific.Checked = True Then
'       If MDC_SetMod.Value_ChkYn("MDC_PAYPIC", "EmpID", Val(oForm.Items("Code").Specific.String) * -1) = True Then
'            Call Sbo_Application.StatusBar.SetText("직인이미지를 등록하세요.", bmt_Short, smt_Error)
'            Exit Function
'        End If
'    ElseIf oForm.Items("IMGCHK").Specific.Checked = False Then
'        Call Picture_Delet
'    end if
    
    If Trim$(oDS_PH_PY005A.GetValue("U_WCHCLT", 0)) = "" Then
        oDS_PH_PY005A.setValue "U_WCHCLT", 0, oForm.Items("Code").Specific.String
        oForm.Items("WCHCLT").Update
    End If
    If Trim$(oDS_PH_PY005A.GetValue("U_SUPCLT", 0)) = "" Then
        oDS_PH_PY005A.setValue "U_SUPCLT", 0, oForm.Items("Code").Specific.String
        oForm.Items("SUPCLT").Update
    End If
    If Trim$(oDS_PH_PY005A.GetValue("U_JUMCLT", 0)) = "" Then
        oDS_PH_PY005A.setValue "U_JUMCLT", 0, oForm.Items("Code").Specific.String
        oForm.Items("JUMCLT").Update
    End If
    
    oForm.PaneLevel = 2
    If Trim$(oDS_PH_PY005A.GetValue("U_AutoChk", 0)) = "" Then
        Call Sbo_Application.StatusBar.SetText("사원번호 구성 체계는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error)
        oForm.Items("AutoChk").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim$(oDS_PH_PY005A.GetValue("U_EmpTLen", 0)) = "" Then
        Call Sbo_Application.StatusBar.SetText("사원번호 자릿수는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error)
        oForm.Items("EmpTLen").CLICK ct_Regular
        Exit Function
    End If
    If Trim$(oDS_PH_PY005A.GetValue("U_EmpType", 0)) = "" Then
        Call Sbo_Application.StatusBar.SetText("결재란 수는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error)
        oForm.Items("EmpType").CLICK ct_Regular
        Exit Function
    End If
    
    
    
    '// Code & Name 생성
    Call oDS_PH_PY005A.setValue("Code", 0, Trim(oDS_PH_PY005A.GetValue("U_CLTCode", 0)))
    Call oDS_PH_PY005A.setValue("NAME", 0, Trim(oDS_PH_PY005A.GetValue("U_CLTName", 0)))
    
'    '//코드 중복 체크
'    If oForm.Mode = fm_ADD_MODE Then
'        sQry = "SELECT CODE FROM [@PH_PY005A] WHERE CODE = '" & Trim(oDS_PH_PY102A.GetValue("U_CLTCode", 0)) & "'"
'        oRecordset.DoQuery sQry
'        If oRecordset.RecordCount > 0 Then
'            Sbo_Application.SetStatusBarMessage "코드가 존재합니다", bmt_Short, True
'            PH_PY002_DataValidCheck = False
'            Exit Function
'        End If
'    End If
    
    PH_PY005_DataValidCheck = True
    Set oRecordSet = Nothing
    Exit Function
   
PH_PY005_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY005_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY005_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY005_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY005_Validate_Error
    PH_PY005_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY005A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY005_Validate = False
        GoTo PH_PY005_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY005_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY005_Validate_Error:
    PH_PY005_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY005_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub Picture_Save()
    
    '직인이미지 등록
        
    Dim L_Code As Double
    Dim SEALIMG As String
    
    L_Code = Val(oForm.Items("Code").Specific.String)
    
    If Val(L_Code) = 0 Then Exit Sub
    
    L_Code = L_Code * -1
    
    If oForm.Items("IMGCHK").Specific.Checked = False Then
        
        Call Sbo_Application.StatusBar.SetText("세무서식에 사용할 직인 등록을 체크하세요.", bmt_Short, smt_Error)
        Exit Sub
        
    End If
    
    SEALIMG = FileListBoxForm.OpenDialog(oForm, "graphic Files (*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF", "파일선택", "C:\")
    
    If Trim$(SEALIMG) = "" Then Exit Sub
        
    Dim sQry            As String

    Dim iFileLen  As Long
    Dim iDataFile As Long
    Dim i         As Integer
    Dim iFrag     As Integer
    Dim iChunks   As Integer
    Dim iChunk()  As Byte

    Const iChunkSize As Integer = 16348

    sQry = "SELECT EmpID, FILLEN, FILIMG FROM MDC_PAYPIC WHERE EmpID = " & L_Code & ""
    Set g_ADORS1 = New ADODB.Recordset
    g_ADORS1.Open sQry, g_ERPDMS, adOpenDynamic, adLockOptimistic

    iDataFile = FreeFile
    
    Open SEALIMG For Binary Access Read As iDataFile
    
    iFileLen = LOF(iDataFile)
    
    If iFileLen > 0 Then
        
        If g_ADORS1.EOF Then g_ADORS1.AddNew
        g_ADORS1.Fields("EmpID").Value = L_Code
        g_ADORS1.Fields("FILLEN").Value = iFileLen    '/ 길이
        iChunks = iFileLen \ iChunkSize
        iFrag = iFileLen Mod iChunkSize
        g_ADORS1.Fields("FILIMG").AppendChunk Null
        ReDim iChunk(iFrag)
        Get iDataFile, , iChunk()
        g_ADORS1.Fields("FILIMG").AppendChunk iChunk()
        ReDim iChunk(iChunkSize)
        For i = 1 To iChunks
            Get iDataFile, , iChunk()
            g_ADORS1.Fields("FILIMG").AppendChunk iChunk()
        Next i
        
    End If
    
    Close iDataFile
    
    g_ADORS1.Update
    
    Set g_ADORS1 = Nothing
    
    Exit Sub
'/ Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
Error_Message:
    Set g_ADORS1 = Nothing
    Sbo_Application.StatusBar.SetText "Picture_Save Error:" & Err.Description, bmt_Short, smt_Error
    
End Sub


Private Sub Picture_Delete()
    
    '직인이미지 삭제
    
    Dim sQry            As String
    Dim L_Code          As Double
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    L_Code = Val(oForm.Items("Code").Specific.String)
    
    If Val(L_Code) = 0 Then Exit Sub
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    L_Code = L_Code * -1
    
    sQry = "DELETE FROM [MDC_PAYPIC] WHERE EmpID = " & L_Code & ""
    oRecordSet.DoQuery sQry
    
    Set oRecordSet = Nothing
    
End Sub

Private Sub Sign_Enabled()
    Dim i       As Integer
    Dim MAXCNT  As Integer
    MAXCNT = Val(oDS_PH_PY005A.GetValue("U_EmpType", 0))
    For i = 1 To 8
        If MAXCNT >= i Then
            oForm.Items("Sign0" & i).Enabled = True
        Else
            oDS_PH_PY005A.setValue "U_Sign0" & i, 0, ""
            oForm.Items("Sign0" & i).Enabled = False
        End If
    Next i
    
End Sub

