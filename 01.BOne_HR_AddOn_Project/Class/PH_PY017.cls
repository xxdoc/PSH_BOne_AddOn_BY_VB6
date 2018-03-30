VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY017.cls
'//  Module         : 근태관리 > 월근태집계
'//  Desc           : 월근태집계처리
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY017A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY017B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY017.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY017_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY017"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY017_CreateItems
    Call PH_PY017_EnableMenus
    Call PH_PY017_SetDocument(oFromDocEntry01)
'    Call PH_PY017_FormResize
    
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

Private Function PH_PY017_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim CLTCOD      As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY017_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY017A = oForm.DataSources.DBDataSources("@PH_PY017A")
    Set oDS_PH_PY017B = oForm.DataSources.DBDataSources("@PH_PY017B")

    
    Set oMat1 = oForm.Items("Mat01").Specific       '//@PH_PY017B
    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    
    
    '// 사업장
    
    Set oCombo = oForm.Items("CLTCOD").Specific
'    oCombo.DataBind.SetBound True, "", "CLTCOD"
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    
    
    '// 접속자에 따른 사업장 선택
    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
    oDS_PH_PY017A.setValue "U_CLTCOD", 0, CLTCOD
    
    oDS_PH_PY017A.setValue "U_YM", 0, Format(Now, "YYYYMM")
    
    'oForm.DataSources.UserDataSources.Item("CLTCOD").Value =
    
    
    '// 년월
    'oForm.Items("YM").Specific.Value =
    
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY017_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY017_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY017_EnableMenus()
    
    On Error GoTo PH_PY017_EnableMenus_Error
    Call oForm.EnableMenu("1281", True)     '//찾기
    Call oForm.EnableMenu("1282", True)     '//신규
    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY017_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY017_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY017_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY017_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY017_FormItemEnabled
        Call PH_PY017_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY017_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY017_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY017_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY017_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim CLTCOD  As String
    
    On Error GoTo PH_PY017_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
'        '// 접속자에 따른 사업장 선택
'        CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
'        oDS_PH_PY017A.setValue "U_CLTCOD", 0, CLTCOD
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        '//년월
        oDS_PH_PY017A.setValue "U_YM", 0, Format(Now, "YYYYMM")

        oMat1.Clear
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY017_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY017_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                        If PH_PY017_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    End If
                        '//해야할일 작업
                End If
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY017_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY017_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY017_FormItemEnabled
                        End If
                    End If
                End If
                If pval.ItemUID = "Btn_CREATE" Then
                    Call PH_PY017_ITEM_CREATE
                End If
                
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat01"
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
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                  
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat01"
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)
                    End If
                End Select

                Select Case pval.ItemUID
                Case "Mat01"
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
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
       
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
'                    If pval.ItemUID = "Mat01" And pval.ColUID = "" Then
'                        Call PH_PY017_AddMatrixRow
'                        Call oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
'                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY017_FormItemEnabled
                Call PH_PY017_AddMatrixRow
                
            End If
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
                Set oDS_PH_PY017A = Nothing
                Set oDS_PH_PY017B = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY017A", "Code")
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
                Call PH_PY017_FormItemEnabled
                Call PH_PY017_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY017_FormItemEnabled
                Call PH_PY017_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY017_FormItemEnabled
                Call PH_PY017_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY017_FormItemEnabled
            Case "1293" '// 행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat1, oDS_PH_PY017B, "U_CODNBR")
                Call PH_PY017_AddMatrixRow
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Case "Mat01"
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

Sub PH_PY017_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY017_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY017B.GetValue("U_MSTCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY017B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY017B.InsertRecord (oRow)
            End If
            oDS_PH_PY017B.Offset = oRow
            oDS_PH_PY017B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY017B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY017B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY017B.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY017B.setValue "U_RspCode", oRow, ""
            oDS_PH_PY017B.setValue "U_StdGDay", oRow, 0
            oDS_PH_PY017B.setValue "U_StdPDay", oRow, 0
            oDS_PH_PY017B.setValue "U_StdNDay", oRow, 0
            oDS_PH_PY017B.setValue "U_GetDay", oRow, 0
            oDS_PH_PY017B.setValue "U_WoHDay", oRow, 0
            oDS_PH_PY017B.setValue "U_PayDay", oRow, 0
            oDS_PH_PY017B.setValue "U_AbsDay", oRow, 0
            oDS_PH_PY017B.setValue "U_Base", oRow, 0
            oDS_PH_PY017B.setValue "U_Extend", oRow, 0
            oDS_PH_PY017B.setValue "U_Midnight", oRow, 0
            oDS_PH_PY017B.setValue "U_EarlyTo", oRow, 0
            oDS_PH_PY017B.setValue "U_Special", oRow, 0
            oDS_PH_PY017B.setValue "U_SpExtend", oRow, 0
            oDS_PH_PY017B.setValue "U_SMidnigh", oRow, 0
            oDS_PH_PY017B.setValue "U_SEarlyTo", oRow, 0
            oDS_PH_PY017B.setValue "U_EduTime", oRow, 0
            oDS_PH_PY017B.setValue "U_LateToC", oRow, 0
            oDS_PH_PY017B.setValue "U_EarlyOfC", oRow, 0
            oDS_PH_PY017B.setValue "U_GoOutC", oRow, 0
            oDS_PH_PY017B.setValue "U_LateToT", oRow, 0
            oDS_PH_PY017B.setValue "U_EarlyOfT", oRow, 0
            oDS_PH_PY017B.setValue "U_GoOutT", oRow, 0
            oDS_PH_PY017B.setValue "U_JCHDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_YCHDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_YCHHGA", oRow, 0
            oDS_PH_PY017B.setValue "U_SNHDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_SNHHGA", oRow, 0
            oDS_PH_PY017B.setValue "U_DNGDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_WHMDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY1", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY2", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY3", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY4", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY5", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY6", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY7", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY8", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY9", oRow, 0
            
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY017B.Offset = oRow - 1
            oDS_PH_PY017B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY017B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY017B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY017B.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY017B.setValue "U_RspCode", oRow, ""
            oDS_PH_PY017B.setValue "U_StdGDay", oRow, 0
            oDS_PH_PY017B.setValue "U_StdPDay", oRow, 0
            oDS_PH_PY017B.setValue "U_StdNDay", oRow, 0
            oDS_PH_PY017B.setValue "U_GetDay", oRow, 0
            oDS_PH_PY017B.setValue "U_WoHDay", oRow, 0
            oDS_PH_PY017B.setValue "U_PayDay", oRow, 0
            oDS_PH_PY017B.setValue "U_AbsDay", oRow, 0
            oDS_PH_PY017B.setValue "U_Base", oRow, 0
            oDS_PH_PY017B.setValue "U_Extend", oRow, 0
            oDS_PH_PY017B.setValue "U_Midnight", oRow, 0
            oDS_PH_PY017B.setValue "U_EarlyTo", oRow, 0
            oDS_PH_PY017B.setValue "U_Special", oRow, 0
            oDS_PH_PY017B.setValue "U_SpExtend", oRow, 0
            oDS_PH_PY017B.setValue "U_SMidnigh", oRow, 0
            oDS_PH_PY017B.setValue "U_SEarlyTo", oRow, 0
            oDS_PH_PY017B.setValue "U_EduTime", oRow, 0
            oDS_PH_PY017B.setValue "U_LateToC", oRow, 0
            oDS_PH_PY017B.setValue "U_EarlyOfC", oRow, 0
            oDS_PH_PY017B.setValue "U_GoOutC", oRow, 0
            oDS_PH_PY017B.setValue "U_LateToT", oRow, 0
            oDS_PH_PY017B.setValue "U_EarlyOfT", oRow, 0
            oDS_PH_PY017B.setValue "U_GoOutT", oRow, 0
            oDS_PH_PY017B.setValue "U_JCHDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_YCHDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_YCHHGA", oRow, 0
            oDS_PH_PY017B.setValue "U_SNHDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_SNHHGA", oRow, 0
            oDS_PH_PY017B.setValue "U_DNGDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_WHMDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY1", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY2", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY3", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY4", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY5", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY6", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY7", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY8", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY9", oRow, 0
            
            oMat1.LoadFromDataSource
            
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
            oDS_PH_PY017B.Offset = oRow
            oDS_PH_PY017B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY017B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY017B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY017B.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY017B.setValue "U_RspCode", oRow, ""
            oDS_PH_PY017B.setValue "U_StdGDay", oRow, 0
            oDS_PH_PY017B.setValue "U_StdPDay", oRow, 0
            oDS_PH_PY017B.setValue "U_StdNDay", oRow, 0
            oDS_PH_PY017B.setValue "U_GetDay", oRow, 0
            oDS_PH_PY017B.setValue "U_WoHDay", oRow, 0
            oDS_PH_PY017B.setValue "U_PayDay", oRow, 0
            oDS_PH_PY017B.setValue "U_AbsDay", oRow, 0
            oDS_PH_PY017B.setValue "U_Base", oRow, 0
            oDS_PH_PY017B.setValue "U_Extend", oRow, 0
            oDS_PH_PY017B.setValue "U_Midnight", oRow, 0
            oDS_PH_PY017B.setValue "U_EarlyTo", oRow, 0
            oDS_PH_PY017B.setValue "U_Special", oRow, 0
            oDS_PH_PY017B.setValue "U_SpExtend", oRow, 0
            oDS_PH_PY017B.setValue "U_SMidnigh", oRow, 0
            oDS_PH_PY017B.setValue "U_SEarlyTo", oRow, 0
            oDS_PH_PY017B.setValue "U_EduTime", oRow, 0
            oDS_PH_PY017B.setValue "U_LateToC", oRow, 0
            oDS_PH_PY017B.setValue "U_EarlyOfC", oRow, 0
            oDS_PH_PY017B.setValue "U_GoOutC", oRow, 0
            oDS_PH_PY017B.setValue "U_LateToT", oRow, 0
            oDS_PH_PY017B.setValue "U_EarlyOfT", oRow, 0
            oDS_PH_PY017B.setValue "U_GoOutT", oRow, 0
            oDS_PH_PY017B.setValue "U_JCHDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_YCHDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_YCHHGA", oRow, 0
            oDS_PH_PY017B.setValue "U_SNHDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_SNHHGA", oRow, 0
            oDS_PH_PY017B.setValue "U_DNGDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_WHMDAY", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY1", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY2", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY3", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY4", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY5", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY6", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY7", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY8", oRow, 0
            oDS_PH_PY017B.setValue "U_EtcDAY9", oRow, 0
            
        oMat1.LoadFromDataSource
        
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY017_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY017_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY017_FormClear()
On Error GoTo PH_PY017_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY017'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY017_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY017_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY017_DataValidCheck() As Boolean
On Error GoTo PH_PY017_DataValidCheck_Error
    PH_PY017_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim tCode   As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
   
   
    '// 라인 ---------------------------
    If oMat1.VisualRowCount > 0 Then
      
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY017_DataValidCheck = False
        Exit Function
    End If
    
    oMat1.FlushToDataSource
    
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY017B.Size > 1 Then oDS_PH_PY017B.RemoveRecord (oDS_PH_PY017B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    '//HEAD TABLE에 키 SET
    tCode = Trim(oForm.Items("CLTCOD").Specific.VALUE) + oForm.Items("YM").Specific.VALUE
    oDS_PH_PY017A.setValue "Code", 0, tCode
    oDS_PH_PY017A.setValue "Name", 0, tCode
    
    PH_PY017_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY017_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY017_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY017_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Function PH_PY017_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY017_Validate_Error
    PH_PY017_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY017A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY017_Validate = False
        GoTo PH_PY017_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY017_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY017_Validate_Error:
    PH_PY017_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY017_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'//행삭제 (FormUID, pval, BubbleEvent, 매트릭스 이름, 디비데이터소스, 데이터 체크 필드명)
Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean, _
                                                    oMat As SAPbouiCOM.Matrix, DBData As SAPbouiCOM.DBDataSource, CheckField As String)

    Dim i As Long
    
    On Error GoTo Raise_EVENT_ROW_DELETE_Error

    If (oLastColRow > 0) Then
        If pval.BeforeAction = True Then
        
        ElseIf pval.BeforeAction = False Then
            If oMat.RowCount <> oMat.VisualRowCount Then
                oMat.FlushToDataSource
    
                While (i <= DBData.Size - 1)
                    If DBData.GetValue(CheckField, i) = "" Then
                        DBData.RemoveRecord (i)
                        i = 0
                    Else
                        i = i + 1
                    End If
                Wend
                
                For i = 0 To DBData.Size
                    Call DBData.setValue("U_LineNum", i, i + 1)
                Next i
                
                oMat.LoadFromDataSource
            End If
        End If
    End If
Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY017_ITEM_CREATE()
    
    Dim i           As Long
    Dim j           As Long
    Dim sPrice      As String
    Dim sFile       As String
    Dim OneRec      As String
    Dim sQry        As String
    Dim CLTCOD      As String
    Dim YM          As String
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oForm       As SAPbouiCOM.Form
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Err
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Set oForm = Sbo_Application.Forms.ActiveForm
    
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    YM = oForm.Items("YM").Specific.VALUE
    
    
    Call oForm.Freeze(True)
    
    sQry = "EXEC [PH_PY017_01] '" & CLTCOD & "', '" & YM & "'"
    oRecordSet.DoQuery sQry
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("자료집계중!", oRecordSet.RecordCount, False)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
        
    If oRecordSet.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet = Nothing
        Exit Sub
    End If
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i + 1 > oDS_PH_PY017B.Size Then
            oDS_PH_PY017B.InsertRecord (i)
        End If
        
        oMat1.AddRow
        oDS_PH_PY017B.Offset = i
        oDS_PH_PY017B.setValue "U_LineNum", i, i + 1
        
        oDS_PH_PY017B.setValue "U_MSTCOD", i, Trim(oRecordSet.Fields("MSTCOD").VALUE)
        oDS_PH_PY017B.setValue "U_MSTNAM", i, Trim(oRecordSet.Fields("FullName").VALUE)
        oDS_PH_PY017B.setValue "U_TeamCode", i, Trim(oRecordSet.Fields("TeamCode").VALUE)
        oDS_PH_PY017B.setValue "U_RspCode", i, Trim(oRecordSet.Fields("RspCode").VALUE)
        oDS_PH_PY017B.setValue "U_ClsCode", i, Trim(oRecordSet.Fields("ClsCode").VALUE)
        oDS_PH_PY017B.setValue "U_StdGDay", i, Trim(oRecordSet.Fields("StdGDay").VALUE)
        oDS_PH_PY017B.setValue "U_StdPDay", i, Trim(oRecordSet.Fields("StdPDay").VALUE)
        oDS_PH_PY017B.setValue "U_StdNDay", i, Trim(oRecordSet.Fields("StdNDay").VALUE)
        oDS_PH_PY017B.setValue "U_GetDay", i, Trim(oRecordSet.Fields("GetDay").VALUE)
        oDS_PH_PY017B.setValue "U_WoHDay", i, Trim(oRecordSet.Fields("WoHDay").VALUE)
        oDS_PH_PY017B.setValue "U_PayDay", i, Trim(oRecordSet.Fields("PayDay").VALUE)
        oDS_PH_PY017B.setValue "U_AbsDay", i, Trim(oRecordSet.Fields("AbsDay").VALUE)
        oDS_PH_PY017B.setValue "U_Base", i, Trim(oRecordSet.Fields("Base").VALUE)
        oDS_PH_PY017B.setValue "U_Extend", i, Trim(oRecordSet.Fields("Extend").VALUE)
        oDS_PH_PY017B.setValue "U_Midnight", i, Trim(oRecordSet.Fields("Midnight").VALUE)
        oDS_PH_PY017B.setValue "U_EarlyTo", i, Trim(oRecordSet.Fields("EarlyTo").VALUE)
        oDS_PH_PY017B.setValue "U_Special", i, Trim(oRecordSet.Fields("Special").VALUE)
        oDS_PH_PY017B.setValue "U_SpExtend", i, Trim(oRecordSet.Fields("SpExtend").VALUE)
        oDS_PH_PY017B.setValue "U_SMidnigh", i, Trim(oRecordSet.Fields("SMidnigh").VALUE)
        oDS_PH_PY017B.setValue "U_SEarlyTo", i, Trim(oRecordSet.Fields("SEarlyTo").VALUE)
        oDS_PH_PY017B.setValue "U_EduTime", i, Trim(oRecordSet.Fields("EduTime").VALUE)
        oDS_PH_PY017B.setValue "U_LateToC", i, Trim(oRecordSet.Fields("LateToC").VALUE)
        oDS_PH_PY017B.setValue "U_EarlyOfC", i, Trim(oRecordSet.Fields("EarlyOfC").VALUE)
        oDS_PH_PY017B.setValue "U_GoOutC", i, Trim(oRecordSet.Fields("GoOutC").VALUE)
        oDS_PH_PY017B.setValue "U_LateToT", i, Trim(oRecordSet.Fields("LateToT").VALUE)
        oDS_PH_PY017B.setValue "U_EarlyOfT", i, Trim(oRecordSet.Fields("EarlyOfT").VALUE)
        oDS_PH_PY017B.setValue "U_GoOutT", i, Trim(oRecordSet.Fields("GoOutT").VALUE)
        oDS_PH_PY017B.setValue "U_JCHDAY", i, Trim(oRecordSet.Fields("JCHDAY").VALUE)
        oDS_PH_PY017B.setValue "U_YCHDAY", i, Trim(oRecordSet.Fields("YCHDAY").VALUE)
        oDS_PH_PY017B.setValue "U_YCHHGA", i, Trim(oRecordSet.Fields("YCHHGA").VALUE)
        oDS_PH_PY017B.setValue "U_SNHDAY", i, Trim(oRecordSet.Fields("SNHDAY").VALUE)
        oDS_PH_PY017B.setValue "U_SNHHGA", i, Trim(oRecordSet.Fields("SNHHGA").VALUE)
        oDS_PH_PY017B.setValue "U_DNGDAY", i, Trim(oRecordSet.Fields("DNGDAY").VALUE)
        oDS_PH_PY017B.setValue "U_WHMDAY", i, Trim(oRecordSet.Fields("WHMDAY").VALUE)
        oDS_PH_PY017B.setValue "U_EtcDAY1", i, Trim(oRecordSet.Fields("EtcDAY1").VALUE)
        oDS_PH_PY017B.setValue "U_EtcDAY2", i, Trim(oRecordSet.Fields("EtcDAY2").VALUE)
        oDS_PH_PY017B.setValue "U_EtcDAY3", i, Trim(oRecordSet.Fields("EtcDAY3").VALUE)
        oDS_PH_PY017B.setValue "U_EtcDAY4", i, Trim(oRecordSet.Fields("EtcDAY4").VALUE)
        oDS_PH_PY017B.setValue "U_EtcDAY5", i, Trim(oRecordSet.Fields("EtcDAY5").VALUE)
        oDS_PH_PY017B.setValue "U_EtcDAY6", i, Trim(oRecordSet.Fields("EtcDAY6").VALUE)
        oDS_PH_PY017B.setValue "U_EtcDAY7", i, Trim(oRecordSet.Fields("EtcDAY7").VALUE)
        oDS_PH_PY017B.setValue "U_EtcDAY8", i, Trim(oRecordSet.Fields("EtcDAY8").VALUE)
        oDS_PH_PY017B.setValue "U_EtcDAY9", i, Trim(oRecordSet.Fields("EtcDAY9").VALUE)
            
        oRecordSet.MoveNext
        
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 조회중...!"
    Next
    
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update

    Call PH_PY017_AddMatrixRow
    
    Call oForm.Freeze(False)
    
    Exit Sub

Err:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
End Sub





