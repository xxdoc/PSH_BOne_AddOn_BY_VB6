VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY009.cls
'//  Module         : 근태관리 > 근태자료UPLOAD
'//  Desc           : 근태자료등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY009A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY009B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY009.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY009_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY009"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY009_CreateItems
    Call PH_PY009_EnableMenus
    Call PH_PY009_SetDocument(oFromDocEntry01)
'    Call PH_PY009_FormResize
    
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

Private Function PH_PY009_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    Dim CLTCOD      As String
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY009_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    Set oDS_PH_PY009A = oForm.DataSources.DBDataSources("@PS_USERDS02")
    Set oDS_PH_PY009B = oForm.DataSources.DBDataSources("@PS_USERDS01")

    
    Set oMat1 = oForm.Items("Mat01").Specific       '//@PH_PY009B
    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    
    
    '// 사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'
'    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
'    oCombo.Select CLTCOD, psk_ByValue
'    oCombo.Select 0, psk_Index
    oForm.Items("CLTCOD").DisplayDesc = True
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY009_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY009_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY009_EnableMenus()
    
    On Error GoTo PH_PY009_EnableMenus_Error
    Call oForm.EnableMenu("1281", False)     '//찾기
    Call oForm.EnableMenu("1282", True)     '//신규
    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY009_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY009_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY009_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY009_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY009_FormItemEnabled
        Call PH_PY009_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY009_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY009_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY009_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY009_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY009_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        oForm.Items("Comments").Specific.VALUE = ""
        oMat1.Clear
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", False)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY009_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY009_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                        If PH_PY009_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    End If
                        '//해야할일 작업
                End If
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY009_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY009_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY009_FormItemEnabled
                        End If
                    End If
                End If
                If pval.ItemUID = "Btn_UPLOAD" Then
                    Call PH_PY009_Excel_Upload
                End If
                If pval.ItemUID = "Btn_Apply" Then
                    Call PH_PY009_DataApply
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
'                        Call PH_PY009_AddMatrixRow
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

                Call PH_PY009_FormItemEnabled
                Call PH_PY009_AddMatrixRow
                
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
                Set oDS_PH_PY009A = Nothing
                Set oDS_PH_PY009B = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY009A", "Code")
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
                Call PH_PY009_FormItemEnabled
                Call PH_PY009_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY009_FormItemEnabled
                Call PH_PY009_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY009_FormItemEnabled
                Call PH_PY009_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY009_FormItemEnabled
            Case "1293" '// 행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat1, oDS_PH_PY009B, "U_CODNBR")
                Call PH_PY009_AddMatrixRow
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

Sub PH_PY009_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY009_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY009B.GetValue("U_ColReg01", oRow - 1)) <> "" Then
            If oDS_PH_PY009B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY009B.InsertRecord (oRow)
            End If
            oDS_PH_PY009B.Offset = oRow
            oDS_PH_PY009B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY009B.setValue "U_ColReg01", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg02", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg03", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg04", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg05", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg06", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg07", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg08", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg09", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg10", oRow, ""
            oDS_PH_PY009B.setValue "U_ColReg11", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY009B.Offset = oRow - 1
            oDS_PH_PY009B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY009B.setValue "U_ColReg01", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg02", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg03", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg04", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg05", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg06", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg07", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg08", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg09", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg10", oRow - 1, ""
            oDS_PH_PY009B.setValue "U_ColReg11", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY009B.Offset = oRow
        oDS_PH_PY009B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY009B.setValue "U_ColReg01", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg02", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg03", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg04", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg05", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg06", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg07", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg08", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg09", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg10", oRow, ""
        oDS_PH_PY009B.setValue "U_ColReg11", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY009_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY009_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY009_FormClear()
On Error GoTo PH_PY009_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY009'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY009_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY009_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY009_DataValidCheck() As Boolean
On Error GoTo PH_PY009_DataValidCheck_Error
    PH_PY009_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
   
   
    '// 라인 ---------------------------
    If oMat1.VisualRowCount > 1 Then
        
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY009_DataValidCheck = False
        Exit Function
    End If
    
    oMat1.FlushToDataSource
    
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY009B.Size > 1 Then oDS_PH_PY009B.RemoveRecord (oDS_PH_PY009B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    
    PH_PY009_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY009_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY009_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY009_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Function PH_PY009_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY009_Validate_Error
    PH_PY009_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY009A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY009_Validate = False
        GoTo PH_PY009_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY009_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY009_Validate_Error:
    PH_PY009_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY009_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Private Sub PH_PY009_Excel_Upload()
    
    Dim i           As Long
    Dim j           As Long
    Dim sPrice      As String
    Dim sFile       As String
    Dim OneRec      As String
    Dim sQry        As String
    Dim CLTCOD      As String
    
    
    Dim xl          As Excel.Application
    Dim xlwb        As Excel.Workbook
    Dim xlsh        As Excel.Worksheet
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oForm       As SAPbouiCOM.Form
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Err
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Set oForm = Sbo_Application.Forms.ActiveForm
    
    
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    
    
    sFile = FileListBoxForm.OpenDialog(FileListBoxForm, "*.xls|*.xls|*.xlsx|*.xlsx|", "파일선택", "C:\")
    
    If sFile = "" Then
        Sbo_Application.StatusBar.SetText "파일을 선택해 주세요.", bmt_Short, smt_Error
        Exit Sub
    Else
        If Mid(Right(sFile, 4), 1, 3) = "xls" Or Mid(Right(sFile, 5), 1, 4) = "xlsx" Then
            oForm.Items("Comments").Specific.VALUE = sFile

        Else
            Sbo_Application.StatusBar.SetText "엑셀파일이 아닙니다.", bmt_Short, smt_Error
            Exit Sub
        End If
    End If
    
    '엑셀 Object 연결
    Set xl = CreateObject("excel.application")
    Set xlwb = xl.Workbooks.Open(sFile, , True)
    Set xlsh = xlwb.Worksheets("Sheet1")
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    For i = 2 To xlsh.UsedRange.Rows.Count
        If i - 2 <> 0 Then
            oDS_PH_PY009B.InsertRecord (i - 2)
        End If
        oDS_PH_PY009B.Offset = i - 2
        Call oDS_PH_PY009B.setValue("U_LineNum", i - 2, i - 1)
        Call oDS_PH_PY009B.setValue("U_ColReg01", i - 2, CLTCOD)
        Call oDS_PH_PY009B.setValue("U_ColReg02", i - 2, xlsh.Cells(i, 1))
        Call oDS_PH_PY009B.setValue("U_ColReg03", i - 2, xlsh.Cells(i, 2))
        Call oDS_PH_PY009B.setValue("U_ColReg04", i - 2, xlsh.Cells(i, 3))
        Call oDS_PH_PY009B.setValue("U_ColReg05", i - 2, xlsh.Cells(i, 4))
        Call oDS_PH_PY009B.setValue("U_ColReg06", i - 2, xlsh.Cells(i, 5))
        Call oDS_PH_PY009B.setValue("U_ColReg07", i - 2, xlsh.Cells(i, 6))
        Call oDS_PH_PY009B.setValue("U_ColReg08", i - 2, xlsh.Cells(i, 7))
        Call oDS_PH_PY009B.setValue("U_ColReg09", i - 2, xlsh.Cells(i, 8))
        Call oDS_PH_PY009B.setValue("U_ColReg10", i - 2, xlsh.Cells(i, 9))
        Call oDS_PH_PY009B.setValue("U_ColReg11", i - 2, xlsh.Cells(i, 10))
        
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & xlsh.UsedRange.Rows.Count - 1 & "건 조회중...!"
    Next i
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update

    Call PH_PY009_AddMatrixRow
    
    Call oForm.Freeze(False)
    Sbo_Application.StatusBar.SetText "엑셀을 불러왔습니다.", bmt_Short, smt_Warning


    '액셀개체 닫음
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    '진행바 초기화
    Exit Sub

Err:
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
End Sub


Private Function PH_PY009_DataApply() As Boolean
    Dim i           As Long
    Dim j           As Long
    Dim LineId      As Long
    Dim ShiftDat    As String
    Dim sQry        As String
    Dim sQry2        As String
    
    Dim CLTCOD      As String
    Dim ymd         As String
    Dim tms         As String
    Dim tmrid       As String
    Dim usrid       As String
    Dim kname       As String
    Dim sabun       As String
    Dim Gubun       As String
    Dim pmode       As String
    Dim appr        As String
    Dim rst         As String
    
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    On Error GoTo PH_PY009_DataApply_Error
    
    PH_PY009_DataApply = False
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            CLTCOD = oMat1.Columns("CLTCOD").Cells(i).Specific.VALUE
            ymd = oMat1.Columns("ymd").Cells(i).Specific.VALUE
            tms = oMat1.Columns("tms").Cells(i).Specific.VALUE
            tmrid = oMat1.Columns("tmrid").Cells(i).Specific.VALUE
            usrid = oMat1.Columns("usrid").Cells(i).Specific.VALUE
            kname = oMat1.Columns("kname").Cells(i).Specific.VALUE
            sabun = oMat1.Columns("sabun").Cells(i).Specific.VALUE
            Gubun = oMat1.Columns("gubun").Cells(i).Specific.VALUE
            pmode = oMat1.Columns("pmode").Cells(i).Specific.VALUE
            appr = oMat1.Columns("appr").Cells(i).Specific.VALUE
            rst = oMat1.Columns("rst").Cells(i).Specific.VALUE
            
            ymd = Mid(ymd, 1, 4) + Mid(ymd, 6, 2) + Mid(ymd, 9, 2)
            tms = Mid(tms, 1, 2) + Mid(tms, 4, 2) + Mid(tms, 7, 2)
            
            sQry = "EXEC PH_PY009 '" & CLTCOD & "','" & ymd & "','" & tms & "','" & tmrid & "','" & usrid & "','"
            sQry = sQry & kname & "','" & sabun & "','" & Gubun & "','" & pmode & "','" & appr & "','" & rst & "'"
            
            
'            sQry = "INSERT INTO ZPH_PY009"
'            sQry = sQry & " ("
'            sQry = sQry & " CLTCOD,"
'            sQry = sQry & " ymd,"
'            sQry = sQry & " tms,"
'            sQry = sQry & " tmrid,"
'            sQry = sQry & " usrid,"
'            sQry = sQry & " kname,"
'            sQry = sQry & " sabun,"
'            sQry = sQry & " gubun,"
'            sQry = sQry & " pmode,"
'            sQry = sQry & " appr,"
'            sQry = sQry & " rst"
'            sQry = sQry & " ) "
'            sQry = sQry & "VALUES("
'            sQry = sQry & "'" & CLTCOD & "',"
'            sQry = sQry & "'" & ymd & "',"
'            sQry = sQry & "'" & tms & "',"
'            sQry = sQry & "'" & tmrid & "',"
'            sQry = sQry & "'" & usrid & "',"
'            sQry = sQry & "'" & kname & "',"
'            sQry = sQry & "'" & sabun & "',"
'            sQry = sQry & "'" & gubun & "',"
'            sQry = sQry & "'" & pmode & "',"
'            sQry = sQry & "'" & appr & "',"
'            sQry = sQry & "'" & rst & "'"
'            sQry = sQry & ")"
            oRecordSet.DoQuery sQry
            ProgressBar01.VALUE = ProgressBar01.VALUE + 1
            ProgressBar01.Text = ProgressBar01.VALUE & "/" & oMat1.VisualRowCount - 1 & "건 저장중...!"
            'Sbo_Application.SetStatusBarMessage "근태기찰 데이터를 입력중입니다..", bmt_Short, False
        Next
        PH_PY009_DataApply = True
        Sbo_Application.SetStatusBarMessage "근태기찰 데이터가 입력 되었습니다.", bmt_Short, False
    End If
    
    Set oRecordSet = Nothing
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Exit Function
    
PH_PY009_DataApply_Error:
    Set oRecordSet = Nothing
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY009_DataApply_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



