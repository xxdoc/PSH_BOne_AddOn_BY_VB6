VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_SY001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_SY001.cls
'//  Module         : 운영관리 > 시스템공통코드등록
'//  Desc           : 시스템공통코드등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form
Public oMat1 As SAPbouiCOM.Matrix
Private oDS_PH_SY001H As SAPbouiCOM.DBDataSource
Private oDS_PH_SY001L As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_SY001.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_SY001_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_SY001"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_SY001_CreateItems
    Call PH_SY001_EnableMenus
    Call PH_SY001_SetDocument(oFromDocEntry01)
    Call PH_SY001_FormResize
    
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

Private Function PH_SY001_CreateItems() As Boolean

    Dim sQry    As String
    Dim oEdit   As SAPbouiCOM.EditText
    
    Dim oRecordset As SAPbobsCOM.Recordset
    
    On Error GoTo PH_SY001_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_SY001H = oForm.DataSources.DBDataSources("@PH_SY001H")
    Set oDS_PH_SY001L = oForm.DataSources.DBDataSources("@PH_SY001L")
    Set oMat1 = oForm.Items("Mat01").Specific
    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    
    PH_SY001_ChooseFromList
    
    Set oEdit = oForm.Items("Name").Specific
    oEdit.ChooseFromListUID = "CFL_FrItemCode"
    oEdit.ChooseFromListAlias = "ItemCode"
    
    Set oRecordset = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_SY001_CreateItems_Error:
    Set oRecordset = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_SY001_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_SY001_ChooseFromList()

    Dim oChooseFromListCollection       As SAPbouiCOM.ChooseFromListCollection
    Dim oConditions                     As SAPbouiCOM.Conditions
    Dim oCondition                      As SAPbouiCOM.Condition
    Dim oChooseFromList                 As SAPbouiCOM.ChooseFromList
    Dim oChooseFromListCreationParams   As SAPbouiCOM.ChooseFromListCreationParams
    
    On Error GoTo PH_SY001_ChooseFromList_Error
'
    
    '//품목코드 From
    Set oChooseFromListCollection = oForm.ChooseFromLists
    Set oChooseFromListCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oChooseFromListCreationParams.ObjectType = lf_Items
    oChooseFromListCreationParams.uniqueID = "CFL_FrItemCode"
    oChooseFromListCreationParams.MultiSelection = True

    Set oChooseFromList = oChooseFromListCollection.Add(oChooseFromListCreationParams)

    Set oConditions = oChooseFromList.GetConditions()
    Set oCondition = oConditions.Add()
    
    Exit Sub
    
PH_SY001_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_SY001_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_SY001_EnableMenus()
    
    On Error GoTo PH_SY001_EnableMenus_Error
    
    
    Call oForm.EnableMenu("1283", True)
    Call oForm.EnableMenu("1293", True)
    Exit Sub
    
PH_SY001_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_SY001_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_SY001_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_SY001_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_SY001_FormItemEnabled
        Call PH_SY001_AddMatrixRow(0, True)
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_SY001_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_SY001_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_SY001_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_SY001_FormItemEnabled()
    
    On Error GoTo PH_SY001_FormItemEnabled_Error
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        
        Call oForm.EnableMenu("1281", True)
        Call oForm.EnableMenu("1282", False)
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        
        Call oForm.EnableMenu("1281", False)
        Call oForm.EnableMenu("1282", True)
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        
        Call oForm.EnableMenu("1281", True)
        Call oForm.EnableMenu("1282", True)
    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_SY001_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_SY001_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_SY001_FormResize()
On Error GoTo PH_SY001_FormResize_Error
    
    Exit Sub
PH_SY001_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_SY001_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            If pval.BeforeAction = True Then
                If pval.ItemUID = "PH_SY001" Then
                    If oForm.Mode = fm_ADD_MODE Then
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                End If
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_SY001_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_SY001_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "PH_SY001" Then
                    If oForm.Mode = fm_ADD_MODE Then
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                End If
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_SY001_FormItemEnabled
                            Call PH_SY001_AddMatrixRow(0, True)
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_SY001_FormItemEnabled
                        End If
                    End If
                End If
            End If
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            If pval.ItemUID = "Mat01" Then
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUID
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Else
                oLastItemUID = pval.ItemUID
                oLastColUID = ""
                oLastColRow = 0
            End If
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
            Exit Sub
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                If pval.ItemUID = "Mat01" Then
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)
                    End If
                End If
                If pval.ItemUID = "Mat01" Then
                    If pval.Row > 0 Then
                        oLastItemUID = pval.ItemUID
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Else
                    oLastItemUID = pval.ItemUID
                    oLastColUID = ""
                    oLastColRow = 0
                End If
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
                    If (pval.ItemUID = "Mat01") Then
                        Call oDS_PH_SY001L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
                        
                        If (pval.ColUID = "Code") Then
                            '//기타작업
                            If oMat1.RowCount = pval.Row And Trim(oDS_PH_SY001L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                                PH_SY001_AddMatrixRow (pval.Row)
                            End If
                        End If
                    
                        oMat1.FlushToDataSource
                        oMat1.LoadFromDataSource
                        oMat1.AutoResizeColumns
                        oForm.Update
                        oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK ct_Regular
                        
                    End If
           
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                Call PH_SY001_FormItemEnabled
                Call PH_SY001_AddMatrixRow(oMat1.VisualRowCount)
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
                Call PH_SY001_FormResize
            End If
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_SY001H", "Name")
            End If
'----------------------------------------------------------
        Case et_PICKER_CLICKED '//37
            
'----------------------------------------------------------
        Case et_GRID_SORT '//38
            
'----------------------------------------------------------
        Case et_Drag '//39
            
    End Select
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i As Long
    On Error GoTo Raise_FormMenuEvent_Error
    
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
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
    Exit Sub
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_SY001_FormItemEnabled
                Call PH_SY001_AddMatrixRow(0, True)
            Case "1284":
            Case "1286":
            Case "1293":
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":
                Call PH_SY001_FormItemEnabled
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":
                Call PH_SY001_FormItemEnabled
                Call PH_SY001_AddMatrixRow(0, True)
            Case "1288", "1289", "1290", "1291":
                Call PH_SY001_FormItemEnabled
        End Select
    End If
    Exit Sub
Raise_FormMenuEvent_Error:
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
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUID
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Else
        oLastItemUID = pval.ItemUID
        oLastColUID = ""
        oLastColRow = 0
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)

    Dim i As Long
    
    On Error GoTo Raise_EVENT_ROW_DELETE_Error

    If (oLastColRow > 0) Then
        If pval.BeforeAction = True Then
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat1.VisualRowCount
                oMat1.Columns("LineNum").Cells(i).Specific.Value = i
            Next i
            oMat1.FlushToDataSource
            Call oDS_PH_SY001L.RemoveRecord(oDS_PH_SY001L.Size - 1)
            oMat1.LoadFromDataSource
            If oMat1.RowCount = 0 Then
                Call PH_SY001_AddMatrixRow(0)
            Else
                If Trim(oDS_PH_SY001L.GetValue("U_Code", oMat1.RowCount - 1)) <> "" Then
                    Call PH_SY001_AddMatrixRow(oMat1.RowCount)
                End If
            End If
        End If
    End If
Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PH_SY001_AddMatrixRow(ByVal oRow As Long, Optional RowInserted As Boolean)
On Error GoTo PH_SY001_AddMatrixRow_Error
    Call oForm.Freeze(True)
    If RowInserted = False Then '//행추가여부
        oDS_PH_SY001L.InsertRecord (oRow)
    End If
    oMat1.AddRow
    oDS_PH_SY001L.Offset = oRow
    oDS_PH_SY001L.setValue "U_LineNum", oRow, oRow + 1
    oMat1.LoadFromDataSource
    Call oForm.Freeze(False)
    Exit Sub
PH_SY001_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_SY001_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_SY001_FormClear()
On Error GoTo PH_SY001_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_SY001'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_SY001_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_SY001_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_SY001_DataValidCheck() As Boolean
On Error GoTo PH_SY001_DataValidCheck_Error
    PH_SY001_DataValidCheck = False
    Dim i As Long
    If Trim(oForm.Items("Code").Specific.Value) = "" Then
        Sbo_Application.SetStatusBarMessage "코드는 필수입니다.", bmt_Short, True
        oForm.Items("Code").CLICK ct_Regular
        PH_SY001_DataValidCheck = False
        Exit Function
    End If
    If Trim(oForm.Items("Name").Specific.Value) = "" Then
        Sbo_Application.SetStatusBarMessage "코드명은 필수입니다.", bmt_Short, True
        oForm.Items("Name").CLICK ct_Regular
        PH_SY001_DataValidCheck = False
        Exit Function
    End If
    If oMat1.VisualRowCount <= 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PH_SY001_DataValidCheck = False
        Exit Function
    End If
    For i = 1 To oMat1.VisualRowCount - 1
        If (oMat1.Columns("Code").Cells(i).Specific.Value = "") Then
            Sbo_Application.SetStatusBarMessage "코드는 필수입니다.", bmt_Short, True
            oMat1.Columns("Code").Cells(i).CLICK ct_Regular
            PH_SY001_DataValidCheck = False
            Exit Function
        End If
        If (oMat1.Columns("Name").Cells(i).Specific.Value = "") Then
            Sbo_Application.SetStatusBarMessage "코드명은 필수입니다.", bmt_Short, True
            oMat1.Columns("Name").Cells(i).CLICK ct_Regular
            PH_SY001_DataValidCheck = False
            Exit Function
        End If
    Next
    Call oDS_PH_SY001L.RemoveRecord(oDS_PH_SY001L.Size - 1) '//마지막행제거
    Call oMat1.LoadFromDataSource
    PH_SY001_DataValidCheck = True
    Exit Function
PH_SY001_DataValidCheck_Error:
    PH_SY001_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_SY001_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_SY001_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordset As SAPbobsCOM.Recordset
    
    On Error GoTo PH_SY001_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("Param01").Specific.Value
    Param02 = oForm.Items("Param01").Specific.Value
    Param03 = oForm.Items("Param01").Specific.Value
    Param04 = oForm.Items("Param01").Specific.Value
    
    sQry = "SELECT 10"
    Call oRecordset.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordset.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_SY001_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordset.RecordCount, False)
    
    For i = 0 To oRecordset.RecordCount - 1
        If i <> 0 Then
            oDS_PH_SY001L.InsertRecord (i)
        End If
        oDS_PH_SY001L.Offset = i
        Call oDS_PH_SY001L.setValue("U_COL01", i, oRecordset.Fields(0).Value)
        Call oDS_PH_SY001L.setValue("U_COL02", i, oRecordset.Fields(1).Value)
        oRecordset.MoveNext
        ProgressBar01.Value = ProgressBar01.Value + 1
        ProgressBar01.Text = ProgressBar01.Value & "/" & oRecordset.RecordCount & "건 조회중...!"
    Next
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordset = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_SY001_MTX01_Exit:
    Set oRecordset = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_SY001_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordset = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_SY001_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_SY001_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_SY001_Validate_Error
    PH_SY001_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordset As SAPbobsCOM.Recordset
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_SY001H] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_SY001_Validate = False
        GoTo PH_SY001_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordset = Nothing
    Exit Function
PH_SY001_Validate_Exit:
    Set oRecordset = Nothing
    Exit Function
PH_SY001_Validate_Error:
    PH_SY001_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_SY001_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function







