VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY312"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY312.cls
'//  Module         : 근태관리 > 마스터등록
'//  Desc           : 버스요금 등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix
Public oMat2 As SAPbouiCOM.Matrix

Private oDS_PH_USERDS01 As SAPbouiCOM.DBDataSource
Private oDS_PH_USERDS02 As SAPbouiCOM.DBDataSource

Private oMat01Row01 As Long

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY312.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY312_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY312"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY312_CreateItems
    Call PH_PY312_EnableMenus
    Call PH_PY312_SetDocument(oFromDocEntry01)
'    Call PH_PY312_FormResize
    
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

Private Function PH_PY312_CreateItems() As Boolean

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
    
    On Error GoTo PH_PY312_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_USERDS01 = oForm.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PH_USERDS02 = oForm.DataSources.DBDataSources("@PS_USERDS02")
    
    
    
    Set oMat1 = oForm.Items("Mat01").Specific
    Set oMat2 = oForm.Items("Mat02").Specific
    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    
    oMat2.SelectionMode = ms_NotSupported
    oMat2.AutoResizeColumns


    
    
    Call oForm.DataSources.UserDataSources.Add("DocDate", dt_DATE, 100)
    Call oForm.DataSources.UserDataSources.Add("Amt1", dt_SHORT_TEXT, 10)
    Call oForm.DataSources.UserDataSources.Add("Amt2", dt_SHORT_TEXT, 10)
    
    oForm.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    oForm.Items("Amt1").Specific.DataBind.SetBound True, "", "Amt1"
    oForm.Items("Amt2").Specific.DataBind.SetBound True, "", "Amt2"
    
    sQry = "Select MAX(DocDate) From YPH_PY312A"
    oRecordSet.DoQuery sQry
    oForm.Items("DocDate").Specific.VALUE = Format(oRecordSet.Fields(0).VALUE, "YYYYMMDD")
    
    
    Call PH_PY312_MTX01
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY312_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY312_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY312_EnableMenus()
    
    On Error GoTo PH_PY312_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY312_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY312_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY312_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY312_SetDocument_Error
    
'    If (oFromDocEntry01 = "") Then
'        Call PH_PY312_FormItemEnabled
''        Call PH_PY312_AddMatrixRow
'    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY312_FormItemEnabled
'        oForm.Items("Code").Specific.Value = oFromDocEntry01
'        oForm.Items("1").CLICK ct_Regular
'    End If
    Exit Sub
    
PH_PY312_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY312_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY312_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY312_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
                
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        Call oForm.EnableMenu("1293", False)    '//행삭제
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", False)     '//문서추가
        Call oForm.EnableMenu("1293", False)    '//행삭제
    ElseIf (oForm.Mode = fm_OK_MODE) Then
                
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", False)     '//문서추가
        Call oForm.EnableMenu("1293", False)    '//행삭제

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY312_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY312_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    If oForm.Mode = fm_ADD_MODE Then
                        
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        
                    End If
                End If
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY312_FormItemEnabled
    '                        Call PH_PY312_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY312_FormItemEnabled
    '                        Call PH_PY312_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
    '                        Call PH_PY312_FormItemEnabled
                        End If
                    End If
                End If
                
                If pval.ItemUID = "Button01" Then
                    Call PH_PY312_MTX01
                End If
                
                If pval.ItemUID = "Button02" Then
                    Dim MSTCOD As String
                    If oMat2.VisualRowCount > 0 Then
                        MSTCOD = oMat2.Columns("MSTCOD").Cells(1).Specific.VALUE
                        
                        Call PH_PY312_UPDATE
                        
                        Call PH_PY312_MTX01
                        
                        Call PH_PY312_MTX02(MSTCOD)
                    End If
                End If
                
                If pval.ItemUID = "Button03" Then
                    If Trim(oForm.Items("DocDate").Specific.VALUE) = "" Then
                        Sbo_Application.SetStatusBarMessage "일자를 입력하지 않습니다.", bmt_Short, True
                        
                        Exit Sub
                    Else
                        If Sbo_Application.MessageBox("현재 일자와 금액으로 일괄 적용하시겠습니까? ", 2, "Yes", "No") = 2 Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        Call PH_PY312_APPLY
                        
                        Call PH_PY312_MTX01
                        
                    End If
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
                        If Val(oMat1.Columns("Amt1").Cells(pval.Row).Specific.VALUE) + Val(oMat1.Columns("Amt2").Cells(pval.Row).Specific.VALUE) > 0 Then
                            Call PH_PY312_MTX02(oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE)
                        Else
                            oMat2.Clear
                            oMat2.FlushToDataSource
                            oMat2.AddRow
                            oMat2.Columns("MSTCOD").Cells(1).Specific.VALUE = oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE
                            oMat2.Columns("FullName").Cells(1).Specific.VALUE = oMat1.Columns("FullName").Cells(pval.Row).Specific.VALUE
                            oMat2.Columns("DocDate").Cells(1).Specific.VALUE = Mid(oForm.Items("DocDate").Specific.VALUE, 1, 4) + "-" + Mid(oForm.Items("DocDate").Specific.VALUE, 5, 2) + "-" + Mid(oForm.Items("DocDate").Specific.VALUE, 7, 2)
                        End If
                    End If
                Case "Mat02"
                    If pval.Row > 0 Then
                        Call oMat2.SelectRow(pval.Row, True, False)
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

                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource
                Call PH_PY312_FormItemEnabled
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
                Set oDS_PH_USERDS01 = Nothing
                Set oDS_PH_USERDS02 = Nothing
                
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
'                oForm.Items.Item("79").Width = oForm.Items.Item("KUKGRD").Left + oForm.Items.Item("KUKGRD").Width - oForm.Items.Item("79").Left + 10
'                oForm.Items.Item("79").Height = oForm.Items.Item("80").Height
'
'                oForm.Items.Item("77").Width = oForm.Items.Item("BUYN20").Left + oForm.Items.Item("BUYN20").Width - oForm.Items.Item("77").Left + 16
'                oForm.Items.Item("77").Height = oForm.Items.Item("78").Height
'
'                oForm.Items.Item("8").Width = oForm.Items.Item("Mat02").Left + oForm.Items.Item("Mat02").Width + 5
'                oForm.Items.Item("8").Height = oForm.Items.Item("1").Top - 80
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY312A", "Code")
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
                Call PH_PY312_FormItemEnabled
'                Call PH_PY312_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY312_FormItemEnabled
'                Call PH_PY312_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY312_FormItemEnabled
'                Call PH_PY312_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY312_FormItemEnabled
            Case "1293" '// 행삭제
                
'                Call PH_PY312_AddMatrixRow
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


Sub PH_PY312_FormClear()
On Error GoTo PH_PY312_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY312'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY312_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY312_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY312_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY312_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    
    sQry = "EXEC PH_PY312_01"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    oMat2.Clear
    oMat2.FlushToDataSource
    oMat2.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY312_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_USERDS01.InsertRecord (i)
        End If
        oDS_PH_USERDS01.Offset = i
        Call oDS_PH_USERDS01.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_USERDS01.setValue("U_ColReg01", i, oRecordSet.Fields(0).VALUE)
        Call oDS_PH_USERDS01.setValue("U_ColReg02", i, oRecordSet.Fields(1).VALUE)
        Call oDS_PH_USERDS01.setValue("U_ColReg03", i, oRecordSet.Fields(2).VALUE)
        Call oDS_PH_USERDS01.setValue("U_ColReg04", i, oRecordSet.Fields(3).VALUE)
        Call oDS_PH_USERDS01.setValue("U_ColReg05", i, oRecordSet.Fields(4).VALUE)
        Call oDS_PH_USERDS01.setValue("U_ColReg06", i, oRecordSet.Fields(5).VALUE)
        Call oDS_PH_USERDS01.setValue("U_ColSum01", i, oRecordSet.Fields(6).VALUE)
        Call oDS_PH_USERDS01.setValue("U_ColSum02", i, oRecordSet.Fields(7).VALUE)
        oRecordSet.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 조회중...!"
    Next
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY312_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY312_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY312_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY312_MTX02(MSTCOD As String)

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
        
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY312_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    
    
    sQry = "EXEC PH_PY312_02 '" & MSTCOD & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat2.Clear
    oMat2.FlushToDataSource
    oMat2.LoadFromDataSource
    
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY312_MTX02_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_USERDS02.InsertRecord (i)
        End If
        oDS_PH_USERDS02.Offset = i
        Call oDS_PH_USERDS02.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_USERDS02.setValue("U_ColReg01", i, oRecordSet.Fields(0).VALUE)
        Call oDS_PH_USERDS02.setValue("U_ColReg02", i, oRecordSet.Fields(1).VALUE)
        Call oDS_PH_USERDS02.setValue("U_ColReg03", i, oRecordSet.Fields(2).VALUE)
        Call oDS_PH_USERDS02.setValue("U_ColSum01", i, oRecordSet.Fields(3).VALUE)
        Call oDS_PH_USERDS02.setValue("U_ColSum02", i, oRecordSet.Fields(4).VALUE)
        oRecordSet.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 조회중...!"
    Next
    oMat2.LoadFromDataSource
    oMat2.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY312_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY312_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY312_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY312_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY312_Validate_Error
    PH_PY312_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY312A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY312_Validate = False
        GoTo PH_PY312_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY312_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY312_Validate_Error:
    PH_PY312_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY312_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY312_UPDATE() As Boolean
    Dim i           As Long
    
    Dim sQry        As String
    Dim MSTCOD      As String
    Dim CLTCOD      As String
    
    Dim DocDate As Date
    Dim Amt1 As Long
    Dim Amt2 As Long
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim IsNew As String '신규여부
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    On Error GoTo PH_PY312_UPDATE_Error
    
    PH_PY312_UPDATE = False
    
    oMat2.FlushToDataSource
    
    MSTCOD = oMat2.Columns("MSTCOD").Cells(1).Specific.VALUE
    
    sQry = " Select Count(*) From YPH_PY312A Where MSTCOD = '" & MSTCOD & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields(0).VALUE > 0 Then
        IsNew = "N"
    Else
        IsNew = "Y"
    End If
    
    
    For i = 1 To oMat2.VisualRowCount
        If Trim(oMat2.Columns("MSTCOD").Cells(i).Specific.VALUE) <> "" Then
        
            MSTCOD = oMat2.Columns("MSTCOD").Cells(i).Specific.VALUE
            DocDate = oMat2.Columns("DocDate").Cells(i).Specific.VALUE
            Amt1 = oMat2.Columns("Amt1").Cells(i).Specific.VALUE
            Amt2 = oMat2.Columns("Amt2").Cells(i).Specific.VALUE
            
            If IsNew = "N" Then
                sQry = " Update YPH_PY312A Set Amt1 = " & Amt1 & ", Amt2 = " & Amt2 & " Where MSTCOD = '" & MSTCOD & "' And DocDate = '" & DocDate & "'"
            Else
                CLTCOD = "1"
                
    '            If oMat1.Columns("InForm02").Cells(pval.Row).Specific.VALUE = "N" Then
    '                '통근버스 사용 Y/N
    '                Amt1 = 0
    '            End If
    '
    '            If oMat1.Columns("InForm03").Cells(pval.Row).Specific.VALUE = "N" Then
    '                '신마산 사용 Y/N
    '                Amt2 = 0
    '            End If
                
                sQry = "INSERT INTO YPH_PY312A"
                sQry = sQry & " ("
                sQry = sQry & " CLTCOD,"
                sQry = sQry & " MSTCOD,"
                sQry = sQry & " DocDate,"
                sQry = sQry & " Amt1,"
                sQry = sQry & " Amt2"
                sQry = sQry & " ) "
                sQry = sQry & "VALUES("
                sQry = sQry & "'" & CLTCOD & "',"
                sQry = sQry & "'" & MSTCOD & "',"
                sQry = sQry & "'" & DocDate & "',"
                sQry = sQry & Amt1 & ","
                sQry = sQry & Amt2 & ")"
                 
            End If
            oRecordSet.DoQuery sQry
        End If
    Next
    
    
    PH_PY312_UPDATE = True
        
    Set oRecordSet = Nothing
    Exit Function
    
PH_PY312_UPDATE_Error:
    Set oRecordSet = Nothing
    PH_PY312_UPDATE = False
    Sbo_Application.SetStatusBarMessage "PH_PY312_UPDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Function PH_PY312_APPLY() As Boolean
    Dim i           As Long
    
    Dim sQry        As String
    
    Dim MSTCOD      As String
    Dim CLTCOD      As String
    Dim DocDate As String
    Dim Amt1 As Double
    Dim Amt2 As Double
    
    Dim InAmt1 As Double
    Dim InAmt2 As Double
    
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    Dim oRecordSet02  As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    Set oRecordSet02 = oCompany.GetBusinessObject(BoRecordset)
    
    
    
    On Error GoTo PH_PY312_APPLY_Error
    
    PH_PY312_APPLY = False
    
    oMat1.FlushToDataSource
    
    DocDate = oForm.Items("DocDate").Specific.VALUE
    Amt1 = oForm.Items("Amt1").Specific.VALUE
    Amt2 = oForm.Items("Amt2").Specific.VALUE
    
    CLTCOD = "1" '//창원사업장
    
    sQry = "Select Count(*) From YPH_PY312A Where DocDate = '" & DocDate & "'"
    oRecordSet01.DoQuery sQry
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oMat1.VisualRowCount, False)
    ProgressBar01.VALUE = 0
    
    If oRecordSet01.Fields(0).VALUE = 0 Then
        '//신규일자 Insert
        For i = 1 To oMat1.VisualRowCount
            MSTCOD = oMat1.Columns("MSTCOD").Cells(i).Specific.VALUE
            
            If oMat1.Columns("InForm02").Cells(i).Specific.VALUE = "Y" Then
                InAmt1 = Amt1
            Else
                InAmt1 = 0
            End If
            
            If oMat1.Columns("InForm03").Cells(i).Specific.VALUE = "Y" Then
                InAmt2 = Amt2
            Else
                InAmt2 = 0
            End If
            
            sQry = "insert into YPH_PY312A "
            sQry = sQry & " ( CLTCOD, "
            sQry = sQry & " MSTCOD, "
            sQry = sQry & " DocDate, "
            sQry = sQry & " Amt1, "
            sQry = sQry & " Amt2 ) "
            sQry = sQry & "VALUES("
            sQry = sQry & "'" & CLTCOD & "',"
            sQry = sQry & "'" & MSTCOD & "',"
            sQry = sQry & "'" & DocDate & "',"
            sQry = sQry & InAmt1 & ","
            sQry = sQry & InAmt2 & ")"
            
            oRecordSet02.DoQuery sQry
            
            ProgressBar01.VALUE = ProgressBar01.VALUE + 1
            ProgressBar01.Text = ProgressBar01.VALUE & "/" & oMat1.VisualRowCount & "건 처리중...!"
        Next
    Else
        '//기존일자에 Update
        For i = 1 To oMat1.VisualRowCount
            MSTCOD = oMat1.Columns("MSTCOD").Cells(i).Specific.VALUE
            
            If oMat1.Columns("InForm02").Cells(i).Specific.VALUE = "Y" Then
                InAmt1 = Amt1
            Else
                InAmt1 = 0
            End If
            
            If oMat1.Columns("InForm03").Cells(i).Specific.VALUE = "Y" Then
                InAmt2 = Amt2
            Else
                InAmt2 = 0
            End If
            
            sQry = " Update YPH_PY312A Set Amt1 = " & InAmt1 & ", Amt2 = " & InAmt2 & " Where MSTCOD = '" & MSTCOD & "' And DocDate = '" & DocDate & "'"
            oRecordSet02.DoQuery sQry
            
            ProgressBar01.VALUE = ProgressBar01.VALUE + 1
            ProgressBar01.Text = ProgressBar01.VALUE & "/" & oMat1.VisualRowCount & "건 처리중...!"
        Next
    End If
    
    
    PH_PY312_APPLY = True
        
    Set oRecordSet01 = Nothing
    Set oRecordSet02 = Nothing
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    
    Exit Function
    
PH_PY312_APPLY_Error:
    Set oRecordSet01 = Nothing
    Set oRecordSet02 = Nothing
    Set ProgressBar01 = Nothing
    PH_PY312_APPLY = False
    Sbo_Application.SetStatusBarMessage "PH_PY312_APPLY_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



