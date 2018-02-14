VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY130.cls
'//  Module         : 급여관리 > 급여마스터관리
'//  Desc           : 팀별 성과급차등 등급등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY130A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY130B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY130.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY130_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY130"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY130_CreateItems
    Call PH_PY130_EnableMenus
    Call PH_PY130_SetDocument(oFromDocEntry01)
'    Call PH_PY130_FormResize
    
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

Private Function PH_PY130_CreateItems() As Boolean

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
    
    On Error GoTo PH_PY130_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY130A = oForm.DataSources.DBDataSources("@PH_PY130A")
    Set oDS_PH_PY130B = oForm.DataSources.DBDataSources("@PH_PY130B")

    
    Set oMat1 = oForm.Items("Mat01").Specific       '//@PH_PY130B

    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns


    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    oForm.AutoManaged = True
    Call AutoManaged(oForm, "YM,JOBGBN,Number")
    
    '// 년월
    Call oDS_PH_PY130A.setValue("U_YM", 0, Format(Now, "YYYYMM"))
    
    '// 지급구분
    Set oCombo = oForm.Items("JOBGBN").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE CODE = 'P212' AND U_UseYN = 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JOBGBN").DisplayDesc = True

    '//사업장
    Set oColumn = oMat1.Columns("CLTCOD")
    sQry = "SELECT BPLId, BPLName FROM [OBPL]"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oColumn.DisplayDesc = True
    
    
    '//등급
    Set oColumn = oMat1.Columns("Grade")
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '82' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oColumn.DisplayDesc = True
    
    Call oForm.DataSources.UserDataSources.Add("S_A", dt_SUM)
    oForm.Items("S_A").Specific.DataBind.SetBound True, "", "S_A"
    
    Call oForm.DataSources.UserDataSources.Add("S_B", dt_SUM)
    oForm.Items("S_B").Specific.DataBind.SetBound True, "", "S_B"
    
    Call oForm.DataSources.UserDataSources.Add("S_C", dt_SUM)
    oForm.Items("S_C").Specific.DataBind.SetBound True, "", "S_C"
    
    Call oForm.DataSources.UserDataSources.Add("S_X", dt_SUM)
    oForm.Items("S_X").Specific.DataBind.SetBound True, "", "S_X"
    
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY130_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY130_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY130_EnableMenus()
    
    On Error GoTo PH_PY130_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY130_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY130_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY130_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY130_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY130_FormItemEnabled
'        Call PH_PY130_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY130_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY130_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY130_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY130_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY130_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    
    
    
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("Btn01").Visible = True
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        Call oForm.EnableMenu("1293", False)    '//행삭제
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("Btn01").Visible = True
        
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        Call oForm.EnableMenu("1293", False)    '//행삭제
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("Btn01").Visible = False
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        Call oForm.EnableMenu("1293", False)    '//행삭제

    End If
    oForm.Items("S_A").Specific.VALUE = 0
    oForm.Items("S_B").Specific.VALUE = 0
    oForm.Items("S_C").Specific.VALUE = 0
    oForm.Items("S_X").Specific.VALUE = 0
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY130_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY130_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim S_A, S_B, S_C, S_X As Long
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY130_DataValidCheck("Y") = False Then
                            BubbleEvent = False
                        End If
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        
                    End If
                End If
                If pval.ItemUID = "Btn01" Then
                    If PH_PY130_DataValidCheck("N") = False Then
                        BubbleEvent = False
                    Else
                        If PH_PY130_Create_Data = False Then
                            BubbleEvent = False
                        End If
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY130_FormItemEnabled
    '                        Call PH_PY130_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY130_FormItemEnabled
    '                        Call PH_PY130_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
    '                        Call PH_PY130_FormItemEnabled
                        End If
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
                    
                    
                    S_A = 0
                    S_B = 0
                    S_C = 0
                    S_X = 0
                        
                    If oMat1.VisualRowCount > 0 Then
                        For i = 0 To oMat1.VisualRowCount - 1
                            If oMat1.Columns("Grade").Cells(i + 1).Specific.VALUE = "A" Then
                                S_A = S_A + 1
                            ElseIf oMat1.Columns("Grade").Cells(i + 1).Specific.VALUE = "B" Then
                                S_B = S_B + 1
                            ElseIf oMat1.Columns("Grade").Cells(i + 1).Specific.VALUE = "C" Then
                                S_C = S_C + 1
                            ElseIf oMat1.Columns("Grade").Cells(i + 1).Specific.VALUE = "X" Then
                                S_X = S_X + 1
                            End If
                            
                        Next i
                        
                        oForm.Items("S_A").Specific.VALUE = S_A
                        oForm.Items("S_B").Specific.VALUE = S_B
                        oForm.Items("S_C").Specific.VALUE = S_C
                        oForm.Items("S_X").Specific.VALUE = S_X
                        
                    Else
                        oForm.Items("S_A").Specific.VALUE = 0
                        oForm.Items("S_B").Specific.VALUE = 0
                        oForm.Items("S_C").Specific.VALUE = 0
                        oForm.Items("S_X").Specific.VALUE = 0
                    End If
                        
                        
                    
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

                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource
                Call PH_PY130_FormItemEnabled
                
                
                S_A = 0
                S_B = 0
                S_C = 0
                S_X = 0
                
                If oMat1.VisualRowCount > 0 Then
                    For i = 0 To oMat1.VisualRowCount - 1
                        If oMat1.Columns("Grade").Cells(i + 1).Specific.VALUE = "A" Then
                            S_A = S_A + 1
                        ElseIf oMat1.Columns("Grade").Cells(i + 1).Specific.VALUE = "B" Then
                            S_B = S_B + 1
                        ElseIf oMat1.Columns("Grade").Cells(i + 1).Specific.VALUE = "C" Then
                            S_C = S_C + 1
                        ElseIf oMat1.Columns("Grade").Cells(i + 1).Specific.VALUE = "X" Then
                            S_X = S_X + 1
                        End If
                        
                    Next i
                    
                    oForm.Items("S_A").Specific.VALUE = S_A
                    oForm.Items("S_B").Specific.VALUE = S_B
                    oForm.Items("S_C").Specific.VALUE = S_C
                    oForm.Items("S_X").Specific.VALUE = S_X
                    
                Else
                    oForm.Items("S_A").Specific.VALUE = 0
                    oForm.Items("S_B").Specific.VALUE = 0
                    oForm.Items("S_C").Specific.VALUE = 0
                    oForm.Items("S_X").Specific.VALUE = 0
                End If
                
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
                Set oDS_PH_PY130A = Nothing
                Set oDS_PH_PY130B = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY130A", "Code")
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
'                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY130A", "Code")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY130_FormItemEnabled
'                Call PH_PY130_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY130_FormItemEnabled
'                Call PH_PY130_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY130_FormItemEnabled
'                Call PH_PY130_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY130_FormItemEnabled
            Case "1293" '// 행삭제
                
'                Call PH_PY130_AddMatrixRow
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

Sub PH_PY130_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY130_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY130B.GetValue("U_CLTCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY130B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY130B.InsertRecord (oRow)
            End If
            oDS_PH_PY130B.Offset = oRow
            oDS_PH_PY130B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY130B.setValue "U_Date", oRow, ""
            oDS_PH_PY130B.setValue "U_Type", oRow, ""
            oDS_PH_PY130B.setValue "U_Comments", oRow, 0
            oDS_PH_PY130B.setValue "U_Close", oRow, 0
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY130B.Offset = oRow - 1
            oDS_PH_PY130B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY130B.setValue "U_Date", oRow - 1, ""
            oDS_PH_PY130B.setValue "U_Type", oRow - 1, ""
            oDS_PH_PY130B.setValue "U_Comments", oRow - 1, 0
            oDS_PH_PY130B.setValue "U_Close", oRow - 1, 0
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY130B.Offset = oRow
        oDS_PH_PY130B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY130B.setValue "U_Date", oRow, ""
        oDS_PH_PY130B.setValue "U_Type", oRow, ""
        oDS_PH_PY130B.setValue "U_Comments", oRow, 0
        oDS_PH_PY130B.setValue "U_Close", oRow, 0
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY130_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY130_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY130_FormClear()
On Error GoTo PH_PY130_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY130'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY130_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY130_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY130_DataValidCheck(ChkYN As String) As Boolean
On Error GoTo PH_PY130_DataValidCheck_Error
    PH_PY130_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim tCode   As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    If ChkYN = "Y" Then
        If Trim(oDS_PH_PY130A.GetValue("U_JOBGBN", 0)) = "" Then
            Sbo_Application.SetStatusBarMessage "지급구분은 필수입니다.", bmt_Short, True
            oForm.Items("JOBGBN").CLICK ct_Regular
            PH_PY130_DataValidCheck = False
            Exit Function
        End If
    End If
    
    If Trim(oDS_PH_PY130A.GetValue("U_YM", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "년월은 필수입니다.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        PH_PY130_DataValidCheck = False
        Exit Function
    End If
    
    If Trim(oDS_PH_PY130A.GetValue("U_Number", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
        oForm.Items("Number").CLICK ct_Regular
        PH_PY130_DataValidCheck = False
        Exit Function
    End If
    
    '// 코드,이름 저장
    tCode = Trim(oDS_PH_PY130A.GetValue("U_YM", 0)) & Trim(oDS_PH_PY130A.GetValue("U_JOBGBN", 0)) & Trim(oDS_PH_PY130A.GetValue("U_Number", 0))
    Call oDS_PH_PY130A.setValue("Code", 0, tCode)
    Call oDS_PH_PY130A.setValue("Name", 0, tCode)
    
    If oForm.Mode = fm_ADD_MODE Then
        '// 데이터 중복 체크
        sQry = "SELECT Code FROM [@PH_PY130A] WHERE Code = '" & tCode & "'"
        oRecordSet.DoQuery sQry
        
        If oRecordSet.RecordCount > 0 Then
            Sbo_Application.SetStatusBarMessage "이미 데이터가 존재합니다.", bmt_Short, True
            PH_PY130_DataValidCheck = False
            Exit Function
        End If
    End If
    
    If ChkYN = "Y" Then
        If oMat1.VisualRowCount = 0 Then
            Sbo_Application.SetStatusBarMessage "데이터가 없습니다. 자료생성을 하기바랍니다", bmt_Short, True
            PH_PY130_DataValidCheck = False
            Exit Function
        End If
        
        For i = 1 To oMat1.VisualRowCount - 1
            '//구분
            If oMat1.Columns("Grade").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "등급은 필수입니다.", bmt_Short, True
                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
                PH_PY130_DataValidCheck = False
                Exit Function
            End If
        Next
    Else
        
    End If
    
    
    
    PH_PY130_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY130_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY130_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY130_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY130_Create_Data() As Boolean
    Dim i       As Long
    Dim sQry    As String
    
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY130_Create_Data_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    
    sQry = "exec [PH_PY130]"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY130_Create_Data_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY130B.InsertRecord (i)
        End If
        oDS_PH_PY130B.Offset = i
        Call oDS_PH_PY130B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY130B.setValue("U_CLTCOD", i, oRecordSet.Fields("CLTCOD").VALUE)
        Call oDS_PH_PY130B.setValue("U_TeamCode", i, oRecordSet.Fields("TeamCode").VALUE)
        Call oDS_PH_PY130B.setValue("U_TeamName", i, oRecordSet.Fields("TeamName").VALUE)
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
    Exit Function
PH_PY130_Create_Data_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
PH_PY130_Create_Data_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY130_Create_Data_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Function PH_PY130_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY130_Validate_Error
    PH_PY130_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY130A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY130_Validate = False
        GoTo PH_PY130_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY130_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY130_Validate_Error:
    PH_PY130_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY130_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


