VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY311"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY311.cls
'//  Module         : 인사관리 > 근태
'//  Desc           : 통근버스운행등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Public oGrid01              As SAPbouiCOM.Grid

Private oDS_PH_PY311A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY311B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long
Private oDocDate As String


Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY311.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY311_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY311"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE '//fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY311_CreateItems
    Call PH_PY311_EnableMenus
    Call PH_PY311_SetDocument(oFromDocEntry01)
'    Call PH_PY311_FormResize
    Call PH_PY311_Load_MonthData
    
    Call PH_PY311_AddMatrixRow
    
    
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

Private Function PH_PY311_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY311_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY311A = oForm.DataSources.DBDataSources("@PH_PY311A")
    Set oDS_PH_PY311B = oForm.DataSources.DBDataSources("@PH_PY311B")

    
    Set oMat1 = oForm.Items("Mat01").Specific       '//@PH_PY311B
    
    Set oGrid01 = oForm.Items("Grid01").Specific
    oForm.DataSources.DataTables.Add ("ZTEMP")
    
    
    oForm.Items("Code").Enabled = False
        
    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns


    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    Dim CLTCOD As String
    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
    
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.Select CLTCOD, psk_ByValue
    
    oForm.Items("CLTCOD").DisplayDesc = True
    
'    oForm.DataSources.UserDataSources.Item("CLTCOD").Value = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
    
    '// 년
    Call oDS_PH_PY311A.setValue("U_DocDate", 0, Format(Now, "YYYYMMDD"))
    
    Set oColumn = oMat1.Columns("BusGubun")
    oColumn.ValidValues.Add "1", "[1]출근"
    oColumn.ValidValues.Add "2", "[2]퇴근"
    
    oColumn.DisplayDesc = True
    
        
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY311_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY311_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY311_EnableMenus()
    
    On Error GoTo PH_PY311_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY311_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY311_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY311_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY311_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY311_FormItemEnabled
'        Call PH_PY311_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY311_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        
        
        oForm.Items("1").CLICK ct_Regular
        oForm.Items("Code").Enabled = False
        Call PH_PY311_AddMatrixRow
        
    End If
    Exit Sub
    
PH_PY311_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY311_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY311_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim CLTCOD As String
    
    On Error GoTo PH_PY311_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        Call oForm.EnableMenu("1293", True)    '//행삭제
        
        
        CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
        Set oCombo = oForm.Items("CLTCOD").Specific
        oCombo.Select CLTCOD, psk_ByValue
        
'        If Now = oDocDate Then
'            Call oDS_PH_PY311A.setValue("U_DocDate", 0, Format(Now, "YYYYMMDD")) 'Format(Now, "YYYYMMDD"))
'        Else
            Call oDS_PH_PY311A.setValue("U_DocDate", 0, oDocDate) 'Format(Now, "YYYYMMDD"))
'        End If
        
    
        Call PH_PY311_Load_MonthData
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        Call oForm.EnableMenu("1293", True)    '//행삭제
        
        CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
        Set oCombo = oForm.Items("CLTCOD").Specific
        oCombo.Select CLTCOD, psk_ByValue
        
'        Call oDS_PH_PY311A.setValue("U_DocDate", 0, Format(Now, "YYYYMMDD"))
'
'        Call PH_PY311_Load_MonthData
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        Call oForm.EnableMenu("1293", True)    '//행삭제
        
        CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
        Set oCombo = oForm.Items("CLTCOD").Specific
        oCombo.Select CLTCOD, psk_ByValue
        
        
'        Call oDS_PH_PY311A.setValue("U_DocDate", 0, Format(Now, "YYYYMMDD"))
        
        Call PH_PY311_Load_MonthData
        
    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY311_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY311_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim j           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim cnt         As Long
    Dim YM As String
    Dim YM1 As String
    Dim BusGubun(4) As String
    Dim BusTime(4) As Integer
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "1"
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY311_DataValidCheck("Y") = False Then
                            BubbleEvent = False
                        End If
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        
                    End If
                Case "Btn01"
                    If oForm.Mode = fm_ADD_MODE Then
                        cnt = oDS_PH_PY311B.Size
                        If cnt > 0 Then
                            For j = 0 To cnt - 1
                                oDS_PH_PY311B.RemoveRecord oDS_PH_PY311B.Size - 1
                            Next j
                            If cnt = 1 Then
                                oDS_PH_PY311B.Clear
                            End If
                        End If
                        oMat1.LoadFromDataSource
                        
                        BusTime(0) = 830
                        BusTime(1) = 830
                        BusTime(2) = 1730
                        BusTime(3) = 2030
                        BusTime(4) = 2030
                        
                        BusGubun(0) = "1"
                        BusGubun(1) = "2"
                        BusGubun(2) = "2"
                        BusGubun(3) = "1"
                        BusGubun(4) = "2"
                        
                        For i = 1 To 5
                            oDS_PH_PY311B.InsertRecord i - 1 '라인추가
                            oDS_PH_PY311B.setValue "U_LineNum", i - 1, i
                            oDS_PH_PY311B.setValue "U_BusTime", i - 1, BusTime(i - 1)
                            oDS_PH_PY311B.setValue "U_BusGubun", i - 1, BusGubun(i - 1)
                        Next
                        
                        oDS_PH_PY311B.InsertRecord i - 1 '라인추가
                        oDS_PH_PY311B.setValue "U_LineNum", i - 1, i
                        
                        oMat1.LoadFromDataSource
                    Else
                        Call MDC_Com.MDC_GF_Message("신규모드일때 입력가능합니다.", "E")
                    End If
                End Select
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY311_FormItemEnabled
    '                        Call PH_PY311_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY311_FormItemEnabled
                            
    '                        Call PH_PY311_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY311_AddMatrixRow
    '                        Call PH_PY311_FormItemEnabled
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
                Case "Grid01"

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
                If pval.ItemUID = "Grid01" Then
                    Call Search_Matrix_Data
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
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "DocDate" Then
                        YM = Left(oForm.Items("DocDate").Specific.VALUE, 6)
                        If oGrid01.Rows.Count > 1 Then
                            YM1 = Left(Format(Trim(oGrid01.DataTable.GetValue(0, 1)), "YYYYMMDD"), 6)
                        Else
                            YM1 = ""
                        End If
                        If YM <> YM1 Then
                            Call PH_PY311_Load_MonthData
                        End If
                    End If
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "BusTime" Then
                            Call PH_PY311_AddMatrixRow
                        End If
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource
                Call PH_PY311_FormItemEnabled
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
                Set oDS_PH_PY311A = Nothing
                Set oDS_PH_PY311B = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY311A", "Code")
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
                Call PH_PY311_FormItemEnabled
                
                
    
'                Call PH_PY311_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY311_FormItemEnabled
'                Call PH_PY311_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY311_FormItemEnabled
'                Call PH_PY311_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY311_FormItemEnabled
            Case "1293" '// 행삭제
                
'                Call PH_PY311_AddMatrixRow
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
                Call Delete_EmptyRow
            Case et_FORM_DATA_UPDATE:   '//35
                Call Delete_EmptyRow
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

Sub PH_PY311_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY311_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY311B.GetValue("U_BusTime", oRow - 1)) <> "" Then
            If oDS_PH_PY311B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY311B.InsertRecord (oRow)
            End If
            oDS_PH_PY311B.Offset = oRow
            oDS_PH_PY311B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY311B.setValue "U_BusTime", oRow, ""
            oDS_PH_PY311B.setValue "U_BusGubun", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY311B.Offset = oRow - 1
            oDS_PH_PY311B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY311B.setValue "U_BusTime", oRow - 1, ""
            oDS_PH_PY311B.setValue "U_BusGubun", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY311B.Offset = oRow
        oDS_PH_PY311B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY311B.setValue "U_BusTime", oRow, ""
        oDS_PH_PY311B.setValue "U_BusGubun", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY311_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY311_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY311_FormClear()
On Error GoTo PH_PY311_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY311'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY311_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY311_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY311_DataValidCheck(ChkYN As String) As Boolean
On Error GoTo PH_PY311_DataValidCheck_Error
    PH_PY311_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim tCode   As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    If ChkYN = "Y" Then
        If Trim(oDS_PH_PY311A.GetValue("U_CLTCOD", 0)) = "" Then
            Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
            oForm.Items("CLTCOD").CLICK ct_Regular
            PH_PY311_DataValidCheck = False
            Exit Function
        End If
    End If
    
    If Trim(oDS_PH_PY311A.GetValue("U_DocDate", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "일자는 필수입니다.", bmt_Short, True
        oForm.Items("DocDate").CLICK ct_Regular
        PH_PY311_DataValidCheck = False
        Exit Function
    End If
    
    
    '// 코드,이름 저장
    'tCode = Trim(oDS_PH_PY311A.GetValue("U_CLTCOD", 0)) & Right(Format(oDS_PH_PY311A.GetValue("U_DocDate", 0), "YYYYMMDD"), 6)
    tCode = Trim(oDS_PH_PY311A.GetValue("U_CLTCOD", 0)) + Right(Trim(oDS_PH_PY311A.GetValue("U_DocDate", 0)), 6)
    Call oDS_PH_PY311A.setValue("Code", 0, tCode)
    Call oDS_PH_PY311A.setValue("Name", 0, tCode)
    
    If oForm.Mode = fm_ADD_MODE Then
        '// 데이터 중복 체크
        sQry = "SELECT Code FROM [@PH_PY311A] WHERE Code = '" & tCode & "'"
        oRecordSet.DoQuery sQry
        
        If oRecordSet.RecordCount > 0 Then
            Sbo_Application.SetStatusBarMessage "이미 데이터가 존재합니다.", bmt_Short, True
            PH_PY311_DataValidCheck = False
            Exit Function
        End If
    End If
    
    If ChkYN = "Y" Then
        If oMat1.VisualRowCount = 0 Then
            Sbo_Application.SetStatusBarMessage "데이터가 없습니다. 월을생성을 하기바랍니다", bmt_Short, True
            PH_PY311_DataValidCheck = False
            Exit Function
        End If
    End If
    
    PH_PY311_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY311_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY311_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY311_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Function PH_PY311_Load_MonthData() As Boolean
    Dim cnt&, i&
    
    Dim CLTCOD      As String
    Dim DocDate      As String
    Dim YM  As String
    Dim YM1  As String
    
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    DocDate = oForm.Items("DocDate").Specific.VALUE
    
    
    sQry = "EXEC PH_PY311 '" & CLTCOD & "', '" & DocDate & "'"
    
    '/ Procedure 실행(Grid 사용)
    oForm.DataSources.DataTables.Item(0).ExecuteQuery (sQry)
    oGrid01.DataTable = oForm.DataSources.DataTables.Item("ZTEMP")
    
    Call GridSetting
   
    
    If oForm.Mode = fm_OK_MODE Then
        oForm.Mode = fm_UPDATE_MODE
    End If

    Sbo_Application.StatusBar.SetText "작업을 완료하였습니다.", bmt_Short, smt_Success
        
    oDocDate = oForm.Items("DocDate").Specific.VALUE
                        
    
    
    PH_PY311_Load_MonthData = True
End Function



Function PH_PY311_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY311_Validate_Error
    PH_PY311_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY311A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY311_Validate = False
        GoTo PH_PY311_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY311_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY311_Validate_Error:
    PH_PY311_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY311_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


'****************************************************************************************************************
'// Grid 꾸며주기
'****************************************************************************************************************
Private Sub GridSetting()
On Error GoTo GridSetting_Error
    Dim i           As Integer
    Dim sColsTitle  As String
    Dim sColsLine   As String
    
    oForm.Freeze True
    
    oGrid01.SelectionMode = ms_Single
    
'    oGrid01.Columns.Item(0).LinkedObjectType = lf_BusinessPartner
'    oGrid01.Columns.Item(4).LinkedObjectType = lf_Invoice
'     oGrid01.Columns.Item(2).LinkedObjectType = lf_Items
       
    For i = 0 To oGrid01.Columns.Count - 1
        sColsTitle = oGrid01.Columns(i).TitleObject.Caption
        
        oGrid01.Columns(i).Editable = False
        
        If oGrid01.DataTable.Columns(i).Type = ft_Float Then
            oGrid01.Columns(i).RightJustified = True
        End If
        
'        If InStr(1, sColsTitle, "T(mm)") > 0 Or _
'           InStr(1, sColsTitle, "W(mm)") > 0 Or _
'           InStr(1, sColsTitle, "L(mm)") > 0 Or _
'           InStr(1, sColsTitle, "Weight") > 0 Or _
'           InStr(1, sColsTitle, "Qty(Kg)") > 0 Then
'            oGrid01.Columns(i).RightJustified = True
'            oGrid01.Columns(i).BackColor = &HE0E0E0
'        End If
'
'        If InStr(1, sColsTitle, "수량") > 0 Or _
'           InStr(1, sColsTitle, "중량") > 0 Then
'            oGrid01.Columns(i).RightJustified = True
'            oGrid01.Columns(i).BackColor = &HFFC0C0
'        End If
    Next i
        
    oForm.Freeze False
    
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
GridSetting_Error:
    MDC_Com.MDC_GF_Message "GridSetting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub Search_Matrix_Data()
    On Error GoTo Search_Matrix_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim DocDate As String
    Dim i                   As Integer
    Dim CLTCOD  As String
    
    
    
    oForm.Freeze True
    
     For i = 0 To oGrid01.Rows.Count - 1
        If oGrid01.Rows.IsSelected(i) = True Then
            CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
            DocDate = Format(Trim(oGrid01.DataTable.GetValue(0, i)), "YYYYMMDD")
            
            If Trim(oGrid01.DataTable.GetValue(4, i)) = "운행" Then
               oForm.Items("Code").Enabled = True
               PH_PY311_SetDocument (CLTCOD + Right(DocDate, 6))
               'PH_PY311_SetDocument (DocDate)
            Else
                If oForm.Mode = fm_ADD_MODE Then
                    oForm.Items("DocDate").Specific.VALUE = DocDate
                Else
'                    oForm.Mode = fm_ADD_MODE
'
'                    oForm.Items("DocDate").Specific.Value = DocDate
                End If
            End If
        End If
     Next i
    
    oDocDate = oForm.Items("DocDate").Specific.VALUE
    
    oForm.Freeze False
    Set oRecordSet = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Matrix_Data_Error:
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "Search_Matrix_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error
    Dim i&
    
    oMat1.FlushToDataSource
    
    For i = 0 To oMat1.VisualRowCount - 1
        If Trim(oDS_PH_PY311B.GetValue("U_BusTime", i)) = "" Then
            oDS_PH_PY311B.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat1.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

