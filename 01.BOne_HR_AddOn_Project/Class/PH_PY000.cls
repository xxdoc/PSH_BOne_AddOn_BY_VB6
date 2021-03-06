VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY000.cls
'//  Module         : 인사관리 > 사원마스터등록
'//  Desc           : 사원마스터등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY000A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY000B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY000.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY000_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY000"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY000_CreateItems
    Call PH_PY000_EnableMenus
    Call PH_PY000_SetDocument(oFromDocEntry01)
'    Call PH_PY000_FormResize
    
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

Private Function PH_PY000_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY000_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// 데이터셋정의
    '//----------------------------------------------------------------------------------------------
    '//테이블이 있을경우 데이터셋(Matrix)
    Set oDS_PH_PY000A = oForm.DataSources.DBDataSources("@PH_PY000A")   '//헤더
    Set oDS_PH_PY000B = oForm.DataSources.DBDataSources("@PH_PY000B")   '//라인

    Set oMat1 = oForm.Items("Mat1").Specific       '

    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns

    
    Set oColumn = oMat1.Columns("Value")
    oColumn.DisplayDesc = True
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
     
     
    '//매트릭스컬럼
    Set oColumn = oMat1.Columns("Value")
    oColumn.Editable = True

'    '//UserDataSources
'    Call oForm.DataSources.UserDataSources.Add("     ", dt_SHORT_TEXT, 10)
'    Set oCombo = oForm.Items("    ").Specific
'    oCombo.DataBind.SetBound True, "", "    "
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'    oForm.Items("CLTCOD").DisplayDesc = True
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY000_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY000_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY000_EnableMenus()
    
    On Error GoTo PH_PY000_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY000_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY000_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY000_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY000_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY000_FormItemEnabled
        Call PH_PY000_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY000_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY000_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY000_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY000_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY000_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

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
    
PH_PY000_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY000_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim ErrNum      As Integer
    Dim sTemp       As String
    
    Dim sTable  As String
    Dim sField1 As String
    Dim sField2 As String
                        
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    ErrNum = 0
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUid = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY000_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY000_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUid = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
'                            Call PH_PY000_FormItemEnabled
'                            Call PH_PY000_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
'                            Call PH_PY000_FormItemEnabled
'                            Call PH_PY000_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
'                            Call PH_PY000_FormItemEnabled
                        End If
                    End If
                End If

            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUid
            Case "Mat1", "Grid1"
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
            If pval.BeforeAction = True Then
                Select Case pval.ItemUid
                Case "Mat1", "Grid1"
                    If pval.Row > 0 Then
                        '//Call oMat1.SelectRow(pval.Row, True, False)

                        oLastItemUID = pval.ItemUid
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Case Else
                    oLastItemUID = pval.ItemUid
                    oLastColUID = ""
                    oLastColRow = 0
                End Select
            ElseIf pval.BeforeAction = False Then

            End If
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7
            If pval.BeforeAction = False Then
                If pval.ItemUid = "Mat1" And pval.ColUID = "UserCode" Then
                   
                End If
            End If
            
'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                If pval.ItemChanged = True Then
                    If pval.ItemUid = "Mat1" And pval.ColUID = "UserCode" Then
                        If oMat1.Columns("UserCode").Cells(pval.Row).Specific.Value <> "" Then
                            sTemp = MDC_SetMod.Get_ReData("U_Name", "User_Code", "OUSR", "'" & Trim(oMat1.Columns("UserCode").Cells(pval.Row).Specific.Value) & "'")
                            If sTemp <> "" Then
                                oMat1.Columns("UserName").Cells(pval.Row).Specific.Value = sTemp
                                oMat1.AutoResizeColumns
                            Else
                                oMat1.Columns("UserName").Cells(pval.Row).Specific.Value = ""
                                BubbleEvent = False
                                oMat1.AutoResizeColumns
                            End If
                        End If
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    If pval.ItemUid = "Mat1" And pval.ColUID = "UserCode" Then
                        PH_PY000_AddMatrixRow
                        oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK ct_Regular
                    End If
                    If pval.ItemUid = "CodField" Or pval.ItemUid = "NamField" Then
                        Set oColumn = oMat1.Columns("Value")
                        sTable = oForm.Items("DefTable").Specific.Value
                        sField1 = oForm.Items("CodField").Specific.Value
                        sField2 = oForm.Items("NamField").Specific.Value
                        
                        If sTable <> "" And sField1 <> "" And sField2 <> "" Then
                            If oColumn.ValidValues.Count > 0 Then
                                For i = oColumn.ValidValues.Count - 1 To 0 Step -1
                                    oColumn.ValidValues.Remove i, psk_Index
                                Next i
                            End If
                            
                            If TableFieldCheck(sTable, sField1, sField2) = True Then
                                sQry = "Select " & sField1 & "," & sField2 & " From " & sTable
                                oRecordSet.DoQuery sQry
                                If oRecordSet.RecordCount > 0 Then
                                    Do Until oRecordSet.EOF
                                        oColumn.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
                                        oRecordSet.MoveNext
                                    Loop
                                End If
                            End If
                            
                        Else
                            If oColumn.ValidValues.Count > 0 Then
                                For i = oColumn.ValidValues.Count - 1 To 0 Step -1
                                    oColumn.ValidValues.Remove i, psk_Index
                                Next i
                                oColumn.ValidValues.Add "", ""
                            End If
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

                Call PH_PY000_FormItemEnabled
                Call PH_PY000_AddMatrixRow
'
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
                Set oDS_PH_PY000A = Nothing
                Set oDS_PH_PY000B = Nothing

                Set oMat1 = Nothing
'                Set oGrid1 = Nothing
                
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
    
ErrMessage:
If ErrNum = 1 Then
        Sbo_Application.SetStatusBarMessage "테이블이 존재 하지 않습니다.", bmt_Short, True
    End If
    
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
                Call PH_PY000_FormItemEnabled
                Call PH_PY000_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY000_FormItemEnabled
                Call PH_PY000_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY000_FormItemEnabled
                Call PH_PY000_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY000_FormItemEnabled
            Case "1293" '// 행삭제
                '// [MAT1 용]
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource

                    While (i <= oDS_PH_PY000B.Size - 1)
                        If oDS_PH_PY000B.GetValue("U_UserCode", i) = "" Then
                            oDS_PH_PY000B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend

                    For i = 0 To oDS_PH_PY000B.Size
                        Call oDS_PH_PY000B.setValue("U_LineNum", i, i + 1)
                    Next i

                    oMat1.LoadFromDataSource
                End If
                Call PH_PY000_AddMatrixRow
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    Dim oColumn         As SAPbouiCOM.Column
    Dim sTable          As String
    Dim sField1         As String
    Dim sField2         As String
    Dim i               As Long
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormDataEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
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
                Set oColumn = oMat1.Columns("Value")
                sTable = oForm.Items("DefTable").Specific.Value
                sField1 = oForm.Items("CodField").Specific.Value
                sField2 = oForm.Items("NamField").Specific.Value
                
                If sTable <> "" And sField1 <> "" And sField2 <> "" Then
                    If oColumn.ValidValues.Count > 0 Then
                        For i = oColumn.ValidValues.Count - 1 To 0 Step -1
                            oColumn.ValidValues.Remove i, psk_Index
                        Next i
                    End If
                    
                    If TableFieldCheck(sTable, sField1, sField2) = True Then
                        sQry = "Select " & sField1 & "," & sField2 & " From " & sTable
                        oRecordSet.DoQuery sQry
                        If oRecordSet.RecordCount > 0 Then
                            Do Until oRecordSet.EOF
                                oColumn.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
                                oRecordSet.MoveNext
                            Loop
                        End If
                    End If
                    
                Else
                    If oColumn.ValidValues.Count > 0 Then
                        For i = oColumn.ValidValues.Count - 1 To 0 Step -1
                            oColumn.ValidValues.Remove i, psk_Index
                        Next i
                        oColumn.ValidValues.Add "", ""
                    End If
                End If
                oMat1.AutoResizeColumns
                
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

Sub PH_PY000_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY000_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1 용]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount

    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY000B.GetValue("U_UserCode", oRow - 1)) <> "" Then
            If oDS_PH_PY000B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY000B.InsertRecord (oRow)
            End If
            oDS_PH_PY000B.Offset = oRow
            oDS_PH_PY000B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY000B.setValue "U_UserCode", oRow, ""
            oDS_PH_PY000B.setValue "U_UserName", oRow, ""
            oDS_PH_PY000B.setValue "U_Value", oRow, ""
            oDS_PH_PY000B.setValue "U_Default", oRow, ""
            oDS_PH_PY000B.setValue "U_Comments", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY000B.Offset = oRow - 1
            oDS_PH_PY000B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY000B.setValue "U_UserCode", oRow - 1, ""
            oDS_PH_PY000B.setValue "U_UserName", oRow - 1, ""
            oDS_PH_PY000B.setValue "U_value", oRow - 1, ""
            oDS_PH_PY000B.setValue "U_value", oRow - 1, ""
            oDS_PH_PY000B.setValue "U_value", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY000B.Offset = oRow
        oDS_PH_PY000B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY000B.setValue "U_UserCode", oRow, ""
        oDS_PH_PY000B.setValue "U_UserName", oRow, ""
        oDS_PH_PY000B.setValue "U_Value", oRow, ""
        oDS_PH_PY000B.setValue "U_Default", oRow, ""
        oDS_PH_PY000B.setValue "U_Comments", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY000_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY000_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY000_FormClear()
On Error GoTo PH_PY000_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY000'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY000_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY000_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY000_DataValidCheck() As Boolean
On Error GoTo PH_PY000_DataValidCheck_Error
    PH_PY000_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim iLooper As Long
    Dim vMatrix As SAPbouiCOM.Matrix
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//필수 체크
    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY000A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        Exit Function
'    End If

    oMat1.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY000B.Size > 1 Then oDS_PH_PY000B.RemoveRecord (oDS_PH_PY000B.Size - 1)
    oMat1.LoadFromDataSource
    
    PH_PY000_DataValidCheck = True
    
'    Set vMatrix = oForm.Items("Mat1").Specific
'    vMatrix.Columns("UserCode").TitleObject.Sortable = True
'    vMatrix.Columns("UserCode").TitleObject.Sort gst_Ascending
'    vMatrix.Columns("UserCode").TitleObject.Sortable = False
    
    
    
    
    
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY000_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY000_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY000_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY000_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY000_Validate_Error
    PH_PY000_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY000A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY000_Validate = False
        GoTo PH_PY000_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY000_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY000_Validate_Error:
    PH_PY000_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY000_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PH_PY000_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    
    Dim BPLID           As String
    Dim ItmBsort        As String
    Dim DocDate         As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY000_Print_Report01_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY000_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[S142] 발주서"
    ReportName = "S142_1.rpt"
    sQry = "EXEC PH_PY000_1 '" & oForm.Items("8").Specific.Value & "'"
    
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
    
PH_PY000_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY000_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



