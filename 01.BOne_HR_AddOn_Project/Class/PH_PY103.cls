VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY103.cls
'//  Module         : 급여관리 > 급여관리
'//  Desc           : 공제항목설정
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY103A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY103B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY103.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY103_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY103"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY103_CreateItems
    Call PH_PY103_EnableMenus
    Call PH_PY103_SetDocument(oFromDocEntry01)
'    Call PH_PY103_FormResize
    
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

Private Function PH_PY103_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY103_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY103A = oForm.DataSources.DBDataSources("@PH_PY103A")
    Set oDS_PH_PY103B = oForm.DataSources.DBDataSources("@PH_PY103B")

    
    Set oMat1 = oForm.Items("Mat1").Specific       '//@PH_PY103B

    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns

    Call PH_PY103_ChooseFromList
    
    
    '// 헤더 ------------------------------------------------------------------------------
    '// 사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'    oCombo.Select 0, psk_Index
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// 라인 ------------------------------------------------------------------------------
    
    '//G/L 계정
    Set oColumn = oMat1.Columns("ACTCOD")
    oColumn.ChooseFromListUID = "CFL_GLCode"
    oColumn.ChooseFromListAlias = "AcctCode"
    
    '// 수당코드
    Set oColumn = oMat1.Columns("CSUCOD")
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P143' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = False
    
    '// 고정항목
    Set oColumn = oMat1.Columns("FIXGBN")
    oColumn.ValidValues.Add "Y", "고정"
    oColumn.ValidValues.Add "N", "계산"
    oColumn.ValidValues.Add "V", "변동"
    oColumn.DisplayDesc = True
    
    '// 끝전처리
    Set oColumn = oMat1.Columns("ROUNDT")
    oColumn.ValidValues.Add "R", "반올림"
    oColumn.ValidValues.Add "F", "절사"
    oColumn.ValidValues.Add "C", "절상"
    oColumn.DisplayDesc = True
    
    '// 단위
    Set oColumn = oMat1.Columns("LENGTH")
    oColumn.ValidValues.Add "1", "원"
    oColumn.ValidValues.Add "10", "십원"
    oColumn.ValidValues.Add "100", "백원"
    oColumn.ValidValues.Add "1000", "천원"
    oColumn.DisplayDesc = True
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY103_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY103_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY103_EnableMenus()
    
    On Error GoTo PH_PY103_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY103_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY103_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY103_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY103_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY103_FormItemEnabled
        Call PH_PY103_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY103_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY103_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY103_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY103_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY103_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("YM").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY103_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY103_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                        If PH_PY103_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY103_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY103_FormItemEnabled
                            Call PH_PY103_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY103_FormItemEnabled
                            Call PH_PY103_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY103_FormItemEnabled
                        End If
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
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
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat1" And pval.ColUID = "CSUCOD" Then
                        oMat1.Columns("CSUNAM").Cells(pval.Row).Specific.Value = MDC_SetMod.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat1.Columns("CSUCOD").Cells(pval.Row).Specific.Value & "'", " AND Code = 'P143'")
                        Call PH_PY103_AddMatrixRow
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat1"
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)
                    End If
                End Select
                
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
                    If pval.ItemUID = "Mat1" And pval.ColUID = "CSUNAM" Then
                        If oMat1.Columns("CSUCOD").Cells(pval.Row).Specific.Value <> "" Then
                            Call PH_PY103_AddMatrixRow
                            Call oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                            Call oMat1.AutoResizeColumns
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

                Call PH_PY103_FormItemEnabled
                Call PH_PY103_AddMatrixRow
                
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
                Set oDS_PH_PY103A = Nothing
                Set oDS_PH_PY103B = Nothing
                
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
                If pval.ItemUID = "Mat1" And pval.ColUID = "ACTCOD" Then
                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY103B", "U_ACTCOD,U_ACTNAM", pval.ItemUID, pval.Row)
                End If
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY103A", "Code")      '//접속자 권한에 따른 사업장 보기
        
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY103_FormItemEnabled
                Call PH_PY103_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY103_FormItemEnabled
                Call PH_PY103_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY103_FormItemEnabled
                Call PH_PY103_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY103_FormItemEnabled
            Case "1293" '// 행삭제
                '// [MAT1] 급여 수당
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY103B.Size - 1)
                        If oDS_PH_PY103B.GetValue("U_CSUCOD", i) = "" Then
                            oDS_PH_PY103B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY103B.Size
                        Call oDS_PH_PY103B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY103_AddMatrixRow
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

Sub PH_PY103_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY103_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY103B.GetValue("U_CSUCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY103B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY103B.InsertRecord (oRow)
            End If
            oDS_PH_PY103B.Offset = oRow
            oDS_PH_PY103B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY103B.setValue "U_LINSEQ", oRow, ""
            oDS_PH_PY103B.setValue "U_CSUCOD", oRow, ""
            oDS_PH_PY103B.setValue "U_CSUNAM", oRow, ""
            oDS_PH_PY103B.setValue "U_BNSUSE", oRow, "N"
            oDS_PH_PY103B.setValue "U_FIXGBN", oRow, "N"
            oDS_PH_PY103B.setValue "U_INSLIN", oRow, ""
            oDS_PH_PY103B.setValue "U_ACTCOD", oRow, ""
            oDS_PH_PY103B.setValue "U_ACTNAM", oRow, ""
            oDS_PH_PY103B.setValue "U_ROUNDT", oRow, ""
            oDS_PH_PY103B.setValue "U_LENGTH", oRow, ""
            oDS_PH_PY103B.setValue "U_SILCUN", oRow, ""
            oDS_PH_PY103B.setValue "U_REMARK", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY103B.Offset = oRow - 1
            oDS_PH_PY103B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY103B.setValue "U_LINSEQ", oRow - 1, ""
            oDS_PH_PY103B.setValue "U_CSUCOD", oRow - 1, ""
            oDS_PH_PY103B.setValue "U_CSUNAM", oRow - 1, ""
            oDS_PH_PY103B.setValue "U_BNSUSE", oRow - 1, "N"
            oDS_PH_PY103B.setValue "U_FIXGBN", oRow - 1, "N"
            oDS_PH_PY103B.setValue "U_INSLIN", oRow - 1, ""
            oDS_PH_PY103B.setValue "U_ACTCOD", oRow - 1, ""
            oDS_PH_PY103B.setValue "U_ACTNAM", oRow - 1, ""
            oDS_PH_PY103B.setValue "U_ROUNDT", oRow - 1, ""
            oDS_PH_PY103B.setValue "U_LENGTH", oRow - 1, ""
            oDS_PH_PY103B.setValue "U_SILCUN", oRow - 1, ""
            oDS_PH_PY103B.setValue "U_REMARK", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY103B.Offset = oRow
        oDS_PH_PY103B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY103B.setValue "U_LINSEQ", oRow, ""
        oDS_PH_PY103B.setValue "U_CSUCOD", oRow, ""
        oDS_PH_PY103B.setValue "U_CSUNAM", oRow, ""
        oDS_PH_PY103B.setValue "U_BNSUSE", oRow, "N"
        oDS_PH_PY103B.setValue "U_FIXGBN", oRow, "N"
        oDS_PH_PY103B.setValue "U_INSLIN", oRow, ""
        oDS_PH_PY103B.setValue "U_ACTCOD", oRow, ""
        oDS_PH_PY103B.setValue "U_ACTNAM", oRow, ""
        oDS_PH_PY103B.setValue "U_ROUNDT", oRow, ""
        oDS_PH_PY103B.setValue "U_LENGTH", oRow, ""
        oDS_PH_PY103B.setValue "U_SILCUN", oRow, ""
        oDS_PH_PY103B.setValue "U_REMARK", oRow, ""
        
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY103_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY103_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY103_FormClear()
On Error GoTo PH_PY103_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY103'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY103_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY103_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY103_DataValidCheck() As Boolean
On Error GoTo PH_PY103_DataValidCheck_Error
    PH_PY103_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '//헤더
    If Trim(oDS_PH_PY103A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY103_DataValidCheck = False
        Exit Function
    End If
    
    If Trim(oDS_PH_PY103A.GetValue("U_YM", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "적용시작월은 필수입니다.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        PH_PY103_DataValidCheck = False
        Exit Function
    End If
    
    '// Code & Name 생성
    Call oDS_PH_PY103A.setValue("Code", 0, Trim(oDS_PH_PY103A.GetValue("U_CLTCOD", 0)) & Trim(oDS_PH_PY103A.GetValue("U_YM", 0)))
    Call oDS_PH_PY103A.setValue("NAME", 0, Trim(oDS_PH_PY103A.GetValue("U_CLTCOD", 0)) & Trim(oDS_PH_PY103A.GetValue("U_YM", 0)))
    
    '// 라인 ---------------------------
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            '//구분
            If oMat1.Columns("CSUCOD").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "코드는 필수입니다.", bmt_Short, True
                oMat1.Columns("CSUCOD").Cells(i).CLICK ct_Regular
                PH_PY103_DataValidCheck = False
                Exit Function
            End If
            '//근로자
            If oMat1.Columns("CSUNAM").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "근로자는 필수입니다.", bmt_Short, True
                oMat1.Columns("CSUNAM").Cells(i).CLICK ct_Regular
                PH_PY103_DataValidCheck = False
                Exit Function
            End If
        Next
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY103_DataValidCheck = False
        Exit Function
    End If
    
        oMat1.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY103B.Size > 1 Then oDS_PH_PY103B.RemoveRecord (oDS_PH_PY103B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY103_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY103_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY103_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY103_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY103_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY103_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("Param01").Specific.Value
    Param02 = oForm.Items("Param01").Specific.Value
    Param03 = oForm.Items("Param01").Specific.Value
    Param04 = oForm.Items("Param01").Specific.Value
    
    sQry = "SELECT 10"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY103_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY103B.InsertRecord (i)
        End If
        oDS_PH_PY103B.Offset = i
        Call oDS_PH_PY103B.setValue("U_COL01", i, oRecordSet.Fields(0).Value)
        Call oDS_PH_PY103B.setValue("U_COL02", i, oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
        ProgressBar01.Value = ProgressBar01.Value + 1
        ProgressBar01.Text = ProgressBar01.Value & "/" & oRecordSet.RecordCount & "건 조회중...!"
    Next
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY103_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY103_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY103_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY103_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY103_Validate_Error
    PH_PY103_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY103A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY103_Validate = False
        GoTo PH_PY103_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY103_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY103_Validate_Error:
    PH_PY103_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY103_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PH_PY103_ChooseFromList()

    Dim oChooseFromListCollection       As SAPbouiCOM.ChooseFromListCollection
    Dim oConditions                     As SAPbouiCOM.Conditions
    Dim oCondition                      As SAPbouiCOM.Condition
    Dim oChooseFromList                 As SAPbouiCOM.ChooseFromList
    Dim oChooseFromListCreationParams   As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit                           As SAPbouiCOM.EditText
    Dim oColumn                         As SAPbouiCOM.Column

    On Error GoTo PH_PY103_ChooseFromList_Error
    
    '//비용계정
    Set oChooseFromListCollection = oForm.ChooseFromLists
    Set oChooseFromListCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oChooseFromListCreationParams.ObjectType = lf_GLAccounts
    oChooseFromListCreationParams.uniqueID = "CFL_GLCode"
    oChooseFromListCreationParams.MultiSelection = False

    Set oChooseFromList = oChooseFromListCollection.Add(oChooseFromListCreationParams)

    Set oConditions = oChooseFromList.GetConditions()
    Set oCondition = oConditions.Add()
    
    Exit Sub
    
PH_PY103_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY103_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub






