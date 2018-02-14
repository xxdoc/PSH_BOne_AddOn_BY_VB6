VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY107.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 급상여기준일설정
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// 그리드 사용시
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_PH_PY107     As SAPbouiCOM.DataTable
'
''// 매트릭스 사용시
Public oMat1 As SAPbouiCOM.Matrix
Private oDS_PH_PY107A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY107B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY107.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY107_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY107"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY107_CreateItems
    Call PH_PY107_EnableMenus
    Call PH_PY107_SetDocument(oFromDocEntry01)
'    Call PH_PY107_FormResize
    
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

Private Function PH_PY107_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY107_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
'    '//매트릭스--------------------------------------------------------------------------------------
    Set oDS_PH_PY107A = oForm.DataSources.DBDataSources("@PH_PY107A")   '//헤더
    Set oDS_PH_PY107B = oForm.DataSources.DBDataSources("@PH_PY107B")   '//라인

    Set oMat1 = oForm.Items("Mat1").Specific       '

    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
'    '//----------------------------------------------------------------------------------------------

    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
    '//사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//기준일계산
    Set oCombo = oForm.Items("INPDAY").Specific
    oCombo.ValidValues.Add "1", "30일기준"
    oCombo.ValidValues.Add "2", "말일기준"
    oForm.Items("INPDAY").DisplayDesc = True

    '//급여지급
    Set oColumn = oMat1.Columns("PAYSEL")
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P213' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
            oRecordSet.MoveNext
        Loop
    End If
    oColumn.DisplayDesc = True
    
    '//시작월
    Set oColumn = oMat1.Columns("STRMON")
    oColumn.ValidValues.Add "0", "당월"
    oColumn.ValidValues.Add "-1", "전월"
    oColumn.ValidValues.Add "1", "익월"
    oColumn.DisplayDesc = True
    
    '//지급월
    Set oColumn = oMat1.Columns("JIGMON")
    oColumn.ValidValues.Add "0", "당월"
    oColumn.ValidValues.Add "-1", "전월"
    oColumn.ValidValues.Add "1", "익월"
    oColumn.DisplayDesc = True
    
    '//상여한도월
    Set oColumn = oMat1.Columns("BNSEMM")
    oColumn.ValidValues.Add "0", "당월"
    oColumn.ValidValues.Add "-1", "전월"
    oColumn.ValidValues.Add "1", "익월"
    oColumn.DisplayDesc = True
    
    '//급여한도월
    Set oColumn = oMat1.Columns("PAYEMM")
    oColumn.ValidValues.Add "0", "당월"
    oColumn.ValidValues.Add "-1", "전월"
    oColumn.ValidValues.Add "1", "익월"
    oColumn.DisplayDesc = True
    
    '// 시작일
    Set oColumn = oMat1.Columns("STRDAY")
    For i = 1 To 31
        oColumn.ValidValues.Add i, i & "일"
    Next i
    
    '// 지급일
    Set oColumn = oMat1.Columns("JIGDAY")
    oColumn.ValidValues.Add 0, "말일"
    For i = 1 To 31
        oColumn.ValidValues.Add i, i & "일"
    Next i
    
    '// 상여한도기준일
    Set oColumn = oMat1.Columns("BNSEDD")
    oColumn.ValidValues.Add 0, "말일"
    For i = 1 To 31
        oColumn.ValidValues.Add i, i & "일"
    Next i
    
    '// 급여한도기준일
    Set oColumn = oMat1.Columns("PAYEDD")
    oColumn.ValidValues.Add 0, "말일"
    For i = 1 To 31
        oColumn.ValidValues.Add i, i & "일"
    Next i
    
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY107_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY107_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY107_EnableMenus()
    
    On Error GoTo PH_PY107_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY107_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY107_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY107_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY107_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY107_FormItemEnabled
        Call PH_PY107_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY107_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY107_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY107_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY107_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY107_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("INPDAY").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("INPDAY").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("YM").Enabled = False
        oForm.Items("INPDAY").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY107_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY107_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY107_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY107_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUid = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
'                            Call PH_PY107_FormItemEnabled
'                            Call PH_PY107_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
'                            Call PH_PY107_FormItemEnabled
'                            Call PH_PY107_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
'                            Call PH_PY107_FormItemEnabled
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
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    If pval.ItemUid = "Mat1" And pval.ColUID = "PAYSEL" Then
                        Call PH_PY107_AddMatrixRow
                    End If
                    
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUid
                Case "Mat1"
'                    If pval.Row > 0 Then
'                        Call oMat1.SelectRow(pval.Row, True, False)
'                    End If
                End Select
                
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

                Call PH_PY107_FormItemEnabled
                Call PH_PY107_AddMatrixRow
                
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
'                Set oDS_PH_PY107A = Nothing
'                Set oDS_PH_PY107B = Nothing
                
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY107A", "Code")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY107_FormItemEnabled
                Call PH_PY107_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY107_FormItemEnabled
                Call PH_PY107_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY107_FormItemEnabled
                Call PH_PY107_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY107_FormItemEnabled
            Case "1293" '// 행삭제
                '// [MAT1 용]
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource

                    While (i <= oDS_PH_PY107B.Size - 1)
                        If oDS_PH_PY107B.GetValue("U_PAYSEL", i) = "" Then
                            oDS_PH_PY107B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend

                    For i = 0 To oDS_PH_PY107B.Size
                        Call oDS_PH_PY107B.setValue("U_LineNum", i, i + 1)
                    Next i

                    oMat1.LoadFromDataSource
                End If
                Call PH_PY107_AddMatrixRow
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

Sub PH_PY107_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY107_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1 용]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount

    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY107B.GetValue("U_PAYSEL", oRow - 1)) <> "" Then
            If oDS_PH_PY107B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY107B.InsertRecord (oRow)
            End If
            oDS_PH_PY107B.Offset = oRow
            oDS_PH_PY107B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY107B.setValue "U_PAYSEL", oRow, ""
            oDS_PH_PY107B.setValue "U_STRMON", oRow, ""
            oDS_PH_PY107B.setValue "U_STRDAY", oRow, ""
            oDS_PH_PY107B.setValue "U_JIGMON", oRow, ""
            oDS_PH_PY107B.setValue "U_JIGDAY", oRow, ""
            oDS_PH_PY107B.setValue "U_BNSEMM", oRow, ""
            oDS_PH_PY107B.setValue "U_BNSEDD", oRow, ""
            oDS_PH_PY107B.setValue "U_PAYEMM", oRow, ""
            oDS_PH_PY107B.setValue "U_PAYEDD", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY107B.Offset = oRow - 1
            oDS_PH_PY107B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY107B.setValue "U_PAYSEL", oRow - 1, ""
            oDS_PH_PY107B.setValue "U_STRMON", oRow - 1, ""
            oDS_PH_PY107B.setValue "U_STRDAY", oRow - 1, ""
            oDS_PH_PY107B.setValue "U_JIGMON", oRow - 1, ""
            oDS_PH_PY107B.setValue "U_JIGDAY", oRow - 1, ""
            oDS_PH_PY107B.setValue "U_BNSEMM", oRow - 1, ""
            oDS_PH_PY107B.setValue "U_BNSEDD", oRow - 1, ""
            oDS_PH_PY107B.setValue "U_PAYEMM", oRow - 1, ""
            oDS_PH_PY107B.setValue "U_PAYEDD", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY107B.Offset = oRow
        oDS_PH_PY107B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY107B.setValue "U_PAYSEL", oRow, ""
        oDS_PH_PY107B.setValue "U_STRMON", oRow, ""
        oDS_PH_PY107B.setValue "U_STRDAY", oRow, ""
        oDS_PH_PY107B.setValue "U_JIGMON", oRow, ""
        oDS_PH_PY107B.setValue "U_JIGDAY", oRow, ""
        oDS_PH_PY107B.setValue "U_BNSEMM", oRow, ""
        oDS_PH_PY107B.setValue "U_BNSEDD", oRow, ""
        oDS_PH_PY107B.setValue "U_PAYEMM", oRow, ""
        oDS_PH_PY107B.setValue "U_PAYEDD", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY107_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY107_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY107_FormClear()
On Error GoTo PH_PY107_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY107'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY107_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY107_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY107_DataValidCheck() As Boolean
On Error GoTo PH_PY107_DataValidCheck_Error
    
    Dim i       As Long
    Dim k       As Long
    Dim sQry    As String
    Dim Chk_Data    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    PH_PY107_DataValidCheck = False
'    '//----------------------------------------------------------------------------------
'    '//필수 체크
'    '//----------------------------------------------------------------------------------
    If Trim(oDS_PH_PY107A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        Exit Function
    End If
    
    If MDC_SetMod.ChkYearMonth(oDS_PH_PY107A.GetValue("U_YM", 0)) = False Then
        Sbo_Application.SetStatusBarMessage "적용년도는 필수입니다.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY107A.GetValue("U_INPDAY", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "월급직일수계산은 필수입니다.", bmt_Short, True
        oForm.Items("INPDAY").CLICK ct_Regular
        Exit Function
    End If
'
    Call oDS_PH_PY107A.setValue("Code", 0, Trim$(oDS_PH_PY107A.GetValue("U_CLTCOD", 0)) & Trim$(oDS_PH_PY107A.GetValue("U_YM", 0)))
    Call oDS_PH_PY107A.setValue("Name", 0, Trim$(oDS_PH_PY107A.GetValue("U_CLTCOD", 0)) & Trim$(oDS_PH_PY107A.GetValue("U_YM", 0)))
    
    If oForm.Mode = fm_ADD_MODE Then
        If MDC_SetMod.Value_ChkYn("[@PH_PY107A]", "Code", "'" & Trim$(oDS_PH_PY107A.GetValue("Code", 0)) & "'") = False Then
            Sbo_Application.StatusBar.SetText "이미 저장되어져 있는 헤더의 내용과 일치합니다", bmt_Short, smt_Error
            Exit Function
        End If
    End If
    
    '// 라인 ---------------------------
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            '//구분
            If oMat1.Columns("PAYSEL").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "지급구분은 필수입니다.", bmt_Short, True
                oMat1.Columns("PAYSEL").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            '//시작월
            If oMat1.Columns("STRMON").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "시작월은 필수입니다.", bmt_Short, True
                oMat1.Columns("STRMON").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            '//시작일
            If oMat1.Columns("STRDAY").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "시작일은 필수입니다.", bmt_Short, True
                oMat1.Columns("STRDAY").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            '//지급월
            If oMat1.Columns("JIGMON").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "지급월은 필수입니다.", bmt_Short, True
                oMat1.Columns("JIGMON").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            '//지급일
            If oMat1.Columns("JIGDAY").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "지급일은 필수입니다.", bmt_Short, True
                oMat1.Columns("JIGDAY").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            '//상여한도월
            If oMat1.Columns("BNSEMM").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "상여한도월은 필수입니다.", bmt_Short, True
                oMat1.Columns("BNSEMM").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            '//상여한도일
            If oMat1.Columns("BNSEDD").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "상여한도일은 필수입니다.", bmt_Short, True
                oMat1.Columns("BNSEDD").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            '//급여한도월
            If oMat1.Columns("PAYEMM").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "급여한도월은 필수입니다.", bmt_Short, True
                oMat1.Columns("PAYEMM").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            '//급여한도일
            If oMat1.Columns("PAYEDD").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "급여한도일은 필수입니다.", bmt_Short, True
                oMat1.Columns("PAYEDD").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            
            '//중복체크
            Chk_Data = Trim$(oDS_PH_PY107B.GetValue("U_PAYSEL", i))
            For k = i + 1 To oMat1.VisualRowCount - 2
                oDS_PH_PY107B.Offset = k
                If Trim$(Chk_Data) = Trim$(oDS_PH_PY107B.GetValue("U_PAYSEL", k)) Then
                    Sbo_Application.SetStatusBarMessage "내용이 중복입력되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error
                    oMat1.Columns("PAYSEL").Cells(k + 1).CLICK ct_Regular
                    Exit Function
                End If
            Next k
        Next
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY107_DataValidCheck = False
        Exit Function
    End If
    
    oMat1.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY107B.Size > 1 Then oDS_PH_PY107B.RemoveRecord (oDS_PH_PY107B.Size - 1)
    oMat1.LoadFromDataSource
    
    PH_PY107_DataValidCheck = True
    Set oRecordSet = Nothing
    
    Exit Function
    
PH_PY107_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY107_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY107_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY107_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY107_Validate_Error
    PH_PY107_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY107A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY107_Validate = False
        GoTo PH_PY107_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY107_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY107_Validate_Error:
    PH_PY107_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY107_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


