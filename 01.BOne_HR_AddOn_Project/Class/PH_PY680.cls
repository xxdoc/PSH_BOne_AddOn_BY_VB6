VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY680"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY680.cls
'//  Module         : PH
'//  Desc           : 상벌현황
'//  작성자         : NGY
'//  DATE           : 2012.12.27
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid

''// 그리드 사용시
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_PH_PY680     As SAPbouiCOM.DataTable
'
''// 매트릭스 사용시
'Public oMat1 As SAPbouiCOM.Matrix
'Private oDS_PH_PY680A As SAPbouiCOM.DBDataSource
'Private oDS_PH_PY680B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY680.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY680_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY680"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY680_CreateItems
    Call PH_PY680_EnableMenus
    Call PH_PY680_SetDocument(oFromDocEntry01)
'    Call PH_PY680_FormResize
    
    oForm.Visible = True
    oForm.Update
    oForm.Freeze False
    
    oForm.ActiveItem = "DocDateFr"
    
    Set oXmlDoc = Nothing
    Exit Sub
    
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY680_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY680_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
    
    Set oGrid01 = oForm.Items("Grid01").Specific
    Call oForm.DataSources.DataTables.Add("DataTable")
    oGrid01.SelectionMode = ms_Auto
    
    '//사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// 접속자에 따른 사업장 선택
    oForm.DataSources.UserDataSources.Item("CLTCOD").VALUE = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
    
     '//기준일자
    Call oForm.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 10)
    oForm.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
    oForm.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYYMMDD")

    Call oForm.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 10)
    oForm.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
    oForm.DataSources.UserDataSources.Item("DocDateTo").VALUE = Format(Now, "YYYYMMDD")
    
    '재직구분
    Call oForm.DataSources.UserDataSources.Add("HRStatus", dt_SHORT_TEXT, 10)
    Call oForm.Items("HRStatus").Specific.DataBind.SetBound(True, "", "HRStatus")
    
    '상벌종류
    Call oForm.DataSources.UserDataSources.Add("Type", dt_SHORT_TEXT, 10)
    Call oForm.Items("Type").Specific.DataBind.SetBound(True, "", "Type")
    
    '출력구분
    Call oForm.DataSources.UserDataSources.Add("PrtType", dt_SHORT_TEXT, 10)
    Call oForm.Items("PrtType").Specific.DataBind.SetBound(True, "", "PrtType")
    
    '//////////콤보박스 세팅//////////
    '상벌구분
    Set oCombo = oForm.Items("Div").Specific
    oCombo.ValidValues.Add "%", "전체"
    oCombo.ValidValues.Add "1", "상"
    oCombo.ValidValues.Add "2", "벌"
    oCombo.Select 0, psk_Index
    
    '출력구분
    Set oCombo = oForm.Items("PrtType").Specific
    Call oCombo.ValidValues.Add("1", "부서별")
    Call oCombo.ValidValues.Add("2", "년도별")
    Call oCombo.Select(0, psk_Index)
    
    '재직구분
    sQry = "           SELECT      statusID,"
    sQry = sQry & "                name"
    sQry = sQry & " FROM       [OHST]"
    sQry = sQry & " ORDER BY  statusID"
    
    Call oForm.Items("HRStatus").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm.Items("HRStatus").Specific, sQry, "", False, False)
    Call oForm.Items("HRStatus").Specific.Select(0, psk_Index)
    
    '상벌종류
    sQry = "           SELECT      U_Code,"
    sQry = sQry & "                U_CodeNm"
    sQry = sQry & " FROM       [@PS_HR200L]"
    sQry = sQry & " WHERE      Code = 'P130'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    
    Call oForm.Items("Type").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm.Items("Type").Specific, sQry, "", False, False)
    Call oForm.Items("Type").Specific.Select(0, psk_Index)
    
    '//////////콤보박스 세팅//////////
    
    '//커서를 첫번째 ITEM으로 지정
'    oForm.ActiveItem = "CLTCOD"
        
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY680_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY680_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY680_EnableMenus()
    
    On Error GoTo PH_PY680_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY680_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY680_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY680_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY680_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY680_FormItemEnabled
        Call PH_PY680_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY680_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY680_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY680_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY680_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY680_FormItemEnabled_Error
    
    
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
    
PH_PY680_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY680_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUID = "Btn01" Then
                    Call PH_PY680_Print_Report01
                ElseIf pval.ItemUID = "BtnSearch" Then
                    Call PH_PY680_MTX01
                End If
            ElseIf pval.BeforeAction = False Then


            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat01", "Grid01"
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
'                    Select Case pval.ItemUID
'                           '//사업장이 바뀌면 부서와 담당 재설정
'                           Case "CLTCOD"
'                                '//부서
'                                Set oCombo = oForm.Items("TeamCode").Specific
'                                If oCombo.ValidValues.Count > 0 Then              '//삭제
'                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                                           oCombo.ValidValues.Remove i, psk_Index
'                                   Next i
'                                End If
'                                '//현재 사업장으로 다시 Qry
'                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
'                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
'
'                                '//담당
'                                Set oCombo = oForm.Items("RspCode").Specific
'                                If oCombo.ValidValues.Count > 0 Then              '//삭제
'                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                                           oCombo.ValidValues.Remove i, psk_Index
'                                   Next i
'                                End If
'                                '//현재 사업장으로 다시 Qry
'                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
'                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
'
'                           '//부서가 바뀌면 담당 재설정
'                           Case "TeamCode"
'                                '//담당은 그 부서의 담당만 표시
'                                Set oCombo = oForm.Items("RspCode").Specific
'                                If oCombo.ValidValues.Count > 0 Then              '//삭제
'                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                                           oCombo.ValidValues.Remove i, psk_Index
'                                   Next i
'                                End If
'                                '//현재 사업장으로 다시 Qry
'                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" & oForm.Items.Item("TeamCode").Specific.Value & "' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
'                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
'                           '//근무형태가 바뀌면 근무조 재설정
'                           Case "ShiftDat"
'                                Set oCombo = oForm.Items("GNMUJO").Specific
'                                If oCombo.ValidValues.Count > 0 Then              '//삭제
'                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                                           oCombo.ValidValues.Remove i, psk_Index
'                                   Next i
'                                End If
'                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y' AND U_Char1 = '" & oForm.Items.Item("ShiftDat").Specific.Value & "'"
'                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
'
'                    End Select
                End If
            End If
            
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat01"
'                    If pval.Row > 0 Then
'                        Call oMat1.SelectRow(pval.Row, True, False)
'                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Mat01", "Grid01"
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
                    Select Case pval.ItemUID
                        Case "Code"
                            '//사원명 찿아서 화면 표시 하기
                            sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" & Trim(oForm.Items("Code").Specific.VALUE) & "'"
                            oRecordSet.DoQuery sQry
                            oForm.Items("CodeName").Specific.String = Trim(oRecordSet.Fields("U_FullName").VALUE)
                    End Select
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
               'oMat1.LoadFromDataSource

                Call PH_PY680_FormItemEnabled
                Call PH_PY680_AddMatrixRow
                
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
'                Set oDS_PH_PY680A = Nothing
'                Set oDS_PH_PY680B = Nothing
                
                'Set oMat1 = Nothing
                'Set oGrid1 = Nothing
                
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
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY680_FormItemEnabled
                Call PH_PY680_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY680_FormItemEnabled
                Call PH_PY680_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY680_FormItemEnabled
                Call PH_PY680_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY680_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY680B.Size - 1)
'                        If oDS_PH_PY680B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY680B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY680B.Size
'                        Call oDS_PH_PY680B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
                'End If
                Call PH_PY680_AddMatrixRow
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

Sub PH_PY680_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY680_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1 용]
    'oMat1.FlushToDataSource
    'oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY680B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY680B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY680B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY680B.Offset = oRow
'            oDS_PH_PY680B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY680B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY680B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY680B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY680B.Offset = oRow - 1
'            oDS_PH_PY680B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY680B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY680B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY680B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY680B.Offset = oRow
'        oDS_PH_PY680B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY680B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY680B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY680B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY680_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY680_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY680_FormClear()
On Error GoTo PH_PY680_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY680'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY680_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY680_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY680_DataValidCheck() As Boolean
On Error GoTo PH_PY680_DataValidCheck_Error
    PH_PY680_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    '//----------------------------------------------------------------------------------
'    '//필수 체크
'    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY680A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        PH_PY680_DataValidCheck = False
'        Exit Function
'    End If
'
'    oMat1.FlushToDataSource
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY680B.Size > 1 Then oDS_PH_PY680B.RemoveRecord (oDS_PH_PY680B.Size - 1)
'    oMat1.LoadFromDataSource
    
    PH_PY680_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY680_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY680_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY680_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY680_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY680_Validate_Error
    PH_PY680_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY680A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY680_Validate = False
        GoTo PH_PY680_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY680_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY680_Validate_Error:
    PH_PY680_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY680_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Public Sub PH_PY680_MTX01()
'******************************************************************************
'Function ID : PH_PY680_MTX01()
'해당모듈 : PH_PY680
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY680_MTX01_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String
    Dim DocDateFr As String
    Dim DocDateTo As String
    Dim Div As String
    Dim Code As String
    Dim HRStatus As String
    Dim HRType As String
    Dim PrtType As String
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DocDateFr = Trim(oForm.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm.Items("DocDateTo").Specific.VALUE)
    Div = Trim(oForm.Items("Div").Specific.VALUE)
    Code = Trim(oForm.Items("Code").Specific.VALUE)
    HRStatus = Trim(oForm.Items("HRStatus").Specific.VALUE)
    HRType = Trim(oForm.Items("Type").Specific.VALUE)
    PrtType = Trim(oForm.Items("PrtType").Specific.VALUE)
        
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", RecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "         EXEC [PH_PY680_02] '"
    sQry = sQry & CLTCOD & "', '"
    sQry = sQry & DocDateFr & "',  '"
    sQry = sQry & DocDateTo & "', '"
    sQry = sQry & Div & "', '"
    sQry = sQry & Code & "', '"
    sQry = sQry & HRStatus & "', '"
    sQry = sQry & HRType & "'"
 
    oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable")
    Call oGrid01.DataTable.Clear
    Call oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry)
    
    ProgBar01.VALUE = 100
'    ProgBar01.Text = "비용 집계중...!"
    ProgBar01.Stop
    
    oGrid01.Columns(11).RightJustified = True
    oGrid01.Columns(12).RightJustified = True
    oGrid01.Columns(13).RightJustified = True

    If oGrid01.Rows.Count = 0 Then
        ErrNum = 1
        GoTo PH_PY680_MTX01_Error
    End If
        
    oGrid01.AutoResizeColumns
    oForm.Update
    
    Call oForm.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set RecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY680_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set RecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY680_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Sub PH_PY680_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY680_Print_Report01_Error
    
    Dim CLTCOD           As String
    Dim DocDateFr        As String
    Dim DocDateTo        As String
    Dim Div              As String
    Dim Code             As String
    Dim HRStatus As String
    Dim HRType As String
    Dim PrtType As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DocDateFr = Trim(oForm.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm.Items("DocDateTo").Specific.VALUE)
    Div = Trim(oForm.Items("Div").Specific.VALUE)
    Code = Trim(oForm.Items("Code").Specific.VALUE)
    HRStatus = Trim(oForm.Items("HRStatus").Specific.VALUE)
    HRType = Trim(oForm.Items("Type").Specific.VALUE)
    PrtType = Trim(oForm.Items("PrtType").Specific.VALUE)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY680_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY680] 상벌현황"
    If PrtType = "1" Then
        ReportName = "PH_PY680_01.rpt"
    Else
        ReportName = "PH_PY680_02.rpt"
    End If
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(3)
    ReDim gRpt_Formula_Value(3)
    
    gRpt_Formula(1) = "CLTCOD"
    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
    Call oRecordSet.DoQuery(sQry)
    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
    
    gRpt_Formula(2) = "DocDateFr"
    gRpt_Formula_Value(2) = Format(DocDateFr, "####-##-##")
    
    gRpt_Formula(3) = "DocDateTo"
    gRpt_Formula_Value(3) = Format(DocDateTo, "####-##-##")
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '/ Procedure 실행"
    If PrtType = "1" Then
        sQry = "         EXEC [PH_PY680_01] '"
        sQry = sQry & CLTCOD & "', '"
        sQry = sQry & DocDateFr & "',  '"
        sQry = sQry & DocDateTo & "', '"
        sQry = sQry & Div & "', '"
        sQry = sQry & Code & "', '"
        sQry = sQry & HRStatus & "', '"
        sQry = sQry & HRType & "'"
    Else
        sQry = "         EXEC [PH_PY680_03] '"
        sQry = sQry & CLTCOD & "', '"
        sQry = sQry & DocDateFr & "',  '"
        sQry = sQry & DocDateTo & "', '"
        sQry = sQry & Div & "', '"
        sQry = sQry & Code & "', '"
        sQry = sQry & HRStatus & "', '"
        sQry = sQry & HRType & "'"
    End If
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY680_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "Y", "V", , 1) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
'    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY680_Print_Report01_Error:
    If ErrNum = 1 Then
        Set ProgBar01 = Nothing
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY680_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub

















