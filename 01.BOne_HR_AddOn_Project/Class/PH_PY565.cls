VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY565"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY565.cls
'//  Module         : PH
'//  Desc           : 연장근무자현황
'//  작성자         : NGY
'//  DATE           : 2012.12.06
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// 그리드 사용시
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_PH_PY565     As SAPbouiCOM.DataTable
'
''// 매트릭스 사용시
'Public oMat1 As SAPbouiCOM.Matrix
'Private oDS_PH_PY565A As SAPbouiCOM.DBDataSource
'Private oDS_PH_PY565B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument

    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY565.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY565_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY565"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY565_CreateItems
    Call PH_PY565_EnableMenus
    Call PH_PY565_SetDocument(oFromDocEntry01)
'    Call PH_PY565_FormResize
    
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

Private Function PH_PY565_CreateItems() As Boolean
On Error GoTo PH_PY565_CreateItems_Error

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Call oForm.Freeze(True)
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// 접속자에 따른 권한별 사업장 콤보박스세팅
    Call CLTCOD_Select(oForm, "CLTCOD")
    
    '//기준일자
    Call oForm.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
    oForm.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    oForm.DataSources.UserDataSources.Item("DocDate").VALUE = Format(Now, "YYYYMMDD")
    
    '//부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("TeamCode").Specific
    oCombo.DataBind.SetBound True, "", "TeamCode"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("TeamCode").DisplayDesc = True
    
    '//담당
    Call oForm.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("RspCode").Specific
    oCombo.DataBind.SetBound True, "", "RspCode"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("RspCode").DisplayDesc = True
    
    'G5제외
    Call oForm.DataSources.UserDataSources.Add("G5_YN", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("G5_YN").Specific
    oCheck.DataBind.SetBound True, "", "G5_YN"
    oForm.Items("G5_YN").Specific.Checked = True
    
    '야간제외
    Call oForm.DataSources.UserDataSources.Add("MidNt_YN", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("MidNt_YN").Specific
    oCheck.DataBind.SetBound True, "", "MidNt_YN"
    oForm.Items("MidNt_YN").Specific.Checked = False
    
    '결재박스제외
    Call oForm.DataSources.UserDataSources.Add("Box_YN", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("Box_YN").Specific
    oCheck.DataBind.SetBound True, "", "Box_YN"
    oForm.Items("Box_YN").Specific.Checked = False
    
    '//커서를 첫번째 ITEM으로 지정
    oForm.ActiveItem = "CLTCOD"
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY565_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY565_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY565_EnableMenus()
On Error GoTo PH_PY565_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY565_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY565_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY565_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY565_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY565_FormItemEnabled
        Call PH_PY565_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY565_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY565_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY565_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY565_FormItemEnabled()
On Error GoTo PH_PY565_FormItemEnabled_Error
    
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
    
PH_PY565_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY565_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox

    Dim oRecordSet  As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "Btn01" Then
                    Call PH_PY565_Print_Report01
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
                    Select Case pval.ItemUID
                           '//사업장이 바뀌면 부서와 담당 재설정
                           Case "CLTCOD"
                                '//부서
                                Set oCombo = oForm.Items("TeamCode").Specific
                                If oCombo.ValidValues.Count > 0 Then              '//삭제
                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                           oCombo.ValidValues.Remove i, psk_Index
                                   Next i
                                End If
                                '//현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                                
                                '//담당
                                Set oCombo = oForm.Items("RspCode").Specific
                                If oCombo.ValidValues.Count > 0 Then              '//삭제
                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                           oCombo.ValidValues.Remove i, psk_Index
                                   Next i
                                End If
                                '//현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                                
                           '//부서가 바뀌면 담당 재설정
                           Case "TeamCode"
                                '//담당은 그 부서의 담당만 표시
                                Set oCombo = oForm.Items("RspCode").Specific
                                If oCombo.ValidValues.Count > 0 Then              '//삭제
                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                           oCombo.ValidValues.Remove i, psk_Index
                                   Next i
                                End If
                                '//현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" & oForm.Items.Item("TeamCode").Specific.VALUE & "' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")

                    End Select
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
'                    Select Case pval.ItemUID
'                        Case "Code"
'                            '//사원명 찿아서 화면 표시 하기
'                            sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" & Trim(oForm.Items("Code").Specific.Value) & "'"
'                            oRecordSet.DoQuery sQry
'                            oForm.Items("CodeName").Specific.String = Trim(oRecordSet.Fields("U_FullName").Value)
'                    End Select
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
               'oMat1.LoadFromDataSource

                Call PH_PY565_FormItemEnabled
                Call PH_PY565_AddMatrixRow
                
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
'                Set oDS_PH_PY565A = Nothing
'                Set oDS_PH_PY565B = Nothing
                
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

Raise_FormItemEvent_Error:
    oForm.Freeze (False)
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormMenuEvent_Error

    Dim i As Long
    
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
                Call PH_PY565_FormItemEnabled
                Call PH_PY565_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY565_FormItemEnabled
                Call PH_PY565_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY565_FormItemEnabled
                Call PH_PY565_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY565_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY565B.Size - 1)
'                        If oDS_PH_PY565B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY565B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY565B.Size
'                        Call oDS_PH_PY565B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
                'End If
                Call PH_PY565_AddMatrixRow
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

Sub PH_PY565_AddMatrixRow()
On Error GoTo PH_PY565_AddMatrixRow_Error

    Dim oRow    As Long
    
    Call oForm.Freeze(True)
    
    '//[Mat1 용]
    'oMat1.FlushToDataSource
    'oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY565B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY565B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY565B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY565B.Offset = oRow
'            oDS_PH_PY565B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY565B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY565B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY565B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY565B.Offset = oRow - 1
'            oDS_PH_PY565B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY565B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY565B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY565B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY565B.Offset = oRow
'        oDS_PH_PY565B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY565B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY565B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY565B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY565_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY565_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY565_FormClear()
On Error GoTo PH_PY565_FormClear_Error

    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY565'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY565_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY565_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY565_DataValidCheck() As Boolean
On Error GoTo PH_PY565_DataValidCheck_Error

    PH_PY565_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    '//----------------------------------------------------------------------------------
'    '//필수 체크
'    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY565A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        PH_PY565_DataValidCheck = False
'        Exit Function
'    End If
'
'    oMat1.FlushToDataSource
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY565B.Size > 1 Then oDS_PH_PY565B.RemoveRecord (oDS_PH_PY565B.Size - 1)
'    oMat1.LoadFromDataSource
    
    PH_PY565_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
PH_PY565_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY565_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY565_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Function PH_PY565_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY565_Validate_Error

    PH_PY565_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY565A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY565_Validate = False
        GoTo PH_PY565_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY565_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY565_Validate_Error:
    PH_PY565_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY565_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY565_Print_Report01()
On Error GoTo PH_PY565_Print_Report01_Error

    Dim DocNum      As String
    Dim ErrNum      As Integer
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim sQry        As String
    
    Dim CLTCOD      As String
    Dim TeamCode    As String
    Dim RspCode     As String
    Dim DocDate     As String
    
    Dim G5_YN       As String
    Dim MidNt_YN    As String
    Dim Box_YN      As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY565_Print_Report01_Error
    End If
    
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DocDate = Trim(oForm.Items("DocDate").Specific.VALUE)
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE)
    
    If TeamCode = "" Or Null Then TeamCode = "%"
    If RspCode = "" Then RspCode = "%"
    
    If oForm.DataSources.UserDataSources.Item("G5_YN").VALUE = "Y" Then
        G5_YN = "Y"
    Else
        G5_YN = "N"
    End If
    
    If oForm.DataSources.UserDataSources.Item("MidNt_YN").VALUE = "Y" Then
        MidNt_YN = "Y"
    Else
        MidNt_YN = "N"
    End If
    
    If oForm.DataSources.UserDataSources.Item("Box_YN").VALUE = "Y" Then
        Box_YN = "Y"
    Else
        Box_YN = "N"
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY565] 연장근무자현황"
    ReportName = "PH_PY565_01.rpt"
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)

    gRpt_Formula(1) = "CLTCOD"
    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
    Call oRecordSet.DoQuery(sQry)
    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE

    gRpt_Formula(2) = "DocDate"
    gRpt_Formula_Value(2) = Format(DocDate, "####-##-##")

    '/ SubReport
    ReDim gRpt_SRptSqry(7)
    ReDim gRpt_SRptName(7)

    ReDim gRpt_SFormula(7, 1)
    ReDim gRpt_SFormula_Value(7, 1)

    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""

    gRpt_SFormula(2, 1) = ""
    gRpt_SFormula_Value(2, 1) = ""

    gRpt_SFormula(3, 1) = ""
    gRpt_SFormula_Value(3, 1) = ""
    
    gRpt_SFormula(4, 1) = ""
    gRpt_SFormula_Value(4, 1) = ""
    
    gRpt_SFormula(5, 1) = ""
    gRpt_SFormula_Value(5, 1) = ""
    
    gRpt_SFormula(6, 1) = ""
    gRpt_SFormula_Value(6, 1) = ""
    
    gRpt_SFormula(7, 1) = ""
    gRpt_SFormula_Value(7, 1) = ""

    sQry = "EXEC [PH_PY565_02] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "', '" & RspCode & "'"
    gRpt_SRptSqry(1) = sQry
    gRpt_SRptName(1) = "PH_PY565_SUB1"

    sQry = "EXEC [PH_PY565_03] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "', '" & RspCode & "'"
    gRpt_SRptSqry(2) = sQry
    gRpt_SRptName(2) = "PH_PY565_SUB2"

    sQry = "EXEC [PH_PY565_04] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "', '" & RspCode & "'"
    gRpt_SRptSqry(3) = sQry
    gRpt_SRptName(3) = "PH_PY565_SUB3"
    
    sQry = "EXEC [PH_PY565_05] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "', '" & RspCode & "'"
    gRpt_SRptSqry(4) = sQry
    gRpt_SRptName(4) = "PH_PY565_SUB4"
    
    sQry = "EXEC [PH_PY565_06] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "', '" & RspCode & "'"
    gRpt_SRptSqry(5) = sQry
    gRpt_SRptName(5) = "PH_PY565_SUB5"
    
    sQry = "EXEC [PH_PY565_07] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "', '" & RspCode & "'"
    gRpt_SRptSqry(6) = sQry
    gRpt_SRptName(6) = "PH_PY565_SUB6"
    
    '출장자 추가(2014.08.29 송명규)
    sQry = "EXEC [PH_PY565_08] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "', '" & RspCode & "'"
    gRpt_SRptSqry(7) = sQry
    gRpt_SRptName(7) = "PH_PY565_SUB7"

    '/ Procedure 실행"
    sQry = "EXEC [PH_PY565_01] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "', '" & RspCode & "','" & G5_YN & "','" & MidNt_YN & "','" & Box_YN & "'"
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY565_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V", , 1) = False Then
        'Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
        GoTo PH_PY565_Print_Report01_Error
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY565_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY565_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub









