VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY110.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 개인별상여율등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY110A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY110B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY110.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY110_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY110"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY110_CreateItems
    Call PH_PY110_EnableMenus
    Call PH_PY110_SetDocument(oFromDocEntry01)
'    Call PH_PY110_FormResize
    
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

Private Function PH_PY110_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY110_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY110A = oForm.DataSources.DBDataSources("@PH_PY110A")   '//헤더
    Set oDS_PH_PY110B = oForm.DataSources.DBDataSources("@PH_PY110B")   '//라인

    Set oMat1 = oForm.Items("Mat1").Specific       '

    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns

    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
    '//사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    
     

    '// 지급구분
    Set oCombo = oForm.Items("JOBGBN").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.Select "1", psk_ByValue
    oForm.Items("JOBGBN").DisplayDesc = True
    
    '// 지급대상
    Set oCombo = oForm.Items("JOBTRG").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.Select "1", psk_ByValue
    oForm.Items("JOBTRG").DisplayDesc = True


    '// 라인 ------------------------------------------------------------------------------------------------
    '// 부서명
'    Set oColumn = oMat1.Columns("DPTNAM")
'    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='1' AND U_UseYN= 'Y'"
'    oRecordset.DoQuery sQry
'    If oRecordset.RecordCount > 0 Then
'        Do Until oRecordset.EOF
'            oColumn.ValidValues.Add Trim$(oRecordset.Fields(0).Value), Trim$(oRecordset.Fields(1).Value)
'            oRecordset.MoveNext
'        Loop
'    End If
'    oColumn.DisplayDesc = True

    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY110_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY110_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function




Private Sub PH_PY110_EnableMenus()
    
    On Error GoTo PH_PY110_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY110_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY110_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY110_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY110_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY110_FormItemEnabled
        Call PH_PY110_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY110_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY110_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY110_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY110_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oConditions As SAPbouiCOM.Conditions
    
    On Error GoTo PH_PY110_FormItemEnabled_Error
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("JOBGBN").Enabled = True
        oForm.Items("JOBTRG").Enabled = True
        oForm.Items("Btn_Excel").Enabled = True
        oForm.Items("Btn_Create").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        '/ 귀속년월
        oForm.Items("YM").Specific.Value = Format$(Now, "YYYYMM")
        
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("JOBGBN").Enabled = True
        oForm.Items("JOBTRG").Enabled = True
        oForm.Items("Btn_Excel").Enabled = False
        oForm.Items("Btn_Create").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("YM").Enabled = False
        oForm.Items("JOBGBN").Enabled = False
        oForm.Items("JOBTRG").Enabled = False
        oForm.Items("Btn_Excel").Enabled = True
        oForm.Items("Btn_Create").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY110_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY110_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    If oForm.Mode = fm_ADD_MODE Or fm_UPDATE_MODE Then
                        If PH_PY110_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    End If
                End If
                If pval.ItemUID = "Btn_Excel" Then
                    Call PH_PY110_Excel_Upload
                    Call PH_PY110_AddMatrixRow
'                    If Execution = False Then
'                        BubbleEvent = False
'                    End If
                End If
                If pval.ItemUID = "Btn_Create" Then
                    Call MSTCOD_Setting
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If pval.ActionSuccess = True Then
                        Call PH_PY110_FormItemEnabled
                        
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True And pval.ItemUID = "Mat1" And pval.ColUID = "MSTCOD" And pval.CharPressed = 9 Then
                If oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.Value = "" Then
                    oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                Else
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.Value + "'") = True Then
                        oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
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

            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat1"
                    If pval.Row > 0 Then
                        '//Call oMat1.SelectRow(pval.Row, True, False)

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
                    If pval.ItemUID = "Mat1" And pval.ColUID = "MSTCOD" Then
                        If oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.Value <> "" Then
                            oMat1.Columns("MSTNAM").Cells(pval.Row).Specific.Value = MDC_SetMod.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.Value & "'")
                            oMat1.Columns("DPTCOD").Cells(pval.Row).Specific.Value = MDC_SetMod.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.Value & "'")
                            oMat1.Columns("DPTNAM").Cells(pval.Row).Specific.Value = MDC_SetMod.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat1.Columns("DPTCOD").Cells(pval.Row).Specific.Value & "'", " AND Code = '1'")
                        End If
                        Call PH_PY110_AddMatrixRow
                        oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK ct_Regular
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY110_FormItemEnabled
                Call PH_PY110_AddMatrixRow

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
                Set oDS_PH_PY110A = Nothing
                Set oDS_PH_PY110B = Nothing

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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY110A", "Code")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY110_FormItemEnabled
                Call PH_PY110_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY110_FormItemEnabled
                Call PH_PY110_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY110_FormItemEnabled
                Call PH_PY110_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY110_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource

                    While (i <= oDS_PH_PY110B.Size - 1)
                        If oDS_PH_PY110B.GetValue("U_MSTCOD", i) = "" Then
                            oDS_PH_PY110B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend

                    For i = 0 To oDS_PH_PY110B.Size
                        Call oDS_PH_PY110B.setValue("U_LineNum", i, i + 1)
                    Next i

                    oMat1.LoadFromDataSource
                End If
                Call PH_PY110_AddMatrixRow
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

Sub PH_PY110_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY110_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 용]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount

    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY110B.GetValue("U_MSTCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY110B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY110B.InsertRecord (oRow)
            End If
            oDS_PH_PY110B.Offset = oRow
            oDS_PH_PY110B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY110B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY110B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY110B.setValue "U_DPTCOD", oRow, ""
            oDS_PH_PY110B.setValue "U_DPTCOD", oRow, ""
            oDS_PH_PY110B.setValue "U_BNSRAT", oRow, 0
            oDS_PH_PY110B.setValue "U_BNSAMT", oRow, 0
            oDS_PH_PY110B.setValue "U_REMARK", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY110B.Offset = oRow - 1
            oDS_PH_PY110B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY110B.setValue "U_MSTCOD", oRow - 1, ""
            oDS_PH_PY110B.setValue "U_MSTNAM", oRow - 1, ""
            oDS_PH_PY110B.setValue "U_DPTCOD", oRow - 1, ""
            oDS_PH_PY110B.setValue "U_DPTCOD", oRow - 1, ""
            oDS_PH_PY110B.setValue "U_BNSRAT", oRow - 1, 0
            oDS_PH_PY110B.setValue "U_BNSAMT", oRow - 1, 0
            oDS_PH_PY110B.setValue "U_REMARK", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY110B.Offset = oRow
        oDS_PH_PY110B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY110B.setValue "U_MSTCOD", oRow, ""
        oDS_PH_PY110B.setValue "U_MSTNAM", oRow, ""
        oDS_PH_PY110B.setValue "U_DPTCOD", oRow, ""
        oDS_PH_PY110B.setValue "U_DPTCOD", oRow, ""
        oDS_PH_PY110B.setValue "U_BNSRAT", oRow, 0
        oDS_PH_PY110B.setValue "U_BNSAMT", oRow, 0
        oDS_PH_PY110B.setValue "U_REMARK", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY110_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY110_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY110_FormClear()
On Error GoTo PH_PY110_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY110'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY110_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY110_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY110_DataValidCheck() As Boolean
On Error GoTo PH_PY110_DataValidCheck_Error
    PH_PY110_DataValidCheck = False
    Dim i           As Long
    Dim k           As Long
    Dim sCode       As String
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//필수체크
    '//----------------------------------------------------------------------------------
    '// 헤더
    If Trim(oDS_PH_PY110A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY110A.GetValue("U_YM", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "귀속년월은 필수입니다.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY110A.GetValue("U_JOBGBN", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "지급구분은 필수입니다.", bmt_Short, True
        oForm.Items("JOBGBN").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY110A.GetValue("U_JOBTRG", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "지급대상은 필수입니다.", bmt_Short, True
        oForm.Items("JOBTRG").CLICK ct_Regular
        Exit Function
    End If
    
    oMat1.FlushToDataSource
    '// 라인 ---------------------------
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            '//구분
            If oMat1.Columns("MSTCOD").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "사번은 필수입니다.", bmt_Short, True
                oMat1.Columns("MSTCOD").Cells(i).CLICK ct_Regular
                Exit Function
            End If
           
            '//중복체크
            For k = i + 1 To oMat1.VisualRowCount - 2
                If Trim$(oDS_PH_PY110B.GetValue("U_MSTCOD", i)) = Trim$(oDS_PH_PY110B.GetValue("U_MSTCOD", k)) Then
                    Sbo_Application.SetStatusBarMessage "내용이 중복입력되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error
                    oMat1.Columns("MSTCOD").Cells(k + 1).CLICK ct_Regular
                    Exit Function
                End If
            Next k
        Next
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY110_DataValidCheck = False
        Exit Function
    End If
    
    sCode = Trim$(oDS_PH_PY110A.GetValue("U_CLTCOD", 0)) & Trim$(oDS_PH_PY110A.GetValue("U_YM", 0)) _
            & Trim$(oDS_PH_PY110A.GetValue("U_JOBGBN", 0)) & Trim$(oDS_PH_PY110A.GetValue("U_JOBTRG", 0))
    
    oDS_PH_PY110A.setValue "Code", 0, sCode
    oDS_PH_PY110A.setValue "Name", 0, sCode
    
'    oMat1.FlushToDataSource
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY110B.Size > 1 Then oDS_PH_PY110B.RemoveRecord (oDS_PH_PY110B.Size - 1)
'    oMat1.LoadFromDataSource

    '//데이터 중복 체크
    If MDC_SetMod.Get_ReData("Code", "Code", "[@PH_PY110A]", "'" & "sCode" & "'") <> "" Then
        Sbo_Application.SetStatusBarMessage "기존에 데이터가 존재합니다.", bmt_Short, True
        PH_PY110_DataValidCheck = False
        Exit Function
    End If
    
    PH_PY110_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY110_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY110_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY110_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY110_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY110_Validate_Error
    PH_PY110_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY110A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY110_Validate = False
        GoTo PH_PY110_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY110_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY110_Validate_Error:
    PH_PY110_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY110_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PH_PY110_Excel_Upload()
    
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim CheckLine   As Long
    Dim CheckYN     As Boolean
    
    Dim oBNSRAT     As Double   '/적용율
    Dim oBNSAMT     As Double   '/적용금액
    
    Dim sFile       As String
    Dim sQry        As String
    
    Dim xl          As Excel.Application
    Dim xlwb        As Excel.Workbook
    Dim xlsh        As Excel.Worksheet
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oForm       As SAPbouiCOM.Form
    
    Dim V_StatusCnt&, oProValue&, TOTCNT&, tRow&  '//progbar
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Err
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    Set oForm = Sbo_Application.Forms.ActiveForm
    
'    oForm.Freeze (True)
    
    sFile = FileListBoxForm.OpenDialog(FileListBoxForm, "*.xls", "파일선택", "C:\")
    
    If sFile = "" Then
        Sbo_Application.StatusBar.SetText "파일을 선택해 주세요.", bmt_Short, smt_Error
        Exit Sub
    Else
        If Right(Replace(sFile, Chr$(0), ""), 3) <> "xls" And Right(Replace(sFile, Chr$(0), ""), 4) <> "xlsx" Then
            Sbo_Application.StatusBar.SetText "엑셀파일이 아닙니다.", bmt_Short, smt_Error
            Exit Sub
        End If
    End If
    
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    
    '엑셀 Object 연결
    Set xl = CreateObject("excel.application")
    Set xlwb = xl.Workbooks.Open(sFile, , True)
    Set xlsh = xlwb.Worksheets("상여율등록")
    
    If xlsh.Cells(1, 1).Value <> "사번" Then
        Sbo_Application.StatusBar.SetText "A열 첫번째 행 타이틀은 사번", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 2).Value <> "성명" Then
        Sbo_Application.StatusBar.SetText "B열 두번째 행 타이틀은 성명", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 3).Value <> "적용률" Then
        Sbo_Application.StatusBar.SetText "C열 세번째 행 타이틀은 적용률", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 4).Value <> "적용금액" Then
        Sbo_Application.StatusBar.SetText "D열 세번째 행 타이틀은 적용금액", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 5).Value <> "비고" Then
        Sbo_Application.StatusBar.SetText "E열 세번째 행 타이틀은 비고", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    '프로그레스 바    ///////////////////////////////////////
    Set oProgBar = Sbo_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, False)
    
    '최대값 구하기
    TOTCNT = xlsh.UsedRange.Rows.Count
   
    V_StatusCnt = Round(TOTCNT / 50, 0)
    oProValue = 1
    tRow = 1
    '/////////////////////////////////////////////////////
    
    
    CheckLine = 0
    For i = 2 To xlsh.UsedRange.Rows.Count
        
        
        '//사번 존재 여부 체크
        If Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim(xlsh.Cells(i, 1)) & "'") = True Then
            Sbo_Application.StatusBar.SetText i & "행의 [" & xlsh.Cells(i, 1) & " ] 사번이 존재하지 않습니다.", bmt_Short, smt_Error
            GoTo Err
        Else
            CheckYN = False
            For j = 0 To oDS_PH_PY110B.Size - 1
                If Trim(xlsh.Cells(i, 1)) = Trim(oDS_PH_PY110B.GetValue("U_MSTCOD", j)) Then
                    CheckYN = True
                    CheckLine = j
                End If
            Next j
            
            '//마지막행 제거
            If Trim(oDS_PH_PY110B.GetValue("U_MSTCOD", oDS_PH_PY110B.Size - 1)) = "" Then
                oDS_PH_PY110B.RemoveRecord (oDS_PH_PY110B.Size - 1)
            End If
            
            '// 사원마스터에서 사번에 대한 정보 가져오기
            sQry = "select U_FullName, U_TeamCode, U_CodeNm "
            sQry = sQry & " FROM [@PH_PY001A] T0 INNER JOIN [@PS_HR200L] T1 ON T0.U_TeamCode = T1.U_Code"
            sQry = sQry & " WHERE T0.Code='" & Trim(xlsh.Cells(i, 1)) & "'"
            oRecordSet.DoQuery sQry
            
            '//Mat1에 업로드엑셀 사번이 존재 유무 CheckYN
            If CheckYN = True Then
                oDS_PH_PY110B.setValue "U_MSTNAM", CheckLine, Trim$(oRecordSet.Fields(0).Value)
                oDS_PH_PY110B.setValue "U_DPTCOD", CheckLine, Trim$(oRecordSet.Fields(1).Value)
                oDS_PH_PY110B.setValue "U_DPTNAM", CheckLine, Trim$(oRecordSet.Fields(2).Value)
                oDS_PH_PY110B.setValue "U_BNSRAT", CheckLine, xlsh.Cells(i, 3)
                oDS_PH_PY110B.setValue "U_BNSAMT", CheckLine, xlsh.Cells(i, 4)
                oDS_PH_PY110B.setValue "U_REMARK", CheckLine, xlsh.Cells(i, 5)
            Else        '//새로 사번 추가
'                If oDS_PH_PY110B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY110B.InsertRecord (oDS_PH_PY110B.Size)
'                End If
                oDS_PH_PY110B.Offset = oDS_PH_PY110B.Size - 1
                oDS_PH_PY110B.setValue "U_MSTCOD", oDS_PH_PY110B.Size - 1, xlsh.Cells(i, 1)
                oDS_PH_PY110B.setValue "U_MSTNAM", oDS_PH_PY110B.Size - 1, Trim$(oRecordSet.Fields(0).Value)
                oDS_PH_PY110B.setValue "U_DPTCOD", oDS_PH_PY110B.Size - 1, Trim$(oRecordSet.Fields(1).Value)
                oDS_PH_PY110B.setValue "U_DPTNAM", oDS_PH_PY110B.Size - 1, Trim$(oRecordSet.Fields(2).Value)
                oDS_PH_PY110B.setValue "U_BNSRAT", oDS_PH_PY110B.Size - 1, xlsh.Cells(i, 3)
                oDS_PH_PY110B.setValue "U_BNSAMT", oDS_PH_PY110B.Size - 1, xlsh.Cells(i, 4)
                oDS_PH_PY110B.setValue "U_REMARK", oDS_PH_PY110B.Size - 1, xlsh.Cells(i, 5)
            End If
            
            If (TOTCNT > 50 And tRow = oProValue * V_StatusCnt) Or TOTCNT <= 50 Then
                oProgBar.Text = tRow & "/ " & TOTCNT & " 건 처리중...!"
                oProValue = oProValue + 1
                oProgBar.Value = oProValue
            End If
            tRow = tRow + 1
        End If
    Next i
    
    '//라인번호 재정의
    For i = 0 To oDS_PH_PY110B.Size - 1
        Call oDS_PH_PY110B.setValue("U_LineNum", i, i + 1)
    Next i
    
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    
    oProgBar.Stop
    
    Sbo_Application.StatusBar.SetText "엑셀을 불러왔습니다.", bmt_Short, smt_Success


    '액셀개체 닫음
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    '진행바 초기화
    
    Set oProgBar = Nothing
'    oForm.Freeze (False)
    Exit Sub

Err:
'    oForm.Freeze (False)
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
End Sub

Private Sub MSTCOD_Setting()
On Error GoTo Error_Message
    Dim i           As Long
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim iRow        As Long
    Dim MatRowCnt   As Long:    Dim oRow  As Long
    Dim ErrNum      As Integer
    
    Dim CLTCOD      As String   '//사업장
    Dim JOBYMM      As String   '//귀속년월
    Dim JOBTRG      As String   '//지급대상
    
    Dim V_StatusCnt&, oProValue&, TOTCNT&, tRow&  '//progbar
    
    Dim CheckValue  As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    oForm.Freeze (True)
    
    CLTCOD = Trim$(oDS_PH_PY110A.GetValue("U_CLTCOD", 0))
    JOBYMM = Trim$(oDS_PH_PY110A.GetValue("U_YM", 0))
    JOBTRG = Trim$(oDS_PH_PY110A.GetValue("U_JOBTRG", 0))
    If JOBYMM = "" Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    sQry = "EXEC PH_PY110 '" & CLTCOD & "','" & JOBYMM & "','" & JOBTRG & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    
    Set oProgBar = Sbo_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, False)
    
    '최대값 구하기 ///////////////////////////////////////
    TOTCNT = oRecordSet.RecordCount
   
    V_StatusCnt = Round(TOTCNT / 50, 0)
    oProValue = 1
    tRow = 1
    '/////////////////////////////////////////////////////

    
    oDS_PH_PY110B.RemoveRecord (oDS_PH_PY110B.Size - 1)
    
    MatRowCnt = oMat1.VisualRowCount - 2
    oRow = oMat1.VisualRowCount - 1
    Do Until oRecordSet.EOF
        CheckValue = "N"
        For iRow = 0 To MatRowCnt
'            oDS_PH_PY110B.Offset = iRow
            If Trim$(oRecordSet.Fields("MSTCOD").Value) = Trim$(oDS_PH_PY110B.GetValue("U_MSTCOD", iRow)) Then
                CheckValue = "Y"
                Exit For
            End If
        Next iRow
        
        If CheckValue = "N" Then
            oDS_PH_PY110B.InsertRecord oRow
            oDS_PH_PY110B.Offset = oRow
            oDS_PH_PY110B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY110B.setValue "U_MSTCOD", oRow, Trim(oRecordSet.Fields("MSTCOD").Value)    '사원번호
            oDS_PH_PY110B.setValue "U_MSTNAM", oRow, Trim(oRecordSet.Fields("MSTNAM").Value)    '성명
            oDS_PH_PY110B.setValue "U_DPTCOD", oRow, Trim(oRecordSet.Fields("DPTCOD").Value)    '부서
            oDS_PH_PY110B.setValue "U_DPTNAM", oRow, Trim(oRecordSet.Fields("DPTNAM").Value)    '부서명
            oDS_PH_PY110B.setValue "U_BNSRAT", oRow, "100"                                '적용율(%)
            oDS_PH_PY110B.setValue "U_BNSAMT", oRow, "0"                                  '적용금액
            oDS_PH_PY110B.setValue "U_REMARK", oRow, ""
            oRow = oRow + 1
        End If
        
        tRow = tRow + 1
        If (TOTCNT > 50 And tRow = oProValue * V_StatusCnt) Or TOTCNT <= 50 Then
            oProgBar.Text = tRow & "/ " & TOTCNT & " 건 처리중...!"
            oProValue = oProValue + 1
            oProgBar.Value = oProValue
        End If
        
        oRecordSet.MoveNext
    Loop
    
    Sbo_Application.StatusBar.SetText "작업이 완료되었습니다.", bmt_Short, smt_Success
     
    If oRow = MatRowCnt + 1 Then
        ErrNum = 3
        GoTo Error_Message
    End If
'    oForm.Freeze (False)
    
   
    
    oMat1.LoadFromDataSource
    Call PH_PY110_AddMatrixRow
    
    oProgBar.Stop
    Set oProgBar = Nothing
    Set oRecordSet = Nothing
    
    Exit Sub
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "귀속년월은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "생성할 대상자가 없습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "이미 모든 대상자가 생성되어 있어 추가생성된 대상자가 없습니다.", bmt_Short, smt_Success
    Else
        Sbo_Application.StatusBar.SetText "MSTCOD_Setting Error : " & Err.Description, bmt_Short, smt_Error
    End If
    
'    oForm.Freeze (False)
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    Set oRecordSet = Nothing
End Sub
