VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY678"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PH_PY678.cls
'//  Module             : 인사관리>기타관리
'//  Desc               : 당직근무자 일괄 등록
'//  FormType           : PH_PY678
'//  Create Date(Start) : 2013.10.30
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm           As SAPbouiCOM.Form

Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PH_PY678B  As SAPbouiCOM.DBDataSource '라인

Private oLast_Item_UID As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oSeq&
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i         As Long
    Dim oInnerXml As String
    Dim oXmlDoc   As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY678.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY678_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01, "PH_PY678" '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
'    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm.Freeze True
    Call PH_PY678_CreateItems
    Call PH_PY678_ComboBox_Setting
    Call PH_PY678_CF_ChooseFromList
    Call PH_PY678_EnableMenus
    Call PH_PY678_SetDocument(oFromDocEntry01)
    Call PH_PY678_FormResize

    oForm.Items("FrDt").Specific.VALUE = Format(Date, "yyyyMM01")
    oForm.Items("ToDt").Specific.VALUE = Format(Date, "yyyyMMdd")

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

'/아이템 변경 이벤트
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error

    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            Call Raise_EVENT_ITEM_PRESSED(FormUID, pval, BubbleEvent)
        Case et_KEY_DOWN: '//2
            Call Raise_EVENT_KEY_DOWN(FormUID, pval, BubbleEvent)
        Case et_COMBO_SELECT: '//5
            Call Raise_EVENT_COMBO_SELECT(FormUID, pval, BubbleEvent)
        Case et_CLICK: '//6
            Call Raise_EVENT_CLICK(FormUID, pval, BubbleEvent)
        Case et_DOUBLE_CLICK: '//7
            Call Raise_EVENT_DOUBLE_CLICK(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LINK_PRESSED '//8
            Call Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, pval, BubbleEvent)
        Case et_VALIDATE: '//10
            Call Raise_EVENT_VALIDATE(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LOAD: '//11
            Call Raise_EVENT_MATRIX_LOAD(FormUID, pval, BubbleEvent)
        Case et_FORM_ACTIVATE: '//18
            '//et_FORM_ACTIVATE
        Case et_FORM_DEACTIVATE: '//19
            '//et_FORM_DEACTIVATE
        Case et_FORM_RESIZE '//20
            Call Raise_EVENT_RESIZE(FormUID, pval, BubbleEvent)
        Case et_CHOOSE_FROM_LIST '//27
            Call Raise_EVENT_CHOOSE_FROM_LIST(FormUID, pval, BubbleEvent)
        Case et_GOT_FOCUS: '//3
            Call Raise_EVENT_GOT_FOCUS(FormUID, pval, BubbleEvent)
        Case et_LOST_FOCUS: '//4
            '//et_LOST_FOCUS
        Case et_FORM_UNLOAD: '//17
            Call Raise_EVENT_FORM_UNLOAD(FormUID, pval, BubbleEvent)
    End Select
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormMenuEvent_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PH_PY678_FormItemEnabled
                Call PH_PY678_AddMatrixRow
            Case "1282": '추가
                Call PH_PY678_FormItemEnabled
                Call PH_PY678_AddMatrixRow
            
'                oMat01.Clear
'                oDS_PH_PY032A.Clear

'                Call PH_PY032_LoadCaption
'                Call PH_PY032_FormItemEnabled
                '//Call PH_PY032_FormItemEnabled '//UDO방식
                '//Call PH_PY032_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY032_FormItemEnabled
        End Select
    End If
    Exit Sub
Raise_FormMenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then '//BeforeAction = False
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
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
            oLast_Item_UID = pval.ItemUID
            oLast_Col_UID = pval.ColUID
            oLast_Col_Row = pval.Row
        End If
    Else
        oLast_Item_UID = pval.ItemUID
        oLast_Col_UID = ""
        oLast_Col_Row = 0
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error

Dim ErrNum As Integer

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Btn_Search" Then '조회버튼
        
            Call PH_PY678_MTX01
        
        ElseIf pval.ItemUID = "Btn_Save" Then '저장버튼
        
            Call PH_PY678_Save
            
        End If
        
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "1" Then
        
            If oForm.Mode = fm_ADD_MODE Then
            
                If pval.ActionSuccess = True Then
                    oForm.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1291") '//이동(최종데이타)
                ElseIf pval.Action_Success = False Then
                    Call PH_PY678_FormItemEnabled
                    Call PH_PY678_AddMatrixRow
                End If
                
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            
                If pval.ActionSuccess = True Then
                    Call PH_PY678_FormItemEnabled
                    Call PH_PY678_AddMatrixRow
                End If
                
            ElseIf oForm.Mode = fm_OK_MODE Then
            
                If pval.ActionSuccess = True Then
                    Call PH_PY678_FormItemEnabled
                End If
                
            End If
            
        End If
        
    End If


'    If pval.BeforeAction = True Then
'
'        If pval.ItemUID = "PH_PY678" Then
'            If oForm.Mode = fm_ADD_MODE Then
'            ElseIf oForm.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm.Mode = fm_OK_MODE Then
'            End If
'        End If
'
'        If pval.ItemUID = "1" Then '/추가/확인 버튼클릭
'
'            If oForm.Mode = fm_ADD_MODE Then
'
'                If PH_PY678_HeaderSpaceLineDel() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
''                If PH_PY678_DataCheck() = False Then
''                    BubbleEvent = False
''                    Exit Sub
''                End If
'
''                Call PH_PY032_FormReset
'                oForm.Mode = fm_ADD_MODE
'
''                Call PH_PY032_LoadCaption
''                Call PH_PY032_MTX01
'
'                oLast_Mode = oForm.Mode
'
'            ElseIf oForm.Mode = fm_UPDATE_MODE Then
'
'                If PH_PY678_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
''                If PH_PY032_DataCheck() = False Then
''                    BubbleEvent = False
''                    Exit Sub
''                End If
'
'                If PH_PY678_UpdateData() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
'                Call PH_PY678_FormReset
'                oForm.Mode = fm_ADD_MODE
'
'                Call PH_PY678_MTX01
'
''                oForm.Items("GCode").Click ct_Regular
'            End If
'
'        ElseIf pval.ItemUID = "Btn01" Then '/조회
'
'            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
'
'            Call PH_PY678_MTX01
'
'        End If
'
'    ElseIf pval.BeforeAction = False Then
'        If pval.ItemUID = "PH_PY678" Then
'            If oForm.Mode = fm_ADD_MODE Then
'            ElseIf oForm.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm.Mode = fm_OK_MODE Then
'            End If
'        End If
'    End If
    
    Exit Sub
    
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error

    If pval.BeforeAction = True Then
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "Mat01", "MSTCOD1") '사번1
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "Mat01", "MSTCOD2") '사번2
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                oLast_Item_UID = pval.ItemUID
                oLast_Col_UID = pval.ColUID
                oLast_Col_Row = pval.Row
                
                Call oMat01.SelectRow(pval.Row, True, False)
            End If
        Else
            oLast_Item_UID = pval.ItemUID
            oLast_Col_UID = ""
            oLast_Col_Row = 0
        End If
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
'        Call PH_PY678_FlushToItemValue(pval.ItemUID)
        
    End If

    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        
    End If
    Exit Sub
Raise_EVENT_DOUBLE_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_DOUBLE_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LINK_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LINK_PRESSED_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Call oForm.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
        
            Call PH_PY678_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
        
'            If (pval.ItemUID = "Mat01") Then
''                If (pval.ColUID = "ItemCode") Then
''                    '//기타작업
''                    Call oDS_PH_PY032B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
''                    If oMat01.RowCount = pval.Row And Trim(oDS_PH_PY032B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
''                        PH_PY032_AddMatrixRow (pval.Row)
''                    End If
''                Else
''                    Call oDS_PH_PY032B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
''                End If
'            Else
'
'                Call PH_PY678_FlushToItemValue(pval.ItemUID)
'
'            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm.Update
        End If
        
    ElseIf pval.BeforeAction = False Then

    End If
    
    Call oForm.Freeze(False)
    
    Exit Sub
    
Raise_EVENT_VALIDATE_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PH_PY678_FormItemEnabled
        Call PH_PY678_AddMatrixRow   '//UDO방식
        Call oMat01.AutoResizeColumns
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY678_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "ItemCode") Then
'            Dim oDataTable01 As SAPbouiCOM.DataTable
'            Set oDataTable01 = pval.SelectedObjects
'            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY032A", "U_CardCode,U_CardName")
'        End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oLast_Item_UID = pval.ItemUID
            oLast_Col_UID = pval.ColUID
            oLast_Col_Row = pval.Row
        End If
    Else
        oLast_Item_UID = pval.ItemUID
        oLast_Col_UID = ""
        oLast_Col_Row = 0
    End If
    Exit Sub
Raise_EVENT_GOT_FOCUS_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_GOT_FOCUS_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_FORM_UNLOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_FORM_UNLOAD_Error
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
        RemoveForms oFormUniqueID01
        Set oForm = Nothing
        Set oMat01 = Nothing
    End If
    Exit Sub
Raise_EVENT_FORM_UNLOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_UNLOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If (oLast_Col_Row > 0) Then
        If pval.BeforeAction = True Then
'            If (PH_PY032_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PH_PY678B.RemoveRecord(oDS_PH_PY678B.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PH_PY678_AddMatrixRow
            Else
                If Trim(oDS_PH_PY678B.GetValue("U_DocNo", oMat01.RowCount - 1)) <> "" Then
                    Call PH_PY678_AddMatrixRow
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY678_CreateItems() As Boolean
On Error GoTo PH_PY678_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY678B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")

    '기간(시작)
    Call oForm.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    
    '기간(종료)
    Call oForm.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY678_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY678_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY678_ComboBox_Setting()
On Error GoTo PH_PY678_ComboBox_Setting_Error

    Dim oColumn As SAPbouiCOM.Column
    Dim sQry    As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
 
    '//////////조회정보//////////
'    '부서
'    Call oForm.Items("STeamCd").Specific.ValidValues.Add("%", "전체")
'    sQry = "            SELECT      U_Code AS [Code],"
'    sQry = sQry & "                 U_CodeNm As [Name]"
'    sQry = sQry & "  FROM       [@PS_HR200L]"
'    sQry = sQry & "  WHERE      Code = '1'"
'    sQry = sQry & "                 AND U_UseYN = 'Y'"
'    sQry = sQry & "  ORDER BY  U_Seq"
'    Call MDC_SetMod.Set_ComboList(oForm.Items("STeamCd").Specific, sQry, "", False, False)
'    Call oForm.Items("STeamCd").Specific.Select(0, psk_Index)

    '//////////매트릭스//////////
    '사업장
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Set oColumn = oMat01.Columns("DutyGbn1")
    oColumn.ValidValues.Add "A", "종일당직"
    oColumn.ValidValues.Add "D", "주간당직"
    oColumn.ValidValues.Add "N", "야간당직"
    
    Set oColumn = oMat01.Columns("DutyGbn2")
    oColumn.ValidValues.Add "A", "종일당직"
    oColumn.ValidValues.Add "D", "주간당직"
    oColumn.ValidValues.Add "N", "야간당직"
    
    Call oForm.Freeze(False)
    Set oColumn = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY678_ComboBox_Setting_Error:

    Call oForm.Freeze(False)
    Set oColumn = Nothing
    Set oRecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PH_PY678_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY678_CF_ChooseFromList()
On Error GoTo PH_PY678_CF_ChooseFromList_Error

    Exit Sub
PH_PY678_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY678_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY678_EnableMenus()
On Error GoTo PH_PY678_EnableMenus_Error
    
    Call oForm.EnableMenu("1283", False) '삭제
    Call oForm.EnableMenu("1286", False) '닫기(미지원)
    Call oForm.EnableMenu("1287", False) '복제
    Call oForm.EnableMenu("1285", False) '복원
    Call oForm.EnableMenu("1284", False) '취소
    Call oForm.EnableMenu("1293", False) '행삭제(미지원)
    Call oForm.EnableMenu("1281", False)
    Call oForm.EnableMenu("1282", True)
    
    Exit Sub
    
PH_PY678_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY678_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY678_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY678_SetDocument_Error

    If (oFromDocEntry01 = "") Then
        Call PH_PY678_FormItemEnabled
'        Call PH_PY678_AddMatrixRow' 최초 첫행 추가
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY678_FormItemEnabled
        oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub

PH_PY678_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY678_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY678_FormResize()
On Error GoTo PH_PY678_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PH_PY678_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY678_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'/폼의 아이템 사용지정
Private Sub PH_PY678_FormItemEnabled()
On Error GoTo PH_PY678_FormItemEnabled_Error
    If (oForm.Mode = fm_ADD_MODE) Then
        
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("FrDt").Enabled = True
        oForm.Items("ToDt").Enabled = True
        
        oForm.Items("Mat01").Enabled = True
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
     
    End If
    
    Exit Sub
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY678_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY678_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PH_PY678_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY678_FlushToItemValue_Error

    Dim i       As Integer
    Dim ErrNum  As Integer
    Dim sQry    As String
    Dim CLTCOD  As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        
        Case "Mat01"
            
            Call oMat01.FlushToDataSource
            
            If oCol = "MSTCOD1" Then
                Call oDS_PH_PY678B.setValue("U_ColReg04", oRow - 1, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oDS_PH_PY678B.GetValue("U_ColReg03", oRow - 1) & "'"))
            End If
            
            If oCol = "MSTCOD2" Then
                Call oDS_PH_PY678B.setValue("U_ColReg07", oRow - 1, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oDS_PH_PY678B.GetValue("U_ColReg06", oRow - 1) & "'"))
            End If
            
            Call oMat01.LoadFromDataSource
            
            Call oMat01.AutoResizeColumns
        
        Case "CLTCOD"
        
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PH_PY678_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PH_PY678_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

Private Sub PH_PY678_AddMatrixRow()
On Error GoTo PH_PY678_AddMatrixRow_Error
    
    Dim oRow As Long
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat01.FlushToDataSource
    oRow = oMat01.VisualRowCount
    
    If oMat01.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY678B.GetValue("U_DocNo", oRow - 1)) <> "" Then
            If oDS_PH_PY678B.Size <= oMat01.VisualRowCount Then
                Call oDS_PH_PY678B.InsertRecord(oRow)
            End If
            oDS_PH_PY678B.Offset = oRow
            Call oDS_PH_PY678B.setValue("U_LineNum", oRow, oRow + 1)
            Call oDS_PH_PY678B.setValue("U_ColDt01", oRow, "")
            Call oDS_PH_PY678B.setValue("U_ColReg01", oRow, "")
            Call oDS_PH_PY678B.setValue("U_ColReg02", oRow, "")
            Call oMat01.LoadFromDataSource
        Else
            oDS_PH_PY678B.Offset = oRow - 1
            Call oDS_PH_PY678B.setValue("U_LineNum", oRow - 1, oRow)
            Call oDS_PH_PY678B.setValue("U_ColDt01", oRow - 1, "")
            Call oDS_PH_PY678B.setValue("U_ColReg01", oRow - 1, "")
            Call oDS_PH_PY678B.setValue("U_ColReg02", oRow - 1, "")
            Call oMat01.LoadFromDataSource
        End If
    ElseIf oMat01.VisualRowCount = 0 Then
        oDS_PH_PY678B.Offset = oRow
        Call oDS_PH_PY678B.setValue("U_LineNum", oRow, oRow + 1)
        Call oDS_PH_PY678B.setValue("U_ColDt01", oRow, "")
        Call oDS_PH_PY678B.setValue("U_ColReg01", oRow, "")
        Call oDS_PH_PY678B.setValue("U_ColReg02", oRow, "")
        Call oMat01.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY678_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY678_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'Private Sub PH_PY678_FormClear()
'On Error GoTo PH_PY678_FormClear_Error
'    Dim DocEntry As String
'    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY678'", "")
'    If DocEntry = 0 Then
'        oForm.Items("DocEntry").Specific.VALUE = 1
'    Else
'        oForm.Items("DocEntry").Specific.VALUE = DocEntry
'    End If
'    Exit Sub
'PH_PY678_FormClear_Error:
'    Sbo_Application.SetStatusBarMessage "PH_PY678_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Sub

'Private Function PH_PY678_DataValidCheck() As Boolean
'On Error GoTo PH_PY678_DataValidCheck_Error
'    PH_PY678_DataValidCheck = False
'    Dim i       As Long
'    Dim sQry    As String
'    Dim oRecordSet  As SAPbobsCOM.Recordset
'
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'    '사업장
'    If Trim(oDS_PH_PY678A.GetValue("U_CLTCOD", 0)) = "" Then
'        Call Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", bmt_Short, True)
'        Call oForm.Items("CLTCOD").CLICK(ct_Regular)
'        PH_PY678_DataValidCheck = False
'        Exit Function
'    End If
'
'    '팀
'    If Trim(oDS_PH_PY678A.GetValue("U_TeamCode", 0)) = "" Then
'        Call Sbo_Application.SetStatusBarMessage("팀정보는 필수입니다.", bmt_Short, True)
'        Call oForm.Items("TeamCode").CLICK(ct_Regular)
'        PH_PY678_DataValidCheck = False
'        Exit Function
'    End If
'
'    '기간(시작)
'    If Trim(oDS_PH_PY678A.GetValue("U_FrDt", 0)) = "" Then
'        Call Sbo_Application.SetStatusBarMessage("기간(시작) 필수입니다.", bmt_Short, True)
'        Call oForm.Items("FrDt").CLICK(ct_Regular)
'        PH_PY678_DataValidCheck = False
'        Exit Function
'    End If
'
'    '기간(종료)
'    If Trim(oDS_PH_PY678A.GetValue("U_ToDt", 0)) = "" Then
'        Call Sbo_Application.SetStatusBarMessage("기간(종료)는 필수입니다.", bmt_Short, True)
'        Call oForm.Items("ToDt").CLICK(ct_Regular)
'        PH_PY678_DataValidCheck = False
'        Exit Function
'    End If
'
'    '라인
'    If oMat01.VisualRowCount > 1 Then
'        For i = 1 To oMat01.VisualRowCount - 1
'
'            '배부규칙
'            If oMat01.Columns("ProfCode").Cells(i).Specific.VALUE = "" And oMat01.Columns("DocNo").Cells(i).Specific.VALUE <> 0 Then
'                Call Sbo_Application.SetStatusBarMessage("배부규칙은 필수입니다.", bmt_Short, True)
'                Call oMat01.Columns("ProfCode").Cells(i).CLICK(ct_Regular)
'                PH_PY678_DataValidCheck = False
'                Exit Function
'            End If
'
'            '적요
'            If oMat01.Columns("LineMemo").Cells(i).Specific.VALUE = "" And oMat01.Columns("DocNo").Cells(i).Specific.VALUE <> 0 Then
'                Call Sbo_Application.SetStatusBarMessage("적요는 필수입니다.", bmt_Short, True)
'                Call oMat01.Columns("LineMemo").Cells(i).CLICK(ct_Regular)
'                PH_PY678_DataValidCheck = False
'                Exit Function
'            End If
'
'        Next
'    Else
'        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
'        PH_PY678_DataValidCheck = False
'        Exit Function
'    End If
'
'        oMat01.FlushToDataSource
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY678B.Size > 1 Then oDS_PH_PY678B.RemoveRecord (oDS_PH_PY678B.Size - 1)
'
'    oMat01.LoadFromDataSource
'
'    PH_PY678_DataValidCheck = True
'
'    Exit Function
'
'    Set oRecordSet = Nothing
'
'
'PH_PY678_DataValidCheck_Error:
'    Set oRecordSet = Nothing
'    PH_PY678_DataValidCheck = False
'    Sbo_Application.SetStatusBarMessage "PH_PY678_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Function

Public Sub PH_PY678_MTX01()
'******************************************************************************
'Function ID : PH_PY678_MTX01()
'해당모듈    : PH_PY678
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY678_MTX01_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As Integer
    
    Dim CLTCOD  As String '사업장
    Dim FrDt    As String '기간(시작)
    Dim ToDt    As String '기간(종료)
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    FrDt = oForm.Items("FrDt").Specific.VALUE '기간(시작)
    ToDt = oForm.Items("ToDt").Specific.VALUE '기간(종료)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "            EXEC [PH_PY678_01] "
    sQry = sQry & "'" & CLTCOD & "',"  '사업장
    sQry = sQry & "'" & FrDt & "',"  '기간(시작)
    sQry = sQry & "'" & ToDt & "'"  '기간(종료)
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PH_PY678B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
'        oForm.Mode = fm_ADD_MODE
        
        Call PH_PY678_FormItemEnabled
        Call PH_PY678_AddMatrixRow
        
        GoTo PH_PY678_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PH_PY678B.Size Then
            Call oDS_PH_PY678B.InsertRecord(i)
        End If
        
        oMat01.AddRow
        oDS_PH_PY678B.Offset = i
        
        Call oDS_PH_PY678B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY678B.setValue("U_ColDt01", i, Trim(oRecordSet01.Fields("Date").VALUE)) '일자
        Call oDS_PH_PY678B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DayName").VALUE)) '요일
        Call oDS_PH_PY678B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("DayType").VALUE)) '요일구분
        Call oDS_PH_PY678B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("MSTCOD1").VALUE)) '사번
        Call oDS_PH_PY678B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("MSTNAM1").VALUE)) '성명
        Call oDS_PH_PY678B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("DutyGbn1").VALUE)) '당직근무구분
        Call oDS_PH_PY678B.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("MSTCOD2").VALUE)) '사번2
        Call oDS_PH_PY678B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("MSTNAM2").VALUE)) '성명2
        Call oDS_PH_PY678B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("DutyGbn2").VALUE)) '당직근무구분2
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    Call PH_PY678_AddMatrixRow
    ProgBar01.Stop
    Call oForm.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY678_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY678_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PH_PY678_Save()
'******************************************************************************
'Function ID : PH_PY678_Save()
'해당모듈    : PH_PY678
'기능        : 데이터 저장
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY678_Save_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As Integer
    
    Dim CLTCOD  As String '사업장
    Dim StdDate As String '기준일
    Dim MSTCOD  As String '사번
    Dim DutyGbn As String '당직근무형태
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 1
     
        CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
        StdDate = oDS_PH_PY678B.GetValue("U_ColDt01", i) '기준일
        MSTCOD = Trim(oDS_PH_PY678B.GetValue("U_ColReg03", i)) '사번
        DutyGbn = Trim(oDS_PH_PY678B.GetValue("U_ColReg05", i)) '당직근무 형태
             
        If Trim(MSTCOD) <> "" Then
             
            sQry = "            EXEC [PH_PY678_02] "
            sQry = sQry & "'" & CLTCOD & "',"  '사업장
            sQry = sQry & "'" & StdDate & "',"  '기준일
            sQry = sQry & "'" & MSTCOD & "',"  '사번
            sQry = sQry & "'" & DutyGbn & "'"  '당직근무형태
            
            Call oRecordSet01.DoQuery(sQry)
        End If
     
        '토요일 주간당직 등록
        MSTCOD = Trim(oDS_PH_PY678B.GetValue("U_ColReg06", i)) '사번
        DutyGbn = Trim(oDS_PH_PY678B.GetValue("U_ColReg08", i)) '당직근무 형태
        
        If Trim(MSTCOD) <> "" Then
             
            sQry = "            EXEC [PH_PY678_02] "
            sQry = sQry & "'" & CLTCOD & "',"  '사업장
            sQry = sQry & "'" & StdDate & "',"  '기준일
            sQry = sQry & "'" & MSTCOD & "',"  '사번
            sQry = sQry & "'" & DutyGbn & "'"  '당직근무형태
            
            Call oRecordSet01.DoQuery(sQry)
        End If
        
'        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 저장중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    ProgBar01.Stop
    Call oForm.Freeze(False)
    
    Call MDC_Com.MDC_GF_Message("저장 완료", "S")
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY678_Save_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
       
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY678_Save_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub


