VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SM010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_SM010.cls
'//  Module         : 재고관리 > 재고장
'//  Desc           : 재고장
'//  FormType       : PS_SM010
'//  Create Date    : 2010.08.20
'//  Copyright  (c) Morning Data
'//********************************************************************************

'//사용방법
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_SM010
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String
Public oBaseItmBsort01 As String

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    BatchNum As String
    Quantity As Long
    OPORNo As Long
    POR1No As Long
    Check As Boolean
    OPDNNo As Long
    PDN1No As Long
End Type
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String, Optional ByVal oItmBsort02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SM010.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_SM010_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Set oBaseForm01 = oForm02
    oBaseItemUID01 = oItemUID02
    oBaseColUID01 = oColUID02
    oBaseColRow01 = oColRow02
    oBaseTradeType01 = oTradeType02
    oBaseItmBsort01 = oItmBsort02
    
    Call PS_SM010_CreateItems
    Call PS_SM010_ComboBox_Setting
    Call PS_SM010_CF_ChooseFromList
    Call PS_SM010_FormItemEnabled
    Call PS_SM010_EnableMenus
    '//Call PS_SM010_FormClear '//UDO방식일때
    '//Call PS_SM010_AddMatrixRow(0, True) '//UDO방식일때
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
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
Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_SM010_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_SM010_FormItemEnabled '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
'            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
'            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
'            MenuCreationParams01.uniqueID = "MenuUID"
'            MenuCreationParams01.String = "메뉴명"
'            MenuCreationParams01.Enabled = True
'            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
'        End If
    ElseIf pval.BeforeAction = False Then
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'                Call Sbo_Application.Menus.RemoveEx("MenuUID")
'        End If
    End If
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_SM010_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_SM010_SetBaseForm '//부모폼에입력
                If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
                    Call oForm01.Close
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_SM010_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_SM010" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SM010_FormItemEnabled
'                    Call PS_SM010_FormClear '//UDO방식일때
'                    Call PS_SM010_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SM010_FormItemEnabled
'                End If
'            End If
'        End If
    End If
    Exit Sub
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    If pval.BeforeAction = True Then
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemName", "") '//사용자값활성
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "ItmBsort" Then
            Dim i As Long
            For i = 0 To oForm01.Items("ItmMsort").Specific.ValidValues.Count - 1
                Call oForm01.Items("ItmMsort").Specific.ValidValues.Remove(0, psk_Index)
            Next
            Call oForm01.Items("ItmMsort").Specific.ValidValues.Add("선택", "선택")
            Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" & oForm01.Items("ItmBsort").Specific.Selected.VALUE & "' ORDER BY U_Code", "", False, False)
            If oForm01.Items("ItmMsort").Specific.ValidValues.Count > 0 Then
                Call oForm01.Items("ItmMsort").Specific.Select(0, psk_Index)
            End If
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Grid01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row > 0 Then
                    
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "Grid01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row >= 0 Then
                    oForm01.Items("LItemCod").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(0, ot_SelectionOrder)).VALUE 'Item(pval.Row, ot_SelectionOrder)).VALUE
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Grid01" Then
            If pval.Row = -1 Then
'                oGrid01.Columns(pval.ColUID).TitleObject.Sortable = True
                
            Else
                If oGrid01.Rows.SelectedRows.Count > 0 Then
                    Call PS_SM010_SetBaseForm '//부모폼에입력
                    If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
                        Call oForm01.Close
                    End If
                Else
                    BubbleEvent = False
                End If
            End If
        End If
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
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_SM010_FormItemEnabled
        '//Call PS_SM010_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
'        oForm01.Items("Mat01").Top = 70
'        oForm01.Items("Mat01").Height = (oForm01.Height / 2) - 70
'        oForm01.Items("Mat01").Left = 7
'        oForm01.Items("Mat01").Width = oForm01.Width - 21
'        oForm01.Items("Mat02").Top = (oForm01.Height / 2) + 10
'        oForm01.Items("Mat02").Height = (oForm01.Height / 2) - 75
'        oForm01.Items("Mat02").Left = 7
'        oForm01.Items("Mat02").Width = oForm01.Width - 21
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    Dim oDataTable01  As SAPbouiCOM.DataTable
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        If (pval.ItemUID = "ItemCode") Then
            Set oDataTable01 = pval.SelectedObjects
            If oDataTable01 Is Nothing Then
            Else
                oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
                oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
            End If
        End If
        oForm01.Update
    End If
    Set oDataTable01 = Nothing
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
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
        Set oForm01 = Nothing
        Set oGrid01 = Nothing
    End If
    Exit Sub
Raise_EVENT_FORM_UNLOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_UNLOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
    '        For i = 1 To oMat01.VisualRowCount
    '            oMat01.Columns("COL01").Cells(i).Specific.Value = i
    '        Next i
    '        oMat01.FlushToDataSource
    '        Call oDS_PS_SM010L.RemoveRecord(oDS_PS_SM010L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_SM010_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_SM010_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_SM010_CreateItems() As Boolean
On Error GoTo PS_SM010_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    'oGrid01.SelectionMode = ms_NotSupported
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 30)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    
    Call oForm01.DataSources.UserDataSources.Add("StockType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("StockType").Specific.DataBind.SetBound(True, "", "StockType")
    Call oForm01.DataSources.UserDataSources.Add("TradeType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("TradeType").Specific.DataBind.SetBound(True, "", "TradeType")
    
    Call oForm01.DataSources.UserDataSources.Add("ItemGpCd", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItemGpCd").Specific.DataBind.SetBound(True, "", "ItemGpCd")
    Call oForm01.DataSources.UserDataSources.Add("ItmBsort", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItmBsort").Specific.DataBind.SetBound(True, "", "ItmBsort")
    Call oForm01.DataSources.UserDataSources.Add("ItmMsort", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItmMsort").Specific.DataBind.SetBound(True, "", "ItmMsort")
    Call oForm01.DataSources.UserDataSources.Add("Size", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Size").Specific.DataBind.SetBound(True, "", "Size")
    Call oForm01.DataSources.UserDataSources.Add("ItemType", dt_SHORT_TEXT, 50)
    Call oForm01.Items("ItemType").Specific.DataBind.SetBound(True, "", "ItemType")
    Call oForm01.DataSources.UserDataSources.Add("Mark", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Mark").Specific.DataBind.SetBound(True, "", "Mark")

    Dim oChkBox As SAPbouiCOM.CheckBox

    Call oForm01.DataSources.UserDataSources.Add("Check01", dt_SHORT_TEXT, 1)
    Set oChkBox = oForm01.Items("ChkBox01").Specific
    oChkBox.ValOn = "Y"
    oChkBox.ValOff = "N"
    oChkBox.DataBind.SetBound True, , "Check01"
    oForm01.DataSources.UserDataSources("Check01").VALUE = "N" '미체크로 값을 주고 폼을 로드
    
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_SM010_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM010_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SM010_ComboBox_Setting()
On Error GoTo PS_SM010_ComboBox_Setting_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SM010", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SM010", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_SM010", "Mat01", "ItemCode")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM010", "StockType", "", "1", "재고있는품목")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM010", "StockType", "", "2", "전체")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("StockType").Specific, "PS_SM010", "StockType")
    Call oForm01.Items("StockType").Specific.Select("2", psk_ByValue)
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM010", "TradeType", "", "", "전체")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM010", "TradeType", "", "1", "일반")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM010", "TradeType", "", "2", "임가공")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("TradeType").Specific, "PS_SM010", "TradeType")
    Call oForm01.Items("TradeType").Specific.Select(oBaseTradeType01, psk_ByValue)
    
    Call oForm01.Items("ItmBsort").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", False, False)
    Call oForm01.Items("ItmMsort").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", "", False, False)
    Call oForm01.Items("ItemType").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType").Specific, "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", "", False, False)
    Call oForm01.Items("Mark").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", "", False, False)
    Call oForm01.Items("ItemGpCd").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemGpCd").Specific, "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]", "", False, False)
    Call oForm01.Items("ItmBsort").Specific.Select(0, psk_Index)
    Call oForm01.Items("ItmMsort").Specific.Select(0, psk_Index)
    Call oForm01.Items("ItemType").Specific.Select(0, psk_Index)
    Call oForm01.Items("Mark").Specific.Select(0, psk_Index)
    Call oForm01.Items("ItemGpCd").Specific.Select(0, psk_Index)
    '// 창고
    Set oCombo = oForm01.Items("WhsCode").Specific
    sQry = "SELECT WhsCode, WhsName From OWHS"
    oRecordSet01.DoQuery sQry
    
    Call oCombo.ValidValues.Add("선택", "선택")
    
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oCombo.Select(0, psk_Index)
    
    oForm01.Items("ItemCode").Click ct_Regular
    oForm01.Items("TradeType").Enabled = False
    
    
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM010_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM010_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SM010_CF_ChooseFromList()
On Error GoTo PS_SM010_CF_ChooseFromList_Error
    '//ChooseFromList 설정
'    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
'    Dim oCons               As SAPbouiCOM.Conditions
'    Dim oCon                As SAPbouiCOM.Condition
'    Dim oCFL                As SAPbouiCOM.ChooseFromList
'    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
'    Dim oEdit               As SAPbouiCOM.EditText
'    Dim oColumn             As SAPbouiCOM.Column
'
'    Set oEdit = oForm01.Items("ItemCode").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = "4"
'    oCFLCreationParams.uniqueID = "CFLITEMCD"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)
'
''    Set oCons = oCFL.GetConditions()
''    Set oCon = oCons.Add()
''    oCon.Alias = "CardType"
''    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
''    oCon.CondVal = "C"
''    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLITEMCD"
'    oEdit.ChooseFromListAlias = "ItemCode"
    Exit Sub
PS_SM010_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM010_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SM010_FormItemEnabled()
On Error GoTo PS_SM010_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseItmBsort01 = "" Then
        
        Else
            Call oForm01.Items("ItmBsort").Specific.Select(oBaseItmBsort01, psk_ByValue)
            oForm01.Items("ItmBsort").Enabled = False
        End If
        '//각모드에따른 아이템설정
        '//Call PS_SM010_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM010_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM010_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SM010_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SM010_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_SM010L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_SM010L.Offset = oRow
'    oDS_PS_SM010L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM010_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM010_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SM010_FormClear()
On Error GoTo PS_SM010_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SM010'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SM010_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM010_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SM010_EnableMenus()
On Error GoTo PS_SM010_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_SM010_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM010_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SM010_DataValidCheck() As Boolean
On Error GoTo PS_SM010_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_SM010_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_SM010_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_SM010_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_SM010_FormClear
    Exit Function
PS_SM010_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM010_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SM010_MTX01()
On Error GoTo PS_SM010_MTX01_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim Param08 As String
    Dim Param09 As String
    Dim Param10 As String
    Dim Param11 As String
    
    Param01 = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    Param02 = Trim(oForm01.Items("StockType").Specific.Selected.VALUE)
    Param03 = Trim(oForm01.Items("TradeType").Specific.Selected.VALUE)
    Param04 = Trim(oForm01.Items("ItmBsort").Specific.Selected.VALUE)
    Param05 = Trim(oForm01.Items("ItmMsort").Specific.Selected.VALUE)
    Param06 = Trim(oForm01.Items("Size").Specific.VALUE)
    Param07 = Trim(oForm01.Items("ItemType").Specific.Selected.VALUE)
    Param08 = Trim(oForm01.Items("Mark").Specific.Selected.VALUE)
    Param09 = Trim(oForm01.Items("ItemName").Specific.VALUE)
    Param10 = Trim(oForm01.Items("ItemGpCd").Specific.Selected.VALUE)
    Param11 = Trim(oForm01.Items("WhsCode").Specific.Selected.VALUE)
    
    If oBaseForm01 Is Nothing Then
        Query01 = "EXEC PS_SM010_01 '" & Param01 & "','','','" & Param02 & "','','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "','" & Param11 & "'"
    ElseIf oBaseForm01.Type = "149" Or oBaseForm01.Type = "139" Or oBaseForm01.Type = "140" Or oBaseForm01.Type = "180" Or oBaseForm01.Type = "133" Or oBaseForm01.Type = "179" Or oBaseForm01.Type = "60091" Then
        Query01 = "EXEC PS_SM010_01 '" & Param01 & "','Y','','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "','" & Param11 & "'" '//판매Y,구매,재고타입(1:재고있는것만,2:전체),거래타입(1:일반,2:임가공)
    ElseIf oBaseForm01.Type = "142" Or oBaseForm01.Type = "143" Or oBaseForm01.Type = "182" Or oBaseForm01.Type = "141" Or oBaseForm01.Type = "181" Or oBaseForm01.Type = "60092" Then
        Query01 = "EXEC PS_SM010_01 '" & Param01 & "','','Y','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "','" & Param11 & "'" '//판매,구매Y
    Else
        Query01 = "EXEC PS_SM010_01 '" & Param01 & "','','','" & Param02 & "','','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "','" & Param11 & "'"
    End If
    
    oGrid01.DataTable.Clear
    
    Call oForm01.DataSources.DataTables.Item("DataTable").ExecuteQuery(Query01)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
    If oGrid01.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_SM010_MTX01_Exit
    End If
        
    oGrid01.AutoResizeColumns
    oForm01.Update
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM010_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM010_MTX01_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM010_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SM010_DI_API() As Boolean
'On Error GoTo PS_SM010_DI_API_Error
'    PS_SM010_DI_API = True
'    Dim i, j As Long
'    Dim oDIObject As SAPbobsCOM.Documents
'    Dim RetVal As Long
'    Dim LineNumCount As Long
'    Dim ResultDocNum As Long
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Company.StartTransaction
'
'    ReDim ItemInformation(0)
'    ItemInformationCount = 0
'    For i = 1 To oMat01.VisualRowCount
'        ReDim Preserve ItemInformation(ItemInformationCount)
'        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).Check = False
'        ItemInformationCount = ItemInformationCount + 1
'    Next
'
'    LineNumCount = 0
'    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
'    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
'    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
'    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.Value, "&&&&-&&-&&")
'    For i = 0 To UBound(ItemInformation)
'        If ItemInformation(i).Check = True Then
'            GoTo Continue_First
'        End If
'        If i <> 0 Then
'            oDIObject.Lines.Add
'        End If
'        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
'        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
'        oDIObject.Lines.BaseType = "22"
'        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
'        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
'        For j = i To UBound(ItemInformation)
'            If ItemInformation(j).Check = True Then
'                GoTo Continue_Second
'            End If
'            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
'                GoTo Continue_Second
'            End If
'            '//같은것
'            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
'            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
'            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
'            oDIObject.Lines.BatchNumbers.Add
'            ItemInformation(j).PDN1No = LineNumCount
'            ItemInformation(j).Check = True
'Continue_Second:
'        Next
'        LineNumCount = LineNumCount + 1
'Continue_First:
'    Next
'    RetVal = oDIObject.Add
'    If RetVal = 0 Then
'        ResultDocNum = Sbo_Company.GetNewObjectKey
'        For i = 0 To UBound(ItemInformation)
'            Call oDS_PS_SM010L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_SM010L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
'    Else
'        GoTo PS_SM010_DI_API_Error
'    End If
'
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_Commit
'    End If
'    oMat01.LoadFromDataSource
'    oMat01.AutoResizeColumns
'
'    Set oDIObject = Nothing
'    Exit Function
'PS_SM010_DI_API_DI_Error:
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
'    PS_SM010_DI_API = False
'    Set oDIObject = Nothing
'    Exit Function
'PS_SM010_DI_API_Error:
'    Sbo_Application.SetStatusBarMessage "PS_SM010_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    PS_SM010_DI_API = False
End Function

Private Sub PS_SM010_SetBaseForm()
On Error GoTo PS_SM010_SetBaseForm_Error
    Dim i As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim sQry02 As String
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Today   As String
    Dim ItemCode As String
        
        
    If oBaseForm01 Is Nothing Then
    '//DoNothing
    ElseIf oBaseForm01.TypeEx = "721" Then '//입고
        Set oBaseMat01 = oBaseForm01.Items("13").Specific '//부모폼의매트릭스
        For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
            oBaseMat01.Columns("1").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
            oBaseColRow01 = oBaseColRow01 + 1
        Next
    ElseIf oBaseForm01.TypeEx = "PS_MM180" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "BItemCod" Then
            oBaseForm01.Items("BItemCod").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_SD091" Then '//사용하려는폼의 폼타입
        Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
        For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
            oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
            oBaseColRow01 = oBaseColRow01 + 1
        Next

    ElseIf oBaseForm01.TypeEx = "PS_QM010" Then '//사용하려는폼의 폼타입
        oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
    ElseIf oBaseForm01.TypeEx = "PS_PP077" Then '//사용하려는폼의 폼타입
        oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
    ElseIf oBaseForm01.TypeEx = "PS_PP078" Then '//사용하려는폼의 폼타입
        oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
    '//RYJ
    ElseIf oBaseForm01.TypeEx = "PS_MM005" Then '//사용하려는폼의 폼타입
'        If oBaseItemUID01 = "ItemCode" Then
'            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
'        ElseIf oBaseItemUID01 = "Mat01" Then
'            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
'            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
'                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
'                oBaseColRow01 = oBaseColRow01 + 1
'            Next
'        End If

        
        
        
        sQry = "           SELECT   CASE"
        sQry = sQry & "                 WHEN LEFT(U_ItmBsort, 1) = '4' THEN '20'"
        sQry = sQry & "                 WHEN LEFT(U_ItmBsort, 1) = '3' THEN '10'"
        sQry = sQry & "                 WHEN LEFT(U_ItmBsort, 1) = '2' THEN '50'"
        sQry = sQry & "             END"
        sQry = sQry & " FROM    OITM"
        sQry = sQry & " WHERE   ItemCode = '" & oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE & "'"
        
        Call oRecordSet01.DoQuery(sQry)
        
        sQry02 = " Select convert(char(8),GetDate(),112) "
        oRecordset02.DoQuery sQry02
        Today = Trim(oRecordset02.Fields(0).VALUE)
        
'        If Today >= "20150701" Then
'            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
'                ItemCode = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
'                If Trim(oRecordSet01.Fields(0).VALUE) = "20" Then
'                    If Len(ItemCode) <> 11 Then
'                        oForm01.DataSources.UserDataSources("Check01").VALUE = "Y"
'                        Call MDC_Com.MDC_GF_Message("부자재청구는 통합부자재 코드만 가능합니다. 확인바랍니다.", "E")
'                        Exit Sub
'                    End If
'                End If
'            Next
'        End If
        
        
        
        If oBaseForm01.Items("OrdType").Specific.VALUE = "10" Or oBaseForm01.Items("OrdType").Specific.VALUE = "20" Or oBaseForm01.Items("OrdType").Specific.VALUE = "50" Then
        
            If oRecordSet01.Fields(0).VALUE = oBaseForm01.Items("OrdType").Specific.VALUE Then
                If oBaseItemUID01 = "ItemCode" Then
                    oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                ElseIf oBaseItemUID01 = "Mat01" Then
                    Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
                    For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                        oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                        oBaseColRow01 = oBaseColRow01 + 1
                    Next
                End If
            Else
                Call MDC_Com.MDC_GF_Message("품의와 품목구분이 일치하지않아 등록할 수 없습니다. 확인하십시오.", "E")
                Exit Sub
            End If
        Else
            If oBaseItemUID01 = "ItemCode" Then
                oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
            ElseIf oBaseItemUID01 = "Mat01" Then
                Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
                For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                    oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                    oBaseColRow01 = oBaseColRow01 + 1
                Next
            End If
        End If
    ElseIf oBaseForm01.TypeEx = "PS_MM030" Then '//사용하려는폼의 폼타입
    
        If oBaseItemUID01 = "ItemCode" Then
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_PP043" Then '//사용하려는폼의 폼타입
    
        If oBaseItemUID01 = "ItemCode" Then
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("OrdMgNum").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_MM070" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_SD010" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_PP010" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_PP020" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    '//RYJ
    ElseIf oBaseForm01.TypeEx = "PS_MM005" Then '//사용하려는폼의 폼타입
        
    ElseIf oBaseForm01.TypeEx = "PS_MM030" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_MM070" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_SD010" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_PP010" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_PP020" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_QM060" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    
    ElseIf oBaseForm01.TypeEx = "PS_PP950" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        End If
    ElseIf oBaseForm01.TypeEx = "PS_MM002" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("MItemCod").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_PP130" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                If oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = "" Then
                    oBaseMat01.Columns("ItmMsort").Cells(oBaseColRow01).Specific.VALUE = Mid(oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE, 1, 5)
                End If
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_PP940" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                If oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = "" Then
                    oBaseMat01.Columns("ItmMsort").Cells(oBaseColRow01).Specific.VALUE = Mid(oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE, 1, 5)
                End If
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_QM903" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                If oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = "" Then
                    oBaseMat01.Columns("ItmBsort").Cells(oBaseColRow01).Specific.VALUE = Mid(oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE, 1, 3)
                    oBaseMat01.Columns("ItmMsort").Cells(oBaseColRow01).Specific.VALUE = Mid(oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE, 1, 5)
                End If
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_MM097" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_PP083" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_SD005" Then '//업체별 단가등록
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_SD006" Then '//업체별 판매계획등록
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_PP025" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_QM005" Then '분말 제품검사사양서 등록 //사용하려는폼의 폼타입
        oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
    ElseIf oBaseForm01.TypeEx = "PS_MM100" Then '//사용하려는폼의 폼타입
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_MM007" Then '//통구실 품목코드 생성요청
        


        
        
        
        sQry = "           SELECT   CASE"
        sQry = sQry & "                 WHEN LEFT(U_ItmBsort, 1) = '4' THEN '20'"
        sQry = sQry & "                 WHEN LEFT(U_ItmBsort, 1) = '3' THEN '10'"
        sQry = sQry & "                 WHEN LEFT(U_ItmBsort, 1) = '2' THEN '50'"
        sQry = sQry & "             END"
        sQry = sQry & " FROM    OITM"
        sQry = sQry & " WHERE   ItemCode = '" & oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE & "'"
        
        Call oRecordSet01.DoQuery(sQry)
        
        sQry02 = " Select convert(char(8),GetDate(),112) "
        oRecordset02.DoQuery sQry02
        Today = Trim(oRecordset02.Fields(0).VALUE)
        
        If Today >= "20150701" Then
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                ItemCode = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                If Trim(oRecordSet01.Fields(0).VALUE) = "20" Then
                    If Len(ItemCode) <> 11 Then
                        oForm01.DataSources.UserDataSources("Check01").VALUE = "Y"
                        Call MDC_Com.MDC_GF_Message("통합부자재 코드로만 선택이 가능합니다. 확인바랍니다.", "E")
                        Exit Sub
                    End If
                End If
            Next
        End If
        
        If oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            
        
            
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("MATNR").Cells(oBaseColRow01).Specific.VALUE = Left(oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE, 7)
                oBaseMat01.Columns("MAKTX").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품명").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE & " : "
                oBaseMat01.Columns("MEINS").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("단위").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseMat01.Columns("MATKL").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("분류코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    ElseIf oBaseForm01.TypeEx = "PS_QM011" Then '//고객불만등록
        If oBaseItemUID01 = "ItemCode" Then
            oBaseForm01.Items("ItemCode").Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
        ElseIf oBaseItemUID01 = "Mat01" Then
            Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//부모폼의매트릭스
            For i = 0 To oGrid01.Rows.SelectedRows.Count - 1 '//선택된행의수
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oGrid01.DataTable.Columns("품목코드").Cells(oGrid01.Rows.SelectedRows.Item(i, ot_SelectionOrder)).VALUE
                oBaseColRow01 = oBaseColRow01 + 1
            Next
        End If
    End If
    Exit Sub
PS_SM010_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM010_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



