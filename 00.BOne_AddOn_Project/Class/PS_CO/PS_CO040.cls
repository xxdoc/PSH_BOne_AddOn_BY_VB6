VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_CO040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_CO040.cls
'//  Module         : 기본관리 > 기본
'//  Desc           : 기본
'//  FormType       : PS_CO040
'//  Create Date    : 2010.10.25
'//  Create By      : Kim Dong Sub
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_CO040H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_CO040L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    LotNo As String
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
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_CO040.srf")
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
    
    oFormUniqueID01 = "PS_CO040_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "Code" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_CO040_CreateItems
    Call PS_CO040_ComboBox_Setting
    Call PS_CO040_CF_ChooseFromList
    Call PS_CO040_EnableMenus
    Call PS_CO040_SetDocument(oFromDocEntry01)
    Call PS_CO040_FormResize
    
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
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_CO040_FormItemEnabled '//UDO방식
'                Call PS_CO040_AddMatrixRow(oMat01.VisualRowCount, False) '//UDO방식
            Case "1282": '추가
                Call PS_CO040_FormItemEnabled '//UDO방식
                Call PS_CO040_AddMatrixRow(0, True) '//UDO방식
            Case "1283":
'                oForm01.Mode = fm_ADD_MODE
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_CO040_FormItemEnabled
'                Call PS_CO040_AddMatrixRow(oMat01.VisualRowCount, False) '//UDO방식
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

    ElseIf pval.BeforeAction = False Then

    End If
    If pval.ItemUID = "Mat01" Then
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
        If pval.ItemUID = "PS_CO040" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_CO040_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_CO040_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_CO040_FormItemEnabled
                Call PS_CO040_AddMatrixRow(0, True) '//UDO방식일때
            End If
            If oForm01.Mode = fm_OK_MODE Then
'                Call PS_CO040_AddMatrixRow(oDS_PS_CO040L.Offset + 1, False)
            End If
'            If oForm01.Mode = fm_FIND_MODE And pval.Action_Success = False Then
'                Call PS_CO040_AddMatrixRow(oMat01.RowCount, True)
'            End If
        End If
        If pval.ItemUID = "PS_CO040" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If

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

    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "CCCode") Then
                    '//기타작업
                    Call oDS_PS_CO040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_CO040L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_CO040_AddMatrixRow (pval.Row)
                    End If
                Else
                    Call oDS_PS_CO040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            Else
                If (pval.ItemUID = "Code") Then
                    Call oDS_PS_CO040H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "Name") Then
                    Call oDS_PS_CO040H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                Else
                
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oForm01.Update
        End If
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
        Call PS_CO040_FormItemEnabled
        Call PS_CO040_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_CO040_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
    Dim oDataTable01 As SAPbouiCOM.DataTable
    Dim i As Integer
    Dim oRowCount As Integer
        If (pval.ItemUID = "Mat01") Then
            If (pval.ColUID = "CCCode") Then
                oMat01.FlushToDataSource
                Set oDataTable01 = pval.SelectedObjects
                If pval.Row <= "1" Then
                    oRowCount = 0
                Else
                    oRowCount = pval.Row - 1
                End If
                
                For i = 0 To oDataTable01.Rows.Count - 1
                    Call oDS_PS_CO040L.setValue("U_CCCode", oRowCount, oDataTable01.Columns("PrcCode").Cells(i).VALUE)
                    Call oDS_PS_CO040L.setValue("U_CCName", oRowCount, oDataTable01.Columns("PrcName").Cells(i).VALUE)
                    Call oDS_PS_CO040L.setValue("U_GrpCode", oRowCount, oDataTable01.Columns("GrpCode").Cells(i).VALUE)
                   
                    Call PS_CO040_AddMatrixRow(oRowCount + 1, False)
                    oRowCount = pval.Row + i
                Next

                oMat01.LoadFromDataSource
                Set oDataTable01 = Nothing
                oMat01.LoadFromDataSource
                If oForm01.Mode = fm_OK_MODE Then
                    oForm01.Mode = fm_UPDATE_MODE
                End If
           End If
       End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat01" Then
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
        Set oMat01 = Nothing
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
'            If (PS_CO040_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineId").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_CO040L.RemoveRecord(oDS_PS_CO040L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_CO040_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_CO040L.GetValue("U_CCCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_CO040_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_CO040_CreateItems() As Boolean
On Error GoTo PS_CO040_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_CO040H = oForm01.DataSources.DBDataSources("@PS_CO040H")
    Set oDS_PS_CO040L = oForm01.DataSources.DBDataSources("@PS_CO040L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
'    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
'    Call oForm01.DataSources.UserDataSources.Add("WhsCode", dt_SHORT_TEXT, 100)
'    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
'    Call oForm01.Items("WhsCode").Specific.DataBind.SetBound(True, "", "WhsCode")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_CO040_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO040_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_CO040_ComboBox_Setting()
On Error GoTo PS_CO040_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    Call oForm01.Freeze(False)
    Exit Sub
PS_CO040_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO040_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_CO040_CF_ChooseFromList()
On Error GoTo PS_CO040_CF_ChooseFromList_Error
'   ChooseFromList 설정
'   코스트센터 - 매트릭스
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column
    
    Set oColumn = oMat01.Columns("CCCode")
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
    
    oCFLCreationParams.ObjectType = 61
    oCFLCreationParams.uniqueID = "CFLCCCode"
    oCFLCreationParams.MultiSelection = True
    Set oCFL = oCFLs.Add(oCFLCreationParams)
    
    ' Choose from list 에 조건을 줄 경우
    ' choosefromlist가 화면에 나오면 서식세팅으로 원하는 필드값 추가 가능
    Set oCons = oCFL.GetConditions()
    Set oCon = oCons.Add()
    oCon.Alias = "DimCode"                                               ' Condition Field
    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL            ' Equal
    oCon.CondVal = "1"                                                   ' Condition Value
    oCFL.SetConditions oCons
    
    oColumn.ChooseFromListUID = "CFLCCCode"
    oColumn.ChooseFromListAlias = "PrcCode"

    Exit Sub
PS_CO040_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO040_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_CO040_FormItemEnabled()
On Error GoTo PS_CO040_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_CO040_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("Code").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("Code").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("Code").Enabled = True
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//찾기
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_CO040_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO040_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_CO040_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_CO040_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_CO040L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_CO040L.Offset = oRow
    oDS_PS_CO040L.setValue "LineId", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_CO040_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO040_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_CO040_FormClear()
On Error GoTo PS_CO040_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO040'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_CO040_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO040_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_CO040_EnableMenus()
On Error GoTo PS_CO040_EnableMenus_Error
    '//메뉴활성화 Navigation
    Call oForm01.EnableMenu("1288", True)
    Call oForm01.EnableMenu("1289", True)
    Call oForm01.EnableMenu("1290", True)
    Call oForm01.EnableMenu("1291", True)
    Call oForm01.EnableMenu("1293", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    '//Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_CO040_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO040_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_CO040_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_CO040_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_CO040_FormItemEnabled
        Call PS_CO040_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_CO040_FormItemEnabled
        oForm01.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_CO040_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO040_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_CO040_DataValidCheck() As Boolean
On Error GoTo PS_CO040_DataValidCheck_Error
    PS_CO040_DataValidCheck = False
    Dim i As Long
    If oForm01.Items("Code").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "코스트센터 그룹 코드는 필수입니다.", bmt_Short, True
        oForm01.Items("WhsCode").Click ct_Regular
        PS_CO040_DataValidCheck = False
        Exit Function
    End If
    
    If oForm01.Items("Name").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "코스트센터 그룹 이름은 필수입니다.", bmt_Short, True
        oForm01.Items("Name").Click ct_Regular
        PS_CO040_DataValidCheck = False
        Exit Function
    End If
    
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_CO040_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("CCCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "코스트센터 코드는 필수입니다.", bmt_Short, True
            oMat01.Columns("CCCode").Cells(i).Click ct_Regular
            PS_CO040_DataValidCheck = False
            Exit Function
        End If
    Next
   
    Call oDS_PS_CO040L.RemoveRecord(oDS_PS_CO040L.Size - 1)

    Call oMat01.LoadFromDataSource
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_CO040_FormClear
    End If
  
    PS_CO040_DataValidCheck = True
    Exit Function
PS_CO040_DataValidCheck_Error:
    PS_CO040_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_CO040_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PS_CO040_FormResize()
On Error GoTo PS_CO040_FormResize_Error
    
    Exit Sub
PS_CO040_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO040_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_CO040_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_CO040_Validate_Error
    PS_CO040_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "수정" Then
'        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
'        Dim Exist As Boolean
'        Exist = False
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            Exist = False
'            For j = 1 To oMat01.RowCount - 1
'                '//라인번호가 같고, 품목코드가 같으면 존재하는행 , LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
'                If Val(RecordSet01.Fields(1).Value) = Val(oMat01.Columns("U_LineNum").Cells(j).Specific.Value) And RecordSet01.Fields(2).Value = oMat01.Columns("1").Cells(j).Specific.Value And oMat01.Columns("U_LineNum").Cells(j).Specific.Value <> "" Then
'                    Exist = True
'                End If
'            Next
'            If (Exist = False) Then '//삭제된 행중
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "삭제된행이 다른사용자에 의해 출하,선출요청되었습니다. 적용할수 없습니다.", "W"
'                    PS_CO040_Validate = False
'                    GoTo PS_CO040_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
'        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
'            Else
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(oForm01.Items("8").Specific.Value) & "' AND U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할수 없습니다.", "W"
'                    PS_CO040_Validate = False
'                    GoTo PS_CO040_Validate_Exit
'                End If
'            End If
'        End If
    ElseIf ValidateType = "취소" Then
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
'                PS_CO040_Validate = False
'                GoTo PS_CO040_Validate_Exit
'            End If
'            RecordSet01.MoveNext
'        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_CO040_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_CO040_Validate_Error:
    PS_CO040_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_CO040_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



