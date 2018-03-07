VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FI190"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_FI190.cls
'//  Module         : FI
'//  Desc           : ����������ȸ(EXCEL�ٿ�)
'//  FormType       : PS_FI190
'//  Create Date    : 2014.11.21
'//********************************************************************************

'//�����
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_FI190
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid

Public oBaseForm01 As SAPbouiCOM.Form '//�θ���
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String
Public oBaseItmBsort01 As String

Private oLastItemUID01 As String 'Ŭ�������� ������ ������ ������ Uid��
Private oLastColUID01 As String '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLastColRow01 As Long '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

'//����ڱ���ü
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
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String, Optional ByVal oItmBsort02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FI190.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '��Ʈ������ Ÿ��Ʋ���̿� �����̸� ����
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_FI190_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO����϶�
        
    oForm01.Freeze True
    Set oBaseForm01 = oForm02
    oBaseItemUID01 = oItemUID02
    oBaseColUID01 = oColUID02
    oBaseColRow01 = oColRow02
    oBaseTradeType01 = oTradeType02
    oBaseItmBsort01 = oItmBsort02
    
    Call PS_FI190_CreateItems
    Call PS_FI190_ComboBox_Setting
    Call PS_FI190_CF_ChooseFromList
    Call PS_FI190_FormItemEnabled
    Call PS_FI190_EnableMenus
    '//Call PS_FI190_FormClear '//UDO����϶�
    '//Call PS_FI190_AddMatrixRow(0, True) '//UDO����϶�
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    
    '���� Ŀ����ġ
    oForm01.ActiveItem = "BPLId"
    
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
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
                '//Call PS_FI190_FormItemEnabled '//UDO���
            Case "1282": '�߰�
                '//Call PS_FI190_FormItemEnabled '//UDO���
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
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
'            MenuCreationParams01.String = "�޴���"
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
               Call PS_FI190_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button02" Then
            Call oForm01.Close

        End If

    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_FI190" Then
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//����ڰ�Ȱ��
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//����ڰ�Ȱ��
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//����ڰ�Ȱ��
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemName", "") '//����ڰ�Ȱ��
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
'        If pval.ItemUID = "ItmBsort" Then
'            Dim i As Long
'            For i = 0 To oForm01.Items("ItmMsort").Specific.ValidValues.Count - 1
'                Call oForm01.Items("ItmMsort").Specific.ValidValues.Remove(0, psk_Index)
'            Next
'            Call oForm01.Items("ItmMsort").Specific.ValidValues.Add("%", "����")
'            Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" & oForm01.Items("ItmBsort").Specific.Selected.VALUE & "' ORDER BY U_Code", "", False, False)
'            If oForm01.Items("ItmMsort").Specific.ValidValues.Count > 0 Then
'                Call oForm01.Items("ItmMsort").Specific.Select(0, psk_Index)
'            End If
'        End If
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
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
'    If pval.BeforeAction = True Then
'        If pval.ItemUID = "Grid01" Then
'            If pval.Row = -1 Then
''                oGrid01.Columns(pval.ColUID).TitleObject.Sortable = True
'
'            Else
'                If oGrid01.Rows.SelectedRows.Count > 0 Then
'                    Call PS_FI190_SetBaseForm '//�θ������Է�
'                    If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
'                        Call oForm01.Close
'                    End If
'                Else
'                    BubbleEvent = False
'                End If
'            End If
'        End If
'    ElseIf pval.BeforeAction = False Then
'
'    End If
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
        Call PS_FI190_FormItemEnabled
        '//Call PS_FI190_AddMatrixRow(oMat01.VisualRowCount) '//UDO���
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
'        If (pval.ItemUID = "ItemCode") Then
'            Set oDataTable01 = pval.SelectedObjects
'            If oDataTable01 Is Nothing Then
'            Else
'                oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
'                oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
'            End If
'        End If
'        oForm01.Update
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
            '//������� ��������ɿ��ΰ˻�
        ElseIf pval.BeforeAction = False Then
    '        For i = 1 To oMat01.VisualRowCount
    '            oMat01.Columns("COL01").Cells(i).Specific.Value = i
    '        Next i
    '        oMat01.FlushToDataSource
    '        Call oDS_PS_FI190L.RemoveRecord(oDS_PS_FI190L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_FI190_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_�����÷�", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_FI190_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_FI190_CreateItems() As Boolean
On Error GoTo PS_FI190_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    'oGrid01.SelectionMode = ms_NotSupported
    
    '�Ⱓ
    Call oForm01.DataSources.UserDataSources.Add("DocDatefr", dt_DATE, 10)
    Call oForm01.Items("DocDatefr").Specific.DataBind.SetBound(True, "", "DocDatefr")
    oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYYMM") + "01"
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateto", dt_DATE, 10)
    Call oForm01.Items("DocDateto").Specific.DataBind.SetBound(True, "", "DocDateto")
    oForm01.DataSources.UserDataSources.Item("DocDateto").VALUE = Format(Now, "YYYYMMDD")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_FI190_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI190_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_FI190_ComboBox_Setting()
On Error GoTo PS_FI190_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//�޺��� �⺻������

    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL]", "", False, False)
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
'    Call oForm01.Items("Div").Specific.ValidValues.Add("1", "������ҹ�αݾ׿���")
'    Call oForm01.Items("Div").Specific.ValidValues.Add("2", "��ǥ��ȣ�����ڷ�")
'    Call oForm01.Items("Div").Specific.ValidValues.Add("3", "������������ �󰢺а���ȸ")
'    Call oForm01.Items("Div").Specific.ValidValues.Add("4", "�������������ȸ")
'    Call oForm01.Items("Div").Specific.ValidValues.Add("5", "�����԰�����(â��)")
'    Call oForm01.Items("Div").Specific.Select(0, psk_Index)
'
    Call oForm01.Freeze(False)
    Exit Sub
PS_FI190_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI190_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FI190_CF_ChooseFromList()
On Error GoTo PS_FI190_CF_ChooseFromList_Error
    '//ChooseFromList ����
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
PS_FI190_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI190_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FI190_FormItemEnabled()
On Error GoTo PS_FI190_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
'        If oBaseItmBsort01 = "" Then
'
'        Else
'            Call oForm01.Items("ItmBsort").Specific.Select(oBaseItmBsort01, psk_ByValue)
'            oForm01.Items("ItmBsort").Enabled = False
'        End If
        '//����忡���� �����ۼ���
        '//Call PS_FI190_FormClear '//UDO���
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//����忡���� �����ۼ���
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_FI190_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI190_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FI190_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_FI190_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//���߰�����
'        oDS_PS_FI190L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_FI190L.Offset = oRow
'    oDS_PS_FI190L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_FI190_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI190_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FI190_FormClear()
On Error GoTo PS_FI190_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_FI190'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_FI190_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI190_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FI190_EnableMenus()
On Error GoTo PS_FI190_EnableMenus_Error
    '//�޴�Ȱ��ȭ
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//�޴�����
    Exit Sub
PS_FI190_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI190_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_FI190_DataValidCheck() As Boolean
On Error GoTo PS_FI190_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "â���� �ʼ��Դϴ�.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_FI190_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "������ �������� �ʽ��ϴ�.", bmt_Short, True
'        PS_FI190_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "ǰ���� �ʼ��Դϴ�.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_FI190_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_FI190_FormClear
    Exit Function
PS_FI190_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI190_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_FI190_MTX01()
On Error GoTo PS_FI190_MTX01_Error
    '//��Ʈ������ ������ �ε�
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String
    Dim DocDateFr As String
    Dim DocDateTo As String
    Dim Div As String
    
    BPLID = oForm01.Items("BPLId").Specific.Selected.VALUE
    DocDateFr = oForm01.Items("DocDatefr").Specific.VALUE
    DocDateTo = oForm01.Items("DocDateto").Specific.VALUE
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", 100, False)
    
    Query01 = "EXEC PS_FI190_01 '" & BPLID & "','" & DocDateFr & "','" & DocDateTo & "'"
    
    oGrid01.DataTable.Clear
    
    Call oForm01.DataSources.DataTables.Item("DataTable").ExecuteQuery(Query01)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
    oGrid01.Columns(24).RightJustified = True
    oGrid01.Columns(25).RightJustified = True
    oGrid01.Columns(26).RightJustified = True
    oGrid01.Columns(27).RightJustified = True
    oGrid01.Columns(28).RightJustified = True
    oGrid01.Columns(29).RightJustified = True
    
    If oGrid01.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "E")
        GoTo PS_FI190_MTX01_Exit
    End If
        
    oGrid01.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_FI190_MTX01_Exit:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_FI190_MTX01_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI190_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_FI190_DI_API() As Boolean
'On Error GoTo PS_FI190_DI_API_Error
'    PS_FI190_DI_API = True
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
'    oDIObject.BPL_IDAssignedToInvoice = oForm01.Items("BPLId").Specific.Selected.Value
'    oDIObject.CardCode = oForm01.Items("CardCode").Specific.Value
'    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.Value, "&&&&-&&-&&")
'    For i = 0 To UBound(ItemInformation)
'        If ItemInformation(i).Check = True Then
'            GoTo Continue_First
'        End If
'        If i <> 0 Then
'            oDIObject.Lines.Add
'        End If
'        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
'        oDIObject.Lines.WarehouseCode = oForm01.Items("WhsCode").Specific.Value
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
'            '//������
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
'            Call oDS_PS_FI190L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_FI190L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
'    Else
'        GoTo PS_FI190_DI_API_Error
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
'PS_FI190_DI_API_DI_Error:
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
'    PS_FI190_DI_API = False
'    Set oDIObject = Nothing
'    Exit Function
'PS_FI190_DI_API_Error:
'    Sbo_Application.SetStatusBarMessage "PS_FI190_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    PS_FI190_DI_API = False
End Function




