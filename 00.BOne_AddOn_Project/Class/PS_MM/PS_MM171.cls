VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM171"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM171.cls
'//  Module         : ���Ű��� > ��ǥ�μ�
'//  Desc           : ��ü��� ��� �԰� ����Ʈ
'//  FormType       : PS_MM171
'//  Create Date    : 2010.08.31
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oBaseForm01 As SAPbouiCOM.Form '//�θ���
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_MM171H As SAPbouiCOM.DBDataSource '������
Private oDS_PS_MM171L As SAPbouiCOM.DBDataSource '��϶���

Private oLastItemUID01 As String 'Ŭ�������� ������ ������ ������ Uid��
Private oLastColUID01 As String '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLastColRow01 As Long '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

'//����ڱ���ü
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
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(ByRef oForm02 As SAPbouiCOM.Form, ByVal oItemUID02 As String, ByVal oColUID02 As String, ByVal oColRow02 As Long)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM171.srf")
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
    
    oFormUniqueID01 = "PS_MM171_" & GetTotalFormsCount
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
    
    Call PS_MM171_CreateItems
    Call PS_MM171_ComboBox_Setting
'    Call PS_MM171_CF_ChooseFromList
    Call PS_MM171_FormItemEnabled
    Call PS_MM171_EnableMenus
    '//Call PS_MM171_FormClear '//UDO����϶�
    '//Call PS_MM171_AddMatrixRow(0, True) '//UDO����϶�
    
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
            '//et_MATRIX_LINK_PRESSED
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
            '//et_CHOOSE_FROM_LIST
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
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
                '//Call PS_MM171_FormItemEnabled '//UDO���
            Case "1282": '�߰�
                '//Call PS_MM171_FormItemEnabled '//UDO���
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
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    Dim i As Long
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_MM171_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Btn02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_MM171_SetBaseForm
                '//�θ����� �Է��ϴ� �۾�
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_MM171" Then
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, "ItemCode", "") '//����ڰ�Ȱ��
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, "Mat01", "ItemCode") '//����ڰ�Ȱ��
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
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
Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    Dim oDataTable01  As SAPbouiCOM.DataTable
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        If (pval.ItemUID = "GRDocNum") Then
            Set oDataTable01 = pval.SelectedObjects
            oForm01.DataSources.UserDataSources("GRDocNum").Value = oDataTable01.Columns(0).Cells(0).Value
        End If
        oForm01.Update
    End If
    Set oDataTable01 = Nothing
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    If pval.BeforeAction = True Then
'        If pval.ItemChanged = True Then
'            If (pval.ItemUID = "Mat01") Then
'                Call oDS_PS_MM171L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'            Else
'                If (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_MM171H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_MM171H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_MM171H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                End If
'            End If
'            oMat01.LoadFromDataSource
'        End If
    ElseIf pval.BeforeAction = False Then

    End If
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_MM171_FormItemEnabled
        '//Call PS_MM171_AddMatrixRow(0, True) '//UDO���
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    If pval.BeforeAction = True Then
        '//������� ��������ɿ��ΰ˻�
    ElseIf pval.BeforeAction = False Then
'        For i = 1 To oMat01.VisualRowCount
'            oMat01.Columns("COL01").Cells(i).Specific.Value = i
'        Next i
'        oMat01.Clear
'        oMat01.FlushToDataSource
'        Call oDS_ZYM30L.RemoveRecord(oDS_ZYM30L.Size - 1)
'        oMat01.LoadFromDataSource
'        If oMat01.RowCount = 0 Then
'            Call PS_MM171_AddMatrixRow(0)
'        Else
'            If Trim(oDS_MM171L.GetValue("U_�����÷�", oMat01.RowCount - 1)) <> "" Then
'                Call PS_MM171_AddMatrixRow(oMat01.RowCount)
'            End If
'        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_MM171_CreateItems() As Boolean
On Error GoTo PS_MM171_CreateItems_Error
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '//Set oDS_PS_MM171H = oForm01.DataSources.DBDataSources("@PS_MM171H")
    Set oDS_PS_MM171L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 30)
    oForm01.Items("CardCode").Specific.DataBind.SetBound True, "", "CardCode"
    oForm01.DataSources.UserDataSources.Item("CardCode").Value = oBaseForm01.Items("CardCode").Specific.Value

    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 30)
    oForm01.Items("BPLId").Specific.DataBind.SetBound True, "", "BPLId"
    oForm01.DataSources.UserDataSources.Item("BPLId").Value = oBaseForm01.Items("BPLId").Specific.Value
        
    Set oRecordSet01 = Nothing
    Exit Function
PS_MM171_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_MM171_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_MM171_ComboBox_Setting()
On Error GoTo PS_MM171_ComboBox_Setting_Error
    '//�޺��� �⺻������
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_MM171", "Mat01", "ItemCode", "01", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_MM171", "Mat01", "ItemCode", "02", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PS_MM171", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_MM171", "ItemCode", "", "01", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_MM171", "ItemCode", "", "02", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PS_MM171", "ItemCode")
    
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Exit Sub
PS_MM171_ComboBox_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM171_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM171_CF_ChooseFromList()
On Error GoTo PS_MM171_CF_ChooseFromList_Error
    '//ChooseFromList ����
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

    Set oEdit = oForm01.Items("PODocNum").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = "20"
    oCFLCreationParams.uniqueID = "CFLPDNNO"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oEdit.ChooseFromListUID = "CFLPDNNO"
    oEdit.ChooseFromListAlias = "CardName"

'    Set oEdit = oForm01.Items("WhsCode").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = "64"
'    oCFLCreationParams.uniqueID = "CFLWHSCODE"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLWHSCODE"
'    oEdit.ChooseFromListAlias = "WhsCode"
    Exit Sub
PS_MM171_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM171_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM171_FormItemEnabled()
On Error GoTo PS_MM171_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
        '//Call PS_MM171_FormClear '//UDO���
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//����忡���� �����ۼ���
    End If
    Exit Sub
PS_MM171_FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM171_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM171_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_MM171_AddMatrixRow_Error
'    If RowIserted = False Then '//���߰�����
'        oDS_PS_MM171L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_MM171L.Offset = oRow
'    oDS_PS_MM171L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Exit Sub
PS_MM171_AddMatrixRow_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM171_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM171_FormClear()
On Error GoTo PS_MM171_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM171'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.Value = 1
    Else
        oForm01.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PS_MM171_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM171_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM171_EnableMenus()
On Error GoTo PS_MM171_EnableMenus_Error
    '//�޴�Ȱ��ȭ
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//�޴�����
    Exit Sub
PS_MM171_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM171_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_MM171_DataValidCheck() As Boolean
On Error GoTo PS_MM171_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "â���� �ʼ��Դϴ�.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_MM171_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "������ �������� �ʽ��ϴ�.", bmt_Short, True
'        PS_MM171_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "ǰ���� �ʼ��Դϴ�.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_MM171_DataValidCheck = False
'            Exit Function
'        End If
'    Next
    
    Call PS_MM171_FormClear
    Exit Function
PS_MM171_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM171_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_MM171_MTX01()
On Error GoTo PS_MM171_MTX01_Error
    '//��Ʈ������ ������ �ε�
    Dim i As Integer
    Dim j As Integer
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    
    Param01 = oForm01.Items("CardCode").Specific.Value
    Param02 = oForm01.Items("BPLId").Specific.Value
      
    Query01 = "EXEC PS_MM171_01 '" & Param01 & "','" & Param02 & "'"
    Call RecordSet01.DoQuery(Query01)
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "E")
        GoTo PS_MM171_MTX01_Exit
    End If
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", RecordSet01.RecordCount, False)
    
    j = 0
    
    For i = 0 To RecordSet01.RecordCount - 1
        If RecordSet01.Fields("RepayP").Value > 0 Then
        
            If j <> 0 Then
                oDS_PS_MM171L.InsertRecord (j)
            End If
            
            oDS_PS_MM171L.Offset = j
            Call oDS_PS_MM171L.setValue("U_LineNum", j, j + 1)
            Call oDS_PS_MM171L.setValue("U_ColReg01", j, False)
            Call oDS_PS_MM171L.setValue("U_ColReg02", j, RecordSet01.Fields("GRDocNum").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg03", j, RecordSet01.Fields("GRLinNum").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg04", j, RecordSet01.Fields("CardCode").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg05", j, RecordSet01.Fields("CardName").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg06", j, RecordSet01.Fields("ItemCode").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg07", j, RecordSet01.Fields("ItemName").Value)
            Call oDS_PS_MM171L.setValue("U_ColSum01", j, RecordSet01.Fields("LinTotal").Value)
            Call oDS_PS_MM171L.setValue("U_ColDt01", j, Format(RecordSet01.Fields("ImDate").Value, "YYYYMMDD"))
            Call oDS_PS_MM171L.setValue("U_ColDt02", j, Format(RecordSet01.Fields("DueDate").Value, "YYYYMMDD"))
            Call oDS_PS_MM171L.setValue("U_ColReg08", j, RecordSet01.Fields("LateDay").Value)
            Call oDS_PS_MM171L.setValue("U_ColSum02", j, RecordSet01.Fields("RepayP").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg09", j, RecordSet01.Fields("DocType").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg10", j, RecordSet01.Fields("CntcName").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg11", j, RecordSet01.Fields("PODocNum").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg12", j, RecordSet01.Fields("Unit").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg13", j, RecordSet01.Fields("Size").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg14", j, RecordSet01.Fields("Qty").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg15", j, RecordSet01.Fields("Unweight").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg16", j, RecordSet01.Fields("ItmBsort").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg17", j, RecordSet01.Fields("ItmMsort").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg18", j, RecordSet01.Fields("ItemType").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg19", j, RecordSet01.Fields("Quality").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg20", j, RecordSet01.Fields("Mark").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg21", j, RecordSet01.Fields("CallSize").Value)
            Call oDS_PS_MM171L.setValue("U_ColReg22", j, RecordSet01.Fields("ObasUnit").Value)
            j = j + 1
        End If
        
        RecordSet01.MoveNext
        ProgressBar01.Value = ProgressBar01.Value + 1
        ProgressBar01.Text = ProgressBar01.Value & "/" & RecordSet01.RecordCount & "�� ��ȸ��...!"
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Exit Sub
PS_MM171_MTX01_Exit:
    Set RecordSet01 = Nothing
    Exit Sub
PS_MM171_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_MM171_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_MM171_DI_API() As Boolean
On Error GoTo PS_MM171_DI_API_Error
    PS_MM171_DI_API = True
    Dim i, j As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim LineNumCount As Long
    Dim ResultDocNum As Long
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.Value
        ItemInformation(ItemInformationCount).LotNo = oMat01.Columns("LotNo").Cells(i).Specific.Value
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.Value
        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.Value
        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.Value
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = oForm01.Items("BPLId").Specific.Selected.Value
    oDIObject.CardCode = oForm01.Items("CardCode").Specific.Value
    '//oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.Value, "YYYY-MM-DD")
    For i = 0 To UBound(ItemInformation)
        If ItemInformation(i).Check = True Then
            GoTo Continue_First
        End If
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = oForm01.Items("WhsCode").Specific.Value
        oDIObject.Lines.BaseType = "22"
        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
        For j = i To UBound(ItemInformation)
            If ItemInformation(j).Check = True Then
                GoTo Continue_Second
            End If
            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
                GoTo Continue_Second
            End If
            '//������
            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).LotNo
            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.Add
            ItemInformation(j).PDN1No = LineNumCount
            ItemInformation(j).Check = True
Continue_Second:
        Next
        LineNumCount = LineNumCount + 1
Continue_First:
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To UBound(ItemInformation)
            Call oDS_PS_MM171L.setValue("U_OPDNNo", i, ResultDocNum)
            Call oDS_PS_MM171L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        Next
    Else
        GoTo PS_MM171_DI_API_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_MM171_DI_API_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_MM171_DI_API = False
    Set oDIObject = Nothing
    Exit Function
PS_MM171_DI_API_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM171_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_MM171_DI_API = False
End Function

Private Sub PS_MM171_SetBaseForm()
On Error GoTo PS_SM010_SetBaseForm_Error
    If oForm01.Mode = fm_ADD_MODE Then
        Dim i As Long
        Dim j As Long
        Dim ErrNum As Long
        Dim oMat02 As SAPbouiCOM.Matrix
        Set oMat02 = oBaseForm01.Items("Mat01").Specific
        For i = 1 To oMat01.VisualRowCount
            For j = 1 To oMat02.VisualRowCount
                If oMat01.Columns("GRDocNum").Cells(i).Specific.Value = oMat02.Columns("GRDocNum").Cells(j).Specific.Value Then
                    ErrNum = 1
                    GoTo PS_SM010_SetBaseForm_Error
                End If
            Next
                If oMat01.Columns("V_0").Cells(i).Specific.Checked = True Then
                    oMat02.Columns("GRDocNum").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("GRDocNum").Cells(i).Specific.Value
                    oMat02.Columns("GRLinNum").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("GRLinNum").Cells(i).Specific.Value
                    oMat02.Columns("PODocNum").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("PODocNum").Cells(i).Specific.Value
                    oMat02.Columns("ItemCode").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("ItemCode").Cells(i).Specific.Value
                    oMat02.Columns("ItemName").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("ItemName").Cells(i).Specific.Value
                    oMat02.Columns("LinTotal").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("LinTotal").Cells(i).Specific.Value
                    oMat02.Columns("ImDate").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("ImDate").Cells(i).Specific.Value
                    oMat02.Columns("DueDate").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("DueDate").Cells(i).Specific.Value
                    oMat02.Columns("LateDay").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("LateDay").Cells(i).Specific.Value
                    oMat02.Columns("Qty").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("Qty").Cells(i).Specific.Value
                    oMat02.Columns("Unit").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("Unit").Cells(i).Specific.Value
                    oMat02.Columns("Size").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("Size").Cells(i).Specific.Value
                                 
                    oMat02.Columns("Weight").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("Unweight").Cells(i).Specific.Value
                    oMat02.Columns("RepayP").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("RepayP").Cells(i).Specific.Value
                    Call oMat02.Columns("ItmBsort").Cells(oBaseColRow01).Specific.Select(oMat01.Columns("ItmBsort").Cells(i).Specific.Value, psk_ByValue)
                    Call oMat02.Columns("ItmMsort").Cells(oBaseColRow01).Specific.Select(oMat01.Columns("ItmMsort").Cells(i).Specific.Value, psk_ByValue)
                    Call oMat02.Columns("ItemType").Cells(oBaseColRow01).Specific.Select(oMat01.Columns("ItemType").Cells(i).Specific.Value, psk_ByValue)
                    Call oMat02.Columns("Quality").Cells(oBaseColRow01).Specific.Select(oMat01.Columns("Quality").Cells(i).Specific.Value, psk_ByValue)
                    Call oMat02.Columns("Mark").Cells(oBaseColRow01).Specific.Select(oMat01.Columns("Mark").Cells(i).Specific.Value, psk_ByValue)
                    oMat02.Columns("CallSize").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("CallSize").Cells(i).Specific.Value
                    Call oMat02.Columns("ObasUnit").Cells(oBaseColRow01).Specific.Select(oMat01.Columns("ObasUnit").Cells(i).Specific.Value, psk_ByValue)
        
                   
            
                    oBaseColRow01 = oBaseColRow01 + 1
                End If
        Next
        oForm01.Close
        '//�θ����� �Է��ϴ� �۾�
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
    ElseIf oForm01.Mode = fm_OK_MODE Then
    End If
    Exit Sub
PS_SM010_SetBaseForm_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�԰� ��ȣ�� �̹� �ֽ��ϴ�. Ȯ�� �� �缱�� �ϼ���! ", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PS_SM010_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub


