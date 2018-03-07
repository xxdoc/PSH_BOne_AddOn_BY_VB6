VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FI215"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File : PS_FI215.cls
'//  Module : �繫����>�ֵ��>
'//  Desc : ä�� ������ ����
'//  FormType : PS_FI215
'//  Create Date : 2012.08.21
'//  Modifier : Song Myounggyu
'//  Modify Date :
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_FI215H As SAPbouiCOM.DBDataSource '������
Private oDS_PS_FI215L As SAPbouiCOM.DBDataSource '��϶���
'Private oDS_PS_USERDS01 As SAPbouiCOM.DBDataSource    '��϶���

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
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FI215.srf")
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
    
    oFormUniqueID01 = "PS_FI215_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO����϶�
        
    oForm01.Freeze True
    Call PS_FI215_CreateItems
    Call PS_FI215_ComboBox_Setting
    Call PS_FI215_Initial_Setting
    Call PS_FI215_CF_ChooseFromList
    Call PS_FI215_EnableMenus
    Call PS_FI215_SetDocument(oFromDocEntry01)
    Call PS_FI215_FormResize
    
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
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
            
            Case "7169": '���� ��������
            
                '���� �������� ���� �� ��Ʈ������ ���� ������ �࿡ �� �� �߰�
                Call PS_FI215_AddMatrixRow(oMat01.VisualRowCount)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
            
            Case "7169": '���� ��������
            
                '���� �������� ���� ó��
                Call oForm01.Freeze(True)
                Call oDS_PS_FI215L.RemoveRecord(oDS_PS_FI215L.Size - 1)
                Call oMat01.LoadFromDataSource
                Call oForm01.Freeze(False)
            
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
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_FI215_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    Call PS_FI215_MTX01 '��Ʈ������ ������ �ε�
                End If
            End If
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_FI215_Print_Report01
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
        ElseIf pval.ItemUID = "Btn_Print" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
        
                Call PS_FI215_Print_Report01
            
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_FI215" Then
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
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '�ŷ�ó�ڵ� ���˼�ġ Ȱ��
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") 'ǰ���ڵ�(�۹�) ���˼�ġ Ȱ��
        
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
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                
                Call oMat01.SelectRow(pval.Row, True, False)
            
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
            End If
        Else
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = ""
            oLastColRow01 = 0
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row = 0 Then
            
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource
            
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
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "CardCode") Then
                oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf (pval.ItemUID = "ItemCode") Then
                oQuery01 = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf (pval.ItemUID = "CntcCode") Then
                oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
            
            oForm01.Items(pval.ItemUID).Click ct_Regular
        End If
    ElseIf pval.BeforeAction = False Then

    End If
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_FI215_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_FI215_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
    
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
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_FI215_CreateItems() As Boolean
On Error GoTo PS_FI215_CreateItems_Error
    Call oForm01.Freeze(True)
'    Dim oQuery01 As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Set oDS_PS_FI215H = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    Set oDS_PS_FI215L = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    
    '��Ʈ���� �ʱ�ȭ
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '�����_S
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    '�����_E
    
    '������ ����_S
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    '������ ����_E

    '������ ����_S
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    '������ ����_E
    
    '�ŷ�ó�ڵ�_S
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CardCode").Specific.DataBind.SetBound(True, "", "CardCode")
    '�ŷ�ó�ڵ�_E
    
    '�ŷ�ó��_S
    Call oForm01.DataSources.UserDataSources.Add("CardName", dt_SHORT_TEXT, 200)
    Call oForm01.Items("CardName").Specific.DataBind.SetBound(True, "", "CardName")
    '�ŷ�ó��_E
    
    'AR��������_S
    Call oForm01.DataSources.UserDataSources.Add("DocStatus", dt_SHORT_TEXT, 10)
    Call oForm01.Items("DocStatus").Specific.DataBind.SetBound(True, "", "DocStatus")
    'AR��������_E
    
    Call oForm01.Freeze(False)
    Exit Function
PS_FI215_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI215_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_FI215_ComboBox_Setting()
On Error GoTo PS_FI215_ComboBox_Setting_Error
    
    Call oForm01.Freeze(True)
     
        '����� �޺��ڽ� ����_S
        Call oForm01.Items("BPLId").Specific.ValidValues.Add("%", "��ü")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
        Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        '����� �޺��ڽ� ����_S
        
        '���ֹ������� ����_S
        Call oForm01.Items("DocStatus").Specific.ValidValues.Add("%", "��ü")
        Call oForm01.Items("DocStatus").Specific.ValidValues.Add("O", "�̰�")
        Call oForm01.Items("DocStatus").Specific.ValidValues.Add("C", "����")
        Call oForm01.Items("DocStatus").Specific.Select(0, psk_Index) '�⺻���� ��ü ����
        '���ֹ������� ����_E
        
        '��Ʈ���� �÷��� �޺��ڽ� ����_S
        Dim oColumn As SAPbouiCOM.Column
'        Dim lRecordSet As SAPbobsCOM.Recordset '���ڵ��
'        Dim lQuery As String '������Ʈ��
'
'        Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

        '��������
        Set oColumn = oMat01.Columns("DocStatus")
        Call oColumn.ValidValues.Add("O", "�̰�")
        Call oColumn.ValidValues.Add("C", "����")
        '��������
        '��Ʈ���� �÷��� �޺��ڽ� ����_E
        
        
    Call oForm01.Freeze(False)
    Exit Sub
PS_FI215_ComboBox_Setting_Error:
'    Set lRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI215_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FI215_Initial_Setting()
On Error GoTo PS_FI215_Initial_Setting_Error
    '����� ������� �Ҽ� ����� ����
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '��¥ ����
    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYYMMDD")
    
    '��¥ ��ĭ���� ����
'    oForm01.Items("ToDt").Specific.VALUE = ""
'    oForm01.Items("FrDt").Specific.VALUE = ""
 
    Exit Sub
PS_FI215_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI215_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_FI215_CF_ChooseFromList()
On Error GoTo PS_FI215_CF_ChooseFromList_Error
    Exit Sub
PS_FI215_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI215_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FI215_FormItemEnabled()
On Error GoTo PS_FI215_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_FI215_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI215_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FI215_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_FI215_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//���߰�����
        oDS_PS_FI215L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_FI215L.Offset = oRow
'    oDS_PS_FI215L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_FI215_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI215_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FI215_FormClear()
On Error GoTo PS_FI215_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_FI215'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_FI215_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI215_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FI215_EnableMenus()
On Error GoTo PS_FI215_EnableMenus_Error
    Exit Sub
PS_FI215_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI215_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FI215_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_FI215_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_FI215_FormItemEnabled

    Else

    End If
    Exit Sub
PS_FI215_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI215_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_FI215_DataValidCheck() As Boolean
On Error GoTo PS_FI215_DataValidCheck_Error
    PS_FI215_DataValidCheck = False
    Dim i As Long

'    If oForm01.Items("DocDateTo").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "�������� To�� �ʼ��Դϴ�.", bmt_Short, True
'        oForm01.Items("DocDateTo").Click ct_Regular
'        PS_FI215_DataValidCheck = False
'        Exit Function
'    End If

    PS_FI215_DataValidCheck = True
    Exit Function
PS_FI215_DataValidCheck_Error:
    PS_FI215_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_FI215_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_FI215_MTX01()
On Error GoTo PS_FI215_MTX01_Error
    '//��Ʈ������ ������ �ε�
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId As String '�����
    Dim FrDt As String '�����Ͻ���
    Dim ToDt As String '����������
    Dim CardCode As String '�ŷ�ó
    Dim DocStatus As String '��������
    
    BPLId = oForm01.Items("BPLId").Specific.Selected.VALUE
    FrDt = oForm01.Items("FrDt").Specific.VALUE
    ToDt = oForm01.Items("ToDt").Specific.VALUE
    CardCode = oForm01.Items("CardCode").Specific.VALUE
    DocStatus = IIf(oForm01.Items("DocStatus").Specific.Selected.VALUE = "%", "", oForm01.Items("DocStatus").Specific.Selected.VALUE)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", RecordSet01.RecordCount, False) '������ ������ �� ���� ���α׷��� ����
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_FI215_01 '" & BPLId & "','" & FrDt & "','" & ToDt & "','" & CardCode & "','" & DocStatus & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_FI215_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_FI215L.InsertRecord(loopCount)
        End If
        oDS_PS_FI215L.Offset = loopCount
        
        Call oDS_PS_FI215L.setValue("U_LineNum", loopCount, loopCount + 1) '���ι�ȣ
        Call oDS_PS_FI215L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("DocEntry").VALUE) 'AR�����ȣ
        Call oDS_PS_FI215L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("DocDate").VALUE) '������
        Call oDS_PS_FI215L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("DueDate").VALUE) '������
        Call oDS_PS_FI215L.setValue("U_ColReg12", loopCount, RecordSet01.Fields("TaxDate").VALUE) '������
        Call oDS_PS_FI215L.setValue("U_ColReg13", loopCount, RecordSet01.Fields("PayDate").VALUE) '���޿�����
        Call oDS_PS_FI215L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("CardCode").VALUE) '�ŷ�ó�ڵ�
        Call oDS_PS_FI215L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("CardName").VALUE) '�ŷ�ó��
        Call oDS_PS_FI215L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("Currency").VALUE) '��ȭ
        Call oDS_PS_FI215L.setValue("U_ColSum01", loopCount, RecordSet01.Fields("LineTotal").VALUE) '�ݾ�
        Call oDS_PS_FI215L.setValue("U_ColSum02", loopCount, RecordSet01.Fields("VatSum").VALUE) '�ΰ���
        Call oDS_PS_FI215L.setValue("U_ColSum03", loopCount, RecordSet01.Fields("Total").VALUE) '�Ѱ�
        Call oDS_PS_FI215L.setValue("U_ColPrc01", loopCount, RecordSet01.Fields("TotalFC").VALUE) '�Ѱ�(��ȭ)
        Call oDS_PS_FI215L.setValue("U_ColReg07", loopCount, RecordSet01.Fields("ReceiptsDt").VALUE) '�Ա�����
        Call oDS_PS_FI215L.setValue("U_ColReg08", loopCount, RecordSet01.Fields("DelayDay").VALUE) '�����ϼ�
        Call oDS_PS_FI215L.setValue("U_ColSum04", loopCount, RecordSet01.Fields("Receipts").VALUE) '�Աݾ�
        Call oDS_PS_FI215L.setValue("U_ColPrc02", loopCount, RecordSet01.Fields("ReceiptsFC").VALUE) '�Աݾ�(��ȭ)
        Call oDS_PS_FI215L.setValue("U_ColSum05", loopCount, RecordSet01.Fields("AdjAmt").VALUE) '�����ݾ�
        Call oDS_PS_FI215L.setValue("U_ColPrc03", loopCount, RecordSet01.Fields("AdjAmtFC").VALUE) '�����ݾ�(��ȭ)
        Call oDS_PS_FI215L.setValue("U_ColSum06", loopCount, RecordSet01.Fields("RecTotal").VALUE) 'ȸ���ݾ��Ѱ�
        Call oDS_PS_FI215L.setValue("U_ColPrc04", loopCount, RecordSet01.Fields("RecTotalFC").VALUE) 'ȸ���ݾ��Ѱ�(��ȭ)
        Call oDS_PS_FI215L.setValue("U_ColReg09", loopCount, RecordSet01.Fields("PayMth").VALUE) '���޼���
        Call oDS_PS_FI215L.setValue("U_ColReg10", loopCount, RecordSet01.Fields("BoeDueDate").VALUE) '����������
        Call oDS_PS_FI215L.setValue("U_ColReg11", loopCount, RecordSet01.Fields("DocStatus").VALUE) 'AR���幮������
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "�� ��ȸ��...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_FI215_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "W")
    Exit Sub
PS_FI215_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI215_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_FI215_FormResize()
On Error GoTo PS_FI215_FormResize_Error
    
    Exit Sub
PS_FI215_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI215_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FI215_Print_Report01()
On Error GoTo Print_Query_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String

    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim Sub_sQry        As String
    
'    Dim BPLId           As String
'    Dim CardCode       As String
'    Dim Pumtxt        As String


    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    '// ��ȸ���ǹ�
    
'    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
'    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
'    Pumtxt = Trim(oForm01.Items("Pumtxt").Specific.VALUE)
'
'    If Pumtxt = "" Then Pumtxt = "%"
'    If CardCode = "" Then CardCode = "%"
'
    
    Dim BPLId As String '�����
    Dim ItemClass As String 'ǰ�񱸺�
    Dim TradeType As String '�ŷ�����
    Dim FrDt As String '�����Ͻ���
    Dim ToDt As String '����������
    Dim CardCode As String '�ŷ�ó
    Dim ItemCode As String 'ǰ���ڵ�(�۹�)
    Dim DocStatus As String '��������
    Dim Chk01 As String '�����
    Dim Chk02 As String '�̳�ǰ
    
    BPLId = oForm01.Items("BPLId").Specific.Selected.VALUE
    ItemClass = IIf(oForm01.Items("ItemClass").Specific.Selected.VALUE = "%", "", oForm01.Items("ItemClass").Specific.Selected.VALUE)
    TradeType = IIf(oForm01.Items("TradeType").Specific.Selected.VALUE = "%", "", oForm01.Items("TradeType").Specific.Selected.VALUE)
    FrDt = oForm01.Items("FrDt").Specific.VALUE
    ToDt = oForm01.Items("ToDt").Specific.VALUE
    CardCode = oForm01.Items("CardCode").Specific.VALUE
    ItemCode = oForm01.Items("ItemCode").Specific.VALUE
    DocStatus = IIf(oForm01.Items("DocStatus").Specific.Selected.VALUE = "%", "", oForm01.Items("DocStatus").Specific.Selected.VALUE)
    Chk01 = IIf(oForm01.Items("Chk01").Specific.Checked = True, "1", "0")
    Chk02 = IIf(oForm01.Items("Chk02").Specific.Checked = True, "1", "0")
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_FI215] ����Ʈ"
    ReportName = "PS_FI215.rpt"
 
    '// Formula �����ʵ�
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure ����"
    sQry = "EXEC PS_FI215_01 '" & BPLId & "','" & ItemClass & "','" & TradeType & "','" & FrDt & "','" & ToDt & "','" & CardCode & "','" & ItemCode & "','" & DocStatus & "','" & Chk01 & "','" & Chk02 & "'"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query�� �������� 'Y'��...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub



