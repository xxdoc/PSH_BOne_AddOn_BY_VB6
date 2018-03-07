VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FX210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_FX210.cls
'//  Module         : ����Ʈ > �����ڻ� FX
'//  Desc           : �����ڻ��̷���Ȳ
'//  FormType       : PS_FX210
'//  Create Date    : 2012.08.31
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_FX210H As SAPbouiCOM.DBDataSource '������
Private oDS_PS_FX210L As SAPbouiCOM.DBDataSource '��϶���

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
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FX210.srf")
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
    
    oFormUniqueID01 = "PS_FX210_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO����϶�
        
    oForm01.Freeze True
    Call PS_FX210_CreateItems
    Call PS_FX210_ComboBox_Setting
    Call PS_FX210_Initial_Setting
    Call PS_FX210_CF_ChooseFromList
    Call PS_FX210_EnableMenus
    Call PS_FX210_SetDocument(oFromDocEntry01)
    Call PS_FX210_FormResize
    
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
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
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
                If PS_FX210_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_FX210_Print_Report01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_FX210" Then
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "BPLId", "") '//����ڰ�Ȱ��
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//����ڰ�Ȱ��
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '//����ڰ�Ȱ��
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//����ڰ�Ȱ��
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
            If (pval.ItemUID = "BPLId") Then
                oQuery01 = "SELECT BPLName, BPLId FROM [OBPL] WHERE BPLid = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("BPLName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
          '  ElseIf (pval.ItemUID = "CardCode") Then
          '      oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
          '      oRecordSet01.DoQuery oQuery01
          '      oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
          '  ElseIf (pval.ItemUID = "ItemCode") Then
          '     oQuery01 = "SELECT ItemName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
          '      oRecordSet01.DoQuery oQuery01
          '      oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
          '  ElseIf (pval.ItemUID = "CntcCode") Then
          '      oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
          '      oRecordSet01.DoQuery oQuery01
          '      oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
            oForm01.Items(pval.ItemUID).Click ct_Regular
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
        Call PS_FX210_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_FX210_FormResize
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


Private Function PS_FX210_CreateItems() As Boolean
On Error GoTo PS_FX210_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 100)
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 100)
    Call oForm01.DataSources.UserDataSources.Add("HisType", dt_SHORT_TEXT, 100)
    
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    Call oForm01.Items("DocDateFr").Specific.DataBind.SetBound(True, "", "DocDateFr")
    Call oForm01.Items("DocDateTo").Specific.DataBind.SetBound(True, "", "DocDateTo")
    Call oForm01.Items("HisType").Specific.DataBind.SetBound(True, "", "HisType")
    
    oForm01.Items("DocDateFr").Specific.VALUE = Format(Now, "YYYY") + "0101"
    
    oForm01.Items("DocDateTo").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_FX210_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX210_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_FX210_ComboBox_Setting()
On Error GoTo PS_FX210_ComboBox_Setting_Error
     Call oForm01.Freeze(True)

        Call oForm01.Items("BPLId").Specific.ValidValues.Add("", "")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
        Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        
        Call oForm01.Items("HisType").Specific.ValidValues.Add("%", "��ü")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("HisType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'", "", False, False)
        Call oForm01.Items("HisType").Specific.Select(0, psk_Index)

    Call oForm01.Freeze(False)
    Exit Sub
    
PS_FX210_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX210_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX210_Initial_Setting()
On Error GoTo PS_FX210_Initial_Setting_Error
'  �����
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 ' �μ���
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD()
    Exit Sub
PS_FX210_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX210_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_FX210_CF_ChooseFromList()
On Error GoTo PS_FX210_CF_ChooseFromList_Error
    Exit Sub
PS_FX210_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX210_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX210_FormItemEnabled()
On Error GoTo PS_FX210_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_FX210_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX210_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX210_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_FX210_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    Call oForm01.Freeze(False)
    Exit Sub
PS_FX210_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX210_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX210_FormClear()
On Error GoTo PS_FX210_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_FX210'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_FX210_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX210_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX210_EnableMenus()
On Error GoTo PS_FX210_EnableMenus_Error
    Exit Sub
PS_FX210_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX210_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX210_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_FX210_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_FX210_FormItemEnabled

    Else

    End If
    Exit Sub
PS_FX210_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX210_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_FX210_DataValidCheck() As Boolean
On Error GoTo PS_FX210_DataValidCheck_Error
    PS_FX210_DataValidCheck = False
    Dim i As Long

    If oForm01.Items("BPLId").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "������� �ʼ��Դϴ�.", bmt_Short, True
        oForm01.Items("BPLId").Click ct_Regular
        PS_FX210_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("DocDateFr").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "�̷����� Fr�� �ʼ��Դϴ�.", bmt_Short, True
        oForm01.Items("DocDateFr").Click ct_Regular
        PS_FX210_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("DocDateTo").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "�̷����� To�� �ʼ��Դϴ�.", bmt_Short, True
        oForm01.Items("DocDateTo").Click ct_Regular
        PS_FX210_DataValidCheck = False
        Exit Function
    End If
    PS_FX210_DataValidCheck = True
    Exit Function
PS_FX210_DataValidCheck_Error:
    PS_FX210_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_FX210_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PS_FX210_FormResize()
On Error GoTo PS_FX210_FormResize_Error
    
    Exit Sub
PS_FX210_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX210_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_FX210_Print_Report01()
On Error GoTo PS_FX210_Print_Report01_Error
    Dim ErrNum          As Integer
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim sQry01                      As String
    Dim BPLId                       As String
    Dim DocDateFr                   As String
    Dim DocDateTo                   As String
    Dim HisType                     As String
    Dim oRecordSet        As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    
    '// ��ȸ���ǹ�
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    HisType = Trim(oForm01.Items("HisType").Specific.VALUE)
    
    If HisType = "" Then HisType = "%"
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "�����ڻ��̷���Ȳ [PS_FX210]"
    ReportName = "PS_FX210_01.rpt"
    
    '// Formula �����ʵ�
    ReDim gRpt_Formula(3)
    ReDim gRpt_Formula_Value(3)
    
    gRpt_Formula(1) = "BPLId"
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" & BPLId & "'"
    Call oRecordSet.DoQuery(sQry)
    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
    
    gRpt_Formula(2) = "DocDateFr"
    gRpt_Formula_Value(2) = Format(DocDateFr, "0000-00-00")
    
    gRpt_Formula(3) = "DocDateTo"
    gRpt_Formula_Value(3) = Format(DocDateTo, "0000-00-00")
    
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure ����"
    sQry = "EXEC [PS_FX210_01] '" & BPLId & "', '" & DocDateFr & "', '" & DocDateTo & "', '" & HisType & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PS_FX210_Print_Report01_Error
    End If

    '/ Action (sub_query�� �������� 'Y'��...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_FX210_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX210_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



