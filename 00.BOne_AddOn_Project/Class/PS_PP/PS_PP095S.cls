VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP095S"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP095S.cls
'//  Module         : MG�ŷ�����ǥ��Ͻ� ȣ��
'//  Desc           : MG�ŷ�����ǥ���
'//  FormType       : PS_PP095S
'//  Create Date    : 2013.09.26
'//  Copyright  (c) Morning Data
'//********************************************************************************

'//�����
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_SM010
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)


Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix
Private oDS_PS_PP095SH As SAPbouiCOM.DBDataSource '������
Private oDS_PS_PP095SL As SAPbouiCOM.DBDataSource '��϶���



Public oBaseForm01 As SAPbouiCOM.Form '//�θ���
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseBPLId01 As String

Private oLastItemUID01 As String 'Ŭ�������� ������ ������ ������ Uid��
Private oLastColUID01 As String '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLastColRow01 As Long '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oMat01Row01 As Long
Private oMat02Row02 As Long

Private sSeq As Long

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
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oFromBPLId02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP095S.srf")
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
    
    oFormUniqueID01 = "PS_PP095S_" & GetTotalFormsCount
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
    
    oBaseBPLId01 = oFromBPLId02
    
    Call PS_PP095S_CreateItems
    Call PS_PP095S_ComboBox_Setting
    Call PS_PP095S_CF_ChooseFromList
    Call PS_PP095S_FormItemEnabled
    Call PS_PP095S_EnableMenus
    '//Call PS_PP095S_FormClear '//UDO����϶�
    '//Call PS_PP095S_AddMatrixRow(0, True) '//UDO����϶�
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
    'oForm01.Items("Button01").Click ct_Regular
    
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
                '//Call PS_PP095S_FormItemEnabled '//UDO���
            Case "1282": '�߰�
                '//Call PS_PP095S_FormItemEnabled '//UDO���
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
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_PP095S_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_PP095S_SetBaseForm '//�θ������Է�
                Call oForm01.Close
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_PP095S_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//�ؾ����� �۾�
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP095S" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP095S_FormItemEnabled
'                    Call PS_PP095S_FormClear '//UDO����϶�
'                    Call PS_PP095S_AddMatrixRow(oMat01.RowCount, True) '//UDO����϶�
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP095S_FormItemEnabled
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//����ڰ�Ȱ��
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//����ڰ�Ȱ��
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//����ڰ�Ȱ��
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemName", "") '//����ڰ�Ȱ��
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
        If pval.ItemUID = "Mat01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row > 0 Then
                    Call oMat01.SelectRow(pval.Row, True, False)
                    oMat01Row01 = pval.Row
'                    oMat01.SelectionMode = ms_Auto
                    
                    
'                    If GetItem_ManBtchNum(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) = "Y" Then '//��ġ�� ����ϴ�ǰ��
'                        Call PS_PP095S_MTX02
'                    ElseIf GetItem_ManBtchNum(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) = "N" Then '//��ġ�� ������� �ʴ�ǰ��
'                        Call PS_PP095S_MTX02
'                    End If
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        
'        If pval.ItemUID = "Opt01" Then
'            oForm01.Settings.MatrixUID = "Mat01"
'            oForm01.Settings.Enabled = True
'            oForm01.Settings.EnableRowFormat = True
'        End If
'        If pval.ItemUID = "Opt02" Then
'            oForm01.Settings.MatrixUID = "Mat02"
'            oForm01.Settings.Enabled = True
'            oForm01.Settings.EnableRowFormat = True
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "Mat01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row > 0 Then
'                    If oMat01.Columns("CHK").Cells(pval.Row).Specific.Checked = True Then
'                        Call oMat01.SelectRow(pval.Row, True, True)
'                    Else
'                        Call oMat01.SelectRow(pval.Row, True, False)
'                    End If
'                    If (pval.ColUID = "CHK") Then
'                        If oMat01.Columns("CHK").Cells(pval.Row).Specific.Checked = True Then
'                            oMat01.Columns("ChkSeq").Cells(pval.Row).Specific.VALUE = sSeq
'                            sSeq = sSeq + 1
'                        Else
'                            oMat01.Columns("ChkSeq").Cells(pval.Row).Specific.VALUE = 999
'                        End If
'                    End If
'                    If oMat01.Columns("CHK").Cells(pval.Row).Specific.VALUE
                    
'                    If GetItem_ManBtchNum(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) = "Y" Then '//��ġ�� ����ϴ�ǰ��
'                        Call PS_PP095S_MTX02
'                    ElseIf GetItem_ManBtchNum(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) = "N" Then '//��ġ�� ������� �ʴ�ǰ��
'                        Call PS_PP095S_MTX02
'                    End If
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
    
    Dim sQry  As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If pval.ItemUID = "ItemCode" Then
               sQry = "Select ItemName From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.String) & "'"
               oRecordSet.DoQuery sQry
               oForm01.Items("ItemName").Specific.String = Trim(oRecordSet.Fields(0).VALUE)
            
            End If
            If pval.ItemUID = "Mat01" Then
                
            
            End If
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
        Call PS_PP095S_FormItemEnabled
        '//Call PS_PP095S_AddMatrixRow(oMat01.VisualRowCount) '//UDO���
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

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    Dim oDataTable01  As SAPbouiCOM.DataTable
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "ItemCode") Then
'            Set oDataTable01 = pval.SelectedObjects
'            If oDataTable01 Is Nothing Then
'            Else
'                oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'                oForm01.DataSources.UserDataSources("ItemName").Value = oDataTable01.Columns(1).Cells(0).Value
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
            '//������� ��������ɿ��ΰ˻�
        ElseIf pval.BeforeAction = False Then
    '        For i = 1 To oMat01.VisualRowCount
    '            oMat01.Columns("COL01").Cells(i).Specific.Value = i
    '        Next i
    '        oMat01.FlushToDataSource
    '        Call oDS_PS_PP095SL.RemoveRecord(oDS_PS_PP095SL.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_PP095S_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_�����÷�", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_PP095S_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP095S_CreateItems() As Boolean
On Error GoTo PS_PP095S_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP095SH = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    
    
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
   ' Call oForm01.DataSources.UserDataSources.Add("CpCode", dt_SHORT_TEXT, 10)
   ' Call oForm01.Items("CpCode").Specific.DataBind.SetBound(True, "", "CpCode")
        
    
    oForm01.Items("Mat01").Enabled = False
    
    sSeq = 1 'Ŭ������
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP095S_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP095S_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP095S_ComboBox_Setting()
On Error GoTo PS_PP095S_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//�޺��� �⺻������
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP095S", "Mat01", "ItemCode", "01", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP095S", "Mat01", "ItemCode", "02", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP095S", "Mat01", "ItemCode")
    
    'Call oForm01.Items("BPLId").Specific.ValidValues.Add("����", "����")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL ORDER BY BPLId", "", False, False)
    
   ' Call oForm01.Items("CpCode").Specific.ValidValues.Add("����", "����")
   ' Call MDC_SetMod.Set_ComboList(oForm01.Items("CpCode").Specific, "SELECT U_CpCode, U_CpName FROM [@PS_PP001L] Where U_ItmBSort = '104' ORDER BY U_CpCode", "", False, False)
    
   
    
    Call oForm01.Items("BPLId").Specific.Select(oBaseBPLId01, psk_ByDescription)
   ' Call oForm01.Items("CpCode").Specific.Select(oBaseCpCode01, psk_ByDescription)
    
        
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP095S_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP095S_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP095S_CF_ChooseFromList()
On Error GoTo PS_PP095S_CF_ChooseFromList_Error
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
PS_PP095S_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095S_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP095S_FormItemEnabled()
On Error GoTo PS_PP095S_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
        '//Call PS_PP095S_FormClear '//UDO���
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//����忡���� �����ۼ���
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP095S_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP095S_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP095S_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP095S_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//���߰�����
'        oDS_PS_PP095SL.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_PP095SL.Offset = oRow
'    oDS_PS_PP095SL.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP095S_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP095S_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP095S_FormClear()
On Error GoTo PS_PP095S_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP095S'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP095S_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095S_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP095S_EnableMenus()
On Error GoTo PS_PP095S_EnableMenus_Error
    '//�޴�Ȱ��ȭ
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//�޴�����
    Exit Sub
PS_PP095S_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095S_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP095S_DataValidCheck() As Boolean
On Error GoTo PS_PP095S_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "â���� �ʼ��Դϴ�.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_PP095S_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "������ �������� �ʽ��ϴ�.", bmt_Short, True
'        PS_PP095S_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "ǰ���� �ʼ��Դϴ�.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_PP095S_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_PP095S_FormClear
    Exit Function
PS_PP095S_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095S_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP095S_MTX01()
On Error GoTo PS_PP095S_MTX01_Error
    '//��Ʈ������ ������ �ε�
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    
    Param01 = oForm01.Items("BPLId").Specific.Selected.VALUE
    Param02 = oForm01.Items("ItemCode").Specific.VALUE
    
    Query01 = "EXEC PS_PP095S_01 '" & Param01 & "','" & Param02 & "'"
    
'    If oBaseForm01 Is Nothing Then
'        Query01 = "EXEC PS_PP095S_01 '" & Param01 & "','','','" & Param02 & "','','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "'"
'    ElseIf oBaseForm01.Type = "149" Or oBaseForm01.Type = "139" Or oBaseForm01.Type = "140" Or oBaseForm01.Type = "180" Or oBaseForm01.Type = "133" Or oBaseForm01.Type = "179" Or oBaseForm01.Type = "60091" Then
'        Query01 = "EXEC PS_PP095S_01 '" & Param01 & "','Y','','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "'" '//�Ǹ�Y,����,���Ÿ��(1:����ִ°͸�,2:��ü),�ŷ�Ÿ��(1:�Ϲ�,2:�Ӱ���)
'    ElseIf oBaseForm01.Type = "142" Or oBaseForm01.Type = "143" Or oBaseForm01.Type = "182" Or oBaseForm01.Type = "141" Or oBaseForm01.Type = "181" Or oBaseForm01.Type = "60092" Then
'        Query01 = "EXEC PS_PP095S_01 '" & Param01 & "','','Y','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "'" '//�Ǹ�,����Y
'    Else
'        Query01 = "EXEC PS_PP095S_01 '" & Param01 & "','','','" & Param02 & "','','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "'"
'    End If
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    
    If (RecordSet01.RecordCount = 0) Then
        oForm01.Items("Mat01").Enabled = False
        Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "W")
        GoTo PS_PP095S_MTX01_Exit
    Else
        oForm01.Items("Mat01").Enabled = True
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_PP095SH.InsertRecord (i)
        End If
        oDS_PS_PP095SH.Offset = i
        Call oDS_PS_PP095SH.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP095SH.setValue("U_ColReg01", i, "N")
        Call oDS_PS_PP095SH.setValue("U_ColReg02", i, RecordSet01.Fields("PackNo").VALUE)
    '    Call oDS_PS_PP095SH.setValue("U_ColReg03", i, RecordSet01.Fields("LotNo").VALUE)
        Call oDS_PS_PP095SH.setValue("U_ColReg04", i, RecordSet01.Fields("ItemCode").VALUE)
        Call oDS_PS_PP095SH.setValue("U_ColReg05", i, RecordSet01.Fields("ItemName").VALUE)
        Call oDS_PS_PP095SH.setValue("U_ColReg06", i, RecordSet01.Fields("Weight").VALUE)
        Call oDS_PS_PP095SH.setValue("U_ColReg07", i, RecordSet01.Fields("ProDate").VALUE)
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "�� ��ȸ��...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP095S_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP095S_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP095S_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP095S_SetBaseForm()
On Error GoTo PS_PP095S_SetBaseForm_Error
    Dim i As Long
    Dim j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim S_Weight As Double
    
    Dim PackNo As String
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    If oBaseForm01.TypeEx = "PS_PP095" Then
        Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//��Ʈ����
'        For i = 1 To oMat01.RowCount '//ǰ����ǰ��
'            If oMat01.Columns("CHK").Cells(i).Specific.Checked = True Then
'                If (Val(oMat01.Columns("SelQty").Cells(i).Specific.Value) <= 0) Then
'                    '//������ ���õ��� ����ǰ��
'                Else
'                    oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.Value = oMat01.Columns("ItemCode").Cells(i).Specific.Value '//ǰ��
'                    oBaseMat01.Columns("OutWhCd").Cells(oBaseColRow01).Specific.Value = oMat02.Columns("WhsCode").Cells(oMat01Row01).Specific.Value '//���â��
'                    oBaseMat01.Columns("Qty").Cells(oBaseColRow01).Specific.Value = Val(oMat01.Columns("SelQty").Cells(i).Specific.Value) '//���� '//������ �����ϸ� �߷��� �ڵ�����ȴ�.
'                    oBaseMat01.Columns("Unweight").Cells(oBaseColRow01).Specific.Value = Val(MDC_PS_Common.GetItem_UnWeight(oMat01.Columns("ItemCode").Cells(i).Specific.Value)) '//����
''                    oBaseMat01.Columns("14").Cells(oBaseColRow01).Specific.Value = MDC_PS_Common.GetValue("EXEC PS_SBO_GETPRICE '" & oBaseForm01.Items("4").Specific.Value & "','" & oMat01.Columns("ItemCode").Cells(i).Specific.Value & "'", 0, 1)
'                    oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01 + 1).Click ct_Regular
'                    oBaseColRow01 = oBaseColRow01 + 1
'                End If
'            End If
'        Next
        
        
                
        For i = 1 To oMat01.RowCount '//��ġ����ǰ��
            If oMat01.Columns("CHK").Cells(i).Specific.Checked = True Then
'                If (Val(oMat01.Columns("SelQty").Cells(i).Specific.VALUE) <= 0) Then
'                    '//������ ���õ��� ����ǰ��
'                Else
                PackNo = oMat01.Columns("PackNo").Cells(i).Specific.VALUE
                
                Query01 = " SELECT LotNo    = b.U_LotNo"
                Query01 = Query01 + " FROM [@PS_PP090H] a INNER JOIN [@PS_PP090L] b ON a.DocEntry = b.DocEntry AND a.CanCeled = 'N'"
                Query01 = Query01 + " WHERE a.U_BPLId    = '" + oForm01.Items("BPLId").Specific.VALUE + "'"
                Query01 = Query01 + " AND a.U_PackNo = '" + PackNo + "'"
                
                Call RecordSet01.DoQuery(Query01)
                
                For j = 0 To RecordSet01.RecordCount - 1
                    oBaseMat01.Columns("OrdNum").Cells(oBaseColRow01).Specific.VALUE = RecordSet01.Fields(0).VALUE
                    
                    'oBaseMat01.Columns("OrdNum").Cells(oBaseColRow01).Specific.VALUE = oMat01.Columns("OrdNum").Cells(i).Specific.VALUE '//ǰ��
                    
                    RecordSet01.MoveNext
                    oBaseColRow01 = oBaseColRow01 + 1
                Next
                    
'                End If
            End If
        Next
        
        For i = 1 To oBaseMat01.RowCount
            S_Weight = S_Weight + oBaseMat01.Columns("Weight").Cells(i).Specific.VALUE
        Next
        
        oBaseForm01.Items("S_Weight").Specific.VALUE = S_Weight
        
        
    End If
    Exit Sub
PS_PP095S_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095S_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub




