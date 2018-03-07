VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD059"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_SD059.cls
'//  Module : �ǸŰ���>�ֵ��>�������
'//  Desc : ����Ű��������
'//  FormType : PS_SD059
'//  Create Date : 2015.02.28
'//  Modified Date : 2015.02.28
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_SD059H As SAPbouiCOM.DBDataSource '������
Private oDS_PS_SD059L As SAPbouiCOM.DBDataSource '��϶���

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

Private oDocType01 As String
Private oDocEntry01 As String
Private oFormMode01 As SAPbouiCOM.BoFormMode

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(Optional ByVal prmManageNo As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD059.srf")
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
    
    oFormUniqueID01 = "PS_SD059_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO����϶�
        
    oForm01.Freeze True
    Call PS_SD059_CreateItems
    Call PS_SD059_ComboBox_Setting
    Call PS_SD059_CF_ChooseFromList
    Call PS_SD059_EnableMenus
    Call PS_SD059_SetDocument(prmManageNo)
    Call PS_SD059_FormResize
    
    oForm01.EnableMenu ("1283"), True         '// ����
    oForm01.EnableMenu ("1287"), True         '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1284"), False        '// ���
    oForm01.EnableMenu ("1293"), True         '// �����
    
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
                Call PS_SD059_FormItemEnabled '//UDO���
            Case "1282": '�߰�
                Call PS_SD059_FormItemEnabled '//UDO���
                Call PS_SD059_Initial_Setting
'                Call PS_SD059_AddMatrixRow(0, True) '//UDO���
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                Call PS_SD059_FormItemEnabled
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

    Dim SavedPointer As Integer

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "PS_SD059" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_SD059_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_SD059_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_SD059" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_SD059_FormItemEnabled
'                    Call PS_SD059_AddMatrixRow(0, True) '//UDO����϶�
                    Call PS_SD059_Initial_Setting
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_SD059_FormItemEnabled
                End If
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
    
        If pval.ItemUID = "Mat01" Then
        
        ElseIf pval.ItemUID = "CardCode" Then

            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '�ؿܰŷ�ó ���˼�ġ ����

        ElseIf pval.ItemUID = "CurrAmt" Then

            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CurrAmt", "") '�ݾ���ȭ ���˼�ġ ����
        

        ElseIf pval.ItemUID = "Currency" Then

            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Currency", "") '��ȭ��ȭ ���˼�ġ ����
        
        End If
        
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
    
        If pval.ItemChanged = True Then
            
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
'        If pval.ItemUID = "Mat01" Then
'            If pval.Row > 0 Then
'                Call oMat01.SelectRow(pval.Row, True, False)
'            End If
'        End If
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat01.SelectRow(pval.Row, True, False)
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
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
    
        Call PS_SD059_FlushToItemValue(pval.ItemUID)
    
    End If

    Exit Sub
    
Raise_EVENT_VALIDATE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_SD059_FormItemEnabled
        Call PS_SD059_AddMatrixRow(oMat01.VisualRowCount) '//UDO���
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_SD059_FormResize
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
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//������� ��������ɿ��ΰ˻�
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_SD059L.RemoveRecord(oDS_PS_SD059L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_SD059_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_SD059L.GetValue("U_Code", oMat01.RowCount - 1)) <> "" Then
                    Call PS_SD059_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_SD059_CreateItems() As Boolean
On Error GoTo PS_SD059_CreateItems_Error
   
    Set oDS_PS_SD059H = oForm01.DataSources.DBDataSources("@PS_SD059H")
'    Set oDS_PS_SD059L = oForm01.DataSources.DBDataSources("@PS_SD059L")
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
'    oMat01.SelectionMode = ms_NotSupported
'    oMat01.AutoResizeColumns
    
    Exit Function
PS_SD059_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD059_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SD059_ComboBox_Setting()
On Error GoTo PS_SD059_ComboBox_Setting_Error
        
    '//�޺��� �⺻������
    'Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    'Dim oRecordSet01 As SAPbobsCOM.Recordset
        
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD059", "Mat01", "UseYN", "Y", "Y")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD059", "Mat01", "UseYN", "N", "N")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("UseYN"), "PS_SD059", "Mat01", "UseYN")
    
    '���������_S
    sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ExptAgt").Specific, sQry, MDC_PS_Common.User_BPLId, False, False)
    '���������_E
    
    '����ȭ��_S
    sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ExptOwn").Specific, sQry, MDC_PS_Common.User_BPLId, False, False)
    '����ȭ��_E
    
    '������_S
    sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Maker").Specific, sQry, MDC_PS_Common.User_BPLId, False, False)
    '������_E
    
'    '�з�_S
'    sQry = "            SELECT      U_Code,"
'    sQry = sQry & "                 U_CodeNm"
'    sQry = sQry & "  FROM       [@PS_GA050L]"
'    sQry = sQry & "  WHERE      Code = '12'"
'    sQry = sQry & "                 AND U_UseYN = 'Y'"
'    sQry = sQry & "  ORDER BY  U_Seq"
'    Call oForm01.Items("Ctgr").Specific.ValidValues.Add("%", "����")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("Ctgr").Specific, sQry, "%", False, False)
'    '�з�_E
    
'    Set oCombo = Nothing
'    Set oRecordSet01 = Nothing

    Exit Sub
PS_SD059_ComboBox_Setting_Error:
'    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD059_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD059_CF_ChooseFromList()
On Error GoTo PS_SD059_CF_ChooseFromList_Error
    Exit Sub
PS_SD059_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD059_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD059_FormItemEnabled()
On Error GoTo PS_SD059_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
'
        oForm01.Items("DocEntry").Enabled = False
'        oForm01.Items("MngNo").Enabled = True
'        oForm01.Items("Mat01").Enabled = True
        Call PS_SD059_FormClear '//UDO���
        Call oForm01.EnableMenu("1281", True) '//ã��
        Call oForm01.EnableMenu("1282", False) '//�߰�
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
'        oForm01.Items("MngNo").Enabled = True
'        oForm01.Items("Mat01").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//ã��
        Call oForm01.EnableMenu("1282", True) '//�߰�
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//����忡���� �����ۼ���
        
        oForm01.Items("DocEntry").Enabled = False
'        oForm01.Items("MngNo").Enabled = False
'        oForm01.Items("Mat01").Enabled = True
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD059_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD059_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD059_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD059_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//���߰�����
        oDS_PS_SD059L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD059L.Offset = oRow
    oDS_PS_SD059L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD059_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD059_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD059_FormClear()
On Error GoTo PS_SD059_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD059'", "")
    
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SD059_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD059_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD059_EnableMenus()
On Error GoTo PS_SD059_EnableMenus_Error
    '//�޴�Ȱ��ȭ
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//�޴�����
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//�޴�����
    Exit Sub
PS_SD059_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD059_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD059_Initial_Setting()
On Error GoTo PS_SD059_Initial_Setting_Error

    '���������
    Call oForm01.Items("ExptAgt").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

    '����ȭ��
    Call oForm01.Items("ExptOwn").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '������
    Call oForm01.Items("Maker").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

'    '�������з�
'    Call oForm01.Items("Ctgr").Specific.Select("%", psk_ByValue)
'
'    '������
'    Call oForm01.Items("Maker").Specific.Select("%", psk_ByValue)
'
'    'OS
'    Call oForm01.Items("OS").Specific.Select("%", psk_ByValue)
'
'    '���Կ���
'    Call oForm01.Items("PchsYN").Specific.Select("%", psk_ByValue)

    Exit Sub
PS_SD059_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD059_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD059_SetDocument(ByVal prmManageNo As String)
On Error GoTo PS_SD059_SetDocument_Error
    If (prmManageNo = "") Then
        Call PS_SD059_FormItemEnabled
'        Call PS_SD059_AddMatrixRow(0, True) '//UDO����϶�
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_SD059_FormItemEnabled
'        oForm01.Items("MngNo").Specific.VALUE = prmManageNo '������ȣ
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_SD059_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD059_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD059_DataValidCheck() As Boolean
On Error GoTo PS_SD059_DataValidCheck_Error
    PS_SD059_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_SD059_FormClear
    End If
    
'    '�з� �̼��� ��
'    If oForm01.Items("Ctgr").Specific.VALUE = "%" Then
'        Call Sbo_Application.SetStatusBarMessage("�з��� ���õ��� �ʾҽ��ϴ�.", bmt_Short, True)
'        PS_SD059_DataValidCheck = False
'        Exit Function
'    End If
'
'    '����� �̼��� ��
'    If oForm01.Items("BPLId").Specific.VALUE = "%" Then
'        Call Sbo_Application.SetStatusBarMessage("������� ���õ��� �ʾҽ��ϴ�.", bmt_Short, True)
'        PS_SD059_DataValidCheck = False
'        Exit Function
'    End If
'
'    '������ȣ ���Է� ��
'    If oForm01.Items("MngNo").Specific.VALUE = "" Then
'        Call Sbo_Application.SetStatusBarMessage("������ȣ�� �Էµ��� �ʾҽ��ϴ�.", bmt_Short, True)
'        PS_SD059_DataValidCheck = False
'        Exit Function
'    End If
'
'    '�𵨸� ���Է� ��
'    If oForm01.Items("ModelNm").Specific.VALUE = "" Then
'        Call Sbo_Application.SetStatusBarMessage("�𵨸��� �Էµ��� �ʾҽ��ϴ�.", bmt_Short, True)
'        PS_SD059_DataValidCheck = False
'        Exit Function
'    End If
'
'    '������ �̼��� ��
'    If oForm01.Items("Maker").Specific.VALUE = "%" Then
'        Call Sbo_Application.SetStatusBarMessage("�����簡 ���õ��� �ʾҽ��ϴ�.", bmt_Short, True)
'        PS_SD059_DataValidCheck = False
'        Exit Function
'    End If
'
'    '���Կ��� �̼��� ��
'    If oForm01.Items("PchsYN").Specific.VALUE = "%" Then
'        Call Sbo_Application.SetStatusBarMessage("���Կ��ΰ� ���õ��� �ʾҽ��ϴ�.", bmt_Short, True)
'        PS_SD059_DataValidCheck = False
'        Exit Function
'    End If
    
'    '�������� ���Է� ��
'    If oMat01.VisualRowCount = 1 Then
'        Sbo_Application.SetStatusBarMessage "������ �������� �ʽ��ϴ�.", bmt_Short, True
'        PS_SD059_DataValidCheck = False
'        Exit Function
'    End If
'
'    For i = 1 To oMat01.VisualRowCount - 1
'        If (oMat01.Columns("Code").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "�ڵ�� �ʼ��Դϴ�.", bmt_Short, True
'            oMat01.Columns("Code").Cells(i).Click ct_Regular
'            PS_SD059_DataValidCheck = False
'            Exit Function
'        End If
'
'        If (oMat01.Columns("CodeNm").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "�ڵ���� �ʼ��Դϴ�.", bmt_Short, True
'            oMat01.Columns("CodeNm").Cells(i).Click ct_Regular
'            PS_SD059_DataValidCheck = False
'            Exit Function
'        End If
'    Next

'    Call oMat01.FlushToDataSource
'    Call oDS_PS_SD059L.RemoveRecord(oDS_PS_SD059L.Size - 1)
'    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_SD059_FormClear
    End If
    
    PS_SD059_DataValidCheck = True
    Exit Function
PS_SD059_DataValidCheck_Error:
    PS_SD059_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_SD059_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD059_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_SD059_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        
        Case "CardCode"
        
        oForm01.Items("CardName").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'") '�ؿܰŷ�ó��
        
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_SD059_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_SD059_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

Private Sub PS_SD059_MTX01()
On Error GoTo PS_SD059_MTX01_Error
    '//��Ʈ������ ������ �ε�
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Param01 = oForm01.Items("Param01").Specific.VALUE
    Param02 = oForm01.Items("Param01").Specific.VALUE
    Param03 = oForm01.Items("Param01").Specific.VALUE
    Param04 = oForm01.Items("Param01").Specific.VALUE
    
    Query01 = "SELECT 10"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "E")
        GoTo PS_SD059_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_SD059L.InsertRecord (i)
        End If
        oDS_PS_SD059L.Offset = i
        Call oDS_PS_SD059L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_SD059L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_SD059_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_SD059_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD059_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_SD059_FormResize()
On Error GoTo PS_SD059_FormResize_Error
    
    Exit Sub
PS_SD059_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD059_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD059_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_SD059_Validate_Error
    PS_SD059_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "����" Then
        '//������ ���� ã�Ƽ� �������ɼ� �˻� , ���� �Էµ����� ������ �Ұ����ϵ��� ������ �ʿ��ϴٸ� �������� ã�±��� ����
    ElseIf ValidateType = "�����" Then
        '//������� ��������ɿ��ΰ˻�
    ElseIf ValidateType = "���" Then
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_SD059_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_SD059_Validate_Error:
    PS_SD059_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_SD059_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
