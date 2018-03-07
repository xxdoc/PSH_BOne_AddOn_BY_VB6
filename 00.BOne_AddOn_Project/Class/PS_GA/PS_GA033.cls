VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA033"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_GA033.cls
'//  Module             : �ѹ����� > ���� �� ������
'//  Desc               : ���� ��ȸ �� ���
'//  FormType           : PS_GA033
'//  Create Date(Start) : 2014.03.31
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oGrid01         As SAPbouiCOM.Grid

Private oDS_PS_GA033L  As SAPbouiCOM.DBDataSource '��϶���

Private oLastItemUID01 As String 'Ŭ�������� ������ ������ ������ Uid��
Private oLastColUID01  As String '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLastColRow01  As Long '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA033.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '��Ʈ������ Ÿ��Ʋ���̿� �����̸� ����
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_GA033_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO����϶�
        
    oForm01.Freeze True
    Call PS_GA033_CreateItems
    Call PS_GA033_ComboBox_Setting
    Call PS_GA033_Initial_Setting
    Call PS_GA033_CF_ChooseFromList
    Call PS_GA033_EnableMenus
    Call PS_GA033_SetDocument(oFromDocEntry01)
    Call PS_GA033_FormResize

'    Call PS_GA033_Add_MatrixRow(0, True)
    Call PS_GA033_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1287"), False        '// ����
    oForm01.EnableMenu ("1285"), False        '// ����
    oForm01.EnableMenu ("1284"), False         '// ���
    oForm01.EnableMenu ("1293"), False         '// �����
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

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

Public Sub PS_GA033_MTX01()
'******************************************************************************
'Function ID : PS_GA033_MTX01()
'�ش���    : PS_GA033
'���        : ������ ��ȸ
'�μ�        : ����
'��ȯ��      : ����
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo PS_GA033_MTX01_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID      As String '�����
    Dim StdYear    As String '���س⵵
    
    BPLID = oForm01.Items("BPLID").Specific.Selected.VALUE
    StdYear = oForm01.Items("StdYear").Specific.VALUE
        
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", RecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "      EXEC [PS_GA033_01] '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & StdYear & "'"
 
    Call oGrid01.DataTable.Clear
    Call oForm01.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
    ProgBar01.VALUE = 100
'    ProgBar01.Text = "��� ������...!"
    ProgBar01.Stop
    
'    oGrid01.Columns(2).RightJustified = True
    oGrid01.Columns(3).RightJustified = True
    oGrid01.Columns(4).RightJustified = True
    oGrid01.Columns(5).RightJustified = True
    oGrid01.Columns(6).RightJustified = True
    oGrid01.Columns(7).RightJustified = True
    oGrid01.Columns(8).RightJustified = True
    oGrid01.Columns(9).RightJustified = True
    oGrid01.Columns(10).RightJustified = True
    oGrid01.Columns(11).RightJustified = True
    oGrid01.Columns(12).RightJustified = True
    oGrid01.Columns(13).RightJustified = True
    oGrid01.Columns(14).RightJustified = True
    oGrid01.Columns(15).RightJustified = True
'    oGrid01.Columns(16).RightJustified = True
'    oGrid01.Columns(17).RightJustified = True
'    oGrid01.Columns(18).RightJustified = True
'    oGrid01.Columns(19).RightJustified = True
'    oGrid01.Columns(20).RightJustified = True
'    oGrid01.Columns(21).RightJustified = True
'    oGrid01.Columns(22).RightJustified = True
'    oGrid01.Columns(23).RightJustified = True
'    oGrid01.Columns(24).RightJustified = True
'    oGrid01.Columns(25).RightJustified = True
'    oGrid01.Columns(26).RightJustified = True
'    oGrid01.Columns(27).RightJustified = True
'    oGrid01.Columns(28).RightJustified = True
'    oGrid01.Columns(29).RightJustified = True
'    oGrid01.Columns(30).RightJustified = True
'    oGrid01.Columns(31).RightJustified = True
'    oGrid01.Columns(32).RightJustified = True
'    oGrid01.Columns(33).RightJustified = True
'    oGrid01.Columns(34).RightJustified = True
'    oGrid01.Columns(35).RightJustified = True
'    oGrid01.Columns(36).RightJustified = True
'    oGrid01.Columns(37).RightJustified = True
'    oGrid01.Columns(38).RightJustified = True
'    oGrid01.Columns(39).RightJustified = True
'    oGrid01.Columns(40).RightJustified = True
'    oGrid01.Columns(41).RightJustified = True
'    oGrid01.Columns(42).RightJustified = True
'    oGrid01.Columns(43).RightJustified = True
'    oGrid01.Columns(44).RightJustified = True
'    oGrid01.Columns(45).RightJustified = True
'    oGrid01.Columns(46).RightJustified = True
    
    If oGrid01.Rows.Count = 0 Then
        ErrNum = 1
        GoTo PS_GA033_MTX01_Error
    End If
        
    oGrid01.AutoResizeColumns
    oForm01.Update
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set RecordSet01 = Nothing
    
    Exit Sub

PS_GA033_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set RecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("��ȸ ����� �����ϴ�. Ȯ���ϼ���.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA033_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Function PS_GA033_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_GA033_HeaderSpaceLineDel()
'�ش���    : PS_GA033
'���        : �ʼ��Է»��� üũ
'�μ�        : ����
'��ȯ��      : True:�ʼ��Է»����� ��� �Է�, Fasle:�ʼ��Է»��� �� �ϳ��� �Է����� �ʾ���
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo PS_GA033_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
'        Case Trim(oForm01.Items("BPLID").Specific.VALUE) = "%" '�����
'            ErrNum = 1
'            GoTo PS_GA033_HeaderSpaceLineDel_Error '
    End Select

    PS_GA033_HeaderSpaceLineDel = True
    
    Exit Function

PS_GA033_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
'        MDC_Com.MDC_GF_Message "������� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
'        Call oForm01.Items("BPLID").Click(ct_Regular)
    End If
    PS_GA033_HeaderSpaceLineDel = False
End Function

Private Sub PS_GA033_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_GA033_FlushToItemValue_Error

    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim sQry     As String
    Dim ItemCode As String
    Dim Qty      As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID

        Case "AcctCode01"
            
            oForm01.Items("AcctName01").Specific.VALUE = MDC_GetData.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" & oForm01.Items("AcctCode01").Specific.VALUE & "'") '����

        Case "AcctCode02"
            
            oForm01.Items("AcctName02").Specific.VALUE = MDC_GetData.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" & oForm01.Items("AcctCode02").Specific.VALUE & "'") '��������
            
        Case "AcctCode03"
            
            oForm01.Items("AcctName03").Specific.VALUE = MDC_GetData.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" & oForm01.Items("AcctCode03").Specific.VALUE & "'") '���ΰ�������
        
    End Select

    Set oRecordSet01 = Nothing

    Exit Sub
    
PS_GA033_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("���� �޽���", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA033_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/���� ������ �������
Sub PS_GA033_FormItemEnabled()
On Error GoTo PS_GA033_FormItemEnabled_Error

    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    
    Exit Sub

PS_GA033_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_GA033_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/������ ���� �̺�Ʈ
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

Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
            Case "1282": '�߰�
                '/�߰���ư Ŭ���� ��Ʈ���� insertrow
                
'                Call PS_GA033_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
            
            Case "7169": '���� ��������
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
                '//Call PS_GA033_FormItemEnabled '//UDO���
            Case "1282": '�߰�
'                oMat01.Clear
'                oDS_PS_GA033H.Clear

'                Call PS_GA033_LoadCaption
'                Call PS_GA033_FormItemEnabled
                '//Call PS_GA033_FormItemEnabled '//UDO���
                '//Call PS_GA033_AddMatrixRow(0, True) '//UDO���
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                '//Call PS_GA033_FormItemEnabled
                
            Case "7169": '���� ��������
                
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
    
        If pval.ItemUID = "PS_GA033" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
            
        If pval.ItemUID = "BtnSearch" Then '������ȸ
        
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_GA033_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        
        ElseIf pval.ItemUID = "BtnPrint1" Then '�Ϲ� ���
        
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_GA033_Print_Report01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA033" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode01", "") '//����ڰ�Ȱ��
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode02", "") '//����ڰ�Ȱ��
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode03", "") '//����ڰ�Ȱ��
        
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//����ڰ�Ȱ��
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PS_GA033_FlushToItemValue(pval.ItemUID)
    
    End If

    Exit Sub
    
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
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

    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
        
            Call PS_GA033_FlushToItemValue(pval.ItemUID)
                
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
        Call PS_GA033_FormItemEnabled
        '//Call PS_GA033_AddMatrixRow(oMat01.VisualRowCount) '//UDO���
    End If
    
    Exit Sub
    
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error

    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_GA033_FormResize
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
'            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_GA033H", "U_CardCode,U_CardName")
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
'            If (PS_GA033_Validate("�����") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//������� ��������ɿ��ΰ˻�
            
        ElseIf pval.BeforeAction = False Then
'            For i = 1 To oMat01.VisualRowCount
'                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
'            Next i
'            oMat01.FlushToDataSource
'            Call oDS_PS_GA033L.RemoveRecord(oDS_PS_GA033L.Size - 1)
'            oMat01.LoadFromDataSource
            
'            If oMat01.RowCount = 0 Then
'                Call PS_GA033_Add_MatrixRow(0)
'            Else
'                If Trim(oDS_PS_GA033L.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_GA033_Add_MatrixRow(oMat01.RowCount)
'                End If
'            End If

        End If
    End If
    
    Exit Sub
    
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_GA033_CreateItems() As Boolean
On Error GoTo PS_GA033_CreateItems_Error

    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    
    '�����_S
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
    
    '���س⵵
    Call oForm01.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 4)
    Call oForm01.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")
   
    Call oForm01.Freeze(False)
    
    Exit Function
    
PS_GA033_CreateItems_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA033_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/�޺��ڽ� set
Sub PS_GA033_ComboBox_Setting()
On Error GoTo PS_GA033_ComboBox_Setting_Error

    Dim oCombo  As SAPbouiCOM.ComboBox
    Dim sQry    As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '//////////�⺻����//////////S
    '�����
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
    
PS_GA033_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA033_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA033_Initial_Setting()
On Error GoTo PS_GA033_Initial_Setting_Error
    '����� ������� �Ҽ� ����� ����
    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 
    '��¥ ����
'    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYY") + "1231"
'    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYY") + "0101"
    oForm01.Items("StdYear").Specific.VALUE = Format(Now, "YYYY")
    
    Exit Sub
    
PS_GA033_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA033_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA033_CF_ChooseFromList()
On Error GoTo PS_GA033_CF_ChooseFromList_Error

    Exit Sub
    
PS_GA033_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA033_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA033_EnableMenus()
On Error GoTo PS_GA033_EnableMenus_Error
    
    Exit Sub
    
PS_GA033_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA033_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA033_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_GA033_SetDocument_Error

    If (oFromDocEntry01 = "") Then
        Call PS_GA033_FormItemEnabled
        '//Call PS_GA033_AddMatrixRow(0, True) '//UDO����϶�
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_GA033_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    
    Exit Sub
    
PS_GA033_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA033_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA033_FormResize()
On Error GoTo PS_GA033_FormResize_Error
    
    Exit Sub
    
PS_GA033_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA033_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA033_Print_Report01()
On Error GoTo Print_Query_Error

    Dim DocNum      As String
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim sQry        As String

    Dim i           As Integer
    Dim ErrNum      As Integer
    Dim Sub_sQry    As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    Dim BPLID      As String '�����
    Dim StdYear    As String '���س⵵
    
    BPLID = oForm01.Items("BPLID").Specific.Selected.VALUE
    StdYear = oForm01.Items("StdYear").Specific.VALUE
    
    WinTitle = "[PS_GA033] ����Ʈ"
    
    ReportName = "PS_GA033_01.rpt"
    
    '����
    sQry = "      EXEC [PS_GA033_02] '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & StdYear & "'"
    
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
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", oRecordSet.RecordCount, False)

'    Call oRecordSet.DoQuery(sQry)
'
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query�� �������� 'Y'��...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
'    ProgBar01.Text = "��� ������...!"
    ProgBar01.Stop
    
    Set ProgBar01 = Nothing
    Set oRecordSet = Nothing
    
    Exit Sub

Print_Query_Error:
    Set ProgBar01 = Nothing
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_GA033_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub





