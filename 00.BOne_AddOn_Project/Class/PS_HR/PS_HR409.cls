VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR409"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_HR409.cls
'//  Module         : HR
'//  Description    : ��й�ȣ����
'//  FormType       : PS_HR409
'//  Create Date    : 2012.12.02
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_HR409H As SAPbouiCOM.DBDataSource '������
Private oDS_PS_HR409L As SAPbouiCOM.DBDataSource '��϶���

Private oLastItemUID01 As String 'Ŭ�������� ������ ������ ������ Uid��
Private oLastColUID01 As String '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLastColRow01 As Long '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��
Private ItemInformationCount As Long
'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR409.srf")
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
    
    oFormUniqueID01 = "PS_HR409_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    'oForm01.DataBrowser.BrowseBy = "Code" '//UDO����϶�

    oForm01.EnableMenu ("1293"), False     '// �����
    oForm01.EnableMenu ("1287"), False     '// ����
    oForm01.EnableMenu ("1284"), False     '// ���
    oForm01.Freeze True
    
    Call CreateItems
    Call ComboBox_Setting
    
    
    'Call AddMatrixRow(0, True)
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
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim Cnt             As Integer
    Dim BPLId           As String
    Dim MSTCOD          As String
    Dim FULLNAME        As String
    Dim PassWd          As String
    Dim PassWd1         As String
    Dim PassWd2         As String
    Dim vReturnValue$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                       
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "MSTCOD" Then
                        If oForm01.Items("MSTCOD").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    BPLId = oForm01.Items("BPLId").Specific.VALUE
                    MSTCOD = oForm01.Items("MSTCOD").Specific.VALUE
                    FULLNAME = oForm01.Items("FULLNAME").Specific.VALUE
                    PassWd = oForm01.Items("PassWd").Specific.VALUE
                    PassWd1 = oForm01.Items("ChPWd1").Specific.VALUE
                    PassWd2 = oForm01.Items("ChPWd2").Specific.VALUE
                    
                    
                    sQry = " Select Count(*) From Z_PS_HRPASS Where BPLId = '" & BPLId & "' and MSTCOD = '" & MSTCOD & "'"
                    
                    oRecordSet01.DoQuery sQry
                    
                    Cnt = oRecordSet01.Fields(0).VALUE
                    
                    If Cnt > 0 Then '
                        '//������ Update
                        
                        sQry = " Select PassWd From Z_PS_HRPASS Where BPLId = '" & BPLId & "' and MSTCOD = '" & MSTCOD & "'"
                    
                        oRecordSet01.DoQuery sQry
                        
                        If oRecordSet01.Fields(0).VALUE <> PassWd Then
                            Sbo_Application.SetStatusBarMessage "��й�ȣ�� Ʋ���ϴ�.. Ȯ�ιٶ��ϴ�.", bmt_Short, True
                        Else
                            If PassWd1 <> PassWd2 Then
                                Sbo_Application.SetStatusBarMessage "�����й�ȣ�� �����й�ȣȮ�� ��ȣ�� Ʋ���ϴ�.. Ȯ�ιٶ��ϴ�.", bmt_Short, True
                            Else
                                sQry = "Update Z_PS_HRPASS set PassWd = '" & PassWd1 & "' Where BPLId = '" & BPLId & "' and MSTCOD = '" & MSTCOD & "'"
                                oRecordSet01.DoQuery sQry
                                Sbo_Application.SetStatusBarMessage "���� ó���� �Ͽ����ϴ�.. Ȯ�ιٶ��ϴ�.", bmt_Short, True
                            End If
                        End If
                    Else
                        If PassWd1 <> PassWd2 Then
                            Sbo_Application.SetStatusBarMessage "�����й�ȣ�� �����й�ȣȮ�� ��ȣ�� Ʋ���ϴ�.. Ȯ�ιٶ��ϴ�.", bmt_Short, True
                        Else
                            '//����� ������� insert
                            vReturnValue = Sbo_Application.MessageBox("��й�ȣ ��ϵ� ����� �ƴմϴ�. ����Ͻðڽ��ϱ�? (�űԺ���� 1111 �Դϴ�.)", 1, "&Ȯ��", "&���")
                            If vReturnValue = 1 Then
                                sQry = " insert into Z_PS_HRPASS Values ('" & BPLId & "','" & MSTCOD & "','" & FULLNAME & "','1111')"
                                oRecordSet01.DoQuery sQry
                                Sbo_Application.SetStatusBarMessage "�ű� ó���� �Ͽ����ϴ�.. Ȯ�ιٶ��ϴ�.", bmt_Short, True
                            Else
                            End If
                        End If
                    End If
'                    If oForm01.Mode = fm_ADD_MODE Then
''                        oForm01.Mode = fm_OK_MODE
'                        FormItemEnabled
''                        Call Sbo_Application.ActivateMenuItem("1282")
'                    ElseIf oForm01.Mode = fm_OK_MODE Then
'                        FormItemEnabled
'                        Call AddMatrixRow(0, True)
'                    End If
                End If
            
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then        '�̷����ڰ� �ٲ�� �� �� �� ���⵵��
                   If pval.ItemUID = "MSTCOD" Then
                      FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                   End If
                End If
            
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim BPLId           As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim sSeq$

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
        
' Matrix �ʵ忡 ���� ���� â ����ֱ�
   Select Case oUID
        Case "MSTCOD"
            BPLId = oForm01.Items("BPLId").Specific.VALUE
            sQry = "Select a.U_FULLNAME From OHEM a  "
            sQry = sQry & " Where a.U_MSTCOD = '" & Trim(oForm01.Items("MSTCOD").Specific.VALUE) & "' And a.branch = '" & oForm01.Items("BPLId").Specific.VALUE & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("FULLNAME").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
'            oMat01.LoadFromDataSource
   End Select
   
   
    Set oRecordSet01 = Nothing
    Exit Sub
       
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�ڵ带 Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo MenuEvent_Error
    
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim i&
    
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
            Case "1287": '����
   
            Case "1282": '�߰�
   
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư             '�߰���ư Ŭ���� Matrix�� �� �߰�
   
        End Select
    End If
    Exit Sub
MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function CreateItems() As Boolean
On Error GoTo CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
       
'    Set oDS_PS_HR409H = oForm01.DataSources.DBDataSources("@PS_HR409H")
'    Set oDS_PS_HR409L = oForm01.DataSources.DBDataSources("@PS_HR409L")

       
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 1)
    oForm01.Items("BPLId").Specific.DataBind.SetBound True, "", "BPLId"
    
    Call oForm01.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    oForm01.Items("MSTCOD").Specific.DataBind.SetBound True, "", "MSTCOD"
    
    Call oForm01.DataSources.UserDataSources.Add("PassWd", dt_SHORT_TEXT, 10)
    oForm01.Items("PassWd").Specific.DataBind.SetBound True, "", "PassWd"
    
    Call oForm01.DataSources.UserDataSources.Add("ChPWd1", dt_SHORT_TEXT, 10)
    oForm01.Items("ChPWd1").Specific.DataBind.SetBound True, "", "ChPWd1"
    
    Call oForm01.DataSources.UserDataSources.Add("ChPWd2", dt_SHORT_TEXT, 10)
    oForm01.Items("ChPWd2").Specific.DataBind.SetBound True, "", "ChPWd2"
    
    
'    oMat01.SelectionMode = ms_NotSupported
       
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error

    
    
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
             
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    
    '//���̵� ����� ����
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Call oForm01.Freeze(False)
    Exit Sub
ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
        'Call FormClear '//UDO���
        Call oForm01.EnableMenu("1281", False) '//ã��
        Call oForm01.EnableMenu("1282", False) '//�߰�
        oForm01.Items("Code").Enabled = False
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
        Call oForm01.EnableMenu("1281", False) '//ã��
        oForm01.Items("Code").Enabled = False
        Call oForm01.EnableMenu("1282", False) '//�߰�
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("Code").Enabled = False
        Call oForm01.EnableMenu("1282", False) '//�߰�
      
        
        '//����忡���� �����ۼ���
    End If
    Call oForm01.Freeze(False)
    Exit Sub
FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//�۾�
    ElseIf (eventInfo.BeforeAction = False) Then
        '//�۾�
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormDataEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_FormDataEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

