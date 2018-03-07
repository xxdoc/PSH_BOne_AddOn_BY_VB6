VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY681"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PH_PY681.cls
'//  Module : �λ���� > ���°��� > ���¸���Ʈ
'//  Desc : ��ٹ��ϼ���Ȳ
'//  FormType : PH_PY681
'//  Create Date(Start) : 2014.05.08
'//  Create Date(End) : 2014.05.12
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid

Private oLastItemUID01 As String 'Ŭ�������� ������ ������ ������ Uid��
Private oLastColUID01 As String '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLastColRow01 As Long '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml As String
    Dim oXmlDoc As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY681.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '��Ʈ������ Ÿ��Ʋ���̿� �����̸� ����
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY681_" & GetTotalFormsCount
    Call AddForms(Me, oFormUniqueID01, "PH_PY681") '//���߰�
    Call Sbo_Application.LoadBatchActions(oXmlDoc.xml)
    '�� �Ҵ�
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    '//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO����϶�
        
    Call oForm.Freeze(True)
    Call PH_PY681_CreateItems
    Call PH_PY681_ComboBox_Setting
    Call PH_PY681_CF_ChooseFromList
    Call PH_PY681_EnableMenus
    Call PH_PY681_SetDocument(oFromDocEntry01)
    Call PH_PY681_FormResize
    
    Call oForm.EnableMenu("1283", False) '����
    Call oForm.EnableMenu("1286", False) '�ݱ�
    Call oForm.EnableMenu("1287", False) '����
    Call oForm.EnableMenu("1285", False) '����
    Call oForm.EnableMenu("1284", False) '���
    Call oForm.EnableMenu("1293", False) '�����
    Call oForm.EnableMenu("1281", False)
    Call oForm.EnableMenu("1282", True)
  
    oForm.Items("YYYY").Specific.VALUE = Format(Now, "YYYY") - 1
  
    Call oForm.Update
    Call oForm.Freeze(False)
    
    oForm.Visible = True
    Set oXmlDoc = Nothing

    Exit Sub
LoadForm_Error:
    Call oForm.Update
    Call oForm.Freeze(False)
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Call Sbo_Application.SetStatusBarMessage("Form_Load Error:" & Err.Description, bmt_Short, True)
End Sub

Public Sub PH_PY681_MTX01()
'******************************************************************************
'Function ID : PH_PY681_MTX01()
'�ش��� : PH_PY681
'��� : ������ ��ȸ
'�μ� : ����
'��ȯ�� : ����
'Ư�̻��� : ����
'******************************************************************************
On Error GoTo PH_PY681_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String '�����
    Dim TeamCode As String '��
    Dim RspCode As String '���
    Dim ClsCode As String '��
    Dim MSTCOD As String '�����ȣ
    Dim yyyy As String '���س⵵
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '�����
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '��
    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE) '���
    ClsCode = Trim(oForm.Items("ClsCode").Specific.VALUE) '��
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '�����ȣ
    yyyy = Trim(oForm.Items("YYYY").Specific.VALUE) '���س⵵
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", 100, False)
    
    Call oForm.Freeze(True)
    
    sQry = "                EXEC [PH_PY681_01] "
    sQry = sQry & "'" & CLTCOD & "'," '�����
    sQry = sQry & "'" & TeamCode & "',"  '��
    sQry = sQry & "'" & RspCode & "',"  '���
    sQry = sQry & "'" & ClsCode & "',"  '��
    sQry = sQry & "'" & MSTCOD & "',"  '���
    sQry = sQry & "'" & yyyy & "'"  '���س⵵
    
    oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable")
    Call oGrid01.DataTable.Clear
    Call oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry)
    
    ProgBar01.VALUE = 100
    ProgBar01.Text = "��ȸ��...!"
    ProgBar01.Stop
    
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

    If oGrid01.Rows.Count = 0 Then
        ErrNum = 1
        GoTo PH_PY681_MTX01_Error
    End If
        
    Call oGrid01.AutoResizeColumns
    Call oForm.Update
    
    Call oForm.Freeze(False)
  
    Set ProgBar01 = Nothing
'    Set RecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY681_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
'    Set RecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("��ȸ ����� �����ϴ�. Ȯ���ϼ���.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY681_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Function PH_PY681_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PH_PY681_HeaderSpaceLineDel()
'�ش��� : PH_PY681
'��� : �ʼ��Է»��� üũ
'�μ� : ����
'��ȯ�� : True:�ʼ��Է»����� ��� �Է�, Fasle:�ʼ��Է»��� �� �ϳ��� �Է����� �ʾ���
'Ư�̻��� : ����
'******************************************************************************
On Error GoTo PH_PY681_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm.Items("YYYY").Specific.VALUE) = "" '���س⵵
            ErrNum = 1
            GoTo PH_PY681_HeaderSpaceLineDel_Error
'        Case Trim(oForm.Items("DestNo2").Specific.VALUE) = "" '�����ȣ2
'            ErrNum = 2
'            GoTo PH_PY681_HeaderSpaceLineDel_Error
'        Case Trim(oForm.Items("MSTCOD").Specific.VALUE) = "" '�����ȣ
'            ErrNum = 3
'            GoTo PH_PY681_HeaderSpaceLineDel_Error
'        Case Trim(oForm.Items("FrDate").Specific.VALUE) = "" '��������
'            ErrNum = 4
'            GoTo PH_PY681_HeaderSpaceLineDel_Error
'        Case Trim(oForm.Items("FrTime").Specific.VALUE) = "" '���۽ð�
'            ErrNum = 5
'            GoTo PH_PY681_HeaderSpaceLineDel_Error
'        Case Trim(oForm.Items("ToDate").Specific.VALUE) = "" '��������
'            ErrNum = 6
'            GoTo PH_PY681_HeaderSpaceLineDel_Error
'        Case Trim(oForm.Items("ToTime").Specific.VALUE) = "" '����ð�
'            ErrNum = 7
'            GoTo PH_PY681_HeaderSpaceLineDel_Error
    End Select

    PH_PY681_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY681_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���س⵵�� �ʼ���ȸ �����Դϴ�. Ȯ���ϼ���.", "E"
        Call oForm.Items("YYYY").CLICK(ct_Regular)
'    ElseIf ErrNum = 2 Then
'        MDC_Com.MDC_GF_Message "�����ȣ2�� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
'        Call oForm.Items("DestNo2").CLICK(ct_Regular)
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "�����ȣ�� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
'        Call oForm.Items("MSTCOD").CLICK(ct_Regular)
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "�������ڴ� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
'        Call oForm.Items("FrDate").CLICK(ct_Regular)
'    ElseIf ErrNum = 5 Then
'        MDC_Com.MDC_GF_Message "���۽ð��� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
'        Call oForm.Items("FrTime").CLICK(ct_Regular)
'    ElseIf ErrNum = 6 Then
'        MDC_Com.MDC_GF_Message "�������ڴ� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
'        Call oForm.Items("FrDate").CLICK(ct_Regular)
'    ElseIf ErrNum = 7 Then
'        MDC_Com.MDC_GF_Message "����ð��� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
'        Call oForm.Items("FrTime").CLICK(ct_Regular)
    End If
    PH_PY681_HeaderSpaceLineDel = False
End Function

'/ ��Ʈ���� �ʼ� ���� check
Private Function PH_PY681_MatrixSpaceLineDel() As Boolean
On Error GoTo PH_PY681_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PH_PY681_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY681_MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "�� ������ ����ڵ尡 �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "�� ������ �ð��� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "�� ������ ������ڰ� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "�� ������ �񰡵��ڵ尡 �����ϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "PH_PY681_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PH_PY681_MatrixSpaceLineDel = False
End Function

Private Sub PH_PY681_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY681_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String
    Dim TeamCode As String
    Dim RspCode As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    
    Call oForm.Freeze(True)
    
    Select Case oUID
            
        Case "MSTCOD"
        
            oForm.Items("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.VALUE & "'") '����
            
        Case "CLTCOD"
            
            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
            
            If oForm.Items("TeamCode").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("TeamCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '�μ��޺�����
            Call oForm.Items("TeamCode").Specific.ValidValues.Add("%", "��ü")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '1'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char2 = '" & CLTCOD & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("TeamCode").Specific, sQry, "", False, False)
            Call oForm.Items("TeamCode").Specific.Select(0, psk_Index)
            
        Case "TeamCode"
            
            TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
            
            If oForm.Items("RspCode").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("RspCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("RspCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '����޺�����
            Call oForm.Items("RspCode").Specific.ValidValues.Add("%", "��ü")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '2'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char1 = '" & TeamCode & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("RspCode").Specific, sQry, "", False, False)
            Call oForm.Items("RspCode").Specific.Select(0, psk_Index)
        
        Case "RspCode"
            
            TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
            RspCode = Trim(oForm.Items("RspCode").Specific.VALUE)
            
            If oForm.Items("ClsCode").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("ClsCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("ClsCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '���޺�����
            Call oForm.Items("ClsCode").Specific.ValidValues.Add("%", "��ü")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '9'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char1 = '" & RspCode & "'"
            sQry = sQry & "                 AND U_Char2 = '" & TeamCode & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("ClsCode").Specific, sQry, "", False, False)
            Call oForm.Items("ClsCode").Specific.Select(0, psk_Index)
        
    End Select

    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PH_PY681_FlushToItemValue_Error:

    Call oForm.Freeze(False)
'    Call ProgBar01.Stop
'    ProgBar01.VALUE = 100
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing

    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("PH_PY681_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If

End Sub

'/���� ������ �������
Sub PH_PY681_FormItemEnabled()
On Error GoTo PH_PY681_FormItemEnabled_Error
    If (oForm.Mode = fm_ADD_MODE) Then
    
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_FIND_MODE) Then
    
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_OK_MODE) Then
    
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
    End If
    
    Exit Sub
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY681_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY681_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/������ ���� �̺�Ʈ
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error

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
Raise_FormItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormMenuEvent_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
                '/�߰���ư Ŭ���� ��Ʈ���� insertrow
                
'                Call PH_PY681_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
'                oForm.Mode = fm_ADD_MODE
'                BubbleEvent = False
'                Call PH_PY681_LoadCaption
                
                'oForm.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
            
            Case "7169": '���� ��������
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
            Case "1281": 'ã��
                '//Call PH_PY681_FormItemEnabled '//UDO���
            Case "1282": '�߰�
'                oMat01.Clear
'                oDS_PH_PY681A.Clear

'                Call PH_PY681_LoadCaption
'                Call PH_PY681_FormItemEnabled
                '//Call PH_PY681_FormItemEnabled '//UDO���
                '//Call PH_PY681_AddMatrixRow(0, True) '//UDO���
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                '//Call PH_PY681_FormItemEnabled
                
            Case "7169": '���� ��������
            
        End Select
    End If
    Exit Sub
Raise_FormMenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    
        If pval.ItemUID = "PH_PY681" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
            End If
        End If
            
        If pval.ItemUID = "BtnSearch" Then '/��ȸ
            
            If PH_PY681_HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            End If
        
            Call PH_PY681_MTX01
        
        ElseIf pval.ItemUID = "BtnPrint" Then

            If PH_PY681_HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            End If

            Call PH_PY681_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PH_PY681" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD", "") '���
        
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
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        
        Call PH_PY681_FlushToItemValue(pval.ItemUID)
        
    End If

    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm.Freeze(False)
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

    Call oForm.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
            
            Call PH_PY681_FlushToItemValue(pval.ItemUID)
            
        End If
        
    ElseIf pval.BeforeAction = False Then

    End If
    
    Call oForm.Freeze(False)
    
    Exit Sub
    
Raise_EVENT_VALIDATE_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PH_PY681_FormItemEnabled
        '//Call PH_PY681_AddMatrixRow(oMat01.VisualRowCount) '//UDO���
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY681_FormResize
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
'            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY681A", "U_CardCode,U_CardName")
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
        Set oForm = Nothing
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

        ElseIf pval.BeforeAction = False Then
            
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY681_CreateItems() As Boolean
On Error GoTo PH_PY681_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm.Items("Grid01").Specific
    Call oForm.DataSources.DataTables.Add("DataTable")
    oGrid01.SelectionMode = ms_Auto

    '�����
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    
    '��
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Call oForm.Items("TeamCode").Specific.DataBind.SetBound(True, "", "TeamCode")
    
    '���
    Call oForm.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 10)
    Call oForm.Items("RspCode").Specific.DataBind.SetBound(True, "", "RspCode")
    
    '��
    Call oForm.DataSources.UserDataSources.Add("ClsCode", dt_SHORT_TEXT, 10)
    Call oForm.Items("ClsCode").Specific.DataBind.SetBound(True, "", "ClsCode")
    
    '���
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 20)
    Call oForm.Items("MSTCOD").Specific.DataBind.SetBound(True, "", "MSTCOD")
    
    '����
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 50)
    Call oForm.Items("MSTNAM").Specific.DataBind.SetBound(True, "", "MSTNAM")
    
    '���س⵵
    Call oForm.DataSources.UserDataSources.Add("YYYY", dt_SHORT_TEXT, 4)
    Call oForm.Items("YYYY").Specific.DataBind.SetBound(True, "", "YYYY")
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY681_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY681_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/�޺��ڽ� set
Sub PH_PY681_ComboBox_Setting()
On Error GoTo PH_PY681_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)

    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY681_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY681_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY681_CF_ChooseFromList()
On Error GoTo PH_PY681_CF_ChooseFromList_Error

    Exit Sub
PH_PY681_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY681_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY681_EnableMenus()
On Error GoTo PH_PY681_EnableMenus_Error
    
    Exit Sub
PH_PY681_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY681_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY681_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY681_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PH_PY681_FormItemEnabled
        '//Call PH_PY681_AddMatrixRow(0, True) '//UDO����϶�
    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY681_FormItemEnabled
'        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm.Items("1").Click ct_Regular
    End If
    Exit Sub
PH_PY681_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY681_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY681_FormResize()
On Error GoTo PH_PY681_FormResize_Error
    
    Exit Sub
PH_PY681_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY681_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY681_Print_Report01()
    On Error GoTo PH_PY681_Print_Report01_Error

    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim CLTCOD As String '�����
    Dim TeamCode As String '��
    Dim RspCode As String '���
    Dim ClsCode As String '��
    Dim MSTCOD As String '�����ȣ
    Dim yyyy As String '���س⵵
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '�����
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '��
    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE) '���
    ClsCode = Trim(oForm.Items("ClsCode").Specific.VALUE) '��
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '�����ȣ
    yyyy = Trim(oForm.Items("YYYY").Specific.VALUE) '���س⵵

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", 100, False)

     '/ ODBC ���� üũ
    If ConnectODBC = False Then
        GoTo PH_PY681_Print_Report01_Error
    End If

    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY681] ��ٹ��ϼ� ��Ȳ"

    ReportName = "PH_PY681_01.rpt"

    '// Formula �����ʵ�
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)

    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)

    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)

    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""

    '/ Procedure ����"
    sQry = "                EXEC [PH_PY681_02] "
    sQry = sQry & "'" & CLTCOD & "'," '�����
    sQry = sQry & "'" & TeamCode & "',"  '��
    sQry = sQry & "'" & RspCode & "',"  '���
    sQry = sQry & "'" & ClsCode & "',"  '��
    sQry = sQry & "'" & MSTCOD & "',"  '���
    sQry = sQry & "'" & yyyy & "'"  '���س⵵

    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V", "", 2) = False Then
        GoTo PH_PY681_Print_Report01_Error
    End If

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Exit Sub

PH_PY681_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("Print_Query_Error:" & Err.Number & " - " & Err.Description, "E")
    
End Sub








