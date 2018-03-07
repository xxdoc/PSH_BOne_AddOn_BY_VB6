VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_QM210.cls
'//  Module : ǰ������ > �系ǥ�ذ���
'//  Desc : �系ǥ�ص����Ȳ
'//  FormType : PS_QM210
'//  Create Date(Start) : 2012.05.08
'//  Create Date(End) : 2012.05.08
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM210H As SAPbouiCOM.DBDataSource '������
Private oDS_PS_QM210L As SAPbouiCOM.DBDataSource '��϶���
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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM210.srf")
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
    
    oFormUniqueID01 = "PS_QM210_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO����϶�
        
    oForm01.Freeze True
    Call PS_QM210_CreateItems
    Call PS_QM210_ComboBox_Setting
    Call PS_QM210_Initial_Setting
    Call PS_QM210_CF_ChooseFromList
    Call PS_QM210_EnableMenus
    Call PS_QM210_SetDocument(oFromDocEntry01)
    Call PS_QM210_FormResize
    
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
                Call PS_QM210_AddMatrixRow(oMat01.VisualRowCount)
                
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
                Call oDS_PS_QM210L.RemoveRecord(oDS_PS_QM210L.Size - 1)
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
    
        If pval.ItemUID = "btnSearch" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_QM210_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    Call PS_QM210_MTX01 '��Ʈ������ ������ �ε�
                End If
            End If
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_QM210_Print_Report01
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
        ElseIf pval.ItemUID = "btnPrint" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
        
                Call PS_QM210_Print_Report01
            
            End If
            
        ElseIf pval.ItemUID = "btnPrint2" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
        
                Call PS_QM210_Print_Report02
            
            End If
            
        End If
        
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "PS_QM210" Then
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
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "DocCode", "") '�����ڵ� ���˼�ġ Ȱ��
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "DeptCode", "") '�μ� ���˼�ġ Ȱ��
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error
    Dim i&
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    'Dim SumWeight As Currency, SumQty As Long
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$, BPLID$
    'Dim Calculate_Weight As Double
    Dim Seq&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Dim sCount As Integer, sSeq As Integer
    
    Dim AttBPLId As String
    Dim AttTeam As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemUID = "AttBPLId" Then
'            oForm01.Freeze True
'            oMat01.Clear
'            oDS_PS_QM210L.Clear
'            oForm01.Freeze False
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "AttBPLId" Then
            oForm01.Freeze True
            
            '//�ش������� ���� Reset
            sCount = oForm01.Items("AttTeam").Specific.ValidValues.Count
            sSeq = sCount
            For i = 1 To sCount
                oForm01.Items("AttTeam").Specific.ValidValues.Remove sSeq - 1, psk_Index
                sSeq = sSeq - 1
            Next i
            
            AttBPLId = oForm01.Items("AttBPLId").Specific.VALUE
            
            '//�ش������� ������� Set
            sQry = "SELECT b.U_Code, b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code And a.Name = '�μ�' And b.U_Char2 = '" + AttBPLId + "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("AttTeam").Specific.ValidValues.Add "-", "����"
            
            Do Until oRecordSet01.EOF
                
                oForm01.Items("AttTeam").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                oRecordSet01.MoveNext
                
            Loop
            
            oForm01.Items("AttTeam").Specific.Select 0, psk_Index
            
            oForm01.Freeze False
        ElseIf pval.ItemUID = "AttTeam" Then
            oForm01.Freeze True
            
            '//�ش����� ����� Reset
            sCount = oForm01.Items("AttRsp").Specific.ValidValues.Count
            sSeq = sCount
            For i = 1 To sCount
                oForm01.Items("AttRsp").Specific.ValidValues.Remove sSeq - 1, psk_Index
                sSeq = sSeq - 1
            Next i
            
            AttTeam = oForm01.Items("AttTeam").Specific.VALUE
            
            '//�ش������� ������� Set
            sQry = "SELECT b.U_Code, b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code And a.Name = '���' And b.U_Char1 = '" + AttTeam + "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("AttRsp").Specific.ValidValues.Add "-", "����"
            
            Do Until oRecordSet01.EOF
                
                oForm01.Items("AttRsp").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                oRecordSet01.MoveNext
                
            Loop
            
            oForm01.Items("AttRsp").Specific.Select 0, psk_Index
            
            oForm01.Freeze False
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

    Dim oTempClass As Variant

    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "DocEntry" Then
                Set oTempClass = New PS_MM070
                Call oTempClass.LoadForm(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
            End If
        End If
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
        
            If (pval.ItemUID = "DeptCode") Then
            
                oQuery01 = "                   SELECT      T1.U_CodeNm AS [Name]"
                oQuery01 = oQuery01 & "  FROM        [@PS_HR200H] AS T0"
                oQuery01 = oQuery01 & "                  INNER JOIN"
                oQuery01 = oQuery01 & "                  [@PS_HR200L] AS T1"
                oQuery01 = oQuery01 & "                      ON T0.Code = T1.Code"
                oQuery01 = oQuery01 & "  WHERE       T1.U_UseYN = 'Y'"
                oQuery01 = oQuery01 & "                  AND T1.U_Char2 = 2"
                oQuery01 = oQuery01 & "                  AND T0.Name = '�μ�'"
                oQuery01 = oQuery01 & "                  AND T1.U_Code = '" & oForm01.Items("DeptCode").Specific.VALUE & "'"
                    
                Call oRecordSet01.DoQuery(oQuery01)
                oForm01.Items("DeptName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
'            ElseIf (pval.ItemUID = "ItemCode") Then
'
'                oQuery01 = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
'                oRecordSet01.DoQuery oQuery01
'                oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'
'            ElseIf (pval.ItemUID = "CntcCode") Then
'
'                oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
'                oRecordSet01.DoQuery oQuery01
'                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
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
        Call PS_QM210_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_QM210_FormResize
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


Private Function PS_QM210_CreateItems() As Boolean
On Error GoTo PS_QM210_CreateItems_Error
    Call oForm01.Freeze(True)
'    Dim oQuery01 As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Set oDS_PS_QM210H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_QM210L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '��Ʈ���� �ʱ�ȭ
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '�����_S
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    '�����_E
    
    '����Ÿ��_S
    Call oForm01.DataSources.UserDataSources.Add("DocType", dt_SHORT_TEXT, 100)
    Call oForm01.Items("DocType").Specific.DataBind.SetBound(True, "", "DocType")
    '����Ÿ��_E
    
    '�����ڵ�_S
    Call oForm01.DataSources.UserDataSources.Add("DocCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("DocCode").Specific.DataBind.SetBound(True, "", "DocCode")
    '�����ڵ�_E
    
    '�μ�_S
    Call oForm01.DataSources.UserDataSources.Add("DeptCode", dt_SHORT_TEXT, 200)
    Call oForm01.Items("DeptCode").Specific.DataBind.SetBound(True, "", "DeptCode")
    '�μ�_E
    
    '�μ���_S
    Call oForm01.DataSources.UserDataSources.Add("DeptName", dt_SHORT_TEXT, 200)
    Call oForm01.Items("DeptName").Specific.DataBind.SetBound(True, "", "DeptName")
    '�μ���_E
    
    '��⹮�� ����_S
    Call oForm01.DataSources.UserDataSources.Add("chkDelYN", dt_SHORT_TEXT, 1)
    Call oForm01.Items("chkDelYN").Specific.DataBind.SetBound(True, "", "chkDelYN")
    '��⹮�� ����_E
    
    Call oForm01.Freeze(False)
    Exit Function
PS_QM210_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM210_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_QM210_ComboBox_Setting()
On Error GoTo PS_QM210_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
     
    '�����_S
    Call oForm01.Items("BPLId").Specific.ValidValues.Add("", "����")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
    Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
    '�����_E
     
    '����Ÿ��_S
    sQry = "            SELECT      T1.U_Minor,"
    sQry = sQry & "                 T1.U_CdName"
    sQry = sQry & "  FROM       [@PS_SY001H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_SY001L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      T0.Code = 'Q200'"
    
    Call oForm01.Items("DocType").Specific.ValidValues.Add("", "����")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("DocType").Specific, sQry, "", False, False)
    Call oForm01.Items("DocType").Specific.Select(0, psk_Index)
    '����Ÿ��_E
        
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
    
PS_QM210_ComboBox_Setting_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM210_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM210_Initial_Setting()
On Error GoTo PS_QM210_Initial_Setting_Error

'    '����� ������� �Ҽ� ����� ����
'    Call oForm01.Items("AttBPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
'
'    'üũ�ڽ� ����
'    'oForm01.Items("Chk01").Specific.Checked = True
'
'    '��¥ ����
'    oForm01.Items("StdYM").Specific.VALUE = Left(Format(Date, "YYYYMMDD"), 4) + "-" + Mid(Format(Date, "YYYYMMDD"), 5, 2)
''    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM") + "01"
 
    Exit Sub
PS_QM210_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM210_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_QM210_CF_ChooseFromList()
On Error GoTo PS_QM210_CF_ChooseFromList_Error
    Exit Sub
PS_QM210_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM210_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM210_FormItemEnabled()
On Error GoTo PS_QM210_FormItemEnabled_Error

    Call oForm01.Freeze(True)
    
    If (oForm01.Mode = fm_ADD_MODE) Then

        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue) '�α����� ������� ����η� �⺻ ����

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    
    Call oForm01.Freeze(False)
    
    Exit Sub
    
PS_QM210_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM210_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM210_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM210_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//���߰�����
        oDS_PS_QM210L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_QM210L.Offset = oRow
'    oDS_PS_QM210L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM210_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM210_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM210_FormClear()
On Error GoTo PS_QM210_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM210'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_QM210_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM210_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM210_EnableMenus()
On Error GoTo PS_QM210_EnableMenus_Error
    Exit Sub
PS_QM210_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM210_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM210_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM210_SetDocument_Error

    If (oFromDocEntry01 = "") Then
    
        Call PS_QM210_FormItemEnabled

    Else

    End If
    
    Exit Sub
    
PS_QM210_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM210_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_QM210_DataValidCheck() As Boolean
On Error GoTo PS_QM210_DataValidCheck_Error
    PS_QM210_DataValidCheck = False
    Dim i As Long

'    If oForm01.Items("DocDateTo").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "�������� To�� �ʼ��Դϴ�.", bmt_Short, True
'        oForm01.Items("DocDateTo").Click ct_Regular
'        PS_QM210_DataValidCheck = False
'        Exit Function
'    End If

    PS_QM210_DataValidCheck = True
    Exit Function
PS_QM210_DataValidCheck_Error:
    PS_QM210_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_QM210_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_QM210_MTX01()
On Error GoTo PS_QM210_MTX01_Error
    '//��Ʈ������ ������ �ε�
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As Integer '�����
    Dim DocType As String '����Ÿ��
    Dim DocCode As String '�����ڵ�
    Dim DeptCode As String '�μ�
    Dim chkDelYN As String '��⹮������
    
    BPLID = oForm01.Items("BPLId").Specific.Selected.VALUE '�����
    DocType = oForm01.Items("DocType").Specific.Selected.VALUE '����Ÿ��
    DocCode = oForm01.Items("DocCode").Specific.VALUE '�����ڵ�
    DeptCode = oForm01.Items("DeptCode").Specific.VALUE '�μ�
    chkDelYN = IIf(oForm01.Items("chkDelYN").Specific.Checked = True, "Y", "N") '��⹮������
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", RecordSet01.RecordCount, False) '������ ������ �� ���� ���α׷��� ����
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_QM210_01 '"
    Query01 = Query01 & BPLID & "','"
    Query01 = Query01 & DocType & "','"
    Query01 = Query01 & DocCode & "','"
    Query01 = Query01 & DeptCode & "','"
    Query01 = Query01 & chkDelYN & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_QM210_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_QM210L.InsertRecord(loopCount)
        End If
        oDS_PS_QM210L.Offset = loopCount

        Call oDS_PS_QM210L.setValue("U_LineNum", loopCount, loopCount + 1) '���ι�ȣ
        Call oDS_PS_QM210L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("DocType").VALUE) '����Ÿ��(��)
        Call oDS_PS_QM210L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("DocCode").VALUE) '�����ڵ�
        Call oDS_PS_QM210L.setValue("U_ColReg10", loopCount, RecordSet01.Fields("StdName").VALUE) 'ǥ�ع�����
        Call oDS_PS_QM210L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("DeptCode").VALUE) '�μ��ڵ�
        Call oDS_PS_QM210L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("DeptName").VALUE) '�μ���
        Call oDS_PS_QM210L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("CrtDate").VALUE) '��������
        Call oDS_PS_QM210L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("EmpCode").VALUE) '����ڻ��
        Call oDS_PS_QM210L.setValue("U_ColReg07", loopCount, RecordSet01.Fields("EmpName").VALUE) '����ڼ���
        Call oDS_PS_QM210L.setValue("U_ColReg08", loopCount, RecordSet01.Fields("CrtCmt").VALUE) '��������
        
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
    
PS_QM210_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "W")
    Exit Sub
PS_QM210_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM210_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_QM210_FormResize()
On Error GoTo PS_QM210_FormResize_Error
    
    Exit Sub
PS_QM210_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM210_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM210_Print_Report01()
On Error GoTo Print_Query_Error

    Dim DocNum As String
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String
    
    Dim BPLID As Integer '�����
    Dim DocType As String '����Ÿ��
    Dim DocCode As String '�����ڵ�
    Dim DeptCode As String '�μ�
    Dim chkDelYN As String '��⹮������
    
    BPLID = oForm01.Items("BPLId").Specific.Selected.VALUE '�����
    DocType = oForm01.Items("DocType").Specific.Selected.VALUE '����Ÿ��
    DocCode = oForm01.Items("DocCode").Specific.VALUE '�����ڵ�
    DeptCode = oForm01.Items("DeptCode").Specific.VALUE '�μ�
    chkDelYN = IIf(oForm01.Items("chkDelYN").Specific.Checked = True, "Y", "N") '��⹮������

    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", 100, False)
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC

    WinTitle = "[PS_QM210] ����Ʈ"
    ReportName = "PS_QM210_01.rpt"
 
    '// Formula �����ʵ�
    ReDim gRpt_Formula(0)
    ReDim gRpt_Formula_Value(0)
    
    '// SubReport
    ReDim gRpt_SRptSqry(0)
    ReDim gRpt_SRptName(0)
    
    ReDim gRpt_SFormula(0, 0)
    ReDim gRpt_SFormula_Value(0, 0)
    
    '/ Procedure ����"
    sQry = "         EXEC PS_QM210_01 '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & DocType & "','"
    sQry = sQry & DocCode & "','"
    sQry = sQry & DeptCode & "','"
    sQry = sQry & chkDelYN & "'"
    
    Call oRecordSet.DoQuery(sQry)
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query�� �������� 'Y'��...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub PS_QM210_Print_Report02()
On Error GoTo PS_QM210_Print_Report02_Error

    Dim DocNum As String
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String
    
    Dim BPLID As Integer '�����
    Dim DocType As String '����Ÿ��
    Dim DocCode As String '�����ڵ�
    Dim DeptCode As String '�μ�
    Dim chkDelYN As String '��⹮������
    
    BPLID = oForm01.Items("BPLId").Specific.Selected.VALUE '�����
    DocType = oForm01.Items("DocType").Specific.Selected.VALUE '����Ÿ��
    DocCode = oForm01.Items("DocCode").Specific.VALUE '�����ڵ�
    DeptCode = oForm01.Items("DeptCode").Specific.VALUE '�μ�
    chkDelYN = IIf(oForm01.Items("chkDelYN").Specific.Checked = True, "1", "0") '��⹮������
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", 100, False)

    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC

    WinTitle = "[PS_QM210_02] ǥ�ص����Ȳ"
    ReportName = "PS_QM210_02.rpt"
 
    '// Formula �����ʵ�
    ReDim gRpt_Formula(0)
    ReDim gRpt_Formula_Value(0)
    
    '// SubReport
    ReDim gRpt_SRptSqry(0)
    ReDim gRpt_SRptName(0)
    
    ReDim gRpt_SFormula(0, 0)
    ReDim gRpt_SFormula_Value(0, 0)
    
    '/ Procedure ����"
    sQry = "EXEC PS_QM210_02 '" & BPLID & "','" & DocType & "','" & DocCode & "','" & DeptCode & "','" & chkDelYN & "'"

    Call oRecordSet.DoQuery(sQry)
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PS_QM210_Print_Report02_Error
    End If

    '/ Action (sub_query�� �������� 'Y'��...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM210_Print_Report02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_QM210_Print_Report02_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

