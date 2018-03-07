VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM055"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_QM055.cls
'//  Module         : QM
'//  Desc           : ��ǰ ��ǰ�˻��缭 ���
'//  FormType       : PS_QM055
'//  Create Date    : 2013.12.26
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit         '��� ���ؿ��� �ش��ϴ� ����� ��� ������ ���������� �����ϵ��� �� �� ����մϴ�.

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM055H       As SAPbouiCOM.DBDataSource    '������ 'Represents a SAP Business One database table to be attached to a form.
Private oDS_PS_QM055L       As SAPbouiCOM.DBDataSource    '��϶���


Private oLast_Item_UID   As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID    As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row    As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��
Private oSeq&
Private TmpCode As String

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oBPLId01 As String, Optional ByVal oItemCode01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM055.srf")
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
    
    oFormUniqueID01 = "PS_QM055_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    ' ������ �Ӽ� � ��ü ������ �����մϴ�

    oForm01.SupportedModes = -1
    
    If Trim(oItemCode01) = "" Then
        oForm01.Mode = fm_ADD_MODE
    Else
        oForm01.Mode = fm_FIND_MODE
    End If
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm01.DataBrowser.BrowseBy = "DocEntry"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////D////////////////////////////////////////////////
    
    
    oForm01.Freeze True
    Call CreateItems
    
    oForm01.EnableMenu ("1283"), True      '// ����
    oForm01.EnableMenu ("1293"), True      '// �����
    oForm01.EnableMenu ("1287"), True      '// ����
    oForm01.EnableMenu ("1284"), False     '// ���
    
    
    
    Call ComboBox_Setting
    Call CF_ChooseFromList
    Call Initial_Setting
    Call FormItemEnabled
    Call FormClear '//UDO����϶�
    Call AddMatrixRow(0, oMat01.RowCount, True) '//UDO����϶�
    oForm01.Update      ' Reloads the form from the application.
                        ' Use this method to update an object from the database.

'    If oBaseItemCode01 <> "" Then
'        oForm01.Items("1").Click ct_Regular
'    End If
    
    oForm01.Freeze False
       
    oForm01.Visible = True  ' Indicates whether the form is visible.
    
    
    
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
    Dim sQry As String
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    
    Dim ItemCode As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                            Call FormClear
                            '// Input Code, Name
                            
                        End If
'                        If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
'
'                            Dim OutSize As String
'
'                            OutSize = oForm01.Items("OutSize").Specific.VALUE
'                            ItemCode = oForm01.Items("ItemCode").Specific.VALUE
'                            sQry = "Update [OITM] set U_OutSize = '" & OutSize & "' Where ItemCode = '" & ItemCode & "'"
'                            oRecordSet01.DoQuery sQry
'                        End If
                 
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                ' ��ǰ�ڵ�
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "ItemCode" Then
                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "InspEmp" Then
                        If oForm01.Items("InspEmp").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ConfEmp" Then
                        If oForm01.Items("ConfEmp").Specific.VALUE = "" Then
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
                 oLast_Item_UID = pval.ItemUID

            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
            '//��Ʈ������ ������ �ε�
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
              If pval.ItemUID = "1" Then
            
                If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                    oForm01.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                    
                ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = False Then
                    FormItemEnabled
                    AddMatrixRow 1, oMat01.RowCount, True
                End If
                
                
                
              End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                
                Call oForm01.Freeze(True)
                If pval.ItemChanged = True Then
                    ' ��ǰ�ڵ�
                    If pval.ItemUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "InspEmp" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ConfEmp" Then
                        FlushToItemValue pval.ItemUID
                    
                    ElseIf (pval.ItemUID = "Mat01") Then
                        Call oDS_PS_QM055L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM055L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                           AddMatrixRow 1, oMat01.VisualRowCount, True
                           Call oDS_PS_QM055L.setValue("U_Seqno", pval.Row - 1, oMat01.Columns("LineNum").Cells(pval.Row).Specific.VALUE)
                        End If
                        oMat01.LoadFromDataSource
                        oMat01.AutoResizeColumns
                        oForm01.Update
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click
                    End If
                    
                    
                End If
                
'                If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM055L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                   AddMatrixRow 1, oMat01.VisualRowCount, True
'                End If
                Call oForm01.Freeze(False)
            Case et_MATRIX_LOAD: '//11
                AddMatrixRow 1, oMat01.VisualRowCount, True
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
                oLast_Item_UID = pval.ItemUID
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
Call oForm01.Freeze(False)
Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�ű�(�߰�)��忡���� ������ �� �����ϴ�. ��ǰ��ȸ �� ó���ϼ���.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
        
    
End Sub


Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    If (pval.BeforeAction = True) Then      '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
            Case "1293"     '�����
        End Select
        
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1281": 'ã��
                FormItemEnabled
            Case "1282": '�߰�
                FormItemEnabled
                FormClear
                AddMatrixRow 0, oMat01.RowCount, True
                oForm01.Items("ItemCode").Click ct_Collapsed
                
            Case "1287": '����
                oForm01.Freeze True
                Call FormClear
                oForm01.Items("ItemCode").Enabled = True
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_QM055L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                               
            Case "1293"     '�����
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource    ' DBDataSource�� ���ڵ尡 ���� �� �����.
                    Call oDS_PS_QM055L.RemoveRecord(oDS_PS_QM055L.Size - 1)     ' ���ڵ� �� ���� �����.
                    oMat01.LoadFromDataSource   ' DBDataSource�� ��Ʈ������ �ø���
                    If oMat01.RowCount = 0 Then
                        Call AddMatrixRow(1, 0, True)
                    Else
                        If Trim(oDS_PS_QM055L.GetValue("U_InspItem", oMat01.RowCount - 1)) <> "" Then
                            Call AddMatrixRow(1, oMat01.RowCount, True)

                        End If
                    End If
                End If
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
                FormItemEnabled
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

Function CreateItems() As Boolean
On Error GoTo CreateItems_Error
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_QM055H = oForm01.DataSources.DBDataSources("@PS_QM055H")
    Set oDS_PS_QM055L = oForm01.DataSources.DBDataSources("@PS_QM055L")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    
    
    Set oRecordSet01 = Nothing
    Exit Function
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//�޺��� �⺻������
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("AQLYN").Specific
    oCombo.ValidValues.Add "N", "N"
    oCombo.ValidValues.Add "Y", "Y"
    oCombo.Select 0, psk_Index
    
    Exit Sub
ComboBox_Setting_Error:
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList ����
    Exit Sub
CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub Initial_Setting()
On Error GoTo Initial_Setting_Error
    '//���̵� ����� ����
    Dim oCombo          As SAPbouiCOM.ComboBox
    Set oCombo = oForm01.Items("BPLId").Specific
    
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
Exit Sub
Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "Initial_Setting_Error: " & Err.Number & " _ " & Err.Description, bmt_Short, True
End Sub


Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("ItemName").Enabled = True
        oForm01.Items("Size").Enabled = True
        oForm01.Items("OutSize").Enabled = True
        oForm01.Items("Inspect").Enabled = True
        oForm01.Items("InspEmp").Enabled = True
        oForm01.Items("ConfEmp").Enabled = True
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
        '//����忡���� �����ۼ���
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("ItemName").Enabled = True
        oForm01.Items("Size").Enabled = True
        oForm01.Items("OutSize").Enabled = True
        oForm01.Items("Inspect").Enabled = True
        oForm01.Items("InspEmp").Enabled = True
        oForm01.Items("ConfEmp").Enabled = True
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//����忡���� �����ۼ���
        '//����忡���� �����ۼ���
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("ItemName").Enabled = True
        oForm01.Items("Size").Enabled = True
        oForm01.Items("OutSize").Enabled = True
        oForm01.Items("Inspect").Enabled = True
        oForm01.Items("InspEmp").Enabled = True
        oForm01.Items("ConfEmp").Enabled = True
    End If
    Exit Sub
FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oSeq As Integer, ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo AddMatrixRow_Error
   Select Case oSeq
       Case 0:
            oMat01.AddRow
            oDS_PS_QM055L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
       Case 1:
            oDS_PS_QM055L.InsertRecord oRow
            oDS_PS_QM055L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
   End Select
'AddMatrixRow_Error:
'    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM055'", "")
    If DocNum = 0 Then
'        oForm01.Items("DocEntry").Specific.String = 1
        oDS_PS_QM055H.setValue "DocEntry", 0, "1"
        oDS_PS_QM055H.setValue "Code", 0, "1"
        oDS_PS_QM055H.setValue "Name", 0, "1"
    Else
'        oForm01.Items("DocEntry").Specific.String = DocNum
        oDS_PS_QM055H.setValue "DocEntry", 0, DocNum    ' ȭ�鿡 ������ �ȵǱ� ����
        oDS_PS_QM055H.setValue "Code", 0, DocNum
        oDS_PS_QM055H.setValue "Name", 0, DocNum
    End If
          
    
    Exit Sub
FormClear_Error:
    Sbo_Application.SetStatusBarMessage "FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function DataValidCheck() As Boolean
On Error GoTo DataValidCheck_Error
    '//��ȿ���˻�
    Exit Function
DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub MTX01()
On Error GoTo MTX01_Error
    '//��Ʈ������ ������ �ε�
    Exit Sub
MTX01_Error:
Sbo_Application.SetStatusBarMessage "MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    oForm01.Freeze True
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_QM055H.GetValue("U_ItemCode", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_QM055H.GetValue("U_Inspect", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select
    oForm01.Freeze False
    HeaderSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "��ǰ�ڵ�� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "�˻����� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i             As Long
    Dim ErrNum        As Integer
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    Dim sQry          As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    ErrNum = 0
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '// ȭ����� ��Ʈ������ �Էµ� ������ ��� ������ͼҽ��� �ѱ��
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat01.FlushToDataSource
    
    '// ����
    If oMat01.VisualRowCount <= 1 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '// �Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '// �̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    If oMat01.VisualRowCount > 0 Then
        If oDS_PS_QM055L.GetValue("U_InspItem", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_QM055L.RemoveRecord oMat01.VisualRowCount - 1
        End If
    End If
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '���� �����Ͽ����� DB������ �ҽ��� �ٽ� �����´�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat01.LoadFromDataSource
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    Dim i$
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim MItemCod$
    Dim Qty As Currency
    Dim Calculate_Weight As Currency
    Dim vReturnValue$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "ItemCode"
            sQry = "Select FrgnName, U_Size, U_OutSize From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oDS_PS_QM055H.setValue "U_ItemName", 0, Trim(oRecordSet01.Fields(0).VALUE)
            oDS_PS_QM055H.setValue "U_Size", 0, Trim(oRecordSet01.Fields(1).VALUE)
            oDS_PS_QM055H.setValue "U_OutSize", 0, Trim(oRecordSet01.Fields(2).VALUE)
           
        Case "InspEmp"
            sQry = "Select U_FullName From [@PH_PY001A] Where Code = '" & Trim(oForm01.Items("InspEmp").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oDS_PS_QM055H.setValue "U_InspName", 0, Trim(oRecordSet01.Fields(0).VALUE)
        Case "ConfEmp"
            sQry = "Select U_FullName From [@PH_PY001A] Where Code = '" & Trim(oForm01.Items("ConfEmp").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oDS_PS_QM055H.setValue "U_ConfName", 0, Trim(oRecordSet01.Fields(0).VALUE)
    End Select
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
End Sub
Private Function Item_Error_Message(ErrNum As Integer)
On Error GoTo Item_Error_Message

    If ErrNum = 1 Then
        GoTo Item_Error_Message
    End If

Item_Error_Message:

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "��ǰ�ڵ带 ���� �Է��ϼ���!", "E"
    End If
    
End Function