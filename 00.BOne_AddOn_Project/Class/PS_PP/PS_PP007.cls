VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP007.cls
'//  Module         : PP
'//  Desc           : ���������
'//  FormType       : PS_PP007
'//  Create Date    : 2010.09.27
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit         '��� ���ؿ��� �ش��ϴ� ����� ��� ������ ���������� �����ϵ��� �� �� ����մϴ�.

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_PP007H       As SAPbouiCOM.DBDataSource    '������ 'Represents a SAP Business One database table to be attached to a form.
Private oDS_PS_PP007L       As SAPbouiCOM.DBDataSource    '��϶���
Private oDS_PS_PP007M       As SAPbouiCOM.DBDataSource

Private oLast_Item_UID   As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID    As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row    As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��
Private oSeq&

'Private oFormMode As String

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP007.srf")
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
    
    oFormUniqueID01 = "PS_PP007_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    ' ������ �Ӽ� � ��ü ������ �����մϴ�

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
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
    Call FormItemEnabled
    Call FormClear '//UDO����϶�
    Call AddMatrixRow(0, oMat01.RowCount, True) '//UDO����϶�
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//�޴�����
    
    oForm01.Update      ' Reloads the form from the application.
                        ' Use this method to update an object from the database.

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
                        
                        '// Input Code, Name
                        oDS_PS_PP007H.setValue "Code", 0, Trim(oDS_PS_PP007H.GetValue("U_ItmBsort", 0))
                        oDS_PS_PP007H.setValue "Name", 0, Trim(oDS_PS_PP007H.GetValue("U_ItmBname", 0))
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
            
                ' ��ǰ�з��ڵ�
                If oForm01.Items("ItmBsort").Specific.VALUE = "" Then
                    If pval.ItemUID = "ItmBsort" And pval.CharPressed = 9 Then
                        oForm01.Items("ItmBsort").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' �����ڵ�
                If pval.ItemUID = "Mat01" And pval.ColUID = "CpCode" And pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                 End If
                
                
                
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then                '
                    If pval.ItemUID = "Mat01" And pval.ColUID = "JsCode" Then
                        FlushToItemValue pval.ItemUID, pval.ColUID, pval.Row, pval.ColUID
                    End If
                End If

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
                If pval.ItemChanged = True Then
                    ' �����з�
                    If pval.ItemUID = "ItmBsort" And pval.ItemChanged = True Then
                        sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oForm01.Items("ItmBname").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                    If (pval.ItemUID = "Mat01" And (pval.ColUID = "CpCode")) Then
                        sQry = "Select U_CpName From [@PS_PP001L] Where U_CpCode = '" & Trim(oMat01.Columns("CpCode").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oMat01.Columns("CpName").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    End If
                    
                End If

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
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                oForm01.Items("ItmBsort").Click ct_Regular
                
            Case "1287": '����
                oForm01.Freeze True
                oDS_PS_PP007H.setValue "Code", 0, ""
                oDS_PS_PP007H.setValue "Name", 0, ""
                oDS_PS_PP007H.setValue "U_ItmBsort", 0, ""
                oDS_PS_PP007H.setValue "U_ItmBname", 0, ""
         
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_PP007L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("LineNum").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        AddMatrixRow 1, oMat01.RowCount, True
'                    End If
'                End If
                
            Case "1293"     '�����
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource        ' DBDataSource�� ���ڵ尡 ���� �� �����.
                    Call oDS_PS_PP007L.RemoveRecord(oDS_PS_PP007L.Size - 1)     ' ���ڵ� �� ���� �����.
                    oMat01.LoadFromDataSource       ' DBDataSource�� ��Ʈ������ �ø���
                    If oMat01.RowCount = 0 Then
                        Call AddMatrixRow(1, 0, True)
                    Else
                        If Trim(oDS_PS_PP007L.GetValue("U_JsCode", oMat01.RowCount - 1)) <> "" Then
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
'                FormItemEnabled
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
    
    Set oDS_PS_PP007H = oForm01.DataSources.DBDataSources("@PS_PP007H")
    Set oDS_PS_PP007L = oForm01.DataSources.DBDataSources("@PS_PP007L")
    
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// �����
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordset01.Fields(0).Value), Trim(oRecordset01.Fields(1).Value)
        oRecordSet01.MoveNext
    Loop
        
    '//���̵� ����� ����
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
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

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
'        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("ItmBsort").Enabled = True
'        oForm01.Items("ItmMcode").Enabled = True
'        oForm01.Items("CntcCode").Enabled = True
'        oMat01.Columns("Price1").Editable = True
'        oMat01.Columns("Price2").Editable = True
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
'        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("ItmBsort").Enabled = True
'        oForm01.Items("ItmMcode").Enabled = True
'        oForm01.Items("CntcCode").Enabled = True
'        oMat01.Columns("Price1").Editable = False
'        oMat01.Columns("Price2").Editable = False
        
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//����忡���� �����ۼ���
'            oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("ItmBsort").Enabled = False
'        oForm01.Items("ItmMcode").Enabled = False
'        oForm01.Items("CntcCode").Enabled = False
'        oMat01.Columns("Price1").Editable = True
'        oMat01.Columns("Price2").Editable = True
    End If
    Exit Sub
FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oSeq As Integer, ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo AddMatrixRow_Error
   Select Case oSeq
       Case 0:
            'oMat01.AddRow
            'oMat01.Columns("LineNum").Cells(oRow + 1).Specific.Value = oRow + 1
            oDS_PS_PP007L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
       Case 1:
            oDS_PS_PP007L.InsertRecord oRow
            oDS_PS_PP007L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
   End Select
'AddMatrixRow_Error:
'    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP007'", "")
    If DocNum = 0 Then
'        oForm01.Items("DocEntry").Specific.String = 1
        oDS_PS_PP007H.setValue "DocEntry", 0, "1"
    Else
'        oForm01.Items("DocEntry").Specific.String = DocNum
        oDS_PS_PP007H.setValue "DocEntry", 0, DocNum    ' ȭ�鿡 ������ �ȵǱ� ����
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


Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_PP007H.GetValue("U_ItmBsort", 0) = "" Or oDS_PS_PP007H.GetValue("U_ItmBname", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    HeaderSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "��з� �Ǵ� ��з����� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
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
        For i = 0 To oMat01.VisualRowCount - 2
            oDS_PS_PP007L.Offset = i
            Select Case True
                Case oDS_PS_PP007L.GetValue("U_JsCode", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
                Case oDS_PS_PP007L.GetValue("U_JsName", i) = ""
                    ErrNum = 3
                    GoTo MatrixSpaceLineDel_Error
                Case oDS_PS_PP007L.GetValue("U_CpCode", i) = ""
                    ErrNum = 4
                    GoTo MatrixSpaceLineDel_Error
    
                Case oDS_PS_PP007L.GetValue("U_CpName", i) = ""
                    ErrNum = 5
                    GoTo MatrixSpaceLineDel_Error

            End Select
        Next i
        
        If oDS_PS_PP007L.GetValue("U_JsCode", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_PP007L.RemoveRecord oMat01.VisualRowCount - 1
        End If
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '���� �����Ͽ����� DB������ �ҽ��� �ٽ� �����´�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "��������ڵ� �����ʹ� �ʼ��Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "�������� �����ʹ� �ʼ��Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "�����ڵ� �����ʹ� �ʼ��Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "������ �����ʹ� �ʼ��Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, ByVal oCID As String, Optional oRow As Long, Optional oCol As String)
    Dim i$
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "Mat01"
            Select Case oCID
                Case "JsCode"
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 2) And _
                    Trim(oMat01.Columns("JsCode").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        AddMatrixRow 1, oMat01.RowCount, True
                        oMat01.Columns("JsCode").Cells(oRow).Click ct_Regular
                    End If

           End Select
    End Select
    Set oRecordSet01 = Nothing
End Sub
