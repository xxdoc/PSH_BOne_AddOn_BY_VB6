VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM260"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM260.cls
'//  Module         : QM
'//  Description    : �˻�������������
'//  FormType       : PS_QM260
'//  Create Date    : 2013.5.18
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM260H As SAPbouiCOM.DBDataSource '������
Private oDS_PS_QM260L As SAPbouiCOM.DBDataSource '��϶���

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM260.srf")
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
    
    oFormUniqueID01 = "PS_QM260_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "Code" '//UDO����϶�

    oForm01.EnableMenu ("1293"), True     '// �����
    oForm01.EnableMenu ("1287"), True     '// ����
    oForm01.EnableMenu ("1284"), True     '// ���
    oForm01.Freeze True
    
    Call CreateItems
    Call ComboBox_Setting
    
    Call SetDocument(oFromDocEntry01)
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
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim i&
    Dim sCount&, sSeq&
    
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                  
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        

                        If Len(Trim(oForm01.Items("YM").Specific.VALUE) + Trim(oForm01.Items("Seq").Specific.VALUE)) <> 8 Then
                            Sbo_Application.SetStatusBarMessage "���, ������ �ڸ����� 8�ڸ����� �մϴ�.", bmt_Short, True
                            BubbleEvent = False
                            Exit Sub
                        End If
                        sQry = "Select Right('00000000' + Convert(Nvarchar(10),AutoKey),8) From ONNM Where ObjectCode = 'PS_QM260' "
                        oRecordSet01.DoQuery sQry
                        oForm01.Items("Code").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)

                    
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                  
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "ItemCode" Then
                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
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
             If pval.ItemChanged = True Then
               If pval.ColUID = "HisDate" Then
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
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    FormItemEnabled
                    Call AddMatrixRow(0, True)
'                    If oForm01.Mode = fm_ADD_MODE Then
'                        FormItemEnabled
'                    ElseIf oForm01.Mode = fm_OK_MODE Then
'                        FormItemEnabled
'                        Call AddMatrixRow(0, True)
'                    End If
                End If
                If pval.BeforeAction = False And pval.ItemChanged = True Then
                   If pval.ColUID = "HisDate" Then
                      FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                   End If
                End If
            Case et_KEY_DOWN: '//2
        
            Case et_COMBO_SELECT: '//5
                
            Case et_CLICK: '//6
                
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    ElseIf pval.ItemUID = "YM" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    ElseIf pval.ItemUID = "Seq" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    End If
                    
                    If pval.ItemUID = "Mat01" Then
                        '�̷����ڰ� �ٲ�� �� �� �� ���⵵��
                        If pval.ColUID = "HisDate" Then
                           FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    End If
                End If
               
            Case et_MATRIX_LOAD: '//11
                Call AddMatrixRow(oMat01.VisualRowCount, False)
                Call FormItemEnabled
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
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
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim SerialNo    As String
    Dim sSeq$

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oUID = "ItemCode" Then
        sQry = "Select ItemName From OITM Where ItemCode = '" & oForm01.Items("ItemCode").Specific.VALUE & "'"
        oRecordSet01.DoQuery sQry
        oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    End If
    
    If oUID = "YM" Then
        SerialNo = Mid(oForm01.Items("YM").Specific.VALUE, 3, 2) + "-" + Right(Trim(oForm01.Items("YM").Specific.VALUE), 2) + "-" + oForm01.Items("Seq").Specific.VALUE
        oForm01.Items("SerialNo").Specific.VALUE = SerialNo
    End If
    
    If oUID = "Seq" Then
        SerialNo = Mid(oForm01.Items("YM").Specific.VALUE, 3, 2) + "-" + Right(Trim(oForm01.Items("YM").Specific.VALUE), 2) + "-" + oForm01.Items("Seq").Specific.VALUE
        oForm01.Items("SerialNo").Specific.VALUE = SerialNo
    End If
    
    If oUID = "Mat01" Then
    
        oMat01.FlushToDataSource
            
        ' Matrix �ʵ忡 ���� ���� â ����ֱ�
        Select Case oCol
            Case "HisDate"
              oMat01.LoadFromDataSource
              
                '--------------------------------------------------------------------------------------------
                If oRow = oMat01.RowCount And Trim(oDS_PS_QM260L.GetValue("U_HisDate", oRow - 1)) <> "" Then
                '// ���� ���� �߰�
                Call AddMatrixRow(0, False)
                oMat01.Columns("HisDate").Cells(oRow).Click ct_Regular
                End If
                
        End Select
    End If
    Set oRecordSet01 = Nothing
    Exit Sub
       
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�̷����ڸ� Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
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
                  Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
                 oForm01.DataBrowser.BrowseBy = "Code" '//UDO����϶�
            Case "1282": '�߰�
                 oForm01.DataBrowser.BrowseBy = "Code" '//UDO����϶�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource        ' DBDataSource�� ���ڵ尡 ���� �� �����.
                    Call oDS_PS_QM260L.RemoveRecord(oDS_PS_QM260L.Size - 1)     ' ���ڵ� �� ���� �����.
                    oMat01.LoadFromDataSource       ' DBDataSource�� ��Ʈ������ �ø���
                    If oMat01.RowCount = 0 Then
'                        Call AddMatrixRow(1, 0, True)
                        Call AddMatrixRow(1, 0)
                    Else
                        If Trim(oDS_PS_QM260L.GetValue("U_HisDate", oMat01.RowCount - 1)) <> "" Then
'                            Call AddMatrixRow(1, oMat01.RowCount, True)
                            Call AddMatrixRow(1, oMat01.RowCount)
                        End If
                    End If
                End If
            Case "1281": 'ã��
                 Call AddMatrixRow(0, True) '//UDO���
                 Call FormItemEnabled
            Case "1287": '����
                oForm01.Freeze True
                oDS_PS_QM260H.setValue "Code", 0, ""
                oDS_PS_QM260H.setValue "U_Seq", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_QM260L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
            Case "1282": '�߰�
                 Call FormItemEnabled '//UDO���
                 Call AddMatrixRow(0, True) '//UDO���
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư             '�߰���ư Ŭ���� Matrix�� �� �߰�
                 Call FormItemEnabled
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
       
    Set oDS_PS_QM260H = oForm01.DataSources.DBDataSources("@PS_QM260H")
    Set oDS_PS_QM260L = oForm01.DataSources.DBDataSources("@PS_QM260L")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oForm01.Items("YM").Specific.VALUE = Format(Now, "YYYYMM")
'    Call oForm01.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
'    oForm01.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
'    oForm01.DataSources.UserDataSources.Item("DocDate").Value = Format(Now, "YYYYMMDD")
    
'    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    
    
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
        Call oForm01.EnableMenu("1281", True) '//ã��
        Call oForm01.EnableMenu("1282", False) '//�߰�
        oForm01.Items("YM").Enabled = True
        oForm01.Items("Seq").Enabled = True
        oForm01.Items("GaugeNo").Enabled = True
        oForm01.Items("Spec").Enabled = True
        oForm01.Items("TestLoc1").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
        Call oForm01.EnableMenu("1281", True) '//ã��
        Call oForm01.EnableMenu("1282", True) '//�߰�
        oForm01.Items("YM").Enabled = True
        oForm01.Items("Seq").Enabled = True
        oForm01.Items("GaugeNo").Enabled = True
        oForm01.Items("Code").Enabled = True
        oForm01.Items("Spec").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("YM").Enabled = True
        oForm01.Items("Seq").Enabled = True
        oForm01.Items("GaugeNo").Enabled = True
        
        Call oForm01.EnableMenu("1282", True) '//�߰�
        oForm01.Items("Spec").Enabled = True
        
        '//����忡���� �����ۼ���
    End If
    Call oForm01.Freeze(False)
    Exit Sub
FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo AddMatrixRow_Error
    
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//���߰�����
        oRow = oMat01.RowCount
        oDS_PS_QM260L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_QM260L.Offset = oRow
    oDS_PS_QM260L.setValue "LineId", oRow, oRow + 1
    oDS_PS_QM260L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//������� ��������ɿ��ΰ˻�
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineId").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_QM260L.RemoveRecord(oDS_PS_QM260L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call AddMatrixRow(0)
            Else
                If Trim(oDS_PS_QM260L.GetValue("U_HisDate", oMat01.RowCount - 1)) <> "" Then
                    Call AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

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
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 Then
        If oDS_PS_QM260L.GetValue("U_HisDate", 0) = "" Then
            ErrNum = 1
            GoTo MatrixSpaceLineDel_Error
        End If
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '// �Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '// �̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 2
            oDS_PS_QM260L.Offset = i
            Select Case True
                Case oDS_PS_QM260L.GetValue("U_HisDate", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
                Case oDS_PS_QM260L.GetValue("U_Location", i) = ""
                    ErrNum = 3
                    GoTo MatrixSpaceLineDel_Error
            End Select
        Next i
        
        If oDS_PS_QM260L.GetValue("U_HisDate", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_QM260L.RemoveRecord oMat01.VisualRowCount - 1
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
        MDC_Com.MDC_GF_Message "���ε���Ÿ�� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "�̷����ڴ� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "�����ġ�� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_QM260H.GetValue("U_GaugeNo", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
'        Case oDS_PS_QM260H.GetValue("U_Number", 0) = ""
'            ErrNum = 2
'            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_QM260H.GetValue("U_Spec", 0) = ""
            ErrNum = 3
        Case oDS_PS_QM260H.GetValue("U_TestLoc1", 0) = ""
            ErrNum = 4
    End Select
   
   HeaderSpaceLineDel = True
   
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "��������ȣ�� �ʼ��Է� �����Դϴ�.", "E"
    End If
    If ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "������ �ʼ��Է� �����Դϴ�.", "E"
    End If
    If ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "ġ���� �ʼ��Է� �����Դϴ�.", "E"
    End If
    If ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "�˻�������� �ʼ��Է� �����Դϴ�.", "E"
    End If
    If ErrNum = "" Then
        MDC_Com.MDC_GF_Message "������ �Ǿ����ϴ�.", "W"
    End If
    HeaderSpaceLineDel = False
End Function

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

Private Sub SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo SetDocument_Error

    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim sCount&, sSeq&, i&
    
    If (oFromDocEntry01 = "") Then
        Call FormItemEnabled
        Call AddMatrixRow(0, True)
    Else
        oForm01.Mode = fm_FIND_MODE
        Call FormItemEnabled
        oForm01.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
        oForm01.Items("Code").Enabled = False
        
    End If
    Exit Sub
SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
