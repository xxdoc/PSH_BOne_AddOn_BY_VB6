VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM025"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM025.cls
'//  Module         : QM
'//  Description    : �˻缺���� ��ǰó����
'//  FormType       : PS_QM025
'//  Create Date    : 2011.01.20
'//  Modified Date  :
'//  Creator        : NGY
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_QM025L As SAPbouiCOM.DBDataSource    '��϶���

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&
Private oForm01_Mode&
Private oSeq&

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM025.srf")
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
    
    oFormUniqueID01 = "PS_QM025_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    
    oForm01_Mode = fm_ADD_MODE '//����ʱ�ȭ
    oForm01.Freeze True
        
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    
    Call LoadCaption
    
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False         '// �ݱ�
    oForm01.EnableMenu ("1287"), False        '// ����
    oForm01.EnableMenu ("1284"), False         '// ���
    oForm01.EnableMenu ("1293"), False         '// �����
        
    oForm01.Update
    oForm01.Freeze False
    oForm01.Visible = True
    
    Set oXmlDoc01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    If (oForm01 Is Nothing) = False Then
        Set oForm01 = Nothing
    End If
    MDC_Com.MDC_GF_Message "LoadForm_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim TempForm01 As Variant
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
        
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "Mat01" Then
                        If oMat01.Columns("CardCode").Cells(pval.Row).Specific.VALUE = "" Then
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
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If oForm01_Mode = fm_UPDATE_MODE Then
                        If CH_QM020(pval) = False Then '//�׸� Update
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oForm01_Mode = fm_OK_MODE
                        oMat01.Clear
                        oDS_PS_QM025L.Clear
                        Call LoadData
                        Call LoadCaption
                    ElseIf oForm01_Mode = fm_OK_MODE Then
                        oForm01.Close
                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    If HeaderSpaceLineDel() = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call LoadData
                    Call LoadCaption
                ElseIf pval.ItemUID = "Btn03" Then
                    If oForm01_Mode = fm_UPDATE_MODE Then
                        If CH_CARDCODE(pval) = False Then    '//��ǰó �ϰ�����
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
                    
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "BPLId" Then
                    oMat01.Clear
                    oDS_PS_QM025L.Clear
                    oForm01_Mode = fm_OK_MODE
                    Call LoadCaption
                End If
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "CardCode" Then
                            oForm01_Mode = fm_UPDATE_MODE
                            Call LoadCaption
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                            oForm01.Freeze False
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_QM025L = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    If ErrNum = 101 Then
        ErrNum = 0
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
        BubbleEvent = False
    Else
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
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
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//�۾�
    ElseIf (eventInfo.BeforeAction = False) Then
        '//�۾�
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("PackNoF").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("PackNoT").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "PACKING ���۹�ȣ�� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "PACKING �����ȣ�� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_PS_QM025L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
        
    '// ��Ʈ���� ��ü �Ҵ�
    Set oMat01 = oForm01.Items("Mat01").Specific
        
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

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
        oRecordSet01.MoveNext
    Loop
        
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Initialization()
On Error GoTo Initialization_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//���̵� ����� ����
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
'    oForm01.Items("PackNoF").Specific.VALUE = "20100101001"
'    oForm01.Items("PackNoT").Specific.VALUE = "20111231001"
    '//���̵� ��� ����
    'oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
    '//���̵� �μ� ����
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim ItemCode$, Qty&
    Dim Calculate_Weight As Currency
    Dim Calculate_Qty As Currency
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim CardCode$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "Mat01"
            If oCol = "CardCode" Then
                CardCode = Trim(oMat01.Columns("CardCode").Cells(oRow).Specific.VALUE)
                sQry = "Select CardName From OCRD Where CardCode = '" & Trim(CardCode) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("CardName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
'                For i = 1 To oMat01.RowCount
'                    oMat01.Columns("CardCode").Cells(i).Specific.VALUE = CardCode
'                    oMat01.Columns("CardName").Cells(i).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'                Next
            End If
        
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01_Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "����"
    ElseIf oForm01_Mode = fm_OK_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "Ȯ��"
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId$, PackNoF$, PackNoT$
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    PackNoF = Trim(oForm01.Items("PackNoF").Specific.VALUE)
    PackNoT = Trim(oForm01.Items("PackNoT").Specific.VALUE)
    
    sQry = "EXEC [PS_QM025_01] '" & BPLId & "', '" & PackNoF & "', '" & PackNoT & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_QM025L.Clear
        
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
        oForm01_Mode = fm_OK_MODE
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
'    Dim ProgBar01 As SAPbouiCOM.ProgressBar
'    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_QM025L.Size Then
            oDS_PS_QM025L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_QM025L.Offset = i
        oDS_PS_QM025L.setValue "U_LineNum", i, i + 1
        oDS_PS_QM025L.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("U_OrdNum").VALUE)
        oDS_PS_QM025L.setValue "U_ColDt01", i, Format(Trim(oRecordSet01.Fields("U_InspDate").VALUE), "YYYYMMDD")
        oDS_PS_QM025L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("U_CardCode").VALUE)
        oDS_PS_QM025L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("U_CardName").VALUE)
        oDS_PS_QM025L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
        oDS_PS_QM025L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("U_ItemName").VALUE)
        
        oRecordSet01.MoveNext
'        ProgBar01.VALUE = ProgBar01.VALUE + 1
'        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "�� ��ȸ��...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns

    'ProgBar01.Stop
    oForm01.Freeze False
    oForm01_Mode = fm_ADD_MODE
    
'    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
'    ProgBar01.Stop
 '   Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function CH_QM020(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo CH_QM020_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    Dim BPLId$, PackNoF$, PackNoT$
    Dim OrdNum$, CardCode$, CardName$
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    PackNoF = Trim(oForm01.Items("PackNoF").Specific.VALUE)
    PackNoT = Trim(oForm01.Items("PackNoT").Specific.VALUE)
    
    oMat01.FlushToDataSource
    
    
    Sbo_Company.StartTransaction
    
    For i = 0 To oMat01.RowCount - 1
        CardCode = Trim(oDS_PS_QM025L.GetValue("U_ColReg02", i)) '//��ǰó�ڵ�
        sQry = "Select CardName From OCRD Where CardCode = '" & CardCode & "'"
        oRecordSet01.DoQuery sQry
        
        If oRecordSet01.RecordCount = 0 Then
            MDC_Com.MDC_GF_Message "��ǰó�� �߸��� �ڷᰡ �ֽ��ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
            Set oRecordSet01 = Nothing
            If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
            CH_QM020 = False
            Exit Function
        End If
    Next
    
    For i = 0 To oMat01.RowCount - 1
        
        OrdNum = Trim(oDS_PS_QM025L.GetValue("U_ColReg01", i)) '//�۾����ù�ȣ
        CardCode = Trim(oDS_PS_QM025L.GetValue("U_ColReg02", i)) '//��ǰó�ڵ�
        sQry = "Select CardName From OCRD Where CardCode = '" & CardCode & "'"
        oRecordSet01.DoQuery sQry
        CardName = Trim(oRecordSet01.Fields(0).VALUE)
        
        
        sQry = "Update [@PS_QM020H]"
        sQry = sQry & " set U_CardCode = '" + CardCode + "',"
        sQry = sQry & " U_CardName = '" + CardName + "'"
        sQry = sQry & " Where U_OrdNum = '" + OrdNum + "'"
        
        oRecordSet01.DoQuery sQry
    Next
    
        
    Sbo_Company.EndTransaction wf_Commit
    

    
    MDC_Com.MDC_GF_Message "�˻缺���� ��ǰó���� �Ϸ�!", "S"
    Set oRecordSet01 = Nothing
    CH_QM020 = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CH_QM020_Error:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Set oRecordSet01 = Nothing
    CH_QM020 = False
    MDC_Com.MDC_GF_Message "CH_QM020_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function CH_CARDCODE(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo CH_CARDCODE_Error

    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim CardCode$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    CardCode = Trim(oMat01.Columns("CardCode").Cells(1).Specific.VALUE)
    sQry = "Select CardName From OCRD Where CardCode = '" & Trim(CardCode) & "'"
    oRecordSet01.DoQuery sQry
                    
    For i = 1 To oMat01.RowCount
        oMat01.Columns("CardCode").Cells(i).Specific.VALUE = CardCode
        oMat01.Columns("CardName").Cells(i).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    Next
            
    Set oRecordSet01 = Nothing
    CH_CARDCODE = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CH_CARDCODE_Error:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Set oRecordSet01 = Nothing
    CH_CARDCODE = False
    MDC_Com.MDC_GF_Message "CH_CARDCODE_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

