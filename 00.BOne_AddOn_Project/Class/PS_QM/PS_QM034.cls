VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM034"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM034.cls
'//  Module         : QM
'//  Description    : M/G��������԰˻���ص��
'//  FormType       : PS_QM034
'//  Create Date    : 2013.08.20
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM034H       As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_QM034L       As SAPbouiCOM.DBDataSource    '��϶���

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM034.srf")
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
    
    oFormUniqueID01 = "PS_QM034_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm01.DataBrowser.BrowseBy = "Code"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call FormClear
'    Call Matrix_AddRow(1, 0, True) 'oMat01
'    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), True         '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1287"), True         '// ����
    oForm01.EnableMenu ("1284"), False        '// ���
    oForm01.EnableMenu ("1293"), False        '// �����
        
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_PS_QM034H = oForm01.DataSources.DBDataSources("@PS_QM034H")
'    Set oDS_PS_QM034L = oForm01.DataSources.DBDataSources("@PS_QM034L")
    
    '// ��Ʈ���� ��ü �Ҵ�
'    Set oMat01 = oForm01.Items("Mat01").Specific
'    oDS_PS_QM034H.setValue "U_RegiDate", 0, Format(Now, "YYYYMMDD")
   
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
    Dim oRecordSet      As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// �����
'    Set oCombo = oForm01.Items("BPLId").Specific
'    sQry = "SELECT BPLId, BPLName From [OBPL]"
'    oRecordset01.DoQuery sQry
'    Do Until oRecordset01.EOF
'        oCombo.ValidValues.Add Trim(oRecordset01.Fields(0).Value), Trim(oRecordset01.Fields(1).Value)
'        oRecordset01.MoveNext
'    Loop
'    Call oForm01.Items("TS_Gbn").Specific.ValidValues.Add("10", "Kgf/m2")
'    Call oForm01.Items("TS_Gbn").Specific.ValidValues.Add("20", "N/mm2")
    
    
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                        If oForm01.Mode = fm_ADD_MODE Then
                            If Check_Exist = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//���
'                    If pval.ItemUID = "CardCode" Then
'                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
'                    End If
'                    If pval.ItemUID = "ItemCode" Then
'                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
''                            Sbo_Application.ActivateMenuItem ("7425")
''                            BubbleEvent = False
'                            Dim ChildForm01 As Variant
'                            Set ChildForm01 = New PS_SM010
'                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
'                        End If
'                    End If
                    
                    
                    
                    '//����
'                    If pval.ItemUID = "Mat01" Then
'                        If pval.ColUID = "PP070No" Then
'                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.Value = "" Then
'                                Sbo_Application.ActivateMenuItem ("7425")
'                                BubbleEvent = False
'                            End If
'                        End If
'                    End If
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
        
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                        FormItemEnabled
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
                    End If
'                ElseIf pval.ItemUID = "Btn03" Then
'                    If HeaderSpaceLineDel = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
'                    Call PS_QM034_Print_Report01
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                 If pval.ItemUID = "TS_Gbn" Then
                    If Trim(oForm01.Items("TS_Gbn").Specific.VALUE) = "10" Then
                        oForm01.Items("s6").Specific.Caption = "���尭��(kgf/��)"
                    ElseIf Trim(oForm01.Items("TS_Gbn").Specific.VALUE) = "20" Then
                        oForm01.Items("s6").Specific.Caption = "���尭��(N/��)"
                    End If
                    
                 End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//���
'                    If pval.ItemUID = "CardCode" Or pval.ItemUID = "ItemCode" Then
'                        FlushToItemValue pval.ItemUID
'                    End If
                    '//����
'                    If pval.ItemUID = "Mat01" And (pval.ColUID = "PP070No") Then
'                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
'                    End If
                End If
            Case et_MATRIX_LOAD: '//11
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
            Case "1285": '����
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1285": '����
            Case "1287": '����
'                oForm01.Freeze True
'                oDS_PS_QM034H.setValue "U_CardCode", 0, ""
'                oDS_PS_QM034H.setValue "U_CardName", 0, ""
'                oDS_PS_QM034H.setValue "U_ItemCode", 0, ""
'                oDS_PS_QM034H.setValue "U_ItemName", 0, ""
'                FormItemEnabled
'                Call FormClear
'                oForm01.Items("CardCode").Click ct_Regular
'                oForm01.Freeze False
            
            Case "1293": '�����
'                If oMat01.RowCount <> oMat01.VisualRowCount Then
'                    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'                    '//�Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
'                    '//�̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
'                    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'                    For i = 0 To oMat01.VisualRowCount - 1
'                        oMat01.Columns("LineId").Cells(i + 1).Specific.Value = i + 1
'                    Next i
'
'                    oMat01.FlushToDataSource
'                    oDS_PS_QM034L.RemoveRecord oDS_PS_QM034L.Size - 1   '// Mat1�� ����������(�����) ����
'                    oMat01.Clear
'                    oMat01.LoadFromDataSource
'                End If
            
            Case "1281": 'ã��
'                FormItemEnabled
'                oForm01.Items("CardCode").Click ct_Regular
                
            Case "1282": '�߰�
'                Call FormItemEnabled
'                Call FormClear
'                oForm01.Items("CardCode").Click ct_Regular
                
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("PP070No").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        If oDS_PS_QM034H.GetValue("Status", 0) = "O" Then
'                            Matrix_AddRow 1, oMat01.RowCount, False
'                        End If
'                    End If
'                End If
  
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        
    ElseIf (eventInfo.BeforeAction = False) Then
        '//�۾�
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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

Private Sub FormItemEnabled()
'    If oForm01.Mode = fm_FIND_MODE Then
'        oForm01.Items("CardCode").Enabled = True
'        oForm01.Items("ItemCode").Enabled = True
'
'    ElseIf oForm01.Mode = fm_ADD_MODE Then
'        oForm01.Items("CardCode").Enabled = True
'        oForm01.Items("ItemCode").Enabled = True
'
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("CardCode").Enabled = False
'        oForm01.Items("ItemCode").Enabled = False
'    End If
'
'    If Trim(oForm01.Items("TS_Gbn").Specific.VALUE) = "10" Then
'        oForm01.Items("s6").Specific.Caption = "���尭��(kgf/��)"
'    ElseIf Trim(oForm01.Items("TS_Gbn").Specific.VALUE) = "20" Then
'        oForm01.Items("s6").Specific.Caption = "���尭��(N/��)"
'    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM034'", "")
    If DocNum = 0 Then
        oForm01.Items("Code").Specific.VALUE = 1
        oDS_PS_QM034H.setValue "Code", 0, 1
    Else
        'oForm01.Items("Code").Specific.Value = DocNum
        oDS_PS_QM034H.setValue "Code", 0, DocNum
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oMat As Integer, ByVal oRow As Long, Optional Insert_YN As Boolean)
On Error GoTo Matrix_AddRow_Error
        
    Select Case oMat
        Case 1: 'oMat01
'            If Insert_YN = False Then
'                oRow = oMat01.RowCount
'                oDS_PS_QM034L.InsertRecord (oRow)
'            End If
'            '���Գ���
'            oDS_PS_QM034L.Offset = oRow
'            oDS_PS_QM034L.setValue "LineId", oRow, oRow + 1
'            oDS_PS_QM034L.setValue "U_PP070No", oRow, ""
'            oDS_PS_QM034L.setValue "U_ItemCode", oRow, ""
'            oDS_PS_QM034L.setValue "U_ItemName", oRow, ""
'            oDS_PS_QM034L.setValue "U_Size", oRow, ""
'            oDS_PS_QM034L.setValue "U_Mark", oRow, ""
'            oDS_PS_QM034L.setValue "U_Qty", oRow, ""
'            oDS_PS_QM034L.setValue "U_Weight", oRow, ""
'            oDS_PS_QM034L.setValue "U_DocDate", oRow, ""
'            oMat01.LoadFromDataSource
        
    End Select
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Matrix_AddRow_Error:
    MDC_Com.MDC_GF_Message "Matrix_AddRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
'        Case "CardCode"
'            sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" & Trim(oDS_PS_QM034H.GetValue("U_CardCode", 0)) & "'"
'            oRecordSet.DoQuery sQry
'            oDS_PS_QM034H.setValue "U_CardName", 0, Trim(oRecordSet.Fields(0).VALUE)
'
'
'        Case "ItemCode"
'            sQry = "Select ItemName From OITM Where ItemCode = '" & Trim(oDS_PS_QM034H.GetValue("U_ItemCode", 0)) & "'"
'            oRecordSet.DoQuery sQry
'            oDS_PS_QM034H.setValue "U_ItemName", 0, Trim(oRecordSet.Fields(0).VALUE)
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
        
        End Select
    End If
     
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
    '    Case Trim(oDS_PS_QM034H.GetValue("U_CardCode", 0)) = ""
    '        ErrNum = 1
    '        GoTo HeaderSpaceLineDel_Error
    '    Case Trim(oDS_PS_QM034H.GetValue("U_CardName", 0)) = ""
    '        ErrNum = 2
    '        GoTo HeaderSpaceLineDel_Error
    '    Case Trim(oDS_PS_QM034H.GetValue("U_ItemCode", 0)) = ""
    '        ErrNum = 3
    '        GoTo HeaderSpaceLineDel_Error
    '    Case Trim(oDS_PS_QM034H.GetValue("U_ItemName", 0)) = ""
    '        ErrNum = 4
    '        GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������ �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "�������� �����ϴ�. �����ڵ带 Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "��ǰ�� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "��ǰ���� �����ϴ�. ��ǰ�ڵ带 Ȯ���Ͽ� �ֽʽÿ�.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// ����
    If oMat01.VisualRowCount = 1 Then '// MAT01�� ���� �ִ��� Ȯ�� (ErrorNumber : 1)
'        ErrNum = 1
'        GoTo MatrixSpaceLineDel_Error
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//������ �� �ϳ��� ���� i=0���� �����ϹǷ� �ϳ��� ���Ƿ�
    '//oMat01.RowCount - 2�� �ȴ�..�ݵ�� ��� ���� �ϴ� �ʼ����� Ȯ���Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    If oMat01.VisualRowCount > 0 Then
'        For i = 0 To oMat01.VisualRowCount - 2      '// Mat1�� �Է°��� �ùٸ��� ������ Ȯ�� (ErrorNumber : 2)
'            oDS_PS_QM034L.Offset = i
'            If Trim(oDS_PS_QM034L.GetValue("U_PP070No", i)) = "" Then
'                ErrNum = 2
'                oMat01.Columns("PP070No").Cells(i + 1).Click ct_Regular
'                GoTo MatrixSpaceLineDel_Error
'            End If
'        Next i
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//�Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '//�̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'    If oMat01.VisualRowCount > 0 Then
'        oDS_PS_QM034L.RemoveRecord oDS_PS_QM034L.Size - 1   '// Mat1�� ����������(�����) ����
'    End If
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '���� �����Ͽ����� DB������ �ҽ��� �ٽ� �����´�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���� �����Ͱ� �����ϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "��ũ���幮�� ��ȣ�� �ʼ��Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function Check_Exist() As Boolean
On Error GoTo Check_Exist_Error
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    
'    sQry = "select * from [@PS_QM034H] where U_CardCode = '" & Trim(oDS_PS_QM034H.GetValue("U_CardCode", 0)) & "' "
'    sQry = sQry & "and U_ItemCode = '" & Trim(oDS_PS_QM034H.GetValue("U_ItemCode", 0)) & "'"
'    oRecordSet.DoQuery sQry
'
'    If oRecordSet.RecordCount <> 0 Then
'        ErrNum = 1
'        GoTo Check_Exist_Error
'    End If
    
    Set oRecordSet = Nothing
    Check_Exist = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Check_Exist_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�̹� ������ '���� & ��ǰ'�� �����մϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Check_Exist = False
End Function

'Private Sub PS_QM034_Print_Report01()
'On Error GoTo PS_QM034_Print_Report01_Error
'    Dim WinTitle                    As String
'    Dim ReportName                  As String
'    Dim sQry                        As String
'    Dim sQry01                      As String
'    Dim i                           As Long
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'
'    Dim CardCode$, ItemCode$
'
'    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
'    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
'
'    If CardCode = "" Then CardCode = "%"
'    If ItemCode = "" Then ItemCode = "%"
'
'    Call ConnectODBC
'    WinTitle = "[PS_QM034_01] �˻�԰�"
'    ReportName = "PS_QM034_01.rpt"
'
'    sQry = "EXEC [PS_QM034_01] '" & CardCode & "','" & ItemCode & "'"
'
'    ReDim gRpt_Formula(1)
'    ReDim gRpt_Formula_Value(1)
'
'    ReDim gRpt_SRptSqry(1)
'    ReDim gRpt_SRptName(1)
'    ReDim gRpt_SFormula(1, 1)
'    ReDim gRpt_SFormula_Value(1, 1)
'
'    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
'        Sbo_Application.SetStatusBarMessage "gCryReport_Action : ����!", bmt_Short, True
'    End If
'    Exit Sub
'PS_QM034_Print_Report01_Error:
'    Sbo_Application.SetStatusBarMessage "PS_QM034_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Sub


