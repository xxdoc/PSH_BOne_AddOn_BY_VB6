VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_CO120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_CO120.cls
'//  Module         : CO
'//  Description    : ������ �������
'//  FormType       : PS_CO120
'//  Create Date    : 2010.11.17
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_CO120H As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_CO120L As SAPbouiCOM.DBDataSource    '��϶���

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_CO120.srf")
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
    
    oFormUniqueID01 = "PS_CO120_" & GetTotalFormsCount
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
'    Call Add_MatrixRow(0, True)
    Call FormItemEnabled
    
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
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim Calculate_Weight As Double
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_CO111
                    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                        Dim YM$, BPLID$, Code$
                        YM = Trim(oDS_PS_CO120H.GetValue("U_YM", 0))
                        BPLID = Trim(oDS_PS_CO120H.GetValue("U_BPLId", 0))
                        Code = YM & BPLID
                        Call oDS_PS_CO120H.setValue("Code", 0, Code)
                        Call oDS_PS_CO120H.setValue("Name", 0, Code)
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
                
                If pval.Row = 0 Then

                    '����
                    oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                    oMat01.FlushToDataSource
    
                End If
            
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
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                ElseIf pval.ItemUID = "Btn01" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call LoadData
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
'et_MATRIX_LOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_MATRIX_LOAD: '//11
                Add_MatrixRow oMat01.RowCount, False
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
                Set oDS_PS_CO120H = Nothing
                Set oDS_PS_CO120L = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
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
                oForm01.Freeze True
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_CO120L.RemoveRecord oDS_PS_CO120L.Size - 1       '// Mat01�� ����������(�����) ����
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    
                    If oMat01.Columns("OrdMgNum").Cells(oMat01.RowCount).Specific.VALUE <> "" Then
                        Call Add_MatrixRow(oMat01.RowCount, False)
                    End If
                End If
                oForm01.Freeze False
            Case "1281": 'ã��
                oForm01.Freeze True
                Call FormItemEnabled
'                oForm01.Items("CycleCod").Click ct_Regular
                oForm01.Freeze False
            Case "1282": '�߰�
                oForm01.Freeze True
                Call FormItemEnabled
                Add_MatrixRow 0, True
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                oForm01.Freeze True
                Call FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("CycleCod").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        Add_MatrixRow oMat01.RowCount, False
'                    End If
'                End If
                oForm01.Freeze False
            Case "1287": '// ����
                oForm01.Freeze True
                oDS_PS_CO120H.setValue "Code", 0, ""
                oDS_PS_CO120H.setValue "Name", 0, ""
                oDS_PS_CO120H.setValue "U_YM", 0, ""
                oDS_PS_CO120H.setValue "U_BPLId", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_CO120L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    oForm01.Freeze False
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_PS_CO120H = oForm01.DataSources.DBDataSources("@PS_CO120H")
    Set oDS_PS_CO120L = oForm01.DataSources.DBDataSources("@PS_CO120L")
    
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
    
    '// �������
'    Set oCombo = oForm01.Items("ClsPrd").Specific
'    sQry = "SELECT Code, Name From [OFPR]"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
    
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

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList ����
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CF_ChooseFromList_Error:
    MDC_Com.MDC_GF_Message "CF_ChooseFromList_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//���߰�����
        oDS_PS_CO120L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_CO120L.offset = oRow
    oDS_PS_CO120L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sRow As Long
    Dim sSeq$
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Select Case oUID
        Case "Mat01"
'            If oCol = "CycleCod" Then
'                oForm01.Freeze True
'                oMat01.FlushToDataSource
'
'                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("CycleCod").Cells(oRow).Specific.Value) <> "" Then
'                    oMat01.FlushToDataSource
'                    Call Add_MatrixRow(oMat01.RowCount, False)
'                    oMat01.Columns("CycleCod").Cells(oRow).Click ct_Regular
'                End If
'
''                sQry = "Select ItemName, FrgnName From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.Value) & "'"
''                oRecordSet01.DoQuery sQry
''                oMat01.Columns("ItemName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(0).Value)
''                oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(1).Value)
''
''                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
'                oForm01.Freeze False
'            End If
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���Ű��������� ��ҵǾ��ų� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_CO120H.GetValue("U_YM", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_CO120H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "��������� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "������� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
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
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 Then
'        If oDS_PS_CO120L.GetValue("U_CycleCod", 0) = "" Then
'            ErrNum = 2
'            GoTo MatrixSpaceLineDel_Error
'        End If
    End If
    
    For i = 0 To oMat01.VisualRowCount - 2
'        Select Case True
'            Case oDS_PS_CO120L.GetValue("U_ItemCode", i) = ""
'                ErrNum = 2
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_CO120L.GetValue("U_Qty", i) = "" Or oDS_PS_CO120L.GetValue("U_Qty", i) = 0
'                ErrNum = 3
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_CO120L.GetValue("U_Weight", i) = ""
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_CO120L.GetValue("U_Price", i) = 0
'                ErrNum = 5
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_CO120L.GetValue("U_LinTotal", i) = 0
'                ErrNum = 6
'                GoTo MatrixSpaceLineDel_Error
'        End Select
    Next
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "ù���ο� ��λ���Ŭ �ڵ尡 �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "������ �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "�߷��� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "�ܰ��� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "�ݾ��� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error
    Dim i&
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_CO120L.GetValue("U_CycleCod", i)) = "" Then
            oDS_PS_CO120L.RemoveRecord i   '// Mat01�� ����������(�����) ����
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim YM$, BPLID$
    
    YM = Trim(oForm01.Items("YM").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)

    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet01.RecordCount, False)

    sQry = "EXEC [PS_CO120_01] '" & YM & "','" & BPLID & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_CO120L.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Call oForm01.Freeze(False)
        Exit Sub
'    ElseIf (oRecordSet01.Fields("ErrMessage").VALUE <> "") Then
'        Call Sbo_Application.MessageBox(oRecordSet01.Fields("ErrMessage").VALUE)
'        Set oRecordSet01 = Nothing
'        Call oForm01.Freeze(False)
'        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_CO120L.Size Then
            oDS_PS_CO120L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_CO120L.offset = i
        oDS_PS_CO120L.setValue "U_LineNum", i, i + 1
        oDS_PS_CO120L.setValue "U_POEntry", i, Trim(oRecordSet01.Fields("POEntry").VALUE)
        oDS_PS_CO120L.setValue "U_POLine", i, Trim(oRecordSet01.Fields("POLine").VALUE)
        oDS_PS_CO120L.setValue "U_Sequence", i, Trim(oRecordSet01.Fields("Sequence").VALUE)
        oDS_PS_CO120L.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)
        oDS_PS_CO120L.setValue "U_ItemName", i, Trim(oRecordSet01.Fields("ItemName").VALUE)
        oDS_PS_CO120L.setValue "U_CpCode", i, Trim(oRecordSet01.Fields("CpCode").VALUE)
        oDS_PS_CO120L.setValue "U_CpName", i, Trim(oRecordSet01.Fields("CpName").VALUE)
        oDS_PS_CO120L.setValue "U_CCCode", i, Trim(oRecordSet01.Fields("CCCode").VALUE)
        oDS_PS_CO120L.setValue "U_CCName", i, Trim(oRecordSet01.Fields("CCName").VALUE)
        oDS_PS_CO120L.setValue "U_ProdQty", i, Trim(oRecordSet01.Fields("ProdQty").VALUE)
        oDS_PS_CO120L.setValue "U_DefQty", i, Trim(oRecordSet01.Fields("DefQty").VALUE)
        oDS_PS_CO120L.setValue "U_Cost", i, Trim(oRecordSet01.Fields("Cost").VALUE)
        oDS_PS_CO120L.setValue "U_Scrap", i, Trim(oRecordSet01.Fields("Scrap").VALUE)
        oDS_PS_CO120L.setValue "U_Loss", i, Trim(oRecordSet01.Fields("Loss").VALUE)

        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "�� ��ȸ��...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub