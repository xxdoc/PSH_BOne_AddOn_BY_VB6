VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_HR300.cls
'//  Module         : HR
'//  Description    : �濵��ħ���
'//  FormType       : PS_HR300
'//  Create Date    : 2012.2.11
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************


Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_HR300H As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_HR300L As SAPbouiCOM.DBDataSource    '��϶���

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&
Private oSeq&
Private oCount&

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR300.srf")
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
    
    oFormUniqueID01 = "PS_HR300_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormClear
    Call Add_MatrixRow(0, True)
        
    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1287"), False        '// ����
    oForm01.EnableMenu ("1284"), True         '// ���
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
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Dim SumWeight As Currency, SumQty As Long
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$, BPLId$
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                    
                        '���� ���س⵵, ���� ������� ������ �̹� �ԷµǾ����� üũ(fm_ADD_MODE�ϰ�츸)
                        If oForm01.Mode = fm_ADD_MODE Then
                            If CheckDocument = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                            oMat01.FlushToDataSource

                            oMat01.LoadFromDataSource
                        End If

                        Call Delete_EmptyRow
                        oLast_Mode = oForm01.Mode
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        oLast_Mode = oForm01.Mode
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    
                End If
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "BPLId" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_HR300L.Clear
                    If oForm01.Mode = fm_ADD_MODE Then
                        Call Add_MatrixRow(0, False)
                    End If
                                        
                    oForm01.Freeze False
                End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then

        
                    If (pval.ItemUID = "Mat01") Then
                        If pval.ColUID = "MngPlc" Then
                            oMat01.FlushToDataSource
                            
                            Call oDS_PS_HR300L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            If oMat01.RowCount = pval.Row And Trim(oDS_PS_HR300L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                                Add_MatrixRow (pval.Row)
                            End If
                            oMat01.LoadFromDataSource
                        End If
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                    Else
                        
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
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_OK_MODE Then
                        If oLast_Mode = fm_UPDATE_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            oLast_Mode = 100
                        ElseIf oLast_Mode = fm_FIND_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            FormItemEnabled
                            oLast_Mode = 100
                        End If
                    ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    
                End If
            Case et_MATRIX_LOAD: '//11
'                Add_MatrixRow oMat01.RowCount, False
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
                Set oDS_PS_HR300H = Nothing
                Set oDS_PS_HR300L = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    oForm01.Freeze False
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
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
'[1284:���] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1284": '���
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
            Case "1286": '�ݱ�
'[1293:�����] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '�����
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_HR300L.RemoveRecord oDS_PS_HR300L.Size - 1       '// Mat01�� ����������(�����) ����
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
'[1281:ã��] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": 'ã��
                oForm01.Freeze True
                Call FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                
                '//���̵� ����� ����
                oForm01.Items("BPLId").Specific.Select MDC_PS_Common.User_BPLId, psk_ByValue
                oForm01.Freeze False
'[1282:�߰�] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '�߰�
                Call Initialization
                Call FormItemEnabled
                Call FormClear
                oDS_PS_HR300H.setValue "U_StdYear", 0, Format(Now, "YYYY")
                Call Add_MatrixRow(0, True)
'                oForm01.Items("BPLId").Click ct_Collapsed
'[1288~1291:�׺���̼�] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                Call FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("MngPlc").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_HR300H.GetValue("Status", 0) = "O" Then
                            Add_MatrixRow oMat01.RowCount, False
                        End If
                    End If
                End If
        End Select
    End If
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
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
    Set oDS_PS_HR300H = oForm01.DataSources.DBDataSources("@PS_HR300H")
    Set oDS_PS_HR300L = oForm01.DataSources.DBDataSources("@PS_HR300L")
    
    '// ��Ʈ���� ��ü �Ҵ�
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oDS_PS_HR300H.setValue "U_StdYear", 0, Format(Now, "YYYY")
    
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
    
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
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
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("StdYear").Enabled = True
                
        oForm01.Items("Mat01").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("StdYear").Enabled = True
        
        oForm01.Items("Mat01").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("StdYear").Enabled = True
        oForm01.Items("Mat01").Enabled = True
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_HR300'", "")
    If DocNum = 0 Then
        oForm01.Items("DocNum").Specific.VALUE = 1
    Else
        oForm01.Items("DocNum").Specific.VALUE = DocNum
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//���߰�����
        oDS_PS_HR300L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_HR300L.Offset = oRow
    oDS_PS_HR300L.setValue "U_LineNum", oRow, oRow + 1
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
    Dim SumWeight As Currency, SumQty As Long
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
'        Case "ItmBsort"
'            sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" & Trim(oDS_PS_HR300H.GetValue("U_ItmBsort", 0)) & "'"
'            oRecordSet01.DoQuery sQry
'
'            oDS_PS_HR300H.setValue "U_ItmBname", 0, Trim(oRecordSet01.Fields(0).VALUE)
'        Case "Mat01"
'            Select Case oCol
'                Case "MachCode"
'                    oForm01.Freeze True
'                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("MachCode").Cells(oRow).Specific.VALUE) <> "" Then
'                        oMat01.FlushToDataSource
'                        Add_MatrixRow oMat01.RowCount, False
'                        oMat01.Columns("MachCode").Cells(oRow).Click ct_Regular
'                    End If
'
'                    sQry = "Select a.U_MachCode, a.U_MachName "
'                    sQry = sQry & "From [@PS_PP130H] a "
'                    sQry = sQry & "Where a.U_MachCode = '" & Trim(oMat01.Columns("MachCode").Cells(oRow).Specific.VALUE) & "'"
'                    oRecordSet01.DoQuery sQry
'
'                    If oRecordSet01.RecordCount = 0 Then
'                        MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
'                        Set oRecordSet01 = Nothing
'                        oForm01.Freeze False
'                        Exit Sub
'                    End If
'
'                    oMat01.Columns("MachName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_MachName").VALUE)
'                    oMat01.Columns("MachCode").Cells(oRow).Click ct_Regular
'
'                    oForm01.Freeze False
'            End Select
    End Select

'    Set oRecordset = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_HR300H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_HR300H.GetValue("U_StdYear", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "���س⵵ �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// ����
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 And oDS_PS_HR300L.GetValue("U_MngPlc", 0) = "" Then
        ErrNum = 2
        GoTo MatrixSpaceLineDel_Error
    End If
    
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case oDS_PS_HR300L.GetValue("U_MngPlc", i) = ""
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
        End Select
    Next
    oMat01.LoadFromDataSource

    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Or ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "���� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "�濵��ħ������ �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
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
        If Trim(oDS_PS_HR300L.GetValue("U_MngPlc", i)) = "" Then
            oDS_PS_HR300L.RemoveRecord i   '// Mat01�� ����������(�����) ����
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function CheckDocument() As Boolean
'******************************************************************************
'Function ID : CheckDocument()
'�ش��� : PS_HR300
'��� : �ش� ���س⵵�� �Էµ� �ڷᰡ �ִ��� �˻�
'�μ� : ����
'��ȯ�� : ����
'Ư�̻��� : ����
'******************************************************************************
On Error GoTo CheckDocument_Error
    
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Query01 = "                 SELECT       COUNT(*)"
    Query01 = Query01 & "  FROM        [@PS_HR300H] AS T0"
    Query01 = Query01 & "                  INNER JOIN"
    Query01 = Query01 & "                  [@PS_HR300H] AS T1"
    Query01 = Query01 & "                      ON T0.DocEntry = T1.DocEntry"
    Query01 = Query01 & "  WHERE       T0.[Status] = 'O'"
    Query01 = Query01 & "                  AND T0.U_BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "'"
    Query01 = Query01 & "                  AND T0.U_StdYear = '" & oForm01.Items("StdYear").Specific.VALUE & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    If RecordSet01.Fields(0).VALUE = 0 Then
        CheckDocument = True
    Else
        CheckDocument = False
        Call MDC_Com.MDC_GF_Message("�ش� ���س⵵�� �̹� [�濵��ħ]�� ��ϵǾ����ϴ�. Ȯ���Ͻʽÿ�.", "E")
    End If
    
    Set RecordSet01 = Nothing
    Exit Function
CheckDocument_Error:
    Set RecordSet01 = Nothing
    Call Sbo_Application.SetStatusBarMessage("CheckDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function
