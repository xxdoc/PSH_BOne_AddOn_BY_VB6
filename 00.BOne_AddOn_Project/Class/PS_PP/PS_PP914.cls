VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP914"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP914.cls
'//  Module         : PP
'//  Description    : ������񺰰�������Ȳ
'//  FormType       : PS_PP914
'//  Create Date    : 2011.06.15
'//  Modified Date  :
'//  Creator        : NGY
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid

Public oDS_PS_PP914A As SAPbouiCOM.DataTable

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP914.srf")
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
    
    oFormUniqueID01 = "PS_PP914_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_OK_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1287"), False        '// ����
    oForm01.EnableMenu ("1284"), True         '// ���
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

Sub Initialization()
On Error GoTo Initialization_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//���̵� ����� ����
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//���̵� ��� ����
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error

    Dim optBtn As SAPbouiCOM.OptionBtn
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    oGrid01.SelectionMode = ms_Auto
    
    Call oForm01.DataSources.DataTables.Add("PS_PP914A")
    
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("PS_PP914A")
    
    Set oDS_PS_PP914A = oForm01.DataSources.DataTables.Item("PS_PP914A")
    
    '//������� �ҽ� ��ü �Ҵ�
'    Set oDS_PS_PP914H = oForm01.DataSources.DBDataSources("@PS_PP914H")
'    Set oDS_PS_PP914L = oForm01.DataSources.DBDataSources("@PS_PP914L")
    
    '// ��Ʈ���� ��ü �Ҵ�
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 10)
    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
    oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYYMMDD")

    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 10)
    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
    oForm01.DataSources.UserDataSources.Item("DocDateTo").VALUE = Format(Now, "YYYYMMDD")
    
    Call oForm01.DataSources.UserDataSources.Add("RadioBtn", dt_SHORT_TEXT, 1)
    
    Set optBtn = oForm01.Items("Rad01").Specific          '������� ������ ��Ȳ
    optBtn.ValOn = "A"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.Selected = True
        
    Set optBtn = oForm01.Items("Rad02").Specific          '������� �񰡵� ����
    optBtn.ValOn = "B"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.GroupWith ("Rad01")
    
    Set optBtn = oForm01.Items("Rad03").Specific          '������ �񰡵� ����1
    optBtn.ValOn = "C"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.GroupWith ("Rad02")
    
    Set optBtn = oForm01.Items("Rad04").Specific          '������ �񰡵� ����2
    optBtn.ValOn = "D"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn"
    optBtn.GroupWith ("Rad03")
    
     
    Set optBtn = Nothing
   
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
       
'    Call oForm01.Freeze(True)
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
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
                If pval.ItemUID = "BtnSearch" Then
'                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
'                        If HeaderSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
'                    End If
                
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call PS_PP914_MTX01
                    End If
                
                ElseIf pval.ItemUID = "BtnPrint" Then '��¹�ư Ŭ����
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call Print_Query
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10                                  '���� â ����� ��Ī �־��ֱ�
                    If pval.ItemUID = "FailCode" Then
                       FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
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
        
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
'                If pval.ItemUID = "1" Then
'                    If oForm01.Mode = fm_ADD_MODE Then
'                        oForm01.Mode = fm_OK_MODE
'                        Call Sbo_Application.ActivateMenuItem("1282")
'                    ElseIf oForm01.Mode = fm_OK_MODE Then
'                        FormItemEnabled
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
'                    End If
'                End If
            Case et_KEY_DOWN: '//2
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
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
'                Set oMat01 = Nothing
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

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
    
    Dim CardCode$, CardName$, ItemMsort$, MsortName$, Mark$, MarkName$, ItemType$, TypeName$
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
            
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
'        Case Trim(oDS_PS_PP914H.GetValue("U_BPLId", 0)) = ""
'            ErrNum = 1
'            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub PS_PP914_MTX01()
On Error GoTo PS_PP914_MTX01_Error

    '//��Ʈ������ ������ �ε�
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim Query02 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLId As String
    Dim Part As String
    Dim DocDateFr As String
    Dim DocDateTo As String
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    Part = Trim(oForm01.Items("Part").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", 100, False)
    
    If oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "A" Then
    
        Query01 = "EXEC [PS_PP914_51] '" & BPLId & "', '" & Part & "', '" & DocDateFr & "', '" & DocDateTo & "'"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "B" Then
    
        Query01 = "EXEC [PS_PP914_52] '" & BPLId & "', '" & Part & "', '" & DocDateFr & "', '" & DocDateTo & "'"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "C" Then
    
        Query01 = "EXEC [PS_PP914_53] '" & BPLId & "', '" & Part & "', '" & DocDateFr & "', '" & DocDateTo & "'"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "D" Then
    
        Query01 = "EXEC [PS_PP914_53] '" & BPLId & "', '" & Part & "', '" & DocDateFr & "', '" & DocDateTo & "'"
        
    End If
    
    Call oGrid01.DataTable.Clear
    Call oDS_PS_PP914A.ExecuteQuery(Query01)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")


    If oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "A" Then
    
        oGrid01.Columns(4).RightJustified = True
        oGrid01.Columns(5).RightJustified = True
        oGrid01.Columns(6).RightJustified = True
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "B" Then
    
        oGrid01.Columns(4).RightJustified = True
        oGrid01.Columns(7).RightJustified = True
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "C" Then
    
        oGrid01.Columns(4).RightJustified = True
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "D" Then
    
        oGrid01.Columns(4).RightJustified = True
        
    End If


'    oGrid01.Columns(14).RightJustified = True
'    oGrid01.Columns(15).RightJustified = True
'    oGrid01.Columns(16).RightJustified = True
'    oGrid01.Columns(17).RightJustified = True
'    oGrid01.Columns(18).RightJustified = True
'    oGrid01.Columns(19).RightJustified = True
'    oGrid01.Columns(20).RightJustified = True
'    oGrid01.Columns(21).RightJustified = True
'    oGrid01.Columns(22).RightJustified = True
'    oGrid01.Columns(23).RightJustified = True
'    oGrid01.Columns(24).RightJustified = True
'    oGrid01.Columns(25).RightJustified = True
'    oGrid01.Columns(26).RightJustified = True
'    oGrid01.Columns(27).RightJustified = True
'    oGrid01.Columns(29).RightJustified = True
'    oGrid01.Columns(30).RightJustified = True
'    oGrid01.Columns(31).RightJustified = True
'    oGrid01.Columns(32).RightJustified = True
'    oGrid01.Columns(33).RightJustified = True
'
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[���]��, ���
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[���]��, ���
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[�Ϸ�]��, ���
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) 'ǰ����, ���
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '���԰���, ���
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '����(ǰ��-���԰�), �ϴ�
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '�˼��԰���, ���
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '����(���԰�-ǰ��), �ϴ�
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '�Ѽҿ���, ����
        
    If oGrid01.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "E")
        GoTo PS_PP914_MTX01_Exit
    End If

    oGrid01.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP914_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP914_MTX01_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_PP914_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Print_Query()
On Error GoTo Print_Query_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim sQry01          As String
    Dim Sub_sQry        As String
    
    Dim BPLId           As String
    Dim DocDateFr       As String
    Dim DocDateTo       As String
    Dim Part            As String

    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", 100, False)

    Call ConnectODBC
    
    '// ��ȸ���ǹ�
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    Part = Trim(oForm01.Items("Part").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    
    
'    If oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "A" Then
'       WinTitle = "������񺰰�������Ȳ[PS_PP914_01]"
'       ReportName = "PS_PP914_01.RPT"
'    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "B" Then
'       WinTitle = "������񺰺񰡵�����[PS_PP914_02]"
'       ReportName = "PS_PP914_02.RPT"
'    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "C" Then
'       WinTitle = "�������񰡵�����[PS_PP914_03]"
'       ReportName = "PS_PP914_03.RPT"
'    End If
    
    '// Formula �����ʵ�
    ReDim gRpt_Formula(3)
    ReDim gRpt_Formula_Value(3)
    
    gRpt_Formula(1) = "DocDateFr"
    gRpt_Formula_Value(1) = Format(DocDateFr, "0000-00-00")
    gRpt_Formula(2) = "DocDateTo"
    gRpt_Formula_Value(2) = Format(DocDateTo, "0000-00-00")
    gRpt_Formula(3) = "BPLId"
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    sQry01 = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
    Call oRecordSet01.DoQuery(sQry01)
    gRpt_Formula_Value(3) = oRecordSet01.Fields(0).VALUE
    Set oRecordSet01 = Nothing
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure ����"
    If oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "A" Then
    
        WinTitle = "������񺰰�������Ȳ[PS_PP914_01]"
        ReportName = "PS_PP914_01.RPT"
    
        sQry = "EXEC [PS_PP914_01] '" & BPLId & "', '" & Part & "', '" & DocDateFr & "', '" & DocDateTo & "'"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "B" Then
    
        WinTitle = "������񺰺񰡵�����[PS_PP914_02]"
        ReportName = "PS_PP914_02.RPT"
    
        sQry = "EXEC [PS_PP914_02] '" & BPLId & "', '" & Part & "', '" & DocDateFr & "', '" & DocDateTo & "'"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "C" Then
    
        WinTitle = "�������񰡵�����[PS_PP914_03]"
        ReportName = "PS_PP914_03.RPT"
    
        sQry = "EXEC [PS_PP914_03] '" & BPLId & "', '" & Part & "', '" & DocDateFr & "', '" & DocDateTo & "'"
        
    ElseIf oForm01.DataSources.UserDataSources("RadioBtn").VALUE = "D" Then
    
        WinTitle = "�������񰡵�����[PS_PP914_03]"
        ReportName = "PS_PP914_04.RPT"
    
        sQry = "EXEC [PS_PP914_03] '" & BPLId & "', '" & Part & "', '" & DocDateFr & "', '" & DocDateTo & "'"
        
    End If

'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

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









