VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM015.cls
'//  Module         : MM
'//  Description    : ���ձ��� ����ǰ���ۼ�
'//  FormType       : PS_MM015
'//  Create Date    : 2011.02.21
'//  Modified Date  :
'//  Creator        : Noh Geun Yong
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Public oMat02               As SAPbouiCOM.Matrix
Public oGrid01              As SAPbouiCOM.Grid
Private oDS_PS_MM015H       As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_MM015L       As SAPbouiCOM.DBDataSource    '������


Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&
Private VIDocNum$, VODocNum$

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM015.srf")
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
    
    oFormUniqueID01 = "PS_MM015_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
'    Call FormClear
'    Call Matrix_AddRow(1, 0, True) 'oMat01
'    FormItemEnabled

    'oForm01.Items("MovDocNo").Click ct_Regular
    
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1287"), False        '// ����
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
    
    Set oDS_PS_MM015H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_MM015L = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    '// ��Ʈ���� ��ü �Ҵ�
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    
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
    
    oCombo.Select "0", psk_Index
    
    
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
    Dim vReturnValue$
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                
                If pval.ItemUID = "Search" Then '��ȸ��ư
                    Call Search_Data
                
                ElseIf pval.ItemUID = "Save" Then '���������ư
                    vReturnValue = Sbo_Application.MessageBox("�� �����͸� �߰��� �Ŀ��� ������ �� �����ϴ�. ����ϰڽ��ϱ�?", 1, "&Ȯ��", "&���")
                    If vReturnValue = 1 Then
                        If Save_Data() = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
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
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                
            Case et_CLICK: '//6
                          
                If pval.ItemUID = "Mat01" Then
                    
                    Call Search_Matrix_Data(pval.Row)
                End If
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//���
                    
                End If
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
                Call Form_Resize(FormUID, pval, BubbleEvent)
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
    Dim vReturnValue$
    
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

Private Sub FormItemEnabled()
    If oForm01.Mode = fm_FIND_MODE Then
        
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oMat As Integer, ByVal oRow As Long, Optional Insert_YN As Boolean)
On Error GoTo Matrix_AddRow_Error
        
    Select Case oMat
        
    End Select
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Matrix_AddRow_Error:
    MDC_Com.MDC_GF_Message "Matrix_AddRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub Form_Resize(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Form_Resize_Error
    Dim i   As Integer
    
    

    
    oForm01.Items("Mat01").Top = (oForm01.Height / 2) - 15
    oForm01.Items("Mat01").Height = (oForm01.Height / 2) - 38
    oForm01.Items("Mat01").Left = 6
    oForm01.Items("Mat01").Width = oForm01.Width - 21
    
    oMat01.AutoResizeColumns
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Form_Resize_Error:
    Sbo_Application.SetStatusBarMessage "Form_Resize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Search_Data()
    On Error GoTo Search_Data_Error
    Dim oRecordSet01          As SAPbobsCOM.Recordset
    Dim i                   As Integer
    Dim sQry                As String
    Dim BPLId               As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Freeze True
    
    oMat01.Clear
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    
    sQry = "EXEC PS_MM015_01 '" & BPLId & "'"
        
   oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_MM015H.Clear
        
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
        oForm01.Freeze False
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_MM015H.Size Then
            oDS_PS_MM015H.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_MM015H.Offset = i
        oDS_PS_MM015H.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("BEDAT").VALUE)
        oDS_PS_MM015H.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("EBELN").VALUE)
        oDS_PS_MM015H.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("Lifnr").VALUE)
        oDS_PS_MM015H.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("CardCode").VALUE)
        oDS_PS_MM015H.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("CardName").VALUE)
        oDS_PS_MM015H.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields("Cnt").VALUE)
        
        oRecordSet01.MoveNext
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    
    oMat02.Clear
    oForm01.Items("S_MENGE").Specific.VALUE = ""
    oForm01.Items("S_NETWR").Specific.VALUE = ""
            
    oForm01.Freeze False
    

    Set oRecordSet01 = Nothing
    Exit Sub
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Data_Error:
        Set oRecordSet01 = Nothing
        MDC_Com.MDC_GF_Message "Search_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub Search_Matrix_Data(ClickRow As Long)
    On Error GoTo Search_Matrix_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim cnt                 As Integer
    Dim S_MENGE               As Double 'ǰ���߷�
    Dim S_NETWR              As Double 'ǰ�Ǳݾ�
    
    
    Dim EBELN               As String
    
    EBELN = Trim(oMat01.Columns("EBELN").Cells(ClickRow).Specific.VALUE)
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    
    '/ Procedure ����
    sQry = "EXEC PS_MM015_02 '" & EBELN & "'"
    oRecordSet.DoQuery sQry

    
     'Matrix�� ���===========================================================================================
     '// Line �ʱ�ȭ
     cnt = oDS_PS_MM015L.Size
     If cnt > 0 Then
         For j = 0 To cnt - 1
             oDS_PS_MM015L.RemoveRecord oDS_PS_MM015L.Size - 1
         Next j
         If cnt = 1 Then
             oDS_PS_MM015L.Clear
         End If
     End If
     oMat02.LoadFromDataSource
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         
    '//Matrix�� Data �ѷ��ش�
    j = 1
     Do Until oRecordSet.EOF
             
         If oDS_PS_MM015L.Size < j Then
             oDS_PS_MM015L.InsertRecord j - 1 '�����߰�
         End If
         '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         oDS_PS_MM015L.setValue "U_ColReg01", j - 1, oRecordSet.Fields("BPLId").VALUE
         oDS_PS_MM015L.setValue "U_ColReg02", j - 1, oRecordSet.Fields("DocNum").VALUE
         oDS_PS_MM015L.setValue "U_ColReg03", j - 1, oRecordSet.Fields("LineNum").VALUE
         oDS_PS_MM015L.setValue "U_ColReg04", j - 1, oRecordSet.Fields("CardCode").VALUE
         oDS_PS_MM015L.setValue "U_ColReg05", j - 1, oRecordSet.Fields("Purchase").VALUE
         oDS_PS_MM015L.setValue "U_ColReg06", j - 1, oRecordSet.Fields("PQType").VALUE
         oDS_PS_MM015L.setValue "U_ColReg07", j - 1, oRecordSet.Fields("itemCode").VALUE
         oDS_PS_MM015L.setValue "U_ColQTy01", j - 1, oRecordSet.Fields("PQTy").VALUE
         oDS_PS_MM015L.setValue "U_ColQTy02", j - 1, oRecordSet.Fields("Weight").VALUE
         oDS_PS_MM015L.setValue "U_ColQTy03", j - 1, oRecordSet.Fields("MENGE").VALUE 'ǰ�Ǽ���
         oDS_PS_MM015L.setValue "U_ColQTy04", j - 1, oRecordSet.Fields("NETWR").VALUE 'ǰ�Ǳݾ�
         oDS_PS_MM015L.setValue "U_ColReg08", j - 1, oRecordSet.Fields("CntcCode").VALUE
         oDS_PS_MM015L.setValue "U_ColReg09", j - 1, oRecordSet.Fields("BEDAT").VALUE
         oDS_PS_MM015L.setValue "U_ColReg10", j - 1, oRecordSet.Fields("EINDT").VALUE
         
         S_MENGE = S_MENGE + oRecordSet.Fields("MENGE").VALUE
         
         S_NETWR = S_NETWR + oRecordSet.Fields("NETWR").VALUE

         j = j + 1
         oRecordSet.MoveNext
     Loop
     oMat02.LoadFromDataSource
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    oForm01.Items("S_MENGE").Specific.VALUE = S_MENGE
    oForm01.Items("S_NETWR").Specific.VALUE = S_NETWR
    
    
    oForm01.Freeze False
    Set oRecordSet = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Matrix_Data_Error:
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "Search_Matrix_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function Save_Data() As Boolean
    On Error GoTo Save_Data_Error
    Dim DI_oPurchaseOrders As SAPbobsCOM.Documents
    Dim oRecordSet        As SAPbobsCOM.Recordset
    Dim RetVal            As Long
    Dim sQry                As String
    Dim i                   As Integer
    Dim j                   As Integer
    Dim ErrNum              As Integer
    Dim PkWt                As Double
    Dim NPkWt               As Double
    Dim OPkWt               As Double
    Dim PODocEntry$, DocNum$, LineNum$, CardCode$, BPLId$, DocDate$, CntcCode$, DueDate$
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat02.FlushToDataSource
    
    If oMat02.VisualRowCount > 0 Then
    
        
        i = 0
    
        oDS_PS_MM015L.Offset = i
        
        
        If Sbo_Company.InTransaction = True Then Sbo_Company.EndTransaction wf_RollBack
        Sbo_Company.StartTransaction

        Set DI_oPurchaseOrders = Nothing
        Set DI_oPurchaseOrders = Sbo_Company.GetBusinessObject(oPurchaseOrders)
        
        DocNum = Trim(oDS_PS_MM015L.GetValue("U_ColReg02", i))
        LineNum = Trim(oDS_PS_MM015L.GetValue("U_ColReg03", i))
        CardCode = Trim(oDS_PS_MM015L.GetValue("U_ColReg04", i))
        BPLId = Trim(oDS_PS_MM015L.GetValue("U_ColReg01", i))
'        DocDate = Format(Now, "YYYY-MM-DD")
        
        
'         oDS_PS_MM015L.setValue "U_ColReg01", j - 1, oRecordset.Fields("BPLId").VALUE
'         oDS_PS_MM015L.setValue "U_ColReg02", j - 1, oRecordset.Fields("DocNum").VALUE
'         oDS_PS_MM015L.setValue "U_ColReg03", j - 1, oRecordset.Fields("LineNum").VALUE
'         oDS_PS_MM015L.setValue "U_ColReg04", j - 1, oRecordset.Fields("CardCode").VALUE
'         oDS_PS_MM015L.setValue "U_ColReg05", j - 1, oRecordset.Fields("Purchase").VALUE
'         oDS_PS_MM015L.setValue "U_ColReg06", j - 1, oRecordset.Fields("PQType").VALUE
'         oDS_PS_MM015L.setValue "U_ColReg07", j - 1, oRecordset.Fields("itemCode").VALUE
'         oDS_PS_MM015L.setValue "U_ColQTy01", j - 1, oRecordset.Fields("PQTy").VALUE
'         oDS_PS_MM015L.setValue "U_ColQTy02", j - 1, oRecordset.Fields("Weight").VALUE
'         oDS_PS_MM015L.setValue "U_ColQTy03", j - 1, oRecordset.Fields("MENGE").VALUE 'ǰ�Ǽ���
'         oDS_PS_MM015L.setValue "U_ColQTy04", j - 1, oRecordset.Fields("NETWR").VALUE 'ǰ�Ǳݾ�
'         oDS_PS_MM015L.setValue "U_ColReg08", j - 1, oRecordset.Fields("CntcCode").VALUE
'         oDS_PS_MM015L.setValue "U_ColReg09", j - 1, oRecordset.Fields("BEDAT").VALUE
'         oDS_PS_MM015L.setValue "U_ColReg10", j - 1, oRecordset.Fields("EINDT").VALUE
'
        
        
        DocDate = Format(Trim(oDS_PS_MM015L.GetValue("U_ColReg09", i)), "YYYY-MM-DD")
        DueDate = Format(Trim(oDS_PS_MM015L.GetValue("U_ColReg10", i)), "YYYY-MM-DD")
        CntcCode = Trim(oDS_PS_MM015L.GetValue("U_ColReg08", i))
        With DI_oPurchaseOrders
            .CardCode = CardCode
            .BPL_IDAssignedToInvoice = BPLId
            .DocDate = DocDate
            .DocDueDate = DueDate
            
            sQry = "Select empID From OHEM Where U_MSTCOD = '" & CntcCode & "'"
            oRecordSet.DoQuery sQry
            .DocumentsOwner = Trim(oRecordSet.Fields("empID").VALUE)
            
            .UserFields("U_reType").VALUE = Trim(oDS_PS_MM015L.GetValue("U_ColReg06", i)) 'Trim(oRecordSet01.Fields("U_PQType").VALUE)
            .UserFields("U_okYN").VALUE = "N"
            
            .UserFields("U_OrdTyp").VALUE = Trim(oDS_PS_MM015L.GetValue("U_ColReg05", i)) 'Trim(oRecordSet01.Fields("U_Purchase").VALUE)
            
            .Lines.SetCurrentLine i
            .Lines.ItemCode = Trim(oDS_PS_MM015L.GetValue("U_ColReg07", i))
            .Lines.Quantity = Trim(oDS_PS_MM015L.GetValue("U_ColQTy03", i)) 'Trim(oRecordSet01.Fields("U_E_MENGE").VALUE)
            .Lines.LineTotal = Trim(oDS_PS_MM015L.GetValue("U_ColQTy04", i)) 'Trim(oRecordSet01.Fields("U_E_NETWR").VALUE)
            .Lines.WarehouseCode = "10" + BPLId
            .Lines.UserFields("U_Qty").VALUE = Trim(oDS_PS_MM015L.GetValue("U_ColQTy03", i)) 'Trim(oRecordSet01.Fields("U_E_MENGE").VALUE)
            .Lines.UserFields("U_MM010Doc").VALUE = DocNum
            .Lines.UserFields("U_MM010Lin").VALUE = LineNum
            .Lines.UserFields("U_Auto").VALUE = "N"
                        
            If oMat02.VisualRowCount > 1 Then
                For i = 1 To oMat02.VisualRowCount - 1
                    DocNum = Trim(oDS_PS_MM015L.GetValue("U_ColReg02", i))
                    LineNum = Trim(oDS_PS_MM015L.GetValue("U_ColReg03", i))
                    
                    If i > 0 Then .Lines.Add
                    
                    .Lines.ItemCode = Trim(oDS_PS_MM015L.GetValue("U_ColReg07", i))
                    .Lines.Quantity = Trim(oDS_PS_MM015L.GetValue("U_ColQTy03", i))
                    .Lines.LineTotal = Trim(oDS_PS_MM015L.GetValue("U_ColQTy04", i))
                    .Lines.WarehouseCode = "10" + BPLId
                    .Lines.UserFields("U_Qty").VALUE = Trim(oDS_PS_MM015L.GetValue("U_ColQTy03", i))
                    .Lines.UserFields("U_MM010Doc").VALUE = DocNum
                    .Lines.UserFields("U_MM010Lin").VALUE = LineNum
                    '.Lines.UserFields("U_Auto").Value =
                    
                Next
            End If
        End With
                
            
        RetVal = DI_oPurchaseOrders.Add
        If (RetVal <> 0) Then
            If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
        Else
            Sbo_Company.GetNewObjectCode PODocEntry
            sQry = "EXEC [PS_INTERFACE_01] '" & PODocEntry & "'"
            oRecordSet.DoQuery sQry
        
            If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_Commit
            
            Set DI_oPurchaseOrders = Nothing
            
            oMat02.Clear
            oForm01.Items("S_MENGE").Specific.VALUE = ""
            oForm01.Items("S_NETWR").Specific.VALUE = ""
            
            
            sQry = "Update [@PS_MM010L] Set  U_GuBun = '3' "
            sQry = sQry & "From  [@PS_MM010L] a Inner Join [@PS_Mm010H] b On a.DocEntry = b.DocEntry "
            sQry = sQry & "Where a.U_GuBun = '2' And  b.U_PQType = '20' And b.CanCeled = 'N' And Isnull(a.U_POYesNo, 'N') = 'Y'"

    '//�ش��κп� ���¸� 3���� �ٲٸ� �� ������ ��

            oRecordSet.DoQuery sQry
            
            Call Search_Data
        End If
    End If
        
    

    Set DI_oPurchaseOrders = Nothing
    
    Set oRecordSet = Nothing
    Save_Data = True
    
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Save_Data_Error:
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "Save_Data_Error:" & Err.Number & " - " & Err.Description, "E"
    Save_Data = False
End Function