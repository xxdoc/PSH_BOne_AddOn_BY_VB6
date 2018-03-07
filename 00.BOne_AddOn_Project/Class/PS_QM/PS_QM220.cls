VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM220.cls
'//  Module         : QM
'//  Description    : M/G�˻缺����ȸ
'//  FormType       : PS_QM220
'//  Create Date    : 2012.05.08
'//  Modified Date  :
'//  Creator        : NGY
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_TEMPTABLE As SAPbouiCOM.DBDataSource

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM220.srf")
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
    
    oFormUniqueID01 = "PS_QM220_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
        
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    oForm01.Freeze True
    
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
                  
    FormItemEnabled
    
    oForm01.EnableMenu ("1281"), False        '// ã��
    oForm01.EnableMenu ("1282"), False        '// �߰�
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1287"), False        '// ����
    oForm01.EnableMenu ("1285"), False        '// ����
    oForm01.EnableMenu ("1284"), False        '// ���
    oForm01.EnableMenu ("1293"), False         '// �����
        
    oForm01.Update
    oForm01.Freeze False
    oForm01.Visible = True
    
    Set oXmlDoc01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadForm_Error:
    If oForm01 Is Nothing Then
    Else
        oForm01.Update
        oForm01.Freeze False
    End If
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
    Dim TempForm01 As Variant
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                       End If
                    ElseIf pval.ItemUID = "ItemCode" Then
                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If oMat01.Columns("PQDocNum").Cells(pval.Row).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
                
'et_COMBO_SELECT ////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE//////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn_ret" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call Display_MatrixData
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "Mat01" And pval.ColUID = "Admision" Then
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                End If
            Case et_CLICK: '//6
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row = "0" And pval.ColUID = "Admision" Then
                    Dim Admision$
                    oForm01.Freeze True
                    oMat01.FlushToDataSource
                    If Trim(oDS_PS_TEMPTABLE.GetValue("U_sField08", 0)) = "" Or Trim(oDS_PS_TEMPTABLE.GetValue("U_sField08", 0)) = "N" Then
                        Admision = "Y"
                    ElseIf Trim(oDS_PS_TEMPTABLE.GetValue("U_sField08", 0)) = "Y" Then
                        Admision = "N"
                    End If
                    For i = 0 To oMat01.VisualRowCount - 1
                        oDS_PS_TEMPTABLE.setValue "U_sField08", i, Admision
                    Next i
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
'et_MATRIX_LINK_PRESSED /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_MATRIX_LINK_PRESSED '//8
         
            Case et_VALIDATE: '//10
         
                
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
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    If oForm01 Is Nothing Then
    Else
        oForm01.Freeze False
    End If
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
             Case "7169" '���� ��������
                Call Add_MatrixRow(oMat01.VisualRowCount)
'            Case "1284": '���
'            Case "1286": '�ݱ�
'            Case "1293": '�����
'            Case "1281": 'ã��
'            Case "1282": '�߰�
'            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
             Case "7169" '���� ��������
                '���� �������� ���� ó��
                Call oForm01.Freeze(True)
                Call oDS_PS_TEMPTABLE.RemoveRecord(oDS_PS_TEMPTABLE.Size - 1)
                Call oMat01.LoadFromDataSource
                Call oForm01.Freeze(False)
'            Case "1284": '���
'            Case "1286": '�ݱ�
'            Case "1293": '�����
'            Case "1281": 'ã��
'            Case "1282": '�߰�
'            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_PS_TEMPTABLE = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// ��Ʈ���� ��ü �Ҵ�
    Set oMat01 = oForm01.Items("Mat01").Specific
        
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 8)
    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
    oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 8)
    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
    oForm01.DataSources.UserDataSources.Item("DocDateTo").VALUE = Date
    
    
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
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
   ' Dim DocNum As String
   ' DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM030'", "")
   ' If DocNum = 0 Then
   '     oForm01.Items("DocNum").Specific.VALUE = 1
   ' Else
   '     oForm01.Items("DocNum").Specific.VALUE = DocNum
   ' End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//���߰�����
        oDS_PS_TEMPTABLE.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_TEMPTABLE.Offset = oRow
    oDS_PS_TEMPTABLE.setValue "U_LineNum", oRow, oRow + 1
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
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01"

    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���Ű��������� ��ҵǾ��ų� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "E"
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
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
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
    End If
    
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function Display_MatrixData() As Boolean
On Error GoTo Display_MatrixData_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Dim BPLID$, DocDateFr$, DocDateTo$, ItemCode$, CardCode$
    Dim sCnt&
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    
    If ItemCode = "" Then ItemCode = "%"
    If CardCode = "" Then CardCode = "%"
    
    sQry = "EXEC [PS_QM220_01] '" & BPLID & "', '" & DocDateFr & "', '" & DocDateTo & "', '" & ItemCode & "', '" & CardCode & "'"
    oRecordSet01.DoQuery sQry
    
    oDS_PS_TEMPTABLE.Clear
    
    sCnt = 0
    If oRecordSet01.RecordCount > 0 Then
        Do Until oRecordSet01.EOF
            oDS_PS_TEMPTABLE.InsertRecord (sCnt)
            oDS_PS_TEMPTABLE.Offset = sCnt
            
            oDS_PS_TEMPTABLE.setValue "U_LineNum", sCnt, sCnt + 1
            oDS_PS_TEMPTABLE.setValue "U_ColReg01", sCnt, Trim(oRecordSet01.Fields(0).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg02", sCnt, Trim(oRecordSet01.Fields(1).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg03", sCnt, Trim(oRecordSet01.Fields(2).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg04", sCnt, Format(Trim(oRecordSet01.Fields(3).VALUE), "YYYY-MM-DD")
            oDS_PS_TEMPTABLE.setValue "U_ColReg05", sCnt, Trim(oRecordSet01.Fields(4).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg06", sCnt, Trim(oRecordSet01.Fields(5).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg07", sCnt, Trim(oRecordSet01.Fields(6).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg08", sCnt, Trim(oRecordSet01.Fields(7).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg09", sCnt, Trim(oRecordSet01.Fields(8).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg10", sCnt, Trim(oRecordSet01.Fields(9).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg11", sCnt, Trim(oRecordSet01.Fields(10).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg12", sCnt, Trim(oRecordSet01.Fields(11).VALUE)
            oDS_PS_TEMPTABLE.setValue "U_ColReg13", sCnt, Trim(oRecordSet01.Fields(12).VALUE)
            
            
            sCnt = sCnt + 1
            oRecordSet01.MoveNext
        Loop
        
    End If
    
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    Display_MatrixData = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Display_MatrixData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "��ȸ�� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "Display_MatrixData_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Display_MatrixData = False
End Function



