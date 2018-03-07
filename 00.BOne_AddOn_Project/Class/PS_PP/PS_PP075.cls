VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP075"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP075.cls
'//  Module         : PP
'//  Description    : �����̵����
'//  FormType       : PS_PP075
'//  Create Date    : 2010.10.21
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_PP075H       As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_PP075L       As SAPbouiCOM.DBDataSource    '��϶���

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP075.srf")
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
    
    oFormUniqueID01 = "PS_PP075_" & GetTotalFormsCount
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
    Call FormClear
    Call Make_MovDocNo
    Call Matrix_AddRow(1, 0, True) 'oMat01
'    FormItemEnabled
    
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_PS_PP075H = oForm01.DataSources.DBDataSources("@PS_PP075H")
    Set oDS_PS_PP075L = oForm01.DataSources.DBDataSources("@PS_PP075L")
    
    '// ��Ʈ���� ��ü �Ҵ�
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oDS_PS_PP075H.setValue "U_RegiDate", 0, Format(Now, "YYYYMMDD")
    
    '�����
    oDS_PS_PP075H.setValue "U_CntcCode", 0, User_MSTCOD()
    FlushToItemValue "CntcCode"
    
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
    
    Call oCombo.Select(User_BPLId(), psk_ByValue)
    
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
                        
                        If UpDate_PP070("A") = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                
                ElseIf pval.ItemUID = "Print" Then '��¹�ư Ŭ����
                    Call Print_Query
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//���
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    '//����
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "PP070No" Then
                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
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
                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//���
                    If pval.ItemUID = "CntcCode" Or pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "RegiDate" Then
                        Call Make_MovDocNo
                    End If
                    '//����
                    If pval.ItemUID = "Mat01" And (pval.ColUID = "PP070No") Then
                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
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
                If UpDate_PP070("B") = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1286": '�ݱ�
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1285": '����
                If UpDate_PP070("A") = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1285": '����
            Case "1293": '�����
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                    '//�Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
                    '//�̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
                    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineId").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
    
                    oMat01.FlushToDataSource
                    oDS_PS_PP075L.RemoveRecord oDS_PS_PP075L.Size - 1   '// Mat1�� ����������(�����) ����
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
            
            Case "1281": 'ã��
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                
            Case "1282": '�߰�
                Call FormItemEnabled
                Call FormClear
                oDS_PS_PP075H.setValue "U_RegiDate", 0, Format(Now, "YYYYMMDD")
                Call Matrix_AddRow(1, 0, True) 'oMat01
                Call Make_MovDocNo
                
                Dim oCombo          As SAPbouiCOM.ComboBox
                Set oCombo = oForm01.Items("BPLId").Specific
                Call oCombo.Select(User_BPLId(), psk_Index)
                
                oDS_PS_PP075H.setValue "U_CntcCode", 0, User_MSTCOD()
                FlushToItemValue "CntcCode"

            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("PP070No").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_PP075H.GetValue("Status", 0) = "O" Then
                            Matrix_AddRow 1, oMat01.RowCount, False
                        End If
                    End If
                End If
  
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
        oForm01.Items("DocNum").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DeliArea").Enabled = True
        oForm01.Items("CarNo").Enabled = True
        oForm01.Items("TransCom").Enabled = True
        oForm01.Items("Fee").Enabled = True
        oForm01.Items("RegiDate").Enabled = True
        oMat01.Columns("PP070No").Editable = True
        
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DeliArea").Enabled = True
        oForm01.Items("CarNo").Enabled = True
        oForm01.Items("TransCom").Enabled = True
        oForm01.Items("Fee").Enabled = True
        oForm01.Items("RegiDate").Enabled = True
        oMat01.Columns("PP070No").Editable = True
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        If oForm01.Items("Canceled").Specific.VALUE = "Y" Then
            oForm01.Items("Comments").Click ct_Regular
            oForm01.Items("DocNum").Enabled = False
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("CntcCode").Enabled = False
            oForm01.Items("CardCode").Enabled = False
            oForm01.Items("DeliArea").Enabled = False
            oForm01.Items("CarNo").Enabled = False
            oForm01.Items("TransCom").Enabled = False
            oForm01.Items("Fee").Enabled = False
            oForm01.Items("RegiDate").Enabled = False
            oMat01.Columns("PP070No").Editable = False
        Else
            oForm01.Items("DocNum").Enabled = True
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("DeliArea").Enabled = True
            oForm01.Items("CarNo").Enabled = True
            oForm01.Items("TransCom").Enabled = True
            oForm01.Items("Fee").Enabled = True
            oForm01.Items("RegiDate").Enabled = True
            oMat01.Columns("PP070No").Editable = False
        End If
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP075'", "")
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
'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oMat As Integer, ByVal oRow As Long, Optional Insert_YN As Boolean)
On Error GoTo Matrix_AddRow_Error
        
    Select Case oMat
        Case 1: 'oMat01
            If Insert_YN = False Then
                oRow = oMat01.RowCount
                oDS_PS_PP075L.InsertRecord (oRow)
            End If
            '���Գ���
            oDS_PS_PP075L.Offset = oRow
            oDS_PS_PP075L.setValue "LineId", oRow, oRow + 1
            oDS_PS_PP075L.setValue "U_PP070No", oRow, ""
            oDS_PS_PP075L.setValue "U_ItemCode", oRow, ""
            oDS_PS_PP075L.setValue "U_ItemName", oRow, ""
            oDS_PS_PP075L.setValue "U_Size", oRow, ""
            oDS_PS_PP075L.setValue "U_Mark", oRow, ""
            oDS_PS_PP075L.setValue "U_Qty", oRow, ""
            oDS_PS_PP075L.setValue "U_Weight", oRow, ""
            oDS_PS_PP075L.setValue "U_DocDate", oRow, ""
            oMat01.LoadFromDataSource

'        Case 2: 'oMat02
'            If Insert_YN = False Then
'                oRow = oMat2.RowCount
'                oDS_ZPP140M.InsertRecord (oRow)
'            End If
'            '���⳻��
'            oDS_ZPP140M.Offset = oRow
'            oDS_ZPP140M.setValue "LineId", oRow, oRow + 1
'            oDS_ZPP140M.setValue "U_ConfDate", oRow, ""
'            oDS_ZPP140M.setValue "U_ConfNo", oRow, ""
'            oDS_ZPP140M.setValue "U_Size", oRow, ""
'            oDS_ZPP140M.setValue "U_ExpQty", oRow, ""
'            oDS_ZPP140M.setValue "U_RfndQty", oRow, ""
'            oMat02.LoadFromDataSource
        
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
        Case "CntcCode"
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_PP075H.GetValue("U_CntcCode", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_PP075H.setValue "U_CntcName", 0, Trim(oRecordSet.Fields(0).VALUE)
        Case "CardCode"
            sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" & Trim(oDS_PS_PP075H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_PP075H.setValue "U_CardName", 0, Trim(oRecordSet.Fields(0).VALUE)
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
            Case "PP070No"
                oMat01.FlushToDataSource
                oDS_PS_PP075L.Offset = oRow - 1
                'oMat01.SetLineData oRow
                
                i = InStr(oMat01.Columns("PP070No").Cells(oRow).Specific.String, "-")
                
                DocNum = Left(oMat01.Columns("PP070No").Cells(oRow).Specific.String, i - 1)
                LineId = Mid(oMat01.Columns("PP070No").Cells(oRow).Specific.String, i + 1)
                
                sQry = "select b.U_ItemCode, b.U_ItemName, c.U_Size, c.U_Mark, d.name, b.U_SelQty, b.U_SelWt "
                sQry = sQry & "from [@PS_PP070H] a inner join [@PS_PP070L] b on a.docentry=b.docentry "
                sQry = sQry & "left  join OITM c on b.U_ItemCode=c.ItemCode "
                sQry = sQry & "left  join [@PSH_MARK] d on c.U_Mark=d.Code "
                sQry = sQry & "Where a.DocNum = '" & DocNum & "'"
                sQry = sQry & "and b.LineId = '" & LineId & "'"
                oRecordSet.DoQuery sQry
'                If oRecordSet.RecordCount = 0 Then
'                    Sbo_Application.SetStatusBarMessage "�ڽ�Ʈ���͸� Ȯ���Ͽ� �ֽʽÿ�. ", bmt_Short, True
'                    oMat01.Columns("CCCode").Cells(oRow).Click ct_Regular
'                Else
                    oDS_PS_PP075L.setValue "U_ItemCode", oRow - 1, oRecordSet.Fields(0).VALUE
                    oDS_PS_PP075L.setValue "U_ItemName", oRow - 1, oRecordSet.Fields(1).VALUE
                    oDS_PS_PP075L.setValue "U_Size", oRow - 1, oRecordSet.Fields(2).VALUE
                    oDS_PS_PP075L.setValue "U_Mark", oRow - 1, oRecordSet.Fields(4).VALUE
                    oDS_PS_PP075L.setValue "U_Qty", oRow - 1, oRecordSet.Fields(5).VALUE
                    oDS_PS_PP075L.setValue "U_Weight", oRow - 1, oRecordSet.Fields(6).VALUE
'                    oDS_PS_PP075L.setValue "U_DocDate", oRow - 1, oRecordSet.Fields(7).Value
                    oMat01.SetLineData oRow
'                End If
                
                '--------------------------------------------------------------------------------------------
                If oRow = oMat01.RowCount And Trim(oDS_PS_PP075L.GetValue("U_PP070No", oRow - 1)) <> "" Then
                '// ���� ���� �߰�
                    Call Matrix_AddRow(1, 0, False)
                    oMat01.Columns("PP070No").Cells(oRow).Click ct_Regular
                End If

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
        Case Trim(oDS_PS_PP075H.GetValue("U_BPLId", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP075H.GetValue("U_CntcCode", 0)) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP075H.GetValue("U_CntcName", 0)) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_PP075H.GetValue("U_CardCode", 0)) = ""
'            ErrNum = 4
'            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_PP075H.GetValue("U_CardName", 0)) = ""
'            ErrNum = 5
'            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_PP075H.GetValue("U_CarNo", 0)) = ""
'            ErrNum = 6
'            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_PP075H.GetValue("U_Fee", 0)) = "" Or Trim(oDS_PS_PP075H.GetValue("U_Fee", 0)) = 0
'            ErrNum = 7
'            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP075H.GetValue("U_RegiDate", 0)) = ""
            ErrNum = 8
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP075H.GetValue("U_MovDocNo", 0)) = ""
            ErrNum = 9
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "����ڴ� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "�����ڸ��� �����ϴ�. ������ڵ带 Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "��ǰó�� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "��ǰó���� �����ϴ�. ��ǰó�ڵ带 Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "������ȣ�� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "������ ���ų� ���� 0 �Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "������ڴ� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "�̵���Ϲ�ȣ�� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
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
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//������ �� �ϳ��� ���� i=0���� �����ϹǷ� �ϳ��� ���Ƿ�
    '//oMat01.RowCount - 2�� �ȴ�..�ݵ�� ��� ���� �ϴ� �ʼ����� Ȯ���Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 2      '// Mat1�� �Է°��� �ùٸ��� ������ Ȯ�� (ErrorNumber : 2)
            oDS_PS_PP075L.Offset = i
            If Trim(oDS_PS_PP075L.GetValue("U_PP070No", i)) = "" Then
                ErrNum = 2
                oMat01.Columns("PP070No").Cells(i + 1).Click ct_Regular
                GoTo MatrixSpaceLineDel_Error
            End If
        Next i
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//�Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '//�̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    If oMat01.VisualRowCount > 0 Then
        oDS_PS_PP075L.RemoveRecord oDS_PS_PP075L.Size - 1   '// Mat1�� ����������(�����) ����
    End If
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

Private Sub Make_MovDocNo() '�̵���Ϲ�ȣ����
On Error GoTo Make_MovDocNo_Error
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    '/ Procedure ����
    sQry = "EXEC PS_PP075_01 '" & oDS_PS_PP075H.GetValue("U_RegiDate", 0) & "'"
    oRecordSet.DoQuery sQry

    oDS_PS_PP075H.setValue "U_MovDocNo", 0, Trim(oRecordSet.Fields(0).VALUE)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Make_MovDocNo_Error:
    MDC_Com.MDC_GF_Message "Make_MovDocNo_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function UpDate_PP070(ByVal A_B As String) As Boolean
On Error GoTo UpDate_PP070_Error
    Dim i&, j&, DocEntry$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
        
    Select Case A_B
        Case "A": '[@PS_PP070L]-MovDocNo �� �̵���Ϲ�ȣ update

            For i = 0 To oMat01.VisualRowCount - 1
                oDS_PS_PP075L.Offset = i
                    
                j = InStr(Trim(oDS_PS_PP075L.GetValue("U_PP070No", i)), "-")
                DocEntry = Left(Trim(oDS_PS_PP075L.GetValue("U_PP070No", i)), j - 1)
                LineId = Mid(Trim(oDS_PS_PP075L.GetValue("U_PP070No", i)), j + 1)
                
                sQry = "update [@PS_PP070L] set U_MovDocNo = '" & Trim(oDS_PS_PP075H.GetValue("U_MovDocNo", 0)) & "' "
                sQry = sQry & "Where docentry = '" & DocEntry & "' "
                sQry = sQry & "and LineId = '" & LineId & "'"
                oRecordSet.DoQuery sQry
            Next i
        
        Case "B": '[@PS_PP070L]-MovDocNo �� �̵���Ϲ�ȣ update ���
            
            For i = 0 To oMat01.VisualRowCount - 2
                oDS_PS_PP075L.Offset = i
                
                j = InStr(Trim(oDS_PS_PP075L.GetValue("U_PP070No", i)), "-")
                DocEntry = Left(Trim(oDS_PS_PP075L.GetValue("U_PP070No", i)), j - 1)
                LineId = Mid(Trim(oDS_PS_PP075L.GetValue("U_PP070No", i)), j + 1)
                
                '//��ҽ� ����������(PS_PP077)���� ���Ǿ����� ��ҺҰ�**********************************
                sQry = "select DocNum from [@PS_PP077H] "
                sQry = sQry & "where isnull(Canceled,'') <> 'Y' "
                sQry = sQry & "and U_MovDocNo = '" & Trim(oDS_PS_PP075H.GetValue("U_MovDocNo", 0)) & "' "
                sQry = sQry & "and U_PP070No = '" & DocEntry & "' "
                sQry = sQry & "and U_PP070NoL = '" & LineId & "'"
                oRecordSet.DoQuery sQry
                
                If oRecordSet.RecordCount <> 0 Then '���Ǿ�����
                    ErrNum = 1
                    GoTo UpDate_PP070_Error
                Else
                    sQry = "update [@PS_PP070L] set U_MovDocNo = Null "
                    sQry = sQry & "Where docentry = '" & DocEntry & "' "
                    sQry = sQry & "and LineId = '" & LineId & "'"
                    oRecordSet.DoQuery sQry
                End If
            Next i
        
    End Select
    UpDate_PP070 = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
UpDate_PP070_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "����ó�����[PS_PP077]���� ���Ǿ������ϴ�. ����Ҽ� �����ϴ�.", "E"
    Else
        MDC_Com.MDC_GF_Message "UpDate_PP070_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    UpDate_PP070 = False
End Function

Private Sub Print_Query()
On Error GoTo Print_Query_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim sQry01          As String
    Dim Sub_sQry        As String
    Dim DocNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    '// ��ȸ���ǹ�
    DocNum = Trim(oForm01.Items("DocNum").Specific.VALUE)

    sQry01 = "SELECT COUNT(*) AS COUNT FROM [@PS_PP075H] H, [@PS_PP075L] L  WHERE H.DocEntry = L.DocEntry and H.DocNum = '" & DocNum & "'"
    Call oRecordSet01.DoQuery(sQry01)
    If Trim(oRecordSet01.Fields(0).VALUE) > 10 Then
        '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WinTitle = "[PS_PP075] �������/������"
        ReportName = "PS_PP075_05.RPT"
        
        '// Formula �����ʵ�
        ReDim gRpt_Formula(1)
        ReDim gRpt_Formula_Value(1)
        
        '// SubReport
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        
        ReDim gRpt_SFormula(1, 1)
        ReDim gRpt_SFormula_Value(1, 1)
        
        gRpt_SFormula(1, 1) = ""
        gRpt_SFormula_Value(1, 1) = ""
        
        '/ Procedure ����
        sQry = "EXEC [PS_PP075_05] '" & DocNum & "'"
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount = 0 Then
            ErrNum = 1
            GoTo Print_Query_Error
        End If
        
    Else
        
        '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WinTitle = "[PS_PP075] �������/������"
        ReportName = "PS_PP075_01.RPT"
        
        '// Formula �����ʵ�
        ReDim gRpt_Formula(1)
        ReDim gRpt_Formula_Value(1)
        
        '// SubReport
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        
        ReDim gRpt_SFormula(1, 1)
        ReDim gRpt_SFormula_Value(1, 1)
        
        gRpt_SFormula(1, 1) = ""
        gRpt_SFormula_Value(1, 1) = ""
        
        '/ Procedure ����
        sQry = "EXEC [PS_PP075_02] '" & DocNum & "'"
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount = 0 Then
            ErrNum = 1
            GoTo Print_Query_Error
        End If
        
    End If

    '/ Action (sub_query�� �������� 'Y'��...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

