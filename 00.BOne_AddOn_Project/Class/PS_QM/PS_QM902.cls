VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM902"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM902.cls
'//  Module         : FI
'//  Description    : ���к�� ����ܰ����
'//  FormType       : PS_QM902
'//  Create Date    : 2011.05.12
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM902H       As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_QM902L       As SAPbouiCOM.DBDataSource    '��϶���

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM902.srf")
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
    
    oFormUniqueID01 = "PS_QM902_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm01.DataBrowser.BrowseBy = "DocEntry"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call FormClear
'    Call Matrix_AddRow(1, 0, True) 'oMat01
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_PS_QM902H = oForm01.DataSources.DBDataSources("@PS_QM902H")
    Set oDS_PS_QM902L = oForm01.DataSources.DBDataSources("@PS_QM902L")
    
    '// ��Ʈ���� ��ü �Ҵ�
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oDS_PS_QM902H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")

    
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
    
    '// ������ ��з�
'    Set oCombo = oForm01.Items("ItmBsort").Specific
'    sQry = "select Code, Name from [@PSH_ITMBSORT] where Code in ('101','102','104') order by Code"
'    oRecordSet01.DoQuery sQry
'    Call oCombo.ValidValues.Add("000", "����")
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
'    Call oCombo.Select(0, psk_Index)
    
     '// �����
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    oMat01.Columns("Staff").ValidValues.Add "", "����"
    oMat01.Columns("Staff").ValidValues.Add "1", "1��"
    oMat01.Columns("Staff").ValidValues.Add "2", "2��"
    
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

                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//���
                    If pval.ItemUID = "U_ItmBsort" Then
                        If oForm01.Items("U_ItmBsort").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    
                    If pval.ItemUID = "ItmMsort" Then
                        If oForm01.Items("ItmMsort").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    '//����
'                    If pval.ItemUID = "Mat01" Then
'                        If pval.ColUID = "PP070No" Then
'                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.VALUE = "" Then
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
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//���
'                    If pval.ItemUID = "ItmMsort" Then
'                        oMat01.Clear
'                    End If
                    If pval.ItemUID = "U_ItmBsort" Then
                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                    End If
                    If pval.ItemUID = "ItmMsort" Then
                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                    End If
                    
                    '//����
'                    If pval.ItemUID = "Mat01" And (pval.ColUID = "DocDate" Or pval.ColUID = "LoanAmt" Or pval.ColUID = "RepayAmt") Then
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
                    oDS_PS_QM902L.RemoveRecord oDS_PS_QM902L.Size - 1   '// Mat1�� ����������(�����) ����
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
            
            Case "1281": 'ã��
                FormItemEnabled
                oForm01.Items("DocEntry").Click ct_Regular
                
            Case "1282": '�߰�
                Call FormItemEnabled
                Call FormClear
                oDS_PS_QM902H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
'                Call Matrix_AddRow(1, 0, True) 'oMat01
                
                '//-- Combo Box �ʱ�ȭ
                Dim oCombo          As SAPbouiCOM.ComboBox
                '// �����
                Set oCombo = oForm01.Items("BPLId").Specific
                oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
                
                Set oCombo = Nothing

'                oForm01.Items("DocDate").Click ct_Regular

            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("DocDate").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
'                        Matrix_AddRow 1, oMat01.RowCount, False
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
    If oForm01.Mode = fm_FIND_MODE Then
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("U_ItmBsort").Enabled = True
        oForm01.Items("ItmMsort").Enabled = True
        
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("U_ItmBsort").Enabled = True
        oForm01.Items("ItmMsort").Enabled = True
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("U_ItmBsort").Enabled = False
        oForm01.Items("ItmMsort").Enabled = False

    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM902'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
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
                oDS_PS_QM902L.InsertRecord (oRow)
            End If
            '����
            oDS_PS_QM902L.Offset = oRow
            oDS_PS_QM902L.setValue "LineId", oRow, oRow + 1
            oDS_PS_QM902L.setValue "U_ItemCode", oRow, ""
            oDS_PS_QM902L.setValue "U_ItemName", oRow, ""
            oDS_PS_QM902L.setValue "U_CpCode", oRow, ""
            oDS_PS_QM902L.setValue "U_CpName", oRow, ""
            oDS_PS_QM902L.setValue "U_Price", oRow, ""
            oMat01.LoadFromDataSource

        
    End Select
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Matrix_AddRow_Error:
    MDC_Com.MDC_GF_Message "Matrix_AddRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, j&, Cnt&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim Balance         As Currency
    Dim RBalance        As Currency
    
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "U_ItmBsort"
            '//�ߺз��� �˻�
            oForm01.Freeze True
            
            sQry = "select Name from [@PSH_ITMBSORT] where Code = '" & Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_QM902H.setValue "U_ItmBname", 0, Trim(oRecordSet.Fields(0).VALUE)
            oMat01.LoadFromDataSource
            
            oForm01.Freeze False
        Case "ItmMsort"
            '//�ߺз��� �˻�
            sQry = "select U_CodeName from [@PSH_ITMMSORT] where U_Code = '" & Trim(oDS_PS_QM902H.GetValue("U_ItmMsort", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_QM902H.setValue "U_ItmMname", 0, Trim(oRecordSet.Fields(0).VALUE)
            
            
            '//��з�, �ߺз��� �ش��ϴ� ������ ���
             'Matrix�� ���===========================================================================================
             oForm01.Freeze True
             '// Line �ʱ�ȭ
             Cnt = oDS_PS_QM902L.Size
             If Cnt > 0 Then
                 For j = 0 To Cnt - 1
                     oDS_PS_QM902L.RemoveRecord oDS_PS_QM902L.Size - 1
                 Next j
                 If Cnt = 1 Then
                     oDS_PS_QM902L.Clear
                 End If
             End If
             oMat01.LoadFromDataSource
             '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) = "101" Or Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) = "102" Then '101:����(��ǰ), 102:��ǰ(��ǰ)�� ��쿡��
                sQry = "select ItemCode, ItemName from [OITM] where U_ItmBsort = '" & Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) & "' "
                sQry = sQry & "and U_ItmMsort = '" & Trim(oDS_PS_QM902H.GetValue("U_ItmMsort", 0)) & "' "
                sQry = sQry & "order by ItemCode"
                oRecordSet.DoQuery sQry
                     
                '//Matrix�� Data �ѷ��ش�
                 j = 1
                 Do Until oRecordSet.EOF
                         
                     If oDS_PS_QM902L.Size < j Then
                         oDS_PS_QM902L.InsertRecord j - 1 '�����߰�
                     End If
                     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                     oDS_PS_QM902L.setValue "U_LineNum", j - 1, j
                     oDS_PS_QM902L.setValue "LineId", j - 1, j
                     oDS_PS_QM902L.setValue "U_ItemCode", j - 1, oRecordSet.Fields("ItemCode").VALUE
                     oDS_PS_QM902L.setValue "U_ItemName", j - 1, oRecordSet.Fields("ItemName").VALUE
'                     oDS_PS_QM902L.setValue "U_CpCode", j - 1, oRecordSet.Fields("CpCode").VALUE
'                     oDS_PS_QM902L.setValue "U_CpName", j - 1, oRecordSet.Fields("CpName").VALUE
                     
                     j = j + 1
                     oRecordSet.MoveNext
                 Loop
            ElseIf Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) = "104" Then '104:��Ƽ������(��ǰ)
                sQry = "select U_CpCode,U_CpName from [@PS_PP001L] where Code = 'CP501' order by U_CpCode"
                oRecordSet.DoQuery sQry
                 
                '//Matrix�� Data �ѷ��ش�
                 j = 1
                 Do Until oRecordSet.EOF
                         
                     If oDS_PS_QM902L.Size < j Then
                         oDS_PS_QM902L.InsertRecord j - 1 '�����߰�
                     End If
                     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                     oDS_PS_QM902L.setValue "U_LineNum", j - 1, j
                     oDS_PS_QM902L.setValue "LineId", j - 1, j
                     oDS_PS_QM902L.setValue "U_CpCode", j - 1, oRecordSet.Fields("U_CpCode").VALUE
                     oDS_PS_QM902L.setValue "U_CpName", j - 1, oRecordSet.Fields("U_CpName").VALUE
                     
                     j = j + 1
                     oRecordSet.MoveNext
                 Loop
            ElseIf Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) = "101" Or Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) = "102" Or Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) = "111" Or Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) = "601" Then '�ݼӺи�
                sQry = "select U_CpCode,U_CpName from [@PS_PP001L] where Code = 'CP801' order by U_CpCode"
                oRecordSet.DoQuery sQry
                 
                '//Matrix�� Data �ѷ��ش�
                 j = 1
                 Do Until oRecordSet.EOF
                         
                     If oDS_PS_QM902L.Size < j Then
                         oDS_PS_QM902L.InsertRecord j - 1 '�����߰�
                     End If
                     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                     oDS_PS_QM902L.setValue "U_LineNum", j - 1, j
                     oDS_PS_QM902L.setValue "LineId", j - 1, j
                     oDS_PS_QM902L.setValue "U_CpCode", j - 1, oRecordSet.Fields("U_CpCode").VALUE
                     oDS_PS_QM902L.setValue "U_CpName", j - 1, oRecordSet.Fields("U_CpName").VALUE
                     
                     j = j + 1
                     oRecordSet.MoveNext
                Loop
            End If
            
            oMat01.LoadFromDataSource
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            oForm01.Freeze False
            
            
            
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
'            Case "DocDate"
'                oMat01.FlushToDataSource
'                oDS_PS_QM902L.Offset = oRow - 1
'                'oMat01.SetLineData oRow
'
'                '--------------------------------------------------------------------------------------------
'                If oRow = oMat01.RowCount And Trim(oDS_PS_QM902L.GetValue("U_DocDate", oRow - 1)) <> "" Then
'                '// ���� ���� �߰�
'                    Call Matrix_AddRow(1, 0, False)
'                    oMat01.Columns("DocDate").Cells(oRow).Click ct_Regular
'                End If
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
        Case Trim(oDS_PS_QM902H.GetValue("U_DocDate", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM902H.GetValue("U_ItmBsort", 0)) = "000"
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM902H.GetValue("U_ItmMsort", 0)) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM902H.GetValue("U_ItmMname", 0)) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������ڴ� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "��з��� �ʼ������Դϴ�. �����Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "�ߺз��� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "�ߺз����� �����ϴ�. �ߺз��ڵ带 Ȯ���Ͽ� �ֽʽÿ�.", "E"
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
    If oMat01.VisualRowCount = 0 Then '// MAT01�� ���� �ִ��� Ȯ�� (ErrorNumber : 1)
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//������ �� �ϳ��� ���� i=0���� �����ϹǷ� �ϳ��� ���Ƿ�
    '//oMat01.RowCount - 2�� �ȴ�..�ݵ�� ��� ���� �ϴ� �ʼ����� Ȯ���Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'    If oMat01.VisualRowCount > 0 Then
'        For i = 0 To oMat01.VisualRowCount - 2      '// Mat1�� �Է°��� �ùٸ��� ������ Ȯ�� (ErrorNumber : 2)
'            oDS_PS_QM902L.Offset = i
'            If Trim(oDS_PS_QM902L.GetValue("U_DocDate", i)) = "" Then
'                ErrNum = 2
'                oMat01.Columns("DocDate").Cells(i + 1).Click ct_Regular
'                GoTo MatrixSpaceLineDel_Error
'            End If
'        Next i
'    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//�Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '//�̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'    If oMat01.VisualRowCount > 0 Then
'        oDS_PS_QM902L.RemoveRecord oDS_PS_QM902L.Size - 1   '// Mat1�� ����������(�����) ����
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
        MDC_Com.MDC_GF_Message "����(����)�� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function