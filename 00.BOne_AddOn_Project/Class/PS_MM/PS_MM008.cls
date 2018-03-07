VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM008.cls
'//  Module         : MM
'//  Description    : ��ǰ���� ���� ���
'//  FormType       : PS_MM008
'//  Create Date    : 2013.12.14
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01             As SAPbouiCOM.Matrix
Private oDS_PS_MM008H As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_MM008L As SAPbouiCOM.DBDataSource    '��϶���

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&

Private oYYYYMM$

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM008.srf")
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
    
    oFormUniqueID01 = "PS_MM008_" & GetTotalFormsCount
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
    Call Initialization
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim ItemCode        As String
    Dim CpCode        As String
    Dim rate As Currency
    Dim StdWgt As Currency
                            
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
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
                        
                        Dim Code$, BPLID$, YM$, ItmBSort$
                        BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
                        YM = Trim(Right(oForm01.Items("YM").Specific.VALUE, 4))
                        ItmBSort = Trim(oForm01.Items("ItmBsort").Specific.VALUE)
                        
                        Code = YM & BPLID & ItmBSort
                        Call oDS_PS_MM008H.setValue("Code", 0, Code)
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "ItmBsort" Then
                        If oForm01.Items("ItmBsort").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "ItemCode" Then
                            If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ColUID = "CpCode" Then
                            If oMat01.Columns("CpCode").Cells(pval.Row).Specific.VALUE = "" Then
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
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "ItmBsort" Then
                        oForm01.Freeze True
                        
                        sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" & Trim(oForm01.Items("ItmBsort").Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        oForm01.Items("ItmBname").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        oForm01.Freeze False
                    ElseIf pval.ItemUID = "Mat01" Then
                        oMat01.FlushToDataSource
                        If pval.ColUID = "BfWgt" Or pval.ColUID = "AfWgt" Then
                            oForm01.Freeze True
                            If Val(oMat01.Columns("AfWgt").Cells(pval.Row).Specific.VALUE) <> 0 Or Val(oMat01.Columns("BfWgt").Cells(pval.Row).Specific.VALUE) <> 0 Then
                                Call oDS_PS_MM008L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                Call oDS_PS_MM008L.setValue("U_Scrap", pval.Row - 1, Val(oMat01.Columns("BfWgt").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns("AfWgt").Cells(pval.Row).Specific.VALUE))  '��ũ����
                            Else
                                Call oDS_PS_MM008L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                            End If
                            oMat01.LoadFromDataSource
                            oForm01.Freeze False
                        
                            oMat01.Columns(pval.ColUID).Cells(pval.Row).Click
                           
                            
                        ElseIf pval.ColUID = "ItemCode" Then
                            If (pval.Row = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) <> "" Then
                                oMat01.FlushToDataSource
                                Call Add_MatrixRow(oMat01.RowCount, False)
                                oMat01.Columns("ItemCode").Cells(pval.Row).Click ct_Regular
                            End If
                            
                            ItemCode = Trim(oDS_PS_MM008L.GetValue("U_ItemCode", pval.Row - 1))
                            sQry = "Select a.ItemCode, a.ItemName "
                            sQry = sQry & " From OITM a "
                            sQry = sQry & " Where a.ItemCode = '" & ItemCode & "'"
                            oRecordSet01.DoQuery sQry
                            
                                
                            oForm01.Freeze True
                            oMat01.FlushToDataSource
                            
                            If oRecordSet01.RecordCount = 0 Then
                            
                                '��Ʈ������ �����͸� ���� ���ε��ϸ� �̺�Ʈ�� ����Ǳ� ������ DataSource�� ���ε��ϴ� ������� ����(2011.11.22 �۸���)
                                Call oDS_PS_MM008L.setValue("U_ItemCode", pval.Row - 1, "")
                                Call oDS_PS_MM008L.setValue("U_ItemName", pval.Row - 1, "")
                                Call oDS_PS_MM008L.setValue("U_Cpcode", pval.Row - 1, "")
                                Call oDS_PS_MM008L.setValue("U_CpName", pval.Row - 1, "")
                                Call oDS_PS_MM008L.setValue("U_BfWgt", pval.Row - 1, 0)
                                Call oDS_PS_MM008L.setValue("U_AfWgt", pval.Row - 1, 0)
                                Call oDS_PS_MM008L.setValue("U_Scrap", pval.Row - 1, 0)
                                Call oDS_PS_MM008L.setValue("U_Rate", pval.Row - 1, 0)
                                
                                MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
                                
                                Set oRecordSet01 = Nothing
                                oMat01.LoadFromDataSource
                                oForm01.Freeze False
                                Exit Sub
                            
                                
                            End If
                            
                            
                            '��Ʈ������ �����͸� ���� ���ε��ϸ� �̺�Ʈ�� ����Ǳ� ������ DataSource�� ���ε��ϴ� ������� ����(2011.11.22 �۸���)
                            Call oDS_PS_MM008L.setValue("U_ItemCode", pval.Row - 1, Trim(oRecordSet01.Fields("ItemCode").VALUE)) 'ǰ���ڵ�
                            Call oDS_PS_MM008L.setValue("U_ItemName", pval.Row - 1, Trim(oRecordSet01.Fields("ItemName").VALUE)) 'ǰ���̸�
                            
                            oMat01.LoadFromDataSource
                            oForm01.Freeze False
                            
                            oMat01.Columns(pval.ColUID).Cells(pval.Row).Click
                        ElseIf pval.ColUID = "CpCode" Then
                            CpCode = Trim(oDS_PS_MM008L.GetValue("U_CpCode", pval.Row - 1))
                            sQry = "Select U_CpName"
                            sQry = sQry & " From [@PS_PP001L] "
                            sQry = sQry & " Where U_CpCode = '" & CpCode & "'"
                            oRecordSet01.DoQuery sQry
                            oForm01.Freeze True
                            oMat01.FlushToDataSource
                        
                            If oRecordSet01.RecordCount = 0 Then
                                Call oDS_PS_MM008L.setValue("U_CpName", pval.Row - 1, "")  '������
                            Else
                                Call oDS_PS_MM008L.setValue("U_CpName", pval.Row - 1, Trim(oRecordSet01.Fields("U_CpName").VALUE)) '������
                            End If
                            
                            
                            oMat01.LoadFromDataSource
                            oForm01.Freeze False
                            
                            oMat01.Columns(pval.ColUID).Cells(pval.Row).Click
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
                Set oDS_PS_MM008H = Nothing
                Set oDS_PS_MM008L = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
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
'[1293:�����] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '�����
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_MM008L.RemoveRecord oDS_PS_MM008L.Size - 1       '// Mat01�� ����������(�����) ����
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
            Case "1281": 'ã��
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
'                oForm01.Items("CycleCod").Click ct_Regular
                oForm01.Freeze False
            Case "1282": '�߰�
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
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
                oDS_PS_MM008H.setValue "Code", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_MM008L.setValue "Code", i, ""
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
    Set oDS_PS_MM008H = oForm01.DataSources.DBDataSources("@PS_MM008H")
    Set oDS_PS_MM008L = oForm01.DataSources.DBDataSources("@PS_MM008L")
    
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
    
    oForm01.Items("YM").Specific.VALUE = Left(Format(Now, "YYYYMMDD"), 6)
    Call Add_MatrixRow(0, True)
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
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo Add_MatrixRow_Error
'    If RowIserted = False Then '//���߰�����
'        oDS_PS_MM008L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_MM008L.Offset = oRow
'    oDS_PS_MM008L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
'    Exit Sub
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Add_MatrixRow_Error:
'    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
'End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_MM008H.GetValue("U_ItmBsort", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM008H.GetValue("U_BPLId", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM008H.GetValue("U_YM", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Len(Trim(oForm01.Items("YM").Specific.VALUE)) <> 6
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "��з��� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "������� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "����� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "����� 6�ڸ� YYYYMM �������� �Է��ؾ��մϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function


Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//���߰�����
        oDS_PS_MM008L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM008L.Offset = oRow
    oDS_PS_MM008L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
