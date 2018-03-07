VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM153"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM153.cls
'//  Module         : MM
'//  Description    : �����԰� ǰ���԰� Ȯ��
'//  FormType       : PS_MM153
'//  Create Date    : 2013.10.14
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oGrid1           As SAPbouiCOM.Grid

Public oDS_PS_MM153H     As SAPbouiCOM.DataTable

'Private oDS_PS_MM153H As SAPbouiCOM.DBDataSource    '������

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&
Private oForm01_Mode&
Private oLast_RightClick_CgNum&

Private oMenuUID$

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm(Optional MenuUID$)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM153.srf")
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
    
    oFormUniqueID01 = "PS_MM153_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oMenuUID = MenuUID
    
    oForm01.Mode = fm_OK_MODE
'    oForm01_Mode = fm_FIND_MODE
    oForm01.EnableMenu ("1281"), False        '// ã��
    oForm01.EnableMenu ("1282"), False        '// �߰�
    oForm01.EnableMenu ("1293"), False         '// �����

    oMenuUID = MenuUID
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting(oMenuUID)
    Call Initialization
    Call LoadCaption
    Call FormItemEnabled(oMenuUID)
    
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1287"), False        '// ����
    oForm01.EnableMenu ("1285"), False        '// ����
    oForm01.EnableMenu ("1284"), False         '// ���
            
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
    Dim SumQty As Long, SumWeight As Currency
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim RFC_Sender$
    Dim Seq&
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    Dim ChildForm02 As Variant
    Set ChildForm02 = New PS_SM030
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If oForm01_Mode = fm_UPDATE_MODE Then
                        
                        If Update_PurchaseDemand(pval) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        oForm01_Mode = fm_OK_MODE
                        Call LoadCaption
                    ElseIf oForm01_Mode = fm_OK_MODE Then
                        oForm01.Close
                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    Call LoadData
                    
                    oForm01_Mode = fm_OK_MODE
                    Call LoadCaption
                ElseIf pval.ItemUID = "Btn03" Then
                    If oForm01.DataSources.DataTables.Item(0).Rows.Count > 0 Then
                        oForm01.Freeze True
                        For i = 0 To oForm01.DataSources.DataTables.Item(0).Rows.Count - 1
                            If oDS_PS_MM153H.Columns.Item("û������").Cells(i).VALUE = "N" Then
                                If oGrid1.DataTable.GetValue("�˻翩��", i) = "Y" Then
                                    oGrid1.DataTable.Columns("�˻翩��").Cells.Item(i).VALUE = "N"
                                    oDS_PS_MM153H.Columns.Item("�˻���").Cells(i).VALUE = ""
                                Else
                                    oGrid1.DataTable.Columns("�˻翩��").Cells.Item(i).VALUE = "Y"
                                    oDS_PS_MM153H.Columns.Item("�˻���").Cells(i).VALUE = Format(Now, "YYYYMMDD")
                                End If
                            End If
                        Next
                        oForm01.Freeze False
                    End If
                    oForm01_Mode = fm_UPDATE_MODE
                    Call LoadCaption
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ItemCode" Then
                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
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
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "BPLId" Then
                    oForm01.Freeze True
                    oDS_PS_MM153H.Clear
                    
                    oForm01.Freeze False
                ElseIf pval.ItemUID = "Grid01" Then
                    
                    
                    If oDS_PS_MM153H.Columns.Item("û������").Cells(pval.Row).VALUE = "Y" Then
                        oDS_PS_MM153H.Columns.Item("�˻翩��").Cells(pval.Row).VALUE = oDS_PS_MM153H.Columns.Item("OKYN").Cells(pval.Row).VALUE
                        oDS_PS_MM153H.Columns.Item("�˻���").Cells(pval.Row).VALUE = oDS_PS_MM153H.Columns.Item("OKDate").Cells(pval.Row).VALUE
                    Else
                        If oDS_PS_MM153H.Columns.Item("�˻翩��").Cells(pval.Row).VALUE = "Y" Then
                            oDS_PS_MM153H.Columns.Item("�˻���").Cells(pval.Row).VALUE = Format(Now, "YYYYMMDD")
                        Else
                            oDS_PS_MM153H.Columns.Item("�˻���").Cells(pval.Row).VALUE = ""
                        End If
                    End If
                    oForm01_Mode = fm_UPDATE_MODE
                    Call LoadCaption
                End If
            Case et_CLICK: '//6
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7

            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Grid01" Then
                        
                        If oDS_PS_MM153H.Columns.Item("û������").Cells(pval.Row).VALUE = "Y" Then
                            oDS_PS_MM153H.Columns.Item("�˻翩��").Cells(pval.Row).VALUE = oDS_PS_MM153H.Columns.Item("OKYN").Cells(pval.Row).VALUE
                            oDS_PS_MM153H.Columns.Item("�˻���").Cells(pval.Row).VALUE = oDS_PS_MM153H.Columns.Item("OKDate").Cells(pval.Row).VALUE
                        Else
                            
                        End If
                        
                        oForm01_Mode = fm_UPDATE_MODE
                        Call LoadCaption
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
                
                Set oDS_PS_MM153H = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    oForm01.Freeze False
    Set ProgressBar01 = Nothing
    Set oRecordSet01 = Nothing
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
    Dim ReturnValue&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
'[1293:�����] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                
'[1281:ã��] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": 'ã��
                
'[1282:�߰�] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '�߰�
                
'[1288~1291:�׺���̼�] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
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
        If eventInfo.Row > 0 Then
            
        End If
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
    'Set oDS_PS_MM153H = oForm01.DataSources.DBDataSources("@PS_MM153H")
    
    Set oGrid1 = oForm01.Items("Grid01").Specific
    
    oForm01.DataSources.DataTables.Add ("PS_MM153")
    
    oGrid1.DataTable = oForm01.DataSources.DataTables.Item("PS_MM153")
    Set oDS_PS_MM153H = oForm01.DataSources.DataTables.Item("PS_MM153")
    
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting(Optional MenuUID$)
On Error GoTo ComboBox_Setting_Error
    '//�޺��� �⺻������
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 8)
    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
    
    oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Date 'Format(Date - 1, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 8)
    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
    
    oForm01.DataSources.UserDataSources.Item("DocDateTo").VALUE = Date

    '// �����
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    
    '//���翩��
    Set oCombo = oForm01.Items("OKYN").Specific
    oCombo.ValidValues.Add "Y", "����"
    oCombo.ValidValues.Add "N", "�̰���"
    oCombo.Select 1, psk_Index

    
    
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
    
    '//���̵� ��� ����
    oForm01.Items("CardCode").Specific.VALUE = ""
    
    
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

Sub FormItemEnabled(Optional MenuUID$)
On Error GoTo FormItemEnabled_Error
'    If MenuUID = "PS_MM153_1" Then '����
'        If oForm01_Mode = fm_FIND_MODE Then
'            oForm01.Items("CntcCode").Click ct_Regular
'            oForm01.Items("OrdType").Enabled = True
'            oForm01.Items("BPLId").Enabled = True
'            oForm01.Items("CntcCode").Enabled = True
'            oForm01.Items("DeptCode").Enabled = True
'            oForm01.Items("DocDateFr").Enabled = True
'            oForm01.Items("DocDateTo").Enabled = True
'            oForm01.Items("CgNumFr").Enabled = True
'            oForm01.Items("CgNumTo").Enabled = True
'            oForm01.Items("ItemCode").Enabled = True
'            oForm01.Items("ItmBSort").Enabled = True
'            oForm01.Items("ItmMSort").Enabled = True
''            oForm01.Items("ItemType").Enabled = True
'            oForm01.Items("OKYN").Enabled = True
'            oForm01.Items("Btn02").Enabled = True
'
'
'        End If
'    Else
'
'    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim ItemCode$, Qty&, SumQty As Long, SumWeight As Currency
    Dim Calculate_Weight As Currency
    Dim Calculate_Qty As Currency
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        
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
'    Select Case True
'        Case Trim(oForm01.Items("CardCode").Specific.VALUE) = ""
'            ErrNum = 1
'            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
'            ErrNum = 2
'            GoTo HeaderSpaceLineDel_Error
'    End Select
'
'    If oForm01_Mode = fm_ADD_MODE Then
'        If Trim(oForm01.Items("CardCode").Specific.VALUE) = "" Then
'            ErrNum = 3
'            GoTo HeaderSpaceLineDel_Error
'        ElseIf Trim(oForm01.Items("BPLId").Specific.VALUE) = "" Then
'            ErrNum = 4
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�ŷ�ó�� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "������� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "û������ �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "û���μ��� �ʼ������Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Public Function Update_PurchaseDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Update_PurchaseDemand_Error
    Dim i As Integer, j&, K&
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim OkYN$, OkDate$, DocNo$, LineNum$
    
        
    If oForm01.DataSources.DataTables.Item(0).Rows.Count > 0 Then
        For i = 0 To oForm01.DataSources.DataTables.Item(0).Rows.Count - 1
            If oDS_PS_MM153H.Columns.Item("û������").Cells(i).VALUE = "N" Then
                OkYN = oDS_PS_MM153H.Columns.Item("�˻翩��").Cells(i).VALUE
                OkDate = oDS_PS_MM153H.Columns.Item("�˻���").Cells(i).VALUE
                DocNo = oDS_PS_MM153H.Columns.Item("�԰���ȣ").Cells(i).VALUE
                LineNum = oDS_PS_MM153H.Columns.Item("�԰�����").Cells(i).VALUE
                
                         
                sQry = "UPDATE [@PS_MM152L] "
                sQry = sQry & "SET "
                sQry = sQry & "U_QCOKYN = '" & OkYN & "', "
                If OkDate = "" Then
                    sQry = sQry & "U_QCOKDate = NULL "
                Else
                    sQry = sQry & "U_QCOKDate = '" & OkDate & "' "
                End If
                sQry = sQry & " Where DocEntry = '" & DocNo & "' and U_LineNum = '" & LineNum & "' "
        
                RecordSet01.DoQuery sQry
            End If
            
        Next i
        'Sbo_Application.SetStatusBarMessage "���ſ�û������ ����Ǿ����ϴ�.", bmt_Short, False
        MDC_Com.MDC_GF_Message "�����԰� ǰ���˻� �Ϸ�!", "S"
        
        oForm01.Items("Btn02").Click ct_Regular
        
    Else
        MDC_Com.MDC_GF_Message "�����Ͱ� �������� �ʽ��ϴ�.!", "S"
        'Sbo_Application.SetStatusBarMessage "�����Ͱ� �������� �ʽ��ϴ�.", bmt_Short, True
    End If
    
    Set RecordSet01 = Nothing
    Update_PurchaseDemand = True
    oForm01_Mode = fm_OK_MODE
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_PurchaseDemand_Error:
    Update_PurchaseDemand = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Update_PurchaseDemand_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String, SumQty As Currency, SumWeight As Currency
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID$, CardCode$, OkYN$
    Dim Calculate_Qty As Long
    Dim DocDateFr$, DocDateTo$
    
    Dim iRow As Long
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    
    OkYN = Trim(oForm01.Items("OKYN").Specific.VALUE)
    
    If BPLID = "" Then BPLID = "%"
    If CardCode = "" Then CardCode = "%"
    If DocDateFr = "" Then DocDateFr = Format(Date - 90, "YYYY-MM") & "-01"
    'If DocDateTo = "" Then DocDateTo = "20991231"
    If DocDateTo = "" Then DocDateTo = Format(Now, "YYYY-MM-DD")
    
    If OkYN = "" Or OkYN = "ALL" Then OkYN = "%"
    
    sQry = "EXEC [PS_MM153_01] '" & BPLID & "','" & CardCode & "','" & DocDateFr & "','" & DocDateTo & "','" & OkYN & "'"
    
    Call oDS_PS_MM153H.ExecuteQuery(sQry)
    
    iRow = oForm01.DataSources.DataTables.Item(0).Rows.Count
    
    Call TitleSetting(iRow)
    
    'oRecordSet01.DoQuery sQry
    
'    If (oRecordSet01.RecordCount = 0) Then
'        MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
'        Set oRecordSet01 = Nothing
'        Exit Sub
'    End If
    
    oForm01.Freeze True
    oForm01.Freeze False
  
    'Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    oForm01.Freeze False
    'Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub TitleSetting(iRow As Long)
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
    Dim ColumnCnt          As Long
     
    Dim COLNAM(8)       As String
    Dim BPLID          As String
    
    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
        
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo error_Message
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Freeze True
    
    ColumnCnt = oDS_PS_MM153H.Columns.Item("ColumnCnt").Cells(0).VALUE
'    COLNAM(0) = "������ȣ-��"
'    COLNAM(1) = "ǰ���ڵ�"
'    COLNAM(2) = "ǰ���"
'    COLNAM(3) = "�����"
'    COLNAM(4) = "����"
'    COLNAM(5) = "����/�߷�"
'    COLNAM(6) = "û����ȣ"
'    COLNAM(7) = "û������"
'    COLNAM(8) = "��������"
'    COLNAM(9) = "û���ڸ�"
    
    BPLID = oForm01.Items("BPLId").Specific.VALUE
    
    
    
    For i = 0 To ColumnCnt  'UBound(COLNAM)
        'oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        
        Select Case oGrid1.Columns.Item(i).TitleObject.Caption 'COLNAM(i)
            Case "�˻���"
                oGrid1.Columns.Item(i).Editable = True
                oGrid1.Columns.Item(i).RightJustified = True
            Case "�˻翩��"
                oGrid1.Columns.Item(i).Editable = True
                oGrid1.Columns.Item(i).Type = gct_ComboBox
                Set oComboCol = oGrid1.Columns.Item("�˻翩��")
                
                oComboCol.ValidValues.Add "Y", "�˻�"
                oComboCol.ValidValues.Add "N", "�̰˻�"
        
                oComboCol.DisplayType = cdt_Description
            Case Else
                oGrid1.Columns.Item(i).Editable = False
        End Select
        
        
        
    Next i
    
    oGrid1.AutoResizeColumns

    oForm01.Freeze False

    Set oColumn = Nothing

    Exit Sub

error_Message:
    oForm01.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub


Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01_Mode = fm_OK_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "Ȯ��"
    ElseIf oForm01_Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "����"
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

