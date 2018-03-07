VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM170.cls
'//  Module         : MM
'//  Desc           : ��ü��ݵ��
'//  FormType       : PS_MM170
'//  Create Date    : 2010.08.31
'//  Modified Date  :
'//  Creator        : DONG SUB KIM
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oForm02              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_MM170H       As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_MM170L       As SAPbouiCOM.DBDataSource    '��϶���

Private oLast_Item_UID   As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID    As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row    As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oSeq&
'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM170.srf")
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
    
    oFormUniqueID01 = "PS_MM170_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
        
    oForm01.Freeze True
    
    Call oForm01.EnableMenu("1293", True)
    Call CreateItems
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////D////////////////////////////////////////////////
    
    Call ComboBox_Setting
    Call CF_ChooseFromList
    Call Initial_Setting
    Call FormItemEnabled       '//���� ������ Ȱ��ȭ ��Ʈ��
    Call FormClear             '//UDO����϶�
    Call AddMatrixRow(0, oMat01.RowCount, True) '//UDO����϶� - ��Ʈ���� Call Function
    
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1287"), True         '// ����
    oForm01.EnableMenu ("1284"), True         '// ���
    oForm01.EnableMenu ("1293"), True         '// �����
    
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//�޴�����
    oForm01.Update
    oForm01.Freeze False
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
Dim i As Integer
Dim sQry                As String
Dim oRecordSet01        As SAPbobsCOM.Recordset
Dim TempForm01 As Variant
Dim ErrNum As Integer

Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)   '// ��ü ���� �� ������ �Ҵ�
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
        
            Case et_ITEM_PRESSED: '//1
            If pval.ItemUID = "1" Then
                If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                   If HeaderSpaceLineDel = False Then
                        BubbleEvent = False                   ' BubbleEvent = True �̸�, ����ڿ��� ������� �Ѱ��ش�. BeforeAction = True�� ��츸 ����.
                        Exit Sub
                    End If
                    If MatrixSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            ElseIf pval.ItemUID = "Btn01" Then      '// ǰ�Ǽ� ���
            
            ElseIf pval.ItemUID = "Btn02" Then      '// DI API - �а� ����
               If oForm01.Mode = fm_OK_MODE Then
                  If oForm01.Items("JdtDate").Specific.VALUE = "" Then
                     ErrNum = 1
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  ElseIf oForm01.Items("Status").Specific.VALUE = "C" Then
                     ErrNum = 2
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  Else
                    If Create_oJournalEntries(1) = False Then
                      BubbleEvent = False
                      Exit Sub
                    End If
                  End If
               
               Else
                   Call MDC_Com.MDC_GF_Message("���� ������ �� �а� ó�� �ٶ��ϴ�.", "W")
                   BubbleEvent = False
                   Exit Sub
               End If
                       
            ElseIf pval.ItemUID = "Btn03" Then      '// DI API - �а� ���
               If oForm01.Mode = fm_OK_MODE Then
                  If oForm01.Items("JdtDate").Specific.VALUE = "" Then
                     ErrNum = 1
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  ElseIf oForm01.Items("JdtCC").Specific.VALUE <> "Y" Then
                     ErrNum = 3
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  ElseIf oForm01.Items("Status").Specific.VALUE = "C" Then
                     ErrNum = 2
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  Else
                     If Cancel_oJournalEntries(1) = False Then
                        BubbleEvent = False
                        Exit Sub
                     End If
                  End If
               Else
                     Call MDC_Com.MDC_GF_Message("���� ������ �� �а� ó�� �ٶ��ϴ�.", "W")
                     BubbleEvent = False
                     Exit Sub
               End If
               
            Else
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" And pval.ColUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    End If
                End If
            End If
   
               
            Case et_KEY_DOWN: '//2
                ' �ŷ�ó�ڵ�
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    If pval.ItemUID = "CardCode" And pval.CharPressed = 9 Then
                        oForm01.Items("CardCode").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' �԰���ȣ - ������ ȣ��
                If pval.ItemUID = "Mat01" And pval.ColUID = "GRDocNum" And pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Or _
                           oForm01.Items("BPLId").Specific.VALUE = "" Then
                           ErrNum = 4
                           Call Item_Error_Message(ErrNum)
                        Else
                           Set TempForm01 = New PS_MM171
                           Call TempForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                        End If
                           BubbleEvent = False
                    End If
                End If
                               
                ' �����
                If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                    If pval.ItemUID = "CntcCode" And pval.CharPressed = 9 Then
                        oForm01.Items("CntcCode").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' ǰ���ڵ�
                If pval.ItemUID = "Mat01" And pval.ColUID = "ItemCode" And pval.CharPressed = 9 Then
                   If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                       oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                       Sbo_Application.ActivateMenuItem ("7425")
                       BubbleEvent = False
                   End If
                End If
                          
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            If pval.ItemChanged = True Then
                If pval.ItemUID = "Mat01" And pval.ColUID = "GRDocNum" Then
                    FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                End If
            End If
        
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            
                ' ��� �̸� Query
                If pval.ItemUID = "CntcCode" Then
                    sQry = "Select U_FULLNAME, U_MSTCOD From [OHEM] Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
                    Call oRecordSet01.DoQuery(sQry)
                    oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                End If
                                                      
                
                                                      
                ' ������ �ڵ�
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" And pval.ColUID = "GRDocNum" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    End If
                End If
                
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
                oLast_Item_UID = pval.ItemUID
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
        
    ElseIf (pval.BeforeAction = False) Then         '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
            
             ' ���� �� �߰� ����ó��
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                          oForm01.Mode = fm_OK_MODE
                          Call Sbo_Application.ActivateMenuItem("1282")
                    ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = False Then
                        FormItemEnabled
                        AddMatrixRow 1, oMat01.RowCount, True
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.Action_Success = True Then
                    oSeq = 1
                End If
            Case et_COMBO_SELECT: '//5
            
            ' ��ü�ݾ� ��Ʈ���� SUM
            If pval.ItemUID = "Mat01" And pval.ColUID = "RepayYN" Then
                oForm01.Items("Sum").Specific.VALUE = ""
                For i = 1 To oMat01.VisualRowCount - 1
                    If oMat01.Columns("RepayYN").Cells(i).Specific.VALUE = "Y" Then
                        oForm01.Items("Sum").Specific.VALUE = Val(oForm01.Items("Sum").Specific.VALUE) + _
                        Val(oMat01.Columns("RepayP").Cells(i).Specific.VALUE)
                        '//���ڿ� ���� ���Ե� ���ڸ� ������ ������ ���ڰ����� ��ȯ�մϴ�.
                    End If
                Next i
            End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            ' ������ Query
            If pval.ItemUID = "CardCode" And pval.ItemChanged = True Then
                sQry = "Select CardName From [OCRD] Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                oDS_PS_MM170H.setValue "U_CardName", 0, Trim(oRecordSet01.Fields(0).VALUE)
            End If
            
            '//�������ڸ� �а����ڿ� �����ϰ�...
            If pval.ItemUID = "DocDate" Then
                oDS_PS_MM170H.setValue "U_JdtDate", 0, oForm01.Items("DocDate").Specific.VALUE
            End If
            
            Case et_MATRIX_LOAD: '//11
                AddMatrixRow 1, oMat01.VisualRowCount, True
            Case et_FORM_ACTIVATE: '//18
            If oSeq = 1 Then
                oSeq = 0
            End If
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
                
            Case et_GOT_FOCUS: '//3
                oLast_Item_UID = pval.ItemUID

            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
        End Select
      End If
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
        Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
  Dim i&
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '��ݱ�
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư

        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1281": 'ã��
                FormItemEnabled
'                oForm01.Items("ItemCode").Click ct_Regular
            Case "1282": '�߰�
                FormItemEnabled
                FormClear
                AddMatrixRow 0, oMat01.RowCount, True
                oForm01.Items("CardCode").Click ct_Collapsed
                
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("GRDocNum").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        AddMatrixRow 1, oMat01.RowCount, True
                    End If
                End If
            Case "1293": '��ݱ�
               
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD:  '//34 - �߰�
'            FormClear
'            If Create_oJournalEntries(2) = False Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            Case et_FORM_DATA_UPDATE: '//35 - ������Ʈ
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
Raise_FormDataEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//�۾�
    ElseIf (eventInfo.BeforeAction = False) Then
        '//�۾�
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function CreateItems() As Boolean
On Error GoTo CreateItems_Error
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
     
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM170H = oForm01.DataSources.DBDataSources("@PS_MM170H")
    Set oDS_PS_MM170L = oForm01.DataSources.DBDataSources("@PS_MM170L")
    Set oMat01 = oForm01.Items("Mat01").Specific                        '// ��Ʈ���� ������ ��
    
     oDS_PS_MM170H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
     oDS_PS_MM170H.setValue "U_JdtDate", 0, Format(Now, "yyyymmdd")
     
    '// ����� ����Ʈ
    sQry = "SELECT BPLId, BPLName FROM [OBPL]"
    
    oRecordSet01.DoQuery sQry
    
    Do Until oRecordSet01.EOF
        oForm01.Items("BPLId").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '// ��ü�������
    oMat01.Columns("RepayYN").ValidValues.Add "Y", "�ο�"
    oMat01.Columns("RepayYN").ValidValues.Add "N", "����"
    
    Set oRecordSet01 = Nothing
    Exit Function
    
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//�޺��� �⺻������
        Call oForm01.Freeze(True)
    
    '//�޺��� �⺻������
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "Mat01", "ItemCode", "01", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "Mat01", "ItemCode", "02", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_MM130", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "ItemCode", "", "01", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "ItemCode", "", "02", "����ǰ")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_MM130", "ItemCode")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL WHERE BPLId = '1' ORDER BY BPLId", "â��", False, False)

' ǰ���з�
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code")
' ǰ���ߺз�
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code")
' ����Ÿ��
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code")
' ����
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] ORDER BY Code")
' ������ȣ
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Mark"), "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code")
' ���Ա��ش���
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ObasUnit"), "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code")
    
    Call oForm01.Freeze(False)
    Exit Sub
    Exit Sub
ComboBox_Setting_Error:
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList ����
    Exit Sub
CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub Initial_Setting()
On Error GoTo Initial_Setting_Error
 ' �����
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 ' �μ���
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
    Exit Sub
Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("JdtDate").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True
        oForm01.Items("JdtDate").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("JdtDate").Enabled = False
    End If
    
    Exit Sub

FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Sub

Sub AddMatrixRow(ByVal oSeq As Integer, ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo AddMatrixRow_Error

   Select Case oSeq
        Case 0:
            oMat01.AddRow           ' ��Ʈ������ ���ο� �θ� �߰��Ѵ�.
            oDS_PS_MM170L.setValue "LineId", oRow, oRow + 1
            oMat01.LoadFromDataSource
        Case 1:
            oDS_PS_MM170L.InsertRecord oRow
            oDS_PS_MM170L.setValue "LIneId", oRow, oRow + 1
            oMat01.LoadFromDataSource
    End Select
'AddMatrixRow_Error:
'    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM170'", "")
    If DocNum = 0 Then
        oDS_PS_MM170H.setValue "DocNum", 0, "1"
    Else
        oDS_PS_MM170H.setValue "DocNum", 0, DocNum    ' ȭ�鿡 ������ �ȵǱ� ����
    End If
    Exit Sub
FormClear_Error:
    Sbo_Application.SetStatusBarMessage "FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function DataValidCheck() As Boolean
On Error GoTo DataValidCheck_Error
    '//��ȿ���˻�
    Exit Function
DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    '//�����ͺ�ȭ�� ����ó��
    Dim i$
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
    
        Case "Mat01"    'rowcount: �� ī��Ʈ�� ��ȯ, VisualRowCount: ������ �θ� �����ϰ� ���� ���̴� �� ī��Ʈ�� ��ȯ
            If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And _
            Trim(oMat01.Columns("GRDocNum").Cells(oRow).Specific.VALUE) <> "" Then
                oMat01.FlushToDataSource      '������ �ҽ��� ����� ��Ʈ�����κ��� ������ �ҽ� ���ڵ�� �� �θ� �����Ѵ�.
                AddMatrixRow 1, oMat01.RowCount, True
                oMat01.Columns("GRDocNum").Cells(oRow).Click ct_Regular 'Column: Į�� ������Ʈ�� collection�� ��ȯ�Ѵ�.
            End If
    End Select
    Set oRecordSet01 = Nothing
End Sub

Private Sub MTX01()
On Error GoTo MTX01_Error
    '//��Ʈ������ ������ �ε�
    Exit Sub
MTX01_Error:
    Sbo_Application.SetStatusBarMessage "MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_MM170H.GetValue("U_CardCode", 0) = "" Or _
             oDS_PS_MM170H.GetValue("U_BPLId", 0) = "" Or _
             oDS_PS_MM170H.GetValue("U_DocDate", 0) = ""
        
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    HeaderSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�ŷ�ó�ڵ�, �����, ��û���ڴ� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function
        
Private Function MatrixSpaceLineDel() As Boolean
'------------------------------------------------------------------------------
' ������ �������� ��ȿ���� �����Ѵ�
'------------------------------------------------------------------------------
On Error GoTo MatrixSpaceLineDel_Error
    Dim i           As Long
    Dim K           As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    Dim oRow        As Integer
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    ErrNum = 0
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '// ȭ����� ��Ʈ������ �Էµ� ������ ��� ������ͼҽ��� �ѱ��
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    ' Flushes current data from the user interface to the bounded data source, as follows:
    ' 1. Cleans the data source.
    ' 2. Copies each row from the matrix to the corresponding data source record.
    oMat01.FlushToDataSource

    '// ����
    If oMat01.VisualRowCount <= 1 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '// �Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '// �̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    If oMat01.VisualRowCount > 0 Then
        
    For i = 0 To oMat01.VisualRowCount - 2
        oDS_PS_MM170L.Offset = i
            Select Case True
                Case oDS_PS_MM170L.GetValue("U_GRDocNum", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
    
                Case oDS_PS_MM170L.GetValue("U_RepayYN", i) = ""
                    ErrNum = 3
                    GoTo MatrixSpaceLineDel_Error
                    
                Case oDS_PS_MM170L.GetValue("U_RepayP", i) = ""
                    ErrNum = 4
                    GoTo MatrixSpaceLineDel_Error
            End Select
     Next i
     End If

    If oDS_PS_MM170L.GetValue("U_GRDocNum", oMat01.VisualRowCount - 1) = "" Then
        oDS_PS_MM170L.RemoveRecord oMat01.VisualRowCount - 1
    End If
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '���� �����Ͽ����� DB������ �ҽ��� �ٽ� �����´�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "�԰����� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "��ü������� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "��ü�ݾ� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function Exist_YN(DocNum$) As String

    Dim oRecordSet01   As SAPbobsCOM.Recordset
    Dim sQry           As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT Top 1 T1.DocNum FROM [@PS_MM170H] T1 "
    sQry = sQry & " WHERE T1.DocNum  = '" & DocNum & "'"
    oRecordSet01.DoQuery sQry
    
    Do Until oRecordSet01.EOF
        Exist_YN = oRecordSet01(0).VALUE
        oRecordSet01.MoveNext
    Loop
    
    If Trim(Exist_YN) = "" Then
        Exist_YN = ""
        Exit Function
    End If
    
    Set oRecordSet01 = Nothing
End Function

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If pval.BeforeAction = True Then
        If Sbo_Application.MessageBox("���� ���� �Ͻðڽ��ϱ�?", 1, "OK", "NO") <> 1 Then
            BubbleEvent = False
        End If
        '//������� ��������ɿ��ΰ˻�
    ElseIf pval.BeforeAction = False Then
        For i = 1 To oMat01.VisualRowCount
            oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
        Next i
'        oMat01.Clear
        oMat01.FlushToDataSource
        Call oDS_PS_MM170L.RemoveRecord(oDS_PS_MM170L.Size - 1)
        oMat01.LoadFromDataSource
        If oMat01.RowCount = 0 Then
            Call AddMatrixRow(0, oMat01.RowCount, True)
        Else
            If Trim(oDS_PS_MM170L.GetValue("U_ItemCode", oMat01.RowCount - 1)) <> "" Then
                Call AddMatrixRow(1, oMat01.RowCount, True)
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
        

Private Function Item_Error_Message(ErrNum As Integer)
On Error GoTo Item_Error_Message

GoTo Item_Error_Message

Item_Error_Message:
 If ErrNum = 1 Then
    MDC_Com.MDC_GF_Message "�а�ó������ ���� �Է��ϼ���.", "E"
 ElseIf ErrNum = 2 Then
    MDC_Com.MDC_GF_Message "������ Close �Ǵ� Cancel �Ǿ����ϴ�.", "E"
 ElseIf ErrNum = 3 Then
    MDC_Com.MDC_GF_Message "�а�����:Y�� �� ��� �� �� �ֽ��ϴ�.", "E"
  ElseIf ErrNum = 4 Then
    MDC_Com.MDC_GF_Message "�ŷ�ó�ڵ��, ������� ���� �Է��ϼ���.", "E"
 End If
End Function
        
        
Private Function Create_oJournalEntries(ChkType As Integer) As Boolean
On Error GoTo Create_oJournalEntries
    Dim f_oJournalEntries As SAPbobsCOM.JournalEntries
    
    Dim i             As Long
    Dim ErrNum        As Integer
    Dim ErrCode       As Long
    Dim ErrMsg        As String
    Dim RetVal&
    Dim sTransId$
        
    Dim SCardCode$, sDocDate$
    Dim sSum            As Currency
    Dim sCC             As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    If (Sbo_Company.InTransaction = True) Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    SCardCode = Trim(oDS_PS_MM170H.GetValue("U_CardCode", 0))
    sDocDate = Format(oDS_PS_MM170H.GetValue("U_JdtDate", 0), "0000-00-00")
    sSum = Trim(oDS_PS_MM170H.GetValue("U_Sum", 0))
 
    Set f_oJournalEntries = Sbo_Company.GetBusinessObject(oJournalEntries)
        
    With f_oJournalEntries
         .ReferenceDate = sDocDate            '������
         .DueDate = sDocDate
         .TaxDate = sDocDate
         
         .Lines.Add
         .Lines.SetCurrentLine 0
         .Lines.AccountCode = "21101010"     '��������
         .Lines.ShortName = SCardCode        'G/L����/BP �ڵ�
         .Lines.Debit = sSum                 '����
         .Lines.LineMemo = "��ü���"
         
         .Lines.SetCurrentLine 1
         .Lines.AccountCode = "43125020"     '������-��Ÿ
         .Lines.ShortName = "43125020"
         .Lines.Credit = sSum                '�뺯
         .Lines.LineMemo = "��ü���"
         .UserFields("U_BPLId").VALUE = oForm01.Items("BPLId").Specific.VALUE '�����
         
    End With
    '// �Ϸ�
    RetVal = f_oJournalEntries.Add
    If (0 <> RetVal) Then
        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
        GoTo Create_oJournalEntries
    End If

    sCC = "Y"
    
    If ChkType = 1 Then
        Sbo_Company.GetNewObjectCode sTransId
        sQry = "Update [@PS_MM170H] Set U_JdtNo = '" & sTransId & "', U_JdtDate = '" & sDocDate & "', U_Sum = '" & sSum & "', U_JdtCC = '" & sCC & "' "
        sQry = sQry & "Where DocNum = '" & Trim(oDS_PS_MM170H.GetValue("DocNum", 0)) & "'"
        oRecordSet01.DoQuery sQry
        If (Sbo_Company.InTransaction = True) Then
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If

  oDS_PS_MM170H.setValue "U_JdtNo", 0, sTransId
  oDS_PS_MM170H.setValue "U_JdtDate", 0, Format(Now, "yyyymmdd")
  oDS_PS_MM170H.setValue "U_Sum", 0, sSum
  oDS_PS_MM170H.setValue "U_JdtCC", 0, sCC

  Set oRecordSet01 = Nothing
  Set f_oJournalEntries = Nothing
  Create_oJournalEntries = True
  
  oForm01.Items("Btn02").Enabled = False
  oForm01.Items("Btn03").Enabled = True
  
  Exit Function
  
'///////////////////////////////////////////////////////////////////////////////////////
Create_oJournalEntries:
    Set oRecordSet01 = Nothing
    Set f_oJournalEntries = Nothing
    If Sbo_Company.InTransaction Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Create_oJournalEntries = False
    MDC_Com.MDC_GF_Message "Create_oJournalEntries:" & Err.Description & ErrMsg, "E"
End Function
        
        
Private Function Cancel_oJournalEntries(ChkType As Integer) As Boolean
On Error GoTo Cancel_oJournalEntries
    Dim f_oJournalEntries As SAPbobsCOM.JournalEntries
    
    Dim i             As Long
    Dim ErrNum        As Integer
    Dim ErrCode       As Long
    Dim ErrMsg        As String
    Dim RetVal&
    Dim sTransId$
        
    Dim SCardCode$, sDocDate$
    Dim sSum            As Currency
    Dim sCC             As String
        
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    If (Sbo_Company.InTransaction = True) Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    
    Sbo_Company.StartTransaction
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    Set f_oJournalEntries = Nothing
    Set f_oJournalEntries = Sbo_Company.GetBusinessObject(oJournalEntries)
        
    If (f_oJournalEntries.GetByKey(Trim(oDS_PS_MM170H.GetValue("U_JdtNo", 0))) = False) Then
        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
        GoTo Cancel_oJournalEntries
    End If
    
    '// �Ϸ�
    RetVal = f_oJournalEntries.Cancel
    If (0 <> RetVal) Then
        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
        GoTo Cancel_oJournalEntries
    End If
    
    sCC = "N"
    
    If ChkType = 1 Then
        Sbo_Company.GetNewObjectCode sTransId
        sQry = "Update [@PS_MM170H] Set U_JdtCan = '" & sTransId & "', U_Sum = '" & sSum & "', U_JdtCC = '" & sCC & "' "
        sQry = sQry & "Where DocNum = '" & Trim(oDS_PS_MM170H.GetValue("DocNum", 0)) & "'"
        oRecordSet01.DoQuery sQry
        
        If (Sbo_Company.InTransaction = True) Then
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If
    
'    oForm01.Update
  oDS_PS_MM170H.setValue "U_JdtCan", 0, sTransId
  oDS_PS_MM170H.setValue "U_Sum", 0, sSum
  oDS_PS_MM170H.setValue "U_JdtCC", 0, sCC
  
  Set oRecordSet01 = Nothing
  Set f_oJournalEntries = Nothing
  Cancel_oJournalEntries = True
  
  oForm01.Items("Btn02").Enabled = False
  oForm01.Items("Btn03").Enabled = False
  
  Exit Function
  
'///////////////////////////////////////////////////////////////////////////////////////
Cancel_oJournalEntries:
    Set oRecordSet01 = Nothing
    Set f_oJournalEntries = Nothing
    If Sbo_Company.InTransaction Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Cancel_oJournalEntries = False
    MDC_Com.MDC_GF_Message "Create_oJournalEntries:" & Err.Description & ErrMsg, "E"
End Function