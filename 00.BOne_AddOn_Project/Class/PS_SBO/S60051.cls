VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S60051"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : S60051.cls
'//  Module         :
'//  Description    : �ڱݰ���>��������-��������
'//  FormType       : 60051
'//  Create Date    : 2011.02.17
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix

Private oMat01Row           As Long

Public Sub LoadForm(ByVal oFormUniqueID As String)
    oFormUniqueID01 = oFormUniqueID
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    Set oMat01 = oForm01.Items("5").Specific
    Call CreateItems
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        If oForm01.Items("4").Specific.Selected.VALUE = "F" Then '�ε��� ���
                            If Create_oJournalEntries(2) = False Then
                                BubbleEvent = False
                                Exit Sub
                            Else
                                MDC_Com.MDC_GF_Message "�ε��������� �̵��� �Ϸ�Ǿ����ϴ�.", "S"
                            End If
                        
                            oForm01.Items("38").Click ct_Regular
                            BubbleEvent = False
                        End If
                    End If
                End If
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
        End Select
        
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_FORM_LOAD
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
                If pval.ItemUID = "38" Then
                    Call oForm01.Items("4").Specific.ValidValues.Add("F", "�ε�")
                End If
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
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
'                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
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
            Case "1285": '����
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1285": '����
            Case "1287": '����
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
'                oForm01.Freeze True
'                Call oForm01.Items("BPLId01").Specific.Select(User_BPLId(), psk_ByValue)
'                oForm01.Items("6").Click ct_Regular
'                oForm01.Freeze False
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
        If eventInfo.ItemUID = "76" Then
            If eventInfo.Row > 0 Then
                oMat01Row = eventInfo.Row
            End If
        End If
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    Dim stdItem         As Object
    Dim newItem         As Object
    Dim oEdit           As SAPbouiCOM.EditText
    Dim optBtn          As SAPbouiCOM.OptionBtn
    Dim oCombo01        As SAPbouiCOM.ComboBox
    Dim oCombo02        As SAPbouiCOM.ComboBox
    Dim oStatic         As SAPbouiCOM.StaticText
    
    Dim oRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)

    oForm01.Freeze True
    
    '//�����-ComboBox
'    Set newItem = oForm01.Items.Add("Static01", it_STATIC)
'    Set stdItem = oForm01.Items("2006")
'    newItem.Left = stdItem.Left + 93
'    newItem.Top = stdItem.Top
'    newItem.Height = stdItem.Height
'    newItem.Width = stdItem.Width
''    newItem.LinkTo = "TradeType"
'    Set oStatic = newItem.Specific
'    oStatic.Caption = "�����"
'
'    Set newItem = oForm01.Items.Add("BPLId01", it_COMBO_BOX)
'    Set stdItem = oForm01.Items("2007")
'    newItem.Left = stdItem.Left + 93
'    newItem.Top = stdItem.Top
'    newItem.Height = stdItem.Height
'    newItem.Width = stdItem.Width + 40
'    newItem.DisplayDesc = True
'    Set oCombo01 = newItem.Specific
'    Call oCombo01.DataBind.SetBound(True, "OJDT", "U_BPLId")
'
'    sQry = "select BPLId, BPLName from [OBPL] order by BPLId"
'    oRecordset.DoQuery sQry
'    Do Until oRecordset.EOF
'        oCombo01.ValidValues.Add Trim(oRecordset.Fields(0).VALUE), Trim(oRecordset.Fields(1).VALUE)
'        oRecordset.MoveNext
'    Loop
'    Call oCombo01.Select(User_BPLId(), psk_ByValue)
    
    
'    sQry = "select BPLId, BPLName from [OBPL] order by BPLId"
'    oRecordset.DoQuery sQry
'    Do Until oRecordset.EOF
'        oCombo02.ValidValues.Add Trim(oRecordset.Fields(0).Value), Trim(oRecordset.Fields(1).Value)
'        oRecordset.MoveNext
'    Loop

    oForm01.Freeze False
    
    Set oCombo01 = Nothing
    Set oCombo02 = Nothing
    Set oRecordset = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    Set oCombo01 = Nothing
    Set oCombo02 = Nothing
    Set oRecordset = Nothing
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, j&, DocNum$, LineId$, Counts$, cnt$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Dim oRecordset    As SAPbobsCOM.Recordset
        
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID

    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "76" Then
        Select Case oCol
        
        End Select
    End If
     
    Set oRecordset = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub Form_Resize(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Form_Resize_Error
    
'     oForm01.Items("Static01").Left = oForm01.Items("2006").Left + 93
'     oForm01.Items("BPLId01").Left = oForm01.Items("2007").Left + 93
'     oForm01.Items("Static02").Top = oForm01.Items("2001").Top
'     oForm01.Items("Static02").Left = oForm01.Items("2001").Left + 161
'     oForm01.Items("BPLId02").Top = oForm01.Items("2000").Top
'     oForm01.Items("BPLId02").Left = oForm01.Items("2000").Left + 161

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Form_Resize_Error:
    Sbo_Application.SetStatusBarMessage "Form_Resize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function Create_oJournalEntries(ChkType As Integer) As Boolean
On Error GoTo Create_oJournalEntries_Error
    Dim f_oJournalEntries As SAPbobsCOM.JournalEntries         '//�а����� ��ü
    
    Dim oRow1&, oRow2&
    Dim i               As Integer
    Dim j               As Integer
    Dim K               As Integer
    Dim ErrNum          As Integer
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim RetVal&
    Dim VTransId$
        
    Dim vBoeKey$, vBPLId$
    
    Dim oRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String
     
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Sbo_Company.StartTransaction
    Set f_oJournalEntries = Nothing
    Set f_oJournalEntries = Sbo_Company.GetBusinessObject(oJournalEntries)
    
    i = 0
    With f_oJournalEntries
    
    '// Header  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    .ReferenceDate = Format(Trim(oForm01.Items("55").Specific.VALUE), "0000-00-00")
    .DueDate = Format(Trim(oForm01.Items("55").Specific.VALUE), "0000-00-00")
    .TaxDate = Format(Trim(oForm01.Items("61").Specific.VALUE), "0000-00-00")
    .Memo = "�߽ɿ��� �ε������̵�"
    
    
        '// Line /////////////////////////////////////////////////////////////
        For j = 1 To oMat01.VisualRowCount
            If Trim(oMat01.Columns("9").Cells(j).Specific.Checked) = True Then
               
                sQry = "select BoeKey from OBOE where BoeType = 'I' and BoeNum = '" & Trim(oMat01.Columns("7").Cells(j).Specific.VALUE) & "'"
                oRecordset.DoQuery sQry
                vBoeKey = oRecordset.Fields("BoeKey").VALUE
                                
                sQry = "select U_BPLId from ORCT where BoeAbs = '" & vBoeKey & "'"
                oRecordset.DoQuery sQry
                vBPLId = oRecordset.Fields("U_BPLId").VALUE
                '��ǥ��� �����
                .UserFields("U_BPLId").VALUE = vBPLId
               
                '//����(Debit)--------------------------------------------------------
                .Lines.Add
                .Lines.SetCurrentLine i
    
                '.Lines.AccountCode = "5201270100"
                .Lines.ShortName = Trim(oMat01.Columns("28").Cells(j).Specific.VALUE)
                .Lines.ControlAccount = "11104070" '�ε�����
                .Lines.Debit = Val(oMat01.Columns("2").Cells(j).Specific.VALUE)
                .Lines.Reference1 = vBoeKey
                .Lines.LineMemo = "�������� ��ȣ(" & Trim(oMat01.Columns("7").Cells(j).Specific.VALUE) & ") : �߽ɿ��� �ε��̵�"
                .Lines.UserFields("U_BPLId").VALUE = vBPLId
                
                i = i + 1
                
                '//�뺯(Credit)--------------------------------------------------------
                .Lines.Add
                .Lines.SetCurrentLine i
                
                '.Lines.AccountCode = "5201270100"
                .Lines.ShortName = Trim(oMat01.Columns("28").Cells(j).Specific.VALUE)
                .Lines.ControlAccount = "11104060" '��������
                .Lines.Credit = Val(oMat01.Columns("2").Cells(j).Specific.VALUE)
                .Lines.Reference1 = vBoeKey
                .Lines.LineMemo = "�������� ��ȣ(" & Trim(oMat01.Columns("7").Cells(j).Specific.VALUE) & ") : �߽ɿ��� �ε��̵�"
                .Lines.UserFields("U_BPLId").VALUE = vBPLId
                
                i = i + 1
                
            End If
        Next j
    End With
                
    '// �Ϸ� ///////////////////////////////////////////////////////////////
    RetVal = f_oJournalEntries.Add
    If (0 <> RetVal) Then
        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
        GoTo Create_oJournalEntries_Error
    End If

    If ChkType = 1 Then
        Sbo_Company.EndTransaction wf_RollBack
    ElseIf ChkType = 2 Then

        Sbo_Company.GetNewObjectCode VTransId
                    
        For j = 1 To oMat01.VisualRowCount
            If Trim(oMat01.Columns("9").Cells(j).Specific.Checked) = True Then
                '// OBOE ���� Upate
                sQry = "update OBOE set BoeStatus='F'" 'F : �ε�����
                sQry = sQry & " where BoeType='I'"    ' I : ��������
                sQry = sQry & " and BoeNum='" & Trim(oMat01.Columns("7").Cells(j).Specific.VALUE) & "'"
                oRecordset.DoQuery sQry
                
                '// �������� Insert
                sQry = "select BoeKey from OBOE where BoeType = 'I' and BoeNum = '" & Trim(oMat01.Columns("7").Cells(j).Specific.VALUE) & "'"
                oRecordset.DoQuery sQry
                vBoeKey = oRecordset.Fields("BoeKey").VALUE

                sQry = "insert into Z60051 values('" & VTransId & "','" & vBoeKey & "','" & Trim(oMat01.Columns("7").Cells(j).Specific.VALUE) & "')"
                oRecordset.DoQuery sQry
            End If
        Next j
        
        Sbo_Company.EndTransaction wf_Commit
    End If

    Set oRecordset = Nothing
    Set f_oJournalEntries = Nothing
    Create_oJournalEntries = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Create_oJournalEntries_Error:
    Set oRecordset = Nothing
    Set f_oJournalEntries = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Create_oJournalEntries = False
    Sbo_Application.SetStatusBarMessage "Create_oJournalEntries:" & Err.Description & ErrMsg, bmt_Short, True
End Function