VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY510.cls
'//  Module         : ��õ¡��>�ٷμҵ�
'//  Desc           : �����ٹ��� �ϰ�����
'//  FormType       : 2010110510
'//  Create Date    : 2010.01.05
'//  Modified Date  :
'//  Creator        : Choi Dong Kwon
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit
Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oGrid           As SAPbouiCOM.Grid
Private oDS_ZPY510      As SAPbouiCOM.DataTable
Private mJSNYER         As String

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY510.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    
    '------------------------------------------------------------------------
    '//�������� ��Ʈ������ Ʋ��쿡 ������ó�� �ε� �ǵ��� ���� ���
    '------------------------------------------------------------------------
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY510_" & GetTotalFormsCount
    
    '�� �Ҵ�
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '--------------------------------------------------------------------------------------------------------------
    '�÷��ǿ� ���� ��´�   **�÷����̶� ��ü�� ��� ���� �迭�μ� ���⼭�� Ȱ��ȭ�Ǿ��� �ִ� ���� ��� �ִ�
    '--------------------------------------------------------------------------------------------------------------
    
    AddForms Me, oFormUniqueID, "ZPY510"
    oForm.SupportedModes = -1
    oForm.Mode = fm_OK_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    'oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    
    Call CreateItems
    
    oForm.EnableMenu ("1293"), False  '/ �����
    oForm.EnableMenu ("1284"), False  '/ ���

    oForm.Update
    oForm.Visible = True
    
    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
     Sbo_Application.StatusBar.SetText "�����ٹ��� �ϰ������� �����ų �� �����ϴ�.", bmt_Short, smt_Error
     If (oForm Is Nothing) = False Then
        Set oForm = Nothing
    End If
End Sub

'*******************************************************************
' Item Initial
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oEdit           As SAPbouiCOM.EditText
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '// UserDataSources
    With oForm.DataSources.UserDataSources
        .Add "JSNYER", dt_SHORT_TEXT, 4
        .Add "MSTCOD", dt_SHORT_TEXT, 8
        .Add "MSTNAM", dt_SHORT_TEXT, 30
    End With
    
    Set oEdit = oForm.Items("JSNYER").Specific
    oEdit.DataBind.SetBound True, "", "JSNYER"
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    
    oForm.DataSources.UserDataSources("JSNYER").ValueEx = Format$(Now, "YYYY")
    
    '//�����
    Set oCombo = oForm.Items("FCLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "��ü"
    oForm.Items("FCLTCOD").DisplayDesc = True
    oCombo.Select "%", psk_ByValue
    
    Set oCombo = oForm.Items("TCLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "��ü"
    oForm.Items("TCLTCOD").DisplayDesc = True
    oCombo.Select "%", psk_ByValue
    
    '/ Grid
    Set oGrid = oForm.Items("Grid1").Specific
    Call oForm.DataSources.DataTables.Add("ZPY510")
    
    Set oDS_ZPY510 = oForm.DataSources.DataTables.Item("ZPY510")
    Call oDS_ZPY510.ExecuteQuery("EXEC ZPY510 '1900', '', '', ''")
    oGrid.DataTable = oDS_ZPY510
    
    Call titleSetting
        
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TitleSetting
' Author    : Choi Dong Kwon
' Date      : 2008-07-15
' Purpose   : Grid�� Column Title ����
'---------------------------------------------------------------------------------------
'
Private Sub titleSetting()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim ErrNum          As Integer
    Dim i               As Integer
    
    Dim COLNAM(24)      As String

'/ Initial
    ErrNum = 0
    
    oForm.Freeze True
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'/  �÷���
    COLNAM(0) = "����"
    COLNAM(1) = "�����ȣ"
    COLNAM(2) = "����"
    COLNAM(3) = "�������"
    COLNAM(4) = "�ڻ��ڵ�"
    COLNAM(5) = "�����ڻ��ڵ�"
    COLNAM(6) = "����ڹ�ȣ"
    COLNAM(7) = "�ͼӽ�����"
    COLNAM(8) = "�ͼ�������"
    COLNAM(9) = "���ذ��������"
    COLNAM(10) = "���ذ���������"
    COLNAM(11) = "�޿��ݾ�"
    COLNAM(12) = "�󿩱ݾ�"
    COLNAM(13) = "������"
    COLNAM(14) = "�ֽĸż����ñ��������"
    COLNAM(15) = "�츮������������"
    COLNAM(16) = "������Ѱ�"
    COLNAM(17) = "�ǰ�����"
    COLNAM(18) = "���뺸��"
    COLNAM(19) = "���ο���"
    COLNAM(20) = "���ݺ����"
    COLNAM(21) = "�ҵ漼"
    COLNAM(22) = "�ֹμ�"
    COLNAM(23) = "��Ư��"
    COLNAM(24) = "��������"
    
    '// �÷��� ����
    For i = 0 To 24
        oGrid.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        
        If i >= 11 Then
            oGrid.Columns.Item(i).RightJustified = True
        Else
            oGrid.Columns.Item(i).RightJustified = False
        End If
        
        If i > 0 Then
            oGrid.Columns.Item(i).Editable = False
        Else
            oGrid.Columns.Item(i).Editable = True
        End If
        
    Next i
    
    '// Grid�� �÷��� ComboBox, CheckBox ����
    Call Grid_Col_Define
    oGrid.AutoResizeColumns

'/
    oForm.Freeze False

    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    oForm.Freeze False
    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
    
    Else
        Sbo_Application.StatusBar.SetText "titleSetting ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub

Private Sub Grid_Display()
On Error GoTo Error_Message
    Dim sQry        As String
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim iRow        As Integer
    
    Dim MSTCOD      As String
    Dim FCLTCOD     As String
    Dim TCLTCOD     As String
    
'/  Default Value
    ErrNum = 0
    iRow = 0

    mJSNYER = Trim$(oForm.DataSources.UserDataSources("JSNYER").ValueEx)
    MSTCOD = Trim$(oForm.DataSources.UserDataSources("MSTCOD").ValueEx)

'/ Check
    Select Case True
    Case oForm.Items("FCLTCOD").Specific.Selected Is Nothing
        ErrNum = 1
        GoTo Error_Message
    Case oForm.Items("TCLTCOD").Specific.Selected Is Nothing
        ErrNum = 1
        GoTo Error_Message
    Case mJSNYER = ""
        ErrNum = 2
        GoTo Error_Message
    End Select
    FCLTCOD = oForm.Items("FCLTCOD").Specific.Selected.Value
    TCLTCOD = oForm.Items("TCLTCOD").Specific.Selected.Value
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Mode = fm_OK_MODE
'/ ��ȸ
    sQry = " Exec ZPY510 '" & Trim$(mJSNYER) & "', '" & Trim$(MSTCOD) & "', '" & _
                              Trim$(FCLTCOD) & "', '" & Trim$(TCLTCOD) & "'"
    Debug.Print sQry
    Call oDS_ZPY510.ExecuteQuery(sQry)
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    If iRow = 1 Then
        oRecordSet.DoQuery sQry
        iRow = oRecordSet.RecordCount
    End If

    Sbo_Application.StatusBar.SetText iRow & " ���� �ֽ��ϴ�.", bmt_Short, smt_Success
    Set oRecordSet = Nothing
    Call titleSetting
    
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "�ڻ� �ڵ�� �ʼ��Դϴ�. �����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Success
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "�ͼӳ⵵�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Success
    Else
        Sbo_Application.StatusBar.SetText "Grid_Display Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)

    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    '// ��� �Է½� ���� ��ȸ
    If oUID = "MSTCOD" Then
        If oForm.Items(oUID).Specific.Value = "" Then
            oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = ""
        Else
            oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = _
            MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", oForm.Items(oUID).Specific.Value)
        End If
        
        oForm.Update
    End If
    oForm.Freeze False
End Sub

'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
'et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_ITEM_PRESSED
            If pval.BeforeAction Then
                If pval.ItemUID = "Btn1" Then  '// ���� ���� ��ư
                    Call Grid_Display
                    BubbleEvent = False
                ElseIf pval.ItemUID = "Btn2" Then  '// File Upload ��ư
                    Call Grid_Save
                    BubbleEvent = False
                ElseIf pval.ItemUID = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then   '/ ChooseBtn�������Ʈ
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUID = "Grid1" And pval.ColUID = "U_CHECK" And pval.Row = -1 Then                   '/ ��üüũ
                    Call CheckAll
                End If
            End If

'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
            If pval.BeforeAction = False Then
                If pval.ItemUID = "Grid1" And pval.ColUID = "U_MSTCOD" And pval.CharPressed = 9 Then
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oDS_ZPY510.GetValue("U_MSTCOD", pval.Row) + "'") = True Or _
                        oDS_ZPY510.GetValue("U_MSTCOD", pval.Row) = "" Then
                        oGrid.Columns.Item("U_MSTCOD").CLICK pval.Row
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
            
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUID = "MSTCOD" Then
                    FlushToItemValue pval.ItemUID
                ElseIf pval.ItemUID = "Grid1" Then
                    Call FlushToItemValue(pval.ColUID, pval.Row)
                End If
            End If

'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            '----------------------------------------------------
            '�÷��ǿ��� ������ ��� �޸� ����
            '----------------------------------------------------
            If pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oDS_ZPY510 = Nothing
                Set oGrid = Nothing
                Set oForm = Nothing
            End If
    End Select

    Exit Sub
'//////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)

    If pval.BeforeAction = True Then
        Exit Sub
    End If

    Select Case pval.MenuUID
        Case "1287" '/ ����
        Case "1281", "1282":
        Case "1288" To "1291":
        Case "1293"
    End Select
    Exit Sub
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    Dim i       As Long
    Dim sQry    As String
    Dim oCombo  As SAPbouiCOM.ComboBox
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormDataEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select

    End If
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
    
Raise_FormDataEvent_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

Private Sub CheckAll()
    Dim CheckType   As String
    Dim oRow        As Long

    oForm.Freeze True
    CheckType = "Y"
    For oRow = 0 To oGrid.Rows.Count - 1
        If Trim$(oDS_ZPY510.GetValue("U_CHECK", oRow)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For oRow = 0 To oGrid.Rows.Count - 1
        oDS_ZPY510.Rows.Offset = oRow
        If CheckType = "N" Then
            oDS_ZPY510.setValue "U_CHECK", oRow, "Y"
        Else
            oDS_ZPY510.setValue "U_CHECK", oRow, "N"
        End If
    Next
    oForm.Freeze False
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Grid_Col_Define
' Author    : Choi Dong Kwon
' Date      : 2008-07-14
' Purpose   : Grid�� Column�鿡 ���Ͽ� LinkButton, ComboBox, CheckBox���� ����
'---------------------------------------------------------------------------------------
'
Private Sub Grid_Col_Define()

On Error GoTo Error_Message

    Dim sQry            As String
    Dim ErrNum          As Integer

    Dim oColumn         As SAPbouiCOM.GridColumn
    Dim oEditCol        As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oColumn = oGrid.Columns.Item("U_CHECK")
    oColumn.Type = gct_CheckBox
    
    '// ��������� LinkButton �߰�
    Set oEditCol = oGrid.Columns.Item("EMPID")
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    
    '// �ڻ��ڵ�
    '// EditText Column => ComboBox Column���� ����
    Set oColumn = oGrid.Columns.Item("CLTCOD")
    oColumn.Type = gct_ComboBox
    
    Set oComboCol = oGrid.Columns.Item("CLTCOD")
    oComboCol.DisplayType = cdt_Description
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oComboCol.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
        oRecordSet.MoveNext
    Loop
    
    Set oColumn = oGrid.Columns.Item("JCLTCOD")
    oColumn.Type = gct_ComboBox
    
    Set oComboCol = oGrid.Columns.Item("JCLTCOD")
    oComboCol.DisplayType = cdt_Description
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oComboCol.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
        oRecordSet.MoveNext
    Loop
    
    Set oColumn = Nothing
    Set oEditCol = Nothing
    Set oComboCol = Nothing
    Set oRecordSet = Nothing
    Exit Sub

'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Set oEditCol = Nothing
    Set oComboCol = Nothing
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
    
    Else
        Sbo_Application.StatusBar.SetText "Grid_Col_Define ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Grid_Save
' Author    : Choi Dong Kwon
' Date      : 2008-07-15
' Purpose   : Grid�� ������ �ϰ� �����ϴ� ���ν���
'---------------------------------------------------------------------------------------
'
Private Sub Grid_Save()

On Error GoTo Error_Message

    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim ErrNum      As Integer
    
    Dim oRow        As Long
    Dim UserId      As Long
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    UserId = oCompany.UserSignature
    
    ErrNum = 0
    For oRow = 0 To oGrid.Rows.Count - 1
    
        '// üũ�� �ุ ����
        If oDS_ZPY510.GetValue("U_CHECK", oRow) = "Y" Then
        
            oCompany.StartTransaction       '/ Ʈ����� ����
            
            sQry = "EXEC ZPY510_1 '" & mJSNYER & "', " & _
                                 "'" & oDS_ZPY510.GetValue("MSTCOD", oRow) & "', " & _
                                 "'" & oDS_ZPY510.GetValue("CLTCOD", oRow) & "', " & _
                                 "'" & oDS_ZPY510.GetValue("JCLTCOD", oRow) & "', " & _
                                       CStr(UserId) & " "
            oRecordSet.DoQuery sQry
            Debug.Print sQry
            oCompany.EndTransaction wf_Commit   '/ Ʈ����� ����
        End If
    Next
    
    Call Grid_Display
    Sbo_Application.StatusBar.SetText "�����ٹ��� �ϰ������� �Ϸ�Ǿ����ϴ�.", bmt_Short, smt_Success
    'oForm.Mode = fm_OK_MODE
    Set oRecordSet = Nothing
    Exit Sub
    
Error_Message:
    
    oCompany.EndTransaction wf_RollBack '/ Ʈ����� RollBack
    
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "Grid_Save ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error

End Sub