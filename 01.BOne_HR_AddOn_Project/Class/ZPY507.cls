VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY507"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY507.cls
'//  Module         : �λ����>��õ¡��>�ٷμҵ�
'//  Desc           : ���� ��� ��ȸ(��ü)
'//  FormType       : 2010110507
'//  Create Date    : 2009.12.13
'//  Modified Date  :
'//  Creator        : Choi Dong Kwon
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit
Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oGrid1          As SAPbouiCOM.Grid
Private oDS_ZPY507      As SAPbouiCOM.DataTable

Private Sub titleSetting()
On Error GoTo Error_Message
    Dim sQry            As String
    Dim iCol            As Long
    
    Dim COLNAM(150)     As String
    Dim oEditCol        As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    Dim oColumn         As SAPbouiCOM.GridColumn
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
'/ Initial
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    COLNAM(0) = "����⵵"
    COLNAM(1) = "���걸��"
    COLNAM(2) = "�Ű�����"
    COLNAM(3) = "��������"
    COLNAM(4) = "�����"
    COLNAM(5) = "�����ȣ"
    COLNAM(6) = "�������"
    COLNAM(7) = "�����"
    COLNAM(8) = "�����ڷ�No"
    COLNAM(9) = "�ҵ��׸�No"
    COLNAM(10) = "��������No"
    COLNAM(11) = "���װ��No"
    COLNAM(12) = "�Ƿ��No"
    COLNAM(13) = "��α�No"
    COLNAM(14) = "����.����No"
    COLNAM(15) = "���ٹ����Ѱ�"
    COLNAM(16) = "���ٹ����Ѱ�"
    COLNAM(17) = "�������"
    COLNAM(18) = "�ѱ޿�"
    COLNAM(19) = "�ٷμҵ����"
    COLNAM(20) = "�ٷμҵ�ݾ�"
    COLNAM(21) = "���ΰ����ݾ�"
    COLNAM(22) = "���������"
    COLNAM(23) = "����ڰ�����"
    COLNAM(24) = "�ξ簡����"
    COLNAM(25) = "�ξ簡������"
    COLNAM(26) = "��ο���ο�"
    COLNAM(27) = "��ο�����"
    COLNAM(28) = "������ο�"
    COLNAM(29) = "����ΰ�����"
    COLNAM(30) = "�γ�������"
    COLNAM(31) = "�γ��ڰ�����"
    COLNAM(32) = "�ڳ�����ο�"
    COLNAM(33) = "�ڳ��������"
    COLNAM(34) = "����Ծ��ο�"
    COLNAM(35) = "����Ծ����"
    COLNAM(36) = "���ڳ��ο�"
    COLNAM(37) = "���ڳ����"
    COLNAM(38) = "���ο���"
    COLNAM(39) = "��Ÿ����(����������)"
    COLNAM(40) = "��Ÿ����(���ο���)"
    COLNAM(41) = "��Ÿ����(�縳�б�����������)"
    COLNAM(42) = "��Ÿ����(������ü������)"
    COLNAM(43) = "��������(�ٷ��������޿������)"
    COLNAM(44) = "��������(���б���ΰ���)"
    COLNAM(45) = "�����(�ǰ�����)"
    COLNAM(46) = "�����(���뺸��)"
    COLNAM(47) = "�����(���强����)"
    COLNAM(48) = "�����(���������)"
    COLNAM(49) = "�Ƿ������ݾ�"
    COLNAM(50) = "����������ݾ�"
    COLNAM(51) = "�����������Աݿ����ݻ�ȯ-������"
    COLNAM(52) = "�����������Աݿ����ݻ�ȯ-������"
    COLNAM(53) = "������"
    COLNAM(54) = "����������ڻ�ȯ��-15��̸�"
    COLNAM(55) = "����������ڻ�ȯ��-29������"
    COLNAM(56) = "����������ڻ�ȯ��-30���̻�"
    COLNAM(57) = "��αݰ����ݾ�"
    COLNAM(58) = "ȥ��,�̻�,��ʺ�"
    COLNAM(59) = "Ư��������"
    COLNAM(60) = "ǥ�ذ���"
    COLNAM(61) = "�����ҵ�ݾ�"
    COLNAM(62) = "���ο����������"
    COLNAM(63) = "��������ҵ����"
    COLNAM(64) = "�ұ�������αݼҵ����"
    COLNAM(65) = "���ø�������(û������)"
    COLNAM(66) = "���ø�������(����û����������)"
    COLNAM(67) = "���ø�������(������ø�������)"
    COLNAM(68) = "���ø�������(�ٷ������ø�������)"
    COLNAM(69) = "�����������ڰ���"
    COLNAM(70) = "�ſ�ī��ҵ����"
    COLNAM(71) = "�츮�������հ���"
    COLNAM(72) = "����ֽ�������ҵ����"
    COLNAM(73) = "���������߼ұ������"
    COLNAM(74) = "��Ÿ�ҵ������"
    COLNAM(75) = "���ռҵ����ǥ��"
    COLNAM(76) = "���⼼��"
    COLNAM(77) = "�ҵ��"
    COLNAM(78) = "��Ư��"
    COLNAM(79) = "��������"
    COLNAM(80) = "���鼼�װ�"
    COLNAM(81) = "�ٷμҵ漼�װ���"
    COLNAM(82) = "�������հ���"
    COLNAM(83) = "�������Ա�"
    COLNAM(84) = "�����ó�ڱ�"
    COLNAM(85) = "�ܱ�����"
    COLNAM(86) = "���װ�����"
    COLNAM(87) = "�����ҵ漼"
    COLNAM(88) = "�����ֹμ�"
    COLNAM(89) = "������Ư��"
    COLNAM(90) = "��(��)�ٹ���_�ҵ漼"
    COLNAM(91) = "��(��)�ٹ���_�ֹμ�"
    COLNAM(92) = "��(��)�ٹ���_��Ư��"
    COLNAM(93) = "��(��)�ٹ���_�ҵ漼"
    COLNAM(94) = "��(��)�ٹ���_�ֹμ�"
    COLNAM(95) = "��(��)�ٹ���_��Ư��"
    COLNAM(96) = "�����ҵ漼"
    COLNAM(97) = "�����ֹμ�"
    COLNAM(98) = "������Ư��"

    For iCol = 0 To 98
        oGrid1.Columns.Item(iCol).Editable = False
        oGrid1.Columns.Item(iCol).TitleObject.Caption = COLNAM(iCol)
        If iCol >= 8 Then
            oGrid1.Columns.Item(iCol).RightJustified = True
        End If
        If oCompany.version >= "860040" Then                       '2007B PL18 �̻� �϶�(2007A ������ Ȯ�� �ʿ�)
             oGrid1.Columns.Item(iCol).TitleObject.Sortable = True
        End If

        
    Next iCol

    '// Link Button
    Set oEditCol = oGrid1.Columns.Item("EMPID")     '// �������
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    
    Set oEditCol = oGrid1.Columns.Item("DOCNO1")    '// �����ڷ�No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO2")    '// �ҵ��׸�No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO3")    '// ��������No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO4")    '// ���װ��No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO5")    '// �Ƿ��No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO6")    '// ��α�No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO7")    '// ����.����No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    
    '// ComboBox
    Set oColumn = oGrid1.Columns.Item("JSNGBN")     '// ���걸��
    oColumn.Type = gct_ComboBox
    
    Set oComboCol = oGrid1.Columns.Item("JSNGBN")
    oComboCol.DisplayType = cdt_Description
    oComboCol.ValidValues.Add "1", "��������(������)"
    oComboCol.ValidValues.Add "2", "�ߵ�����(������)"

    
    Set oColumn = oGrid1.Columns.Item("CLTCOD")     '// �����
    oColumn.Type = gct_ComboBox
    
    Set oComboCol = oGrid1.Columns.Item("CLTCOD")
    oComboCol.DisplayType = cdt_Description
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oComboCol.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
            oRecordSet.MoveNext
        Loop
    End If
   
    oGrid1.AutoResizeColumns
    oForm.Freeze False

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
    Sbo_Application.StatusBar.SetText "titleSetting ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)

    Select Case oUID
      Case "MSTCOD"
        If oForm.Items(oUID).Specific.String = "" Then
            oForm.DataSources.UserDataSources("MSTNAM").ValueEx = ""
        Else
            oForm.DataSources.UserDataSources("MSTNAM").ValueEx = MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(oUID).Specific.String & "'", "")
        End If
    End Select
    oForm.Update
    
End Sub
'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error
    
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
            If pval.ItemUID = "Btn01" Then  '// ã�� ��ư
                Call Grid_Display
                BubbleEvent = False
            ElseIf pval.ItemUID = "CBtn01" And oForm.Items("MSTCOD").Enabled = True Then   '/ ChooseBtn �������Ʈ
                oForm.Items("MSTCOD").CLICK ct_Regular
                Sbo_Application.ActivateMenuItem ("7425")
                BubbleEvent = False
            End If
        End If
        
'et_MATRIX_LINK_PRESSED'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_MATRIX_LINK_PRESSED
        If pval.BeforeAction Then
            If pval.ItemUID = "Grid1" And Left$(pval.ColUID, 5) = "DOCNO" Then
                Call UserFormLink(pval.ColUID, pval.Row)
                BubbleEvent = False
            End If
        End If

'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_VALIDATE
        If pval.BeforeAction = False And pval.ItemChanged = True Then
            If pval.ItemUID = "MSTCOD" Then
                Call FlushToItemValue(pval.ItemUID)
            End If
        End If

'et_FORM_UNLOAD'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_FORM_UNLOAD
        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        '�÷��ǿ��� ������ ��� �޸� ����
        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        If pval.BeforeAction = False Then
            RemoveForms oFormUniqueID
            Set oDS_ZPY507 = Nothing
            Set oGrid1 = Nothing
            Set oForm = Nothing
        End If
    End Select

    Exit Sub
'//////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub UserFormLink(LinkName As String, LineNum As Long)
    Dim oTmpObject  As Variant
    Dim DocNum      As String
    Dim JSNYER      As String
    Dim CLTCOD      As String
    Dim MSTCOD      As String
    
    If Trim$(LinkName) <> "" Then
        Select Case LinkName
        Case "DOCNO1"
            Set oTmpObject = New ZPY343
            DocNum = oDS_ZPY507.GetValue("DOCNO1", LineNum)
            If Trim$(DocNum) <> "" Then
                Call oTmpObject.LoadForm(DocNum)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO2"
            Set oTmpObject = New ZPY501
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO3"
            Set oTmpObject = New ZPY502
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO4"
            Set oTmpObject = New ZPY504
            DocNum = oDS_ZPY507.GetValue("DOCNO4", LineNum)
            If Trim$(DocNum) <> "" Then
                Call oTmpObject.LoadForm(DocNum)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO5"
            Set oTmpObject = New ZPY506
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO6"
            Set oTmpObject = New ZPY505
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO7"
            Set oTmpObject = New ZPY508
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        End Select
    
    End If
    Set oTmpObject = Nothing
End Sub

Private Sub Grid_Display()
On Error GoTo Error_Message
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset

    Dim FYEAR   As String:  Dim TYEAR   As String
    Dim JIGFDAT As String:  Dim JIGTDAT As String
    Dim SINFYMM As String:  Dim SINTYMM As String
    Dim MSTCOD  As String:  Dim MSTNAM  As String
    Dim CLTCOD  As String:  Dim JSNGBN  As String
    Dim PILMED  As Double:  Dim PILGBU  As Double
    Dim iRow    As Long
'/ Check
    ErrNum = 0
    iRow = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    With oForm.DataSources.UserDataSources
        FYEAR = Trim$(.Item("FYEAR").ValueEx)
        TYEAR = Trim$(.Item("TYEAR").ValueEx)
        JIGFDAT = Trim$(.Item("JIGFDAT").ValueEx)
        JIGTDAT = Trim$(.Item("JIGTDAT").ValueEx)
        SINFYMM = Trim$(.Item("SINFYMM").ValueEx)
        SINTYMM = Trim$(.Item("SINTYMM").ValueEx)
        MSTCOD = Trim$(.Item("MSTCOD").ValueEx)
        MSTNAM = Trim$(.Item("MSTNAM").ValueEx)
        PILMED = Val(.Item("PILMED").ValueEx)
        PILGBU = Val(.Item("PILGBU").ValueEx)
    End With
    CLTCOD = oForm.Items("CLTCOD").Specific.Selected.Value
    JSNGBN = oForm.Items("JSNGBN").Specific.Selected.Value
    
    If FYEAR = "" Or TYEAR = "" Then        '// ����⵵ üũ(�Է� ���� ��� ����)
        ErrNum = 1
        GoTo Error_Message
    End If
    If JIGFDAT = "" Or JIGTDAT = "" Then    '// �������� üũ(�Է� ���� ��� ��ü��ȸ)
        JIGFDAT = "1900-01-01"
        JIGTDAT = "2999-12-31"
    End If
    If SINFYMM = "" Or SINTYMM = "" Then    '// �Ű���� üũ(�Է� ���� ��� ��ü��ȸ)
        SINFYMM = "190001"
        SINTYMM = "299912"
    End If

    '/ ��ȸ
    sQry = "Exec ZPY507 N'" & FYEAR & "', N'" & TYEAR & "', " & _
                       "N'" & JIGFDAT & "', N'" & JIGTDAT & "', " & _
                       "N'" & CLTCOD & "', N'" & MSTCOD & "', " & _
                       "N'" & MSTNAM & "', N'" & JSNGBN & "', " & _
                       "N'" & SINFYMM & "', N'" & SINTYMM & "', " & _
                        CStr(PILMED) & ", " & CStr(PILGBU)
    oDS_ZPY507.ExecuteQuery (sQry)
    iRow = oDS_ZPY507.Rows.Count
    If iRow = 1 Then
        oRecordSet.DoQuery sQry
        iRow = oRecordSet.RecordCount
    End If

    If iRow > 0 Then
        Sbo_Application.StatusBar.SetText iRow & " ���� �ֽ��ϴ�.", bmt_Short, smt_Success
    Else
        Sbo_Application.StatusBar.SetText "��ȸ�� �����Ͱ� �����ϴ�.", bmt_Short, smt_Success
    End If
    Set oRecordSet = Nothing
    
    Call titleSetting
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "���س⵵�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Success
    Else
        Sbo_Application.StatusBar.SetText "Grid_Display Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
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

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY507.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)

    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//�������� ��Ʈ������ Ʋ��쿡 ������ó�� �ε� �ǵ��� ���� ���
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY507_" & GetTotalFormsCount
    
    '�� �Ҵ�
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '�÷��ǿ� ���� ��´�   **�÷����̶� ��ü�� ��� ���� �迭�μ� ���⼭�� Ȱ��ȭ�Ǿ��� �ִ� ���� ��� �ִ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    AddForms Me, oFormUniqueID, "ZPY507"
    oForm.SupportedModes = -1

    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    'oForm.DataBrowser.BrowseBy = "Code"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    Call CreateItems

    oForm.EnableMenu ("1281"), False    '/ �߰�
    oForm.EnableMenu ("1282"), False    '/ �߰�

    oForm.Mode = fm_OK_MODE
    oForm.Freeze False
    oForm.Update
    oForm.Visible = True

    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
    Sbo_Application.StatusBar.SetText "Form_Load Error:" & Err.Description, bmt_Short, smt_Error
     If (oForm Is Nothing) = False Then
        oForm.Freeze False
        Set oForm = Nothing
    End If
End Sub

Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oEdit           As SAPbouiCOM.EditText
    Dim sQry            As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '/UserDataSource ����
    With oForm.DataSources.UserDataSources
        .Add "FYEAR", dt_SHORT_TEXT, 4
        .Add "TYEAR", dt_SHORT_TEXT, 4
        .Add "JIGFDAT", dt_DATE
        .Add "JIGTDAT", dt_DATE
        .Add "MSTCOD", dt_SHORT_TEXT, 8
        .Add "MSTNAM", dt_LONG_TEXT, 200
        .Add "SINFYMM", dt_SHORT_TEXT, 6
        .Add "SINTYMM", dt_SHORT_TEXT, 6
        .Add "PILMED", dt_SUM
        .Add "PILGBU", dt_SUM
    End With
    
    Set oEdit = oForm.Items("FYEAR").Specific   '// ���س⵵(From)
    oEdit.DataBind.SetBound True, "", "FYEAR"
    
    Set oEdit = oForm.Items("TYEAR").Specific   '// ���س⵵(To)
    oEdit.DataBind.SetBound True, "", "TYEAR"
    
    Set oEdit = oForm.Items("JIGFDAT").Specific '// ��������(From)
    oEdit.DataBind.SetBound True, "", "JIGFDAT"
    
    Set oEdit = oForm.Items("JIGTDAT").Specific '// ��������(To)
    oEdit.DataBind.SetBound True, "", "JIGTDAT"
    
    Set oEdit = oForm.Items("MSTCOD").Specific  '// ���
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    
    Set oEdit = oForm.Items("MSTNAM").Specific  '// ����
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    
    Set oEdit = oForm.Items("SINFYMM").Specific '// �Ű����(From)
    oEdit.DataBind.SetBound True, "", "SINFYMM"
    
    Set oEdit = oForm.Items("SINTYMM").Specific '// �Ű����(To)
    oEdit.DataBind.SetBound True, "", "SINTYMM"
    
    Set oEdit = oForm.Items("PILMED").Specific  '// �Ƿ�������
    oEdit.DataBind.SetBound True, "", "PILMED"
    
    Set oEdit = oForm.Items("PILGBU").Specific  '// ��αݰ�����
    oEdit.DataBind.SetBound True, "", "PILGBU"
    

    '//�����
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "��ü"
    oForm.Items("CLTCOD").DisplayDesc = True
    oCombo.Select "%", psk_ByValue

    '// ���걸��
    Set oCombo = oForm.Items("JSNGBN").Specific
    oCombo.ValidValues.Add "%", "���"
    oCombo.ValidValues.Add "1", "��������(������)"
    oCombo.ValidValues.Add "2", "�ߵ�����(������)"
    oCombo.Select 0, psk_Index
    

    oForm.DataSources.UserDataSources("FYEAR").ValueEx = ZPAY_GBL_JSNYER
    oForm.DataSources.UserDataSources("TYEAR").ValueEx = ZPAY_GBL_JSNYER
    
'// ������� �ҽ� ��ü �Ҵ�
    Set oGrid1 = oForm.Items("Grid1").Specific
    Set oDS_ZPY507 = oForm.DataSources.DataTables.Add("ZPY507")
    oDS_ZPY507.ExecuteQuery ("Exec ZPY507 '1900', '1900', NULL, NULL, '%', '%', '%', '%', '', '', 0, 0")
    oGrid1.DataTable = oDS_ZPY507

    Call titleSetting
    
    oForm.ActiveItem = "FYEAR"
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    
    Sbo_Application.StatusBar.SetText "CreateItems ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub