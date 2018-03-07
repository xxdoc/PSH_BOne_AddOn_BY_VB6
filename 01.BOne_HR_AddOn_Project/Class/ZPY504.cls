VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY504"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY504.cls
'//  Module         : �λ����>�������
'//  Desc           : ��������ȸ
'//  FormType       : 2010110504
'//  Create Date    : 2005.01.20
'//  Modified Date  :
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oDS_ZPY504H     As SAPbouiCOM.DBDataSource    '�ý����ڵ� ���

Private Last_Item       As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oOLDCHK         As String

Private Sub FormItemEnabled()
    If oForm.Mode = fm_FIND_MODE Or oForm.Mode = fm_ADD_MODE Then
        oForm.Items("DocNum").Enabled = True
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = True
        oForm.Items("JSNGBN").Enabled = True
        oForm.Items("ENDCHK").Enabled = True
    ElseIf oForm.Mode = fm_OK_MODE Then
        oForm.Items("DocNum").Enabled = False
        oForm.Items("JSNYER").Enabled = False
        oForm.Items("MSTCOD").Enabled = False
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("JSNGBN").Enabled = False
        '// �޿��������Ȱ��� ��Ȱ��ȭ
        oOLDCHK = oDS_ZPY504H.GetValue("U_ENDCHK", 0)
        If MDC_SetMod.Get_ReData("U_ENDCHK", "U_JOBYER", "[@ZPY509L]", "'" & oDS_ZPY504H.GetValue("U_JSNYER", 0) & "'", " AND Code = '" & oDS_ZPY504H.GetValue("U_CLTCOD", 0) & "'") = "Y" Then
            oForm.Items("ENDCHK").Enabled = False
        Else
            oForm.Items("ENDCHK").Enabled = True
        End If
        
    End If
End Sub

Private Sub UserFormLink(LinkName$)
    Dim oTmpObject  As Variant

    If oForm.Items("CLTCOD").Specific.Selected Is Nothing Then Exit Sub

    If oForm.Mode <> fm_FIND_MODE Then
        If LinkName = "LBtn1" Then
'            DOCNUM = MDC_SetMod.Get_ReData("DocNum", "U_JSNYER", "[@ZPY501H]", _
'                    "'" & oForm.Items("JSNYER").Specific.String & "'", _
'                    " AND U_MSTCOD = '" & oForm.Items("MSTCOD").Specific.String & "'")
'            Call MDC_SetMod.UserLink(DOCNUM, "ZPY501", "D")
            Set oTmpObject = New ZPY501     'ZPY501(2000060501) :
            Call oTmpObject.LoadForm(oForm.Items("JSNYER").Specific.String, oForm.Items("MSTCOD").Specific.String, oForm.Items("CLTCOD").Specific.Selected.Value)
            Sbo_Application.Forms.ActiveForm.Select
            Call Get_FormColor
        ElseIf LinkName = "LBtn2" Then
            Set oTmpObject = New ZPY505     'ZPY505(2000060505) :
            Call oTmpObject.LoadForm(oForm.Items("JSNYER").Specific.String, oForm.Items("MSTCOD").Specific.String, oForm.Items("CLTCOD").Specific.Selected.Value)
            Sbo_Application.Forms.ActiveForm.Select
            Call Get_FormColor
        ElseIf LinkName = "LBtn3" Then
            Set oTmpObject = New ZPY506     'ZPY506(2000060506) :
            Call oTmpObject.LoadForm(oForm.Items("JSNYER").Specific.String, oForm.Items("MSTCOD").Specific.String, oForm.Items("CLTCOD").Specific.Selected.Value)
            Sbo_Application.Forms.ActiveForm.Select
            Call Get_FormColor
        End If
    End If
    Set oTmpObject = Nothing
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
                If pval.ItemUID = "CBtn1" Then
                    If oForm.Items("MSTCOD").Enabled = True Then
                        oForm.Items("MSTCOD").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                ElseIf pval.ItemUID = "LBtn1" Or pval.ItemUID = "LBtn2" Or pval.ItemUID = "LBtn3" Then  '// Link Button
                    Call UserFormLink(pval.ItemUID)
                ElseIf pval.ItemUID = "1" And (oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE) Then
                    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                    '//�߰� �� ������Ʈ�ÿ�
                    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            Else
                If pval.ItemUID = "1" And oForm.Mode = fm_OK_MODE Then
                    Call FormItemEnabled
                ElseIf pval.ItemUID = "Folder1" Then
                    oForm.PaneLevel = 1
                ElseIf pval.ItemUID = "Folder2" Then
                    oForm.PaneLevel = 2
                End If
            End If
            
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True And _
            (pval.ItemUID = "MSTCOD" Or pval.ItemUID = "JSNYER" Or pval.ItemUID = "JSNMON") Then
                FlushToItemValue pval.ItemUID
            End If
            
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
            If pval.BeforeAction = True And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" And oForm.Mode <> fm_FIND_MODE Then
                If Last_Item = "MSTCOD" Then
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(Last_Item).Specific.String & "'", "") = True _
                    And oForm.Items(Last_Item).Specific.String <> "" And Last_Item <> pval.ItemUID Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
            
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
             If pval.BeforeAction = True And pval.ItemUID = "JSNYER" And pval.CharPressed = 9 Then
                If Len(Trim$(oForm.Items("JSNYER").Specific.String)) = 0 Then
                    Sbo_Application.StatusBar.SetText "�ͼӳ⵵�� �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 Then
                If Trim$(oForm.Items("MSTCOD").Specific.String) = "" Then
                    Sbo_Application.StatusBar.SetText "�����ȣ�� �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                    BubbleEvent = False
                Else
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(pval.ItemUID).Specific.String & "'", "") = True Then
                        oForm.Items("MSTCOD").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUID
                End If
            Else
                Last_Item = pval.ItemUID
            End If
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            '�÷��ǿ��� ������ ��� �޸� ����
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            If pval.BeforeAction = False Then
               RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_ZPY504H = Nothing
            End If
'et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_MATRIX_LOAD:
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:", bmt_Short, smt_Error
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim DocNum          As String
    
    ErrNum = 0
 '/ Check
    Select Case True
      Case Trim$(oDS_ZPY504H.GetValue("U_ENDCHK", 0)) = "Y" And Trim$(oOLDCHK) = "Y"
           ErrNum = 1
           GoTo Error_Message
      Case oDS_ZPY504H.GetValue("U_MSTCOD", 0) = ""
           ErrNum = 2
           GoTo Error_Message
      Case oForm.Items("JSNGBN").Specific.Selected Is Nothing
           ErrNum = 3
           GoTo Error_Message
      Case Trim$(oForm.Items("JSNMON").Specific.Value) = ""
           ErrNum = 4
           GoTo Error_Message
      Case Trim$(oForm.Items("SINYMM").Specific.Value) = ""
           ErrNum = 5
           GoTo Error_Message
      Case oForm.Items("CLTCOD").Specific.Selected Is Nothing
           ErrNum = 6
           GoTo Error_Message
      Case Trim$(oForm.Items("JSNYER").Specific.Value) = ""
           ErrNum = 7
           GoTo Error_Message
    End Select
    
    DocNum = Exist_YN(oDS_ZPY504H.GetValue("U_JSNYER", 0), oDS_ZPY504H.GetValue("U_MSTCOD", 0), oDS_ZPY504H.GetValue("U_CLTCOD", 0))
    If Trim$(DocNum) <> "" And Trim$(oDS_ZPY504H.GetValue("DocNum", 0)) <> Trim$(DocNum) Then
        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        '���������Ͱ� �����ϴµ� �ڱ� �ڽ��� ���� �ڱ��ڽ��� �ƴ϶��(���������� ����Ѱ� �ƴϸ� �ϳ��� �����ؾ���)
        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        ErrNum = 12
        GoTo Error_Message
    End If
    
   
    HeaderSpaceLineDel = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "��� �ڷ��Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "�����ȣ�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "���� ������ �ʼ��Դϴ�. �����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "�ͼ� ���� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "���� ������ �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "�ڻ� �ڵ�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "���� �⵵�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 8 Then
        Sbo_Application.StatusBar.SetText "������ȣ" & DocNum & " ��(��) �����Ͱ� ��ġ�մϴ�. ������� �ʽ��ϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    HeaderSpaceLineDel = False
End Function


Private Function Exist_YN(JOBYER$, MSTCOD$, CLTCOD$) As String
'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'������ �������� ���������Ͱ� �ִ��� Ȯ���Ѵ�
'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT Top 1 T1.DocNum FROM [@ZPY504H] T1 "
    sQry = sQry & " WHERE T1.U_JSNYER = N'" & Trim$(JOBYER) & "'"
    sQry = sQry & "   AND T1.U_MSTCOD = N'" & Trim$(MSTCOD) & "'"
    sQry = sQry & "   AND T1.U_CLTCOD = N'" & Trim$(CLTCOD) & "'"
    oRecordSet.DoQuery sQry
    
    Do Until oRecordSet.EOF
        Exist_YN = oRecordSet(0).Value
        oRecordSet.MoveNext
    Loop
    
    If Trim$(Exist_YN) = "" Then
        Exist_YN = ""
        Exit Function
    End If
    
    Set oRecordSet = Nothing
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim oMast           As ZPAY_g_EmpID
    
    Select Case oUID
      Case "JSNYER"
            If Trim$(oForm.Items(oUID).Specific.String) <> "" Then
                ZPAY_GBL_JSNYER = oForm.Items(oUID).Specific.String
            Else
                oDS_ZPY504H.setValue "U_JSNYER", 0, ZPAY_GBL_JSNYER
            End If
            oForm.Items(oUID).Update
      Case "JSNMON"
            If Trim$(oForm.Items(oUID).Specific.Value) <> "" Then
                oDS_ZPY504H.setValue "U_JSNMON", 0, Format$(oForm.Items(oUID).Specific.Value, "00")
            End If
            oForm.Items(oUID).Update
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oDS_ZPY504H.setValue "U_MSTCOD", 0, ""
                oDS_ZPY504H.setValue "U_MSTNAM", 0, ""
                oDS_ZPY504H.setValue "U_EmpID", 0, ""
            Else
                oDS_ZPY504H.setValue "U_MSTCOD", 0, UCase$(oForm.Items(oUID).Specific.String)
                oMast = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oDS_ZPY504H.setValue "U_MSTNAM", 0, oMast.MSTNAM
                oDS_ZPY504H.setValue "U_EmpID", 0, oMast.EmpID
            End If

            oForm.Items("MSTNAM").Update
            oForm.Items("EmpID").Update
            oForm.Items(oUID).Update
    End Select
 
End Sub
'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    
    If pval.BeforeAction = True Then
        Select Case pval.MenuUID
            Case "1283" '/ ����
                If Trim$(oDS_ZPY504H.GetValue("U_ENDCHK", 0)) = "Y" Then
                    Sbo_Application.StatusBar.SetText "��� �ڷ��Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                    BubbleEvent = False
                    Exit Sub
                Else
                    If Sbo_Application.MessageBox("���� ȭ�鳻����ü�� ���� �Ͻðڽ��ϱ�? ������ �� �����ϴ�.", 2, "Yes", "No") = 2 Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            Case Else
                Exit Sub
        End Select
    Else
    
        Select Case pval.MenuUID
            Case "1287" '/ ����
            Case "1283" '/ ����
                 FormItemEnabled
            Case "1281", "1282" '/ ã��, �߰�
                FormItemEnabled
                oDS_ZPY504H.setValue "U_JSNYER", 0, ZPAY_GBL_JSNYER
                oForm.Items("JSNYER").Update
                oForm.Items("JSNYER").CLICK ct_Regular
            Case "1288" To "1291"
                FormItemEnabled
            Case "1293"
        End Select
    End If
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
Public Sub LoadForm(Optional DocNum As String)
'Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY504.srf")
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
    
    oFormUniqueID = "ZPY504_" & GetTotalFormsCount
    
    '�� �Ҵ�
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '�÷��ǿ� ���� ��´�   **�÷����̶� ��ü�� ��� ���� �迭�μ� ���⼭�� Ȱ��ȭ�Ǿ��� �ִ� ���� ��� �ִ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    AddForms Me, oFormUniqueID, "ZPY504"
    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    
    Call CreateItems
    
    oForm.EnableMenu ("1281"), True  '/ ã��
    oForm.EnableMenu ("1282"), False '/ �߰�
    oForm.EnableMenu ("1284"), False '/ ���
    oForm.EnableMenu ("1293"), False '/ �����
    
    If DocNum <> "" Then
        Call ShowSource(DocNum)
    End If
    
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

Private Sub ShowSource(DocNum$)
    oForm.Items("DocNum").Enabled = True
    oForm.Items("DocNum").Specific.Value = DocNum
    
    oForm.Items("1").CLICK ct_Regular
End Sub

'*******************************************************************
'
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oFolder     As SAPbouiCOM.Folder
    Dim iCol        As Integer
    
    oForm.PaneLevel = 2
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oFolder = oForm.Items("Folder1").Specific
    
    '// DB DataSource ��ü �Ҵ�
    Set oDS_ZPY504H = oForm.DataSources.DBDataSources("@ZPY504H")
    
    '// Folder
    oForm.DataSources.UserDataSources.Add "FolderD", dt_SHORT_TEXT, 1
    For iCol = 1 To 2
        Set oFolder = oForm.Items("Folder" & iCol).Specific
        oFolder.DataBind.SetBound True, "", "FolderD"
        If iCol = 2 Then
            oFolder.Select
            oFolder.GroupWith ("Folder" & iCol - 1)
            
        End If
        oForm.Items("Folder" & iCol).Enabled = True
        oForm.Items("Folder" & iCol).AffectsFormMode = False
    Next iCol
    
    '// ���걸��
    Set oCombo = oForm.Items("JSNGBN").Specific
    oCombo.ValidValues.Add "1", "��������(������)"
    oCombo.ValidValues.Add "2", "�ߵ�����(������)"
    
    
    '// �����
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
    Loop
    
  '/ Check ��ư
    Set oCheck = oForm.Items("ENDCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    
    Set oCheck = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oCheck = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error:" & Err.Description, bmt_Short, smt_Error
End Sub

