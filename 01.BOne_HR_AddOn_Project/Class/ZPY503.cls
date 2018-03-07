VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY503"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY503.cls
'//  Module         : �λ����>�������
'//  Desc           : ���� ���� ���
'//  FormType       : 2000060503
'//  Create Date    : 2006.01.20
'//  Modified Date  :
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form

Private oJsnYear        As String
Private oSMonth         As String
Private oEMonth         As String

Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     'Ŭ�������� ������ ������ ������ Uid��

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm()
'Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY503.srf")
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
    
    oFormUniqueID = "ZPY503_" & GetTotalFormsCount
    
    '�� �Ҵ�
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '�÷��ǿ� ���� ��´�   **�÷����̶� ��ü�� ��� ���� �迭�μ� ���⼭�� Ȱ��ȭ�Ǿ��� �ִ� ���� ��� �ִ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    AddForms Me, oFormUniqueID, "ZPY503"
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    'oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    CreateItems
    oForm.Freeze False
    
    oForm.EnableMenu ("1281"), False  '/ ã��
    oForm.EnableMenu ("1282"), True  '/ �߰�
    oForm.EnableMenu ("1284"), False '/ ���
    oForm.EnableMenu ("1293"), False '/ �����
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
                If pval.ItemUID = "CBtn1" Then   '/ ChooseBtn�������Ʈ
                    If oForm.Items("MSTCOD").Enabled = True Then
                        oForm.Items("MSTCOD").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                ElseIf pval.ItemUID = "1" And oForm.Mode = fm_ADD_MODE Then
                    If Execution = False Then
                        BubbleEvent = False
                    Else
                        BubbleEvent = False
                        oForm.Mode = fm_OK_MODE
                    End If
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True And _
              (pval.ItemUID = "JSNYER" Or pval.ItemUID = "SMonth" Or pval.ItemUID = "EMonth" Or _
              pval.ItemUID = "MSTCOD" Or pval.ItemUID = "JSNMON" Or pval.ItemUID = "JIGDAT") Then
               FlushToItemValue pval.ItemUID
            End If
'et_COMBO_SELECT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_COMBO_SELECT
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUID = "JSNGBN" Or pval.ItemUID = "JSNMON" Then
                    FlushToItemValue pval.ItemUID
                End If
                If pval.ItemUID = "CLTCOD" Then
                    '//�⺻���� - �μ�1 (����忡 ���� �μ�����)
                    Set oCombo = oForm.Items("DPTSTR").Specific

                    If oCombo.ValidValues.Count > 0 Then
                        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                            oCombo.ValidValues.Remove i, psk_Index
                        Next i
                        oCombo.ValidValues.Add "%", "��ü"
                        oCombo.Select 0, psk_Index
                    End If
                    
                    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                    sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                    sQry = sQry & " ORDER BY U_Code"
                    Call SetReDataCombo(oForm, sQry, oCombo)
                    oCombo.ValidValues.Add "%", "��ü"
                    oCombo.Select "%", psk_ByValue
                    
                    '//�⺻���� - �μ�2 (����忡 ���� �μ�����)
                    Set oCombo = oForm.Items("DPTEND").Specific

                    If oCombo.ValidValues.Count > 0 Then
                        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                            oCombo.ValidValues.Remove i, psk_Index
                        Next i
                        oCombo.ValidValues.Add "%", "��ü"
                        oCombo.Select 0, psk_Index
                    End If
                    
                    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                    sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                    sQry = sQry & " ORDER BY U_Code"
                    Call SetReDataCombo(oForm, sQry, oCombo)
                    oCombo.ValidValues.Add "%", "��ü"
                    oCombo.Select "%", psk_ByValue
                    
                End If
            End If
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
             If pval.BeforeAction = True And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" Then
                If Last_Item = "JSNYER" Then     '/����⵵
                    If Trim$(oForm.Items(Last_Item).Specific.VALUE) <> "" Then
                        If MDC_SetMod.ChkYearMonth(Trim$(CStr((oForm.Items(Last_Item).Specific.VALUE))) & "01") = False Then
                            oForm.Items(Last_Item).Update
                            Sbo_Application.StatusBar.SetText "����⵵�� Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                            BubbleEvent = False
                        End If
                    End If
                ElseIf Last_Item = "SMonth" Or Last_Item = "EMonth" Then
'                    If Trim$(oForm.Items(Last_Item).Specific.Value) <> "" Then
'                        If MDC_SetMod.ChkYearMonth(oJsnYear & Format$(oForm.Items(Last_Item).Specific.Value, "00")) = False Then
'                            oForm.Items(Last_Item).Update
'                            Sbo_Application.StatusBar.SetText "�����Ⱓ�� Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
'                            BubbleEvent = False
'                        End If
'                    End If
                ElseIf Last_Item = "MSTCOD" Then
                    If Trim$(oForm.Items(Last_Item).Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items(Last_Item).Specific.String) & "'", "") = True Then
                        oForm.Items(Last_Item).Update
                        Sbo_Application.StatusBar.SetText "�����ȣ�� Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                        BubbleEvent = False
                    End If
                End If
            End If
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
             If pval.BeforeAction = True And pval.ItemUID = "JSNYER" And pval.CharPressed = 9 Then
                If Len(Trim$(oForm.Items(pval.ItemUID).Specific.String)) < 4 Then
                    oForm.Items(pval.ItemUID).Specific.VALUE = Format$(oForm.Items(pval.ItemUID).Specific.VALUE, "2000")
                End If
                If MDC_SetMod.ChkYearMonth(Trim$(CStr((oForm.Items(pval.ItemUID).Specific.VALUE))) & "01") = False Then
                    Sbo_Application.StatusBar.SetText "����⵵�� Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 Then
                If Trim$(oForm.Items("MSTCOD").Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'", "") = True Then
                    Sbo_Application.StatusBar.SetText "�����ȣ�� Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                    BubbleEvent = False
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
                Set oMat1 = Nothing
            End If
    End Select
    
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
            oForm.Items("JSNYER").CLICK ct_Regular
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

Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo1         As SAPbouiCOM.ComboBox
    Dim oCombo2         As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oEdit           As SAPbouiCOM.EditText
    Dim oColumn         As SAPbouiCOM.Column
    Dim sQry            As String
    Dim i               As Integer
    Dim STDMON          As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oMat1 = oForm.Items("Mat1").Specific
        
    Call oForm.DataSources.UserDataSources.Add("JSNYER", dt_SHORT_TEXT, 4)      '/ �ͼӳ⵵
    Call oForm.DataSources.UserDataSources.Add("JSNMON", dt_SHORT_TEXT, 2)      '/ �ͼӿ�
    Call oForm.DataSources.UserDataSources.Add("JSNGBN", dt_SHORT_TEXT, 10)     '/ ����
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)      '/ �����
    Call oForm.DataSources.UserDataSources.Add("DPTSTR", dt_SHORT_TEXT, 10)     '/ �μ��ڵ�
    Call oForm.DataSources.UserDataSources.Add("DPTEND", dt_SHORT_TEXT, 10)
    Call oForm.DataSources.UserDataSources.Add("SMonth", dt_DATE)     '/ ������
    Call oForm.DataSources.UserDataSources.Add("EMonth", dt_DATE)      '/ ������
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 8)
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 30)
    Call oForm.DataSources.UserDataSources.Add("SINYMM", dt_SHORT_TEXT, 6)    '/ �Ű�����
    Call oForm.DataSources.UserDataSources.Add("JIGDAT", dt_DATE)      '/ ��������
    
    Set oEdit = oForm.Items("JSNYER").Specific
    oEdit.DataBind.SetBound True, "", "JSNYER"
    Set oCombo1 = oForm.Items("JSNMON").Specific
    oCombo1.DataBind.SetBound True, "", "JSNMON"
    Set oEdit = oForm.Items("SMonth").Specific
    oEdit.DataBind.SetBound True, "", "SMonth"
    Set oEdit = oForm.Items("EMonth").Specific
    oEdit.DataBind.SetBound True, "", "EMonth"
    Set oEdit = oForm.Items("SINYMM").Specific
    oEdit.DataBind.SetBound True, "", "SINYMM"
    Set oEdit = oForm.Items("JIGDAT").Specific
    oEdit.DataBind.SetBound True, "", "JIGDAT"
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    Set oCombo1 = oForm.Items("JSNGBN").Specific
    oCombo1.DataBind.SetBound True, "", "JSNGBN"
    
    Call oForm.DataSources.UserDataSources.Add("Col0", dt_SHORT_TEXT)
    Call oForm.DataSources.UserDataSources.Add("Col1", dt_SHORT_TEXT)
    
    Set oColumn = oMat1.Columns("Col0")
    oColumn.DataBind.SetBound True, "", "Col0"
    
    Set oColumn = oMat1.Columns("Col1")
    oColumn.DataBind.SetBound True, "", "Col1"
 
    '// ���걸��
    Set oCombo1 = oForm.Items("JSNGBN").Specific
    oCombo1.ValidValues.Add "1", "��������(������)"
    oCombo1.ValidValues.Add "2", "�ߵ�����(������)"
'    sQry = " SELECT U_Minor, U_CdName FROM [@ZPY001L] WHERE Code='P192' ORDER BY U_Minor "
'    oRecordSet.DoQuery sQry
'    Do Until oRecordSet.EOF
'        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
'        oRecordSet.MoveNext
'    Loop
    '// �ͼӿ���
    Set oCombo1 = oForm.Items("JSNMON").Specific
    For i = 1 To 12
        STDMON = Format$(i, "00")
        oCombo1.ValidValues.Add STDMON, STDMON
    Next i
    '// �����
    Set oCombo1 = oForm.Items("CLTCOD").Specific
    oCombo1.DataBind.SetBound True, "", "CLTCOD"
    sQry = "SELECT Code,Name FROM [@PH_PY005A] "
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    
    

    '// �μ�
    Set oCombo1 = oForm.Items("DPTSTR").Specific
    oCombo1.DataBind.SetBound True, "", "DPTSTR"
    oForm.Items("DPTSTR").DisplayDesc = True
    '// �μ�
    Set oCombo1 = oForm.Items("DPTEND").Specific
    oCombo1.DataBind.SetBound True, "", "DPTEND"
    oForm.Items("DPTEND").DisplayDesc = True
    
    oForm.DataSources.UserDataSources("JSNMON").ValueEx = Format$(Now, "MM")
    oForm.ActiveItem = "JSNYER"
    
    Set oEdit = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Function Execution() As Boolean
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim TOTCNT&, MSTCNT&
    Dim DPTSTR          As String * 8
    Dim DPTEND          As String * 8
    Dim MSTCOD          As String * 8
    Dim STRDAT          As String * 10
    Dim ENDDAT          As String * 10
    Dim CLTCOD          As String
    Dim BPLID           As String
    Dim JSNGBN          As String
    
    ErrNum = 0
    '/ �ʼ�Check /
    If Trim$(oForm.Items("JSNYER").Specific.String) = "" Then     '/ ����⵵
        ErrNum = 1
        GoTo Error_Message
    ElseIf oForm.Items("JSNGBN").Specific.Selected Is Nothing Then
        ErrNum = 7
        GoTo Error_Message
    ElseIf Trim$(oForm.Items("SMonth").Specific.VALUE) = "" Or _
           Trim$(oForm.Items("EMonth").Specific.VALUE) = "" Then
        ErrNum = 2
        GoTo Error_Message
    ElseIf oForm.Items("SMonth").Specific.VALUE > oForm.Items("EMonth").Specific.VALUE Then
        ErrNum = 3
        GoTo Error_Message
    ElseIf oForm.Items("CLTCOD").Specific.Selected Is Nothing Then
        ErrNum = 8
        GoTo Error_Message
'    ElseIf oForm.Items("BPLId").Specific.Selected Is Nothing Then
'        ErrNum = 4
'        GoTo Error_Message
    ElseIf oForm.Items("DPTSTR").Specific.Selected Is Nothing Or _
        oForm.Items("DPTEND").Specific.Selected Is Nothing Then
        ErrNum = 5
        GoTo Error_Message
    ElseIf Len(oForm.Items("SINYMM").Specific.VALUE) <> 6 Then
        ErrNum = 9
        GoTo Error_Message
    ElseIf oForm.Items("JSNMON").Specific.Selected Is Nothing Then
        ErrNum = 10
        GoTo Error_Message
    ElseIf Len(oForm.Items("JIGDAT").Specific.VALUE) = 0 Then
        ErrNum = 11
        GoTo Error_Message
    End If
    DPTSTR = oForm.Items("DPTSTR").Specific.Selected.VALUE
    DPTEND = oForm.Items("DPTEND").Specific.Selected.VALUE
    MSTCOD = oForm.Items("MSTCOD").Specific.String
    STRDAT = Format$(oForm.Items("SMonth").Specific.VALUE, "0000-00-00")
    ENDDAT = Format$(oForm.Items("EMonth").Specific.VALUE, "0000-00-00")
    CLTCOD = oForm.Items("CLTCOD").Specific.Selected.VALUE
'    BPLID = oForm.Items("BPLId").Specific.Selected.VALUE
    If Trim$(DPTSTR) = "-1" Then DPTSTR = "00000001"
    If Trim$(DPTEND) = "-1" Then DPTEND = "ZZZZZZZZ"
    If Trim$(MSTCOD) = "" Then MSTCOD = "%"
    JSNGBN = oForm.Items("JSNGBN").Specific.Selected.VALUE
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'/ �ش�⵵ ���긶������ Ȯ�ο���
    If MDC_SetMod.Get_ReData("U_ENDCHK", "U_JOBYER", "[@ZPY509L]", "'" & Trim$(oJsnYear) & "'", " AND Code = '" & Trim$(CLTCOD) & "'") = "Y" Then
        ErrNum = 13
        GoTo Error_Message
    End If
    
    '/ ���꼼�װ�� ����� ��ȸ
    sQry = " EXEC ZPY503_1 '" & Trim$(oJsnYear) & "', '" & Trim$(JSNGBN) & "', '" & STRDAT & "', '" & ENDDAT & "','" & Trim$(CLTCOD) & "', '" & Trim$(DPTSTR) & "', '" & Trim$(DPTEND) & "','" & Trim$(MSTCOD) & "' "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 6
        GoTo Error_Message
    End If
   '//
    oMat1.Clear
    TOTCNT = 0: MSTCNT = 0
    Do Until oRecordSet.EOF
        TOTCNT = TOTCNT + 1
        MSTCOD = oRecordSet.Fields(0).VALUE
        
    '/ �ش�⵵ ���긶������ Ȯ�ο���
        If MDC_SetMod.Get_ReData("U_ENDCHK", "U_JSNYER", "[@ZPY504H]", "'" & Trim$(oJsnYear) & "'", " AND U_MSTCOD = '" & Trim$(MSTCOD) & "' AND U_CLTCOD = '" & Trim$(CLTCOD) & "'") = "Y" Then
           oForm.DataSources.UserDataSources("Col0").VALUE = TOTCNT
           oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & oRecordSet.Fields("MSTNAM").VALUE & " ���װ�� ����! ����ڷ�************"
           oMat1.AddRow
        Else
            '/ �����ڷ������ ����������Ȯ��
             If MDC_SetMod.Value_ChkYn("[@ZPY343H]", "U_JsnYear", "'" & Trim$(oJsnYear) & "'", " AND  U_MstCode = '" & Trim$(MSTCOD) & "' AND U_CLTCOD = '" & Trim$(CLTCOD) & "'") = True Then
                    oForm.DataSources.UserDataSources("Col0").VALUE = TOTCNT
                    oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & oRecordSet.Fields("MSTNAM").VALUE & " ���װ�� ����! �����ڷ���� ����"
                    oMat1.AddRow
            '/ �ҵ�����׸��Ͽ� ����������Ȯ��
            ElseIf MDC_SetMod.Value_ChkYn("[@ZPY501H]", "U_JSNYER", "'" & Trim$(oJsnYear) & "'", " AND  U_MSTCOD = '" & Trim$(MSTCOD) & "' AND U_CLTCOD = '" & Trim$(CLTCOD) & "'") = True Then
                    oForm.DataSources.UserDataSources("Col0").VALUE = TOTCNT
                    oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & oRecordSet.Fields("MSTNAM").VALUE & " ���װ�� ����! �ҵ�����׸��� ����"
                    oMat1.AddRow
            '/ �޿��⺻��� ����������Ȯ�� (2010.03.03 �ֵ��� �߰�)
            ElseIf MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(MSTCOD) & "'") = True Then
                    oForm.DataSources.UserDataSources("Col0").VALUE = TOTCNT
                    oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & oRecordSet.Fields("MSTNAM").VALUE & " ���װ�� ����! �޿��⺻��� ����"
                    oMat1.AddRow
            Else
                '/ ���꼼�װ������
                If Execution_Save(oJsnYear, oRecordSet.Fields("CLTCOD").VALUE, MSTCOD) = True Then
                    MSTCNT = MSTCNT + 1
                    oForm.DataSources.UserDataSources("Col0").VALUE = TOTCNT
                    oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & oRecordSet.Fields("MSTNAM").VALUE & " ���װ�� �Ϸ�."
                    oMat1.AddRow
                Else
                    oForm.DataSources.UserDataSources("Col0").VALUE = TOTCNT
                    oForm.DataSources.UserDataSources("Col1").VALUE = MSTCOD & ": " & oRecordSet.Fields("MSTNAM").VALUE & " ���װ�� ����! **************"
                    oMat1.AddRow
                End If
             End If
        
        End If
        oRecordSet.MoveNext
    Loop
  '/
    Sbo_Application.StatusBar.SetText "( " & MSTCNT & "/" & TOTCNT & " )�� �۾��� �Ϸ��Ͽ����ϴ�.", bmt_Short, smt_Success
    Set oRecordSet = Nothing
    Execution = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "����⵵�� �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "�����Ⱓ�� �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "���ۿ����� ������� �۽��ϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "������ �����ϼ���. �����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "������ �μ������� �����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "���װ���� ����ڷᰡ �����ϴ�. ���ڷ� ������ ���� �Ͻʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "���� ������ �ʼ��Դϴ�. �����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 8 Then
        Sbo_Application.StatusBar.SetText "�ڻ� �ڵ�� �ʼ��Դϴ�. �����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 9 Then
        Sbo_Application.StatusBar.SetText "�Ű� ������ �ʼ��Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 10 Then
        Sbo_Application.StatusBar.SetText "�ͼ� ���� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 11 Then
        Sbo_Application.StatusBar.SetText "���� ���ڴ� �ʼ��Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 13 Then
        Sbo_Application.StatusBar.SetText "��� �ڷ��Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Execution ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    Execution = False
End Function
Private Function Execution_Save(JSNYER$, CLTCOD$, MSTCOD$) As Boolean
On Error GoTo Error_Message
    Dim sRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Dim JSNGBN          As String
    Dim JSNMON          As String
    Dim SINYMM          As String
    Dim JIGDAT          As String
    
    '// Default
    JSNGBN = oForm.Items("JSNGBN").Specific.Selected.VALUE
    JSNMON = oForm.Items("JSNMON").Specific.Selected.VALUE
    SINYMM = oForm.Items("SINYMM").Specific.VALUE
    JIGDAT = oForm.Items("JIGDAT").Specific.VALUE
    JIGDAT = Format$(JIGDAT, "0000-00-00")
    JSNMON = Format$(JSNMON, "00")
    
    '// ����� ���� ���װ�����
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    'Exec dbo.MDC_ZPY503_05  '2005', '9603004', '1','12','2008-01-01'
    sQry = "Exec ZPY503_" & Mid$(JSNYER, 3, 2) & Space$(1) & "'" & JSNYER & "', '" & Trim$(CLTCOD) & "','" & Trim$(MSTCOD) & "', '" & Trim$(JSNGBN) & "' , '" & _
            Trim$(JSNMON) & "', '" & Trim$(SINYMM) & "', '" & Trim$(JIGDAT) & "', " & oCompany.UserSignature
            
    sRecordset.DoQuery sQry
    If sRecordset.RecordCount <= 0 Then
        Execution_Save = False
        Exit Function
    End If
    
    Set sRecordset = Nothing

    Execution_Save = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set sRecordset = Nothing
    Sbo_Application.StatusBar.SetText "Execution_Save ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
    
    Execution_Save = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim MstInfo         As ZPAY_g_EmpID
    Dim JIGDAT          As String
    Select Case oUID
      Case "JSNYER"
            If Trim$(oForm.Items(oUID).Specific.String) <> "" Then
                ZPAY_GBL_JSNYER = oForm.Items(oUID).Specific.String
            Else
                oForm.DataSources.UserDataSources("JSNYER").ValueEx = ZPAY_GBL_JSNYER
            End If
            oJsnYear = oForm.DataSources.UserDataSources("JSNYER").ValueEx
      Case "SMonth"
            oSMonth = oForm.Items(oUID).Specific.VALUE
            oForm.DataSources.UserDataSources("SMonth").ValueEx = oSMonth
      Case "EMonth"
            oEMonth = oForm.Items(oUID).Specific.VALUE
            oForm.DataSources.UserDataSources("EMonth").ValueEx = oEMonth
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oForm.Items(oUID).Specific.String = ""
                oForm.DataSources.UserDataSources("MSTNAM").ValueEx = ""
            Else
                oForm.Items(oUID).Specific.String = UCase$(oForm.Items(oUID).Specific.String)
                MstInfo = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oForm.DataSources.UserDataSources("MSTNAM").ValueEx = MstInfo.MSTNAM
            End If
            oForm.Items("MSTNAM").Update
      Case "JSNGBN"
            If Trim$(oJsnYear) = "" Then oJsnYear = Format$(Now, "yyyy")
            oForm.Freeze True
            If oForm.Items(oUID).Specific.Selected.VALUE = "1" Then
                oForm.DataSources.UserDataSources("JSNMON").ValueEx = "12"
                oForm.DataSources.UserDataSources("SMonth").ValueEx = oJsnYear & "0101"
                oForm.DataSources.UserDataSources("EMonth").ValueEx = oJsnYear & "1231"
            Else
                oForm.DataSources.UserDataSources("JSNMON").ValueEx = ""
                oForm.DataSources.UserDataSources("SMonth").ValueEx = ""
                oForm.DataSources.UserDataSources("EMonth").ValueEx = ""
            End If
            oForm.Items("SMonth").Update
            oForm.Items("EMonth").Update
            oForm.Freeze False
        Case "JSNMON"
            If Trim$(oForm.DataSources.UserDataSources(oUID).ValueEx) = "" Then
                oForm.DataSources.UserDataSources("SMonth").ValueEx = ""
                oForm.DataSources.UserDataSources("EMonth").ValueEx = ""
                oForm.DataSources.UserDataSources("SINYMM").ValueEx = ""
                oForm.DataSources.UserDataSources("JIGDAT").ValueEx = ""
            Else
                If Not oForm.Items("JSNGBN").Specific.Selected Is Nothing Then
                    oForm.DataSources.UserDataSources("SMonth").ValueEx = oJsnYear & Trim$(oForm.DataSources.UserDataSources("JSNMON").ValueEx) & "01"
                    oForm.DataSources.UserDataSources("EMonth").ValueEx = oJsnYear & Trim$(oForm.DataSources.UserDataSources("JSNMON").ValueEx) & MDC_SetMod.Lday(oJsnYear & Trim$(oForm.DataSources.UserDataSources("JSNMON").ValueEx))
                    oForm.DataSources.UserDataSources("SINYMM").ValueEx = ""
                    oForm.DataSources.UserDataSources("JIGDAT").ValueEx = ""
                End If
            End If
            oForm.Items("SMonth").Update
            oForm.Items("EMonth").Update
        Case "JIGDAT"
            JIGDAT = oForm.DataSources.UserDataSources("JIGDAT").ValueEx
            If Trim$(JIGDAT) = "" Then
                oForm.DataSources.UserDataSources("JIGDAT").ValueEx = ""
                oForm.DataSources.UserDataSources("SINYMM").ValueEx = ""
            Else
                If Right$(JIGDAT, 2) <= "10" Then
                    oForm.DataSources.UserDataSources("SINYMM").ValueEx = Left$(JIGDAT, 6)
                Else
                    oForm.DataSources.UserDataSources("SINYMM").ValueEx = Format$(DateAdd("M", 1, Format$(Left$(JIGDAT, 6) & "01", "0000-00-00")), "YYYYMM")
                End If
            End If
            
    End Select
    oForm.Items(oUID).Update
End Sub


