VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY985"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : PH_PY985.cls
'//  Module         : �λ����>�������
'//  Desc           : �Ƿ�����޸������ڷ� �����ü����
'//  FormType       :
'//  Create Date    : 2014.01.28
'//  Modified Date  : 2017.01.24
'//  Creator        : Hys
'//  Modifier       :
'//  Copyright  (c) Poongsan Holdings
'//****************************************************************************

Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private sRecordset      As SAPbobsCOM.Recordset
Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     'Ŭ�������� ������ ������ ������ Uid��

Private CLTCOD          As String
Private yyyy            As String
Private HtaxID          As String
Private DocDate         As String
Private oFilePath       As String

Private FILNAM      As String * 30    '��  ��  ��
Private MaxRow      As Long
Private BUSCNT      As Integer  '/ B���ڵ��Ϸù�ȣ
Private BUSTOT      As Integer  '/ B���ڵ��Ѱ���

Private NEWCNT      As Integer
Private OLDCNT      As Integer
Private C_SAUP      As String
Private C_YYYY      As String
Private C_SABUN     As String
Private E_BUYCNT    As String
Private C_BUYCNT    As String


'2013����� 250 BYTE
'2015����� 251 BYTE
'2017����� 251 BYTE

Private Type A_record
        A001   As String * 1     '���ڵ屸�� 'A'
        A002   As String * 2     '�ڷᱸ��   '26'
        A003   As String * 3     '������
        A004   As String * 6     '�Ϸù�ȣ
        A005   As String * 8     '��������
        A006   As String * 10    '����ڹ�ȣ
        A007   As String * 20    'Ȩ�ؽ�ID
        A008   As String * 4     '�������α׷��ڵ�
        A009   As String * 10    '����ڹ�ȣ
        A010   As String * 40    '���θ�(��ȣ)
        A011   As String * 13    '�ҵ������ֹε�Ϲ�ȣ
        A012   As String * 1     '��,�ܱ���
        A013   As String * 30    '����
        A014   As String * 10    '����ó����ڵ�Ϲ�ȣ
        A015   As String * 40    '����ó��ȣ
        A016   As String * 1     '�Ƿ������ڵ�
        A017   As String * 5     '�Ǽ�
        A018   As String * 11    '���ޱݾ�
        A019   As String * 1     '���ӽü���
        A020   As String * 13    '�ֹε�Ϲ�ȣ
        A021   As String * 1     '��,�ܱ����ڵ�
        A022   As String * 1     '���ε��ش翩��
        A023   As String * 1     '������Ⱓ�ڵ�
        A024   As String * 19    '����
End Type
Dim A_rec As A_record

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY985.srf")
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
    
    oFormUniqueID = "PH_PY985_" & GetTotalFormsCount
    
    '�� �Ҵ�
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '�÷��ǿ� ���� ��´�   **�÷����̶� ��ü�� ��� ���� �迭�μ� ���⼭�� Ȱ��ȭ�Ǿ��� �ִ� ���� ��� �ִ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    AddForms Me, oFormUniqueID, "PH_PY985"
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    
    oForm.Freeze True
    CreateItems
    oForm.Freeze False
    
    oForm.EnableMenu ("1281"), False '/ ã��
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
'
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                                BubbleEvent = False
                                Exit Sub
                        End If
                    End If
                End If
                If pval.ItemUID = "Btn01" Then
                   If File_Create = False Then
                       BubbleEvent = False
                       Exit Sub
                   Else
                       BubbleEvent = False
                       oForm.Mode = fm_OK_MODE
                   End If
                
                End If
            Else
            End If
            
        Case et_COMBO_SELECT:
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID
                           '//������� �ٲ��
                           Case "CLTCOD"
                                sQry = "SELECT U_HomeTId, U_ChgDpt, U_ChgName, U_ChgTel  FROM [@PH_PY005A] WHERE U_CLTCode = '" & Trim(oForm.Items("CLTCOD").Specific.VALUE) & "'"
                                oRecordSet.DoQuery sQry
                                oForm.Items("HtaxID").Specific.String = Trim(oRecordSet.Fields("U_HomeTId").VALUE)
                    End Select
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
    
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
    
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
    
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
    
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
            oForm.Items("JsnYear").CLICK ct_Regular
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
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    oForm.Items("CLTCOD").DisplayDesc = True
    '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
    Call CLTCOD_Select(oForm, "CLTCOD")
    
    oForm.Items("YYYY").Specific.String = Format(Now, "YYYY") - 1    '�⵵ �⺻�⵵���� - 1
    
    Call oForm.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)     '��������
    oForm.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Function File_Create() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oStr            As String
    Dim sQry            As String
    
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    'ȭ�麯���� ���������� MOVE
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = Trim(oForm.Items("YYYY").Specific.VALUE)
    HtaxID = Trim(oForm.Items("HtaxID").Specific.VALUE)
    DocDate = Trim(oForm.Items("DocDate").Specific.VALUE)
    
    ErrNum = 0
    
    '/ Question
    If Sbo_Application.MessageBox("�Ƿ�� �Ű������� �����Ͻðڽ��ϱ�?", 2, "&Yes!", "&No") = 2 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    '/ A RECORD ó��
    If File_Create_A_record = False Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    Close #1
    
    Sbo_Application.StatusBar.SetText "�����ü������ ���������� �Ϸ�Ǿ����ϴ�.", bmt_Short, smt_Success
    File_Create = True
    Set sRecordset = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set sRecordset = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "����Ͽ����ϴ�.", bmt_Short, smt_Success
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "A���ڵ� ���� ����.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "File_Create ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    File_Create = False
End Function
Private Function File_Create_A_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim PRTDAT          As String
    Dim saup            As String
    Dim CheckA          As String
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CheckA = False  '/üũ�ʿ�����
    ErrNum = 0
    
    '/ A_RECORE QUERY
    sQry = "EXEC PH_PY985_A '" & CLTCOD & "', '" & yyyy & "', '" & HtaxID & "', '" & DocDate & "'"
    oRecordSet.DoQuery sQry
        
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        'PATH�� �����̸� �����
        saup = oRecordSet.Fields("A009").VALUE  '����ڹ�ȣ
        oFilePath = "C:\BANK\CA" + Mid$(saup, 1, 7) + "." + Mid$(saup, 8, 3)
        Close #1
        Open oFilePath For Output As #1
        
        Dim ProgressBar01 As SAPbouiCOM.ProgressBar
        Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("�ۼ�����!", oRecordSet.RecordCount, False)
    
        NEWCNT = 0
        
        Do Until oRecordSet.EOF
        
            NEWCNT = NEWCNT + 1     '/ �Ϸù�ȣ
            
            'A RECORD MOVE
            
            A_rec.A001 = oRecordSet.Fields("A001").VALUE
            A_rec.A002 = oRecordSet.Fields("A002").VALUE
            A_rec.A003 = oRecordSet.Fields("A003").VALUE
            A_rec.A004 = Format$(NEWCNT, String$(Len(A_rec.A004), "0"))  '/ �Ϸù�ȣ
            A_rec.A005 = oRecordSet.Fields("A005").VALUE
            A_rec.A006 = oRecordSet.Fields("A006").VALUE
            A_rec.A007 = oRecordSet.Fields("A007").VALUE
            A_rec.A008 = oRecordSet.Fields("A008").VALUE
            A_rec.A009 = oRecordSet.Fields("A009").VALUE
            A_rec.A010 = oRecordSet.Fields("A010").VALUE
            A_rec.A011 = oRecordSet.Fields("A011").VALUE
            A_rec.A012 = oRecordSet.Fields("A012").VALUE
            A_rec.A013 = oRecordSet.Fields("A013").VALUE
            A_rec.A014 = oRecordSet.Fields("A014").VALUE
            A_rec.A015 = oRecordSet.Fields("A015").VALUE
            A_rec.A016 = oRecordSet.Fields("A016").VALUE
            A_rec.A017 = Format$(oRecordSet.Fields("A017").VALUE, String$(Len(A_rec.A017), "0"))
            A_rec.A018 = Format$(oRecordSet.Fields("A018").VALUE, String$(Len(A_rec.A018), "0"))
            A_rec.A019 = oRecordSet.Fields("A019").VALUE
            A_rec.A020 = oRecordSet.Fields("A020").VALUE
            A_rec.A021 = oRecordSet.Fields("A021").VALUE
            A_rec.A022 = oRecordSet.Fields("A022").VALUE
            A_rec.A023 = oRecordSet.Fields("A023").VALUE
            A_rec.A024 = oRecordSet.Fields("A024").VALUE
            
            
            
            Print #1, MDC_SetMod.sStr(A_rec.A001) & MDC_SetMod.sStr(A_rec.A002) & MDC_SetMod.sStr(A_rec.A003) & MDC_SetMod.sStr(A_rec.A004) & _
                      MDC_SetMod.sStr(A_rec.A005) & MDC_SetMod.sStr(A_rec.A006) & MDC_SetMod.sStr(A_rec.A007) & MDC_SetMod.sStr(A_rec.A008) & _
                      MDC_SetMod.sStr(A_rec.A009) & MDC_SetMod.sStr(A_rec.A010) & MDC_SetMod.sStr(A_rec.A011) & MDC_SetMod.sStr(A_rec.A012) & _
                      MDC_SetMod.sStr(A_rec.A013) & MDC_SetMod.sStr(A_rec.A014) & MDC_SetMod.sStr(A_rec.A015) & MDC_SetMod.sStr(A_rec.A016) & _
                      MDC_SetMod.sStr(A_rec.A017) & MDC_SetMod.sStr(A_rec.A018) & MDC_SetMod.sStr(A_rec.A019) & MDC_SetMod.sStr(A_rec.A020) & _
                      MDC_SetMod.sStr(A_rec.A021) & MDC_SetMod.sStr(A_rec.A022) & MDC_SetMod.sStr(A_rec.A023) & MDC_SetMod.sStr(A_rec.A024)
                      
            oRecordSet.MoveNext
            
            ProgressBar01.VALUE = ProgressBar01.VALUE + 1
            ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "�� �ۼ���........!"
        
            
        Loop
    End If
    
    If CheckA = False Then
        File_Create_A_record = True
    Else
        File_Create_A_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "�Ƿ���ڷᰡ �������� �ʽ��ϴ�. ����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("A���ڵ����: " & Err.Description, False, True)
    End If
    
    File_Create_A_record = False
    
End Function


Private Sub Matrix_AddRow(ByVal MatrixMsg As String, Optional Insert_YN As Boolean, Optional MatrixErr As Boolean)
    If MatrixErr = True Then
        oForm.DataSources.UserDataSources("Col0").VALUE = "??"
    Else
        oForm.DataSources.UserDataSources("Col0").VALUE = ""
    End If
    oForm.DataSources.UserDataSources("Col1").VALUE = MatrixMsg
    If Insert_YN = True Then
        oMat1.AddRow
        MaxRow = MaxRow + 1
    End If
    oMat1.SetLineData MaxRow
End Sub


'ȭ�麯�� CHECK
Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel
    Dim ErrNum      As Integer
    
    ErrNum = 0
    '/ �ʼ�Check
    If oForm.Items("HtaxID").Specific.VALUE = "" Then
        ErrNum = 1
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("DocDate").Specific.VALUE = "" Then
        ErrNum = 2
        GoTo HeaderSpaceLineDel
    End If
    
    HeaderSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
HeaderSpaceLineDel:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "Ȩ�ؽ�ID(5�ڸ��̻�)�� Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "�������ڴ� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel ���� �� ������ �߻��߽��ϴ�." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    
    HeaderSpaceLineDel = False
End Function

