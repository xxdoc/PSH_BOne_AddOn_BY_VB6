VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY509"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY509.cls
'//  Module         : ��õ¡��>�ٷμҵ�
'//  Desc           : �����ڷ� �����۾�
'//  FormType       : 2010110509
'//  Create Date    : 2009.02.13
'//  Modified Date  :
'//  Creator        : Choi Dong Kwon
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oDS_ZPY509H     As SAPbouiCOM.DBDataSource
Private oDS_ZPY509L     As SAPbouiCOM.DBDataSource

Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     'Ŭ�������� ������ ������ ������ Uid��
Private Col_Last_Uid    As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private Col_Last_Row    As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY509.srf")
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
    
    oFormUniqueID = "ZPY509_" & GetTotalFormsCount
    
    '�� �Ҵ�
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '�÷��ǿ� ���� ��´�   **�÷����̶� ��ü�� ��� ���� �迭�μ� ���⼭�� Ȱ��ȭ�Ǿ��� �ִ� ���� ��� �ִ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    AddForms Me, oFormUniqueID, "ZPY509"
    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm.DataBrowser.BrowseBy = "Code"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    Call CreateItems
    Call FormItemEnabled
    
    oForm.EnableMenu ("1293"), False  '/ �����
    oForm.EnableMenu ("1283"), True  '/ ����
    oForm.EnableMenu ("1287"), False  '/ ����
    oForm.EnableMenu ("1284"), False  '/ ���
    
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

'*******************************************************************
'
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oColumn         As SAPbouiCOM.Column
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ Matrix
    Set oMat1 = oForm.Items("Mat1").Specific
    
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_ZPY509H = oForm.DataSources.DBDataSources("@ZPY509H")
    Set oDS_ZPY509L = oForm.DataSources.DBDataSources("@ZPY509L")
    
    '//�����
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
'    oCombo.ValidValues.Add "%", "��ü"
    oForm.Items("CLTCOD").DisplayDesc = True
'    oCombo.Select "%", psk_ByValue
    
    '/ �����ڷ�
    Set oColumn = oMat1.Columns("Col5")
    oColumn.ValOff = "N"
    oColumn.ValOn = "Y"
    
    '/ �����ڷ�
    Set oColumn = oMat1.Columns("Col2")
    oColumn.ValOff = "N"
    oColumn.ValOn = "Y"
    
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error

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
            If pval.ItemUid = "1" Then
                '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                '//�߰��� ������Ʈ�ÿ�
                '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If MDC_SetMod.Value_ChkYn("[@ZPY509H]", "Code", "'" & oForm.Items("CLTCOD").Specific.Selected.Value & "'") = False Then
                            Sbo_Application.StatusBar.SetText "�̹� ����Ǿ��� �ִ� ����� ����� ��ġ�մϴ�", bmt_Short, smt_Error
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    ElseIf MatrixSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call Batch_EndCheck
                    End If
                End If
            ElseIf pval.ItemUid = "Btn1" Then
                Call Create_Year
            End If
        Else
            If pval.ItemUid = "1" And pval.ActionSuccess = True And oForm.Mode = fm_ADD_MODE Then
                oForm.Mode = fm_OK_MODE
                Call Sbo_Application.ActivateMenuItem("1282")
            ElseIf pval.ItemUid = "Mat1" And (pval.ColUID = "Col2" Or pval.ColUID = "Col5") Then
                Call FlushToItemValue(pval.ColUID, pval.Row)
            End If
        End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_VALIDATE
        If pval.ItemUid = "Mat1" And pval.ColUID = "Col1" Then
            Call FlushToItemValue(pval.ColUID, pval.Row)
        End If
    
'et_CLICK'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_CLICK
        If pval.FormUID = oForm.uniqueID And pval.BeforeAction = True And Last_Item = "Mat1" And Col_Last_Uid = "Col1" And Col_Last_Row > 0 _
          And (Col_Last_Uid <> pval.ColUID Or Col_Last_Row <> pval.Row) And pval.ItemUid <> "1000001" And pval.ItemUid <> "2" Then
            If Col_Last_Row > oMat1.VisualRowCount Then
                Exit Sub
            End If
        ElseIf pval.FormUID = oForm.uniqueID And pval.BeforeAction = True And pval.ItemUid = "Mat1" And pval.Row > 0 Then
            Last_Item = pval.ItemUid
            Col_Last_Row = pval.Row
            Col_Last_Uid = pval.ColUID
        End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_GOT_FOCUS
        If Last_Item = "Mat1" Then
            If pval.Row > 0 Then
                Last_Item = pval.ItemUid
                Col_Last_Row = pval.Row
                Col_Last_Uid = pval.ColUID
            End If
        Else
            Last_Item = pval.ItemUid
            Col_Last_Row = 0
            Col_Last_Uid = ""
        End If
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_FORM_UNLOAD
        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        '�÷��ǿ��� ������ ��� �޸� ����
        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        If pval.BeforeAction = False Then
            RemoveForms oFormUniqueID
            Set oForm = Nothing
            Set oDS_ZPY509H = Nothing
            Set oDS_ZPY509L = Nothing
            Set oMat1 = Nothing
        End If
'et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_MATRIX_LOAD:
        If pval.BeforeAction = False Then
            Call FormItemEnabled
            Call Matrix_AddRow(oMat1.VisualRowCount, False)
        End If
            
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:", bmt_Short, smt_Error
End Sub

'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    
    If pval.BeforeAction = True Then
        Select Case pval.MenuUID
            Case "1283" '/ ����
                If Sbo_Application.MessageBox("���� ȭ�鳻����ü�� ���� �Ͻðڽ��ϱ�? ������ �� �����ϴ�.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1288", "1289", "1290", "1291"
                Call AuthorityCheck(oForm, "CLTCOD", "@ZPY509H", "Code")       '//������ ���ѿ� ���� ����� ����
                
            Case Else
                Exit Sub
        End Select
    Else
        Select Case pval.MenuUID
            Case "1287" '/ ����
            Case "1281", "1282":
                Call FormItemEnabled
                If pval.MenuUID = "1282" Then
                    Call Matrix_AddRow(0, True)
    
                End If
    '        Case "1283" '/ ����
    '             FormItemEnabled
            Case "1288" To "1291":
                Call FormItemEnabled
    '        Case "1293" '/ �����
        
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
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oRow As Long, Optional Insert_YN As Boolean)
    If Insert_YN = False Then
        oDS_ZPY509L.InsertRecord (oRow)
    End If
    oDS_ZPY509L.Offset = oRow
    oDS_ZPY509L.setValue "U_LineNum", oRow, oRow + 1
    oDS_ZPY509L.setValue "U_JOBYER", oRow, ""
    oDS_ZPY509L.setValue "U_MONCHK", oRow, "N"
    oDS_ZPY509L.setValue "U_ENDCHK", oRow, "N"
    oDS_ZPY509L.setValue "U_LGNADM", oRow, ""
    oDS_ZPY509L.setValue "U_MODDAT", oRow, ""
    oMat1.LoadFromDataSource
End Sub
Private Sub FormItemEnabled()
    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_FIND_MODE Then
        oForm.Items("CLTCOD").Enabled = True
    Else
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("Btn1").Enabled = False
    End If
    If oForm.Mode <> fm_FIND_MODE Then
        oForm.Items("Btn1").Enabled = True
    Else
        oForm.Items("Btn1").Enabled = False
    End If
    '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    
    ErrNum = 0
 
 '/ Check
    Select Case True
    Case oForm.Items("CLTCOD").Specific.Selected Is Nothing
         ErrNum = 1
         GoTo Error_Message
    End Select
    
    Call oDS_ZPY509H.setValue("Code", 0, oDS_ZPY509H.GetValue("U_CLTCOD", 0))
    
    HeaderSpaceLineDel = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "�ڻ��ڵ�� �ʼ��Դϴ�. �����Ͽ� �ֽʽÿ�", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel Error : " & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    HeaderSpaceLineDel = False
End Function
Private Function MatrixSpaceLineDel() As Boolean
'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'������ �������� ��ȿ���� �����Ѵ�
'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
On Error GoTo Error_Message
    Dim i           As Long
    Dim k           As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    
    ErrNum = 0
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    'ȭ����� ��Ʈ������ �Էµ� ������ ��� ������ͼҽ��� �ѱ��
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat1.FlushToDataSource
        
    If oMat1.RowCount = 1 Then          '// Mat1�� ���� �ִ��� Ȯ�� (ErrorNumber : 1)
        ErrNum = 1
        GoTo Error_Message
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//������ �� �ϳ��� ���� i=0���� �����ϹǷ� �ϳ��� ���Ƿ�
    '//oMat1.RowCount - 2�� �ȴ�..�ݵ�� ��� ���� �ϴ� �ʼ����� Ȯ���Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    For i = 0 To oMat1.VisualRowCount - 2      '// Mat1�� �Է°��� �ùٸ��� ������ Ȯ�� (ErrorNumber : 3)
        oDS_ZPY509L.Offset = i
        If Trim$(oDS_ZPY509L.GetValue("U_JOBYER", i)) = "" Then
            ErrNum = 2
            oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
            GoTo Error_Message

        Else
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            '�ߺ�üũ�۾�
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            Chk_Data = Trim$(oDS_ZPY509L.GetValue("U_JOBYER", i))
            For k = i + 1 To oMat1.VisualRowCount - 2
                oDS_ZPY509L.Offset = k
                If Trim$(Chk_Data) = Trim$(oDS_ZPY509L.GetValue("U_JOBYER", k)) Then
                    ErrNum = 3
                    oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            Next k
        End If
    Next i
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//�Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '//�̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oDS_ZPY509L.RemoveRecord oDS_ZPY509L.Size - 1   '// Mat1�� ����������(�����) ����
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '���� �����Ͽ����� DB������ �ҽ��� �ٽ� �����´�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat1.LoadFromDataSource
    
    MatrixSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "�Է��� �����Ͱ� �����ϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "�ͼӳ⵵�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "�ͼӳ⵵�� �ߺ��ԷµǾ����ϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "MatrixSpaceLineDel Error : " & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    MatrixSpaceLineDel = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : Create_Year
' DateTime  : 2009-02-16
' Author    :
' Purpose   : ���� ����
'---------------------------------------------------------------------------------------
'
Private Sub Create_Year()
On Error GoTo Error_Message

    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CLTCOD      As String
    Dim iRow        As Long
    Dim kRow        As Long
    Dim MaxRow      As Long
    Dim check       As String
    
    Dim NowYer      As Integer
    
    CLTCOD = Trim$(oDS_ZPY509H.GetValue("U_CLTCOD", 0))
    NowYer = MDC_SetMod.Get_ReData("CONVERT(CHAR(4), GETDATE(), 120)", "1", "OADM", "1")
    MaxRow = oMat1.VisualRowCount - 1
    
    oMat1.FlushToDataSource
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
        
    '// Matrix �ǹؿ� �ִ� ���� ����
    If Trim$(oDS_ZPY509L.GetValue("U_JOBYER", MaxRow)) = "" Then
        oDS_ZPY509L.RemoveRecord MaxRow
        MaxRow = MaxRow - 1
    End If
    
Create_Step1:   '/ ���� ���굥���Ϳ� ���� ���������� ����
    sQry = "EXEC ZPY509_1 '" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        GoTo Create_Step2
    End If
    
    Do Until oRecordSet.EOF
        check = "N"
        '// �����Ϸ��� ������ �̹� Matrix�� �����ϴ��� ��
        If oDS_ZPY509L.Size > 0 Then
            For kRow = 0 To oDS_ZPY509L.Size - 1
                If Val(oDS_ZPY509L.GetValue("U_JOBYER", kRow)) = Val(oRecordSet.Fields("U_JSNYER").Value) Then
                    check = "Y"
                End If
            Next kRow
        End If
        '// �����Ϸ��� ������ �̹� Matrix�� �����ϴ� ��� ���� ����
        If check = "N" Then
            MaxRow = MaxRow + 1
            oDS_ZPY509L.InsertRecord MaxRow
            oDS_ZPY509L.Offset = MaxRow
            oDS_ZPY509L.setValue "U_LINENUM", MaxRow, MaxRow + 1
            oDS_ZPY509L.setValue "U_JOBYER", MaxRow, Val(oRecordSet.Fields("U_JSNYER").Value)
            oDS_ZPY509L.setValue "U_MONCHK", MaxRow, "N"
            oDS_ZPY509L.setValue "U_ENDCHK", MaxRow, "N"
            oDS_ZPY509L.setValue "U_LGNADM", MaxRow, ""
            oDS_ZPY509L.setValue "U_MODDAT", MaxRow, ""
        End If
        oRecordSet.MoveNext
    Loop

Create_Step2:   '/ ���غ��� ������ 10�Ⱓ�� ���� ���������� ����

    For iRow = NowYer To NowYer + 10
        check = "N"
        '// �����Ϸ��� ������ �̹� Matrix�� �����ϴ��� ��
        If oDS_ZPY509L.Size > 0 Then
            For kRow = 0 To oDS_ZPY509L.Size - 1
                If Val(oDS_ZPY509L.GetValue("U_JOBYER", kRow)) = iRow Then
                    check = "Y"
                End If
            Next kRow
        End If
        '// �����Ϸ��� ������ �̹� Matrix�� �����ϴ� ��� ���� ����
        If check = "N" Then
            MaxRow = MaxRow + 1
            oDS_ZPY509L.InsertRecord MaxRow
            oDS_ZPY509L.Offset = MaxRow
            oDS_ZPY509L.setValue "U_LINENUM", MaxRow, MaxRow + 1
            oDS_ZPY509L.setValue "U_JOBYER", MaxRow, iRow
            oDS_ZPY509L.setValue "U_MONCHK", MaxRow, "N"
            oDS_ZPY509L.setValue "U_ENDCHK", MaxRow, "N"
            oDS_ZPY509L.setValue "U_LGNADM", MaxRow, ""
            oDS_ZPY509L.setValue "U_MODDAT", MaxRow, ""
        End If
    Next
    
    '// Matrix�ؿ� ���� �߰�
    Call Matrix_AddRow(MaxRow + 1, False)
    
    Exit Sub
Error_Message:
    Sbo_Application.StatusBar.SetText "Create_Year Error : " & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim MONCHK  As String
    Dim ENDCHK  As String
    
    Select Case oUID
    Case "Col1"
        oMat1.FlushToDataSource
        
        oDS_ZPY509L.Offset = oRow - 1
        
        If oRow = oMat1.RowCount And Trim$(oDS_ZPY509L.GetValue("U_JOBYER", oRow - 1)) <> "" Then
            Matrix_AddRow oRow
            oMat1.Columns("Col1").Cells(oRow).CLICK ct_Regular
        End If
        oMat1.LoadFromDataSource
    Case "Col2", "Col5"
    
        oForm.Freeze True
        oMat1.FlushToDataSource
        
        MONCHK = Trim$(oDS_ZPY509L.GetValue("U_MONCHK", oRow - 1))
        ENDCHK = Trim$(oDS_ZPY509L.GetValue("U_ENDCHK", oRow - 1))
        
        oDS_ZPY509L.Offset = oRow - 1
        '// �����ڷᰡ Y�̸� �����ڷ�� �ڵ����� Y�� ����
        '// �����ڷ� �Ǵ� �����ڷῡ Y�� üũ�ϸ� �ڵ����� ����ڿ� �������ڸ� ����������� ����
        If ENDCHK = "Y" Then
            oDS_ZPY509L.setValue "U_MONCHK", oRow - 1, "Y"
        End If
        oDS_ZPY509L.setValue "U_LGNADM", oRow - 1, oCompany.UserName
        oDS_ZPY509L.setValue "U_MODDAT", oRow - 1, Format$(Now, "YYYYMMDD")
        oMat1.LoadFromDataSource
        oForm.Freeze False
        
    End Select

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Batch_EndCheck
' DateTime  : 2009-02-16
' Author    :
' Purpose   : �����۾� ó��
'---------------------------------------------------------------------------------------
'
Private Sub Batch_EndCheck()
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim iRow        As Long
    
    Dim CLTCOD      As String
    Dim JOBYER      As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CLTCOD = Trim$(oDS_ZPY509H.GetValue("U_CLTCOD", 0))
    
    oCompany.StartTransaction
    For iRow = 0 To oMat1.VisualRowCount - 1
        If Trim$(oDS_ZPY509L.GetValue("U_MODDAT", iRow)) = Format$(Now, "YYYYMMDD") Then
            JOBYER = Trim$(oDS_ZPY509L.GetValue("U_JOBYER", iRow))
            If Trim$(oDS_ZPY509L.GetValue("U_ENDCHK", iRow)) = "Y" Then
                sQry = "EXEC ZPY509 '" & CLTCOD & "', '" & JOBYER & "', '1'"
                oRecordSet.DoQuery sQry
            ElseIf Trim$(oDS_ZPY509L.GetValue("U_MONCHK", iRow)) = "Y" Then
                sQry = "EXEC ZPY509 '" & CLTCOD & "', '" & JOBYER & "', '2'"
                oRecordSet.DoQuery sQry
            Else
                sQry = "EXEC ZPY509 '" & CLTCOD & "', '" & JOBYER & "', '3'"
                oRecordSet.DoQuery sQry
            End If
        End If
    Next iRow
    oCompany.EndTransaction wf_Commit
    
    Set oRecordSet = Nothing
    
    Exit Sub
Error_Message:
    If oCompany.InTransaction Then
        oCompany.EndTransaction wf_RollBack
    End If
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "Batch_EndCheck Error : " & Space$(5) & Err.Description, bmt_Short, smt_Error

End Sub