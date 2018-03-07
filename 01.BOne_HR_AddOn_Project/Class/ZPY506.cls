VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY506"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY506.cls
'//  Module         : �λ����>�������
'//  Desc           : �Ƿ����� ���
'//  FormType       : 2000060506
'//  Create Date    : 2006.01.15
'//  Modified Date  : 2012.02.03
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************

Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oDS_ZPY506H     As SAPbouiCOM.DBDataSource    '�ý����ڵ� ���
Private oDS_ZPY506L     As SAPbouiCOM.DBDataSource    '�ý����ڵ� ����
Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     'Ŭ�������� ������ ������ ������ Uid��
Private Col_Last_Uid    As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private Col_Last_Row    As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��
Private oOLDCHK         As String


Private Sub FormItemEnabled()
    If oForm.Mode = fm_FIND_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = True
        oForm.Items("DocNum").Enabled = True
        oForm.Items("ENDCHK").Enabled = True
    ElseIf oForm.Mode = fm_ADD_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        oForm.Items("ENDCHK").Enabled = True
    ElseIf oForm.Mode = fm_OK_MODE Then
        oForm.Items("JSNYER").Enabled = False
        oForm.Items("MSTCOD").Enabled = False
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        '// �⸶���Ȱ��� ��Ȱ��ȭ
        oOLDCHK = oDS_ZPY506H.GetValue("U_ENDCHK", 0)
        If MDC_SetMod.Get_ReData("U_ENDCHK", "U_JOBYER", "[@ZPY509L]", "'" & oDS_ZPY506H.GetValue("U_JSNYER", 0) & "'", " AND Code = '" & oDS_ZPY506H.GetValue("U_CLTCOD", 0) & "'") = "Y" Then
            oForm.Items("ENDCHK").Enabled = False
        Else
            oForm.Items("ENDCHK").Enabled = True
        End If
        
    End If
End Sub
Private Sub ShowSource(JSNYER$, MSTCOD$, CLTCOD$)
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim DocEntry$
    Dim oMast           As ZPAY_g_EmpID
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    sQry = "SELECT DocEntry FROM [@ZPY506H]"
    sQry = sQry & "   WHERE U_JSNYER = N'" & JSNYER & "'"
    sQry = sQry & "   AND   U_MSTCOD = N'" & MSTCOD & "'"
    sQry = sQry & "   AND   U_CLTCOD = N'" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            DocEntry = oRecordSet.Fields(0).Value
            oRecordSet.MoveNext
        Loop
        oForm.Items("JSNYER").Specific.Value = JSNYER
        oForm.Items("MSTCOD").Specific.String = MSTCOD
        
        Call oForm.Items("CLTCOD").Specific.Select(CLTCOD, psk_ByValue)
        oForm.Items("DocNum").Specific.Value = DocEntry
        
        oForm.Items("DocNum").Update
        oMat1.LoadFromDataSource
        oForm.Update
        oForm.Items("1").CLICK ct_Regular
        
    Else
        oForm.Mode = fm_OK_MODE
        Call Sbo_Application.ActivateMenuItem("1282")
        
        Call oDS_ZPY506H.setValue("U_JSNYER", 0, JSNYER)
        Call oDS_ZPY506H.setValue("U_MSTCOD", 0, MSTCOD)
        Call oDS_ZPY506H.setValue("U_CLTCOD", 0, CLTCOD)
        oMast = MDC_SetMod.Get_EmpID_InFo(MSTCOD)
        Call oDS_ZPY506H.setValue("U_MSTNAM", 0, oMast.MSTNAM)
        Call oDS_ZPY506H.setValue("U_EmpID", 0, oMast.EmpID)
        
        oForm.Update
        
        Sbo_Application.SendKeys "{TAB}"
    End If
    
    Set oRecordSet = Nothing
    
End Sub
Private Sub FormClear()
    Dim DocNum&
    
    DocNum = MDC_SetMod.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'ZPY506'", "")
    
    If DocNum = 0 Then
        oForm.Items("DocNum").Specific.String = 1
    Else
        oForm.Items("DocNum").Specific.String = DocNum
    End If
    Call FlushToItemValue("JSNYER")
    
End Sub

Private Sub Display_GibuMan(sRow As Long)
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim MSTCOD          As String
    Dim JSNYER          As String
    Dim FAMNAM          As String
    
    
    JSNYER = oDS_ZPY506H.GetValue("U_JSNYER", 0)
    MSTCOD = oDS_ZPY506H.GetValue("U_MSTCOD", 0)
    FAMNAM = oDS_ZPY506L.GetValue("U_FAMNAM", sRow)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

  '// �Ƿ����Է½� �ҵ��ڷ���ȭ���� �ξ簡�������� ���ο� �ش��ϴ� ������ ���� ǥ������.
   sQry = "SELECT  T0.U_FAMNAM AS FAMNAM, T0.U_CHKINT AS INTGBN, T0.U_FAMPER AS PERNBR, "
   sQry = sQry & " T0.U_CHKCOD AS GWANGE "
   sQry = sQry & " FROM [@ZPY501L] T0 INNER JOIN [@ZPY501H] T1 ON T0.DocEntry = T1.DocEntry"
   sQry = sQry & " WHERE T1.U_JSNYER = '" & Trim$(JSNYER) & "'"
   sQry = sQry & " AND   T1.U_MSTCOD =  '" & Trim$(MSTCOD) & "'"
   sQry = sQry & " AND   T0.U_FAMNAM =  '" & Trim$(FAMNAM) & "'"
   oRecordSet.DoQuery sQry
   If oRecordSet.RecordCount > 0 Then
        oDS_ZPY506L.setValue "U_INTGBN", sRow, oRecordSet.Fields("INTGBN").Value
        oDS_ZPY506L.setValue "U_PERNBR", sRow, oRecordSet.Fields("PERNBR").Value
        oDS_ZPY506L.setValue "U_GWANGE", sRow, oRecordSet.Fields("GWANGE").Value
   End If

    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oRecordSet = Nothing
     Sbo_Application.StatusBar.SetText "Display_GibuMan Error:" & Err.Number & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub
Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim i&
    Dim oMast           As ZPAY_g_EmpID
    Dim TOTCNT As Double
    Dim TOTAMT As Double
    
    Select Case oUID
      Case "JSNYER"
            If Trim$(oForm.Items(oUID).Specific.String) <> "" Then
                ZPAY_GBL_JSNYER = oForm.Items(oUID).Specific.String
            Else
                oDS_ZPY506H.setValue "U_JSNYER", 0, ZPAY_GBL_JSNYER
            End If
            oForm.Items(oUID).Update
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oDS_ZPY506H.setValue "U_MSTCOD", 0, ""
                oDS_ZPY506H.setValue "U_MSTNAM", 0, ""
                oDS_ZPY506H.setValue "U_EmpID", 0, ""
                oDS_ZPY506H.setValue "U_CLTCOD", 0, ""
            Else
                oDS_ZPY506H.setValue "U_MSTCOD", 0, UCase$(oForm.Items(oUID).Specific.String)
                oMast = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oDS_ZPY506H.setValue "U_MSTNAM", 0, oMast.MSTNAM
                oDS_ZPY506H.setValue "U_EmpID", 0, oMast.EmpID
                oDS_ZPY506H.setValue "U_CLTCOD", 0, oMast.CLTCOD
            End If
            oForm.Items("MSTNAM").Update
            oForm.Items("EmpID").Update
            oForm.Items("CLTCOD").Update
            oForm.Items(oUID).Update
    End Select
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    Select Case oUID
        Case "Col3", "Col4", "Col12"
            oMat1.FlushToDataSource

            For i = 1 To oMat1.VisualRowCount
                TOTCNT = TOTCNT + Val(oDS_ZPY506L.GetValue("U_MEDCNT", i - 1))
                TOTAMT = TOTAMT + Val(oDS_ZPY506L.GetValue("U_MEDAMT", i - 1))
            Next i
            oDS_ZPY506H.setValue "U_TOTCNT", 0, TOTCNT
            oDS_ZPY506H.setValue "U_TOTAMT", 0, TOTAMT
            oForm.Items("TOTCNT").Update
            oForm.Items("TOTAMT").Update
            
            oDS_ZPY506L.Offset = oRow - 1
'            oMat1.SetLineData oRow
        Case "Col11"
             oDS_ZPY506L.Offset = oRow - 1
             oDS_ZPY506L.setValue "U_FAMNAM", oRow - 1, oMat1.Columns(oUID).Cells(oRow).Specific.Value
             Call Display_GibuMan(oRow - 1)
             oMat1.SetLineData oRow
'        Case "Col1", "Col4"
'            oMat1.FlushToDataSource
'            oDS_ZPY506L.Offset = oRow - 1
'
'            If oRow = oMat1.RowCount And Val(oDS_ZPY506L.GetValue("U_MEDAMT", oRow - 1)) <> 0 Then
'                Matrix_AddRow oRow
'                oMat1.Columns("Col1").Cells(oRow).Click ct_Regular
'            End If
    End Select
    If Left$(oUID, 3) = "Col" Then
     oMat1.FlushToDataSource
            If oRow = oMat1.RowCount And Val(oDS_ZPY506L.GetValue("U_MEDAMT", oRow - 1)) <> 0 Then
                Matrix_AddRow oRow
                oMat1.Columns("Col4").Cells(oRow).CLICK ct_Regular
            End If
    End If
End Sub

Private Function MatrixSpaceLineDel() As Boolean
'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'������ �������� ��ȿ���� �����Ѵ�
'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
On Error GoTo Error_Message
    Dim iRow        As Long
    Dim kRow        As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    Dim GovidChk    As String
    
    ErrNum = 0
    '/ ����κ� üũ
    Select Case True
      Case Len(Trim$(oDS_ZPY506H.GetValue("U_JSNYER", 0))) <> 4
           ErrNum = 8
           GoTo Error_Message
      Case oDS_ZPY506H.GetValue("U_MSTCOD", 0) = ""
           ErrNum = 5
           GoTo Error_Message
      Case oDS_ZPY506H.GetValue("U_CLTCOD", 0) = ""
           ErrNum = 12
           GoTo Error_Message
           
    End Select
    '/ �ֹι�ȣüũ����
    GovidChk = MDC_SetMod.Get_ReData("ISNULL(T0.U_GovIDChk,'N')", "T1.U_MSTCOD", "[@PH_PY005A] T0 INNER JOIN [@PH_PY001A] T1 ON T0.CODE = T1.U_CLTCOD", "'" & Trim$(oDS_ZPY506H.GetValue("U_MSTCOD", 0)) & "'", "")
    
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
    For iRow = 0 To oMat1.VisualRowCount - 2      '// Mat1�� �Է°��� �ùٸ��� ������ Ȯ�� (ErrorNumber : 3)
        oDS_ZPY506L.Offset = iRow
        '// �Ƿ�������� ����������.
        If Trim$(oDS_ZPY506L.GetValue("U_MEDCOD", iRow)) <> "1" And Trim$(oDS_ZPY506L.GetValue("U_MEDNBR", iRow)) = "" Then   '/
            ErrNum = 6
            oMat1.Columns("Col1").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY506L.GetValue("U_MEDCOD", iRow)) <> "1" And Val(oDS_ZPY506L.GetValue("U_MEDCNT", iRow)) + Val(oDS_ZPY506L.GetValue("U_MEDCNT2", iRow)) = 0 Then
            ErrNum = 4
            oMat1.Columns("Col3").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Val(oDS_ZPY506L.GetValue("U_MEDAMT", iRow)) + Val(oDS_ZPY506L.GetValue("U_MEDAMT2", iRow)) = 0 Then
            ErrNum = 5
            oMat1.Columns("Col4").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY506L.GetValue("U_PERNBR", iRow)) = "" Then
            ErrNum = 7
            oMat1.Columns("Col6").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY506L.GetValue("U_DAECHK", iRow)) = "" Then  '/ �ش翩��
            ErrNum = 10
            oMat1.Columns("Col6").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        Else
            '// 6.�ֹι�ȣ ���� üũ
            If Trim$(GovidChk) = "Y" And Len(oDS_ZPY506L.GetValue("U_PERNBR", iRow)) > 0 Then
                If GovIDCheck(oDS_ZPY506L.GetValue("U_PERNBR", iRow)) = False Then
                    ErrNum = 11
                    oMat1.Columns("Col6").Cells(iRow + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            End If
'            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'            '�ߺ�üũ�۾�(2012.02.03_�ּ�ó�� ����� ��Ϲ�ȣ�� �Է� ���ϴ� ��쵵 ����)
'            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'            Chk_Data = Trim$(oDS_ZPY506L.GetValue("U_MEDCOD", iRow)) & Trim$(oDS_ZPY506L.GetValue("U_MEDNBR", iRow)) & Trim$(oDS_ZPY506L.GetValue("U_PERNBR", iRow))
'            For kRow = iRow + 1 To oMat1.VisualRowCount - 2
'                oDS_ZPY506L.Offset = kRow
'                If Trim$(Chk_Data) = Trim$(oDS_ZPY506L.GetValue("U_MEDCOD", kRow)) & Trim$(oDS_ZPY506L.GetValue("U_MEDNBR", kRow)) & Trim$(oDS_ZPY506L.GetValue("U_PERNBR", kRow)) Then
'                    ErrNum = 3
'                    oMat1.Columns("Col1").Cells(iRow + 1).CLICK ct_Regular
'                    GoTo Error_Message
'                End If
'            Next kRow
        End If
    Next iRow
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '//�Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '//�̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oDS_ZPY506L.RemoveRecord oDS_ZPY506L.Size - 1   '// Mat1�� ����������(�����) ����
    
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
        Sbo_Application.StatusBar.SetText "����ó��ȣ�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "�Ƿ�������� ����ó�� �ߺ��ԷµǾ����ϴ�. ������� ����ó�� �����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "���ްǼ��� 0���Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "�Ƿ�� ���ޱݾ��� 0�Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "����ó����ڹ�ȣ�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "�Ƿ�����޴������ �ֹε�Ϲ�ȣ�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 8 Then
        Sbo_Application.StatusBar.SetText "�ͼӳ⵵�� Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 9 Then
        Sbo_Application.StatusBar.SetText "�����ڵ�� �ʼ��Դϴ�. �����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 10 Then
        Sbo_Application.StatusBar.SetText "�ش翩�δ� �ʼ��Դϴ�. �����Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 11 Then
        Sbo_Application.StatusBar.SetText "�ֹε�Ϲ�ȣ�� Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 12 Then
        Sbo_Application.StatusBar.SetText "�ڻ��ڵ�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "MatrixSpaceLineDel Error:" & Err.Number & Space$(10) & Err.Description, bmt_Short, smt_Error
    
    End If
    MatrixSpaceLineDel = False
End Function
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
                If pval.ItemUID = "1" Then
                        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                        '//�߰��� �����̽ÿ�
                        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                        If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                            If oForm.Mode = fm_ADD_MODE Then
                                If MDC_SetMod.Value_ChkYn("[@ZPY506H]", "U_JSNYER", "'" & oForm.Items("JSNYER").Specific.String & "'", " AND U_MSTCOD = '" & oForm.Items("MSTCOD").Specific.String & "'") = False Then
                                    Sbo_Application.StatusBar.SetText "�̹� ����Ǿ��� �ִ� ����� ����� ��ġ�մϴ�.", bmt_Short, smt_Error
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        If Trim$(oDS_ZPY506H.GetValue("U_ENDCHK", 0)) = "Y" And Trim$(oOLDCHK) = "Y" Then
                            Sbo_Application.StatusBar.SetText "��� �ڷ��Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                             BubbleEvent = False
                             Exit Sub
                        ElseIf MatrixSpaceLineDel = False Then
                                BubbleEvent = False
                            End If
                        End If
                ElseIf pval.ItemUID = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then '/ ChooseBtn�������Ʈ
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            Else
                If pval.ItemUID = "1" And pval.ActionSuccess = True And oForm.Mode = fm_ADD_MODE Then
                    oForm.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                End If
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
            If pval.FormUID = oForm.uniqueID And pval.BeforeAction = True And Last_Item = "Mat1" And Col_Last_Uid = "Col1" And Col_Last_Row > 0 _
              And (Col_Last_Uid <> pval.ColUID Or Col_Last_Row <> pval.Row) And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" Then
                If Col_Last_Row > oMat1.VisualRowCount Then
                    Exit Sub
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True And _
            (pval.ItemUID = "MSTCOD" Or pval.ItemUID = "JSNYER") Then
                FlushToItemValue pval.ItemUID
            ElseIf pval.BeforeAction = False And pval.ItemChanged = True And pval.ItemUID = "Mat1" _
            And (pval.ColUID = "Col3" Or pval.ColUID = "Col4" Or pval.ColUID = "Col1" Or pval.ColUID = "Col2" Or pval.ColUID = "Col11") Then
                FlushToItemValue pval.ColUID, pval.Row
            End If
        
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            '//�߰���忡�� �ڵ��̺�Ʈ�� �ڵ忡�� �Ͼ� ������
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
             If pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE Then
                If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(pval.ItemUID).Specific.String & "'", "") = True Then
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                Else
                    If oMat1.RowCount > 0 Then
                        oMat1.Columns("Col1").Cells(oMat1.VisualRowCount).CLICK ct_Regular
                        BubbleEvent = False
                    End If
                End If
             ElseIf pval.BeforeAction = True And pval.ColUID = "Col1" And pval.CharPressed = 9 Then
                If Trim$(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = "" Then
                        '1-����û�� �ڷ�� ����ڹ�ȣ �ʿ����.
                Else
                    '/ ����ڹ�ȣ üũ
                    If Trim$(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) <> "" And Len(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) <= 12 Then
                        If MDC_SetMod.TaxNoCheck(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = False Then
                            Call Sbo_Application.StatusBar.SetText("����ڹ�ȣ�� Ʋ���ϴ�. Ȯ���ϼ���.", bmt_Short, smt_Error)
                            BubbleEvent = False
                        End If
                    End If
                End If
             ElseIf pval.BeforeAction = True And pval.ColUID = "Col6" And pval.CharPressed = 9 Then
                If Trim$(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = "" Then
                    Call Sbo_Application.StatusBar.SetText("����� �ֹε�Ϲ�ȣ�� �ʼ��Դϴ�. �Է��ϼ���.", bmt_Short, smt_Error)
                    BubbleEvent = False
                Else
                    '/ �ֹε�Ϲ�ȣ üũ
                    If GovIDCheck(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = False Then
                        Call Sbo_Application.StatusBar.SetText("�ֹε�Ϲ�ȣ�� Ʋ���ϴ�. Ȯ���ϼ���.", bmt_Short, smt_Error)
                        BubbleEvent = False
                    End If
                End If
             End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUID
                    Col_Last_Row = pval.Row
                    Col_Last_Uid = pval.ColUID
                End If
            Else
                Last_Item = pval.ItemUID
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
                Set oDS_ZPY506H = Nothing
                Set oDS_ZPY506L = Nothing
                Set oMat1 = Nothing
            End If
'et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_MATRIX_LOAD:
            If pval.BeforeAction = False Then
                FormItemEnabled
                Call Matrix_AddRow(oMat1.VisualRowCount)
            End If
            
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:", bmt_Short, smt_Error
End Sub
'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oRow As Long, Optional Insert_YN As Boolean)
        If Insert_YN = False Then
            oDS_ZPY506L.InsertRecord (oRow)
        End If
        oDS_ZPY506L.Offset = oRow
        oDS_ZPY506L.setValue "U_LineNum", oRow, oRow + 1
        oDS_ZPY506L.setValue "U_MEDNAM", oRow, ""
        oDS_ZPY506L.setValue "U_MEDNBR", oRow, ""
        oDS_ZPY506L.setValue "U_MEDCNT2", oRow, ""
        oDS_ZPY506L.setValue "U_MEDAMT2", oRow, ""
        oDS_ZPY506L.setValue "U_MEDCNT", oRow, ""
        oDS_ZPY506L.setValue "U_MEDAMT", oRow, ""
        oDS_ZPY506L.setValue "U_GWANGE", oRow, ""
        oDS_ZPY506L.setValue "U_MEDCOD", oRow, ""
        oDS_ZPY506L.setValue "U_FAMNAM", oRow, ""
        oDS_ZPY506L.setValue "U_PERNBR", oRow, ""
        oDS_ZPY506L.setValue "U_DAECHK", oRow, ""
        oDS_ZPY506L.setValue "U_INTGBN", oRow, "1"
        oMat1.LoadFromDataSource
End Sub
'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i&
    
    If pval.BeforeAction = True Then
        Select Case pval.MenuUID
            Case "1283" '/ ����
                If Trim$(oDS_ZPY506H.GetValue("U_ENDCHK", 0)) = "Y" Then
                    Sbo_Application.StatusBar.SetText "��� �ڷ��Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                    BubbleEvent = False
                    Exit Sub
                Else
                    If Sbo_Application.MessageBox("���� ȭ�鳻����ü�� ���� �Ͻðڽ��ϱ�? ������ �� �����ϴ�.", 2, "Yes", "No") = 2 Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            Case "1288", "1289", "1290", "1291"
                Call AuthorityCheck(oForm, "CLTCOD", "@ZPY506H", "DocNum")      '//������ ���ѿ� ���� ����� ����
            
            Case Else
                Exit Sub
        End Select
    Else
    
        Select Case pval.MenuUID
            Case "1287" '/ ����
            Case "1283" '/ ����
                 FormItemEnabled
            Case "1281", "1282":
                FormItemEnabled
                If pval.MenuUID = "1282" Then
                    FormClear
                    Call Matrix_AddRow(0, True)
                    oForm.Items("JSNYER").CLICK ct_Regular
                End If
            Case "1288" To "1291":
            Case "1293"
                If oMat1.RowCount <> oMat1.VisualRowCount Then
                    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                    '//�Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
                    '//�̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
                    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
                    For i = 0 To oMat1.VisualRowCount - 1
                        oMat1.Columns("Col0").Cells(i + 1).Specific.Value = i + 1
                    Next i
                    
                    oMat1.FlushToDataSource
                    oDS_ZPY506L.RemoveRecord oDS_ZPY506L.Size - 1   '// Mat1�� ����������(�����) ����
                    oMat1.Clear
                    oMat1.LoadFromDataSource
                    
                End If
                Call FlushToItemValue("Col3", 1)
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
' .srf ���Ϸκ��� ���� �ε��Ѵ�
'*******************************************************************
Public Sub LoadForm(Optional JSNYER$, Optional MSTCOD$, Optional CLTCOD$)
'Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY506.srf")
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
    
    oFormUniqueID = "ZPY506_" & GetTotalFormsCount
    
    '�� �Ҵ�
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '�÷��ǿ� ���� ��´�   **�÷����̶� ��ü�� ��� ���� �迭�μ� ���⼭�� Ȱ��ȭ�Ǿ��� �ִ� ���� ��� �ִ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    AddForms Me, oFormUniqueID, "ZPY506"
    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    CreateItems
    
    oForm.EnableMenu ("1293"), True  '/ �����
    oForm.EnableMenu ("1283"), True  '/ ����
    oForm.EnableMenu ("1284"), False  '/ ���
    
    
    If JSNYER <> "" Then
        Call ShowSource(JSNYER, MSTCOD, CLTCOD)
    End If
    
    oForm.Freeze False
    oForm.Update
    'oForm.Visible = True

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
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oColumn         As SAPbouiCOM.Column
    Dim oCheck          As SAPbouiCOM.CheckBox
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_ZPY506H = oForm.DataSources.DBDataSources("@ZPY506H")
    Set oDS_ZPY506L = oForm.DataSources.DBDataSources("@ZPY506L")
    
    Set oMat1 = oForm.Items("Mat1").Specific
    
    '// �����
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
    Loop
    
    '// �����ڵ�
    Set oColumn = oMat1.Columns("Col5")
    oColumn.ValidValues.Add "0", "����"
    oColumn.ValidValues.Add "1", "�ҵ����� ��������"
    oColumn.ValidValues.Add "2", "������� ��������"
    oColumn.ValidValues.Add "3", "�����"
    oColumn.ValidValues.Add "4", "������"
    oColumn.ValidValues.Add "5", "�����ڸ�"
    oColumn.ValidValues.Add "6", "��Ÿ"
 
    '// ���ε��ش翩��
    Set oColumn = oMat1.Columns("Col7")
    oColumn.ValidValues.Add "1", "����,���,65���̻���"
    oColumn.ValidValues.Add "2", "�׿ܱ⺻�������"
 
    '// ���ܱ����ڵ�
    Set oColumn = oMat1.Columns("Col10")
    oColumn.ValidValues.Add "1", "������"
    oColumn.ValidValues.Add "9", "�ܱ���"
 
    '// �����ڵ�
    Set oColumn = oMat1.Columns("Col12")
    oColumn.ValidValues.Add "1", "����û���� �����ϴ� �Ƿ�� �ڷ�"
    oColumn.ValidValues.Add "2", "���ΰǰ���������� �Ƿ��δ������"
    oColumn.ValidValues.Add "3", "������꼭, �������꼭"
    oColumn.ValidValues.Add "4", "�����޿���� ������"
    oColumn.ValidValues.Add "5", "��Ÿ �Ƿ�� ������"
    
    oMat1.Columns("Col8").Visible = False
    oMat1.Columns("Col9").Visible = False
 
   '/ Check ��ư
    Set oCheck = oForm.Items("ENDCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"

    Set oCheck = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oCheck = Nothing
    Set oColumn = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error:" & Err.Description, bmt_Short, smt_Error
End Sub
