Attribute VB_Name = "MDC_SetMod"
'//////////////////////////////////////////////////////////////////////////////////////
'// ������: ������� �븮 �迵ȣ                                                     //
'// �Ⱓ: 2005.4~2005.5 ����                                                         //
'// ���������Ʈ: Sap�ڸ��� ���Ĵٵ尳��                                             //
'// �̸�:neverdie74@hanmail.net                                                      //
'//////////////////////////////////////////////////////////////////////////////////////
Option Explicit

Public Function Value_ChkYn(Tablename$, ColumnName$, Key_Str$, Optional And_Line$) As Boolean
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '���̺��� ������ ���� �Է°��� �����ϴ����� üũ�Ѵ�
    '�μ�����:���̺��̸�,�÷��̸�,���縦 Ȯ���ؾ� �ϴ�Ű��,�÷��� ������ Ÿ��
    '���� �÷��� ����Ÿ���ϰ�찡 �ƴϸ� Key_Str�� �յڿ� "'"�� �ٿ� �־�� �Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    Dim s_Recordset      As SAPbobsCOM.Recordset
    Dim sSQL$
    Dim Count_Chk%

    If Key_Str <> "" Then
        sSQL = "SELECT count(*) FROM " + Tablename + " Where " + ColumnName + "=" + CStr(Key_Str) + ""
        sSQL = sSQL + And_Line

        Set s_Recordset = Sbo_Company.GetBusinessObject(BoRecordset)
        s_Recordset.DoQuery sSQL

        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        '�������� �������� Ȯ��
        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        Count_Chk = s_Recordset.Fields(0).VALUE

        If Count_Chk > 0 Then
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            '������ ���� Ű������ ������ ����
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            Value_ChkYn = False
        Else
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            '�������� �ʴ°�
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            Value_ChkYn = True
        End If

        Set s_Recordset = Nothing
    Else
        Value_ChkYn = True
    End If
End Function


'///////////////////////////
'--------------------------
'/��ġ�� ���� �׸����� üũ
'--------------------------
'//////////////////////////
Public Function BachUseChk(ItemCode$) As Boolean
    Dim sSQL         As String
    Dim oRecordSet   As SAPbobsCOM.Recordset

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    sSQL = ""
    sSQL = sSQL & " SELECT  ManBtchNum "
    sSQL = sSQL & " FROM    [OITM]  "
    sSQL = sSQL & " WHERE   ItemCode = '" & ItemCode & "' "

    oRecordSet.DoQuery sSQL
    
    If oRecordSet.Fields("ManBtchNum").VALUE = "Y" Then
        BachUseChk = True
    Else
        BachUseChk = False
    End If

    Set oRecordSet = Nothing

End Function

'///////////////////////////
'--------------------------
'/�ø����� ���� �׸����� üũ
'--------------------------
'//////////////////////////
Public Function SerialUseChk(ItemCode$) As Boolean
    Dim sSQL         As String
    Dim oRecordSet   As SAPbobsCOM.Recordset

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    sSQL = ""
    sSQL = sSQL & " SELECT  ManSerNum "
    sSQL = sSQL & " FROM    [OITM]  "
    sSQL = sSQL & " WHERE   ItemCode = '" & ItemCode & "' "

    oRecordSet.DoQuery sSQL
    
    If oRecordSet.Fields(0).VALUE = "Y" Then
        SerialUseChk = True
    Else
        SerialUseChk = False
    End If

    Set oRecordSet = Nothing

End Function

'///////////////////////////////
'------------------------------
'/ �ø��� Set
'------------------------------
'//////////////////////////////
Public Function SetSerial(ByVal ItemCode$, ByVal Qty As Long, ByRef oGenEntrySerialNumbers As SAPbobsCOM.SerialNumbers) As String
    Dim sSQL         As String
    Dim SerialNum    As String
    Dim SerialHead   As String
    Dim SerialSeq    As String
    Dim SeqLen       As Long
    Dim SerialInit   As String
    Dim YearWeek     As String
    Dim i, j         As Long
    Dim oRecordSet   As SAPbobsCOM.Recordset

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    '--------------------------
    '/ ǰ���� �̴ϼ��� ��ȸ
    '---------------------------
        
    sSQL = ""
    sSQL = sSQL & " SELECT  U_ManserNum "
    sSQL = sSQL & " FROM    [OITM]  "
    sSQL = sSQL & " WHERE   ItemCode = '" & ItemCode & "' "

    oRecordSet.DoQuery sSQL
    
    SerialInit = oRecordSet.Fields(0).VALUE

    '-------------------------
    '/ Year(2), Week
    '-------------------------
    sSQL = ""
    sSQL = sSQL & " SELECT   CONVERT(VARCHAR(2), RIGHT(DATEPART(year, GETDATE()),2)) "
    sSQL = sSQL & "       +  CONVERT(VARCHAR(2), DATEPART(week, GETDATE())) AS 'YearWeek'"
    
    oRecordSet.DoQuery sSQL

    YearWeek = oRecordSet.Fields(0).VALUE

    '--------------------------
    '/ ������ ����� ���� Max
    '--------------------------
    sSQL = ""
    sSQL = sSQL & " SELECT  Max(SuppSerial)"
    sSQL = sSQL & " FROM    [OSRI]    "
    sSQL = sSQL & " WHERE   SuppSerial like '" & SerialInit & YearWeek & "%'"

    oRecordSet.DoQuery sSQL
    
    SerialNum = Trim(oRecordSet.Fields(0).VALUE)
        
    If SerialNum = "" Then
        '/ ������ �ش��ϴ� �ø����� ���°��
        
        For i = 1 To Qty
            If i > 1 Then
                oGenEntrySerialNumbers.Add
            End If
            
            SerialSeq = i
            
            For j = 1 To 4 - Len(SerialSeq)
                SerialSeq = "0" & SerialSeq
            Next
            
            oGenEntrySerialNumbers.ManufacturerSerialNumber = SerialInit & YearWeek & SerialSeq
        
            If i = 1 Then SetSerial = SerialInit & YearWeek & SerialSeq
        Next
    Else
        '/ ������ �ø����� �ִ°��
        SerialHead = Mid(SerialNum, 1, Len(SerialNum) - 4)
        
        SerialSeq = Val(Right(SerialNum, 4)) + 1
        
        For i = 1 To Qty
            If i > 1 Then
                oGenEntrySerialNumbers.Add
            End If
        
            For j = 1 To 4 - Len(SerialSeq)
                SerialSeq = "0" & SerialSeq
            Next
            
            oGenEntrySerialNumbers.ManufacturerSerialNumber = SerialHead & SerialSeq
                
            If i = 1 Then SetSerial = SerialHead & SerialSeq
        
            SerialSeq = Val(SerialSeq) + 1
        Next
        
        
    End If
    
    'GetSerial = SerialNum
    
'    If Trim(GetSerial) = "" Then
'        GetSerial = "F"
'    End If
    
    Set oRecordSet = Nothing
End Function
Public Function Get_TitleNameQC(LCode$, SD380Name$) As String
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    'ǰ�������� text�� ��������
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    sQry = "        SELECT  U_Text  "
    sQry = sQry & " FROM    [dbo].[@ZSY004L] T1 "
    sQry = sQry & "         INNER JOIN "
    sQry = sQry & "         [dbo].[@ZSY004H] T2  "
    sQry = sQry & "             ON  T1.Code = T2.Code "
    sQry = sQry & " WHERE   T1.U_Language = " & Z_Language & " "
    sQry = sQry & "         AND   T2.U_TextID = '" & LCode & "' "

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Get_TitleNameQC = oRecordSet(0).VALUE
    Else
        Get_TitleNameQC = SD380Name
    End If

    Set oRecordSet = Nothing
End Function

Public Function gCryReport_Action(RptTitle$, RptName$, SRptChk$, rQry$, Optional RptCnt$, Optional FormulaChk$ = "N", Optional ActionT$, Optional ByVal ExportString As String) As Boolean
'/***********************************************************************/
'// ��� : CRYSTALREPORT VER10 ���
'// Creator        : Ham Mi Kyoung
'// ���� : RptTitle- �̸�����âŸ��Ʋ, RptName-����Ʈ��,SRptChk-���긮��Ʈ�������(Y/N),  rQry-����Ʈ���ǹ�,
'//        RptCnt - ����� �̸�����â ����, FormulaChk-Formula�������(Y/N),
'//        ActionT(P/V)-P:�̸�����â���� �ٷ� �μ�,V-�̸�����
'// ������ �ʿ��� ��� �ݵ�� ǥ��ȭ���� �����Ұ�!
'// Copyright  (c) Morning Data
'/***********************************************************************/
    Dim i           As Integer
    Dim j           As Integer
    Dim K           As Integer
    Dim l           As Integer
    Dim x           As Integer
    Dim y           As Integer
    Dim ErrNum      As Integer
    Dim FormulaCnt  As Integer
    Dim SubReptCnt  As Integer
    Dim sFormulaCnt As Integer
    
    ErrNum = 0
    '/ Check
    FormulaCnt = UBound(gRpt_Formula)
    If SRptChk = "Y" Then
        SubReptCnt = UBound(gRpt_SRptName)
    End If
    
    Set g_ADORS1 = New ADODB.Recordset
    Set g_ADORS2 = New ADODB.Recordset
    
    
    g_ADORS1.Open rQry, ZG_CRWDSN, adOpenKeyset, adLockBatchOptimistic
'    If g_ADORS1.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo error_Message
'    End If
 '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    Set g_CApp = New CRAXDDRT.Application
    If RptCnt = "" Or RptCnt = "1" Then
        Set g_GCrview = frmRPT_View1.CRViewer1
    ElseIf RptCnt = "2" Then
        Set g_GCrview = frmRPT_View2.CRViewer1
    ElseIf RptCnt = "3" Then
        Set g_GCrview = frmRPT_View3.CRViewer1
    End If
    Set g_Report = g_CApp.OpenReport(ShareFolderPath & "ReportPS\" & RptName)
    g_Report.Database.Tables.Item(1).SetDataSource g_ADORS1, 3
    g_Report.DiscardSavedData
    g_Report.VerifyOnEveryPrint = True
    
    '//����Ʈ ������� ���ο� ����
    
    If MDC_PS_Common.GetValue("SELECT U_PRTYN FROM [OHEM] WHERE U_MSTCOD = '" & MDC_PS_Common.User_MSTCOD & "'", 0, 1) = "Y" Then
        g_Report.PrinterSetup (0)
    End If
'/ SubReport /
    If SRptChk = "Y" Then
        Set g_CrSections = g_Report.Sections
         For i = 1 To g_CrSections.Count
            Set g_CrSection = g_CrSections.Item(i)
            Set g_CrReportObjs = g_CrSection.ReportObjects
            For K = 1 To g_CrReportObjs.Count
                If g_CrReportObjs.Item(K).Kind = crSubreportObject Then
                    Set g_CrSubReportObj = g_CrReportObjs.Item(K)
                    Set g_CrSubReport = g_CrSubReportObj.OpenSubreport
                    For j = 1 To SubReptCnt
                        If g_CrSubReportObj.SubreportName = Trim(gRpt_SRptName(j)) Then
                            g_ADORS2.Open gRpt_SRptSqry(j), g_ERPDMS, adOpenKeyset, adLockBatchOptimistic
                            g_CrSubReport.Database.Tables.Item(1).SetDataSource g_ADORS2, 3
                            g_ADORS2.Close
                        '/ SubFormula //////////////////////////////////////////////////////////////
                            If gRpt_SFormula(j, 1) <> "" Then
                                 g_CrSubReport.FormulaSyntax = crCrystalSyntaxFormula
                                 For x = 1 To g_CrSubReport.FormulaFields.Count
                                    Set g_cFormula = g_CrSubReport.FormulaFields.Item(x)
                                    sFormulaCnt = UBound(gRpt_SFormula, 2)
                                    For y = 1 To sFormulaCnt
                                        If g_cFormula.FormulaFieldName = Trim(gRpt_SFormula(j, y)) Then
                                            g_cFormula.Text = "'" & gRpt_SFormula_Value(j, y) & "'"
                                        End If
                                    Next y
                                Next x
                            End If
                        '///////////////////////////////////////////////////////////////////////////
                        End If
                    Next j
                End If
            Next K
        Next i
    End If
    
 '/ Formula /
    If FormulaCnt >= 1 Then
         g_Report.FormulaSyntax = crCrystalSyntaxFormula
         For i = 1 To g_Report.FormulaFields.Count
            Set g_cFormula = g_Report.FormulaFields.Item(i)
            
            For K = 1 To FormulaCnt
                 If g_cFormula.FormulaFieldName = Trim(gRpt_Formula(K)) Then
                     g_cFormula.Text = "'" & gRpt_Formula_Value(K) & "'"
                 End If
            Next K
        Next i
    End If
    
    'PDF Export_S(2017.05.17 �۸�� �߰�)
    If ExportString <> "" Then '�ܺ����Ϸ� Export�� �̸����� â ȣ�� ����

        g_Report.ExportOptions.PDFExportAllPages = True
        g_Report.ExportOptions.DestinationType = crEDTDiskFile
        g_Report.ExportOptions.DiskFileName = ExportString
        g_Report.ExportOptions.FormatType = crEFTPortableDocFormat
        
        Call g_Report.Export(False)
    'PDF Export_E(2017.05.17 �۸�� �߰�)
    Else
    
        '/ Report Viewer Show /
        If ActionT = "P" Then
            g_Report.PrintOut False
        Else
            If RptCnt = "" Or RptCnt = "1" Then
                frmRPT_View1.Show
                frmRPT_View1.Caption = RptTitle
            ElseIf RptCnt = "2" Then
                frmRPT_View2.Show
                frmRPT_View2.Caption = RptTitle
            ElseIf RptCnt = "3" Then
                frmRPT_View3.Show
                frmRPT_View3.Caption = RptTitle
            End If
        '/ Action
           g_GCrview.ReportSource = g_Report
           g_GCrview.ViewReport
           g_GCrview.Zoom (100)
        End If

    End If
    
 '/ Init_Crystal
    Set g_CApp = Nothing
    Set g_GCrview = Nothing
    Set g_Report = Nothing
    Set g_CrSections = Nothing
    Set g_CrSection = Nothing
    Set g_CrReportObjs = Nothing
    Set g_CrSubReportObj = Nothing
    Set g_CrSubReport = Nothing
    Set g_cFormula = Nothing
    Set g_ADORS1 = Nothing
    Set g_ADORS2 = Nothing
 '/ End
    gCryReport_Action = True
    Exit Function
'/ Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
error_Message:
    Set g_CApp = Nothing
    Set g_GCrview = Nothing
    Set g_Report = Nothing
    Set g_CrSections = Nothing
    Set g_CrSection = Nothing
    Set g_CrReportObjs = Nothing
    Set g_CrSubReportObj = Nothing
    Set g_CrSubReport = Nothing
    Set g_cFormula = Nothing
    Set g_ADORS1 = Nothing
    Set g_ADORS2 = Nothing

    If ErrNum = 1 Then
        Sbo_Application.SetStatusBarMessage "��ȸ�ڷᰡ �����ϴ�.", bmt_Short, True
    Else
        Sbo_Application.SetStatusBarMessage "Print_Query : " & Space(10) & Err.Description, bmt_Short, True
    End If
gCryReport_Action = False

End Function

Public Function SetDrive(ByVal ServerIP$, ByVal ShrName$, ByVal UserID$, ByVal PWD$, Optional ByVal DriverName$ = "W:") As Long
    Dim ws
    Dim oDrives
    Dim i        As Long
    Dim IpName   As String
    Dim result   As Long

On Error GoTo Err

    '/ �̹� ����� ��Ʈ��ũ ����̹��� �ִٸ� ����.
    Set ws = CreateObject("WScript.Network")

    Set oDrives = ws.EnumNetworkDrives

    For i = 0 To oDrives.Count - 1 Step 2
         ws.RemoveNetworkDrive oDrives.Item(i)
    Next

    '��Ʈ��ũ ����̺� ����
    'ex)ws.MapNetworkDrive "[���õ���̺�������]", "[���ݵ���̺��]", false, "[���̵�]", "[��ȣ]"
    '���� �۾��� ini�� �ҽ��� �ھƼ� ��� �ϸ� �ڵ����� ����̹��� ���� ����� �Ǳ���
    ws.MapNetworkDrive DriverName, ServerIP & ShrName, False, UserID, PWD
   
    SetDrive = 0

    Sbo_Application.SetStatusBarMessage "�������� ������ �Ǿ����ϴ� !.", bmt_Short, False
    Exit Function
Err:
    SetDrive = -1
End Function

'///////////////////////////
'--------------------------
'/��ġ�� ���� �׸����� üũ
'--------------------------
'//////////////////////////
Public Function BatchUseChk(ItemCode$) As Boolean
    Dim sSQL         As String
    Dim oRecordSet   As SAPbobsCOM.Recordset

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    sSQL = ""
    sSQL = sSQL & " SELECT  ManBtchNum "
    sSQL = sSQL & " FROM    [OITM]  "
    sSQL = sSQL & " WHERE   ItemCode = '" & ItemCode & "' "

    oRecordSet.DoQuery sSQL
    
    If oRecordSet.Fields("ManBtchNum").VALUE = "Y" Then
        BatchUseChk = True
    Else
        BatchUseChk = False
    End If

    Set oRecordSet = Nothing

End Function

Public Function BatchOpenQtyChk(ItemCode$, UseQty As Double, Optional WH$) As Boolean
    Dim sSQL         As String
    Dim oRecordSet   As SAPbobsCOM.Recordset

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    sSQL = ""
    sSQL = sSQL & " SELECT  SUM(Quantity) "
    sSQL = sSQL & " FROM    [OIBT]  "
    sSQL = sSQL & " WHERE   ItemCode = '" & ItemCode & "' "
    
    If WH <> "" Then
        sSQL = sSQL & " AND   WhsCode = '" & WH & "' "
    End If
    
    sSQL = sSQL & " GROUP BY  ItemCode "

    oRecordSet.DoQuery sSQL
    
    If oRecordSet.Fields(0).VALUE >= UseQty Then
        BatchOpenQtyChk = True
    Else
        BatchOpenQtyChk = False
    End If

    Set oRecordSet = Nothing

End Function

Public Sub Remove_ComboList(Lst As Object)
    Dim i As Long
    Dim ComBox          As New SAPbouiCOM.ComboBox
    Set ComBox = Lst
    For i = 1 To ComBox.ValidValues.Count
        Call ComBox.ValidValues.Remove(0, psk_Index)
    Next
    Set ComBox = Nothing
End Sub

Public Sub Set_ComboList(Lst As Object, sSQL As String, Optional TValue As String = "", Optional Reset As Boolean, Optional SetF As Boolean)
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    'ComboBox Object,Query,SD380 value
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    Dim ComBox      As New SAPbouiCOM.ComboBox
    Dim s_Recordset As SAPbobsCOM.Recordset
    
    Set s_Recordset = Sbo_Company.GetBusinessObject(BoRecordset)
    s_Recordset.DoQuery sSQL
    Set ComBox = Lst
    
    If Reset = True Then
        While ComBox.ValidValues.Count > 0
            Call ComBox.ValidValues.Remove(0, psk_Index)
        Wend
    End If
    
    If SetF = True Then
        ComBox.ValidValues.Add "", ""
    End If
    
    While Not s_Recordset.EOF
        ComBox.ValidValues.Add CStr(s_Recordset.Fields(0).VALUE), CStr(s_Recordset.Fields(1).VALUE)    'Value,Description
        s_Recordset.MoveNext
    Wend
    
    If TValue <> "" Then
        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        'Sets SD380 Value
        '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        ComBox.Select TValue, psk_ByDescription
    End If
    Set ComBox = Nothing
    Set s_Recordset = Nothing
End Sub

