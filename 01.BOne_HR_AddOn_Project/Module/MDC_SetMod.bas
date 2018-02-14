Attribute VB_Name = "MDC_SetMod"
'//Function List
'Value_ChkYn                        ���̺��� ������ ���� �Է°��� �����ϴ����� üũ�Ѵ�
'Get_ReData                         �ܼ����� �����Ѵ�.
'MDC_ActUserTableWithReturnValue    User Defined Window�� Pop-up Windowó�� ȣ���ϰ�,
'                                   Code List��������� �θ�â�� �ݿ��ϱ� ���Ͽ� ������ ������ �θ���踦 �����Ѵ�
'MDC_CF_DBDatasourceReturn          CHOOSEFROMLIST�� ���� ����(�ھ�)
'MDC_CH_UserDatasourceReturn        CHOOSEFROMLIST�� ���� ����(���������)
'gCryReport_Action                  ũ����Ż����Ʈ ���
'MDC_GF_Message                     �޼��� ���
'MDC_GF_Nz                          NULL �� üũ
'SetMessage_Err                     �����޼��� ���
'SetMessage_Ok                      �����޼��� ���
'SetDrive                           ��Ʈ��ũ ����̺� ����
'RemoveBlankRecord                  UDO���� �߻��ϴ� ���ڵ带 �ʼ����� üũ �� �����
'Get_CustTitle                      ������ text�� ��������
'SetMatrix_Sorter                   Matrix & Grid �������

Option Explicit

'// Combo Box Setting
Public Sub GP_SetComboList(fCombo As SAPbouiCOM.ComboBox, _
                                fSQL As String, _
                                Optional fComboAdd As Boolean = False)
    'Function ID : GetListIndex
    '��    ��    :
    '��    ��    : Lst
    '�� ȯ ��    : None
    'Ư�̻���    : �޺��ڽ��� ���� �� ������ �ý��� �ڵ忡�� ������ �����Ѵ�
    Dim fRecordset As SAPbobsCOM.Recordset
    
    Set fRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    fRecordset.DoQuery fSQL
    
    If fComboAdd = True Then
        Call fCombo.ValidValues.Add("", "")
    End If
    
    Do Until fRecordset.EOF
        Call fCombo.ValidValues.Add(fRecordset.Fields(0).VALUE, fRecordset.Fields(1).VALUE)
        fRecordset.MoveNext
    Loop
    
    Set fRecordset = Nothing
    
End Sub


Public Function Get_ReData(oReColumn$, oColumn$, oTable$, oTaValue$, Optional AndLine$) As Variant
    '------------------------------------------------
    '��ȯ�÷�,���� �÷�,���̺�,���ǰ�,AND ��
    '------------------------------------------------
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    sQry = "SELECT " & oReColumn & " FROM " & oTable
    sQry = sQry & " WHERE " & oColumn & " = " & oTaValue
    If AndLine <> "" Then
        sQry = sQry & AndLine
    End If
    oRecordSet.DoQuery sQry

    Do Until oRecordSet.EOF
        Get_ReData = oRecordSet(0).VALUE
        oRecordSet.MoveNext
    Loop

    Set oRecordSet = Nothing
End Function

'---------------------------------------------------------------------------------------
''//    CHOOSEFROMLIST�� ���� ����
'       MDC_GP_ChooseFromList_DBDatasourceReturn(PVAL, FORMUID, ���̺��̸�, ������ �÷�,
'       MATRIX, ���� ROW, ���ι�ȣ�÷�, üũ�ڽ��� ��� �÷���, üũ�ڽ� �ʱⰪ)
'---------------------------------------------------------------------------------------
Public Sub MDC_CF_DBDatasourceReturn(pval As SAPbouiCOM.IItemEvent, _
                                                      MDC_pFormUID As String, _
                                                      MDC_pTableName As String, _
                                                      Optional ByVal MDC_sUDS As String = "", _
                                                      Optional ByVal MDC_pMatrix As String = "", _
                                                      Optional ByVal MDC_pRow As Integer = 0, _
                                                      Optional ByVal MDC_pSeqNoUDS As String = "", _
                                                      Optional ByVal MDC_pFieldName As String = "", _
                                                      Optional ByVal MDC_pFieldValue As String = "")

    Dim MDC_oCFLEvento  As SAPbouiCOM.IChooseFromListEvent
    Dim MDC_sCFLID      As String
    Dim MDC_oCFL        As SAPbouiCOM.ChooseFromList
    Dim MDC_oDataTable  As SAPbouiCOM.DataTable
    
    Dim MDC_pForm       As SAPbouiCOM.Form
    Dim MDC_oMatrix     As SAPbouiCOM.Matrix
    Dim MDC_oDBTable    As SAPbouiCOM.DBDataSource
    
    Dim MDC_iLooper     As Integer
    Dim MDC_jLooper     As Integer
    Dim MDC_sTemp01
    Dim MDC_sTemp02
    
    
    Set MDC_pForm = Sbo_Application.Forms.Item(MDC_pFormUID)

    Set MDC_oCFLEvento = pval
    
    Set MDC_oDataTable = MDC_oCFLEvento.SelectedObjects
    
    MDC_sCFLID = MDC_oCFLEvento.ChooseFromListUID
    '// ��ҹ�ư Ŭ����
    If MDC_oDataTable Is Nothing Then
        Exit Sub
    End If
    
    Set MDC_oCFL = MDC_pForm.ChooseFromLists.Item(MDC_sCFLID)
    
    Set MDC_oDBTable = MDC_pForm.DataSources.DBDataSources.Item(MDC_pTableName)

    If MDC_pMatrix <> "" Then Set MDC_oMatrix = MDC_pForm.Items(MDC_pMatrix).Specific
    MDC_sTemp01 = Split(MDC_sUDS, ",")
    
    If MDC_pMatrix <> "" And MDC_pRow > 0 Then
    
        For MDC_jLooper = 0 To MDC_oDataTable.Rows.Count - 1
            
            If MDC_jLooper > 0 Then
                If MDC_pSeqNoUDS <> "" Then
                    MDC_oDBTable.InsertRecord (MDC_pRow + MDC_jLooper - 1)
                    MDC_oDBTable.Offset = MDC_pRow + MDC_jLooper - 1
                    MDC_oDBTable.setValue MDC_pSeqNoUDS, MDC_pRow + MDC_jLooper - 1, MDC_pRow + MDC_jLooper
                Else
                    MDC_oDBTable.InsertRecord (MDC_pRow + MDC_jLooper - 1)
                    MDC_oDBTable.Offset = MDC_pRow + MDC_jLooper - 1
                End If
            Else
                MDC_oDBTable.Offset = MDC_pRow + MDC_jLooper - 1
            End If
            
            For MDC_iLooper = 0 To UBound(MDC_sTemp01)
                If MDC_oCFL.ObjectType = "171" Then   '// �������Ÿ�ϰ�� �� + �̸�
                    If MDC_iLooper = 0 Then
                        MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), MDC_pRow + MDC_jLooper - 1, MDC_oDataTable.GetValue("U_MSTCOD", MDC_jLooper)
                    ElseIf MDC_iLooper = 1 Then
                        MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), MDC_pRow + MDC_jLooper - 1, MDC_oDataTable.GetValue("U_FULLNAME", MDC_jLooper)
                    ElseIf MDC_iLooper = 2 Then
                        MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), MDC_pRow + MDC_jLooper - 1, MDC_oDataTable.GetValue("U_TeamCode", MDC_jLooper)
                    ElseIf MDC_iLooper = 3 Then
                        MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), MDC_pRow + MDC_jLooper - 1, MDC_oDataTable.GetValue("U_TeamCode", MDC_jLooper)
                    End If
                Else
                    MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), MDC_pRow + MDC_jLooper - 1, MDC_oDataTable.GetValue(MDC_iLooper, MDC_jLooper)
                End If
            Next MDC_iLooper
            
            If MDC_pFieldName <> "" And MDC_pFieldValue <> "" Then MDC_oDBTable.setValue MDC_pFieldName, MDC_pRow + MDC_jLooper - 1, MDC_pFieldValue
            
            MDC_oMatrix.LoadFromDataSource
        Next MDC_jLooper
    Else
        MDC_sTemp02 = ""
        For MDC_jLooper = 0 To MDC_oDataTable.Rows.Count - 1
            For MDC_iLooper = 0 To UBound(MDC_sTemp01)
            
                Select Case MDC_oCFL.ObjectType
                Case "171"

                Case Else
                    MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), 0, MDC_oDataTable.GetValue(MDC_iLooper, MDC_jLooper)
                End Select
                
            Next MDC_iLooper
            
        Next MDC_jLooper
        
        
        
    End If
    
    Exit Sub

End Sub

'---------------------------------------------------------------------------------------
''//    CHOOSEFROMLIST�� ���� ����
'       MDC_GP_ChooseFromList_UserDatasourceReturn(PVAL, FORMUID, ���̺��̸�)
'       # ���̺��� ������� ����ϴ� ������
'---------------------------------------------------------------------------------------
Public Sub MDC_CH_UserDatasourceReturn(pval As SAPbouiCOM.IItemEvent, _
                                                      MDC_pFormUID As String, _
                                                      ByVal MDC_sUDS As String, _
                                                      Optional MDC_GirdName As String)

    Dim MDC_oCFLEvento As SAPbouiCOM.IChooseFromListEvent
    Dim MDC_sCFLID     As String
    Dim MDC_oCFL       As SAPbouiCOM.ChooseFromList
    Dim MDC_oDataTable As SAPbouiCOM.DataTable
    
    Dim MDC_pForm      As SAPbouiCOM.Form
    
    Dim MDC_iLooper    As Integer
    Dim MDC_jLooper    As Integer
    Dim MDC_sTemp01
    
    Set MDC_pForm = Sbo_Application.Forms.Item(MDC_pFormUID)

    Set MDC_oCFLEvento = pval
    Set MDC_oDataTable = MDC_oCFLEvento.SelectedObjects
    '// ��ҹ�ư Ŭ����
    If MDC_oDataTable Is Nothing Then
        Exit Sub
    End If
    MDC_sCFLID = MDC_oCFLEvento.ChooseFromListUID
    Set MDC_oCFL = MDC_pForm.ChooseFromLists.Item(MDC_sCFLID)
    
    MDC_sTemp01 = Split(MDC_sUDS, ",")

    For MDC_iLooper = 0 To UBound(MDC_sTemp01)
        Select Case MDC_oCFL.ObjectType
            Case "MDC_MM_CBP101"
                
            Case Else
                MDC_pForm.DataSources.UserDataSources.Item(MDC_sTemp01(MDC_iLooper)).ValueEx = Trim(MDC_oDataTable.GetValue(MDC_iLooper, 0))
        End Select
    Next MDC_iLooper

End Sub


'/***********************************************************************/
'// ��� : CRYSTALREPORT VER10 ���
'// ���� : RptTitle- �̸�����âŸ��Ʋ, RptName-����Ʈ��,SRptChk-���긮��Ʈ�������(Y/N),  rQry-����Ʈ���ǹ�,
'//        RptCnt - ����� �̸�����â ����, FormulaChk-Formula�������(Y/N),
'//        ActionT(P/V)-P:�̸�����â���� �ٷ� �μ�,V-�̸�����, PrintSetup(TRUE,FALSE), F:PDF�� �μ�
'//        BlobName : ����Ʈ���� ����ϴ� �ΰ��ʵ��, BlobTop,left, Width, Height: �ΰ�����
'// Copyright  (c) Morning Data
'/***********************************************************************/
Public Function gCryReport_Action(RptTitle As String, RptName As String, SRptChk As String, rQry As String, Optional RptCnt As String, Optional FormulaChk As String, _
                                    Optional ActionT As String, Optional DiskFileNam As String, Optional PortraitYN As Long, Optional Qty As Long) As Boolean
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim x           As Long
    Dim y           As Long
    Dim ErrNum      As Long
    Dim FormulaCnt  As Long
    Dim SubReptCnt  As Long
    Dim sFormulaCnt As Long
On Error GoTo Error_Message
    ErrNum = 0
    '/ Check
    If FormulaChk = "" Then FormulaChk = "N"
    FormulaCnt = UBound(gRpt_Formula)
    If SRptChk = "Y" Then
        SubReptCnt = UBound(gRpt_SRptName)
    End If
    If ActionT = "" Then
        ActionT = "P"  '/ �̸�������
    End If
   
    Set g_ADORS1 = New ADODB.Recordset
    Set g_ADORS2 = New ADODB.Recordset

     g_ADORS1.Open rQry, g_ERPDMS, adOpenKeyset, adLockBatchOptimistic

 '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    Set g_CApp = New CRAXDDRT.Application

    If RptCnt = "" Or RptCnt = "1" Then
        Set g_GCrview = frmRPT_View1.CRViewer1
    ElseIf RptCnt = "2" Then
        Set g_GCrview = frmRPT_View2.CRViewer1
    ElseIf RptCnt = "3" Then
        Set g_GCrview = frmRPT_View3.CRViewer1
    End If
    Set g_Report = g_CApp.OpenReport(MDC_Globals.SP_Path & "\" & SP_Report & "\" & RptName)
   
    g_Report.Database.Tables.Item(1).SetDataSource g_ADORS1, 3
    g_Report.DiscardSavedData
    
    If PortraitYN = 1 Then
        g_Report.PaperOrientation = crPortrait  ' ���ι���
    ElseIf PortraitYN = 2 Then
        g_Report.PaperOrientation = crLandscape ' ���ι���
    Else
        g_Report.PaperOrientation = crDefaultPaperOrientation
    End If

    If MDC_PS_Common.GetValue("SELECT U_PRTYN FROM [OHEM] WHERE U_MSTCOD = '" & MDC_PS_Common.User_MSTCOD & "'", 0, 1) = "Y" Then
        g_Report.PrinterSetup (0)
    End If
    
'/ SubReport /
    If SRptChk = "Y" Then  '/ ���긮��Ʈ�� �ְų� �̹����ΰ��������������� �������
        Set g_CrSections = g_Report.Sections
         For i = 1 To g_CrSections.Count
            Set g_CrSection = g_CrSections.Item(i)
            Set g_CrReportObjs = g_CrSection.ReportObjects
          '/** Kind(1-Formula, 2-Cation Text, 3-crLineObject,4-crBoxObject, 9-crBlobFieldObject)  **/
            For k = 1 To g_CrReportObjs.Count
                '//
                If g_CrReportObjs.Item(k).Kind = crSubreportObject Then
                    Set g_CrSubReportObj = g_CrReportObjs.Item(k)
                    Set g_CrSubReport = g_CrSubReportObj.OpenSubreport
                    For j = 1 To SubReptCnt
                        If g_CrSubReportObj.SubreportName = Trim(gRpt_SRptName(j)) Then
                            g_ADORS2.Open gRpt_SRptSqry(j), g_ERPDMS, adOpenKeyset, adLockBatchOptimistic
                            g_CrSubReport.Database.Tables.Item(1).SetDataSource g_ADORS2, 3
                            g_ADORS2.Close
                        '/ SubFormula //
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
                        End If
                    Next j
                End If
            Next k
        Next i
    End If
 '/ Formula /
    If FormulaCnt >= 1 Then
         g_Report.FormulaSyntax = crCrystalSyntaxFormula
         For i = 1 To g_Report.FormulaFields.Count
            Set g_cFormula = g_Report.FormulaFields.Item(i)
           
            For k = 1 To FormulaCnt
                 If g_cFormula.FormulaFieldName = Trim(gRpt_Formula(k)) Then
                     g_cFormula.Text = "'" & gRpt_Formula_Value(k) & "'"
                 End If
            Next k
        Next i
    End If
   
 '/ Report Action /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    If ActionT = "P" Then
    '/ Report Viewer Show /
        g_Report.PrintOutEx , Qty, , , , RptName
        
    ElseIf ActionT = "F" Then
    '/ Report File Export Show /
        g_Report.ExportOptions.DestinationType = crEDTDiskFile              '/ �����ѵ�ũ�� ����
        g_Report.ExportOptions.FormatType = crEFTPortableDocFormat          '/ PDF���Ϸ� ����(crEFTPortableDocFormat)
        g_Report.ExportOptions.DiskFileName = DiskFileNam
        g_Report.DisplayProgressDialog = False
        g_Report.Export False
    Else

    '/ Report Viewer Show /
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
    
 '/ Init_Crystal
    'g_ADORS1.Close
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
Error_Message:
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
       ' Sbo_Application.StatusBar.SetText MDC_SetMod.Get_TitleName("E0001", "��ȸ�ڷᰡ �����ϴ�. "), bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Print_Query : " & Space(10) & Err.Description, bmt_Short, smt_Error
    End If
    gCryReport_Action = False
End Function

Public Function SetDrive(Drive As String, Path As String, ID As String, PassWord As String) As Long

    Dim WS       As Object
    
    On Error GoTo Err
    
    '��Ʈ��ũ ����̺� ����
    Set WS = CreateObject("WScript.Network")
    
    If GetNetDrives(WS) = False Then
        WS.MapNetworkDrive Drive, Path, False, ID, PassWord
    End If
    
    SetDrive = 0
    Exit Function
    
Err:
    SetDrive = -1
    If Err.Description = " -2147024811" Then
        
    ElseIf Err.Number = "-2147023677" Then
        SetDrive = 0
    End If
End Function

Public Function GetNetDrives(WS) As Boolean  '��ũ��ũ ����̺갡 ���� �ϴ��� Ȯ��

    On Error GoTo Err
    
    Dim oDrive
    Dim i As Long
    Dim sTmp As String
    Dim sResult() As String

    Set oDrive = WS.EnumNetworkDrives
    ReDim sResult(oDrive.Count / 2, 1) As String
    
    For i = 0 To oDrive.Count - 1 Step 2
        sResult(i, 0) = oDrive.Item(i)
        sResult(i, 1) = oDrive.Item(i + 1)
        If sResult(i, 0) = "Q:" Then
            GetNetDrives = True
            Exit Function
        End If
    Next
    GetNetDrives = False
    
    Exit Function
Err:
    GetNetDrives = False
End Function

'Function ID : RemoveBlankRecord
'��    ��    : UDO���� �߻��ϴ� ���ڵ带 �ʼ����� üũ �� �����
'��    ��    : Matrix ��ü,DBDataSource ��ü,�ʼ� Field Name,�������μ� ���ι�ȣ ���Ž�
'�� ȯ ��    : ����
Public Sub RemoveBlankRecord(ByVal oMat As SAPbouiCOM.Matrix, ByVal oDS As SAPbouiCOM.DBDataSource, SpecialAlias As String, Optional LineNumAlias As String)

    Dim i As Long
    
    On Error GoTo Error_Message
    
    oMat.FlushToDataSource

    Do Until oDS.Size = i
        If Trim(oDS.GetValue(SpecialAlias, i)) = "" Then
                oDS.RemoveRecord i
        Else
            i = i + 1
        End If
    Loop
        
    If LineNumAlias <> "" Then
        For i = 0 To oDS.Size - 1
            oDS.setValue LineNumAlias, i, i + 1
        Next i
    End If
    
    oMat.LoadFromDataSource
    Set oMat = Nothing
    Set oDS = Nothing
    Exit Sub
    
Error_Message:
    Set oMat = Nothing
    Set oDS = Nothing
    Sbo_Application.StatusBar.SetText Err.Description, bmt_Short, smt_Error
End Sub


Public Sub SetMatrix_Sorter(oForm As SAPbouiCOM.Form, _
                             MatrixUID As String, _
                             MatrixType As BoFormItemTypes, _
                             Optional bSort As Boolean = True)
    On Error GoTo Error_Message

    Dim iLooper As Long
    Dim vMatrix As SAPbouiCOM.Matrix
    Dim vGrid   As SAPbouiCOM.Grid

    'MDC_Matrix_Sorter oform,"Mat1",it_MATRIX
    ' Calll  MDC_Matrix_Sorter (oform,"Mat1",it_MATRIX)

    '//---------------------------------------------------------------------------------------------------------
    '// Matrix & Grid ��������� 2007 PL18�̻�������� �����ϹǷ� �ڵ� �� üũ �ʼ�
    '//---------------------------------------------------------------------------------------------------------
    
    If oCompany.version >= "860040" Then '2007B PL18 �̻� �϶�(2007A ������ Ȯ�� �ʿ�)
 
        Select Case MatrixType

            Case it_GRID
                Set vGrid = oForm.Items(MatrixUID).Specific
                For iLooper = 0 To vGrid.Columns.Count - 1
                    vGrid.Columns(iLooper).TitleObject.Sortable = bSort
                Next iLooper
        
            Case it_MATRIX
                Set vMatrix = oForm.Items(MatrixUID).Specific
                For iLooper = 0 To vMatrix.Columns.Count - 1
                    vMatrix.Columns(iLooper).TitleObject.Sortable = bSort
                Next iLooper

        End Select
    End If

    Exit Sub
Error_Message:
    Sbo_Application.StatusBar.SetText Err.Description, bmt_Short
End Sub


Public Function ChkYearMonth(YearMon$) As Boolean
    Dim oYear$
    Dim oMonth$

    If Len(YearMon) < 6 Then
        ChkYearMonth = False
        Exit Function
    End If
    oYear = Mid(YearMon, 1, 4)
    If MDC_Com.uISNUMERIC(oYear, "0", "INT") < 2000 Or MDC_Com.uISNUMERIC(oYear, "0", "INT") > 3000 Then
        ChkYearMonth = False
        Exit Function
    End If
    oMonth = Mid(YearMon, 5, 2)
    If MDC_Com.uISNUMERIC(oMonth, "0", "INT") < 1 Or MDC_Com.uISNUMERIC(oMonth, "0", "INT") > 12 Then
        ChkYearMonth = False
        Exit Function
    End If
    ChkYearMonth = True
End Function

Public Function Get_ItemName(ItemCode$) As String
'// ǰ���� ��ȯ�մϴ�
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim Sql        As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Sql = "select ItemName from OITM WHERE ItemCode='" & ItemCode & "'"
    oRecordSet.DoQuery Sql
    Do Until oRecordSet.EOF
        Get_ItemName = oRecordSet.Fields(0).VALUE      '/
        oRecordSet.MoveNext
    Loop
    
    Set oRecordSet = Nothing
    Exit Function

LenDecimal_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "ǰ���� �����ü� �����ϴ�." & Space(10) & Err.Description, bmt_Short, True
End Function

' Function  : GF_DLookup
' ��    ��  : Ư�� ���̺��� Ư�� �ʵ尪�� ���ǿ� ���� ��ȯ�Ѵ�.
' Ex)   GF_DLookup("ItemName", "OITM", "ItemCode = 'CDM2608209044'")
'        Item Code�� CDM2608209044�� ItemName �����
Public Function GF_DLookup(sField As String, Optional sTbl As String = "", Optional sWhere As String = "") As Variant

    Dim oTRecordset             As SAPbobsCOM.Recordset
    Dim sSQL                    As String

    Set oTRecordset = oCompany.GetBusinessObject(BoRecordset)

    If sTbl = "" Then
        sSQL = "SELECT " & sField
    Else
        If sWhere = "" Then
            sSQL = "SELECT " & sField & " FROM " & sTbl
        Else
            sSQL = "SELECT " & sField & " FROM " & sTbl & " WHERE " & sWhere
        End If
    End If

    oTRecordset.DoQuery sSQL

    If oTRecordset.EOF = False Then
        GF_DLookup = oTRecordset.Fields(0).VALUE
    Else
        GF_DLookup = ""
    End If

    Set oTRecordset = Nothing

End Function


'--------------------------------------------------------------------------------------
'//     NULL �� üũ
'--------------------------------------------------------------------------------------
Public Function GF_Nz(pAnyData) As Currency
    
    On Error GoTo Err_Disp
    
    If pAnyData = "" Then pAnyData = 0
    
    If Not IsNumeric(pAnyData) Then pAnyData = 0
    
    GF_Nz = IIf(IsNull(pAnyData), 0, pAnyData)
    
Exit Function

Err_Disp:
    
    pAnyData = 0
    
End Function


'//************************
'// Adding a ComboBox item
'//***********************
Public Sub GP_CreateComboBox(ByVal pForm As SAPbouiCOM.Form, ByVal pItemName As String _
                                , ByVal pFieldName As String, ByVal pTop As Integer _
                                , ByVal pLeft As Integer, ByVal pWidth As Integer, ByVal pHeight As Integer _
                                , Optional ByVal pFromPane As Integer = 0 _
                                , Optional ByVal pToPane As Integer = 0 _
                                , Optional ByVal pTableName As String _
                                , Optional ByVal pEnabled As Boolean = True _
                                , Optional ByVal pAffectsFormMode As Boolean = True)
                                
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oItem       As SAPbouiCOM.Item
    
    Set oItem = pForm.Items.Add(pItemName, it_COMBO_BOX)
    oItem.Enabled = pEnabled
    oItem.AffectsFormMode = pAffectsFormMode
    
    Call GP_SetItemDefaultSetting(oItem, pTop, pLeft, pWidth, pHeight, pFromPane, pToPane)
    Set oCombo = oItem.Specific
    oItem.DisplayDesc = True

    If pFieldName <> "" Then
        If pTableName <> "" Then
            oCombo.DataBind.SetBound True, pTableName, pFieldName
        Else
            oCombo.DataBind.SetBound True, "", pFieldName
        End If
    End If
    
    Set oItem = Nothing: Set oCombo = Nothing
    
End Sub

'-----------------------------------------------------------------------------------------
'   ComboBox �ڽ� ������ ä���
'-----------------------------------------------------------------------------------------
Public Sub SetReDataCombo(ByVal pForm As SAPbouiCOM.Form, ByVal pSQL As String, pCombo As SAPbouiCOM.ComboBox, Optional AddSpace As String)
    
    Dim i           As Long
    Dim sQry        As String
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//���� �޺� ������ ����
    If pCombo.ValidValues.Count > 0 Then
        For i = 0 To pCombo.ValidValues.Count - 1
            pCombo.ValidValues.Remove 0, psk_Index
        Next
    End If
    If AddSpace = "Y" Then
        Call pCombo.ValidValues.Add("", "")
    End If
    
    oRecordSet.DoQuery pSQL
    
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            pCombo.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
'        pForm.Items(pCombo).DisplayDesc = True
    End If
    
End Sub


'-----------------------------------------------------------------------------------------
'   ComboBox �ڽ� ������ ä���
'-----------------------------------------------------------------------------------------
'   sType = "A" -> ADD (�� �� �����Ϳ� -1�� �߰��Ѵ�.)
'   : ��ȸ ȭ���̳� �ʼ� �Է»����� �ƴҰ�� �� ���ڿ��� �߰��Ѵ�.
'   (���ǻ���) - �����ʵ�� �߰��� �����ʹ� �⺻���� -1 �� ���� ������ ������ �ۼ���
'   �̸� �ݿ��ؾ� �Ѵ�.
'-----------------------------------------------------------------------------------------
Public Sub GP_SetComboBox(ByVal pForm As SAPbouiCOM.Form, pSQL As String, pCombo As SAPbouiCOM.ComboBox, Optional pType As String)
    
    Dim oTRecordset  As SAPbobsCOM.Recordset
    Dim iLooper     As Integer
        
    'Call GP_ComboClear(pCombo)
    Set oTRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    oTRecordset.DoQuery pSQL
    
    If UCase$(pType) = "ADD" Then
        pCombo.ValidValues.Add "", ""
    End If
    
    For iLooper = 1 To oTRecordset.RecordCount
            
        pCombo.ValidValues.Add CStr(oTRecordset.Fields.Item(0).VALUE), CStr(oTRecordset.Fields.Item(1).VALUE)
        oTRecordset.MoveNext
        
    Next iLooper
    
    '-----------------------------------------------------------
    '// �޺��ڽ����� ù��° �����͸� �����ϰ� �Ѵ�.
    '   �̰� Ǯ���ָ� ��ũ��ų�� ������ ������ �߻�
    'If pCombo.ValidValues.Count > 0 Then
'        Call pCombo.Select(0, psk_Index)
    'End If
    '-----------------------------------------------------------
    Set oTRecordset = Nothing: Set pCombo = Nothing
    
End Sub

Public Sub GP_CreateStaticText(ByVal pForm As SAPbouiCOM.Form, ByVal pItemName As String _
                                , ByVal pLinkTo As String, ByVal pCaption As String, ByVal pTop As Integer _
                                , ByVal pLeft As Integer, ByVal pWidth As Integer, ByVal pHeight As Integer _
                                , Optional ByVal pFromPane As Integer = 0 _
                                , Optional ByVal pToPane As Integer = 0)
    Dim oStatic As SAPbouiCOM.StaticText
    Dim oItem   As SAPbouiCOM.Item
    
    Set oItem = pForm.Items.Add(pItemName, it_STATIC)
    Call GP_SetItemDefaultSetting(oItem, pTop, pLeft, pWidth, pHeight, pFromPane, pToPane)
    oItem.LinkTo = pLinkTo
    Set oStatic = oItem.Specific
    oStatic.Caption = pCaption
    Set oStatic = Nothing: Set oItem = Nothing
    
End Sub


'/**********************
'// Setting an Item properties
'//*********************
Public Sub GP_SetItemDefaultSetting(ByVal pItem As SAPbouiCOM.Item, ByVal pTop As Integer _
                        , ByVal pLeft As Integer, ByVal pWidth As Integer, ByVal pHeight As Integer _
                        , Optional ByVal pFromPane As Integer = 0 _
                        , Optional ByVal pToPane As Integer = 0)
                        
    pItem.Left = pLeft
    pItem.Width = pWidth
    pItem.Top = pTop
    pItem.Height = pHeight
    pItem.FromPane = pFromPane
    pItem.ToPane = pToPane
End Sub


'//************************
'// Adding a Ext Edit item
'//***********************
'//************************
'// Adding a Text Edit item
'//***********************
Public Sub GP_CreateTextEdit(ByVal pForm As SAPbouiCOM.Form, ByVal pItemName As String _
                        , ByVal pFieldName As String, ByVal pTop As Integer, ByVal pLeft As Integer _
                        , ByVal pWidth As Integer, ByVal pHeight As Integer _
                        , Optional ByVal pFromPane As Integer = 0 _
                        , Optional ByVal pToPane As Integer = 0, Optional ByVal pTableName As String _
                        , Optional ByVal pRightJustified As Boolean = False _
                        , Optional ByVal pEnabled As Boolean = True _
                        , Optional ByVal pAffectsFormMode As Boolean = True _
                        , Optional ByVal pObjectType As Integer = -1 _
                        , Optional ByVal pDescription As String = "" _
                        , Optional ByVal pType As String)

    Dim oEdit   As SAPbouiCOM.EditText
    Dim oItem   As SAPbouiCOM.Item
    
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    
    If pType = "EXTEDIT" Then
        Set oItem = pForm.Items.Add(pItemName, it_EXTEDIT)
    Else
        Set oItem = pForm.Items.Add(pItemName, it_EDIT)
    End If
    
    oItem.Enabled = pEnabled
    oItem.RightJustified = pRightJustified
    oItem.AffectsFormMode = pAffectsFormMode
    oItem.Description = pDescription
    oItem.TextStyle = ts_EXTEND
    Call GP_SetItemDefaultSetting(oItem, pTop, pLeft, pWidth, pHeight, pFromPane, pToPane)
    Set oEdit = oItem.Specific
    
    If pFieldName <> "" Then
        If pTableName <> "" Then
            oEdit.DataBind.SetBound True, pTableName, pFieldName
        Else
            oEdit.DataBind.SetBound True, "", pFieldName
        End If
    End If
    
    If pObjectType >= 0 Then
                
        Set oCFLs = pForm.ChooseFromLists
        Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

        If pObjectType = 4 Or pObjectType = 999 Then
            oCFLCreationParams.MultiSelection = True
        Else
            oCFLCreationParams.MultiSelection = False
        End If
        oCFLCreationParams.ObjectType = pObjectType
        oCFLCreationParams.uniqueID = "CFL" & pItemName
        
        Select Case pObjectType
        Case 999, 992, 993: oCFLCreationParams.ObjectType = 4
        Case 998, 997, 996, 995, 994: oCFLCreationParams.ObjectType = 2
        Case Else: oCFLCreationParams.ObjectType = pObjectType
        End Select

        Set oCFL = oCFLs.Add(oCFLCreationParams)
        Set oCons = oCFL.GetConditions()
        
        Select Case pObjectType
        Case 1      '// GL����Ÿ
        
        Case 2, 996     '// ���Ǹ���Ÿ
            Select Case pForm.Type
            Case 146
                Set oCon = oCons.Add()
                oCon.Alias = "CardType"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "S"
            Case Else
                Set oCon = oCons.Add()
                oCon.Alias = "CardType"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "C"
                oCon.Relationship = cr_OR
                Set oCon = oCons.Add()
                oCon.Alias = "CardType"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "S"
            End Select
        Case 997, 994      '// ���Ǹ���Ÿ - ��
            Set oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
        Case 998, 995     '// ���Ǹ���Ÿ - ���޾�ü
            Set oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
        Case 4, 999     '// ǰ�񸶽�Ÿ
            Set oCon = oCons.Add()
            oCon.Alias = "InVntItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
        Case 992, 993     '// BOM(��ȹ���� ���޹���� �ۼ�)
            Set oCon = oCons.Add()
            oCon.Alias = "Prcrmntmtd"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "M"
        Case 17     '// �Ǹſ���
        
        Case 64      '// â��
'            Set oCon = oCons.Add()
'            oCon.Alias = "WhsCode"
'            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
'            oCon.CondVal = "01"
        Case 171    '// �������Ÿ
        Case 202    '// �������
            Set oCon = oCons.Add()
            oCon.Alias = "Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "R"
            oCon.Relationship = cr_AND

            Set oCon = oCons.Add()
            oCon.ComparedAlias = "PlannedQty"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_THAN
            oCon.CondVal = "CmpltQty"
        End Select

        oCFL.SetConditions oCons
        oEdit.ChooseFromListUID = "CFL" & pItemName
        
        '---------------------------------------------------------------------
        '   ChooseFromListAlias �� �� ChooseFromListUID ������ �����ؾ� �Ѵ�.
        '---------------------------------------------------------------------
        Select Case pObjectType
        Case 1      '// GL����Ÿ
            oEdit.ChooseFromListAlias = "AcctCode"
        Case 2, 994, 995    '// ���Ǹ���Ÿ
            oEdit.ChooseFromListAlias = "CardCode"
        Case 998, 997, 996
            oEdit.ChooseFromListAlias = "CardName"
        Case 4, 992     '// ǰ�񸶽�Ÿ
            oEdit.ChooseFromListAlias = "ItemCode"
        Case 999, 993     '// ǰ�񸶽�Ÿ
            oEdit.ChooseFromListAlias = "ItemName"
        Case 17    '// �Ǹſ���
            oEdit.ChooseFromListAlias = "DocEntry"
        Case 52     '// ǰ��׷�
            oEdit.ChooseFromListAlias = "ItmsGrpCod"
        Case 64     '// â��
            oEdit.ChooseFromListAlias = "WhsCode"
        Case 171    '// �������Ÿ
            oEdit.ChooseFromListAlias = "EmpID"
        Case 59    '// �԰�
            oEdit.ChooseFromListAlias = "DocEntry"
        Case 60    '// ���
            oEdit.ChooseFromListAlias = "DocEntry"
        Case 202    '// �������
            oEdit.ChooseFromListAlias = "DocEntry"
        End Select
        
    End If
    
    Set oEdit = Nothing: Set oItem = Nothing
    
End Sub

'// Matrix Combo Box Setting
Public Sub GP_MatrixSetMatComboList(fCombo As SAPbouiCOM.Column, _
                                        fSQL As String, _
                                        Optional AndLine$, _
                                        Optional AddSpace$)
    'Function ID : GetListIndex
    '��    ��    :
    '��    ��    : Lst
    '�� ȯ ��    : None
    'Ư�̻���    : �޺��ڽ��� ���� �� ������ �ý��� �ڵ忡�� ������ �����Ѵ�
    
    Dim fRecordset As SAPbobsCOM.Recordset
    
    Set fRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    fRecordset.DoQuery fSQL

    If AddSpace = "Y" Then
        Call fCombo.ValidValues.Add("", "")
    End If
    Do Until fRecordset.EOF
        Call fCombo.ValidValues.Add(fRecordset.Fields(0).VALUE, fRecordset.Fields(1).VALUE)
        fRecordset.MoveNext
    Loop
        
    Set fRecordset = Nothing

End Sub

'-----------------------------------------------------------------------------------------
'   �׺���̼� ��Ʈ�� ���� ���̱�/���߱� �Լ�
'   -> �̸�����, ���, �����, ã��, �߰�, ����, ����, ��ó��, �ǳ�, ���
'-----------------------------------------------------------------------------------------
Public Sub GP_EnableMenus(eForm As SAPbouiCOM.Form, _
                              ByVal bPreview As Boolean, _
                              ByVal bPrint As Boolean, _
                              ByVal bDeleteRow As Boolean, _
                              ByVal bFind As Boolean, _
                              ByVal bAdd As Boolean, _
                              ByVal bNextRecord As Boolean, _
                              ByVal bPreviousRecord As Boolean, _
                              ByVal bFirstRecord As Boolean, _
                              ByVal bLastRecord As Boolean, _
                              ByVal bCancel As Boolean, _
                              Optional ByVal bRowAdd As Boolean = False, _
                              Optional ByVal bDuplicate As Boolean = False, _
                              Optional ByVal bRemove As Boolean = False, _
                              Optional ByVal bRowClose As Boolean = False, _
                              Optional ByVal bClose As Boolean = False)

'    If Left$(eForm.Type, 2) = "20" Then
        eForm.EnableMenu "519", bPreview         '// �μ�[�̸�����]
        eForm.EnableMenu "520", bPrint           '// �μ�[���]
        eForm.EnableMenu "1293", bDeleteRow      '// �����
        eForm.EnableMenu "1281", bFind           '// ����ã��
        eForm.EnableMenu "1282", bAdd            '// �����߰�
        eForm.EnableMenu "1283", bRemove         '// ��������(������ ������ ���)
        eForm.EnableMenu "1286", bClose          '// �����ݱ�
        eForm.EnableMenu "1288", bNextRecord     '// ����
        eForm.EnableMenu "1289", bPreviousRecord '// ����
        eForm.EnableMenu "1290", bFirstRecord    '// ��ó��
        eForm.EnableMenu "1291", bLastRecord     '// �ǳ�
        eForm.EnableMenu "1284", bCancel         '// �������
        eForm.EnableMenu "1292", bRowAdd         '// ���߰�
        eForm.EnableMenu "1287", bDuplicate      '// ��������
        eForm.EnableMenu "1299", bRowClose       '// ��ݱ�
'    End If
End Sub

'// �߰��� ���� AUTOKEY�� ���� ��� ���������� 1�� �߰�
Public Function GF_AUTOKEYAdd(sObjectCode As String, sTablename As String)

    Dim sQry            As String
    Dim gRecordset      As SAPbobsCOM.Recordset
    Dim sAutoKey        As Long
    Dim sDocEntry       As Long
    
    Set gRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT AUTOKEY FROM ONNM WHERE OBJECTCODE = " & Chr$(39) & sObjectCode & Chr$(39)
    gRecordset.DoQuery sQry
    sAutoKey = GF_Nz(gRecordset.Fields(0).VALUE)
    
    sQry = "SELECT MAX(DOCENTRY) + 1 FROM " & sTablename
    gRecordset.DoQuery sQry
    sDocEntry = GF_Nz(gRecordset.Fields(0).VALUE)
    
    If sDocEntry = 0 Or sDocEntry = 1 Then
        sDocEntry = sDocEntry + 1
    End If
    
    If sDocEntry > sAutoKey Then
        sQry = "UPDATE ONNM SET AUTOKEY = " & sAutoKey + (sDocEntry - sAutoKey)
        sQry = sQry & " WHERE OBJECTCODE = " & Chr$(39) & sObjectCode & Chr$(39)
        gRecordset.DoQuery sQry
    End If
    
    Set gRecordset = Nothing
    
    Exit Function

End Function

Public Function OFPR_PeriodStatus(sDocDate As String) As String

    Dim sSQL        As String
    Dim sRecordset  As SAPbobsCOM.Recordset
    
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    sSQL = "SELECT PERIODSTAT FROM OFPR "
    sSQL = sSQL & " WHERE " & Chr$(39) & sDocDate & Chr$(39)
    sSQL = sSQL & " Between F_REFDATE And T_REFDATE"
    sRecordset.DoQuery sSQL
    
    OFPR_PeriodStatus = sRecordset.Fields(0).VALUE
    
    Set sRecordset = Nothing
    
    Exit Function
    
End Function

Public Function AutoManaged(oForm As SAPbouiCOM.Form, ByRef sItem As String)

    Dim i       As Long
    Dim Item()  As String
    
    Item = Split(Replace(sItem, " ", ""), ",")
    
    oForm.AutoManaged = True
    
    For i = 0 To UBound(Item)
        oForm.Items.Item(Item(i)).SetAutoManagedAttribute ama_Editable, afm_Add, mvb_True
        oForm.Items.Item(Item(i)).SetAutoManagedAttribute ama_Editable, afm_Find, mvb_True
        oForm.Items.Item(Item(i)).SetAutoManagedAttribute ama_Editable, afm_Ok, mvb_False
        
    Next i
    
End Function


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
        sSQL = "SELECT count(*) FROM " + Tablename + " Where " + ColumnName + "=" + CStr(Key_Str)
        If And_Line <> "" Then
           sSQL = sSQL + And_Line
        End If

        Set s_Recordset = oCompany.GetBusinessObject(BoRecordset)
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
    Else
      Value_ChkYn = True
    End If
    Set s_Recordset = Nothing
End Function

Public Function Get_EmpID_InFo(EmpCode$) As ZPAY_g_EmpID
    '/ ������� ��ȸ  /
    Dim F_EmpID     As ZPAY_g_EmpID
    Dim Rs          As New SAPbobsCOM.Recordset
    Dim Sql$
    
    
    Set Rs = oCompany.GetBusinessObject(BoRecordset)
    
    Sql = "SELECT  T0.U_EmpId,"                             '//�������
    Sql = Sql & " T0.U_FullName,"                           '//�����
    Sql = Sql & " T0.Code,"                             '//�����ȣ
    Sql = Sql & " T0.U_CLTCOD,"                             '//�����
    Sql = Sql & " T0.U_TeamCode,"                           '//�μ�
    Sql = Sql & " T0.U_RspCode,"                            '//���
    Sql = Sql & " T0.U_ClsCode,"                            '//��
    Sql = Sql & " Substring(replace(Convert(VarChar(10), T0.U_StartDat, 20), '-', ''), 1, 8) AS INPDAT,"            '//�Ի�����
    Sql = Sql & " Substring(replace(Convert(VarChar(10), T0.U_TermDate , 20), '-', ''), 1, 8) AS OUTDAT,"           '//�������
    Sql = Sql & " Substring(replace(Convert(VarChar(10), T0.U_GRPDAT , 20), '-', ''), 1, 8) AS GRPDAT,"              '//�׷��Ի���
'    Sql = Sql & " Substring(replace(Convert(VarChar(10), T0.U_BALYMD,  20), '-', ''), 1, 8) AS BALYMD,"            '//�����߷�����
'    Sql = Sql & " T0.U_BALCOD,"                            '//�����߷ɺμ�
    Sql = Sql & " T0.U_JIGTYP,"                             '//��������
    Sql = Sql & " T2.posID,"                                '//����(��å)�ڵ�
    Sql = Sql & " T0.U_HOBONG ,"                            '//ȣ��
    Sql = Sql & " T0.U_STDAMT ,"                            '//�޿��⺻��
    Sql = Sql & " T0.U_PAYTYP,"                                '//�޿�����
    Sql = Sql & " T0.U_PAYSEL ,"                            '//�޿����޴��
    Sql = Sql & " T0.U_GBHSEL ,"                            '//��뺸�迩��
    Sql = Sql & " T0.U_govid ,"                             '//�ֹι�ȣ
    Sql = Sql & " T0.U_sex ,"                               '//����
    Sql = Sql & " Substring(replace(Convert(VarChar(10), T0.U_RETDAT,  20), '-', ''), 1, 8) AS RETDAT,"             '//�߰�������
    Sql = Sql & " T0.U_JIGCOD,"                            '//�����ڵ�
    Sql = Sql & " (Case T0.U_BAEWOO When 'Y' then 1 else 0 end) AS U_BAEWOO,"                                       '//�����
    Sql = Sql & " ISNULL(T0.U_BUYNSU, 0) AS U_BUYNSU,"      '//�ξ簡��
    Sql = Sql & " ISNULL(T0.U_DAGYSU, 0) AS U_DAGYSU,"      '//���ڳ�
    Sql = Sql & " ISNULL((Select Convert(Char(8),MAX(Dateadd(dd, 1, U_ENDRET)),112) From [@PH_PY115A] Where U_MSTCOD = T0.Code), Convert(Char(8),Isnull(U_RetDat,U_STARTDAT),112)) As ENDRET "
    Sql = Sql & " FROM [@PH_PY001A] T0  LEFT JOIN [OUDP] T1 ON T0.U_TeamCode = T1.Code"
    Sql = Sql & " LEFT JOIN [OHPS] T2 ON T0.U_Position = T2.PosID"
'    Sql = Sql & " LEFT JOIN   (SELECT T0.*, T1.U_RelCd"
'    Sql = Sql & " FROM [@PH_PY001A] T0 INNER JOIN [@PS_HR200L] T1 ON T0.U_PAYTYP = T1.U_Code AND T1.Code = 'P132') T3 ON T0.U_MSTCOD = T3.Code"
    Sql = Sql & " WHERE T0.Code = '" & EmpCode & "'"
    Sql = Sql & " ORDER BY T0.Code"
    Rs.DoQuery Sql
    
    If Rs.RecordCount = 0 Then
       With F_EmpID
            .EmpID = Space(0)
            .MSTNAM = Space(0)
            .MSTCOD = Space(0)
            .CLTCOD = Space(0)
            .TeamCode = Space(0)
            .RspCode = Space(0)
            .CLTCOD = Space(0)
            .StartDate = Space(0)
            .TermDate = Space(0)
            .BALCOD = Space(0)
            .BALYMD = Space(0)
            .JIGTYP = Space(0)
            .PAYTYP = Space(0)
            .PAYSEL = Space$(0)
            .Position = Space(0)
            .HOBONG = Space(0)
            .STDAMT = 0
            .GBHSEL = Space(0)
            .PERNBR = Space(0)
            .Sex = ""
            .RETDAT = ""
            .JIGCOD = ""
            .GONCNT = 0
            .DAGYSU = 0
            .GRPDAT = Space(0)
            .ENDRET = Space(0)
       End With
    Else
       Do Until Rs.EOF
        With F_EmpID
            .EmpID = Rs.Fields("U_EmpID").VALUE         '//�������
            .MSTNAM = Rs.Fields("U_FullName").VALUE     '//�����
            .MSTCOD = Rs.Fields("Code").VALUE       '//����ڵ�
            .CLTCOD = Rs.Fields("U_CLTCOD").VALUE       '//�����
            .TeamCode = Rs.Fields("U_TeamCode").VALUE   '//�μ�
            .RspCode = Rs.Fields("U_RspCode").VALUE     '//���
            .ClsCode = Rs.Fields("U_ClsCode").VALUE     '//��
            .StartDate = Rs.Fields("INPDAT").VALUE      '//�Ի�����
            .TermDate = Rs.Fields("OUTDAT").VALUE       '//�������
'            .BALYMD = Rs.Fields("U_BALYMD").Value       '//�����߷�����
'            .BALCOD = Rs.Fields("U_BALCOD").Value       '//�����߷ɺμ�
            .JIGTYP = Rs.Fields("U_JIGTYP").VALUE       '//����
            .Position = Rs.Fields("PosID").VALUE        '//����
            .HOBONG = Rs.Fields("U_Hobong").VALUE       '//ȣ��
            .STDAMT = Rs.Fields("U_STDAMT").VALUE       '//�⺻��
            .PAYTYP = Rs.Fields("U_PAYTYP").VALUE       '//�޿�����
            .PAYSEL = Rs.Fields("U_PAYSEL").VALUE       '//�޿������ϱ���
            .GBHSEL = Rs.Fields("U_GBHSEL").VALUE       '//��뺸�賳�Կ���
            .PERNBR = Rs.Fields("U_govid").VALUE        '//�ֹι�ȣ
            .Sex = Rs.Fields("U_SEX").VALUE             '//����
            .RETDAT = Rs.Fields("RETDAT").VALUE         '//�ߵ���������
            .JIGCOD = Rs.Fields("U_JIGCOD").VALUE       '//����
            .GONCNT = 1 + Rs.Fields("U_BAEWOO").VALUE + Rs.Fields("U_BUYNSU").VALUE '//�ξ簡��
            .DAGYSU = Rs.Fields("U_DAGYSU").VALUE       '//���ڳ����
            .GRPDAT = Rs.Fields("GRPDAT").VALUE         '//�׷��Ի�����
            .ENDRET = Rs.Fields("ENDRET").VALUE        '//��������
        End With
        Rs.MoveNext
     Loop
        
    End If
    Get_EmpID_InFo = F_EmpID
    Set Rs = Nothing
End Function


Public Function Get_PayLockInfo(sJOBYMM As String, sJOBTYP As String, sJOBGBN As String, sPAYSEL As String) As Boolean
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "       SELECT ISNULL(U_ENDCHK, 'N') "
    sQry = sQry & "FROM   [@ZPY307L] "
    sQry = sQry & "WHERE  Code = '" & Left(sJOBYMM, 4) & "' "
    sQry = sQry & "AND    U_JOBYMM = '" & sJOBYMM & "' "
    If Trim(sJOBTYP) <> "%" And Trim(sJOBTYP) <> "" Then
        sQry = sQry & "AND   (CASE WHEN U_JOBTYP = '%' THEN '" & sJOBTYP & "' ELSE U_JOBTYP END) LIKE '" & sJOBTYP & "' "
    End If
    If Trim(sJOBGBN) <> "%" And Trim(sJOBGBN) <> "" Then
        sQry = sQry & "AND   (CASE WHEN U_JOBGBN = '%' THEN '" & sJOBGBN & "' ELSE U_JOBTYP END) LIKE '" & sJOBGBN & "' "
    End If
    If Trim(sPAYSEL) <> "%" And Trim(sPAYSEL) <> "" Then
        sQry = sQry & "AND   (CASE WHEN U_PAYSEL = '%' THEN '" & sPAYSEL & "' ELSE U_JOBTYP END) LIKE '" & sPAYSEL & "' "
    End If
    
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        Get_PayLockInfo = False
    ElseIf oRecordSet.Fields(0).VALUE = "N" Then
        Get_PayLockInfo = False
    Else
        Get_PayLockInfo = True
    End If

    Set oRecordSet = Nothing

End Function


Public Function TaxNoCheck(ByVal strNo As String) As Boolean
'*******************************************************
' ����ڹ�ȣ ����üũ
'********************************************************
   
   Const COMPNO_LEN      As Byte = 10 '����ڹ�ȣ�� ����

   Dim blnRet            As Boolean   '�����
   Dim aryNo(COMPNO_LEN) As Byte      '���ڿ� �迭
   Dim bytCntNo          As Byte      '��������
   Dim intMod            As Integer   '����������
   Dim intInt            As Integer   '�Ҽ������� ���簪
   Dim intSub            As Integer   '�����
   Dim BUSNBR            As String    '����ڹ�ȣ

    BUSNBR = Replace(strNo, "-", "")
   '����ڹ�ȣ�� ���̰� 10�ڸ����
   If (Len(Trim(BUSNBR)) = COMPNO_LEN) Then
      '������ ���鼭 ����Ʈ�迭�� �����
      For bytCntNo = 1 To COMPNO_LEN
         aryNo(bytCntNo) = Val(Mid(BUSNBR, bytCntNo, 1))
      Next bytCntNo
      '������ ���ڸ� ���Ѵ�

      intMod = ((aryNo(1) * 1) + (aryNo(2) * 3) + (aryNo(3) * 7) + (aryNo(4) * 1) + _
                (aryNo(5) * 3) + (aryNo(6) * 7) + (aryNo(7) * 1) + (aryNo(8) * 3)) Mod COMPNO_LEN

      '�Ҽ������ϸ� �����Ͽ� ���Ѵ�
      intInt = Int(aryNo(9) * 5 / COMPNO_LEN)
      '������� ���Ѵ�
      intSub = (aryNo(9) * 5) - (intInt * 10)

      intSub = (intMod + intInt + intSub) Mod 10

      intSub = IIf((intSub = 0), 10, intSub)

      'üũ���� Ȯ���Ͽ� ������ �Ǻ��Ѵ�

      blnRet = (aryNo(COMPNO_LEN) = (COMPNO_LEN - intSub))
   Else
      blnRet = False
   End If
 '����� �����Ѵ�
  TaxNoCheck = blnRet
End Function


Public Function RInt(Dub As Double, ByVal oPnt As Integer, ByVal Rtype As String)
   Dim Rub  As Double
   Dim Cub  As Double
   Dim Pnt  As Integer
   If Dub = 0 Then
        RInt = 0
        Exit Function
   End If
   Pnt = CInt(oPnt)
   Select Case Pnt
   Case 1       '/ 1��
        Rub = 0.5
        Cub = 0.9
   Case 10   '/ 10��
        Rub = 5
        Cub = 9
   Case 100
        Rub = 50
        Cub = 90
   Case 1000
        Rub = 500
        Cub = 900
   End Select
   
'///////////////////////////////////////////////////
'/ ����ó��(R:�ݿø�, F:����, C:�ø�)
'///////////////////////////////////////////////////
   Select Case Trim(Rtype)
   Case "R"
        RInt = Int((Dub + Rub) / Pnt) * Pnt
   Case "C"
        RInt = Int((Dub + Cub) / Pnt) * Pnt
   Case "F"
        RInt = Int(Dub / Pnt) * Pnt
   End Select
   
End Function


'---------------------------------------------------------------------------------------
' Procedure : Term2
' Author    :
' Date      :
' Purpose   : �ټӱⰣ ���� ����ϴ� �Լ�
' Remark    : Term�Լ��� ������ ���� ����
'           : �������ڸ� �������� ��¥�� 1��, 1����, 1�Ͼ� ���ؼ� �������ڰ� �ɶ����� ī��Ʈ�ؼ�
'             �����
'---------------------------------------------------------------------------------------
'
Public Sub Term2(STRDAT As String, ENDDAT As String)
    Dim CHKDAY  As String
    Dim CHKDAY1 As String
    Dim ENDDAT1 As String
    Dim TempCnt As Integer

    ZPAY_GBL_GNSYER = 0:   ZPAY_GBL_GNMYER = 0
    ZPAY_GBL_GNSMON = 0:   ZPAY_GBL_GNMMON = 0
    ZPAY_GBL_GNSDAY = 0:   ZPAY_GBL_GNMDAY = 0
    ENDDAT1 = Format(DateAdd("d", 1, Format(ENDDAT, "0000-00-00")), "YYYYMMDD")
    
    CHKDAY1 = STRDAT
    '// �ټӳ�� üũ
    TempCnt = 0
    Do Until CHKDAY > ENDDAT1
        TempCnt = TempCnt + 1
        CHKDAY = Format(DateAdd("yyyy", TempCnt, Format(CHKDAY1, "0000-00-00")), "YYYYMMDD")
    Loop
    ZPAY_GBL_GNSYER = TempCnt - 1
    CHKDAY1 = Format(DateAdd("yyyy", ZPAY_GBL_GNSYER, Format(CHKDAY1, "0000-00-00")), "YYYYMMDD")
    CHKDAY = CHKDAY1
    
    '// �ټӿ��� üũ
    TempCnt = 0
    Do Until CHKDAY > ENDDAT1
        TempCnt = TempCnt + 1
        CHKDAY = Format(DateAdd("m", TempCnt, Format(CHKDAY1, "0000-00-00")), "YYYYMMDD")
    Loop
    ZPAY_GBL_GNSMON = TempCnt - 1
    CHKDAY1 = Format(DateAdd("m", ZPAY_GBL_GNSMON, Format(CHKDAY1, "0000-00-00")), "YYYYMMDD")
    CHKDAY = CHKDAY1
    
    '// �ټ��ϼ� üũ
    TempCnt = 0
    Do Until CHKDAY > ENDDAT1
        TempCnt = TempCnt + 1
        CHKDAY = Format(DateAdd("d", TempCnt, Format(CHKDAY1, "0000-00-00")), "YYYYMMDD")
    Loop
    ZPAY_GBL_GNSDAY = TempCnt - 1
    CHKDAY = Format(DateAdd("d", ZPAY_GBL_GNSDAY, Format(CHKDAY1, "0000-00-00")), "YYYYMMDD")
    
'/ �ټӿ��� /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
   ZPAY_GBL_GNMYER = ZPAY_GBL_GNSYER
'/ �ټӿ���
   ZPAY_GBL_GNMMON = ZPAY_GBL_GNSYER * 12 + ZPAY_GBL_GNSMON
   
End Sub

Public Function IInt(Dub As Double, Pnt As Double)
   Dim SDub As String * 20
   Dim TDub As Double
   Dim Tub  As Double
'/
   Tub = IIf(Dub >= 0, (Dub / Pnt), (Dub / Pnt * -1))
   SDub = Format$(Tub, "0000000000000.000000")
   Mid$(SDub, 14, 7) = Space$(7)
   TDub = Val(SDub)
   IInt = IIf(Dub >= 0, (TDub * Pnt), (TDub * Pnt * -1))
End Function


Public Function Get_GabGunSe_Table(ByRef GABGUN As Double, ByRef JUMINN As Double, ByVal oINCOME As Double, _
                                   ByVal oInWON%, ByVal oChlWON%, ByVal JOBYMM As String, _
                                   ByVal oKUKAMT As Double, ByVal PAY_001 As String) As Variant
On Error GoTo Error_Message
    Dim Rs     As SAPbobsCOM.Recordset
    Dim sQry    As String
    
    Dim WK_INCOME As Double
    Dim WK_GULTAX As Double
    
 '/ Initial
    WK_INCOME = 0
    WK_GULTAX = 0
        
    Set Rs = oCompany.GetBusinessObject(BoRecordset)
    GABGUN = 0
    JUMINN = 0
 '/ �����޾�
   If oINCOME <= 0 Then
       Get_GabGunSe_Table = "�����ݾ��� 0���� �۰ų� �����ϴ�. Ȯ���Ͽ� �ּ���."
       Exit Function
    End If
    
    WK_INCOME = oINCOME
        
    If JOBYMM <= "201201" Then
        
        '/ 1000�����ʰ���
        If oINCOME > 10000000 Then
           GABGUN = IInt(((oINCOME - 10000000) * 0.95) * 0.35, 10)
           WK_INCOME = 10000000
        End If

    Else
                
        If oINCOME > 28000000 Then
            
            '/ 2800�����ʰ���
            GABGUN = 5985000 + IInt(((oINCOME - 28000000) * 0.95) * 0.38, 10)
            WK_INCOME = 10000000

        ElseIf oINCOME > 10000000 Then
            
            '/ 1000�����ʰ���
            GABGUN = IInt(((oINCOME - 10000000) * 0.95) * 0.35, 10)
            WK_INCOME = 10000000

        End If

    End If
    
    If JOBYMM >= "201101" And oChlWON > 0 Then
        oInWON = oInWON + oChlWON - 1
        oChlWON = 0
    End If
   
'/ ���̼�������ǥ ��ϵ� ���̺� ���� /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    sQry = " SELECT TOP 1 ISNULL(T0.U_CODAVR,0) AS U_CODAVR,"
    sQry = sQry & "       ISNULL(CASE WHEN " & oInWON & " <= 1  THEN U_BY01ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 2  THEN U_BY02ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 3  AND " & oChlWON & "  < 2 THEN U_BY03ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 3  AND " & oChlWON & " >= 2 THEN U_BY03DJ"
    sQry = sQry & "                   WHEN " & oInWON & "  = 4  AND " & oChlWON & "  < 2 THEN U_BY04ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 4  AND " & oChlWON & " >= 2 THEN U_BY04DJ"
    sQry = sQry & "                   WHEN " & oInWON & "  = 5  AND " & oChlWON & "  < 2 THEN U_BY05ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 5  AND " & oChlWON & " >= 2 THEN U_BY05DJ"
    sQry = sQry & "                   WHEN " & oInWON & "  = 6  AND " & oChlWON & "  < 2 THEN U_BY06ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 6  AND " & oChlWON & " >= 2 THEN U_BY06DJ"
    sQry = sQry & "                   WHEN " & oInWON & "  = 7  AND " & oChlWON & "  < 2 THEN U_BY07ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 7  AND " & oChlWON & " >= 2 THEN U_BY07DJ"
    sQry = sQry & "                   WHEN " & oInWON & "  = 8  AND " & oChlWON & "  < 2 THEN U_BY08ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 8  AND " & oChlWON & " >= 2 THEN U_BY08DJ"
    sQry = sQry & "                   WHEN " & oInWON & "  = 9  AND " & oChlWON & "  < 2 THEN U_BY09ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 9  AND " & oChlWON & " >= 2 THEN U_BY09DJ"
    sQry = sQry & "                   WHEN " & oInWON & "  = 10 AND " & oChlWON & "  < 2 THEN U_BY10ST"
    sQry = sQry & "                   WHEN " & oInWON & "  = 10 AND " & oChlWON & " >= 2 THEN U_BY10DJ"
    sQry = sQry & "                   WHEN " & oInWON & " >= 11 AND " & oChlWON & "  < 2 THEN U_BY11ST"
    sQry = sQry & "                   WHEN " & oInWON & " >= 11 AND " & oChlWON & " >= 2 THEN U_BY11DJ"
    sQry = sQry & "                   ELSE 0 END, 0) AS U_GABGUB "
    sQry = sQry & " FROM [@ZPY301L] T0 WHERE   T0.CODE <= '" & JOBYMM & "'"
    sQry = sQry & " AND     T0.U_CODFRS <= " & WK_INCOME & " AND     T0.U_CODTOM >  " & WK_INCOME & ""
    sQry = sQry & " ORDER BY T0.Code Desc"
    Rs.DoQuery sQry
    If Rs.RecordCount <> 0 Then
        WK_GULTAX = Rs.Fields("U_GABGUB").VALUE
    End If
    
'/ ���ټ� /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    
    GABGUN = IInt(GABGUN + WK_GULTAX, 10)
'/
   If GABGUN < 1000 Then GABGUN = 0
'/ ����ҵ漼(�ֹμ�) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
   JUMINN = IInt(GABGUN * 0.1, 10)
 '/ END
    Set Rs = Nothing
    
    Exit Function
'/ Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
Error_Message:
    Set Rs = Nothing
    Sbo_Application.StatusBar.SetText "[�������߻�: Get_GabGunSe_Table()]" & Space(10) & Err.Description, bmt_Short, smt_Error
End Function


Public Function Get_GabGunSe(ByRef GABGUN As Double, ByRef JUMINN As Double, ByVal oINCOME As Double, ByVal oInWON%, ByVal oChlWON%, ByVal JOBYMM As String, ByVal oKUKAMT As Double, ByVal PAY_001 As String) As Variant
On Error GoTo Error_Message
    Dim Rs     As SAPbobsCOM.Recordset
    Dim sQry    As String
    
    Dim WS_INCOME As Double
    Dim WK_INCOME As Double
    Dim WK_GNLOSD As Double
    Dim WK_SANTAX As Double
    Dim WK_TAXGON As Double
    Dim WK_KUKAMT As Double
    
    Dim WK_GULTAX As Double
    
 '/ Initial
    WK_INCOME = 0: WK_GNLOSD = 0: WK_SANTAX = 0: WK_TAXGON = 0: WK_KUKAMT = 0
    WS_INCOME = 0: WK_GULTAX = 0
        
    Set Rs = oCompany.GetBusinessObject(BoRecordset)
    GABGUN = 0
    JUMINN = 0
 '/ �����޾�
   If oINCOME <= 0 Then
       Get_GabGunSe = "�����ݾ��� 0���� �۰ų� �����ϴ�. Ȯ���Ͽ� �ּ���."
       Exit Function
    End If
'/  ���̼�������ǥ ���� ��հ��� ����� ���
    If PAY_001 = "2" Or PAY_001 = "3" Then
        sQry = " SELECT TOP 1 ISNULL(T0.U_CODAVR,0) AS U_CODAVR FROM [@ZPY301L] T0 WHERE   T0.CODE <= '" & JOBYMM & "'"
        sQry = sQry & " AND     T0.U_CODFRS <= " & oINCOME & " AND     T0.U_CODTOM >  " & oINCOME & ""
        sQry = sQry & " ORDER BY T0.Code Desc"
        Rs.DoQuery sQry
        If Rs.RecordCount <> 0 Then
            oINCOME = Rs.Fields("U_CODAVR").VALUE
            oKUKAMT = oINCOME
            WS_INCOME = oINCOME
        End If
    End If

    WK_INCOME = oINCOME
    WS_INCOME = oINCOME
    
    If JOBYMM <= "201201" Then
        
        '/ 1000�����ʰ���
        If oINCOME > 10000000 Then
           GABGUN = IInt(((oINCOME - 10000000) * 0.95) * 0.35, 10)
           WK_INCOME = 10000000
           WS_INCOME = 10000000
        End If

    Else

        If oINCOME > 28000000 Then
            
            '/ 2800�����ʰ���
            GABGUN = 5985000 + IInt(((oINCOME - 28000000) * 0.95) * 0.38, 10)
            
        ElseIf oINCOME > 10000000 Then
            
            '/ 1000�����ʰ���
            GABGUN = IInt(((oINCOME - 10000000) * 0.95) * 0.35, 10)
            
        End If

    End If
    
   '// 2008�������(������ü ������ ��������)
'   If Left(JOBYMM, 4) = "2008" Then
'        Select Case Trim(MDC_COMpanyGubun)
'        Case "OBS"
'        WS_INCOME = oINCOME
'        End Select
'   End If

'/
   WK_INCOME = WK_INCOME * 12
'/ �ٷμҵ����(2007.01����) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
If Left(JOBYMM, 4) <= "2008" Then
    '(�ٷμҵ�: 500��������              ���װ���
    '           500�����ʰ�~1500��������  500����+(�ٷμҵ�- 500����)*50%
    '          1500�����ʰ�~3000�������� 1000����+(�ٷμҵ�-1500����)*15%)
    '          4500��������              1225����+(�ٷμҵ�-3000����)*10%)
    '          4500�����ʰ�              1375����+(�ٷμҵ�-4500����)* 5%) �ѵ�����
     If WK_INCOME <= 5000000 Then
        WK_GNLOSD = WK_INCOME
     ElseIf WK_INCOME <= 15000000 Then
        WK_GNLOSD = 5000000 + (WK_INCOME - 5000000) * 0.5
     ElseIf WK_INCOME <= 30000000 Then        '3000
        WK_GNLOSD = 10000000 + (WK_INCOME - 15000000) * 0.15
     ElseIf WK_INCOME <= 45000000 Then        '4500
        WK_GNLOSD = 12250000 + (WK_INCOME - 30000000) * 0.1
     Else
        WK_GNLOSD = 13750000 + (WK_INCOME - 45000000) * 0.05
     End If
Else
    '/2009�� �ٷμҵ�����ݾ� ����
    '(�ٷμҵ�: 500��������              ����*80%
    '           500�����ʰ�~1500��������  400����+(�ٷμҵ�- 500����)*50%
    '          1500�����ʰ�~3000��������  900����+(�ٷμҵ�-1500����)*15%)
    '          4500��������              1125����+(�ٷμҵ�-3000����)*10%)
    '          4500�����ʰ�              1275����+(�ٷμҵ�-4500����)* 5%) �ѵ�����
     If WK_INCOME <= 5000000 Then
        WK_GNLOSD = WK_INCOME
     ElseIf WK_INCOME <= 15000000 Then
        WK_GNLOSD = 4000000 + (WK_INCOME - 5000000) * 0.5
     ElseIf WK_INCOME <= 30000000 Then        '3000
        WK_GNLOSD = 9000000 + (WK_INCOME - 15000000) * 0.15
     ElseIf WK_INCOME <= 45000000 Then        '4500
        WK_GNLOSD = 11250000 + (WK_INCOME - 30000000) * 0.1
     Else
        WK_GNLOSD = 12750000 + (WK_INCOME - 45000000) * 0.05
     End If
        
End If
    
'/ �ٷμҵ�ݾ� ( �ٷμҵ�-�ٷμҵ���� ) /
   WK_INCOME = WK_INCOME - WK_GNLOSD
'/ �⺻���� /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
   If Trim(JOBYMM) <= "200812" Then
    '  �������� 1�δ� 100����
    '  WK_INCOME = WK_INCOME - 1000000                   '/ 1.��        �� /
       WK_INCOME = WK_INCOME - (1000000 * oInWON)          '/ 2.�ξ簡������ /
   Else
    '  �������� 1�δ� 150����
    '  WK_INCOME = WK_INCOME - 1500000                   '/ 1.��        �� /
       WK_INCOME = WK_INCOME - (1500000 * oInWON)          '/ 2.�ξ簡������ /
   End If
'//(2007.01���� ���泻�� //////////////////////////////////////////////////////////////////////
'// �Ҽ��������߰����� ����
'// ���ڳ��߰����� �ż�: 20�������ڳడ 2�� 50����, 2���ʰ� 50���� +(2���ʰ��ο���*100����)
'//////////////////////////////////////////////////////////////////////////////////////////////
''/ �Ҽ��ο��߰����� /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
''  �Ҽ����� 1�� 100����, 2�� 50����
'   Select Case (oInWON)
'     Case 1: WK_INCOME = WK_INCOME - 1000000
'     Case 2: WK_INCOME = WK_INCOME - 500000
'   End Select
'/ ���ڳ��߰����� /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
   If (oChlWON) > 1 And oInWON > 2 Then
        If (oChlWON) <= 2 Then
            WK_INCOME = WK_INCOME - 500000
        Else
            WK_INCOME = WK_INCOME - 500000
            '/ 2009.05���� ������ ���ڳ� ���� 2���̻� �߰��ο��� �����ߴ��Ŵ� �״��
            If (PAY_001 = "1" Or PAY_001 = "2") Then
               WK_INCOME = WK_INCOME - (1000000 * (oChlWON - 2))
            End If
        End If
   End If

'/ Ư������(2�������ΰ��1,200,000 3���̻��ΰ�� 2,400,000)
'/ Ư������-2008��4�����ͺ���� /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'/ (2�������ΰ��: 1,200,000 => 100������ �����޿����� 25/1000�ش��ϴ� �ݾ��� �հ��
'/ (3���̻��ΰ��: 2,400,000 => 240������ �����޿����� 5/100�ش��ϴ� �ݾ��� �հ��+ �����޿��׿��� 4õ�����ʰ��ݾ��� 5/100
'/~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    If Trim(JOBYMM) <= "200712" Then
         If (oInWON) <= 2 Then
            WK_INCOME = WK_INCOME - (1000000 + (WS_INCOME * 12 * 2.5 / 100))
         Else
            WK_INCOME = WK_INCOME - (2400000 + (WS_INCOME * 12 * 5 / 100))
         End If
    Else
         If (oInWON) <= 2 Then
            WK_INCOME = WK_INCOME - (1100000 + (WS_INCOME * 12 * 2.5 / 100))
         Else
            WK_INCOME = WK_INCOME - (2500000 + (WS_INCOME * 12 * 5 / 100))
            If (WS_INCOME * 12) > 40000000 Then
                 WK_INCOME = WK_INCOME - ((WS_INCOME * 12 - 40000000) * 5 / 100)
            End If
         End If
    End If
    
'/ ���ݺ�������(2008.03�������� ���ο��ݵ����, 2008��04������ ���ο��ݺ���������
    If Trim(JOBYMM) <= "200712" Then
        '/ (���ο�������ǥ�� ����� *12)
        sQry = " SELECT  T0.U_EMPAMT, T0.U_FROM, T0.U_TO"
        sQry = sQry & " FROM [@ZPY102L] T0 INNER JOIN [@ZPY102H] T1 ON T0.Code = T1.Code"
        sQry = sQry & " WHERE T1.Code <= '" & JOBYMM & "'"
        sQry = sQry & " AND  T0.U_FROM <= " & WS_INCOME & ""
        sQry = sQry & " AND T0.U_TO > " & WS_INCOME & ""
        sQry = sQry & " ORDER BY T1.Code Desc"
        Rs.DoQuery sQry
        If Rs.RecordCount <> 0 Then
          WK_INCOME = IInt(WK_INCOME - (Rs.Fields("U_EMPAMT").VALUE * 12), 1)
        End If
    Else        '// 2008�� 4������
        sQry = " SELECT TOP 1 U_EMPRAT, U_FROM, U_TO FROM [@ZPY102H] "
        sQry = sQry & " WHERE CODE >= '200804' ORDER BY CODE DESC"
        Rs.DoQuery sQry
        If Rs.RecordCount <> 0 Then
            If Val(oKUKAMT) < Rs.Fields("U_FROM").VALUE Then
               WK_KUKAMT = Rs.Fields("U_FROM").VALUE
            ElseIf Rs.Fields("U_TO").VALUE > 0 And Val(oKUKAMT) > Rs.Fields("U_TO").VALUE Then
               WK_KUKAMT = Rs.Fields("U_TO").VALUE
            Else
               WK_KUKAMT = Val(oKUKAMT)
            End If
            WK_KUKAMT = MDC_SetMod.IInt(WK_KUKAMT * 12 * Val(Rs.Fields("U_EMPRAT").VALUE) / 100, 10)
            
            WK_INCOME = WK_INCOME - WK_KUKAMT
        End If
    End If
    '/ ����ǥ�� ( �ٷμҵ�ݾ� - �������� - Ư������ - ��Ÿ�ҵ���� ) /
    If WK_INCOME < 0 Then WK_INCOME = 0
'/ ���⼼�� /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    If Trim(JOBYMM) <= "200812" Then
        '/2008�⵵
        '(����ǥ��:1200��������               ����ǥ��*8%
        '          1200�����ʰ�~4600��������  ����ǥ��*17%-  96����
        '          4600�����ʰ�~8800��������  ����ǥ��*26%- 674����
        '          8800�����ʰ�               ����ǥ��*35%-1766����)
         If WK_INCOME <= 12000000 Then
            WK_SANTAX = WK_INCOME * 0.08 - 0
         ElseIf WK_INCOME <= 46000000 Then
            WK_SANTAX = WK_INCOME * 0.17 - 1080000
         ElseIf WK_INCOME <= 88000000 Then
            WK_SANTAX = WK_INCOME * 0.26 - 5220000
         Else
            WK_SANTAX = WK_INCOME * 0.35 - 13140000
         End If
    ElseIf Trim(JOBYMM) = "200912" Then
        '/2009�⵵
        '(����ǥ��:1200��������               ����ǥ��*6%
        '          1200�����ʰ�~4600��������  ����ǥ��*16%-  72����
        '          4600�����ʰ�~8800��������  ����ǥ��*26%- 616����
        '          8800�����ʰ�               ����ǥ��*35%-1666����)
         If WK_INCOME <= 12000000 Then
            WK_SANTAX = WK_INCOME * 0.06 - 0
         ElseIf WK_INCOME <= 46000000 Then
            WK_SANTAX = WK_INCOME * 0.16 - 1200000
         ElseIf WK_INCOME <= 88000000 Then
            WK_SANTAX = WK_INCOME * 0.25 - 5340000
         Else
            WK_SANTAX = WK_INCOME * 0.35 - 14140000
         End If
    ElseIf Trim(JOBYMM) <= "201201" Then
        '/2010�⵵
        '(����ǥ��:1200��������               ����ǥ��*6%
        '          1200�����ʰ�~4600��������  ����ǥ��*15%-  72����
        '          4600�����ʰ�~8800��������  ����ǥ��*24%- 582����
        '          8800�����ʰ�               ����ǥ��*35%-1590����)
         If WK_INCOME <= 12000000 Then
            WK_SANTAX = WK_INCOME * 0.06 - 0
         ElseIf WK_INCOME <= 46000000 Then
            WK_SANTAX = WK_INCOME * 0.15 - 1080000
         ElseIf WK_INCOME <= 88000000 Then
            WK_SANTAX = WK_INCOME * 0.24 - 5220000
         Else
            WK_SANTAX = WK_INCOME * 0.35 - 14900000
         End If
    
    Else
        
        '/2012�⵵
        '(����ǥ��:1200��������               ����ǥ��*6%
        '          1200�����ʰ�~4600��������  ����ǥ��*15%-  72����
        '          4600�����ʰ�~8800��������  ����ǥ��*24%-  582����
        '          8000�����ʰ�~3��� ����    ����ǥ��*35%-  1590����)
        '          3��� �ʰ�                 ����ǥ��*38%-  9010����)
         If WK_INCOME <= 12000000 Then
            WK_SANTAX = WK_INCOME * 0.06 - 0
         ElseIf WK_INCOME <= 46000000 Then
            WK_SANTAX = (WK_INCOME - 12000000) * 0.15 + 720000
         ElseIf WK_INCOME <= 88000000 Then
            WK_SANTAX = (WK_INCOME - 46000000) * 0.24 + 5820000
         ElseIf WK_INCOME <= 300000000 Then
            WK_SANTAX = (WK_INCOME - 88000000) * 0.35 + 15900000
         Else
            WK_SANTAX = (WK_INCOME - 300000000) * 0.38 + 90100000
         End If
         
    End If

   WK_SANTAX = IInt(WK_SANTAX, 1)
'/ ���װ���(2007.01 ����) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'  50��������  ���⼼�� * 55%
'  50�����ʰ�  275000 + (���⼼��-500000) * 30%
'  ���װ����ѵ���: 45�����ѵ�
   If WK_SANTAX <= 500000 Then
      WK_TAXGON = WK_SANTAX * 0.55
   Else
      WK_TAXGON = 275000 + (WK_SANTAX - 500000) * 0.3
   End If
'/
   WK_TAXGON = IInt(WK_TAXGON, 1)
   If WK_TAXGON > 500000 Then WK_TAXGON = 500000

'/ �������� ( ���⼼�� - ���װ��� �� ���� ) /
    WK_GULTAX = IInt((WK_SANTAX - WK_TAXGON) / 12, 10)

'/ ���ټ� /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    
    GABGUN = IInt(GABGUN + WK_GULTAX, 10)
'/
   If GABGUN < 1000 Then GABGUN = 0
'/ ����ҵ漼(�ֹμ�) /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
   JUMINN = IInt(GABGUN * 0.1, 10)
 '/ END
    Set Rs = Nothing
    
    Exit Function
'/ Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
Error_Message:
    Set Rs = Nothing
    Sbo_Application.StatusBar.SetText "[�������߻�: Get_GabGunSe()]" & Space(10) & Err.Description, bmt_Short, smt_Error
End Function

Public Function TermDay(StrDate As String, EndDate As String) As Integer
    Dim STRDAT  As Date
    Dim ENDDAT  As Date
    Dim i       As Integer
    If IsDate(Format(StrDate, "0000-00-00")) = False Or IsDate(Format(EndDate, "0000-00-00")) = False Then
        TermDay = 0
        Exit Function
    End If
    STRDAT = Format(StrDate, "0000-00-00")
    ENDDAT = Format(EndDate, "0000-00-00")
    i = DateDiff("d", STRDAT, ENDDAT) + 1
    TermDay = i
End Function


Public Function Lday(YMM As String)
'/ END Day
   Select Case True
     Case IsDate(Format(Mid(YMM, 1, 6) & "31", "0000-00-00")): Lday = "31"
     Case IsDate(Format(Mid(YMM, 1, 6) & "30", "0000-00-00")): Lday = "30"
     Case IsDate(Format(Mid(YMM, 1, 6) & "29", "0000-00-00")): Lday = "29"
     Case IsDate(Format(Mid(YMM, 1, 6) & "28", "0000-00-00")): Lday = "28"
     Case Else:                                                Lday = Space(0)
   End Select
End Function

Public Function CreateFolder(FileName$) As String
On Error GoTo Error_Message
    Dim fs As New FileSystemObject
    

    If fs.FolderExists(FileName) = False Then fs.CreateFolder FileName
        
    CreateFolder = ""
    
    Set fs = Nothing
    Exit Function
Error_Message:
    Set fs = Nothing
    CreateFolder = Err.Number & Space(1) & Err.Description
End Function

Public Function sStr(GetStr As String) As String
    sStr = LeftB(StrConv(GetStr, vbFromUnicode), Len(GetStr))
    sStr = Left(StrConv(sStr, vbUnicode), Len(GetStr))
    If Asc(Right(sStr, 1)) = 0 Then Mid(sStr, Len(sStr), 1) = Space(1)
End Function

Public Sub Get_FormColor()
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '����� ���� ��������� ������ �⺻ ������ ����������
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
        Dim oRecordSet      As SAPbobsCOM.Recordset
        Dim sQry            As String
        Dim StringColor     As String   '��� ����
        
        Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
        sQry = "Select Color from OADM"
        oRecordSet.DoQuery sQry
        
        Do Until oRecordSet.EOF
            StringColor = Trim(oRecordSet.Fields(0).VALUE)
            oRecordSet.MoveNext
        Loop
        
        If StringColor = 1 Then
            Sbo_Application.ActivateMenuItem ("5633")
        ElseIf StringColor = 2 Then
            Sbo_Application.ActivateMenuItem ("5634")
        ElseIf StringColor = 3 Then
            Sbo_Application.ActivateMenuItem ("5635")
        ElseIf StringColor = 4 Then
            Sbo_Application.ActivateMenuItem ("5636")
        ElseIf StringColor = 5 Then
            Sbo_Application.ActivateMenuItem ("5637")
        ElseIf StringColor = 6 Then
            Sbo_Application.ActivateMenuItem ("5638")
        ElseIf StringColor = 7 Then
            Sbo_Application.ActivateMenuItem ("5639")
        ElseIf StringColor = 8 Then
            Sbo_Application.ActivateMenuItem ("5640")
        ElseIf StringColor = 9 Then
            Sbo_Application.ActivateMenuItem ("5641")
        End If
        
    Set oRecordSet = Nothing
End Sub

Public Function Get_UserName(oUserSign$)
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry$
    
    If oUserSign <> "" Then
        sQry = "SELECT U_NAME FROM OUSR"
        sQry = sQry + " WHERE USERID = '" & oUserSign & "'"
        
        Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
        oRecordSet.DoQuery sQry
        Do Until oRecordSet.EOF
            Get_UserName = oRecordSet.Fields(0).VALUE
            oRecordSet.MoveNext
        Loop
        If Trim(Get_UserName) = "" Then
            Get_UserName = ""
        End If
    Else
        Get_UserName = ""
    End If
    
    Set oRecordSet = Nothing
End Function

Public Function Month_LastDay(JOBDAT As String) As String
    '***************************************************************************
    'Function ID : Month_LastDay(�λ�޿���⿡�� ���)
    '��    ��    : �ش�� �ϼ� ����
    '��    ��    :
    '�� ȯ ��    :
    'Ư�̻���    :
    '***************************************************************************
    '/ �ش���� ������ ����
    Select Case True
      Case IsDate(Format(JOBDAT & "31", "0000-00-00")): Month_LastDay = 31
      Case IsDate(Format(JOBDAT & "30", "0000-00-00")): Month_LastDay = 30
      Case IsDate(Format(JOBDAT & "29", "0000-00-00")): Month_LastDay = 29
      Case IsDate(Format(JOBDAT & "28", "0000-00-00")): Month_LastDay = 28
    End Select

End Function


Public Function TableFieldCheck(sTable As String, sField1 As String, Optional sField2 As String) As Boolean
    '***************************************************************************
    'Function ID :TableFieldCheck
    '��    ��    : ���̺� ���� ������ �ش� ���̺��� �ʵ�� ���� üũ
    '��    ��    :sTable, sField, sField2
    '�� ȯ ��    :True,False
    'Ư�̻���    :
    '***************************************************************************
    Dim i           As Long
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    TableFieldCheck = False
    
    sQry = "SELECT * FROM sysobjects WHERE name = '" & sTable & "' AND xtype='U'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        Sbo_Application.SetStatusBarMessage "�Է��Ͻ� [" & sTable & "���̺��� ���� ���� �ʽ��ϴ�.", bmt_Short, True
        Exit Function
    End If
    
    sQry = "select * from INFORMATION_SCHEMA.COLUMNS where table_name='" & sTable & "' and column_name= '" & sField1 & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        Sbo_Application.SetStatusBarMessage "�Է��Ͻ� [" & sField1 & "] �ʵ���� ���� ���� �ʽ��ϴ�.", bmt_Short, True
        Exit Function
    End If
    
    sQry = "select * from INFORMATION_SCHEMA.COLUMNS where table_name='" & sTable & "' and column_name= '" & sField2 & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        Sbo_Application.SetStatusBarMessage "�Է��Ͻ� [" & sField2 & "] �ʵ���� ���� ���� �ʽ��ϴ�.", bmt_Short, True
        Exit Function
    End If
    
    TableFieldCheck = True

End Function



Public Sub AuthorityCheck(oForm As SAPbouiCOM.Form, Item As String, Table As String, DocType As String)

    '***************************************************************************
    'Function ID :AuthorityCheck
    '��    ��    :������ ���ѿ� ���� ������ ����
    '��    ��    :oForm, Item(������ �޴� ������) , Table( ��) @PH_PY001 ) , sType ( ������ : Code, ���� : DocEntry )
    '�� ȯ ��    :Code or DocEntry
    'Ư�̻���    :
    '***************************************************************************
    Dim i           As Long
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    AuthorityCheck = False
    
    sQry = "UPDATE [" & Table & "] SET U_NaviDoc = NULL"
    oRecordSet.DoQuery sQry
            
    sQry = "UPDATE [" & Table & "] SET U_NaviDoc = " & DocType & " WHERE U_" & Item & " IN ("
    sQry = sQry & " SELECT U_Value"
    sQry = sQry & " From [@PH_PY000B] T0 INNER JOIN [@PH_PY000A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE T1.Code = '" & Item & "' AND T0.U_UserCode = '" & oCompany.UserName & "' Group By U_Value)"
    
    oRecordSet.DoQuery sQry

    oForm.DataBrowser.BrowseBy = "NaviDoc"
End Sub


Public Sub CLTCOD_Select(oForm As SAPbouiCOM.Form, Item As String, Optional AuthorityUse As Boolean = True)

    '***************************************************************************
    'Function ID :AuthorityCheck
    '��    ��    :������ ���ѿ� ���� ����� �޺��ڽ� ����
    '��    ��    :oForm, item (CLTCOD), AuthorityUse (True:���ѿ� ���������ǥ�� , False:��ü�����ǥ��)
    '�� ȯ ��    :Code or DocEntry
    'Ư�̻���    :
    '***************************************************************************
    Dim i           As Long
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oCombo = oForm.Items(Item).Specific
    If oCombo.ValidValues.Count > 0 Then
        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
            oCombo.ValidValues.Remove i, psk_Index
        Next
        
    End If
    
    If AuthorityUse = True Then
        sQry = " SELECT T2.Code ,T2.Name"
        sQry = sQry & " From [@PH_PY000B] T0 INNER JOIN [@PH_PY000A] T1 ON T0.Code = T1.Code"
        sQry = sQry & " INNER JOIN [@PH_PY005A] T2 ON T0.U_Value = T2.Code"
        sQry = sQry & " WHERE T1.Code = 'CLTCOD' AND T0.U_UserCode = '" & oCompany.UserName & "'"
        sQry = sQry & " GROUP BY T2.Code ,T2.Name ORDER BY T2.Code"
        
        oRecordSet.DoQuery sQry
        
        
        If oRecordSet.RecordCount > 0 Then
            Do Until oRecordSet.EOF
                oCombo.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
                oRecordSet.MoveNext
            Loop
            oCombo.Select "" & MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'") & "", psk_ByValue
        Else
            oCombo.ValidValues.Add "", "-"
        End If
    Else '//false
        sQry = "SELECT Code, Name FROM [@PH_PY005A] "
        oRecordSet.DoQuery sQry
        
        If oRecordSet.RecordCount > 0 Then
            Do Until oRecordSet.EOF
                oCombo.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
                oRecordSet.MoveNext
            Loop
        Else
            oCombo.ValidValues.Add "", "-"
        End If
    End If
End Sub

Public Sub PAY_Matrix_AddCol(oMatrix As SAPbouiCOM.Matrix, Col$, iE%, Tn$, Wt As Double, Et As Boolean, St As Boolean, BouYN As Boolean, TableNAM$, FieldNAM$)
'On Error GoTo error_Message
    '***************************************************************************
    'Function ID : PAY_Matrix_AddCol(�λ�޿���⿡�� ���)
    '��    ��    : ��Ʈ������ �÷� �߰�
    '�� �� ��    : �Թ̰�
    'Ư�̻���    : ��ȸ�÷��� ���� ��� Col00��ȣ�÷��� ��ũ�������ͷ� �����
    '              ������ �̰ɷ� ���� ���� ��������~
    'Mat:��Ʈ����uid, col:�÷�Uid, iE:�÷�����-[edit(16),�޺�(113), üũ(122), ��ũ(116)], Tn:�÷�Ÿ��Ʋ��, Wt:�ʺ�, Et:Editable true/false��
    'bouYN:DataBind ����, TableNam:���̺��, FieldNam:�ʵ��, Rt:���������Ŀ���
    '***************************************************************************
    
    Dim oCols    As SAPbouiCOM.Columns
    Dim oCol     As SAPbouiCOM.Column
    
    Set oCols = oMatrix.Columns
    
    Set oCol = oCols.Add(Col, iE)
    oCol.DataBind.SetBound BouYN, TableNAM, FieldNAM     '/ UI����� UserDataSources bound�����������
    oCol.TitleObject.Caption = Tn
    oCol.Width = Wt
    oCol.Editable = Et
    oCol.RightJustified = St
    Set oCols = Nothing
    Set oCol = Nothing
    Exit Sub
'/ Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
Error_Message:
    Set oCols = Nothing
    Set oCol = Nothing
    Sbo_Application.Err_Message ("[�������߻�: PAY_Matrix_AddCol()]" & Space(10) & Err.Description)
End Sub


Public Sub Set_ComboList(Lst As Object, sSQL As String, Optional TValue As String = "", Optional Reset As Boolean, Optional SetF As Boolean)
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    'ComboBox Object,Query,SD380 value
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    Dim ComBox          As New SAPbouiCOM.ComboBox
    Dim s_Recordset      As SAPbobsCOM.Recordset
    
    Set s_Recordset = oCompany.GetBusinessObject(BoRecordset)
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
