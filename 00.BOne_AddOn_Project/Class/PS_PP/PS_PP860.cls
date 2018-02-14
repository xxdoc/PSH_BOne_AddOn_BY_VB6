VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP860"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP860.cls
'//  Module         : PP
'//  Desc           : 월생산계획대비실적조회
'//  FormType       : PS_PP860
'//  Create Date    : 2011.10.18   NGY
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_PP860H As SAPbouiCOM.DBDataSource '등록헤더
'Private oDS_PS_PP860H As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    LotNo As String
    Quantity As Long
    OPORNo As Long
    POR1No As Long
    Check As Boolean
    OPDNNo As Long
    PDN1No As Long
End Type

Private oLast_Mode&

Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP860.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP860_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_FIND_MODE
    '//oForm01.DataBrowser.BrowseBy = "BPLId" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP860_CreateItems
    Call PS_PP860_ComboBox_Setting
    Call PS_PP860_CF_ChooseFromList
    Call PS_PP860_EnableMenus
    Call PS_PP860_SetDocument(oFromDocEntry01)
    Call PS_PP860_FormResize
    
    oForm01.Items("YYYYMM").Click ct_Regular
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub
'//조회데이타 가져오기
Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId$, YYYYMM$, Part$
    
    Dim SItemCode, DocDateF$, DocDateT$, SGbn$
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    YYYYMM = Trim(oForm01.Items("YYYYMM").Specific.VALUE)
    Part = Trim(oForm01.Items("Part").Specific.VALUE)
    
    oForm01.Freeze True
    
    If Left(Part, 3) = "124" Then '부품
       sQry = "EXEC [PS_PP860_02] '" & BPLId & "','" & YYYYMM & "','" & Part & "'"
    Else
       sQry = "EXEC [PS_PP860_01] '" & BPLId & "','" & YYYYMM & "','" & Part & "'"
    End If
    
    oRecordSet01.DoQuery sQry

    oMat01.Clear
    oDS_PS_PP860H.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

'    If (oRecordSet01.RecordCount = 0) Then
'         MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
'        Set oRecordSet01 = Nothing
'        oForm01.Mode = fm_ADD_MODE
'
'        Call Add_MatrixRow(0, True)
'        Call LoadCaption
'        Exit Sub
'    End If

    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)

    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP860H.Size Then
            oDS_PS_PP860H.InsertRecord (i)
        End If

        oMat01.AddRow
        oDS_PS_PP860H.Offset = i

        'oDS_PS_PP860H.setValue "U_LineNum", i, i + 1
        oDS_PS_PP860H.setValue "U_ColRgl01", i, Trim(oRecordSet01.Fields(0).VALUE)
        oDS_PS_PP860H.setValue "U_ColRgl02", i, Trim(oRecordSet01.Fields(1).VALUE)
        oDS_PS_PP860H.setValue "U_ColRgl03", i, Trim(oRecordSet01.Fields(2).VALUE)
        oDS_PS_PP860H.setValue "U_ColRgl04", i, Trim(oRecordSet01.Fields(3).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields(4).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields(5).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields(6).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields(7).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields(8).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields(9).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg07", i, Trim(oRecordSet01.Fields(10).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg08", i, Trim(oRecordSet01.Fields(11).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg09", i, Trim(oRecordSet01.Fields(12).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg10", i, Trim(oRecordSet01.Fields(13).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg11", i, Trim(oRecordSet01.Fields(14).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg12", i, Trim(oRecordSet01.Fields(15).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg13", i, Trim(oRecordSet01.Fields(16).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg14", i, Trim(oRecordSet01.Fields(17).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg15", i, Trim(oRecordSet01.Fields(18).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg16", i, Trim(oRecordSet01.Fields(19).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg17", i, Trim(oRecordSet01.Fields(20).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg18", i, Trim(oRecordSet01.Fields(21).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg19", i, Trim(oRecordSet01.Fields(22).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg20", i, Trim(oRecordSet01.Fields(23).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg21", i, Trim(oRecordSet01.Fields(24).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg22", i, Trim(oRecordSet01.Fields(25).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg23", i, Trim(oRecordSet01.Fields(26).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg24", i, Trim(oRecordSet01.Fields(27).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg25", i, Trim(oRecordSet01.Fields(28).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg26", i, Trim(oRecordSet01.Fields(29).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg27", i, Trim(oRecordSet01.Fields(30).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg28", i, Trim(oRecordSet01.Fields(31).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg29", i, Trim(oRecordSet01.Fields(32).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg30", i, Trim(oRecordSet01.Fields(33).VALUE)
        oDS_PS_PP860H.setValue "U_ColReg31", i, Trim(oRecordSet01.Fields(34).VALUE)
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    
    
    
    
    
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub CalcData() '//일일생산량 계산
On Error GoTo CalcData_Error
    Dim i As Integer
    Dim D(30)
    Dim N(30)
    
    Dim DInwon, NInwon, MkCapa
    
    Dim Lastdd As Integer
        
    Dim YYYYMM$
    
    Dim sQry01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    YYYYMM = Trim(oForm01.Items("YYYYMM").Specific.VALUE) & "01"
    
    sQry01 = "Select right(Convert(Char(8),DateAdd(dd, -1, DateAdd(mm, 1,'" & YYYYMM & "')),112),2)"
    
    RecordSet01.DoQuery sQry01
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        MDC_Com.MDC_GF_Message "날짜 계산이 틀립니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set RecordSet01 = Nothing
        GoTo CalcData_Error
    Else
        Lastdd = Trim(RecordSet01.Fields(0).VALUE)
    End If

    
    Call oForm01.Freeze(True)
    For i = 1 To Lastdd
        D(i - 1) = Trim(oForm01.Items("D" & i).Specific.VALUE)
        N(i - 1) = Trim(oForm01.Items("N" & i).Specific.VALUE)
    Next
    
    For i = 1 To oMat01.RowCount '//
        DInwon = oMat01.Columns("DInwon").Cells(i).Specific.VALUE
        NInwon = oMat01.Columns("NInwon").Cells(i).Specific.VALUE
        MkCapa = oMat01.Columns("MkCapa").Cells(i).Specific.VALUE
        
        Call oDS_PS_PP860H.setValue("U_ColSum01", i - 1, DInwon)
        Call oDS_PS_PP860H.setValue("U_ColSum02", i - 1, NInwon)
        Call oDS_PS_PP860H.setValue("U_ColSum03", i - 1, MkCapa)
        Call oDS_PS_PP860H.setValue("U_ColReg01", i - 1, (Val(D(0)) + Val(N(0))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg02", i - 1, (Val(D(1)) + Val(N(1))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg03", i - 1, (Val(D(2)) + Val(N(2))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg04", i - 1, (Val(D(3)) + Val(N(3))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg05", i - 1, (Val(D(4)) + Val(N(4))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg06", i - 1, (Val(D(5)) + Val(N(5))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg07", i - 1, (Val(D(6)) + Val(N(6))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg08", i - 1, (Val(D(7)) + Val(N(7))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg09", i - 1, (Val(D(8)) + Val(N(8))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg10", i - 1, (Val(D(9)) + Val(N(9))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg11", i - 1, (Val(D(10)) + Val(N(10))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg12", i - 1, (Val(D(11)) + Val(N(11))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg13", i - 1, (Val(D(12)) + Val(N(12))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg14", i - 1, (Val(D(13)) + Val(N(13))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg15", i - 1, (Val(D(14)) + Val(N(14))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg16", i - 1, (Val(D(15)) + Val(N(15))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg17", i - 1, (Val(D(16)) + Val(N(16))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg18", i - 1, (Val(D(17)) + Val(N(17))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg19", i - 1, (Val(D(18)) + Val(N(18))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg20", i - 1, (Val(D(19)) + Val(N(19))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg21", i - 1, (Val(D(20)) + Val(N(20))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg22", i - 1, (Val(D(21)) + Val(N(21))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg23", i - 1, (Val(D(22)) + Val(N(22))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg24", i - 1, (Val(D(23)) + Val(N(23))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg25", i - 1, (Val(D(24)) + Val(N(24))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg26", i - 1, (Val(D(25)) + Val(N(25))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg27", i - 1, (Val(D(26)) + Val(N(26))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg28", i - 1, (Val(D(27)) + Val(N(27))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg29", i - 1, (Val(D(28)) + Val(N(28))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg30", i - 1, (Val(D(29)) + Val(N(29))) * Val(MkCapa))
        Call oDS_PS_PP860H.setValue("U_ColReg31", i - 1, (Val(D(30)) + Val(N(30))) * Val(MkCapa))

    Next
    
    
    
    
    oForm01.Mode = fm_ADD_MODE
    
        

    oMat01.Clear
    'oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CalcData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "CalcData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP860H.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_PP860H.Offset = oRow
    oDS_PS_PP860H.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'// 데이타 Update

'// 데이타 insert
Public Function Add_PurchaseDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_PurchaseDemand_Error
    Dim i As Integer
    Dim j As Integer
    Dim sQry01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim sQry02 As String
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim YYYYMM$, BPLId$, DocEntry$, ItmBsort&, ItmBName&, ItmMsort$, ItmMname$, ItemCode$, ItemName$, DocDate$, ymd$
    Dim DInwon, NInwon, MkCapa

    Dim Lastdd
    
    Dim D(30)
    Dim N(30)
    
    Dim MP(30)
        
    BPLId = "1"
    YYYYMM = Trim(oForm01.Items("YYYYMM").Specific.VALUE)
    
    ymd = Trim(oForm01.Items("YYYYMM").Specific.VALUE) & "01"
    
    ItmBsort = Trim(oForm01.Items("ItmBsort").Specific.VALUE)
    
    
    '//저장시 삭제후 재등록
    sQry02 = "Delete From [@PS_PP860H] Where Convert(Char(6),U_DocDate,112) ='" & YYYYMM & "'"
    RecordSet02.DoQuery sQry02
'    If Trim(RecordSet02.Fields(0).VALUE) = 0 Then
'        MDC_Com.MDC_GF_Message "기존자료 삭제오류입니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
'        Set RecordSet02 = Nothing
'        GoTo Add_PurchaseDemand_Error
'    End If
    
    '//해당월 최종일자 계산
    sQry01 = "Select right(Convert(Char(8),DateAdd(dd, -1, DateAdd(mm, 1,'" & ymd & "')),112),2)"
    RecordSet01.DoQuery sQry01
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        MDC_Com.MDC_GF_Message "날짜 계산이 틀립니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set RecordSet01 = Nothing
        GoTo Add_PurchaseDemand_Error
    Else
        Lastdd = Trim(RecordSet01.Fields(0).VALUE)
    End If
    
    
    '주야간 근무시간
    For i = 1 To Lastdd
        D(i - 1) = Trim(oForm01.Items("D" & i).Specific.VALUE)
        N(i - 1) = Trim(oForm01.Items("N" & i).Specific.VALUE)
    Next
    
    For i = 1 To oMat01.RowCount '//
        ItmMsort = oMat01.Columns("ItmMsort").Cells(i).Specific.VALUE
        ItmMname = oMat01.Columns("ItmMname").Cells(i).Specific.VALUE
        ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemName = oMat01.Columns("ItemName").Cells(i).Specific.VALUE
        DInwon = oMat01.Columns("DInwon").Cells(i).Specific.VALUE
        NInwon = oMat01.Columns("NInwon").Cells(i).Specific.VALUE
        MkCapa = oMat01.Columns("MkCapa").Cells(i).Specific.VALUE
        
        
        
        
        For j = 1 To Lastdd
            sQry01 = "Select IsNull(Max(DocEntry), 0) From [@PS_PP860H]"
            RecordSet01.DoQuery sQry01
            If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
                DocEntry = 1
            Else
                DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
            End If
            
            
            MP(j - 1) = oMat01.Columns("D" & j).Cells(i).Specific.VALUE
            
            If j < 10 Then
                DocDate = YYYYMM & "0" & j
            Else
                DocDate = YYYYMM & j
            End If
                
            
            sQry02 = "INSERT INTO [@PS_PP860H]"
            sQry02 = sQry02 & " ("
            sQry02 = sQry02 & " DocEntry,"
            sQry02 = sQry02 & " DocNum,"
            sQry02 = sQry02 & " U_BPLId,"
            sQry02 = sQry02 & " U_DocDate,"
            sQry02 = sQry02 & " U_ItmBsort,"
            sQry02 = sQry02 & " U_ItmBname,"
            sQry02 = sQry02 & " U_ItmMsort,"
            sQry02 = sQry02 & " U_ItmMname,"
            sQry02 = sQry02 & " U_ItemCode,"
            sQry02 = sQry02 & " U_ItemName,"
            sQry02 = sQry02 & " U_DInwon,"
            sQry02 = sQry02 & " U_NInwon,"
            sQry02 = sQry02 & " U_MkCapa,"
            sQry02 = sQry02 & " U_PQTy,"
            sQry02 = sQry02 & " U_DTime,"
            sQry02 = sQry02 & " U_NTime"
            sQry02 = sQry02 & " ) "
            sQry02 = sQry02 & "VALUES("
            sQry02 = sQry02 & DocEntry & ","
            sQry02 = sQry02 & DocEntry & ","
            sQry02 = sQry02 & "'" & BPLId & "',"
            sQry02 = sQry02 & "'" & DocDate & "',"
            sQry02 = sQry02 & "'" & ItmBsort & "',"
            sQry02 = sQry02 & "'" & ItmBsort & "',"
            sQry02 = sQry02 & "'" & ItmMsort & "',"
            sQry02 = sQry02 & "'" & ItmMname & "',"
            sQry02 = sQry02 & "'" & ItemCode & "',"
            sQry02 = sQry02 & "'" & ItemName & "',"
            sQry02 = sQry02 & "'" & DInwon & "',"
            sQry02 = sQry02 & "'" & NInwon & "',"
            sQry02 = sQry02 & "'" & Val(MkCapa) & "',"
            sQry02 = sQry02 & "'" & Val(MP(j - 1)) & "',"
            sQry02 = sQry02 & "'" & Val(D(j - 1)) & "',"
            sQry02 = sQry02 & "'" & Val(N(j - 1)) & "'"
            sQry02 = sQry02 & ")"
            RecordSet02.DoQuery sQry02
        Next
        
    Next
    
    
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"

    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Add_PurchaseDemand = True
    
    
    
    
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_PurchaseDemand_Error:
    Add_PurchaseDemand = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Add_PurchaseDemand_Error:" & Err.Number & " - " & Err.Description, "E"
End Function
'/ 입력필수 사항 Check
Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("YYYYMM").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Part").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
  
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "년월은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "담당은 필수사항입니다. 확인하세요.", "E"
    End If
    HeaderSpaceLineDel = False
End Function
'/ 메트릭스 필수 사항 check
Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String


    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 사원코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 시간이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 등록일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 비가동코드가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function
'/폼의 아이템 사용지정
Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        'oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub



'/아이템 변경 이벤트
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error


    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            Call Raise_EVENT_ITEM_PRESSED(FormUID, pval, BubbleEvent)
        Case et_KEY_DOWN: '//2
            If pval.CharPressed = 9 Then
                If pval.ItemUID = "ItemCode" Then
                    If oForm01.Items("ItemCode").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                 
                End If
            End If
           
            Call Raise_EVENT_KEY_DOWN(FormUID, pval, BubbleEvent)
        Case et_COMBO_SELECT: '//5
            Call Raise_EVENT_COMBO_SELECT(FormUID, pval, BubbleEvent)
        Case et_CLICK: '//6
            Call Raise_EVENT_CLICK(FormUID, pval, BubbleEvent)
        Case et_DOUBLE_CLICK: '//7
            Call Raise_EVENT_DOUBLE_CLICK(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LINK_PRESSED '//8
            Call Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, pval, BubbleEvent)
        Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    
                End If
           ' Call Raise_EVENT_VALIDATE(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LOAD: '//11
            Call Raise_EVENT_MATRIX_LOAD(FormUID, pval, BubbleEvent)
        Case et_FORM_ACTIVATE: '//18
            '//et_FORM_ACTIVATE
        Case et_FORM_DEACTIVATE: '//19
            '//et_FORM_DEACTIVATE
        Case et_FORM_RESIZE '//20
            Call Raise_EVENT_RESIZE(FormUID, pval, BubbleEvent)
        Case et_CHOOSE_FROM_LIST '//27
            Call Raise_EVENT_CHOOSE_FROM_LIST(FormUID, pval, BubbleEvent)
        Case et_GOT_FOCUS: '//3
            Call Raise_EVENT_GOT_FOCUS(FormUID, pval, BubbleEvent)
        Case et_LOST_FOCUS: '//4
            '//et_LOST_FOCUS
        Case et_FORM_UNLOAD: '//17
            Call Raise_EVENT_FORM_UNLOAD(FormUID, pval, BubbleEvent)
    End Select
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error

Dim sQry As String
Dim RecordSet01 As SAPbobsCOM.Recordset
Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
               
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_PP860_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_PP860H.Clear

'                Call LoadCaption
'                Call FormItemEnabled
                '//Call PS_PP860_FormItemEnabled '//UDO방식
                '//Call PS_PP860_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_PP860_FormItemEnabled
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
Raise_FormDataEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
    End If
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "PS_PP860" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
                
        If pval.ItemUID = "Btn_save" Then '/추가/확인 버튼클릭
            If oForm01.Mode = fm_ADD_MODE Then
                If HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If Add_PurchaseDemand(pval) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                Call FormReset
                oForm01.Mode = fm_ADD_MODE
                
'                Call LoadCaption
                Call LoadData
                
                oLast_Mode = oForm01.Mode
                
                
            End If
            
        ElseIf pval.ItemUID = "Btn_Calc" Then '/계산
             If HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
                
             End If
            Call CalcData
        ElseIf pval.ItemUID = "Btn_ret" Then '/조회
            If HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            End If
            Call PS_PP860_MatrixColumnSetting(oForm01.Items("YYYYMM").Specific.VALUE & "01")
            Call LoadData
            
            
            'Call PS_PP860_AddMatrixRow(oMat01.VisualRowCount)
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP860" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    If pval.BeforeAction = True Then
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
'                Call oMat01.SelectRow(pval.Row, True, False)
'
'                oForm01.Items("DocEntry").Specific.VALUE = oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE
''                oForm01.Items("BPLId").Specific.VALUE = oMat01.Columns("BPLId").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("DocDate").Specific.VALUE = oMat01.Columns("DocDate").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("ItemCode").Specific.VALUE = oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("Qty").Specific.VALUE = oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("Amt").Specific.VALUE = oMat01.Columns("Amt").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("Comments").Specific.VALUE = oMat01.Columns("Comments").Cells(pval.Row).Specific.VALUE
'                'oForm01.Items("Gbn").Specific.VALUE = "1" 'oMat01.Columns("Gbn").Cells(pval.Row).Specific.VALUE
'                Call oForm01.Items("Gbn").Specific.Select(oMat01.Columns("Gbn").Cells(pval.Row).Specific.VALUE, psk_ByValue)
'                'oForm01.Items("CpCode").Specific.Select 1, psk_Index
'                oForm01.Mode = fm_UPDATE_MODE
'                Call LoadCaption
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

Dim i&
Dim sQry            As String

Dim sCount&, sSeq&
Dim sCode As String
Dim sCpCode As String

Dim oRecordSet01      As SAPbobsCOM.Recordset

Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_DOUBLE_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_DOUBLE_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LINK_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LINK_PRESSED_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
'/현재 사용안함
Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    If pval.BeforeAction = True Then
'        If pval.ItemChanged = True Then
'            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_PP860H.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP860H.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_PP860_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_PP860H.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                End If
'            Else
'                If (pval.ItemUID = "DocEntry") Then
'                    Call oDS_PS_PP860H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_PP860H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_PP860H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_PP860H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                End If
'            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
'        End If
    ElseIf pval.BeforeAction = False Then

    End If
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_PP860_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP860_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "ItemCode") Then
'            Dim oDataTable01 As SAPbouiCOM.DataTable
'            Set oDataTable01 = pval.SelectedObjects
'            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP860H", "U_CardCode,U_CardName")
'        End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
    End If
    Exit Sub
Raise_EVENT_GOT_FOCUS_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_GOT_FOCUS_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_FORM_UNLOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_FORM_UNLOAD_Error
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
        RemoveForms oFormUniqueID01
        Set oForm01 = Nothing
        Set oMat01 = Nothing
    End If
    Exit Sub
Raise_EVENT_FORM_UNLOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_UNLOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
'            If (PS_PP860_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_PP860H.RemoveRecord(oDS_PS_PP860H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_PP860H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP860_CreateItems() As Boolean
On Error GoTo PS_PP860_CreateItems_Error
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP860H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
'    Set oDS_PS_PP860H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '/유저데이타 속성 선언 날짜형식
    
    Set oRecordSet01 = Nothing
    Exit Function
PS_PP860_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP860_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
'/콤보박스 set
Sub PS_PP860_ComboBox_Setting()
On Error GoTo PS_PP860_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Items("YYYYMM").Specific.VALUE = Format(Date, "YYYYMM")
    
    Call oForm01.Items("BPLId").Specific.ValidValues.Add("", "")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
    
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_PP860_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP860_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP860_CF_ChooseFromList()
On Error GoTo PS_PP860_CF_ChooseFromList_Error

    Exit Sub
PS_PP860_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP860_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP860_FormItemEnabled()
On Error GoTo PS_PP860_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_PP860_FormClear '//UDO방식
'        Call oForm01.EnableMenu("1281", True) '//찾기
'        Call oForm01.EnableMenu("1282", False) '//추가
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
'        Call oForm01.EnableMenu("1281", False) '//찾기
'        Call oForm01.EnableMenu("1282", True) '//추가
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Exit Sub
PS_PP860_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP860_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP860_EnableMenus()
On Error GoTo PS_PP860_EnableMenus_Error
    
    Exit Sub
PS_PP860_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP860_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP860_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP860_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP860_FormItemEnabled
        '//Call PS_PP860_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP860_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP860_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP860_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP860_FormResize()
On Error GoTo PS_PP860_FormResize_Error
    
    Exit Sub
PS_PP860_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP860_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_PP860_MatrixColumnSetting(ByVal pYM As String)
On Error GoTo PS_PP860_MatrixColumnSetting_Error
    
    Call oForm01.Freeze(True)
    
        Dim Dt As String
        Dt = CStr(DateAdd("D", -1, DateAdd("M", 1, Format(Left(pYM, 6), "&&&&-&&-01")))) '해당 월의 마지막 날
   
        Dim LastDay As Integer
        LastDay = CInt(Right(Dt, 2))
        
        Dim loopCount As Integer
        
        For loopCount = 1 To oMat01.Columns.Count - 1 'LineNum는 제외하기 위해서 1부터 시작
        
                oMat01.Columns(loopCount).Editable = True

       
        Next
        
'        oMat01.AutoResizeColumns
        
        Dim DisableColumn As Integer
        Dim DisableColumnString As String
        
        Dim DayNum As String
        Dim temp As String
        Dim DayName As String
        
        For loopCount = 1 To LastDay
            If loopCount < 10 Then
                temp = "0"
            Else
                temp = ""
            End If
            
            DayNum = Weekday(Format(Left(pYM, 6) & temp & CStr(loopCount), "&&&&-&&-&&"))
            Select Case DayNum
                Case "1"
                    DayName = "일"
                Case "2"
                    DayName = "월"
                Case "3"
                    DayName = "화"
                Case "4"
                    DayName = "수"
                Case "5"
                    DayName = "목"
                Case "6"
                    DayName = "금"
                Case "7"
                    DayName = "토"
            End Select
                
                
            DisableColumnString = "D" & temp & CStr(loopCount)
            oMat01.Columns.Item(DisableColumnString).TitleObject.Caption = CStr(loopCount) & "일(" & DayName & ")"
'            If DayName = "토" Or DayName = "일" Then
'                oMat01.Columns.Item(DisableColumnString).ForeColor = 254 '빨간색
'            Else
'                oMat01.Columns.Item(DisableColumnString).ForeColor = 0 '검은색
'            End If
            
            If DayName = "일" Then
                oMat01.Columns(DisableColumnString).BackColor = RGB(255, 0, 0) '빨간색
            ElseIf DayName = "토" Then
                oMat01.Columns(DisableColumnString).BackColor = RGB(0, 128, 255) '하늘색

            Else
                oMat01.Columns(DisableColumnString).BackColor = RGB(255, 255, 255) '흰색
            End If
                            
        Next
        
        
        If LastDay <> 31 Then
        
            DisableColumn = 31 - LastDay
        
            For loopCount = 0 To DisableColumn - 1
            
                DisableColumnString = "D" & CStr(31 - loopCount)
            
                oMat01.Columns.Item(DisableColumnString).Editable = False '해당월의 말일이 존재하지 않으면 막음
            
            Next
        
        End If
        
    Call oForm01.Freeze(False)
    oMat01.AutoResizeColumns
    Exit Sub
PS_PP860_MatrixColumnSetting_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP860_MatrixColumnSetting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
