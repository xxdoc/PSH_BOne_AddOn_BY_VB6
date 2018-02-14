VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP060.cls
'//  Module         : 생산관리 > 작업외 공수 등록
'//  Desc           : 기본
'//  FormType       : PS_PP060
'//  Create Date    : 2010.10.13 노근용
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_PP060H As SAPbouiCOM.DBDataSource '등록헤더
'/Private oDS_PS_PP060L As SAPbouiCOM.DBDataSource '등록라인

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
'/ 추가 확인 갱신 버튼 이름 변경
Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("Btn_save").Specific.Caption = "추가"
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("Btn_save").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn_save").Specific.Caption = "수정"
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'//조회데이타 가져오기
Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim OrdType$, BPLId$, CpGbn$, CpCode$, CntcCode$, WorkGbn$, DocDateFr$, DocDateTo$
    
    OrdType = Trim(oForm01.Items("OrdType").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CpGbn = Trim(oForm01.Items("CpGbn").Specific.VALUE)
    CpCode = Trim(oForm01.Items("CpCode").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    WorkGbn = Trim(oForm01.Items("WorkGbn").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    
    If BPLId = "" Then BPLId = "%"
    If CpGbn = "" Then CpGbn = "%"
    If CpCode = "" Then CpCode = "%"
    If CntcCode = "" Then CntcCode = "%"
    If WorkGbn = "" Then WorkGbn = "%"
    If DocDateFr = "" Then DocDateFr = "19000101"
    If DocDateTo = "" Then DocDateTo = "20991231"
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    sQry = "EXEC [PS_PP060_01] '" & OrdType & "', '" & BPLId & "','" & CpGbn & "','" & CpCode & "','" & CntcCode & "','" & WorkGbn & "','" & DocDateFr & "','" & DocDateTo & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_PP060H.Clear
    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
                
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        oForm01.Mode = fm_ADD_MODE
        
        Call Add_MatrixRow(0, True)
        Call LoadCaption
        Exit Sub
    End If
    
    oForm01.Freeze True
        
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP060H.Size Then
            oDS_PS_PP060H.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_PP060H.Offset = i
        
        Call oDS_PS_PP060H.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP060H.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '문서번호
        Call oDS_PS_PP060H.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("DocNum").VALUE)) '문서번호2
        '선택
        Call oDS_PS_PP060H.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("U_CpGbn").VALUE)) '공정구분
        Call oDS_PS_PP060H.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("U_CpCode").VALUE)) '공정코드
        Call oDS_PS_PP060H.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("U_CpName").VALUE)) '공정명
        Call oDS_PS_PP060H.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("U_CntcCode").VALUE)) '사원코드
        Call oDS_PS_PP060H.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("U_CntcName").VALUE)) '사원명
        Call oDS_PS_PP060H.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("U_FixCode").VALUE)) '자산코드
        Call oDS_PS_PP060H.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("U_FixName").VALUE)) '장비명
        Call oDS_PS_PP060H.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("U_FixCode2").VALUE)) '자산번호(구)
        Call oDS_PS_PP060H.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("U_ItmBsort").VALUE)) '금형번호
        Call oDS_PS_PP060H.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("U_BPLId").VALUE)) '사업장
        Call oDS_PS_PP060H.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")) '등록일자
        Call oDS_PS_PP060H.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("U_NCode").VALUE)) '비가동코드
        Call oDS_PS_PP060H.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("U_WorkNote").VALUE)) '작업내용
        Call oDS_PS_PP060H.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("U_WorkTime").VALUE)) '작업시간
        Call oDS_PS_PP060H.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("U_WorkGbn").VALUE)) '작업구분
        Call oDS_PS_PP060H.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("U_OrdType").VALUE)) '문서구분
        Call oDS_PS_PP060H.setValue("U_ColReg20", i, Trim(oRecordSet01.Fields("ItmRegYN").VALUE)) '자재등록여부
        
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

Public Sub DeleteData()
On Error GoTo DeleteData_Error
    Dim i As Integer
    Dim sQry As String
    Dim ItmRegYN As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, DocNum$, BPLId$, CpGbn$, CpCode$, CntcCode$, WorkGbn$, DocDateFr$, DocDateTo$, OrdType$
    Dim Check As String
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 1
        Check = Trim(oDS_PS_PP060H.GetValue("U_ColReg03", i))
        If Check = "Y" Then
            DocEntry = Trim(oDS_PS_PP060H.GetValue("U_ColReg01", i))
            DocNum = Trim(oDS_PS_PP060H.GetValue("U_ColReg02", i))
            ItmRegYN = Trim(oDS_PS_PP060H.GetValue("U_ColReg20", i))
            
            sQry = "Delete From [@PS_PP060H] where DocEntry = '" & DocEntry & "' and DocNum = '" & DocNum & "'"
            
            If ItmRegYN = "N" Then
                Call oRecordSet01.DoQuery(sQry) '소요자재가 등록되지 않은 행만 삭제
            End If
            
        End If
  
    Next i
     
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    Call Add_MatrixRow(0, True)
    
    Exit Sub

    OrdType = Trim(oForm01.Items("OrdType").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CpGbn = Trim(oForm01.Items("CpGbn").Specific.VALUE)
    CpCode = Trim(oForm01.Items("CpCode").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    WorkGbn = Trim(oForm01.Items("WorkGbn").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    
    
    If BPLId = "" Then BPLId = "%"
    If CpGbn = "" Then CpGbn = "%"
    If CpCode = "" Then CpCode = "%"
    If CntcCode = "" Then CntcCode = "%"
    If WorkGbn = "" Then WorkGbn = "%"
    If DocDateFr = "" Then DocDateFr = "19000101"
    If DocDateTo = "" Then DocDateTo = "20991231"
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    sQry = "EXEC [PS_PP060_01] '" & OrdType & "','" & BPLId & "','" & CpGbn & "','" & CpCode & "','" & CntcCode & "','" & WorkGbn & "','" & DocDateFr & "','" & DocDateTo & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_PP060H.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP060H.Size Then
            oDS_PS_PP060H.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_PP060H.Offset = i
        
        Call oDS_PS_PP060H.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP060H.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '문서번호
        Call oDS_PS_PP060H.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("DocNum").VALUE)) '문서번호2
        '선택
        Call oDS_PS_PP060H.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("U_CpGbn").VALUE)) '공정구분
        Call oDS_PS_PP060H.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("U_CpCode").VALUE)) '공정코드
        Call oDS_PS_PP060H.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("U_CpName").VALUE)) '공정명
        Call oDS_PS_PP060H.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("U_CntcCode").VALUE)) '사원코드
        Call oDS_PS_PP060H.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("U_CntcName").VALUE)) '사원명
        Call oDS_PS_PP060H.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("U_FixCode").VALUE)) '자산코드
        Call oDS_PS_PP060H.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("U_FixName").VALUE)) '장비명
        Call oDS_PS_PP060H.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("U_FixCode2").VALUE)) '자산번호(구)
        Call oDS_PS_PP060H.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("U_ItmBsort").VALUE)) '금형번호
        Call oDS_PS_PP060H.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("U_BPLId").VALUE)) '사업장
        Call oDS_PS_PP060H.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")) '등록일자
        Call oDS_PS_PP060H.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("U_NCode").VALUE)) '비가동코드
        Call oDS_PS_PP060H.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("U_WorkNote").VALUE)) '작업내용
        Call oDS_PS_PP060H.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("U_WorkTime").VALUE)) '작업시간
        Call oDS_PS_PP060H.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("U_WorkGbn").VALUE)) '작업구분
        Call oDS_PS_PP060H.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("U_OrdType").VALUE)) '문서구분
        Call oDS_PS_PP060H.setValue("U_ColReg20", i, Trim(oRecordSet01.Fields("ItmRegYN").VALUE)) '자재등록여부
        
'        'oDS_PS_PP060H.setValue "U_LineNum", i, i + 1
'        oDS_PS_PP060H.setValue "U_BPLId", i, Trim(oRecordSet01.Fields("U_BPLId").VALUE)
'        oDS_PS_PP060H.setValue "U_LineNum", i, Trim(oRecordSet01.Fields("U_LineNum").VALUE)
'        'oDS_PS_PP060H.setValue "DocNum", i, Trim(oRecordSet01.Fields("U_DocNum").Value)
'        oDS_PS_PP060H.setValue "U_CpGbn", i, Trim(oRecordSet01.Fields("U_CpGbn").VALUE)
'        oDS_PS_PP060H.setValue "U_CpCode", i, Trim(oRecordSet01.Fields("U_CpCode").VALUE)
'        oDS_PS_PP060H.setValue "U_CpName", i, Trim(oRecordSet01.Fields("U_CpName").VALUE)
'        oDS_PS_PP060H.setValue "U_CntcCode", i, Trim(oRecordSet01.Fields("U_CntcCode").VALUE)
'        oDS_PS_PP060H.setValue "U_CntcName", i, Trim(oRecordSet01.Fields("U_CntcName").VALUE)
'        oDS_PS_PP060H.setValue "U_ItmBsort", i, Trim(oRecordSet01.Fields("U_ItmBsort").VALUE)
'        oDS_PS_PP060H.setValue "U_DocDate", i, Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")
'        oDS_PS_PP060H.setValue "U_NCode", i, Trim(oRecordSet01.Fields("U_NCode").VALUE)
'        oDS_PS_PP060H.setValue "U_WorkNote", i, Trim(oRecordSet01.Fields("U_WorkNote").VALUE)
'        oDS_PS_PP060H.setValue "U_WorkTime", i, Trim(oRecordSet01.Fields("U_WorkTime").VALUE)
'        oDS_PS_PP060H.setValue "U_WorkGbn", i, Trim(oRecordSet01.Fields("U_WorkGbn").VALUE)
'        oDS_PS_PP060H.setValue "U_OrdType", i, Trim(oRecordSet01.Fields("U_OrdType").VALUE)
        
        
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
DeleteData_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "DeleteData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP060H.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_PP060H.Offset = oRow
    oDS_PS_PP060H.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'// 데이타 UPDATE
Public Function UpdateData(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo UpdateData_Error
    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, DocNum$, BPLId$, LineNum$, WorkTime As Currency
    Dim CpGbn$, CpCode$, CpName$, CntcCode$, CntcName$, ItmBsort$, DocDate$, WorkGbn$, WorkNote$, OrdType$, NCode$
    Dim Check$
    Dim FixCode As String
    Dim FixName As String
    Dim FixCode2 As String
    
    oMat01.FlushToDataSource
    
    OrdType = Trim(oForm01.Items("OrdType").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CpGbn = Trim(oForm01.Items("CpGbn").Specific.VALUE)
    CpCode = Trim(oForm01.Items("CpCode").Specific.VALUE)
    CpName = Trim(oForm01.Items("CpName").Specific.VALUE)
    WorkGbn = Trim(oForm01.Items("WorkGbn").Specific.VALUE)
    
    For i = 1 To oMat01.RowCount
        Check = oMat01.Columns("Check").Cells(i).Specific.Checked
        If Check = True Then
            j = j + 1
        End If
    Next
    
    If j <= 0 Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!", "E"
        Exit Function
    End If
        
    
    For i = 1 To oMat01.RowCount
        Check = oMat01.Columns("Check").Cells(i).Specific.Checked
        If Check = True Then
            j = j + 1
            DocEntry = oMat01.Columns("DocEntry").Cells(i).Specific.VALUE
            
            CntcCode = oMat01.Columns("CntcCode").Cells(i).Specific.VALUE
            CntcName = oMat01.Columns("CntcName").Cells(i).Specific.VALUE
            
            FixCode = oMat01.Columns("FixCode").Cells(i).Specific.VALUE
            FixName = oMat01.Columns("FixName").Cells(i).Specific.VALUE
            FixCode2 = oMat01.Columns("FixCode2").Cells(i).Specific.VALUE
            
            ItmBsort = oMat01.Columns("ItmBsort").Cells(i).Specific.VALUE
            NCode = oMat01.Columns("NCode").Cells(i).Specific.VALUE
            WorkNote = oMat01.Columns("WorkNote").Cells(i).Specific.VALUE
            WorkTime = oMat01.Columns("WorkTime").Cells(i).Specific.VALUE
            DocDate = oMat01.Columns("DocDate").Cells(i).Specific.VALUE
        
            sQry = "Update [@PS_PP060H]"
            sQry = sQry & " set "
            sQry = sQry & " U_CntcCode = '" & CntcCode & "',"
            sQry = sQry & " U_CntcName = '" & CntcName & "',"
            sQry = sQry & " U_FixCode = '" & FixCode & "',"
            sQry = sQry & " U_FixName = '" & FixName & "',"
            sQry = sQry & " U_FixCode2 = '" & FixCode2 & "',"
            sQry = sQry & " U_ItmBsort = '" & ItmBsort & "',"
            sQry = sQry & " U_DocDate  = '" & DocDate & "',"
            sQry = sQry & " U_CpCode  = '" & CpCode & "',"
            sQry = sQry & " U_CpName  = '" & CpName & "',"
            sQry = sQry & " U_NCode = '" & NCode & "',"
            sQry = sQry & " U_WorkNote = '" & WorkNote & "',"
            sQry = sQry & " U_WorkTime = '" & WorkTime & "'"
            
            sQry = sQry & " Where DocEntry = '" & DocEntry & "'"
            RecordSet01.DoQuery sQry
        End If
    Next
    
    
            
    MDC_Com.MDC_GF_Message "작업외 공수수정 완료!", "S"
    
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    Set RecordSet01 = Nothing
   ' Add_PurchaseDemand = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Updatedata_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

'// 데이타 INSERT
Public Function Add_PurchaseDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_PurchaseDemand_Error
    Dim i As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, DocNum$, OrdType$, BPLId$, LineNum$, WorkTime As Currency
    Dim CpGbn$, CpCode$, CpName$, CntcCode$, CntcName$, ItmBsort$, DocDate$, WorkGbn$, WorkNote$, NCode$
    Dim FixCode As String
    Dim FixName As String
    Dim FixCode2 As String
        
    oMat01.FlushToDataSource
    
    OrdType = Trim(oForm01.Items("OrdType").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CpGbn = Trim(oForm01.Items("CpGbn").Specific.VALUE)
    CpCode = Trim(oForm01.Items("CpCode").Specific.VALUE)
    CpName = Trim(oForm01.Items("CpName").Specific.VALUE)
    WorkGbn = Trim(oForm01.Items("WorkGbn").Specific.VALUE)
    
    If OrdType = "20" Then
        '//비가동입력
        CpGbn = ""
        CpCode = ""
        CpName = ""
    End If
    
    For i = 0 To oMat01.RowCount - 2
        DocDate = Trim(oDS_PS_PP060H.GetValue("U_ColDt01", i))
    
        sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_PP060H] where Left(DocEntry, 6) = Left('" & DocDate & "', 6)"
        RecordSet01.DoQuery sQry
        If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
            DocEntry = Left(DocDate, 6) + "0001"
        Else
            DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
        End If
        
        DocNum = i + 1
        LineNum = i + 1
        
        CntcCode = Trim(oDS_PS_PP060H.GetValue("U_ColReg07", i))
        CntcName = Trim(oDS_PS_PP060H.GetValue("U_ColReg08", i))
        FixCode = Trim(oDS_PS_PP060H.GetValue("U_ColReg09", i))
        FixName = Trim(oDS_PS_PP060H.GetValue("U_ColReg10", i))
        FixCode2 = Trim(oDS_PS_PP060H.GetValue("U_ColReg11", i))
        ItmBsort = Trim(oDS_PS_PP060H.GetValue("U_ColReg12", i))
        NCode = Trim(oDS_PS_PP060H.GetValue("U_ColReg15", i))
        WorkNote = Trim(oDS_PS_PP060H.GetValue("U_ColReg16", i))
        WorkTime = Trim(oDS_PS_PP060H.GetValue("U_ColQty01", i))
        
        sQry = "INSERT INTO [@PS_PP060H]"
        sQry = sQry & " ("
        sQry = sQry & " DocEntry,"
        sQry = sQry & " DocNum,"
        sQry = sQry & " U_BPLId,"
        sQry = sQry & " U_LineNum,"
        sQry = sQry & " U_CpGbn,"
        sQry = sQry & " U_CpCode,"
        sQry = sQry & " U_CpName,"
        sQry = sQry & " U_CntcCode,"
        sQry = sQry & " U_CntcName,"
        sQry = sQry & " U_FixCode,"
        sQry = sQry & " U_FixName,"
        sQry = sQry & " U_FixCode2,"
        sQry = sQry & " U_ItmBsort,"
        sQry = sQry & " U_DocDate,"
        sQry = sQry & " U_WorkNote,"
        sQry = sQry & " U_WorkTime,"
        sQry = sQry & " U_WorkGbn,"
        sQry = sQry & " U_OrdType,"
        sQry = sQry & " U_NCode,"
        sQry = sQry & " U_ItmRegYN"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("
        sQry = sQry & DocEntry & ","
        sQry = sQry & DocNum & ","
        sQry = sQry & "'" & BPLId & "',"
        sQry = sQry & "'" & LineNum & "',"
        sQry = sQry & "'" & CpGbn & "',"
        sQry = sQry & "'" & CpCode & "',"
        sQry = sQry & "'" & CpName & "',"
        sQry = sQry & "'" & CntcCode & "',"
        sQry = sQry & "'" & CntcName & "',"
        sQry = sQry & "'" & FixCode & "',"
        sQry = sQry & "'" & FixName & "',"
        sQry = sQry & "'" & FixCode2 & "',"
        sQry = sQry & "'" & ItmBsort & "',"
        sQry = sQry & "'" & DocDate & "',"
        sQry = sQry & "'" & WorkNote & "',"
        sQry = sQry & "'" & WorkTime & "',"
        sQry = sQry & "'" & WorkGbn & "',"
        sQry = sQry & "'" & OrdType & "',"
        sQry = sQry & "'" & NCode & "',"
        sQry = sQry & "'N'"
        sQry = sQry & ")"
        RecordSet01.DoQuery sQry
    Next
    
    MDC_Com.MDC_GF_Message "작업외 공수등록 완료!", "S"
    
    Set RecordSet01 = Nothing
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
        Case Trim(oForm01.Items("OrdType").Specific.VALUE) = "10" And Trim(oForm01.Items("CpGbn").Specific.VALUE) = "%"
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("BPLId").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("OrdType").Specific.VALUE) = "10" And Trim(oForm01.Items("CpGbn").Specific.VALUE) <> "60" And Trim(oForm01.Items("CpGbn").Specific.VALUE) <> "70" And Trim(oForm01.Items("CpGbn").Specific.VALUE) <> "80" And Trim(oForm01.Items("CpCode").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("WorkGbn").Specific.VALUE) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("OrdType").Specific.VALUE) = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "공정구분은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "공정코드는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "작업구분은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "문서구분은 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
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

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '/ 헤드사항을 입력
    Dim OrdType$, BPLId$, CpGbn$, CpCode$, CpName$, WorkGbn$
    
    OrdType = Trim(oForm01.Items("OrdType").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CpGbn = Trim(oForm01.Items("CpGbn").Specific.VALUE)
    CpCode = Trim(oForm01.Items("CpCode").Specific.VALUE)
    CpName = Trim(oForm01.Items("CpName").Specific.VALUE)
    WorkGbn = Trim(oForm01.Items("WorkGbn").Specific.VALUE)
    
    '/메트릭스 읽을때 선언
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 2
        If Trim(oDS_PS_PP060H.GetValue("U_ColReg07", i)) = "" Then
            oDS_PS_PP060H.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource

    ErrNum = 0
    
    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case oDS_PS_PP060H.GetValue("U_ColReg07", i) = ""
                ErrNum = 2
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_PP060H.GetValue("U_ColQty01", i) = ""
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_PP060H.GetValue("U_ColDt01", i) = ""
                ErrNum = 4
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_PP060H.GetValue("U_ColReg15", i) = ""
                If OrdType = "20" Then
                    ErrNum = 5
                    GoTo MatrixSpaceLineDel_Error
                End If
        End Select
    Next
    
    '/메트릭스의 값변경후 선언
    oMat01.LoadFromDataSource

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

'/코드의 코드명 보여주기
Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim ItemCode$, Qty&
    Dim Calculate_Weight As Double
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CntcCode" '/사번
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01" '/메트릭스
            oForm01.Freeze True
            If oCol = "CntcCode" Then '/사번코드
                
'                If oForm01.Mode = fm_ADD_MODE Then
'                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("CntcCode").Cells(oRow).Specific.Value) <> "" Then

                        sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oMat01.Columns("CntcCode").Cells(oRow).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry

                        oMat01.Columns("CntcName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        oMat01.Columns("DocDate").Cells(oRow).Specific.VALUE = Format(Date - 1, "YYYYMMDD")

                        oMat01.FlushToDataSource
'                        Add_MatrixRow oMat01.RowCount, False
                        If oMat01.RowCount = oRow And Trim(oDS_PS_PP060H.GetValue("U_ColReg07", oRow - 1)) <> "" Then
                            Call Add_MatrixRow(oRow, False) 'True)
                        End If
                        
'                    End If
'                End If
                
            ElseIf oCol = "FixCode" Then
            
                sQry = "           SELECT       U_FixName AS [FixName],"
                sQry = sQry & "                 U_TempChr1 As FixCode2"
                sQry = sQry & " FROM        [@PS_FX005H]"
                sQry = sQry & " WHERE       U_FixCode + '-' + U_SubCode = '" & Trim(oMat01.Columns("FixCode").Cells(oRow).Specific.VALUE) & "'"
                Call oRecordSet01.DoQuery(sQry)
                
                oMat01.Columns("FixName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("FixName").VALUE)
                oMat01.Columns("FixCode2").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("FixCode2").VALUE)
                
                Call oMat01.FlushToDataSource
                
            End If
            
            Call oMat01.AutoResizeColumns
            oForm01.Freeze False
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/폼의 아이템 사용지정
Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
        oForm01.Items("OrdType").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CpGbn").Enabled = True
        oForm01.Items("CpCode").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("WorkGbn").Enabled = True
        oForm01.Items("DocDateFr").Enabled = True
        oForm01.Items("DocDateTo").Enabled = True
        oMat01.Columns("BPLId").Editable = False
        oMat01.Columns("CpGbn").Editable = False
        oMat01.Columns("CpCode").Editable = False
        oMat01.Columns("CpName").Editable = False
        oMat01.Columns("CntcCode").Editable = True
        oMat01.Columns("CntcName").Editable = False
        oMat01.Columns("ItmBsort").Editable = True
        oMat01.Columns("DocDate").Editable = True
        oMat01.Columns("WorkNote").Editable = True
        oMat01.Columns("WorkTime").Editable = True
        oMat01.Columns("WorkGbn").Editable = False
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("OrdType").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CpGbn").Enabled = True
        oForm01.Items("CpCode").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("WorkGbn").Enabled = True
        oForm01.Items("DocDateFr").Enabled = True
        oForm01.Items("DocDateTo").Enabled = True
        oMat01.Columns("BPLId").Editable = False
        oMat01.Columns("CpGbn").Editable = False
        oMat01.Columns("CpCode").Editable = False
        oMat01.Columns("CpName").Editable = False
        oMat01.Columns("CntcCode").Editable = False
        oMat01.Columns("CntcName").Editable = False
        oMat01.Columns("ItmBsort").Editable = False
        oMat01.Columns("DocDate").Editable = False
        oMat01.Columns("WorkNote").Editable = False
        oMat01.Columns("WorkTime").Editable = False
        oMat01.Columns("WorkGbn").Editable = False
        oForm01.Items("Btn_ret").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP060.srf")
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
    
    oFormUniqueID01 = "PS_PP060_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP060_CreateItems
    Call PS_PP060_ComboBox_Setting
    Call PS_PP060_CF_ChooseFromList
    Call PS_PP060_EnableMenus
    Call PS_PP060_SetDocument(oFromDocEntry01)
    Call PS_PP060_FormResize
    
    
    Call Add_MatrixRow(0, True)
    Call LoadCaption
    Call FormItemEnabled
    
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True
        
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
    '//일자 Set
     oForm01.Items("DocDateFr").Specific.VALUE = Format(Date - 1, "YYYYMMDD")
     oForm01.Items("DocDateTo").Specific.VALUE = Format(Date, "YYYYMMDD")
'     oMat01.Columns("LineNum").Cells(1).Specific.Value = 1
     
     Call oForm01.Items("CntcCode").Click
     
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub
'/아이템 변경 이벤트
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error


    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            Call Raise_EVENT_ITEM_PRESSED(FormUID, pval, BubbleEvent)
        Case et_KEY_DOWN: '//2
            If pval.CharPressed = 9 Then
                If pval.ItemUID = "CntcCode" Then
                    If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                 ElseIf pval.ItemUID = "Mat01" Then '/메트릭스
                    If pval.ColUID = "CpCode" Then
                        If oMat01.Columns("CpCode").Cells(pval.Row).Specific.VALUE = "" Then
                           Sbo_Application.ActivateMenuItem ("7425")
                           BubbleEvent = False
                        End If
'                        Call Add_MatrixRow(0, True)
                
                    ElseIf pval.ColUID = "CntcCode" Then '/사번
                        If oMat01.Columns("CntcCode").Cells(pval.Row).Specific.VALUE = "" Then
                           Sbo_Application.ActivateMenuItem ("7425")
                           BubbleEvent = False
                        End If
                    ElseIf pval.ColUID = "FixCode" Then '자산번호
                        If oMat01.Columns("FixCode").Cells(pval.Row).Specific.VALUE = "" Then
                           Sbo_Application.ActivateMenuItem ("7425")
                           BubbleEvent = False
                        End If
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
                    If pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "CntcCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "WorkTime" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "FixCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "WorkNote" Then
                            Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                        End If
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                        Else
                            oForm01.Mode = fm_UPDATE_MODE
                            ' Call LoadCaption
                        End If
                    End If
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
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                Call Add_MatrixRow(0, True)
                BubbleEvent = False
                Call LoadCaption
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_PP060_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_PP060H.Clear

'                Call LoadCaption
'                Call FormItemEnabled
                '//Call PS_PP060_FormItemEnabled '//UDO방식
                '//Call PS_PP060_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_PP060_FormItemEnabled
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
        If pval.ItemUID = "PS_PP060" Then
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
                If MatrixSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                If Add_PurchaseDemand(pval) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Add_MatrixRow 0, True

'                        Call Delete_EmptyRow
                oLast_Mode = oForm01.Mode
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If UpdateData(pval) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                Call LoadData
            End If
            
        ElseIf pval.ItemUID = "Btn_ret" Then '/조회
            If HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            End If
            
            oForm01.Mode = fm_UPDATE_MODE
            
            Call LoadCaption
            Call LoadData
        ElseIf pval.ItemUID = "Btn_del" Then '/삭제
'            If DeleteBefor = False Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            oForm01.Mode = fm_OK_MODE '/fm_VIEW_MODE
            
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?                              ※소요자재가 등록된 행은 삭제가 불가능합니다. 소요자재 정보를 먼저 삭제하세요.", "1", "예", "아니오") = "1" Then
            
                Call DeleteData
                Call LoadData
                
            Else
            
            End If
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP060" Then
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "FixCode") '자산번호 포맷서치
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
                Call oMat01.SelectRow(pval.Row, True, False)
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
Dim SCpCode As String

Dim oRecordSet01      As SAPbobsCOM.Recordset

Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
            If pval.ItemUID = "CpGbn" Then
                

                sCount = oForm01.Items("CpCode").Specific.ValidValues.Count
                sSeq = sCount
                For i = 1 To sCount
                    oForm01.Items("CpCode").Specific.ValidValues.Remove sSeq - 1, psk_Index
                    sSeq = sSeq - 1
                Next i

                '//공정구분에 따른 공정코드변경

                Select Case oForm01.Items("CpGbn").Specific.VALUE
                    Case "10"
                        '//멀티게이지-금형래핑
                        Call oForm01.Items("WorkGbn").Specific.Select("104", psk_ByValue)
                        oForm01.Items("ItmBsort").Enabled = True
                        oMat01.Columns("ItmBsort").Editable = True
                    Case "20"
                        '//멀티게이지-외경연삭
                        Call oForm01.Items("WorkGbn").Specific.Select("104", psk_ByValue)
                        oForm01.Items("ItmBsort").Enabled = False
                        oMat01.Columns("ItmBsort").Editable = False
                    Case "30"
                        '//멀티게이지-포장
                        Call oForm01.Items("WorkGbn").Specific.Select("104", psk_ByValue)
                        oForm01.Items("ItmBsort").Enabled = False
                        oMat01.Columns("ItmBsort").Editable = False
                    Case "40"
                        '//휘팅-바렐
                        Call oForm01.Items("WorkGbn").Specific.Select("101", psk_ByValue)
                        oForm01.Items("ItmBsort").Enabled = False
                        oMat01.Columns("ItmBsort").Editable = False
                    Case "50"
                        '//휘팅-포장
                        Call oForm01.Items("WorkGbn").Specific.Select("101", psk_ByValue)
                        oForm01.Items("ItmBsort").Enabled = False
                        oMat01.Columns("ItmBsort").Editable = False
                    Case "60", "70", "80"
                        '//검사공수 입력
                        Call oForm01.Items("WorkGbn").Specific.Select("105", psk_ByValue)
                    Case Else

                End Select


 
            ElseIf pval.ItemUID = "WorkGbn" Then
                If oForm01.Items("OrdType").Specific.VALUE = "20" Then
                    '//비가동
                    
                Else
                    '//실동
                    sCount = oForm01.Items("CpCode").Specific.ValidValues.Count
                    sSeq = sCount
                    For i = 1 To sCount
                        oForm01.Items("CpCode").Specific.ValidValues.Remove sSeq - 1, psk_Index
                        sSeq = sSeq - 1
                    Next i
                    Select Case oForm01.Items("WorkGbn").Specific.VALUE
                        Case "101"
                            sCode = "CP301"
                            SCpCode = "%"
                        Case "102"
                            sCode = "CP401"
                            SCpCode = "%"
                        Case "104"
                            sCode = "CP501"
                            SCpCode = "%"
                        Case "105", "106"
                            sCode = "%"
                            SCpCode = "%"
                            
                        Case "107"
                            sCode = "CP101"
                            SCpCode = "%"
                            
                    End Select
                    
                    sQry = "SELECT U_CpCode, U_CpName From [@PS_PP001L] Where Code like '" & sCode & "' and U_CpCode like '" & SCpCode & "' Order by Code"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("CpCode").Specific.ValidValues.Add "", ""
                    Do Until oRecordSet01.EOF
                        
                        oForm01.Items("CpCode").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                        oRecordSet01.MoveNext
                    Loop
                    Select Case oForm01.Items("CpGbn").Specific.VALUE
                            Case "10"
                                '//v-mill
                                SCpCode = "CP50101"
                            Case "20"
                                '// 콤보박스에 값을 지정 멀티게이지-외경연삭일때는 FRM공정을 SD380
                                SCpCode = "CP50104"
                            Case "30"
                                '// 콤보박스에 값을 지정 멀티게이지-포장일때는 FRM공정을 SD380
                                SCpCode = "CP50107"
                            Case "40" '//휘팅바렐
                                SCpCode = "CP30112"
                            Case "50" '//휘팅포장
                                SCpCode = "CP30114"
                            Case "60", "70", "80"
                                oForm01.Items("CpCode").Specific.Select 0, psk_Index
                            Case Else
                               '// 콤보박스에 첫데이타를 SD380로
                               oForm01.Items("CpCode").Specific.Select 1, psk_Index
                               SCpCode = oForm01.Items("CpCode").Specific.VALUE
        
                               sQry = "SELECT U_CpName From [@PS_PP001L] Where U_CpCode = '" & SCpCode & "' Order by Code"
                                oRecordSet01.DoQuery sQry
                                oForm01.Items("CpName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    
                        End Select
                        
                        If oForm01.Items("CpGbn").Specific.VALUE <> "60" And oForm01.Items("CpGbn").Specific.VALUE <> "70" And oForm01.Items("CpGbn").Specific.VALUE <> "80" Then
                        
                            Call oForm01.Items("CpCode").Specific.Select(SCpCode, psk_ByValue)
                            sQry = "SELECT U_CpName From [@PS_PP001L] Where U_CpCode = '" & SCpCode & "' Order by Code"
                            oRecordSet01.DoQuery sQry
                            oForm01.Items("CpName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                        End If
                End If
            ElseIf pval.ItemUID = "OrdType" Then
            
                If oForm01.Items("OrdType").Specific.VALUE = "10" Then
                    '// 실동입력
                    oForm01.Items("CpGbn").Enabled = True
'                    oForm01.Items("CntcCode").Enabled = True '사번 TextBox 활성, 비활성 기능 제외(2013.05.03 송명규 추가)
                    oForm01.Items("CpCode").Enabled = True
                    oForm01.Items("ItmBsort").Enabled = True
                    oMat01.Columns("NCode").Editable = False
                Else
                    '// 비가동입력
                    oForm01.Items("CpGbn").Enabled = False
                    
                    Call oForm01.Items("CpGbn").Specific.Select("", psk_ByValue)
'                    oForm01.Items("CntcCode").Enabled = False '사번 TextBox 활성, 비활성 기능 제외(2013.05.03 송명규 추가)
                    
                    oForm01.Items("CpCode").Enabled = False
                    Call oForm01.Items("CpCode").Specific.Select("", psk_ByValue)
                    oForm01.Items("ItmBsort").Enabled = False
                    oMat01.Columns("NCode").Editable = True
                    
                    
                End If
            ElseIf pval.ItemUID = "CpCode" Then '// 공정코드
                sQry = "SELECT U_CpName From [@PS_PP001L] Where U_CpCode = '" & Trim(oForm01.Items("CpCode").Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oForm01.Items("CpName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If



    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
            If pval.Row = -1 Then
'                oGrid01.Columns(pval.ColUID).TitleObject.Sortable = True
                
            Else
'                If oGrid01.Rows.SelectedRows.Count > 0 Then
                
                    Call PS_PP060_OpenItemRegist(pval.Row)

'                Else
'                    BubbleEvent = False
'                End If
            End If
        End If
    
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

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        
    End If
    
    Call oForm01.Freeze(False)
    
    Exit Sub
Raise_EVENT_VALIDATE_Error:

    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_PP060_FormItemEnabled
        '//Call PS_PP060_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP060_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP060H", "U_CardCode,U_CardName")
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
'            If (PS_PP060_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_PP060H.RemoveRecord(oDS_PS_PP060H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_PP060H.GetValue("U_ColReg07", oMat01.RowCount - 1)) <> "" Then
                    Call Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP060_CreateItems() As Boolean
On Error GoTo PS_PP060_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP060H = oForm01.DataSources.DBDataSources("@PS_USERDS01")

    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    oMat01.Columns("DocNum").Visible = False 'DocNum Hidden 처리(2015.01.27 송명규)
    
    '/유저데이타 속성 선언 날짜형식
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 8)
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 8)
    Call oForm01.Items("DocDateFr").Specific.DataBind.SetBound(True, "", "DocDateFr")
    Call oForm01.Items("DocDateTo").Specific.DataBind.SetBound(True, "", "DocDateTo")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP060_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP060_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP060_ComboBox_Setting()
On Error GoTo PS_PP060_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP060", "CpGbn", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP060", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP060", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP060", "CpGbn", "", "10", "멀티게이지-금형래핑")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP060", "CpGbn", "", "20", "멀티게이지-외경연삭")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("CpGbn").Specific, "10", "멀티게이지")
        
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Set oCombo = oForm01.Items("BPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    
    oForm01.Items("WorkGbn").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("WorkGbn").Specific, "select Code, Name from [@PSH_ITMBSORT] Where U_PudYN = 'Y' order by Code", "101", False, False)
    
    
    'Call MDC_SetMod.Set_ComboList(oMat01.Columns("NCode").Cells( Items("WorkGbn").Specific, "select Code, Name from [@PSH_ITMBSORT] Where U_PudYN = 'Y' order by Code", "101", False, False)
    
    Set oCombo = oForm01.Items("OrdType").Specific
    oCombo.ValidValues.Add "10", "실동"
    oCombo.ValidValues.Add "20", "비가동"
    oCombo.Select "0", psk_Index
    
    '//공정구분 입력(시스템코드로 수정, 2014.11.18 송명규)
    sQry = "           SELECT      U_Minor,"
    sQry = sQry & "                U_CdName"
    sQry = sQry & " FROM       [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'P206'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    
    Call oForm01.Items("CpGbn").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CpGbn").Specific, sQry, "", False, False)
    Call oForm01.Items("CpGbn").Specific.Select(0, psk_Index)
    
'    Set oCombo = oForm01.Items("CpGbn").Specific
'    oCombo.ValidValues.Add "", ""
'    oCombo.ValidValues.Add "10", "멀티게이지-금형래핑"
'    oCombo.ValidValues.Add "20", "멀티게이지-외경연삭"
'    oCombo.ValidValues.Add "30", "멀티게이지-포장"
'    oCombo.ValidValues.Add "40", "휘팅-바렐"
'    oCombo.ValidValues.Add "50", "휘팅-포장"
'    oCombo.ValidValues.Add "60", "품질검사공수"
'    oCombo.ValidValues.Add "70", "정비공수"
    
    '공정구분(매트릭스) 2013.05.01 송명규 추가
'    Call oMat01.Columns("CpGbn").ValidValues.Add("10", "멀티게이지-금형래핑")
'    Call oMat01.Columns("CpGbn").ValidValues.Add("20", "멀티게이지-외경연삭")
'    Call oMat01.Columns("CpGbn").ValidValues.Add("30", "멀티게이지-포장")
'    Call oMat01.Columns("CpGbn").ValidValues.Add("40", "휘팅-바렐")
'    Call oMat01.Columns("CpGbn").ValidValues.Add("50", "휘팅-포장")
'    Call oMat01.Columns("CpGbn").ValidValues.Add("60", "품질검사공수")
'    Call oMat01.Columns("CpGbn").ValidValues.Add("70", "정비공수")
    sQry = "           SELECT      U_Minor,"
    sQry = sQry & "                U_CdName"
    sQry = sQry & " FROM       [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'P206'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CpGbn"), sQry)
    
    '작업구분(매트릭스) 2013.05.01 송명규 추가
    sQry = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' ORDER BY Code"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("WorkGbn"), sQry)
    
    '문서구분(매트릭스) 2013.05.01 송명규 추가
    Call oMat01.Columns("OrdType").ValidValues.Add("10", "실동")
    Call oMat01.Columns("OrdType").ValidValues.Add("20", "비가동")
    
    '사업장(매트릭스) 2013.05.01 송명규 추가
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'P005' Order By U_Minor "
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("NCode").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("CpCode").Specific, "select U_CpCode, U_CpName from [@PS_PP001L] Where Code = 'CP301' ", "CP30112", False, False)
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_PP060_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP060_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP060_CF_ChooseFromList()
On Error GoTo PS_PP060_CF_ChooseFromList_Error

    Exit Sub
PS_PP060_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP060_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP060_FormItemEnabled()
On Error GoTo PS_PP060_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_PP060_FormClear '//UDO방식
'        Call oForm01.EnableMenu("1281", True) '//찾기
'        Call oForm01.EnableMenu("1282", False) '//추가
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
'        Call oForm01.EnableMenu("1281", False) '//찾기
'        Call oForm01.EnableMenu("1282", True) '//추가
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP060_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP060_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP060_EnableMenus()
On Error GoTo PS_PP060_EnableMenus_Error
    
    Exit Sub
PS_PP060_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP060_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP060_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP060_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP060_FormItemEnabled
        '//Call PS_PP060_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP060_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP060_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP060_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP060_FormResize()
On Error GoTo PS_PP060_FormResize_Error
    
    Exit Sub
PS_PP060_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP060_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_PP060_OpenItemRegist(ByVal pRow As Integer)
On Error GoTo PS_PP060_OpenItemRegist_Error
    
    Dim oTempClass As Variant
    
    Dim StdNo As String '기준문서번호
    Dim BPLId As String '사업장
    Dim CpCode As String '공정코드
    Dim CpName As String '공정명
    Dim WkCode As String '작업자사번
    Dim WkName As String '작업자명
    Dim FixCode As String '자산코드
    Dim FixName As String '자산명
    Dim DocDate As String '등록일자
    Dim WorkTime As String '작업시간
    Dim WorkNote As String '작업내용
    
    StdNo = Trim(oMat01.Columns("DocEntry").Cells(pRow).Specific.VALUE)
    BPLId = Trim(oMat01.Columns("BPLId").Cells(pRow).Specific.VALUE)
    CpCode = Trim(oMat01.Columns("CpCode").Cells(pRow).Specific.VALUE)
    CpName = Trim(oMat01.Columns("CpName").Cells(pRow).Specific.VALUE)
    WkCode = Trim(oMat01.Columns("CntcCode").Cells(pRow).Specific.VALUE)
    WkName = Trim(oMat01.Columns("CntcName").Cells(pRow).Specific.VALUE)
    FixCode = Trim(oMat01.Columns("FixCode").Cells(pRow).Specific.VALUE)
    FixName = Trim(oMat01.Columns("FixName").Cells(pRow).Specific.VALUE)
    DocDate = Trim(oMat01.Columns("DocDate").Cells(pRow).Specific.VALUE)
    WorkTime = Trim(oMat01.Columns("WorkTime").Cells(pRow).Specific.VALUE)
    WorkNote = Trim(oMat01.Columns("WorkNote").Cells(pRow).Specific.VALUE)

    Set oTempClass = New PS_PP061
    Call oTempClass.LoadForm(StdNo, BPLId, CpCode, CpName, WkCode, WkName, FixCode, FixName, DocDate, WorkTime, WorkNote)

'    Call Sbo_Application.MessageBox(DocEntry)
    
    Exit Sub

PS_PP060_OpenItemRegist_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP060_OpenItemRegist_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

