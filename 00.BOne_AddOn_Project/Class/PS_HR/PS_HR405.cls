VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_HR405.cls
'//  Module         : HR
'//  Desc           : 전문직 인사평가 참고사항
'//  FormType       : PS_HR405
'//  Create Date    : 2012.09.11 N.G.Y
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_HR405H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_HR405L As SAPbouiCOM.DBDataSource '등록라인

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

Private gYear As String
Private gMSTCOD As String
Private gFULLNAME As String
Private gEmpNo1 As String
Private gEmpName1 As String
Private gPassWd As String

Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long
'/ 추가 확인 갱신 버튼 이름 변경
Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("Btn_save").Specific.Caption = "추가"
        oForm01.Items("Btn_del").Enabled = False
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("Btn_save").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn_save").Specific.Caption = "수정"
        oForm01.Items("Btn_del").Enabled = True
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
    
    Dim SBPLID$, SYear$, SMSTCOD$, SEmpNo1$
    
    SBPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    SYear = Trim(oForm01.Items("Year").Specific.VALUE)
    SMSTCOD = Trim(oForm01.Items("MSTCOD").Specific.VALUE)
    SEmpNo1 = Trim(oForm01.Items("EmpNo1").Specific.VALUE)
    
    If SEmpNo1 = "" Then SEmpNo1 = "%"
         
    sQry = "EXEC [PS_HR405_01] '" & SBPLID & "','" & SYear & "', '" & SMSTCOD & "', '" & SEmpNo1 & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_HR405L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
                
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        oForm01.Mode = fm_ADD_MODE
        
        Call Add_MatrixRow(0, True)
        Call LoadCaption
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_HR405L.Size Then
            oDS_PS_HR405L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_HR405L.Offset = i
        
        oDS_PS_HR405L.setValue "U_LineNum", i, i + 1
        oDS_PS_HR405L.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields(0).VALUE) '//BPLId
        oDS_PS_HR405L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields(1).VALUE) '//DocEntry
        oDS_PS_HR405L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields(2).VALUE) '//Year
        oDS_PS_HR405L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields(3).VALUE) '//FULLNAME
        oDS_PS_HR405L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields(4).VALUE) '//EmpName1
        oDS_PS_HR405L.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields(5).VALUE) '//RateCode
        oDS_PS_HR405L.setValue "U_ColReg07", i, Trim(oRecordSet01.Fields(6).VALUE) '//RateName
        oDS_PS_HR405L.setValue "U_ColReg08", i, Trim(oRecordSet01.Fields(7).VALUE) '//DocDate
        oDS_PS_HR405L.setValue "U_ColReg09", i, Trim(oRecordSet01.Fields(8).VALUE) '//Contents
        oDS_PS_HR405L.setValue "U_ColSum01", i, Trim(oRecordSet01.Fields(9).VALUE) '//EvalScr
        
        
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
    
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
Public Sub DeleteData()
On Error GoTo DeleteData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$
    Dim Check As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
        gYear = Trim(oForm01.Items("Year").Specific.VALUE)
        gMSTCOD = Trim(oForm01.Items("MSTCOD").Specific.VALUE)
        gFULLNAME = Trim(oForm01.Items("FULLNAME").Specific.VALUE)
        gEmpNo1 = Trim(oForm01.Items("EmpNo1").Specific.VALUE)
        gEmpName1 = Trim(oForm01.Items("EmpName1").Specific.VALUE)
        gPassWd = Trim(oForm01.Items("PassWd").Specific.VALUE)
    
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        
        sQry = "Select Count(*) From [@PS_HR405H] where DocEntry = '" & DocEntry & "'"
        oRecordSet01.DoQuery sQry
        
        If (oRecordSet01.RecordCount = 0) Then
            MDC_Com.MDC_GF_Message "삭제대상이 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
            Set oRecordSet01 = Nothing
            oForm01.Mode = fm_ADD_MODE
        Else
            sQry = "Delete From [@PS_HR405H] where DocEntry = '" & DocEntry & "'"
            oRecordSet01.DoQuery sQry
        End If
    End If
    
    Call FormReset
    
    oForm01.Mode = fm_ADD_MODE
    
    oForm01.Items("Btn_ret").Click ct_Regular
    

'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
DeleteData_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "DeleteData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'/메트릭스 Row추가
Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_HR405L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_HR405L.Offset = oRow
    oDS_PS_HR405L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'// 데이타 Update
Public Function UpdateData(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo UpdateData_Error
    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, BPLID$, Year$, MSTCOD$, FULLNAME$, EmpNo1$, EmpName1$, RateCode$, RateMNm$, DocDate$
    Dim Contents$, EvalScr$
    
    
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    Year = Trim(oForm01.Items("Year").Specific.VALUE)
    MSTCOD = Trim(oForm01.Items("MSTCOD").Specific.VALUE)
    FULLNAME = Trim(oForm01.Items("FULLNAME").Specific.VALUE)
    EmpNo1 = Trim(oForm01.Items("EmpNo1").Specific.VALUE)
    EmpName1 = Trim(oForm01.Items("EmpName1").Specific.VALUE)
    RateCode = Trim(oForm01.Items("RateCode").Specific.VALUE)
    RateMNm = Trim(oForm01.Items("RateMNm").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    Contents = Trim(oForm01.Items("Contents").Specific.VALUE)
    EvalScr = Trim(oForm01.Items("EvalScr").Specific.VALUE)
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!", "E"
        UpdateData = False
        Exit Function
    End If
        
    
    sQry = "Update [@PS_HR405H]"
    sQry = sQry & " set "
    sQry = sQry & " U_BPLId = '" & BPLID & "',"
    sQry = sQry & " U_Year = '" & Year & "',"
    sQry = sQry & " U_MSTCOD = '" & MSTCOD & "',"
    sQry = sQry & " U_FULLNAME = '" & FULLNAME & "',"
    sQry = sQry & " U_Empno1 = '" & EmpNo1 & "',"
    sQry = sQry & " U_EmpName1 = '" & EmpName1 & "',"
    sQry = sQry & " U_RateCode  = '" & RateCode & "',"
    sQry = sQry & " U_RateMNm  = '" & RateMNm & "',"
    sQry = sQry & " U_DocDate  = '" & DocDate & "',"
    sQry = sQry & " U_Contents = '" & Contents & "',"
    sQry = sQry & " U_EvalScr = '" & EvalScr & "'"
            
    sQry = sQry & " Where DocEntry = '" & DocEntry & "'"
    RecordSet01.DoQuery sQry
    
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    gYear = Trim(oForm01.Items("Year").Specific.VALUE)
    gMSTCOD = Trim(oForm01.Items("MSTCOD").Specific.VALUE)
    gFULLNAME = Trim(oForm01.Items("FULLNAME").Specific.VALUE)
    gEmpNo1 = Trim(oForm01.Items("EmpNo1").Specific.VALUE)
    gEmpName1 = Trim(oForm01.Items("EmpName1").Specific.VALUE)
    gPassWd = Trim(oForm01.Items("PassWd").Specific.VALUE)
    
    Set RecordSet01 = Nothing
    UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Updatedata_Error:" & Err.Number & " - " & Err.Description, "E"
End Function
'// 데이타 insert
Public Function Add_PurchaseDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_PurchaseDemand_Error
    Dim i As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
   Dim DocEntry$, BPLID$, Year$, MSTCOD$, FULLNAME$, EmpNo1$, EmpName1$, RateCode$, RateMNm$, DocDate$
    Dim Contents$, EvalScr$
    
    
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    Year = Trim(oForm01.Items("Year").Specific.VALUE)
    MSTCOD = Trim(oForm01.Items("MSTCOD").Specific.VALUE)
    FULLNAME = Trim(oForm01.Items("FULLNAME").Specific.VALUE)
    EmpNo1 = Trim(oForm01.Items("EmpNo1").Specific.VALUE)
    EmpName1 = Trim(oForm01.Items("EmpName1").Specific.VALUE)
    RateCode = Trim(oForm01.Items("RateCode").Specific.VALUE)
    RateMNm = Trim(oForm01.Items("RateMNm").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    Contents = Trim(oForm01.Items("Contents").Specific.VALUE)
    EvalScr = Trim(oForm01.Items("EvalScr").Specific.VALUE)
    sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_HR405H]"
    RecordSet01.DoQuery sQry
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
        
'    sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_Code),3)) + 1)),3) From [@PS_HR405H] Where U_BPLId = '" & BPLId & "' And U_GCode = '" & GCode & "'"
'    RecordSet01.DoQuery sQry
'    If Trim(RecordSet01.Fields(0).VALUE) = "" Then
'        Code = BPLId + GCode + "001"
'    Else
'        Code = BPLId + GCode + Trim(RecordSet01.Fields(0).VALUE)
'    End If
    
    If PasswordChk(pval) = False Then
        Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
        oForm01.Items("PassWd").Click ct_Regular
    Else
        
    
    
    
        sQry = "INSERT INTO [@PS_HR405H]"
        sQry = sQry & " ("
        sQry = sQry & " DocEntry,"
        sQry = sQry & " DocNum,"
        sQry = sQry & " U_BPLId,"
        sQry = sQry & " U_Year,"
        sQry = sQry & " U_MSTCOD,"
        sQry = sQry & " U_FULLNAME,"
        sQry = sQry & " U_EmpNo1,"
        sQry = sQry & " U_EmpName1,"
        sQry = sQry & " U_RateCode,"
        sQry = sQry & " U_RateMNm,"
        sQry = sQry & " U_DocDate,"
        sQry = sQry & " U_Contents,"
        sQry = sQry & " U_EvalScr"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("
        sQry = sQry & DocEntry & ","
        sQry = sQry & DocEntry & ","
        sQry = sQry & "'" & BPLID & "',"
        sQry = sQry & "'" & Year & "',"
        sQry = sQry & "'" & MSTCOD & "',"
        sQry = sQry & "'" & FULLNAME & "',"
        sQry = sQry & "'" & EmpNo1 & "',"
        sQry = sQry & "'" & EmpName1 & "',"
        sQry = sQry & "'" & RateCode & "',"
        sQry = sQry & "'" & RateMNm & "',"
        sQry = sQry & "'" & DocDate & "',"
        sQry = sQry & "'" & Contents & "',"
        sQry = sQry & "'" & EvalScr & "'"
        sQry = sQry & ")"
        RecordSet02.DoQuery sQry
        
        
        MDC_Com.MDC_GF_Message "등록 완료!", "S"
        
        gYear = Year
        gMSTCOD = MSTCOD
        gFULLNAME = FULLNAME
        gEmpNo1 = EmpNo1
        gEmpName1 = EmpName1
        gPassWd = Trim(oForm01.Items("PassWd").Specific.VALUE)
        
        Set RecordSet01 = Nothing
        Set RecordSet02 = Nothing
        Add_PurchaseDemand = True
    
        Call FormReset
        
    End If
    
    
    
    
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
        Case Trim(oForm01.Items("Year").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("MSTCOD").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("FULLNAME").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("EmpNo1").Specific.VALUE) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("EmpName1").Specific.VALUE) = 0
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("RateCode").Specific.VALUE) = 0
            ErrNum = 6
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("RateMNm").Specific.VALUE) = 0
            ErrNum = 7
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("DocDate").Specific.VALUE) = 0
            ErrNum = 8
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Contents").Specific.VALUE) = 0
            ErrNum = 9
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "년도는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "평가자는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "평가자명은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "피평가자는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "피평가자명은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "평가항목은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "평가항목명은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "평가일은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "평가내용은 필수사항입니다. 확인하세요.", "E"
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
        Case "MSTCOD" '/평가자
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("MSTCOD").Specific.VALUE) & "'"
            sQry = sQry & " And  branch = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
            oRecordSet01.DoQuery sQry
            oForm01.Items("FULLNAME").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
'            oForm01.Items("SEmpNo").Specific.VALUE = Trim(oForm01.Items("EmpNo").Specific.VALUE)
'            oForm01.Items("SEmpName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
'        Case "SEmpNo" '/평가자
'            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("SEmpNo").Specific.VALUE) & "'"
'            sQry = sQry & " And  branch = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
'            oRecordSet01.DoQuery sQry
'            oForm01.Items("SEmpName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'
'            oForm01.Items("EmpNo").Specific.VALUE = Trim(oForm01.Items("SEmpNo").Specific.VALUE)
'            oForm01.Items("EmpName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "EmpNo1" '/피평가자
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("EmpNo1").Specific.VALUE) & "'"
            sQry = sQry & " And  branch = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
            oRecordSet01.DoQuery sQry
            oForm01.Items("EmpName1").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
'            oForm01.Items("SEmpNo1").Specific.VALUE = Trim(oForm01.Items("EmpNo1").Specific.VALUE)
'            oForm01.Items("SEmpName1").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
'        Case "SEmpNo1" '/피평가자
'            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("SEmpNo1").Specific.VALUE) & "'"
'            sQry = sQry & " And  branch = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
'            oRecordSet01.DoQuery sQry
'            oForm01.Items("SEmpName1").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'
'            oForm01.Items("EmpNo1").Specific.VALUE = Trim(oForm01.Items("SEmpNo1").Specific.VALUE)
'            oForm01.Items("EmpName1").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "RateCode" '//평가항목
            sQry = "SELECT  T1.U_RateMNm FROM [@PS_HR400H] AS T0 Inner Join [@PS_HR400L] AS T1 On T0.Code = T1.Code "
            sQry = sQry & " WHERE   T0.U_BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
            sQry = sQry & " And T0.U_Year = '" & oForm01.Items("Year").Specific.VALUE & "' "
            sQry = sQry & " And T1.U_RateCode = '" & oForm01.Items("RateCode").Specific.VALUE & "'"
            
            oRecordSet01.DoQuery sQry
            oForm01.Items("RateMNm").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    End Select

'   Set oRecordSet = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'/폼의 아이템 사용지정
Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR405.srf")
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
    
    oFormUniqueID01 = "PS_HR405_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_HR405_CreateItems
    Call PS_HR405_ComboBox_Setting
    Call PS_HR405_CF_ChooseFromList
    Call PS_HR405_EnableMenus
    Call PS_HR405_SetDocument(oFromDocEntry01)
    Call PS_HR405_FormResize
    
    
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
        
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_HR405H]"
    RecordSet01.DoQuery sQry
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
        
    
        
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
    '//일자 Set
     oForm01.Items("DocDate").Specific.VALUE = Format(Date, "YYYYMMDD")

     'oForm01.Items("GCode").Click ct_Regular
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
                If pval.ItemUID = "MSTCOD" Then
                    If oForm01.Items("MSTCOD").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
'                If pval.ItemUID = "SEmpNo" Then
'                    If oForm01.Items("SEmpNo").Specific.VALUE = "" Then
'                        Sbo_Application.ActivateMenuItem ("7425")
'                        BubbleEvent = False
'                    End If
'                End If
                
                If pval.ItemUID = "EmpNo1" Then
                    If oForm01.Items("EmpNo1").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If

                End If
                
'                If pval.ItemUID = "SEmpNo1" Then
'                    If oForm01.Items("SEmpNo1").Specific.VALUE = "" Then
'                        Sbo_Application.ActivateMenuItem ("7425")
'                        BubbleEvent = False
'                    End If
'
'                End If
                If pval.ItemUID = "RateCode" Then
                    If oForm01.Items("RateCode").Specific.VALUE = "" Then
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
                    If pval.ItemUID = "MSTCOD" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "EmpNo1" Then
                        FlushToItemValue pval.ItemUID
                    End If
'                    If pval.ItemUID = "SEmpNo" Then
'                        FlushToItemValue pval.ItemUID
'                    End If
'                    If pval.ItemUID = "SEmpNo1" Then
'                        FlushToItemValue pval.ItemUID
'                    End If
                    If pval.ItemUID = "RateCode" Then
                        FlushToItemValue pval.ItemUID
                    End If
                End If
            Call Raise_EVENT_VALIDATE(FormUID, pval, BubbleEvent)
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
                
                
                Call FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call LoadCaption
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_HR405_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_HR405H.Clear

'                Call LoadCaption
'                Call FormItemEnabled
                '//Call PS_HR405_FormItemEnabled '//UDO방식
                '//Call PS_HR405_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_HR405_FormItemEnabled
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
        If pval.ItemUID = "PS_HR405" Then
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

                Call FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call LoadCaption
                Call LoadData
                
                oLast_Mode = oForm01.Mode
                
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If UpdateData(pval) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call FormReset
                oForm01.Mode = fm_ADD_MODE
                Call LoadCaption
                Call LoadData
                
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "Btn_ret" Then '/조회
            If PasswordChk(pval) = False Then
                Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
                oForm01.Items("PassWd").Click ct_Regular
            Else
                Call LoadData
            End If
            
        ElseIf pval.ItemUID = "Btn_del" Then '/삭제
'            If DeleteBefor = False Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            
            Call DeleteData
            Call LoadData
            oForm01.Mode = fm_OK_MODE '/fm_VIEW_MODE
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_HR405" Then
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

Dim i As Integer
Dim sQry As String
Dim oRecordSet01 As SAPbobsCOM.Recordset
Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

Dim DocEntry$
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            
            
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)
                
                DocEntry = oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE
                
                If Trim(DocEntry) <> "" Then
                    sQry = "EXEC [PS_HR405_02] '" & DocEntry & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If (oRecordSet01.RecordCount = 0) Then
                        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                        Set oRecordSet01 = Nothing
                        oForm01.Mode = fm_ADD_MODE
                        
                        Call LoadCaption
                        Exit Sub
                    End If
                    
                    oForm01.Freeze True
                    
                    Call oDS_PS_HR405H.setValue("DocEntry", 0, Trim(oRecordSet01.Fields("DocEntry").VALUE))
                    Call oDS_PS_HR405H.setValue("U_BPLId", 0, Trim(oRecordSet01.Fields("BPLId").VALUE))
                    Call oDS_PS_HR405H.setValue("U_Year", 0, Trim(oRecordSet01.Fields("Year").VALUE))
                    Call oDS_PS_HR405H.setValue("U_MSTCOD", 0, Trim(oRecordSet01.Fields("MSTCOD").VALUE))
                    Call oDS_PS_HR405H.setValue("U_FULLNAME", 0, Trim(oRecordSet01.Fields("FULLNAME").VALUE))
                    Call oDS_PS_HR405H.setValue("U_EmpNo1", 0, Trim(oRecordSet01.Fields("EmpNo1").VALUE))
                    Call oDS_PS_HR405H.setValue("U_EmpName1", 0, Trim(oRecordSet01.Fields("EmpName1").VALUE))
                    Call oDS_PS_HR405H.setValue("U_RateCode", 0, Trim(oRecordSet01.Fields("RateCode").VALUE))
                    Call oDS_PS_HR405H.setValue("U_RateMNm", 0, Trim(oRecordSet01.Fields("RateMNm").VALUE))
                    Call oDS_PS_HR405H.setValue("U_DocDate", 0, Format(Trim(oRecordSet01.Fields("DocDate").VALUE), "YYYYMMDD"))
                    Call oDS_PS_HR405H.setValue("U_Contents", 0, Trim(oRecordSet01.Fields("Contents").VALUE))
                    Call oDS_PS_HR405H.setValue("U_EvalScr", 0, Trim(oRecordSet01.Fields("EvalScr").VALUE))
                    
'                    oForm01.Items("SEmpNo").Specific.VALUE = Trim(oRecordSet01.Fields("EmpNo").VALUE)
'                    oForm01.Items("SEmpName").Specific.VALUE = Trim(oRecordSet01.Fields("EmpName").VALUE)
'
'                    oForm01.Items("SEmpNo1").Specific.VALUE = Trim(oRecordSet01.Fields("EmpNo1").VALUE)
'                    oForm01.Items("SEmpName1").Specific.VALUE = Trim(oRecordSet01.Fields("EmpName1").VALUE)
                    
                    oForm01.Freeze False
                    oForm01.Mode = fm_UPDATE_MODE
                    
                    
                    Call LoadCaption
                End If
                
                
                
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
Dim ClasCode As String
Dim BPLID As String

Dim oRecordSet01      As SAPbobsCOM.Recordset

Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
'        If pval.ItemUID = "BPLId" Then
'            BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
'
'        End If
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
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_HR405L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_HR405L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_HR405_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_HR405L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
'                If (pval.ItemUID = "Year") Then
'                    oForm01.Items("SYear").Specific.VALUE = Trim(oForm01.Items("Year").Specific.VALUE)
'                End If
                
                If (pval.ItemUID = "PassWd") Then
    
                    If PasswordChk(pval) = False Then
                        Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
                        oForm01.Items("PassWd").Click ct_Regular
                    Else
'                        oForm01.Items("SPassWd").Specific.VALUE = Trim(oForm01.Items("PassWd").Specific.VALUE)
                    End If

'                    sQry = "Select Count(*) From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("EmpNo").Specific.VALUE) & "'"
'                    sQry = sQry & " And  branch = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
'                    sQry = sQry & " And  govId = '" & oForm01.Items("PassWd").Specific.VALUE & "' "
'                    oRecordSet01.DoQuery sQry
'
'                    If Trim(oRecordSet01.Fields(0).VALUE) <= 0 Then
'                        Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
'                        oForm01.Items("PassWd").Click ct_Regular
'                    Else
'                        oForm01.Items("SPassWd").Specific.VALUE = Trim(oForm01.Items("PassWd").Specific.VALUE)
'                    End If
                End If
                
'                If (pval.ItemUID = "SPassWd") Then
'
'
'                    oForm01.Items("PassWd").Specific.VALUE = Trim(oForm01.Items("SPassWd").Specific.VALUE)
'                End If
            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
        End If
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
        Call PS_HR405_FormItemEnabled
        '//Call PS_HR405_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_HR405_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_HR405H", "U_CardCode,U_CardName")
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
'            If (PS_HR405_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_HR405H.RemoveRecord(oDS_PS_HR405H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_HR405H.GetValue("U_MSTCOD", oMat01.RowCount - 1)) <> "" Then
                    Call Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_HR405_CreateItems() As Boolean
On Error GoTo PS_HR405_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
         
    
    
    Set oDS_PS_HR405H = oForm01.DataSources.DBDataSources("@PS_HR405H")
    Set oDS_PS_HR405L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    '// 메트릭스 개체 할당
  
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
      
    Call oForm01.DataSources.UserDataSources.Add("PassWd", dt_SHORT_TEXT, 10)
    Call oForm01.Items("PassWd").Specific.DataBind.SetBound(True, "", "PassWd")
    
    oForm01.Items("Year").Specific.VALUE = Format(Now, "YYYY")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_HR405_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR405_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
'/콤보박스 set
Sub PS_HR405_ComboBox_Setting()
On Error GoTo PS_HR405_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
        
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_HR405_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR405_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR405_CF_ChooseFromList()
On Error GoTo PS_HR405_CF_ChooseFromList_Error

    Exit Sub
PS_HR405_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR405_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR405_FormItemEnabled()
On Error GoTo PS_HR405_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_HR405_FormClear '//UDO방식
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
PS_HR405_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR405_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR405_EnableMenus()
On Error GoTo PS_HR405_EnableMenus_Error
    
    Exit Sub
PS_HR405_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR405_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR405_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_HR405_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_HR405_FormItemEnabled
        '//Call PS_HR405_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_HR405_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_HR405_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR405_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR405_FormResize()
On Error GoTo PS_HR405_FormResize_Error
    
    Exit Sub
PS_HR405_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR405_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub FormReset()
On Error GoTo FormReset_Error

Dim sQry As String
Dim RecordSet01 As SAPbobsCOM.Recordset
Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
   

sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_HR405H]"
RecordSet01.DoQuery sQry
If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
    oForm01.Items("DocEntry").Specific.VALUE = 1
Else
    oForm01.Items("DocEntry").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE) + 1
End If
Call oDS_PS_HR405H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId)
'Call oDS_PS_HR405H.setValue("U_Year", 0, "")
Call oDS_PS_HR405H.setValue("U_Year", 0, gYear)
Call oDS_PS_HR405H.setValue("U_MSTCOD", 0, gMSTCOD)
'Call oDS_PS_HR405H.setValue("U_EmpNo", 0, "")
Call oDS_PS_HR405H.setValue("U_FULLNAME", 0, gFULLNAME)
'Call oDS_PS_HR405H.setValue("U_EmpName", 0, "")
'Call oDS_PS_HR405H.setValue("U_EmpNo1", 0, "")
'Call oDS_PS_HR405H.setValue("U_EmpName1", 0, "")
Call oDS_PS_HR405H.setValue("U_EmpNo1", 0, gEmpNo1)
Call oDS_PS_HR405H.setValue("U_EmpName1", 0, gEmpName1)
Call oDS_PS_HR405H.setValue("U_RateCode", 0, "")
Call oDS_PS_HR405H.setValue("U_RateMNm", 0, "")
Call oDS_PS_HR405H.setValue("U_DocDate", 0, Format(Date, "YYYYMMDD"))
Call oDS_PS_HR405H.setValue("U_Contents", 0, "")
Call oDS_PS_HR405H.setValue("U_EvalScr", 0, 0)


oForm01.Items("PassWd").Specific.VALUE = gPassWd


oForm01.Mode = fm_ADD_MODE

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormReset_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


'// 패스워드 Check
Public Function PasswordChk(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo PasswordChk_Error
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim MSTCOD$, PassWd$
        
    MSTCOD = Trim(oForm01.Items("MSTCOD").Specific.VALUE)
    PassWd = Trim(oForm01.Items("PassWd").Specific.VALUE)
    
    If Trim(MSTCOD) = "" Then
        MDC_Com.MDC_GF_Message "사번이 없습니다. 입력바랍니다!", "E"
        PasswordChk = False
        Exit Function
    End If
        
    sQry = "Select Count(*) From Z_PS_HRPASS Where MSTCOD = '" & Trim(oForm01.Items("MSTCOD").Specific.VALUE) & "'"
    sQry = sQry & " And  BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
    sQry = sQry & " And  PassWd = '" & oForm01.Items("PassWd").Specific.VALUE & "' "
    RecordSet01.DoQuery sQry
                
    If Trim(RecordSet01.Fields(0).VALUE) <= 0 Then
        PasswordChk = False
    Else
        PasswordChk = True
    End If
        
    Set RecordSet01 = Nothing
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PasswordChk_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PasswordChk_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

