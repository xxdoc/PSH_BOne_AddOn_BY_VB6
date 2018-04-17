VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FX005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_FX005.cls
'//  Module         : FX
'//  Desc           : 고정자산등록
'//  FormType       : PS_FX005
'//  Create Date    : 2012.07.17 N.G.Y
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_FX005H  As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_FX005L  As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private oLast_Mode     As String

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

    Dim i    As Integer
    Dim sQry As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim SBPLID    As String
    Dim SClasCod  As String
    Dim SFixCode  As String
    Dim SFixName  As String
    Dim STempChr1 As String
    
    SBPLID = Trim(oForm01.Items("SBPLId").Specific.VALUE)
    SClasCod = Trim(oForm01.Items("SClasCod").Specific.VALUE)
    SFixCode = Trim(oForm01.Items("SFixCode").Specific.VALUE)
    SFixName = Trim(oForm01.Items("SFixName").Specific.VALUE)
    STempChr1 = Trim(oForm01.Items("STempChr1").Specific.VALUE)
    
    If SClasCod = "" Then SClasCod = "%"
    If SFixCode = "" Then SFixCode = "%"
    If SFixName = "" Then SFixName = "%"
    If STempChr1 = "" Then STempChr1 = "%"
 
    sQry = "      EXEC [PS_FX005_01] '"
    sQry = sQry & SBPLID & "','"
    sQry = sQry & SClasCod & "', '"
    sQry = sQry & SFixCode & "', '"
    sQry = sQry & SFixName & "', '"
    sQry = sQry & STempChr1 & "'"
    
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_FX005L.Clear
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
        If i + 1 > oDS_PS_FX005L.Size Then
            oDS_PS_FX005L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_FX005L.offset = i
        
        oDS_PS_FX005L.setValue "U_LineNum", i, i + 1
        oDS_PS_FX005L.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields(0).VALUE) '//BPLId
        oDS_PS_FX005L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields(1).VALUE) '//DocEntry
        oDS_PS_FX005L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields(2).VALUE) '//FixCode
        oDS_PS_FX005L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields(3).VALUE) '//SubCode
        oDS_PS_FX005L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields(4).VALUE) '//FixName
        oDS_PS_FX005L.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields(5).VALUE) '//ClasCode
        oDS_PS_FX005L.setValue "U_ColReg07", i, Trim(oRecordSet01.Fields(6).VALUE) '//PostDate
        oDS_PS_FX005L.setValue "U_ColQty01", i, Trim(oRecordSet01.Fields(7).VALUE) '//PostQty
        oDS_PS_FX005L.setValue "U_ColSum01", i, Trim(oRecordSet01.Fields(8).VALUE) '//PostAmt
        oDS_PS_FX005L.setValue "U_ColReg08", i, Trim(oRecordSet01.Fields(9).VALUE) '//TeamNm
        oDS_PS_FX005L.setValue "U_ColReg09", i, Trim(oRecordSet01.Fields(10).VALUE) '//RspNm
        oDS_PS_FX005L.setValue "U_ColReg10", i, Trim(oRecordSet01.Fields(11).VALUE) '//StopYN
'        oDS_PS_FX005L.setValue "U_ColSum01", i, Trim(oRecordSet01.Fields(9).VALUE)
'        oDS_PS_FX005L.setValue "U_ColReg09", i, Trim(oRecordSet01.Fields(10).VALUE)
'        oDS_PS_FX005L.setValue "U_ColDt01", i, Format(Trim(oRecordSet01.Fields(11).VALUE), "YYYYMMDD")
        
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

LoadData_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub DeleteData()
On Error GoTo DeleteData_Error

    Dim i    As Integer
    Dim sQry As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    Dim FixCode  As String
    Dim Check    As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        FixCode = Trim(oForm01.Items("FixCode").Specific.VALUE)
        
        sQry = "        Select      Count(*) "
        sQry = sQry & " From        [@PS_FX005H] "
        sQry = sQry & " where       DocEntry = '" & DocEntry & "' "
        sQry = sQry & "             And U_FixCode = '" & FixCode & "'"
        
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            MDC_Com.MDC_GF_Message "삭제대상이 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
            Set oRecordSet01 = Nothing
            oForm01.Mode = fm_ADD_MODE
        Else
            sQry = "Delete From [@PS_FX005H] where DocEntry = '" & DocEntry & "' And U_FixCode = '" & FixCode & "'"
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

DeleteData_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "DeleteData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PS_FX005L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_FX005L.offset = oRow
    oDS_PS_FX005L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub

Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function UpdateData(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo UpdateData_Error

    Dim i    As Integer
    Dim j    As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    Dim BPLID    As String
    Dim ClasCode As String
    Dim FixCode  As String
    Dim FixName  As String
    Dim PostDate As String
    Dim PostAmt  As String
    Dim PostQty  As String
    Dim LongYear As String
    Dim DepCode  As String
    Dim DepRate  As String
    Dim CardCode As String
    Dim CardName As String
    Dim Kw       As String
    Dim Volt     As String
    Dim TeamCode As String
    Dim TeamNm   As String
    Dim RspCode  As String
    Dim RspNm    As String
    Dim PrcCode  As String
    Dim PrcName  As String
    Dim Comments As String
    Dim Place    As String
    Dim TempChr1 As String
    Dim StopYN   As String
    Dim SubDiv   As String
    Dim SubCode  As String
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    ClasCode = Trim(oForm01.Items("ClasCode").Specific.VALUE)
    SubDiv = Trim(oForm01.Items("SubDiv").Specific.VALUE)
    FixCode = Trim(oForm01.Items("FixCode").Specific.VALUE)
    SubCode = Trim(oForm01.Items("SubCode").Specific.VALUE)
    FixName = Trim(oForm01.Items("FixName").Specific.VALUE)
    PostDate = Trim(oForm01.Items("PostDate").Specific.VALUE)
    PostAmt = Trim(oForm01.Items("PostAmt").Specific.VALUE)
    PostQty = Trim(oForm01.Items("PostQty").Specific.VALUE)
    LongYear = Trim(oForm01.Items("LongYear").Specific.VALUE)
    DepCode = Trim(oForm01.Items("DepCode").Specific.VALUE)
    DepRate = Trim(oForm01.Items("DepRate").Specific.VALUE)
    If oForm01.Items("StopYN").Specific.Checked = True Then
        StopYN = "Y"
    Else
        StopYN = "N"
    End If
    
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    CardName = Trim(oForm01.Items("CardName").Specific.VALUE)
    Kw = Trim(oForm01.Items("Kw").Specific.VALUE)
    Volt = Trim(oForm01.Items("Volt").Specific.VALUE)
    TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE)
    TeamNm = Trim(oForm01.Items("TeamNm").Specific.VALUE)
    RspCode = Trim(oForm01.Items("RspCode").Specific.VALUE)
    RspNm = Trim(oForm01.Items("RspNm").Specific.VALUE)
    PrcCode = Trim(oForm01.Items("PrcCode").Specific.VALUE)
    PrcName = Trim(oForm01.Items("PrcName").Specific.VALUE)
    Comments = Trim(oForm01.Items("Comments").Specific.VALUE)
    Place = Trim(oForm01.Items("Place").Specific.VALUE)
    TempChr1 = Trim(oForm01.Items("TempChr1").Specific.VALUE)
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!", "E"
        UpdateData = False
        Exit Function
    End If

    sQry = "        Update [@PS_FX005H]"
    sQry = sQry & " set "
    sQry = sQry & " U_BPLId = '" & BPLID & "',"
    sQry = sQry & " U_ClasCode = '" & ClasCode & "',"
    sQry = sQry & " U_SubDiv = '" & SubDiv & "',"
    sQry = sQry & " U_FixCode = '" & FixCode & "',"
    sQry = sQry & " U_SubCode = '" & SubCode & "',"
    sQry = sQry & " U_FixName = '" & FixName & "',"
    sQry = sQry & " U_PostDate = '" & PostDate & "',"
    sQry = sQry & " U_PostAmt = '" & PostAmt & "',"
    sQry = sQry & " U_PostQty  = '" & PostQty & "',"
    sQry = sQry & " U_LongYear  = '" & LongYear & "',"
    sQry = sQry & " U_DepCode  = '" & DepCode & "',"
    sQry = sQry & " U_DepRate = '" & DepRate & "',"
    sQry = sQry & " U_StopYN = '" & StopYN & "',"
    sQry = sQry & " U_CardCode  = '" & CardCode & "',"
    sQry = sQry & " U_CardName  = '" & CardName & "',"
    sQry = sQry & " U_Kw  = '" & Kw & "',"
    sQry = sQry & " U_Volt  = '" & Volt & "',"
    sQry = sQry & " U_TeamCode  = '" & TeamCode & "',"
    sQry = sQry & " U_TeamNm  = '" & TeamNm & "',"
    sQry = sQry & " U_RspCode  = '" & RspCode & "',"
    sQry = sQry & " U_RspNm  = '" & RspNm & "',"
    sQry = sQry & " U_PrcCode  = '" & PrcCode & "',"
    sQry = sQry & " U_PrcName  = '" & PrcName & "',"
    sQry = sQry & " U_Place  = '" & Place & "',"
    sQry = sQry & " U_TempChr1  = '" & TempChr1 & "',"
    sQry = sQry & " U_Comments  = '" & Comments & "'"
            
    sQry = sQry & " Where DocEntry = '" & DocEntry & "'"
    RecordSet01.DoQuery sQry
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    UpdateData = True
    Exit Function

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
    
   Dim DocEntry$, BPLID$, ClasCode$, FixCode$, FixName$, PostDate$, PostAmt$, PostQty$, LongYear$, DepCode$
    Dim DepRate$, CardCode$, CardName$, Kw$, Volt$, TeamCode$, TeamNm$, RspCode$, RspNm$, PrcCode$, PrcName$, Comments$
    Dim Place$, TempChr1$, StopYN$
    Dim SubDiv$, SubCode$
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    ClasCode = Trim(oForm01.Items("ClasCode").Specific.VALUE)
    SubDiv = Trim(oForm01.Items("SubDiv").Specific.VALUE)
    FixCode = Trim(oForm01.Items("FixCode").Specific.VALUE)
    FixName = Trim(oForm01.Items("FixName").Specific.VALUE)
    SubCode = Trim(oForm01.Items("SubCode").Specific.VALUE)
    PostDate = Trim(oForm01.Items("PostDate").Specific.VALUE)
    PostAmt = Trim(oForm01.Items("PostAmt").Specific.VALUE)
    PostQty = Trim(oForm01.Items("PostQty").Specific.VALUE)
    LongYear = Trim(oForm01.Items("LongYear").Specific.VALUE)
    DepCode = Trim(oForm01.Items("DepCode").Specific.VALUE)
    DepRate = Trim(oForm01.Items("DepRate").Specific.VALUE)
    If oForm01.Items("StopYN").Specific.Checked = True Then
        StopYN = "Y"
    Else
        StopYN = "N"
    End If
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    CardName = Trim(oForm01.Items("CardName").Specific.VALUE)
    Kw = Trim(oForm01.Items("Kw").Specific.VALUE)
    Volt = Trim(oForm01.Items("Volt").Specific.VALUE)
    TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE)
    TeamNm = Trim(oForm01.Items("TeamNm").Specific.VALUE)
    RspCode = Trim(oForm01.Items("RspCode").Specific.VALUE)
    RspNm = Trim(oForm01.Items("RspNm").Specific.VALUE)
    PrcCode = Trim(oForm01.Items("PrcCode").Specific.VALUE)
    PrcName = Trim(oForm01.Items("PrcName").Specific.VALUE)
    Comments = Trim(oForm01.Items("Comments").Specific.VALUE)
    Place = Trim(oForm01.Items("Place").Specific.VALUE)
    TempChr1 = Trim(oForm01.Items("TempChr1").Specific.VALUE)
    
    sQry = "Select right('00' + Convert(Nvarchar(3),Convert(Numeric(3,0), Isnull(Max(U_SubCode),'000')) + 1),3) From [@PS_FX005H] Where U_FixCode = '" & FixCode & "'"
    RecordSet01.DoQuery sQry
    
    SubCode = Trim(RecordSet01.Fields(0).VALUE)
    
    sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_FX005H]"
    RecordSet01.DoQuery sQry
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
        
'    sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_Code),3)) + 1)),3) From [@PS_FX005H] Where U_BPLId = '" & BPLId & "' And U_GCode = '" & GCode & "'"
'    RecordSet01.DoQuery sQry
'    If Trim(RecordSet01.Fields(0).VALUE) = "" Then
'        Code = BPLId + GCode + "001"
'    Else
'        Code = BPLId + GCode + Trim(RecordSet01.Fields(0).VALUE)
'    End If
    
    sQry = "INSERT INTO [@PS_FX005H]"
    sQry = sQry & " ("
    sQry = sQry & " DocEntry,"
    sQry = sQry & " DocNum,"
    sQry = sQry & " U_BPLId,"
    sQry = sQry & " U_ClasCode,"
    sQry = sQry & " U_SubDiv,"
    sQry = sQry & " U_FixCode,"
    sQry = sQry & " U_SubCode,"
    sQry = sQry & " U_FixName,"
    sQry = sQry & " U_PostDate,"
    sQry = sQry & " U_PostAmt,"
    sQry = sQry & " U_PostQty,"
    sQry = sQry & " U_LongYear,"
    sQry = sQry & " U_DepCode,"
    sQry = sQry & " U_DepRate,"
    sQry = sQry & " U_StopYN,"
    sQry = sQry & " U_CardCode,"
    sQry = sQry & " U_CardName,"
    sQry = sQry & " U_Kw,"
    sQry = sQry & " U_Volt,"
    sQry = sQry & " U_TeamCode,"
    sQry = sQry & " U_TeamNm,"
    sQry = sQry & " U_RspCode,"
    sQry = sQry & " U_RspNm,"
    sQry = sQry & " U_PrcCode,"
    sQry = sQry & " U_PrcName,"
    sQry = sQry & " U_Place,"
    sQry = sQry & " U_Comments,"
    sQry = sQry & " U_TempChr1"
    sQry = sQry & " ) "
    sQry = sQry & "VALUES("
    sQry = sQry & DocEntry & ","
    sQry = sQry & DocEntry & ","
    sQry = sQry & "'" & BPLID & "',"
    sQry = sQry & "'" & ClasCode & "',"
    sQry = sQry & "'" & SubDiv & "',"
    sQry = sQry & "'" & FixCode & "',"
    sQry = sQry & "'" & SubCode & "',"
    sQry = sQry & "'" & FixName & "',"
    sQry = sQry & "'" & PostDate & "',"
    sQry = sQry & "'" & PostAmt & "',"
    sQry = sQry & "'" & PostQty & "',"
    sQry = sQry & "'" & LongYear & "',"
    sQry = sQry & "'" & DepCode & "',"
    sQry = sQry & "'" & DepRate & "',"
    sQry = sQry & "'" & StopYN & "',"
    sQry = sQry & "'" & CardCode & "',"
    sQry = sQry & "'" & CardName & "',"
    sQry = sQry & "'" & Kw & "',"
    sQry = sQry & "'" & Volt & "',"
    sQry = sQry & "'" & TeamCode & "',"
    sQry = sQry & "'" & TeamNm & "',"
    sQry = sQry & "'" & RspCode & "',"
    sQry = sQry & "'" & RspNm & "',"
    sQry = sQry & "'" & PrcCode & "',"
    sQry = sQry & "'" & PrcName & "',"
    sQry = sQry & "'" & Place & "',"
    sQry = sQry & "'" & Comments & "',"
    sQry = sQry & "'" & TempChr1 & "'"
    sQry = sQry & ")"
    RecordSet02.DoQuery sQry
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Add_PurchaseDemand = True
    
    Exit Function

Add_PurchaseDemand_Error:
    Add_PurchaseDemand = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Add_PurchaseDemand_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("ClasCode").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("FixCode").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("FixName").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("PostDate").Specific.VALUE) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("PostQty").Specific.VALUE) = 0
            ErrNum = 8
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
    Exit Function

HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "자산분류는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "자산코드는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "자산명은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "취득일자는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "취득수량은 필수사항입니다. 확인하세요.", "E"
    End If
    HeaderSpaceLineDel = False
End Function

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

    Dim i               As String
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim ItemCode        As String
    Dim Qty             As Long
    Dim Calculate_Weight As Double
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CardCode" '/거래처
        
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "FixCode"
        
            sQry = "Select right('00' + Convert(Nvarchar(3),Convert(Numeric(3,0), Max(U_SubCode)) + 1),3) From [@PS_FX005H] Where U_FixCode = '" & oForm01.Items("FixCode").Specific.VALUE & "'"
            oRecordSet01.DoQuery sQry
                
            oForm01.Items("SubCode").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "TeamCode" '//팀
        
            sQry = "SELECT  T1.U_CodeNm FROM [@PS_HR200H] AS T0 Inner Join [@PS_HR200L] AS T1 ON T0.Code = T1.Code "
            sQry = sQry & " WHERE   T0.Name = '부서' AND T1.U_UseYN = 'Y' AND T1.U_Char2 = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
            sQry = sQry & " And T1.U_Code = '" & oForm01.Items("TeamCode").Specific.VALUE & "'"
            
            oRecordSet01.DoQuery sQry
            oForm01.Items("TeamNm").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "RspCode" '//담당
        
            sQry = "SELECT  T1.U_CodeNm FROM [@PS_HR200H] AS T0 Inner Join [@PS_HR200L] AS T1 ON T0.Code = T1.Code "
            sQry = sQry & " WHERE   T0.Name = '담당' AND T1.U_UseYN = 'Y' AND T1.U_Char2 = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
            sQry = sQry & " And T1.U_Char1 = '" & oForm01.Items("TeamCode").Specific.VALUE & "'"
            sQry = sQry & " And T1.U_Code = '" & oForm01.Items("RspCode").Specific.VALUE & "'"
            
            oRecordSet01.DoQuery sQry
            oForm01.Items("RspNm").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "PrcCode" '//
        
            sQry = "SELECT OCRNAME FROM OOCR WHERE DIMCODE = '1' And OCRCODE = '" & Trim(oForm01.Items("PrcCode").Specific.VALUE) & "'"
            
            oRecordSet01.DoQuery sQry
            oForm01.Items("PrcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FX005.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_FX005_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_FX005_CreateItems
    Call PS_FX005_ComboBox_Setting
    Call PS_FX005_CF_ChooseFromList
    Call PS_FX005_EnableMenus
    Call PS_FX005_SetDocument(oFromDocEntry01)
    Call PS_FX005_FormResize
    
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
        
    sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_FX005H]"
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
     oForm01.Items("PostDate").Specific.VALUE = Format(Date, "YYYYMMDD")

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
            If (pval.BeforeAction = True) Then
                If pval.CharPressed = 9 Then
                    
                    If pval.ItemUID = "CardCde" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
    
                    End If
                    If pval.ItemUID = "TeamCode" Then
                        If oForm01.Items("TeamCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
    
                    End If
                    If pval.ItemUID = "RspCode" Then
                        If oForm01.Items("RspCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
    
                    End If
                    If pval.ItemUID = "PrcCode" Then
                        If oForm01.Items("PrcCode").Specific.VALUE = "" Then
                            
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                        
    
                    End If
                    If pval.ItemUID = "FixCode" Then
                        If oForm01.Items("FixCode").Specific.VALUE = "" Then
                            If Trim(oForm01.Items("ClasCode").Specific.VALUE) = "" Then
                                MDC_Com.MDC_GF_Message "분류가 선택되지 않았습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                            Else
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
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
            If (pval.BeforeAction = False) Then
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "TeamCode" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "RspCode" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "PrcCode" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "FixCode" Then
                        FlushToItemValue pval.ItemUID
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
                '//Call PS_FX005_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_FX005H.Clear

'                Call LoadCaption
'                Call FormItemEnabled
                '//Call PS_FX005_FormItemEnabled '//UDO방식
                '//Call PS_FX005_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_FX005_FormItemEnabled
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
        If pval.ItemUID = "PS_FX005" Then
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
        
            Call LoadData
            
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
        If pval.ItemUID = "PS_FX005" Then
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
            
                Call oMat01.SelectRow(pval.Row, True, False)
                
                Call PS_FX005_DisplayFixData(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE)
        
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX005_DisplayFixData(ByVal pDocEntry As Integer)
'******************************************************************************
'Function ID : PS_FX005_DisplayFixData()
'해당모듈    : PS_FX005
'기능        : 고정자산 정보를 헤더에 출력
'인수        : 관리번호
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_FX005_DisplayFixData_Error

    Dim i        As Integer
    Dim sQry     As String
    Dim DocEntry As String

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    oForm01.Freeze True
    
    DocEntry = pDocEntry ' oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE
    
    If Trim(DocEntry) <> "" Then
        
        sQry = "EXEC [PS_FX005_02] '" & DocEntry & "'"
        oRecordSet01.DoQuery sQry

        If (oRecordSet01.RecordCount = 0) Then
            MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
            Set oRecordSet01 = Nothing
            oForm01.Mode = fm_ADD_MODE

            Call LoadCaption
            Exit Sub
        End If

        Call oDS_PS_FX005H.setValue("U_Pic01", 0, "")
        Call oDS_PS_FX005H.setValue("U_Pic02", 0, "")
        Call oDS_PS_FX005H.setValue("U_Pic03", 0, "")
        Call oDS_PS_FX005H.setValue("U_Pic04", 0, "")
        Call oDS_PS_FX005H.setValue("U_Pic05", 0, "")
        Call oDS_PS_FX005H.setValue("U_Pic06", 0, "")

        Call oDS_PS_FX005H.setValue("U_Pic01", 0, "\\191.1.1.220\Asset_Pic\" & Trim(oRecordSet01.Fields("FixCode").VALUE) & "_" & Trim(oRecordSet01.Fields("SubCode").VALUE) & "_01.BMP")
        Call oDS_PS_FX005H.setValue("U_Pic02", 0, "\\191.1.1.220\Asset_Pic\" & Trim(oRecordSet01.Fields("FixCode").VALUE) & "_" & Trim(oRecordSet01.Fields("SubCode").VALUE) & "_02.BMP")
        Call oDS_PS_FX005H.setValue("U_Pic03", 0, "\\191.1.1.220\Asset_Pic\" & Trim(oRecordSet01.Fields("FixCode").VALUE) & "_" & Trim(oRecordSet01.Fields("SubCode").VALUE) & "_03.BMP")
        Call oDS_PS_FX005H.setValue("U_Pic04", 0, "\\191.1.1.220\Asset_Pic\" & Trim(oRecordSet01.Fields("FixCode").VALUE) & "_" & Trim(oRecordSet01.Fields("SubCode").VALUE) & "_04.BMP")
        Call oDS_PS_FX005H.setValue("U_Pic05", 0, "\\191.1.1.220\Asset_Pic\" & Trim(oRecordSet01.Fields("FixCode").VALUE) & "_" & Trim(oRecordSet01.Fields("SubCode").VALUE) & "_05.BMP")
        Call oDS_PS_FX005H.setValue("U_Pic06", 0, "\\191.1.1.220\Asset_Pic\" & Trim(oRecordSet01.Fields("FixCode").VALUE) & "_" & Trim(oRecordSet01.Fields("SubCode").VALUE) & "_06.BMP")
        
        Call oDS_PS_FX005H.setValue("DocEntry", 0, Trim(oRecordSet01.Fields("DocEntry").VALUE))
        Call oDS_PS_FX005H.setValue("U_BPLId", 0, Trim(oRecordSet01.Fields("BPLId").VALUE))
        Call oDS_PS_FX005H.setValue("U_ClasCode", 0, Trim(oRecordSet01.Fields("ClasCode").VALUE))
        Call oDS_PS_FX005H.setValue("U_SubDiv", 0, Trim(oRecordSet01.Fields("SubDiv").VALUE)) 'Sub자산 여부(2015.07.06 송명규 추가)
        Call oDS_PS_FX005H.setValue("U_FixCode", 0, Trim(oRecordSet01.Fields("FixCode").VALUE))
        Call oDS_PS_FX005H.setValue("U_SubCode", 0, Trim(oRecordSet01.Fields("SubCode").VALUE))
        Call oDS_PS_FX005H.setValue("U_FixName", 0, Trim(oRecordSet01.Fields("FixName").VALUE))
        Call oDS_PS_FX005H.setValue("U_PostDate", 0, Format(Trim(oRecordSet01.Fields("PostDate").VALUE), "YYYYMMDD"))
        Call oDS_PS_FX005H.setValue("U_PostQty", 0, Trim(oRecordSet01.Fields("PostQty").VALUE))
        Call oDS_PS_FX005H.setValue("U_PostAmt", 0, Trim(oRecordSet01.Fields("PostAmt").VALUE))
        Call oDS_PS_FX005H.setValue("U_StopYN", 0, Trim(oRecordSet01.Fields("StopYN").VALUE))
        Call oDS_PS_FX005H.setValue("U_LongYear", 0, Trim(oRecordSet01.Fields("LongYear").VALUE))
        Call oDS_PS_FX005H.setValue("U_DepCode", 0, Trim(oRecordSet01.Fields("DepCode").VALUE))
        Call oDS_PS_FX005H.setValue("U_DepRate", 0, Trim(oRecordSet01.Fields("DepRate").VALUE))
        Call oDS_PS_FX005H.setValue("U_CardCode", 0, Trim(oRecordSet01.Fields("CardCode").VALUE))
        Call oDS_PS_FX005H.setValue("U_CardName", 0, Trim(oRecordSet01.Fields("CardName").VALUE))
        Call oDS_PS_FX005H.setValue("U_Kw", 0, Trim(oRecordSet01.Fields("Kw").VALUE))
        Call oDS_PS_FX005H.setValue("U_Volt", 0, Trim(oRecordSet01.Fields("Volt").VALUE))
        Call oDS_PS_FX005H.setValue("U_TeamCode", 0, Trim(oRecordSet01.Fields("TeamCode").VALUE))
        Call oDS_PS_FX005H.setValue("U_TeamNm", 0, Trim(oRecordSet01.Fields("TeamNm").VALUE))
        Call oDS_PS_FX005H.setValue("U_RspCode", 0, Trim(oRecordSet01.Fields("RspCode").VALUE))
        Call oDS_PS_FX005H.setValue("U_RspNm", 0, Trim(oRecordSet01.Fields("RspNm").VALUE))
        Call oDS_PS_FX005H.setValue("U_Place", 0, Trim(oRecordSet01.Fields("Place").VALUE))
        Call oDS_PS_FX005H.setValue("U_Comments", 0, Trim(oRecordSet01.Fields("Comments").VALUE))
        Call oDS_PS_FX005H.setValue("U_TempChr1", 0, Trim(oRecordSet01.Fields("TempChr1").VALUE))
        Call oDS_PS_FX005H.setValue("U_PrcCode", 0, Trim(oRecordSet01.Fields("PrcCode").VALUE))
        Call oDS_PS_FX005H.setValue("U_PrcName", 0, Trim(oRecordSet01.Fields("PrcName").VALUE))

    End If

    oForm01.Freeze False
    oForm01.Mode = fm_UPDATE_MODE
    Call LoadCaption

    Set oRecordSet01 = Nothing

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Exit Sub

PS_FX005_DisplayFixData_Error:

    oForm01.Freeze False
    
    Set oRecordSet01 = Nothing

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_FX005_DisplayFixData_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

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
        If pval.ItemUID = "ClasCode" Or pval.ItemUID = "BPLId" Then
            If oForm01.Mode = fm_ADD_MODE Then
                '//신규모드일때
                If Trim(oForm01.Items("SubDiv").Specific.VALUE) = "N" Then
                    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
                    ClasCode = Trim(oForm01.Items("ClasCode").Specific.VALUE) 'Selected.
                    sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_FixCode),3)) + 1)),3) From [@PS_FX005H] Where U_BPLId = '" & BPLID & "' And U_ClasCode = '" & ClasCode & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(BPLID) <> "" And Trim(ClasCode) <> "%" Then
                        If Trim(oRecordSet01.Fields(0).VALUE) = "" Then
                            oForm01.Items("FixCode").Specific.String = BPLID + ClasCode + "001"
                        Else
                            oForm01.Items("FixCode").Specific.String = BPLID + ClasCode + Trim(oRecordSet01.Fields(0).VALUE)
                        End If
                        oForm01.Items("SubCode").Specific.VALUE = "001"
                    End If
                End If
            End If
        ElseIf pval.ItemUID = "SubDiv" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If Trim(oForm01.Items("SubDiv").Specific.VALUE) = "Y" Then
                    oForm01.Items("FixCode").Enabled = True
                    oForm01.Items("FixCode").Specific.VALUE = ""
                Else
                    oForm01.Items("FixCode").Enabled = False
                    oForm01.Items("SubCode").Specific.VALUE = "001"
                    
                    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
                    ClasCode = Trim(oForm01.Items("ClasCode").Specific.VALUE) 'Selected.
                    
                    sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_FixCode),3)) + 1)),3) From [@PS_FX005H] Where U_BPLId = '" & BPLID & "' And U_ClasCode = '" & ClasCode & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(BPLID) <> "" And Trim(ClasCode) <> "%" Then
                        If Trim(oRecordSet01.Fields(0).VALUE) = "" Then
                            oForm01.Items("FixCode").Specific.String = BPLID + ClasCode + "001"
                        Else
                            oForm01.Items("FixCode").Specific.String = BPLID + ClasCode + Trim(oRecordSet01.Fields(0).VALUE)
                        End If
                        oForm01.Items("SubCode").Specific.VALUE = "001"
                    End If
                End If
            End If
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
        
        If pval.ItemUID = "Pic01" Or pval.ItemUID = "Pic02" Or pval.ItemUID = "Pic03" Or pval.ItemUID = "Pic04" Or pval.ItemUID = "Pic05" Or pval.ItemUID = "Pic06" Then
        
            Call PS_FX005_LoadPic(pval.ItemUID)
            Call PS_FX005_DisplayFixData(oForm01.Items("DocEntry").Specific.VALUE)
            BubbleEvent = False
            
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
'/현재 사용안함
Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            
'            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_FX005L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_FX005L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_FX005_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_FX005L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
'            Else
'                If (pval.ItemUID = "DocEntry") Then
'                    Call oDS_PS_FX005H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
'                ElseIf (pval.ItemUID = "TeamCode") Then
'                    Call oDS_PS_FX005H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
'                    Call oDS_PS_FX005H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
'                Else
'                    Call oDS_PS_FX005H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
'                End If
'            End If
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
        Call PS_FX005_FormItemEnabled
        '//Call PS_FX005_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_FX005_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_FX005H", "U_CardCode,U_CardName")
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
'            If (PS_FX005_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_FX005H.RemoveRecord(oDS_PS_FX005H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_FX005H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_FX005_CreateItems() As Boolean
On Error GoTo PS_FX005_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_FX005H = oForm01.DataSources.DBDataSources("@PS_FX005H")
    Set oDS_PS_FX005L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    '// 메트릭스 개체 할당
  
    oForm01.Items("FixCode").Enabled = False
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_FX005_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX005_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
'/콤보박스 set
Sub PS_FX005_ComboBox_Setting()
On Error GoTo PS_FX005_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
        
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Set oCombo = oForm01.Items("SBPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Set oCombo = oForm01.Items("SubDiv").Specific
    oCombo.ValidValues.Add "N", "메인자산"
    oCombo.ValidValues.Add "Y", "Sub자산"
    oCombo.Select 0, psk_Index
    
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ClasCode").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'FX001' order by U_Minor", "", False, False)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SClasCod").Specific, "Select U_Minor = '%', U_CdName = '전체' Union All SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'FX001' order by U_Minor", "", False, False)
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("DepCode").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'FX003' order by U_Minor", "", False, False)
'    Call oForm01.Items("DepCode").Specific.Select(0, psk_ByValue)
    
    'Call oForm01.Items("ClasCode").Specific.Select("%", psk_ByValue)
    Call oForm01.Items("SClasCod").Specific.Select("%", psk_ByValue)
    
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("Sect").Specific, "Select U_Minor = '%', U_CdName = '선택' Union All SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q010' order by U_Minor", "", False, False)
'    Call oForm01.Items("Sect").Specific.Select("%", psk_ByValue)
    
    sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'FX001' order by U_Minor"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("ClasCode").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_FX005_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX005_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX005_CF_ChooseFromList()
On Error GoTo PS_FX005_CF_ChooseFromList_Error

    Exit Sub
PS_FX005_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX005_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX005_FormItemEnabled()
On Error GoTo PS_FX005_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_FX005_FormClear '//UDO방식
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
PS_FX005_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX005_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX005_EnableMenus()
On Error GoTo PS_FX005_EnableMenus_Error
    
    Exit Sub
PS_FX005_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX005_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX005_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_FX005_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_FX005_FormItemEnabled
        '//Call PS_FX005_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_FX005_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_FX005_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX005_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX005_FormResize()
On Error GoTo PS_FX005_FormResize_Error
    
    Exit Sub
PS_FX005_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX005_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub FormReset()
On Error GoTo FormReset_Error

Dim sQry As String
Dim RecordSet01 As SAPbobsCOM.Recordset
Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
   
'
'
'
''Call oForm01.Items("ClasCode").Specific.Select("", psk_ByValue)
'Call oForm01.Items("SClasCod").Specific.Select("%", psk_ByValue)
'oForm01.Items("FixCode").Specific.VALUE = ""
'oForm01.Items("FixName").Specific.VALUE = ""
'
'oForm01.Items("CardCode").Specific.VALUE = ""
'oForm01.Items("CardName").Specific.VALUE = ""
'
'oForm01.Items("LongYear").Specific.VALUE = 0
'oForm01.Items("DepRate").Specific.VALUE = 0
'
'oForm01.Items("TeamCode").Specific.VALUE = ""
'oForm01.Items("TeamNm").Specific.VALUE = ""
'oForm01.Items("RspCode").Specific.VALUE = ""
'oForm01.Items("RspNm").Specific.VALUE = ""
'
'oForm01.Items("PostQty").Specific.VALUE = 0
'oForm01.Items("PostAmt").Specific.VALUE = 0
'oForm01.Items("PrcCode").Specific.VALUE = ""
'oForm01.Items("PrcName").Specific.VALUE = ""
'oForm01.Items("Place").Specific.VALUE = ""
'oForm01.Items("TempChr1").Specific.VALUE = ""
'oForm01.Items("Comments").Specific.VALUE = ""

sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_FX005H]"
RecordSet01.DoQuery sQry
If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
    oForm01.Items("DocEntry").Specific.VALUE = 1
Else
    oForm01.Items("DocEntry").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE) + 1
End If
Call oDS_PS_FX005H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId)
Call oDS_PS_FX005H.setValue("U_ClasCode", 0, "")
Call oDS_PS_FX005H.setValue("U_SubDiv", 0, "N") '서브자산여부추가(2017.09.05 송명규) (N:메인자산, Y:서브자산)
Call oDS_PS_FX005H.setValue("U_FixCode", 0, "")
Call oDS_PS_FX005H.setValue("U_FixName", 0, "")
Call oDS_PS_FX005H.setValue("U_PostDate", 0, Format(Date, "YYYYMMDD"))
Call oDS_PS_FX005H.setValue("U_PostQty", 0, 0)
Call oDS_PS_FX005H.setValue("U_PostAmt", 0, 0)
Call oDS_PS_FX005H.setValue("U_StopYN", 0, "N")
Call oDS_PS_FX005H.setValue("U_LongYear", 0, 0)

Call oDS_PS_FX005H.setValue("U_DepRate", 0, 0)
Call oDS_PS_FX005H.setValue("U_CardCode", 0, "")
Call oDS_PS_FX005H.setValue("U_CardName", 0, "")
Call oDS_PS_FX005H.setValue("U_Kw", 0, 0)
Call oDS_PS_FX005H.setValue("U_Volt", 0, 0)
Call oDS_PS_FX005H.setValue("U_TeamCode", 0, "")
Call oDS_PS_FX005H.setValue("U_TeamNm", 0, "")
Call oDS_PS_FX005H.setValue("U_RspCode", 0, "")
Call oDS_PS_FX005H.setValue("U_RspNm", 0, "")
Call oDS_PS_FX005H.setValue("U_Place", 0, "")
Call oDS_PS_FX005H.setValue("U_Comments", 0, "")
Call oDS_PS_FX005H.setValue("U_TempChr1", 0, "")
Call oDS_PS_FX005H.setValue("U_PrcCode", 0, "")
Call oDS_PS_FX005H.setValue("U_PrcName", 0, "")


oForm01.Mode = fm_ADD_MODE

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormReset_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


'****************************************************************************************************************
'// Grid 꾸며주기
'****************************************************************************************************************
Private Sub GridSetting()
On Error GoTo GridSetting_Error
'    Dim i           As Integer
'    Dim sColsTitle  As String
'    Dim sColsLine   As String
'
'    oForm01.Freeze True
'
'    oGrid01.SelectionMode = ms_Single
'
''    oGrid01.Columns.Item(0).LinkedObjectType = lf_BusinessPartner
''    oGrid01.Columns.Item(4).LinkedObjectType = lf_Invoice
'     oGrid01.Columns.Item(2).LinkedObjectType = lf_Items
'
'    For i = 0 To oGrid01.Columns.Count - 1
'        sColsTitle = oGrid01.Columns(i).TitleObject.Caption
'
'        oGrid01.Columns(i).Editable = False
'
'        If oGrid01.DataTable.Columns(i).Type = ft_Float Then
'            oGrid01.Columns(i).RightJustified = True
'        End If
'
''        If InStr(1, sColsTitle, "T(mm)") > 0 Or _
''           InStr(1, sColsTitle, "W(mm)") > 0 Or _
''           InStr(1, sColsTitle, "L(mm)") > 0 Or _
''           InStr(1, sColsTitle, "Weight") > 0 Or _
''           InStr(1, sColsTitle, "Qty(Kg)") > 0 Then
''            oGrid01.Columns(i).RightJustified = True
''            oGrid01.Columns(i).BackColor = &HE0E0E0
''        End If
''
''        If InStr(1, sColsTitle, "수량") > 0 Or _
''           InStr(1, sColsTitle, "중량") > 0 Then
''            oGrid01.Columns(i).RightJustified = True
''            oGrid01.Columns(i).BackColor = &HFFC0C0
''        End If
'    Next i
'
'    oForm01.Freeze False
    
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
GridSetting_Error:
    MDC_Com.MDC_GF_Message "GridSetting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Public Sub PS_FX005_LoadPic(ByVal pPictureControlName As String) ' 고정자산 이미지 등록
On Error GoTo PS_FX005_LoadPic_Error

    Dim i           As Integer
    Dim sFilePath   As String
    Dim sFileName   As String
    Dim sQry        As String
    Dim SaveFolders As String
    Dim objFile     As Object
    Dim FSO         As FileSystemObject

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set FSO = New FileSystemObject
    
    '//
    If Trim(oDS_PS_FX005H.GetValue("U_FixCode", 0)) = "" Then Exit Sub
    
    '//
    SaveFolders = "\\191.1.1.220\Asset_Pic"
    
    'If Dir(SaveFolders, vbDirectory) = "" Then MkDir SaveFolders
    
    '//사진 불러오기
    sFilePath = FileListBoxForm.OpenDialog(FileListBoxForm, "*.BMP|*.BMP", "파일선택", "C:\")
    If sFilePath = "" Then Exit Sub
    sFilePath = Replace(sFilePath, Chr$(0), "")
    sFileName = Trim$(Mid(sFilePath, InStr(sFilePath, "\") + 1, Len(sFilePath) - 1))
    If UCase(Right(sFileName, 3)) <> "BMP" Then
        Sbo_Application.SetStatusBarMessage "BMP 확장자만 가능합니다.", bmt_Short, True
        Exit Sub
    End If
    
    Dim imageFileName As String
    
    If pPictureControlName = "Pic01" Then
        imageFileName = "_01.BMP"
    ElseIf pPictureControlName = "Pic02" Then
        imageFileName = "_02.BMP"
    ElseIf pPictureControlName = "Pic03" Then
        imageFileName = "_03.BMP"
    ElseIf pPictureControlName = "Pic04" Then
        imageFileName = "_04.BMP"
    ElseIf pPictureControlName = "Pic05" Then
        imageFileName = "_05.BMP"
    ElseIf pPictureControlName = "Pic06" Then
        imageFileName = "_06.BMP"
    End If
    
    '//서버에 기존 파일 체크
    For Each objFile In FSO.GetFolder(SaveFolders).Files
        If InStr(Trim$(UCase(objFile.Name)), UCase$(Trim(oDS_PS_FX005H.GetValue("U_FixCode", 0)) & "_" & Trim(oDS_PS_FX005H.GetValue("U_SubCode", 0)) & imageFileName)) > 0 Then
            If Sbo_Application.MessageBox("현재 자산의 사진 파일이 있습니다. 교체하시겠습니까?", 2, "Yes", "No") = 1 Then
                '//기존 파일 삭제
                Kill SaveFolders & "\" & objFile.Name
            Else
                Exit Sub
            End If
        End If
    Next objFile
    
    Call FileCopy(sFilePath, SaveFolders & "\" & Trim(oDS_PS_FX005H.GetValue("U_FixCode", 0)) & "_" & Trim(oDS_PS_FX005H.GetValue("U_SubCode", 0)) & imageFileName)
    
    sQry = "      EXEC [PS_FX005_03] '"
    sQry = sQry & pPictureControlName & "','"
    sQry = sQry & Trim(oDS_PS_FX005H.GetValue("U_FixCode", 0)) & "', '"
    sQry = sQry & Trim(oDS_PS_FX005H.GetValue("U_SubCode", 0)) & "'"
    Call oRecordSet01.DoQuery(sQry)
    
    Call Sbo_Application.MessageBox("사진이 업로드 되었습니다.")
    If oForm01.Mode = fm_OK_MODE Then
        oForm01.Mode = fm_UPDATE_MODE
    End If
    
    Set FSO = Nothing
    
    Exit Sub
    
PS_FX005_LoadPic_Error:
    Set FSO = Nothing
    Sbo_Application.SetStatusBarMessage "PS_FX005_LoadPic_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
