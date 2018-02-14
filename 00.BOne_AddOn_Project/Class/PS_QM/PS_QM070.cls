VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM070"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_QM070.cls
'//  Module         : QM
'//  Desc           : 불합리적출등록
'//  FormType       : PS_QM070
'//  Create Date    : 2011.03.30 N.G.Y
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM070H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_QM070L As SAPbouiCOM.DBDataSource '등록라인

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
    
    Dim sBPLId$, DocDateF$, DocDateT$, SProYN$, SPDeptCd$
    
    sBPLId = Trim(oForm01.Items("SBPLId").Specific.VALUE)
    DocDateF = Trim(oForm01.Items("DocDateF").Specific.VALUE)
    DocDateT = Trim(oForm01.Items("DocDateT").Specific.VALUE)
    SProYN = Trim(oForm01.Items("SProYN").Specific.VALUE)
    SPDeptCd = Trim(oForm01.Items("SPDeptCd").Specific.VALUE)
    
    
 
    If DocDateF = "" Then DocDateF = "19000101"
    If DocDateT = "" Then DocDateT = "20991231"
    
    sQry = "EXEC [PS_QM070_01] '" & sBPLId & "','" & DocDateF & "','" & DocDateT & "','" & SProYN & "','" & SPDeptCd & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_QM070L.Clear
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
        If i + 1 > oDS_PS_QM070L.Size Then
            oDS_PS_QM070L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_QM070L.Offset = i
        
        oDS_PS_QM070L.setValue "U_LineNum", i, i + 1
        oDS_PS_QM070L.setValue "U_ColNum01", i, Trim(oRecordSet01.Fields(0).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields(1).VALUE) '//사업장
        oDS_PS_QM070L.setValue "U_ColDt01", i, Format(Trim(oRecordSet01.Fields(2).VALUE), "YYYYMMDD")
        oDS_PS_QM070L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields(3).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields(4).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields(5).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields(6).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields(7).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg07", i, Trim(oRecordSet01.Fields(8).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg08", i, Trim(oRecordSet01.Fields(9).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg09", i, Trim(oRecordSet01.Fields(10).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg10", i, Trim(oRecordSet01.Fields(11).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg11", i, Trim(oRecordSet01.Fields(12).VALUE)
        oDS_PS_QM070L.setValue "U_ColDt02", i, Format(Trim(oRecordSet01.Fields(13).VALUE), "YYYYMMDD")
        oDS_PS_QM070L.setValue "U_ColReg12", i, Trim(oRecordSet01.Fields(14).VALUE)
        oDS_PS_QM070L.setValue "U_ColReg13", i, Trim(oRecordSet01.Fields(15).VALUE)
        oDS_PS_QM070L.setValue "U_ColSum01", i, Trim(oRecordSet01.Fields(16).VALUE)
        
        
        
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
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, ItemCode$
    Dim Check As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        sQry = "Select Count(*) From [@PS_QM070H] where DocEntry = '" & DocEntry & "'"
        oRecordSet01.DoQuery sQry
        
        If (oRecordSet01.RecordCount = 0) Then
            MDC_Com.MDC_GF_Message "삭제대상이 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
            Set oRecordSet01 = Nothing
            oForm01.Mode = fm_ADD_MODE
        Else
            sQry = "Delete From [@PS_QM070H] where DocEntry = '" & DocEntry & "'"
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
        oDS_PS_QM070L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_QM070L.Offset = oRow
    oDS_PS_QM070L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'// 데이타 Update
Public Function Updatedata(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Updatedata_Error
    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, BPLId$, DocType$, JukCode$, JukName$, TypeCode$, TypeName$, DeptCd$, DeptNm$, UnitName$, NGText$, CntcCode$, CntcName$
    Dim DocDate$, PDeptCd$, PDeptNm$, ProYN$, ProDate$, ProText$, ProCode$, ProName$, CR$, MA$, MI$, OB$
    Dim Amt As Currency
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    DocType = Trim(oForm01.Items("DocType").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    'JukCode = Trim(oForm01.Items("JukCode").Specific.VALUE)
    'JukName = Trim(oForm01.Items("JukName").Specific.VALUE)
    'TypeCode = Trim(oForm01.Items("TypeCode").Specific.VALUE)
    'TypeName = Trim(oForm01.Items("TypeName").Specific.VALUE)
    DeptCd = Trim(oForm01.Items("DeptCd").Specific.VALUE)
'    DeptNm = Trim(oForm01.Items("DeptNm").Specific.VALUE)
    UnitName = Trim(oForm01.Items("UnitName").Specific.VALUE)
    NGText = Trim(oForm01.Items("NGText").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE)
    PDeptCd = Trim(oForm01.Items("PDeptCd").Specific.VALUE)
'    PDeptNm = Trim(oForm01.Items("PDeptNm").Specific.VALUE)
    ProYN = Trim(oForm01.Items("ProYN").Specific.VALUE)
    ProDate = Trim(oForm01.Items("ProDate").Specific.VALUE)
    ProText = Trim(oForm01.Items("ProText").Specific.VALUE)
    Amt = Trim(oForm01.Items("Amt").Specific.VALUE)
    ProCode = Trim(oForm01.Items("ProCode").Specific.VALUE)
    ProName = Trim(oForm01.Items("ProName").Specific.VALUE)
    CR = Trim(oForm01.Items("CR").Specific.VALUE)
    MA = Trim(oForm01.Items("MA").Specific.VALUE)
    MI = Trim(oForm01.Items("MI").Specific.VALUE)
    OB = Trim(oForm01.Items("OB").Specific.VALUE)
    
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!", "E"
        Updatedata = False
        Exit Function
    End If
        
    
    sQry = "Update [@PS_QM070H]"
    sQry = sQry & " set "
    sQry = sQry & " U_BPLId = '" & BPLId & "',"
    sQry = sQry & " U_DocType = '" & DocType & "',"
    sQry = sQry & " U_DocDate = '" & DocDate & "',"
    'sQry = sQry & " U_JukCode  = '" & JukCode & "',"
    'sQry = sQry & " U_JukName  = '" & JukName & "',"
    'sQry = sQry & " U_TypeCode  = '" & TypeCode & "',"
    'sQry = sQry & " U_TypeName  = '" & TypeName & "',"
    sQry = sQry & " U_DeptCd  = '" & DeptCd & "',"
    sQry = sQry & " U_DeptNm  = '" & DeptNm & "',"
    sQry = sQry & " U_UnitName  = '" & UnitName & "',"
    sQry = sQry & " U_NGText  = '" & NGText & "',"
    sQry = sQry & " U_CntcCode  = '" & CntcCode & "',"
    sQry = sQry & " U_CntcName  = '" & CntcName & "',"
    sQry = sQry & " U_PDeptCd  = '" & PDeptCd & "',"
    sQry = sQry & " U_PDeptNm  = '" & PDeptNm & "',"
    sQry = sQry & " U_ProYN  = '" & ProYN & "',"
    sQry = sQry & " U_ProDate  = '" & ProDate & "',"
    sQry = sQry & " U_ProText  = '" & ProText & "',"
    sQry = sQry & " U_Amt = '" & Amt & "',"
    sQry = sQry & " U_CR = '" & CR & "',"
    sQry = sQry & " U_MA = '" & MA & "',"
    sQry = sQry & " U_MI = '" & MI & "',"
    sQry = sQry & " U_OB = '" & OB & "',"
    sQry = sQry & " U_ProCode = '" & ProCode & "',"
    sQry = sQry & " U_ProName = '" & ProName & "'"
            
    sQry = sQry & " Where DocEntry = '" & DocEntry & "'"
    RecordSet01.DoQuery sQry
    
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    Updatedata = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Updatedata_Error:
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
    
    Dim DocEntry$, BPLId$, DocType$, JukCode$, JukName$, TypeCode$, TypeName$, DeptCd$, DeptNm$, UnitName$, NGText$, CntcCode$, CntcName$
    Dim DocDate$, PDeptCd$, PDeptNm$, ProYN$, ProDate$, ProText$, ProCode$, ProName$, CR$, MA$, MI$, OB$
    Dim Amt As Currency
    Dim LineNum$
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    DocType = Trim(oForm01.Items("DocType").Specific.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    'JukCode = Trim(oForm01.Items("JukCode").Specific.VALUE)
    'JukName = Trim(oForm01.Items("JukName").Specific.VALUE)
    'TypeCode = Trim(oForm01.Items("TypeCode").Specific.VALUE)
    'TypeName = Trim(oForm01.Items("TypeName").Specific.VALUE)
    DeptCd = Trim(oForm01.Items("DeptCd").Specific.VALUE)
    'DeptNm = Trim(oForm01.Items("DeptNm").Specific.VALUE)
    UnitName = Trim(oForm01.Items("UnitName").Specific.VALUE)
    NGText = Trim(oForm01.Items("NGText").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE)
    PDeptCd = Trim(oForm01.Items("PDeptCd").Specific.VALUE)
    'PDeptNm = Trim(oForm01.Items("PDeptNm").Specific.VALUE)
    ProYN = Trim(oForm01.Items("ProYN").Specific.VALUE)
    ProDate = Trim(oForm01.Items("ProDate").Specific.VALUE)
    ProText = Trim(oForm01.Items("ProText").Specific.VALUE)
    Amt = Trim(oForm01.Items("Amt").Specific.VALUE)
    ProCode = Trim(oForm01.Items("ProCode").Specific.VALUE)
    ProName = Trim(oForm01.Items("ProName").Specific.VALUE)
    CR = Trim(oForm01.Items("CR").Specific.VALUE)
    MA = Trim(oForm01.Items("MA").Specific.VALUE)
    MI = Trim(oForm01.Items("MI").Specific.VALUE)
    OB = Trim(oForm01.Items("OB").Specific.VALUE)
    
    sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM070H]"
    RecordSet01.DoQuery sQry
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "Select IsNull(Max(U_LineNum), 0) From [@PS_QM070H] Where U_BPLId = '" & BPLId & "' And U_DocDate = '" & DocDate & "'"
    RecordSet01.DoQuery sQry
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        LineNum = 1
    Else
        LineNum = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    
    sQry = "INSERT INTO [@PS_QM070H]"
    sQry = sQry & " ("
    sQry = sQry & " DocEntry,"
    sQry = sQry & " DocNum,"
    sQry = sQry & " U_BPLId,"
    sQry = sQry & " U_DocType,"
    sQry = sQry & " U_DocDate,"
    sQry = sQry & " U_LineNum,"
    sQry = sQry & " U_JukCode,"
    sQry = sQry & " U_JukName,"
    sQry = sQry & " U_TypeCode,"
    sQry = sQry & " U_TypeName,"
    sQry = sQry & " U_DeptCd,"
'    sQry = sQry & " U_DeptNm,"
    sQry = sQry & " U_UnitName,"
    sQry = sQry & " U_NGText,"
    sQry = sQry & " U_CntcCode,"
    sQry = sQry & " U_CntcName,"
    sQry = sQry & " U_PDeptCd,"
'    sQry = sQry & " U_PDeptNm,"
    sQry = sQry & " U_ProYN,"
    sQry = sQry & " U_ProDate,"
    sQry = sQry & " U_ProText,"
    sQry = sQry & " U_Amt,"
    sQry = sQry & " U_ProCode,"
    sQry = sQry & " U_ProName,"
    sQry = sQry & " U_CR,"
    sQry = sQry & " U_MA,"
    sQry = sQry & " U_MI,"
    sQry = sQry & " U_OB"
    sQry = sQry & " ) "
    sQry = sQry & "VALUES("
    sQry = sQry & DocEntry & ","
    sQry = sQry & DocEntry & ","
    sQry = sQry & "'" & BPLId & "',"
    sQry = sQry & "'" & DocType & "',"
    sQry = sQry & "'" & DocDate & "',"
    sQry = sQry & "'" & LineNum & "',"
    sQry = sQry & "'" & JukCode & "',"
    sQry = sQry & "'" & JukName & "',"
    sQry = sQry & "'" & TypeCode & "',"
    sQry = sQry & "'" & TypeName & "',"
    sQry = sQry & "'" & DeptCd & "',"
 '   sQry = sQry & "'" & DeptNm & "',"
    sQry = sQry & "'" & UnitName & "',"
    sQry = sQry & "'" & NGText & "',"
    sQry = sQry & "'" & CntcCode & "',"
    sQry = sQry & "'" & CntcName & "',"
    sQry = sQry & "'" & PDeptCd & "',"
 '   sQry = sQry & "'" & PDeptNm & "',"
    sQry = sQry & "'" & ProYN & "',"
    sQry = sQry & "'" & ProDate & "',"
    sQry = sQry & "'" & ProText & "',"
    sQry = sQry & "'" & Amt & "',"
    sQry = sQry & "'" & ProCode & "',"
    sQry = sQry & "'" & ProName & "',"
    sQry = sQry & "'" & CR & "',"
    sQry = sQry & "'" & MA & "',"
    sQry = sQry & "'" & MI & "',"
    sQry = sQry & "'" & OB & "'"
    sQry = sQry & ")"
    RecordSet01.DoQuery sQry
    
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
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
        Case Trim(oForm01.Items("DocDate").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("UnitName").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "입력일자는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "적출장소(부위)는 필수사항입니다. 확인하세요.", "E"
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
        Case "CntcCode"  '/사번
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "ProCode"  '/사번
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("ProCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("ProName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
            
            
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
        'oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        'oForm01.Items("ItemCode").Enabled = True
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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM070.srf")
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
    
    oFormUniqueID01 = "PS_QM070_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM070_CreateItems
    Call PS_QM070_ComboBox_Setting
    Call PS_QM070_CF_ChooseFromList
    Call PS_QM070_EnableMenus
    Call PS_QM070_SetDocument(oFromDocEntry01)
    Call PS_QM070_FormResize
    
    
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
        
    sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM070H]"
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

     oForm01.Items("DocDate").Click ct_Regular
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
                 
                ElseIf pval.ItemUID = "ProCode" Then
                    If oForm01.Items("ProCode").Specific.VALUE = "" Then
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
                    If pval.ItemUID = "CntcCode" Or pval.ItemUID = "ProCode" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    
                    If pval.ItemUID = "DocDate" Then
                    
                        Dim sQry As String
                        Dim DocDate As String
                        Dim RecordSet01 As SAPbobsCOM.Recordset
                        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                        
                        DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
                        
                        sQry = "Select left('" & DocDate & "',6) + '01', Convert(char(8),Dateadd(dd, -1, left(convert(char(8),Dateadd(mm, 1, '" & DocDate & "'),112), 6) + '01'),112)"
                        RecordSet01.DoQuery sQry
                        oForm01.Items("DocDateF").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                        oForm01.Items("DocDateT").Specific.VALUE = Trim(RecordSet01.Fields(1).VALUE)
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
                
                oForm01.Items("ItemCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_QM070_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_QM070H.Clear

'                Call LoadCaption
'                Call FormItemEnabled
                '//Call PS_QM070_FormItemEnabled '//UDO방식
                '//Call PS_QM070_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_QM070_FormItemEnabled
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
        If pval.ItemUID = "PS_QM070" Then
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
                If Updatedata(pval) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call FormReset
                oForm01.Mode = fm_ADD_MODE
                Call LoadCaption
                Call LoadData
                
                
                oForm01.Items("DocDate").Click ct_Regular
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
        ElseIf pval.ItemUID = "Btn_ex" Then '/첨부파일열기
'            If oForm01.Mode = fm_UPDATE_MODE Then
            
            
            




'                Dim strFile As String
'                strFile = "\\191.1.1.220\불합리적출\" & oForm01.Items("DocDate").Specific.VALUE & ".ppt"
'               ' ShellExecute Me.hWnd, vbNullString, """" & strFile & """", vbNullString, vbNullString, 3
'                ShellExecute 0, vbNullString, """" & strFile & """", vbNullString, vbNullString, 3
'            End If
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM070" Then
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
Dim sQry As String
Dim oRecordSet01 As SAPbobsCOM.Recordset
Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
Dim DocEntry$
    

    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)
                DocEntry = oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE
                
                sQry = "EXEC [PS_QM070_02] '" & DocEntry & "'"
                oRecordSet01.DoQuery sQry
                
                If (oRecordSet01.RecordCount = 0) Then
                    MDC_Com.MDC_GF_Message "결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                    Set oRecordSet01 = Nothing
                    oForm01.Mode = fm_ADD_MODE
                    
                    Call LoadCaption
                    Call FormReset
                    Exit Sub
                End If
                
'                oDS_PS_QM070L.setValue "U_ColNum01", i, Trim(oRecordSet01.Fields(0).VALUE)
                oDS_PS_QM070H.setValue "DocEntry", 0, Trim(oRecordSet01.Fields(0).VALUE)
                oDS_PS_QM070H.setValue "U_BPLId", 0, Trim(oRecordSet01.Fields(1).VALUE)
                oDS_PS_QM070H.setValue "U_DocType", 0, Trim(oRecordSet01.Fields(2).VALUE)
                oDS_PS_QM070H.setValue "U_DocDate", 0, Format(Trim(oRecordSet01.Fields(3).VALUE), "YYYYMMDD")
                oDS_PS_QM070H.setValue "U_LineNum", 0, Trim(oRecordSet01.Fields(4).VALUE)
                oDS_PS_QM070H.setValue "U_CR", 0, Trim(oRecordSet01.Fields(5).VALUE)
                oDS_PS_QM070H.setValue "U_MA", 0, Trim(oRecordSet01.Fields(6).VALUE)
                oDS_PS_QM070H.setValue "U_MI", 0, Trim(oRecordSet01.Fields(7).VALUE)
                oDS_PS_QM070H.setValue "U_OB", 0, Trim(oRecordSet01.Fields(8).VALUE)
                oDS_PS_QM070H.setValue "U_DeptCd", 0, Trim(oRecordSet01.Fields(9).VALUE)
                oDS_PS_QM070H.setValue "U_UnitName", 0, Trim(oRecordSet01.Fields(10).VALUE)
                oDS_PS_QM070H.setValue "U_NGText", 0, Trim(oRecordSet01.Fields(11).VALUE)
                oDS_PS_QM070H.setValue "U_CntcCode", 0, Trim(oRecordSet01.Fields(12).VALUE)
                oDS_PS_QM070H.setValue "U_CntcName", 0, Trim(oRecordSet01.Fields(13).VALUE)
                oDS_PS_QM070H.setValue "U_PDeptCd", 0, Trim(oRecordSet01.Fields(14).VALUE)
                oDS_PS_QM070H.setValue "U_ProYN", 0, Trim(oRecordSet01.Fields(15).VALUE)
                oDS_PS_QM070H.setValue "U_ProDate", 0, Format(Trim(oRecordSet01.Fields(16).VALUE), "YYYYMMDD")
                oDS_PS_QM070H.setValue "U_ProText", 0, Trim(oRecordSet01.Fields(17).VALUE)
                oDS_PS_QM070H.setValue "U_ProCode", 0, Trim(oRecordSet01.Fields(18).VALUE)
                oDS_PS_QM070H.setValue "U_ProName", 0, Trim(oRecordSet01.Fields(19).VALUE)
                oDS_PS_QM070H.setValue "U_Amt", 0, Trim(oRecordSet01.Fields(10).VALUE)
'                oForm01.Items("DocEntry").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'                oForm01.Items("BPLId").Specific.VALUE = Trim(oRecordSet01.Fields(1).VALUE)
'                oForm01.Items("DocType").Specific.VALUE = Trim(oRecordSet01.Fields(2).VALUE)
'                oForm01.Items("DocDate").Specific.VALUE = Trim(oRecordSet01.Fields(3).VALUE)
'                oForm01.Items("CR").Specific.VALUE = Trim(oRecordSet01.Fields(4).VALUE)
'                oForm01.Items("MA").Specific.VALUE = Trim(oRecordSet01.Fields(5).VALUE)
'                oForm01.Items("MI").Specific.VALUE = Trim(oRecordSet01.Fields(6).VALUE)
'                oForm01.Items("OB").Specific.VALUE = Trim(oRecordSet01.Fields(7).VALUE)
'
                
'                oForm01.Items("DeptNm").Specific.VALUE = oMat01.Columns("DeptNm").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("UnitName").Specific.VALUE = oMat01.Columns("UnitName").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("NGText").Specific.VALUE = oMat01.Columns("NGText").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("CntcName").Specific.VALUE = oMat01.Columns("CntcName").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("PDeptNm").Specific.VALUE = oMat01.Columns("PDeptNm").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("ProYN").Specific.VALUE = oMat01.Columns("ProYn").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("ProDate").Specific.VALUE = oMat01.Columns("ProDate").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("ProText").Specific.VALUE = oMat01.Columns("ProText").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("Amt").Specific.VALUE = oMat01.Columns("Amt").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("ProName").Specific.VALUE = oMat01.Columns("ProName").Cells(pval.Row).Specific.VALUE
'
'                oForm01.Items("JukCode").Specific.VALUE = oMat01.Columns("JukName").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("JukName").Specific.VALUE = oMat01.Columns("JukName").Cells(pval.Row).Specific.VALUE
'                oForm01.Items("TypeName").Specific.VALUE = oMat01.Columns("TypeName").Cells(pval.Row).Specific.VALUE
                
                oForm01.Mode = fm_UPDATE_MODE
                Call LoadCaption
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

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
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
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
'        If pval.ItemChanged = True Then
'            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_QM070L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM070L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_QM070_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_QM070L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                End If
'            Else
'                If (pval.ItemUID = "DocEntry") Then
'                    Call oDS_PS_QM070H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_QM070H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_QM070H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_QM070H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                End If
'            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
'        End If
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
        Call PS_QM070_FormItemEnabled
        '//Call PS_QM070_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_QM070_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_QM070H", "U_CardCode,U_CardName")
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
'            If (PS_QM070_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_QM070H.RemoveRecord(oDS_PS_QM070H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_QM070H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_QM070_CreateItems() As Boolean
On Error GoTo PS_QM070_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
         
    
    
    Set oDS_PS_QM070H = oForm01.DataSources.DBDataSources("@PS_QM070H")
    Set oDS_PS_QM070L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '/유저데이타 속성 선언 날짜형식
    Call oForm01.DataSources.UserDataSources.Add("DocDateF", dt_DATE, 8)
    Call oForm01.DataSources.UserDataSources.Add("DocDateT", dt_DATE, 8)
    Call oForm01.Items("DocDateF").Specific.DataBind.SetBound(True, "", "DocDateF")
    Call oForm01.Items("DocDateT").Specific.DataBind.SetBound(True, "", "DocDateT")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_QM070_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM070_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
'/콤보박스 set
Sub PS_QM070_ComboBox_Setting()
On Error GoTo PS_QM070_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
        
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    Set oCombo = oForm01.Items("DocType").Specific
    oCombo.ValidValues.Add "10", "당직"
    oCombo.ValidValues.Add "20", "일반"
    oCombo.Select "0", psk_Index
    
    Set oCombo = oForm01.Items("ProYN").Specific
    oCombo.ValidValues.Add "N", "미완료"
    oCombo.ValidValues.Add "Y", "완료"
    oCombo.Select "0", psk_Index
    
    Set oCombo = oForm01.Items("SProYN").Specific
    oCombo.ValidValues.Add "N", "미완료"
    oCombo.ValidValues.Add "Y", "완료"
    oCombo.Select "0", psk_Index
    
    Set oCombo = oForm01.Items("DeptCd").Specific
    sQry = "Select PrcCode, PrcName From [OPRC] Where left(GrpCode,1) in ('1', '2') Order by PrcCode"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("DeptCd").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    Set oCombo = oForm01.Items("PDeptCd").Specific
    sQry = "Select PrcCode, PrcName From [OPRC] Where left(GrpCode,1) in ('1', '2') Order by PrcCode"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("PDeptCd").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    Set oCombo = oForm01.Items("SPDeptCd").Specific
    sQry = "Select PrcCode, PrcName From [OPRC] Where left(GrpCode,1) in ('1', '2') Order by PrcCode"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//적출코드
    oForm01.Items("CR").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CR").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q002' order by U_Minor", "", False, False)
    oForm01.Items("MA").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("MA").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q003' order by U_Minor", "", False, False)
    oForm01.Items("MI").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("MI").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q004' order by U_Minor", "", False, False)
    oForm01.Items("OB").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OB").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q005' order by U_Minor", "", False, False)
    
    
    sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'Q002' Order By U_Minor "
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("CR").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'Q003' Order By U_Minor "
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("MA").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'Q004' Order By U_Minor "
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("MI").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'Q005' Order By U_Minor "
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("OB").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    
    
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_QM070_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM070_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM070_CF_ChooseFromList()
On Error GoTo PS_QM070_CF_ChooseFromList_Error

    Exit Sub
PS_QM070_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM070_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM070_FormItemEnabled()
On Error GoTo PS_QM070_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_QM070_FormClear '//UDO방식
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
PS_QM070_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM070_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM070_EnableMenus()
On Error GoTo PS_QM070_EnableMenus_Error
    
    Exit Sub
PS_QM070_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM070_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM070_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM070_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM070_FormItemEnabled
        '//Call PS_QM070_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_QM070_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_QM070_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM070_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM070_FormResize()
On Error GoTo PS_QM070_FormResize_Error
    
    Exit Sub
PS_QM070_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM070_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub FormReset()
On Error GoTo FormReset_Error

Dim sQry As String
Dim RecordSet01 As SAPbobsCOM.Recordset
Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)



oDS_PS_QM070H.setValue "U_CR", 0, ""
oDS_PS_QM070H.setValue "U_MA", 0, ""
oDS_PS_QM070H.setValue "U_MI", 0, ""
oDS_PS_QM070H.setValue "U_OB", 0, ""
oDS_PS_QM070H.setValue "U_DeptCd", 0, ""
oDS_PS_QM070H.setValue "U_UnitName", 0, ""
oDS_PS_QM070H.setValue "U_NGText", 0, ""
oDS_PS_QM070H.setValue "U_CntcCode", 0, ""
oDS_PS_QM070H.setValue "U_CntcName", 0, ""
oDS_PS_QM070H.setValue "U_PDeptCd", 0, ""
oDS_PS_QM070H.setValue "U_ProYN", 0, "N"
oDS_PS_QM070H.setValue "U_ProDate", 0, ""
oDS_PS_QM070H.setValue "U_ProText", 0, ""
oDS_PS_QM070H.setValue "U_ProCode", 0, ""
oDS_PS_QM070H.setValue "U_ProName", 0, ""
oDS_PS_QM070H.setValue "U_Amt", 0, 0

        
sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM070H]"
RecordSet01.DoQuery sQry
If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
    oForm01.Items("DocEntry").Specific.VALUE = 1
Else
    oForm01.Items("DocEntry").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE) + 1
End If

oForm01.Mode = fm_ADD_MODE

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormReset_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


