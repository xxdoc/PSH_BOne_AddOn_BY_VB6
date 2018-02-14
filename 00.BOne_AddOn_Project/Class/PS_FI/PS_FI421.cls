VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FI421"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_FI421.cls
'//  Module : 재무관리 > 애드온 > 입금표
'//  Desc : 입금표등록
'//  FormType : PS_FI421
'//  Create Date(Start) : 2013.04.23
'//  Create Date(End) :
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_FI421H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_FI421L As SAPbouiCOM.DBDataSource '등록라인

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
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FI421.srf")
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
    
    oFormUniqueID01 = "PS_FI421_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_FI421_CreateItems
    Call PS_FI421_ComboBox_Setting
    Call PS_FI421_CF_ChooseFromList
    Call PS_FI421_EnableMenus
    Call PS_FI421_SetDocument(oFromDocEntry01)
    Call PS_FI421_FormResize

    Call PS_FI421_Add_MatrixRow(0, True)
    Call PS_FI421_LoadCaption
    Call PS_FI421_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_FI421H]"
    Call RecordSet01.DoQuery(sQry)
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_FI421H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_FI421H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
        
    Call PS_FI421_FormReset
        
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing

    '발행일자 설정
    Call oDS_PS_FI421H.setValue("U_DocDate", 0, Format(Date, "YYYYMMDD"))
    
    '발행일자FR
    oForm01.Items("SDocDateFr").Specific.VALUE = Format(Date, "YYYYMM01")
    
    '발행일자TO
    oForm01.Items("SDocDateTo").Specific.VALUE = Format(Date, "YYYYMMDD")
    
    '사번 포커스
    oForm01.Items("CntcCode").Click ct_Regular
     
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FI421_LoadCaption()
'******************************************************************************
'Function ID : PS_FI421_LoadCaption()
'해당모듈 : PS_FI421
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FI421_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "추가"
        oForm01.Items("BtnDelete").Enabled = False
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
        oForm01.Items("BtnDelete").Enabled = True
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_FI421_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_FI421_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_FI421L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_FI421L.Offset = oRow
    oDS_PS_FI421L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_FI421_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_FI421_MTX01()
'******************************************************************************
'Function ID : PS_FI421_MTX01()
'해당모듈 : PS_FI421
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FI421_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim sDocEntry As String '관리번호
    Dim SSerialNo As String '일련번호
    Dim sBPLId As String '사업장
    Dim SRspCode As String '담당
    Dim SCntcCode As String '사번
    Dim SCardCode As String '거래처
    Dim SDocDateFr As String '발행일자
    Dim SDocDateTo As String '발행일자TO
    Dim SAmount As Currency '공급가액
    Dim SVatTax As Currency '부가가치세
    Dim SContents As String '내용
    
    sDocEntry = Trim(oForm01.Items("SDocEntry").Specific.VALUE) '관리번호
    SSerialNo = Trim(oForm01.Items("SSerialNo").Specific.VALUE) '일련번호
    sBPLId = Trim(oForm01.Items("SBPLId").Specific.VALUE) '사업장
    SRspCode = Trim(oForm01.Items("SRspCode").Specific.VALUE) '담당
    SCntcCode = Trim(oForm01.Items("SCntcCode").Specific.VALUE) '사번
    SCardCode = Trim(oForm01.Items("SCardCode").Specific.VALUE) '거래처
    SDocDateFr = Trim(oForm01.Items("SDocDateFr").Specific.VALUE) '발행일자FR
    SDocDateTo = Trim(oForm01.Items("SDocDateTo").Specific.VALUE) '발행일자TO
    SAmount = Trim(oForm01.Items("SAmount").Specific.VALUE) '공급가액
    SVatTax = Trim(oForm01.Items("SVatTax").Specific.VALUE) '부가가치세
    SContents = Trim(oForm01.Items("SContents").Specific.VALUE) '내용
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "                EXEC [PS_FI421_01]"
    sQry = sQry & "'" & sDocEntry & "',"
    sQry = sQry & "'" & SSerialNo & "',"
    sQry = sQry & "'" & sBPLId & "',"
    sQry = sQry & "'" & SRspCode & "',"
    sQry = sQry & "'" & SCntcCode & "',"
    sQry = sQry & "'" & SCardCode & "',"
    sQry = sQry & "'" & SDocDateFr & "',"
    sQry = sQry & "'" & SDocDateTo & "',"
    sQry = sQry & "'" & SAmount & "',"
    sQry = sQry & "'" & SVatTax & "',"
    sQry = sQry & "'" & SContents & "'"
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_FI421L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
        Call PS_FI421_Add_MatrixRow(0, True)
        Call PS_FI421_LoadCaption
        
        GoTo PS_FI421_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_FI421L.Size Then
            oDS_PS_FI421L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_FI421L.Offset = i
        
        Call oDS_PS_FI421L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_FI421L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PS_FI421L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("SerialNo").VALUE)) '일련번호
        Call oDS_PS_FI421L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("BPLId").VALUE)) '사업장
        Call oDS_PS_FI421L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("RspCode").VALUE)) '담당
        Call oDS_PS_FI421L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("CntcCode").VALUE)) '사번
        Call oDS_PS_FI421L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("CntcName").VALUE)) '성명
        Call oDS_PS_FI421L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("CardCode").VALUE)) '거래처
        Call oDS_PS_FI421L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("CardName").VALUE)) '거래처명
        Call oDS_PS_FI421L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("DocDate").VALUE), "YYYYMMDD")) '발행일자
        Call oDS_PS_FI421L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Amount").VALUE)) '공급가액
        Call oDS_PS_FI421L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("VatTax").VALUE)) '부가가치세
        Call oDS_PS_FI421L.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("Contents").VALUE)) '내용
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_FI421_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_FI421_DeleteData()
'******************************************************************************
'Function ID : PS_FI421_DeleteData()
'해당모듈 : PS_FI421
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FI421_DeleteData_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PS_FI421H] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_FI421_DeleteData_Error
        Else
            sQry = "DELETE FROM [@PS_FI421H] WHERE DocEntry = '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
'    Call PS_FI421_FormReset
    
'    oForm01.Mode = fm_ADD_MODE
    
'    Call oForm01.Items("BtnSearch").Click(ct_Regular)
    
'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PS_FI421_Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_FI421_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_FI421_UpdateData() As Boolean
'******************************************************************************
'Function ID : PS_FI421_UpdateData()
'해당모듈 : PS_FI421
'기능 : 기본정보를 수정
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FI421_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim SerialNo As String '일련번호
    Dim BPLId As String '사업장
    Dim RspCode As String '담당
    Dim CntcCode As String '사번
    Dim CntcName As String '성명
    Dim CardCode As String '거래처
    Dim CardName As String '거래처명
    Dim DocDate As String '발행일자
    Dim Amount As Currency '공급가액
    Dim VatTax As Currency '부가가치세
    Dim Contents As String '내용
    Dim UserSign As String 'UserSign
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE) '관리번호
    SerialNo = Trim(oForm01.Items("SerialNo").Specific.VALUE)  '일련번호
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)  '사업장
    RspCode = Trim(oForm01.Items("RspCode").Specific.VALUE)  '담당
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)  '사번
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE)  '성명
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)  '거래처
    CardName = Trim(oForm01.Items("CardName").Specific.VALUE)  '거래처명
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)  '발행일자
    Amount = Trim(oForm01.Items("Amount").Specific.VALUE)  '공급가액
    VatTax = Trim(oForm01.Items("VatTax").Specific.VALUE)  '부가가치세
    Contents = Trim(oForm01.Items("Contents").Specific.VALUE)  '내용
    UserSign = Sbo_Company.UserSignature
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
        PS_FI421_UpdateData = False
        Exit Function
    End If
    
    sQry = "                EXEC [PS_FI421_03]"
    sQry = sQry & "'" & DocEntry & "',"
    sQry = sQry & "'" & SerialNo & "',"
    sQry = sQry & "'" & BPLId & "',"
    sQry = sQry & "'" & RspCode & "',"
    sQry = sQry & "'" & CntcCode & "',"
    sQry = sQry & "'" & CntcName & "',"
    sQry = sQry & "'" & CardCode & "',"
    sQry = sQry & "'" & CardName & "',"
    sQry = sQry & "'" & DocDate & "',"
    sQry = sQry & "'" & Amount & "',"
    sQry = sQry & "'" & VatTax & "',"
    sQry = sQry & "'" & Contents & "'"
    
    Call RecordSet01.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PS_FI421_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_FI421_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_FI421_AddData() As Boolean
'******************************************************************************
'Function ID : PS_FI421_AddData()
'해당모듈 : PS_FI421
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FI421_AddData_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim SerialNo As String '일련번호
    Dim BPLId As String '사업장
    Dim RspCode As String '담당
    Dim CntcCode As String '사번
    Dim CntcName As String '성명
    Dim CardCode As String '거래처
    Dim CardName As String '거래처명
    Dim DocDate As String '발행일자
    Dim Amount As Currency '공급가액
    Dim VatTax As Currency '부가가치세
    Dim Contents As String '내용
    Dim UserSign As String 'UserSign
    
    SerialNo = Trim(oForm01.Items("SerialNo").Specific.VALUE)  '일련번호
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)  '사업장
    RspCode = Trim(oForm01.Items("RspCode").Specific.VALUE)  '담당
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)  '사번
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE)  '성명
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)  '거래처
    CardName = Trim(oForm01.Items("CardName").Specific.VALUE)  '거래처명
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)  '발행일자
    Amount = Trim(oForm01.Items("Amount").Specific.VALUE)  '공급가액
    VatTax = Trim(oForm01.Items("VatTax").Specific.VALUE)  '부가가치세
    Contents = Trim(oForm01.Items("Contents").Specific.VALUE)  '내용
    UserSign = Sbo_Company.UserSignature
    
    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_FI421H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "                EXEC [PS_FI421_02]"
    sQry = sQry & "'" & DocEntry & "',"
    sQry = sQry & "'" & SerialNo & "',"
    sQry = sQry & "'" & BPLId & "',"
    sQry = sQry & "'" & RspCode & "',"
    sQry = sQry & "'" & CntcCode & "',"
    sQry = sQry & "'" & CntcName & "',"
    sQry = sQry & "'" & CardCode & "',"
    sQry = sQry & "'" & CardName & "',"
    sQry = sQry & "'" & DocDate & "',"
    sQry = sQry & "'" & Amount & "',"
    sQry = sQry & "'" & VatTax & "',"
    sQry = sQry & "'" & Contents & "',"
    sQry = sQry & "'" & UserSign & "'"
    
    
    Call RecordSet02.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_FI421_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_AddData_Error:
    PS_FI421_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PS_FI421_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_FI421_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_FI421_HeaderSpaceLineDel()
'해당모듈 : PS_FI421
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FI421_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("RspCode").Specific.VALUE) = "%" '담당
            ErrNum = 1
            GoTo PS_FI421_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("CardCode").Specific.VALUE) = "" '거래처
            ErrNum = 2
            GoTo PS_FI421_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("DocDate").Specific.VALUE) = "%" '발행일자
            ErrNum = 3
            GoTo PS_FI421_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Amount").Specific.VALUE) = "0" '공급가액
            ErrNum = 4
            GoTo PS_FI421_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Contents").Specific.VALUE) = "" '내용
            ErrNum = 5
            GoTo PS_FI421_HeaderSpaceLineDel_Error
    End Select

    PS_FI421_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "담당은 필수사항입니다. 선택하세요.", "E"
        Call oForm01.Items("RspCode").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "거래처는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("CardCode").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "발행일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("DocDate").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "공급가액은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Amount").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "내용은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Contents").Click(ct_Regular)
    End If
    PS_FI421_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PS_FI421_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_FI421_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PS_FI421_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PS_FI421_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_FI421_MatrixSpaceLineDel = False
End Function

Private Sub PS_FI421_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_FI421_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    Dim Qty As Integer
    Dim Calculate_Weight As Double
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Select Case oUID
'
'        Case "BOSHour"
'
'    End Select

    '시각 데이터의 Validate를 체크하기 위한 구문_S
    If oUID = "BOSHour" Or oUID = "BOEHour" Then
        
        If IsNumeric(oForm01.Items(oUID).Specific.VALUE) = True Then
        
            If Val(oForm01.Items(oUID).Specific.VALUE) >= 0 And Val(oForm01.Items(oUID).Specific.VALUE) <= 9 Then
            
                Call oDS_PS_FI421H.setValue("U_" & oUID, 0, "0" & CStr(oForm01.Items(oUID).Specific.VALUE))
                
            ElseIf Val(oForm01.Items(oUID).Specific.VALUE) > 24 Then
            
                ErrNum = 2
                Call oDS_PS_FI421H.setValue("U_" & oUID, 0, "")
                GoTo PS_FI421_FlushToItemValue_Error
                
            Else
            
                Call oDS_PS_FI421H.setValue("U_" & oUID, 0, CStr(oForm01.Items(oUID).Specific.VALUE))
                
            End If
        
        Else
        
            ErrNum = 1
            Call oDS_PS_FI421H.setValue("U_" & oUID, 0, "")
            GoTo PS_FI421_FlushToItemValue_Error
        
        End If
            
    ElseIf oUID = "BOSMin" Or oUID = "BOEMin" Then
        
        If IsNumeric(oForm01.Items(oUID).Specific.VALUE) = True Then
        
            If Val(oForm01.Items(oUID).Specific.VALUE) >= 0 And Val(oForm01.Items(oUID).Specific.VALUE) <= 9 Then
            
                Call oDS_PS_FI421H.setValue("U_" & oUID, 0, "0" & CStr(oForm01.Items(oUID).Specific.VALUE))
                
            ElseIf Val(oForm01.Items(oUID).Specific.VALUE) >= 60 Then
            
                ErrNum = 3
                Call oDS_PS_FI421H.setValue("U_" & oUID, 0, "")
                GoTo PS_FI421_FlushToItemValue_Error
                
            Else
            
                Call oDS_PS_FI421H.setValue("U_" & oUID, 0, CStr(oForm01.Items(oUID).Specific.VALUE))
                
            End If
        
        Else
        
            ErrNum = 1
            Call oDS_PS_FI421H.setValue("U_" & oUID, 0, "")
            GoTo PS_FI421_FlushToItemValue_Error
        
        End If
            
    End If
    '시각 데이터의 Validate를 체크하기 위한 구문_E

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_FI421_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("시각은 숫자만 입력이 가능합니다.", "E")
    ElseIf ErrNum = 2 Then
        Call MDC_Com.MDC_GF_Message("시각(시)는 24미만의 값만 입력이 가능합니다.", "E")
    ElseIf ErrNum = 3 Then
        Call MDC_Com.MDC_GF_Message("시각(분)은 60미만의 값만 입력이 가능합니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_FI421_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_FI421_FormItemEnabled()
On Error GoTo PS_FI421_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_FI421_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/아이템 변경 이벤트
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error

    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            Call Raise_EVENT_ITEM_PRESSED(FormUID, pval, BubbleEvent)
        Case et_KEY_DOWN: '//2
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
                
                Call PS_FI421_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_FI421_LoadCaption
                
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
                '//Call PS_FI421_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_FI421H.Clear

'                Call PS_FI421_LoadCaption
'                Call PS_FI421_FormItemEnabled
                '//Call PS_FI421_FormItemEnabled '//UDO방식
                '//Call PS_FI421_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_FI421_FormItemEnabled
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
    
        If pval.ItemUID = "PS_FI421" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_FI421_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If PS_FI421_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_FI421_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_FI421_LoadCaption
                Call PS_FI421_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_FI421_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                If PS_FI421_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_FI421_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_FI421_LoadCaption
                Call PS_FI421_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
                
            Call PS_FI421_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_FI421_LoadCaption
            Call PS_FI421_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_FI421_DeleteData
                Call PS_FI421_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_FI421_LoadCaption
                Call PS_FI421_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnPrint" Then '입금표출력
        
            Call PS_FI421_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_FI421" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//사용자값활성(사번)
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '//사용자값활성(거래처)
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "SCntcCode", "") '//사용자값활성(사번)
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "SCardCode", "") '//사용자값활성(거래처)
        
        
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
                
                Call oForm01.Freeze(True)

                'DataSource를 이용하여 각 컨트롤에 값을 출력
                Call oDS_PS_FI421H.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
                Call oDS_PS_FI421H.setValue("U_SerialNo", 0, oMat01.Columns("SerialNo").Cells(pval.Row).Specific.VALUE) '일련번호
                Call oDS_PS_FI421H.setValue("U_BPLId", 0, oMat01.Columns("BPLId").Cells(pval.Row).Specific.VALUE) '사업장
                Call oDS_PS_FI421H.setValue("U_RspCode", 0, oMat01.Columns("RspCode").Cells(pval.Row).Specific.VALUE) '담당
                Call oDS_PS_FI421H.setValue("U_CntcCode", 0, oMat01.Columns("CntcCode").Cells(pval.Row).Specific.VALUE) '사번
                Call oDS_PS_FI421H.setValue("U_CntcName", 0, oMat01.Columns("CntcName").Cells(pval.Row).Specific.VALUE) '성명
                Call oDS_PS_FI421H.setValue("U_CardCode", 0, oMat01.Columns("CardCode").Cells(pval.Row).Specific.VALUE) '거래처
                Call oDS_PS_FI421H.setValue("U_CardName", 0, oMat01.Columns("CardName").Cells(pval.Row).Specific.VALUE) '거래처명
                Call oDS_PS_FI421H.setValue("U_DocDate", 0, oMat01.Columns("DocDate").Cells(pval.Row).Specific.VALUE) '발행일자
                Call oDS_PS_FI421H.setValue("U_Amount", 0, oMat01.Columns("Amount").Cells(pval.Row).Specific.VALUE) '공급가액
                Call oDS_PS_FI421H.setValue("U_VatTax", 0, oMat01.Columns("VatTax").Cells(pval.Row).Specific.VALUE) '부가가치세
                Call oDS_PS_FI421H.setValue("U_Contents", 0, oMat01.Columns("Contents").Cells(pval.Row).Specific.VALUE) '내용
                
                oForm01.Mode = fm_UPDATE_MODE
                Call PS_FI421_LoadCaption
                
                Call oForm01.Freeze(False)
                
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

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_FI421L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_FI421L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_FI421_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_FI421L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PS_FI421_FlushToItemValue(pval.ItemUID)
                
                If pval.ItemUID = "CntcCode" Then
                
                    oForm01.Items("CntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '성명
                    
                ElseIf pval.ItemUID = "SCntcCode" Then
                
                    oForm01.Items("SCntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("SCntcCode").Specific.VALUE & "'") '성명
                
                ElseIf pval.ItemUID = "CardCode" Then
                
                    oForm01.Items("CardName").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items("CardCode").Specific.VALUE & "'") '거래처명
                    
                ElseIf pval.ItemUID = "SCardCode" Then
                
                    oForm01.Items("SCardName").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items("SCardCode").Specific.VALUE & "'") '거래처명
                
                End If
                
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
        Call PS_FI421_FormItemEnabled
        '//Call PS_FI421_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_FI421_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_FI421H", "U_CardCode,U_CardName")
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
'            If (PS_FI421_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_FI421H.RemoveRecord(oDS_PS_FI421H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_FI421_Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_FI421H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_FI421_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_FI421_CreateItems() As Boolean
On Error GoTo PS_FI421_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_FI421H = oForm01.DataSources.DBDataSources("@PS_FI421H")
    Set oDS_PS_FI421L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '관리번호
    Call oForm01.DataSources.UserDataSources.Add("SDocEntry", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SDocEntry").Specific.DataBind.SetBound(True, "", "SDocEntry")
    
    '일련번호
    Call oForm01.DataSources.UserDataSources.Add("SSerialNo", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SSerialNo").Specific.DataBind.SetBound(True, "", "SSerialNo")
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("SBPLId", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SBPLId").Specific.DataBind.SetBound(True, "", "SBPLId")
    '사업장_E
    
    '담당
    Call oForm01.DataSources.UserDataSources.Add("SRspCode", dt_SHORT_TEXT, 5)
    Call oForm01.Items("SRspCode").Specific.DataBind.SetBound(True, "", "SRspCode")
    
    '사번
    Call oForm01.DataSources.UserDataSources.Add("SCntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SCntcCode").Specific.DataBind.SetBound(True, "", "SCntcCode")
    
    '성명
    Call oForm01.DataSources.UserDataSources.Add("SCntcName", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SCntcName").Specific.DataBind.SetBound(True, "", "SCntcName")
    
    '거래처
    Call oForm01.DataSources.UserDataSources.Add("SCardCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SCardCode").Specific.DataBind.SetBound(True, "", "SCardCode")
    
    '거래처명
    Call oForm01.DataSources.UserDataSources.Add("SCardName", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SCardName").Specific.DataBind.SetBound(True, "", "SCardName")
    
    '발행일자FR
    Call oForm01.DataSources.UserDataSources.Add("SDocDateFr", dt_DATE)
    Call oForm01.Items("SDocDateFr").Specific.DataBind.SetBound(True, "", "SDocDateFr")
    
    '발행일자TO
    Call oForm01.DataSources.UserDataSources.Add("SDocDateTo", dt_DATE)
    Call oForm01.Items("SDocDateTo").Specific.DataBind.SetBound(True, "", "SDocDateTo")
    
    '공급가액
    Call oForm01.DataSources.UserDataSources.Add("SAmount", dt_SUM)
    Call oForm01.Items("SAmount").Specific.DataBind.SetBound(True, "", "SAmount")
    
    '부가가치세
    Call oForm01.DataSources.UserDataSources.Add("SVatTax", dt_SUM)
    Call oForm01.Items("SVatTax").Specific.DataBind.SetBound(True, "", "SVatTax")

    '내용
    Call oForm01.DataSources.UserDataSources.Add("SContents", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SContents").Specific.DataBind.SetBound(True, "", "SContents")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_FI421_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI421_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_FI421_ComboBox_Setting()
On Error GoTo PS_FI421_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
        
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call oForm01.Items("SBPLId").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("BPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    Set oCombo = oForm01.Items("SBPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '담당(기본정보)
    Call oForm01.Items("RspCode").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("RspCode").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'F003'", "1", False, False)

    Set oCombo = oForm01.Items("RspCode").Specific
    Call oCombo.Select("%", psk_ByValue)
    
    '담당(조회조건)
    Call oForm01.Items("SRspCode").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SRspCode").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'F003'", "1", False, False)

    Set oCombo = oForm01.Items("SRspCode").Specific
    Call oCombo.Select("%", psk_ByValue)
    
    '매트릭스
    '사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    '담당
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("RspCode"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'F003'")
    
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_FI421_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FI421_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FI421_CF_ChooseFromList()
On Error GoTo PS_FI421_CF_ChooseFromList_Error

    Exit Sub
PS_FI421_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI421_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FI421_EnableMenus()
On Error GoTo PS_FI421_EnableMenus_Error
    
    Exit Sub
PS_FI421_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI421_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FI421_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_FI421_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_FI421_FormItemEnabled
        '//Call PS_FI421_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_FI421_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_FI421_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI421_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FI421_FormResize()
On Error GoTo PS_FI421_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_FI421_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_FI421_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_FI421_FormReset()
'******************************************************************************
'Function ID : PS_FI421_FormReset()
'해당모듈 : PS_FI421
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FI421_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_FI421H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_FI421H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_FI421H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    Call oDS_PS_FI421H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId) '사업장
    Call oDS_PS_FI421H.setValue("U_RspCode", 0, "%") '담당
    Call oDS_PS_FI421H.setValue("U_CntcCode", 0, "") '사번
    Call oDS_PS_FI421H.setValue("U_CntcName", 0, "") '성명
    Call oDS_PS_FI421H.setValue("U_CardCode", 0, "") '거래처코드
    Call oDS_PS_FI421H.setValue("U_CardName", 0, "") '거래처명
    Call oDS_PS_FI421H.setValue("U_DocDate", 0, Format(Date, "YYYYMMDD")) '발행일자
    Call oDS_PS_FI421H.setValue("U_Amount", 0, "0") '공급가액
    Call oDS_PS_FI421H.setValue("U_VatTax", 0, "0") '부가가치세
    Call oDS_PS_FI421H.setValue("U_Contents", 0, "") '내용
    '일련번호
    Call PS_FI421_GetSerialNo
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FI421_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_FI421_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_FI421_GetSerialNo()
'******************************************************************************
'Function ID : PS_FI421_GetSerialNo()
'해당모듈 : PH_PY030
'기능 : 출장번호 생성
'인수 : 없음
'반환값 : 합계 금액
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FI421_GetSerialNo_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ErrNum As Integer
    
    Dim DocDate As String
    Dim BPLId As String
    
    BPLId = oForm01.Items("BPLId").Specific.VALUE
    DocDate = Left(Trim(oForm01.Items("DocDate").Specific.VALUE), 6)
    
    sQry = "EXEC PS_FI421_05 '" & BPLId & "', '" & DocDate & "'"
    Call oRecordSet01.DoQuery(sQry)

    oForm01.Items("SerialNo").Specific.VALUE = oRecordSet01.Fields("SerialNo").VALUE
    
    Set oRecordSet01 = Nothing

    Exit Sub

PS_FI421_GetSerialNo_Error:
    If ErrNum = 1 Then
    Else
        MDC_Com.MDC_GF_Message "PS_FI421_GetSerialNo_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    Set oRecordSet01 = Nothing
    
End Sub

Private Sub PS_FI421_Print_Report01()

    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_FI421_Print_Report01_Error
    
    Dim DocEntry As String
    Dim BPLId As String
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
     '/ ODBC 연결 체크
    Call ConnectODBC
    
    '//인자 MOVE , Trim 시키기..
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PS_FI421] 입금표"
    
    If BPLId = "1" Then '창원
        ReportName = "PS_FI421_01.rpt"
    ElseIf BPLId = "2" Then '동래
        ReportName = "PS_FI421_02.rpt"
    ElseIf BPLId = "3" Then '사상
        ReportName = "PS_FI421_03.rpt"
    End If
    
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    sQry = "EXEC [PS_FI421_90] '" & BPLId & "','" & DocEntry & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PS_FI421_Print_Report01_Error
    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_FI421_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_FI421_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub
