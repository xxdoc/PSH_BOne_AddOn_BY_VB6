VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FX010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_FX010.cls
'//  Module : 자산관리 > 자산이력관리
'//  Desc : 자산이력등록
'//  FormType : PS_FX010
'//  Create Date(Start) : 2012.07.23
'//  Create Date(End) : 2012.07.23
'//  Creator : N.G.Y
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_FX010H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_FX010L As SAPbouiCOM.DBDataSource '등록라인

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

Private oBPLId As String
Private oDocdate As String
Private oHisType As String
Private oClasCode As String


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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FX010.srf")
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
    
    oFormUniqueID01 = "PS_FX010_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_FX010_CreateItems
    Call PS_FX010_ComboBox_Setting
    Call PS_FX010_CF_ChooseFromList
    Call PS_FX010_EnableMenus
    Call PS_FX010_SetDocument(oFromDocEntry01)
    Call PS_FX010_FormResize

    Call PS_FX010_Add_MatrixRow(0, True)
    Call PS_FX010_LoadCaption
    Call PS_FX010_FormItemEnabled
    
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
        
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_FX010H]"
    Call RecordSet01.DoQuery(sQry)
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_FX010H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_FX010H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
        
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing

    '발생일자 설정
    Call oDS_PS_FX010H.setValue("U_DocDate", 0, Format(Date, "YYYYMMDD"))
    
    
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX010_LoadCaption()
'******************************************************************************
'Function ID : PS_FX010_LoadCaption()
'해당모듈 : PS_FX010
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX010_LoadCaption_Error

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
PS_FX010_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_FX010_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_FX010_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_FX010L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_FX010L.Offset = oRow
    oDS_PS_FX010L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX010_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_FX010_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_FX010_MTX01()
'******************************************************************************
'Function ID : PS_FX010_MTX01()
'해당모듈 : PS_FX010
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX010_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim SBPLID As String '사업장
    Dim SHisType As String '이력구분
    Dim DocDateF As String '일자_F
    Dim DocDateT As String '일자_T
    
    SBPLID = Trim(oForm01.Items("SBPLId").Specific.VALUE)
    SHisType = Trim(oForm01.Items("SHisType").Specific.VALUE)
    DocDateF = Trim(oForm01.Items("DocDateF").Specific.VALUE)
    DocDateT = Trim(oForm01.Items("DocDateT").Specific.VALUE)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "EXEC [PS_FX010_01] '" & SBPLID & "','" & DocDateF & "','" & DocDateT & "','" & SHisType & "'"
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_FX010L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
        Call PS_FX010_Add_MatrixRow(0, True)
        Call PS_FX010_LoadCaption
        
        GoTo PS_FX010_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_FX010L.Size Then
            oDS_PS_FX010L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_FX010L.Offset = i
        
        Call oDS_PS_FX010L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_FX010L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("BPLId").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("DocDate").VALUE), "YYYYMMDD"))
        Call oDS_PS_FX010L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("HIsType").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("ClasCode").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("AmtYN").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("FixCode").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("SubCode").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("FixName").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("Qty").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Amt").VALUE))
        Call oDS_PS_FX010L.setValue("U_ColTxt01", i, Trim(oRecordSet01.Fields("Comments").VALUE))
        
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
PS_FX010_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_FX010_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_FX010_DeleteData()
'******************************************************************************
'Function ID : PS_FX010_DeleteData()
'해당모듈 : PS_FX010
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX010_DeleteData_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PS_FX010H] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_FX010_DeleteData_Error
        Else
            sQry = "DELETE FROM [@PS_FX010H] WHERE DocEntry = '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
'    Call PS_FX010_FormReset
    
'    oForm01.Mode = fm_ADD_MODE
    
'    Call oForm01.Items("BtnSearch").Click(ct_Regular)
    
'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PS_FX010_Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX010_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_FX010_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_FX010_UpdateData() As Boolean
'******************************************************************************
'Function ID : PS_FX010_UpdateData()
'해당모듈 : PS_FX010
'기능 : 기본정보를 수정
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX010_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim BPLID As String '사업장
    Dim HisType As String '이력구분
    Dim ClasCode As String '자산구분
    Dim DocDate As String '이력일자
    Dim FixCode As String '자산코드
    Dim SubCode As String '자산순번
    Dim FixName As String '자산명
    Dim AmtYN As String '자본적지출여부
    Dim Qty$ '수량
    Dim Amt$ '금액
    Dim Comments As String '비고사항
    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    HisType = Trim(oForm01.Items("HisType").Specific.VALUE) '이력구분
    ClasCode = Trim(oForm01.Items("ClasCode").Specific.VALUE) '자산구분
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE) '이력일자
    AmtYN = Trim(oForm01.Items("AmtYN").Specific.VALUE) '자본적지출여부
    FixCode = Trim(oForm01.Items("FixCode").Specific.VALUE) '자산코드
    SubCode = Trim(oForm01.Items("SubCode").Specific.VALUE) '자산순번
    FixName = Trim(oForm01.Items("FixName").Specific.VALUE) '자산명
    Qty = Trim(oForm01.Items("Qty").Specific.VALUE) '수량
    Amt = Trim(oForm01.Items("Amt").Specific.VALUE) '금액
    Comments = Trim(oForm01.Items("Comments").Specific.VALUE) '비고
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
        PS_FX010_UpdateData = False
        Exit Function
    End If
    
    sQry = "           UPDATE   [@PS_FX010H]"
    sQry = sQry & " SET         U_BPLId = '" & BPLID & "',"
    sQry = sQry & "              U_HisType = '" & HisType & "',"
    sQry = sQry & "              U_ClasCode = '" & ClasCode & "',"
    sQry = sQry & "              U_DocDate = '" & DocDate & "',"
    sQry = sQry & "              U_AmtYN = '" & AmtYN & "',"
    sQry = sQry & "              U_FixCode = '" & FixCode & "',"
    sQry = sQry & "              U_SubCode = '" & SubCode & "',"
    sQry = sQry & "              U_FixName = '" & FixName & "',"
    sQry = sQry & "              U_Qty  = '" & Qty & "',"
    sQry = sQry & "              U_Amt  = '" & Amt & "',"
    sQry = sQry & "              U_Comments = '" & Comments & "'"
    sQry = sQry & " WHERE    DocEntry = '" & DocEntry & "'"
    
    Call RecordSet01.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PS_FX010_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX010_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_FX010_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_FX010_AddData() As Boolean
'******************************************************************************
'Function ID : PS_FX010_AddData()
'해당모듈 : PS_FX010
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX010_AddData_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim BPLID As String '사업장
    Dim HisType As String '이력구분
    Dim ClasCode As String '자산구분
    Dim DocDate As String '이력일자
    Dim FixCode As String '자산코드
    Dim SubCode As String '자산순번
    Dim FixName As String '자산명
    Dim AmtYN As String '자본적지출여부
    Dim Qty$  '수량
    Dim Amt$  '금액
    Dim Comments As String '비고사항
    
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    HisType = Trim(oForm01.Items("HisType").Specific.VALUE) '이력구분
    ClasCode = Trim(oForm01.Items("ClasCode").Specific.VALUE) '자산구분
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE) '이력일자
    AmtYN = Trim(oForm01.Items("AmtYN").Specific.VALUE) '자본적지출여부
    FixCode = Trim(oForm01.Items("FixCode").Specific.VALUE) '자산코드
    SubCode = Trim(oForm01.Items("SubCode").Specific.VALUE) '자산순번
    FixName = Trim(oForm01.Items("FixName").Specific.VALUE) '자산명
    Qty = Trim(oForm01.Items("Qty").Specific.VALUE) '수량
    Amt = Trim(oForm01.Items("Amt").Specific.VALUE) '금액
    Comments = Trim(oForm01.Items("Comments").Specific.VALUE) '비고
    
    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_FX010H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "           INSERT INTO [@PS_FX010H]"
    sQry = sQry & " ("
    sQry = sQry & "     DocEntry,"
    sQry = sQry & "     DocNum,"
    sQry = sQry & "     U_BPLId,"
    sQry = sQry & "     U_HIsType,"
    sQry = sQry & "     U_ClasCode,"
    sQry = sQry & "     U_DocDate,"
    sQry = sQry & "     U_FixCode,"
    sQry = sQry & "     U_SubCode,"
    sQry = sQry & "     U_FixName,"
    sQry = sQry & "     U_AmtYN,"
    sQry = sQry & "     U_Qty,"
    sQry = sQry & "     U_Amt,"
    sQry = sQry & "     U_Comments"
    sQry = sQry & " )"
    sQry = sQry & " VALUES"
    sQry = sQry & " ("
    sQry = sQry & DocEntry & ","
    sQry = sQry & DocEntry & ","
    sQry = sQry & "'" & BPLID & "',"
    sQry = sQry & "'" & HisType & "',"
    sQry = sQry & "'" & ClasCode & "',"
    sQry = sQry & "'" & DocDate & "',"
    sQry = sQry & "'" & FixCode & "',"
    sQry = sQry & "'" & SubCode & "',"
    sQry = sQry & "'" & FixName & "',"
    sQry = sQry & "'" & AmtYN & "',"
    sQry = sQry & "'" & Qty & "',"
    sQry = sQry & "'" & Amt & "',"
    sQry = sQry & "'" & Comments & "'"
    sQry = sQry & ")"
    
    Call RecordSet02.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_FX010_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX010_AddData_Error:
    PS_FX010_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PS_FX010_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_FX010_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_FX010_HeaderSpaceLineDel()
'해당모듈 : PS_FX010
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX010_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("HisType").Specific.VALUE) = "" '이력구분
            ErrNum = 1
            GoTo PS_FX010_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ClasCode").Specific.VALUE) = "" '자산분류
            ErrNum = 2
            GoTo PS_FX010_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("FixCode").Specific.VALUE) = "" '자산코드
            ErrNum = 3
            GoTo PS_FX010_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("FixName").Specific.VALUE) = "" '자산명
            ErrNum = 4
            GoTo PS_FX010_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("DocDate").Specific.VALUE) = "" '이력일자
            ErrNum = 5
            GoTo PS_FX010_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("AmtYN").Specific.VALUE) = "Y" '자본적지출 = 'Y'
            If oForm01.Items("Amt").Specific.VALUE = 0 Then
                ErrNum = 6
                GoTo PS_FX010_HeaderSpaceLineDel_Error
            End If
    End Select

    PS_FX010_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX010_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "이력구분은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("HisType").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "자산분류는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ClasCode").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "자산코드는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("FixCode").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "자산명은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("FixName").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "이력일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("DocDate").Click(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "자본적지출시 금액은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Amt").Click(ct_Regular)
    End If
    PS_FX010_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PS_FX010_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_FX010_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PS_FX010_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX010_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PS_FX010_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_FX010_MatrixSpaceLineDel = False
End Function

Private Sub PS_FX010_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_FX010_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    Dim Qty As Integer
    Dim Calculate_Weight As Double
    Dim FixCode$, SubCode$
    
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CFixCode"
            FixCode = Left(oForm01.Items("CFixCode").Specific.VALUE, 6)
            SubCode = Right(oForm01.Items("CFixCode").Specific.VALUE, 3)
            '
            oForm01.Items("FixCode").Specific.VALUE = FixCode
            oForm01.Items("SubCode").Specific.VALUE = SubCode
            
            sQry = "Select U_FixName From [@PS_FX005H] Where U_FixCode = '" & FixCode & "'"
            sQry = sQry & " and U_SubCode = '" & SubCode & "'"
            oRecordSet01.DoQuery sQry
            oForm01.Items("FixName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "SubCode" '/자산코드
            
    End Select


   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_FX010_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("시각은 숫자만 입력이 가능합니다.", "E")
    ElseIf ErrNum = 2 Then
        Call MDC_Com.MDC_GF_Message("시각(시)는 24미만의 값만 입력이 가능합니다.", "E")
    ElseIf ErrNum = 3 Then
        Call MDC_Com.MDC_GF_Message("시각(분)은 60미만의 값만 입력이 가능합니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_FX010_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_FX010_FormItemEnabled()
On Error GoTo PS_FX010_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX010_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_FX010_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
                Call PS_FX010_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_FX010_LoadCaption
                
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
                '//Call PS_FX010_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_FX010H.Clear

'                Call PS_FX010_LoadCaption
'                Call PS_FX010_FormItemEnabled
                '//Call PS_FX010_FormItemEnabled '//UDO방식
                '//Call PS_FX010_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_FX010_FormItemEnabled
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
    
        If pval.ItemUID = "PS_FX010" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_FX010_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If PS_FX010_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_FX010_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_FX010_LoadCaption
                Call PS_FX010_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_FX010_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                If PS_FX010_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_FX010_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_FX010_LoadCaption
                Call PS_FX010_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
                
            Call PS_FX010_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_FX010_LoadCaption
            Call PS_FX010_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_FX010_DeleteData
                Call PS_FX010_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_FX010_LoadCaption
                Call PS_FX010_MTX01
                
            Else
            
            End If
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_FX010" Then
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
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "CFixCode" Then
                If oForm01.Items("CFixCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
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
                
                Call oForm01.Freeze(True)

                'DataSource를 이용하여 각 컨트롤에 값을 출력
                Call oDS_PS_FX010H.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
                Call oDS_PS_FX010H.setValue("U_BPLId", 0, oMat01.Columns("BPLId").Cells(pval.Row).Specific.VALUE) '사업장
                Call oDS_PS_FX010H.setValue("U_HisType", 0, oMat01.Columns("HisType").Cells(pval.Row).Specific.VALUE) '이력구분
                Call oDS_PS_FX010H.setValue("U_ClasCode", 0, oMat01.Columns("ClasCode").Cells(pval.Row).Specific.VALUE) '자산분류
                Call oDS_PS_FX010H.setValue("U_AmtYN", 0, oMat01.Columns("AmtYN").Cells(pval.Row).Specific.VALUE) '자본적지출여부
                Call oDS_PS_FX010H.setValue("U_DocDate", 0, oMat01.Columns("DocDate").Cells(pval.Row).Specific.VALUE) '이력일자
                Call oDS_PS_FX010H.setValue("U_FixCode", 0, oMat01.Columns("FixCode").Cells(pval.Row).Specific.VALUE) '자산코드
                Call oDS_PS_FX010H.setValue("U_SubCode", 0, oMat01.Columns("SubCode").Cells(pval.Row).Specific.VALUE) '자산순번
                Call oDS_PS_FX010H.setValue("U_FixName", 0, oMat01.Columns("FixName").Cells(pval.Row).Specific.VALUE) '자산명
                Call oDS_PS_FX010H.setValue("U_Qty", 0, oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) '수량
                Call oDS_PS_FX010H.setValue("U_Amt", 0, oMat01.Columns("Amt").Cells(pval.Row).Specific.VALUE) '금액
                Call oDS_PS_FX010H.setValue("U_Comments", 0, oMat01.Columns("Comments").Cells(pval.Row).Specific.VALUE) '비고
                
                oForm01.Items("CFixCode").Specific.VALUE = oMat01.Columns("FixCode").Cells(pval.Row).Specific.VALUE & "-" & oMat01.Columns("SubCode").Cells(pval.Row).Specific.VALUE
                
                oForm01.Mode = fm_UPDATE_MODE
                Call PS_FX010_LoadCaption
                
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

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "BPLId" Then
            oBPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
        End If
        
        If pval.ItemUID = "ClasCode" Then
            oClasCode = oForm01.Items("ClasCode").Specific.VALUE
        End If
        
        If pval.ItemUID = "ClasCode" Then
            oHisType = oForm01.Items("HisType").Specific.VALUE
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
'                    Call oDS_PS_FX010L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_FX010L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_FX010_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_FX010L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PS_FX010_FlushToItemValue(pval.ItemUID)
                
                If pval.ItemUID = "BPLId" Then
                    oBPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
                End If
                
                If pval.ItemUID = "DocDate" Then
                    oDocdate = Trim(oForm01.Items("DocDate").Specific.VALUE)
                End If
                
                If pval.ItemUID = "HisType" Then
                    oHisType = oForm01.Items("HisType").Specific.VALUE
                End If
                
                If pval.ItemUID = "ClasCode" Then
                    oClasCode = oForm01.Items("ClasCode").Specific.VALUE
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
        Call PS_FX010_FormItemEnabled
        '//Call PS_FX010_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_FX010_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_FX010H", "U_CardCode,U_CardName")
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
'            If (PS_FX010_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_FX010H.RemoveRecord(oDS_PS_FX010H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_FX010_Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_FX010H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_FX010_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_FX010_CreateItems() As Boolean
On Error GoTo PS_FX010_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_FX010H = oForm01.DataSources.DBDataSources("@PS_FX010H")
    Set oDS_PS_FX010L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("SBPLId", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SBPLId").Specific.DataBind.SetBound(True, "", "SBPLId")
    '사업장_E
    
    
    '이력구분_S
    Call oForm01.DataSources.UserDataSources.Add("SHisType", dt_SHORT_TEXT, 100)
    Call oForm01.Items("SHisType").Specific.DataBind.SetBound(True, "", "SHisType")
    '이력구분_E
    
    '자산코드(검색용)_S
    Call oForm01.DataSources.UserDataSources.Add("CFixCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("CFixCode").Specific.DataBind.SetBound(True, "", "CFixCode")
    '자산코드(검색용)_E
    
    
    '이력일자_From_S
    Call oForm01.DataSources.UserDataSources.Add("DocDateF", dt_DATE)
    Call oForm01.Items("DocDateF").Specific.DataBind.SetBound(True, "", "DocDateF")
    '이력일자_From_E
    
    
    
    '이력일자_To_S
    Call oForm01.DataSources.UserDataSources.Add("DocDateT", dt_DATE)
    Call oForm01.Items("DocDateT").Specific.DataBind.SetBound(True, "", "DocDateT")
    '이력일자_To_E
    
    oForm01.Items("DocDateF").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    oForm01.Items("DocDateT").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_FX010_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX010_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_FX010_ComboBox_Setting()
On Error GoTo PS_FX010_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
        
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("BPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    Set oCombo = oForm01.Items("SBPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '이력구분
    Call oForm01.Items("HisType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("HisType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'", "", False, False)
    Call oForm01.Items("HisType").Specific.Select(0, psk_Index)
    
    '자산분류
    Call oForm01.Items("ClasCode").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ClasCode").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX001'", "", False, False)
    Call oForm01.Items("ClasCode").Specific.Select(0, psk_Index)
    
    '이력구분(조회조건)
    Call oForm01.Items("SHisType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SHisType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'", "", False, False)
    Call oForm01.Items("SHisType").Specific.Select(0, psk_Index)
    
    Call oForm01.Items("AmtYN").Specific.ValidValues.Add("N", "N")
    Call oForm01.Items("AmtYN").Specific.ValidValues.Add("Y", "Y")
    Call oForm01.Items("AmtYN").Specific.Select(0, psk_Index)
    
    '사업장(매트릭스)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    '이력구분(매트릭스)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("HisType"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'")
    '자산구분(매트릭스)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ClasCode"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX001'")
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_FX010_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX010_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX010_CF_ChooseFromList()
On Error GoTo PS_FX010_CF_ChooseFromList_Error

    Exit Sub
PS_FX010_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX010_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX010_EnableMenus()
On Error GoTo PS_FX010_EnableMenus_Error
    
    Exit Sub
PS_FX010_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX010_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX010_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_FX010_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_FX010_FormItemEnabled
        '//Call PS_FX010_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_FX010_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_FX010_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX010_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX010_FormResize()
On Error GoTo PS_FX010_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_FX010_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX010_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_FX010_FormReset()
'******************************************************************************
'Function ID : PS_FX010_FormReset()
'해당모듈 : PS_FX010
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX010_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_FX010H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_FX010H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_FX010H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    If oBPLId = "" Then oBPLId = MDC_PS_Common.User_BPLId
    If oHisType = "" Then oHisType = "%"
    If oClasCode = "" Then oClasCode = "%"
    If oDocdate = "" Then oDocdate = Format(oDocdate, "YYYYMMDD")
    
    Call oDS_PS_FX010H.setValue("U_BPLId", 0, oBPLId) '사업장
    Call oDS_PS_FX010H.setValue("U_HisType", 0, oHisType) '이력구분
    Call oDS_PS_FX010H.setValue("U_ClasCode", 0, oClasCode) '자산분류
    Call oDS_PS_FX010H.setValue("U_AmtYN", 0, "N") '자본적지출여부
    Call oDS_PS_FX010H.setValue("U_DocDate", 0, oDocdate)  '이력일자
    Call oDS_PS_FX010H.setValue("U_FixCode", 0, "") '자산코드
    Call oDS_PS_FX010H.setValue("U_SubCode", 0, "") '자산순번
    Call oDS_PS_FX010H.setValue("U_FixName", 0, "") '자산명
    Call oDS_PS_FX010H.setValue("U_Qty", 0, 0) '수량
    Call oDS_PS_FX010H.setValue("U_Amt", 0, 0) '금액
    Call oDS_PS_FX010H.setValue("U_Comments", 0, "") '비고사항
    
    oForm01.Items("CFixCode").Specific.VALUE = ""
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX010_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_FX010_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


