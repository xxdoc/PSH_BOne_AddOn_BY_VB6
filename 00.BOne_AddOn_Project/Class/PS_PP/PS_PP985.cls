VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP985"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_PP985.cls
'//  Module : 생산관리>풍산홀딩스애드온>생산리포트
'//  Desc : 작번별 진행금액 일괄조회
'//  FormType : PS_PP985
'//  Create Date(Start) : 2013.01.02
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
Public oMat02 As SAPbouiCOM.Matrix
Public oMat03 As SAPbouiCOM.Matrix
Public oMat04 As SAPbouiCOM.Matrix
Public oMat05 As SAPbouiCOM.Matrix
Public oMat06 As SAPbouiCOM.Matrix

Private oDS_PS_PP985H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP985A As SAPbouiCOM.DBDataSource '자재비
Private oDS_PS_PP985B As SAPbouiCOM.DBDataSource '자체가공비
Private oDS_PS_PP985C As SAPbouiCOM.DBDataSource '외주가공비
Private oDS_PS_PP985D As SAPbouiCOM.DBDataSource '외주제작비
Private oDS_PS_PP985E As SAPbouiCOM.DBDataSource '설계비
Private oDS_PS_PP985F As SAPbouiCOM.DBDataSource '기성매출


Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 매트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 매트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private ItemInformationCount As Long

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP985.srf")
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
    
    oFormUniqueID01 = "PS_PP985_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP985_CreateItems
'    Call PS_PP985_ComboBox_Setting
'    Call PS_PP985_CF_ChooseFromList
'    Call PS_PP985_EnableMenus
'    Call PS_PP985_SetDocument(oFromDocEntry01)
'    Call PS_PP985_FormResize

'    Call PS_PP985_Add_MatrixRow01(0, True)
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True
        
    oForm01.Update
    oForm01.Freeze False
'
'    '품목코드 포커스
'    Call oForm01.Items("ItemCode").Click
    
    oMat01.Columns("PurchaseNm").Visible = False '품의구분은 굳이 필요 없을 것 같아 Hidden 처리
    
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

'/매트릭스 Row추가
Sub PS_PP985_Add_MatrixRow01(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP985_Add_MatrixRow01_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP985A.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_PP985A.Offset = oRow
    oDS_PS_PP985A.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Add_MatrixRow01_Error:
    MDC_Com.MDC_GF_Message "PS_PP985_Add_MatrixRow01_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/매트릭스 Row추가
Sub PS_PP985_Add_MatrixRow02(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP985_Add_MatrixRow02_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP985B.InsertRecord (oRow)
    End If
    
    oMat02.AddRow
    oDS_PS_PP985B.Offset = oRow
    oDS_PS_PP985B.setValue "U_LineNum", oRow, oRow + 1
    
    oMat02.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Add_MatrixRow02_Error:
    MDC_Com.MDC_GF_Message "PS_PP985_Add_MatrixRow02_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/매트릭스 Row추가
Sub PS_PP985_Add_MatrixRow03(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP985_Add_MatrixRow03_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP985C.InsertRecord (oRow)
    End If
    
    oMat03.AddRow
    oDS_PS_PP985C.Offset = oRow
    oDS_PS_PP985C.setValue "U_LineNum", oRow, oRow + 1
    
    oMat03.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Add_MatrixRow03_Error:
    MDC_Com.MDC_GF_Message "PS_PP985_Add_MatrixRow03_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/매트릭스 Row추가
Sub PS_PP985_Add_MatrixRow04(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP985_Add_MatrixRow04_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP985D.InsertRecord (oRow)
    End If
    
    oMat04.AddRow
    oDS_PS_PP985D.Offset = oRow
    oDS_PS_PP985D.setValue "U_LineNum", oRow, oRow + 1
    
    oMat04.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Add_MatrixRow04_Error:
    MDC_Com.MDC_GF_Message "PS_PP985_Add_MatrixRow04_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/매트릭스 Row추가
Sub PS_PP985_Add_MatrixRow05(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP985_Add_MatrixRow05_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP985E.InsertRecord (oRow)
    End If
    
    oMat05.AddRow
    oDS_PS_PP985E.Offset = oRow
    oDS_PS_PP985E.setValue "U_LineNum", oRow, oRow + 1
    
    oMat05.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Add_MatrixRow05_Error:
    MDC_Com.MDC_GF_Message "PS_PP985_Add_MatrixRow05_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/매트릭스 Row추가
Sub PS_PP985_Add_MatrixRow06(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP985_Add_MatrixRow06_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP985F.InsertRecord (oRow)
    End If
    
    oMat06.AddRow
    oDS_PS_PP985F.Offset = oRow
    oDS_PS_PP985F.setValue "U_LineNum", oRow, oRow + 1
    
    oMat06.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Add_MatrixRow06_Error:
    MDC_Com.MDC_GF_Message "PS_PP985_Add_MatrixRow06_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_PP985_Accumulate()
'******************************************************************************
'Function ID : PS_PP985_Accumulate()
'해당모듈 : PS_PP985
'기능 : 데이터 집계 및 테이블 저장
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP985_Accumulate_Error

    Dim sQry As String
    Dim ErrNum As Integer

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CntcCode As String '조회자 사번
    Dim OrdNum As String '메인작번
    Dim SubNo1 As String '서브작번1
    Dim SubNo2 As String '서브작번2
    Dim FrDt As String '기간(시작)
    Dim ToDt As String '기간(종료)
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '메인작번
    SubNo1 = Trim(oForm01.Items("SubNo1").Specific.VALUE) '서브작번1
    SubNo2 = Trim(oForm01.Items("SubNo2").Specific.VALUE) '서브작번2
    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE) '기간(시작)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE) '기간(종료)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("집계 시작!", oRecordSet01.RecordCount, False)
    
    '메인정보 집계
    sQry = "         EXEC [PS_PP985_91] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "','" '서브작번2
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "'" '기간(종료)
    
    Call oRecordSet01.DoQuery(sQry)
    
    '자재비 집계
    sQry = "         EXEC [PS_PP985_92] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "','" '서브작번2
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "'" '기간(종료)
    
    Call oRecordSet01.DoQuery(sQry)
    
    '자체가공비 집계
    sQry = "         EXEC [PS_PP985_93] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "','" '서브작번2
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "'" '기간(종료)
    
    Call oRecordSet01.DoQuery(sQry)
    
    '외주가공비 집계
    sQry = "         EXEC [PS_PP985_94] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "','" '서브작번2
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "'" '기간(종료)
    
    Call oRecordSet01.DoQuery(sQry)
    
    '외주제작비 집계
    sQry = "         EXEC [PS_PP985_95] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "','" '서브작번2
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "'" '기간(종료)
    
    Call oRecordSet01.DoQuery(sQry)
    
    '설계비 집계
    sQry = "         EXEC [PS_PP985_96] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "','" '서브작번2
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "'" '기간(종료)
    
    Call oRecordSet01.DoQuery(sQry)
    
    '기성매출 집계
    sQry = "         EXEC [PS_PP985_97] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "','" '서브작번2
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "'" '기간(종료)
    
    Call oRecordSet01.DoQuery(sQry)
    
    
    ProgBar01.VALUE = 100
    ProgBar01.Text = "비용 집계중...!"
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("집계가 완료되었습니다. 조회를 시작하십시오.", "W")

    Set oRecordSet01 = Nothing

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Accumulate_Error:
    ProgBar01.VALUE = 100
    ProgBar01.Stop
     Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP985_Accumulate_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_PP985_Select0()
'******************************************************************************
'Function ID : PS_PP985_Select0()
'해당모듈 : PS_PP985
'기능 : 주 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP985_Select0_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim CntcCode As String '조회자 사번
    Dim OrdNum As String '메인작번
    Dim SubNo1 As String '서브작번1
    Dim SubNo2 As String '서브작번2
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '메인작번
    SubNo1 = Trim(oForm01.Items("SubNo1").Specific.VALUE) '서브작번1
    SubNo2 = Trim(oForm01.Items("SubNo2").Specific.VALUE) '서브작번2
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_PP985_01] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "'" '서브작번2
    
    Call oRecordSet01.DoQuery(sQry)
    
'    oMat01.Clear
'    oDS_PS_PP985A.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
'        oForm01.Mode = fm_ADD_MODE
'        Call PS_PP985_Add_MatrixRow01(0, True)
        GoTo PS_PP985_Select0_Error
        
    End If
    
'    For i = 0 To oRecordSet01.RecordCount - 1
'
'        If i + 1 > oDS_PS_PP985A.Size Then
'            oDS_PS_PP985A.InsertRecord (i)
'        End If
'
'        oMat01.AddRow
'        oDS_PS_PP985A.Offset = i
'
'        Call oDS_PS_PP985A.setValue("U_LineNum", i, i + 1)
'        Call oDS_PS_PP985A.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("CpCode").VALUE)) '공정코드
'        Call oDS_PS_PP985A.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("CpName").VALUE)) '공정명
'        Call oDS_PS_PP985A.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("WkTime").VALUE)) '실동공수
'        Call oDS_PS_PP985A.setValue("U_ColQty02", i, Trim(oRecordSet01.Fields("Price").VALUE)) '공정단가
'        Call oDS_PS_PP985A.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Amt").VALUE)) '가공비(실동)
'        Call oDS_PS_PP985A.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("FirstWkDt").VALUE)) '최초작업일
'        Call oDS_PS_PP985A.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("LastWkDt").VALUE)) '최종작업일
'        Call oDS_PS_PP985A.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)) '품목(작번)
'        Call oDS_PS_PP985A.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품목명
'
'        oRecordSet01.MoveNext
'        ProgBar01.VALUE = ProgBar01.VALUE + 1
'        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
'
'    Next
'
'    oMat01.LoadFromDataSource
'    oMat01.AutoResizeColumns
    
    '수주금액
    oForm01.Items("SjAmt").Specific.VALUE = oRecordSet01.Fields("SjAmt").VALUE
    '자재비
    oForm01.Items("MatAmt").Specific.VALUE = oRecordSet01.Fields("MatAmt").VALUE
    '자체가공비
    oForm01.Items("GagongAmt").Specific.VALUE = oRecordSet01.Fields("GagongAmt").VALUE
    '외주가공비
    oForm01.Items("OutgAmt").Specific.VALUE = oRecordSet01.Fields("OutgAmt").VALUE
    '외주제작비
    oForm01.Items("OutmAmt").Specific.VALUE = oRecordSet01.Fields("OutmAmt").VALUE
    '설계비
    oForm01.Items("DrawAmt").Specific.VALUE = oRecordSet01.Fields("DrawAmt").VALUE
    '계
    oForm01.Items("Total").Specific.VALUE = oRecordSet01.Fields("Total").VALUE
    '생산완료여부
    oForm01.Items("PCYN").Specific.VALUE = oRecordSet01.Fields("PCYN").VALUE
    '생산완료일자
    oForm01.Items("EndDate").Specific.VALUE = oRecordSet01.Fields("EndDate").VALUE
    '수주일
    oForm01.Items("SjDocDate").Specific.VALUE = oRecordSet01.Fields("SjDocDate").VALUE
    '작업지시일
    oForm01.Items("WODate").Specific.VALUE = oRecordSet01.Fields("WODate").VALUE
    '납기일
    oForm01.Items("SjDueDate").Specific.VALUE = oRecordSet01.Fields("SjDueDate").VALUE
    '협력업체명(외주제작)
    oForm01.Items("OtCardNm").Specific.VALUE = oRecordSet01.Fields("OtCardNm").VALUE
    
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Select0_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("데이터가 존재하지 않습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP985_Select0_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP985_Select1()
'******************************************************************************
'Function ID : PS_PP985_Select1()
'해당모듈 : PS_PP985
'기능 : 자재비 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP985_Select1_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CntcCode As String '조회자 사번
    Dim OrdNum As String '메인작번
    Dim SubNo1 As String '서브작번1
    Dim SubNo2 As String '서브작번2
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '메인작번
    SubNo1 = Trim(oForm01.Items("SubNo1").Specific.VALUE) '서브작번1
    SubNo2 = Trim(oForm01.Items("SubNo2").Specific.VALUE) '서브작번2
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_PP985_02] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "'" '서브작번2
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_PP985A.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_PP985_Add_MatrixRow01(0, True)
        GoTo PS_PP985_Select1_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
    
        If i + 1 > oDS_PS_PP985A.Size Then
            oDS_PS_PP985A.InsertRecord (i)
        End If
    
        oMat01.AddRow
        oDS_PS_PP985A.Offset = i
    
        Call oDS_PS_PP985A.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP985A.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작번
        Call oDS_PS_PP985A.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("PurchaseNm").VALUE)) '품의구분
        Call oDS_PS_PP985A.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)) '자재품목코드
        Call oDS_PS_PP985A.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '자재품명
        Call oDS_PS_PP985A.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("MatAmt").VALUE)) '자재비
        
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
PS_PP985_Select1_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("자재비 조회 내역 결과가 없습니다.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP985_Select1_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP985_Select2()
'******************************************************************************
'Function ID : PS_PP985_Select2()
'해당모듈 : PS_PP985
'기능 : 자체가공비 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP985_Select2_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CntcCode As String '조회자 사번
    Dim OrdNum As String '메인작번
    Dim SubNo1 As String '서브작번1
    Dim SubNo2 As String '서브작번2
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '메인작번
    SubNo1 = Trim(oForm01.Items("SubNo1").Specific.VALUE) '서브작번1
    SubNo2 = Trim(oForm01.Items("SubNo2").Specific.VALUE) '서브작번2
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_PP985_03] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "'" '서브작번2
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat02.Clear
    oDS_PS_PP985B.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_PP985_Add_MatrixRow01(0, True)
        GoTo PS_PP985_Select2_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
    
        If i + 1 > oDS_PS_PP985B.Size Then
            oDS_PS_PP985B.InsertRecord (i)
        End If
    
        oMat02.AddRow
        oDS_PS_PP985B.Offset = i
    
        Call oDS_PS_PP985B.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP985B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작번
        Call oDS_PS_PP985B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("CpName").VALUE)) '공정명
        Call oDS_PS_PP985B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("StdTime").VALUE)) '표준공수
        Call oDS_PS_PP985B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("WkTime").VALUE)) '실동공수
        Call oDS_PS_PP985B.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Amt").VALUE)) '가공비(실동)
        Call oDS_PS_PP985B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("CompltDt").VALUE)) '완료요구일
        Call oDS_PS_PP985B.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("FirstWkDt").VALUE)) '최초작업일
        Call oDS_PS_PP985B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("LastWkDt").VALUE)) '최종작업일
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Select2_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("자재비 조회 내역 결과가 없습니다.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP985_Select2_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP985_Select3()
'******************************************************************************
'Function ID : PS_PP985_Select3()
'해당모듈 : PS_PP985
'기능 : 외주가공비 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP985_Select3_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CntcCode As String '조회자 사번
    Dim OrdNum As String '메인작번
    Dim SubNo1 As String '서브작번1
    Dim SubNo2 As String '서브작번2
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '메인작번
    SubNo1 = Trim(oForm01.Items("SubNo1").Specific.VALUE) '서브작번1
    SubNo2 = Trim(oForm01.Items("SubNo2").Specific.VALUE) '서브작번2
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_PP985_04] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "'" '서브작번2
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat03.Clear
    oDS_PS_PP985C.Clear
    oMat03.FlushToDataSource
    oMat03.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_PP985_Add_MatrixRow01(0, True)
        GoTo PS_PP985_Select3_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
    
        If i + 1 > oDS_PS_PP985C.Size Then
            oDS_PS_PP985C.InsertRecord (i)
        End If
    
        oMat03.AddRow
        oDS_PS_PP985C.Offset = i
    
        Call oDS_PS_PP985C.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP985C.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작번
        Call oDS_PS_PP985C.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("CpCode").VALUE)) '공정코드
        Call oDS_PS_PP985C.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("CpName").VALUE)) '공정명
        Call oDS_PS_PP985C.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명
        Call oDS_PS_PP985C.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Amt").VALUE)) '금액
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat03.LoadFromDataSource
    oMat03.AutoResizeColumns
    
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Select3_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("자재비 조회 내역 결과가 없습니다.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP985_Select3_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP985_Select4()
'******************************************************************************
'Function ID : PS_PP985_Select4()
'해당모듈 : PS_PP985
'기능 : 외주제작비 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP985_Select4_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim CntcCode As String '조회자 사번
    Dim OrdNum As String '메인작번
    Dim SubNo1 As String '서브작번1
    Dim SubNo2 As String '서브작번2
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '메인작번
    SubNo1 = Trim(oForm01.Items("SubNo1").Specific.VALUE) '서브작번1
    SubNo2 = Trim(oForm01.Items("SubNo2").Specific.VALUE) '서브작번2
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_PP985_05] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "'" '서브작번2
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat04.Clear
    oDS_PS_PP985D.Clear
    oMat04.FlushToDataSource
    oMat04.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_PP985_Add_MatrixRow01(0, True)
        GoTo PS_PP985_Select4_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
    
        If i + 1 > oDS_PS_PP985D.Size Then
            oDS_PS_PP985D.InsertRecord (i)
        End If
    
        oMat04.AddRow
        oDS_PS_PP985D.Offset = i
    
        Call oDS_PS_PP985D.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP985D.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작번
        Call oDS_PS_PP985D.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("CpCode").VALUE)) '공정코드
        Call oDS_PS_PP985D.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("CpName").VALUE)) '공정명
        Call oDS_PS_PP985D.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명
        Call oDS_PS_PP985D.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Amt").VALUE)) '금액
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat04.LoadFromDataSource
    oMat04.AutoResizeColumns
    
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Select4_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("외주제작비 조회 내역 결과가 없습니다.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP985_Select4_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP985_Select5()
'******************************************************************************
'Function ID : PS_PP985_Select5()
'해당모듈 : PS_PP985
'기능 : 설계비 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP985_Select5_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CntcCode As String '조회자 사번
    Dim OrdNum As String '메인작번
    Dim SubNo1 As String '서브작번1
    Dim SubNo2 As String '서브작번2
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '메인작번
    SubNo1 = Trim(oForm01.Items("SubNo1").Specific.VALUE) '서브작번1
    SubNo2 = Trim(oForm01.Items("SubNo2").Specific.VALUE) '서브작번2
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_PP985_06] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "'" '서브작번2
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat05.Clear
    oDS_PS_PP985E.Clear
    oMat05.FlushToDataSource
    oMat05.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_PP985_Add_MatrixRow01(0, True)
        GoTo PS_PP985_Select5_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
    
        If i + 1 > oDS_PS_PP985E.Size Then
            oDS_PS_PP985E.InsertRecord (i)
        End If
    
        oMat05.AddRow
        oDS_PS_PP985E.Offset = i
    
        Call oDS_PS_PP985E.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP985E.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작번
        Call oDS_PS_PP985E.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocDate").VALUE)) '일자
        Call oDS_PS_PP985E.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("WorkCode").VALUE)) '작업자사번
        Call oDS_PS_PP985E.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("WorkName").VALUE)) '작업자명
        Call oDS_PS_PP985E.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("PQty").VALUE)) '도면매수
        Call oDS_PS_PP985E.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Amt").VALUE)) '설계비
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat05.LoadFromDataSource
    oMat05.AutoResizeColumns
    
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Select5_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("설계비 조회 내역 결과가 없습니다.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP985_Select5_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP985_Select6()
'******************************************************************************
'Function ID : PS_PP985_Select6()
'해당모듈 : PS_PP985
'기능 : 기성매출 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP985_Select6_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CntcCode As String '조회자 사번
    Dim OrdNum As String '메인작번
    Dim SubNo1 As String '서브작번1
    Dim SubNo2 As String '서브작번2
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '메인작번
    SubNo1 = Trim(oForm01.Items("SubNo1").Specific.VALUE) '서브작번1
    SubNo2 = Trim(oForm01.Items("SubNo2").Specific.VALUE) '서브작번2
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_PP985_07] '"
    sQry = sQry & CntcCode & "','" '조회자사번
    sQry = sQry & OrdNum & "','" '메인작번
    sQry = sQry & SubNo1 & "','" '서브작번1
    sQry = sQry & SubNo2 & "'" '서브작번2
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat06.Clear
    oDS_PS_PP985F.Clear
    oMat06.FlushToDataSource
    oMat06.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_PP985_Add_MatrixRow01(0, True)
        GoTo PS_PP985_Select6_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
    
        If i + 1 > oDS_PS_PP985F.Size Then
            oDS_PS_PP985F.InsertRecord (i)
        End If
    
        oMat06.AddRow
        oDS_PS_PP985F.Offset = i
    
        Call oDS_PS_PP985F.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP985F.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("StdYear").VALUE)) '년도
        Call oDS_PS_PP985F.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Month01").VALUE)) '1월
        Call oDS_PS_PP985F.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("Month02").VALUE)) '2월
        Call oDS_PS_PP985F.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("Month03").VALUE)) '3월
        Call oDS_PS_PP985F.setValue("U_ColSum04", i, Trim(oRecordSet01.Fields("Month04").VALUE)) '4월
        Call oDS_PS_PP985F.setValue("U_ColSum05", i, Trim(oRecordSet01.Fields("Month05").VALUE)) '5월
        Call oDS_PS_PP985F.setValue("U_ColSum06", i, Trim(oRecordSet01.Fields("Month06").VALUE)) '6월
        Call oDS_PS_PP985F.setValue("U_ColSum07", i, Trim(oRecordSet01.Fields("Month07").VALUE)) '7월
        Call oDS_PS_PP985F.setValue("U_ColSum08", i, Trim(oRecordSet01.Fields("Month08").VALUE)) '8월
        Call oDS_PS_PP985F.setValue("U_ColSum09", i, Trim(oRecordSet01.Fields("Month09").VALUE)) '9월
        Call oDS_PS_PP985F.setValue("U_ColSum10", i, Trim(oRecordSet01.Fields("Month10").VALUE)) '10월
        Call oDS_PS_PP985F.setValue("U_ColSum11", i, Trim(oRecordSet01.Fields("Month11").VALUE)) '11월
        Call oDS_PS_PP985F.setValue("U_ColSum12", i, Trim(oRecordSet01.Fields("Month12").VALUE)) '12월
        Call oDS_PS_PP985F.setValue("U_ColSum13", i, Trim(oRecordSet01.Fields("Total").VALUE)) '계
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat06.LoadFromDataSource
    oMat06.AutoResizeColumns
    
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_Select6_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        'Call MDC_Com.MDC_GF_Message("기성매출 조회 내역 결과가 없습니다.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP985_Select6_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Function PS_PP985_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_PP985_HeaderSpaceLineDel()
'해당모듈 : PS_PP985
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP985_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("OrdNum").Specific.VALUE) = "" '작번
            ErrNum = 1
            GoTo PS_PP985_HeaderSpaceLineDel_Error
    End Select

    PS_PP985_HeaderSpaceLineDel = True
Exit Function

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("메인 작번은 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("OrdNum").Click(ct_Regular)
    End If
    PS_PP985_HeaderSpaceLineDel = False
End Function

''/ 매트릭스 필수 사항 check
'Private Function PS_PP985_MatrixSpaceLineDel() As Boolean
'On Error GoTo PS_PP985_MatrixSpaceLineDel_Error
'
'    Dim i As Long
'    Dim ErrNum As Integer
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Dim sQry As String
'
'    Set oRecordSet01 = Nothing
'    PS_PP985_MatrixSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_PP985_MatrixSpaceLineDel_Error:
'    Set oRecordSet01 = Nothing
'    If ErrNum = 1 Then
'        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 2 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 사원코드가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 시간이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 등록일자가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 5 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 비가동코드가 없습니다. 확인하세요.", "E"
'    Else
'        MDC_Com.MDC_GF_Message "PS_PP985_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'    PS_PP985_MatrixSpaceLineDel = False
'End Function

Private Sub PS_PP985_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_PP985_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    Dim Qty As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim OrdNum As String
    Dim SubNo1 As String
    Dim SubNo2 As String
    
    If oUID = "OrdNum" Or oUID = "SubNo1" Or oUID = "SubNo2" Then
        
        OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
        SubNo1 = oForm01.Items("SubNo1").Specific.VALUE
        SubNo2 = oForm01.Items("SubNo2").Specific.VALUE

        sQry = "           SELECT   CASE"
        sQry = sQry & "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakMyung"
        sQry = sQry & "             END AS [ItemName],"
        sQry = sQry & "             CASE"
        sQry = sQry & "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakSize"
        sQry = sQry & "             END AS [SPEC]"
        sQry = sQry & " FROM     [@PS_PP020H] AS T0"
        sQry = sQry & " WHERE   T0.U_JakName = '" & OrdNum & "'"
        sQry = sQry & "             AND T0.U_SubNo1 = CASE WHEN '" & SubNo1 & "' = '' THEN '00' ELSE '" & SubNo1 & "' END"
        sQry = sQry & "             AND T0.U_SubNo2 = CASE WHEN '" & SubNo2 & "' = '' THEN '000' ELSE '" & SubNo2 & "' END"

        Call oRecordSet01.DoQuery(sQry)

        oForm01.Items("ItemName").Specific.VALUE = oRecordSet01.Fields("ItemName").VALUE
        oForm01.Items("Spec").Specific.VALUE = oRecordSet01.Fields("Spec").VALUE
    
    End If
    
    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_PP985_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("시각은 숫자만 입력이 가능합니다.", "E")
'    ElseIf ErrNum = 2 Then
'        Call MDC_Com.MDC_GF_Message("시각(시)는 24미만의 값만 입력이 가능합니다.", "E")
'    ElseIf ErrNum = 3 Then
'        Call MDC_Com.MDC_GF_Message("시각(분)은 60미만의 값만 입력이 가능합니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP985_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_PP985_FormItemEnabled()
On Error GoTo PS_PP985_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP985_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_PP985_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                '/추가버튼 클릭시 매트릭스 insertrow
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
'                oForm01.Mode = fm_ADD_MODE
'                BubbleEvent = False
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm01.Freeze(True)
                
                If oForm01.Settings.MatrixUID = "Mat01" Then
                
                    Call PS_PP985_Add_MatrixRow01(oMat01.VisualRowCount)
                    
                End If
                
                Call oForm01.Freeze(False)

        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_PP985_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_PP985H.Clear

'                Call PS_PP985_LoadCaption
'                Call PS_PP985_FormItemEnabled
                '//Call PS_PP985_FormItemEnabled '//UDO방식
                '//Call PS_PP985_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_PP985_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                
                If oForm01.Settings.MatrixUID = "Mat01" Then
                
                    Call oDS_PS_PP985A.RemoveRecord(oDS_PS_PP985A.Size - 1)
                    Call oMat01.LoadFromDataSource
                    
                End If
                Call oForm01.Freeze(False)
            
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
    
        If pval.ItemUID = "PS_PP985" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAccm" Then
            
            If PS_PP985_HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            Else
                Call PS_PP985_Accumulate
            End If
        
        End If
        
        If pval.ItemUID = "BtnSearch" Then '/조회
            If PS_PP985_HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            Else
                Call PS_PP985_Select0 '메인 조회
                Call PS_PP985_Select1 '자재비 조회
                Call PS_PP985_Select2 '자체가공비 조회
                Call PS_PP985_Select3 '외주가공비 조회
                Call PS_PP985_Select4 '외주제작비 조회
                Call PS_PP985_Select5 '설계비 조회
                Call PS_PP985_Select6 '기성매출 조회
            End If
            
        End If
        
        If pval.ItemUID = "BtnPrint" Then '출력
            If PS_PP985_HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            Else
                Call PS_PP985_Print_Report01
            End If
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP985" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CpCode", "") '//사용자값활성
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    Dim True_False$

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Opt01" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Mat01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oMat04.AutoResizeColumns
            oMat05.AutoResizeColumns
            oMat06.AutoResizeColumns
            Call oForm01.Freeze(False)
        End If
        If pval.ItemUID = "Opt02" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Mat02"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oMat04.AutoResizeColumns
            oMat05.AutoResizeColumns
            oMat06.AutoResizeColumns
            Call oForm01.Freeze(False)
        End If
        If pval.ItemUID = "Opt03" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Mat03"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oMat04.AutoResizeColumns
            oMat05.AutoResizeColumns
            oMat06.AutoResizeColumns
            Call oForm01.Freeze(False)
        End If
        If pval.ItemUID = "Opt04" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Mat04"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oMat04.AutoResizeColumns
            oMat05.AutoResizeColumns
            oMat06.AutoResizeColumns
            Call oForm01.Freeze(False)
        End If
        If pval.ItemUID = "Opt05" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Mat05"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oMat04.AutoResizeColumns
            oMat05.AutoResizeColumns
            oMat06.AutoResizeColumns
            Call oForm01.Freeze(False)
        End If
        If pval.ItemUID = "Opt06" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Mat06"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oMat04.AutoResizeColumns
            oMat05.AutoResizeColumns
            oMat06.AutoResizeColumns
            Call oForm01.Freeze(False)
        End If
    
        If pval.ItemUID = "Mat01" Then

            If pval.Row > 0 Then

                Call oMat01.SelectRow(pval.Row, True, False)

            End If
        
        ElseIf pval.ItemUID = "Mat02" Then

            If pval.Row > 0 Then

                Call oMat02.SelectRow(pval.Row, True, False)

            End If

        ElseIf pval.ItemUID = "Mat03" Then

            If pval.Row > 0 Then

                Call oMat03.SelectRow(pval.Row, True, False)

            End If
        
        ElseIf pval.ItemUID = "Mat04" Then

            If pval.Row > 0 Then

                Call oMat04.SelectRow(pval.Row, True, False)

            End If
        
        ElseIf pval.ItemUID = "Mat05" Then

            If pval.Row > 0 Then

                Call oMat05.SelectRow(pval.Row, True, False)

            End If
        
        ElseIf pval.ItemUID = "Mat06" Then

            If pval.Row > 0 Then

                Call oMat06.SelectRow(pval.Row, True, False)

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

    If pval.Before_Action = True Then
    
    ElseIf pval.Before_Action = False Then
    
        If pval.ItemChanged = True Then
            Call PS_PP985_FlushToItemValue(pval.ItemUID)
        End If
    
    End If

    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row = 0 Then
            
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource
            
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
    
        If pval.ItemChanged = True Then

            Call PS_PP985_FlushToItemValue(pval.ItemUID)
                
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
        Call PS_PP985_FormItemEnabled
        '//Call PS_PP985_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP985_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP985H", "U_CardCode,U_CardName")
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
'            If (PS_PP985_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_PP985H.RemoveRecord(oDS_PS_PP985H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_PP985_Add_MatrixRow01(0)
            Else
                If Trim(oDS_PS_PP985H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_PP985_Add_MatrixRow01(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP985_CreateItems() As Boolean
On Error GoTo PS_PP985_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_PP985H = oForm01.DataSources.DBDataSources("@PS_PP985H")
    Set oDS_PS_PP985A = oForm01.DataSources.DBDataSources("@PS_USERDS06")
    Set oDS_PS_PP985B = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    Set oDS_PS_PP985C = oForm01.DataSources.DBDataSources("@PS_USERDS03")
    Set oDS_PS_PP985D = oForm01.DataSources.DBDataSources("@PS_USERDS04")
    Set oDS_PS_PP985E = oForm01.DataSources.DBDataSources("@PS_USERDS05")
    Set oDS_PS_PP985F = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    
    '//////////매트릭스 개체 할당//////////S
    '자재비 매트릭스
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_Auto
    oMat01.AutoResizeColumns
    '자체가공비 매트릭스
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_Auto
    oMat02.AutoResizeColumns
    '외주가공비 매트릭스
    Set oMat03 = oForm01.Items("Mat03").Specific
    oMat03.SelectionMode = ms_Auto
    oMat03.AutoResizeColumns
    '외주제작비 매트릭스
    Set oMat04 = oForm01.Items("Mat04").Specific
    oMat04.SelectionMode = ms_Auto
    oMat04.AutoResizeColumns
    '설계비 매트릭스
    Set oMat05 = oForm01.Items("Mat05").Specific
    oMat05.SelectionMode = ms_Auto
    oMat05.AutoResizeColumns
    '기성매출 매트릭스
    Set oMat06 = oForm01.Items("Mat06").Specific
    oMat06.SelectionMode = ms_Auto
    oMat06.AutoResizeColumns
    '//////////매트릭스 개체 할당//////////E
    
    '//////////각 매트릭스 서식세팅 선택용 라디오버튼//////////S
    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt01").Specific.DataBind.SetBound(True, "", "Opt01")
    
    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt02").Specific.DataBind.SetBound(True, "", "Opt02")
    
    Call oForm01.DataSources.UserDataSources.Add("Opt03", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt03").Specific.DataBind.SetBound(True, "", "Opt03")
    
    Call oForm01.DataSources.UserDataSources.Add("Opt04", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt04").Specific.DataBind.SetBound(True, "", "Opt04")
    
    Call oForm01.DataSources.UserDataSources.Add("Opt05", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt05").Specific.DataBind.SetBound(True, "", "Opt05")
    
    Call oForm01.DataSources.UserDataSources.Add("Opt06", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt06").Specific.DataBind.SetBound(True, "", "Opt06")
    
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt03")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt04")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt05")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt06")
    '//////////각 매트릭스 서식세팅 선택용 라디오버튼//////////E
    
    '//////////조회//////////S
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdNum").Specific.DataBind.SetBound(True, "", "OrdNum")
    
    '서브작번1
    Call oForm01.DataSources.UserDataSources.Add("SubNo01", dt_SHORT_TEXT, 2)
    Call oForm01.Items("SubNo1").Specific.DataBind.SetBound(True, "", "SubNo01")
    
    '서브작번2
    Call oForm01.DataSources.UserDataSources.Add("SubNo02", dt_SHORT_TEXT, 3)
    Call oForm01.Items("SubNo2").Specific.DataBind.SetBound(True, "", "SubNo02")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("Spec", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Spec").Specific.DataBind.SetBound(True, "", "Spec")
    
    '기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    
    '기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    '//////////조회//////////E
    
    '//////////조회 결과//////////S
    '수주금액
    Call oForm01.DataSources.UserDataSources.Add("SjAmt", dt_SUM)
    Call oForm01.Items("SjAmt").Specific.DataBind.SetBound(True, "", "SjAmt")
    
    '자재비
    Call oForm01.DataSources.UserDataSources.Add("MatAmt", dt_SUM)
    Call oForm01.Items("MatAmt").Specific.DataBind.SetBound(True, "", "MatAmt")
    
    '자체가공비
    Call oForm01.DataSources.UserDataSources.Add("GagongAmt", dt_SUM)
    Call oForm01.Items("GagongAmt").Specific.DataBind.SetBound(True, "", "GagongAmt")
    
    '외주가공비
    Call oForm01.DataSources.UserDataSources.Add("OutgAmt", dt_SUM)
    Call oForm01.Items("OutgAmt").Specific.DataBind.SetBound(True, "", "OutgAmt")
    
    '외주제작비
    Call oForm01.DataSources.UserDataSources.Add("OutmAmt", dt_SUM)
    Call oForm01.Items("OutmAmt").Specific.DataBind.SetBound(True, "", "OutmAmt")
    
    '설계비
    Call oForm01.DataSources.UserDataSources.Add("DrawAmt", dt_SUM)
    Call oForm01.Items("DrawAmt").Specific.DataBind.SetBound(True, "", "DrawAmt")
    
    '계
    Call oForm01.DataSources.UserDataSources.Add("Total", dt_SUM)
    Call oForm01.Items("Total").Specific.DataBind.SetBound(True, "", "Total")
    
    '생산완료여부
    Call oForm01.DataSources.UserDataSources.Add("PCYN", dt_SHORT_TEXT, 10)
    Call oForm01.Items("PCYN").Specific.DataBind.SetBound(True, "", "PCYN")
    
    '생산완료일자
    Call oForm01.DataSources.UserDataSources.Add("EndDate", dt_SHORT_TEXT, 10)
    Call oForm01.Items("EndDate").Specific.DataBind.SetBound(True, "", "EndDate")
    
    '수주일
    Call oForm01.DataSources.UserDataSources.Add("SjDocDate", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SjDocDate").Specific.DataBind.SetBound(True, "", "SjDocDate")
    
    '작업지시일
    Call oForm01.DataSources.UserDataSources.Add("WODate", dt_SHORT_TEXT, 10)
    Call oForm01.Items("WODate").Specific.DataBind.SetBound(True, "", "WODate")
    
    '납기일
    Call oForm01.DataSources.UserDataSources.Add("SjDueDate", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SjDueDate").Specific.DataBind.SetBound(True, "", "SjDueDate")
    
    '협력업체(외주제작)
    Call oForm01.DataSources.UserDataSources.Add("OtCardNm", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OtCardNm").Specific.DataBind.SetBound(True, "", "OtCardNm")
    '//////////조회 결과//////////E
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_PP985_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP985_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_PP985_ComboBox_Setting()
On Error GoTo PS_PP985_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '사업장
'    Call oForm01.Items("BPLID").Specific.ValidValues.Add("%", "전체")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
'
'    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '거래처구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'C100'"
    Call oForm01.Items("CardType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType").Specific, sQry, "", False, False)
    Call oForm01.Items("CardType").Specific.Select("%", psk_ByValue)
    
    '품목구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'S002'"
    Call oForm01.Items("ItemType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType").Specific, sQry, "", False, False)
    Call oForm01.Items("ItemType").Specific.Select("%", psk_ByValue)
    
    '생산완료여부
    Call oForm01.Items("WCYN").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("WCYN").Specific.ValidValues.Add("B", "미완료")
    Call oForm01.Items("WCYN").Specific.ValidValues.Add("C", "완료")
    Call oForm01.Items("WCYN").Specific.Select("%", psk_ByValue)
    
    '소속반
    sQry = "           SELECT      T0.U_Code AS [Code],"
    sQry = sQry & "                T0.U_CodeNm AS [Value]"
    sQry = sQry & " FROM       [@PS_HR200L] AS T0"
    sQry = sQry & " WHERE      T0.Code = '9'"
    sQry = sQry & "                AND T0.U_Char3 = '2'"
    sQry = sQry & "                AND T0.U_Char2 = '2800'"
    sQry = sQry & "                AND T0.U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  T0.U_Seq"
    Call oForm01.Items("ClsCode").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ClsCode").Specific, sQry, "", False, False)
    Call oForm01.Items("ClsCode").Specific.Select("%", psk_ByValue)
       
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_PP985_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP985_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP985_CF_ChooseFromList()
On Error GoTo PS_PP985_CF_ChooseFromList_Error

    Exit Sub
PS_PP985_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP985_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP985_EnableMenus()
On Error GoTo PS_PP985_EnableMenus_Error
    
    Exit Sub
PS_PP985_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP985_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP985_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP985_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP985_FormItemEnabled
        '//Call PS_PP985_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP985_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP985_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP985_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP985_FormResize()
On Error GoTo PS_PP985_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_PP985_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP985_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP985_Print_Report01()
On Error GoTo Print_Query_Error
    Dim DocNum As String
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String
    Dim FrDt As String
    Dim ToDt As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim CntcCode As String '사업장
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP985] 작번별 진행금액 일괄조회"
    ReportName = "PS_PP985_01.rpt"
 
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)

    '// SubReport
    ReDim gRpt_SRptSqry(6)
    ReDim gRpt_SRptName(6)
    
    ReDim gRpt_SFormula(6, 1)
    ReDim gRpt_SFormula_Value(6, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    gRpt_SFormula(2, 1) = ""
    gRpt_SFormula_Value(2, 1) = ""
    
    gRpt_SFormula(3, 1) = ""
    gRpt_SFormula_Value(3, 1) = ""
    
    gRpt_SFormula(4, 1) = ""
    gRpt_SFormula_Value(4, 1) = ""
    
    gRpt_SFormula(5, 1) = ""
    gRpt_SFormula_Value(5, 1) = ""
    
    gRpt_SFormula(6, 1) = ""
    gRpt_SFormula_Value(6, 1) = ""
    
    'Sub_Procedure 실행(자재비)
    Sub_sQry = "EXEC [PS_PP985_52] '" & CntcCode & "'"
    gRpt_SRptSqry(1) = Sub_sQry
    gRpt_SRptName(1) = "PS_PP985_SUB_01"
    
    'Sub_Procedure 실행(자체가공비)
    Sub_sQry = "EXEC [PS_PP985_53] '" & CntcCode & "'"
    gRpt_SRptSqry(2) = Sub_sQry
    gRpt_SRptName(2) = "PS_PP985_SUB_02"

    'Sub_Procedure 실행(외주가공비)
    Sub_sQry = "EXEC [PS_PP985_54] '" & CntcCode & "'"
    gRpt_SRptSqry(3) = Sub_sQry
    gRpt_SRptName(3) = "PS_PP985_SUB_03"

    'Sub_Procedure 실행(외주제작비)
    Sub_sQry = "EXEC [PS_PP985_55] '" & CntcCode & "'"
    gRpt_SRptSqry(4) = Sub_sQry
    gRpt_SRptName(4) = "PS_PP985_SUB_04"
    
    'Sub_Procedure 실행(설계비)
    Sub_sQry = "EXEC [PS_PP985_56] '" & CntcCode & "'"
    gRpt_SRptSqry(5) = Sub_sQry
    gRpt_SRptName(5) = "PS_PP985_SUB_05"
    
    'Sub_Procedure 실행(기성매출)
    Sub_sQry = "EXEC [PS_PP985_57] '" & CntcCode & "'"
    gRpt_SRptSqry(6) = Sub_sQry
    gRpt_SRptName(6) = "PS_PP985_SUB_06"
    
    '/ Procedure 실행"
    sQry = "EXEC [PS_PP985_51] '"
    sQry = sQry & CntcCode & "','"
    sQry = sQry & FrDt & "','"
    sQry = sQry & ToDt & "'"

'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "", "N", "V") = False Then
        GoTo Print_Query_Error
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    ProgBar01.VALUE = 100
    ProgBar01.Stop
     Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_PP985_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub












