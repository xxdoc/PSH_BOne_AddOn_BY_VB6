VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP989"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_PP989.cls
'//  Module             : 생산관리>풍산홀딩스애드온>생산리포트
'//  Desc               : 기간별 표준공수 대비 실동공수 조회
'//  FormType           : PS_PP989
'//  Create Date(Start) : 2014.08.25
'//  Create Date(End)   : 2014.08.26
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01  As String
Public oForm01          As SAPbouiCOM.Form

Public oMat02           As SAPbouiCOM.Matrix

Private oDS_PS_PP989O   As SAPbouiCOM.DBDataSource

Public oGrid01          As SAPbouiCOM.Grid

Public oDS_PS_PP989L    As SAPbouiCOM.DataTable

Private oLastItemUID01  As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01   As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01   As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
'    Dim ProgBar01 As SAPbouiCOM.ProgressBar
'    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("로딩 중...", 100, False)
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP989.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP989_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

    oForm01.Freeze True
    
    Call PS_PP989_CreateItems
    Call PS_PP989_ComboBox_Setting
    Call PS_PP989_CF_ChooseFromList
    Call PS_PP989_FormItemEnabled
    Call PS_PP989_EnableMenus
    Call PS_PP989_SetDocument(oFromDocEntry01)
    Call PS_PP989_FormResize

'    Call PS_PP989_Add_MatrixRow01(0, True)
    
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
    
    '기간
    oForm01.Items("FrDt01").Specific.VALUE = Format(Now, "YYYYMM01")
    oForm01.Items("ToDt01").Specific.VALUE = Format(Now, "YYYYMMDD")

    oForm01.Items("FrDt02").Specific.VALUE = Format(Now, "YYYYMM01")
    oForm01.Items("ToDt02").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oForm01.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
'    ProgBar01.VALUE = 100
'    ProgBar01.Stop
'    Set ProgBar01 = Nothing
    
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    
'    ProgBar01.VALUE = 100
'    ProgBar01.Stop
'    Set ProgBar01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

'/메트릭스 Row추가
Sub PS_PP989_Add_MatrixRow01(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP989_Add_MatrixRow01_Error

    If RowIserted = False Then '//행추가여부
        oDS_PS_PP989O.InsertRecord (oRow)
    End If
    
    oMat02.AddRow
    oDS_PS_PP989O.offset = oRow
    oDS_PS_PP989O.setValue "U_LineNum", oRow, oRow + 1
    
    oMat02.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP989_Add_MatrixRow01_Error:
    MDC_Com.MDC_GF_Message "PS_PP989_Add_MatrixRow01_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_PP989_MTX01()
On Error GoTo PS_PP989_MTX01_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i       As Long
    Dim Query01 As String
    Dim ErrNum  As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID    As String
    Dim FrDt     As String
    Dim ToDt     As String
    Dim OrdGbn   As String
    Dim OrdNum   As String '작번
    Dim OrdSub1  As String '서브작번1
    Dim OrdSub2  As String '서브작번2
    Dim CntcCode As String '담당
    Dim CardCode As String '거래처
    
    BPLID = oForm01.Items("BPLID01").Specific.VALUE
    FrDt = oForm01.Items("FrDt01").Specific.VALUE
    ToDt = oForm01.Items("ToDt01").Specific.VALUE
    OrdGbn = oForm01.Items("OrdGbn01").Specific.VALUE
    OrdNum = oForm01.Items("OrdNum01").Specific.VALUE
    OrdSub1 = oForm01.Items("OrdSub101").Specific.VALUE
    OrdSub2 = oForm01.Items("OrdSub201").Specific.VALUE
    CntcCode = oForm01.Items("CntcCode01").Specific.VALUE
    CardCode = oForm01.Items("CardCode01").Specific.VALUE
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
        
    Query01 = "         EXEC PS_PP989_01 '"
    Query01 = Query01 & BPLID & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "','"
    Query01 = Query01 & OrdGbn & "','"
    Query01 = Query01 & OrdNum & "','"
    Query01 = Query01 & OrdSub1 & "','"
    Query01 = Query01 & OrdSub2 & "','"
    Query01 = Query01 & CntcCode & "','"
    Query01 = Query01 & CardCode & "'"
    
    Call oGrid01.DataTable.Clear
    Call oDS_PS_PP989L.ExecuteQuery(Query01)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    oGrid01.Columns(16).RightJustified = True
    oGrid01.Columns(17).RightJustified = True
'    oGrid01.Columns(9).RightJustified = True
'
'    oGrid01.Columns(10).RightJustified = True
'    oGrid01.Columns(12).RightJustified = True
'    oGrid01.Columns(13).RightJustified = True
'    oGrid01.Columns(15).RightJustified = True
'    oGrid01.Columns(16).RightJustified = True
'    oGrid01.Columns(18).RightJustified = True
'    oGrid01.Columns(19).RightJustified = True
'    oGrid01.Columns(21).RightJustified = True
'    oGrid01.Columns(22).RightJustified = True
'    oGrid01.Columns(24).RightJustified = True
'    oGrid01.Columns(25).RightJustified = True
'    oGrid01.Columns(27).RightJustified = True
'    oGrid01.Columns(28).RightJustified = True
'    oGrid01.Columns(30).RightJustified = True
'    oGrid01.Columns(31).RightJustified = True
'    oGrid01.Columns(33).RightJustified = True
'    oGrid01.Columns(34).RightJustified = True
'    oGrid01.Columns(36).RightJustified = True
'    oGrid01.Columns(37).RightJustified = True
'    oGrid01.Columns(39).RightJustified = True
'    oGrid01.Columns(40).RightJustified = True
'    oGrid01.Columns(42).RightJustified = True
'    oGrid01.Columns(43).RightJustified = True
'
'    oGrid01.Columns(45).RightJustified = True
'    oGrid01.Columns(46).RightJustified = True
'    oGrid01.Columns(48).RightJustified = True
'    oGrid01.Columns(49).RightJustified = True
'    oGrid01.Columns(51).RightJustified = True
'    oGrid01.Columns(52).RightJustified = True
'    oGrid01.Columns(54).RightJustified = True
'    oGrid01.Columns(55).RightJustified = True
'    oGrid01.Columns(57).RightJustified = True
'    oGrid01.Columns(58).RightJustified = True
'    oGrid01.Columns(60).RightJustified = True
'    oGrid01.Columns(61).RightJustified = True
'    oGrid01.Columns(63).RightJustified = True
'    oGrid01.Columns(64).RightJustified = True
'    oGrid01.Columns(66).RightJustified = True
'    oGrid01.Columns(67).RightJustified = True
'    oGrid01.Columns(69).RightJustified = True
'    oGrid01.Columns(70).RightJustified = True
'    oGrid01.Columns(72).RightJustified = True
'    oGrid01.Columns(73).RightJustified = True
'    oGrid01.Columns(75).RightJustified = True
'    oGrid01.Columns(76).RightJustified = True
'    oGrid01.Columns(78).RightJustified = True
'    oGrid01.Columns(79).RightJustified = True
'    oGrid01.Columns(81).RightJustified = True
'    oGrid01.Columns(82).RightJustified = True
'
'    oGrid01.Columns(84).RightJustified = True
'    oGrid01.Columns(85).RightJustified = True
'    oGrid01.Columns(87).RightJustified = True
'    oGrid01.Columns(88).RightJustified = True
'    oGrid01.Columns(90).RightJustified = True
'    oGrid01.Columns(91).RightJustified = True
'    oGrid01.Columns(93).RightJustified = True
'    oGrid01.Columns(94).RightJustified = True
'    oGrid01.Columns(96).RightJustified = True
'    oGrid01.Columns(97).RightJustified = True
'    oGrid01.Columns(99).RightJustified = True
'    oGrid01.Columns(100).RightJustified = True
'    oGrid01.Columns(102).RightJustified = True
'    oGrid01.Columns(103).RightJustified = True
'    oGrid01.Columns(105).RightJustified = True
'    oGrid01.Columns(106).RightJustified = True
'    oGrid01.Columns(108).RightJustified = True
'    oGrid01.Columns(109).RightJustified = True
'    oGrid01.Columns(111).RightJustified = True
'    oGrid01.Columns(112).RightJustified = True
'    oGrid01.Columns(114).RightJustified = True
'    oGrid01.Columns(115).RightJustified = True
'    oGrid01.Columns(117).RightJustified = True
'    oGrid01.Columns(118).RightJustified = True
'    oGrid01.Columns(120).RightJustified = True
'    oGrid01.Columns(121).RightJustified = True
'
'    oGrid01.Columns(123).RightJustified = True
'    oGrid01.Columns(124).RightJustified = True
'    oGrid01.Columns(126).RightJustified = True
'    oGrid01.Columns(127).RightJustified = True
'    oGrid01.Columns(129).RightJustified = True
'    oGrid01.Columns(130).RightJustified = True
'    oGrid01.Columns(132).RightJustified = True
'    oGrid01.Columns(133).RightJustified = True
'    oGrid01.Columns(135).RightJustified = True
'    oGrid01.Columns(136).RightJustified = True
'    oGrid01.Columns(138).RightJustified = True
'    oGrid01.Columns(139).RightJustified = True
'    oGrid01.Columns(141).RightJustified = True
'    oGrid01.Columns(142).RightJustified = True

'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid01.Rows.Count = 1 Then
        ErrNum = 1
        GoTo PS_PP989_MTX01_Error:
    End If
        
    oGrid01.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP989_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP989_MTX01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    If ErrNum = 1 Then
    
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
    
    Else

        Sbo_Application.SetStatusBarMessage "PS_PP989_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
    End If
    
    Set RecordSet01 = Nothing
    
    Call oForm01.Freeze(False)
    
End Sub

Public Sub PS_PP989_MTX02()
'******************************************************************************
'Function ID : PS_PP989_MTX02()
'해당모듈    : PS_PP989
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP989_MTX02_Error

    Dim i       As Integer
    Dim sQry    As String
    
    Dim ErrNum  As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID     As String '사업장
    Dim TeamCode  As String '팀
    Dim RspCode   As String '담당
    Dim ClsCode   As String '반
    Dim CardType  As String '거래처구분
    Dim ItemType  As String '품목구분
    Dim WCYN      As String '생산완료여부
    Dim DateStd   As String '일자기준
    Dim FrDt      As String '기간(Fr)
    Dim ToDt      As String '기간(To)
    Dim ItemCode  As String '품목코드(작번)
    Dim CpCode    As String '공정
    Dim OrdGbn    As String '작업구분
    Dim CntcCode  As String '조회자 사번
    
    Dim TStdTime  As Double '표준공수계(시간)
    Dim TStdAmt   As Currency '표준공수계(금액)
    Dim TWkTime   As Double '실동공수계(시간)
    Dim TWkAmt    As Currency '실동공수계(금액)
    
    BPLID = Trim(oForm01.Items("BPLID02").Specific.VALUE)  '사업장
    TeamCode = Trim(oForm01.Items("TeamCode02").Specific.VALUE)  '팀
    RspCode = Trim(oForm01.Items("RspCode02").Specific.VALUE)  '담당
    ClsCode = Trim(oForm01.Items("ClsCode02").Specific.VALUE)  '반
    CardType = Trim(oForm01.Items("CardType02").Specific.VALUE)  '거래처구분
    ItemType = Trim(oForm01.Items("ItemType02").Specific.VALUE)  '품목구분
    WCYN = Trim(oForm01.Items("WCYN02").Specific.VALUE)  '생산완료여부
    DateStd = Trim(oForm01.Items("DateStd02").Specific.VALUE)  '일자기준
    FrDt = Trim(oForm01.Items("FrDt02").Specific.VALUE)  '기간(Fr)
    ToDt = Trim(oForm01.Items("ToDt02").Specific.VALUE)  '기간(To)
    ItemCode = Trim(oForm01.Items("ItemCode02").Specific.VALUE)  '품목코드(작번)
    CpCode = Trim(oForm01.Items("CpCode02").Specific.VALUE)  '공정
    OrdGbn = Trim(oForm01.Items("OrdGbn02").Specific.VALUE)  '작업구분
    CntcCode = MDC_PS_Common.User_MSTCOD '조회자 사번
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "      EXEC [PS_PP989_02] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & TeamCode & "','" '팀
    sQry = sQry & RspCode & "','" '담당
    sQry = sQry & ClsCode & "','" '반
    sQry = sQry & CardType & "','" '거래처구분
    sQry = sQry & ItemType & "','" '품목구분
    sQry = sQry & WCYN & "','" '생산완료여부
    sQry = sQry & DateStd & "','" '일자기준
    sQry = sQry & FrDt & "','" '기간(Fr)
    sQry = sQry & ToDt & "','" '기간(To)
    sQry = sQry & ItemCode & "','" '품목코드(작번)
    sQry = sQry & CpCode & "','" '공정
    sQry = sQry & OrdGbn & "','" '작업구분
    sQry = sQry & CntcCode & "'" '조회자사번
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat02.Clear
    oDS_PS_PP989O.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
        Call PS_PP989_Add_MatrixRow01(0, True)
        GoTo PS_PP989_MTX02_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
    
        If i + 1 > oDS_PS_PP989O.Size Then
            oDS_PS_PP989O.InsertRecord (i)
        End If
    
        oMat02.AddRow
        oDS_PS_PP989O.offset = i
    
        Call oDS_PS_PP989O.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP989O.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("CpCode").VALUE)) '공정코드
        Call oDS_PS_PP989O.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("CpName").VALUE)) '공정명
        Call oDS_PS_PP989O.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("DocDate").VALUE)) '작업일자
        Call oDS_PS_PP989O.setValue("U_ColPrc01", i, Trim(oRecordSet01.Fields("Price").VALUE)) '공정단가
        Call oDS_PS_PP989O.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("StdTime").VALUE)) '표준공수
        Call oDS_PS_PP989O.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("StdAmt").VALUE)) '표준공수금액
        Call oDS_PS_PP989O.setValue("U_ColQty02", i, Trim(oRecordSet01.Fields("WkTime").VALUE)) '실동공수
        Call oDS_PS_PP989O.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("WkAmt").VALUE)) '실동공수금액
        
        TStdTime = TStdTime + IIf(Trim(oRecordSet01.Fields("DocDate").VALUE) <> "[소 계]", IIf(Trim(oRecordSet01.Fields("DocDate").VALUE) <> "[합 계]", Trim(oRecordSet01.Fields("StdTime").VALUE), 0), 0) '표준공수계(시간)
        TStdAmt = TStdAmt + IIf(Trim(oRecordSet01.Fields("DocDate").VALUE) <> "[소 계]", IIf(Trim(oRecordSet01.Fields("DocDate").VALUE) <> "[합 계]", Trim(oRecordSet01.Fields("StdAmt").VALUE), 0), 0) '표준공수계(금액)
        TWkTime = TWkTime + IIf(Trim(oRecordSet01.Fields("DocDate").VALUE) <> "[소 계]", IIf(Trim(oRecordSet01.Fields("DocDate").VALUE) <> "[합 계]", Trim(oRecordSet01.Fields("WkTime").VALUE), 0), 0) '실동공수계(시간)
        TWkAmt = TWkAmt + IIf(Trim(oRecordSet01.Fields("DocDate").VALUE) <> "[소 계]", IIf(Trim(oRecordSet01.Fields("DocDate").VALUE) <> "[합 계]", Trim(oRecordSet01.Fields("WkAmt").VALUE), 0), 0) '실동공수계(금액)
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    
    oForm01.Items("TStdTime02").Specific.VALUE = TStdTime '표준공수계(시간)
    oForm01.Items("TStdAmt02").Specific.VALUE = TStdAmt '표준공수계(금액)
    oForm01.Items("TWkTime02").Specific.VALUE = TWkTime '실동공수계(시간)
    oForm01.Items("TWkAmt02").Specific.VALUE = TWkAmt '실동공수계(금액)
    
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP989_MTX02_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP989_MTX02_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Function PS_PP989_HeaderSpaceLineDel02() As Boolean
'******************************************************************************
'Function ID : PS_PP989_HeaderSpaceLineDel02()
'해당모듈    : PS_PP989
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP989_HeaderSpaceLineDel02_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
'        Case Trim(oForm01.Items("StdYear").Specific.VALUE) = "" '기준년도
'            ErrNum = 1
'            GoTo PS_PP989_HeaderSpaceLineDel02_Error
    End Select

    PS_PP989_HeaderSpaceLineDel02 = True
Exit Function

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP989_HeaderSpaceLineDel02_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "기준년도는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("FrDate").Click(ct_Regular)
    End If
    PS_PP989_HeaderSpaceLineDel02 = False
End Function

''/ 메트릭스 필수 사항 check
'Private Function PS_PP989_MatrixSpaceLineDel() As Boolean
'On Error GoTo PS_PP989_MatrixSpaceLineDel_Error
'
'    Dim i As Long
'    Dim ErrNum As Integer
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Dim sQry As String
'
'    Set oRecordSet01 = Nothing
'    PS_PP989_MatrixSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_PP989_MatrixSpaceLineDel_Error:
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
'        MDC_Com.MDC_GF_Message "PS_PP989_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'    PS_PP989_MatrixSpaceLineDel = False
'End Function

Private Sub PS_PP989_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_PP989_FlushToItemValue_Error

    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim sQry     As String
    Dim ItemCode As String
    Dim Qty      As Integer
    
    Dim BPLID    As String
    Dim TeamCode As String
    Dim RspCode  As String
    
    Dim OrdNum   As String
    Dim SubNo1   As String
    Dim SubNo2   As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
    
        Case "BPLID02"
            
            BPLID = Trim(oForm01.Items("BPLID02").Specific.VALUE)
            
            If oForm01.Items("TeamCode02").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("TeamCode02").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("TeamCode02").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '부서콤보세팅
            Call oForm01.Items("TeamCode02").Specific.ValidValues.Add("%", "전체")
            sQry = "        SELECT      U_Code AS [Code],"
            sQry = sQry & "             U_CodeNm As [Name]"
            sQry = sQry & " FROM        [@PS_HR200L]"
            sQry = sQry & " WHERE       Code = '1'"
            sQry = sQry & "             AND U_UseYN = 'Y'"
            sQry = sQry & "             AND U_Char2 = '" & BPLID & "'"
            sQry = sQry & " ORDER BY    U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("TeamCode02").Specific, sQry, "", False, False)
            Call oForm01.Items("TeamCode02").Specific.Select(0, psk_Index)
            
'            Call oForm01.Items("RspCode02").Specific.ValidValues.Add("%", "전체") '담당 전체 추가
'            Call oForm01.Items("RspCode02").Specific.Select(0, psk_Index)
'            Call oForm01.Items("ClsCode02").Specific.ValidValues.Add("%", "전체") '반 전체 추가
'            Call oForm01.Items("ClsCode02").Specific.Select(0, psk_Index)
            
        Case "TeamCode02"
            
            TeamCode = Trim(oForm01.Items("TeamCode02").Specific.VALUE)
            
            If oForm01.Items("RspCode02").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("RspCode02").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("RspCode02").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '담당콤보세팅
            Call oForm01.Items("RspCode02").Specific.ValidValues.Add("%", "전체")
            sQry = "        SELECT      U_Code AS [Code],"
            sQry = sQry & "             U_CodeNm As [Name]"
            sQry = sQry & " FROM        [@PS_HR200L]"
            sQry = sQry & " WHERE       Code = '2'"
            sQry = sQry & "             AND U_UseYN = 'Y'"
            sQry = sQry & "             AND U_Char1 = '" & TeamCode & "'"
            sQry = sQry & " ORDER BY    U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("RspCode02").Specific, sQry, "", False, False)
            Call oForm01.Items("RspCode02").Specific.Select(0, psk_Index)
        
        Case "RspCode02"
            
            TeamCode = Trim(oForm01.Items("TeamCode02").Specific.VALUE)
            RspCode = Trim(oForm01.Items("RspCode02").Specific.VALUE)
            
            If oForm01.Items("ClsCode02").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("ClsCode02").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("ClsCode02").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '반콤보세팅
            Call oForm01.Items("ClsCode02").Specific.ValidValues.Add("%", "전체")
            sQry = "        SELECT      U_Code AS [Code],"
            sQry = sQry & "             U_CodeNm As [Name]"
            sQry = sQry & " FROM        [@PS_HR200L]"
            sQry = sQry & " WHERE       Code = '9'"
            sQry = sQry & "             AND U_UseYN = 'Y'"
            sQry = sQry & "             AND U_Char1 = '" & RspCode & "'"
            sQry = sQry & "             AND U_Char2 = '" & TeamCode & "'"
            sQry = sQry & " ORDER BY    U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("ClsCode02").Specific, sQry, "", False, False)
            Call oForm01.Items("ClsCode02").Specific.Select(0, psk_Index)
    
        Case "ItemCode02"
            
            sQry = "        SELECT      FrgnName,"
            sQry = sQry & "             U_Size"
            sQry = sQry & " FROM        OITM"
            sQry = sQry & " WHERE       ItemCode = '" & Trim(oForm01.Items("ItemCode02").Specific.VALUE) & "'"
            Call oRecordSet01.DoQuery(sQry)

            oForm01.Items("ItemName02").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            oForm01.Items("ItemSpec02").Specific.VALUE = Trim(oRecordSet01.Fields(1).VALUE)
            
        Case "CpCode02"
        
            sQry = "        SELECT      U_CpName"
            sQry = sQry & " FROM        [@PS_PP001L]"
            sQry = sQry & " WHERE       U_CpCode = '" & Trim(oForm01.Items("CpCode02").Specific.VALUE) & "'"
            Call oRecordSet01.DoQuery(sQry)

            oForm01.Items("CpName02").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "CntcCode01"
        
            sQry = "        SELECT      U_FullName"
            sQry = sQry & " FROM        [@PH_PY001A]"
            sQry = sQry & " WHERE       Code = '" & Trim(oForm01.Items("CntcCode01").Specific.VALUE) & "'"
            Call oRecordSet01.DoQuery(sQry)

            oForm01.Items("CntcName01").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
        Case "CardCode01"
        
            sQry = "        SELECT      CardName"
            sQry = sQry & " FROM        [OCRD]"
            sQry = sQry & " WHERE       CardCode = '" & Trim(oForm01.Items("CardCode01").Specific.VALUE) & "'"
            Call oRecordSet01.DoQuery(sQry)

            oForm01.Items("CardName01").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
    End Select


    If oUID = "OrdNum01" Or oUID = "OrdSub101" Or oUID = "OrdSub201" Then
    
        OrdNum = oForm01.Items("OrdNum01").Specific.VALUE
        SubNo1 = oForm01.Items("OrdSub101").Specific.VALUE
        SubNo2 = oForm01.Items("OrdSub201").Specific.VALUE

        sQry = "        SELECT      CASE"
        sQry = sQry & "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakMyung"
        sQry = sQry & "             END AS [ItemName],"
        sQry = sQry & "             CASE"
        sQry = sQry & "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakSize"
        sQry = sQry & "             END AS [SPEC]"
        sQry = sQry & " FROM        [@PS_PP020H] AS T0"
        sQry = sQry & " WHERE       T0.U_JakName = '" & OrdNum & "'"
        sQry = sQry & "             AND T0.U_SubNo1 = CASE WHEN '" & SubNo1 & "' = '' THEN '00' ELSE '" & SubNo1 & "' END"
        sQry = sQry & "             AND T0.U_SubNo2 = CASE WHEN '" & SubNo2 & "' = '' THEN '000' ELSE '" & SubNo2 & "' END"

        Call oRecordSet01.DoQuery(sQry)

        oForm01.Items("ItemName01").Specific.VALUE = oRecordSet01.Fields("ItemName").VALUE
'        oForm01.Items("Spec").Specific.VALUE = oRecordSet01.Fields("Spec").VALUE
    
    End If

    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_PP989_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("시각은 숫자만 입력이 가능합니다.", "E")
'    ElseIf ErrNum = 2 Then
'        Call MDC_Com.MDC_GF_Message("시각(시)는 24미만의 값만 입력이 가능합니다.", "E")
'    ElseIf ErrNum = 3 Then
'        Call MDC_Com.MDC_GF_Message("시각(분)은 60미만의 값만 입력이 가능합니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP989_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_PP989_FormItemEnabled()
On Error GoTo PS_PP989_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat02.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True

        Call oForm01.Items("BPLID02").Specific.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
        Call PS_PP989_FlushToItemValue("BPLID02") '팀, 담당, 반 콤보박스 강제 설정

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP989_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_PP989_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
'                oMat02.Clear
'                oMat02.FlushToDataSource
'                oMat02.LoadFromDataSource
                
'                oForm01.Mode = fm_ADD_MODE
'                BubbleEvent = False
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm01.Freeze(True)
                Call PS_PP989_Add_MatrixRow01(oMat02.VisualRowCount)
                Call oForm01.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_PP989_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat02.Clear
'                oDS_PS_PP989H.Clear

'                Call PS_PP989_LoadCaption
'                Call PS_PP989_FormItemEnabled
                '//Call PS_PP989_FormItemEnabled '//UDO방식
                '//Call PS_PP989_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_PP989_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_PP989O.RemoveRecord(oDS_PS_PP989O.Size - 1)
                Call oMat02.LoadFromDataSource
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
    If pval.ItemUID = "Mat02" Then
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
    
        If pval.ItemUID = "PS_PP989" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnSrch01" Then '/조회
            If PS_PP989_HeaderSpaceLineDel02 = False Then
                BubbleEvent = False
                Exit Sub
            Else
                Call PS_PP989_MTX01
            End If
            
        ElseIf pval.ItemUID = "BtnPrt01" Then '출력
            
            Call PS_PP989_Print_Report01
            
        End If
        
        If pval.ItemUID = "BtnSrch02" Then '/조회
            If PS_PP989_HeaderSpaceLineDel02 = False Then
                BubbleEvent = False
                Exit Sub
            Else
                Call PS_PP989_MTX02
            End If
            
        ElseIf pval.ItemUID = "BtnPrt02" Then '출력
            
            Call PS_PP989_Print_Report02
            
        End If
  
    ElseIf pval.BeforeAction = False Then
    
        '폴더를 사용할 때는 필수 소스_S
        'Folder01이 선택되었을 때
        If pval.ItemUID = "Folder01" Then
            
            oForm01.PaneLevel = 1
            oForm01.DefButton = "BtnSrch01"
            
        End If
        
        'Folder02가 선택되었을 때
        If pval.ItemUID = "Folder02" Then
            
            oForm01.PaneLevel = 2
            oForm01.DefButton = "BtnSrch02"
            
        End If
        '폴더를 사용할 때는 필수 소스_E
    
        If pval.ItemUID = "PS_PP989" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode01", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode01", "") '//사용자값활성
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode02", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CpCode02", "") '//사용자값활성
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat02" Then

            If pval.Row > 0 Then

                Call oMat02.SelectRow(pval.Row, True, False)

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
            Call PS_PP989_FlushToItemValue(pval.ItemUID)
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
    
        If pval.ItemUID = "Mat02" Then
        
            If pval.Row = 0 Then
            
                oMat02.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat02.FlushToDataSource
            
            Else
            
                Call PS_PP989_GetDetail(oMat02.Columns("CpCode").Cells(pval.Row).Specific.VALUE, oMat02.Columns("CpName").Cells(pval.Row).Specific.VALUE)
            
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

            Call PS_PP989_FlushToItemValue(pval.ItemUID)
                
'            oMat02.LoadFromDataSource
'            oMat02.AutoResizeColumns
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
'        Call PS_PP989_FormItemEnabled
        '//Call PS_PP989_AddMatrixRow(oMat02.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP989_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP989H", "U_CardCode,U_CardName")
'        End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat02" Then
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
        Set oMat02 = Nothing
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
'            If (PS_PP989_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat02.VisualRowCount
                oMat02.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat02.FlushToDataSource
            Call oDS_PS_PP989L.RemoveRecord(oDS_PS_PP989L.Size - 1)
            oMat02.LoadFromDataSource
            If oMat02.RowCount = 0 Then
                Call PS_PP989_Add_MatrixRow01(0)
            Else
                If Trim(oDS_PS_PP989L.GetValue("U_CntcCode", oMat02.RowCount - 1)) <> "" Then
                    Call PS_PP989_Add_MatrixRow01(oMat02.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP989_CreateItems() As Boolean
On Error GoTo PS_PP989_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    
    Call oForm01.DataSources.DataTables.Add("PS_PP989L")
    
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("PS_PP989L")
    
    Set oDS_PS_PP989L = oForm01.DataSources.DataTables.Item("PS_PP989L")
    
    Set oDS_PS_PP989O = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_Auto
    oMat02.AutoResizeColumns
    
    '//////////작번별 집계//////////_S
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID01", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID01").Specific.DataBind.SetBound(True, "", "BPLID01")

    '기간(Fr)
    Call oForm01.DataSources.UserDataSources.Add("FrDt01", dt_DATE)
    Call oForm01.Items("FrDt01").Specific.DataBind.SetBound(True, "", "FrDt01")

    '기간(To)
    Call oForm01.DataSources.UserDataSources.Add("ToDt01", dt_DATE)
    Call oForm01.Items("ToDt01").Specific.DataBind.SetBound(True, "", "ToDt01")
    
    '작업구분
    Call oForm01.DataSources.UserDataSources.Add("OrdGbn01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdGbn01").Specific.DataBind.SetBound(True, "", "OrdGbn01")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum01", dt_SHORT_TEXT, 50)
    Call oForm01.Items("OrdNum01").Specific.DataBind.SetBound(True, "", "OrdNum01")
    
    '서브작번1
    Call oForm01.DataSources.UserDataSources.Add("OrdSub101", dt_SHORT_TEXT, 2)
    Call oForm01.Items("OrdSub101").Specific.DataBind.SetBound(True, "", "OrdSub101")
    
    '서브작번2
    Call oForm01.DataSources.UserDataSources.Add("OrdSub201", dt_SHORT_TEXT, 3)
    Call oForm01.Items("OrdSub201").Specific.DataBind.SetBound(True, "", "OrdSub201")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("ItemName01", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName01").Specific.DataBind.SetBound(True, "", "ItemName01")
    
    '담당자사번
    Call oForm01.DataSources.UserDataSources.Add("CntcCode01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcCode01").Specific.DataBind.SetBound(True, "", "CntcCode01")
    
    '담당자성명
    Call oForm01.DataSources.UserDataSources.Add("CntcName01", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CntcName01").Specific.DataBind.SetBound(True, "", "CntcName01")
    
    '거래처코드
    Call oForm01.DataSources.UserDataSources.Add("CardCode01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardCode01").Specific.DataBind.SetBound(True, "", "CardCode01")
    
    '거래처명
    Call oForm01.DataSources.UserDataSources.Add("CardName01", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CardName01").Specific.DataBind.SetBound(True, "", "CardName01")
    '//////////작번별 집계//////////_E
    
    '//////////공정별 집계//////////_S
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID02", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID02").Specific.DataBind.SetBound(True, "", "BPLID02")
    
    '팀
    Call oForm01.DataSources.UserDataSources.Add("TeamCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("TeamCode02").Specific.DataBind.SetBound(True, "", "TeamCode02")
    
    '담당
    Call oForm01.DataSources.UserDataSources.Add("RspCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("RspCode02").Specific.DataBind.SetBound(True, "", "RspCode02")
    
    '소속반
    Call oForm01.DataSources.UserDataSources.Add("ClsCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ClsCode02").Specific.DataBind.SetBound(True, "", "ClsCode02")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardType02").Specific.DataBind.SetBound(True, "", "CardType02")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemType02").Specific.DataBind.SetBound(True, "", "ItemType02")
    
    '생산완료여부
    Call oForm01.DataSources.UserDataSources.Add("WCYN02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("WCYN02").Specific.DataBind.SetBound(True, "", "WCYN02")
    
    '일자기준
    Call oForm01.DataSources.UserDataSources.Add("DateStd02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("DateStd02").Specific.DataBind.SetBound(True, "", "DateStd02")
    
    '기간(Fr)
    Call oForm01.DataSources.UserDataSources.Add("FrDt02", dt_DATE)
    Call oForm01.Items("FrDt02").Specific.DataBind.SetBound(True, "", "FrDt02")
    
    '기간(To)
    Call oForm01.DataSources.UserDataSources.Add("ToDt02", dt_DATE)
    Call oForm01.Items("ToDt02").Specific.DataBind.SetBound(True, "", "ToDt02")
    
    '품목(작번)
    Call oForm01.DataSources.UserDataSources.Add("ItemCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemCode02").Specific.DataBind.SetBound(True, "", "ItemCode02")
    
    '품목명
    Call oForm01.DataSources.UserDataSources.Add("ItemName02", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName02").Specific.DataBind.SetBound(True, "", "ItemName02")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("ItemSpec02", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemSpec02").Specific.DataBind.SetBound(True, "", "ItemSpec02")
    
    '공정
    Call oForm01.DataSources.UserDataSources.Add("CpCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CpCode02").Specific.DataBind.SetBound(True, "", "CpCode02")
    
    '공정명
    Call oForm01.DataSources.UserDataSources.Add("CpName02", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CpName02").Specific.DataBind.SetBound(True, "", "CpName02")
    
    '표준공수계(시간)
    Call oForm01.DataSources.UserDataSources.Add("TStdTime02", dt_QUANTITY)
    Call oForm01.Items("TStdTime02").Specific.DataBind.SetBound(True, "", "TStdTime02")
    
    '표준공수계(금액)
    Call oForm01.DataSources.UserDataSources.Add("TStdAmt02", dt_SUM)
    Call oForm01.Items("TStdAmt02").Specific.DataBind.SetBound(True, "", "TStdAmt02")
    
    '실동공수계(시간)
    Call oForm01.DataSources.UserDataSources.Add("TWkTime02", dt_QUANTITY)
    Call oForm01.Items("TWkTime02").Specific.DataBind.SetBound(True, "", "TWkTime02")
    
    '실동공수계(금액)
    Call oForm01.DataSources.UserDataSources.Add("TWkAmt02", dt_SUM)
    Call oForm01.Items("TWkAmt02").Specific.DataBind.SetBound(True, "", "TWkAmt02")
    
    '작업구분
    Call oForm01.DataSources.UserDataSources.Add("OrdGbn02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdGbn02").Specific.DataBind.SetBound(True, "", "OrdGbn02")
    '//////////공정별 집계//////////_E
    
    Call oForm01.Freeze(False)
    
    Exit Function
PS_PP989_CreateItems_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP989_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_PP989_ComboBox_Setting()
On Error GoTo PS_PP989_ComboBox_Setting_Error

    Dim sQry As String
    
    Dim BPLID As String
    BPLID = MDC_PS_Common.User_BPLId
    
    Call oForm01.Freeze(True)
    
    '//////////작번별 집계//////////_S
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID01").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, False, False)
    
    '작업구분
    sQry = "        SELECT      Code, "
    sQry = sQry & "             Name "
    sQry = sQry & " FROM        [@PSH_ITMBSORT]"
    sQry = sQry & " WHERE       U_PudYN = 'Y'"
    sQry = sQry & " ORDER BY    Code"
    Call oForm01.Items("OrdGbn01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdGbn01").Specific, sQry, "", False, False)
    Call oForm01.Items("OrdGbn01").Specific.Select("%", psk_ByValue)
    '//////////작번별 집계//////////_E
    
    '//////////공정별 집계//////////_S
    '사업장2
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID02").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
    '거래처구분
    sQry = "        SELECT      U_Minor AS [Code], "
    sQry = sQry & "             U_CdName AS [Name]"
    sQry = sQry & " FROM        [@PS_SY001L]"
    sQry = sQry & " WHERE       Code = 'C100'"
    Call oForm01.Items("CardType02").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType02").Specific, sQry, "", False, False)
    Call oForm01.Items("CardType02").Specific.Select("%", psk_ByValue)
    
    '품목구분
    sQry = "        SELECT      U_Minor AS [Code], "
    sQry = sQry & "             U_CdName AS [Name]"
    sQry = sQry & " FROM        [@PS_SY001L]"
    sQry = sQry & " WHERE       Code = 'S002'"
    Call oForm01.Items("ItemType02").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType02").Specific, sQry, "", False, False)
    Call oForm01.Items("ItemType02").Specific.Select("%", psk_ByValue)
    
    '생산완료여부
    Call oForm01.Items("WCYN02").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("WCYN02").Specific.ValidValues.Add("B", "미완료")
    Call oForm01.Items("WCYN02").Specific.ValidValues.Add("C", "완료")
    Call oForm01.Items("WCYN02").Specific.Select("%", psk_ByValue)
    
    '일자기준
    Call oForm01.Items("DateStd02").Specific.ValidValues.Add("01", "작업지시")
    Call oForm01.Items("DateStd02").Specific.ValidValues.Add("02", "작업일보")
'    Call oForm01.Items("DateStd").Specific.ValidValues.Add("03", "생산완료")
    Call oForm01.Items("DateStd02").Specific.Select("02", psk_ByValue)
    
    '작업구분
    sQry = "        SELECT      Code, "
    sQry = sQry & "             Name "
    sQry = sQry & " FROM        [@PSH_ITMBSORT]"
    sQry = sQry & " WHERE       U_PudYN = 'Y'"
    sQry = sQry & " ORDER BY    Code"
    Call oForm01.Items("OrdGbn02").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdGbn02").Specific, sQry, "", False, False)
    Call oForm01.Items("OrdGbn02").Specific.Select("%", psk_ByValue)
    '//////////공정별 집계//////////_E
    
'    '담당콤보세팅
'    Call oForm01.Items("RspCode").Specific.ValidValues.Add("%", "전체")
'    sQry = "            SELECT      U_Code AS [Code],"
'    sQry = sQry & "                 U_CodeNm As [Name]"
'    sQry = sQry & "  FROM       [@PS_HR200L]"
'    sQry = sQry & "  WHERE      Code = '2'"
'    sQry = sQry & "                 AND U_UseYN = 'Y'"
'    sQry = sQry & "                 AND U_Char1 = '2800'"
'    sQry = sQry & "  ORDER BY  U_Seq"
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("RspCode").Specific, sQry, "", False, False)
'    Call oForm01.Items("RspCode").Specific.Select(0, psk_Index)
    
'    '소속반
'    sQry = "           SELECT      T0.U_Code AS [Code],"
'    sQry = sQry & "                T0.U_CodeNm AS [Value]"
'    sQry = sQry & " FROM       [@PS_HR200L] AS T0"
'    sQry = sQry & " WHERE      T0.Code = '9'"
'    sQry = sQry & "                AND T0.U_Char3 = '2'"
'    sQry = sQry & "                AND T0.U_Char2 = '2800'"
'    sQry = sQry & "                AND T0.U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY  T0.U_Seq"
'    Call oForm01.Items("ClsCode").Specific.ValidValues.Add("%", "전체")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("ClsCode").Specific, sQry, "", False, False)
'    Call oForm01.Items("ClsCode").Specific.Select("%", psk_ByValue)
       
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_PP989_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP989_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP989_CF_ChooseFromList()
On Error GoTo PS_PP989_CF_ChooseFromList_Error

    Exit Sub
PS_PP989_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP989_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP989_EnableMenus()
On Error GoTo PS_PP989_EnableMenus_Error
    
    Exit Sub
PS_PP989_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP989_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP989_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP989_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP989_FormItemEnabled
        '//Call PS_PP989_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP989_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP989_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP989_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP989_FormResize()
On Error GoTo PS_PP989_FormResize_Error
    
    '그룹박스 크기 동적 할당
    oForm01.Items("GrpBox01").Height = oForm01.Items("Grid01").Height + 85
    oForm01.Items("GrpBox01").Width = oForm01.Items("Grid01").Width + 30
    
    If oGrid01.Columns.Count > 0 Then
        Call oGrid01.AutoResizeColumns
    End If
    
    Call oMat02.AutoResizeColumns
    
    Exit Sub
PS_PP989_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP989_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_PP989_GetDetail(ByVal pCpCode As String, ByVal pCpName As String)
On Error GoTo PS_PP989_GetDetail_Error

    Dim loopCount As Integer
    
    Dim BPLID     As String '사업장
    Dim TeamCode  As String '팀
    Dim RspCode   As String '담당
    Dim ClsCode   As String '반
    Dim CardType  As String '거래처구분
    Dim ItemType  As String '품목구분
    Dim WCYN      As String '생산완료여부
    Dim DateStd   As String '일자기준
    Dim FrDt      As String '기간(Fr)
    Dim ToDt      As String '기간(To)
    Dim ItemCode  As String '품목코드(작번)
    Dim CpCode    As String '공정
    Dim CpName    As String '공정명
    Dim OrdGbn    As String '작업구분

    Dim oTempClass As Variant

'    For loopCount = 0 To oGrid01.Rows.Count - 1
'        If oGrid01.Rows.IsSelected(loopCount) = True Then
'
'            '/ Procedure 실행
''             sQry = "EXEC PS_PP361_02 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "'"
'''            sQry = "EXEC PS_PP361_02 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "','" & SYYYYMM & "','" & EYYYYMM & "'"
''            oRecordSet.DoQuery sQry
''
''            '// Procedure 실행(Grid 사용)
''            oForm01.DataSources.DataTables.Item("ZTEMP02").ExecuteQuery (sQry)
''            oGrid02.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP02")
'
''            SumPkQty = Trim(oGrid02.DataTable.GetValue(7, i))
''            SumPkWt = Trim(oGrid02.DataTable.GetValue(8, i))
'
'            ItemCode = Trim(oGrid01.DataTable.GetValue(0, loopCount))
'            ItemName = Trim(oGrid01.DataTable.GetValue(1, loopCount))
'
'        End If
'    Next

    BPLID = Trim(oForm01.Items("BPLID02").Specific.VALUE)  '사업장
    TeamCode = Trim(oForm01.Items("TeamCode02").Specific.VALUE)  '팀
    RspCode = Trim(oForm01.Items("RspCode02").Specific.VALUE)  '담당
    ClsCode = Trim(oForm01.Items("ClsCode02").Specific.VALUE)  '반
    CardType = Trim(oForm01.Items("CardType02").Specific.VALUE)  '거래처구분
    ItemType = Trim(oForm01.Items("ItemType02").Specific.VALUE)  '품목구분
    WCYN = Trim(oForm01.Items("WCYN02").Specific.VALUE)  '생산완료여부
    DateStd = Trim(oForm01.Items("DateStd02").Specific.VALUE)  '일자기준
    FrDt = Trim(oForm01.Items("FrDt02").Specific.VALUE)  '기간(Fr)
    ToDt = Trim(oForm01.Items("ToDt02").Specific.VALUE)  '기간(To)
    ItemCode = Trim(oForm01.Items("ItemCode02").Specific.VALUE)  '품목코드(작번)
    CpCode = pCpCode  '공정
    CpName = pCpName '공정명
    OrdGbn = Trim(oForm01.Items("OrdGbn02").Specific.VALUE)  '작업구분

    Set oTempClass = New PS_PP990
    Call oTempClass.LoadForm(BPLID, TeamCode, RspCode, ClsCode, CardType, ItemType, WCYN, DateStd, FrDt, ToDt, ItemCode, CpCode, CpName, OrdGbn)
    
    Exit Sub

PS_PP989_GetDetail_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP989_GetDetail_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP989_Print_Report01()
On Error GoTo Print_Query_Error

    Dim DocNum     As String
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String

    Dim i          As Integer
    Dim ErrNum     As Integer
    Dim Sub_sQry   As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim BPLID    As String '사업장
    Dim TeamCode As String '팀
    Dim RspCode  As String '담당
    Dim ClsCode  As String '반
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim WCYN     As String '생산완료여부
    Dim DateStd  As String '일자기준
    Dim FrDt     As String '기간(Fr)
    Dim ToDt     As String '기간(To)
    Dim ItemCode As String '품목코드(작번)
    Dim CpCode   As String '공정
    Dim OrdGbn   As String '작업구분
    Dim CntcCode As String '조회자 사번
    
    BPLID = Trim(oForm01.Items("BPLID04").Specific.VALUE)  '사업장
    TeamCode = Trim(oForm01.Items("TeamCode04").Specific.VALUE)  '팀
    RspCode = Trim(oForm01.Items("RspCode04").Specific.VALUE)  '담당
    ClsCode = Trim(oForm01.Items("ClsCode04").Specific.VALUE)  '반
    CardType = Trim(oForm01.Items("CardType04").Specific.VALUE)  '거래처구분
    ItemType = Trim(oForm01.Items("ItemType04").Specific.VALUE)  '품목구분
    WCYN = Trim(oForm01.Items("WCYN04").Specific.VALUE)  '생산완료여부
    DateStd = Trim(oForm01.Items("DateStd04").Specific.VALUE)  '일자기준
    FrDt = Trim(oForm01.Items("FrDt04").Specific.VALUE)  '기간(Fr)
    ToDt = Trim(oForm01.Items("ToDt04").Specific.VALUE)  '기간(To)
    ItemCode = Trim(oForm01.Items("ItemCode04").Specific.VALUE)  '품목코드(작번)
    CpCode = Trim(oForm01.Items("CpCode04").Specific.VALUE)  '공정
    OrdGbn = Trim(oForm01.Items("OrdGbn04").Specific.VALUE)  '작업구분
    CntcCode = MDC_PS_Common.User_MSTCOD '조회자 사번
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP989] 레포트"
    ReportName = "PS_PP989_01.rpt"
 
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
    sQry = "      EXEC [PS_PP989_51] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & TeamCode & "','" '팀
    sQry = sQry & RspCode & "','" '담당
    sQry = sQry & ClsCode & "','" '반
    sQry = sQry & CardType & "','" '거래처구분
    sQry = sQry & ItemType & "','" '품목구분
    sQry = sQry & WCYN & "','" '생산완료여부
    sQry = sQry & DateStd & "','" '일자기준
    sQry = sQry & FrDt & "','" '기간(Fr)
    sQry = sQry & ToDt & "','" '기간(To)
    sQry = sQry & ItemCode & "','" '품목코드(작번)
    sQry = sQry & CpCode & "','" '공정
    sQry = sQry & OrdGbn & "','" '작업구분
    sQry = sQry & CntcCode & "'" '조회자사번

'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
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
        MDC_Com.MDC_GF_Message "PS_PP989_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub PS_PP989_Print_Report02()
On Error GoTo Print_Query_Error

    Dim DocNum     As String
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String

    Dim i          As Integer
    Dim ErrNum     As Integer
    Dim Sub_sQry   As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim BPLID    As String '사업장
    Dim TeamCode As String '팀
    Dim RspCode  As String '담당
    Dim ClsCode  As String '반
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim WCYN     As String '생산완료여부
    Dim DateStd  As String '일자기준
    Dim FrDt     As String '기간(Fr)
    Dim ToDt     As String '기간(To)
    Dim ItemCode As String '품목코드(작번)
    Dim CpCode   As String '공정
    Dim OrdGbn   As String '작업구분
    Dim CntcCode As String '조회자 사번
    
    BPLID = Trim(oForm01.Items("BPLID02").Specific.VALUE)  '사업장
    TeamCode = Trim(oForm01.Items("TeamCode02").Specific.VALUE)  '팀
    RspCode = Trim(oForm01.Items("RspCode02").Specific.VALUE)  '담당
    ClsCode = Trim(oForm01.Items("ClsCode02").Specific.VALUE)  '반
    CardType = Trim(oForm01.Items("CardType02").Specific.VALUE)  '거래처구분
    ItemType = Trim(oForm01.Items("ItemType02").Specific.VALUE)  '품목구분
    WCYN = Trim(oForm01.Items("WCYN02").Specific.VALUE)  '생산완료여부
    DateStd = Trim(oForm01.Items("DateStd02").Specific.VALUE)  '일자기준
    FrDt = Trim(oForm01.Items("FrDt02").Specific.VALUE)  '기간(Fr)
    ToDt = Trim(oForm01.Items("ToDt02").Specific.VALUE)  '기간(To)
    ItemCode = Trim(oForm01.Items("ItemCode02").Specific.VALUE)  '품목코드(작번)
    CpCode = Trim(oForm01.Items("CpCode02").Specific.VALUE)  '공정
    OrdGbn = Trim(oForm01.Items("OrdGbn02").Specific.VALUE)  '작업구분
    CntcCode = MDC_PS_Common.User_MSTCOD '조회자 사번
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP989] 레포트"
    ReportName = "PS_PP989_01.rpt"
 
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
    sQry = "      EXEC [PS_PP989_52] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & TeamCode & "','" '팀
    sQry = sQry & RspCode & "','" '담당
    sQry = sQry & ClsCode & "','" '반
    sQry = sQry & CardType & "','" '거래처구분
    sQry = sQry & ItemType & "','" '품목구분
    sQry = sQry & WCYN & "','" '생산완료여부
    sQry = sQry & DateStd & "','" '일자기준
    sQry = sQry & FrDt & "','" '기간(Fr)
    sQry = sQry & ToDt & "','" '기간(To)
    sQry = sQry & ItemCode & "','" '품목코드(작번)
    sQry = sQry & CpCode & "','" '공정
    sQry = sQry & OrdGbn & "','" '작업구분
    sQry = sQry & CntcCode & "'" '조회자사번

'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
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
        MDC_Com.MDC_GF_Message "PS_PP989_Print_Report02_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub












