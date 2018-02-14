VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_GA022.cls
'//  Module : 총무관리 > 외부인출입관리
'//  Desc : 일일방문객조회및출력
'//  FormType : PS_GA022
'//  Create Date(Start) : 2013.12.02
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
Private oDS_PS_GA022H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_GA022L As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA022.srf")
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
    
    oFormUniqueID01 = "PS_GA022_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_GA022_CreateItems
    Call PS_GA022_ComboBox_Setting
    Call PS_GA022_CF_ChooseFromList
    Call PS_GA022_EnableMenus
    Call PS_GA022_SetDocument(oFromDocEntry01)
    Call PS_GA022_FormResize

'    Call PS_GA022_Add_MatrixRow(0, True)
    
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
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing

'    '발생일자 설정
'    Call oDS_PS_GA022H.setValue("U_BOSDate", 0, Format(Date, "YYYYMMDD"))
'    '종료일자 설정
'    Call oDS_PS_GA022H.setValue("U_BOEDate", 0, Format(Date, "YYYYMMDD"))

     'oForm01.Items("GCode").Click ct_Regular
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

'/메트릭스 Row추가
Sub PS_GA022_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA022_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_GA022L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_GA022L.offset = oRow
    oDS_PS_GA022L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA022_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_GA022_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_GA022_MTX01()
'******************************************************************************
'Function ID : PS_GA022_MTX01()
'해당모듈 : PS_GA022
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA022_MTX01_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim SZGODAYF As String '기준일자F
    Dim SZGODAYT As String '기준일자T
    Dim SBPLID As String '사업장
  
    SZGODAYF = Trim(oForm01.Items("SZGODAYF").Specific.VALUE) '기준일자F
    SZGODAYT = Trim(oForm01.Items("SZGODAYT").Specific.VALUE) '기준일자T
    SBPLID = Trim(oForm01.Items("SBPLID").Specific.VALUE) '사업장
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Dim CntcCode As String '조회자사번
    Dim BPLID As String '사업장
    Dim ZGODAY As String '기준일자
    Dim ZDPTCD As String '부서
    Dim ZSECNM As String '담당
    Dim ZEMPIE As String '관련자사번
    Dim ZEMPNM As String '관련자성명
    Dim ZPERNP As String '확인자사번
    Dim ZUSRNN As String '방문자성명
    Dim ZBUSNM As String '업체명
    Dim ZOTHCNT As String '기타인원
    Dim ZBIRDT As String '생년월일
    Dim ZRESID As String '주소
    Dim ZCARJO As String '차량종류
    Dim ZCARNO As String '차량번호
    Dim ZPOINT As String '장소
    Dim ZNUM As String '코드
    Dim ZMOKJU As String '출입목적
    Dim ZPLAG As String '블랙박스유무
    Dim ZJEDAY As String '서약서일자
    Dim ZGOTIME As String '출입시간(시작)
    Dim ZOUTIME As String '출입시간(종료)
    Dim ZSPCIS As String '비고
    
    Call oForm01.Freeze(True)
    
    '사업장에 따른 구분
    If SBPLID = "2" Then '기계사업부 일 때

        Dim Client As String '클라이언트(운영용:210, 테스트용:710)
        Dim ServerIP As String '서버IP(운영용:192.1.11.3, 테스트용:192.1.11.7)
        Dim Plant As String '공장코드(부산사업장:9300)
        Dim E_MESSAGE As String '에러메시지
        Dim oTable As Table '테이블 변수
        
        'Real
        Client = "210"
        ServerIP = "192.1.11.3"

        'Test
'        Client = "710"
'        ServerIP = "192.1.11.7"


        Plant = "9300"
        '1. 연결
        Dim oSapConnection As New SAPFunctions 'R3연결 객체
'        Set oSapConnection = CreateObject("SAP.Functions")
        oSapConnection.Connection.User = "ifuser"
        oSapConnection.Connection.Password = "pdauser"
        oSapConnection.Connection.Client = Client
        oSapConnection.Connection.ApplicationServer = ServerIP
        oSapConnection.Connection.language = "KO"
        oSapConnection.Connection.SystemNumber = "00"

        If Not oSapConnection.Connection.Logon(0, True) Then
            ErrNum = 1
            GoTo PS_GA022_MTX01_Error
        End If

        '2. SAP R3 함수 호출
        Dim oFunction As New SAPFunctionsOCX.Function 'R3함수용 Function 객체
        Set oFunction = oSapConnection.Add("ZHR_INTF_VISITOR")

        oFunction.Exports("I_WERKS") = Plant '플랜트 : 기계사업부만 등록하므로 9300으로 Fix
        oFunction.Exports("I_ZDATF") = SZGODAYF '기준일자F
        oFunction.Exports("I_ZDATT") = SZGODAYT '기준일자T

        '-------------------------------------------------------------
        ' 함수를 Call 하면 oTable 객체의 내용을 그대로 SAP R3에 전달.
        ' 그러면 SAP R3는 oTable의 각각의 값들에 대해서 처리 결과를 리턴.
        '-------------------------------------------------------------
        If Not (oFunction.Call) Then

            oSapConnection.Connection.Logoff
            ErrNum = 99
            GoTo PS_GA022_MTX01_Error

        Else

            '-------------------------------------------------------------
            ' SAP함수 호출한 객체에서 테이블 객체 불러와서 값 할당
            '-------------------------------------------------------------
            Set oTable = oFunction.Tables.Item("T_ZVISITOR")
            E_MESSAGE = Trim(oFunction.Imports("E_MESSAGE").VALUE) '메시지

        End If

        CntcCode = MDC_PS_Common.User_MSTCOD '조회자사번

        '조회자가 조회하여 등록된 실테이블 내용 삭제
        sQry = "           DELETE"
        sQry = sQry & " FROM [@PS_GA022H]"
        sQry = sQry & " WHERE   U_CntcCode = '" & CntcCode & "'"
        
        Call oRecordSet01.DoQuery(sQry)

        oMat01.Clear
        oDS_PS_GA022L.Clear
        oMat01.FlushToDataSource
        oMat01.LoadFromDataSource

        For loopCount = 0 To oTable.RowCount - 1
        
            If loopCount + 1 > oDS_PS_GA022L.Size Then
                oDS_PS_GA022L.InsertRecord (loopCount)
            End If

            oMat01.AddRow
            oDS_PS_GA022L.offset = loopCount
            
            BPLID = "2" '사업장
            ZGODAY = Format(oTable(loopCount + 1, "ZGODAY"), "YYYYMMDD") '기준일자
            ZDPTCD = oTable(loopCount + 1, "ZDPTCD") '부서
            ZSECNM = oTable(loopCount + 1, "ZSECNM") '담당
            ZEMPIE = oTable(loopCount + 1, "ZEMPIE") '관련자사번
            ZEMPNM = oTable(loopCount + 1, "ZEMPNM") '관련자성명
            ZPERNP = oTable(loopCount + 1, "ZPERNP") '확인자사번
            ZUSRNN = oTable(loopCount + 1, "ZUSRNN") '방문자성명
            ZBUSNM = oTable(loopCount + 1, "ZBUSNM") '업체명
            ZOTHCNT = oTable(loopCount + 1, "ZOTHCNT") '기타인원
'            If oTable(loopCount + 1, "ZBIRDT") = "000000" Then
'                ZBIRDT = "19000101"
'            Else
'                ZBIRDT = oTable(loopCount + 1, "ZBIRDT") '생년월일
'            End If
            ZBIRDT = oTable(loopCount + 1, "ZBIRDT") '생년월일
            ZRESID = oTable(loopCount + 1, "ZRESID") '주소
            ZCARJO = oTable(loopCount + 1, "ZCARJO") '차량종류
            ZCARNO = oTable(loopCount + 1, "ZCARNO") '차량번호
            ZPOINT = oTable(loopCount + 1, "ZPOINT") '장소
            ZNUM = oTable(loopCount + 1, "ZNUM") '코드
            ZMOKJU = oTable(loopCount + 1, "ZMOKJU") '출입목적
            ZPLAG = oTable(loopCount + 1, "ZPLAG") '블랙박스유무
            ZJEDAY = Format(oTable(loopCount + 1, "ZJEDAY"), "YYYYMMDD") '서약서일자
            ZGOTIME = oTable(loopCount + 1, "ZGOTIME") '출입시간(시작)
            ZOUTIME = oTable(loopCount + 1, "ZOUTIME") '출입시간(종료)
            ZSPCIS = oTable(loopCount + 1, "ZSPCIS") '비고
            
            Call oDS_PS_GA022L.setValue("U_LineNum", loopCount, loopCount + 1)
'            Call oDS_PS_GA022L.setValue("U_ColReg02", loopCount, oTable(loopCount + 1, "WERKS")) '사업장
            Call oDS_PS_GA022L.setValue("U_ColReg02", loopCount, BPLID) '사업장
            Call oDS_PS_GA022L.setValue("U_ColDt01", loopCount, ZGODAY) '기준일자
            Call oDS_PS_GA022L.setValue("U_ColReg05", loopCount, ZDPTCD) '부서
            Call oDS_PS_GA022L.setValue("U_ColReg06", loopCount, ZSECNM) '담당
            Call oDS_PS_GA022L.setValue("U_ColReg07", loopCount, ZEMPIE) '관련자사번
            Call oDS_PS_GA022L.setValue("U_ColReg08", loopCount, ZEMPNM) '관련자성명
            Call oDS_PS_GA022L.setValue("U_ColReg09", loopCount, ZPERNP) '확인자사번
            Call oDS_PS_GA022L.setValue("U_ColReg10", loopCount, ZUSRNN) '방문자성명
            Call oDS_PS_GA022L.setValue("U_ColReg11", loopCount, ZBUSNM) '업체명
            Call oDS_PS_GA022L.setValue("U_ColReg12", loopCount, ZOTHCNT) '기타인원
            Call oDS_PS_GA022L.setValue("U_ColReg13", loopCount, ZBIRDT) '생년월일
            Call oDS_PS_GA022L.setValue("U_ColReg14", loopCount, ZRESID) '주소
            Call oDS_PS_GA022L.setValue("U_ColReg15", loopCount, ZCARJO) '차량종류
            Call oDS_PS_GA022L.setValue("U_ColReg16", loopCount, ZCARNO) '차량번호
            Call oDS_PS_GA022L.setValue("U_ColReg17", loopCount, ZPOINT) '장소
            Call oDS_PS_GA022L.setValue("U_ColReg18", loopCount, ZNUM) '코드
            Call oDS_PS_GA022L.setValue("U_ColReg19", loopCount, ZMOKJU) '출입목적
            Call oDS_PS_GA022L.setValue("U_ColReg20", loopCount, ZPLAG) '블랙박스유무
            Call oDS_PS_GA022L.setValue("U_ColDt03", loopCount, ZJEDAY) '서약서일자
            Call oDS_PS_GA022L.setValue("U_ColReg22", loopCount, ZGOTIME) '출입시간(시작)
            Call oDS_PS_GA022L.setValue("U_ColReg23", loopCount, ZOUTIME) '출입시간(종료)
            Call oDS_PS_GA022L.setValue("U_ColReg24", loopCount, ZSPCIS) '비고
            
            '실테이블에 저장(리포트조회를 위함)
            sQry = "                EXEC [PS_GA022_02] "
            sQry = sQry & "'" & CntcCode & "',"  '조회자사번
            sQry = sQry & "'" & BPLID & "',"  '사업장
            sQry = sQry & "'" & ZGODAY & "',"  '기준일자
            sQry = sQry & "'" & ZDPTCD & "',"  '부서
            sQry = sQry & "'" & ZSECNM & "',"  '담당
            sQry = sQry & "'" & ZEMPIE & "',"  '관련자사번
            sQry = sQry & "'" & ZEMPNM & "',"  '관련자성명
            sQry = sQry & "'" & ZPERNP & "',"  '확인자사번
            sQry = sQry & "'" & ZUSRNN & "',"  '방문자성명
            sQry = sQry & "'" & ZBUSNM & "',"  '업체명
            sQry = sQry & "'" & ZOTHCNT & "',"  '기타인원
            sQry = sQry & "'" & ZBIRDT & "',"  '생년월일
            sQry = sQry & "'" & ZRESID & "',"  '주소
            sQry = sQry & "'" & ZCARJO & "',"  '차량종류
            sQry = sQry & "'" & ZCARNO & "',"  '차량번호
            sQry = sQry & "'" & ZPOINT & "',"  '장소
            sQry = sQry & "'" & ZNUM & "',"  '코드
            sQry = sQry & "'" & ZMOKJU & "',"  '출입목적
            sQry = sQry & "'" & ZPLAG & "',"  '블랙박스유무
            sQry = sQry & "'" & ZJEDAY & "',"  '서약서일자
            sQry = sQry & "'" & ZGOTIME & "',"  '출입시간(시작)
            sQry = sQry & "'" & ZOUTIME & "',"  '출입시간(종료)
            sQry = sQry & "'" & ZSPCIS & "'"  '비고
            
            Call oRecordset02.DoQuery(sQry)
            
            ProgBar01.VALUE = ProgBar01.VALUE + 1
            ProgBar01.Text = ProgBar01.VALUE & "/" & oTable.RowCount & "건 조회중...!"
        
        Next

    Else '기계사업부 이외(요구사항 발생 시 구현을 위해 남겨둠, 송명규 2013.12.03)
    
'        sQry = "         EXEC [PS_GA022_01] '"
'        sQry = sQry & SZGODAYF & "','"
'        sQry = sQry & SZGODAYT & "','"
'        sQry = sQry & SBPLID & "','"
'        sQry = sQry & SZDPTCD & "','"
'        sQry = sQry & SZSECCD & "','"
'        sQry = sQry & SZEMPIE & "'"
'
'        Call oRecordSet01.DoQuery(sQry)
'
'        oMat01.Clear
'        oDS_PS_GA022L.Clear
'        oMat01.FlushToDataSource
'        oMat01.LoadFromDataSource
'
'        If (oRecordSet01.RecordCount = 0) Then
'
'            ErrNum = 1
'
'            oForm01.Mode = fm_ADD_MODE
'
'    '        Call PS_GA022_Add_MatrixRow(0, True)
'
'
'            GoTo PS_GA022_MTX01_Error
'
'            Exit Sub
'        End If
'
'        For loopCount = 0 To oRecordSet01.RecordCount - 1
'            If loopCount + 1 > oDS_PS_GA022L.Size Then
'                oDS_PS_GA022L.InsertRecord (loopCount)
'            End If
'
'            oMat01.AddRow
'            oDS_PS_GA022L.Offset = loopCount
'
'            Call oDS_PS_GA022L.setValue("U_LineNum", loopCount, loopCount + 1)
'            Call oDS_PS_GA022L.setValue("U_ColReg01", loopCount, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
'            Call oDS_PS_GA022L.setValue("U_ColReg02", loopCount, Trim(oRecordSet01.Fields("BPLID").VALUE)) '사업장
'            Call oDS_PS_GA022L.setValue("U_ColDt01", loopCount, Format(Trim(oRecordSet01.Fields("ZGODAY").VALUE), "YYYYMMDD")) '기준일자
'            Call oDS_PS_GA022L.setValue("U_ColReg03", loopCount, Trim(oRecordSet01.Fields("ZEMPIE").VALUE)) '관련자사번
'            Call oDS_PS_GA022L.setValue("U_ColReg04", loopCount, Trim(oRecordSet01.Fields("ZEMPNM").VALUE)) '관련자성명
'            Call oDS_PS_GA022L.setValue("U_ColReg05", loopCount, Trim(oRecordSet01.Fields("ZDPTCD").VALUE)) '관련자팀(CD)
'            Call oDS_PS_GA022L.setValue("U_ColReg06", loopCount, Trim(oRecordSet01.Fields("ZSECCD").VALUE)) '관련자담당(CD)
'            Call oDS_PS_GA022L.setValue("U_ColReg07", loopCount, Trim(oRecordSet01.Fields("ZPOINT").VALUE)) '장소
'            Call oDS_PS_GA022L.setValue("U_ColReg08", loopCount, Trim(oRecordSet01.Fields("ZUSRNN").VALUE)) '방문자성명
'            Call oDS_PS_GA022L.setValue("U_ColDt02", loopCount, Format(Trim(oRecordSet01.Fields("ZBIRDT").VALUE), "YYYYMMDD")) '생년월일
'            Call oDS_PS_GA022L.setValue("U_ColReg10", loopCount, Trim(oRecordSet01.Fields("ZBUSNM").VALUE)) '업체명
'            Call oDS_PS_GA022L.setValue("U_ColReg11", loopCount, Trim(oRecordSet01.Fields("ZRESID").VALUE)) '주소
'            Call oDS_PS_GA022L.setValue("U_ColTm01", loopCount, Trim(oRecordSet01.Fields("ZGOTIME").VALUE)) '출입시간(Fr)
'            Call oDS_PS_GA022L.setValue("U_ColTm02", loopCount, Trim(oRecordSet01.Fields("ZOUTIME").VALUE)) '출입시간(To)
'            Call oDS_PS_GA022L.setValue("U_ColReg14", loopCount, Trim(oRecordSet01.Fields("ZCARJO").VALUE)) '차량종류
'            Call oDS_PS_GA022L.setValue("U_ColReg17", loopCount, Trim(oRecordSet01.Fields("ZCARNO").VALUE)) '차량번호
'            Call oDS_PS_GA022L.setValue("U_ColReg15", loopCount, Trim(oRecordSet01.Fields("ZMOKJU").VALUE)) '출입목적
'            Call oDS_PS_GA022L.setValue("U_ColReg16", loopCount, Trim(oRecordSet01.Fields("ZSPCIU").VALUE)) '방문객비고
'
'            oRecordSet01.MoveNext
'            ProgBar01.VALUE = ProgBar01.VALUE + 1
'            ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
'
'        Next
    
    End If
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Set oRecordset02 = Nothing
    oSapConnection.Connection.Logoff
    Set oTable = Nothing
    Set oSapConnection = Nothing
    Set oFunction = Nothing
       
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA022_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Set oRecordset02 = Nothing
    oSapConnection.Connection.Logoff
    Set oTable = Nothing
    Set oSapConnection = Nothing
    Set oFunction = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA022_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'Private Function PS_GA022_HeaderSpaceLineDel() As Boolean
''******************************************************************************
''Function ID : PS_GA022_HeaderSpaceLineDel()
''해당모듈 : PS_GA022
''기능 : 필수입력사항 체크
''인수 : 없음
''반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
''특이사항 : 없음
''******************************************************************************
'On Error GoTo PS_GA022_HeaderSpaceLineDel_Error
'
'    Dim ErrNum As Integer
'    ErrNum = 0
'
'    Select Case True
'        Case Trim(oForm01.Items("Loc").Specific.VALUE) = "" '발생위치
'            ErrNum = 1
'            GoTo PS_GA022_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BORsn").Specific.VALUE) = "" '정전사유
'            ErrNum = 2
'            GoTo PS_GA022_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOCls").Specific.VALUE) = "%" '정전구분
'            ErrNum = 3
'            GoTo PS_GA022_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOSDate").Specific.VALUE) = "" '발생일자
'            ErrNum = 4
'            GoTo PS_GA022_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOSHour").Specific.VALUE) = "" '발생시각(시)
'            ErrNum = 5
'            GoTo PS_GA022_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOSMin").Specific.VALUE) = "" '발생시각(분)
'            ErrNum = 6
'            GoTo PS_GA022_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOEDate").Specific.VALUE) = "" '종료일자
'            ErrNum = 7
'            GoTo PS_GA022_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOEHour").Specific.VALUE) = "" '종료시각(시)
'            ErrNum = 8
'            GoTo PS_GA022_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOEMin").Specific.VALUE) = "" '종료시각(분)
'            ErrNum = 9
'            GoTo PS_GA022_HeaderSpaceLineDel_Error
'    End Select
'
'    PS_GA022_HeaderSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_GA022_HeaderSpaceLineDel_Error:
'    If ErrNum = 1 Then
'        MDC_Com.MDC_GF_Message "발생위치는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("Loc").Click(ct_Regular)
'    ElseIf ErrNum = 2 Then
'        MDC_Com.MDC_GF_Message "정전사유는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BORsn").Click(ct_Regular)
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "정전구분은 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOCls").Click(ct_Regular)
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "발생일자는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOSDate").Click(ct_Regular)
'    ElseIf ErrNum = 5 Then
'        MDC_Com.MDC_GF_Message "발생시각(시)는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOSHour").Click(ct_Regular)
'    ElseIf ErrNum = 6 Then
'        MDC_Com.MDC_GF_Message "발생시각(분)은 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOSMin").Click(ct_Regular)
'    ElseIf ErrNum = 7 Then
'        MDC_Com.MDC_GF_Message "종료일자는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOEDate").Click(ct_Regular)
'    ElseIf ErrNum = 8 Then
'        MDC_Com.MDC_GF_Message "종료일자(시)는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOEHour").Click(ct_Regular)
'    ElseIf ErrNum = 9 Then
'        MDC_Com.MDC_GF_Message "종료일자(분)은 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOEMin").Click(ct_Regular)
'    End If
'    PS_GA022_HeaderSpaceLineDel = False
'End Function
'
''/ 메트릭스 필수 사항 check
'Private Function PS_GA022_MatrixSpaceLineDel() As Boolean
'On Error GoTo PS_GA022_MatrixSpaceLineDel_Error
'
'    Dim i As Long
'    Dim ErrNum As Integer
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Dim sQry As String
'
'    Set oRecordSet01 = Nothing
'    PS_GA022_MatrixSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_GA022_MatrixSpaceLineDel_Error:
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
'        MDC_Com.MDC_GF_Message "PS_GA022_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'    PS_GA022_MatrixSpaceLineDel = False
'End Function

Private Sub PS_GA022_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_GA022_FlushToItemValue_Error

Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    Dim Qty As Integer
    Dim Calculate_Weight As Double
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String
    Dim ZDPTCD As String
    Dim ZSECCD As String
    
    Select Case oUID

'        Case "SBPLID"
'
'            BPLID = Trim(oForm01.Items("SBPLID").Specific.VALUE)
'
'            If oForm01.Items("SZDPTCD").Specific.ValidValues.Count > 0 Then
'                For i = oForm01.Items("SZDPTCD").Specific.ValidValues.Count - 1 To 0 Step -1
'                    Call oForm01.Items("SZDPTCD").Specific.ValidValues.Remove(i, psk_Index)
'                Next i
'            End If
'
'            '부서콤보세팅
'            Call oForm01.Items("SZDPTCD").Specific.ValidValues.Add("%", "전체")
'            sQry = "            SELECT      U_Code AS [Code],"
'            sQry = sQry & "                 U_CodeNm As [Name]"
'            sQry = sQry & "  FROM       [@PS_HR200L]"
'            sQry = sQry & "  WHERE      Code = '1'"
'            sQry = sQry & "                 AND U_UseYN = 'Y'"
'            sQry = sQry & "                 AND U_Char2 = '" & BPLID & "'"
'            sQry = sQry & "  ORDER BY  U_Seq"
'            Call MDC_SetMod.Set_ComboList(oForm01.Items("SZDPTCD").Specific, sQry, "", False, False)
'            Call oForm01.Items("SZDPTCD").Specific.Select(0, psk_Index)
'
'        Case "SZDPTCD"
'
'            ZDPTCD = Trim(oForm01.Items("SZDPTCD").Specific.VALUE)
'
'            If oForm01.Items("SZSECCD").Specific.ValidValues.Count > 0 Then
'                For i = oForm01.Items("SZSECCD").Specific.ValidValues.Count - 1 To 0 Step -1
'                    Call oForm01.Items("SZSECCD").Specific.ValidValues.Remove(i, psk_Index)
'                Next i
'            End If
'
'            '담당콤보세팅
'            Call oForm01.Items("SZSECCD").Specific.ValidValues.Add("%", "전체")
'            sQry = "            SELECT      U_Code AS [Code],"
'            sQry = sQry & "                 U_CodeNm As [Name]"
'            sQry = sQry & "  FROM       [@PS_HR200L]"
'            sQry = sQry & "  WHERE      Code = '2'"
'            sQry = sQry & "                 AND U_UseYN = 'Y'"
'            sQry = sQry & "                 AND U_Char1 = '" & ZDPTCD & "'"
'            sQry = sQry & "  ORDER BY  U_Seq"
'            Call MDC_SetMod.Set_ComboList(oForm01.Items("SZSECCD").Specific, sQry, "", False, False)
'            Call oForm01.Items("SZSECCD").Specific.Select(0, psk_Index)
'
'        Case "SZEMPIE"
'
'            oForm01.Items("SZEMPNM").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" & oForm01.Items("SZEMPIE").Specific.VALUE & "'") '성명
        
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_GA022_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("시각은 숫자만 입력이 가능합니다.", "E")
    ElseIf ErrNum = 2 Then
        Call MDC_Com.MDC_GF_Message("시각(시)는 24미만의 값만 입력이 가능합니다.", "E")
    ElseIf ErrNum = 3 Then
        Call MDC_Com.MDC_GF_Message("시각(분)은 60미만의 값만 입력이 가능합니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA022_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_GA022_FormItemEnabled()
On Error GoTo PS_GA022_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA022_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_GA022_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
'                oForm01.Mode = fm_ADD_MODE
'                BubbleEvent = False
                
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
                '//Call PS_GA022_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_GA022H.Clear

'                Call PS_GA022_LoadCaption
'                Call PS_GA022_FormItemEnabled
                '//Call PS_GA022_FormItemEnabled '//UDO방식
                '//Call PS_GA022_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_GA022_FormItemEnabled
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
    
        If pval.ItemUID = "PS_GA022" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnSearch" Then '/조회
                
            Call PS_GA022_MTX01
            
        ElseIf pval.ItemUID = "BtnPrint" Then '출력
            
            Call PS_GA022_Print_Report01
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA022" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "SZEMPIE", "") '//사용자값활성
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

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PS_GA022_FlushToItemValue(pval.ItemUID)
    
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

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_GA022L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_GA022L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_GA022_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_GA022L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PS_GA022_FlushToItemValue(pval.ItemUID)
                
'                If pval.ItemUID = "BOSHour" Then
'                    Call PS_GA022_FlushToItemValue(pval.ItemUID)
'                ElseIf (pval.ItemUID = "CardCode") Then
'
'                Else
'
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
        Call PS_GA022_FormItemEnabled
        '//Call PS_GA022_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_GA022_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_GA022H", "U_CardCode,U_CardName")
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
'            If (PS_GA022_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_GA022H.RemoveRecord(oDS_PS_GA022H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_GA022_Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_GA022H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_GA022_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_GA022_CreateItems() As Boolean
On Error GoTo PS_GA022_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_GA022H = oForm01.DataSources.DBDataSources("@PS_GA022H")
    Set oDS_PS_GA022L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '기준일자_From_S
    Call oForm01.DataSources.UserDataSources.Add("SZGODAYF", dt_DATE)
    Call oForm01.Items("SZGODAYF").Specific.DataBind.SetBound(True, "", "SZGODAYF")
    '기준일자_From_E
    
    '기준일자_To_S
    Call oForm01.DataSources.UserDataSources.Add("SZGODAYT", dt_DATE)
    Call oForm01.Items("SZGODAYT").Specific.DataBind.SetBound(True, "", "SZGODAYT")
    '기준일자_To_E
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("SBPLID", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SBPLID").Specific.DataBind.SetBound(True, "", "SBPLID")
    '사업장_E
    
    '기준일자(Fr)
    oForm01.Items("SZGODAYF").Specific.VALUE = Format(Date, "YYYYMMDD")
    '기준일자(To)
    oForm01.Items("SZGODAYT").Specific.VALUE = Format(Date, "YYYYMMDD")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_GA022_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA022_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_GA022_ComboBox_Setting()
On Error GoTo PS_GA022_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
        
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("SBPLID").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '//////////매트릭스//////////S
    '사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLID"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    '부서
    sQry = "            SELECT      U_Code AS [Code],"
    sQry = sQry & "                 U_CodeNm As [Name]"
    sQry = sQry & "  FROM       [@PS_HR200L]"
    sQry = sQry & "  WHERE      Code = '1'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ZDPTCD"), sQry)
    
    '담당
'    sQry = "            SELECT      U_Code AS [Code],"
'    sQry = sQry & "                 U_CodeNm As [Name]"
'    sQry = sQry & "  FROM       [@PS_HR200L]"
'    sQry = sQry & "  WHERE      Code = '2'"
'    sQry = sQry & "                 AND U_UseYN = 'Y'"
'    sQry = sQry & "  ORDER BY  U_Seq"
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ZSECCD"), sQry)
    
    '장소
    sQry = "            SELECT      U_Code AS [Code],"
    sQry = sQry & "                 U_CodeNm As [Name]"
    sQry = sQry & "  FROM       [@PS_GA050L]"
    sQry = sQry & "  WHERE      Code = 'G003'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ZPOINT"), sQry)
    
    '차량종류
    sQry = "            SELECT      U_Code AS [Code],"
    sQry = sQry & "                 U_CodeNm As [Name]"
    sQry = sQry & "  FROM       [@PS_GA050L]"
    sQry = sQry & "  WHERE      Code = 'G002'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ZCARJO"), sQry)
    
    '출입목적
    sQry = "            SELECT      U_Code AS [Code],"
    sQry = sQry & "                 U_CodeNm As [Name]"
    sQry = sQry & "  FROM       [@PS_GA050L]"
    sQry = sQry & "  WHERE      Code = 'G004'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ZMOKJU"), sQry)
    
    '블랙박스유무
    
    
    '//////////매트릭스//////////E
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_GA022_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA022_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA022_CF_ChooseFromList()
On Error GoTo PS_GA022_CF_ChooseFromList_Error

    Exit Sub
PS_GA022_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA022_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA022_EnableMenus()
On Error GoTo PS_GA022_EnableMenus_Error
    
    Exit Sub
PS_GA022_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA022_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA022_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_GA022_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_GA022_FormItemEnabled
        '//Call PS_GA022_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_GA022_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_GA022_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA022_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA022_FormResize()
On Error GoTo PS_GA022_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_GA022_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA022_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA022_Print_Report01()
On Error GoTo Print_Query_Error
    Dim DocNum As String
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    Dim CntcCode As String
    Dim BPLID As String '사업장
    Dim ZGODAYF As String '기준일자F
    Dim ZGODAYT As String '기준일자T
    
    CntcCode = MDC_PS_Common.User_MSTCOD '조회자사번
    BPLID = Trim(oForm01.Items("SBPLID").Specific.VALUE) '사업장
    ZGODAYF = Trim(oForm01.Items("SZGODAYF").Specific.VALUE) '기준일자F
    ZGODAYT = Trim(oForm01.Items("SZGODAYT").Specific.VALUE) '기준일자T
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_GA022] 레포트"
    ReportName = "PS_GA022_01.rpt"
 
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
    sQry = "         EXEC [PS_GA022_01] '"
    sQry = sQry & CntcCode & "','"
    sQry = sQry & BPLID & "','"
    sQry = sQry & ZGODAYF & "','"
    sQry = sQry & ZGODAYT & "'"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_GA022_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub






