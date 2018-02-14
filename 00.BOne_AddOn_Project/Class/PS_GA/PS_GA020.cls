VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_GA020.cls
'//  Module : 총무관리 > 외부인출입관리
'//  Desc : 출입예정자등록
'//  FormType : PS_GA020
'//  Create Date(Start) : 2013.11.18
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
Private oDS_PS_GA020H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_GA020L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private gDeleteRowCount As Integer '행삭제 카운트



'//사용자구조체
Private Type DeleteRowInfo
    DocEntry As Integer
End Type

Private oLast_Mode&

Private DeleteRow() As DeleteRowInfo

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA020.srf")
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
    
    oFormUniqueID01 = "PS_GA020_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_GA020_CreateItems
    Call PS_GA020_ComboBox_Setting
    Call PS_GA020_CF_ChooseFromList
    Call PS_GA020_EnableMenus
    Call PS_GA020_SetDocument(oFromDocEntry01)
    Call PS_GA020_FormResize

'    Call PS_GA020_Add_MatrixRow(0, True)
    Call PS_GA020_LoadCaption
    Call PS_GA020_FormItemEnabled
    
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

    '관리번호
'    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_GA020H]"
'    Call RecordSet01.DoQuery(sQry)
'    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'        Call oDS_PS_GA020H.setValue("DocEntry", 0, 1)
'    Else
'        Call oDS_PS_GA020H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
'    End If
    
    '기준일자 설정
    Call oDS_PS_GA020H.setValue("U_ZGODAY", 0, Format(Date, "YYYYMMDD"))

    '기준일자(Fr)
    oForm01.Items("SZGODAY").Specific.VALUE = Format(Date, "YYYYMMDD")
'    '기준일자(To)
'    oForm01.Items("SZGODAYT").Specific.VALUE = Format(Date, "YYYYMMDD")

    '선택행번호 Visible = False
    oForm01.Items("RowNo").Visible = False
    oForm01.Items("Static24").Visible = False
    
    '관련자사번 클릭(화면 최초 로드시)
    oForm01.Items("ZEMPIE").Click ct_Regular
    
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

Private Sub PS_GA020_LoadCaption()
'******************************************************************************
'Function ID : PS_GA020_LoadCaption()
'해당모듈 : PS_GA020
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA020_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
    
        oForm01.Items("BtnAdd").Specific.Caption = "추가"
        oForm01.Items("BtnAddRow").Specific.Caption = "행추가"

'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"

    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
    
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
        oForm01.Items("BtnAddRow").Specific.Caption = "행수정"
        
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_GA020_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA020_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_GA020L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_GA020L.offset = oRow
    oDS_PS_GA020L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_GA020_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_GA020_MTX01()
'******************************************************************************
'Function ID : PS_GA020_MTX01()
'해당모듈 : PS_GA020
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA020_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim SZGODAY As String '기준일자
    Dim SBPLID As String '사업장
    Dim SZDPTCD As String '관련자팀
    Dim SZSECCD As String '관련자담당
    Dim SZEMPIE As String '관련자사번
    
    SZGODAY = Trim(oForm01.Items("SZGODAY").Specific.VALUE) '기준일자
    SBPLID = Trim(oForm01.Items("SBPLID").Specific.VALUE) '사업장
    SZDPTCD = Trim(oForm01.Items("SZDPTCD").Specific.VALUE) '관련자팀
    SZSECCD = Trim(oForm01.Items("SZSECCD").Specific.VALUE) '관련자담당
    SZEMPIE = Trim(oForm01.Items("SZEMPIE").Specific.VALUE) '관련자사번
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_GA020_01] '"
    sQry = sQry & SZGODAY & "','"
    sQry = sQry & SBPLID & "','"
    sQry = sQry & SZDPTCD & "','"
    sQry = sQry & SZSECCD & "','"
    sQry = sQry & SZEMPIE & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_GA020L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_GA020_Add_MatrixRow(0, True)
        Call PS_GA020_LoadCaption
        
        GoTo PS_GA020_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_GA020L.Size Then
            oDS_PS_GA020L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_GA020L.offset = i
        
        Call oDS_PS_GA020L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_GA020L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PS_GA020L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("BPLID").VALUE)) '사업장
        Call oDS_PS_GA020L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("ZGODAY").VALUE), "YYYYMMDD")) '기준일자
        Call oDS_PS_GA020L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("ZEMPIE").VALUE)) '관련자사번
        Call oDS_PS_GA020L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("ZEMPNM").VALUE)) '관련자성명
        Call oDS_PS_GA020L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("ZDPTCD").VALUE)) '관련자팀(CD)
        Call oDS_PS_GA020L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("ZSECCD").VALUE)) '관련자담당(CD)
        Call oDS_PS_GA020L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("ZPOINT").VALUE)) '장소
        Call oDS_PS_GA020L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("ZUSRNN").VALUE)) '방문자성명
        Call oDS_PS_GA020L.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("ZBIRDT").VALUE), "YYYYMMDD")) '생년월일
        Call oDS_PS_GA020L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("ZBUSNM").VALUE)) '업체명
        Call oDS_PS_GA020L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("ZRESID").VALUE)) '주소
        Call oDS_PS_GA020L.setValue("U_ColTm01", i, Trim(oRecordSet01.Fields("ZGOTIME").VALUE)) '출입시간(Fr)
        Call oDS_PS_GA020L.setValue("U_ColTm02", i, Trim(oRecordSet01.Fields("ZOUTIME").VALUE)) '출입시간(To)
        Call oDS_PS_GA020L.setValue("U_ColReg14", i, Trim(oRecordSet01.Fields("ZCARJO").VALUE)) '차량종류
        Call oDS_PS_GA020L.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("ZCARNO").VALUE)) '차량번호
        Call oDS_PS_GA020L.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("ZMOKJU").VALUE)) '출입목적
        Call oDS_PS_GA020L.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("ZSPCIU").VALUE)) '방문객비고
    
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
PS_GA020_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA020_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_GA020_DeleteData()
'******************************************************************************
'Function ID : PS_GA020_DeleteData()
'해당모듈 : PS_GA020
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA020_DeleteData_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PS_GA020H] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_GA020_DeleteData_Error
        Else
            sQry = "EXEC PS_GA020_04 '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
'    Call PS_GA020_FormReset
    
'    oForm01.Mode = fm_ADD_MODE
    
'    Call oForm01.Items("BtnSearch").Click(ct_Regular)
    
'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PS_GA020_Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA020_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_GA020_UpdateData() As Boolean
'******************************************************************************
'Function ID : PS_GA020_UpdateData()
'해당모듈 : PS_GA020
'기능 : 기본정보를 수정
'인수 : 없음
'반환값 : 없음
'특이사항 : AddData 함수에 Update 기능 포함(쿼리에서 제어)되어 본 Class에서는 사용 안함
'******************************************************************************
On Error GoTo PS_GA020_UpdateData_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim BPLID As String '사업장
    Dim ZDPTCD As String '관련자팀(코드)
    Dim ZDPTNM As String '관련자팀(명)
    Dim ZSECCD As String '관련자담당(코드)
    Dim ZSECNM As String '관련자담당(명)
    Dim ZEMPIE As String '관련자사번
    Dim ZEMPNM As String '관련자성명
    Dim ZGODAY As String '기준일자
    Dim ZGOTIME As String '출입시간(Fr)
    Dim ZOUTIME As String '출입시간(To)
    Dim ZUSRNN As String '방문자성명
    Dim ZBIRDT As String '생년월일
    Dim ZBUSNM As String '업체명
    Dim ZRESID As String '주소
    Dim ZCARJO As String '차량종류
    Dim ZMOKJU As String '출입목적
    Dim ZPOINT As String '장소
    Dim ZSPCIU As String '비고
    
    For loopCount = 0 To oMat01.RowCount - 1
    
        DocEntry = oDS_PS_GA020L.GetValue("U_ColReg01", loopCount) '관리번호
        BPLID = oDS_PS_GA020L.GetValue("U_ColReg02", loopCount) '사업장
        ZDPTCD = oDS_PS_GA020L.GetValue("U_ColReg05", loopCount) '관련자팀(코드)
        ZSECCD = oDS_PS_GA020L.GetValue("U_ColReg06", loopCount) '관련자담당(코드)
        ZEMPIE = oDS_PS_GA020L.GetValue("U_ColReg03", loopCount) '관련자사번
        ZEMPNM = oDS_PS_GA020L.GetValue("U_ColReg04", loopCount) '관련자성명
        ZGODAY = oDS_PS_GA020L.GetValue("U_ColDt01", loopCount) '기준일자
        ZGOTIME = oDS_PS_GA020L.GetValue("U_ColTm01", loopCount) '출입시간(Fr)
        ZOUTIME = oDS_PS_GA020L.GetValue("U_ColTm02", loopCount) '출입시간(To)
        ZUSRNN = oDS_PS_GA020L.GetValue("U_ColReg08", loopCount) '방문자성명
        ZBIRDT = oDS_PS_GA020L.GetValue("U_ColDt02", loopCount) '생년월일
        ZBUSNM = oDS_PS_GA020L.GetValue("U_ColReg10", loopCount) '업체명
        ZRESID = oDS_PS_GA020L.GetValue("U_ColReg11", loopCount) '주소
        ZCARJO = oDS_PS_GA020L.GetValue("U_ColReg14", loopCount) '차량종류
        ZMOKJU = oDS_PS_GA020L.GetValue("U_ColReg15", loopCount) '출입목적
        ZPOINT = oDS_PS_GA020L.GetValue("U_ColReg07", loopCount) '장소
        ZSPCIU = oDS_PS_GA020L.GetValue("U_ColReg16", loopCount) '비고
        
        If Trim(DocEntry) = "" Then
            MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
            PS_GA020_UpdateData = False
            Exit Function
        End If
        
        '관련자팀(명)
        sQry = "            SELECT       T0.U_CodeNm AS [ZDPTNM]"
        sQry = sQry & "  FROM        [@PS_HR200L] AS T0"
        sQry = sQry & "  WHERE       T0.Code = '1'"
        sQry = sQry & "                  AND T0.U_Code = '" & ZDPTCD & "'"
        Call RecordSet01.DoQuery(sQry)
        ZDPTNM = Trim(RecordSet01.Fields(0).VALUE)
        
        '관련자담당(명)
        sQry = "            SELECT       T0.U_CodeNm AS [ZSECNM]"
        sQry = sQry & "  FROM        [@PS_HR200L] AS T0"
        sQry = sQry & "  WHERE       T0.Code = '2'"
        sQry = sQry & "                  AND T0.U_Code = '" & ZSECCD & "'"
        Call RecordSet01.DoQuery(sQry)
        ZSECNM = Trim(RecordSet01.Fields(0).VALUE)
        
        sQry = "                EXEC [PS_GA020_03] "
        sQry = sQry & "'" & DocEntry & "'," '관리번호
        sQry = sQry & "'" & BPLID & "',"  '사업장
        sQry = sQry & "'" & ZDPTCD & "',"  '관련자팀(코드)
        sQry = sQry & "'" & ZDPTNM & "',"  '관련자팀(명)
        sQry = sQry & "'" & ZSECCD & "',"  '관련자담당(코드)
        sQry = sQry & "'" & ZSECNM & "',"  '관련자담당(명)
        sQry = sQry & "'" & ZEMPIE & "',"  '관련자사번
        sQry = sQry & "'" & ZEMPNM & "',"  '관련자성명
        sQry = sQry & "'" & ZGODAY & "',"  '기준일자
        sQry = sQry & "'" & ZGOTIME & "',"  '출입시간(Fr)
        sQry = sQry & "'" & ZOUTIME & "',"  '출입시간(To)
        sQry = sQry & "'" & ZUSRNN & "',"  '방문자성명
        sQry = sQry & "'" & ZBIRDT & "',"  '생년월일
        sQry = sQry & "'" & ZBUSNM & "',"  '업체명
        sQry = sQry & "'" & ZRESID & "',"  '주소
        sQry = sQry & "'" & ZCARJO & "',"  '차량종류
        sQry = sQry & "'" & ZMOKJU & "',"  '출입목적
        sQry = sQry & "'" & ZPOINT & "',"  '장소
        sQry = sQry & "'" & ZSPCIU & "'"  '비고
        
        Call RecordSet01.DoQuery(sQry)
    
    Next
    
    '행 삭제 처리한 내역 DB에서 DELETE
    For loopCount = 0 To UBound(DeleteRow)
    
        sQry = "EXEC PS_GA020_04 '" & DeleteRow(loopCount).DocEntry & "'"
        Call RecordSet01.DoQuery(sQry)
    
    Next
    
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PS_GA020_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA020_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_GA020_AddRowData() As Boolean
'******************************************************************************
'Function ID : PS_GA020_AddRowData()
'해당모듈 : PS_GA020
'기능 : 행추가 및 데이터 바인딩
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA020_AddRowData_Error
    
    Dim RowCount As Integer
    
    Dim DocEntry As Integer
    Dim BPLID As String '사업장
    Dim ZDPTCD As String '관련자팀(코드)
    Dim ZDPTNM As String '관련자팀(명)
    Dim ZSECCD As String '관련자담당(코드)
    Dim ZSECNM As String '관련자담당(명)
    Dim ZEMPIE As String '관련자사번
    Dim ZEMPNM As String '관련자성명
    Dim ZGODAY As String '기준일자
    Dim ZGOTIME As String '출입시간(Fr)
    Dim ZOUTIME As String '출입시간(To)
    Dim ZUSRNN As String '방문자성명
    Dim ZBIRDT As String '생년월일
    Dim ZBUSNM As String '업체명
    Dim ZRESID As String '주소
    Dim ZCARJO As String '차량종류
    Dim ZCARNO As String '차량번호
    Dim ZMOKJU As String '출입목적
    Dim ZPOINT As String '장소
    Dim ZSPCIU As String '비고
    Dim UserSign As String '입력자

    DocEntry = 0 '행추가시는 관리번호를 "0"으로 세팅 후 입력 시점에 Max(DocEntry) 구해서 각각 UPDATE
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    ZDPTCD = Trim(oForm01.Items("ZDPTCD").Specific.VALUE) '관련자팀(코드)
    ZSECCD = Trim(oForm01.Items("ZSECCD").Specific.VALUE) '관련자담당(코드)
    ZEMPIE = Trim(oForm01.Items("ZEMPIE").Specific.VALUE) '관련자사번
    ZEMPNM = Trim(oForm01.Items("ZEMPNM").Specific.VALUE) '관련자성명
    ZGODAY = Trim(oForm01.Items("ZGODAY").Specific.VALUE) '기준일자
    ZGOTIME = Trim(oForm01.Items("ZGOTIME").Specific.VALUE) '출입시간(Fr)
    ZOUTIME = Trim(oForm01.Items("ZOUTIME").Specific.VALUE) '출입시간(To)
    ZUSRNN = Trim(oForm01.Items("ZUSRNN").Specific.VALUE) '방문자성명
    ZBIRDT = Trim(oForm01.Items("ZBIRDT").Specific.VALUE) '생년월일
    ZBUSNM = Trim(oForm01.Items("ZBUSNM").Specific.VALUE) '업체명
    ZRESID = Trim(oForm01.Items("ZRESID").Specific.VALUE) '주소
    ZCARJO = Trim(oForm01.Items("ZCARJO").Specific.VALUE) '차량종류
    ZCARNO = Trim(oForm01.Items("ZCARNO").Specific.VALUE) '차량번호
    ZMOKJU = Trim(oForm01.Items("ZMOKJU").Specific.VALUE) '출입목적
    ZPOINT = Trim(oForm01.Items("ZPOINT").Specific.VALUE) '장소
    ZSPCIU = Trim(oForm01.Items("ZSPCIU").Specific.VALUE) '비고
    UserSign = Sbo_Company.UserSignature
    
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    RowCount = oMat01.RowCount
    
    If RowCount + 1 > oDS_PS_GA020L.Size Then
        Call oDS_PS_GA020L.InsertRecord(RowCount)
    End If
    
    Call oMat01.AddRow
    oDS_PS_GA020L.offset = RowCount
    
    Call oDS_PS_GA020L.setValue("U_LineNum", RowCount, RowCount + 1)
    Call oDS_PS_GA020L.setValue("U_ColReg01", RowCount, DocEntry) '관리번호
    Call oDS_PS_GA020L.setValue("U_ColReg02", RowCount, BPLID) '사업장
    Call oDS_PS_GA020L.setValue("U_ColDt01", RowCount, ZGODAY) '기준일자
    Call oDS_PS_GA020L.setValue("U_ColReg03", RowCount, ZEMPIE) '관련자사번
    Call oDS_PS_GA020L.setValue("U_ColReg04", RowCount, ZEMPNM) '관련자성명
    Call oDS_PS_GA020L.setValue("U_ColReg05", RowCount, ZDPTCD) '관련자팀(CD)
    Call oDS_PS_GA020L.setValue("U_ColReg06", RowCount, ZSECCD) '관련자담당(CD)
    Call oDS_PS_GA020L.setValue("U_ColReg07", RowCount, ZPOINT) '장소
    Call oDS_PS_GA020L.setValue("U_ColReg08", RowCount, ZUSRNN) '방문자성명
    Call oDS_PS_GA020L.setValue("U_ColDt02", RowCount, ZBIRDT) '생년월일
    Call oDS_PS_GA020L.setValue("U_ColReg10", RowCount, ZBUSNM) '업체명
    Call oDS_PS_GA020L.setValue("U_ColReg11", RowCount, ZRESID) '주소
    Call oDS_PS_GA020L.setValue("U_ColTm01", RowCount, ZGOTIME) '출입시간(Fr)
    Call oDS_PS_GA020L.setValue("U_ColTm02", RowCount, ZOUTIME) '출입시간(To)
    Call oDS_PS_GA020L.setValue("U_ColReg14", RowCount, ZCARJO) '차량종류
    Call oDS_PS_GA020L.setValue("U_ColReg17", RowCount, ZCARNO) '차량번호
    Call oDS_PS_GA020L.setValue("U_ColReg15", RowCount, ZMOKJU) '출입목적
    Call oDS_PS_GA020L.setValue("U_ColReg16", RowCount, ZSPCIU) '방문객비고
    
    Call oMat01.LoadFromDataSource
    Call oMat01.AutoResizeColumns
    
'    Call oMat01.LoadFromDataSource
    

    PS_GA020_AddRowData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_AddRowData_Error:
    PS_GA020_AddRowData = False
    MDC_Com.MDC_GF_Message "PS_GA020_AddRowData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_GA020_UpdateRowData() As Boolean
'******************************************************************************
'Function ID : PS_GA020_UpdateRowData()
'해당모듈 : PS_GA020
'기능 : 행추가 및 데이터 바인딩
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA020_UpdateRowData_Error
    
    Dim RowCount As Integer
    
    Dim DocEntry As Integer
    Dim BPLID As String '사업장
    Dim ZDPTCD As String '관련자팀(코드)
    Dim ZDPTNM As String '관련자팀(명)
    Dim ZSECCD As String '관련자담당(코드)
    Dim ZSECNM As String '관련자담당(명)
    Dim ZEMPIE As String '관련자사번
    Dim ZEMPNM As String '관련자성명
    Dim ZGODAY As String '기준일자
    Dim ZGOTIME As String '출입시간(Fr)
    Dim ZOUTIME As String '출입시간(To)
    Dim ZUSRNN As String '방문자성명
    Dim ZBIRDT As String '생년월일
    Dim ZBUSNM As String '업체명
    Dim ZRESID As String '주소
    Dim ZCARJO As String '차량종류
    Dim ZCARNO As String '차량번호
    Dim ZMOKJU As String '출입목적
    Dim ZPOINT As String '장소
    Dim ZSPCIU As String '비고
    Dim UserSign As String '입력자

    '전역변수에 저장된 행번호 저장
    RowCount = oForm01.Items("RowNo").Specific.VALUE - 1

'    DocEntry = 0 '행추가시는 관리번호를 "0"으로 세팅 후 입력 시점에 Max(DocEntry) 구해서 각각 UPDATE
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    ZDPTCD = Trim(oForm01.Items("ZDPTCD").Specific.VALUE) '관련자팀(코드)
    ZSECCD = Trim(oForm01.Items("ZSECCD").Specific.VALUE) '관련자담당(코드)
    ZEMPIE = Trim(oForm01.Items("ZEMPIE").Specific.VALUE) '관련자사번
    ZEMPNM = Trim(oForm01.Items("ZEMPNM").Specific.VALUE) '관련자성명
    ZGODAY = Trim(oForm01.Items("ZGODAY").Specific.VALUE) '기준일자
    ZGOTIME = Trim(oForm01.Items("ZGOTIME").Specific.VALUE) '출입시간(Fr)
    ZOUTIME = Trim(oForm01.Items("ZOUTIME").Specific.VALUE) '출입시간(To)
    ZUSRNN = Trim(oForm01.Items("ZUSRNN").Specific.VALUE) '방문자성명
    ZBIRDT = Trim(oForm01.Items("ZBIRDT").Specific.VALUE) '생년월일
    ZBUSNM = Trim(oForm01.Items("ZBUSNM").Specific.VALUE) '업체명
    ZRESID = Trim(oForm01.Items("ZRESID").Specific.VALUE) '주소
    ZCARJO = Trim(oForm01.Items("ZCARJO").Specific.VALUE) '차량종류
    ZCARNO = Trim(oForm01.Items("ZCARNO").Specific.VALUE) '차량번호
    ZMOKJU = Trim(oForm01.Items("ZMOKJU").Specific.VALUE) '출입목적
    ZPOINT = Trim(oForm01.Items("ZPOINT").Specific.VALUE) '장소
    ZSPCIU = Trim(oForm01.Items("ZSPCIU").Specific.VALUE) '비고
    UserSign = Sbo_Company.UserSignature

'    Call oDS_PS_GA020L.setValue("U_LineNum", , RowCount + 1)
'    Call oDS_PS_GA020L.setValue("U_ColReg01", RowCount, DocEntry) '관리번호
    Call oDS_PS_GA020L.setValue("U_ColReg02", RowCount, BPLID) '사업장
    Call oDS_PS_GA020L.setValue("U_ColDt01", RowCount, ZGODAY) '기준일자
    Call oDS_PS_GA020L.setValue("U_ColReg03", RowCount, ZEMPIE) '관련자사번
    Call oDS_PS_GA020L.setValue("U_ColReg04", RowCount, ZEMPNM) '관련자성명
    Call oDS_PS_GA020L.setValue("U_ColReg05", RowCount, ZDPTCD) '관련자팀(CD)
    Call oDS_PS_GA020L.setValue("U_ColReg06", RowCount, ZSECCD) '관련자담당(CD)
    Call oDS_PS_GA020L.setValue("U_ColReg07", RowCount, ZPOINT) '장소
    Call oDS_PS_GA020L.setValue("U_ColReg08", RowCount, ZUSRNN) '방문자성명
    Call oDS_PS_GA020L.setValue("U_ColDt02", RowCount, ZBIRDT) '생년월일
    Call oDS_PS_GA020L.setValue("U_ColReg10", RowCount, ZBUSNM) '업체명
    Call oDS_PS_GA020L.setValue("U_ColReg11", RowCount, ZRESID) '주소
    Call oDS_PS_GA020L.setValue("U_ColTm01", RowCount, ZGOTIME) '출입시간(Fr)
    Call oDS_PS_GA020L.setValue("U_ColTm02", RowCount, ZOUTIME) '출입시간(To)
    Call oDS_PS_GA020L.setValue("U_ColReg14", RowCount, ZCARJO) '차량종류
    Call oDS_PS_GA020L.setValue("U_ColReg17", RowCount, ZCARNO) '차량번호
    Call oDS_PS_GA020L.setValue("U_ColReg15", RowCount, ZMOKJU) '출입목적
    Call oDS_PS_GA020L.setValue("U_ColReg16", RowCount, ZSPCIU) '방문객비고
    
    Call oMat01.LoadFromDataSource
    Call oMat01.AutoResizeColumns
    
'    Call oMat01.LoadFromDataSource
    

    PS_GA020_UpdateRowData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_UpdateRowData_Error:
    PS_GA020_UpdateRowData = False
    MDC_Com.MDC_GF_Message "PS_GA020_UpdateRowData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_GA020_AddData() As Boolean
'******************************************************************************
'Function ID : PS_GA020_AddData()
'해당모듈 : PS_GA020
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA020_AddData_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim BPLID As String '사업장
    Dim ZDPTCD As String '관련자팀(코드)
    Dim ZDPTNM As String '관련자팀(명)
    Dim ZSECCD As String '관련자담당(코드)
    Dim ZSECNM As String '관련자담당(명)
    Dim ZEMPIE As String '관련자사번
    Dim ZEMPNM As String '관련자성명
    Dim ZGODAY As String '기준일자
    Dim ZGOTIME As String '출입시간(Fr)
    Dim ZOUTIME As String '출입시간(To)
    Dim ZUSRNN As String '방문자성명
    Dim ZBIRDT As String '생년월일
    Dim ZBUSNM As String '업체명
    Dim ZRESID As String '주소
    Dim ZCARJO As String '차량종류
    Dim ZCARNO As String '차량번호
    Dim ZMOKJU As String '출입목적
    Dim ZPOINT As String '장소
    Dim ZSPCIU As String '비고
    Dim UserSign As String '입력자
    
    For loopCount = 0 To oMat01.RowCount - 1
        
        DocEntry = Trim(oDS_PS_GA020L.GetValue("U_ColReg01", loopCount)) '관리번호
        BPLID = Trim(oDS_PS_GA020L.GetValue("U_ColReg02", loopCount)) '사업장
        ZDPTCD = Trim(oDS_PS_GA020L.GetValue("U_ColReg05", loopCount)) '관련자팀(코드)
        ZSECCD = Trim(oDS_PS_GA020L.GetValue("U_ColReg06", loopCount)) '관련자담당(코드)
        ZEMPIE = Trim(oDS_PS_GA020L.GetValue("U_ColReg03", loopCount)) '관련자사번
        ZEMPNM = Trim(oDS_PS_GA020L.GetValue("U_ColReg04", loopCount)) '관련자성명
        ZGODAY = Trim(oDS_PS_GA020L.GetValue("U_ColDt01", loopCount)) '기준일자
        ZGOTIME = Trim(oDS_PS_GA020L.GetValue("U_ColTm01", loopCount)) '출입시간(Fr)
        ZOUTIME = Trim(oDS_PS_GA020L.GetValue("U_ColTm02", loopCount)) '출입시간(To)
        ZUSRNN = Trim(oDS_PS_GA020L.GetValue("U_ColReg08", loopCount)) '방문자성명
        ZBIRDT = Trim(oDS_PS_GA020L.GetValue("U_ColDt02", loopCount)) '생년월일
        ZBUSNM = Trim(oDS_PS_GA020L.GetValue("U_ColReg10", loopCount)) '업체명
        ZRESID = Trim(oDS_PS_GA020L.GetValue("U_ColReg11", loopCount)) '주소
        ZCARJO = Trim(oDS_PS_GA020L.GetValue("U_ColReg14", loopCount)) '차량종류
        ZCARNO = Trim(oDS_PS_GA020L.GetValue("U_ColReg17", loopCount)) '차량번호
        ZMOKJU = Trim(oDS_PS_GA020L.GetValue("U_ColReg15", loopCount)) '출입목적
        ZPOINT = Trim(oDS_PS_GA020L.GetValue("U_ColReg07", loopCount)) '장소
        ZSPCIU = Trim(oDS_PS_GA020L.GetValue("U_ColReg16", loopCount)) '비고
        UserSign = Sbo_Company.UserSignature
        
        '관련자팀(명)
        sQry = "            SELECT       T0.U_CodeNm AS [ZDPTNM]"
        sQry = sQry & "  FROM        [@PS_HR200L] AS T0"
        sQry = sQry & "  WHERE       T0.Code = '1'"
        sQry = sQry & "                  AND T0.U_Code = '" & ZDPTCD & "'"
        Call RecordSet01.DoQuery(sQry)
        ZDPTNM = Trim(RecordSet01.Fields(0).VALUE)
        
        '관련자담당(명)
        sQry = "            SELECT       T0.U_CodeNm AS [ZSECNM]"
        sQry = sQry & "  FROM        [@PS_HR200L] AS T0"
        sQry = sQry & "  WHERE       T0.Code = '2'"
        sQry = sQry & "                  AND T0.U_Code = '" & ZSECCD & "'"
        Call RecordSet01.DoQuery(sQry)
        ZSECNM = Trim(RecordSet01.Fields(0).VALUE)
        
        sQry = "                EXEC [PS_GA020_02] "
        sQry = sQry & "'" & DocEntry & "'," '관리번호
        sQry = sQry & "'" & BPLID & "',"  '사업장
        sQry = sQry & "'" & ZDPTCD & "',"  '관련자팀(코드)
        sQry = sQry & "'" & ZDPTNM & "',"  '관련자팀(명)
        sQry = sQry & "'" & ZSECCD & "',"  '관련자담당(코드)
        sQry = sQry & "'" & ZSECNM & "',"  '관련자담당(명)
        sQry = sQry & "'" & ZEMPIE & "',"  '관련자사번
        sQry = sQry & "'" & ZEMPNM & "',"  '관련자성명
        sQry = sQry & "'" & ZGODAY & "',"  '기준일자
        sQry = sQry & "'" & ZGOTIME & "',"  '출입시간(Fr)
        sQry = sQry & "'" & ZOUTIME & "',"  '출입시간(To)
        sQry = sQry & "'" & ZUSRNN & "',"  '방문자성명
        sQry = sQry & "'" & ZBIRDT & "',"  '생년월일
        sQry = sQry & "'" & ZBUSNM & "',"  '업체명
        sQry = sQry & "'" & ZRESID & "',"  '주소
        sQry = sQry & "'" & ZCARJO & "',"  '차량종류
        sQry = sQry & "'" & ZCARNO & "',"  '차량번호
        sQry = sQry & "'" & ZMOKJU & "',"  '출입목적
        sQry = sQry & "'" & ZPOINT & "',"  '장소
        sQry = sQry & "'" & ZSPCIU & "',"  '비고
        sQry = sQry & "'" & UserSign & "'"  'UserSign
        
        Call RecordSet02.DoQuery(sQry)
    
    Next

    'R3데이터 Interface
    If BPLID = "2" Then '기계사업부 일 때만 R3데이터 전달

        Dim Client As String '클라이언트(운영용:210, 테스트용:710)
        Dim ServerIP As String '서버IP(운영용:192.1.11.3, 테스트용:192.1.11.7)
        Dim Plant As String '공장코드(부산사업장:9300)
        Dim E_MESSAGE As String '에러메시지

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
            GoTo PS_GA020_AddData_Error
        End If

        '2. SAP R3 함수 호출
        Dim oFunction As New SAPFunctionsOCX.Function 'Object 'R3함수용 Function 객체
        Set oFunction = oSapConnection.Add("ZHR_INTF_VISIT")

        oFunction.Exports("I_WERKS") = Plant '플랜트 : 기계사업부만 등록하므로 9300으로 Fix
        oFunction.Exports("I_ZDATE") = ZGODAY '기준일자
        oFunction.Exports("I_ZEMPIE") = ZEMPIE '관련자사번

        '2-1. 테이블
        Dim oTable As Table
        
        '-------------------------------------------------------------
        ' SAP함수 호출한 객체에서 테이블 객체 불러와서 값 할당
        '-------------------------------------------------------------
        Set oTable = oFunction.Tables("T_ZVISIT")

        For loopCount = 0 To oMat01.RowCount - 1

            '관련자담당(명)
            sQry = "            SELECT       T0.U_CodeNm AS [ZSECNM]"
            sQry = sQry & "  FROM        [@PS_HR200L] AS T0"
            sQry = sQry & "  WHERE       T0.Code = '2'"
            sQry = sQry & "                  AND T0.U_Code = '" & oDS_PS_GA020L.GetValue("U_ColReg06", loopCount) & "'"
            Call RecordSet01.DoQuery(sQry)
            ZSECNM = Trim(RecordSet01.Fields(0).VALUE)

            oTable.Rows.Add
            oTable(loopCount + 1, "MANDT") = Client '클라이언트
            oTable(loopCount + 1, "WERKS") = Plant '플랜트
            oTable(loopCount + 1, "ZSAGU") = "3" '사업장(R3) '풍산홀딩스 부산사업장(3)으로 Fix
            oTable(loopCount + 1, "ZGODAY") = oDS_PS_GA020L.GetValue("U_ColDt01", loopCount) '기준일자
            oTable(loopCount + 1, "ZSEQ") = "" '순번
            oTable(loopCount + 1, "ZUPSEQ") = "" '업체순번
            oTable(loopCount + 1, "ZEMPIE") = oDS_PS_GA020L.GetValue("U_ColReg03", loopCount) '관련자사번
            oTable(loopCount + 1, "ZEMPNM") = oDS_PS_GA020L.GetValue("U_ColReg04", loopCount) '관련자성명
            oTable(loopCount + 1, "ZDPTCD") = oDS_PS_GA020L.GetValue("U_ColReg05", loopCount) '부서코드
            oTable(loopCount + 1, "ZSECNM") = ZSECNM '담당명
            oTable(loopCount + 1, "ZPERNP") = "" '확인자사번
            oTable(loopCount + 1, "ZUSRNN") = oDS_PS_GA020L.GetValue("U_ColReg08", loopCount) '방문자성명
            oTable(loopCount + 1, "ZBUSNM") = oDS_PS_GA020L.GetValue("U_ColReg10", loopCount) '업체명
            oTable(loopCount + 1, "ZBIRDT") = oDS_PS_GA020L.GetValue("U_ColDt02", loopCount)  '생년월일
            oTable(loopCount + 1, "ZRESID") = oDS_PS_GA020L.GetValue("U_ColReg11", loopCount)  '주소
            oTable(loopCount + 1, "ZCARJO") = oDS_PS_GA020L.GetValue("U_ColReg14", loopCount)  '차량종류
            oTable(loopCount + 1, "ZCARNO") = oDS_PS_GA020L.GetValue("U_ColReg17", loopCount)  '차량번호
            oTable(loopCount + 1, "ZPOINT") = oDS_PS_GA020L.GetValue("U_ColReg07", loopCount)  '장소
            oTable(loopCount + 1, "ZMOKJU") = oDS_PS_GA020L.GetValue("U_ColReg15", loopCount)  '출입목적
            oTable(loopCount + 1, "ZGOTIME") = oDS_PS_GA020L.GetValue("U_ColTm01", loopCount) & "00"  '출입시간(시작)
            oTable(loopCount + 1, "ZOUTIME") = oDS_PS_GA020L.GetValue("U_ColTm02", loopCount) & "00"  '출입시간(종료)
            oTable(loopCount + 1, "ZSPCIU") = oDS_PS_GA020L.GetValue("U_ColReg16", loopCount)  '방문객비고
            oTable(loopCount + 1, "ZSUNIN") = "" '승인
            oTable(loopCount + 1, "ZSUNTXT") = "" '승인여부
            oTable(loopCount + 1, "ZPLG") = "" '추가,삭제
            oTable(loopCount + 1, "XLOEK") = "" '품목 삭제

        Next

        '-------------------------------------------------------------
        ' 함수를 Call 하면 oTable 객체의 내용을 그대로 SAP R3에 전달.
        ' 그러면 SAP R3는 oTable의 각각의 값들에 대해서 처리 결과를 리턴.
        '-------------------------------------------------------------
        If Not (oFunction.Call) Then

            oSapConnection.Connection.Logoff
            ErrNum = 99
            GoTo PS_GA020_AddData_Error

        Else

            E_MESSAGE = Trim(oFunction.Imports("E_MESSAGE").VALUE) '메시지

        End If

        '행 삭제 처리한 내역 DB에서 DELETE(행삭제한 내용이 존재할 때만)
        If gDeleteRowCount > 0 Then
        
            For loopCount = 0 To UBound(DeleteRow)
        
                sQry = "EXEC PS_GA020_04 '" & DeleteRow(loopCount).DocEntry & "'"
                Call RecordSet01.DoQuery(sQry)
        
            Next
            
        End If

    End If

    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    oSapConnection.Connection.Logoff
    
    PS_GA020_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_AddData_Error:

    If ErrNum = 1 Then
    
        Call MDC_Com.MDC_GF_Message("풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.", "E")

    ElseIf ErrNum = 99 Then
    
        Call MDC_Com.MDC_GF_Message(oFunction.Exception, "E")
        
    Else
    
        Call MDC_Com.MDC_GF_Message("PS_GA020_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
        
    End If
    
    PS_GA020_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    oSapConnection.Connection.Logoff
    
    
End Function

Private Function PS_GA020_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_GA020_HeaderSpaceLineDel()
'해당모듈 : PS_GA020
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA020_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("BPLID").Specific.VALUE) = "%" '사업장
            ErrNum = 1
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZDPTCD").Specific.VALUE) = "%" '관련자팀
            ErrNum = 2
            GoTo PS_GA020_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("ZSECCD").Specific.VALUE) = "%" '관련자담당
'            ErrNum = 3
'            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZEMPIE").Specific.VALUE) = "" '관련자사번
            ErrNum = 4
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZGODAY").Specific.VALUE) = "" '기준일자
            ErrNum = 5
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZGOTIME").Specific.VALUE) = "" '출입시간(Fr)
            ErrNum = 6
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZOUTIME").Specific.VALUE) = "" '출입시간(To)
            ErrNum = 7
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZUSRNN").Specific.VALUE) = "" '방문자성명
            ErrNum = 8
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZBUSNM").Specific.VALUE) = "" '업체명
            ErrNum = 9
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZRESID").Specific.VALUE) = "" '주소
            ErrNum = 10
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZCARJO").Specific.VALUE) = "%" '차량종류
            ErrNum = 11
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZMOKJU").Specific.VALUE) = "%" '출입목적
            ErrNum = 12
            GoTo PS_GA020_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZPOINT").Specific.VALUE) = "%" '장소
            ErrNum = 13
            GoTo PS_GA020_HeaderSpaceLineDel_Error
    End Select

    PS_GA020_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("BPLID").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "관련자팀은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZDPTCD").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "관련자담당은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZSECCD").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "관련자사번은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZEMPIE").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "기준일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZGODAY").Click(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "출입시간(시작)은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZGOTIME").Click(ct_Regular)
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "출입시간(종료)은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZOUTIME").Click(ct_Regular)
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "방문자성명은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZUSRNN").Click(ct_Regular)
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "업체명은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZBUSNM").Click(ct_Regular)
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "주소는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZRESID").Click(ct_Regular)
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "차량종류는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZCARJO").Click(ct_Regular)
    ElseIf ErrNum = 12 Then
        MDC_Com.MDC_GF_Message "출입목적은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZMOKJU").Click(ct_Regular)
    ElseIf ErrNum = 13 Then
        MDC_Com.MDC_GF_Message "장소는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZPOINT").Click(ct_Regular)
    End If
    PS_GA020_HeaderSpaceLineDel = False
End Function


Private Function PS_GA020_CheckBeforeSearch() As Boolean
'******************************************************************************
'Function ID : PS_GA020_CheckBeforeSearch()
'해당모듈 : PS_GA020
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA020_CheckBeforeSearch_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("SZGODAY").Specific.VALUE) = "" '기준일자
            ErrNum = 1
            GoTo PS_GA020_CheckBeforeSearch_Error
        Case Trim(oForm01.Items("SBPLID").Specific.VALUE) = "%" '사업장
            ErrNum = 2
            GoTo PS_GA020_CheckBeforeSearch_Error
        Case Trim(oForm01.Items("SZDPTCD").Specific.VALUE) = "%" '관련자팀
            ErrNum = 3
            GoTo PS_GA020_CheckBeforeSearch_Error
'        Case Trim(oForm01.Items("SZSECCD").Specific.VALUE) = "%" '관련자담당
'            ErrNum = 4
'            GoTo PS_GA020_CheckBeforeSearch_Error
        Case Trim(oForm01.Items("SZEMPIE").Specific.VALUE) = "" '관련자사번
            ErrNum = 5
            GoTo PS_GA020_CheckBeforeSearch_Error
    End Select

    PS_GA020_CheckBeforeSearch = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_CheckBeforeSearch_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "기준일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("SZGODAY").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("SBPLID").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "관련자팀은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("SZDPTCD").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "관련자담당은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("SZSECCD").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "관련자사번은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("SZEMPIE").Click(ct_Regular)
    End If
    PS_GA020_CheckBeforeSearch = False
End Function


'/ 메트릭스 필수 사항 check
Private Function PS_GA020_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_GA020_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PS_GA020_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PS_GA020_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_GA020_MatrixSpaceLineDel = False
End Function

Private Sub PS_GA020_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_GA020_FlushToItemValue_Error

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

        Case "BPLID"
            
            'BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE)
            BPLID = Trim(oDS_PS_GA020H.GetValue("U_BPLID", 0))
            
            If oForm01.Items("ZDPTCD").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("ZDPTCD").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("ZDPTCD").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '부서콤보세팅
            Call oForm01.Items("ZDPTCD").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '1'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char2 = '" & BPLID & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("ZDPTCD").Specific, sQry, "", False, False)
            Call oForm01.Items("ZDPTCD").Specific.Select(0, psk_Index)
            
        Case "ZDPTCD"
            
            'ZDPTCD = Trim(oForm01.Items("ZDPTCD").Specific.VALUE)
            ZDPTCD = Trim(oDS_PS_GA020H.GetValue("U_ZDPTCD", 0))
            
            If oForm01.Items("ZSECCD").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("ZSECCD").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("ZSECCD").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '담당콤보세팅
            Call oForm01.Items("ZSECCD").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '2'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char1 = '" & ZDPTCD & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("ZSECCD").Specific, sQry, "", False, False)
            Call oForm01.Items("ZSECCD").Specific.Select(0, psk_Index)
        
        Case "SBPLID"
        
            BPLID = Trim(oForm01.Items("SBPLID").Specific.VALUE)
            
            If oForm01.Items("SZDPTCD").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("SZDPTCD").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("SZDPTCD").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '부서콤보세팅
            Call oForm01.Items("SZDPTCD").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '1'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char2 = '" & BPLID & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("SZDPTCD").Specific, sQry, "", False, False)
            Call oForm01.Items("SZDPTCD").Specific.Select(0, psk_Index)
        
        Case "SZDPTCD"
        
            ZDPTCD = Trim(oForm01.Items("SZDPTCD").Specific.VALUE)
            
            If oForm01.Items("SZSECCD").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("SZSECCD").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("SZSECCD").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '담당콤보세팅
            Call oForm01.Items("SZSECCD").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '2'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char1 = '" & ZDPTCD & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("SZSECCD").Specific, sQry, "", False, False)
            Call oForm01.Items("SZSECCD").Specific.Select(0, psk_Index)
        
        Case "ZEMPIE"
        
            Call oDS_PS_GA020H.setValue("U_ZEMPNM", 0, MDC_GetData.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" & oForm01.Items("ZEMPIE").Specific.VALUE & "'")) '성명
        
        Case "SZEMPIE"
        
            oForm01.Items("SZEMPNM").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" & oForm01.Items("SZEMPIE").Specific.VALUE & "'") '성명
        
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_GA020_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("에러 메시지", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA020_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_GA020_FormItemEnabled()
On Error GoTo PS_GA020_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_GA020_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
                Call PS_GA020_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_GA020_LoadCaption
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
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
                '//Call PS_GA020_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_GA020H.Clear

'                Call PS_GA020_LoadCaption
'                Call PS_GA020_FormItemEnabled
                '//Call PS_GA020_FormItemEnabled '//UDO방식
                '//Call PS_GA020_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_GA020_FormItemEnabled
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
    
        If pval.ItemUID = "PS_GA020" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAddRow" Then '행추가(행수정)
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_GA020_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If PS_GA020_AddRowData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_GA020_LoadCaption
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_GA020_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                If PS_GA020_UpdateRowData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
'                Call PS_GA020_FormReset
'                oForm01.Mode = fm_ADD_MODE
                
'                Call PS_GA020_LoadCaption
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnAdd" Then '추가(수정) 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
'                If PS_GA020_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
                
                If PS_GA020_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_GA020_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA020_LoadCaption
'                Call PS_GA020_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
'                If PS_GA020_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                '쿼리에서 로직 처리
                '여러 행을 동시에 처리하므로 DocEntry별로 행이 존재하면 UPDATE, 존재하지 않으면 INSERT
                If PS_GA020_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_GA020_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA020_LoadCaption
'                Call PS_GA020_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회

            If PS_GA020_CheckBeforeSearch = False Then
                BubbleEvent = False
                Exit Sub
            End If

            Call PS_GA020_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_GA020_LoadCaption
            Call PS_GA020_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_GA020_DeleteData
                Call PS_GA020_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_GA020_LoadCaption
'                Call PS_GA020_MTX01
                
            Else
            
            End If
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA020" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ZEMPIE", "") '//사용자값활성
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
    
        
        
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row > 0 Then
            
                Call oMat01.SelectRow(pval.Row, True, False)
                
                Call oForm01.Freeze(True)

                'DataSource를 이용하여 각 컨트롤에 값을 출력
                Call oDS_PS_GA020H.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
                Call oDS_PS_GA020H.setValue("U_BPLID", 0, oMat01.Columns("BPLID").Cells(pval.Row).Specific.VALUE) '사업장
                Call PS_GA020_FlushToItemValue("BPLID") '관련자사업장에 따른 팀 세팅
                Call oDS_PS_GA020H.setValue("U_ZDPTCD", 0, oMat01.Columns("ZDPTCD").Cells(pval.Row).Specific.VALUE) '관련자팀
                Call PS_GA020_FlushToItemValue("ZDPTCD") '관련자팀에 따른 담당 세팅
                Call oDS_PS_GA020H.setValue("U_ZSECCD", 0, oMat01.Columns("ZSECCD").Cells(pval.Row).Specific.VALUE) '관련자담당
                Call oDS_PS_GA020H.setValue("U_ZEMPIE", 0, oMat01.Columns("ZEMPIE").Cells(pval.Row).Specific.VALUE) '관련자사번
                Call oDS_PS_GA020H.setValue("U_ZEMPNM", 0, oMat01.Columns("ZEMPNM").Cells(pval.Row).Specific.VALUE) '관련자성명
                Call oDS_PS_GA020H.setValue("U_ZGODAY", 0, oMat01.Columns("ZGODAY").Cells(pval.Row).Specific.VALUE) '기준일자
                Call oDS_PS_GA020H.setValue("U_ZGOTIME", 0, oMat01.Columns("ZGOTIME").Cells(pval.Row).Specific.VALUE) '출입시간(Fr)
                Call oDS_PS_GA020H.setValue("U_ZOUTIME", 0, oMat01.Columns("ZOUTIME").Cells(pval.Row).Specific.VALUE) '출입시간(To)
                Call oDS_PS_GA020H.setValue("U_ZUSRNN", 0, oMat01.Columns("ZUSRNN").Cells(pval.Row).Specific.VALUE) '방문자성명
                Call oDS_PS_GA020H.setValue("U_ZBIRDT", 0, oMat01.Columns("ZBIRDT").Cells(pval.Row).Specific.VALUE) '생년월일
                Call oDS_PS_GA020H.setValue("U_ZBUSNM", 0, oMat01.Columns("ZBUSNM").Cells(pval.Row).Specific.VALUE) '업체명
                Call oDS_PS_GA020H.setValue("U_ZRESID", 0, oMat01.Columns("ZRESID").Cells(pval.Row).Specific.VALUE) '주소
                Call oDS_PS_GA020H.setValue("U_ZCARJO", 0, oMat01.Columns("ZCARJO").Cells(pval.Row).Specific.VALUE) '차량종류
                Call oDS_PS_GA020H.setValue("U_ZCARNO", 0, oMat01.Columns("ZCARNO").Cells(pval.Row).Specific.VALUE) '차량목적
                Call oDS_PS_GA020H.setValue("U_ZMOKJU", 0, oMat01.Columns("ZMOKJU").Cells(pval.Row).Specific.VALUE) '출입목적
                Call oDS_PS_GA020H.setValue("U_ZPOINT", 0, oMat01.Columns("ZPOINT").Cells(pval.Row).Specific.VALUE) '장소
                Call oDS_PS_GA020H.setValue("U_ZSPCIU", 0, oMat01.Columns("ZSPCIU").Cells(pval.Row).Specific.VALUE) '비고
                
                oForm01.Mode = fm_UPDATE_MODE
                Call PS_GA020_LoadCaption
                
                '선택한 행의 번호를 임시 컨트롤에 저장
                oForm01.Items("RowNo").Specific.VALUE = pval.Row
                
                Call oForm01.Freeze(False)
                
            End If
        End If
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PS_GA020_FlushToItemValue(pval.ItemUID)
    
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
        
            Call PS_GA020_FlushToItemValue(pval.ItemUID)
                
        End If
        
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
        
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
        Call PS_GA020_FormItemEnabled
        '//Call PS_GA020_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_GA020_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_GA020H", "U_CardCode,U_CardName")
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
'            If (PS_GA020_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
            
            gDeleteRowCount = gDeleteRowCount + 1
            
            '삭제한 행의 DocEntry 저장_S
            ReDim Preserve DeleteRow(gDeleteRowCount - 1)
            DeleteRow(gDeleteRowCount - 1).DocEntry = oDS_PS_GA020L.GetValue("U_ColReg01", CInt(Trim(oForm01.Items("RowNo").Specific.VALUE)) - 1)
            '삭제한 행의 DocEntry 저장_E
            
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_GA020L.RemoveRecord(oDS_PS_GA020L.Size - 1)
            oMat01.LoadFromDataSource
            

            
'            If oMat01.RowCount = 0 Then
'                Call PS_GA020_Add_MatrixRow(0)
'            Else
'                If Trim(oDS_PS_GA020L.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_GA020_Add_MatrixRow(oMat01.RowCount)
'                End If
'            End If

        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_GA020_CreateItems() As Boolean
On Error GoTo PS_GA020_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_GA020H = oForm01.DataSources.DBDataSources("@PS_GA020H")
    Set oDS_PS_GA020L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("SBPLID", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SBPLID").Specific.DataBind.SetBound(True, "", "SBPLID")
    '사업장_E
    
    '기준일자(Fr)
    Call oForm01.DataSources.UserDataSources.Add("SZGODAY", dt_DATE)
    Call oForm01.Items("SZGODAY").Specific.DataBind.SetBound(True, "", "SZGODAY")
    
'    '기준일자(To)
'    Call oForm01.DataSources.UserDataSources.Add("SZGODAYT", dt_DATE)
'    Call oForm01.Items("SZGODAYT").Specific.DataBind.SetBound(True, "", "SZGODAYT")
    
    '관련자사번
    Call oForm01.DataSources.UserDataSources.Add("SZEMPIE", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SZEMPIE").Specific.DataBind.SetBound(True, "", "SZEMPIE")
    
    '관련자성명
    Call oForm01.DataSources.UserDataSources.Add("SZEMPNM", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SZEMPNM").Specific.DataBind.SetBound(True, "", "SZEMPNM")
    
    '관련자팀
    Call oForm01.DataSources.UserDataSources.Add("SZDPTCD", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SZDPTCD").Specific.DataBind.SetBound(True, "", "SZDPTCD")
    
    '관련자담당
    Call oForm01.DataSources.UserDataSources.Add("SZSECCD", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SZSECCD").Specific.DataBind.SetBound(True, "", "SZSECCD")
    
'    '방문자성명
'    Call oForm01.DataSources.UserDataSources.Add("SZUSRNN", dt_SHORT_TEXT, 20)
'    Call oForm01.Items("SZUSRNN").Specific.DataBind.SetBound(True, "", "SZUSRNN")
'
'    '업체명
'    Call oForm01.DataSources.UserDataSources.Add("SZBUSNM", dt_SHORT_TEXT, 50)
'    Call oForm01.Items("SZBUSNM").Specific.DataBind.SetBound(True, "", "SZBUSNM")
'
'    '차량종류
'    Call oForm01.DataSources.UserDataSources.Add("SZCARJO", dt_SHORT_TEXT, 20)
'    Call oForm01.Items("SZCARJO").Specific.DataBind.SetBound(True, "", "SZCARJO")
'
'    '출입목적
'    Call oForm01.DataSources.UserDataSources.Add("SZMOKJU", dt_SHORT_TEXT, 20)
'    Call oForm01.Items("SZMOKJU").Specific.DataBind.SetBound(True, "", "SZMOKJU")
'
'    '장소
'    Call oForm01.DataSources.UserDataSources.Add("SZPOINT", dt_SHORT_TEXT, 20)
'    Call oForm01.Items("SZPOINT").Specific.DataBind.SetBound(True, "", "SZPOINT")

    '선택행번호
    Call oForm01.DataSources.UserDataSources.Add("RowNo", dt_SHORT_TEXT, 20)
    Call oForm01.Items("RowNo").Specific.DataBind.SetBound(True, "", "RowNo")
    
    

    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_GA020_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA020_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_GA020_ComboBox_Setting()
On Error GoTo PS_GA020_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '//////////기본정보//////////S
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("BPLID").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '관련자팀(FlushToItemValue)
    
    '관련자담당(FlushToItemValue)
    
    '차량종류
    Call oForm01.Items("ZCARJO").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT      U_Code AS [Code],"
    sQry = sQry & "                U_CodeNm As Name"
    sQry = sQry & " FROM       [@PS_GA050L]"
    sQry = sQry & " WHERE      Code = 'G002'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ZCARJO").Specific, sQry, "1", False, False)
    Call oForm01.Items("ZCARJO").Specific.Select("%", psk_ByValue)
    
    '장소(ZPOINT)
    Call oForm01.Items("ZPOINT").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT      U_Code AS [Code],"
    sQry = sQry & "                U_CodeNm As Name"
    sQry = sQry & " FROM       [@PS_GA050L]"
    sQry = sQry & " WHERE      Code = 'G003'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ZPOINT").Specific, sQry, "1", False, False)
    Call oForm01.Items("ZPOINT").Specific.Select("%", psk_ByValue)
      
    '출입목적
    Call oForm01.Items("ZMOKJU").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT      U_Code AS [Code],"
    sQry = sQry & "                U_CodeNm As Name"
    sQry = sQry & " FROM       [@PS_GA050L]"
    sQry = sQry & " WHERE      Code = 'G004'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ZMOKJU").Specific, sQry, "1", False, False)
    Call oForm01.Items("ZMOKJU").Specific.Select("%", psk_ByValue)
    '//////////기본정보//////////E
    
    '//////////조회정보//////////S
    '사업장
    Call oForm01.Items("SBPLID").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("SBPLID").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
'    '차량종류
'    Call oForm01.Items("SZCARJO").Specific.ValidValues.Add("%", "전체")
'    sQry = "           SELECT      U_Code AS [Code],"
'    sQry = sQry & "                U_CodeNm As Name"
'    sQry = sQry & " FROM       [@PS_GA050L]"
'    sQry = sQry & " WHERE      Code = 'G002'"
'    sQry = sQry & "                AND U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY  U_Seq"
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("SZCARJO").Specific, sQry, "1", False, False)
'    Call oForm01.Items("SZCARJO").Specific.Select("%", psk_ByValue)
'
'    '장소(SZPOINT)
'    Call oForm01.Items("SZPOINT").Specific.ValidValues.Add("%", "전체")
'    sQry = "           SELECT      U_Code AS [Code],"
'    sQry = sQry & "                U_CodeNm As Name"
'    sQry = sQry & " FROM       [@PS_GA050L]"
'    sQry = sQry & " WHERE      Code = 'G003'"
'    sQry = sQry & "                AND U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY  U_Seq"
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("SZPOINT").Specific, sQry, "1", False, False)
'    Call oForm01.Items("SZPOINT").Specific.Select("%", psk_ByValue)
'
'    '출입목적
'    Call oForm01.Items("SZMOKJU").Specific.ValidValues.Add("%", "전체")
'    sQry = "           SELECT      U_Code AS [Code],"
'    sQry = sQry & "                U_CodeNm As Name"
'    sQry = sQry & " FROM       [@PS_GA050L]"
'    sQry = sQry & " WHERE      Code = 'G004'"
'    sQry = sQry & "                AND U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY  U_Seq"
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("SZMOKJU").Specific, sQry, "1", False, False)
'    Call oForm01.Items("SZMOKJU").Specific.Select("%", psk_ByValue)
    '//////////조회정보//////////E
    
    '//////////매트릭스//////////S
    '사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLID"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    '관련자팀
    sQry = "            SELECT      U_Code AS [Code],"
    sQry = sQry & "                 U_CodeNm As [Name]"
    sQry = sQry & "  FROM       [@PS_HR200L]"
    sQry = sQry & "  WHERE      Code = '1'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ZDPTCD"), sQry)
    
    '관련자담당
    sQry = "            SELECT      U_Code AS [Code],"
    sQry = sQry & "                 U_CodeNm As [Name]"
    sQry = sQry & "  FROM       [@PS_HR200L]"
    sQry = sQry & "  WHERE      Code = '2'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ZSECCD"), sQry)
    
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
    
    '//////////매트릭스//////////E

    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_GA020_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA020_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA020_CF_ChooseFromList()
On Error GoTo PS_GA020_CF_ChooseFromList_Error

    Exit Sub
PS_GA020_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA020_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA020_EnableMenus()
On Error GoTo PS_GA020_EnableMenus_Error
    
    Exit Sub
PS_GA020_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA020_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA020_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_GA020_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_GA020_FormItemEnabled
        '//Call PS_GA020_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_GA020_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_GA020_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA020_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA020_FormResize()
On Error GoTo PS_GA020_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_GA020_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA020_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_GA020_FormReset()
'******************************************************************************
'Function ID : PS_GA020_FormReset()
'해당모듈 : PS_GA020
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA020_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
'    '관리번호
'    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_GA020H]"
'    Call RecordSet01.DoQuery(sQry)
'
'    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'        Call oDS_PS_GA020H.setValue("DocEntry", 0, 1)
'    Else
'        Call oDS_PS_GA020H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
'    End If
    
    Call oDS_PS_GA020H.setValue("DocEntry", 0, "") '관리번호
    Call oDS_PS_GA020H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId) '사업장
    Call oDS_PS_GA020H.setValue("U_ZGODAY", 0, Format(Date, "YYYYMMDD")) '기준일자
    Call oDS_PS_GA020H.setValue("U_ZEMPIE", 0, "") '관련자사번
    Call oDS_PS_GA020H.setValue("U_ZEMPNM", 0, "") '관련자성명
    Call PS_GA020_FlushToItemValue("BPLID") '관련자사업장에 따른 팀 세팅
    Call oDS_PS_GA020H.setValue("U_ZDPTCD", 0, "%") '관련자팀
    Call PS_GA020_FlushToItemValue("ZDPTCD") '관련자팀에 따른 담당 세팅
    Call oDS_PS_GA020H.setValue("U_ZSECCD", 0, "%") '관련자담당
    Call oDS_PS_GA020H.setValue("U_ZPOINT", 0, "%") '장소
    Call oDS_PS_GA020H.setValue("U_ZUSRNN", 0, "") '방문자성명
    Call oDS_PS_GA020H.setValue("U_ZBIRDT", 0, "") '생년월일
    Call oDS_PS_GA020H.setValue("U_ZBUSNM", 0, "") '업체명
    Call oDS_PS_GA020H.setValue("U_ZRESID", 0, "") '주소
    Call oDS_PS_GA020H.setValue("U_ZGOTIME", 0, "") '출입시간(Fr)
    Call oDS_PS_GA020H.setValue("U_ZOUTIME", 0, "") '출입시간(To)
    Call oDS_PS_GA020H.setValue("U_ZCARJO", 0, "%") '차량종류
    Call oDS_PS_GA020H.setValue("U_ZCARNO", 0, "") '차량번호
    Call oDS_PS_GA020H.setValue("U_ZMOKJU", 0, "%") '출입목적
    Call oDS_PS_GA020H.setValue("U_ZSPCIU", 0, "") '방문객비고
    
    gDeleteRowCount = 0
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA020_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA020_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


