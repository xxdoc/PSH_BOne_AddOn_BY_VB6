VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PH_PY030.cls
'//  Module             : 인사관리>기타관리
'//  Desc               : 공용등록
'//  FormType           : PH_PY030
'//  Create Date(Start) : 2013.03.06
'//  Create Date(End)   : 2013.03.16
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01  As String
Public oForm            As SAPbouiCOM.Form
Public oMat01           As SAPbouiCOM.Matrix
Private oDS_PH_PY030A   As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PH_PY030B   As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01  As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01   As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01   As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode      As Long


'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i         As Long
    Dim oInnerXml As String
    Dim oXmlDoc   As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY030.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY030_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01, "PH_PY030" '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    '//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm.Freeze True
    Call PH_PY030_CreateItems
    Call PH_PY030_ComboBox_Setting
    Call PH_PY030_CF_ChooseFromList
    Call PH_PY030_EnableMenus
    Call PH_PY030_SetDocument(oFromDocEntry01)
    Call PH_PY030_FormResize
    

'    Call PH_PY030_Add_MatrixRow(0, True)
    Call PH_PY030_LoadCaption
    Call PH_PY030_FormItemEnabled
    
    oForm.EnableMenu ("1283"), False        '// 삭제
    oForm.EnableMenu ("1286"), False        '// 닫기
    oForm.EnableMenu ("1287"), False        '// 복제
    oForm.EnableMenu ("1285"), False        '// 복원
    oForm.EnableMenu ("1284"), False         '// 취소
    oForm.EnableMenu ("1293"), False         '// 행삭제
    oForm.EnableMenu ("1281"), False
    oForm.EnableMenu ("1282"), True

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
        
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY030A]"
    Call RecordSet01.DoQuery(sQry)
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PH_PY030A.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PH_PY030A.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    oMat01.Columns("Check").Visible = False '선택 체크박스 Visible = False
    
    Call PH_PY030_FormReset '폼초기화 추가(2013.01.29 송명규)
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing

    '기간
    oForm.Items("SFrDate").Specific.VALUE = Format(Now, "yyyyMMdd")
    oForm.Items("SToDate").Specific.VALUE = Format(Now, "yyyyMMdd")
    '사번 포커스
    Call oForm.Items("MSTCOD").CLICK
    
    Exit Sub
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY030_LoadCaption()
'******************************************************************************
'Function ID : PH_PY030_LoadCaption()
'해당모듈    : PH_PY030
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_LoadCaption_Error

    Call oForm.Freeze(True)

    If oForm.Mode = fm_ADD_MODE Then
        oForm.Items("BtnAdd").Specific.Caption = "추가"
        oForm.Items("BtnDelete").Enabled = False
'    ElseIf oForm.Mode = fm_OK_MODE Then
'        oForm.Items("BtnAdd").Specific.Caption = "확인"
    ElseIf oForm.Mode = fm_UPDATE_MODE Then
        oForm.Items("BtnAdd").Specific.Caption = "수정"
        oForm.Items("BtnDelete").Enabled = True
    End If
    
    Call oForm.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY030_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PH_PY030_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PH_PY030_Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PH_PY030B.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PH_PY030B.Offset = oRow
    oDS_PH_PY030B.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY030_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PH_PY030_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PH_PY030_MTX01()
'******************************************************************************
'Function ID : PH_PY030_MTX01()
'해당모듈    : PH_PY030
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_MTX01_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim sDocEntry   As String '관리번호
    Dim sCLTCOD     As String '사업장
    Dim SDestNo1    As String '출장번호1
    Dim SDestNo2    As String '출장번호2
    Dim sMSTCOD     As String '사원번호
    Dim SDestinat   As String '출장지
    Dim SDest2      As String '출장지상세
    Dim SCoCode     As String '작번
    Dim SFrDate     As String '시작일자
    Dim SToDate     As String '종료일자
    Dim SObject     As String '목적
    Dim SComments   As String '비고
    Dim SRegCls     As String '등록구분
    Dim SObjCls     As String '목적구분
    Dim SDestCode   As String '출장지역
    Dim SDestDiv    As String '출장구분
    Dim SVehicle    As String '차량구분
    Dim sTeamCode   As String '팀(2013.06.01 송명규 추가)

    sDocEntry = Trim(oForm.Items("SDocEntry").Specific.VALUE) '관리번호
    sCLTCOD = Trim(oForm.Items("SCLTCOD").Specific.VALUE) '사업장
    SDestNo1 = Trim(oForm.Items("SDestNo1").Specific.VALUE) '출장번호1
    SDestNo2 = Trim(oForm.Items("SDestNo2").Specific.VALUE) '출장번호2
    sMSTCOD = Trim(oForm.Items("SMSTCOD").Specific.VALUE) '사원번호
    SDestinat = Trim(oForm.Items("SDestinat").Specific.VALUE) '출장지
    SDest2 = Trim(oForm.Items("SDest2").Specific.VALUE) '출장지상세
    SCoCode = Trim(oForm.Items("SCoCode").Specific.VALUE) '작번
    SFrDate = Replace(Trim(oForm.Items("SFrDate").Specific.VALUE), ".", "") '시작일자
    SToDate = Replace(Trim(oForm.Items("SToDate").Specific.VALUE), ".", "") '종료일자
    SObject = Trim(oForm.Items("SObject").Specific.VALUE) '목적
    SComments = Trim(oForm.Items("SComments").Specific.VALUE) '비고
    SRegCls = Trim(oForm.Items("SRegCls").Specific.VALUE) '등록구분
    SObjCls = Trim(oForm.Items("SObjCls").Specific.VALUE) '목적구분
    SDestCode = Trim(oForm.Items("SDestCode").Specific.VALUE) '출장지역
    SDestDiv = Trim(oForm.Items("SDestDiv").Specific.VALUE) '출장구분
    SVehicle = Trim(oForm.Items("SVehicle").Specific.VALUE) '차량구분
    sTeamCode = Trim(oForm.Items("STeamCode").Specific.VALUE) '팀
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "            EXEC [PH_PY030_01] "
    sQry = sQry & "'" & sDocEntry & "'," '관리번호
    sQry = sQry & "'" & sCLTCOD & "',"  '사업장
    sQry = sQry & "'" & SDestNo1 & "',"  '출장번호1
    sQry = sQry & "'" & SDestNo2 & "',"  '출장번호2
    sQry = sQry & "'" & sMSTCOD & "',"  '사원번호
    sQry = sQry & "'" & SDestinat & "',"  '출장지
    sQry = sQry & "'" & SDest2 & "',"  '출장지상세
    sQry = sQry & "'" & SCoCode & "',"  '작번
    sQry = sQry & "'" & SFrDate & "',"  '시작일자
    sQry = sQry & "'" & SToDate & "',"  '종료일자
    sQry = sQry & "'" & SObject & "',"  '목적
    sQry = sQry & "'" & SComments & "',"  '비고
    sQry = sQry & "'" & SRegCls & "',"  '등록구분
    sQry = sQry & "'" & SObjCls & "',"  '목적구분
    sQry = sQry & "'" & SDestCode & "',"  '출장지역
    sQry = sQry & "'" & SDestDiv & "',"  '출장구분
    sQry = sQry & "'" & SVehicle & "',"  '차량구분
    sQry = sQry & "'" & sTeamCode & "'" '팀
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PH_PY030B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm.Mode = fm_ADD_MODE
        
'        Call PH_PY030_Add_MatrixRow(0, True)
        Call PH_PY030_LoadCaption
        
        GoTo PH_PY030_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PH_PY030B.Size Then
            oDS_PH_PY030B.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PH_PY030B.Offset = i
        
        Call oDS_PH_PY030B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY030B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PH_PY030B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("CLTCOD").VALUE)) '사업장
        Call oDS_PH_PY030B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("DestNo1").VALUE)) '출장번호1
        Call oDS_PH_PY030B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("DestNo2").VALUE)) '출장번호2
        Call oDS_PH_PY030B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("MSTCOD").VALUE)) '사원번호
        Call oDS_PH_PY030B.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("MSTNAM").VALUE)) '사원성명
        Call oDS_PH_PY030B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("Destinat").VALUE)) '출장지
        Call oDS_PH_PY030B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("Dest2").VALUE)) '출장지상세
        Call oDS_PH_PY030B.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("CoCode").VALUE)) '작번
        Call oDS_PH_PY030B.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("FrDate").VALUE)) '시작일자
        Call oDS_PH_PY030B.setValue("U_ColTm01", i, Trim(oRecordSet01.Fields("FrTime").VALUE)) '시작시각
        Call oDS_PH_PY030B.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("ToDate").VALUE)) '종료일자
        Call oDS_PH_PY030B.setValue("U_ColTm02", i, Trim(oRecordSet01.Fields("ToTime").VALUE)) '종료시각
        Call oDS_PH_PY030B.setValue("U_ColReg14", i, Trim(oRecordSet01.Fields("Object").VALUE)) '목적
        Call oDS_PH_PY030B.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("Comments").VALUE)) '비고
        Call oDS_PH_PY030B.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("RegCls").VALUE)) '등록구분
        Call oDS_PH_PY030B.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("ObjCls").VALUE)) '목적구분
        Call oDS_PH_PY030B.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("DestCode").VALUE)) '출장지역
        Call oDS_PH_PY030B.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("DestDiv").VALUE)) '출장구분
        Call oDS_PH_PY030B.setValue("U_ColReg20", i, Trim(oRecordSet01.Fields("Vehicle").VALUE)) '차량구분
        Call oDS_PH_PY030B.setValue("U_ColPrc01", i, Trim(oRecordSet01.Fields("FuelPrc").VALUE)) '1L단가
        Call oDS_PH_PY030B.setValue("U_ColReg22", i, Trim(oRecordSet01.Fields("FuelType").VALUE)) '유류
        Call oDS_PH_PY030B.setValue("U_ColNum01", i, Trim(oRecordSet01.Fields("Distance").VALUE)) '거리
        Call oDS_PH_PY030B.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("TransExp").VALUE)) '교통비
        Call oDS_PH_PY030B.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("DayExp").VALUE)) '일비
        Call oDS_PH_PY030B.setValue("U_ColReg23", i, Trim(oRecordSet01.Fields("FoodNum").VALUE)) '식수
        Call oDS_PH_PY030B.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("FoodExp").VALUE)) '식비
        Call oDS_PH_PY030B.setValue("U_ColSum04", i, Trim(oRecordSet01.Fields("ParkExp").VALUE)) '주차비
        Call oDS_PH_PY030B.setValue("U_ColSum05", i, Trim(oRecordSet01.Fields("TollExp").VALUE)) '도로비
        Call oDS_PH_PY030B.setValue("U_ColSum06", i, Trim(oRecordSet01.Fields("TotalExp").VALUE)) '합계
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    Call oForm.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY030_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY030_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PH_PY030_DeleteData()
'******************************************************************************
'Function ID : PH_PY030_DeleteData()
'해당모듈    : PH_PY030
'기능        : 기본정보 삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_DeleteData_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PH_PY030A] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm.Mode = fm_ADD_MODE
            GoTo PH_PY030_DeleteData_Error
        Else
            sQry = "EXEC PH_PY030_04 '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
'    Call PH_PY030_FormReset
    
'    oForm.Mode = fm_ADD_MODE
    
'    Call oForm.Items("BtnSearch").Click(ct_Regular)
    
'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PH_PY030_Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY030_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY030_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PH_PY030_UpdateData() As Boolean
'******************************************************************************
'Function ID : PH_PY030_UpdateData()
'해당모듈    : PH_PY030
'기능        : 기본정보 수정
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_UpdateData_Error

    Dim i    As Integer
    Dim j    As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry    As Integer '관리번호
    Dim CLTCOD      As String '사업장
    Dim DestNo1     As String '출장번호1
    Dim DestNo2     As String '출장번호2
    Dim MSTCOD      As String '사원번호
    Dim MSTNAM      As String '사원성명
    Dim Destinat    As String '출장지
    Dim Dest2       As String '출장지상세
    Dim CoCode      As String '작번
    Dim FrDate      As String '시작일자
    Dim FrTime      As String '시작시각
    Dim ToDate      As String '종료일자
    Dim ToTime      As String '종료시각
    Dim Object      As String '목적
    Dim Comments    As String '비고
    Dim RegCls      As String '등록구분
    Dim ObjCls      As String '목적구분
    Dim DestCode    As String '출장지역
    Dim DestDiv     As String '출장구분
    Dim Vehicle     As String '차량구분
    Dim FuelPrc     As Double '1L단가
    Dim FuelType    As String '유류
    Dim Distance    As Double '거리
    Dim TransExp    As Double '교통비
    Dim DayExp      As Double '일비
    Dim FoodNum     As String '식수
    Dim FoodExp     As Double '식비
    Dim ParkExp     As Double '주차비
    Dim TollExp     As Double '도로비
    Dim TotalExp    As Double '합계

    DocEntry = Trim(oForm.Items("DocEntry").Specific.VALUE) '관리번호
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    DestNo1 = Trim(oForm.Items("DestNo1").Specific.VALUE) '출장번호1
    DestNo2 = Trim(oForm.Items("DestNo2").Specific.VALUE) '출장번호2
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사원번호
    MSTNAM = Trim(oForm.Items("MSTNAM").Specific.VALUE) '사원성명
    Destinat = Trim(oForm.Items("Destinat").Specific.VALUE) '출장지
    Dest2 = Trim(oForm.Items("Dest2").Specific.VALUE) '출장지상세
    CoCode = Trim(oForm.Items("CoCode").Specific.VALUE) '작번
    FrDate = Trim(oForm.Items("FrDate").Specific.VALUE) '시작일자
    FrTime = Trim(oForm.Items("FrTime").Specific.VALUE) '시작시각
    ToDate = Trim(oForm.Items("ToDate").Specific.VALUE) '종료일자
    ToTime = Trim(oForm.Items("ToTime").Specific.VALUE) '종료시각
    Object = Trim(oForm.Items("Object").Specific.VALUE) '목적
    Comments = Trim(oForm.Items("Comments").Specific.VALUE) '비고
    RegCls = Trim(oForm.Items("RegCls").Specific.VALUE) '등록구분
    ObjCls = Trim(oForm.Items("ObjCls").Specific.VALUE) '목적구분
    DestCode = Trim(oForm.Items("DestCode").Specific.VALUE) '출장지역
    DestDiv = Trim(oForm.Items("DestDiv").Specific.VALUE) '출장구분
    Vehicle = Trim(oForm.Items("Vehicle").Specific.VALUE) '차량구분
    FuelPrc = Trim(oForm.Items("FuelPrc").Specific.VALUE) '1L단가
    FuelType = Trim(oForm.Items("FuelType").Specific.VALUE) '유류
    Distance = Trim(oForm.Items("Distance").Specific.VALUE) '거리
    TransExp = Trim(oForm.Items("TransExp").Specific.VALUE) '교통비
    DayExp = Trim(oForm.Items("DayExp").Specific.VALUE) '일비
    FoodNum = Trim(oForm.Items("FoodNum").Specific.VALUE) '식수
    FoodExp = Trim(oForm.Items("FoodExp").Specific.VALUE) '식비
    ParkExp = Trim(oForm.Items("ParkExp").Specific.VALUE) '주차비
    TollExp = Trim(oForm.Items("TollExp").Specific.VALUE) '도로비
    TotalExp = Trim(oForm.Items("TotalExp").Specific.VALUE) '합계
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
        PH_PY030_UpdateData = False
        Exit Function
    End If
    
    sQry = "            EXEC [PH_PY030_03] "
    sQry = sQry & "'" & DocEntry & "'," '관리번호
    sQry = sQry & "'" & CLTCOD & "',"  '사업장
    sQry = sQry & "'" & DestNo1 & "',"  '출장번호1
    sQry = sQry & "'" & DestNo2 & "',"  '출장번호2
    sQry = sQry & "'" & MSTCOD & "',"  '사원번호
    sQry = sQry & "'" & MSTNAM & "',"  '사원성명
    sQry = sQry & "'" & Destinat & "',"  '출장지
    sQry = sQry & "'" & Dest2 & "',"  '출장지상세
    sQry = sQry & "'" & CoCode & "',"  '작번
    sQry = sQry & "'" & FrDate & "',"  '시작일자
    sQry = sQry & "'" & FrTime & "',"  '시작시각
    sQry = sQry & "'" & ToDate & "',"  '종료일자
    sQry = sQry & "'" & ToTime & "',"  '종료시각
    sQry = sQry & "'" & Object & "',"  '목적
    sQry = sQry & "'" & Comments & "',"  '비고
    sQry = sQry & "'" & RegCls & "',"  '등록구분
    sQry = sQry & "'" & ObjCls & "',"  '목적구분
    sQry = sQry & "'" & DestCode & "',"  '출장지역
    sQry = sQry & "'" & DestDiv & "',"  '출장구분
    sQry = sQry & "'" & Vehicle & "',"  '차량구분
    sQry = sQry & "'" & FuelPrc & "',"  '1L단가
    sQry = sQry & "'" & FuelType & "',"  '유류
    sQry = sQry & "'" & Distance & "',"  '거리
    sQry = sQry & "'" & TransExp & "',"  '교통비
    sQry = sQry & "'" & DayExp & "',"  '일비
    sQry = sQry & "'" & FoodNum & "',"  '식수
    sQry = sQry & "'" & FoodExp & "',"  '식비
    sQry = sQry & "'" & ParkExp & "',"  '주차비
    sQry = sQry & "'" & TollExp & "',"  '도로비
    sQry = sQry & "'" & TotalExp & "'"  '합계
    
    Call RecordSet01.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PH_PY030_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY030_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY030_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PH_PY030_AddData() As Boolean
'******************************************************************************
'Function ID : PH_PY030_AddData()
'해당모듈    : PH_PY030
'기능        : 데이터 INSERT
'인수        : 없음
'반환값      : 성공여부
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_AddData_Error

    Dim i    As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry    As Integer '관리번호
    Dim CLTCOD      As String '사업장
    Dim DestNo1     As String '출장번호1
    Dim DestNo2     As String '출장번호2
    Dim MSTCOD      As String '사원번호
    Dim MSTNAM      As String '사원성명
    Dim Destinat    As String '출장지
    Dim Dest2       As String '출장지상세
    Dim CoCode      As String '작번
    Dim FrDate      As String '시작일자
    Dim FrTime      As String '시작시각
    Dim ToDate      As String '종료일자
    Dim ToTime      As String '종료시각
    Dim Object      As String '목적
    Dim Comments    As String '비고
    Dim RegCls      As String '등록구분
    Dim ObjCls      As String '목적구분
    Dim DestCode    As String '출장지역
    Dim DestDiv     As String '출장구분
    Dim Vehicle     As String '차량구분
    Dim FuelPrc     As Double '1L단가
    Dim FuelType    As String '유류
    Dim Distance    As Double '거리
    Dim TransExp    As Double '교통비
    Dim DayExp      As Double '일비
    Dim FoodNum     As String '식수
    Dim FoodExp     As Double '식비
    Dim ParkExp     As Double '주차비
    Dim TollExp     As Double '도로비
    Dim TotalExp    As Double '합계
    Dim UserSign    As String 'UserSign

    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    DestNo1 = Trim(oForm.Items("DestNo1").Specific.VALUE) '출장번호1
    DestNo2 = Trim(oForm.Items("DestNo2").Specific.VALUE) '출장번호2
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사원번호
    MSTNAM = Trim(oForm.Items("MSTNAM").Specific.VALUE) '사원성명
    Destinat = Trim(oForm.Items("Destinat").Specific.VALUE) '출장지
    Dest2 = Trim(oForm.Items("Dest2").Specific.VALUE) '출장지상세
    CoCode = Trim(oForm.Items("CoCode").Specific.VALUE) '작번
    FrDate = Trim(oForm.Items("FrDate").Specific.VALUE) '시작일자
    FrTime = Trim(oForm.Items("FrTime").Specific.VALUE) '시작시각
    ToDate = Trim(oForm.Items("ToDate").Specific.VALUE) '종료일자
    ToTime = Trim(oForm.Items("ToTime").Specific.VALUE) '종료시각
    Object = Trim(oForm.Items("Object").Specific.VALUE) '목적
    Comments = Trim(oForm.Items("Comments").Specific.VALUE) '비고
    RegCls = Trim(oForm.Items("RegCls").Specific.VALUE) '등록구분
    ObjCls = Trim(oForm.Items("ObjCls").Specific.VALUE) '목적구분
    DestCode = Trim(oForm.Items("DestCode").Specific.VALUE) '출장지역
    DestDiv = Trim(oForm.Items("DestDiv").Specific.VALUE) '출장구분
    Vehicle = Trim(oForm.Items("Vehicle").Specific.VALUE) '차량구분
    FuelPrc = Trim(oForm.Items("FuelPrc").Specific.VALUE) '1L단가
    FuelType = Trim(oForm.Items("FuelType").Specific.VALUE) '유류
    Distance = Trim(oForm.Items("Distance").Specific.VALUE) '거리
    TransExp = Trim(oForm.Items("TransExp").Specific.VALUE) '교통비
    DayExp = Trim(oForm.Items("DayExp").Specific.VALUE) '일비
    FoodNum = Trim(oForm.Items("FoodNum").Specific.VALUE) '식수
    FoodExp = Trim(oForm.Items("FoodExp").Specific.VALUE) '식비
    ParkExp = Trim(oForm.Items("ParkExp").Specific.VALUE) '주차비
    TollExp = Trim(oForm.Items("TollExp").Specific.VALUE) '도로비
    TotalExp = Trim(oForm.Items("TotalExp").Specific.VALUE) '합계
    UserSign = oCompany.UserSignature
    
    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY030A]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "            EXEC [PH_PY030_02] "
    sQry = sQry & "'" & DocEntry & "'," '관리번호
    sQry = sQry & "'" & CLTCOD & "',"  '사업장
    sQry = sQry & "'" & DestNo1 & "',"  '출장번호1
    sQry = sQry & "'" & DestNo2 & "',"  '출장번호2
    sQry = sQry & "'" & MSTCOD & "',"  '사원번호
    sQry = sQry & "'" & MSTNAM & "',"  '사원성명
    sQry = sQry & "'" & Destinat & "',"  '출장지
    sQry = sQry & "'" & Dest2 & "',"  '출장지상세
    sQry = sQry & "'" & CoCode & "',"  '작번
    sQry = sQry & "'" & FrDate & "',"  '시작일자
    sQry = sQry & "'" & FrTime & "',"  '시작시각
    sQry = sQry & "'" & ToDate & "',"  '종료일자
    sQry = sQry & "'" & ToTime & "',"  '종료시각
    sQry = sQry & "'" & Object & "',"  '목적
    sQry = sQry & "'" & Comments & "',"  '비고
    sQry = sQry & "'" & RegCls & "',"  '등록구분
    sQry = sQry & "'" & ObjCls & "',"  '목적구분
    sQry = sQry & "'" & DestCode & "',"  '출장지역
    sQry = sQry & "'" & DestDiv & "',"  '출장구분
    sQry = sQry & "'" & Vehicle & "',"  '차량구분
    sQry = sQry & "'" & FuelPrc & "',"  '1L단가
    sQry = sQry & "'" & FuelType & "',"  '유류
    sQry = sQry & "'" & Distance & "',"  '거리
    sQry = sQry & "'" & TransExp & "',"  '교통비
    sQry = sQry & "'" & DayExp & "',"  '일비
    sQry = sQry & "'" & FoodNum & "',"  '식수
    sQry = sQry & "'" & FoodExp & "',"  '식비
    sQry = sQry & "'" & ParkExp & "',"  '주차비
    sQry = sQry & "'" & TollExp & "',"  '도로비
    sQry = sQry & "'" & TotalExp & "',"  '합계
    sQry = sQry & "'" & UserSign & "'"  'UserSign
    
    Call RecordSet02.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PH_PY030_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY030_AddData_Error:
    PH_PY030_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY030_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PH_PY030_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PH_PY030_HeaderSpaceLineDel()
'해당모듈    : PH_PY030
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm.Items("DestNo1").Specific.VALUE) = "" '출장번호1
            ErrNum = 1
            GoTo PH_PY030_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("DestNo2").Specific.VALUE) = "" '출장번호2
            ErrNum = 2
            GoTo PH_PY030_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("MSTCOD").Specific.VALUE) = "" '사원번호
            ErrNum = 3
            GoTo PH_PY030_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrDate").Specific.VALUE) = "" '시작일자
            ErrNum = 4
            GoTo PH_PY030_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrTime").Specific.VALUE) = "" '시작시각
            ErrNum = 5
            GoTo PH_PY030_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToDate").Specific.VALUE) = "" '종료일자
            ErrNum = 6
            GoTo PH_PY030_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToTime").Specific.VALUE) = "" '종료시각
            ErrNum = 7
            GoTo PH_PY030_HeaderSpaceLineDel_Error
    End Select

    PH_PY030_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY030_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출장번호1은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("DestNo1").CLICK(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "출장번호2는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("DestNo2").CLICK(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "사원번호는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("MSTCOD").CLICK(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "시작일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("FrDate").CLICK(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "시작시각은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("FrTime").CLICK(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "종료일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("FrDate").CLICK(ct_Regular)
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "종료시각은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("FrTime").CLICK(ct_Regular)
    End If
    PH_PY030_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PH_PY030_MatrixSpaceLineDel() As Boolean
On Error GoTo PH_PY030_MatrixSpaceLineDel_Error

    Dim i      As Long
    Dim ErrNum As Integer
    Dim sQry   As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Nothing
    PH_PY030_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY030_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PH_PY030_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PH_PY030_MatrixSpaceLineDel = False
End Function

Private Sub PH_PY030_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY030_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim ErrNum    As Integer
    Dim sQry      As String
    Dim ItemCode  As String
    Dim oCombo    As SAPbouiCOM.ComboBox
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case oUID

        Case "FuelPrc"
        
            Call PH_PY030_CalculateTransExp
            Call PH_PY030_CalculateTotalExp
            
        Case "Distance"
        
            Call PH_PY030_CalculateTransExp
            Call PH_PY030_CalculateTotalExp
            
        Case "TransExp"
        
            Call PH_PY030_CalculateTotalExp
            
        Case "DayExp"
        
            Call PH_PY030_CalculateTotalExp
            
        Case "FoodExp"
        
            Call PH_PY030_CalculateTotalExp
        
        Case "ParkExp"
        
            Call PH_PY030_CalculateTotalExp
            
        Case "TollExp"
        
            Call PH_PY030_CalculateTotalExp
        
        Case "FrDate"
        
            Call PH_PY030_GetDestNo
        
        Case "FuelType"
        
            Call PH_PY030_GetFuelPrc
'            Call PH_PY030_CalculateTransExp
'            Call PH_PY030_CalculateTotalExp
        
        Case "CLTCOD"
            
            Call PH_PY030_GetDestNo '출장번호생성
        
        Case "FoodNum"
        
            Call PH_PY030_CalculateFoodExp '식비 계산
            Call PH_PY030_CalculateTotalExp
        
        Case "SCLTCOD"
        
            Set oCombo = oForm.Items("STeamCode").Specific
        
            If oCombo.ValidValues.Count > 0 Then
                For loopCount = oCombo.ValidValues.Count - 1 To 0 Step -1
                    Call oCombo.ValidValues.Remove(loopCount, psk_Index)
                Next
            End If
            
            Call oCombo.ValidValues.Add("%", "전체")
            sQry = "        SELECT      U_Code,"
            sQry = sQry & "             U_CodeNm"
            sQry = sQry & " FROM        [@PS_HR200L]"
            sQry = sQry & " WHERE       Code = '1'"
            sQry = sQry & "             AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
            sQry = sQry & "             AND U_UseYN = 'Y'"
            sQry = sQry & " ORDER BY    U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("STeamCode").Specific, sQry)
            Call oCombo.Select(0, psk_Index)
            
            oForm.Items("STeamCode").DisplayDesc = True
        Case "Destinat"
            If Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) = "1" Then
                sQry = "Select Distance = U_Num2 From [@PS_HR200L] Where Code = 'P217' And U_Code = '" & oForm.Items.Item("Destinat").Specific.VALUE & "' and U_Char1 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                Call oRecordSet01.DoQuery(sQry)
                oForm.Items.Item("Distance").Specific.VALUE = oRecordSet01.Fields("Distance").VALUE
            End If
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
    
PH_PY030_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PH_PY030_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PH_PY030_FormItemEnabled()
On Error GoTo PH_PY030_FormItemEnabled_Error

    If (oForm.Mode = fm_ADD_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_FIND_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_OK_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        Call CLTCOD_Select(oForm, "SCLTCOD")
        
    End If
    
    Exit Sub

PH_PY030_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY030_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/아이템 변경 이벤트
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error

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

Raise_FormItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormMenuEvent_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
                Call PH_PY030_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PH_PY030_LoadCaption
                
                'oForm.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm.Freeze(True)
                Call PH_PY030_Add_MatrixRow(oMat01.VisualRowCount)
                Call oForm.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PH_PY030_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PH_PY030A.Clear

'                Call PH_PY030_LoadCaption
'                Call PH_PY030_FormItemEnabled
                '//Call PH_PY030_FormItemEnabled '//UDO방식
                '//Call PH_PY030_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY030_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm.Freeze(True)
                Call oDS_PH_PY030B.RemoveRecord(oDS_PH_PY030B.Size - 1)
                Call oMat01.LoadFromDataSource
                Call oForm.Freeze(False)
            
        End Select
    End If
    Exit Sub
Raise_FormMenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    
        If pval.ItemUID = "PH_PY030" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm.Mode = fm_ADD_MODE Then
            
                If PH_PY030_HeaderSpaceLineDel() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                If PH_PY030_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                If PH_PY030_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PH_PY030_FormReset
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY030_LoadCaption
                Call PH_PY030_MTX01
                
                oLast_Mode = oForm.Mode
                
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            
                If PH_PY030_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PH_PY030_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                If PH_PY030_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PH_PY030_FormReset
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY030_LoadCaption
                Call PH_PY030_MTX01
                
'                oForm.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
            Call PH_PY030_FormReset
            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PH_PY030_LoadCaption
            Call PH_PY030_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PH_PY030_DeleteData
                Call PH_PY030_FormReset
                oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PH_PY030_LoadCaption
                Call PH_PY030_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            Call PH_PY030_Print_Report01
        
        ElseIf pval.ItemUID = "BtnPrint2" Then
        
            Call PH_PY030_Print_Report02
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PH_PY030" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD", "") '기본정보-사번
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "SMSTCOD", "") '조회조건-사번
        
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
                
                Call oForm.Freeze(True)

                'DataSource를 이용하여 각 컨트롤에 값을 출력
                Call oDS_PH_PY030A.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
                Call oDS_PH_PY030A.setValue("U_CLTCOD", 0, oMat01.Columns("CLTCOD").Cells(pval.Row).Specific.VALUE) '사업장
                Call oDS_PH_PY030A.setValue("U_DestNo1", 0, oMat01.Columns("DestNo1").Cells(pval.Row).Specific.VALUE) '출장번호1
                Call oDS_PH_PY030A.setValue("U_DestNo2", 0, oMat01.Columns("DestNo2").Cells(pval.Row).Specific.VALUE) '출장번호2
                Call oDS_PH_PY030A.setValue("U_MSTCOD", 0, oMat01.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE) '사원번호
                Call oDS_PH_PY030A.setValue("U_MSTNAM", 0, oMat01.Columns("MSTNAM").Cells(pval.Row).Specific.VALUE) '사원성명
                Call oDS_PH_PY030A.setValue("U_Destinat", 0, oMat01.Columns("Destinat").Cells(pval.Row).Specific.VALUE) '출장지
                Call oDS_PH_PY030A.setValue("U_Dest2", 0, oMat01.Columns("Dest2").Cells(pval.Row).Specific.VALUE) '출장지상세
                Call oDS_PH_PY030A.setValue("U_CoCode", 0, oMat01.Columns("CoCode").Cells(pval.Row).Specific.VALUE) '작번
                Call oDS_PH_PY030A.setValue("U_FrDate", 0, Replace(oMat01.Columns("FrDate").Cells(pval.Row).Specific.VALUE, ".", "")) '시작일자
                Call oDS_PH_PY030A.setValue("U_FrTime", 0, oMat01.Columns("FrTime").Cells(pval.Row).Specific.VALUE) '시작시각
                Call oDS_PH_PY030A.setValue("U_ToDate", 0, Replace(oMat01.Columns("ToDate").Cells(pval.Row).Specific.VALUE, ".", "")) '종료일자
                Call oDS_PH_PY030A.setValue("U_ToTime", 0, oMat01.Columns("ToTime").Cells(pval.Row).Specific.VALUE) '종료시각
                Call oDS_PH_PY030A.setValue("U_Object", 0, oMat01.Columns("Object").Cells(pval.Row).Specific.VALUE) '목적
                Call oDS_PH_PY030A.setValue("U_Comments", 0, oMat01.Columns("Comments").Cells(pval.Row).Specific.VALUE) '비고
                Call oDS_PH_PY030A.setValue("U_RegCls", 0, oMat01.Columns("RegCls").Cells(pval.Row).Specific.VALUE) '등록구분
                Call oDS_PH_PY030A.setValue("U_ObjCls", 0, oMat01.Columns("ObjCls").Cells(pval.Row).Specific.VALUE) '목적구분
                Call oDS_PH_PY030A.setValue("U_DestCode", 0, oMat01.Columns("DestCode").Cells(pval.Row).Specific.VALUE) '출장지역
                Call oDS_PH_PY030A.setValue("U_DestDiv", 0, oMat01.Columns("DestDiv").Cells(pval.Row).Specific.VALUE) '출장구분
                Call oDS_PH_PY030A.setValue("U_Vehicle", 0, oMat01.Columns("Vehicle").Cells(pval.Row).Specific.VALUE) '차량구분
                Call oDS_PH_PY030A.setValue("U_FuelPrc", 0, oMat01.Columns("FuelPrc").Cells(pval.Row).Specific.VALUE) '1L단가
                Call oDS_PH_PY030A.setValue("U_FuelType", 0, oMat01.Columns("FuelType").Cells(pval.Row).Specific.VALUE) '유류
                Call oDS_PH_PY030A.setValue("U_Distance", 0, oMat01.Columns("Distance").Cells(pval.Row).Specific.VALUE) '거리
                Call oDS_PH_PY030A.setValue("U_TransExp", 0, oMat01.Columns("TransExp").Cells(pval.Row).Specific.VALUE) '교통비
                Call oDS_PH_PY030A.setValue("U_DayExp", 0, oMat01.Columns("DayExp").Cells(pval.Row).Specific.VALUE) '일비
                Call oDS_PH_PY030A.setValue("U_FoodNum", 0, oMat01.Columns("FoodNum").Cells(pval.Row).Specific.VALUE) '식수
                Call oDS_PH_PY030A.setValue("U_FoodExp", 0, oMat01.Columns("FoodExp").Cells(pval.Row).Specific.VALUE) '식비
                Call oDS_PH_PY030A.setValue("U_ParkExp", 0, oMat01.Columns("ParkExp").Cells(pval.Row).Specific.VALUE) '주차비
                Call oDS_PH_PY030A.setValue("U_TollExp", 0, oMat01.Columns("TollExp").Cells(pval.Row).Specific.VALUE) '도로비
                Call oDS_PH_PY030A.setValue("U_TotalExp", 0, oMat01.Columns("TotalExp").Cells(pval.Row).Specific.VALUE) '합계
                
                oForm.Mode = fm_UPDATE_MODE
                Call PH_PY030_LoadCaption
                
                Call oForm.Freeze(False)
                
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PH_PY030_FlushToItemValue(pval.ItemUID)
        
    End If

    Exit Sub
    
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm.Freeze(False)
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

    Call oForm.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PH_PY030B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PH_PY030B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PH_PY030_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PH_PY030B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PH_PY030_FlushToItemValue(pval.ItemUID)
                
                If pval.ItemUID = "MSTCOD" Then
                
                    oForm.Items("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.VALUE & "'") '성명
                    
                ElseIf pval.ItemUID = "SMSTCOD" Then
                
                    oForm.Items("SMSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & oForm.Items("SMSTCOD").Specific.VALUE & "'") '성명
                    
                End If
                
            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm.Update
        End If
        
    ElseIf pval.BeforeAction = False Then

    End If
    
    Call oForm.Freeze(False)
    
    Exit Sub
    
Raise_EVENT_VALIDATE_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PH_PY030_FormItemEnabled
        '//Call PH_PY030_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
    
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error

    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY030_FormResize
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
'            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY030A", "U_CardCode,U_CardName")
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
        Set oForm = Nothing
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
'            If (PH_PY030_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PH_PY030A.RemoveRecord(oDS_PH_PY030A.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PH_PY030_Add_MatrixRow(0)
            Else
                If Trim(oDS_PH_PY030A.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PH_PY030_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
    
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY030_CreateItems() As Boolean
On Error GoTo PH_PY030_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY030A = oForm.DataSources.DBDataSources("@PH_PY030A")
    Set oDS_PH_PY030B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '관리번호
    Call oForm.DataSources.UserDataSources.Add("SDocEntry", dt_SHORT_TEXT, 100)
    Call oForm.Items("SDocEntry").Specific.DataBind.SetBound(True, "", "SDocEntry")
    
    '사업장
    Call oForm.DataSources.UserDataSources.Add("SCLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("SCLTCOD").Specific.DataBind.SetBound(True, "", "SCLTCOD")
    
    '출장번호1
    Call oForm.DataSources.UserDataSources.Add("SDestNo1", dt_SHORT_TEXT, 10)
    Call oForm.Items("SDestNo1").Specific.DataBind.SetBound(True, "", "SDestNo1")
    
    '출장번호2
    Call oForm.DataSources.UserDataSources.Add("SDestNo2", dt_SHORT_TEXT, 10)
    Call oForm.Items("SDestNo2").Specific.DataBind.SetBound(True, "", "SDestNo2")
    
    '사원번호
    Call oForm.DataSources.UserDataSources.Add("SMSTCOD", dt_SHORT_TEXT, 20)
    Call oForm.Items("SMSTCOD").Specific.DataBind.SetBound(True, "", "SMSTCOD")
    
    '사원성명
    Call oForm.DataSources.UserDataSources.Add("SMSTNAM", dt_SHORT_TEXT, 50)
    Call oForm.Items("SMSTNAM").Specific.DataBind.SetBound(True, "", "SMSTNAM")
    
    '출장지
    Call oForm.DataSources.UserDataSources.Add("SDestinat", dt_SHORT_TEXT, 10)
    Call oForm.Items("SDestinat").Specific.DataBind.SetBound(True, "", "SDestinat")
    
    '출장지상세
    Call oForm.DataSources.UserDataSources.Add("SDest2", dt_SHORT_TEXT, 100)
    Call oForm.Items("SDest2").Specific.DataBind.SetBound(True, "", "SDest2")
    
    '작번
    Call oForm.DataSources.UserDataSources.Add("SCoCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("SCoCode").Specific.DataBind.SetBound(True, "", "SCoCode")
    
    '시작월
    Call oForm.DataSources.UserDataSources.Add("SFrDate", dt_DATE)
    Call oForm.Items("SFrDate").Specific.DataBind.SetBound(True, "", "SFrDate")
    
    '종료월
    Call oForm.DataSources.UserDataSources.Add("SToDate", dt_DATE)
    Call oForm.Items("SToDate").Specific.DataBind.SetBound(True, "", "SToDate")
    
    '목적
    Call oForm.DataSources.UserDataSources.Add("SObject", dt_LONG_TEXT, 100)
    Call oForm.Items("SObject").Specific.DataBind.SetBound(True, "", "SObject")
    
    '비고
    Call oForm.DataSources.UserDataSources.Add("SComments", dt_LONG_TEXT, 100)
    Call oForm.Items("SComments").Specific.DataBind.SetBound(True, "", "SComments")
    
    '등록구분
    Call oForm.DataSources.UserDataSources.Add("SRegCls", dt_SHORT_TEXT, 10)
    Call oForm.Items("SRegCls").Specific.DataBind.SetBound(True, "", "SRegCls")
    
    '목적구분
    Call oForm.DataSources.UserDataSources.Add("SObjCls", dt_SHORT_TEXT, 10)
    Call oForm.Items("SObjCls").Specific.DataBind.SetBound(True, "", "SObjCls")
    
    '출장지역
    Call oForm.DataSources.UserDataSources.Add("SDestCode", dt_SHORT_TEXT, 10)
    Call oForm.Items("SDestCode").Specific.DataBind.SetBound(True, "", "SDestCode")
    
    '출장구분
    Call oForm.DataSources.UserDataSources.Add("SDestDiv", dt_SHORT_TEXT, 10)
    Call oForm.Items("SDestDiv").Specific.DataBind.SetBound(True, "", "SDestDiv")
    
    '차량구분
    Call oForm.DataSources.UserDataSources.Add("SVehicle", dt_SHORT_TEXT, 10)
    Call oForm.Items("SVehicle").Specific.DataBind.SetBound(True, "", "SVehicle")
    
    '팀
    Call oForm.DataSources.UserDataSources.Add("STeamCode", dt_SHORT_TEXT, 10)
    Call oForm.Items("STeamCode").Specific.DataBind.SetBound(True, "", "STeamCode")
    
'    '사업장_S
'    Call oForm.DataSources.UserDataSources.Add("SBPLId", dt_SHORT_TEXT, 10)
'    Call oForm.Items("SBPLId").Specific.DataBind.SetBound(True, "", "SBPLId")
'    '사업장_E
'
'    '사번_S
'    Call oForm.DataSources.UserDataSources.Add("SCntcCode", dt_SHORT_TEXT, 20)
'    Call oForm.Items("SCntcCode").Specific.DataBind.SetBound(True, "", "SCntcCode")
'    '사번_E
'
'    '성명_S
'    Call oForm.DataSources.UserDataSources.Add("SCntcName", dt_SHORT_TEXT, 50)
'    Call oForm.Items("SCntcName").Specific.DataBind.SetBound(True, "", "SCntcName")
'    '성명_E
'
'    '지급일자(시작)_S
'    Call oForm.DataSources.UserDataSources.Add("SPrvdDtFr", dt_DATE)
'    Call oForm.Items("SPrvdDtFr").Specific.DataBind.SetBound(True, "", "SPrvdDtFr")
'    '지급일자(시작)_E
'
'    '지급일자(종료)_S
'    Call oForm.DataSources.UserDataSources.Add("SPrvdDtTo", dt_DATE)
'    Call oForm.Items("SPrvdDtTo").Specific.DataBind.SetBound(True, "", "SPrvdDtTo")
'    '지급일자(종료)_E
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
    
PH_PY030_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY030_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PH_PY030_ComboBox_Setting()
On Error GoTo PH_PY030_ComboBox_Setting_Error

    Dim oCombo  As SAPbouiCOM.ComboBox
    Dim sQry    As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)

    '//////////기본정보//////////
    '출장지
    Call oForm.Items("Destinat").Specific.ValidValues.Add("%", "선택")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P217'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("Destinat").Specific, sQry, "", False, False)
    Call oForm.Items("Destinat").Specific.Select(0, psk_Index)

    '등록구분
    Call oForm.Items("RegCls").Specific.ValidValues.Add("%", "선택")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P223'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("RegCls").Specific, sQry, "", False, False)
    Call oForm.Items("RegCls").Specific.Select(0, psk_Index)
    
    '목적구분
    Call oForm.Items("ObjCls").Specific.ValidValues.Add("%", "선택")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P224'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("ObjCls").Specific, sQry, "", False, False)
    Call oForm.Items("ObjCls").Specific.Select(0, psk_Index)
    
    '출장지역
    Call oForm.Items("DestCode").Specific.ValidValues.Add("%", "선택")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P225'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("DestCode").Specific, sQry, "", False, False)
    Call oForm.Items("DestCode").Specific.Select(0, psk_Index)
    
    '출장구분
    Call oForm.Items("DestDiv").Specific.ValidValues.Add("%", "선택")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P216'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("DestDiv").Specific, sQry, "", False, False)
    Call oForm.Items("DestDiv").Specific.Select(0, psk_Index)
    
    '차량구분
    Call oForm.Items("Vehicle").Specific.ValidValues.Add("%", "선택")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P218'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("Vehicle").Specific, sQry, "", False, False)
    Call oForm.Items("Vehicle").Specific.Select(0, psk_Index)
    
    '유류
    Call oForm.Items("FuelType").Specific.ValidValues.Add("%", "선택")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P226'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("FuelType").Specific, sQry, "", False, False)
    Call oForm.Items("FuelType").Specific.Select(0, psk_Index)
    
    '식수
'    Call oForm.Items("FoodNum").Specific.ValidValues.Add("0", "선택")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P227'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("FoodNum").Specific, sQry, "", False, False)
    Call oForm.Items("FoodNum").Specific.Select(0, psk_Index)
    
    '//////////조회정보//////////
    '출장지
    Call oForm.Items("SDestinat").Specific.ValidValues.Add("%", "전체")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P217'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("SDestinat").Specific, sQry, "", False, False)
    Call oForm.Items("SDestinat").Specific.Select(0, psk_Index)

    '등록구분
    Call oForm.Items("SRegCls").Specific.ValidValues.Add("%", "전체")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P223'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("SRegCls").Specific, sQry, "", False, False)
    Call oForm.Items("SRegCls").Specific.Select(0, psk_Index)
    
    '목적구분
    Call oForm.Items("SObjCls").Specific.ValidValues.Add("%", "전체")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P224'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("SObjCls").Specific, sQry, "", False, False)
    Call oForm.Items("SObjCls").Specific.Select(0, psk_Index)
    
    '출장지역
    Call oForm.Items("SDestCode").Specific.ValidValues.Add("%", "전체")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P225'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("SDestCode").Specific, sQry, "", False, False)
    Call oForm.Items("SDestCode").Specific.Select(0, psk_Index)
    
    '출장구분
    Call oForm.Items("SDestDiv").Specific.ValidValues.Add("%", "전체")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P216'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("SDestDiv").Specific, sQry, "", False, False)
    Call oForm.Items("SDestDiv").Specific.Select(0, psk_Index)
    
    '차량구분
    Call oForm.Items("SVehicle").Specific.ValidValues.Add("%", "전체")
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P218'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("SVehicle").Specific, sQry, "", False, False)
    Call oForm.Items("SVehicle").Specific.Select(0, psk_Index)
    
    '//////////매트릭스//////////
    '사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    '출장지
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P217'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Destinat"), sQry)
    
    '등록구분
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P223'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("RegCls"), sQry)
    
    '목적구분
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P224'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ObjCls"), sQry)
    
    '출장지역
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P225'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("DestCode"), sQry)
    
    '출장구분
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P216'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("DestDiv"), sQry)
    
    '차량구분
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P218'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Vehicle"), sQry)
    
    '유류
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P226'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("FuelType"), sQry)
    
    '식수
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm As [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P227'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("FoodNum"), sQry)
    
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY030_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY030_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY030_CF_ChooseFromList()
On Error GoTo PH_PY030_CF_ChooseFromList_Error

    Exit Sub
PH_PY030_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY030_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY030_EnableMenus()
On Error GoTo PH_PY030_EnableMenus_Error
    
    Exit Sub
PH_PY030_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY030_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY030_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY030_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PH_PY030_FormItemEnabled
        '//Call PH_PY030_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY030_FormItemEnabled
'        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm.Items("1").Click ct_Regular
    End If
    Exit Sub
PH_PY030_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY030_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY030_FormResize()
On Error GoTo PH_PY030_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PH_PY030_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY030_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PH_PY030_FormReset()
'******************************************************************************
'Function ID : PH_PY030_FormReset()
'해당모듈    : PH_PY030
'기능        : 화면 초기화
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY030A]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PH_PY030A.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PH_PY030A.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    Dim User_BPLID As String
    User_BPLID = MDC_PS_Common.User_BPLID
    
    '//////////기준정보//////////
    Call oDS_PH_PY030A.setValue("U_CLTCOD", 0, User_BPLID) '사업장
    Call oDS_PH_PY030A.setValue("U_DestNo1", 0, "") '출장번호1
    Call oDS_PH_PY030A.setValue("U_DestNo2", 0, "") '출장번호2
    Call oDS_PH_PY030A.setValue("U_MSTCOD", 0, "") '사원번호
    Call oDS_PH_PY030A.setValue("U_MSTNAM", 0, "") '사원성명
    Call oDS_PH_PY030A.setValue("U_Destinat", 0, "%") '출장지
    Call oDS_PH_PY030A.setValue("U_Dest2", 0, "") '출장지상세
    Call oDS_PH_PY030A.setValue("U_CoCode", 0, "") '작번
    Call oDS_PH_PY030A.setValue("U_FrDate", 0, Format(Now, "yyyyMMdd")) '시작일자
    Call oDS_PH_PY030A.setValue("U_FrTime", 0, "") '시작시각
    Call oDS_PH_PY030A.setValue("U_ToDate", 0, Format(Now, "yyyyMMdd")) '종료일자
    Call oDS_PH_PY030A.setValue("U_ToTime", 0, "") '종료시각
    Call oDS_PH_PY030A.setValue("U_Object", 0, "") '목적
    Call oDS_PH_PY030A.setValue("U_Comments", 0, "") '비고
    Call oDS_PH_PY030A.setValue("U_RegCls", 0, "01") '등록구분
    Call oDS_PH_PY030A.setValue("U_ObjCls", 0, "%") '목적구분
    Call oDS_PH_PY030A.setValue("U_DestCode", 0, "01") '출장지역
    Call oDS_PH_PY030A.setValue("U_DestDiv", 0, "%") '출장구분
    Call oDS_PH_PY030A.setValue("U_Vehicle", 0, "%") '차량구분
    Call oDS_PH_PY030A.setValue("U_FuelPrc", 0, 0) '1L단가
    Call oDS_PH_PY030A.setValue("U_FuelType", 0, "%") '유류
    Call oDS_PH_PY030A.setValue("U_Distance", 0, 0) '거리
    Call oDS_PH_PY030A.setValue("U_TransExp", 0, 0) '교통비
    Call oDS_PH_PY030A.setValue("U_DayExp", 0, 0) '일비
    Call oDS_PH_PY030A.setValue("U_FoodNum", 0, "0") '식수
    Call oDS_PH_PY030A.setValue("U_FoodExp", 0, 0) '식비
    Call oDS_PH_PY030A.setValue("U_ParkExp", 0, 0) '주차비
    Call oDS_PH_PY030A.setValue("U_TollExp", 0, 0) '도로비
    Call oDS_PH_PY030A.setValue("U_TotalExp", 0, 0) '합계
    '출장번호
    Call PH_PY030_GetDestNo
    
    '//////////조회정보//////////
'    Call oForm.Items("SCLTCOD").Specific.Select(User_BPLID, psk_ByValue) '사업장
'    Call oForm.Items("SDestinat").Specific.Select(0, psk_Index) '출장지
'    Call oForm.Items("SRegCls").Specific.Select(0, psk_Index) '등록구분
'    Call oForm.Items("SObjCls").Specific.Select(0, psk_Index) '목적구분
'    Call oForm.Items("SDestCode").Specific.Select(0, psk_Index) '출장지역
'    Call oForm.Items("SDestDiv").Specific.Select(0, psk_Index) '출장구분
'    Call oForm.Items("SVehicle").Specific.Select(0, psk_Index) '차량구분
'    '기간(월)
'    oForm.Items("SFrDate").Specific.VALUE = Format(Now, "YYYY.MM")
'    oForm.Items("SToDate").Specific.VALUE = Format(Now, "YYYY.MM")
    
    Call oForm.Items("MSTCOD").CLICK
    
    Set RecordSet01 = Nothing
    Call oForm.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY030_FormReset_Error:
    Call oForm.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY030_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PH_PY030_CalculateTransExp()
'******************************************************************************
'Function ID : PH_PY030_CalculateTransExp()
'해당모듈    : PH_PY030
'기능        : 교통비 계산
'인수        : 없음
'반환값      : 합계 금액
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_CalculateTransExp_Error

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Dim ErrNum As Integer

    Dim FuelPrc As Double '유류단가
    Dim Distance As Double '거리
    Dim TransExp As Double '교통비
    
    FuelPrc = CDbl(oForm.Items("FuelPrc").Specific.VALUE)
    Distance = CDbl(oForm.Items("Distance").Specific.VALUE)

    TransExp = ((FuelPrc * Distance * 0.1) \ 10) * 10 '원단위 절사
    
    oForm.Items("TransExp").Specific.VALUE = TransExp

    Set oRecordSet01 = Nothing

    Exit Sub

PH_PY030_CalculateTransExp_Error:
    If ErrNum = 1 Then
    Else
        MDC_Com.MDC_GF_Message "PH_PY030_CalculateTransExp_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    Set oRecordSet01 = Nothing
End Sub

Private Sub PH_PY030_CalculateTotalExp()
'******************************************************************************
'Function ID : PH_PY030_CalculateTotalExp()
'해당모듈    : PH_PY030
'기능        : 합계 금액 계산
'인수        : 없음
'반환값      : 합계 금액
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_CalculateTotalExp_Error

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Dim ErrNum   As Integer
    Dim TransExp As Double '교통비
    Dim DayExp   As Double '일비
    Dim FoodExp  As Double '식비
    Dim ParkExp  As Double '주차비
    Dim TollExp  As Double '도로비
    Dim TotalExp As Double '합계
    
    TransExp = Trim(oForm.Items("TransExp").Specific.VALUE)
    DayExp = Trim(oForm.Items("DayExp").Specific.VALUE)
    FoodExp = Trim(oForm.Items("FoodExp").Specific.VALUE)
    ParkExp = Trim(oForm.Items("ParkExp").Specific.VALUE)
    TollExp = Trim(oForm.Items("TollExp").Specific.VALUE)
    TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp
    
    Call oDS_PH_PY030A.setValue("U_TotalExp", 0, TotalExp)

    Set oRecordSet01 = Nothing

    Exit Sub

PH_PY030_CalculateTotalExp_Error:
    If ErrNum = 1 Then
    Else
        MDC_Com.MDC_GF_Message "PH_PY030_CalculateTotalExp_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    Set oRecordSet01 = Nothing
    
End Sub

Private Sub PH_PY030_GetDestNo()
'******************************************************************************
'Function ID : PH_PY030_GetDestNo()
'해당모듈    : PH_PY030
'기능        : 출장번호 생성
'인수        : 없음
'반환값      : 합계 금액
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_GetDestNo_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Dim ErrNum As Integer
    Dim FrDate As String
    Dim CLTCOD As String
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    FrDate = Left(Trim(oForm.Items("FrDate").Specific.VALUE), 6)
    
    sQry = "      EXEC PH_PY030_05 '"
    sQry = sQry & CLTCOD & "','"
    sQry = sQry & FrDate & "'"
    Call oRecordSet01.DoQuery(sQry)

    oForm.Items("DestNo1").Specific.VALUE = FrDate
    oForm.Items("DestNo2").Specific.VALUE = Trim(oRecordSet01.Fields("DestNo2").VALUE)
    
    Set oRecordSet01 = Nothing

    Exit Sub

PH_PY030_GetDestNo_Error:
    If ErrNum = 1 Then
    Else
        MDC_Com.MDC_GF_Message "PH_PY030_GetDestNo_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    Set oRecordSet01 = Nothing
    
End Sub

Private Sub PH_PY030_GetFuelPrc()
'******************************************************************************
'Function ID : PH_PY030_GetFuelPrc()
'해당모듈    : PH_PY030
'기능        : 유류단가 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_GetFuelPrc_Error

    Dim loopCount As Integer
    Dim sQry      As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD    As String
    Dim StdYear   As String
    Dim StdMonth  As String
    Dim FuelType  As String
    Dim FuelPrice As Double
    
    CLTCOD = Trim(oDS_PH_PY030A.GetValue("U_CLTCOD", 0)) '사업장
    StdYear = Left(Trim(oDS_PH_PY030A.GetValue("U_FrDate", 0)), 4)
    StdMonth = Mid(Trim(oDS_PH_PY030A.GetValue("U_FrDate", 0)), 5, 2)
    FuelType = Trim(oDS_PH_PY030A.GetValue("U_FuelType", 0)) '유류
    
    sQry = "        SELECT      T0.U_Year AS [StdYear],"
    sQry = sQry & "             T1.U_Month AS [StdMonth],"
    sQry = sQry & "             T1.U_Gasoline AS [Gasoline],"
    sQry = sQry & "             T1.U_Diesel AS [Diesel],"
    sQry = sQry & "             T1.U_LPG AS [LPG]"
    sQry = sQry & " FROM        [@PH_PY007A] AS T0"
    sQry = sQry & "             INNER JOIN"
    sQry = sQry & "             [@PH_PY007B] AS T1"
    sQry = sQry & "                 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE       T0.U_CLTCOD = '" & CLTCOD & "'"
    sQry = sQry & "             AND T0.U_Year = '" & StdYear & "'"
    sQry = sQry & "             AND T1.U_Month = '" & StdMonth & "'"
    
    Call oRecordSet.DoQuery(sQry)
    
    If FuelType = "1" Then '휘발유
        FuelPrice = oRecordSet.Fields("Gasoline").VALUE
    ElseIf FuelType = "2" Then '가스
        FuelPrice = oRecordSet.Fields("LPG").VALUE
    ElseIf FuelType = "3" Then '경유
        FuelPrice = oRecordSet.Fields("Diesel").VALUE
    Else
        FuelPrice = 0
    End If

'    Call oDS_PH_PY030A.setValue("U_FuelPrc", 0, FuelPrice)
    
    oForm.Items("FuelPrc").Specific.VALUE = FuelPrice
    Call oForm.Items("Distance").CLICK
    
    Exit Sub
    
PH_PY030_GetFuelPrc_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY030_GetFuelPrc_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY030_CalculateFoodExp()
'******************************************************************************
'Function ID : PH_PY030_CalculateFoodExp()
'해당모듈    : PH_PY030
'기능        : 식비 계산
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY030_CalculateFoodExp_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim MSTCOD  As String '사번
    Dim FoodNum As Integer '식수
    Dim FoodPrc As Double '당일식비
    Dim FoodExp As Double '전체식비
    
    If Trim(oDS_PH_PY030A.GetValue("U_MSTCOD", 0)) = "" And Trim(oDS_PH_PY030A.GetValue("U_FoodNum", 0)) <> "0" Then '사번을 선택하지 않으면
        ErrNum = 1
        GoTo PH_PY030_CalculateFoodExp_Error
    End If
    
    MSTCOD = Trim(oDS_PH_PY030A.GetValue("U_MSTCOD", 0)) '사번
    FoodNum = Trim(oDS_PH_PY030A.GetValue("U_FoodNum", 0)) '식수
    
    sQry = "        SELECT      T1.U_Num4 AS [FoodPrc]"
    sQry = sQry & " FROM        [@PH_PY001A] AS T0"
    sQry = sQry & "             LEFT JOIN"
    sQry = sQry & "             [@PS_HR200L] AS T1"
    sQry = sQry & "                 ON T0.U_JIGCOD = T1.U_Code"
    sQry = sQry & "                 AND T1.Code = 'P232'"
    sQry = sQry & "                 AND T1.U_UseYN = 'Y'"
    sQry = sQry & " WHERE       T0.Code = '" & MSTCOD & "'"

    Call oRecordSet.DoQuery(sQry)
    
    FoodPrc = oRecordSet.Fields("FoodPrc").VALUE
    FoodExp = FoodPrc * FoodNum
    
    Call oDS_PH_PY030A.setValue("U_FoodExp", 0, FoodExp)
    Call oForm.Items("FoodExp").CLICK

    Exit Sub
    
PH_PY030_CalculateFoodExp_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.SetStatusBarMessage "사원을 먼저 선택하세요.", bmt_Short, True
        Call oDS_PH_PY030A.setValue("U_FoodNum", 0, "0") '0식 선택
        Call oForm.Items("MSTCOD").CLICK
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY030_CalculateFoodExp_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub


Private Sub PH_PY030_Print_Report01()
On Error GoTo PH_PY030_Print_Report01_Error

    Dim DocNum     As String
    Dim ErrNum     As Integer
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String
    Dim CLTCOD     As String
    Dim DestNo1    As String
    Dim DestNo2    As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY030_Print_Report01_Error
    End If
    
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DestNo1 = Trim(oForm.Items("DestNo1").Specific.VALUE)
    DestNo2 = Trim(oForm.Items("DestNo2").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY030] 공용증"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PH_PY030_01.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PH_PY030_02.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PH_PY030_03.rpt"
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
    sQry = "      EXEC [PH_PY030_90] '"
    sQry = sQry & CLTCOD & "','"
    sQry = sQry & DestNo1 & "','"
    sQry = sQry & DestNo2 & "'"
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY030_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY030_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY030_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub

Private Sub PH_PY030_Print_Report02()
On Error GoTo PH_PY030_Print_Report02_Error
    
    Dim DocNum      As String
    Dim ErrNum      As Integer
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim sQry        As String
    Dim DocEntry    As String '관리번호
    Dim CLTCOD      As String '사업장
    Dim DestNo1     As String '출장번호1
    Dim DestNo2     As String '출장번호2
    Dim MSTCOD      As String '사번
    Dim Destinat    As String '출장지(%)
    Dim Dest2       As String '출장지상세
    Dim CoCode      As String '작번(%)
    Dim FrDt        As String '기간(Fr)
    Dim ToDt        As String '기간(To)
    Dim Object      As String '목적(%)
    Dim Comments    As String '비고(%)
    Dim TeamCode    As String '팀
    Dim RegCls      As String '등록구분
    Dim ObjCls      As String '목적구분
    Dim DestCode    As String '출장지역
    Dim DestDiv     As String '출장구분
    Dim Vehicle     As String '차량구분
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY030_Print_Report02_Error
    End If
    
    DocEntry = oForm.Items("SDocEntry").Specific.VALUE '관리번호
    CLTCOD = Trim(oForm.Items("SCLTCOD").Specific.Selected.VALUE) '사업장
    DestNo1 = Trim(oForm.Items("SDestNo1").Specific.VALUE) '출장번호1
    DestNo2 = Trim(oForm.Items("SDestNo2").Specific.VALUE) '출장번호2
    MSTCOD = Trim(oForm.Items("SMSTCOD").Specific.VALUE) '사번
    Destinat = Trim(oForm.Items("SDestinat").Specific.Selected.VALUE) '출장지(%)
    Dest2 = Trim(oForm.Items("SDest2").Specific.VALUE) '출장지상세
    CoCode = Trim(oForm.Items("SCoCode").Specific.VALUE) '작번(%)
    FrDt = Trim(oForm.Items("SFrDate").Specific.VALUE) '기간(Fr)
    ToDt = Trim(oForm.Items("SToDate").Specific.VALUE) '기간(To)
    Object = Trim(oForm.Items("SObject").Specific.VALUE) '목적(%)
    Comments = Trim(oForm.Items("SComments").Specific.VALUE) '비고(%)
    TeamCode = Trim(oForm.Items("STeamCode").Specific.VALUE) '팀
    RegCls = Trim(oForm.Items("SRegCls").Specific.Selected.VALUE) '등록구분
    ObjCls = Trim(oForm.Items("SObjCls").Specific.Selected.VALUE) '목적구분
    DestCode = Trim(oForm.Items("SDestCode").Specific.Selected.VALUE) '출장지역
    DestDiv = Trim(oForm.Items("SDestDiv").Specific.Selected.VALUE) '출장구분
    Vehicle = Trim(oForm.Items("SVehicle").Specific.Selected.VALUE) '차량구분
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY030] 공용여비 신청서"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PH_PY030_04.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PH_PY030_05.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PH_PY030_06.rpt"
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
    sQry = "      EXEC [PH_PY030_91] '"
    sQry = sQry & DocEntry & "','" '관리번호
    sQry = sQry & CLTCOD & "','" '사업장
    sQry = sQry & DestNo1 & "','" '출장번호1
    sQry = sQry & DestNo2 & "','" '출장번호2
    sQry = sQry & MSTCOD & "','" '사번
    sQry = sQry & Destinat & "','" '출장지(%)
    sQry = sQry & Dest2 & "','" '출장지상세
    sQry = sQry & CoCode & "','" '작번(%)
    sQry = sQry & FrDt & "','" '기간(Fr)
    sQry = sQry & ToDt & "','" '기간(To)
    sQry = sQry & Object & "','" '목적(%)
    sQry = sQry & Comments & "','" '비고(%)
    sQry = sQry & TeamCode & "','" '팀
    sQry = sQry & RegCls & "','" '등록구분
    sQry = sQry & ObjCls & "','" '목적구분
    sQry = sQry & DestCode & "','" '출장지역
    sQry = sQry & DestDiv & "','" '출장구분
    sQry = sQry & Vehicle & "'" '차량구분
    
'    Call oRecordSet.DoQuery(sQry)
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY030_Print_Report02_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V", , 2) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY030_Print_Report02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY030_Print_Report02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub

