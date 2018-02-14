VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PH_PY203.cls
'//  Module             : 인사관리>인사관리>인사정보관리
'//  Desc               : 교육실적등록
'//  FormType           : PH_PY203
'//  Create Date(Start) : 2016.12.01
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm           As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
'Private oDS_PH_PY203A As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PH_PY203B  As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i         As Long
    Dim oInnerXml As String
    Dim oXmlDoc   As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY203.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY203_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01, "PH_PY203" '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    '//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm.Freeze True
    Call PH_PY203_CreateItems
    Call PH_PY203_ComboBox_Setting
    Call PH_PY203_CF_ChooseFromList
    Call PH_PY203_EnableMenus
    Call PH_PY203_SetDocument(oFromDocEntry01)
    Call PH_PY203_FormResize
    
    Call PH_PY203_Add_MatrixRow(0, True)
    Call PH_PY203_LoadCaption
    Call PH_PY203_FormItemEnabled
    
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

    Call PH_PY203_FormReset '폼초기화
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing

    '사번 포커스
'    Call oForm.Items("MSTCOD").CLICK
    
    Exit Sub
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY203_LoadCaption()
'******************************************************************************
'Function ID : PH_PY203_LoadCaption()
'해당모듈    : PH_PY203
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY203_LoadCaption_Error

    Call oForm.Freeze(True)

    If oForm.Mode = fm_ADD_MODE Then
        oForm.Items("BtnAdd").Specific.Caption = "추가"
'        oForm.Items("BtnDelete").Enabled = False
'    ElseIf oForm.Mode = fm_OK_MODE Then
'        oForm.Items("BtnAdd").Specific.Caption = "확인"
    ElseIf oForm.Mode = fm_UPDATE_MODE Then
        oForm.Items("BtnAdd").Specific.Caption = "수정"
'        oForm.Items("BtnDelete").Enabled = True
    End If
    
    Call oForm.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PH_PY203_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PH_PY203_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PH_PY203B.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PH_PY203B.Offset = oRow
    Call oDS_PH_PY203B.setValue("U_LineNum", oRow, oRow + 1)
    Call oDS_PH_PY203B.setValue("U_ColReg01", oRow, "Y")
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PH_PY203_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PH_PY203_MTX01()
'******************************************************************************
'Function ID : PH_PY203_MTX01()
'해당모듈    : PH_PY203
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY203_MTX01_Error

    Dim i        As Integer
    Dim sQry     As String
    Dim ErrNum   As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD   As String '사업장
    Dim TeamCode As String '부서
    Dim MSTCOD   As String '사번
    Dim StdYear  As String '기준년도
    Dim StdMt    As String '기준월
    Dim Status   As String '재직여부
    Dim EduName  As String '교육명
    Dim EduOrg   As String '교육기관
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '부서
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사번
    StdYear = Trim(oForm.Items("StdYear").Specific.VALUE) '기준년도
    StdMt = Trim(oForm.Items("StdMt").Specific.VALUE) '기준월
    Status = Trim(oForm.Items("Status").Specific.VALUE) '재직여부
    EduName = Trim(oForm.Items("EduName").Specific.VALUE) '교육명
    EduOrg = Trim(oForm.Items("EduOrg").Specific.VALUE) '교육기관
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "      EXEC [PH_PY203_01] '"
    sQry = sQry & CLTCOD & "','"  '사업장
    sQry = sQry & TeamCode & "','"  '부서
    sQry = sQry & MSTCOD & "','"  '사번
    sQry = sQry & StdYear & "','"  '기준년도
    sQry = sQry & StdMt & "','" '기준월
    sQry = sQry & Status & "','" '재직여부
    sQry = sQry & EduName & "','" '교육명
    sQry = sQry & EduOrg & "'" '교육기관
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PH_PY203B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm.Mode = fm_ADD_MODE
        
        Call PH_PY203_Add_MatrixRow(0, True)
        Call PH_PY203_LoadCaption
        
        GoTo PH_PY203_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PH_PY203B.Size Then
            oDS_PH_PY203B.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PH_PY203B.Offset = i
        
        Call oDS_PH_PY203B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY203B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PH_PY203B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("MSTCOD").VALUE)) '사번
        Call oDS_PH_PY203B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("MSTNAM").VALUE)) '성명
        Call oDS_PH_PY203B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("EduType").VALUE)) '분류
        Call oDS_PH_PY203B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("EduName").VALUE)) '교육명
        Call oDS_PH_PY203B.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("FrDt").VALUE), "YYYYMMDD")) '기간시작
        Call oDS_PH_PY203B.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("ToDt").VALUE), "YYYYMMDD")) '기간종료
        Call oDS_PH_PY203B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("EduOrg").VALUE)) '교육기관
        Call oDS_PH_PY203B.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("EduLoc").VALUE)) '교육장소
        Call oDS_PH_PY203B.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("EduHour").VALUE)) '교육시간
        Call oDS_PH_PY203B.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("EduAmt").VALUE)) '교육비
        Call oDS_PH_PY203B.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("EduAmt2").VALUE)) '출장비
        Call oDS_PH_PY203B.setValue("U_ColRgL01", i, Trim(oRecordSet01.Fields("Comment").VALUE)) '비고
        Call oDS_PH_PY203B.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("PlnIDX").VALUE)) '교육계획번호
        Call oDS_PH_PY203B.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("LineID").VALUE)) 'LineID
        
        'Call oDS_PH_PY203B.setValue("U_ColSum06", i, Trim(oRecordSet01.Fields("TotalExp").VALUE)) '합계
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    Call PH_PY203_Add_MatrixRow(oMat01.VisualRowCount)
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    
    Call oForm.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY203_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PH_PY203_DeleteData()
'******************************************************************************
'Function ID : PH_PY203_DeleteData()
'해당모듈    : PH_PY203
'기능        : 기본정보 삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY203_DeleteData_Error

    Dim i      As Integer
    Dim sQry   As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim MSTCOD As String
    Dim LineId As String
    Dim PlnIDX As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("삭제 중...", 100, False)
    
    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 2 '마직막 빈행 제외를 위해 2를 뺌
        If Trim(oDS_PH_PY203B.GetValue("U_ColReg01", i)) = "Y" Then
            
            MSTCOD = Trim(oDS_PH_PY203B.GetValue("U_ColReg02", i)) '사번
            LineId = IIf(Trim(oDS_PH_PY203B.GetValue("U_ColReg12", i)) = "", -1, Trim(oDS_PH_PY203B.GetValue("U_ColReg12", i))) '라인ID
            PlnIDX = Trim(oDS_PH_PY203B.GetValue("U_ColReg10", i)) '교육계획번호
            
            sQry = "            EXEC PH_PY203_04 "
            sQry = sQry & "'" & MSTCOD & "'," '사번
            sQry = sQry & "'" & LineId & "'," '라인ID
            sQry = sQry & "'" & PlnIDX & "'" '교육계획번호
            
            Call oRecordSet01.DoQuery(sQry)
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_DeleteData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY203_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PH_PY203_UpdateData() As Boolean
'******************************************************************************
'Function ID : PH_PY203_UpdateData()
'해당모듈    : PH_PY203
'기능        : 기본정보를 수정
'인수        : 없음
'반환값      : 없음
'특이사항    : 이 Function 사용 안함
'******************************************************************************
On Error GoTo PH_PY203_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer '관리번호
    Dim CLTCOD   As String '사업장
    Dim DestNo1  As String '출장번호1
    Dim DestNo2  As String '출장번호2
    Dim MSTCOD   As String '사원번호
    Dim MSTNAM   As String '사원성명
    Dim Destinat As String '출장지
    Dim Dest2    As String '출장지상세
    Dim CoCode   As String '작번
    Dim FrDate   As String '시작일자
    Dim FrTime   As String '시작시각
    Dim ToDate   As String '종료일자
    Dim ToTime   As String '종료시각
    Dim Object   As String '목적
    Dim Comments As String '비고
    Dim RegCls   As String '등록구분
    Dim ObjCls   As String '목적구분
    Dim DestCode As String '출장지역
    Dim DestDiv  As String '출장구분
    Dim Vehicle  As String '차량구분
    Dim FuelPrc  As Double '1L단가
    Dim FuelType As String '유류
    Dim Distance As Double '거리
    Dim TransExp As Double '교통비
    Dim DayExp   As Double '일비
    Dim FoodNum  As String '식수
    Dim FoodExp  As Double '식비
    Dim ParkExp  As Double '주차비
    Dim TollExp  As Double '도로비
    Dim TotalExp As Double '합계

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
        PH_PY203_UpdateData = False
        Exit Function
    End If
    
    sQry = "            EXEC [PH_PY203_03] "
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
    PH_PY203_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY203_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PH_PY203_AddData() As Boolean
'******************************************************************************
'Function ID : PH_PY203_AddData()
'해당모듈    : PH_PY203
'기능        : 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY203_AddData_Error

    Dim i       As Integer
    Dim sQry    As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD  As String '사업장코드
    Dim MSTCOD  As String '사번
    Dim MSTNAM  As String '성명
    Dim EduType As String '분류
    Dim EduName As String '교육명
    Dim FrDt    As String '기간시작
    Dim ToDt    As String '기간종료
    Dim EduOrg  As String '교육기관
    Dim EduLoc  As String '교육장소
    Dim EduHour As Double '교육시간
    Dim EduAmt  As Double '교육비
    Dim EduAmt2 As Double '출장비
    Dim Comment As String '비고
    Dim PlnIDX  As String '교육계획번호
    Dim LineId  As String '라인ID
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, False)
    
    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 2 '마직막 빈행 제외를 위해 2를 뺌
        If Trim(oDS_PH_PY203B.GetValue("U_ColReg01", i)) = "Y" Then
            
            CLTCOD = oForm.DataSources.UserDataSources("CLTCOD").VALUE '사업장코드
            MSTCOD = Trim(oDS_PH_PY203B.GetValue("U_ColReg02", i)) '사번
            MSTNAM = Trim(oDS_PH_PY203B.GetValue("U_ColReg03", i)) '성명
            EduType = Trim(oDS_PH_PY203B.GetValue("U_ColReg04", i)) '분류
            EduName = Trim(oDS_PH_PY203B.GetValue("U_ColReg05", i)) '교육명
            FrDt = Trim(oDS_PH_PY203B.GetValue("U_ColDt01", i)) '기간시작
            ToDt = Trim(oDS_PH_PY203B.GetValue("U_ColDt02", i)) '기간종료
            EduOrg = Trim(oDS_PH_PY203B.GetValue("U_ColReg08", i)) '교육기관
            EduLoc = Trim(oDS_PH_PY203B.GetValue("U_ColReg09", i))  '교육장소
            EduHour = Trim(oDS_PH_PY203B.GetValue("U_ColQty01", i)) '교육시간
            EduAmt = Trim(oDS_PH_PY203B.GetValue("U_ColSum01", i)) '교육비
            EduAmt2 = Trim(oDS_PH_PY203B.GetValue("U_ColSum02", i)) '출장비
            Comment = Trim(oDS_PH_PY203B.GetValue("U_ColRgL01", i)) '비고
            PlnIDX = Trim(oDS_PH_PY203B.GetValue("U_ColReg10", i)) '교육계획번호
            LineId = IIf(Trim(oDS_PH_PY203B.GetValue("U_ColReg12", i)) = "", -1, Trim(oDS_PH_PY203B.GetValue("U_ColReg12", i))) '라인ID
            
            sQry = "            EXEC [PH_PY203_02] "
            sQry = sQry & "'" & CLTCOD & "'," '사업장
            sQry = sQry & "'" & MSTCOD & "'," '사번
            sQry = sQry & "'" & MSTNAM & "'," '성명
            sQry = sQry & "'" & EduType & "',"  '분류
            sQry = sQry & "'" & EduName & "',"  '교육명
            sQry = sQry & "'" & FrDt & "',"  '기간시작
            sQry = sQry & "'" & ToDt & "',"  '기간종료
            sQry = sQry & "'" & EduOrg & "',"  '교육기관
            sQry = sQry & "'" & EduLoc & "',"  '교육장소
            sQry = sQry & "'" & EduHour & "',"  '교육시간
            sQry = sQry & "'" & EduAmt & "',"  '교육비
            sQry = sQry & "'" & EduAmt2 & "',"  '출장비
            sQry = sQry & "'" & Comment & "',"  '비고
            sQry = sQry & "'" & PlnIDX & "',"  '교육계획번호
            sQry = sQry & "'" & LineId & "'" '라인ID
            
            Call RecordSet01.DoQuery(sQry)
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    MDC_Com.MDC_GF_Message "저장 완료!", "S"
    
    Set RecordSet01 = Nothing
    
    PH_PY203_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PH_PY203_AddData = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY203_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PH_PY203_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PH_PY203_HeaderSpaceLineDel()
'해당모듈 : PH_PY203
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY203_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm.Items("DestNo1").Specific.VALUE) = "" '출장번호1
            ErrNum = 1
            GoTo PH_PY203_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("DestNo2").Specific.VALUE) = "" '출장번호2
            ErrNum = 2
            GoTo PH_PY203_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("MSTCOD").Specific.VALUE) = "" '사원번호
            ErrNum = 3
            GoTo PH_PY203_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrDate").Specific.VALUE) = "" '시작일자
            ErrNum = 4
            GoTo PH_PY203_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrTime").Specific.VALUE) = "" '시작시각
            ErrNum = 5
            GoTo PH_PY203_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToDate").Specific.VALUE) = "" '종료일자
            ErrNum = 6
            GoTo PH_PY203_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToTime").Specific.VALUE) = "" '종료시각
            ErrNum = 7
            GoTo PH_PY203_HeaderSpaceLineDel_Error
    End Select

    PH_PY203_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_HeaderSpaceLineDel_Error:
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
    PH_PY203_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PH_PY203_MatrixSpaceLineDel() As Boolean
On Error GoTo PH_PY203_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    ErrNum = 0

    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 2 '마직막 빈행 제외를 위해 2를 뺌
        If Trim(oDS_PH_PY203B.GetValue("U_ColReg01", i)) = "Y" Then
            
            If Trim(oDS_PH_PY203B.GetValue("U_ColReg02", i)) = "" Then  '사번
                ErrNum = 1
                GoTo PH_PY203_MatrixSpaceLineDel_Error
            ElseIf Trim(oDS_PH_PY203B.GetValue("U_ColReg03", i)) = "" Then '성명
                ErrNum = 2
                GoTo PH_PY203_MatrixSpaceLineDel_Error
            ElseIf Trim(oDS_PH_PY203B.GetValue("U_ColReg04", i)) = "" Then '분류
                ErrNum = 3
                GoTo PH_PY203_MatrixSpaceLineDel_Error
            ElseIf Trim(oDS_PH_PY203B.GetValue("U_ColReg05", i)) = "" Then '교육명
                ErrNum = 4
                GoTo PH_PY203_MatrixSpaceLineDel_Error
            ElseIf Trim(oDS_PH_PY203B.GetValue("U_ColDt01", i)) = "" Then '기간시작
                ErrNum = 5
                GoTo PH_PY203_MatrixSpaceLineDel_Error
            ElseIf Trim(oDS_PH_PY203B.GetValue("U_ColDt02", i)) = "" Then '기간종료
                ErrNum = 6
                GoTo PH_PY203_MatrixSpaceLineDel_Error
'            ElseIf Trim(oDS_PH_PY203B.GetValue("U_ColReg08", i)) = "" Then '교육기관
'                ErrNum = 7
'                GoTo PH_PY203_MatrixSpaceLineDel_Error
'            ElseIf Trim(oDS_PH_PY203B.GetValue("U_ColReg09", i)) = "" Then '교육장소
'                ErrNum = 8
'                GoTo PH_PY203_MatrixSpaceLineDel_Error
            End If

        End If
    Next

    PH_PY203_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_MatrixSpaceLineDel_Error:

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 사번이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 성명이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 분류가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 교육명이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 기간시작이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 기간종료가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 교육기관이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 교육장소가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PH_PY203_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PH_PY203_MatrixSpaceLineDel = False
    
End Function

Private Sub PH_PY203_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY203_FlushToItemValue_Error

    Dim loopCount  As Integer
    Dim ErrNum     As Integer
    Dim sQry       As String
    Dim ItemCode   As String
    
    Dim CLTCOD     As String
    
    Dim MSTCOD     As String '사번
    Dim MSTNAM     As String '성명
    Dim EduType    As String '교육구분
    Dim EduName    As String '교육명
    Dim EduFrDt    As String '교육시작일
    Dim EduToDt    As String '교육종료일
    Dim EduOrg     As String '교육기관
    Dim EduLoc     As String '교육장소
    Dim EduHour    As Double '교육시간
    Dim EduAmount  As Double '교육비
    Dim EduAmount2 As Double '출장비(0)
    Dim Comment    As String '비고('')
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case oUID

        Case "Mat01"
        
            Call oMat01.FlushToDataSource
        
            If oCol = "MSTCOD" Then
            
                Call oDS_PH_PY203B.setValue("U_ColReg02", oRow - 1, oMat01.Columns(oCol).Cells(oRow).Specific.VALUE)
                Call oDS_PH_PY203B.setValue("U_ColReg03", oRow - 1, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'"))
                If oMat01.RowCount = oRow And Trim(oDS_PH_PY203B.GetValue("U_ColReg02", oRow - 1)) <> "" Then
                    Call PH_PY203_Add_MatrixRow(oRow)
                End If
            
            ElseIf oCol = "PlnIDX" Then
            
                '교육계획에서 자료 조회
                sQry = "        SELECT  T0.MSTCOD AS [MSTCOD], " '사번
                sQry = sQry & "         T0.MSTNAM AS [MSTNAM], " '성명
                sQry = sQry & "         T0.Class AS [EduType], " '교육구분
                sQry = sQry & "         T0.EduName AS [EduName], " '교육명
                sQry = sQry & "         T0.EduFrDt AS [EduFrDt], " '교육시작일
                sQry = sQry & "         T0.EduToDt AS [EduToDt], " '교육종료일
                sQry = sQry & "         T0.EduOrg AS [EduOrg], " '교육기관
                sQry = sQry & "         '' AS [EduLoc], " '교육장소
                sQry = sQry & "         T0.EduHour AS [EduHour], " '교육시간
                sQry = sQry & "         T0.EduAmount AS [EduAmount], " '교육비
                sQry = sQry & "         0 AS [EduAmount2], " '출장비
                sQry = sQry & "         '' AS [Comment] " '비고
                sQry = sQry & " FROM    Z_PH_PY204 AS T0 "
                sQry = sQry & " WHERE   T0.IDX = '" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'"
            
                Call oRecordSet01.DoQuery(sQry)
            
                MSTCOD = Trim(oRecordSet01.Fields("MSTCOD").VALUE) '사번
                MSTNAM = Trim(oRecordSet01.Fields("MSTNAM").VALUE) '성명
                EduType = Trim(oRecordSet01.Fields("EduType").VALUE) '교육구분
                EduName = Trim(oRecordSet01.Fields("EduName").VALUE) '교육명
                EduFrDt = Format(Trim(oRecordSet01.Fields("EduFrDt").VALUE), "YYYYMMDD") '교육시작일
                EduToDt = Format(Trim(oRecordSet01.Fields("EduToDt").VALUE), "YYYYMMDD") '교육종료일
                EduOrg = Trim(oRecordSet01.Fields("EduOrg").VALUE) '교육기관
                EduLoc = Trim(oRecordSet01.Fields("EduLoc").VALUE) '교육장소
                EduHour = Trim(oRecordSet01.Fields("EduHour").VALUE) '교육시간
                EduAmount = Trim(oRecordSet01.Fields("EduAmount").VALUE) '교육비
                EduAmount2 = Trim(oRecordSet01.Fields("EduAmount2").VALUE) '출장비
                Comment = Trim(oRecordSet01.Fields("Comment").VALUE) '비고
            
                Call oDS_PH_PY203B.setValue("U_ColReg02", oRow - 1, MSTCOD) '사번
                Call oDS_PH_PY203B.setValue("U_ColReg03", oRow - 1, MSTNAM) '성명
                Call oDS_PH_PY203B.setValue("U_ColReg04", oRow - 1, EduType) '교육구분
                Call oDS_PH_PY203B.setValue("U_ColReg05", oRow - 1, EduName) '교육명
                Call oDS_PH_PY203B.setValue("U_ColDt01", oRow - 1, EduFrDt) '교육시작일
                Call oDS_PH_PY203B.setValue("U_ColDt02", oRow - 1, EduToDt) '교육종료일
                Call oDS_PH_PY203B.setValue("U_ColReg08", oRow - 1, EduOrg) '교육기관
                Call oDS_PH_PY203B.setValue("U_ColReg09", oRow - 1, EduLoc) '교육장소
                Call oDS_PH_PY203B.setValue("U_ColQty01", oRow - 1, EduHour) '교육시간
                Call oDS_PH_PY203B.setValue("U_ColSum01", oRow - 1, EduAmount) '교육비
                Call oDS_PH_PY203B.setValue("U_ColSum02", oRow - 1, EduAmount2) '출장비
                Call oDS_PH_PY203B.setValue("U_ColRgL01", oRow - 1, Comment) '비고
            
                If oMat01.RowCount = oRow And Trim(oDS_PH_PY203B.GetValue("U_ColReg10", oRow - 1)) <> "" Then
                    Call PH_PY203_Add_MatrixRow(oRow)
                End If
            
            End If

            Call oMat01.LoadFromDataSource
            '강제 포커스 이동_S
            Call oMat01.Columns(oCol).Cells(oRow).CLICK(ct_Regular)
            '강제 포커스 이동_E
            Call oMat01.AutoResizeColumns

        Case "CLTCOD"
            
            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
            
            If oForm.Items("TeamCode").Specific.ValidValues.Count > 0 Then
                For loopCount = oForm.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("TeamCode").Specific.ValidValues.Remove(loopCount, psk_Index)
                Next
            End If
            
            '부서콤보세팅
            Call oForm.Items("TeamCode").Specific.ValidValues.Add("%", "전체")
            sQry = "        SELECT      U_Code AS [Code],"
            sQry = sQry & "             U_CodeNm As [Name]"
            sQry = sQry & " FROM        [@PS_HR200L]"
            sQry = sQry & " WHERE       Code = '1'"
            sQry = sQry & "             AND U_UseYN = 'Y'"
            sQry = sQry & "             AND U_Char2 = '" & CLTCOD & "'"
            sQry = sQry & " ORDER BY    U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("TeamCode").Specific, sQry, "", False, False)
            Call oForm.Items("TeamCode").Specific.Select(0, psk_Index)

        Case "MSTCOD"

            oForm.Items("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.VALUE & "'") '성명
            
    End Select

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PH_PY203_FlushToItemValue_Error:

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PH_PY203_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PH_PY203_FormItemEnabled()
On Error GoTo PH_PY203_FormItemEnabled_Error
    If (oForm.Mode = fm_ADD_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_FIND_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_OK_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
    End If
    
    Exit Sub
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY203_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
            
                oForm.Mode = fm_ADD_MODE

                Call PH_PY203_FormReset
                Call PH_PY203_LoadCaption
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
'                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
'                Call oForm.Freeze(True)
'                Call PH_PY203_Add_MatrixRow(oMat01.VisualRowCount)
'                Call oForm.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PH_PY203_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PH_PY203A.Clear

'                Call PH_PY203_LoadCaption
'                Call PH_PY203_FormItemEnabled
                '//Call PH_PY203_FormItemEnabled '//UDO방식
                'Call PH_PY203_Add_MatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY203_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
'                '엑셀 내보내기 이후 처리
'                Call oForm.Freeze(True)
'                Call oDS_PH_PY203B.RemoveRecord(oDS_PH_PY203B.Size - 1)
'                Call oMat01.LoadFromDataSource
'                Call oForm.Freeze(False)
            
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
    
        If pval.ItemUID = "PH_PY203" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm.Mode = fm_ADD_MODE Then
            
'                If PH_PY203_HeaderSpaceLineDel() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

'                If PH_PY203_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                '매트릭스 필수자료 체크
                If PH_PY203_MatrixSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                If PH_PY203_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                Call PH_PY203_FormReset
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY203_LoadCaption
                Call PH_PY203_MTX01
                
                oLast_Mode = oForm.Mode
                
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            
                If PH_PY203_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PH_PY203_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                If PH_PY203_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PH_PY203_FormReset
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY203_LoadCaption
                Call PH_PY203_MTX01
                
'                oForm.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
'            Call PH_PY203_FormReset
            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PH_PY203_LoadCaption
            Call PH_PY203_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PH_PY203_DeleteData
'                Call PH_PY203_FormReset
                oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PH_PY203_LoadCaption
                Call PH_PY203_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnSel" Then '전체선택
        
            Call PH_PY203_CheckAll
        
        ElseIf pval.ItemUID = "BtnExcel" Then '엑셀업로드
        
            Call PH_PY203_Excel_Upload
        
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            Call PH_PY203_Print_Report01
        
        ElseIf pval.ItemUID = "BtnPrint2" Then
        
            Call PH_PY203_Print_Report02
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PH_PY203" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD", "") 'Header-사번
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "Mat01", "MSTCOD") 'Matrix-사번
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "Mat01", "PlnIDX") 'Matrix-교육계획번호
        
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
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "Mat01" Then
        
        Else
        
            Call PH_PY203_FlushToItemValue(pval.ItemUID)
        
        End If
    
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
            
                If pval.ColUID = "MSTCOD" Or pval.ColUID = "PlnIDX" Then
            
                    Call PH_PY203_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                
                    'oMat01.AutoResizeColumns
                
                End If
            
'                oMat01.FlushToDataSource
'
'                If (pval.ColUID = "MSTCOD") Then
'                    Call oDS_PH_PY203B.setValue("U_ColReg02", pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    Call oDS_PH_PY203B.setValue("U_ColReg03", pval.Row - 1, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"))
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PH_PY203B.GetValue("U_ColReg02", pval.Row - 1)) <> "" Then
'                        Call PH_PY203_Add_MatrixRow(pval.Row)
'                    End If
'                Else
''                    Call oDS_PH_PY203B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
'
'                oMat01.LoadFromDataSource
                oMat01.AutoResizeColumns
                
            Else
                
                Call PH_PY203_FlushToItemValue(pval.ItemUID)
                
'                If pval.ItemUID = "MSTCOD" Then
'
'                    oForm.Items("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.VALUE & "'") '성명
'
'                End If
                
            End If

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
        Call PH_PY203_FormItemEnabled
        '//Call PH_PY203_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY203_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY203A", "U_CardCode,U_CardName")
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
'            If (PH_PY203_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PH_PY203B.RemoveRecord(oDS_PH_PY203B.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PH_PY203_Add_MatrixRow(0)
            Else
                If Trim(oDS_PH_PY203B.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PH_PY203_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY203_CreateItems() As Boolean
On Error GoTo PH_PY203_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Set oDS_PH_PY203B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    
    '부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Call oForm.Items("TeamCode").Specific.DataBind.SetBound(True, "", "TeamCode")
    
    '재직여부
    Call oForm.DataSources.UserDataSources.Add("Status", dt_SHORT_TEXT, 10)
    Call oForm.Items("Status").Specific.DataBind.SetBound(True, "", "Status")
    
    '사원번호
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 20)
    Call oForm.Items("MSTCOD").Specific.DataBind.SetBound(True, "", "MSTCOD")
    
    '사원성명
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 50)
    Call oForm.Items("MSTNAM").Specific.DataBind.SetBound(True, "", "MSTNAM")
    
    '기준년도
    Call oForm.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 4)
    Call oForm.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")
    
    '기준월
    Call oForm.DataSources.UserDataSources.Add("StdMt", dt_SHORT_TEXT, 4)
    Call oForm.Items("StdMt").Specific.DataBind.SetBound(True, "", "StdMt")
    
    '교육명
    Call oForm.DataSources.UserDataSources.Add("EduName", dt_SHORT_TEXT, 200)
    Call oForm.Items("EduName").Specific.DataBind.SetBound(True, "", "EduName")
    
    '교육기관
    Call oForm.DataSources.UserDataSources.Add("EduOrg", dt_SHORT_TEXT, 100)
    Call oForm.Items("EduOrg").Specific.DataBind.SetBound(True, "", "EduOrg")
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY203_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY203_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PH_PY203_ComboBox_Setting()
On Error GoTo PH_PY203_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
    '재직여부
    Call oForm.Items("Status").Specific.ValidValues.Add("1", "재직자")
    Call oForm.Items("Status").Specific.ValidValues.Add("2", "퇴직자포함")
    Call oForm.Items("Status").Specific.Select(0, psk_Index)

    '기준월
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm AS [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = '4'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    
    Call oForm.Items("StdMt").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm.Items("StdMt").Specific, sQry, "", False, False)
    Call oForm.Items("StdMt").Specific.Select(0, psk_Index)

    '//////////매트릭스//////////
    
    '분류
    sQry = "        SELECT      edType AS [Code],"
    sQry = sQry & "             name As [Name]"
    sQry = sQry & " FROM        [OHED]"
'    sQry = sQry & "  WHERE      Code = 'P226'"
'    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    edType"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("EduType"), sQry)
    
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY203_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY203_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY203_CF_ChooseFromList()
On Error GoTo PH_PY203_CF_ChooseFromList_Error

    Exit Sub
PH_PY203_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY203_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY203_EnableMenus()
On Error GoTo PH_PY203_EnableMenus_Error
    
    Exit Sub
PH_PY203_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY203_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY203_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY203_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PH_PY203_FormItemEnabled
        '//Call PH_PY203_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY203_FormItemEnabled
'        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm.Items("1").Click ct_Regular
    End If
    Exit Sub
PH_PY203_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY203_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY203_FormResize()
On Error GoTo PH_PY203_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PH_PY203_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY203_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PH_PY203_FormReset()
'******************************************************************************
'Function ID : PH_PY203_FormReset()
'해당모듈 : PH_PY203
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY203_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
    Dim User_BPLID As String
    User_BPLID = MDC_PS_Common.User_BPLID

    '헤더 초기화
    oForm.DataSources.UserDataSources("CLTCOD").VALUE = User_BPLID '사업장
    oForm.DataSources.UserDataSources("TeamCode").VALUE = "%" '부서
    oForm.DataSources.UserDataSources("MSTCOD").VALUE = "" ''사번
    oForm.DataSources.UserDataSources("MSTNAM").VALUE = "" ''성명
    oForm.DataSources.UserDataSources("Status").VALUE = "1" '재직여부
    oForm.DataSources.UserDataSources("StdYear").VALUE = Format(Now, "YYYY") '기준년도
    oForm.DataSources.UserDataSources("EduName").VALUE = "" '교육명
    oForm.DataSources.UserDataSources("EduOrg").VALUE = "" '교육기관
    
    '라인 초기화
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    Call PH_PY203_Add_MatrixRow(0, True)
    
    Set RecordSet01 = Nothing
    Call oForm.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY203_FormReset_Error:
    Call oForm.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY203_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PH_PY203_CheckAll()
On Error GoTo PH_PY203_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PH_PY203B.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PH_PY203B.Offset = loopCount
        If CheckType = "N" Then
            Call oDS_PH_PY203B.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PH_PY203B.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm.Freeze(False)

    Exit Sub
    
PH_PY203_CheckAll_Error:
    Call oForm.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PH_PY203_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Sub PH_PY203_Excel_Upload()
    
    Dim i       As Long
    Dim j       As Long
    Dim sPrice  As String
    Dim sFile   As String
    Dim OneRec  As String
    Dim sQry    As String
    Dim CLTCOD  As String
    
    Dim xl      As Excel.Application
    Dim xlwb    As Excel.Workbook
    Dim xlsh    As Excel.Worksheet
    
    Dim oEdit   As SAPbouiCOM.EditText
    Dim oForm   As SAPbouiCOM.Form
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo Err
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Set oForm = Sbo_Application.Forms.ActiveForm
    
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    
    sFile = FileListBoxForm.OpenDialog(FileListBoxForm, "*.xls", "파일선택", "C:\")
    
    If sFile = "" Then
        Sbo_Application.StatusBar.SetText "파일을 선택해 주세요.", bmt_Short, smt_Error
        Exit Sub
    Else
        If Mid(Right(sFile, 4), 1, 3) = "xls" Or Mid(Right(sFile, 5), 1, 4) = "xlsx" Then
'            oForm.Items("Comments").Specific.VALUE = sFile

        Else
            Sbo_Application.StatusBar.SetText "엑셀파일이 아닙니다.", bmt_Short, smt_Error
            Exit Sub
        End If
    End If
    
    '엑셀 Object 연결
    Set xl = CreateObject("excel.application")
    Set xlwb = xl.Workbooks.Open(sFile, , True)
    Set xlsh = xlwb.Worksheets(1)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    For i = 2 To xlsh.UsedRange.Rows.Count
        If i - 2 <> 0 Then
            oDS_PH_PY203B.InsertRecord (i - 2)
        End If
        oDS_PH_PY203B.Offset = i - 2
        Call oDS_PH_PY203B.setValue("U_LineNum", i - 2, i - 1)
        Call oDS_PH_PY203B.setValue("U_ColReg01", i - 2, "Y")
        Call oDS_PH_PY203B.setValue("U_ColReg02", i - 2, xlsh.Cells(i, 1))
        Call oDS_PH_PY203B.setValue("U_ColReg03", i - 2, xlsh.Cells(i, 2))
        Call oDS_PH_PY203B.setValue("U_ColReg04", i - 2, xlsh.Cells(i, 3))
        Call oDS_PH_PY203B.setValue("U_ColReg05", i - 2, xlsh.Cells(i, 4))
        Call oDS_PH_PY203B.setValue("U_ColDt01", i - 2, Format(xlsh.Cells(i, 5), "YYYYMMDD"))
        Call oDS_PH_PY203B.setValue("U_ColDt02", i - 2, Format(xlsh.Cells(i, 6), "YYYYMMDD"))
        Call oDS_PH_PY203B.setValue("U_ColReg08", i - 2, xlsh.Cells(i, 7))
        Call oDS_PH_PY203B.setValue("U_ColReg09", i - 2, xlsh.Cells(i, 8))
        Call oDS_PH_PY203B.setValue("U_ColQty01", i - 2, xlsh.Cells(i, 9))
        Call oDS_PH_PY203B.setValue("U_ColSum01", i - 2, xlsh.Cells(i, 10))
        Call oDS_PH_PY203B.setValue("U_ColSum02", i - 2, xlsh.Cells(i, 11))
        Call oDS_PH_PY203B.setValue("U_ColRgL01", i - 2, xlsh.Cells(i, 12))
        Call oDS_PH_PY203B.setValue("U_ColReg10", i - 2, xlsh.Cells(i, 13))
        
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & xlsh.UsedRange.Rows.Count - 1 & "건 UPLOAD중...!"
    Next i
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm.Update

    Call PH_PY203_Add_MatrixRow(oMat01.RowCount)
    
    Call oForm.Freeze(False)
    Sbo_Application.StatusBar.SetText "엑셀을 불러왔습니다.", bmt_Short, smt_Warning

    '액셀개체 닫음
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    '진행바 초기화
    Exit Sub

Err:
    
    Call oForm.Freeze(False)

    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
End Sub

Private Sub PH_PY203_Print_Report01()

    Dim DocNum      As String
    Dim ErrNum      As Integer
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY203_Print_Report01_Error
    
    Dim CLTCOD  As String
    Dim DestNo1 As String
    Dim DestNo2 As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY203_Print_Report01_Error
    End If
    
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DestNo1 = Trim(oForm.Items("DestNo1").Specific.VALUE)
    DestNo2 = Trim(oForm.Items("DestNo2").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY203] 공용증"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PH_PY203_01.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PH_PY203_02.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PH_PY203_03.rpt"
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
    sQry = "EXEC [PH_PY203_90] '" & CLTCOD & "','" & DestNo1 & "','" & DestNo2 & "'"
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY203_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY203_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY203_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub

Private Sub PH_PY203_Print_Report02()
On Error GoTo PH_PY203_Print_Report02_Error
    
    Dim DocNum      As String
    Dim ErrNum      As Integer
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim DocEntry    As String '관리번호
    Dim CLTCOD      As String '사업장
    Dim DestNo1     As String '출장번호1
    Dim DestNo2     As String  '출장번호2
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
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY203_Print_Report02_Error
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

    WinTitle = "[PH_PY203] 공용여비 신청서"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PH_PY203_04.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PH_PY203_05.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PH_PY203_06.rpt"
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
    sQry = "      EXEC [PH_PY203_91] '"
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
    
    Call oRecordSet.DoQuery(sQry)
    
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY203_Print_Report02_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V", , 2) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY203_Print_Report02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY203_Print_Report02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub

