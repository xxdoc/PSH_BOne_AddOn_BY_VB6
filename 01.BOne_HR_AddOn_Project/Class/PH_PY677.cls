VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY677"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PH_PY677.cls
'//  Module             : 인사관리>근태관리>근태리포트
'//  Desc               : 일일근태이상자조회
'//  FormType           : PH_PY677
'//  Create Date(Start) : 2013.03.25
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
Private oDS_PH_PY677A  As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PH_PY677B  As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i         As Long
    Dim oInnerXml As String
    Dim oXmlDoc   As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY677.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY677_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01, "PH_PY677" '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    '//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm.Freeze True
    Call PH_PY677_CreateItems
    Call PH_PY677_ComboBox_Setting
    Call PH_PY677_CF_ChooseFromList
    Call PH_PY677_EnableMenus
    Call PH_PY677_SetDocument(oFromDocEntry01)
    Call PH_PY677_FormResize
    
'    Call PH_PY677_Add_MatrixRow(0, True)
'    Call PH_PY677_LoadCaption
    Call PH_PY677_FormItemEnabled
    
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
        
'    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY677A]"
'    Call RecordSet01.DoQuery(sQry)
'    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'        Call oDS_PH_PY677A.setValue("DocEntry", 0, 1)
'    Else
'        Call oDS_PH_PY677A.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
'    End If
'
'    Call PH_PY677_FormReset '폼초기화 추가(2013.01.29 송명규)
    
    oForm.Update
    oForm.Freeze False
    
    '기간(월)
    oForm.Items("FrDate").Specific.VALUE = Format(Now, "YYYYMM01")
    oForm.Items("ToDate").Specific.VALUE = Format(Now, "YYYYMMDD")
    '사번 포커스
    Call oForm.Items("MSTCOD").CLICK
    
    oForm.Items("BtnChk").Specific.Caption = "전체선택"
    
    oForm.Visible = True
    Set oXmlDoc = Nothing

    Exit Sub
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY677_LoadCaption()
'******************************************************************************
'Function ID : PH_PY677_LoadCaption()
'해당모듈    : PH_PY677
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY677_LoadCaption_Error

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
PH_PY677_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub PH_PY677_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
'******************************************************************************
'Function ID : PH_PY677_Add_MatrixRow
'해당모듈    : PH_PY677
'기능        : 메트릭스 행 추가
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY677_Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PH_PY677B.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PH_PY677B.Offset = oRow
    oDS_PH_PY677B.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub

PH_PY677_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PH_PY677_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PH_PY677_MTX01()
'******************************************************************************
'Function ID : PH_PY677_MTX01()
'해당모듈    : PH_PY677
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY677_MTX01_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD   As String '사업장
    Dim FrDate   As String '시작일자
    Dim ToDate   As String '종료일자
    Dim TeamCode As String '부서
    Dim RspCode  As String '담당
    Dim ClsCode  As String '반
    Dim ShiftDat As String '근무형태
    Dim GNMUJO   As String '근무조
    Dim MSTCOD   As String '사원번호
    Dim Class    As String '기찰이상구분(2014.04.10 송명규 추가)
    Dim Confirm  As String '근태기찰정상확인(2013.03.29 송명규 추가)
    Dim WorkType As String '근태구분(2014.05.13 송명규 추가)
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    FrDate = Trim(oForm.Items("FrDate").Specific.VALUE) '시작일자
    ToDate = Trim(oForm.Items("ToDate").Specific.VALUE) '종료일자
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '부서
    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE) '담당
    ClsCode = Trim(oForm.Items("ClsCode").Specific.VALUE) '반
    ShiftDat = Trim(oForm.Items("ShiftDatCd").Specific.VALUE) '근무형태
    GNMUJO = Trim(oForm.Items("GNMUJOCd").Specific.VALUE) '근무조
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사원번호
    Class = Trim(oForm.Items("Class").Specific.VALUE) '기찰이상구분
    Confirm = Trim(oForm.Items("Confirm").Specific.VALUE) '근태기찰정상확인
    WorkType = Trim(oForm.Items("WorkType").Specific.VALUE) '근태구분
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "            EXEC [PH_PY677_01] "
    sQry = sQry & "'" & CLTCOD & "'," '사업장
    sQry = sQry & "'" & FrDate & "',"  '시작일자
    sQry = sQry & "'" & ToDate & "',"  '종료일자
    sQry = sQry & "'" & TeamCode & "',"  '부서
    sQry = sQry & "'" & RspCode & "',"  '담당
    sQry = sQry & "'" & ClsCode & "',"  '반
    sQry = sQry & "'" & ShiftDat & "',"  '근무형태
    sQry = sQry & "'" & GNMUJO & "',"  '근무조
    sQry = sQry & "'" & MSTCOD & "',"  '사원번호
    sQry = sQry & "'" & Class & "',"  '기찰이상구분
    sQry = sQry & "'" & Confirm & "',"  '근태기찰정상확인
    sQry = sQry & "'" & WorkType & "'"  '근태구분
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PH_PY677B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm.Mode = fm_ADD_MODE
        
'        Call PH_PY677_Add_MatrixRow(0, True)
'        Call PH_PY677_LoadCaption
        
        GoTo PH_PY677_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PH_PY677B.Size Then
            oDS_PH_PY677B.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PH_PY677B.Offset = i
        
        Call oDS_PH_PY677B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY677B.setValue("U_ColReg20", i, Trim(oRecordSet01.Fields("Chk").VALUE)) '선택
        Call oDS_PH_PY677B.setValue("U_ColDt05", i, Format(Trim(oRecordSet01.Fields("PosDate").VALUE), "YYYYMMDD")) '일자
        Call oDS_PH_PY677B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("TeamName").VALUE)) '부서
        Call oDS_PH_PY677B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("RspName").VALUE)) '담당
        Call oDS_PH_PY677B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("ClsName").VALUE)) '반
        Call oDS_PH_PY677B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("MSTCOD").VALUE)) '사번
        Call oDS_PH_PY677B.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("MSTNAM").VALUE)) '성명
        Call oDS_PH_PY677B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("ShiftDat").VALUE)) '근무형태
        Call oDS_PH_PY677B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("GNMUJO").VALUE)) '근무조
        Call oDS_PH_PY677B.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("DayWeek").VALUE)) '요일
        Call oDS_PH_PY677B.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("DayType").VALUE)) '요일구분
        Call oDS_PH_PY677B.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("P_GetDt").VALUE), "YYYYMMDD")) '출근일자(계획)
        Call oDS_PH_PY677B.setValue("U_ColTm01", i, Trim(oRecordSet01.Fields("P_GetTime").VALUE)) '출근시각(계획)
        Call oDS_PH_PY677B.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("P_OffDt").VALUE), "YYYYMMDD")) '퇴근일자(계획)
        Call oDS_PH_PY677B.setValue("U_ColTm02", i, Trim(oRecordSet01.Fields("P_OffTime").VALUE)) '퇴근시각(계획)
        Call oDS_PH_PY677B.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("P_Base").VALUE)) '기본(계획)
        Call oDS_PH_PY677B.setValue("U_ColQty02", i, Trim(oRecordSet01.Fields("P_Extend").VALUE)) '연장(계획)
        Call oDS_PH_PY677B.setValue("U_ColQty03", i, Trim(oRecordSet01.Fields("P_MidNt").VALUE)) '심야
        Call oDS_PH_PY677B.setValue("U_ColQty04", i, Trim(oRecordSet01.Fields("P_Special").VALUE)) '특근(계획)
        Call oDS_PH_PY677B.setValue("U_ColQty05", i, Trim(oRecordSet01.Fields("P_SpExtend").VALUE)) '특연(계획)
        Call oDS_PH_PY677B.setValue("U_ColQty06", i, Trim(oRecordSet01.Fields("P_EarlyTo").VALUE)) '조출
        Call oDS_PH_PY677B.setValue("U_ColQty07", i, Trim(oRecordSet01.Fields("P_SEarlyTo").VALUE)) '휴조
        Call oDS_PH_PY677B.setValue("U_ColQty08", i, Trim(oRecordSet01.Fields("P_EducTran").VALUE)) '교육훈련
        Call oDS_PH_PY677B.setValue("U_ColQty09", i, Trim(oRecordSet01.Fields("P_LateTo").VALUE)) '지각
        Call oDS_PH_PY677B.setValue("U_ColQty10", i, Trim(oRecordSet01.Fields("P_EarlyOff").VALUE)) '조퇴
        Call oDS_PH_PY677B.setValue("U_ColReg21", i, Trim(oRecordSet01.Fields("P_WorkType").VALUE)) '근태구분
        Call oDS_PH_PY677B.setValue("U_ColReg22", i, Trim(oRecordSet01.Fields("P_Comment").VALUE)) '비고
        Call oDS_PH_PY677B.setValue("U_ColQty11", i, Trim(oRecordSet01.Fields("P_GoOut").VALUE)) '외출
        Call oDS_PH_PY677B.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("P_Confirm").VALUE)) '기찰정상확인
        Call oDS_PH_PY677B.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("R_GetDt").VALUE)) '출근일자(기찰)
        Call oDS_PH_PY677B.setValue("U_ColTm03", i, Trim(oRecordSet01.Fields("R_GetTime").VALUE)) '출근시각(기찰)
        Call oDS_PH_PY677B.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("R_OffDt").VALUE)) '퇴근일자(기찰)
        Call oDS_PH_PY677B.setValue("U_ColTm04", i, Trim(oRecordSet01.Fields("R_OffTime").VALUE)) '퇴근시각(기찰)
'        Call oDS_PH_PY677B.setValue("U_ColQty05", i, Trim(oRecordSet01.Fields("R_Base").VALUE)) '기본(기찰)
'        Call oDS_PH_PY677B.setValue("U_ColQty06", i, Trim(oRecordSet01.Fields("R_Extend").VALUE)) '연장(기찰)
'        Call oDS_PH_PY677B.setValue("U_ColQty07", i, Trim(oRecordSet01.Fields("R_Special").VALUE)) '특근(기찰)
'        Call oDS_PH_PY677B.setValue("U_ColQty08", i, Trim(oRecordSet01.Fields("R_SpExtend").VALUE)) '특연(기찰)
'        Call oDS_PH_PY677B.setValue("U_ColQty09", i, Trim(oRecordSet01.Fields("R_TotTime").VALUE)) '총근무시간(기찰)
        Call oDS_PH_PY677B.setValue("U_ColQty12", i, Trim(oRecordSet01.Fields("Rotation").VALUE)) '교대일수
        Call oDS_PH_PY677B.setValue("U_ColReg24", i, Trim(oRecordSet01.Fields("R_YN").VALUE)) '기찰완료여부
        Call oDS_PH_PY677B.setValue("U_ColReg25", i, Trim(oRecordSet01.Fields("WkAbCls").VALUE)) '근태이상분류
        Call oDS_PH_PY677B.setValue("U_ColReg26", i, Trim(oRecordSet01.Fields("WkAbCmt").VALUE)) '근태이상사유
        Call oDS_PH_PY677B.setValue("U_ColReg27", i, Trim(oRecordSet01.Fields("ActText").VALUE)) '근무내용
        Call oDS_PH_PY677B.setValue("U_ColReg28", i, Trim(oRecordSet01.Fields("RotateYN").VALUE)) '교대인정
        
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
PH_PY677_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY677_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PH_PY677_DeleteData()
'******************************************************************************
'Function ID : PH_PY677_DeleteData()
'해당모듈    : PH_PY677
'기능        : 기본정보 삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY677_DeleteData_Error

    Dim i      As Integer
    Dim sQry   As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PH_PY677A] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm.Mode = fm_ADD_MODE
            GoTo PH_PY677_DeleteData_Error
        Else
            sQry = "EXEC PH_PY677_04 '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
'    Call PH_PY677_FormReset
    
'    oForm.Mode = fm_ADD_MODE
    
'    Call oForm.Items("BtnSearch").Click(ct_Regular)
    
'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PH_PY677_Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY677_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY677_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PH_PY677_UpdateData() As Boolean
'******************************************************************************
'Function ID : PH_PY677_UpdateData()
'해당모듈    : PH_PY677
'기능        : 기본정보를 수정
'인수        : 없음
'반환값      : 없음
'특이사항    : 1. 근무형태, 근무조 추가(2017.09.05 송명규)
'******************************************************************************
On Error GoTo PH_PY677_UpdateData_Error

    Dim i    As Integer
    Dim j    As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Dim CLTCOD     As String '사업장
    Dim PosDate    As String '일자
    Dim MSTCOD     As String '사번
    Dim ShiftDat   As String '근무형태
    Dim GNMUJO     As String '근무조
    Dim P_WorkType As String '근태구분
    Dim P_Confirm  As String '기찰정상확인
    Dim P_GetTime  As String '출근시각
    Dim P_OffDt    As String '퇴근일자
    Dim P_OffTime  As String '퇴근시각
    Dim P_Base     As Double '기본
    Dim P_Extend   As Double '연장
    Dim P_Special  As Double '특근
    Dim P_SpExtend As Double '특연
    Dim P_Midnight As Double '심야
    Dim P_EarlyTo  As Double '조출
    Dim P_SEarlyTo As Double '휴조
    Dim P_EducTran As Double '교육훈련
    Dim P_LateTo   As Double '지각
    Dim P_EarlyOff As Double '조퇴
    Dim P_GoOut    As Double '외출
    Dim P_Comment  As String '비고
    Dim WkAbCls    As String '근태이상분류
    Dim WkAbCmt    As String '근태이상사유
    Dim ActText    As String '근무내용
    Dim RotateYN   As String '교대인정
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, False)
    
    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PH_PY677B.GetValue("U_ColReg20", i)) = "Y" Then
            
            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
            PosDate = Trim(oDS_PH_PY677B.GetValue("U_ColDt05", i)) '일자
            MSTCOD = Trim(oDS_PH_PY677B.GetValue("U_ColReg05", i)) '사번
            ShiftDat = Trim(oDS_PH_PY677B.GetValue("U_ColReg07", i)) '근무형태
            GNMUJO = Trim(oDS_PH_PY677B.GetValue("U_ColReg08", i)) '근무조
            P_WorkType = Trim(oDS_PH_PY677B.GetValue("U_ColReg21", i)) '근태구분
            P_Confirm = Trim(oDS_PH_PY677B.GetValue("U_ColReg16", i)) '기찰정상확인
            P_GetTime = Trim(oDS_PH_PY677B.GetValue("U_ColTm01", i)) '출근시각
            P_OffDt = Trim(oDS_PH_PY677B.GetValue("U_ColDt02", i)) '퇴근일자
            P_OffTime = Trim(oDS_PH_PY677B.GetValue("U_ColTm02", i)) '퇴근시각
            P_Base = Trim(oDS_PH_PY677B.GetValue("U_ColQty01", i)) '기본
            P_Extend = Trim(oDS_PH_PY677B.GetValue("U_ColQty02", i)) '연장
            P_Special = Trim(oDS_PH_PY677B.GetValue("U_ColQty04", i)) '특근
            P_SpExtend = Trim(oDS_PH_PY677B.GetValue("U_ColQty05", i)) '특연
            P_Midnight = Trim(oDS_PH_PY677B.GetValue("U_ColQty03", i)) '심야
            P_EarlyTo = Trim(oDS_PH_PY677B.GetValue("U_ColQty06", i)) '조출
            P_SEarlyTo = Trim(oDS_PH_PY677B.GetValue("U_ColQty07", i)) '휴조
            P_EducTran = Trim(oDS_PH_PY677B.GetValue("U_ColQty08", i)) '교육훈련
            P_LateTo = Trim(oDS_PH_PY677B.GetValue("U_ColQty09", i)) '지각
            P_EarlyOff = Trim(oDS_PH_PY677B.GetValue("U_ColQty10", i)) '조퇴
            P_GoOut = Trim(oDS_PH_PY677B.GetValue("U_ColQty11", i)) '외출
            P_Comment = Trim(oDS_PH_PY677B.GetValue("U_ColReg22", i)) '비고
            WkAbCls = Trim(oDS_PH_PY677B.GetValue("U_ColReg25", i)) '근태이상분류
            WkAbCmt = Trim(oDS_PH_PY677B.GetValue("U_ColReg26", i)) '근태이상사유
            ActText = Trim(oDS_PH_PY677B.GetValue("U_ColReg27", i)) '근무내용
            RotateYN = Trim(oDS_PH_PY677B.GetValue("U_ColReg28", i)) '교대인정
            
            sQry = "            EXEC [PH_PY677_02] "
            sQry = sQry & "'" & CLTCOD & "'," '사업장
            sQry = sQry & "'" & PosDate & "'," '일자
            sQry = sQry & "'" & MSTCOD & "',"  '사번
            sQry = sQry & "'" & ShiftDat & "',"  '근무형태
            sQry = sQry & "'" & GNMUJO & "',"  '근무조
            sQry = sQry & "'" & P_WorkType & "',"  '근태구분
            sQry = sQry & "'" & P_Confirm & "',"  '기찰정상확인
            sQry = sQry & "'" & P_GetTime & "',"  '출근시각
            sQry = sQry & "'" & P_OffDt & "',"  '퇴근일자
            sQry = sQry & "'" & P_OffTime & "',"  '퇴근시각
            sQry = sQry & "'" & P_Base & "',"  '기본
            sQry = sQry & "'" & P_Extend & "',"  '연장
            sQry = sQry & "'" & P_Special & "',"  '특근
            sQry = sQry & "'" & P_SpExtend & "',"  '특연
            sQry = sQry & "'" & P_Midnight & "',"  '심야
            sQry = sQry & "'" & P_EarlyTo & "',"  '조출
            sQry = sQry & "'" & P_SEarlyTo & "',"  '휴조
            sQry = sQry & "'" & P_EducTran & "',"  '교육훈련
            sQry = sQry & "'" & P_LateTo & "',"  '지각
            sQry = sQry & "'" & P_EarlyOff & "',"  '조퇴
            sQry = sQry & "'" & P_GoOut & "',"  '외출
            sQry = sQry & "'" & P_Comment & "',"  '비고
            sQry = sQry & "'" & WkAbCls & "'," '근태이상분류
            sQry = sQry & "'" & WkAbCmt & "'," '근태이상사유
            sQry = sQry & "'" & ActText & "'," '근무내용
            sQry = sQry & "'" & RotateYN & "'" '교대인정
            
            Call RecordSet01.DoQuery(sQry)
        End If
    Next

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PH_PY677_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY677_UpdateData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY677_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PH_PY677_AddData() As Boolean
'******************************************************************************
'Function ID : PH_PY677_AddData()
'해당모듈    : PH_PY677
'기능        : 데이터 INSERT
'인수        : 없음
'반환값      : 성공여부
'특이사항    : 이 클래스에서는 사용 안함
'******************************************************************************
On Error GoTo PH_PY677_AddData_Error

    Dim i    As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = oCompany.GetBusinessObject(BoRecordset)
    
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
    Dim UserSign As String 'UserSign

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
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY677A]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "            EXEC [PH_PY677_02] "
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
    PH_PY677_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY677_AddData_Error:
    PH_PY677_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY677_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PH_PY677_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PH_PY677_HeaderSpaceLineDel()
'해당모듈    : PH_PY677
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY677_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm.Items("DestNo1").Specific.VALUE) = "" '출장번호1
            ErrNum = 1
            GoTo PH_PY677_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("DestNo2").Specific.VALUE) = "" '출장번호2
            ErrNum = 2
            GoTo PH_PY677_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("MSTCOD").Specific.VALUE) = "" '사원번호
            ErrNum = 3
            GoTo PH_PY677_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrDate").Specific.VALUE) = "" '시작일자
            ErrNum = 4
            GoTo PH_PY677_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrTime").Specific.VALUE) = "" '시작시각
            ErrNum = 5
            GoTo PH_PY677_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToDate").Specific.VALUE) = "" '종료일자
            ErrNum = 6
            GoTo PH_PY677_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToTime").Specific.VALUE) = "" '종료시각
            ErrNum = 7
            GoTo PH_PY677_HeaderSpaceLineDel_Error
    End Select

    PH_PY677_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY677_HeaderSpaceLineDel_Error:
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
    PH_PY677_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PH_PY677_MatrixSpaceLineDel() As Boolean
On Error GoTo PH_PY677_MatrixSpaceLineDel_Error

    Dim i       As Long
    Dim ErrNum  As Integer
    Dim sQry As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
Set oRecordSet01 = Nothing
    PH_PY677_MatrixSpaceLineDel = True
Exit Function

PH_PY677_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PH_PY677_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PH_PY677_MatrixSpaceLineDel = False
End Function

Private Sub PH_PY677_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY677_FlushToItemValue_Error

    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim sQry     As String
    Dim ItemCode As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD   As String
    Dim TeamCode As String
    Dim RspCode  As String
    
    Dim PreWorkType As String
    Dim WorkType    As String
    Dim JanQty      As Double
    Dim ymd         As String
    Dim MSTCOD      As String
    Dim YY          As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    
    Call oForm.Freeze(True)
    
    Select Case oUID
        
        Case "Mat01"
            
            PreWorkType = Trim(oDS_PH_PY677B.GetValue("U_ColReg21", oRow - 1))
            
            Call oMat01.FlushToDataSource
            
            If oCol = "P_GetTime" Then
            
                Call PH_PY677_Time_ReSet(oRow)
                Call oMat01.LoadFromDataSource
                Call oMat01.Columns(oCol).Cells(oRow).CLICK
            
            ElseIf oCol = "P_OffTime" Then
            
                If oDS_PH_PY677B.GetValue("U_ColTm02", oRow - 1) <> "0000" Then
                    Call PH_PY677_Time_Calc_Main(oDS_PH_PY677B.GetValue("U_ColTm02", oRow - 1), oRow)
                    Call oMat01.LoadFromDataSource
                    Call oMat01.Columns(oCol).Cells(oRow).CLICK
                End If
            
            ElseIf oCol = "P_WorkType" Then
                
                WorkType = Trim(oDS_PH_PY677B.GetValue("U_ColReg21", oRow - 1))
                
                Select Case WorkType

                    Case "A01", "A02", "E02", "E03", "F01", "F02", "F03", "F04", "F05" '무단결근, 유계결근, 무급휴일, 휴업, 병가(휴직), 신병휴직, 정직(유결), 가사휴직, 공상휴직(F01) 추가(2017.12.07)
                        
                        Call oDS_PH_PY677B.setValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1)) '퇴근일자(P_OffDt)
                        Call oDS_PH_PY677B.setValue("U_ColTm01", oRow - 1, "00:00") '출근시각(P_GetTime)
                        Call oDS_PH_PY677B.setValue("U_ColTm02", oRow - 1, "00:00") '퇴근시각(P_OffTime)

                        Call PH_PY677_Time_ReSet(oRow)
                        Call oMat01.LoadFromDataSource

                    Case "C02", "D04", "D05", "D06", "D07", "H05" '훈련, 경조휴가, 하기휴가, 특별휴가, 분만휴가, 조합활동

                        Call oDS_PH_PY677B.setValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1)) '퇴근일자(P_OffDt)
                        Call oDS_PH_PY677B.setValue("U_ColTm01", oRow - 1, "00:00") '출근시각(P_GetTime)
                        Call oDS_PH_PY677B.setValue("U_ColTm02", oRow - 1, "00:00") '퇴근시각(P_OffTime)
                        Call oDS_PH_PY677B.setValue("U_ColQty12", oRow - 1, 0) '교대일수

                        Call PH_PY677_Time_ReSet(oRow)
                        Call oMat01.LoadFromDataSource

                    Case "D02", "D09" '연차/반차 휴가
                    
                        '연차/반차 휴가 잔여일수 확인
                        If WorkType = "D02" Then
                            JanQty = 1
                        ElseIf WorkType = "D09" Then
                            JanQty = 0.5
                        End If

                        ymd = Trim(oDS_PH_PY677B.GetValue("U_ColDt05", oRow - 1))
                        CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
                        MSTCOD = Trim(oDS_PH_PY677B.GetValue("U_ColReg05", oRow - 1))
                        
                        Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", oRecordSet01.RecordCount, False)
                        
                        sQry = "      EXEC [PH_PY775_01] '"
                        sQry = sQry & CLTCOD & "','"
                        sQry = sQry & Left(ymd, 4) & "','"
                        sQry = sQry & MSTCOD & "'"
                        
                        Call oRecordSet01.DoQuery(sQry)
                        
                        If oRecordSet01.Fields("jandd").VALUE < JanQty Then
                            
                            ErrNum = 1
                            Call oDS_PH_PY677B.setValue("U_ColReg21", oRow - 1, "A00")
                        
                            Call oMat01.LoadFromDataSource
                            
                            GoTo PH_PY677_FlushToItemValue_Error
                            
                        Else
                            
                            Call oDS_PH_PY677B.setValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1)) '퇴근일자(P_OffDt)
                            Call oDS_PH_PY677B.setValue("U_ColTm01", oRow - 1, "00:00") '출근시각(P_GetTime)
                            Call oDS_PH_PY677B.setValue("U_ColTm02", oRow - 1, "00:00") '퇴근시각(P_OffTime)
                            Call oDS_PH_PY677B.setValue("U_ColQty12", oRow - 1, 1) '교대일수
    
                            Call PH_PY677_Time_ReSet(oRow)
                            
                            Call oMat01.LoadFromDataSource
                            
                        End If
                        
                        ProgBar01.VALUE = 100
                        Call ProgBar01.Stop
                        
                    Case "D08", "D10" '근속보전휴가, 근속보전반차(기계사업부)
                    
                        '근속보전휴가 잔량 확인
                        ymd = Trim(oDS_PH_PY677B.GetValue("U_ColDt05", oRow - 1))
                        CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
                        MSTCOD = Trim(oDS_PH_PY677B.GetValue("U_ColReg05", oRow - 1))

                        If Trim(oDS_PH_PY677B.GetValue("U_ColDt05", oRow - 1)) >= Left(ymd, 4) + "0701" Then
                            YY = Left(ymd, 4)
                        Else
                            YY = CStr(CInt(Left(ymd, 4)) - 1)
                        End If

                        Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", oRecordSet01.RecordCount, False)

                        sQry = "        SELECT     Bqty = ISNULL(( "
                        sQry = sQry & "                               SELECT     ISNULL(c.Qty,0) "
                        sQry = sQry & "                               FROM       [ZPH_GUNSOKCHA] c"
                        sQry = sQry & "                               WHERE      c.Year = '" & YY & "'"
                        sQry = sQry & "                                          AND ("
                        sQry = sQry & "                                                   CASE "
                        sQry = sQry & "                                                       WHEN ISDATE(a.U_GrpDat) = 0 THEN a.U_StartDat "
                        sQry = sQry & "                                                       ELSE A.U_GrpDat "
                        sQry = sQry & "                                                   END"
                        sQry = sQry & "                                              ) BETWEEN c.DocDateFr AND c.DocDateTo"
                        sQry = sQry & "                           ),0), " '발생수량
                        sQry = sQry & "            Sqty = ISNULL(( "
                        sQry = sQry & "                               SELECT     COUNT(*)"
                        sQry = sQry & "                               FROM       [ZPH_PY008] c"
                        sQry = sQry & "                               WHERE      c.CLTCOD = A.U_CLTCOD"
                        sQry = sQry & "                                          AND c.PosDate BETWEEN '" & YY & "' + '0701' AND CONVERT(CHAR(4),CONVERT(INTEGER,'" & YY & "') + 1 ) + '0630'"
                        sQry = sQry & "                                          AND c.MSTCOD  = a.Code"
                        sQry = sQry & "                                          AND c.WorkType = 'D08'"
                        sQry = sQry & "                                          AND c.PosDate <> '" & ymd & "'"
                        sQry = sQry & "                          ),0) " '보전년차 사용수량
                        sQry = sQry & "                    + "
                        sQry = sQry & "                    ISNULL(( "
                        sQry = sQry & "                               SELECT     COUNT(*) / 2.0"
                        sQry = sQry & "                               FROM       [ZPH_PY008] c"
                        sQry = sQry & "                               WHERE      c.CLTCOD = A.U_CLTCOD"
                        sQry = sQry & "                                          AND c.PosDate BETWEEN '" & YY & "' + '0701' AND CONVERT(CHAR(4),CONVERT(INTEGER,'" & YY & "') + 1 ) + '0630'"
                        sQry = sQry & "                                          AND c.MSTCOD = a.Code"
                        sQry = sQry & "                                          AND c.WorkType = 'D10'"
                        sQry = sQry & "                                          AND c.PosDate <> '" & ymd & "'"
                        sQry = sQry & "                           ),0)" '보전반차 사용수량
                        sQry = sQry & " FROM       [@PH_PY001A] a"
                        sQry = sQry & " WHERE      a.U_CLTCOD = '" & CLTCOD & "'"
                        sQry = sQry & "            AND a.Code = '" & MSTCOD & "'"

                        Call oRecordSet01.DoQuery(sQry)

                        If WorkType = "D08" Then
                            JanQty = 1
                        ElseIf WorkType = "D10" Then
                            JanQty = 0.5
                        End If
                        
                        If oRecordSet01.Fields("Bqty").VALUE - oRecordSet01.Fields("Sqty").VALUE < JanQty Then
                        
                            ErrNum = 2
                            Call oDS_PH_PY677B.setValue("U_ColReg21", oRow - 1, "A00")
                            Call oMat01.LoadFromDataSource
                            
                            GoTo PH_PY677_FlushToItemValue_Error
                           
                        Else
                        
                            Call oDS_PH_PY677B.setValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1)) '퇴근일자(P_OffDt)
                            Call oDS_PH_PY677B.setValue("U_ColTm01", oRow - 1, "00:00") '출근시각(P_GetTime)
                            Call oDS_PH_PY677B.setValue("U_ColTm02", oRow - 1, "00:00") '퇴근시각(P_OffTime)
                            Call oDS_PH_PY677B.setValue("U_ColQty12", oRow - 1, 1) '교대일수
    
                            Call PH_PY677_Time_ReSet(oRow)
                            
                            Call oMat01.LoadFromDataSource
                            
                        End If
                        
                        ProgBar01.VALUE = 100
                        Call ProgBar01.Stop
                        
                End Select
            End If
            
'            Call oMat01.LoadFromDataSource
'            Call oMat01.AutoResizeColumns
            
            
        Case "MSTCOD"
        
            oForm.Items("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.VALUE & "'") '성명
            
        Case "ShiftDatCd"
        
            oForm.Items("ShiftDatNm").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L] AS T0", "'" & oForm.Items("ShiftDatCd").Specific.VALUE & "'", " AND T0.Code = 'P154' AND T0.U_UseYN = 'Y'") '근무형태

        Case "GNMUJOCd"

            oForm.Items("GNMUJONm").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L] AS T0", "'" & oForm.Items("GNMUJOCd").Specific.VALUE & "'", " AND T0.Code = 'P155' AND T0.U_UseYN = 'Y'") '근무조
        
        Case "CLTCOD"
            
            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
            
            If oForm.Items("TeamCode").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("TeamCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '부서콤보세팅
            Call oForm.Items("TeamCode").Specific.ValidValues.Add("%", "전체")
            sQry = "        SELECT    U_Code AS [Code],"
            sQry = sQry & "           U_CodeNm As [Name]"
            sQry = sQry & " FROM      [@PS_HR200L]"
            sQry = sQry & " WHERE     Code = '1'"
            sQry = sQry & "           AND U_UseYN = 'Y'"
            sQry = sQry & "           AND U_Char2 = '" & CLTCOD & "'"
            sQry = sQry & " ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("TeamCode").Specific, sQry, "", False, False)
            Call oForm.Items("TeamCode").Specific.Select(0, psk_Index)
            
        Case "TeamCode"
            
            TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
            
            If oForm.Items("RspCode").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("RspCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("RspCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '담당콤보세팅
            Call oForm.Items("RspCode").Specific.ValidValues.Add("%", "전체")
            sQry = "        SELECT    U_Code AS [Code],"
            sQry = sQry & "           U_CodeNm As [Name]"
            sQry = sQry & " FROM      [@PS_HR200L]"
            sQry = sQry & " WHERE     Code = '2'"
            sQry = sQry & "           AND U_UseYN = 'Y'"
            sQry = sQry & "           AND U_Char1 = '" & TeamCode & "'"
            sQry = sQry & " ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("RspCode").Specific, sQry, "", False, False)
            Call oForm.Items("RspCode").Specific.Select(0, psk_Index)
        
        Case "RspCode"
            
            TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
            RspCode = Trim(oForm.Items("RspCode").Specific.VALUE)
            
            If oForm.Items("ClsCode").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("ClsCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("ClsCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '반콤보세팅
            Call oForm.Items("ClsCode").Specific.ValidValues.Add("%", "전체")
            sQry = "        SELECT    U_Code AS [Code],"
            sQry = sQry & "           U_CodeNm As [Name]"
            sQry = sQry & " FROM      [@PS_HR200L]"
            sQry = sQry & " WHERE     Code = '9'"
            sQry = sQry & "           AND U_UseYN = 'Y'"
            sQry = sQry & "           AND U_Char1 = '" & RspCode & "'"
            sQry = sQry & "           AND U_Char2 = '" & TeamCode & "'"
            sQry = sQry & " ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("ClsCode").Specific, sQry, "", False, False)
            Call oForm.Items("ClsCode").Specific.Select(0, psk_Index)
        
    End Select

    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
PH_PY677_FlushToItemValue_Error:

    Call oForm.Freeze(False)
'    Call ProgBar01.Stop
'    ProgBar01.VALUE = 100
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing

    If ErrNum = 1 Then
        Call Sbo_Application.SetStatusBarMessage("연차휴가 잔여일수가 없습니다. 확인바랍니다.", bmt_Short, True)
    ElseIf ErrNum = 2 Then
        Call Sbo_Application.SetStatusBarMessage("보전휴가 잔여일수가 없습니다. 확인바랍니다.", bmt_Short, True)
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY677_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If

End Sub

'/폼의 아이템 사용지정
Sub PH_PY677_FormItemEnabled()
On Error GoTo PH_PY677_FormItemEnabled_Error
    If (oForm.Mode = fm_ADD_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_FIND_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_OK_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
    End If
    
    Exit Sub

PH_PY677_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY677_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
'                Call PH_PY677_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
'                oForm.Mode = fm_ADD_MODE
'                BubbleEvent = False
'                Call PH_PY677_LoadCaption
                
                'oForm.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PH_PY677_Add_MatrixRow(oMat01.VisualRowCount)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PH_PY677_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PH_PY677A.Clear

'                Call PH_PY677_LoadCaption
'                Call PH_PY677_FormItemEnabled
                '//Call PH_PY677_FormItemEnabled '//UDO방식
                '//Call PH_PY677_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY677_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm.Freeze(True)
                Call oDS_PH_PY677B.RemoveRecord(oDS_PH_PY677B.Size - 1)
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
    
        If pval.ItemUID = "PH_PY677" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnUpdate" Then '수정 버튼
        
            If oForm.Mode = fm_ADD_MODE Then
            
                Call PH_PY677_UpdateData
            
'                If PH_PY677_HeaderSpaceLineDel() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
''                If PH_PY677_DataCheck() = False Then
''                    BubbleEvent = False
''                    Exit Sub
''                End If
'
''                If PH_PY677_AddData() = False Then
''                    BubbleEvent = False
''                    Exit Sub
''                End If
'
''                Call PH_PY677_FormReset
'                oForm.Mode = fm_ADD_MODE
'
''                Call PH_PY677_LoadCaption
'                Call PH_PY677_MTX01
'
'                oLast_Mode = oForm.Mode
                
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            
                If PH_PY677_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PH_PY677_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
'                If PH_PY677_UpdateData() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
'                Call PH_PY677_FormReset
'                oForm.Mode = fm_ADD_MODE
'
'                Call PH_PY677_LoadCaption
'                Call PH_PY677_MTX01
                
'                oForm.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
'            Call PH_PY677_FormReset
'            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
'            Call PH_PY677_LoadCaption
            Call PH_PY677_MTX01
            
'        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
'
'            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
'
'                Call PH_PY677_DeleteData
'                Call PH_PY677_FormReset
'                oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
'
'                Call PH_PY677_LoadCaption
'                Call PH_PY677_MTX01
'
'            Else
'
'            End If
        
'        ElseIf pval.ItemUID = "BtnPrint" Then
'
'            Call PH_PY677_Print_Report01
        
        ElseIf pval.ItemUID = "BtnChk" Then
        
'            If oForm.Items("BtnChk").Specific.Caption = "전체선택" Then
'                oForm.Items("BtnChk").Specific.Caption = "전체해제"
'            Else
'                oForm.Items("BtnChk").Specific.Caption = "전체선택"
'            End If
            
            Call PH_PY677_CheckAll
            
        ElseIf pval.ItemUID = "BtnConfirm" Then
        
            Call PH_PY677_ConfirmAll
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PH_PY677" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD", "") '사번
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "ShiftDatCd", "") '근무형태
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "GNMUJOCd", "") '근무조
        
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
                
'                Call oForm.Freeze(True)
'
'                'DataSource를 이용하여 각 컨트롤에 값을 출력
'                Call oDS_PH_PY677A.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
'                Call oDS_PH_PY677A.setValue("U_CLTCOD", 0, oMat01.Columns("CLTCOD").Cells(pval.Row).Specific.VALUE) '사업장
'                Call oDS_PH_PY677A.setValue("U_DestNo1", 0, oMat01.Columns("DestNo1").Cells(pval.Row).Specific.VALUE) '출장번호1
'                Call oDS_PH_PY677A.setValue("U_DestNo2", 0, oMat01.Columns("DestNo2").Cells(pval.Row).Specific.VALUE) '출장번호2
'                Call oDS_PH_PY677A.setValue("U_MSTCOD", 0, oMat01.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE) '사원번호
'                Call oDS_PH_PY677A.setValue("U_MSTNAM", 0, oMat01.Columns("MSTNAM").Cells(pval.Row).Specific.VALUE) '사원성명
'                Call oDS_PH_PY677A.setValue("U_Destinat", 0, oMat01.Columns("Destinat").Cells(pval.Row).Specific.VALUE) '출장지
'                Call oDS_PH_PY677A.setValue("U_Dest2", 0, oMat01.Columns("Dest2").Cells(pval.Row).Specific.VALUE) '출장지상세
'                Call oDS_PH_PY677A.setValue("U_CoCode", 0, oMat01.Columns("CoCode").Cells(pval.Row).Specific.VALUE) '작번
'                Call oDS_PH_PY677A.setValue("U_FrDate", 0, Replace(oMat01.Columns("FrDate").Cells(pval.Row).Specific.VALUE, ".", "")) '시작일자
'                Call oDS_PH_PY677A.setValue("U_FrTime", 0, oMat01.Columns("FrTime").Cells(pval.Row).Specific.VALUE) '시작시각
'                Call oDS_PH_PY677A.setValue("U_ToDate", 0, Replace(oMat01.Columns("ToDate").Cells(pval.Row).Specific.VALUE, ".", "")) '종료일자
'                Call oDS_PH_PY677A.setValue("U_ToTime", 0, oMat01.Columns("ToTime").Cells(pval.Row).Specific.VALUE) '종료시각
'                Call oDS_PH_PY677A.setValue("U_Object", 0, oMat01.Columns("Object").Cells(pval.Row).Specific.VALUE) '목적
'                Call oDS_PH_PY677A.setValue("U_Comments", 0, oMat01.Columns("Comments").Cells(pval.Row).Specific.VALUE) '비고
'                Call oDS_PH_PY677A.setValue("U_RegCls", 0, oMat01.Columns("RegCls").Cells(pval.Row).Specific.VALUE) '등록구분
'                Call oDS_PH_PY677A.setValue("U_ObjCls", 0, oMat01.Columns("ObjCls").Cells(pval.Row).Specific.VALUE) '목적구분
'                Call oDS_PH_PY677A.setValue("U_DestCode", 0, oMat01.Columns("DestCode").Cells(pval.Row).Specific.VALUE) '출장지역
'                Call oDS_PH_PY677A.setValue("U_DestDiv", 0, oMat01.Columns("DestDiv").Cells(pval.Row).Specific.VALUE) '출장구분
'                Call oDS_PH_PY677A.setValue("U_Vehicle", 0, oMat01.Columns("Vehicle").Cells(pval.Row).Specific.VALUE) '차량구분
'                Call oDS_PH_PY677A.setValue("U_FuelPrc", 0, oMat01.Columns("FuelPrc").Cells(pval.Row).Specific.VALUE) '1L단가
'                Call oDS_PH_PY677A.setValue("U_FuelType", 0, oMat01.Columns("FuelType").Cells(pval.Row).Specific.VALUE) '유류
'                Call oDS_PH_PY677A.setValue("U_Distance", 0, oMat01.Columns("Distance").Cells(pval.Row).Specific.VALUE) '거리
'                Call oDS_PH_PY677A.setValue("U_TransExp", 0, oMat01.Columns("TransExp").Cells(pval.Row).Specific.VALUE) '교통비
'                Call oDS_PH_PY677A.setValue("U_DayExp", 0, oMat01.Columns("DayExp").Cells(pval.Row).Specific.VALUE) '일비
'                Call oDS_PH_PY677A.setValue("U_FoodNum", 0, oMat01.Columns("FoodNum").Cells(pval.Row).Specific.VALUE) '식수
'                Call oDS_PH_PY677A.setValue("U_FoodExp", 0, oMat01.Columns("FoodExp").Cells(pval.Row).Specific.VALUE) '식비
'                Call oDS_PH_PY677A.setValue("U_ParkExp", 0, oMat01.Columns("ParkExp").Cells(pval.Row).Specific.VALUE) '주차비
'                Call oDS_PH_PY677A.setValue("U_TollExp", 0, oMat01.Columns("TollExp").Cells(pval.Row).Specific.VALUE) '도로비
'                Call oDS_PH_PY677A.setValue("U_TotalExp", 0, oMat01.Columns("TotalExp").Cells(pval.Row).Specific.VALUE) '합계
'
'                oForm.Mode = fm_UPDATE_MODE
'                Call PH_PY677_LoadCaption
'
'                Call oForm.Freeze(False)
                
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
        
'            If pval.ColUID = "P_Confirm" Or pval.ColUID = "P_WorkType" Then
            If pval.ColUID = "P_WorkType" Then

                Call PH_PY677_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
            
            End If
        
        Else
    
            Call PH_PY677_FlushToItemValue(pval.ItemUID)
        
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
            

                If pval.ColUID = "P_GetTime" Or pval.ColUID = "P_OffTime" Then

                    Call PH_PY677_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                
                End If
                
'                oMat01.LoadFromDataSource
                oMat01.AutoResizeColumns
                
            Else
                
                Call PH_PY677_FlushToItemValue(pval.ItemUID)
                
            End If
            
            oForm.Update
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
        Call PH_PY677_FormItemEnabled
        '//Call PH_PY677_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY677_FormResize
        Call oMat01.AutoResizeColumns
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY677A", "U_CardCode,U_CardName")
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
'            If (PH_PY677_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PH_PY677A.RemoveRecord(oDS_PH_PY677A.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PH_PY677_Add_MatrixRow(0)
            Else
                If Trim(oDS_PH_PY677A.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PH_PY677_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY677_CreateItems() As Boolean
On Error GoTo PH_PY677_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
'    Set oDS_PH_PY677A = oForm.DataSources.DBDataSources("@PH_PY677A")
    Set oDS_PH_PY677B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    
    '시작일자
    Call oForm.DataSources.UserDataSources.Add("FrDate", dt_DATE)
    Call oForm.Items("FrDate").Specific.DataBind.SetBound(True, "", "FrDate")
    
    '종료일자
    Call oForm.DataSources.UserDataSources.Add("ToDate", dt_DATE)
    Call oForm.Items("ToDate").Specific.DataBind.SetBound(True, "", "ToDate")
    
    '부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("TeamCode").Specific.DataBind.SetBound(True, "", "TeamCode")
    
    '담당
    Call oForm.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("RspCode").Specific.DataBind.SetBound(True, "", "RspCode")
    
    '반
    Call oForm.DataSources.UserDataSources.Add("ClsCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("ClsCode").Specific.DataBind.SetBound(True, "", "ClsCode")
    
    '근무형태 (코드)
    Call oForm.DataSources.UserDataSources.Add("ShiftDatCd", dt_SHORT_TEXT, 20)
    Call oForm.Items("ShiftDatCd").Specific.DataBind.SetBound(True, "", "ShiftDatCd")
    
    '근무형태(명)
    Call oForm.DataSources.UserDataSources.Add("ShiftDatNm", dt_SHORT_TEXT, 50)
    Call oForm.Items("ShiftDatNm").Specific.DataBind.SetBound(True, "", "ShiftDatNm")
    
    '근무조(코드)
    Call oForm.DataSources.UserDataSources.Add("GNMUJOCd", dt_SHORT_TEXT, 20)
    Call oForm.Items("GNMUJOCd").Specific.DataBind.SetBound(True, "", "GNMUJOCd")
    
    '근무조(명)
    Call oForm.DataSources.UserDataSources.Add("GNMUJONm", dt_SHORT_TEXT, 50)
    Call oForm.Items("GNMUJONm").Specific.DataBind.SetBound(True, "", "GNMUJONm")
        
    '사원번호
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 20)
    Call oForm.Items("MSTCOD").Specific.DataBind.SetBound(True, "", "MSTCOD")
    
    '사원성명
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 50)
    Call oForm.Items("MSTNAM").Specific.DataBind.SetBound(True, "", "MSTNAM")
    
    '근태기찰정상확인
    Call oForm.DataSources.UserDataSources.Add("Confirm", dt_SHORT_TEXT, 10)
    Call oForm.Items("Confirm").Specific.DataBind.SetBound(True, "", "Confirm")
    
    '기찰이상구분
    Call oForm.DataSources.UserDataSources.Add("Class", dt_SHORT_TEXT, 10)
    Call oForm.Items("Class").Specific.DataBind.SetBound(True, "", "Class")
    
    '근태구분(2014.05.13 송명규 추가)
    Call oForm.DataSources.UserDataSources.Add("WorkType", dt_SHORT_TEXT, 10)
    Call oForm.Items("WorkType").Specific.DataBind.SetBound(True, "", "WorkType")
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY677_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY677_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PH_PY677_ComboBox_Setting()
On Error GoTo PH_PY677_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)

    '헤더
    '기찰이상구분
    Set oCombo = oForm.Items("Class").Specific
    Call oCombo.ValidValues.Add("%", "전체")
    Call oCombo.ValidValues.Add("Y", "근태이상")
    Call oCombo.ValidValues.Add("N", "정상")
    Call oCombo.Select("Y", psk_ByValue)
    
    '근태이상자 확인
    Set oCombo = oForm.Items("Confirm").Specific
    Call oCombo.ValidValues.Add("%", "전체")
    Call oCombo.ValidValues.Add("N", "미확인[N]")
    Call oCombo.ValidValues.Add("Y", "확인[Y]")
    oForm.Items("Confirm").DisplayDesc = True
    Call oCombo.Select("N", psk_ByValue)
    
    '근태구분
    Call oForm.Items("WorkType").Specific.ValidValues.Add("%", "전체")
    sQry = "        SELECT    U_Code AS [Code],"
    sQry = sQry & "           U_CodeNm As [Name]"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = 'P221'"
    sQry = sQry & "           AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("WorkType").Specific, sQry, "", False, False)
    Call oForm.Items("WorkType").Specific.Select(0, psk_Index)

    '//////////매트릭스//////////
    
    '근무형태
    sQry = "        SELECT    U_Code AS [Code],"
    sQry = sQry & "           U_CodeNm As [Name]"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = 'P154'"
    sQry = sQry & "           AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ShiftDat"), sQry)
    
    '근무조
    sQry = "        SELECT    U_Code AS [Code],"
    sQry = sQry & "           U_CodeNm As [Name]"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = 'P155'"
    sQry = sQry & "           AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("GNMUJO"), sQry)
    
    '요일구분
    sQry = "        SELECT    U_Code AS [Code],"
    sQry = sQry & "           U_CodeNm As [Name]"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = 'P202'"
    sQry = sQry & "           AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("DayType"), sQry)
    
    '근태구분
    sQry = "        SELECT    U_Code AS [Code],"
    sQry = sQry & "           U_CodeNm As [Name]"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = 'P221'"
    sQry = sQry & "           AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("P_WorkType"), sQry)
    
    '기찰정상확인
'    Set oCombo = oMat01.Columns("P_Confirm")
    Call oMat01.Columns("P_Confirm").ValidValues.Add("N", "미확인[N]")
    Call oMat01.Columns("P_Confirm").ValidValues.Add("Y", "확인[Y]")
    oMat01.Columns("P_Confirm").DisplayDesc = True
'    Call oCombo.Select(0, psk_Index)

    '근태이상분류
    sQry = "        SELECT    U_Code AS [Code],"
    sQry = sQry & "           U_CodeNm As [Name]"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = 'P237'"
    sQry = sQry & "           AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("WkAbCls"), sQry)
    
    '교대인정(2016.12.07
    Call oMat01.Columns("RotateYN").ValidValues.Add("N", "N")
    Call oMat01.Columns("RotateYN").ValidValues.Add("Y", "Y")
    
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY677_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY677_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY677_CF_ChooseFromList()
On Error GoTo PH_PY677_CF_ChooseFromList_Error

    Exit Sub
PH_PY677_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY677_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY677_EnableMenus()
On Error GoTo PH_PY677_EnableMenus_Error
    
    Exit Sub
PH_PY677_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY677_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY677_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY677_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PH_PY677_FormItemEnabled
        '//Call PH_PY677_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY677_FormItemEnabled
'        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm.Items("1").Click ct_Regular
    End If
    Exit Sub
PH_PY677_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY677_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY677_FormResize()
On Error GoTo PH_PY677_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PH_PY677_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY677_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY677_Time_ReSet(ByVal pRow As Integer)
On Error GoTo PH_PY677_Time_ReSet_Error

    '//시간 초기화
'    Call oForm.Freeze(True)
    
    Call oDS_PH_PY677B.setValue("U_ColQty01", pRow - 1, 0) '기본
    Call oDS_PH_PY677B.setValue("U_ColQty02", pRow - 1, 0) '연장
    Call oDS_PH_PY677B.setValue("U_ColQty03", pRow - 1, 0) '심야
    Call oDS_PH_PY677B.setValue("U_ColQty06", pRow - 1, 0) '조출
    Call oDS_PH_PY677B.setValue("U_ColQty04", pRow - 1, 0) '특근
    Call oDS_PH_PY677B.setValue("U_ColQty05", pRow - 1, 0) '특연
    Call oDS_PH_PY677B.setValue("U_ColQty08", pRow - 1, 0) '교육훈련
    Call oDS_PH_PY677B.setValue("U_ColQty07", pRow - 1, 0) '휴조
    Call oDS_PH_PY677B.setValue("U_ColQty09", pRow - 1, 0) '지각
    Call oDS_PH_PY677B.setValue("U_ColQty10", pRow - 1, 0) '조퇴
    Call oDS_PH_PY677B.setValue("U_ColQty11", pRow - 1, 0) '외출
    
    Call oDS_PH_PY677B.setValue("U_ColTm02", pRow - 1, "00:00") '퇴근시각
    
'    oForm.DataSources.UserDataSources.Item("Base").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("Extend").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("Midnight").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("EarlyTo").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("Special").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("SpExtend").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("EducTran").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("SEarlyTo").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("LateTo").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("EarlyOff").VALUE = 0
'    oForm.DataSources.UserDataSources.Item("GoOut").VALUE = 0
'
'    oForm.DataSources.UserDataSources.Item("GOFrTime").VALUE = "0000"
'    oForm.DataSources.UserDataSources.Item("GOToTime").VALUE = "0000"
'    oForm.DataSources.UserDataSources.Item("GOFrTim2").VALUE = "0000"
'    oForm.DataSources.UserDataSources.Item("GOToTim2").VALUE = "0000"
    
'    Call oForm.Freeze(False)
    Exit Sub

PH_PY677_Time_ReSet_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY677_Time_ReSet_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY677_Time_Calc_Main(ByVal OffTime As String, ByVal pRow As Integer)
On Error GoTo PH_PY677_Time_Calc_Main_Error

    '//근태구분 Select
    Dim i        As Integer
    Dim CLTCOD   As String
    Dim ShiftDat As String '//근무형태
    Dim GNMUJO   As String '//근무조
    Dim DayOff   As String '//평일휴일구분
    
    Dim PosDate As String '//기준일
    Dim GetDate As String '//출근일
    Dim OffDate As String '//퇴근일
    Dim GetTime As String '//출근시간
    'Dim OffTime As String '//퇴근시간
    
    Dim FromTime As String
    Dim ToTime   As String
    
    Dim hTime1 As Double '//오전10분휴식시간
    Dim hTime2 As Double '//점심휴식시간
    Dim hTime3 As Double '//오후 10분 휴식시간
    Dim hTime4 As Double '//저녁휴식시간
    Dim hTime5 As Double '//야간휴식시간
    
    
    Dim NextDay  As String
    Dim TimeType As String
    Dim sQry     As String
    
    Dim STime As String
    Dim ETime As String
    
    Dim EarlyTo  As Double
    Dim SEarlyTo As Double
    Dim Base     As Double
    Dim Special  As Double
    Dim Extend   As Double
    Dim SpExtend As Double
    Dim Midnight As Double
    
    Dim WorkType As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    ShiftDat = Trim(oDS_PH_PY677B.GetValue("U_ColReg07", pRow - 1)) '근무형태
    GNMUJO = Trim(oDS_PH_PY677B.GetValue("U_ColReg08", pRow - 1)) '근무조
    
    PosDate = Trim(oDS_PH_PY677B.GetValue("U_ColDt05", pRow - 1)) '일자
    GetDate = Trim(oDS_PH_PY677B.GetValue("U_ColDt01", pRow - 1)) '출근일자
    OffDate = Trim(oDS_PH_PY677B.GetValue("U_ColDt02", pRow - 1)) '퇴근일자
    DayOff = Trim(oDS_PH_PY677B.GetValue("U_ColReg10", pRow - 1)) '휴일 평일 구분
    GetTime = Trim(CStr(oDS_PH_PY677B.GetValue("U_ColTm01", pRow - 1))) '출근시각
    OffTime = Trim(CStr(oDS_PH_PY677B.GetValue("U_ColTm02", pRow - 1))) '퇴근시각
    
    GetTime = "0" + GetTime
    GetTime = Right(GetTime, 4)
    
    OffTime = "0" + OffTime
    OffTime = Right(OffTime, 4)
    
    WorkType = Trim(CStr(oDS_PH_PY677B.GetValue("U_ColReg21", pRow - 1))) '근태구분
    
    sQry = "        SELECT   U_TimeType, "
    sQry = sQry & "          U_FromTime, "
    sQry = sQry & "          U_ToTime, "
    sQry = sQry & "          U_NextDay "
    sQry = sQry & " FROM     [@PH_PY002A] a "
    sQry = sQry & "          INNER JOIN "
    sQry = sQry & "          [@PH_PY002B] b "
    sQry = sQry & "             On a.Code = b.Code "
    sQry = sQry & " WHERE    a.U_CLTCOD = '" & CLTCOD & "'" '//사업부
    sQry = sQry & "          AND a.U_SType = '" & ShiftDat & "'" '//교대
    sQry = sQry & "          AND a.U_Shift = '" & GNMUJO & "'"  '//조
    sQry = sQry & "          AND b.U_DayType = '" & DayOff & "'" '//평일
  
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
'        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY677_Time_Calc_Main_Exit
    End If
    
    For i = 0 To oRecordSet.RecordCount - 1
        FromTime = CStr(oRecordSet.Fields(1).VALUE)
        FromTime = "0000" + FromTime
        FromTime = Right(FromTime, 4)
        
        ToTime = CStr(oRecordSet.Fields(2).VALUE)
        ToTime = "0000" + ToTime
        ToTime = Right(ToTime, 4)
        
        NextDay = Trim(oRecordSet.Fields(3).VALUE)
        TimeType = Trim(oRecordSet.Fields(0).VALUE)
        
        If NextDay = "" Then
            NextDay = "N"
        End If
        
        If NextDay = "N" Then
            If ToTime = "0000" Then
                ToTime = "2400"
            End If
        End If
        
        Select Case TimeType
            Case "40" '//조출
                If GetDate = PosDate Then '//출근당일이면
                    If GetTime < ToTime Then '//출근시간 < 기준종료시간
                        STime = GetTime
                        ETime = ToTime
                        
                        If DayOff = "1" Then
                            EarlyTo = PH_PY677_Time_Calc(STime, ETime)
                        Else
                            SEarlyTo = PH_PY677_Time_Calc(STime, ETime)
                        End If
                        
                    End If
                End If
            Case "10", "50" '//정상근무시간
                If (GNMUJO = "11" Or GNMUJO = "21") Then
                    Select Case NextDay
                        Case "N" '//당일
                            If PosDate = GetDate Then '//기준일자 = 출근일자
                                '//1교대1조, 2교대 1조당일
                                If GetTime < FromTime Then '//출근시간 < 시작시간
                                    STime = FromTime       '//시작시간
                                Else
                                    STime = GetTime        '//출근시간
                                End If
                                    
                                If GetDate <> OffDate Then '//퇴근일이 틀릴때(다음날 퇴근일때)
                                    ETime = ToTime  '//종료시간
                                Else
                                    If OffTime < ToTime Then '//퇴근시간 < 종료시간
                                        ETime = OffTime      '//퇴근시간
                                    Else
                                        ETime = ToTime       '//종료시간
                                    End If
                                End If
                                
                                If DayOff = 1 Then
                                    Base = Base + PH_PY677_Time_Calc(STime, ETime) '//평일
                                Else
                                    Special = Special + PH_PY677_Time_Calc(STime, ETime) '//휴일
                                End If
                            End If
                        Case "Y"  '//익일
'                            If PosDate = OffDate Then '//기준일자와 퇴근일자가 같으면 계산안함(당일퇴근)
'                                STime = "0000"
'                                ETime = "0000"
'                            Else                      '//익일퇴근일때
'                                If OffTime < ToTime Then  '//퇴근시간 < 종료시간
'                                    STime = "0000"        '//00시 부터
'                                    ETime = OffTime       '//퇴근시간
'                                Else
'                                    STime = "0000"
'                                    ETime = ToTime
'                                End If
'                            End If
'
'                            If DayOff = 1 Then
'                                Base = Base + PH_PY677_Time_Calc(STime, ETime)
'                            Else
'                                Special = Special + PH_PY677_Time_Calc(STime, ETime)
'                            End If
                        End Select
                Else
                    If GNMUJO = "22" Then '//당일 익일 더해야 한다.
                        Select Case NextDay
                            Case "N" '//당일
                                If PosDate = GetDate Then '//기준일 = 출근일
                                    If GetTime < FromTime Then  '//출근시간 < 시작시간
                                        STime = FromTime        '//시작시간
                                    Else
                                        STime = GetTime         '//출근시간
                                    End If
                                    
                                    If GetDate = OffDate Then '//퇴근일이 같으면(기준일 퇴근)
                                        If OffTime < "2400" Then '//퇴근시간 < 24시
                                            ETime = OffTime         '//퇴근시간
                                        Else
                                            ETime = "2400"
                                        End If
                                    Else
                                        ETime = "2400"
                                    End If
                                    
                                    If DayOff = 1 Then
                                        Base = Base + PH_PY677_Time_Calc(STime, ETime)
                                    Else
                                        Special = Special + PH_PY677_Time_Calc(STime, ETime)
                                    End If
                                End If
                                
                            Case "Y" '//익일
                                If PosDate = OffDate Then '//기준일이 같으면 계산안함
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    STime = "0000"      '//시작시간은 00시
                                    
                                    If OffTime < ToTime Then '//퇴근시간 < 종료시간
                                        ETime = OffTime        '//퇴근시간
                                    Else
                                        ETime = ToTime          '//종료시간
                                    End If
                                    
                                End If
                                
                                If DayOff = 1 Then
                                    Base = Base + PH_PY677_Time_Calc(STime, ETime)
                                Else
                                    Special = Special + PH_PY677_Time_Calc(STime, ETime)
                                End If
                                    
                        End Select
                        
                    End If
                    
                End If
            Case "65", "66", "15" '//오전, 오후 휴식시간, 점심시간
                If (GNMUJO = "11" Or GNMUJO = "21" Or GNMUJO = "22") Then '//1교대1조 2교대 1조
                    Select Case NextDay
                        Case "N" '//당일
                            If PosDate <> GetDate Then '//당일 출근이 아니면
                                STime = "0000"
                                ETime = "0000"
                            Else
                                If PosDate = OffDate Then '//당일퇴근
                                    If GetTime < FromTime Then  '//출근시간 < 시작시간
                                        If FromTime <= OffTime Then '//시작시간 <= 퇴근시간
                                            STime = FromTime  '//시작시간
                                            If ToTime < OffTime Then    '//종료시간 < 퇴근시간
                                                ETime = ToTime
                                            Else
                                                ETime = OffTime
                                            End If
                                        Else
                                            STime = "0000"
                                            ETime = "0000"
                                        End If
                                    Else
                                        If GetTime > ToTime Then    '//출근시간 > 종료시간
                                            STime = "0000"
                                            ETime = "0000"
                                        Else
                                            STime = GetTime    '//출근시간
                                            If ToTime < OffTime Then    '//종료시간 < 퇴근시간
                                                ETime = ToTime          '//종료시간
                                            Else
                                                ETime = OffTime         '//퇴근시간
                                            End If
                                        End If
                                    End If
                                Else '//다음날 퇴근
                                    If ToTime < GetTime Then '//종료시간 < 출근시간
                                        STime = "0000"
                                        ETime = "0000"
                                    Else
                                        If GetTime < FromTime Then   '//출근시간 < 시작시간
                                            STime = FromTime
                                            ETime = ToTime
                                            
                                        Else        '//시작시간이후 출근
                                            STime = GetTime
                                        End If
                                        ETime = ToTime
                                    End If
                                                
                                End If
                            End If
                        Case "Y"    '//익일
                            If PosDate = OffDate Then '//기준일 = 퇴근일(당일퇴근)
                                STime = "0000"
                                ETime = "0000"
                            Else
                                If OffTime < FromTime Then '//퇴근시간 < 시작시간
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    STime = FromTime
                                    If OffTime < ToTime Then '//퇴근시간 < 종료시간
                                        ETime = OffTime
                                    Else
                                        ETime = ToTime
                                    End If
                                End If
                            End If
                    End Select

                End If
                        
                hTime1 = PH_PY677_Time_Calc(STime, ETime) '//오전휴식시간

                If GNMUJO = "22" And TimeType = "65" Then '//야간조(2교대) 오전휴식일경우
                    Midnight = Midnight - hTime1
                Else
                    If DayOff = "1" Then
                        Base = Base - hTime1
                    Else
                        Special = Special - hTime1
                    End If
                End If
                
            Case "20", "60" '//연장근무
                If (GNMUJO = "11" Or GNMUJO = "21") Then '//1교대1조 2교대 1조
                    Select Case NextDay
                        Case "N" '//당일
                            If PosDate <> OffDate Then '// 출근일 <> 퇴근일
                                If GetTime < FromTime Then '//출근시간 < 시작시간
                                    STime = FromTime    '//
                                Else
                                    STime = GetTime
                                End If
                                ETime = "2400"
                            Else                            '//당일 퇴근
                                If OffTime < FromTime Then  '//퇴근시간 < 시작시간
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    STime = FromTime    '//종료시간
                                    ETime = OffTime     '//퇴근시간
                                End If
                            End If
                            If DayOff = "1" Then
                                Extend = Extend + PH_PY677_Time_Calc(STime, ETime)
                            Else
                                SpExtend = SpExtend + PH_PY677_Time_Calc(STime, ETime)
                            End If
                        Case "Y" '//익일
                            If PosDate <> OffDate Then  '//기준일 <> 퇴근일
                                STime = "0000"
                                If OffTime < ToTime Then    '//퇴근시간 < 종료시간
                                    ETime = OffTime
                                Else
                                    ETime = ToTime
                                End If
                                
                            Else
                                STime = "0000"
                                ETime = "0000"
                            End If
                            If DayOff = "1" Then
                                Extend = Extend + PH_PY677_Time_Calc(STime, ETime)
                            Else
                                SpExtend = SpExtend + PH_PY677_Time_Calc(STime, ETime)
                            End If
                            
                    End Select
                ElseIf GNMUJO = "22" Then '//2교대 2조
                    Select Case NextDay
                        Case "N" '//당일
                            '//연장근무없다
                        Case "Y" '//익일
                            If PosDate <> OffDate Then '// 기준일 <> 퇴근일
                                If OffTime < FromTime Then  '//퇴근시간 < 시작시간
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    STime = FromTime
                                    ETime = OffTime
                                    
                                End If
                            Else
                                STime = "0000"
                                ETime = "0000"
                            End If
                            
                            If DayOff = "1" Then
                                Extend = Extend + PH_PY677_Time_Calc(STime, ETime)
                            Else
                                SpExtend = SpExtend + PH_PY677_Time_Calc(STime, ETime)
                            End If
                    End Select
                End If
            Case "25" '//저녁휴식
            Case "30" '//심야시간
                If (GNMUJO = "11" Or GNMUJO = "21" Or GNMUJO = "22") Then '//1교대1조 2교대 1조
                    Select Case NextDay
                        Case "N" '//당일
                            If PosDate <> OffDate Then '//기준일자 <> 퇴근일자
                                If GetTime < FromTime Then '//출근시간 < 시작시간
                                    STime = FromTime    '//시작시간
                                Else
                                    STime = GetTime     '//출근시간
                                End If
                                ETime = "2400"
                            Else                            '//당일퇴근시
                                If OffTime < FromTime Then  '//퇴근시간 < 시작시간
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    If GetTime < FromTime Then '//출근시간 < 시작시간
                                        STime = FromTime            '//시작시간
                                    Else
                                        STime = GetTime
                                    End If
                                    
                                    If OffTime < ToTime Then '//퇴근시간 < 종료시간
                                        ETime = OffTime
                                    Else
                                        ETime = ToTime '//종료시간
                                    End If
                                    
                                End If
                            End If
                            Midnight = Midnight + PH_PY677_Time_Calc(STime, ETime)
                        Case "Y" '//익일
                            If PosDate <> OffDate Then '//출근일 <> 퇴근일
                                STime = "0000"
                                If OffTime < ToTime Then    '//퇴근시간 < 종료시간
                                    ETime = OffTime         '//퇴근시간
                                Else
                                    ETime = ToTime          '//종료시간
                                End If
                            Else
                                STime = "0000"
                                ETime = "0000"
                            End If
                            Midnight = Midnight + PH_PY677_Time_Calc(STime, ETime)
                    End Select

                End If
            Case "35" '//야간휴식 수정해야함
                '//hTime5
                Select Case NextDay
                    Case "N" '//당일
                        If PosDate = OffDate Then '//기준일자 = 퇴근일자
                            If OffTime < FromTime Then '//퇴근시간 < 시작시간
                                STime = "0000"
                                ETime = "0000"
                            Else
                                If GetTime < FromTime Then
                                    STime = FromTime
                                Else
                                    STime = GetTime
                                End If

                                If OffTime < ToTime Then
                                    ETime = OffTime
                                Else
                                    ETime = ToTime
                                End If
                            End If
                        Else
                            '//다음날 퇴근
                            If GetTime < FromTime Then '//출근시간 < 시작시간
                                STime = FromTime
                            Else
                                STime = GetTime
                            End If

                            ETime = ToTime
                        End If

                    Case "Y"
                        If PosDate = OffDate Then '//기준일자 = 퇴근일자
                            STime = "0000"
                            ETime = "0000"
                        Else
                            If OffTime < FromTime Then '//퇴근시간 < 시작시간
                                STime = "0000"
                                ETime = "0000"
                            Else
                                STime = FromTime
                                If OffTime < ToTime Then
                                    ETime = OffTime
                                Else
                                    ETime = ToTime
                                End If
                            End If
                        End If
                End Select

                hTime5 = PH_PY677_Time_Calc(STime, ETime)

                If DayOff = "1" Then '//평일
                    If GNMUJO = "22" Then '//2교대 2조(야간조)
                        Base = Base - hTime5 '//기본근무
                        Midnight = Midnight - hTime5 '//심야시간에서 차감
                    Else '//그외 주간조
                        Extend = Extend - hTime5 '//연장근무에서 차감
                        Midnight = Midnight - hTime5 '//심야시간에서 차감
                    End If
                Else
                    '//휴일
                    If GNMUJO = "22" Then
                        Special = Special - hTime5
                        Midnight = Midnight - hTime5 '//심야시간에서 차감
                    Else
                        '//다음날 퇴근은 연장근무임
                        SpExtend = SpExtend - hTime5 '//연장근무에서 차감
                        Midnight = Midnight - hTime5 '//심야시간에서 차감
                    End If
                End If
            
        End Select
        oRecordSet.MoveNext
    Next
    
    Call oDS_PH_PY677B.setValue("U_ColQty06", pRow - 1, PH_PY677_hhmm_Calc(EarlyTo)) '조출
    Call oDS_PH_PY677B.setValue("U_ColQty01", pRow - 1, PH_PY677_hhmm_Calc(Base, WorkType)) '기본
    Call oDS_PH_PY677B.setValue("U_ColQty02", pRow - 1, PH_PY677_hhmm_Calc(Extend)) '연장
    Call oDS_PH_PY677B.setValue("U_ColQty03", pRow - 1, PH_PY677_hhmm_Calc(Midnight)) '심야
    
    Call oDS_PH_PY677B.setValue("U_ColQty07", pRow - 1, PH_PY677_hhmm_Calc(SEarlyTo)) '휴조
    Call oDS_PH_PY677B.setValue("U_ColQty04", pRow - 1, PH_PY677_hhmm_Calc(Special)) '특근
    Call oDS_PH_PY677B.setValue("U_ColQty05", pRow - 1, PH_PY677_hhmm_Calc(SpExtend)) '특근연장
       
'    oForm.Items("EarlyTo").Specific.VALUE = PH_PY677_hhmm_Calc(EarlyTo) '//조출
'    oForm.Items("Base").Specific.VALUE = PH_PY677_hhmm_Calc(Base) '//기본
'    oForm.Items("Extend").Specific.VALUE = PH_PY677_hhmm_Calc(Extend) '//연장
'    oForm.Items("Midnight").Specific.VALUE = PH_PY677_hhmm_Calc(Midnight) '//심야
'
'    oForm.Items("SEarlyTo").Specific.VALUE = PH_PY677_hhmm_Calc(SEarlyTo) '//특근조출
'    oForm.Items("Special").Specific.VALUE = PH_PY677_hhmm_Calc(Special) '//특근
'    oForm.Items("SpExtend").Specific.VALUE = PH_PY677_hhmm_Calc(SpExtend) '//특근연장
    
    Set oRecordSet = Nothing
    Exit Function
PH_PY677_Time_Calc_Main_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY677_Time_Calc_Main_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY677_Time_Calc_Main_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY677_hhmm_Calc(ByVal mTime As Double, Optional ByVal pWorkType As String) As Double
    
    Dim hh As Long
    Dim MM As Double
    
    Dim Ret As Double
    
    
    hh = Int((mTime) / 3600)
    MM = ((mTime) Mod 3600) / 60

    If MM > 0 Then
        If MM > 30 Then
            MM = 1
        Else
            MM = 0.5
        End If
    ElseIf MM = 0 Then
        MM = 0
    Else
        If MM < -30 Then
            MM = -1
        Else
            MM = -0.5
        End If
    End If

    Ret = hh + MM
    '근태구분이 반차일 경우 무조건 4시간 반환(2014.04.21 송명규 추가) 별로 안 좋은 방법인데...
    If pWorkType = "D09" Or pWorkType = "D10" Then
        Ret = 4
    End If
        
    If IsNull(Ret) Then
        Ret = 0
    End If
    
    PH_PY677_hhmm_Calc = Ret
    
    Exit Function
PH_PY677_hhmm_Calc_Exit:
    Exit Function
PH_PY677_hhmm_Calc_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY677_hhmm_Calc_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY677_Time_Calc(ByVal GetTime As String, ByVal OffTime As String) As Double
    
    Dim STime As Long
    Dim ETime As Long
    
    Dim hh As Long
    Dim MM As Double
    
    Dim Ret As Double
    
    
    STime = Int(Mid(GetTime, 1, 2)) * 3600 + Int(Mid(GetTime, 3, 2)) * 60
    
    ETime = Int(Mid(OffTime, 1, 2)) * 3600 + Int(Mid(OffTime, 3, 2)) * 60
    
'    hh = Int((ETime - STime) / 3600)
'    mm = ((ETime - STime) Mod 3600) / 60
'
'    If mm > 0 Then
'        If mm > 30 Then
'            mm = 1
'        Else
'            mm = 0.5
'        End If
'    ElseIf mm = 0 Then
'        mm = 0
'    Else
'        If mm < -30 Then
'            mm = -1
'        Else
'            mm = -0.5
'        End If
'    End If
'
'    Ret = hh + mm
    
    Ret = ETime - STime
    
    If IsNull(Ret) Then
        Ret = 0
    End If
    
    PH_PY677_Time_Calc = Ret
    
    Exit Function
PH_PY677_Time_Calct_Exit:
    Exit Function
PH_PY677_Mark_Set_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY677_Time_Calc_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY677_CheckAll()
On Error GoTo PH_PY677_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PH_PY677B.GetValue("U_ColReg20", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PH_PY677B.Offset = loopCount
        If CheckType = "N" Then
            Call oDS_PH_PY677B.setValue("U_ColReg20", loopCount, "Y")
        Else
            Call oDS_PH_PY677B.setValue("U_ColReg20", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm.Freeze(False)

    Exit Sub
    
PH_PY677_CheckAll_Error:
    Call oForm.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PH_PY677_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Sub PH_PY677_ConfirmAll()
On Error GoTo PH_PY677_ConfirmAll_Error

    Dim ConfirmType As String
    Dim loopCount As Integer
    
    Call oForm.Freeze(True)
    ConfirmType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PH_PY677B.GetValue("U_ColReg16", loopCount)) = "N" Then
            ConfirmType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PH_PY677B.Offset = loopCount
        If ConfirmType = "N" Then
            Call oDS_PH_PY677B.setValue("U_ColReg16", loopCount, "Y")
        Else
            Call oDS_PH_PY677B.setValue("U_ColReg16", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm.Freeze(False)

    Exit Sub
    
PH_PY677_ConfirmAll_Error:
    Call oForm.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PH_PY677_ConfirmAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

'Private Sub PH_PY677_Print_Report01()
'
'    Dim DocNum As String
'    Dim ErrNum As Integer
'    Dim WinTitle As String
'    Dim ReportName As String
'    Dim sQry As String
'    Dim oRecordSet As SAPbobsCOM.Recordset
'
'    On Error GoTo PH_PY677_Print_Report01_Error
'
'    Dim CLTCOD As String
'    Dim DestNo1 As String
'    Dim DestNo2 As String
'
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'     '/ ODBC 연결 체크
'    If ConnectODBC = False Then
'        GoTo PH_PY677_Print_Report01_Error
'    End If
'
'    '//인자 MOVE , Trim 시키기..
'    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
'    DestNo1 = Trim(oForm.Items("DestNo1").Specific.VALUE)
'    DestNo2 = Trim(oForm.Items("DestNo2").Specific.VALUE)
'
'    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'
'    WinTitle = "[PH_PY677] 공용증"
'
'    If CLTCOD = "1" Then '창원
'        ReportName = "PH_PY677_01.rpt"
'    ElseIf CLTCOD = "2" Then '동래
'        ReportName = "PH_PY677_02.rpt"
'    ElseIf CLTCOD = "3" Then '사상
'        ReportName = "PH_PY677_03.rpt"
'    End If
'
'    '// Formula 수식필드
'    ReDim gRpt_Formula(2)
'    ReDim gRpt_Formula_Value(2)
'
'    '// SubReport
'    ReDim gRpt_SRptSqry(1)
'    ReDim gRpt_SRptName(1)
'
'    ReDim gRpt_SFormula(1, 1)
'    ReDim gRpt_SFormula_Value(1, 1)
'
'    gRpt_SFormula(1, 1) = ""
'    gRpt_SFormula_Value(1, 1) = ""
'
'    '/ Procedure 실행"
'    sQry = "EXEC [PH_PY677_90] '" & CLTCOD & "','" & DestNo1 & "','" & DestNo2 & "'"
'
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY677_Print_Report01_Error
'    End If
'
'    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
'        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
'    End If
'
'    Set oRecordSet = Nothing
'    Exit Sub
'
'PH_PY677_Print_Report01_Error:
'    If ErrNum = 1 Then
'        Set oRecordSet = Nothing
'        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
'    Else
'    Set oRecordSet = Nothing
'    Sbo_Application.SetStatusBarMessage "PH_PY677_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    End If
'
'End Sub





