VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY508"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PH_PY508.cls
'//  Module : 인사관리>인사리포트
'//  Desc : 재직증명 등록 및 발급
'//  FormType : PH_PY508
'//  Create Date(Start) : 2015.05.14
'//  Create Date(End) : 2015.05.18
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PH_PY508A As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PH_PY508B As SAPbouiCOM.DBDataSource '등록라인

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
    check As Boolean
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
    Dim oInnerXml As String
    Dim oXmlDoc As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY508.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY508_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01, "PH_PY508" '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    '//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm.Freeze True
    
    Call PH_PY508_CreateItems
    Call PH_PY508_FormItemEnabled '콤보박스 세팅 때문에 실행
    Call PH_PY508_ComboBox_Setting
    Call PH_PY508_CF_ChooseFromList
    Call PH_PY508_EnableMenus
    Call PH_PY508_SetDocument(oFromDocEntry01)
    Call PH_PY508_FormResize
    
'    Call PH_PY508_Add_MatrixRow(0, True)
    Call PH_PY508_LoadCaption
    
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
        
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY508A]"
    Call RecordSet01.DoQuery(sQry)
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PH_PY508A.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PH_PY508A.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    oMat01.Columns("CLTAdrs").Visible = False '사업장주소 Visible False
    oMat01.Columns("RepName").Visible = False '대표이사명 Visible False
    
    Call PH_PY508_FormReset '폼초기화 추가(2013.01.29 송명규)
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing

    '발행일 기간
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

Private Sub PH_PY508_LoadCaption()
'******************************************************************************
'Function ID : PH_PY508_LoadCaption()
'해당모듈 : PH_PY508
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY508_LoadCaption_Error

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
PH_PY508_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PH_PY508_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PH_PY508_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PH_PY508B.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PH_PY508B.Offset = oRow
    oDS_PH_PY508B.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY508_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PH_PY508_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PH_PY508_MTX01()
'******************************************************************************
'Function ID : PH_PY508_MTX01()
'해당모듈 : PH_PY508
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY508_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim sDocEntry As String '관리번호
    Dim sCLTCOD As String '사업장
    Dim SDocNo1 As String '문서번호1
    Dim SDocNo2 As String '문서번호2
    Dim SFrDate As String '발행일(시작)
    Dim SToDate As String '발행일(종료)
    Dim sMSTCOD As String '사번
    Dim sTeamCode As String '소속팀
    Dim sRspCode As String '소속담당
    Dim SUseCmt As String '용도
    
    sDocEntry = Trim(oForm.Items("SDocEntry").Specific.VALUE) '관리번호
    sCLTCOD = Trim(oForm.Items("SCLTCOD").Specific.VALUE) '사업장
    SDocNo1 = Trim(oForm.Items("SDocNo1").Specific.VALUE) '문서번호1
    SDocNo2 = Trim(oForm.Items("SDocNo2").Specific.VALUE) '문서번호2
    SFrDate = Trim(oForm.Items("SFrDate").Specific.VALUE) '발행일(시작)
    SToDate = Trim(oForm.Items("SToDate").Specific.VALUE) '발행일(종료)
    sMSTCOD = Trim(oForm.Items("SMSTCOD").Specific.VALUE) '사번
    sTeamCode = Trim(oForm.Items("STeamCode").Specific.VALUE) '소속팀
    sRspCode = Trim(oForm.Items("SRspCode").Specific.VALUE) '소속담당
    SUseCmt = Trim(oForm.Items("SUseCmt").Specific.VALUE) '용도
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "                EXEC [PH_PY508_01] "
    sQry = sQry & "'" & sDocEntry & "'," '관리번호
    sQry = sQry & "'" & sCLTCOD & "'," '사업장
    sQry = sQry & "'" & SDocNo1 & "'," '문서번호1
    sQry = sQry & "'" & SDocNo2 & "'," '문서번호2
    sQry = sQry & "'" & SFrDate & "'," '발행일(시작)
    sQry = sQry & "'" & SToDate & "'," '발행일(종료)
    sQry = sQry & "'" & sMSTCOD & "',"  '사원번호
    sQry = sQry & "'" & sTeamCode & "'," '소속팀
    sQry = sQry & "'" & sRspCode & "'," '소속담당
    sQry = sQry & "'" & SUseCmt & "'" '용도
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PH_PY508B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm.Mode = fm_ADD_MODE
        
'        Call PH_PY508_Add_MatrixRow(0, True)
        Call PH_PY508_LoadCaption
        
        GoTo PH_PY508_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PH_PY508B.Size Then
            oDS_PH_PY508B.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PH_PY508B.Offset = i
        
        Call oDS_PH_PY508B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY508B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PH_PY508B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PH_PY508B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("CLTCOD").VALUE)) '사업장
        Call oDS_PH_PY508B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("DocNo1").VALUE)) '문서번호1
        Call oDS_PH_PY508B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("DocNo2").VALUE)) '문서번호2
        Call oDS_PH_PY508B.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("CLTAdrs").VALUE)) '사업장주소
        Call oDS_PH_PY508B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("RepName").VALUE)) '대표이사명
        Call oDS_PH_PY508B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("RegDt").VALUE)) '발행일
        Call oDS_PH_PY508B.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("MSTCOD").VALUE)) '사번
        Call oDS_PH_PY508B.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("MSTNAM").VALUE)) '성명
        Call oDS_PH_PY508B.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("TeamCode").VALUE)) '소속팀
        Call oDS_PH_PY508B.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("TeamName").VALUE)) '소속팀명
        Call oDS_PH_PY508B.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("RspCode").VALUE)) '소속담당
        Call oDS_PH_PY508B.setValue("U_ColReg14", i, Trim(oRecordSet01.Fields("RspName").VALUE)) '소속담당명
        Call oDS_PH_PY508B.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("Position").VALUE)) '직책
        Call oDS_PH_PY508B.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("GovID").VALUE)) '주민등록번호
        Call oDS_PH_PY508B.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("CurAdrs").VALUE)) '현주소
        Call oDS_PH_PY508B.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("GrpDat").VALUE)) '입사일
        Call oDS_PH_PY508B.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("UseCmt").VALUE)) '용도
        Call oDS_PH_PY508B.setValue("U_ColReg20", i, Trim(oRecordSet01.Fields("Retire").VALUE)) '퇴사여부
        
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
PH_PY508_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY508_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PH_PY508_DeleteData()
'******************************************************************************
'Function ID : PH_PY508_DeleteData()
'해당모듈 : PH_PY508
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY508_DeleteData_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PH_PY508A] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm.Mode = fm_ADD_MODE
            GoTo PH_PY508_DeleteData_Error
        Else
            sQry = "EXEC PH_PY508_04 '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
'    Call PH_PY508_FormReset
    
'    oForm.Mode = fm_ADD_MODE
    
'    Call oForm.Items("BtnSearch").Click(ct_Regular)
    
'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PH_PY508_Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY508_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY508_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PH_PY508_UpdateData() As Boolean
'******************************************************************************
'Function ID : PH_PY508_UpdateData()
'해당모듈 : PH_PY508
'기능 : 기본정보를 수정
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY508_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer '관리번호
    Dim DocNo1 As String '문서번호1
    Dim DocNo2 As String '문서번호2
    Dim CLTCOD As String '사업장
    Dim CLTAdrs As String '사업장주소
    Dim RepName As String '대표이사명
    Dim RegDt As String '발행일
    Dim MSTCOD As String '사번
    Dim MSTNAM As String '성명
    Dim TeamCode As String '소속팀(코드)
    Dim TeamName As String '소속팀명
    Dim RspCode As String '소속담당(코드)
    Dim RspName As String '소속담당명
    Dim Position As String '직책
    Dim GovID As String '주민등록번호
    Dim CurAdrs As String '현주소
    Dim GrpDat As String '입사일
    Dim UseCmt As String '용도
    Dim Retire As String '퇴사여부
    
    DocEntry = Trim(oForm.Items("DocEntry").Specific.VALUE) '관리번호
    DocNo1 = Trim(oForm.Items("DocNo1").Specific.VALUE)  '문서번호1
    DocNo2 = Trim(oForm.Items("DocNo2").Specific.VALUE)  '문서번호2
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)  '사업장
    CLTAdrs = Trim(oForm.Items("CLTAdrs").Specific.VALUE)  '사업장주소
    RepName = Trim(oForm.Items("RepName").Specific.VALUE) '대표이사명
    RegDt = Trim(oForm.Items("RegDt").Specific.VALUE) '발행일
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사번
    MSTNAM = Trim(oForm.Items("MSTNAM").Specific.VALUE) '성명
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '소속팀(코드)
'    TeamName = Trim(oForm.Items("TeamName").Specific.VALUE) '소속팀명
    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE)  '소속담당(코드)
'    RspName = Trim(oForm.Items("DocEntry").Specific.VALUE)  '소속담당명
    Position = Trim(oForm.Items("Position").Specific.VALUE) '직책
    GovID = Trim(oForm.Items("GovID").Specific.VALUE) '주민등록번호
    CurAdrs = Trim(oForm.Items("CurAdrs").Specific.VALUE) '현주소
    GrpDat = Trim(oForm.Items("GrpDat").Specific.VALUE) '입사일
    UseCmt = Trim(oForm.Items("UseCmt").Specific.VALUE) '용도
    Retire = IIf(oForm.Items("Retire").Specific.Checked = True, "1", "0") '퇴사여부

    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
        PH_PY508_UpdateData = False
        Exit Function
    End If
    
    sQry = "                EXEC [PH_PY508_03] "
    sQry = sQry & "'" & DocEntry & "'," '관리번호
    sQry = sQry & "'" & DocNo1 & "'," '문서번호1
    sQry = sQry & "'" & DocNo2 & "'," '문서번호2
    sQry = sQry & "'" & CLTCOD & "'," '사업장
    sQry = sQry & "'" & CLTAdrs & "'," '사업장주소
    sQry = sQry & "'" & RepName & "'," '대표이사명
    sQry = sQry & "'" & RegDt & "'," '발행일
    sQry = sQry & "'" & MSTCOD & "',"  '사원번호
    sQry = sQry & "'" & MSTNAM & "',"  '사원성명
    sQry = sQry & "'" & TeamCode & "'," '소속팀(코드)
    sQry = sQry & "'" & RspCode & "'," '소속담당(코드)
    sQry = sQry & "'" & Position & "'," '직책
    sQry = sQry & "'" & GovID & "'," '주민등록번호
    sQry = sQry & "'" & CurAdrs & "'," '현주소
    sQry = sQry & "'" & GrpDat & "'," '입사일
    sQry = sQry & "'" & UseCmt & "'," '용도
    sQry = sQry & "'" & Retire & "'" '퇴사여부
    
    Call RecordSet01.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PH_PY508_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY508_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY508_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PH_PY508_AddData() As Boolean
'******************************************************************************
'Function ID : PH_PY508_AddData()
'해당모듈 : PH_PY508
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY508_AddData_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer '관리번호
    Dim DocNo1 As String '문서번호1
    Dim DocNo2 As String '문서번호2
    Dim CLTCOD As String '사업장
    Dim CLTAdrs As String '사업장주소
    Dim RepName As String '대표이사명
    Dim RegDt As String '발행일
    Dim MSTCOD As String '사번
    Dim MSTNAM As String '성명
    Dim TeamCode As String '소속팀(코드)
    Dim TeamName As String '소속팀명
    Dim RspCode As String '소속담당(코드)
    Dim RspName As String '소속담당명
    Dim Position As String '직책
    Dim GovID As String '주민등록번호
    Dim CurAdrs As String '현주소
    Dim GrpDat As String '입사일
    Dim UseCmt As String '용도
    Dim Retire As String '퇴사여부
    Dim UserSign As String 'UserSign

    DocNo1 = Trim(oForm.Items("DocNo1").Specific.VALUE) '문서번호1
    DocNo2 = Trim(oForm.Items("DocNo2").Specific.VALUE) '문서번호2
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    CLTAdrs = Trim(oForm.Items("CLTAdrs").Specific.VALUE) '사업장주소
    RepName = Trim(oForm.Items("RepName").Specific.VALUE) '대표이사명
    RegDt = Trim(oForm.Items("RegDt").Specific.VALUE) '발행일
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사번
    MSTNAM = Trim(oForm.Items("MSTNAM").Specific.VALUE) '성명
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '소속팀(코드)
'    TeamName = Trim(oForm.Items("TeamName").Specific.VALUE) '소속팀명
    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE) '소속담당(코드)
'    RspName = Trim(oForm.Items("RspName").Specific.VALUE) '소속담당명
    Position = Trim(oForm.Items("Position").Specific.VALUE) '직책
    GovID = Trim(oForm.Items("GovID").Specific.VALUE) '주민등록번호
    CurAdrs = Trim(oForm.Items("CurAdrs").Specific.VALUE) '현주소
    GrpDat = Trim(oForm.Items("GrpDat").Specific.VALUE) '입사일
    UseCmt = Trim(oForm.Items("UseCmt").Specific.VALUE) '용도
    Retire = IIf(oForm.Items("Retire").Specific.Checked = True, "1", "0") '퇴사여부
    
    UserSign = oCompany.UserSignature
    
    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY508A]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "                EXEC [PH_PY508_02] "
    sQry = sQry & "'" & DocEntry & "'," '관리번호
    sQry = sQry & "'" & DocNo1 & "'," '문서번호1
    sQry = sQry & "'" & DocNo2 & "'," '문서번호2
    sQry = sQry & "'" & CLTCOD & "'," '사업장
    sQry = sQry & "'" & CLTAdrs & "'," '사업장주소
    sQry = sQry & "'" & RepName & "'," '대표이사명
    sQry = sQry & "'" & RegDt & "'," '발행일
    sQry = sQry & "'" & MSTCOD & "',"  '사원번호
    sQry = sQry & "'" & MSTNAM & "',"  '사원성명
    sQry = sQry & "'" & TeamCode & "'," '소속팀(코드)
    sQry = sQry & "'" & RspCode & "'," '소속담당(코드)
    sQry = sQry & "'" & Position & "'," '직책
    sQry = sQry & "'" & GovID & "'," '주민등록번호
    sQry = sQry & "'" & CurAdrs & "'," '현주소
    sQry = sQry & "'" & GrpDat & "'," '입사일
    sQry = sQry & "'" & UseCmt & "'," '용도
    sQry = sQry & "'" & Retire & "'," '퇴사여부
    sQry = sQry & "'" & UserSign & "'"  'UserSign
    
    Call RecordSet02.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PH_PY508_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY508_AddData_Error:
    PH_PY508_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY508_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PH_PY508_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PH_PY508_HeaderSpaceLineDel()
'해당모듈 : PH_PY508
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY508_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm.Items("DocNo1").Specific.VALUE) = "" '문서번호1
            ErrNum = 1
            GoTo PH_PY508_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("DocNo2").Specific.VALUE) = "" '문서번호2
            ErrNum = 2
            GoTo PH_PY508_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("CLTAdrs").Specific.VALUE) = "" '사업장주소
            ErrNum = 3
            GoTo PH_PY508_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("RepName").Specific.VALUE) = "" '대표이사명
            ErrNum = 4
            GoTo PH_PY508_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("RegDt").Specific.VALUE) = "" '발행일
            ErrNum = 5
            GoTo PH_PY508_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("MSTCOD").Specific.VALUE) = ""  '사번
            ErrNum = 6
            GoTo PH_PY508_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("GovID").Specific.VALUE) = "" '주민등록번호
            ErrNum = 7
            GoTo PH_PY508_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("CurAdrs").Specific.VALUE) = "" '현주소
            ErrNum = 8
            GoTo PH_PY508_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("GrpDat").Specific.VALUE) = "" '입사일
            ErrNum = 9
            GoTo PH_PY508_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("UseCmt").Specific.VALUE) = "" '용도
            ErrNum = 10
            GoTo PH_PY508_HeaderSpaceLineDel_Error
    End Select

    PH_PY508_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY508_HeaderSpaceLineDel_Error:

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "문서번호1은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("DocNo1").CLICK(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "문서번호2는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("DocNo2").CLICK(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "사업장주소는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("CLTAdrs").CLICK(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "대표이사명은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("RepName").CLICK(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "발행일은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("RegDt").CLICK(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "사번은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("MSTCOD").CLICK(ct_Regular)
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "주민등록번호는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("GovID").CLICK(ct_Regular)
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "현주소는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("CurAdrs").CLICK(ct_Regular)
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "입사일은 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("GrpDat").CLICK(ct_Regular)
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "용도는 필수사항입니다. 확인하세요.", "E"
        Call oForm.Items("UseCmt").CLICK(ct_Regular)
    End If
    
    PH_PY508_HeaderSpaceLineDel = False
    
End Function

'/ 메트릭스 필수 사항 check
Private Function PH_PY508_MatrixSpaceLineDel() As Boolean
On Error GoTo PH_PY508_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PH_PY508_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY508_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PH_PY508_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PH_PY508_MatrixSpaceLineDel = False
End Function

Private Sub PH_PY508_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String, Optional pFormMode As String)
On Error GoTo PH_PY508_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    
    Dim TeamCode As String
'    Dim CLTCOD As String
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case oUID
    
        Case "RegDt"
        
            If pFormMode <> "2" Then '업데이트 모드에서는 문서번호 재설정 안함
                Call PH_PY508_GetDocNo
            End If

        Case "CLTCOD"

            Call PH_PY508_GetDocNo
            
'            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)

'            '소속팀
'            Call oForm.Items("TeamCode").Specific.ValidValues.Add("%", "선택")
'            sQry = "            SELECT      U_Code AS [Code],"
'            sQry = sQry & "                 U_CodeNm As [Name]"
'            sQry = sQry & "  FROM       [@PS_HR200L]"
'            sQry = sQry & "  WHERE      Code = '1'"
'            sQry = sQry & "                 AND U_UseYN = 'Y'"
'            sQry = sQry & "                 AND U_Char2 = '" & CLTCOD & "'"
'            sQry = sQry & "  ORDER BY  U_Seq"
'            Call MDC_SetMod.Set_ComboList(oForm.Items("TeamCode").Specific, sQry, "", False, False)
'            Call oForm.Items("TeamCode").Specific.Select("%", psk_ByValue)
'
'            '소속담당
'            Call oForm.Items("RspCode").Specific.ValidValues.Add("%", "선택")
'            sQry = "            SELECT      U_Code AS [Code],"
'            sQry = sQry & "                 U_CodeNm As [Name]"
'            sQry = sQry & "  FROM       [@PS_HR200L]"
'            sQry = sQry & "  WHERE      Code = '2'"
'            sQry = sQry & "                 AND U_UseYN = 'Y'"
'            sQry = sQry & "                 AND U_Char2 = '" & CLTCOD & "'"
'            sQry = sQry & "  ORDER BY  U_Seq"
'            Call MDC_SetMod.Set_ComboList(oForm.Items("RspCode").Specific, sQry, "", False, False)
'            Call oForm.Items("RspCode").Specific.Select(0, psk_Index)
            
        Case "SCLTCOD"

            Set oCombo = oForm.Items("STeamCode").Specific

            If oCombo.ValidValues.Count > 0 Then
                For loopCount = oCombo.ValidValues.Count - 1 To 0 Step -1
                    Call oCombo.ValidValues.Remove(loopCount, psk_Index)
                Next
            End If

            Call oCombo.ValidValues.Add("%", "전체")
            sQry = "           SELECT     U_Code,"
            sQry = sQry & "               U_CodeNm"
            sQry = sQry & " FROM      [@PS_HR200L]"
            sQry = sQry & " WHERE     Code = '1'"
            sQry = sQry & "               AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
            sQry = sQry & "               AND U_UseYN = 'Y'"
            sQry = sQry & " ORDER BY U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("STeamCode").Specific, sQry)
            Call oCombo.Select(0, psk_Index)

'            oForm.Items("STeamCode").DisplayDesc = True

        Case "STeamCode"
            
            TeamCode = Trim(oForm.Items("STeamCode").Specific.VALUE)
            
            If oForm.Items("SRspCode").Specific.ValidValues.Count > 0 Then
                For loopCount = oForm.Items("SRspCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("SRspCode").Specific.ValidValues.Remove(loopCount, psk_Index)
                Next
            End If
            
            '담당콤보세팅
            Call oForm.Items("SRspCode").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '2'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char1 = '" & TeamCode & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("SRspCode").Specific, sQry, "", False, False)
            Call oForm.Items("SRspCode").Specific.Select(0, psk_Index)

        Case "MSTCOD"
        
            sQry = "EXEC PH_PY508_06 '"
            sQry = sQry & oForm.Items("CLTCOD").Specific.VALUE & "','"
            sQry = sQry & oForm.Items("MSTCOD").Specific.VALUE & "'"
            
            Call oRecordSet01.DoQuery(sQry)
            
            Call oDS_PH_PY508A.setValue("U_MSTNAM", 0, Trim(oRecordSet01.Fields("MSTNAM").VALUE)) '성명
            Call oDS_PH_PY508A.setValue("U_TeamCode", 0, Trim(oRecordSet01.Fields("TeamCode").VALUE)) '팀코드
            Call oDS_PH_PY508A.setValue("U_RspCode", 0, Trim(oRecordSet01.Fields("RspCode").VALUE)) '소속담당
            Call oDS_PH_PY508A.setValue("U_Position", 0, Trim(oRecordSet01.Fields("Position").VALUE)) '직책
            Call oDS_PH_PY508A.setValue("U_GovID", 0, Trim(oRecordSet01.Fields("GovID").VALUE)) '주민등록번호
            Call oDS_PH_PY508A.setValue("U_CurAdrs", 0, Trim(oRecordSet01.Fields("CurAdrs").VALUE)) '현주소
            Call oDS_PH_PY508A.setValue("U_GrpDat", 0, Format(Trim(oRecordSet01.Fields("GrpDat").VALUE), "yyyyMMdd")) '입사일
            Call oDS_PH_PY508A.setValue("U_CLTAdrs", 0, Trim(oRecordSet01.Fields("CLTAdrs").VALUE)) '사업장주소
            Call oDS_PH_PY508A.setValue("U_RepName", 0, Trim(oRecordSet01.Fields("RepName").VALUE)) '대표이사명
            
            oForm.Items("Retire").Enabled = False
            
        Case "SMSTCOD"
                
            oForm.Items("SMSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & oForm.Items("SMSTCOD").Specific.VALUE & "'") '성명
            
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PH_PY508_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PH_PY508_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PH_PY508_FormItemEnabled()
On Error GoTo PH_PY508_FormItemEnabled_Error
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
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY508_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY508_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                '/추가버튼 클릭시 메트릭스 insertrow
                
                Call PH_PY508_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PH_PY508_LoadCaption
                
                'oForm.Items("GCode").Click ct_Regular
                
                oForm.Items("Retire").Enabled = True
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm.Freeze(True)
                Call PH_PY508_Add_MatrixRow(oMat01.VisualRowCount)
                Call oForm.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PH_PY508_FormItemEnabled '//UDO방식
            Case "1282": '추가
            
'                oMat01.Clear
'                oDS_PH_PY508A.Clear

'                Call PH_PY508_LoadCaption
'                Call PH_PY508_FormItemEnabled
                '//Call PH_PY508_FormItemEnabled '//UDO방식
                '//Call PH_PY508_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY508_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm.Freeze(True)
                Call oDS_PH_PY508B.RemoveRecord(oDS_PH_PY508B.Size - 1)
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
    
        If pval.ItemUID = "PH_PY508" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm.Mode = fm_ADD_MODE Then
            
                If PH_PY508_HeaderSpaceLineDel() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                If PH_PY508_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                If PH_PY508_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PH_PY508_FormReset
                
                oForm.Items("Retire").Enabled = True
                
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY508_LoadCaption
                Call PH_PY508_MTX01
                
                oLast_Mode = oForm.Mode
                
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            
                If PH_PY508_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PH_PY508_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                If PH_PY508_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PH_PY508_FormReset
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY508_LoadCaption
                Call PH_PY508_MTX01
                
'                oForm.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
            Call PH_PY508_FormReset
            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PH_PY508_LoadCaption
            Call PH_PY508_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PH_PY508_DeleteData
                Call PH_PY508_FormReset
                
                oForm.Items("Retire").Enabled = True
                
                oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PH_PY508_LoadCaption
                Call PH_PY508_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            Call PH_PY508_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PH_PY508" Then
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
                Call oDS_PH_PY508A.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
                Call oDS_PH_PY508A.setValue("U_CLTCOD", 0, oMat01.Columns("CLTCOD").Cells(pval.Row).Specific.VALUE) '사업장
                Call oDS_PH_PY508A.setValue("U_DocNo1", 0, oMat01.Columns("DocNo1").Cells(pval.Row).Specific.VALUE) '문서번호1
                Call oDS_PH_PY508A.setValue("U_DocNo2", 0, oMat01.Columns("DocNo2").Cells(pval.Row).Specific.VALUE) '문서번호2
                Call oDS_PH_PY508A.setValue("U_CLTAdrs", 0, oMat01.Columns("CLTAdrs").Cells(pval.Row).Specific.VALUE) '사업장주소
                Call oDS_PH_PY508A.setValue("U_RepName", 0, oMat01.Columns("RepName").Cells(pval.Row).Specific.VALUE) '대표이사명
                Call oDS_PH_PY508A.setValue("U_RegDt", 0, Replace(oMat01.Columns("RegDt").Cells(pval.Row).Specific.VALUE, ".", "")) '발행일
                Call oDS_PH_PY508A.setValue("U_MSTCOD", 0, oMat01.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE) '사번
                Call oDS_PH_PY508A.setValue("U_MSTNAM", 0, oMat01.Columns("MSTNAM").Cells(pval.Row).Specific.VALUE) '성명
                Call oDS_PH_PY508A.setValue("U_TeamCode", 0, oMat01.Columns("TeamCode").Cells(pval.Row).Specific.VALUE) '소속팀
'                Call oDS_PH_PY508A.setValue("U_ColReg12", 0, oMat01.Columns("TeamName").Cells(pval.Row).Specific.VALUE) '소속팀명
                Call oDS_PH_PY508A.setValue("U_RspCode", 0, oMat01.Columns("RspCode").Cells(pval.Row).Specific.VALUE) '소속담당
'                Call oDS_PH_PY508A.setValue("U_ColReg14", 0, oMat01.Columns("RspName").Cells(pval.Row).Specific.VALUE) '소속담당명
                Call oDS_PH_PY508A.setValue("U_Position", 0, oMat01.Columns("Position").Cells(pval.Row).Specific.VALUE) '직책
                Call oDS_PH_PY508A.setValue("U_GovID", 0, oMat01.Columns("GovID").Cells(pval.Row).Specific.VALUE) '주민등록번호
                Call oDS_PH_PY508A.setValue("U_CurAdrs", 0, oMat01.Columns("CurAdrs").Cells(pval.Row).Specific.VALUE) '현주소
                Call oDS_PH_PY508A.setValue("U_GrpDat", 0, Replace(oMat01.Columns("GrpDat").Cells(pval.Row).Specific.VALUE, ".", "")) '입사일
                Call oDS_PH_PY508A.setValue("U_UseCmt", 0, oMat01.Columns("UseCmt").Cells(pval.Row).Specific.VALUE) '용도
                'Call oDS_PH_PY508A.setValue("U_Retire", 0, IIf(oMat01.Columns("Retire").Cells(pval.Row).Specific.VALUE = "1", oForm.Items("Retire").Specific.Checked = True, oForm.Items("Retire").Specific.Checked = False)) '퇴사여부
                Call oDS_PH_PY508A.setValue("U_Retire", 0, IIf(oMat01.Columns("Retire").Cells(pval.Row).Specific.VALUE = "1", "Y", "N")) '퇴사여부
                
                oForm.Items("Retire").Enabled = False
                
                oForm.Mode = fm_UPDATE_MODE
                Call PH_PY508_LoadCaption
                
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
    
        Call PH_PY508_FlushToItemValue(pval.ItemUID)
        
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
'                    Call oDS_PH_PY508B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PH_PY508B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PH_PY508_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PH_PY508B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PH_PY508_FlushToItemValue(pval.ItemUID, 0, "", pval.FormMode)
                
'                If pval.ItemUID = "MSTCOD" Then
'
'                    oForm.Items("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.VALUE & "'") '성명
'
'                ElseIf pval.ItemUID = "SMSTCOD" Then
'
'                    oForm.Items("SMSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & oForm.Items("SMSTCOD").Specific.VALUE & "'") '성명
'
'                End If
                
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
        Call PH_PY508_FormItemEnabled
        '//Call PH_PY508_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY508_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY508A", "U_CardCode,U_CardName")
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
'            If (PH_PY508_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PH_PY508A.RemoveRecord(oDS_PH_PY508A.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PH_PY508_Add_MatrixRow(0)
            Else
                If Trim(oDS_PH_PY508A.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PH_PY508_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY508_CreateItems() As Boolean
On Error GoTo PH_PY508_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY508A = oForm.DataSources.DBDataSources("@PH_PY508A")
    Set oDS_PH_PY508B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '//////////기본정보//////////_S
    '기본정보는 실제 테이블과 필드 매핑
    '//////////기본정보//////////_E
        
    '//////////조회정보//////////_S
    '관리번호
    Call oForm.DataSources.UserDataSources.Add("SDocEntry", dt_SHORT_TEXT, 100)
    Call oForm.Items("SDocEntry").Specific.DataBind.SetBound(True, "", "SDocEntry")
        
    '사업장
    Call oForm.DataSources.UserDataSources.Add("SCLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("SCLTCOD").Specific.DataBind.SetBound(True, "", "SCLTCOD")
    
    '문서번호1
    Call oForm.DataSources.UserDataSources.Add("SDocNo1", dt_SHORT_TEXT, 4)
    Call oForm.Items("SDocNo1").Specific.DataBind.SetBound(True, "", "SDocNo1")
    
    '문서번호2
    Call oForm.DataSources.UserDataSources.Add("SDocNo2", dt_SHORT_TEXT, 3)
    Call oForm.Items("SDocNo2").Specific.DataBind.SetBound(True, "", "SDocNo2")
    
    '발행일(FR)
    Call oForm.DataSources.UserDataSources.Add("SFrDate", dt_DATE)
    Call oForm.Items("SFrDate").Specific.DataBind.SetBound(True, "", "SFrDate")
    
    '발행일(TO)
    Call oForm.DataSources.UserDataSources.Add("SToDate", dt_DATE)
    Call oForm.Items("SToDate").Specific.DataBind.SetBound(True, "", "SToDate")
    
    '사번
    Call oForm.DataSources.UserDataSources.Add("SMSTCOD", dt_SHORT_TEXT, 20)
    Call oForm.Items("SMSTCOD").Specific.DataBind.SetBound(True, "", "SMSTCOD")
    
    '성명
    Call oForm.DataSources.UserDataSources.Add("SMSTNAM", dt_SHORT_TEXT, 50)
    Call oForm.Items("SMSTNAM").Specific.DataBind.SetBound(True, "", "SMSTNAM")
    
    '소속팀
    Call oForm.DataSources.UserDataSources.Add("STeamCode", dt_SHORT_TEXT, 10)
    Call oForm.Items("STeamCode").Specific.DataBind.SetBound(True, "", "STeamCode")
    
    '소속담당
    Call oForm.DataSources.UserDataSources.Add("SRspCode", dt_SHORT_TEXT, 10)
    Call oForm.Items("SRspCode").Specific.DataBind.SetBound(True, "", "SRspCode")
    
    '용도(%)
    Call oForm.DataSources.UserDataSources.Add("SUseCmt", dt_LONG_TEXT, 100)
    Call oForm.Items("SUseCmt").Specific.DataBind.SetBound(True, "", "SUseCmt")
    '//////////조회정보//////////_E
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY508_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY508_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PH_PY508_ComboBox_Setting()
On Error GoTo PH_PY508_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    
    Dim CLTCOD As String
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    
    Call oForm.Freeze(True)

    '//////////기본정보//////////
    '소속팀
    Call oForm.Items("TeamCode").Specific.ValidValues.Add("%", "선택")
    sQry = "            SELECT      U_Code AS [Code],"
    sQry = sQry & "                 U_CodeNm As [Name]"
    sQry = sQry & "  FROM       [@PS_HR200L]"
    sQry = sQry & "  WHERE      Code = '1'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
'    sQry = sQry & "                 AND U_Char2 = '" & CLTCOD & "'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("TeamCode").Specific, sQry, "", False, False)
    Call oForm.Items("TeamCode").Specific.Select("%", psk_ByValue)
    
    '소속담당
    Call oForm.Items("RspCode").Specific.ValidValues.Add("%", "선택")
    sQry = "            SELECT      U_Code AS [Code],"
    sQry = sQry & "                 U_CodeNm As [Name]"
    sQry = sQry & "  FROM       [@PS_HR200L]"
    sQry = sQry & "  WHERE      Code = '2'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
'    sQry = sQry & "                 AND U_Char2 = '" & CLTCOD & "'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm.Items("RspCode").Specific, sQry, "", False, False)
    Call oForm.Items("RspCode").Specific.Select(0, psk_Index)
    
    '직책
    Call oForm.Items("Position").Specific.ValidValues.Add("%", "선택")
    sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID"
    Call MDC_SetMod.Set_ComboList(oForm.Items("Position").Specific, sQry, "", False, False)
    Call oForm.Items("Position").Specific.Select(0, psk_Index)
    
    '//////////매트릭스//////////
    '사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    '직책
    sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Position"), sQry)
    
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY508_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY508_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY508_CF_ChooseFromList()
On Error GoTo PH_PY508_CF_ChooseFromList_Error

    Exit Sub
PH_PY508_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY508_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY508_EnableMenus()
On Error GoTo PH_PY508_EnableMenus_Error
    
    Exit Sub
PH_PY508_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY508_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY508_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY508_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PH_PY508_FormItemEnabled
        '//Call PH_PY508_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY508_FormItemEnabled
'        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm.Items("1").Click ct_Regular
    End If
    Exit Sub
PH_PY508_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY508_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY508_FormResize()
On Error GoTo PH_PY508_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PH_PY508_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY508_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PH_PY508_FormReset()
'******************************************************************************
'Function ID : PH_PY508_FormReset()
'해당모듈 : PH_PY508
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY508_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY508A]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PH_PY508A.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PH_PY508A.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    Dim User_BPLID As String
    User_BPLID = MDC_PS_Common.User_BPLID
    
    '//////////기준정보//////////
    Call oDS_PH_PY508A.setValue("U_CLTCOD", 0, User_BPLID) '사업장
    Call oDS_PH_PY508A.setValue("U_DocNo1", 0, "") '문서번호1
    Call oDS_PH_PY508A.setValue("U_DocNo2", 0, "") '문서번호2
    Call oDS_PH_PY508A.setValue("U_CLTAdrs", 0, "") '사업장주소
    Call oDS_PH_PY508A.setValue("U_RepName", 0, "") '대표이사명
    Call oDS_PH_PY508A.setValue("U_RegDt", 0, Format(Now, "yyyyMMdd")) '발행일자
    Call oDS_PH_PY508A.setValue("U_MSTCOD", 0, "") '사번
    Call oDS_PH_PY508A.setValue("U_MSTNAM", 0, "") '성명
    Call oDS_PH_PY508A.setValue("U_TeamCode", 0, "%") '소속팀
    Call oDS_PH_PY508A.setValue("U_RspCode", 0, "%") '소속담당
    Call oDS_PH_PY508A.setValue("U_Position", 0, "%") '직책
    Call oDS_PH_PY508A.setValue("U_GovID", 0, "") '주민등록번호
    Call oDS_PH_PY508A.setValue("U_CurAdrs", 0, "") '현주소
    Call oDS_PH_PY508A.setValue("U_GrpDat", 0, Format(Now, "yyyyMMdd")) '입사일
    Call oDS_PH_PY508A.setValue("U_UseCmt", 0, "") '용도
    
    Call oDS_PH_PY508A.setValue("U_Retire", 0, "N") '퇴사자

    '문서번호
    Call PH_PY508_GetDocNo
    
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
PH_PY508_FormReset_Error:
    Call oForm.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY508_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PH_PY508_GetDocNo()
'******************************************************************************
'Function ID : PH_PY508_GetDocNo()
'해당모듈 : PH_PY508
'기능 : 문서번호 생성
'인수 : 없음
'반환값 : 합계 금액
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY508_GetDocNo_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Dim ErrNum As Integer
    
    Dim StdYear As String
    Dim CLTCOD As String
    
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    StdYear = Left(Trim(oForm.Items("RegDt").Specific.VALUE), 4)
    
    sQry = "EXEC PH_PY508_05 '" & CLTCOD & "', '" & StdYear & "'"
    Call oRecordSet01.DoQuery(sQry)

    oForm.Items("DocNo1").Specific.VALUE = StdYear
    oForm.Items("DocNo2").Specific.VALUE = Trim(oRecordSet01.Fields("DocNo2").VALUE)
    
    Set oRecordSet01 = Nothing

    Exit Sub

PH_PY508_GetDocNo_Error:
    If ErrNum = 1 Then
    Else
        MDC_Com.MDC_GF_Message "PH_PY508_GetDocNo_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    Set oRecordSet01 = Nothing
    
End Sub

Private Sub PH_PY508_Print_Report01()
On Error GoTo PH_PY508_Print_Report01_Error

    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim loopCount As Integer
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String
    Dim DocNo1 As String
    Dim DocNo2 As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY508_Print_Report01_Error
    End If
    
'    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
'    DocNo1 = Trim(oForm.Items("DocNo1").Specific.VALUE)
'    DocNo2 = Trim(oForm.Items("DocNo2").Specific.VALUE)
    
    '조회자사번
    Dim MSTCOD As String
    MSTCOD = MDC_PS_Common.User_MSTCOD

    '임시테이블에 check된항목저장하기 위한 기 정보 삭제
    sQry = "DELETE [Z_PH_PY508] WHERE MSTCOD = '" & MSTCOD & "'"
    Call oRecordSet.DoQuery(sQry)
    
    '임시테이블에 check된항목저장
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PH_PY508B.GetValue("U_ColReg01", loopCount)) = "Y" Then
            sQry = "INSERT INTO [Z_PH_PY508] VALUES ('" & MSTCOD & "', '" & Trim(oDS_PH_PY508B.GetValue("U_ColReg02", loopCount)) & "')"
            Call oRecordSet.DoQuery(sQry)
        End If
    Next

    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY508] 재직증명원"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PH_PY508_01.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PH_PY508_02.rpt"
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
    sQry = "         EXEC [PH_PY508_51] '"
'    sQry = sQry & CLTCOD & "','" '사업장
'    sQry = sQry & DocNo1 & "','" '문서번호1
'    sQry = sQry & DocNo2 & "','" '문서번호2
    sQry = sQry & MSTCOD & "'" '사원번호
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY508_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY508_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY508_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub

