VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_QM170.cls
'//  Module : 품질관리>현장 불합리 관리
'//  Desc : 현장 불합리 개선 등록
'//  FormType : PS_QM170
'//  Create Date(Start) : 2016.08.23
'//  Create Date(End) : 2016.08.24
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
'Private oDS_PS_QM170A As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_QM170B As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private oLast_Mode&

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i As Long
    Dim oInnerXml As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM170.srf")
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
    
    oFormUniqueID01 = "PS_QM170_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM170_CreateItems
    Call PS_QM170_ComboBox_Setting
    Call PS_QM170_CF_ChooseFromList
    Call PS_QM170_EnableMenus
    Call PS_QM170_SetDocument(oFromDocEntry01)
    Call PS_QM170_FormResize
    
'    Call PS_QM170_Add_MatrixRow(0, True)
    Call PS_QM170_LoadCaption
    Call PS_QM170_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [Z_PS_QM170_01]"
    Call RecordSet01.DoQuery(sQry)
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        oForm01.DataSources.UserDataSources("DocEntry").VALUE = 1
    Else
        oForm01.DataSources.UserDataSources("DocEntry").VALUE = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    oMat01.Columns("Check").Visible = False '매트릭스 선택 필드 숨김
    Call PS_QM170_FormReset '폼초기화
    
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

Private Sub PS_QM170_LoadCaption()
'******************************************************************************
'Function ID : PS_QM170_LoadCaption()
'해당모듈 : PS_QM170
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM170_LoadCaption_Error

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
PS_QM170_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_QM170_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM170_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM170B.InsertRecord (oRow)
    End If

    oMat01.AddRow
    oDS_PS_QM170B.Offset = oRow
    oDS_PS_QM170B.setValue "U_LineNum", oRow, oRow + 1

    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM170_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_QM170_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_QM170_MTX01()
'******************************************************************************
'Function ID : PS_QM170_MTX01()
'해당모듈 : PS_QM170
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM170_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim sDocEntry As String '관리번호(PK)
    Dim SBPLID As String '사업장코드(PK)
    Dim SRegNo As String '등록번호(PK)
    Dim SFrDate As String '등록일자(From)
    Dim SToDate As String '등록일자(To)
    Dim SCntcCode As String '개선자사번
    Dim STeamCode As String '소속팀코드
    Dim SRspCode As String '소속담당코드
    Dim SClsCode As String '소속반코드
    Dim SFields As String '분야
    Dim SGrade As String '등급
    Dim STitle As String '제목(%)
    Dim SComment As String '비고(%)

    sDocEntry = oForm01.DataSources.UserDataSources("SDocEntry").VALUE '관리번호
    SBPLID = oForm01.DataSources.UserDataSources("SBPLID").VALUE '사업장
    SRegNo = oForm01.DataSources.UserDataSources("SRegNo").VALUE '등록번호
    SFrDate = oForm01.DataSources.UserDataSources("SFrDate").VALUE '시작일자
    SToDate = oForm01.DataSources.UserDataSources("SToDate").VALUE '종료일자
    SCntcCode = oForm01.DataSources.UserDataSources("SCntcCode").VALUE '개선자사번
    STeamCode = oForm01.DataSources.UserDataSources("STeamCode").VALUE '팀
    SRspCode = oForm01.DataSources.UserDataSources("SRspCode").VALUE '담당
    SClsCode = oForm01.DataSources.UserDataSources("SClsCode").VALUE '반
    SFields = oForm01.DataSources.UserDataSources("SFields").VALUE '분야
    SGrade = oForm01.DataSources.UserDataSources("SGrade").VALUE '등급
    STitle = oForm01.DataSources.UserDataSources("STitle").VALUE '제목
    SComment = oForm01.DataSources.UserDataSources("SComment").VALUE '내용&비고
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)

    Call oForm01.Freeze(True)

    sQry = "         EXEC [PS_QM170_01] '"
    sQry = sQry & sDocEntry & "','" '관리번호
    sQry = sQry & SBPLID & "','"  '사업장
    sQry = sQry & SRegNo & "','"   '등록번호
    sQry = sQry & SFrDate & "','"  '시작일자
    sQry = sQry & SToDate & "','"  '종료일자
    sQry = sQry & SCntcCode & "','"  '사원번호
    sQry = sQry & STeamCode & "','"  '팀
    sQry = sQry & SRspCode & "','"  '담당
    sQry = sQry & SClsCode & "','"  '반
    sQry = sQry & SFields & "','"  '분야
    sQry = sQry & SGrade & "','"  '등급
    sQry = sQry & STitle & "','"  '제목
    sQry = sQry & SComment & "'"  '내용&비고
    
    Call oRecordSet01.DoQuery(sQry)

    oMat01.Clear
    oDS_PS_QM170B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then

        ErrNum = 1

        oForm01.Mode = fm_ADD_MODE

'        Call PS_QM170_Add_MatrixRow(0, True)
        Call PS_QM170_LoadCaption

        GoTo PS_QM170_MTX01_Error

        Exit Sub
    End If

    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_QM170B.Size Then
            oDS_PS_QM170B.InsertRecord (i)
        End If

        oMat01.AddRow
        oDS_PS_QM170B.Offset = i

        Call oDS_PS_QM170B.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_QM170B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PS_QM170B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PS_QM170B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("BPLID").VALUE)) '사업장
        Call oDS_PS_QM170B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("RegNo").VALUE)) '등록번호
        Call oDS_PS_QM170B.setValue("U_ColDt01", i, Trim(oRecordSet01.Fields("RegDate").VALUE)) '등록일자
        Call oDS_PS_QM170B.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("CntcCode").VALUE)) '사번
        Call oDS_PS_QM170B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("CntcName").VALUE)) '성명
        Call oDS_PS_QM170B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("TeamCode").VALUE)) '소속팀
        Call oDS_PS_QM170B.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("RspCode").VALUE)) '소속담당
        Call oDS_PS_QM170B.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("ClsCode").VALUE)) '소속반
        Call oDS_PS_QM170B.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("Fields").VALUE)) '분야
        Call oDS_PS_QM170B.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("Grade").VALUE)) '등급
        Call oDS_PS_QM170B.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("Title").VALUE)) '제목
        Call oDS_PS_QM170B.setValue("U_ColReg14", i, Trim(oRecordSet01.Fields("Comment").VALUE)) '내용&비고
        
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
PS_QM170_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing

    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_QM170_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_QM170_DeleteData()
'******************************************************************************
'Function ID : PS_QM170_DeleteData()
'해당모듈 : PS_QM170
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM170_DeleteData_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim DocEntry As String

    If oForm01.Mode = fm_UPDATE_MODE Then

        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)

        sQry = "SELECT COUNT(*) FROM [Z_PS_QM170_01] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)

        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_QM170_DeleteData_Error
        Else
            sQry = "EXEC PS_QM170_04 '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If

    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")

'    Call PS_QM170_FormReset

'    oForm01.Mode = fm_ADD_MODE

'    Call oForm01.Items("BtnSearch").Click(ct_Regular)

'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PS_QM170_Add_MatrixRow(0, True)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM170_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_QM170_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_QM170_UpdateData() As Boolean
'******************************************************************************
'Function ID : PS_QM170_UpdateData()
'해당모듈 : PS_QM170
'기능 : 기본정보를 수정
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM170_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim DocEntry As Integer '관리번호(PK)
    Dim BPLId As String '사업장코드(PK)
    Dim RegNo As String '등록번호(PK)
    Dim RegDate As String '등록일자
    Dim CntcCode As String '개선자사번
    Dim CntcName As String '개선자성명
    Dim TeamCode As String '소속팀코드
    Dim RspCode As String '소속담당코드
    Dim ClsCode As String '소속반코드
    Dim Fields As String '분야
    Dim Grade As String '등급
    Dim Title As String '제목
    Dim Comment As String '비고

    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE) '관리번호
    BPLId = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    RegNo = Trim(oForm01.Items("RegNo").Specific.VALUE) '등록번호
    RegDate = Trim(oForm01.Items("RegDate").Specific.VALUE) '등록일자
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE) '개선자사번
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE) '개선자성명
    TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE) '소속팀코드
    RspCode = Trim(oForm01.Items("RspCode").Specific.VALUE) '소속담당코드
    ClsCode = Trim(oForm01.Items("ClsCode").Specific.VALUE) '소속반코드
    Fields = Trim(oForm01.Items("Fields").Specific.VALUE) '분야
    Grade = Trim(oForm01.Items("Grade").Specific.VALUE) '등급
    Title = Trim(oForm01.Items("Title").Specific.VALUE) '제목
    Comment = Trim(oForm01.Items("Comment").Specific.VALUE) '비고

    sQry = "         EXEC [PS_QM170_03] '"
    sQry = sQry & DocEntry & "','" '관리번호
    sQry = sQry & BPLId & "','"  '사업장
    sQry = sQry & RegNo & "','"  '등록번호
    sQry = sQry & RegDate & "','"  '등록일자
    sQry = sQry & CntcCode & "','"  '사원번호
    sQry = sQry & CntcName & "','"  '사원성명
    sQry = sQry & TeamCode & "','"  '팀
    sQry = sQry & RspCode & "','"  '담당
    sQry = sQry & ClsCode & "','"  '반
    sQry = sQry & Fields & "','"  '분야
    sQry = sQry & Grade & "','"  '등급
    sQry = sQry & Title & "','"  '제목
    sQry = sQry & Comment & "'"  '비고

    Call RecordSet01.DoQuery(sQry)

    MDC_Com.MDC_GF_Message "수정 완료!", "S"

    Set RecordSet01 = Nothing
    PS_QM170_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM170_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_QM170_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_QM170_AddData() As Boolean
'******************************************************************************
'Function ID : PS_QM170_AddData()
'해당모듈 : PS_QM170
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM170_AddData_Error

    Dim i As Integer
    Dim sQry As String

    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim DocEntry As Integer '관리번호(PK)
    Dim BPLId As String '사업장코드(PK)
    Dim RegNo As String '등록번호(PK)
    Dim RegDate As String '등록일자
    Dim CntcCode As String '개선자사번
    Dim CntcName As String '개선자성명
    Dim TeamCode As String '소속팀코드
'    Dim TeamName As String '소속팀명
    Dim RspCode As String '소속담당코드
'    Dim RspName As String '소속담당명
    Dim ClsCode As String '소속반코드
'    Dim ClsName As String '소속반명
    Dim Fields As String '분야
    Dim Grade As String '등급
    Dim Title As String '제목
    Dim Comment As String '비고

    BPLId = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    RegNo = Trim(oForm01.Items("RegNo").Specific.VALUE) '등록번호
    RegDate = Trim(oForm01.Items("RegDate").Specific.VALUE) '등록일자
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE) '개선자사번
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE) '개선자성명
    TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE) '소속팀코드
'    TeamName = Trim(oForm01.Items("TeamName").Specific.VALUE) '소속팀명
    RspCode = Trim(oForm01.Items("RspCode").Specific.VALUE) '소속담당코드
'    RspName = Trim(oForm01.Items("RspName").Specific.VALUE) '소속담당명
    ClsCode = Trim(oForm01.Items("ClsCode").Specific.VALUE) '소속반코드
'    ClsName = Trim(oForm01.Items("ClsName").Specific.VALUE)'소속반명
    Fields = Trim(oForm01.Items("Fields").Specific.VALUE) '분야
    Grade = Trim(oForm01.Items("Grade").Specific.VALUE) '등급
    Title = Trim(oForm01.Items("Title").Specific.VALUE) '제목
    Comment = Trim(oForm01.Items("Comment").Specific.VALUE) '비고

    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[Z_PS_QM170_01]"
    Call RecordSet01.DoQuery(sQry)

    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If

    sQry = "         EXEC [PS_QM170_02] '"
    sQry = sQry & DocEntry & "','" '관리번호
    sQry = sQry & BPLId & "','"  '사업장
    sQry = sQry & RegNo & "','"  '등록번호
    sQry = sQry & RegDate & "','"  '등록일자
    sQry = sQry & CntcCode & "','"  '사원번호
    sQry = sQry & CntcName & "','"  '사원성명
    sQry = sQry & TeamCode & "','"  '팀
    sQry = sQry & RspCode & "','"  '담당
    sQry = sQry & ClsCode & "','"  '반
    sQry = sQry & Fields & "','"  '분야
    sQry = sQry & Grade & "','"  '등급
    sQry = sQry & Title & "','"  '제목
    sQry = sQry & Comment & "'"  '비고
  
    Call RecordSet02.DoQuery(sQry)

    MDC_Com.MDC_GF_Message "등록 완료!", "S"

    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_QM170_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM170_AddData_Error:
    PS_QM170_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PS_QM170_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_QM170_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_QM170_HeaderSpaceLineDel()
'해당모듈 : PS_QM170
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM170_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("RegNo").Specific.VALUE) = "" '등록번호
            ErrNum = 1
            GoTo PS_QM170_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("RegDate").Specific.VALUE) = "" '등록일자
            ErrNum = 2
            GoTo PS_QM170_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("CntcCode").Specific.VALUE) = "" '사원번호
            ErrNum = 3
            GoTo PS_QM170_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Fields").Specific.VALUE) = "%" '분야
            ErrNum = 4
            GoTo PS_QM170_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Grade").Specific.VALUE) = "%" '등급
            ErrNum = 5
            GoTo PS_QM170_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Title").Specific.VALUE) = "" '제목
            ErrNum = 6
            GoTo PS_QM170_HeaderSpaceLineDel_Error
    End Select

    PS_QM170_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM170_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "등록번호는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("RegNo").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "등록일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("RegDate").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "사원번호는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("CntcCode").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "분야는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Fields").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "등급은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Grade").Click(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "제목은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Title").Click(ct_Regular)
    End If
    PS_QM170_HeaderSpaceLineDel = False
End Function

''/ 메트릭스 필수 사항 check
'Private Function PS_QM170_MatrixSpaceLineDel() As Boolean
'On Error GoTo PS_QM170_MatrixSpaceLineDel_Error
'
'    Dim i As Long
'    Dim ErrNum As Integer
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Dim sQry As String
'
'    Set oRecordSet01 = Nothing
'    PS_QM170_MatrixSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_QM170_MatrixSpaceLineDel_Error:
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
'        MDC_Com.MDC_GF_Message "PS_QM170_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'    PS_QM170_MatrixSpaceLineDel = False
'End Function

Private Sub PS_QM170_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_QM170_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID

'        Case "FrDate"
'
'            Call PS_QM170_GetRegNo
'
'        Case "CLTCOD"
'
'            Call PS_QM170_GetRegNo '출장번호생성
        
        Case "SBPLID"
        
            Set oCombo = oForm01.Items("STeamCode").Specific
        
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
            sQry = sQry & "               AND U_Char2 = '" & oForm01.Items.Item("SBPLID").Specific.VALUE & "'"
            sQry = sQry & "               AND U_UseYN = 'Y'"
            sQry = sQry & " ORDER BY U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("STeamCode").Specific, sQry)
            Call oCombo.Select(0, psk_Index)
            
            'oForm01.Items("STeamCode").DisplayDesc = True
            
        Case "STeamCode"
        
            Set oCombo = oForm01.Items("SRspCode").Specific
            
            If oCombo.ValidValues.Count > 0 Then
            
                For loopCount = oCombo.ValidValues.Count - 1 To 0 Step -1
                    Call oCombo.ValidValues.Remove(loopCount, psk_Index)
                Next
            
            End If
            
            Call oCombo.ValidValues.Add("%", "전체")
            sQry = "           SELECT     U_Code,"
            sQry = sQry & "               U_CodeNm"
            sQry = sQry & " FROM      [@PS_HR200L]"
            sQry = sQry & " WHERE     Code = '2'"
            sQry = sQry & "               AND U_Char2 = '" & oForm01.Items.Item("SBPLID").Specific.VALUE & "'"
            sQry = sQry & "               AND U_Char1 = '" & oForm01.Items.Item("STeamCode").Specific.VALUE & "'"
            sQry = sQry & "               AND U_UseYN = 'Y'"
            sQry = sQry & " ORDER BY U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("SRspCode").Specific, sQry)
            Call oCombo.Select(0, psk_Index)
            
        Case "SRspCode"
        
            Set oCombo = oForm01.Items("SClsCode").Specific
            
            If oCombo.ValidValues.Count > 0 Then
            
                For loopCount = oCombo.ValidValues.Count - 1 To 0 Step -1
                    Call oCombo.ValidValues.Remove(loopCount, psk_Index)
                Next
            
            End If
            
            Call oCombo.ValidValues.Add("%", "전체")
            sQry = "           SELECT     U_Code,"
            sQry = sQry & "               U_CodeNm"
            sQry = sQry & " FROM      [@PS_HR200L]"
            sQry = sQry & " WHERE     Code = '9'"
            sQry = sQry & "               AND U_Char3 = '" & oForm01.Items.Item("SBPLID").Specific.VALUE & "'"
            sQry = sQry & "               AND U_Char2 = '" & oForm01.Items.Item("STeamCode").Specific.VALUE & "'"
            sQry = sQry & "               AND U_Char1 = '" & oForm01.Items.Item("SRspCode").Specific.VALUE & "'"
            sQry = sQry & "               AND U_UseYN = 'Y'"
            sQry = sQry & " ORDER BY U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("SClsCode").Specific, sQry)
            Call oCombo.Select(0, psk_Index)

        Case "CntcCode"
        
            oForm01.DataSources.UserDataSources("CntcName").VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '성명
            oForm01.DataSources.UserDataSources("TeamCode").VALUE = MDC_GetData.Get_ReData("CASE WHEN ISNULL(U_TeamCode, '') = '' THEN '%' ELSE U_TeamCode END", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '팀
            oForm01.DataSources.UserDataSources("RspCode").VALUE = MDC_GetData.Get_ReData("CASE WHEN ISNULL(U_RspCode, '') = '' THEN '%' ELSE U_RspCode END", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '담당
            oForm01.DataSources.UserDataSources("ClsCode").VALUE = MDC_GetData.Get_ReData("CASE WHEN ISNULL(U_ClsCode, '') = '' THEN '%' ELSE U_ClsCode END", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '반
        
        Case "SCntcCode"

            oForm01.Items("SCntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("SCntcCode").Specific.VALUE & "'") '성명

    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_QM170_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_QM170_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PS_QM170_FormItemEnabled()
On Error GoTo PS_QM170_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        
'        oForm01.Items("ItemCode").Enabled = True

    ElseIf (oForm01.Mode = fm_OK_MODE) Then
    
'        '// 접속자에 따른 권한별 사업장 콤보박스세팅
'        Call CLTCOD_Select(oForm01, "CLTCOD")
'        Call CLTCOD_Select(oForm01, "SCLTCOD")
        
    End If
    
    Exit Sub
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM170_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_QM170_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
                Call PS_QM170_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_QM170_LoadCaption
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm01.Freeze(True)
                Call PS_QM170_Add_MatrixRow(oMat01.VisualRowCount)
                Call oForm01.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_QM170_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_QM170A.Clear

'                Call PS_QM170_LoadCaption
'                Call PS_QM170_FormItemEnabled
                '//Call PS_QM170_FormItemEnabled '//UDO방식
                '//Call PS_QM170_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_QM170_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_QM170B.RemoveRecord(oDS_PS_QM170B.Size - 1)
                Call oMat01.LoadFromDataSource
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
    
        If pval.ItemUID = "PS_QM170" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_QM170_HeaderSpaceLineDel() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                If PS_QM170_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                If PS_QM170_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_QM170_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_QM170_LoadCaption
                Call PS_QM170_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_QM170_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PS_QM170_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                If PS_QM170_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_QM170_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_QM170_LoadCaption
                Call PS_QM170_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
'            Call PS_QM170_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_QM170_LoadCaption
            Call PS_QM170_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_QM170_DeleteData
                Call PS_QM170_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_QM170_LoadCaption
                Call PS_QM170_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            Call PS_QM170_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM170" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '기본정보-사번
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "SCntcCode", "") '조회조건-사번
        
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

                'DataSource를 이용하여 각 컨트롤에 값을 대입
                oForm01.DataSources.UserDataSources("DocEntry").VALUE = oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE '관리번호
                oForm01.DataSources.UserDataSources("BPLID").VALUE = oMat01.Columns("BPLID").Cells(pval.Row).Specific.VALUE '사업장
                oForm01.DataSources.UserDataSources("RegNo").VALUE = oMat01.Columns("RegNo").Cells(pval.Row).Specific.VALUE '등록번호
                oForm01.DataSources.UserDataSources("RegDate").VALUE = oMat01.Columns("RegDate").Cells(pval.Row).Specific.VALUE '등록일자
                oForm01.DataSources.UserDataSources("CntcCode").VALUE = oMat01.Columns("CntcCode").Cells(pval.Row).Specific.VALUE '사번
                oForm01.DataSources.UserDataSources("CntcName").VALUE = oMat01.Columns("CntcName").Cells(pval.Row).Specific.VALUE '성명
                oForm01.DataSources.UserDataSources("TeamCode").VALUE = oMat01.Columns("TeamCode").Cells(pval.Row).Specific.VALUE '팀
                oForm01.DataSources.UserDataSources("RspCode").VALUE = oMat01.Columns("RspCode").Cells(pval.Row).Specific.VALUE '담당
                oForm01.DataSources.UserDataSources("ClsCode").VALUE = oMat01.Columns("ClsCode").Cells(pval.Row).Specific.VALUE '반
                oForm01.DataSources.UserDataSources("Fields").VALUE = oMat01.Columns("Fields").Cells(pval.Row).Specific.VALUE '분야
                oForm01.DataSources.UserDataSources("Grade").VALUE = oMat01.Columns("Grade").Cells(pval.Row).Specific.VALUE '등급
                oForm01.DataSources.UserDataSources("Title").VALUE = oMat01.Columns("Title").Cells(pval.Row).Specific.VALUE '제목
                oForm01.DataSources.UserDataSources("Comment").VALUE = oMat01.Columns("Comment").Cells(pval.Row).Specific.VALUE '내용&비고

                oForm01.Mode = fm_UPDATE_MODE
                Call PS_QM170_LoadCaption

                Call oForm01.Freeze(False)

            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PS_QM170_FlushToItemValue(pval.ItemUID)
        
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
'                    Call oDS_PS_QM170B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM170B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_QM170_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_QM170B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PS_QM170_FlushToItemValue(pval.ItemUID)
                
'                If pval.ItemUID = "CntcCode" Then
'
'                    oForm01.Items("CntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '성명
'                    '팀
'                    '담당
'                    '반
'
'                ElseIf pval.ItemUID = "SCntcCode" Then
'
'                    oForm01.Items("SCntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("SCntcCode").Specific.VALUE & "'") '성명
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
        Call PS_QM170_FormItemEnabled
        '//Call PS_QM170_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_QM170_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_QM170A", "U_CardCode,U_CardName")
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
'            If (PS_QM170_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
'            For i = 1 To oMat01.VisualRowCount
'                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
'            Next i
'            oMat01.FlushToDataSource
'            Call oDS_PS_QM170A.RemoveRecord(oDS_PS_QM170A.Size - 1)
'            oMat01.LoadFromDataSource
'            If oMat01.RowCount = 0 Then
'                Call PS_QM170_Add_MatrixRow(0)
'            Else
'                If Trim(oDS_PS_QM170A.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_QM170_Add_MatrixRow(oMat01.RowCount)
'                End If
'            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_QM170_CreateItems() As Boolean
On Error GoTo PS_QM170_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_QM170A = oForm01.DataSources.DBDataSources("@PS_QM170A")
'    Set oDS_PS_QM170A = oForm01.DataSources.UserDataSources("
    Set oDS_PS_QM170B = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '//////////기본정보//////////_S
    '관리번호
    Call oForm01.DataSources.UserDataSources.Add("DocEntry", dt_SHORT_TEXT, 100)
    Call oForm01.Items("DocEntry").Specific.DataBind.SetBound(True, "", "DocEntry")
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
    
    '등록번호
    Call oForm01.DataSources.UserDataSources.Add("RegNo", dt_SHORT_TEXT, 10)
    Call oForm01.Items("RegNo").Specific.DataBind.SetBound(True, "", "RegNo")
    
    '등록일자
    Call oForm01.DataSources.UserDataSources.Add("RegDate", dt_DATE)
    Call oForm01.Items("RegDate").Specific.DataBind.SetBound(True, "", "RegDate")
    
    '사번
    Call oForm01.DataSources.UserDataSources.Add("CntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcCode").Specific.DataBind.SetBound(True, "", "CntcCode")
    
    '성명
    Call oForm01.DataSources.UserDataSources.Add("CntcName", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcName").Specific.DataBind.SetBound(True, "", "CntcName")
    
    '소속팀
    Call oForm01.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("TeamCode").Specific.DataBind.SetBound(True, "", "TeamCode")
    
    '소속담당
    Call oForm01.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("RspCode").Specific.DataBind.SetBound(True, "", "RspCode")
    
    '소속반
    Call oForm01.DataSources.UserDataSources.Add("ClsCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ClsCode").Specific.DataBind.SetBound(True, "", "ClsCode")
    
    '분야
    Call oForm01.DataSources.UserDataSources.Add("Fields", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Fields").Specific.DataBind.SetBound(True, "", "Fields")
    
    '등급
    Call oForm01.DataSources.UserDataSources.Add("Grade", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Grade").Specific.DataBind.SetBound(True, "", "Grade")
    
    '제목
    Call oForm01.DataSources.UserDataSources.Add("Title", dt_SHORT_TEXT, 200)
    Call oForm01.Items("Title").Specific.DataBind.SetBound(True, "", "Title")
    
    '내용&비고
    Call oForm01.DataSources.UserDataSources.Add("Comment", dt_SHORT_TEXT, 200)
    Call oForm01.Items("Comment").Specific.DataBind.SetBound(True, "", "Comment")
    '//////////기본정보//////////_E
    
    '//////////조회정보//////////_S
    '관리번호
    Call oForm01.DataSources.UserDataSources.Add("SDocEntry", dt_SHORT_TEXT, 100)
    Call oForm01.Items("SDocEntry").Specific.DataBind.SetBound(True, "", "SDocEntry")
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("SBPLID", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SBPLID").Specific.DataBind.SetBound(True, "", "SBPLID")
    
    '등록번호
    Call oForm01.DataSources.UserDataSources.Add("SRegNo", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SRegNo").Specific.DataBind.SetBound(True, "", "SRegNo")
    
    '등록기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("SFrDate", dt_DATE)
    Call oForm01.Items("SFrDate").Specific.DataBind.SetBound(True, "", "SFrDate")
    
    '등록기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("SToDate", dt_DATE)
    Call oForm01.Items("SToDate").Specific.DataBind.SetBound(True, "", "SToDate")
    
    '사번
    Call oForm01.DataSources.UserDataSources.Add("SCntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SCntcCode").Specific.DataBind.SetBound(True, "", "SCntcCode")
    
    '성명
    Call oForm01.DataSources.UserDataSources.Add("SCntcName", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SCntcName").Specific.DataBind.SetBound(True, "", "SCntcName")
    
    '소속팀
    Call oForm01.DataSources.UserDataSources.Add("STeamCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("STeamCode").Specific.DataBind.SetBound(True, "", "STeamCode")
    
    '소속담당
    Call oForm01.DataSources.UserDataSources.Add("SRspCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SRspCode").Specific.DataBind.SetBound(True, "", "SRspCode")
    
    '소속반
    Call oForm01.DataSources.UserDataSources.Add("SClsCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SClsCode").Specific.DataBind.SetBound(True, "", "SClsCode")
    
    '분야
    Call oForm01.DataSources.UserDataSources.Add("SFields", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SFields").Specific.DataBind.SetBound(True, "", "SFields")
    
    '등급
    Call oForm01.DataSources.UserDataSources.Add("SGrade", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SGrade").Specific.DataBind.SetBound(True, "", "SGrade")
    
    '제목
    Call oForm01.DataSources.UserDataSources.Add("STitle", dt_SHORT_TEXT, 200)
    Call oForm01.Items("STitle").Specific.DataBind.SetBound(True, "", "STitle")
    
    '내용&비고
    Call oForm01.DataSources.UserDataSources.Add("SComment", dt_SHORT_TEXT, 200)
    Call oForm01.Items("SComment").Specific.DataBind.SetBound(True, "", "SComment")
    '//////////조회정보//////////_E
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_QM170_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM170_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_QM170_ComboBox_Setting()
On Error GoTo PS_QM170_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId As String
    BPLId = MDC_PS_Common.User_BPLId
    
    Call oForm01.Freeze(True)

    '//////////기본정보//////////
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLId, False, False)
    
    '분야
    Call oForm01.Items("Fields").Specific.ValidValues.Add("%", "선택")
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName As [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'Q231'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Fields").Specific, sQry, "", False, False)
    Call oForm01.Items("Fields").Specific.Select(0, psk_Index)
    
    '등급
    Call oForm01.Items("Grade").Specific.ValidValues.Add("%", "선택")
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName As [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'Q232'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Grade").Specific, sQry, "", False, False)
    Call oForm01.Items("Grade").Specific.Select(0, psk_Index)
    
    '팀
    Call oForm01.Items("TeamCode").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT     U_Code,"
    sQry = sQry & "               U_CodeNm"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = '1'"
    sQry = sQry & "               AND U_Char2 = '" & oForm01.Items.Item("BPLID").Specific.VALUE & "'"
    sQry = sQry & "               AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("TeamCode").Specific, sQry)
    Call oForm01.Items("TeamCode").Specific.Select(0, psk_Index)
    
    '담당
    Call oForm01.Items("RspCode").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT     U_Code,"
    sQry = sQry & "               U_CodeNm"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = '2'"
    sQry = sQry & "               AND U_Char2 = '" & oForm01.Items.Item("BPLID").Specific.VALUE & "'"
    sQry = sQry & "               AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("RspCode").Specific, sQry)
    Call oForm01.Items("RspCode").Specific.Select(0, psk_Index)
    
    '반
    Call oForm01.Items("ClsCode").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT     U_Code,"
    sQry = sQry & "               U_CodeNm"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = '9'"
    sQry = sQry & "               AND U_Char3 = '" & oForm01.Items.Item("BPLID").Specific.VALUE & "'"
    sQry = sQry & "               AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ClsCode").Specific, sQry)
    Call oForm01.Items("ClsCode").Specific.Select(0, psk_Index)
    
    '//////////조회정보//////////
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLId, False, False)
    
    '분야
    Call oForm01.Items("SFields").Specific.ValidValues.Add("%", "전체")
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName As [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'Q231'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SFields").Specific, sQry, "", False, False)
    Call oForm01.Items("SFields").Specific.Select(0, psk_Index)
    
    '등급
    Call oForm01.Items("SGrade").Specific.ValidValues.Add("%", "전체")
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName As [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'Q232'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SGrade").Specific, sQry, "", False, False)
    Call oForm01.Items("SGrade").Specific.Select(0, psk_Index)
    
    '//////////매트릭스//////////
    '사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLID"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
'
    '소속팀
    sQry = "           SELECT     U_Code,"
    sQry = sQry & "               U_CodeNm"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = '1'"
    sQry = sQry & "               AND U_Char2 = '" & oForm01.Items.Item("BPLID").Specific.VALUE & "'"
    sQry = sQry & "               AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("TeamCode"), sQry)

    '소속담당
    sQry = "           SELECT     U_Code,"
    sQry = sQry & "               U_CodeNm"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = '2'"
    sQry = sQry & "               AND U_Char2 = '" & oForm01.Items.Item("BPLID").Specific.VALUE & "'"
    sQry = sQry & "               AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("RspCode"), sQry)

    '소속반
    sQry = "           SELECT     U_Code,"
    sQry = sQry & "               U_CodeNm"
    sQry = sQry & " FROM      [@PS_HR200L]"
    sQry = sQry & " WHERE     Code = '9'"
    sQry = sQry & "               AND U_Char3 = '" & oForm01.Items.Item("BPLID").Specific.VALUE & "'"
    sQry = sQry & "               AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ClsCode"), sQry)
    
    '분야
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName As [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'Q231'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Fields"), sQry)
    
    '등급
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName As [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'Q232'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Grade"), sQry)

    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_QM170_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM170_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM170_CF_ChooseFromList()
On Error GoTo PS_QM170_CF_ChooseFromList_Error

    Exit Sub
PS_QM170_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM170_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM170_EnableMenus()
On Error GoTo PS_QM170_EnableMenus_Error
    
    Exit Sub
PS_QM170_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM170_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM170_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM170_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM170_FormItemEnabled
        '//Call PS_QM170_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_QM170_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_QM170_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM170_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM170_FormResize()
On Error GoTo PS_QM170_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_QM170_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM170_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_QM170_FormReset()
'******************************************************************************
'Function ID : PS_QM170_FormReset()
'해당모듈 : PS_QM170
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM170_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [Z_PS_QM170_01]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        oForm01.DataSources.UserDataSources("DocEntry").VALUE = 1
    Else
        oForm01.DataSources.UserDataSources("DocEntry").VALUE = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    Dim User_BPLId As String
    User_BPLId = MDC_PS_Common.User_BPLId
    
    '//////////기준정보//////////
    oForm01.DataSources.UserDataSources("BPLID").VALUE = User_BPLId '사업장
    oForm01.DataSources.UserDataSources("RegNo").VALUE = "" '등록번호
    oForm01.DataSources.UserDataSources("RegDate").VALUE = Format(Now, "YYYYMMDD") '등록일자
    oForm01.DataSources.UserDataSources("CntcCode").VALUE = "" '사번
    oForm01.DataSources.UserDataSources("Fields").VALUE = "%" '분야
    oForm01.DataSources.UserDataSources("Grade").VALUE = "%" '등급
    oForm01.DataSources.UserDataSources("TeamCode").VALUE = "%" '팀
    oForm01.DataSources.UserDataSources("RspCode").VALUE = "%" '담당
    oForm01.DataSources.UserDataSources("ClsCode").VALUE = "%" '반
    oForm01.DataSources.UserDataSources("Title").VALUE = "" '제목
    oForm01.DataSources.UserDataSources("Comment").VALUE = "" '내용&비고
    '등록번호
    Call PS_QM170_GetRegNo
    
    '//////////조회정보//////////
    oForm01.DataSources.UserDataSources("SDocEntry").VALUE = "" '관리번호
    Call oForm01.Items("SBPLID").Specific.Select(User_BPLId, psk_ByValue) '사업장 '이벤트 실행을 위하여 UserDataSource 사용 안함
    oForm01.DataSources.UserDataSources("SRegNo").VALUE = "" '등록번호
    oForm01.Items("SFrDate").Specific.VALUE = "" 'Format(Now, "yyyyMM01") '등록기간(시작)
    oForm01.Items("SToDate").Specific.VALUE = "" 'Format(Now, "yyyyMMdd") '등록기간(종료)
    oForm01.DataSources.UserDataSources("SCntcCode").VALUE = "" '사번
    oForm01.DataSources.UserDataSources("SFields").VALUE = "%" '분야
    oForm01.DataSources.UserDataSources("SGrade").VALUE = "%" '등급
    oForm01.DataSources.UserDataSources("STeamCode").VALUE = "%" '팀
    oForm01.DataSources.UserDataSources("SRspCode").VALUE = "%" '담당
    oForm01.DataSources.UserDataSources("SClsCode").VALUE = "%" '반
    oForm01.DataSources.UserDataSources("STitle").VALUE = "" '제목
    oForm01.DataSources.UserDataSources("SComment").VALUE = "" '내용&비고
    
    Call oForm01.Items("CntcCode").Click
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM170_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_QM170_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_QM170_GetRegNo()
'******************************************************************************
'Function ID : PS_QM170_GetRegNo()
'해당모듈 : PS_QM170
'기능 : 등록번호 생성
'인수 : 없음
'반환값 : 합계 금액
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM170_GetRegNo_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ErrNum As Integer
    
    Dim RegDate As String
    Dim BPLId As String
    
    BPLId = oForm01.Items("BPLID").Specific.VALUE
    RegDate = Left(Trim(oForm01.Items("RegDate").Specific.VALUE), 6)
    
    sQry = "EXEC PS_QM170_05 '" & BPLId & "', '" & RegDate & "'"
    Call oRecordSet01.DoQuery(sQry)

    oForm01.DataSources.UserDataSources("RegNo").VALUE = Trim(oRecordSet01.Fields("RegNo").VALUE) ' BPLID & "-" & RegDate

'    oForm01.Items("DestNo1").Specific.VALUE = FrDate
'    oForm01.Items("DestNo2").Specific.VALUE = Trim(oRecordSet01.Fields("DestNo2").VALUE)
    
    Set oRecordSet01 = Nothing

    Exit Sub

PS_QM170_GetRegNo_Error:
    If ErrNum = 1 Then
    Else
        MDC_Com.MDC_GF_Message "PS_QM170_GetRegNo_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    Set oRecordSet01 = Nothing
    
End Sub

Private Sub PS_QM170_Print_Report01()

    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_QM170_Print_Report01_Error
    
    Dim CLTCOD As String
    Dim DestNo1 As String
    Dim DestNo2 As String
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
     '/ ODBC 연결 체크
    Call ConnectODBC
        
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm01.Items("CLTCOD").Specific.VALUE)
    DestNo1 = Trim(oForm01.Items("DestNo1").Specific.VALUE)
    DestNo2 = Trim(oForm01.Items("DestNo2").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PS_QM170] 공용증"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PS_QM170_01.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PS_QM170_02.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PS_QM170_03.rpt"
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
    sQry = "EXEC [PS_QM170_90] '" & CLTCOD & "','" & DestNo1 & "','" & DestNo2 & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PS_QM170_Print_Report01_Error
    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_QM170_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_QM170_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub

