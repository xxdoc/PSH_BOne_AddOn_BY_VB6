VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PH_PY032.cls
'//  Module : 인사관리>기타관리
'//  Desc : 사용외출등록
'//  FormType : PH_PY032
'//  Create Date(Start) : 2013.03.19
'//  Create Date(End) :
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PH_PY032A As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PH_PY032B As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY032.srf")
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
    
    oFormUniqueID01 = "PH_PY032_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01, "PH_PY032" '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    '//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm.Freeze True
    Call PH_PY032_CreateItems
    Call PH_PY032_ComboBox_Setting
    Call PH_PY032_CF_ChooseFromList
    Call PH_PY032_EnableMenus
    Call PH_PY032_SetDocument(oFromDocEntry01)
    Call PH_PY032_FormResize
    

'    Call PH_PY032_Add_MatrixRow(0, True)
    Call PH_PY032_LoadCaption
    Call PH_PY032_FormItemEnabled
    
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
        
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY032A]"
    Call RecordSet01.DoQuery(sQry)
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PH_PY032A.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PH_PY032A.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    Call PH_PY032_FormReset '폼초기화 추가(2013.01.29 송명규)
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing

    '기간(월)
    oForm.Items("SFrDate").Specific.VALUE = Format(Now, "YYYY.MM")
    oForm.Items("SToDate").Specific.VALUE = Format(Now, "YYYY.MM")
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

Private Sub PH_PY032_LoadCaption()
'******************************************************************************
'Function ID : PH_PY032_LoadCaption()
'해당모듈 : PH_PY032
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_LoadCaption_Error

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
PH_PY032_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PH_PY032_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PH_PY032_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PH_PY032B.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PH_PY032B.Offset = oRow
    oDS_PH_PY032B.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY032_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PH_PY032_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PH_PY032_MTX01()
'******************************************************************************
'Function ID : PH_PY032_MTX01()
'해당모듈 : PH_PY032
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim sDocEntry As String '관리번호
    Dim sCLTCOD As String '사업장
    Dim sTeamCd As String
    Dim sMSTCOD As String '사원번호
    Dim SFrDate As String '시작일자
    Dim SToDate As String '종료일자
    Dim SObject As String '목적
    
    sCLTCOD = Trim(oForm.Items("SCLTCOD").Specific.VALUE) '사업장
    sTeamCd = Trim(oForm.Items("STeamCd").Specific.VALUE) '부서
    sMSTCOD = Trim(oForm.Items("SMSTCOD").Specific.VALUE) '사원번호
    SFrDate = Replace(Trim(oForm.Items("SFrDate").Specific.VALUE), ".", "") '시작일자
    SToDate = Replace(Trim(oForm.Items("SToDate").Specific.VALUE), ".", "") '종료일자
    SObject = Trim(oForm.Items("SObject").Specific.VALUE) '목적
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "                EXEC [PH_PY032_01] "
    sQry = sQry & "'" & sDocEntry & "'," '관리번호
    sQry = sQry & "'" & sCLTCOD & "',"  '사업장
    sQry = sQry & "'" & sTeamCd & "',"  '부서
    sQry = sQry & "'" & sMSTCOD & "',"  '사원번호
    sQry = sQry & "'" & SFrDate & "',"  '시작일자
    sQry = sQry & "'" & SToDate & "',"  '종료일자
    sQry = sQry & "'" & SObject & "'" '목적
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PH_PY032B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm.Mode = fm_ADD_MODE
        
'        Call PH_PY032_Add_MatrixRow(0, True)
        Call PH_PY032_LoadCaption
        
        GoTo PH_PY032_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PH_PY032B.Size Then
            oDS_PH_PY032B.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PH_PY032B.Offset = i
        
        Call oDS_PH_PY032B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY032B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PH_PY032B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("CLTCOD").VALUE)) '사업장
        Call oDS_PH_PY032B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("MSTCOD").VALUE)) '사원번호
        Call oDS_PH_PY032B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("MSTNAM").VALUE)) '사원성명
        Call oDS_PH_PY032B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("FrDate").VALUE)) '시작일자
        Call oDS_PH_PY032B.setValue("U_ColTm01", i, Trim(oRecordSet01.Fields("FrTime").VALUE)) '시작시각
        Call oDS_PH_PY032B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("ToDate").VALUE)) '종료일자
        Call oDS_PH_PY032B.setValue("U_ColTm02", i, Trim(oRecordSet01.Fields("ToTime").VALUE)) '종료시각
        Call oDS_PH_PY032B.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("Object").VALUE)) '목적
        
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
PH_PY032_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY032_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PH_PY032_DeleteData()
'******************************************************************************
'Function ID : PH_PY032_DeleteData()
'해당모듈 : PH_PY032
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_DeleteData_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PH_PY032A] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm.Mode = fm_ADD_MODE
            GoTo PH_PY032_DeleteData_Error
        Else
            sQry = "EXEC PH_PY032_04 '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
'    Call PH_PY032_FormReset
    
'    oForm.Mode = fm_ADD_MODE
    
'    Call oForm.Items("BtnSearch").Click(ct_Regular)
    
'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PH_PY032_Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY032_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY032_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PH_PY032_UpdateData() As Boolean
'******************************************************************************
'Function ID : PH_PY032_UpdateData()
'해당모듈 : PH_PY032
'기능 : 기본정보를 수정
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer '관리번호
    Dim CLTCOD As String '사업장
    Dim MSTCOD As String '사원번호
    Dim MSTNAM As String '사원성명
    Dim FrDate As String '시작일자
    Dim FrTime As String '시작시각
    Dim ToDate As String '종료일자
    Dim ToTime As String '종료시각
    Dim Object As String '목적
    
    DocEntry = Trim(oForm.Items("DocEntry").Specific.VALUE) '관리번호
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사원번호
    MSTNAM = Trim(oForm.Items("MSTNAM").Specific.VALUE) '사원성명
    FrDate = Trim(oForm.Items("FrDate").Specific.VALUE) '시작일자
    FrTime = Trim(oForm.Items("FrTime").Specific.VALUE) '시작시각
    ToDate = Trim(oForm.Items("ToDate").Specific.VALUE) '종료일자
    ToTime = Trim(oForm.Items("ToTime").Specific.VALUE) '종료시각
    Object = Trim(oForm.Items("Object").Specific.VALUE) '목적
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
        PH_PY032_UpdateData = False
        Exit Function
    End If
    
    sQry = "                EXEC [PH_PY032_03] "
    sQry = sQry & "'" & DocEntry & "'," '관리번호
    sQry = sQry & "'" & CLTCOD & "',"  '사업장
    sQry = sQry & "'" & MSTCOD & "',"  '사원번호
    sQry = sQry & "'" & MSTNAM & "',"  '사원성명
    sQry = sQry & "'" & FrDate & "',"  '시작일자
    sQry = sQry & "'" & FrTime & "',"  '시작시각
    sQry = sQry & "'" & ToDate & "',"  '종료일자
    sQry = sQry & "'" & ToTime & "',"  '종료시각
    sQry = sQry & "'" & Object & "'"  '목적
    
    Call RecordSet01.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PH_PY032_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY032_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY032_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PH_PY032_AddData() As Boolean
'******************************************************************************
'Function ID : PH_PY032_AddData()
'해당모듈 : PH_PY032
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_AddData_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer '관리번호
    Dim CLTCOD As String '사업장
    Dim MSTCOD As String '사원번호
    Dim MSTNAM As String '사원성명
    Dim FrDate As String '시작일자
    Dim FrTime As String '시작시각
    Dim ToDate As String '종료일자
    Dim ToTime As String '종료시각
    Dim Object As String '목적
    Dim UserSign As String 'UserSign

    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사원번호
    MSTNAM = Trim(oForm.Items("MSTNAM").Specific.VALUE) '사원성명
    FrDate = Trim(oForm.Items("FrDate").Specific.VALUE) '시작일자
    FrTime = Trim(oForm.Items("FrTime").Specific.VALUE) '시작시각
    ToDate = Trim(oForm.Items("ToDate").Specific.VALUE) '종료일자
    ToTime = Trim(oForm.Items("ToTime").Specific.VALUE) '종료시각
    Object = Trim(oForm.Items("Object").Specific.VALUE) '목적
    UserSign = oCompany.UserSignature
    
    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY032A]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "                EXEC [PH_PY032_02] "
    sQry = sQry & "'" & DocEntry & "'," '관리번호
    sQry = sQry & "'" & CLTCOD & "',"  '사업장
    sQry = sQry & "'" & MSTCOD & "',"  '사원번호
    sQry = sQry & "'" & MSTNAM & "',"  '사원성명
    sQry = sQry & "'" & FrDate & "',"  '시작일자
    sQry = sQry & "'" & FrTime & "',"  '시작시각
    sQry = sQry & "'" & ToDate & "',"  '종료일자
    sQry = sQry & "'" & ToTime & "',"  '종료시각
    sQry = sQry & "'" & Object & "',"  '목적
    sQry = sQry & "'" & UserSign & "'"  'UserSign
    
    Call RecordSet02.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PH_PY032_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY032_AddData_Error:
    PH_PY032_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY032_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PH_PY032_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PH_PY032_HeaderSpaceLineDel()
'해당모듈 : PH_PY032
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm.Items("MSTCOD").Specific.VALUE) = "" '사원번호
            ErrNum = 3
            GoTo PH_PY032_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrDate").Specific.VALUE) = "" '시작일자
            ErrNum = 4
            GoTo PH_PY032_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrTime").Specific.VALUE) = "" '시작시각
            ErrNum = 5
            GoTo PH_PY032_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToDate").Specific.VALUE) = "" '종료일자
            ErrNum = 6
            GoTo PH_PY032_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToTime").Specific.VALUE) = "" '종료시각
            ErrNum = 7
            GoTo PH_PY032_HeaderSpaceLineDel_Error
    End Select

    PH_PY032_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY032_HeaderSpaceLineDel_Error:
    If ErrNum = 3 Then
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
    PH_PY032_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PH_PY032_MatrixSpaceLineDel() As Boolean
On Error GoTo PH_PY032_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PH_PY032_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY032_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PH_PY032_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PH_PY032_MatrixSpaceLineDel = False
End Function

Private Sub PH_PY032_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY032_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim sCLTCOD As String
    
    Select Case oUID
        
        Case "SCLTCOD"
            
            sCLTCOD = Trim(oForm.Items("SCLTCOD").Specific.VALUE)
            
            If oForm.Items("STeamCd").Specific.ValidValues.Count > 0 Then
                For i = oForm.Items("STeamCd").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("STeamCd").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '부서콤보세팅
            Call oForm.Items("STeamCd").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '1'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char2 = '" & sCLTCOD & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("STeamCd").Specific, sQry, "", False, False)
            Call oForm.Items("STeamCd").Specific.Select(0, psk_Index)
        
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PH_PY032_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PH_PY032_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PH_PY032_FormItemEnabled()
On Error GoTo PH_PY032_FormItemEnabled_Error
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
PH_PY032_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY032_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
                Call PH_PY032_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PH_PY032_LoadCaption
                
                'oForm.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PH_PY032_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PH_PY032A.Clear

'                Call PH_PY032_LoadCaption
'                Call PH_PY032_FormItemEnabled
                '//Call PH_PY032_FormItemEnabled '//UDO방식
                '//Call PH_PY032_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY032_FormItemEnabled
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
    
        If pval.ItemUID = "PH_PY032" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm.Mode = fm_ADD_MODE Then
            
                If PH_PY032_HeaderSpaceLineDel() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                If PH_PY032_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                If PH_PY032_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PH_PY032_FormReset
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY032_LoadCaption
                Call PH_PY032_MTX01
                
                oLast_Mode = oForm.Mode
                
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            
                If PH_PY032_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PH_PY032_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                If PH_PY032_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PH_PY032_FormReset
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY032_LoadCaption
                Call PH_PY032_MTX01
                
'                oForm.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
            Call PH_PY032_FormReset
            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PH_PY032_LoadCaption
            Call PH_PY032_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PH_PY032_DeleteData
                Call PH_PY032_FormReset
                oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PH_PY032_LoadCaption
                Call PH_PY032_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            Call PH_PY032_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PH_PY032" Then
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
                Call oDS_PH_PY032A.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
                Call oDS_PH_PY032A.setValue("U_CLTCOD", 0, oMat01.Columns("CLTCOD").Cells(pval.Row).Specific.VALUE) '사업장
                Call oDS_PH_PY032A.setValue("U_MSTCOD", 0, oMat01.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE) '사원번호
                Call oDS_PH_PY032A.setValue("U_MSTNAM", 0, oMat01.Columns("MSTNAM").Cells(pval.Row).Specific.VALUE) '사원성명
                Call oDS_PH_PY032A.setValue("U_FrDate", 0, Replace(oMat01.Columns("FrDate").Cells(pval.Row).Specific.VALUE, ".", "")) '시작일자
                Call oDS_PH_PY032A.setValue("U_FrTime", 0, oMat01.Columns("FrTime").Cells(pval.Row).Specific.VALUE) '시작시각
                Call oDS_PH_PY032A.setValue("U_ToDate", 0, Replace(oMat01.Columns("ToDate").Cells(pval.Row).Specific.VALUE, ".", "")) '종료일자
                Call oDS_PH_PY032A.setValue("U_ToTime", 0, oMat01.Columns("ToTime").Cells(pval.Row).Specific.VALUE) '종료시각
                Call oDS_PH_PY032A.setValue("U_Object", 0, oMat01.Columns("Object").Cells(pval.Row).Specific.VALUE) '목적
                
                oForm.Mode = fm_UPDATE_MODE
                Call PH_PY032_LoadCaption
                
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
    
        Call PH_PY032_FlushToItemValue(pval.ItemUID)
        
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
'                    Call oDS_PH_PY032B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PH_PY032B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PH_PY032_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PH_PY032B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PH_PY032_FlushToItemValue(pval.ItemUID)
                
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
        Call PH_PY032_FormItemEnabled
        '//Call PH_PY032_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY032_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY032A", "U_CardCode,U_CardName")
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
'            If (PH_PY032_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PH_PY032A.RemoveRecord(oDS_PH_PY032A.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PH_PY032_Add_MatrixRow(0)
            Else
                If Trim(oDS_PH_PY032A.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PH_PY032_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY032_CreateItems() As Boolean
On Error GoTo PH_PY032_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY032A = oForm.DataSources.DBDataSources("@PH_PY032A")
    Set oDS_PH_PY032B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
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
    
    '부서
    Call oForm.DataSources.UserDataSources.Add("STeamCd", dt_SHORT_TEXT, 10)
    Call oForm.Items("STeamCd").Specific.DataBind.SetBound(True, "", "STeamCd")
    
    '사원번호
    Call oForm.DataSources.UserDataSources.Add("SMSTCOD", dt_SHORT_TEXT, 20)
    Call oForm.Items("SMSTCOD").Specific.DataBind.SetBound(True, "", "SMSTCOD")
    
    '사원성명
    Call oForm.DataSources.UserDataSources.Add("SMSTNAM", dt_SHORT_TEXT, 50)
    Call oForm.Items("SMSTNAM").Specific.DataBind.SetBound(True, "", "SMSTNAM")
    
    '시작월
    Call oForm.DataSources.UserDataSources.Add("SFrDate", dt_SHORT_TEXT, 10)
    Call oForm.Items("SFrDate").Specific.DataBind.SetBound(True, "", "SFrDate")
    
    '종료월
    Call oForm.DataSources.UserDataSources.Add("SToDate", dt_SHORT_TEXT, 10)
    Call oForm.Items("SToDate").Specific.DataBind.SetBound(True, "", "SToDate")
    
    '목적
    Call oForm.DataSources.UserDataSources.Add("SObject", dt_LONG_TEXT, 100)
    Call oForm.Items("SObject").Specific.DataBind.SetBound(True, "", "SObject")
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY032_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY032_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PH_PY032_ComboBox_Setting()
On Error GoTo PH_PY032_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
 
    '//////////조회정보//////////
'    '부서
'    Call oForm.Items("STeamCd").Specific.ValidValues.Add("%", "전체")
'    sQry = "            SELECT      U_Code AS [Code],"
'    sQry = sQry & "                 U_CodeNm As [Name]"
'    sQry = sQry & "  FROM       [@PS_HR200L]"
'    sQry = sQry & "  WHERE      Code = '1'"
'    sQry = sQry & "                 AND U_UseYN = 'Y'"
'    sQry = sQry & "  ORDER BY  U_Seq"
'    Call MDC_SetMod.Set_ComboList(oForm.Items("STeamCd").Specific, sQry, "", False, False)
'    Call oForm.Items("STeamCd").Specific.Select(0, psk_Index)

    '//////////매트릭스//////////
    '사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY032_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY032_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY032_CF_ChooseFromList()
On Error GoTo PH_PY032_CF_ChooseFromList_Error

    Exit Sub
PH_PY032_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY032_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY032_EnableMenus()
On Error GoTo PH_PY032_EnableMenus_Error
    
    Exit Sub
PH_PY032_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY032_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY032_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY032_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PH_PY032_FormItemEnabled
        '//Call PH_PY032_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY032_FormItemEnabled
'        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm.Items("1").Click ct_Regular
    End If
    Exit Sub
PH_PY032_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY032_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY032_FormResize()
On Error GoTo PH_PY032_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
'    oForm.Items("Static15").Left = oForm.Width - oForm.Items("Static07").Left - 49
'    oForm.Items("GrpBox02").Left = oForm.Width - oForm.Items("Static07").Left - 50
    
    Exit Sub
PH_PY032_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY032_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PH_PY032_FormReset()
'******************************************************************************
'Function ID : PH_PY032_FormReset()
'해당모듈 : PH_PY032
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY032A]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PH_PY032A.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PH_PY032A.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    Dim User_BPLID As String
    User_BPLID = MDC_PS_Common.User_BPLID
    
    '//////////기준정보//////////
    Call oDS_PH_PY032A.setValue("U_CLTCOD", 0, User_BPLID) '사업장
    Call oDS_PH_PY032A.setValue("U_MSTCOD", 0, "") '사원번호
    Call oDS_PH_PY032A.setValue("U_MSTNAM", 0, "") '사원성명
    Call oDS_PH_PY032A.setValue("U_FrDate", 0, Format(Now, "yyyyMMdd")) '시작일자
    Call oDS_PH_PY032A.setValue("U_FrTime", 0, "") '시작시각
    Call oDS_PH_PY032A.setValue("U_ToDate", 0, Format(Now, "yyyyMMdd")) '종료일자
    Call oDS_PH_PY032A.setValue("U_ToTime", 0, "") '종료시각
    Call oDS_PH_PY032A.setValue("U_Object", 0, "") '목적
    '출장번호
'    Call PH_PY032_GetDestNo
    
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
PH_PY032_FormReset_Error:
    Call oForm.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY032_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PH_PY032_CalculateTransExp()
'******************************************************************************
'Function ID : PH_PY032_CalculateTransExp()
'해당모듈 : PH_PY032
'기능 : 교통비 계산
'인수 : 없음
'반환값 : 합계 금액
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_CalculateTransExp_Error

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Dim ErrNum As Integer

    Dim FuelPrc As Double '유류단가
    Dim Distance As Double '거리
    Dim TransExp As Double '교통비
    
    FuelPrc = Trim(oForm.Items("FuelPrc").Specific.VALUE)
    Distance = Trim(oForm.Items("Distance").Specific.VALUE)

    TransExp = ((FuelPrc * Distance * 0.1) \ 10) * 10 '원단위 절사
    
    oForm.Items("TransExp").Specific.VALUE = TransExp

    Set oRecordSet01 = Nothing

    Exit Sub

PH_PY032_CalculateTransExp_Error:
    If ErrNum = 1 Then
    Else
        MDC_Com.MDC_GF_Message "PH_PY032_CalculateTransExp_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    Set oRecordSet01 = Nothing
End Sub

Private Sub PH_PY032_CalculateTotalExp()
'******************************************************************************
'Function ID : PH_PY032_CalculateTotalExp()
'해당모듈 : PH_PY032
'기능 : 합계 금액 계산
'인수 : 없음
'반환값 : 합계 금액
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_CalculateTotalExp_Error

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Dim ErrNum As Integer

    Dim TransExp As Double '교통비
    Dim DayExp As Double '일비
    Dim FoodExp As Double '식비
    Dim ParkExp As Double '주차비
    Dim TollExp As Double '도로비
    Dim TotalExp As Double '합계
    
    TransExp = Trim(oForm.Items("TransExp").Specific.VALUE)
    DayExp = Trim(oForm.Items("DayExp").Specific.VALUE)
    FoodExp = Trim(oForm.Items("FoodExp").Specific.VALUE)
    ParkExp = Trim(oForm.Items("ParkExp").Specific.VALUE)
    TollExp = Trim(oForm.Items("TollExp").Specific.VALUE)
    TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp
    
    Call oDS_PH_PY032A.setValue("U_TotalExp", 0, TotalExp)

    Set oRecordSet01 = Nothing

    Exit Sub

PH_PY032_CalculateTotalExp_Error:
    If ErrNum = 1 Then
    Else
        MDC_Com.MDC_GF_Message "PH_PY032_CalculateTotalExp_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    Set oRecordSet01 = Nothing
    
End Sub

Private Sub PH_PY032_GetDestNo()
'******************************************************************************
'Function ID : PH_PY032_GetDestNo()
'해당모듈 : PH_PY032
'기능 : 출장번호 생성
'인수 : 없음
'반환값 : 합계 금액
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_GetDestNo_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Dim ErrNum As Integer
    
    Dim FrDate As String
    Dim CLTCOD As String
    
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    FrDate = Left(Trim(oForm.Items("FrDate").Specific.VALUE), 6)
    
    sQry = "EXEC PH_PY032_05 '" & CLTCOD & "', '" & FrDate & "'"
    Call oRecordSet01.DoQuery(sQry)

    oForm.Items("DestNo1").Specific.VALUE = FrDate
    oForm.Items("DestNo2").Specific.VALUE = Trim(oRecordSet01.Fields("DestNo2").VALUE)
    
    Set oRecordSet01 = Nothing

    Exit Sub

PH_PY032_GetDestNo_Error:
    If ErrNum = 1 Then
    Else
        MDC_Com.MDC_GF_Message "PH_PY032_GetDestNo_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    Set oRecordSet01 = Nothing
    
End Sub

Private Sub PH_PY032_GetFuelPrc()
'******************************************************************************
'Function ID : PH_PY032_GetFuelPrc()
'해당모듈 : PH_PY032
'기능 : 유류단가 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_GetFuelPrc_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim CheckAmt
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String
    Dim StdYear As String
    Dim StdMonth As String
    Dim FuelType As String
    
    Dim FuelPrice As Double
    
    CLTCOD = Trim(oDS_PH_PY032A.GetValue("U_CLTCOD", 0)) '사업장
    StdYear = Left(Trim(oDS_PH_PY032A.GetValue("U_FrDate", 0)), 4)
    StdMonth = Mid(Trim(oDS_PH_PY032A.GetValue("U_FrDate", 0)), 5, 2)
    FuelType = Trim(oDS_PH_PY032A.GetValue("U_FuelType", 0)) '유류
    
    sQry = "           SELECT      T0.U_Year AS [StdYear],"
    sQry = sQry & "                T1.U_Month AS [StdMonth],"
    sQry = sQry & "                T1.U_Gasoline AS [Gasoline],"
    sQry = sQry & "                T1.U_Diesel AS [Diesel],"
    sQry = sQry & "                T1.U_LPG AS [LPG]"
    sQry = sQry & " FROM       [@PH_PY007A] AS T0"
    sQry = sQry & "                INNER JOIN"
    sQry = sQry & "                [@PH_PY007B] AS T1"
    sQry = sQry & "                    ON T0.Code = T1.Code"
    sQry = sQry & " WHERE      T0.U_CLTCOD = '" & CLTCOD & "'"
    sQry = sQry & "                AND T0.U_Year = '" & StdYear & "'"
    sQry = sQry & "                AND T1.U_Month = '" & StdMonth & "'"
    
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

    Call oDS_PH_PY032A.setValue("U_FuelPrc", 0, FuelPrice)
    Call oForm.Items("FuelPrc").CLICK

'    oForm.Items("FuelPrc").Specific.VALUE = FuelPrice

    Exit Sub
    
PH_PY032_GetFuelPrc_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY032_GetFuelPrc_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY032_CalculateFoodExp()
'******************************************************************************
'Function ID : PH_PY032_CalculateFoodExp()
'해당모듈 : PH_PY032
'기능 : 식비 계산
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY032_CalculateFoodExp_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim MSTCOD As String '사번
    Dim FoodNum As Integer '식수
    Dim FoodPrc As Double '당일식비
    Dim FoodExp As Double '전체식비
    
    If Trim(oDS_PH_PY032A.GetValue("U_MSTCOD", 0)) = "" And Trim(oDS_PH_PY032A.GetValue("U_FoodNum", 0)) <> "0" Then '사번을 선택하지 않으면
        ErrNum = 1
        GoTo PH_PY032_CalculateFoodExp_Error
    End If
    
    MSTCOD = Trim(oDS_PH_PY032A.GetValue("U_MSTCOD", 0)) '사번
    FoodNum = Trim(oDS_PH_PY032A.GetValue("U_FoodNum", 0)) '식수
    
    sQry = "            SELECT      T1.U_Num4 AS [FoodPrc]"
    sQry = sQry & "  FROM       [@PH_PY001A] AS T0"
    sQry = sQry & "                 LEFT JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.U_JIGCOD = T1.U_Code"
    sQry = sQry & "                     AND T1.Code = 'P232'"
    sQry = sQry & "                     AND T1.U_UseYN = 'Y'"
    sQry = sQry & "  WHERE      T0.Code = '" & MSTCOD & "'"

    Call oRecordSet.DoQuery(sQry)
    
    FoodPrc = oRecordSet.Fields("FoodPrc").VALUE
    FoodExp = FoodPrc * FoodNum
    
    Call oDS_PH_PY032A.setValue("U_FoodExp", 0, FoodExp)
    Call oForm.Items("FoodExp").CLICK

    Exit Sub
    
PH_PY032_CalculateFoodExp_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.SetStatusBarMessage "사원을 먼저 선택하세요.", bmt_Short, True
        Call oDS_PH_PY032A.setValue("U_FoodNum", 0, "0") '0식 선택
        Call oForm.Items("MSTCOD").CLICK
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY032_CalculateFoodExp_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub


Private Sub PH_PY032_Print_Report01()

    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY032_Print_Report01_Error
    
    Dim DocEntry As Integer
    Dim CLTCOD As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY032_Print_Report01_Error
    End If
    
    '//인자 MOVE , Trim 시키기..
    DocEntry = Trim(oForm.Items("DocEntry").Specific.VALUE)
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY032] 사용외출증"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PH_PY032_01.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PH_PY032_02.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PH_PY032_03.rpt"
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
    sQry = "EXEC [PH_PY032_90] '" & DocEntry & "'"
    
    oRecordSet.DoQuery sQry
    
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY032_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY032_Print_Report01_Error:

    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY032_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub



