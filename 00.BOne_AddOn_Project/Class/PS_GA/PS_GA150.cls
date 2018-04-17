VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_GA150.cls
'//  Module             : 총무관리>안전환경관리>안전보호구지급관리>안전보호구지급등록
'//  Desc               : 안전보호구지급 이력 등록
'//  FormType           : PS_GA150
'//  Create Date(Start) : 2012.09.18
'//  Create Date(End)   : 2012.09.18
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_GA150H  As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_GA150L  As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
'*******************************************************************

' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA150.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_GA150_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_GA150_CreateItems
    Call PS_GA150_ComboBox_Setting
    Call PS_GA150_CF_ChooseFromList
    Call PS_GA150_EnableMenus
    Call PS_GA150_SetDocument(oFromDocEntry01)
    Call PS_GA150_FormResize
    Call PS_GA150_FormReset '폼초기화 추가(2013.01.29 송명규)

'    Call PS_GA150_Add_MatrixRow(0, True)
    Call PS_GA150_LoadCaption
    Call PS_GA150_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False  '// 삭제
    oForm01.EnableMenu ("1286"), False  '// 닫기
    oForm01.EnableMenu ("1287"), False  '// 복제
    oForm01.EnableMenu ("1285"), False  '// 복원
    oForm01.EnableMenu ("1284"), False  '// 취소
    oForm01.EnableMenu ("1293"), False  '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_GA150H]"
    Call RecordSet01.DoQuery(sQry)
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_GA150H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_GA150H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
        
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing

    '지급일자 설정
    oForm01.Items("PrvdDt").Specific.VALUE = Format(Date, "YYYYMMDD")
    '등록일 설정
    oForm01.Items("RegDt").Specific.VALUE = Format(Date, "YYYYMMDD")
    '등록일(조회) 설정
    oForm01.Items("SRegDt").Specific.VALUE = Format(Date, "YYYYMMDD")
    
    Call oForm01.Items("CntcCode").Click
    
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA150_LoadCaption()
'******************************************************************************
'Function ID : PS_GA150_LoadCaption()
'해당모듈    : PS_GA150
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_LoadCaption_Error

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
    
PS_GA150_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub PS_GA150_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
'******************************************************************************
'Function ID : PS_GA150_Add_MatrixRow()
'해당모듈    : PS_GA150
'기능        : 메트릭스 행 추가
'인수        : oRow, RowInserted
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PS_GA150L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_GA150L.offset = oRow
    oDS_PS_GA150L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub

PS_GA150_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_GA150_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_GA150_MTX01()
'******************************************************************************
'Function ID : PS_GA150_MTX01()
'해당모듈    : PS_GA150
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_MTX01_Error

    Dim i      As Integer
    Dim sQry   As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim SBPLID    As String '사업장
    Dim SCntcCode As String '사번
    Dim SPrvdDtFr As String '지급일자(시작)
    Dim SPrvdDtTo As String '지급일자(종료)
    Dim SRegDt    As String '등록일
    Dim ChkMaxDt  As String '최근등록자료 조회
    
    SBPLID = IIf(Trim(oForm01.Items("SBPLId").Specific.VALUE) = "%", "", Trim(oForm01.Items("SBPLId").Specific.VALUE))
    SCntcCode = Trim(oForm01.Items("SCntcCode").Specific.VALUE)
    SPrvdDtFr = Trim(oForm01.Items("SPrvdDtFr").Specific.VALUE)
    SPrvdDtTo = Trim(oForm01.Items("SPrvdDtTo").Specific.VALUE)
    SRegDt = Trim(oForm01.Items("SRegDt").Specific.VALUE)
    
    If oForm01.DataSources.UserDataSources.Item("ChkMaxDt").VALUE = "Y" Then '최근등록자료 조회
        ChkMaxDt = "Y"
    Else
        ChkMaxDt = "N"
    End If
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    If ChkMaxDt = "N" Then
    
        sQry = "      EXEC [PS_GA150_01] '"
        sQry = sQry & SBPLID & "','"
        sQry = sQry & SCntcCode & "','"
        sQry = sQry & SPrvdDtFr & "','"
        sQry = sQry & SPrvdDtTo & "','"
        sQry = sQry & SRegDt & "'"
        
    Else
    
        sQry = "      EXEC [PS_GA150_03] '"
        sQry = sQry & SBPLID & "','"
        sQry = sQry & SCntcCode & "','"
        sQry = sQry & SPrvdDtFr & "','"
        sQry = sQry & SPrvdDtTo & "','"
        sQry = sQry & SRegDt & "'"
    
    End If
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_GA150L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_GA150_Add_MatrixRow(0, True)
        Call PS_GA150_LoadCaption
        
        GoTo PS_GA150_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_GA150L.Size Then
            oDS_PS_GA150L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_GA150L.offset = i
        
        Call oDS_PS_GA150L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_GA150L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PS_GA150L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("BPLId").VALUE)) '사업장
        Call oDS_PS_GA150L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("CntcCode").VALUE)) '사번
        Call oDS_PS_GA150L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("CntcName").VALUE)) '성명
        Call oDS_PS_GA150L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("TeamCd").VALUE)) '팀
        Call oDS_PS_GA150L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("TeamNm").VALUE)) '팀명
        Call oDS_PS_GA150L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("RspCd").VALUE)) '담당
        Call oDS_PS_GA150L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("RspNm").VALUE)) '담당명
        Call oDS_PS_GA150L.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("PrtrCd").VALUE)) '보호구코드
        Call oDS_PS_GA150L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("PrtrNm").VALUE)) '보호구명
        Call oDS_PS_GA150L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("PrtrSpec").VALUE)) '보호구규격
        Call oDS_PS_GA150L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("PrvdDt").VALUE), "yyyyMMdd")) '지급일자
        Call oDS_PS_GA150L.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("LPrvdDt").VALUE), "yyyyMMdd")) '전지급일
        Call oDS_PS_GA150L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Pred").VALUE)) '주기
        Call oDS_PS_GA150L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("Qty").VALUE)) '수량
        Call oDS_PS_GA150L.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("RtrnDt").VALUE)) '반납일자
        Call oDS_PS_GA150L.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("RtrnRsn").VALUE)) '반납사유
        Call oDS_PS_GA150L.setValue("U_ColTxt01", i, Trim(oRecordSet01.Fields("Note").VALUE)) '비고
        Call oDS_PS_GA150L.setValue("U_ColReg14", i, Format(Trim(oRecordSet01.Fields("RegDt").VALUE), "YYYYMMDD")) '등록일(2018.04.02 송명규 추가)
        
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

PS_GA150_MTX01_Error:

    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA150_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_GA150_DeleteData()
'******************************************************************************
'Function ID : PS_GA150_DeleteData()
'해당모듈    : PS_GA150
'기능        : 기본정보 삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_DeleteData_Error

    Dim i      As Integer
    Dim sQry   As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PS_GA150H] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_GA150_DeleteData_Error
        Else
            sQry = "DELETE FROM [@PS_GA150H] WHERE DocEntry = '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Exit Sub

PS_GA150_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA150_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_GA150_UpdateData() As Boolean
'******************************************************************************
'Function ID : PS_GA150_UpdateData()
'해당모듈    : PS_GA150
'기능        : 기본정보 수정
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_UpdateData_Error

    Dim i    As Integer
    Dim j    As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim BPLID    As String '사업장
    Dim CntcCode As String '사번
    Dim CntcName As String '성명
    Dim TeamCd   As String '팀
    Dim TeamNm   As String '팀명
    Dim RspCd    As String '담당
    Dim RspNm    As String '담당명
    Dim PrtrCd   As String '보호구코드
    Dim PrtrNm   As String '보호구명
    Dim PrtrSpec As String '보호구규격
    Dim PrvdDt   As String '지급일자
    Dim Pred     As Integer '주기
    Dim Qty      As Integer '수량
    Dim RtrnDt   As String '반납일자
    Dim RtrnRsn  As String '반납사유
    Dim Note     As String '비고
    Dim RegDt    As String '등록일

    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE) '사번
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE) '성명
    TeamCd = Trim(oForm01.Items("TeamCd").Specific.VALUE) '팀
    TeamNm = Trim(oForm01.Items("TeamNm").Specific.VALUE) '팀명
    RspCd = Trim(oForm01.Items("RspCd").Specific.VALUE) '담당
    RspNm = Trim(oForm01.Items("RspNm").Specific.VALUE) '담당명
    PrtrCd = Trim(oForm01.Items("PrtrCd").Specific.VALUE) '보호구코드
    PrtrNm = Trim(oForm01.Items("PrtrNm").Specific.VALUE) '보호구명
    PrtrSpec = Trim(oForm01.Items("PrtrSpec").Specific.VALUE) '보호구규격
    PrvdDt = Trim(oForm01.Items("PrvdDt").Specific.VALUE) '지급일자
    Pred = Trim(oForm01.Items("Pred").Specific.VALUE) '주기
    Qty = Trim(oForm01.Items("Qty").Specific.VALUE) '수량
    RtrnDt = Trim(oForm01.Items("RtrnDt").Specific.VALUE) '반납일자
    RtrnRsn = Trim(oForm01.Items("RtrnRsn").Specific.VALUE) '반납사유
    Note = Trim(oForm01.Items("Note").Specific.VALUE) '비고
    RegDt = Trim(oForm01.Items("RegDt").Specific.VALUE) '등록일
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
        PS_GA150_UpdateData = False
        Exit Function
    End If
    
    sQry = "        UPDATE   [@PS_GA150H]"
    sQry = sQry & " SET      U_BPLId = '" & BPLID & "'," '사업장
    sQry = sQry & "          U_CntcCode = '" & CntcCode & "'," '사번
    sQry = sQry & "          U_CntcName = '" & CntcName & "',"  '성명
    sQry = sQry & "          U_TeamCd = '" & TeamCd & "',"  '팀
    sQry = sQry & "          U_TeamNm = '" & TeamNm & "',"  '팀명
    sQry = sQry & "          U_RspCd = '" & RspCd & "',"  '담당
    sQry = sQry & "          U_RspNm = '" & RspNm & "',"  '담당명
    sQry = sQry & "          U_PrtrCd = '" & PrtrCd & "',"  '보호구코드
    sQry = sQry & "          U_PrtrNm = '" & PrtrNm & "',"  '보호구명
    sQry = sQry & "          U_PrtrSpec = '" & PrtrSpec & "',"  '보호구규격
    sQry = sQry & "          U_PrvdDt = '" & PrvdDt & "',"  '지급일자
    sQry = sQry & "          U_Pred = '" & Pred & "',"  '주기
    sQry = sQry & "          U_Qty = '" & Qty & "',"  '수량
    sQry = sQry & "          U_RtrnDt = '" & RtrnDt & "',"  '반납일자
    sQry = sQry & "          U_RtrnRsn = '" & RtrnRsn & "',"  '반납사유
    sQry = sQry & "          U_Note = '" & Note & "',"  '비고
    sQry = sQry & "          U_RegDt = '" & RegDt & "'"  '등록일
    sQry = sQry & " WHERE    DocEntry = '" & DocEntry & "'"
    
    Call RecordSet01.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PS_GA150_UpdateData = True
    Exit Function

PS_GA150_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA150_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_GA150_AddData() As Boolean
'******************************************************************************
'Function ID : PS_GA150_AddData()
'해당모듈    : PS_GA150
'기능        : 데이터 INSERT
'인수        : 없음
'반환값      : 성공여부
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_AddData_Error

    Dim i    As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim BPLID    As String '사업장
    Dim CntcCode As String '사번
    Dim CntcName As String '성명
    Dim TeamCd   As String '팀
    Dim TeamNm   As String '팀명
    Dim RspCd    As String '담당
    Dim RspNm    As String '담당명
    Dim PrtrCd   As String '보호구코드
    Dim PrtrNm   As String '보호구명
    Dim PrtrSpec As String '보호구규격
    Dim PrvdDt   As String '지급일자
    Dim LPrvdDt  As String '전지급일
    Dim Pred     As Integer '주기
    Dim Qty      As Integer '수량
    Dim RtrnDt   As String '반납일자
    Dim RtrnRsn  As String '반납사유
    Dim Note     As String '비고
    Dim RegDt    As String '등록일

'    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE) '사번
    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE) '성명
    TeamCd = Trim(oForm01.Items("TeamCd").Specific.VALUE) '팀
    TeamNm = Trim(oForm01.Items("TeamNm").Specific.VALUE) '팀명
    RspCd = Trim(oForm01.Items("RspCd").Specific.VALUE) '담당
    RspNm = Trim(oForm01.Items("RspNm").Specific.VALUE) '담당명
    PrtrCd = Trim(oForm01.Items("PrtrCd").Specific.VALUE) '보호구코드
    PrtrNm = Trim(oForm01.Items("PrtrNm").Specific.VALUE) '보호구명
    PrtrSpec = Trim(oForm01.Items("PrtrSpec").Specific.VALUE) '보호구규격
    PrvdDt = Trim(oForm01.Items("PrvdDt").Specific.VALUE) '지급일자
    LPrvdDt = Trim(oForm01.Items("LPrvdDt").Specific.VALUE) '전지급일
    Pred = Trim(oForm01.Items("Pred").Specific.VALUE) '주기
    Qty = Trim(oForm01.Items("Qty").Specific.VALUE) '수량
    RtrnDt = Trim(oForm01.Items("RtrnDt").Specific.VALUE) '반납일자
    RtrnRsn = Trim(oForm01.Items("RtrnRsn").Specific.VALUE) '반납사유
    Note = Trim(oForm01.Items("Note").Specific.VALUE) '비고
    RegDt = Trim(oForm01.Items("RegDt").Specific.VALUE) '등록일

    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_GA150H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "        INSERT INTO [@PS_GA150H]"
    sQry = sQry & " ("
    sQry = sQry & "     DocEntry,"
    sQry = sQry & "     DocNum,"
    
    sQry = sQry & "     U_BPLId," '사업장"
    sQry = sQry & "     U_CntcCode," '사번"
    sQry = sQry & "     U_CntcName," '성명"
    sQry = sQry & "     U_TeamCd," '팀"
    sQry = sQry & "     U_TeamNm," '팀명"
    sQry = sQry & "     U_RspCd," '담당"
    sQry = sQry & "     U_RspNm," '담당명"
    sQry = sQry & "     U_PrtrCd," '보호구코드"
    sQry = sQry & "     U_PrtrNm," '보호구명"
    sQry = sQry & "     U_PrtrSpec," '보호구규격"
    sQry = sQry & "     U_PrvdDt," '지급일자"
    sQry = sQry & "     U_LPrvdDt," '전지급일"
    sQry = sQry & "     U_Pred," '주기"
    sQry = sQry & "     U_Qty," '수량"
    sQry = sQry & "     U_RtrnDt," '반납일자"
    sQry = sQry & "     U_RtrnRsn," '반납사유"
    sQry = sQry & "     U_Note," '비고"
    sQry = sQry & "     U_RegDt" '등록일
    
    sQry = sQry & " )"
    sQry = sQry & " VALUES"
    sQry = sQry & " ("
    sQry = sQry & DocEntry & ","
    sQry = sQry & DocEntry & ","
    
    sQry = sQry & "'" & BPLID & "'," '사업장"
    sQry = sQry & "'" & CntcCode & "'," '사번"
    sQry = sQry & "'" & CntcName & "'," '성명"
    sQry = sQry & "'" & TeamCd & "'," '팀"
    sQry = sQry & "'" & TeamNm & "'," '팀명"
    sQry = sQry & "'" & RspCd & "'," '담당"
    sQry = sQry & "'" & RspNm & "'," '담당명"
    sQry = sQry & "'" & PrtrCd & "'," '보호구코드"
    sQry = sQry & "'" & PrtrNm & "'," '보호구명"
    sQry = sQry & "'" & PrtrSpec & "'," '보호구규격"
    sQry = sQry & "'" & PrvdDt & "'," '지급일자"
    sQry = sQry & "'" & LPrvdDt & "'," '전지급일"
    sQry = sQry & "'" & Pred & "'," '주기"
    sQry = sQry & "'" & Qty & "'," '수량"
    sQry = sQry & "'" & RtrnDt & "'," '반납일자"
    sQry = sQry & "'" & RtrnRsn & "'," '반납사유"
    sQry = sQry & "'" & Note & "'," '비고"
    sQry = sQry & "'" & RegDt & "'" '등록일
    
    sQry = sQry & ")"
    
    Call RecordSet02.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_GA150_AddData = True

    Exit Function

PS_GA150_AddData_Error:
    PS_GA150_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA150_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_GA150_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_GA150_HeaderSpaceLineDel()
'해당모듈    : PS_GA150
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0
    
    If Trim(oForm01.Items("CntcCode").Specific.VALUE) = "" Then '사번
        ErrNum = 1
        GoTo PS_GA150_HeaderSpaceLineDel_Error
    ElseIf Trim(oForm01.Items("PrtrCd").Specific.VALUE) = "" Then '보호구코드
        ErrNum = 2
        GoTo PS_GA150_HeaderSpaceLineDel_Error
    ElseIf Trim(oForm01.Items("PrvdDt").Specific.VALUE) = "" Then '지급일자
        ErrNum = 3
        GoTo PS_GA150_HeaderSpaceLineDel_Error
    ElseIf Trim(oForm01.Items("Qty").Specific.VALUE) = 0 Then '수량
        ErrNum = 4
        GoTo PS_GA150_HeaderSpaceLineDel_Error
    End If

    PS_GA150_HeaderSpaceLineDel = True
    
    Exit Function

PS_GA150_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사번은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("CntcCode").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "보호구코드는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("PrtrCd").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "지급일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("PrvdDt").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "수량은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Qty").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "부자재청구서 출력시에는 등록일을 필수로 입력해야합니다. 확인하세요.", "E"
        Call oForm01.Items("SRegDt").Click(ct_Regular)
    End If
    PS_GA150_HeaderSpaceLineDel = False
End Function

Private Function PS_GA150_RegDtCheck() As Boolean
'******************************************************************************
'Function ID : PS_GA150_RegDtCheck()
'해당모듈    : PS_GA150
'기능        : 부자재 청구서 출력 시 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_RegDtCheck_Error

    Dim ErrNum As Integer
    ErrNum = 0
    
    If Trim(oForm01.Items("SRegDt").Specific.VALUE) = "" Then '등록일
        ErrNum = 1
        GoTo PS_GA150_RegDtCheck_Error
    End If

    PS_GA150_RegDtCheck = True
    
    Exit Function

PS_GA150_RegDtCheck_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "부자재청구서 출력시에는 등록일을 필수로 입력해야합니다. 확인하세요.", "E"
        Call oForm01.Items("SRegDt").Click(ct_Regular)
    End If
    PS_GA150_RegDtCheck = False
End Function


'/ 메트릭스 필수 사항 check
Private Function PS_GA150_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_GA150_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PS_GA150_MatrixSpaceLineDel = True
    
    Exit Function

PS_GA150_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PS_GA150_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_GA150_MatrixSpaceLineDel = False
End Function

Private Sub PS_GA150_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_GA150_FlushToItemValue_Error

    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim sQry     As String
    Dim ItemCode As String
    Dim Qty      As Integer
    Dim Calculate_Weight As Double
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Select Case oUID
'
'        Case "BOSHour"
'
'    End Select

    Set oRecordSet01 = Nothing

    Exit Sub
    
PS_GA150_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_GA150_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PS_GA150_FormItemEnabled()
On Error GoTo PS_GA150_FormItemEnabled_Error

    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    
    Exit Sub

PS_GA150_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_GA150_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
                Call PS_GA150_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_GA150_LoadCaption
                
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
                '//Call PS_GA150_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_GA150H.Clear

'                Call PS_GA150_LoadCaption
'                Call PS_GA150_FormItemEnabled
                '//Call PS_GA150_FormItemEnabled '//UDO방식
                '//Call PS_GA150_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_GA150_FormItemEnabled
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
    
        If pval.ItemUID = "PS_GA150" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_GA150_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If PS_GA150_DataCheck() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                If PS_GA150_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_GA150_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA150_LoadCaption
                Call PS_GA150_MTX01
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_GA150_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                'UPDATE 모드인경우는 DataCheck(중복 체크)를 할 필요가 없음
'                If PS_GA150_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                If PS_GA150_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_GA150_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA150_LoadCaption
                Call PS_GA150_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
            Call PS_GA150_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_GA150_LoadCaption
            Call PS_GA150_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_GA150_DeleteData
                Call PS_GA150_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_GA150_LoadCaption
                Call PS_GA150_MTX01
                
            Else
            
            End If
            
        ElseIf pval.ItemUID = "BtnPrt" Then
            
            If PS_GA150_RegDtCheck = False Then
                BubbleEvent = False
                Exit Sub
            End If
            
            Call PS_GA150_Print_Report01
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA150" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "PrtrCd", "") '기본정보-보호구코드
        
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

                'DataSource를 이용하여 각 컨트롤에 값을 출력
                Call oDS_PS_GA150H.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
                Call oDS_PS_GA150H.setValue("U_BPLId", 0, oMat01.Columns("BPLId").Cells(pval.Row).Specific.VALUE) '사업장
                Call oDS_PS_GA150H.setValue("U_CntcCode", 0, oMat01.Columns("CntcCode").Cells(pval.Row).Specific.VALUE) '사번
                Call oDS_PS_GA150H.setValue("U_CntcName", 0, oMat01.Columns("CntcName").Cells(pval.Row).Specific.VALUE) '성명
                Call oDS_PS_GA150H.setValue("U_TeamCd", 0, oMat01.Columns("TeamCd").Cells(pval.Row).Specific.VALUE) '팀
                Call oDS_PS_GA150H.setValue("U_TeamNm", 0, oMat01.Columns("TeamNm").Cells(pval.Row).Specific.VALUE) '팀명
                Call oDS_PS_GA150H.setValue("U_RspCd", 0, oMat01.Columns("RspCd").Cells(pval.Row).Specific.VALUE) '담당
                Call oDS_PS_GA150H.setValue("U_RspNm", 0, oMat01.Columns("RspNm").Cells(pval.Row).Specific.VALUE) '담당명
                Call oDS_PS_GA150H.setValue("U_PrtrCd", 0, oMat01.Columns("PrtrCd").Cells(pval.Row).Specific.VALUE) '보호구코드
                Call oDS_PS_GA150H.setValue("U_PrtrNm", 0, oMat01.Columns("PrtrNm").Cells(pval.Row).Specific.VALUE) '보호구명
                Call oDS_PS_GA150H.setValue("U_PrtrSpec", 0, oMat01.Columns("PrtrSpec").Cells(pval.Row).Specific.VALUE) '보호구규격
                Call oDS_PS_GA150H.setValue("U_PrvdDt", 0, oMat01.Columns("PrvdDt").Cells(pval.Row).Specific.VALUE) '지급일자
                Call oDS_PS_GA150H.setValue("U_LPrvdDt", 0, oMat01.Columns("LPrvdDt").Cells(pval.Row).Specific.VALUE) '전지급일
                Call oDS_PS_GA150H.setValue("U_Pred", 0, oMat01.Columns("Pred").Cells(pval.Row).Specific.VALUE) '주기
                Call oDS_PS_GA150H.setValue("U_Qty", 0, oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) '수량
                Call oDS_PS_GA150H.setValue("U_RtrnDt", 0, Replace(oMat01.Columns("RtrnDt").Cells(pval.Row).Specific.VALUE, ".", "")) '반납일자
                Call oDS_PS_GA150H.setValue("U_RtrnRsn", 0, oMat01.Columns("RtrnRsn").Cells(pval.Row).Specific.VALUE)  '반납사유
                Call oDS_PS_GA150H.setValue("U_Note", 0, oMat01.Columns("Note").Cells(pval.Row).Specific.VALUE) '비고
                Call oDS_PS_GA150H.setValue("U_RegDt", 0, Replace(oMat01.Columns("RegDt").Cells(pval.Row).Specific.VALUE, ".", "")) '등록일
                
                oForm01.Mode = fm_UPDATE_MODE
                Call PS_GA150_LoadCaption
                
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
'                    Call oDS_PS_GA150L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_GA150L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_GA150_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_GA150L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PS_GA150_FlushToItemValue(pval.ItemUID)
                
                If pval.ItemUID = "CntcCode" Then
                
                    oForm01.Items("CntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '성명
                    oForm01.Items("TeamCd").Specific.VALUE = MDC_GetData.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '팀코드
                    oForm01.Items("RspCd").Specific.VALUE = MDC_GetData.Get_ReData("U_RspCode", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '담당코드
                    
                    oForm01.Items("TeamNm").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("TeamCd").Specific.VALUE & "'")  '팀명
                    oForm01.Items("RspNm").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("RspCd").Specific.VALUE & "'")  '담당명
                    
                    oForm01.Items("LPrvdDt").Specific.VALUE = Format(MDC_GetData.Get_ReData("MAX(U_PrvdDt)", "U_CntcCode", "[@PS_GA150H]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'", "AND U_PrtrNm = '" & oForm01.Items("PrtrNm").Specific.VALUE & "'"), "YYYYMMDD") '전지급일
                    
                ElseIf pval.ItemUID = "PrtrCd" Then
                
                    oForm01.Items("PrtrNm").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items("PrtrCd").Specific.VALUE & "'")  '보호구명
                    oForm01.Items("PrtrSpec").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" & oForm01.Items("PrtrCd").Specific.VALUE & "'")  '규격
                    
                    oForm01.Items("LPrvdDt").Specific.VALUE = Format(MDC_GetData.Get_ReData("MAX(U_PrvdDt)", "U_CntcCode", "[@PS_GA150H]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'", "AND U_PrtrNm = '" & oForm01.Items("PrtrNm").Specific.VALUE & "'"), "YYYYMMDD") '전지급일
                    
                ElseIf pval.ItemUID = "SCntcCode" Then
                
                    oForm01.Items("SCntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & oForm01.Items("SCntcCode").Specific.VALUE & "'") '성명
                
                End If
                
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
        Call PS_GA150_FormItemEnabled
        '//Call PS_GA150_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_GA150_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_GA150H", "U_CardCode,U_CardName")
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
'            If (PS_GA150_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_GA150H.RemoveRecord(oDS_PS_GA150H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_GA150_Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_GA150H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_GA150_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_GA150_CreateItems() As Boolean
On Error GoTo PS_GA150_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_GA150H = oForm01.DataSources.DBDataSources("@PS_GA150H")
    Set oDS_PS_GA150L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("SBPLId", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SBPLId").Specific.DataBind.SetBound(True, "", "SBPLId")
    '사업장_E

    '사번_S
    Call oForm01.DataSources.UserDataSources.Add("SCntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SCntcCode").Specific.DataBind.SetBound(True, "", "SCntcCode")
    '사번_E
    
    '성명_S
    Call oForm01.DataSources.UserDataSources.Add("SCntcName", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SCntcName").Specific.DataBind.SetBound(True, "", "SCntcName")
    '성명_E
    
    '지급일자(시작)_S
    Call oForm01.DataSources.UserDataSources.Add("SPrvdDtFr", dt_DATE)
    Call oForm01.Items("SPrvdDtFr").Specific.DataBind.SetBound(True, "", "SPrvdDtFr")
    '지급일자(시작)_E
    
    '지급일자(종료)_S
    Call oForm01.DataSources.UserDataSources.Add("SPrvdDtTo", dt_DATE)
    Call oForm01.Items("SPrvdDtTo").Specific.DataBind.SetBound(True, "", "SPrvdDtTo")
    '지급일자(종료)_E
    
    '등록일_S
    Call oForm01.DataSources.UserDataSources.Add("SRegDt", dt_DATE)
    Call oForm01.Items("SRegDt").Specific.DataBind.SetBound(True, "", "SRegDt")
    '등록일_E
    
    '최근등록자료 조회_S
    Call oForm01.DataSources.UserDataSources.Add("ChkMaxDt", dt_SHORT_TEXT, 1)
    Call oForm01.Items("ChkMaxDt").Specific.DataBind.SetBound(True, "", "ChkMaxDt")
    oForm01.Items("ChkMaxDt").Specific.Checked = False
    '최근등록자료 조회_E
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
    
PS_GA150_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA150_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_GA150_ComboBox_Setting()
On Error GoTo PS_GA150_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '기본정보-사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("BPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '조회조건-사업장
    Call oForm01.Items("SBPLId").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("SBPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '매트릭스-사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
    
PS_GA150_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA150_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA150_CF_ChooseFromList()
On Error GoTo PS_GA150_CF_ChooseFromList_Error

    Exit Sub
PS_GA150_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA150_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA150_EnableMenus()
On Error GoTo PS_GA150_EnableMenus_Error
    
    Exit Sub
PS_GA150_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA150_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA150_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_GA150_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_GA150_FormItemEnabled
        '//Call PS_GA150_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_GA150_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_GA150_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA150_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA150_FormResize()
On Error GoTo PS_GA150_FormResize_Error
    
    '그룹박스 크기 동적 할당
    oForm01.Items("GrpBox01").Height = 190
    oForm01.Items("GrpBox01").Width = 450 'oForm01.Items("Mat01").Width / 2 + 30
    
    oForm01.Items("GrpBox02").Height = 190
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_GA150_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA150_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_GA150_FormReset()
'******************************************************************************
'Function ID : PS_GA150_FormReset()
'해당모듈    : PS_GA150
'기능        : 화면 초기화
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_GA150H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_GA150H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_GA150H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    Call oDS_PS_GA150H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId) '사업장
    Call oDS_PS_GA150H.setValue("U_CntcCode", 0, "") '사번
    Call oDS_PS_GA150H.setValue("U_CntcName", 0, "") '성명
    Call oDS_PS_GA150H.setValue("U_TeamCd", 0, "") '팀
    Call oDS_PS_GA150H.setValue("U_TeamNm", 0, "") '팀명
    Call oDS_PS_GA150H.setValue("U_RspCd", 0, "") '담당
    Call oDS_PS_GA150H.setValue("U_RspNm", 0, "") '담당명
    Call oDS_PS_GA150H.setValue("U_PrtrCd", 0, "") '보호구코드
    Call oDS_PS_GA150H.setValue("U_PrtrNm", 0, "") '보호구명
    Call oDS_PS_GA150H.setValue("U_PrtrSpec", 0, "") '보호구규격
    Call oDS_PS_GA150H.setValue("U_PrvdDt", 0, Format(Now, "yyyyMMdd")) '지급일자
    Call oDS_PS_GA150H.setValue("U_LPrvdDt", 0, "") '전지급일
    Call oDS_PS_GA150H.setValue("U_Pred", 0, "") '주기
    Call oDS_PS_GA150H.setValue("U_Qty", 0, 1) '수량
    Call oDS_PS_GA150H.setValue("U_RtrnDt", 0, "") '반납일자
    Call oDS_PS_GA150H.setValue("U_RtrnRsn", 0, "") '반납사유
    Call oDS_PS_GA150H.setValue("U_Note", 0, "") '비고
    Call oDS_PS_GA150H.setValue("U_RegDt", 0, Format(Now, "YYYYMMDD")) '등록일
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

    Exit Sub

PS_GA150_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA150_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function PS_GA150_DataCheck() As Boolean
'******************************************************************************
'Function ID : PS_GA150_DataCheck()
'해당모듈    : PS_GA150
'기능        : 중복데이터 체크
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA150_DataCheck_Error

    PS_GA150_DataCheck = False
    
    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID    As String
    Dim CntcCode As String
    Dim PrvdDt   As String
    Dim ItemCode As String
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    PrvdDt = oForm01.Items("PrvdDt").Specific.VALUE
    ItemCode = oForm01.Items("PrtrCd").Specific.VALUE

    sQry = "      EXEC PS_GA150_02 '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & CntcCode & "','"
    sQry = sQry & PrvdDt & "','"
    sQry = sQry & ItemCode & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    If oRecordSet01.Fields("ReturnValue").VALUE = "True" Then
    
        PS_GA150_DataCheck = True
    
    Else
    
        ErrNum = 1
        GoTo PS_GA150_DataCheck_Error
    
    End If
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA150_DataCheck_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "이미입력된 데이터입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_GA150_DataCheck_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Set oRecordSet01 = Nothing
    PS_GA150_DataCheck = False
End Function

Private Sub PS_GA150_Print_Report01()
On Error GoTo Print_Query_Error

    Dim DocNum     As String
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String

    Dim i As Integer
    Dim ErrNum     As Integer
    Dim Sub_sQry   As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim BPLID   As String '사업장
    Dim RegDt   As String '등록일

    BPLID = Trim(oForm01.Items("SBPLId").Specific.Selected.VALUE) '사업장
    RegDt = Trim(oForm01.Items("SRegDt").Specific.VALUE) '등록일

    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_GA150] 레포트"
    
    ReportName = "PS_GA150_01.rpt"

    sQry = "      EXEC PS_GA150_51 '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & RegDt & "'"
    
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

'    Call oRecordSet.DoQuery(sQry)
'
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

Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_GA150_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

