VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP564"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_PP564.cls
'//  Module             : 생산관리 > 애드온 > 장비이익률관리(기계)
'//  Desc               : 장비별 수주금액 수정관리(영업)
'//  FormType           : PS_PP564
'//  Create Date(Start) : 2017.03.30
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
'Private oDS_PS_PP564A As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP564B  As SAPbouiCOM.DBDataSource '등록라인

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
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP564.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP564_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP564_CreateItems
    Call PS_PP564_ComboBox_Setting
    Call PS_PP564_CF_ChooseFromList
    Call PS_PP564_EnableMenus
    Call PS_PP564_SetDocument(oFromDocEntry01)
    Call PS_PP564_FormResize

    Call PS_PP564_LoadCaption
    Call PS_PP564_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

    Call PS_PP564_FormReset '폼초기화
    
    oForm01.Items("BaseEntry").Visible = False 'BaseEntry 비활성
    oForm01.Items("BaseLine").Visible = False 'BaseLine 비활성
    
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

Private Sub PS_PP564_LoadCaption()
'******************************************************************************
'Function ID : PS_PP564_LoadCaption()
'해당모듈    : PS_PP564
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP564_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "추가"
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
    End If

    Call oForm01.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP564_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_PP564_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP564_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP564B.InsertRecord (oRow)
    End If

    oMat01.AddRow
    oDS_PS_PP564B.Offset = oRow
    oDS_PS_PP564B.setValue "U_LineNum", oRow, oRow + 1

    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP564_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_PP564_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_PP564_MTX01()
'******************************************************************************
'Function ID : PS_PP564_MTX01()
'해당모듈    : PS_PP564
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP564_MTX01_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As Integer

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim SAmtCls  As String '비용구분
    Dim SOrdNum  As String '작번
    Dim SOrdSub1 As String '서브작번1
    Dim SOrdSub2 As String '서브작번2
    Dim SComment As String '비고(%)
    
    SAmtCls = oForm01.DataSources.UserDataSources("SAmtCls").VALUE '비용구분
    SOrdNum = oForm01.DataSources.UserDataSources("SOrdNum").VALUE '작번
    SOrdSub1 = oForm01.DataSources.UserDataSources("SOrdSub1").VALUE '서브작번1
    SOrdSub2 = oForm01.DataSources.UserDataSources("SOrdSub2").VALUE '서브작번2
    SComment = oForm01.DataSources.UserDataSources("SComment").VALUE '비고(%)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)

    Call oForm01.Freeze(True)

    sQry = "      EXEC [PS_PP564_01] '"
    sQry = sQry & SAmtCls & "','" '비용구분
    sQry = sQry & SOrdNum & "','" '작번
    sQry = sQry & SOrdSub1 & "','" '서브작번1
    sQry = sQry & SOrdSub2 & "','" '서브작번2
    sQry = sQry & SComment & "'" '비고
    
    Call oRecordSet01.DoQuery(sQry)

    oMat01.Clear
    oDS_PS_PP564B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then

        ErrNum = 1

        oForm01.Mode = fm_ADD_MODE

'        Call PS_PP564_Add_MatrixRow(0, True)
        Call PS_PP564_LoadCaption

        GoTo PS_PP564_MTX01_Error

        Exit Sub
    End If

    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP564B.Size Then
            oDS_PS_PP564B.InsertRecord (i)
        End If

        oMat01.AddRow
        oDS_PS_PP564B.Offset = i

        Call oDS_PS_PP564B.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP564B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("AmtCls").VALUE)) '비용구분
        Call oDS_PS_PP564B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작지번호
        Call oDS_PS_PP564B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("OrdSub1").VALUE)) '서브작번1
        Call oDS_PS_PP564B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("OrdSub2").VALUE)) '서브작번2
        Call oDS_PS_PP564B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("OrdName").VALUE)) '작지명
        Call oDS_PS_PP564B.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("Amount").VALUE)) '목표금액
        Call oDS_PS_PP564B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("Comment").VALUE)) '비고
        Call oDS_PS_PP564B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("UpdateUser").VALUE)) '수정자
        Call oDS_PS_PP564B.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("UpdateDate").VALUE)) '수정일
        Call oDS_PS_PP564B.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("UpdateTime").VALUE)) '수정시간
        Call oDS_PS_PP564B.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("BaseDL").VALUE)) '기준문서
        Call oDS_PS_PP564B.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("MainOrd").VALUE)) '메인작번
        
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
PS_PP564_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing

    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP564_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP564_DeleteData()
'******************************************************************************
'Function ID : PS_PP564_DeleteData()
'해당모듈    : PS_PP564
'기능        : 기본정보 삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : Delete 사용 안함
'******************************************************************************
On Error GoTo PS_PP564_DeleteData_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As Integer

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim DocEntry As String

    If oForm01.Mode = fm_UPDATE_MODE Then

        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)

        sQry = "SELECT COUNT(*) FROM [Z_PS_PP560_01] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)

        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_PP564_DeleteData_Error
        Else
            sQry = "EXEC PS_PP564_04 '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If

    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")

'    Call PS_PP564_FormReset

'    oForm01.Mode = fm_ADD_MODE

'    Call oForm01.Items("BtnSearch").Click(ct_Regular)

'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PS_PP564_Add_MatrixRow(0, True)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP564_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP564_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_PP564_UpdateData() As Boolean
'******************************************************************************
'Function ID : PS_PP564_UpdateData()
'해당모듈    : PS_PP564
'기능        : 기본정보를 수정
'인수        : 없음
'반환값      : 없음
'특이사항    : Update 사용 안함
'******************************************************************************
On Error GoTo PS_PP564_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

'    Dim DocEntry As Integer '관리번호(PK)
'    Dim BPLID As String '사업장코드(PK)
'    Dim RegNo As String '등록번호(PK)
'    Dim RegDate As String '등록일자
'    Dim CntcCode As String '개선자사번
'    Dim CntcName As String '개선자성명
'    Dim TeamCode As String '소속팀코드
'    Dim RspCode As String '소속담당코드
'    Dim ClsCode As String '소속반코드
'    Dim Fields As String '분야
'    Dim Grade As String '등급
'    Dim Title As String '제목
'    Dim Comment As String '비고
'
'    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE) '관리번호
'    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
'    RegNo = Trim(oForm01.Items("RegNo").Specific.VALUE) '등록번호
'    RegDate = Trim(oForm01.Items("RegDate").Specific.VALUE) '등록일자
'    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE) '개선자사번
'    CntcName = Trim(oForm01.Items("CntcName").Specific.VALUE) '개선자성명
'    TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE) '소속팀코드
'    RspCode = Trim(oForm01.Items("RspCode").Specific.VALUE) '소속담당코드
'    ClsCode = Trim(oForm01.Items("ClsCode").Specific.VALUE) '소속반코드
'    Fields = Trim(oForm01.Items("Fields").Specific.VALUE) '분야
'    Grade = Trim(oForm01.Items("Grade").Specific.VALUE) '등급
'    Title = Trim(oForm01.Items("Title").Specific.VALUE) '제목
'    Comment = Trim(oForm01.Items("Comment").Specific.VALUE) '비고
'
'    sQry = "      EXEC [PS_PP564_03] '"
'    sQry = sQry & DocEntry & "','" '관리번호
'    sQry = sQry & BPLID & "','"  '사업장
'    sQry = sQry & RegNo & "','"  '등록번호
'    sQry = sQry & RegDate & "','"  '등록일자
'    sQry = sQry & CntcCode & "','"  '사원번호
'    sQry = sQry & CntcName & "','"  '사원성명
'    sQry = sQry & TeamCode & "','"  '팀
'    sQry = sQry & RspCode & "','"  '담당
'    sQry = sQry & ClsCode & "','"  '반
'    sQry = sQry & Fields & "','"  '분야
'    sQry = sQry & Grade & "','"  '등급
'    sQry = sQry & Title & "','"  '제목
'    sQry = sQry & Comment & "'"  '비고

    Call RecordSet01.DoQuery(sQry)

    MDC_Com.MDC_GF_Message "수정 완료!", "S"

    Set RecordSet01 = Nothing
    PS_PP564_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP564_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP564_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_PP564_AddData() As Boolean
'******************************************************************************
'Function ID : PS_PP564_AddData()
'해당모듈    : PS_PP564
'기능        : 데이터 INSERT
'인수        : 없음
'반환값      : 성공여부
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP564_AddData_Error

    Dim i    As Integer
    Dim sQry As String

    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim AmtCls       As String '비용구분
    Dim OrdNum       As String '작번
    Dim OrdSub1      As String '서브작번1
    Dim OrdSub2      As String '서브작번2
    Dim OrdName      As String '작지명
    Dim Amount       As Currency '금액
    Dim Comment      As String '비고
    Dim UpdateUserCD As String '등록자 사번
    'Dim UpdateUserNM As String '등록자 성명
    'Dim UpdateDate   As String 'UPDATE DATE
    Dim DocCls       As String 'DocCls
    Dim BaseEntry    As String '기준문서번호
    Dim BaseLine     As String '기준문서라인번호

    AmtCls = Trim(oForm01.Items("AmtCls").Specific.VALUE) '비용구분
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '작번
    OrdSub1 = Trim(oForm01.Items("OrdSub1").Specific.VALUE) '서브작번1
    OrdSub2 = Trim(oForm01.Items("OrdSub2").Specific.VALUE) '서브작번2
    OrdName = Trim(oForm01.Items("OrdName").Specific.VALUE) '작지명
    Amount = Trim(oForm01.Items("Amount").Specific.VALUE) '금액
    Comment = Trim(oForm01.Items("Comment").Specific.VALUE) '비고
    UpdateUserCD = MDC_PS_Common.User_MSTCOD '등록자사번
    '등록자성명 은 프로시저에서 조회
    'UPDATE DATE 은 프로시저에서 선택(GETDATE()이용)
    DocCls = "PP564" 'DocCls (상수)
    BaseEntry = Trim(oForm01.Items("BaseEntry").Specific.VALUE) '기준문서번호
    BaseLine = Trim(oForm01.Items("BaseLine").Specific.VALUE) '기준문서라인번호
    
    sQry = "      EXEC [PS_PP564_02] '"
    sQry = sQry & AmtCls & "','" '비용구분
    sQry = sQry & OrdNum & "','" '작번
    sQry = sQry & OrdSub1 & "','" '서브작번1
    sQry = sQry & OrdSub2 & "','" '서브작번2
    sQry = sQry & OrdName & "','" '작지명
    sQry = sQry & Amount & "','" '금액
    sQry = sQry & Comment & "','" '비고
    sQry = sQry & UpdateUserCD & "','" '등록자사번
    sQry = sQry & DocCls & "','" '자료타입
    sQry = sQry & BaseEntry & "','" '기준문서번호
    sQry = sQry & BaseLine & "'" '기준문서라인번호
    
    Call RecordSet01.DoQuery(sQry)

    MDC_Com.MDC_GF_Message "등록 완료!", "S"

    Set RecordSet01 = Nothing
    
    PS_PP564_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP564_AddData_Error:
    PS_PP564_AddData = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP564_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_PP564_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_PP564_HeaderSpaceLineDel()
'해당모듈    : PS_PP564
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP564_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("AmtCls").Specific.VALUE) = "%" '비용구분
            ErrNum = 1
            GoTo PS_PP564_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("AmtCls").Specific.VALUE = "01") And Trim(oForm01.Items("OrdNum").Specific.VALUE) = "" '일반 & 작번
            ErrNum = 2
            GoTo PS_PP564_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("Amount").Specific.VALUE) = "" '목표금액
            ErrNum = 3
            GoTo PS_PP564_HeaderSpaceLineDel_Error
    End Select

    PS_PP564_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP564_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "비용구분은 필수사항입니다. 확인하세요.", "E"
        'Call oForm01.Items("AmtCls").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "작번은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("OrdNum").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "금액은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("Amount").Click(ct_Regular)
    End If
    PS_PP564_HeaderSpaceLineDel = False
End Function

''/ 메트릭스 필수 사항 check
'Private Function PS_PP564_MatrixSpaceLineDel() As Boolean
'On Error GoTo PS_PP564_MatrixSpaceLineDel_Error
'
'    Dim i As Long
'    Dim ErrNum As Integer
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Dim sQry As String
'
'    Set oRecordSet01 = Nothing
'    PS_PP564_MatrixSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_PP564_MatrixSpaceLineDel_Error:
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
'        MDC_Com.MDC_GF_Message "PS_PP564_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'    PS_PP564_MatrixSpaceLineDel = False
'End Function

Private Sub PS_PP564_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_PP564_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim ErrNum    As Integer
    Dim sQry      As String
    Dim ItemCode  As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID

        Case "BaseDL"
        
            sQry = "        SELECT      T0.U_AmtCls AS [AmtCls],"
            sQry = sQry & "             T0.U_OrdNum AS [OrdNum],"
            sQry = sQry & "             T0.U_OrdSub1 AS [OrdSub1],"
            sQry = sQry & "             T0.U_OrdSub2 AS [OrdSub2],"
            sQry = sQry & "             T0.U_OrdName AS [OrdName],"
            sQry = sQry & "             T0.U_OrdAmt AS [Amount],"
            sQry = sQry & "             T0.U_Comment AS [Comment],"
            sQry = sQry & "             T0.DocEntry AS [BaseEntry],"
            sQry = sQry & "             T0.U_LineUID AS [BaseLine]"
            sQry = sQry & " FROM        [@PS_PP560L] AS T0"
            sQry = sQry & " WHERE       CONVERT(VARCHAR(20), T0.DocEntry) + '-' + CONVERT(VARCHAR(20), T0.U_LineUID) = '" & oForm01.DataSources.UserDataSources("BaseDL").VALUE & "'"
            
            Call oRecordSet01.DoQuery(sQry)
        
            oForm01.Items("AmtCls").Enabled = True
            oForm01.DataSources.UserDataSources("AmtCls").VALUE = oRecordSet01.Fields("AmtCls").VALUE '비용구분
            oForm01.Items("AmtCls").Enabled = False
            oForm01.DataSources.UserDataSources("OrdNum").VALUE = oRecordSet01.Fields("OrdNum").VALUE '작번
            oForm01.DataSources.UserDataSources("OrdSub1").VALUE = oRecordSet01.Fields("OrdSub1").VALUE '서브작번1
            oForm01.DataSources.UserDataSources("OrdSub2").VALUE = oRecordSet01.Fields("OrdSub2").VALUE '서브작번2
            oForm01.DataSources.UserDataSources("OrdName").VALUE = oRecordSet01.Fields("OrdName").VALUE '작명
            oForm01.DataSources.UserDataSources("Amount").VALUE = oRecordSet01.Fields("Amount").VALUE '목표금액
            oForm01.DataSources.UserDataSources("Comment").VALUE = oRecordSet01.Fields("Comment").VALUE '비고
            oForm01.DataSources.UserDataSources("BaseEntry").VALUE = oRecordSet01.Fields("BaseEntry").VALUE 'BaseEntry
            oForm01.DataSources.UserDataSources("BaseLine").VALUE = oRecordSet01.Fields("BaseLine").VALUE 'BaseLine
        
        Case "SOrdNum"

'            oForm01.Items("SCntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("SCntcCode").Specific.VALUE & "'") '성명

    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_PP564_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_PP564_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PS_PP564_FormItemEnabled()
On Error GoTo PS_PP564_FormItemEnabled_Error
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
PS_PP564_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_PP564_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
                Call PS_PP564_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_PP564_LoadCaption
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm01.Freeze(True)
                Call PS_PP564_Add_MatrixRow(oMat01.VisualRowCount)
                Call oForm01.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_PP564_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_PP564A.Clear

'                Call PS_PP564_LoadCaption
'                Call PS_PP564_FormItemEnabled
                '//Call PS_PP564_FormItemEnabled '//UDO방식
                '//Call PS_PP564_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_PP564_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_PP564B.RemoveRecord(oDS_PS_PP564B.Size - 1)
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
    
        If pval.ItemUID = "PS_PP564" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_PP564_HeaderSpaceLineDel() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                If PS_PP564_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                If PS_PP564_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_PP564_FormReset
                oForm01.Mode = fm_ADD_MODE
                
'                Call PS_PP564_LoadCaption
'                Call PS_PP564_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_PP564_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PS_PP564_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                If PS_PP564_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_PP564_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_PP564_LoadCaption
                Call PS_PP564_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
'            Call PS_PP564_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_PP564_LoadCaption
            Call PS_PP564_MTX01
            
'        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
'
'            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
'
'                Call PS_PP564_DeleteData
'                Call PS_PP564_FormReset
'                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
'
'                Call PS_PP564_LoadCaption
'                Call PS_PP564_MTX01
'
'            Else
'
'            End If
        
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            Call PS_PP564_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP564" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "BaseDL", "") '등록정보-문서번호
        
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum", "") '등록정보-작번
        
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
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PS_PP564_FlushToItemValue(pval.ItemUID)
        
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
'                    Call oDS_PS_PP564B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP564B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_PP564_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_PP564B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PS_PP564_FlushToItemValue(pval.ItemUID)
                
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
        Call PS_PP564_FormItemEnabled
        '//Call PS_PP564_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP564_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP564A", "U_CardCode,U_CardName")
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
'            If (PS_PP564_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
'            For i = 1 To oMat01.VisualRowCount
'                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
'            Next i
'            oMat01.FlushToDataSource
'            Call oDS_PS_PP564A.RemoveRecord(oDS_PS_PP564A.Size - 1)
'            oMat01.LoadFromDataSource
'            If oMat01.RowCount = 0 Then
'                Call PS_PP564_Add_MatrixRow(0)
'            Else
'                If Trim(oDS_PS_PP564A.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_PP564_Add_MatrixRow(oMat01.RowCount)
'                End If
'            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP564_CreateItems() As Boolean
On Error GoTo PS_PP564_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01     As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_PP564A = oForm01.DataSources.DBDataSources("@PS_PP564A")
    Set oDS_PS_PP564B = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '//////////등록정보//////////_S
    '문서번호
    Call oForm01.DataSources.UserDataSources.Add("BaseDL", dt_SHORT_TEXT, 20)
    Call oForm01.Items("BaseDL").Specific.DataBind.SetBound(True, "", "BaseDL")
    
    '비용구분
    Call oForm01.DataSources.UserDataSources.Add("AmtCls", dt_SHORT_TEXT, 10)
    Call oForm01.Items("AmtCls").Specific.DataBind.SetBound(True, "", "AmtCls")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdNum").Specific.DataBind.SetBound(True, "", "OrdNum")

    '서브작번1
    Call oForm01.DataSources.UserDataSources.Add("OrdSub1", dt_SHORT_TEXT, 2)
    Call oForm01.Items("OrdSub1").Specific.DataBind.SetBound(True, "", "OrdSub1")
    
    '서브작번2
    Call oForm01.DataSources.UserDataSources.Add("OrdSub2", dt_SHORT_TEXT, 3)
    Call oForm01.Items("OrdSub2").Specific.DataBind.SetBound(True, "", "OrdSub2")
    
    '작지명
    Call oForm01.DataSources.UserDataSources.Add("OrdName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OrdName").Specific.DataBind.SetBound(True, "", "OrdName")
    
    '목표금액
    Call oForm01.DataSources.UserDataSources.Add("Amount", dt_SUM)
    Call oForm01.Items("Amount").Specific.DataBind.SetBound(True, "", "Amount")
    
    '비고
    Call oForm01.DataSources.UserDataSources.Add("Comment", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Comment").Specific.DataBind.SetBound(True, "", "Comment")
    
    '기준문서번호
    Call oForm01.DataSources.UserDataSources.Add("BaseEntry", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BaseEntry").Specific.DataBind.SetBound(True, "", "BaseEntry")
    
    '기준문서라인번호
    Call oForm01.DataSources.UserDataSources.Add("BaseLine", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BaseLine").Specific.DataBind.SetBound(True, "", "BaseLine")
    '//////////등록정보//////////_E

    '//////////조회정보//////////_S
    '비용구분
    Call oForm01.DataSources.UserDataSources.Add("SAmtCls", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SAmtCls").Specific.DataBind.SetBound(True, "", "SAmtCls")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("SOrdNum", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SOrdNum").Specific.DataBind.SetBound(True, "", "SOrdNum")
    
    '서브작번1
    Call oForm01.DataSources.UserDataSources.Add("SOrdSub1", dt_SHORT_TEXT, 2)
    Call oForm01.Items("SOrdSub1").Specific.DataBind.SetBound(True, "", "SOrdSub1")
    
    '서브작번2
    Call oForm01.DataSources.UserDataSources.Add("SOrdSub2", dt_SHORT_TEXT, 3)
    Call oForm01.Items("SOrdSub2").Specific.DataBind.SetBound(True, "", "SOrdSub2")
    
    '작지명
    Call oForm01.DataSources.UserDataSources.Add("SOrdName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("SOrdName").Specific.DataBind.SetBound(True, "", "SOrdName")
    
    '비고
    Call oForm01.DataSources.UserDataSources.Add("SComment", dt_SHORT_TEXT, 100)
    Call oForm01.Items("SComment").Specific.DataBind.SetBound(True, "", "SComment")
    '//////////조회정보//////////_E
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_PP564_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP564_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_PP564_ComboBox_Setting()
On Error GoTo PS_PP564_ComboBox_Setting_Error

    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '//////////등록정보//////////
    '비용구분
    Call oForm01.Items("AmtCls").Specific.ValidValues.Add("%", "선택")
    sQry = "        SELECT      T0.U_Minor, "
    sQry = sQry & "             T0.U_CdName"
    sQry = sQry & " FROM        [@PS_SY001L] T0"
    sQry = sQry & " WHERE       T0.Code = 'P210'"
    sQry = sQry & "             AND T0.U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    T0.U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("AmtCls").Specific, sQry, "", False, False)
    Call oForm01.Items("AmtCls").Specific.Select(0, psk_Index)
    
    '//////////조회정보//////////
    '비용구분
    Call oForm01.Items("SAmtCls").Specific.ValidValues.Add("%", "전체")
    sQry = "        SELECT      T0.U_Minor, "
    sQry = sQry & "             T0.U_CdName"
    sQry = sQry & " FROM        [@PS_SY001L] T0"
    sQry = sQry & " WHERE       T0.Code = 'P210'"
    sQry = sQry & "             AND T0.U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    T0.U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SAmtCls").Specific, sQry, "", False, False)
    Call oForm01.Items("SAmtCls").Specific.Select(0, psk_Index)
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_PP564_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP564_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP564_CF_ChooseFromList()
On Error GoTo PS_PP564_CF_ChooseFromList_Error

    Exit Sub
PS_PP564_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP564_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP564_EnableMenus()
On Error GoTo PS_PP564_EnableMenus_Error
    
    Exit Sub
PS_PP564_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP564_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP564_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP564_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP564_FormItemEnabled
        '//Call PS_PP564_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP564_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP564_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP564_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP564_FormResize()
On Error GoTo PS_PP564_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_PP564_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP564_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_PP564_FormReset()
'******************************************************************************
'Function ID : PS_PP564_FormReset()
'해당모듈    : PS_PP564
'기능        : 화면 초기화
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP564_FormReset_Error

'    Dim sQry        As String
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '관리번호
'    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [Z_PS_PP564_01]"
'    Call RecordSet01.DoQuery(sQry)
'
'    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'        oForm01.DataSources.UserDataSources("DocEntry").VALUE = 1
'    Else
'        oForm01.DataSources.UserDataSources("DocEntry").VALUE = Trim(RecordSet01.Fields(0).VALUE) + 1
'    End If
    
'    Dim User_BPLId As String
'    User_BPLId = MDC_PS_Common.User_BPLId
    
    '//////////기준정보//////////
    oForm01.DataSources.UserDataSources("BaseDL").VALUE = "" '문서번호
    oForm01.DataSources.UserDataSources("AmtCls").VALUE = "%" '비용구분
    oForm01.DataSources.UserDataSources("OrdNum").VALUE = "" '작번
    oForm01.DataSources.UserDataSources("OrdSub1").VALUE = "" '서브작번1
    oForm01.DataSources.UserDataSources("OrdSub2").VALUE = "" '서브작번2
    oForm01.DataSources.UserDataSources("OrdName").VALUE = "" '작지명
    oForm01.DataSources.UserDataSources("Amount").VALUE = 0 '목표금액
    oForm01.DataSources.UserDataSources("Comment").VALUE = "" '비고
    oForm01.DataSources.UserDataSources("BaseEntry").VALUE = "" '기준문서번호
    oForm01.DataSources.UserDataSources("BaseLine").VALUE = "" '기준문서라인번호
    
    '//////////조회정보//////////
    oForm01.DataSources.UserDataSources("SAmtCls").VALUE = "%" '비용구분
    oForm01.DataSources.UserDataSources("SOrdNum").VALUE = "" '작번
    oForm01.DataSources.UserDataSources("SOrdSub1").VALUE = "" '서브작번1
    oForm01.DataSources.UserDataSources("SOrdSub2").VALUE = "" '서브작번2
    oForm01.DataSources.UserDataSources("SOrdName").VALUE = "" '작지명
    oForm01.DataSources.UserDataSources("SComment").VALUE = "" '비고(%)
    
    Call oForm01.Items("BaseDL").Click
    
'    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP564_FormReset_Error:
    Call oForm01.Freeze(False)
'    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP564_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_PP564_Print_Report01()

    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_PP564_Print_Report01_Error
    
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

    WinTitle = "[PS_PP564] 공용증"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PS_PP564_01.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PS_PP564_02.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PS_PP564_03.rpt"
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
    sQry = "EXEC [PS_PP564_90] '" & CLTCOD & "','" & DestNo1 & "','" & DestNo2 & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PS_PP564_Print_Report01_Error
    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_PP564_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_PP564_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub





