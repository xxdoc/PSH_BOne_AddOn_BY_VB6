VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PH_PY015.cls
'//  Module             : 인사 관리>근태관리>연차관리
'//  Desc               : 연차이월등록
'//  FormType           : PH_PY015
'//  Create Date(Start) : 2017.12.26
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01  As String
Public oForm            As SAPbouiCOM.Form
Public oMat01           As SAPbouiCOM.Matrix
'Private oDS_PH_PY015A As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PH_PY015B   As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01  As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01   As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01   As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml   As String
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY015.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY015_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01, "PH_PY015" '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    '//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm.Freeze True
    Call PH_PY015_CreateItems
    Call PH_PY015_ComboBox_Setting
    Call PH_PY015_CF_ChooseFromList
    Call PH_PY015_EnableMenus
    Call PH_PY015_SetDocument(oFromDocEntry01)
    Call PH_PY015_FormResize
'
'    Call PH_PY015_Add_MatrixRow(0, True)
    Call PH_PY015_LoadCaption
    
    oForm.EnableMenu ("1283"), False        '// 삭제
    oForm.EnableMenu ("1286"), False        '// 닫기
    oForm.EnableMenu ("1287"), False        '// 복제
    oForm.EnableMenu ("1285"), False        '// 복원
    oForm.EnableMenu ("1284"), False         '// 취소
    oForm.EnableMenu ("1293"), False         '// 행삭제
    oForm.EnableMenu ("1281"), False
    oForm.EnableMenu ("1282"), True

    Call PH_PY015_FormReset '폼초기화
    
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

Private Sub PH_PY015_LoadCaption()
'******************************************************************************
'Function ID : PH_PY015_LoadCaption()
'해당모듈    : PH_PY015
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY015_LoadCaption_Error

    Call oForm.Freeze(True)

    If oForm.Mode = fm_ADD_MODE Then
        oForm.Items("BtnAdd").Specific.Caption = "저장"
'        oForm.Items("BtnDelete").Enabled = False
'    ElseIf oForm.Mode = fm_OK_MODE Then
'        oForm.Items("BtnAdd").Specific.Caption = "확인"
    ElseIf oForm.Mode = fm_UPDATE_MODE Then
        oForm.Items("BtnAdd").Specific.Caption = "저장"
'        oForm.Items("BtnDelete").Enabled = True
    End If
    
    Call oForm.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY015_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PH_PY015_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PH_PY015_Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PH_PY015B.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PH_PY015B.Offset = oRow
    Call oDS_PH_PY015B.setValue("U_LineNum", oRow, oRow + 1)
    Call oDS_PH_PY015B.setValue("U_ColReg01", oRow, "Y")
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY015_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PH_PY015_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PH_PY015_MTX01()
'******************************************************************************
'Function ID : PH_PY015_MTX01()
'해당모듈    : PH_PY015
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY015_MTX01_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD    As String '사업장
    Dim TeamCode  As String '부서
    Dim RspCode   As String '담당
    Dim ClsCode   As String '반
    Dim StdYear   As String '기준년도
    Dim MSTCOD    As String '사번
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '부서
    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE) '담당
    ClsCode = Trim(oForm.Items("ClsCode").Specific.VALUE) '반
    StdYear = Trim(oForm.Items("StdYear").Specific.VALUE) '기준년도
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사번
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "      EXEC [PH_PY015_01] '"
    sQry = sQry & CLTCOD & "','" '사업장
    sQry = sQry & TeamCode & "','" '부서
    sQry = sQry & RspCode & "','" '담당
    sQry = sQry & ClsCode & "','" '반
    sQry = sQry & StdYear & "','" '기준년도
    sQry = sQry & MSTCOD & "'" '사번
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PH_PY015B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm.Mode = fm_ADD_MODE
        
        Call PH_PY015_Add_MatrixRow(0, True)
        Call PH_PY015_LoadCaption
        
        GoTo PH_PY015_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PH_PY015B.Size Then
            oDS_PH_PY015B.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PH_PY015B.Offset = i
        
        Call oDS_PH_PY015B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY015B.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PH_PY015B.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("TeamName").VALUE)) '부서
        Call oDS_PH_PY015B.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("RspName").VALUE)) '담당
        Call oDS_PH_PY015B.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("ClsName").VALUE)) '반
        Call oDS_PH_PY015B.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("FullName").VALUE)) '성명
        Call oDS_PH_PY015B.setValue("U_ColDt01", i, Trim(Format(oRecordSet01.Fields("ipsail").VALUE, "YYYYMMDD"))) '입사일자
        Call oDS_PH_PY015B.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("sojungil").VALUE)) '근로의무일
        Call oDS_PH_PY015B.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("bicnt").VALUE)) '결근일
        Call oDS_PH_PY015B.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("bRate").VALUE)) '출근율
        Call oDS_PH_PY015B.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("yycha").VALUE)) '연차발생
        Call oDS_PH_PY015B.setValue("U_ColQty02", i, Trim(oRecordSet01.Fields("gunsok").VALUE)) '근속가산
        Call oDS_PH_PY015B.setValue("U_ColQty03", i, Trim(oRecordSet01.Fields("iwol").VALUE)) '전년이월
        Call oDS_PH_PY015B.setValue("U_ColQty04", i, Trim(oRecordSet01.Fields("Tot").VALUE)) '계
        Call oDS_PH_PY015B.setValue("U_ColQty05", i, Trim(oRecordSet01.Fields("useyy").VALUE)) '사용일수
        Call oDS_PH_PY015B.setValue("U_ColQty06", i, Trim(oRecordSet01.Fields("jandd").VALUE)) '정산대상
        Call oDS_PH_PY015B.setValue("U_ColQty07", i, Trim(oRecordSet01.Fields("savedd").VALUE)) '적치
        Call oDS_PH_PY015B.setValue("U_ColQty08", i, Trim(oRecordSet01.Fields("paydd").VALUE)) '임금대치
        Call oDS_PH_PY015B.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("CLTCOD").VALUE)) '사업장
        Call oDS_PH_PY015B.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("StdYear").VALUE)) '기준년도
        Call oDS_PH_PY015B.setValue("U_ColReg20", i, Trim(oRecordSet01.Fields("MSTCOD").VALUE)) '사번
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
'    Call PH_PY015_Add_MatrixRow(oMat01.VisualRowCount)
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    
    Call oForm.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY015_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY015_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PH_PY015_DeleteData()
'******************************************************************************
'Function ID : PH_PY015_DeleteData()
'해당모듈    : PH_PY015
'기능        : 적치 수량 삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY015_DeleteData_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD  As String
    Dim StdYear As String
    Dim MSTCOD  As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("삭제 중...", 100, False)
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PH_PY015B.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            CLTCOD = Trim(oDS_PH_PY015B.GetValue("U_ColReg18", loopCount)) '사업장코드
            StdYear = Trim(oDS_PH_PY015B.GetValue("U_ColReg19", loopCount)) '기준년도
            MSTCOD = Trim(oDS_PH_PY015B.GetValue("U_ColReg20", loopCount)) '사번
                        
            sQry = "      EXEC [PH_PY015_04] '"
            sQry = sQry & CLTCOD & "','" '사업장코드
            sQry = sQry & StdYear & "','" '기준년도
            sQry = sQry & MSTCOD & "'" '사번
            
            Call oRecordSet01.DoQuery(sQry)
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY015_DeleteData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    ElseIf ErrNum = 2 Then
        Call MDC_Com.MDC_GF_Message(loopCount + 1 & "행은 실적이 등록되어 삭제가 불가능합니다. " & loopCount + 1 & "행 이후 삭제는 모두 취소됩니다.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY015_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PH_PY015_AddData() As Boolean
'******************************************************************************
'Function ID : PH_PY015_AddData()
'해당모듈    : PH_PY015
'기능        : 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY015_AddData_Error

    Dim i       As Integer
    Dim sQry    As String
    Dim ErrNum  As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD   As String '사업장코드
    Dim StdYear  As String '기준년도
    Dim MSTCOD   As String '사번
    Dim FullName As String '성명
    Dim SaveDCnt As Single '적치(이월) 수량
    Dim PayDCnt  As Single '임금대치 수량
    Dim UseDCnt  As Single '사용 수량
    Dim UserSign As String '등록자ID
    
    UserSign = oCompany.UserSignature
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)
    
    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PH_PY015B.GetValue("U_ColReg01", i)) = "Y" Then
                        
            CLTCOD = Trim(oDS_PH_PY015B.GetValue("U_ColReg18", i)) '사업장코드
            StdYear = Trim(oDS_PH_PY015B.GetValue("U_ColReg19", i)) '기준년도
            MSTCOD = Trim(oDS_PH_PY015B.GetValue("U_ColReg20", i)) '사번
            FullName = Trim(oDS_PH_PY015B.GetValue("U_ColReg05", i)) '성명
            SaveDCnt = Trim(oDS_PH_PY015B.GetValue("U_ColQty07", i)) '적치(이월) 수량
            PayDCnt = Trim(oDS_PH_PY015B.GetValue("U_ColQty08", i)) '임금대치 수량
            UseDCnt = Trim(oDS_PH_PY015B.GetValue("U_ColQty05", i)) '사용 수량
            
            sQry = "      EXEC [PH_PY015_02] '"
            sQry = sQry & CLTCOD & "','" '사업장코드
            sQry = sQry & StdYear & "','" '기준년도
            sQry = sQry & MSTCOD & "','" '사번
            sQry = sQry & FullName & "','" '성명
            sQry = sQry & SaveDCnt & "','" '적치(이월) 수량
            sQry = sQry & PayDCnt & "','" '임금대치 수량
            sQry = sQry & UseDCnt & "','" '사용수량
            sQry = sQry & UserSign & "'" '등록자ID
            
            If SaveDCnt = 0 And PayDCnt = 0 Then
                            
                If Sbo_Application.MessageBox(FullName & " 사원의 적치수량, 임금대치수량이 0입니다." & Chr(10) & "당해년도 입사자가 아니면 필수로 등록하여야합니다." & Chr(10) & "계속 등록하시겠습니까?", "1", "예", "아니오") = "1" Then
                
                    Call RecordSet01.DoQuery(sQry)
                    
                Else
                
                    ErrNum = "1"
                    GoTo PH_PY015_AddData_Error
                    
                    
                End If
                            
            Else
            
                Call RecordSet01.DoQuery(sQry)
            
            End If
            
            ProgBar01.VALUE = ProgBar01.VALUE + 1
            ProgBar01.Text = ProgBar01.VALUE & "/" & oMat01.VisualRowCount - 1 & "건 저장중..."
            
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    MDC_Com.MDC_GF_Message "저장 완료!", "S"
    
    Set RecordSet01 = Nothing
    
    PH_PY015_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY015_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PH_PY015_AddData = False
    Set RecordSet01 = Nothing
    
    If ErrNum = "1" Then
        Call MDC_Com.MDC_GF_Message("입력이 취소되었습니다.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY015_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Function

Private Function PH_PY015_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PH_PY015_HeaderSpaceLineDel()
'해당모듈    : PH_PY015
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY015_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm.Items("DestNo1").Specific.VALUE) = "" '출장번호1
            ErrNum = 1
            GoTo PH_PY015_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("DestNo2").Specific.VALUE) = "" '출장번호2
            ErrNum = 2
            GoTo PH_PY015_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("MSTCOD").Specific.VALUE) = "" '사원번호
            ErrNum = 3
            GoTo PH_PY015_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrDate").Specific.VALUE) = "" '시작일자
            ErrNum = 4
            GoTo PH_PY015_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("FrTime").Specific.VALUE) = "" '시작시각
            ErrNum = 5
            GoTo PH_PY015_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToDate").Specific.VALUE) = "" '종료일자
            ErrNum = 6
            GoTo PH_PY015_HeaderSpaceLineDel_Error
        Case Trim(oForm.Items("ToTime").Specific.VALUE) = "" '종료시각
            ErrNum = 7
            GoTo PH_PY015_HeaderSpaceLineDel_Error
    End Select

    PH_PY015_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY015_HeaderSpaceLineDel_Error:

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
    PH_PY015_HeaderSpaceLineDel = False
    
End Function

'/ 메트릭스 필수 사항 check
Private Function PH_PY015_MatrixSpaceLineDel() As Boolean
On Error GoTo PH_PY015_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    ErrNum = 0
    
    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PH_PY015B.GetValue("U_ColReg01", i)) = "Y" Then
            
            If Trim(oDS_PH_PY015B.GetValue("U_ColReg07", i)) = 0 And Trim(oDS_PH_PY015B.GetValue("U_ColReg08", i)) = 0 Then '적치, 임금대치 둘다 0인행 체크
                ErrNum = 1
                GoTo PH_PY015_MatrixSpaceLineDel_Error
            'ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg05", i)) = "" Then '성명
            '    ErrNum = 2
            '    GoTo PH_PY015_MatrixSpaceLineDel_Error
            'ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg03", i)) = "" Then '기준년도
            '    ErrNum = 3
            '    GoTo PH_PY015_MatrixSpaceLineDel_Error
            'ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg13", i)) = "" Then '구분
            '    ErrNum = 4
            '    GoTo PH_PY015_MatrixSpaceLineDel_Error
            'ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg15", i)) = "" Then '교육명
            '    ErrNum = 5
            '    GoTo PH_PY015_MatrixSpaceLineDel_Error
            'ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColDt01", i)) = "" Then '교육시작일
            '    ErrNum = 6
            '    GoTo PH_PY015_MatrixSpaceLineDel_Error
            'ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColDt02", i)) = "" Then '교육종료일
            '    ErrNum = 7
            '    GoTo PH_PY015_MatrixSpaceLineDel_Error
            'ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg06", i)) = "" Then '부서
            '    ErrNum = 8
            '    GoTo PH_PY015_MatrixSpaceLineDel_Error
            End If

        End If
    Next

    PH_PY015_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY015_MatrixSpaceLineDel_Error:

    If ErrNum = 1 Then
        'MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 사번이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 2 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 성명이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 기준년도가 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 구분이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 5 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 교육명이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 6 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 교육시작일이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 7 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 교육종료일이 없습니다. 확인하세요.", "E"
'    ElseIf ErrNum = 8 Then
'        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 부서코드가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PH_PY015_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
    PH_PY015_MatrixSpaceLineDel = False
    
End Function

Private Sub PH_PY015_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY015_FlushToItemValue_Error

    Dim loopCount  As Integer
    Dim ErrNum     As Integer
    Dim sQry       As String
    Dim ItemCode   As String
    
    Dim SaveDCnt   As Single '적치수량
    Dim JanDCnt    As Single '정산대상 수량
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case oUID

        Case "Mat01"
        
            Call oMat01.FlushToDataSource

            If oCol = "savedd" Then
            
                SaveDCnt = oMat01.Columns(oCol).Cells(oRow).Specific.VALUE
                JanDCnt = oMat01.Columns("jandd").Cells(oRow).Specific.VALUE
                
                If SaveDCnt > JanDCnt Then
                
                    ErrNum = 1
                    GoTo PH_PY015_FlushToItemValue_Error
                
                Else
                
                    Call oDS_PH_PY015B.setValue("U_ColQty08", oRow - 1, JanDCnt - SaveDCnt)
                
                End If
            
            End If

'            If oCol = "MSTCOD" Then
            
'                Call oDS_PH_PY015B.setValue("U_ColReg04", oRow - 1, oMat01.Columns(oCol).Cells(oRow).Specific.VALUE) '사번
'
'                Call oDS_PH_PY015B.setValue("U_ColReg03", oRow - 1, oForm.Items("StdYear").Specific.VALUE) '기준년도
'
'                '대상자의 인사마스터에서 소속 조회
'                sQry = "        SELECT  T0.U_TeamCode AS [TeamCode], " '부서코드
'                sQry = sQry & "         T1.U_CodeNm AS [TeamName], " '부서명
'                sQry = sQry & "         T0.U_RspCode AS [RspCode], " '담당코드
'                sQry = sQry & "         T2.U_CodeNm AS [RspName], " '담당명
'                sQry = sQry & "         T0.U_ClsCode AS [ClsCode], " '반코드
'                sQry = sQry & "         T3.U_CodeNm AS [ClsName], " '반명
'                sQry = sQry & "         T0.U_FullName AS [FullName], " '성명
'                sQry = sQry & "         T4.U_CodeNm AS [JIGNAM]," '직급
'                sQry = sQry & "         T0.U_CLTCOD AS [CLTCOD]" '소속사업장
'                sQry = sQry & " FROM    [@PH_PY001A] AS T0 "
'                sQry = sQry & "         LEFT JOIN"
'                sQry = sQry & "         [@PS_HR200L] AS T1"
'                sQry = sQry & "             ON T0.U_TeamCode = T1.U_Code"
'                sQry = sQry & "             AND T1.Code = '1'"
'                sQry = sQry & "         LEFT JOIN"
'                sQry = sQry & "         [@PS_HR200L] AS T2"
'                sQry = sQry & "             ON T0.U_RspCode = T2.U_Code"
'                sQry = sQry & "             AND T2.Code = '2'"
'                sQry = sQry & "         LEFT JOIN"
'                sQry = sQry & "         [@PS_HR200L] AS T3"
'                sQry = sQry & "             ON T0.U_ClsCode = T3.U_Code"
'                sQry = sQry & "             AND T3.Code = '9'"
'                sQry = sQry & "         LEFT JOIN"
'                sQry = sQry & "         [@PS_HR200L] AS T4"
'                sQry = sQry & "             ON T0.U_JIGCOD = T4.U_Code"
'                sQry = sQry & "             AND T4.Code = 'P129'"
'                sQry = sQry & " WHERE   T0.Code = '" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'"
'
'                Call oRecordSet01.DoQuery(sQry)
'
'                TeamCode = oRecordSet01.Fields("TeamCode").VALUE
'                TeamName = oRecordSet01.Fields("TeamName").VALUE
'                RspCode = oRecordSet01.Fields("RspCode").VALUE
'                RspName = oRecordSet01.Fields("RspName").VALUE
'                ClsCode = oRecordSet01.Fields("ClsCode").VALUE
'                ClsName = oRecordSet01.Fields("ClsName").VALUE
'                FullName = IIf(oMat01.Columns(oCol).Cells(oRow).Specific.VALUE = "9999999", "전사원", oRecordSet01.Fields("FullName").VALUE)
'                JIGNAM = oRecordSet01.Fields("JIGNAM").VALUE
'                CLTCOD = oRecordSet01.Fields("CLTCOD").VALUE
'
'                Call oDS_PH_PY015B.setValue("U_ColReg06", oRow - 1, TeamCode) '부서코드
'                Call oDS_PH_PY015B.setValue("U_ColReg07", oRow - 1, TeamName) '부서명
'                Call oDS_PH_PY015B.setValue("U_ColReg08", oRow - 1, RspCode) '담당코드
'                Call oDS_PH_PY015B.setValue("U_ColReg09", oRow - 1, RspName) '담당명
'                Call oDS_PH_PY015B.setValue("U_ColReg10", oRow - 1, ClsCode) '반코드
'                Call oDS_PH_PY015B.setValue("U_ColReg11", oRow - 1, ClsName) '반명
'                Call oDS_PH_PY015B.setValue("U_ColReg05", oRow - 1, FullName) '성명
'                Call oDS_PH_PY015B.setValue("U_ColReg12", oRow - 1, JIGNAM) '직급
'                Call oDS_PH_PY015B.setValue("U_ColReg26", oRow - 1, CLTCOD) '사업장
'
'                '행 추가
'                If oMat01.RowCount = oRow And Trim(oDS_PH_PY015B.GetValue("U_ColReg04", oRow - 1)) <> "" Then
'                    Call PH_PY015_Add_MatrixRow(oRow)
'                End If
'
'                Call oMat01.Columns("JIGNAM").Cells(oRow).CLICK(ct_Regular)
'
'            ElseIf oCol = "TeamCode" Then
'
'                Call oDS_PH_PY015B.setValue("U_ColReg07", oRow - 1, IIf(oMat01.Columns(oCol).Cells(oRow).Specific.VALUE = "9999", "전부서", MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'", " AND Code = '1'"))) '부서
'
'            ElseIf oCol = "RspCode" Then
'
'                Call oDS_PH_PY015B.setValue("U_ColReg09", oRow - 1, MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'", " AND Code = '2'")) '담당
'
'            ElseIf oCol = "ClsCode" Then
'
'                Call oDS_PH_PY015B.setValue("U_ColReg11", oRow - 1, MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'", " AND Code = '9'")) '반
'
'            ElseIf oCol = "Class" Then
'
'                Call oDS_PH_PY015B.setValue("U_ColReg14", oRow - 1, MDC_GetData.Get_ReData("name", "edType", "OHED", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'")) '교육구분
'
'            ElseIf oCol = "EduFrDt" Then '교육시작일
'
'                EduMonth = Mid(oMat01.Columns(oCol).Cells(oRow).Specific.VALUE, 5, 2)
'
'                If Val(EduMonth) >= 1 And Val(EduMonth) <= 3 Then
'
'                    EduQuarter = "1"
'                    EduHarf = "상"
'
'                ElseIf Val(EduMonth) >= 4 And Val(EduMonth) <= 6 Then
'
'                    EduQuarter = "2"
'                    EduHarf = "상"
'
'                ElseIf Val(EduMonth) >= 7 And Val(EduMonth) <= 9 Then
'
'                    EduQuarter = "3"
'                    EduHarf = "하"
'
'                ElseIf Val(EduMonth) >= 10 And Val(EduMonth) <= 12 Then
'
'                    EduQuarter = "4"
'                    EduHarf = "하"
'
'                End If
'
'                Call oDS_PH_PY015B.setValue("U_ColSum01", oRow - 1, Val(EduQuarter)) '분기
'                Call oDS_PH_PY015B.setValue("U_ColReg18", oRow - 1, EduHarf) '반기
'                Call oDS_PH_PY015B.setValue("U_ColSum02", oRow - 1, Val(EduMonth)) '월
'
'                'Call oMat01.Columns("EduToDt").Cells(oRow).CLICK(ct_Regular)
'
'            End If
'
            Call oMat01.LoadFromDataSource
            '강제 포커스 이동_S
            Call oMat01.Columns(oCol).Cells(oRow).CLICK(ct_Regular)
            '강제 포커스 이동_E
            Call oMat01.AutoResizeColumns
        
        Case "MSTCOD"

            oForm.Items("FullName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.VALUE & "'") '성명
            
        Case "TeamCode"
        
            oForm.Items("TeamName").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm.Items("TeamCode").Specific.VALUE & "'", "AND Code = '1'") '부서
            
        Case "RspCode"
            
            oForm.Items("RspName").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm.Items("RspCode").Specific.VALUE & "'", "AND Code = '2'") '담당
        
        Case "ClsCode"
        
            oForm.Items("ClsName").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm.Items("ClsCode").Specific.VALUE & "'", "AND Code = '9'") '반
                  
                  
    End Select

    Set oRecordSet01 = Nothing

    Exit Sub
    
PH_PY015_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    If ErrNum = 1 Then
    
        Call MDC_Com.MDC_GF_Message("적치(이월) 수량이 정산대상보다 많습니다. 확인하십시오.", "E")
    
    Else

        Call MDC_Com.MDC_GF_Message("PH_PY015_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    
    End If

End Sub

'/폼의 아이템 사용지정
Sub PH_PY015_FormItemEnabled()
On Error GoTo PH_PY015_FormItemEnabled_Error
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
        
    '매트릭스 컬럼(사업장코드, 기준년도, 사번) 숨김
    oMat01.Columns("CLTCOD").Visible = False
    oMat01.Columns("StdYear").Visible = False
    oMat01.Columns("MSTCOD").Visible = False
    
    Exit Sub
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY015_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY015_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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

                Call PH_PY015_FormReset
                Call PH_PY015_LoadCaption
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm.Freeze(True)
                Call PH_PY015_Add_MatrixRow(oMat01.VisualRowCount)
                Call oForm.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PH_PY015_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PH_PY015A.Clear

'                Call PH_PY015_LoadCaption
'                Call PH_PY015_FormItemEnabled
                '//Call PH_PY015_FormItemEnabled '//UDO방식
                '//Call PH_PY015_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY015_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm.Freeze(True)
                Call oDS_PH_PY015B.RemoveRecord(oDS_PH_PY015B.Size - 1)
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

    Dim ErrorRecordMstNam As String '적치, 임금대치 둘다 0인 행의 사원 이름 저장

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "PH_PY015" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm.Mode = fm_ADD_MODE Then
            
'                If PH_PY015_HeaderSpaceLineDel() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                '매트릭스 필수자료 체크
'                If PH_PY015_MatrixSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If

                If PH_PY015_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

'                Call PH_PY015_FormReset
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY015_LoadCaption
                Call PH_PY015_MTX01
                
                oLast_Mode = oForm.Mode
                
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            
                If PH_PY015_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
'                If PH_PY015_DataCheck() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
                 
                Call PH_PY015_FormReset
                oForm.Mode = fm_ADD_MODE
                
                Call PH_PY015_LoadCaption
                Call PH_PY015_MTX01
                
'                oForm.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
'            Call PH_PY015_FormReset
            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PH_PY015_LoadCaption
            Call PH_PY015_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PH_PY015_DeleteData
'                Call PH_PY015_FormReset
                oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PH_PY015_LoadCaption
                Call PH_PY015_MTX01
                
            Else
            
            End If
        
        ElseIf pval.ItemUID = "BtnSel" Then '전체선택
        
            Call PH_PY015_CheckAll
        
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            Call PH_PY015_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PH_PY015" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "TeamCode", "") 'Header-부서
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "RspCode", "") 'Header-담당
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "ClsCode", "") 'Header-반
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD", "") 'Header-사번
        
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
        
            Call PH_PY015_FlushToItemValue(pval.ItemUID)
        
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
            
                If pval.ColUID = "savedd" Then
            
                    Call PH_PY015_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                
                    'oMat01.AutoResizeColumns
                
                End If
            
'                oMat01.FlushToDataSource
'
'                If (pval.ColUID = "MSTCOD") Then
'                    Call oDS_PH_PY015B.setValue("U_ColReg02", pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    Call oDS_PH_PY015B.setValue("U_ColReg03", pval.Row - 1, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"))
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PH_PY015B.GetValue("U_ColReg02", pval.Row - 1)) <> "" Then
'                        Call PH_PY015_Add_MatrixRow(pval.Row)
'                    End If
'                Else
''                    Call oDS_PH_PY015B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
'
'                oMat01.LoadFromDataSource
                oMat01.AutoResizeColumns
                
            Else
                
                Call PH_PY015_FlushToItemValue(pval.ItemUID)
                
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
        Call PH_PY015_FormItemEnabled
        '//Call PH_PY015_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY015_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY015A", "U_CardCode,U_CardName")
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
'            If (PH_PY015_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PH_PY015B.RemoveRecord(oDS_PH_PY015B.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PH_PY015_Add_MatrixRow(0)
            Else
                If Trim(oDS_PH_PY015B.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PH_PY015_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY015_CreateItems() As Boolean
On Error GoTo PH_PY015_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    Set oDS_PH_PY015B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    
    '부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("TeamCode").Specific.DataBind.SetBound(True, "", "TeamCode")
    
    '부서명
    Call oForm.DataSources.UserDataSources.Add("TeamName", dt_SHORT_TEXT, 200)
    Call oForm.Items("TeamName").Specific.DataBind.SetBound(True, "", "TeamName")
    
    '담당
    Call oForm.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("RspCode").Specific.DataBind.SetBound(True, "", "RspCode")
    
    '담당명
    Call oForm.DataSources.UserDataSources.Add("RspName", dt_SHORT_TEXT, 200)
    Call oForm.Items("RspName").Specific.DataBind.SetBound(True, "", "RspName")

    '반
    Call oForm.DataSources.UserDataSources.Add("ClsCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("ClsCode").Specific.DataBind.SetBound(True, "", "ClsCode")
    
    '반명
    Call oForm.DataSources.UserDataSources.Add("ClsName", dt_SHORT_TEXT, 200)
    Call oForm.Items("ClsName").Specific.DataBind.SetBound(True, "", "ClsName")

    '기준년도
    Call oForm.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 4)
    Call oForm.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")

    '사번
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 20)
    Call oForm.Items("MSTCOD").Specific.DataBind.SetBound(True, "", "MSTCOD")
    
    '성명
    Call oForm.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 50)
    Call oForm.Items("FullName").Specific.DataBind.SetBound(True, "", "FullName")

    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY015_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY015_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PH_PY015_ComboBox_Setting()
On Error GoTo PH_PY015_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
'    '재직여부
'    Call oForm.Items("Status").Specific.ValidValues.Add("1", "재직자")
'    Call oForm.Items("Status").Specific.ValidValues.Add("2", "퇴직자포함")
'    Call oForm.Items("Status").Specific.Select(0, psk_Index)
'
'    '기준월
'    sQry = "        SELECT      U_Code AS [Code],"
'    sQry = sQry & "             U_CodeNm AS [Name]"
'    sQry = sQry & " FROM        [@PS_HR200L]"
'    sQry = sQry & " WHERE       Code = '4'"
'    sQry = sQry & "             AND U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY    U_Seq"
'
'    Call oForm.Items("StdMt").Specific.ValidValues.Add("%", "전체")
'    Call MDC_SetMod.Set_ComboList(oForm.Items("StdMt").Specific, sQry, "", False, False)
'    Call oForm.Items("StdMt").Specific.Select(0, psk_Index)
'
'    '계획구분
'    sQry = "        SELECT      U_Code AS [Code],"
'    sQry = sQry & "             U_CodeNm AS [Name]"
'    sQry = sQry & " FROM        [@PS_HR200L]"
'    sQry = sQry & " WHERE       Code = 'P239'"
'    sQry = sQry & "             AND U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY    U_Seq"
'
'    Call oForm.Items("PlnCls").Specific.ValidValues.Add("%", "전체")
'    Call MDC_SetMod.Set_ComboList(oForm.Items("PlnCls").Specific, sQry, "", False, False)
'    Call oForm.Items("PlnCls").Specific.Select(0, psk_Index)
'
'    '상태
'    Call oForm.Items("DocStatus").Specific.ValidValues.Add("%", "전체")
'    Call oForm.Items("DocStatus").Specific.ValidValues.Add("O", "계획")
'    Call oForm.Items("DocStatus").Specific.ValidValues.Add("P", "완료")
'    Call oForm.Items("DocStatus").Specific.ValidValues.Add("C", "취소")
'    Call oForm.Items("DocStatus").Specific.Select(0, psk_Index)
'
'    '//////////매트릭스//////////_S
'    '상태
'    Call oMat01.Columns("DocStatus").ValidValues.Add("O", "계획")
'    Call oMat01.Columns("DocStatus").ValidValues.Add("P", "완료")
'    Call oMat01.Columns("DocStatus").ValidValues.Add("C", "취소")
'
'    '계획구분
'    sQry = "        SELECT      U_Code AS [Code],"
'    sQry = sQry & "             U_CodeNm AS [Name]"
'    sQry = sQry & " FROM        [@PS_HR200L]"
'    sQry = sQry & " WHERE       Code = 'P239'"
'    sQry = sQry & "             AND U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY    U_Seq"
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("PlnCls"), sQry)
'
'    '이수(수료)증
'    Call oMat01.Columns("Certi").ValidValues.Add("O", "미제출")
'    Call oMat01.Columns("Certi").ValidValues.Add("C", "제출")
'
'    '보고서
'    Call oMat01.Columns("Report").ValidValues.Add("O", "미제출")
'    Call oMat01.Columns("Report").ValidValues.Add("C", "제출")
'
'    '사업장
'    sQry = "        SELECT      BPLId AS [BPLId],"
'    sQry = sQry & "             BPLName AS [BPLName]"
'    sQry = sQry & " FROM        OBPL"
'
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CLTCOD"), sQry)
'    '//////////매트릭스//////////_E
    
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY015_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY015_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY015_CF_ChooseFromList()
On Error GoTo PH_PY015_CF_ChooseFromList_Error

    Exit Sub
PH_PY015_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY015_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY015_EnableMenus()
On Error GoTo PH_PY015_EnableMenus_Error
    
    Exit Sub
PH_PY015_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY015_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY015_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY015_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PH_PY015_FormItemEnabled
        '//Call PH_PY015_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY015_FormItemEnabled
'        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm.Items("1").Click ct_Regular
    End If
    Exit Sub
PH_PY015_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY015_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY015_FormResize()
On Error GoTo PH_PY015_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PH_PY015_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY015_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PH_PY015_FormReset()
'******************************************************************************
'Function ID : PH_PY015_FormReset()
'해당모듈    : PH_PY015
'기능        : 화면 초기화
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY015_FormReset_Error

    Dim sQry        As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
    Dim User_BPLID As String
    User_BPLID = MDC_PS_Common.User_BPLID

    '헤더 초기화
    oForm.DataSources.UserDataSources("CLTCOD").VALUE = User_BPLID '사업장
    oForm.DataSources.UserDataSources("TeamCode").VALUE = "" '부서
    oForm.DataSources.UserDataSources("TeamName").VALUE = "" '부서명
    oForm.DataSources.UserDataSources("RspCode").VALUE = "" '담당
    oForm.DataSources.UserDataSources("RspName").VALUE = "" '담당명
    oForm.DataSources.UserDataSources("ClsCode").VALUE = "" '반
    oForm.DataSources.UserDataSources("ClsName").VALUE = "" '반명
    oForm.DataSources.UserDataSources("StdYear").VALUE = Format(Now, "YYYY") '기준년도
    oForm.DataSources.UserDataSources("MSTCOD").VALUE = "" '사번
    oForm.DataSources.UserDataSources("FullName").VALUE = "" '성명
    
    '라인 초기화
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
'    Call PH_PY015_Add_MatrixRow(0, True)
    
    Set RecordSet01 = Nothing
    Call oForm.Freeze(False)

Exit Sub

PH_PY015_FormReset_Error:
    Call oForm.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PH_PY015_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PH_PY015_CheckAll()
On Error GoTo PH_PY015_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PH_PY015B.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PH_PY015B.Offset = loopCount
        If CheckType = "N" Then
            Call oDS_PH_PY015B.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PH_PY015B.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm.Freeze(False)

    Exit Sub
    
PH_PY015_CheckAll_Error:
    Call oForm.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PH_PY015_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Sub PH_PY015_Print_Report01()
'******************************************************************************
'Function ID : PH_PY015_Print_Report01()
'해당모듈    : PH_PY015
'기능        : 리포트 출력
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY015_Print_Report01_Error

    Dim DocNum      As String
    Dim ErrNum      As Integer
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim CLTCOD      As String
    Dim DestNo1     As String
    Dim DestNo2     As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY015_Print_Report01_Error
    End If
    
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DestNo1 = Trim(oForm.Items("DestNo1").Specific.VALUE)
    DestNo2 = Trim(oForm.Items("DestNo2").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY015]"
    
    If CLTCOD = "1" Then '창원
        ReportName = "PH_PY015_01.rpt"
    ElseIf CLTCOD = "2" Then '동래
        ReportName = "PH_PY015_02.rpt"
    ElseIf CLTCOD = "3" Then '사상
        ReportName = "PH_PY015_03.rpt"
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
    sQry = "EXEC [PH_PY015_90] '" & CLTCOD & "','" & DestNo1 & "','" & DestNo2 & "'"
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY015_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY015_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY015_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub



