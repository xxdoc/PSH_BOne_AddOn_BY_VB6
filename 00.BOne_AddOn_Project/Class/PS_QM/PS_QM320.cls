VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_QM320.cls
'//  Module : 품질관리 -> 멀티게이지성적서관리 -> 11. MG 계량치 분석값 등록
'//  Desc :  MG 계량치 분석값 등록
'//  FormType : PS_QM320
'//  Create Date(Start) : 2017.09.14
'//  Create Date(End) : 2017.09.14
'//  Creator : Hwang youngsu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM320H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_QM320L As SAPbouiCOM.DBDataSource '등록라인

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
    Check As Boolean
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
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM320.srf")
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
    
    oFormUniqueID01 = "PS_QM320_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM320_CreateItems
    Call PS_QM320_ComboBox_Setting
    Call PS_QM320_CF_ChooseFromList
    Call PS_QM320_EnableMenus
    Call PS_QM320_SetDocument(oFromDocEntry01)
    Call PS_QM320_FormResize
    Call PS_QM320_FormReset '폼초기화 추가(2013.01.29 송명규)

'    Call PS_QM320_Add_MatrixRow(0, True)
    Call PS_QM320_LoadCaption
    Call PS_QM320_FormItemEnabled
    
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
        
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_QM320H]"
    Call RecordSet01.DoQuery(sQry)
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_QM320H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_QM320H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
        
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing

    '지급일자 설정
    oForm01.Items("DocDate").Specific.VALUE = Format(Date, "yyyyMMdd")
 '   Call oForm01.Items("MachCode").Click
    
    
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Load_Form Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM320_LoadCaption()
'******************************************************************************
'Function ID : PS_QM320_LoadCaption()
'해당모듈 : PS_QM320
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM320_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "추가"
        oForm01.Items("BtnDelete").Enabled = False
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
        oForm01.Items("BtnDelete").Enabled = True
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM320_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_QM320_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM320_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM320L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_QM320L.Offset = oRow
    oDS_PS_QM320L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM320_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_QM320_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_QM320_MTX01()
'******************************************************************************
'Function ID : PS_QM320_MTX01()
'해당모듈 : PS_QM320
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM320_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim SBPLID As String '사업장
    Dim SMachCod As String '장비코드
    Dim SPrvdDtFr As String '일자(시작)
    Dim SPrvdDtTo As String '일자(종료)
    
    SBPLID = IIf(Trim(oForm01.Items("SBPLId").Specific.VALUE) = "%", "", Trim(oForm01.Items("SBPLId").Specific.VALUE))
    SMachCod = Trim(oForm01.Items("SMachCod").Specific.VALUE)
    SPrvdDtFr = Trim(oForm01.Items("SPrvdDtFr").Specific.VALUE)
    SPrvdDtTo = Trim(oForm01.Items("SPrvdDtTo").Specific.VALUE)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "EXEC [PS_QM320_01] '" & SBPLID & "','" & SMachCod & "','" & SPrvdDtFr & "','" & SPrvdDtTo & "'"
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_QM320L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
        Call PS_QM320_LoadCaption
        
        GoTo PS_QM320_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_QM320L.Size Then
            oDS_PS_QM320L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_QM320L.Offset = i
        
        Call oDS_PS_QM320L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_QM320L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PS_QM320L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("BPLId").VALUE)) '사업장
        Call oDS_PS_QM320L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("MachCode").VALUE)) '장비코드
        Call oDS_PS_QM320L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("DocDate").VALUE), "yyyyMMdd")) '일자
        Call oDS_PS_QM320L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("LotId").VALUE)) '로트번호
        Call oDS_PS_QM320L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("HValue").VALUE)) '열처리 측정값
        Call oDS_PS_QM320L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("DValue").VALUE)) '주간측정값
        Call oDS_PS_QM320L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("NValue").VALUE)) '야간 측정값

        
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
PS_QM320_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_QM320_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_QM320_DeleteData()
'******************************************************************************
'Function ID : PS_QM320_DeleteData()
'해당모듈 : PS_QM320
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM320_DeleteData_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PS_QM320H] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_QM320_DeleteData_Error
        Else
            sQry = "DELETE FROM [@PS_QM320H] WHERE DocEntry = '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM320_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_QM320_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_QM320_UpdateData() As Boolean
'******************************************************************************
'Function ID : PS_QM320_UpdateData()
'해당모듈 : PS_QM320
'기능 : 기본정보를 수정
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM320_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim LotId As String '사업장
    Dim HValue As String '사번
    Dim DValue As String '성명
    Dim NValue As String '팀

    
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    LotId = Trim(oForm01.Items("LotId").Specific.VALUE)   ' 로트번호
    HValue = Trim(oForm01.Items("HValue").Specific.VALUE) '열처리 측정값
    DValue = Trim(oForm01.Items("DValue").Specific.VALUE) '주간 측정값
    NValue = Trim(oForm01.Items("NValue").Specific.VALUE) '야간 측정값

    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
        PS_QM320_UpdateData = False
        Exit Function
    End If
    
    sQry = "           UPDATE   [@PS_QM320H]"
    sQry = sQry & " SET         U_LotId = '" & LotId & "'," '로트번호
    sQry = sQry & "              U_HValue = '" & HValue & "'," '열처리 측정값
    sQry = sQry & "              U_DValue = '" & DValue & "',"  '주간측정값
    sQry = sQry & "              U_NValue = '" & NValue & "'"  '야간측정값

    sQry = sQry & " WHERE    DocEntry = '" & DocEntry & "'"
    
    Call RecordSet01.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PS_QM320_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM320_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_QM320_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_QM320_AddData() As Boolean
'******************************************************************************
'Function ID : PS_QM320_AddData()
'해당모듈 : PS_QM320
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM320_AddData_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim BPLID As String '사업장
    Dim DocDate As String '사번
    Dim MachCode As String '성명
    Dim LotId As String '팀
    Dim HValue As String '팀명
    Dim DValue As String '담당
    Dim NValue As String '담당명
    Dim ErrNum As String '담당명

'    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE) '일자
    MachCode = Trim(oForm01.Items("MachCode").Specific.VALUE) '장비코드
    LotId = Trim(oForm01.Items("LotId").Specific.VALUE) '로트번호
    HValue = Trim(oForm01.Items("HValue").Specific.VALUE) '열처리 측정값
    DValue = Trim(oForm01.Items("DValue").Specific.VALUE) '주간 측정값
    NValue = Trim(oForm01.Items("NValue").Specific.VALUE) '야간 측정값

        sQry = "select * from [@PS_QM320H] where u_DocDate = '" & DocDate & "' and " & "u_MachCode ='" & MachCode & "'"
        Call RecordSet01.DoQuery(sQry)
        
        If (RecordSet01.RecordCount >= 1) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_QM320_AddData_Error
        Else
    
    
    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_QM320H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "           INSERT INTO [@PS_QM320H]"
    sQry = sQry & " ("
    sQry = sQry & "     DocEntry,"
    sQry = sQry & "     DocNum,"
    
    sQry = sQry & "     U_BPLId," '사업장"
    sQry = sQry & "     U_DocDate," '일자"
    sQry = sQry & "     U_MachCode," '장비코드"
    sQry = sQry & "     U_LotId," '로트번호"
    sQry = sQry & "     U_HValue," '열처리측정값"
    sQry = sQry & "     U_DValue," '주간 측정값"
    sQry = sQry & "     U_NValue" '야간 측정값"
    
    sQry = sQry & " )"
    sQry = sQry & " VALUES"
    sQry = sQry & " ("
    sQry = sQry & DocEntry & ","
    sQry = sQry & DocEntry & ","
    
    sQry = sQry & "'" & BPLID & "'," '사업장"
    sQry = sQry & "'" & DocDate & "'," '일자"
    sQry = sQry & "'" & MachCode & "'," '장비코드"
    sQry = sQry & "'" & LotId & "'," '로트번호"
    sQry = sQry & "'" & HValue & "'," '열처리 측정값"
    sQry = sQry & "'" & DValue & "'," '주간 측정값"
    sQry = sQry & "'" & NValue & "'" '야간 측정값"
    
    sQry = sQry & ")"
    
    Call RecordSet02.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    End If
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_QM320_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM320_AddData_Error:
    PS_QM320_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
     If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("중복 입력입니다. 확인하세요.", "W")
    Else
        MDC_Com.MDC_GF_Message "PS_QM320_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
End Function
'/폼의 아이템 사용지정
Sub PS_QM320_FormItemEnabled()
On Error GoTo PS_QM320_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM320_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_QM320_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
                Call PS_QM320_FormReset
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_QM320_LoadCaption
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_QM320_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_QM320H.Clear

'                Call PS_QM320_LoadCaption
'                Call PS_QM320_FormItemEnabled
                '//Call PS_QM320_FormItemEnabled '//UDO방식
                '//Call PS_QM320_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_QM320_FormItemEnabled
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
    
        If pval.ItemUID = "PS_QM320" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then

                If PS_QM320_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_QM320_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_QM320_LoadCaption
                Call PS_QM320_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_QM320_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_QM320_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_QM320_LoadCaption
                Call PS_QM320_MTX01
            
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
            Call PS_QM320_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_QM320_LoadCaption
            Call PS_QM320_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_QM320_DeleteData
                Call PS_QM320_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_QM320_LoadCaption
                Call PS_QM320_MTX01
                
            Else
            
            End If
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM320" Then
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
                Call oDS_PS_QM320H.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
                Call oDS_PS_QM320H.setValue("U_BPLId", 0, oMat01.Columns("BPLId").Cells(pval.Row).Specific.VALUE) '사업장
                Call oDS_PS_QM320H.setValue("U_MachCode", 0, oMat01.Columns("MachCode").Cells(pval.Row).Specific.VALUE) '사번
                Call oDS_PS_QM320H.setValue("U_DocDate", 0, oMat01.Columns("DocDate").Cells(pval.Row).Specific.VALUE) '성명
                Call oDS_PS_QM320H.setValue("U_LotId", 0, oMat01.Columns("LotId").Cells(pval.Row).Specific.VALUE) '팀
                Call oDS_PS_QM320H.setValue("U_HValue", 0, oMat01.Columns("HValue").Cells(pval.Row).Specific.VALUE) '팀명
                Call oDS_PS_QM320H.setValue("U_DValue", 0, oMat01.Columns("DValue").Cells(pval.Row).Specific.VALUE) '팀
                Call oDS_PS_QM320H.setValue("U_NValue", 0, oMat01.Columns("NValue").Cells(pval.Row).Specific.VALUE) '팀명
                          
                
                oForm01.Mode = fm_UPDATE_MODE
                Call PS_QM320_LoadCaption
                
                Call oForm01.Freeze(False)
                
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        Call PS_QM320_DataCheck
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    Dim MachCode As String

    MachCode = Trim(oForm01.Items("MachCode").Specific.VALUE)
    If pval.BeforeAction = False Then
    
        If pval.ItemUID = "MachCode" Then
    
            If PS_QM320_DataCheck() = False Then
                BubbleEvent = False
                Exit Sub
            End If
                    
        End If
    
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
'                    Call oDS_PS_QM320L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM320L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_QM320_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_QM320L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
'
'                If pval.ItemUID = "CntcCode" Then
'
'                    oForm01.Items("CntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '성명
'                    oForm01.Items("TeamCd").Specific.VALUE = MDC_GetData.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '팀코드
'                    oForm01.Items("RspCd").Specific.VALUE = MDC_GetData.Get_ReData("U_RspCode", "Code", "[@PH_PY001A]", "'" & oForm01.Items("CntcCode").Specific.VALUE & "'") '담당코드
'
'                    oForm01.Items("TeamNm").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("TeamCd").Specific.VALUE & "'")  '팀명
'                    oForm01.Items("RspNm").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("RspCd").Specific.VALUE & "'")  '담당명
'
'                ElseIf pval.ItemUID = "PrtrCd" Then
'
'                    oForm01.Items("PrtrNm").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items("PrtrCd").Specific.VALUE & "'")  '보호구명
'                    oForm01.Items("PrtrSpec").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" & oForm01.Items("PrtrCd").Specific.VALUE & "'")  '규격
'
'                ElseIf pval.ItemUID = "SCntcCode" Then
'
'                    oForm01.Items("SCntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & oForm01.Items("SCntcCode").Specific.VALUE & "'") '성명
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
        Call PS_QM320_FormItemEnabled
        '//Call PS_QM320_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_QM320_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_QM320H", "U_CardCode,U_CardName")
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
'            If (PS_QM320_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_QM320H.RemoveRecord(oDS_PS_QM320H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_QM320_Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_QM320H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_QM320_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_QM320_CreateItems() As Boolean
On Error GoTo PS_QM320_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_QM320H = oForm01.DataSources.DBDataSources("@PS_QM320H")
    Set oDS_PS_QM320L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("SBPLId", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SBPLId").Specific.DataBind.SetBound(True, "", "SBPLId")
    '사업장_E

    '사번_S
    Call oForm01.DataSources.UserDataSources.Add("SMachCod", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SMachCod").Specific.DataBind.SetBound(True, "", "SMachCod")
    '사번_E
    
    
    '지급일자(시작)_S
    Call oForm01.DataSources.UserDataSources.Add("SPrvdDtFr", dt_DATE)
    Call oForm01.Items("SPrvdDtFr").Specific.DataBind.SetBound(True, "", "SPrvdDtFr")
    oForm01.Items("SPrvdDtFr").Specific.VALUE = Format(Date, "yyyyMM01")
    '지급일자(시작)_E
    
    '지급일자(종료)_S
    Call oForm01.DataSources.UserDataSources.Add("SPrvdDtTo", dt_DATE)
    Call oForm01.Items("SPrvdDtTo").Specific.DataBind.SetBound(True, "", "SPrvdDtTo")
    oForm01.Items("SPrvdDtTo").Specific.VALUE = Format(Date, "yyyyMMDD")
    '지급일자(종료)_E
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_QM320_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM320_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_QM320_ComboBox_Setting()
On Error GoTo PS_QM320_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '기본정보-사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("BPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    
    Call oForm01.Items("MachCode").Specific.ValidValues.Add("122201", "연속소둔로")
    Call oForm01.Items("MachCode").Specific.ValidValues.Add("122901", "S&D 1호기")
    Call oForm01.Items("MachCode").Specific.ValidValues.Add("122902", "S&D 2호기")
    Call oForm01.Items("MachCode").Specific.ValidValues.Add("122903", "S&D 3호기")
    Call oForm01.Items("MachCode").Specific.ValidValues.Add("122604", "DEGREASER 4호기")
    oForm01.Items("MachCode").Specific.Select "0", psk_Index
    
    
    Call oForm01.Items("SMachCod").Specific.ValidValues.Add("122201", "연속소둔로")
    Call oForm01.Items("SMachCod").Specific.ValidValues.Add("122901", "S&D 1호기")
    Call oForm01.Items("SMachCod").Specific.ValidValues.Add("122902", "S&D 2호기")
    Call oForm01.Items("SMachCod").Specific.ValidValues.Add("122903", "S&D 3호기")
    Call oForm01.Items("SMachCod").Specific.ValidValues.Add("122604", "DEGREASER 4호기")
    oForm01.Items("SMachCod").Specific.Select "0", psk_Index

    '조회조건-사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("SBPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)

    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_QM320_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM320_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM320_CF_ChooseFromList()
On Error GoTo PS_QM320_CF_ChooseFromList_Error

    Exit Sub
PS_QM320_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM320_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM320_EnableMenus()
On Error GoTo PS_QM320_EnableMenus_Error
    
    Exit Sub
PS_QM320_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM320_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM320_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM320_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM320_FormItemEnabled
        '//Call PS_QM320_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_QM320_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_QM320_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM320_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM320_FormResize()
On Error GoTo PS_QM320_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_QM320_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM320_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_QM320_FormReset()
'******************************************************************************
'Function ID : PS_QM320_FormReset()
'해당모듈 : PS_QM320
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM320_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_QM320H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_QM320H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_QM320H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    Call oDS_PS_QM320H.setValue("U_LotId", 0, "") '수량
    Call oDS_PS_QM320H.setValue("U_HValue", 0, "0") '반납일자
    Call oDS_PS_QM320H.setValue("U_DValue", 0, "0") '반납사유
    Call oDS_PS_QM320H.setValue("U_NValue", 0, "0") '비고
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM320_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_QM320_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function PS_QM320_DataCheck() As Boolean
'******************************************************************************
'Function ID : PS_QM320_DataCheck()
'해당모듈 : PS_QM320
'기능 : 중복 데이터 체크 및 S&D 공정일 경우 DValue, NValue/ 열처리 공정일 경우 LOTID, HValue 입력되도록 함.
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM320_DataCheck_Error

    PS_QM320_DataCheck = False

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim MachCode As String

    MachCode = Trim(oForm01.Items("MachCode").Specific.VALUE)
    
    If MachCode = "122201" Then

        oForm01.Items("DValue").Visible = False
        oForm01.Items("NValue").Visible = False
        oForm01.Items("LotId").Visible = True
        oForm01.Items("HValue").Visible = True
        
    Else
        oForm01.Items("DValue").Visible = True
        oForm01.Items("NValue").Visible = True
        oForm01.Items("LotId").Visible = False
        oForm01.Items("HValue").Visible = False

    End If

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_QM320_DataCheck_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "이미입력된 데이터입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_QM320_DataCheck_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Set oRecordSet01 = Nothing
    PS_QM320_DataCheck = False
End Function


