VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_GA101.cls
'//  Module : 총무관리>안전환경관리>무재해목표관리>무재해목표조회
'//  Desc : 무재해목표조회
'//  FormType : PS_GA101
'//  Create Date(Start) : 2012.09.17
'//  Create Date(End) : 2012.09.17
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_GA101H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_GA101L As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA101.srf")
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
    
    oFormUniqueID01 = "PS_GA101_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_GA101_CreateItems
    Call PS_GA101_ComboBox_Setting
    Call PS_GA101_CF_ChooseFromList
    Call PS_GA101_EnableMenus
    Call PS_GA101_SetDocument(oFromDocEntry01)
    Call PS_GA101_FormResize

'    Call PS_GA101_Add_MatrixRow(0, True)
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

        
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing

'    '발생일자 설정
'    Call oDS_PS_GA101H.setValue("U_BOSDate", 0, Format(Date, "YYYYMMDD"))
'    '종료일자 설정
'    Call oDS_PS_GA101H.setValue("U_BOEDate", 0, Format(Date, "YYYYMMDD"))

     'oForm01.Items("GCode").Click ct_Regular
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

'/메트릭스 Row추가
Sub PS_GA101_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA101_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_GA101L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_GA101L.Offset = oRow
    oDS_PS_GA101L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA101_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_GA101_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_GA101_MTX01()
'******************************************************************************
'Function ID : PS_GA101_MTX01()
'해당모듈 : PS_GA101
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA101_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim TrgtCls As String '목표구분
    Dim BPLId As String '사업장
    Dim CmYN As String '달성여부
    Dim TeamCd1 As String '팀1
    Dim RspCd1 As String '담당1
    Dim TeamCd2 As String '팀2
    Dim RspCd2 As String '담당2
    
    TrgtCls = IIf(Trim(oForm01.Items("TrgtCls").Specific.VALUE) = "%", "", Trim(oForm01.Items("TrgtCls").Specific.VALUE)) '목표구분
    BPLId = IIf(Trim(oForm01.Items("BPLId").Specific.VALUE) = "%", "", Trim(oForm01.Items("BPLId").Specific.VALUE)) '사업장
    CmYN = IIf(Trim(oForm01.Items("CmYN").Specific.VALUE) = "%", "", Trim(oForm01.Items("CmYN").Specific.VALUE)) '달성여부
    TeamCd1 = Trim(oForm01.Items("TeamCd1").Specific.VALUE) '팀1
    RspCd1 = Trim(oForm01.Items("RspCd1").Specific.VALUE) '담당1
    TeamCd2 = Trim(oForm01.Items("TeamCd2").Specific.VALUE) '팀2
    RspCd2 = Trim(oForm01.Items("RspCd2").Specific.VALUE) '담당2
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "EXEC [PS_GA101_01] '" & TrgtCls & "','" & BPLId & "','" & CmYN & "','" & TeamCd1 & "','" & RspCd1 & "','" & TeamCd2 & "','" & RspCd2 & "'"
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_GA101L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_GA101_Add_MatrixRow(0, True)
        GoTo PS_GA101_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_GA101L.Size Then
            oDS_PS_GA101L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_GA101L.Offset = i
        
        Call oDS_PS_GA101L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_GA101L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) '관리번호
        Call oDS_PS_GA101L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("TrgtCls").VALUE)) '목표구분
        Call oDS_PS_GA101L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("BPLId").VALUE)) '사업장
        Call oDS_PS_GA101L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("TeamCd1").VALUE)) '팀1
        Call oDS_PS_GA101L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("RspCd1").VALUE)) '담당1
        Call oDS_PS_GA101L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("TeamCd2").VALUE)) '팀2
        Call oDS_PS_GA101L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("RspCd2").VALUE)) '담당2
        Call oDS_PS_GA101L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("TrgtNod").VALUE)) '목표일수
        Call oDS_PS_GA101L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("Grd").VALUE)) '설정등급
        Call oDS_PS_GA101L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("TrgtMtp").VALUE)) '목표배수
        Call oDS_PS_GA101L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("StrDt").VALUE), "yyyyMMdd")) '재개시일자
        Call oDS_PS_GA101L.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("StrPrns").VALUE)) '재개시인원
        Call oDS_PS_GA101L.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("CmDt").VALUE), "yyyyMMdd")) '달성예정일
        Call oDS_PS_GA101L.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("CmYN").VALUE)) '달성여부
        
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
PS_GA101_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA101_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'Private Function PS_GA101_HeaderSpaceLineDel() As Boolean
''******************************************************************************
''Function ID : PS_GA101_HeaderSpaceLineDel()
''해당모듈 : PS_GA101
''기능 : 필수입력사항 체크
''인수 : 없음
''반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
''특이사항 : 없음
''******************************************************************************
'On Error GoTo PS_GA101_HeaderSpaceLineDel_Error
'
'    Dim ErrNum As Integer
'    ErrNum = 0
'
'    Select Case True
'        Case Trim(oForm01.Items("Loc").Specific.VALUE) = "" '발생위치
'            ErrNum = 1
'            GoTo PS_GA101_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BORsn").Specific.VALUE) = "" '정전사유
'            ErrNum = 2
'            GoTo PS_GA101_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOCls").Specific.VALUE) = "%" '정전구분
'            ErrNum = 3
'            GoTo PS_GA101_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOSDate").Specific.VALUE) = "" '발생일자
'            ErrNum = 4
'            GoTo PS_GA101_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOSHour").Specific.VALUE) = "" '발생시각(시)
'            ErrNum = 5
'            GoTo PS_GA101_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOSMin").Specific.VALUE) = "" '발생시각(분)
'            ErrNum = 6
'            GoTo PS_GA101_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOEDate").Specific.VALUE) = "" '종료일자
'            ErrNum = 7
'            GoTo PS_GA101_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOEHour").Specific.VALUE) = "" '종료시각(시)
'            ErrNum = 8
'            GoTo PS_GA101_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("BOEMin").Specific.VALUE) = "" '종료시각(분)
'            ErrNum = 9
'            GoTo PS_GA101_HeaderSpaceLineDel_Error
'    End Select
'
'    PS_GA101_HeaderSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_GA101_HeaderSpaceLineDel_Error:
'    If ErrNum = 1 Then
'        MDC_Com.MDC_GF_Message "발생위치는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("Loc").Click(ct_Regular)
'    ElseIf ErrNum = 2 Then
'        MDC_Com.MDC_GF_Message "정전사유는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BORsn").Click(ct_Regular)
'    ElseIf ErrNum = 3 Then
'        MDC_Com.MDC_GF_Message "정전구분은 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOCls").Click(ct_Regular)
'    ElseIf ErrNum = 4 Then
'        MDC_Com.MDC_GF_Message "발생일자는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOSDate").Click(ct_Regular)
'    ElseIf ErrNum = 5 Then
'        MDC_Com.MDC_GF_Message "발생시각(시)는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOSHour").Click(ct_Regular)
'    ElseIf ErrNum = 6 Then
'        MDC_Com.MDC_GF_Message "발생시각(분)은 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOSMin").Click(ct_Regular)
'    ElseIf ErrNum = 7 Then
'        MDC_Com.MDC_GF_Message "종료일자는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOEDate").Click(ct_Regular)
'    ElseIf ErrNum = 8 Then
'        MDC_Com.MDC_GF_Message "종료일자(시)는 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOEHour").Click(ct_Regular)
'    ElseIf ErrNum = 9 Then
'        MDC_Com.MDC_GF_Message "종료일자(분)은 필수사항입니다. 확인하세요.", "E"
'        Call oForm01.Items("BOEMin").Click(ct_Regular)
'    End If
'    PS_GA101_HeaderSpaceLineDel = False
'End Function
'
''/ 메트릭스 필수 사항 check
'Private Function PS_GA101_MatrixSpaceLineDel() As Boolean
'On Error GoTo PS_GA101_MatrixSpaceLineDel_Error
'
'    Dim i As Long
'    Dim ErrNum As Integer
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Dim sQry As String
'
'    Set oRecordSet01 = Nothing
'    PS_GA101_MatrixSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_GA101_MatrixSpaceLineDel_Error:
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
'        MDC_Com.MDC_GF_Message "PS_GA101_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'    PS_GA101_MatrixSpaceLineDel = False
'End Function

Private Sub PS_GA101_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_GA101_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    Dim Qty As Integer
    Dim Calculate_Weight As Double
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Select Case oUID
'
'        Case "BOSHour"
'
'    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_GA101_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("시각은 숫자만 입력이 가능합니다.", "E")
    ElseIf ErrNum = 2 Then
        Call MDC_Com.MDC_GF_Message("시각(시)는 24미만의 값만 입력이 가능합니다.", "E")
    ElseIf ErrNum = 3 Then
        Call MDC_Com.MDC_GF_Message("시각(분)은 60미만의 값만 입력이 가능합니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA101_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_GA101_FormItemEnabled()
On Error GoTo PS_GA101_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA101_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_GA101_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
'                oForm01.Mode = fm_ADD_MODE
'                BubbleEvent = False
                
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
                '//Call PS_GA101_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_GA101H.Clear

'                Call PS_GA101_LoadCaption
'                Call PS_GA101_FormItemEnabled
                '//Call PS_GA101_FormItemEnabled '//UDO방식
                '//Call PS_GA101_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_GA101_FormItemEnabled
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
    
        If pval.ItemUID = "PS_GA101" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnSearch" Then '/조회
                
            Call PS_GA101_MTX01
            
        ElseIf pval.ItemUID = "BtnPrint" Then '출력
            
            Call PS_GA101_Print_Report01
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA101" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "TeamCd1", "") '기본정보-팀1 포맷서치
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RspCd1", "") '기본정보-담당1 포맷서치
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "TeamCd2", "") '기본정보-팀2 포맷서치
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RspCd2", "") '기본정보-담당2 포맷서치
        
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
'                    Call oDS_PS_GA101L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_GA101L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_GA101_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_GA101L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PS_GA101_FlushToItemValue(pval.ItemUID)
                
                If pval.ItemUID = "TeamCd1" Then
                
                    oForm01.Items("TeamNm1").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("TeamCd1").Specific.VALUE & "'") '팀1
                    
                ElseIf pval.ItemUID = "RspCd1" Then
                
                    oForm01.Items("RspNm1").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("RspCd1").Specific.VALUE & "'")  '담당1
                    
                ElseIf pval.ItemUID = "TeamCd2" Then
                
                    oForm01.Items("TeamNm2").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("TeamCd2").Specific.VALUE & "'") '팀2
                    
                ElseIf pval.ItemUID = "RspCd2" Then
                
                    oForm01.Items("RspNm2").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("RspCd2").Specific.VALUE & "'")  '담당2
                
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
        Call PS_GA101_FormItemEnabled
        '//Call PS_GA101_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_GA101_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_GA101H", "U_CardCode,U_CardName")
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
'            If (PS_GA101_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_GA101H.RemoveRecord(oDS_PS_GA101H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_GA101_Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_GA101H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_GA101_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_GA101_CreateItems() As Boolean
On Error GoTo PS_GA101_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_GA101H = oForm01.DataSources.DBDataSources("@PS_GA101H")
    Set oDS_PS_GA101L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '목표구분_S
    Call oForm01.DataSources.UserDataSources.Add("TrgtCls", dt_SHORT_TEXT, 10)
    Call oForm01.Items("TrgtCls").Specific.DataBind.SetBound(True, "", "TrgtCls")
    '목표구분_E
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    '사업장_E
    
    '달성여부_S
    Call oForm01.DataSources.UserDataSources.Add("CmYN", dt_SHORT_TEXT, 10)
    Call oForm01.Items("CmYN").Specific.DataBind.SetBound(True, "", "CmYN")
    '달성여부_E
    
    '팀코드1_S
    Call oForm01.DataSources.UserDataSources.Add("TeamCd1", dt_SHORT_TEXT, 20)
    Call oForm01.Items("TeamCd1").Specific.DataBind.SetBound(True, "", "TeamCd1")
    '팀코드1_E
    
    '팀명1_S
    Call oForm01.DataSources.UserDataSources.Add("TeamNm1", dt_SHORT_TEXT, 100)
    Call oForm01.Items("TeamNm1").Specific.DataBind.SetBound(True, "", "TeamNm1")
    '팀명1_E
    
    '담당코드1_S
    Call oForm01.DataSources.UserDataSources.Add("RspCd1", dt_SHORT_TEXT, 20)
    Call oForm01.Items("RspCd1").Specific.DataBind.SetBound(True, "", "RspCd1")
    '담당코드1_E
    
    '담당명1_S
    Call oForm01.DataSources.UserDataSources.Add("RspNm1", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RspNm1").Specific.DataBind.SetBound(True, "", "RspNm1")
    '담당명1_E
    
    '팀코드2_S
    Call oForm01.DataSources.UserDataSources.Add("TeamCd2", dt_SHORT_TEXT, 20)
    Call oForm01.Items("TeamCd2").Specific.DataBind.SetBound(True, "", "TeamCd2")
    '팀코드2_E
    
    '팀명2_S
    Call oForm01.DataSources.UserDataSources.Add("TeamNm2", dt_SHORT_TEXT, 100)
    Call oForm01.Items("TeamNm2").Specific.DataBind.SetBound(True, "", "TeamNm2")
    '팀명2_E
    
    '담당코드2_S
    Call oForm01.DataSources.UserDataSources.Add("RspCd2", dt_SHORT_TEXT, 20)
    Call oForm01.Items("RspCd2").Specific.DataBind.SetBound(True, "", "RspCd2")
    '담당코드2_E
    
    '담당명2_S
    Call oForm01.DataSources.UserDataSources.Add("RspNm2", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RspNm2").Specific.DataBind.SetBound(True, "", "RspNm2")
    '담당명2_E
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_GA101_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA101_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_GA101_ComboBox_Setting()
On Error GoTo PS_GA101_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '목표구분
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '10'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    
    Call oForm01.Items("TrgtCls").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("TrgtCls").Specific, sQry, "%", False, False)

    '사업장
    Call oForm01.Items("BPLId").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "%", False, False)
    
    Set oCombo = oForm01.Items("BPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '달성여부
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '11'"
    sQry = sQry & "             AND U_UseYN = 'Y'"

    Call oForm01.Items("CmYN").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CmYN").Specific, sQry, "01", False, False)
    
    
    
    '매트릭스-목표구분
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '10'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("TrgtCls"), sQry)
    
    '매트릭스-사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    '매트릭스-팀1
    sQry = "           SELECT      b.U_Code, "
    sQry = sQry & "                b.U_CodeNm"
    sQry = sQry & " FROM       [@PS_HR200H] a"
    sQry = sQry & "                INNER JOIN"
    sQry = sQry & "                [@PS_HR200L] b"
    sQry = sQry & "                    On a.Code = b.Code"
    sQry = sQry & "                    And a.Name = '부서'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("TeamCd1"), sQry)

    '매트릭스-담당1
    sQry = "           SELECT      b.U_Code, "
    sQry = sQry & "                b.U_CodeNm"
    sQry = sQry & " FROM       [@PS_HR200H] a"
    sQry = sQry & "                INNER JOIN"
    sQry = sQry & "                [@PS_HR200L] b"
    sQry = sQry & "                    On a.Code = b.Code"
    sQry = sQry & "                    And a.Name = '담당'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("RspCd1"), sQry)

    '매트릭스-팀2
    sQry = "           SELECT      b.U_Code, "
    sQry = sQry & "                b.U_CodeNm"
    sQry = sQry & " FROM       [@PS_HR200H] a"
    sQry = sQry & "                INNER JOIN"
    sQry = sQry & "                [@PS_HR200L] b"
    sQry = sQry & "                    On a.Code = b.Code"
    sQry = sQry & "                    And a.Name = '부서'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("TeamCd2"), sQry)

    '매트릭스-담당2
    sQry = "           SELECT      b.U_Code, "
    sQry = sQry & "                b.U_CodeNm"
    sQry = sQry & " FROM       [@PS_HR200H] a"
    sQry = sQry & "                INNER JOIN"
    sQry = sQry & "                [@PS_HR200L] b"
    sQry = sQry & "                    On a.Code = b.Code"
    sQry = sQry & "                    And a.Name = '담당'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("RspCd2"), sQry)
    
    '매트릭스-달성여부
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '11'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CmYN"), sQry)
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_GA101_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA101_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA101_CF_ChooseFromList()
On Error GoTo PS_GA101_CF_ChooseFromList_Error

    Exit Sub
PS_GA101_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA101_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA101_EnableMenus()
On Error GoTo PS_GA101_EnableMenus_Error
    
    Exit Sub
PS_GA101_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA101_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA101_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_GA101_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_GA101_FormItemEnabled
        '//Call PS_GA101_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_GA101_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_GA101_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA101_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA101_FormResize()
On Error GoTo PS_GA101_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_GA101_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA101_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA101_Print_Report01()
On Error GoTo Print_Query_Error
    Dim DocNum As String
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    Dim TrgtCls As String '목표구분
    Dim BPLId As String '사업장
    Dim CmYN As String '달성여부
    Dim TeamCd1 As String '팀1
    Dim RspCd1 As String '담당1
    Dim TeamCd2 As String '팀2
    Dim RspCd2 As String '담당2
    
    TrgtCls = IIf(Trim(oForm01.Items("TrgtCls").Specific.VALUE) = "%", "", Trim(oForm01.Items("TrgtCls").Specific.VALUE)) '목표구분
    BPLId = IIf(Trim(oForm01.Items("BPLId").Specific.VALUE) = "%", "", Trim(oForm01.Items("BPLId").Specific.VALUE)) '사업장
    CmYN = IIf(Trim(oForm01.Items("CmYN").Specific.VALUE) = "%", "", Trim(oForm01.Items("CmYN").Specific.VALUE)) '달성여부
    TeamCd1 = Trim(oForm01.Items("TeamCd1").Specific.VALUE) '팀1
    RspCd1 = Trim(oForm01.Items("RspCd1").Specific.VALUE) '담당1
    TeamCd2 = Trim(oForm01.Items("TeamCd2").Specific.VALUE) '팀2
    RspCd2 = Trim(oForm01.Items("RspCd2").Specific.VALUE) '담당2
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_GA101] 레포트"
    ReportName = "PS_GA101_01.rpt"
 
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
    sQry = "EXEC [PS_GA101_01] '" & TrgtCls & "','" & BPLId & "','" & CmYN & "','" & TeamCd1 & "','" & RspCd1 & "','" & TeamCd2 & "','" & RspCd2 & "'"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_GA101_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub




