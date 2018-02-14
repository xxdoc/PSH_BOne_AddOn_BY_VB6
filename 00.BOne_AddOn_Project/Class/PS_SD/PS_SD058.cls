VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD058"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_SD058.cls
'//  Module : 판매관리 > 애드온 > 기성매출관리
'//  Desc : 외주제작진도관리대상업체조회
'//  FormType : PS_SD058
'//  Create Date(Start) : 2014.06.23
'//  Create Date(End) : 2014.06.23
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_SD058H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_SD058L As SAPbouiCOM.DBDataSource '등록라인

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
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD058.srf")
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
    
    oFormUniqueID01 = "PS_SD058_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    Call oForm01.Freeze(True)
    
    Call PS_SD058_CreateItems
    Call PS_SD058_ComboBox_Setting
    Call PS_SD058_CF_ChooseFromList
    Call PS_SD058_EnableMenus
    Call PS_SD058_SetDocument(oFromDocEntry01)
    Call PS_SD058_FormResize

'    Call PS_SD058_Add_MatrixRow(0, True)
    Call PS_SD058_FormItemEnabled
    Call PS_SD058_Initial_Setting
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True
    
    Call oForm01.Update
    Call oForm01.Freeze(False)
    
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

'/메트릭스 Row추가
Sub PS_SD058_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD058_Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PS_SD058L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_SD058L.Offset = oRow
    Call oDS_PS_SD058L.setValue("U_LineNum", oRow, oRow + 1)
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD058_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_SD058_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_SD058_MTX01()
'******************************************************************************
'Function ID : PS_SD058_MTX01()
'해당모듈 : PS_SD058
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD058_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim StdYM As String '기준년월
    Dim StdAmt As Double '기준품의금액
    Dim OrdNum As String '작번
    Dim CardCode As String '외주업체
        
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    StdYM = Trim(oForm01.Items("StdYM").Specific.VALUE) '기준년월
    StdAmt = Trim(oForm01.Items("StdAmt").Specific.VALUE) '기준품의금액
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '작번
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE) '사용자사번
    
    sQry = "         EXEC [PS_SD058_01] '"
    sQry = sQry & StdYM & "','"
    sQry = sQry & StdAmt & "','"
    sQry = sQry & OrdNum & "','"
    sQry = sQry & CardCode & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_SD058L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_SD058_Add_MatrixRow(0, True)
        
        GoTo PS_SD058_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_SD058L.Size Then
            oDS_PS_SD058L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_SD058L.Offset = i
        
        Call oDS_PS_SD058L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_SD058L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작번(메인)
        Call oDS_PS_SD058L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("FrgnName").VALUE)) '품명(메인)
        Call oDS_PS_SD058L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("Spec").VALUE)) '규격(메인)
        Call oDS_PS_SD058L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("InOut").VALUE)) '자체/외주
        Call oDS_PS_SD058L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)) '품목코드(자재)
        Call oDS_PS_SD058L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명(자재)
        Call oDS_PS_SD058L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("ItemSpec").VALUE)) '규격(자재)
        Call oDS_PS_SD058L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("CardCode").VALUE)) '업체코드
        Call oDS_PS_SD058L.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("CardName").VALUE)) '업체명
        Call oDS_PS_SD058L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("TelNo").VALUE)) '전화번호
        Call oDS_PS_SD058L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("CellNo").VALUE)) '핸드폰번호
        Call oDS_PS_SD058L.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("PreStdYM").VALUE)) '직전등록년월
        Call oDS_PS_SD058L.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("PreCnt").VALUE)) '직전등록회차
        '직전등록회차명
        Call oDS_PS_SD058L.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("PreRR").VALUE)) '직전진도율
        Call oDS_PS_SD058L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("ORDRAmt").VALUE)) '수주금액
        Call oDS_PS_SD058L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("MM030Amt").VALUE)) '품의금액
        Call oDS_PS_SD058L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("MM030Date").VALUE), "YYYYMMDD")) '품의일자
        Call oDS_PS_SD058L.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("MM030Due").VALUE), "YYYYMMDD")) '품의납품일자
        Call oDS_PS_SD058L.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("MM030Entry").VALUE)) '품의문서번호
        Call oDS_PS_SD058L.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("MM030Line").VALUE)) '품의라인번호
           
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
PS_SD058_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_SD058_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Function PS_SD058_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_SD058_HeaderSpaceLineDel()
'해당모듈 : PS_SD058
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD058_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("BPLID").Specific.VALUE) = "%" '사업장
            ErrNum = 1
            GoTo PS_SD058_HeaderSpaceLineDel_Error
    End Select

    PS_SD058_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD058_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("BPLID").Click(ct_Regular)
    End If
    PS_SD058_HeaderSpaceLineDel = False
End Function

Private Function PS_SD058_CheckBeforeSearch(ByVal pItemUID As String) As Boolean
'******************************************************************************
'Function ID : PS_SD058_CheckBeforeSearch()
'해당모듈 : PS_SD058
'기능 : 필수입력사항 체크
'인수 : pItemUID(아이템 UID)
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD058_CheckBeforeSearch_Error

    Dim ErrNum As Integer
    ErrNum = 0

    If pItemUID = "BtnSearch1" Then

        Select Case True
'            Case Trim(oForm01.Items("OrdNum").Specific.VALUE) = "" '작번
'                ErrNum = 1
'                GoTo PS_SD058_CheckBeforeSearch_Error '작번은 제외
            Case Trim(oForm01.Items("StdYM").Specific.VALUE) = "" '기준년월
                ErrNum = 2
                GoTo PS_SD058_CheckBeforeSearch_Error
            Case Trim(oForm01.Items("StdCnt").Specific.VALUE) = "%" '기준회차
                ErrNum = 3
                GoTo PS_SD058_CheckBeforeSearch_Error
        End Select
    
    ElseIf pItemUID = "BtnSearch2" Then
        
        Select Case True
'            Case Trim(oForm01.Items("OrdNumS").Specific.VALUE) = "" '작번
'                ErrNum = 4
'                GoTo PS_SD058_CheckBeforeSearch_Error '작번은 제외
            Case Trim(oForm01.Items("StdYMS").Specific.VALUE) = "" '기준년월
                ErrNum = 5
                GoTo PS_SD058_CheckBeforeSearch_Error
            Case Trim(oForm01.Items("StdCntS").Specific.VALUE) = "%" '기준회차
                ErrNum = 6
                GoTo PS_SD058_CheckBeforeSearch_Error
        End Select
    
    End If

    PS_SD058_CheckBeforeSearch = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD058_CheckBeforeSearch_Error:
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("입력정보의 작번은 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("OrdNum").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        Call MDC_Com.MDC_GF_Message("입력정보의 기준년월은 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("StdYM").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        Call MDC_Com.MDC_GF_Message("입력정보의 기준회차는 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("StdCnt").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        Call MDC_Com.MDC_GF_Message("조회정보의 작번은 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("OrdNumS").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        Call MDC_Com.MDC_GF_Message("조회정보의 기준년월은 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("StdYMS").Click(ct_Regular)
    ElseIf ErrNum = 6 Then
        Call MDC_Com.MDC_GF_Message("조회정보의 기준회차는 필수사항입니다. 확인하세요.", "E")
        Call oForm01.Items("StdCntS").Click(ct_Regular)
    End If
    PS_SD058_CheckBeforeSearch = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PS_SD058_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_SD058_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PS_SD058_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD058_MatrixSpaceLineDel_Error:
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
        MDC_Com.MDC_GF_Message "PS_SD058_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_SD058_MatrixSpaceLineDel = False
End Function

Private Sub PS_SD058_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_SD058_FlushToItemValue_Error

    Dim Price As Double '단가
    Dim RWorkTime As Double '예상추가공수
    Dim Amount As Double '금액
    Dim loopCount As Integer
    Dim totalAmt As Double '금액 합계

    Select Case oUID

        Case "Mat01"
        
'            If oCol = "RWorkTime" Then
'
'                Call oMat01.FlushToDataSource
'
'                RWorkTime = oDS_PS_SD058L.GetValue("U_ColQty01", oRow - 1)
'                Price = oDS_PS_SD058L.GetValue("U_ColSum01", oRow - 1)
'                Amount = RWorkTime * Price
'                Call oDS_PS_SD058L.setValue("U_ColSum02", oRow - 1, Amount)
'
'                For loopCount = 0 To oMat01.VisualRowCount - 1
'
'                    totalAmt = totalAmt + oDS_PS_SD058L.GetValue("U_ColSum02", loopCount)
'
'                Next
'
'                oForm01.Items("Total").Specific.VALUE = totalAmt
'
'                Call oMat01.LoadFromDataSource
'
'            End If
'
'            Call oMat01.AutoResizeColumns

        Case "OrdNum"
        
            oForm01.Items("FrgnName").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
            oForm01.Items("Spec").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
            
        Case "CardCode"
        
            oForm01.Items("CardName").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
        
    End Select

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_SD058_FlushToItemValue_Error:

    Call MDC_Com.MDC_GF_Message("PS_SD058_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PS_SD058_FormItemEnabled()
On Error GoTo PS_SD058_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD058_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_SD058_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub PS_SD058_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD058_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_SD058L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD058L.Offset = oRow
'    oDS_PS_SD058L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD058_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD058_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                
                'oForm01.Items("GCode").Click ct_Regular
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_SD058_AddMatrixRow(oMat01.VisualRowCount)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_SD058_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_SD058H.Clear

'                Call PS_SD058_LoadCaption
'                Call PS_SD058_FormItemEnabled
                '//Call PS_SD058_FormItemEnabled '//UDO방식
                '//Call PS_SD058_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_SD058_FormItemEnabled
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_SD058L.RemoveRecord(oDS_PS_SD058L.Size - 1)
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
    
        If pval.ItemUID = "PS_SD058" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
            
        If pval.ItemUID = "BtnSearch" Then '/조회

            If PS_SD058_CheckBeforeSearch(pval.ItemUID) = False Then
                BubbleEvent = False
                Exit Sub
            End If

            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_SD058_MTX01
        
        ElseIf pval.ItemUID = "BtnPrint" Then '출력
        
            Call PS_SD058_Print_Report01
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        
        If pval.ItemUID = "BtnAdd" Or pval.ItemUID = "BtnDelete" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    
        If pval.ItemUID = "PS_SD058" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    If pval.BeforeAction = True Then
    
        
        
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row > 0 Then
            
                Call oMat01.SelectRow(pval.Row, True, False)
                
            End If
        End If
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PS_SD058_FlushToItemValue(pval.ItemUID)
    
    End If

    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row = 0 Then
            
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource
            
            End If
        
        End If
    
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
            
                Call PS_SD058_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
        
            Else
        
                Call PS_SD058_FlushToItemValue(pval.ItemUID)
            
            End If
                
        End If
        
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
        
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
        Call PS_SD058_FormItemEnabled
        '//Call PS_SD058_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_SD058_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_SD058H", "U_CardCode,U_CardName")
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
        
            
        ElseIf pval.BeforeAction = False Then
            
'            If oMat01.RowCount = 0 Then
'                Call PS_SD058_Add_MatrixRow(0)
'            Else
'                If Trim(oDS_PS_SD058L.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_SD058_Add_MatrixRow(oMat01.RowCount)
'                End If
'            End If

        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SD058_CreateItems() As Boolean
On Error GoTo PS_SD058_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_SD058H = oForm01.DataSources.DBDataSources("@PS_SD058H")
    Set oDS_PS_SD058L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '기준년월
    Call oForm01.DataSources.UserDataSources.Add("StdYM", dt_SHORT_TEXT, 6)
    Call oForm01.Items("StdYM").Specific.DataBind.SetBound(True, "", "StdYM")
    
    '기준품의금액
    Call oForm01.DataSources.UserDataSources.Add("StdAmt", dt_SUM)
    Call oForm01.Items("StdAmt").Specific.DataBind.SetBound(True, "", "StdAmt")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdNum").Specific.DataBind.SetBound(True, "", "OrdNum")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("FrgnName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("FrgnName").Specific.DataBind.SetBound(True, "", "FrgnName")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("Spec", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Spec").Specific.DataBind.SetBound(True, "", "Spec")
    
    '외주업체
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardCode").Specific.DataBind.SetBound(True, "", "CardCode")
    
    '외주업체명
    Call oForm01.DataSources.UserDataSources.Add("CardName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CardName").Specific.DataBind.SetBound(True, "", "CardName")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_SD058_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD058_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_SD058_ComboBox_Setting()
On Error GoTo PS_SD058_ComboBox_Setting_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

'    '//////////입력정보//////////S
'    '기준회차
'    Call oForm01.Items("StdCnt").Specific.ValidValues.Add("%", "선택")
'    sQry = "           SELECT      U_Minor AS [Code],"
'    sQry = sQry & "                U_CdName AS [Name]"
'    sQry = sQry & " FROM       [@PS_SY001L]"
'    sQry = sQry & " WHERE      Code = 'S008'"
'    sQry = sQry & "                AND U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY  U_Seq"
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("StdCnt").Specific, sQry, "", False, False)
'    Call oForm01.Items("StdCnt").Specific.Select("%", psk_ByValue)
'    '//////////입력정보//////////E
'
'    '//////////조회정보//////////S
'    '기준회차
'    Call oForm01.Items("StdCntS").Specific.ValidValues.Add("%", "전체")
'    sQry = "           SELECT      U_Minor AS [Code],"
'    sQry = sQry & "                U_CdName AS [Name]"
'    sQry = sQry & " FROM       [@PS_SY001L]"
'    sQry = sQry & " WHERE      Code = 'S008'"
'    sQry = sQry & "                AND U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY  U_Seq"
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("StdCntS").Specific, sQry, "", False, False)
'    Call oForm01.Items("StdCntS").Specific.Select("%", psk_ByValue)
'    '//////////조회정보//////////E
    
    '//////////매트릭스//////////S
    '기준회차
    sQry = "           SELECT      U_Minor AS [Code],"
    sQry = sQry & "                U_CdName AS [Name]"
    sQry = sQry & " FROM       [@PS_SY001L]"
    sQry = sQry & " WHERE      Code = 'S008'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("PreCnt"), sQry)

'    '품의구분
'    sQry = "           SELECT       Code AS [Code],"
'    sQry = sQry & "                 Name AS [Name]"
'    sQry = sQry & " FROM        [@PSH_ORDTYP]"
'    sQry = sQry & " WHERE       Code IN ('10','20','30','40')" '4개 품의대해서만 조회
'    sQry = sQry & " ORDER BY   Code"
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("POType"), sQry)
    '//////////매트릭스//////////E

    Call oForm01.Freeze(False)
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_SD058_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_SD058_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD058_CF_ChooseFromList()
On Error GoTo PS_SD058_CF_ChooseFromList_Error

    Exit Sub
PS_SD058_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD058_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD058_EnableMenus()
On Error GoTo PS_SD058_EnableMenus_Error
    
    Exit Sub
PS_SD058_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD058_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD058_Initial_Setting()
On Error GoTo PS_SD058_Initial_Setting_Error
    
    '컬럼 Hidden 처리
'    oMat01.Columns("StdYM").Visible = False '기준년월
'    oMat01.Columns("StdCnt").Visible = False '기준회차
    
    oForm01.Items("StdYM").Specific.VALUE = Format(Now, "YYYYMM")
    oForm01.Items("StdAmt").Specific.VALUE = "80000000"
    
    Call oForm01.Items("OrdNum").Click
    
    Exit Sub
PS_SD058_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD058_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD058_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_SD058_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_SD058_FormItemEnabled
        '//Call PS_SD058_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_SD058_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_SD058_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD058_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD058_FormResize()
On Error GoTo PS_SD058_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_SD058_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD058_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD058_Print_Report01()
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
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
            
    Call ConnectODBC
    
    Dim StdYM As String '기준년월
    Dim StdAmt As Double '기준품의금액
    Dim OrdNum As String '작번
    Dim CardCode As String '외주업체
    
    StdYM = Trim(oForm01.Items("StdYM").Specific.VALUE) '기준년월
    StdAmt = Trim(oForm01.Items("StdAmt").Specific.VALUE) '기준품의금액
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '작번
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE) '사용자사번
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
 
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
    
    WinTitle = "[PS_SD058] 외주제작 관리 대상 업체"
'    If PrintOpt = "01" Then '전체 리스트 출력
    
        ReportName = "PS_SD058_01.rpt"
        
        sQry = "         EXEC [PS_SD058_01] '"
        sQry = sQry & StdYM & "','" '기준년월
        sQry = sQry & StdAmt & "','" '기준품의금액
        sQry = sQry & OrdNum & "','" '작번
        sQry = sQry & CardCode & "'" '외주업체
        
'    ElseIf PrintOpt = "02" Then '작번별 리스트 출력
'
'        ReportName = "PS_SD058_02.rpt"
'
'        sQry = "         EXEC [PS_SD058_05] '"
'        sQry = sQry & CntcCode & "','" '사원번호
'        sQry = sQry & StdYMS & "','" '기준년월
'        sQry = sQry & OrdNumS & "'" '작번
'
'    End If
    
'    oRecordSet.DoQuery sQry
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
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_SD058_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
