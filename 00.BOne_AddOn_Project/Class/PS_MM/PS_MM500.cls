VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_MM500.cls
'//  Module : 재고관리 > 애드온 > 자재출고관리
'//  Desc : 제품검사확인(자재담당)
'//  FormType : PS_MM500
'//  Create Date(Start) : 2014.04.01
'//  Create Date(End) :
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix

Private oDS_PS_MM500H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_MM500L As SAPbouiCOM.DBDataSource '등록라인
Private oDS_PS_MM500M As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private ItemInformationCount As Long

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM500.srf")
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
    
    oFormUniqueID01 = "PS_MM500_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_MM500_CreateItems
    Call PS_MM500_ComboBox_Setting
    Call PS_MM500_CF_ChooseFromList
    Call PS_MM500_EnableMenus
    Call PS_MM500_SetDocument(oFromDocEntry01)
    Call PS_MM500_FormResize

'    Call PS_MM500_Add_MatrixRow01(0, True)
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True
        
    oForm01.Update
    oForm01.Freeze False
    
    '검사일자
    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM01")
    oForm01.Items("ToDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    
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
Sub PS_MM500_Add_MatrixRow01(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_MM500_Add_MatrixRow01_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM500L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_MM500L.Offset = oRow
    oDS_PS_MM500L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM500_Add_MatrixRow01_Error:
    MDC_Com.MDC_GF_Message "PS_MM500_Add_MatrixRow01_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_MM500_MTX01()
'******************************************************************************
'Function ID : PS_MM500_MTX01()
'해당모듈 : PS_MM500
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM500_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim FrDt As String '일자(시작)
    Dim ToDt As String '일자(종료)
    Dim BPLId As String '사업장
    Dim CardCode As String '거래처
    Dim CheckYN As String '확인여부
    
    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE) '검사일자(시작)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE) '검사일자(종료)
    BPLId = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE) '납품처
    CheckYN = Trim(oForm01.Items("CheckYN").Specific.VALUE) '확인여부
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_MM500_01] '"
    sQry = sQry & FrDt & "','" '일자(시작)
    sQry = sQry & ToDt & "','" '일자(종료)
    sQry = sQry & BPLId & "','" '사업장
    sQry = sQry & CardCode & "','" '거래처
    sQry = sQry & CheckYN & "'" '확인여부
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_MM500L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_MM500_Add_MatrixRow01(0, True)
        GoTo PS_MM500_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
    
        If i + 1 > oDS_PS_MM500L.Size Then
            oDS_PS_MM500L.InsertRecord (i)
        End If
    
        oMat01.AddRow
        oDS_PS_MM500L.Offset = i
    
        Call oDS_PS_MM500L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_MM500L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PS_MM500L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)) '작번
        Call oDS_PS_MM500L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명
        Call oDS_PS_MM500L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("Spec").VALUE)) '규격
        Call oDS_PS_MM500L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("SD030_DocDate").VALUE)) '요청일
        Call oDS_PS_MM500L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("LotNo").VALUE)) '업체주문번호
        Call oDS_PS_MM500L.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("Weight").VALUE)) '수량
        Call oDS_PS_MM500L.setValue("U_ColQty02", i, Trim(oRecordSet01.Fields("SjWeight").VALUE)) '수주수량
        Call oDS_PS_MM500L.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("SD030_DocEntry").VALUE)) '요청문서번호
        Call oDS_PS_MM500L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("SD030_LineID").VALUE)) '요청라인
        Call oDS_PS_MM500L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("CheckYN").VALUE)) '확인여부
        
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
PS_MM500_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_MM500_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Function PS_MM500_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_MM500_HeaderSpaceLineDel()
'해당모듈 : PS_MM500
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_MM500_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
'        Case Trim(oForm01.Items("StdYear").Specific.VALUE) = "" '기준년도
'            ErrNum = 1
'            GoTo PS_MM500_HeaderSpaceLineDel_Error
    End Select

    PS_MM500_HeaderSpaceLineDel = True
Exit Function

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM500_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "기준년도는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("FrDate").Click(ct_Regular)
    End If
    PS_MM500_HeaderSpaceLineDel = False
End Function

''/ 메트릭스 필수 사항 check
'Private Function PS_MM500_MatrixSpaceLineDel() As Boolean
'On Error GoTo PS_MM500_MatrixSpaceLineDel_Error
'
'    Dim i As Long
'    Dim ErrNum As Integer
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Dim sQry As String
'
'    Set oRecordSet01 = Nothing
'    PS_MM500_MatrixSpaceLineDel = True
'Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PS_MM500_MatrixSpaceLineDel_Error:
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
'        MDC_Com.MDC_GF_Message "PS_MM500_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'    PS_MM500_MatrixSpaceLineDel = False
'End Function

Private Sub PS_MM500_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_MM500_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    Dim Qty As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        
        Case "CardCode"
        
            sQry = "           SELECT       CardName"
            sQry = sQry & " FROM        [OCRD]"
            sQry = sQry & " WHERE       CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry

            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        
        Case "ItemCode"
            
            sQry = "           SELECT       FrgnName"
            sQry = sQry & " FROM        OITM"
            sQry = sQry & " WHERE       ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry

            oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    
'        Case "BPLId", "OrdGbn", "DocDateFr", "DocDateTo", "CntcCode"
'            oForm01.Freeze True
'            If oUID = "CntcCode" Then
'                sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
'                oRecordSet01.DoQuery sQry
'
'                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'            End If
'
'            oMat01.Clear
'            oDS_PS_USERDS01.Clear
'
'            BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
'            OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.VALUE)
'            DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
'            DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
'            CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
'
'            If BPLId = "" Then BPLId = "%"
'            If OrdGbn = "" Then OrdGbn = "%"
'            If DocDateFr = "" Then DocDateFr = "18990101"
'            If DocDateTo = "" Then DocDateTo = "20991231"
'            If CntcCode = "" Then CntcCode = "%"
'
'            sQry = "EXEC [PS_PP045_01] '" & BPLId & "', '" & OrdGbn & "', '" & DocDateFr & "', '" & DocDateTo & "', '" & CntcCode & "'"
'            oRecordSet01.DoQuery sQry
'
'            If oRecordSet01.RecordCount = 0 Then
'                MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
'                Set oRecordSet01 = Nothing
'                oForm01.Freeze False
'                Exit Sub
'            End If
'
'            Dim ProgBar01 As SAPbouiCOM.ProgressBar
'            Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
'
'            For i = 0 To oRecordSet01.RecordCount - 1
'                If i + 1 > oDS_PS_USERDS01.Size Then
'                    oDS_PS_USERDS01.InsertRecord (i)
'                End If
'
'                oMat01.AddRow
'                oDS_PS_USERDS01.Offset = i
'                oDS_PS_USERDS01.setValue "U_LineNum", i, i + 1
'                oDS_PS_USERDS01.setValue "U_ColDt01", i, Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")
'                oDS_PS_USERDS01.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("U_BPLId").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("U_OrdType").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("U_OrdGbn").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("U_WorkCode").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields("U_WorkName").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg07", i, Trim(oRecordSet01.Fields("U_YTime").VALUE)
''----------------------------------------------------------------------------------------------------------
'                oRecordSet01.MoveNext
'                ProgBar01.VALUE = ProgBar01.VALUE + 1
'                ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
'            Next
'            oMat01.LoadFromDataSource
''            oMat01.AutoResizeColumns
'            ProgBar01.Stop
'            oForm01.Freeze False
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_MM500_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("시각은 숫자만 입력이 가능합니다.", "E")
'    ElseIf ErrNum = 2 Then
'        Call MDC_Com.MDC_GF_Message("시각(시)는 24미만의 값만 입력이 가능합니다.", "E")
'    ElseIf ErrNum = 3 Then
'        Call MDC_Com.MDC_GF_Message("시각(분)은 60미만의 값만 입력이 가능합니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_MM500_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_MM500_FormItemEnabled()
On Error GoTo PS_MM500_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM500_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_MM500_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm01.Freeze(True)
                Call PS_MM500_Add_MatrixRow01(oMat01.VisualRowCount)
                Call oForm01.Freeze(False)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_MM500_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_MM500H.Clear

'                Call PS_MM500_LoadCaption
'                Call PS_MM500_FormItemEnabled
                '//Call PS_MM500_FormItemEnabled '//UDO방식
                '//Call PS_MM500_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_MM500_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_MM500L.RemoveRecord(oDS_PS_MM500L.Size - 1)
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
    
        If pval.ItemUID = "PS_MM500" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnSearch" Then '/조회
            If PS_MM500_HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            Else
                Call PS_MM500_MTX01
            End If
            
        ElseIf pval.ItemUID = "BtnUpdate" Then '확인
        
            Call PS_MM500_Update_SD030
        
        End If
        
        If pval.ItemUID = "BtnConfirm" Then '전체확인
        
            Call PS_MM500_ConfirmAll
            
        End If
        
        If pval.ItemUID = "BtnPrint" Then '출력
            If PS_MM500_HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            Else
                Call PS_MM500_Print_Report01
            End If
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_MM500" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '납품처
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '작번
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

    If pval.Before_Action = True Then
    
    ElseIf pval.Before_Action = False Then
    
        If pval.ItemChanged = True Then
            Call PS_MM500_FlushToItemValue(pval.ItemUID)
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

            Call PS_MM500_FlushToItemValue(pval.ItemUID)
                
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
        Call PS_MM500_FormItemEnabled
        '//Call PS_MM500_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_MM500_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_MM500H", "U_CardCode,U_CardName")
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
'            If (PS_MM500_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_MM500H.RemoveRecord(oDS_PS_MM500H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_MM500_Add_MatrixRow01(0)
            Else
                If Trim(oDS_PS_MM500H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_MM500_Add_MatrixRow01(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_MM500_CreateItems() As Boolean
On Error GoTo PS_MM500_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_MM500H = oForm01.DataSources.DBDataSources("@PS_MM500H")
    Set oDS_PS_MM500L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_Auto
    oMat01.AutoResizeColumns
    
    '일자(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    
    '일자(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
    
    '거래처코드
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardCode").Specific.DataBind.SetBound(True, "", "CardCode")

    '거래처명
    Call oForm01.DataSources.UserDataSources.Add("CardName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CardName").Specific.DataBind.SetBound(True, "", "CardName")

    '확인여부
    Call oForm01.DataSources.UserDataSources.Add("CheckYN", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CheckYN").Specific.DataBind.SetBound(True, "", "CheckYN")

    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_MM500_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM500_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_MM500_ComboBox_Setting()
On Error GoTo PS_MM500_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '사업장
    Call oForm01.Items("BPLID").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '확인여부
'    Call oForm01.Items("ChekYN").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("CheckYN").Specific.ValidValues.Add("N", "미확인")
    Call oForm01.Items("CheckYN").Specific.ValidValues.Add("Y", "확인")
    Call oForm01.Items("CheckYN").Specific.Select("N", psk_ByValue)
    
    '매트릭스
    '확인여부
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM500", "Mat01", "CheckYN", "Y", "확인")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM500", "Mat01", "CheckYN", "N", "미확인")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("CheckYN"), "PS_MM500", "Mat01", "CheckYN")
    
'    '소속반
'    sQry = "           SELECT      T0.U_Code AS [Code],"
'    sQry = sQry & "                T0.U_CodeNm AS [Value]"
'    sQry = sQry & " FROM       [@PS_HR200L] AS T0"
'    sQry = sQry & " WHERE      T0.Code = '9'"
'    sQry = sQry & "                AND T0.U_Char3 = '2'"
'    sQry = sQry & "                AND T0.U_Char2 = '2800'"
'    sQry = sQry & "                AND T0.U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY  T0.U_Seq"
'    Call oForm01.Items("ClsCode").Specific.ValidValues.Add("%", "전체")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("ClsCode").Specific, sQry, "", False, False)
'    Call oForm01.Items("ClsCode").Specific.Select("%", psk_ByValue)
       
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_MM500_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM500_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM500_CF_ChooseFromList()
On Error GoTo PS_MM500_CF_ChooseFromList_Error

    Exit Sub
PS_MM500_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM500_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM500_EnableMenus()
On Error GoTo PS_MM500_EnableMenus_Error
    
    Exit Sub
PS_MM500_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM500_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM500_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_MM500_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_MM500_FormItemEnabled
        '//Call PS_MM500_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_MM500_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_MM500_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM500_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM500_FormResize()
On Error GoTo PS_MM500_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_MM500_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM500_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM500_Print_Report01()
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
    
    Dim ChkDateF As String '검사일자(시작)
    Dim ChkDateT As String '검사일자(종료)
    Dim CardCode As String '납품처
    Dim ItemCode As String '작번
    Dim DrawNo As String '도면번호
    
    ChkDateF = Trim(oForm01.Items("ChkDateF").Specific.VALUE)
    ChkDateT = Trim(oForm01.Items("ChkDateT").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    DrawNo = Trim(oForm01.Items("DrawNo").Specific.VALUE)
    
    '조회자사번
    Dim CntcCode As String
    CntcCode = MDC_PS_Common.User_MSTCOD
    
    '임시테이블에 check된항목저장하기 위한 기 정보 삭제
    sQry = "DELETE [Z_PS_MM500] WHERE CntcCode = '" & CntcCode & "'"
    Call oRecordSet.DoQuery(sQry)
    
    '임시테이블에 check된항목저장
    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_MM500L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "INSERT INTO [Z_PS_MM500] VALUES ('" & CntcCode & "', '" & Trim(oDS_PS_MM500L.GetValue("U_ColReg02", i)) & "')"
            Call oRecordSet.DoQuery(sQry)
        End If
    Next
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_MM500] 검사성적서"
    ReportName = "PS_MM500_01.rpt"
 
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
    sQry = "         EXEC [PS_MM500_02] '"
'    sQry = sQry & ChkDateF & "','" '검사일자(시작)
'    sQry = sQry & ChkDateT & "','" '검사일자(종료)
'    sQry = sQry & CardCode & "','" '납품처
'    sQry = sQry & ItemCode & "','" '작번
'    sQry = sQry & DrawNo & "','" '도면번호
    sQry = sQry & CntcCode & "'" '사원번호
    

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
        MDC_Com.MDC_GF_Message "PS_MM500_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub PS_MM500_Update_SD030()
On Error GoTo MM500_Update_SD030_Error

    Dim sQry As String
    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Dim SD030_DocEntry As Integer '문서번호(출하요청)
    Dim SD030_LineID As Integer '라인번호(출하요청)
    Dim CheckYN As String '확인여부
    
    Call oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_MM500L.GetValue("U_ColReg01", i)) = "Y" Then
        
            SD030_DocEntry = oDS_PS_MM500L.GetValue("U_ColReg09", i)
            SD030_LineID = oDS_PS_MM500L.GetValue("U_ColReg10", i)
            CheckYN = oDS_PS_MM500L.GetValue("U_ColReg11", i)
          
            sQry = "         EXEC [PS_MM500_02] '"
            sQry = sQry & SD030_DocEntry & "','" '문서번호
            sQry = sQry & SD030_LineID & "','" '라인번호
            sQry = sQry & CheckYN & "'" '확인여부
    
            Call oRecordSet.DoQuery(sQry)
            
        End If
    Next
    
    Call MDC_Com.MDC_GF_Message("검사여부 확인 완료", "S")
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MM500_Update_SD030_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "수정하는 중 오류가 발생했습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_MM500_Update_SD030_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub PS_MM500_ConfirmAll()
On Error GoTo PS_MM500_ConfirmAll_Error

    Dim ConfirmType As String
    Dim loopCount As Integer
    
    Call oForm01.Freeze(True)
    ConfirmType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_MM500L.GetValue("U_ColReg11", loopCount)) = "N" Then
            ConfirmType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PS_MM500L.Offset = loopCount
        If ConfirmType = "N" Then
            Call oDS_PS_MM500L.setValue("U_ColReg11", loopCount, "Y")
        Else
            Call oDS_PS_MM500L.setValue("U_ColReg11", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)

    Exit Sub
    
PS_MM500_ConfirmAll_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_MM500_ConfirmAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub







