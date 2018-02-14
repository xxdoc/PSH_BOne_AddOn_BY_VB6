VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_CO001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_CO001.cls
'//  Module             : 재무관리 > 애드온 > 원가관리 > 마스터관리
'//  Desc               : 결산마감관리
'//  FormType           : PS_CO001
'//  Create Date(Start) : 2017.11.21
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

Private oDS_PS_CO001L  As SAPbouiCOM.DBDataSource '라인

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
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_CO001.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_CO001_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_CO001_CreateItems
    Call PS_CO001_ComboBox_Setting
    Call PS_CO001_Initial_Setting
    Call PS_CO001_CF_ChooseFromList
    Call PS_CO001_EnableMenus
    Call PS_CO001_SetDocument(oFromDocEntry01)
    Call PS_CO001_FormResize

'    Call PS_CO001_Add_MatrixRow(0, True)
    Call PS_CO001_LoadCaption
    Call PS_CO001_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

'    Dim sQry As String
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    '관리번호
'    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_CO001H]"
'    Call RecordSet01.DoQuery(sQry)
'    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'        Call oDS_PS_CO001H.setValue("DocEntry", 0, 1)
'    Else
'        Call oDS_PS_CO001H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
'    End If
    
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

Private Sub PS_CO001_LoadCaption()
'******************************************************************************
'Function ID : PS_CO001_LoadCaption()
'해당모듈    : PS_CO001
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_CO001_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
    
        oForm01.Items("BtnAdd").Specific.Caption = "저장"

'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"

    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
    
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
        
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_CO001_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_CO001_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_CO001_Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PS_CO001L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_CO001L.offset = oRow
    oDS_PS_CO001L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_CO001_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_CO001_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_CO001_MTX01()
'******************************************************************************
'Function ID : PS_CO001_MTX01()
'해당모듈    : PS_CO001
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_CO001_MTX01_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID      As String '사업장
    Dim StdYear    As String '기준년도
    
    BPLID = oForm01.Items("BPLID").Specific.Selected.VALUE
    StdYear = oForm01.Items("StdYear").Specific.VALUE
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "      EXEC [PS_CO001_01] '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & StdYear & "'"
    
    Call RecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_CO001L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (RecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_CO001_Add_MatrixRow(0, True)
        Call PS_CO001_LoadCaption
        
        GoTo PS_CO001_MTX01_Error
        
        Exit Sub
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount + 1 > oDS_PS_CO001L.Size Then
            Call oDS_PS_CO001L.InsertRecord(loopCount)
        End If
        
        Call oMat01.AddRow
        oDS_PS_CO001L.offset = loopCount
        
        Call oDS_PS_CO001L.setValue("U_LineNum", loopCount, loopCount + 1)
        Call oDS_PS_CO001L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("StdMonth").VALUE) '기준월
        Call oDS_PS_CO001L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("MM005Sts").VALUE) '구매요청
        Call oDS_PS_CO001L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("PP040Sts").VALUE) '작업일보등록
        Call oDS_PS_CO001L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("PP080Sts").VALUE) '생산완료
        Call oDS_PS_CO001L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("SD040Sts").VALUE) '납품처리
        Call oDS_PS_CO001L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("ARSts").VALUE) 'AR송장
        Call oDS_PS_CO001L.setValue("U_ColReg07", loopCount, RecordSet01.Fields("APSts").VALUE) 'AP송장
        Call oDS_PS_CO001L.setValue("U_ColReg08", loopCount, RecordSet01.Fields("OIGESts").VALUE) '출고
        Call oDS_PS_CO001L.setValue("U_ColReg09", loopCount, RecordSet01.Fields("Comments").VALUE) '비고
        
        RecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set RecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_CO001_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set RecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_CO001_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_CO001_AddData()
'******************************************************************************
'Function ID : PS_CO001_AddData()
'해당모듈    : PS_CO001
'기능        : 데이터 INSERT
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_CO001_AddData_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim StdYear  As String '년도
    Dim BPLID    As String '사업장
    Dim StdMonth As String '기준월
    Dim MM005Sts As String '구매요청
    Dim PP040Sts As String '작업일보등록
    Dim PP080Sts As String '생산완료
    Dim SD040Sts As String '납품처리
    Dim ARSts    As String 'AR송장
    Dim APSts    As String 'AP송장
    Dim OIGESts  As String '출고
    Dim Comments As String '비고
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.RowCount - 1
    
        BPLID = oForm01.Items("BPLID").Specific.VALUE '사업장
        StdYear = oForm01.Items("StdYear").Specific.VALUE '년도
        
        StdMonth = Trim(oDS_PS_CO001L.GetValue("U_ColReg01", loopCount)) '기준월
        MM005Sts = Trim(oDS_PS_CO001L.GetValue("U_ColReg02", loopCount)) '구매요청
        PP040Sts = Trim(oDS_PS_CO001L.GetValue("U_ColReg03", loopCount)) '작업일보등록
        PP080Sts = Trim(oDS_PS_CO001L.GetValue("U_ColReg04", loopCount)) '생산완료
        SD040Sts = Trim(oDS_PS_CO001L.GetValue("U_ColReg05", loopCount)) '납품처리
        ARSts = Trim(oDS_PS_CO001L.GetValue("U_ColReg06", loopCount)) 'AR송장
        APSts = Trim(oDS_PS_CO001L.GetValue("U_ColReg07", loopCount)) 'AP송장
        OIGESts = Trim(oDS_PS_CO001L.GetValue("U_ColReg08", loopCount)) '출고
        Comments = Trim(oDS_PS_CO001L.GetValue("U_ColReg09", loopCount)) '비고
        
        sQry = "      EXEC [PS_CO001_02] '"
        sQry = sQry & BPLID & "','" '사업장
        sQry = sQry & StdYear & "','" '년도
        sQry = sQry & StdMonth & "','" '기준월
        sQry = sQry & MM005Sts & "','" '구매요청
        sQry = sQry & PP040Sts & "','" '작업일보등록
        sQry = sQry & PP080Sts & "','" '생산완료
        sQry = sQry & SD040Sts & "','" '납품처리
        sQry = sQry & ARSts & "','" 'AR송장
        sQry = sQry & APSts & "','" 'AP송장
        sQry = sQry & OIGESts & "','" '출고
        sQry = sQry & Comments & "'" '비고
        
        Call RecordSet01.DoQuery(sQry)
    
    Next

    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    
    Exit Sub

PS_CO001_AddData_Error:

    If ErrNum = 1 Then
    
    ElseIf ErrNum = 99 Then
        
    Else
    
        Call MDC_Com.MDC_GF_Message("PS_CO001_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
        
    End If
    
    Set RecordSet01 = Nothing
    
End Sub

Private Function PS_CO001_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_CO001_HeaderSpaceLineDel()
'해당모듈    : PS_CO001
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_CO001_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("BPLID").Specific.VALUE) = "%" '사업장
            ErrNum = 1
            GoTo PS_CO001_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("StdYear").Specific.VALUE) = "" '기준년도
            ErrNum = 2
            GoTo PS_CO001_HeaderSpaceLineDel_Error
    End Select

    PS_CO001_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_CO001_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장을 선택하지 않았습니다. 확인하세요.", "E"
        Call oForm01.Items("BPLID").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "기준년도을 입력하지 않았습니다.. 확인하세요.", "E"
        Call oForm01.Items("StdYear").Click(ct_Regular)
    End If
    PS_CO001_HeaderSpaceLineDel = False
End Function

Private Sub PS_CO001_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_CO001_FlushToItemValue_Error

    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim sQry     As String
    Dim ItemCode As String
    Dim Qty      As Integer
    Dim Total    As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
    
        Case "Mat01"
            
            Call oMat01.FlushToDataSource
            
    End Select

    Set oRecordSet01 = Nothing

    Exit Sub
PS_CO001_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("에러 메시지", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_CO001_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_CO001_FormItemEnabled()
On Error GoTo PS_CO001_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_CO001_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_CO001_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
                
'                Call PS_CO001_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_CO001_LoadCaption
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_CO001_Add_MatrixRow(oMat01.VisualRowCount)
                
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_CO001_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_CO001H.Clear

'                Call PS_CO001_LoadCaption
'                Call PS_CO001_FormItemEnabled
                '//Call PS_CO001_FormItemEnabled '//UDO방식
                '//Call PS_CO001_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_CO001_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_CO001L.RemoveRecord(oDS_PS_CO001L.Size - 1)
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
    
        If pval.ItemUID = "PS_CO001" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '추가(수정) 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_CO001_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_CO001_AddData

'                Call PS_CO001_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_CO001_LoadCaption
                Call PS_CO001_MTX01
                
                'oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
'                If PS_CO001_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                '쿼리에서 로직 처리
                '여러 행을 동시에 처리하므로 DocEntry별로 행이 존재하면 UPDATE, 존재하지 않으면 INSERT
                Call PS_CO001_AddData
                    
'                Call PS_CO001_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_CO001_LoadCaption
                Call PS_CO001_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSrch" Then

'            Call PS_CO001_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            If PS_CO001_HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            End If
            
            Call PS_CO001_LoadCaption
            Call PS_CO001_MTX01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_CO001" Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode01", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode02", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode03", "") '//사용자값활성
        
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
    
        Call PS_CO001_FlushToItemValue(pval.ItemUID)
    
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
        
            If pval.ItemUID = "Mat01" Then
            
                If pval.ColUID = "Month01" Or pval.ColUID = "Month02" Or pval.ColUID = "Month03" Or pval.ColUID = "Month04" Or pval.ColUID = "Month05" Or pval.ColUID = "Month06" Or pval.ColUID = "Month07" Or pval.ColUID = "Month08" Or pval.ColUID = "Month09" Or pval.ColUID = "Month10" Or pval.ColUID = "Month11" Or pval.ColUID = "Month12" Then
            
                    Call PS_CO001_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                
                End If
            
            Else
    
                Call PS_CO001_FlushToItemValue(pval.ItemUID)
            
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
        Call PS_CO001_FormItemEnabled
        '//Call PS_CO001_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    
    Exit Sub
    
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error

    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
    
        Call PS_CO001_FormResize
        
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_CO001H", "U_CardCode,U_CardName")
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
'            If (PS_CO001_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
            
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_CO001L.RemoveRecord(oDS_PS_CO001L.Size - 1)
            oMat01.LoadFromDataSource
            
'            If oMat01.RowCount = 0 Then
'                Call PS_CO001_Add_MatrixRow(0)
'            Else
'                If Trim(oDS_PS_CO001L.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_CO001_Add_MatrixRow(oMat01.RowCount)
'                End If
'            End If

        End If
    End If
    
    Exit Sub
    
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_CO001_CreateItems() As Boolean
On Error GoTo PS_CO001_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_CO001L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
    '사업장_E
    
    '기준년도
    Call oForm01.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 4)
    Call oForm01.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
    
PS_CO001_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO001_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_CO001_ComboBox_Setting()
On Error GoTo PS_CO001_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '//////////기본정보//////////S
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    '매트릭스
    '상태
    sQry = "        SELECT      B.U_Minor, "
    sQry = sQry & "             b.U_CdName"
    sQry = sQry & " FROM        [@PS_SY001H] A"
    sQry = sQry & "             INNER JOIN"
    sQry = sQry & "             [@PS_SY001L] B"
    sQry = sQry & "                 ON A.Code = B.Code"
    sQry = sQry & " WHERE       A.Code = 'F006'"
    
    '구매요청
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("MM005Sts"), sQry)
    
    '작업일보등록
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("PP040Sts"), sQry)
    
    '생산완료
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("PP080Sts"), sQry)
    
    '납품처리
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("SD040Sts"), sQry)
    
    'AR송장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ARSts"), sQry)
    
    'AP송장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("APSts"), sQry)
    
    '출고
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("OIGESts"), sQry)
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_CO001_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO001_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_CO001_Initial_Setting()
On Error GoTo PS_CO001_Initial_Setting_Error
    '사업장 사용자의 소속 사업장 선택
    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 
    '날짜 설정
'    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYY") + "1231"
'    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYY") + "0101"
    oForm01.Items("StdYear").Specific.VALUE = Format(Now, "YYYY")

'    '년도, 사업장 코드는 Hidden 처리
'    oMat01.Columns("StdYear").Visible = False
'    oMat01.Columns("BPLID").Visible = False

    Exit Sub
PS_CO001_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO001_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_CO001_CF_ChooseFromList()
On Error GoTo PS_CO001_CF_ChooseFromList_Error

    Exit Sub
    
PS_CO001_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO001_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_CO001_EnableMenus()
On Error GoTo PS_CO001_EnableMenus_Error
    
    Exit Sub
    
PS_CO001_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO001_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_CO001_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_CO001_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_CO001_FormItemEnabled
        '//Call PS_CO001_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_CO001_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_CO001_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO001_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_CO001_FormResize()
On Error GoTo PS_CO001_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
    
PS_CO001_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO001_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_CO001_FormReset()
'******************************************************************************
'Function ID : PS_CO001_FormReset()
'해당모듈    : PS_CO001
'기능        : 화면 초기화
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_CO001_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
'    '관리번호
'    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_CO001H]"
'    Call RecordSet01.DoQuery(sQry)
'
'    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'        Call oDS_PS_CO001H.setValue("DocEntry", 0, 1)
'    Else
'        Call oDS_PS_CO001H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
'    End If
    
'    Call oDS_PS_CO001H.setValue("DocEntry", 0, "") '관리번호
'    Call oDS_PS_CO001H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId) '사업장
'    Call oDS_PS_CO001H.setValue("U_ZGODAY", 0, Format(Date, "YYYYMMDD")) '기준일자
'    Call oDS_PS_CO001H.setValue("U_ZEMPIE", 0, "") '관련자사번
'    Call oDS_PS_CO001H.setValue("U_ZEMPNM", 0, "") '관련자성명
'    Call PS_CO001_FlushToItemValue("BPLID") '관련자사업장에 따른 팀 세팅
'    Call oDS_PS_CO001H.setValue("U_ZDPTCD", 0, "%") '관련자팀
'    Call PS_CO001_FlushToItemValue("ZDPTCD") '관련자팀에 따른 담당 세팅
'    Call oDS_PS_CO001H.setValue("U_ZSECCD", 0, "%") '관련자담당
'    Call oDS_PS_CO001H.setValue("U_ZPOINT", 0, "%") '장소
'    Call oDS_PS_CO001H.setValue("U_ZUSRNN", 0, "") '방문자성명
'    Call oDS_PS_CO001H.setValue("U_ZBIRDT", 0, "") '생년월일
'    Call oDS_PS_CO001H.setValue("U_ZBUSNM", 0, "") '업체명
'    Call oDS_PS_CO001H.setValue("U_ZRESID", 0, "") '주소
'    Call oDS_PS_CO001H.setValue("U_ZGOTIME", 0, "") '출입시간(Fr)
'    Call oDS_PS_CO001H.setValue("U_ZOUTIME", 0, "") '출입시간(To)
'    Call oDS_PS_CO001H.setValue("U_ZCARJO", 0, "%") '차량종류
'    Call oDS_PS_CO001H.setValue("U_ZCARNO", 0, "") '차량번호
'    Call oDS_PS_CO001H.setValue("U_ZMOKJU", 0, "%") '출입목적
'    Call oDS_PS_CO001H.setValue("U_ZSPCIU", 0, "") '방문객비고
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

    Exit Sub

PS_CO001_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_CO001_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub





