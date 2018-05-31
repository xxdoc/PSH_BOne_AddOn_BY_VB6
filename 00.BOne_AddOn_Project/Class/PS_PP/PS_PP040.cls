VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP040.cls
'//  Module         : 생산관리 > 작업일보등록(작지)
'//  Desc           : 작업일보등록
'//  FormType       : PS_PP040
'//  Create Date    : 2010.08.20
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix
Public oMat03 As SAPbouiCOM.Matrix
Private oDS_PS_PP040H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP040L As SAPbouiCOM.DBDataSource '등록라인
Private oDS_PS_PP040M As SAPbouiCOM.DBDataSource '등록라인
Private oDS_PS_PP040N As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMat01Row01 As Long
Private oMat02Row02 As Long
Private oMat03Row03 As Long

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    BatchNum As String
    Quantity As Long
    OPORNo As Long
    POR1No As Long
    Check As Boolean
    OPDNNo As Long
    PDN1No As Long
End Type
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

Private oDocType01 As String
Private oDocEntry01 As String

Private oOrdGbn As String
Private oSequence As String
Private oDocdate As String
Private oFormMode01 As SAPbouiCOM.BoFormMode
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    Dim MainJob As String
    MainJob = MDC_PS_Common.User_MainJob
    
    If MainJob = "생산팀서무" Then '생산팀서무는 작업일보(작지)의 공정정보 매트릭스 컬럼 세팅을 FIX 시킴(전용화면 사용) (2016.03.16 송명규, 강주란 요청)
        oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP040_01.srf")
    Else
        oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP040.srf")
    End If
    
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
    
    oFormUniqueID01 = "PS_PP040_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
    
    oForm01.Freeze True
    
    Call PS_PP040_CreateItems
    Call PS_PP040_ComboBox_Setting
    Call PS_PP040_CF_ChooseFromList
    Call PS_PP040_EnableMenus
    Call PS_PP040_SetDocument(oFromDocEntry01)
    Call PS_PP040_FormResize
    
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
        
' 작업시간 합계 추가 S
'            Dim i&
'            Dim Total As Currency
'
'
'                For i = 0 To oMat01.VisualRowCount - 1
'
'                    Total = Total + Val(oMat01.Columns("WorkTime").Cells(i + 1).Specific.VALUE)
''                 oMat01.Columns("Total").Cells.Specific.VALUE = Total
'                Next i
'                oForm01.Items("Total").Specific.VALUE = Total
            Call PS_PP040_SumWorkTime
' 작업시간 합계 추가 E

'            Call Raise_EVENT_MATRIX_LOAD(FormUID, pval, BubbleEvent)
        
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
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
                If oForm01.Mode = fm_OK_MODE Then
                    If (PS_PP040_Validate("취소") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If Sbo_Application.MessageBox("정말로 취소하시겠습니까?", "1", "예", "아니오") <> "1" Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                Else
                    MDC_Com.MDC_GF_Message "현재 모드에서는 취소할수 없습니다.", "W"
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call Raise_EVENT_RECORD_MOVE(FormUID, pval, BubbleEvent)
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_PP040_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_PP040_FormItemEnabled '//UDO방식
                Call PS_PP040_AddMatrixRow01(0, True) '//UDO방식
                Call PS_PP040_AddMatrixRow02(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call Raise_EVENT_RECORD_MOVE(FormUID, pval, BubbleEvent)
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
                If (oForm01.Mode <> fm_ADD_MODE) Then
                    If (PS_PP040_FindValidateDocument("@PS_PP040H") = False) Then
                        If Sbo_Application.Menus.Item("1281").Enabled = True Then '//찾기메뉴 활성화일때 수행
                            Sbo_Application.ActivateMenuItem ("1281")
                        Else
                            Sbo_Application.SetStatusBarMessage "관리자에게 문의바랍니다", bmt_Short, True
                        End If
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
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
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
'            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
'            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
'            MenuCreationParams01.uniqueID = "MenuUID"
'            MenuCreationParams01.String = "메뉴명"
'            MenuCreationParams01.Enabled = True
'            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
'        End If
    ElseIf pval.BeforeAction = False Then
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'                Call Sbo_Application.Menus.RemoveEx("MenuUID")
'        End If
    End If
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Or pval.ItemUID = "Mat03" Then
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
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oMat01Row01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat02" Then
        If pval.Row > 0 Then
            oMat02Row02 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat03" Then
        If pval.Row > 0 Then
            oMat03Row03 = pval.Row
        End If
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    Dim vReturnValue$
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("처리 중..", 100, False)
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "PS_PP040" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_PP040_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oOrdGbn = Trim(oForm01.Items("OrdGbn").Specific.VALUE) '//작업구분
                oSequence = oMat01.Columns("Sequence").Cells(1).Specific.VALUE
                oDocdate = Trim(oForm01.Items("DocDate").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_PP040_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "2" Then '//취소버튼 누를시 저장할 자료가 있으면 메시지 표시
            If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                If oMat01.VisualRowCount > 1 Then
                    vReturnValue = Sbo_Application.MessageBox("저장하지 않는 자료가 있습니다. 취소하시겠습니까?", 2, "&확인", "&취소")
                    Select Case vReturnValue
                        Case 1:
                        Case 2:
                                BubbleEvent = False
                                Exit Sub
                    End Select
                End If
            End If
        End If
        
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_PP040_OrderInfoLoad
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                Call PS_PP040_OrderInfoLoad
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button02" Then
            Dim DocEntry$
            Dim LineNum$
            Dim i&
            Dim ErrNum&
            Dim DocNum                  As String
            Dim WinTitle                As String
            Dim ReportName              As String
            Dim oText(1) As String
            
            Dim sQry                    As String
            Dim sQryS                   As String
            Dim sQry1$
            Dim WorkName$
            Dim oRecordSet01            As SAPbobsCOM.Recordset
            Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
            Call ConnectODBC
            
            DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
            For i = 0 To oMat01.VisualRowCount - 1
                If oMat01.IsRowSelected(i + 1) = True Then
                    LineNum = oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE
                End If
            Next
            
            WinTitle = " 공정카드 [PS_PP040]"
            ReportName = "PS_PP040_01.rpt"
               
            sQry1 = "Select U_WorkName From [@PS_PP040M] Where DocEntry = '" & DocEntry & "' And IsNull(U_WorkName, '') <> ''"
            oRecordSet01.DoQuery sQry1
            
            Do Until oRecordSet01.EOF
                WorkName = WorkName + "     " + oRecordSet01.Fields(0).VALUE
                oRecordSet01.MoveNext
            Loop
            
            '//Formula 수식필드
            ReDim gRpt_Formula(1)
            ReDim gRpt_Formula_Value(1)
            
            oText(1) = WorkName
            
            For i = 1 To 1
                If Len("" & i & "") = 1 Then
                    gRpt_Formula(i) = "F0" & i & ""
                Else
                    gRpt_Formula(i) = "F" & i & ""
                End If
                gRpt_Formula_Value(i) = oText(i)
            Next i
            
            '//SubReport
            ReDim gRpt_SRptSqry(1)
            ReDim gRpt_SRptName(1)
            ReDim gRpt_SFormula(1, 1)
            ReDim gRpt_SFormula_Value(1, 1)
            
            gRpt_SFormula(1, 1) = ""
            gRpt_SFormula_Value(1, 1) = ""
            
            sQryS = "EXEC [PS_PP040_06] '" & DocEntry & "', '" & LineNum & "', 'S'"
            
            gRpt_SRptSqry(1) = sQryS
            gRpt_SRptName(1) = "PS_PP040_S1"
            
            '//조회조건문
            sQry = "EXEC [PS_PP040_06] '" & DocEntry & "', '" & LineNum & "', 'M'"
            oRecordSet01.DoQuery sQry
            If oRecordSet01.RecordCount = 0 Then
                MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
                Exit Sub
            End If
            
            '//CR Action
            If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "N", "V") = False Then
                Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
            End If
            Set oRecordSet01 = Nothing
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP040" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    If oOrdGbn = "101" And oSequence = "1" Then
                        oForm01.Mode = fm_FIND_MODE
                        Call PS_PP040_FormItemEnabled
                        oForm01.Items("DocEntry").Specific.VALUE = oDocEntry01
                        oForm01.Items("1").Click ct_Regular
                    Else
                        Call PS_PP040_FormItemEnabled
                        Call PS_PP040_AddMatrixRow01(0, True) '//UDO방식일때
                        Call PS_PP040_AddMatrixRow02(0, True) '//UDO방식일때
                    End If
'
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    If (oFormMode01 = fm_UPDATE_MODE) Then
                        oFormMode01 = fm_OK_MODE
                        oForm01.Mode = fm_FIND_MODE
                        Call PS_PP040_FormItemEnabled
                        oForm01.Items("DocEntry").Specific.VALUE = oDocEntry01
                        oForm01.Items("1").Click ct_Regular
                    End If
                    Call PS_PP040_FormItemEnabled
                End If
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Exit Sub
    
Raise_EVENT_ITEM_PRESSED_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "OrdMgNum" Then
            If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "60" Then '//작업타입이 일반,조정일때
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "OrdMgNum", "") '//사용자값활성
            End If
        End If
        
        
        
        
        
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "OrdMgNum" Then
                If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "60" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "70" Then '//일반,조정, 설계
                    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then
                        Call MDC_Com.MDC_GF_Message("작업구분이 선택되지 않았습니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Items("BPLId").Specific.Selected.VALUE = "선택" Then
                        Call MDC_Com.MDC_GF_Message("사업장이 선택되지 않았습니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Items("ItemCode").Specific.VALUE = "" Then
                        Call MDC_Com.MDC_GF_Message("품목코드가 선택되지 않았습니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Items("OrdNum").Specific.VALUE = "" Then
                        Call MDC_Com.MDC_GF_Message("작지번호가 선택되지 않았습니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Items("PP030HNo").Specific.VALUE = "" Then
                        Call MDC_Com.MDC_GF_Message("작지문서번호가 선택되지 않았습니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "OrdMgNum") '//사용자값활성
                    End If
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//지원
                    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then
                        Call MDC_Com.MDC_GF_Message("작업구분이 선택되지 않았습니다.", "W")
                        oForm01.Items("OrdGbn").Click ct_Regular
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Items("BPLId").Specific.Selected.VALUE = "선택" Then
                        Call MDC_Com.MDC_GF_Message("사업장이 선택되지 않았습니다.", "W")
                        oForm01.Items("BPLId").Click ct_Regular
                        BubbleEvent = False
                        Exit Sub
'                    ElseIf oForm01.Items("ItemCode").Specific.Value = "" Then
'                        Call MDC_Com.MDC_GF_Message("품목코드가 선택되지 않았습니다.", "W")
'                        oForm01.Items("ItemCode").Click ct_Regular
'                        BubbleEvent = False
'                        Exit Sub
                    ElseIf oForm01.Items("OrdNum").Specific.VALUE = "" Then
                        Call MDC_Com.MDC_GF_Message("작지번호가 선택되지 않았습니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "OrdMgNum") '//사용자값활성
                    End If
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//외주
                
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//실적
                
                End If
            
            End If
        End If
        If pval.ItemUID = "Mat02" Then
            If pval.ColUID = "WorkCode" Then
                If Val(oForm01.Items("BaseTime").Specific.VALUE) = 0 Then
                    Call MDC_Com.MDC_GF_Message("기준시간을 입력하지 않았습니다.", "W")
                    oForm01.Items("BaseTime").Click ct_Regular
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        End If
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat02", "WorkCode") '사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat02", "NCode") '사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat03", "FailCode") '사용자값활성
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "MachCode") '설비코드 사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "SubLot") 'sub작지번호 사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "CItemCod") '원재료코드 사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "SCpCode") '지원공정추가(2018.05.30 송명규)
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "UseMCode", "") '작업장비 사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '사용자값활성
    ElseIf pval.BeforeAction = False Then
        '// 화살표 이동 강제 코딩 - 류영조
        If pval.ItemUID = "Mat01" Then
            If pval.CharPressed = 38 Then '//위쪽 화살표
                If pval.Row > 1 And pval.Row <= oMat01.VisualRowCount Then
                    oForm01.Freeze True
                    oMat01.Columns(pval.ColUID).Cells(pval.Row - 1).Click ct_Regular
                    oForm01.Freeze False
                End If
            ElseIf pval.CharPressed = 40 Then '//아래 화살표
                If pval.Row > 0 And pval.Row < oMat01.VisualRowCount Then
                    oForm01.Freeze True
                    oMat01.Columns(pval.ColUID).Cells(pval.Row + 1).Click ct_Regular
                    oForm01.Freeze False
                End If
            End If
            
            '작업시간 입력 시마다 합계 계산(2011.09.26 송명규 추가)
            If pval.ColUID = "WorkTime" And pval.Row <> "0" Then
            
                Call PS_PP040_SumWorkTime
                
            End If
            
         ElseIf pval.ItemUID = "BaseTime" Then
            
            If pval.CharPressed = 9 Then '탭 키 Press
            
                oMat02.Columns("WorkCode").Cells(1).Click
            
            End If
            
        End If
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemChanged = True Then
            Call oForm01.Freeze(True)
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "특정컬럼") Then
                    '//기타작업
                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP040L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        'PS_PP040_AddMatrixRow (pval.Row)
                    End If
                Else
                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                End If
            ElseIf (pval.ItemUID = "Mat02") Then
                If (pval.ColUID = "특정컬럼") Then
                    '//기타작업
                    Call oDS_PS_PP040M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                    If oMat02.RowCount = pval.Row And Trim(oDS_PS_PP040M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        'PS_PP040_AddMatrixRow (pval.Row)
                    End If
                Else
                    Call oDS_PS_PP040M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                End If
            ElseIf (pval.ItemUID = "Mat03") Then
                If (pval.ColUID = "특정컬럼") Then
                Else
                    Call oDS_PS_PP040N.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                End If
            Else
                If (pval.ItemUID = "OrdType") Then
                    Call oDS_PS_PP040H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    If oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "10" Or oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "50" Or oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "60" Or oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "70" Then '//일반,조정,설계
                        '창원은 품목구분 선택하도록 수정 '2015/04/09
                        If oForm01.Items("BPLId").Specific.VALUE = "1" Then
                            oForm01.Items("OrdGbn").Enabled = True
                        Else
                            oForm01.Items("OrdGbn").Enabled = False
                        End If
                        oForm01.Items("BPLId").Enabled = False
                        oForm01.Items("ItemCode").Enabled = False
                    ElseIf oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "20" Then
                        oForm01.Items("OrdGbn").Enabled = True
                        oForm01.Items("BPLId").Enabled = True
                        oForm01.Items("ItemCode").Enabled = True
                    ElseIf oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "30" Then
                        oForm01.Items("OrdGbn").Enabled = False
                        oForm01.Items("BPLId").Enabled = False
                        oForm01.Items("ItemCode").Enabled = False
                    ElseIf oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "40" Then
                        oForm01.Items("OrdGbn").Enabled = False
                        oForm01.Items("BPLId").Enabled = False
                        oForm01.Items("ItemCode").Enabled = False
                    End If
                                        
                    Call oForm01.Items("OrdGbn").Specific.Select(0, psk_Index)
                    oForm01.Items("OrdMgNum").Specific.VALUE = ""
                    'Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
                    oForm01.Items("ItemCode").Specific.VALUE = ""
                    oForm01.Items("ItemName").Specific.VALUE = ""
                    oForm01.Items("OrdNum").Specific.VALUE = ""
                    oForm01.Items("OrdSub1").Specific.VALUE = ""
                    oForm01.Items("OrdSub2").Specific.VALUE = ""
                    oForm01.Items("PP030HNo").Specific.VALUE = ""
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    Call PS_PP040_AddMatrixRow01(0, True)
                    oMat02.Clear
                    oMat02.FlushToDataSource
                    oMat02.LoadFromDataSource
                    Call PS_PP040_AddMatrixRow02(0, True)
                    oMat03.Clear
                    oMat03.FlushToDataSource
                    oMat03.LoadFromDataSource
                ElseIf (pval.ItemUID = "OrdGbn") Then
                    Call oDS_PS_PP040H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    Call PS_PP040_AddMatrixRow01(0, True)
                    oMat02.Clear
                    oMat02.FlushToDataSource
                    oMat02.LoadFromDataSource
                    Call PS_PP040_AddMatrixRow02(0, True)
                    oMat03.Clear
                    oMat03.FlushToDataSource
                    oMat03.LoadFromDataSource
                ElseIf (pval.ItemUID = "BPLId") Then
                    Call oDS_PS_PP040H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    Call PS_PP040_AddMatrixRow01(0, True)
                    oMat02.Clear
                    oMat02.FlushToDataSource
                    oMat02.LoadFromDataSource
                    Call PS_PP040_AddMatrixRow02(0, True)
                    oMat03.Clear
                    oMat03.FlushToDataSource
                    oMat03.LoadFromDataSource
                Else
                    If pval.ItemUID <> "CardType" Then '거래처구분이 아닐 경우만 실행(2012.02.02 송명규 추가)
                        Call oDS_PS_PP040H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    End If
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oMat02.LoadFromDataSource
            oMat02.AutoResizeColumns
            oMat03.LoadFromDataSource
            oMat03.AutoResizeColumns
            oForm01.Update
            If pval.ItemUID = "Mat01" Then
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Collapsed
            ElseIf pval.ItemUID = "Mat02" Then
                oMat02.Columns(pval.ColUID).Cells(pval.Row).Click ct_Collapsed
            ElseIf pval.ItemUID = "Mat03" Then
                oMat03.Columns(pval.ColUID).Cells(pval.Row).Click ct_Collapsed
            Else
                
            End If
            Call oForm01.Freeze(False)
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error
    Dim TempForm01 As Variant
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Opt01" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat02"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Opt02" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat03"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Opt03" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
'        If pval.ItemUID = "Mat01" Then
'            If pval.Row > 0 Then
'                Call oMat01.SelectRow(pval.Row, True, False)
'            End If
'        End If
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)
                oMat01Row01 = pval.Row
            End If
        End If
        If pval.ItemUID = "Mat02" Then
            If pval.Row > 0 Then
                Call oMat02.SelectRow(pval.Row, True, False)
                oMat02Row02 = pval.Row
            End If
        End If
        If pval.ItemUID = "Mat03" Then
            If pval.Row > 0 Then
                Call oMat03.SelectRow(pval.Row, True, False)
                oMat03Row03 = pval.Row
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        '// 작업지시번호 링크 번튼 - 류영조
        If pval.ItemUID = "LBtn01" Then
            Set TempForm01 = New PS_PP030
            Call TempForm01.LoadForm(oForm01.Items("PP030HNo").Specific.VALUE)
            Set TempForm01 = Nothing
        End If
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "60" Then '//작업타입이 일반,조정인경우
                    If oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE = "" Then
                    
                    Else
                        If oMat03.VisualRowCount = 0 Then
                            Call PS_PP040_AddMatrixRow03(0, True)
                        Else
                            Call PS_PP040_AddMatrixRow03(oMat03.VisualRowCount)
                        End If
                        Call oDS_PS_PP040N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP040N.setValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns("CpCode").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP040N.setValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns("CpName").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP040N.setValue("U_OLineNum", oMat03.VisualRowCount - 1, pval.Row)
                        oMat03.LoadFromDataSource
                        oMat03.AutoResizeColumns
'                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
'                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
                        oMat03.Columns("OLineNum").TitleObject.Sortable = True
                        Call oMat03.Columns("OLineNum").TitleObject.Sort(gst_Ascending)
                        oMat03.FlushToDataSource
                    End If
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT지원인경우
                    If oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE = "" Then
                    
                    Else
                        If oMat03.VisualRowCount = 0 Then
                            Call PS_PP040_AddMatrixRow03(0, True)
                        Else
                            Call PS_PP040_AddMatrixRow03(oMat03.VisualRowCount)
                        End If
                        Call oDS_PS_PP040N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP040N.setValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns("CpCode").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP040N.setValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns("CpName").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP040N.setValue("U_OLineNum", oMat03.VisualRowCount - 1, pval.Row)
                        oMat03.LoadFromDataSource
                        oMat03.AutoResizeColumns
'                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
'                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
                        oMat03.Columns("OLineNum").TitleObject.Sortable = True
                        Call oMat03.Columns("OLineNum").TitleObject.Sort(gst_Ascending)
                        oMat03.FlushToDataSource
                    End If
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
                End If
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
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "OrdMgNum" Then
                Set oTempClass = New PS_PP030
                Call oTempClass.LoadForm(Mid(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, 1, InStr(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, "-") - 1))
            End If
            If pval.ColUID = "PP030HNo" Then
                Set oTempClass = New PS_PP030
                Call oTempClass.LoadForm(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
            End If
        End If
        If pval.ItemUID = "Mat03" Then
            If pval.ColUID = "OrdMgNum" Then
                Set oTempClass = New PS_PP030
                Call oTempClass.LoadForm(Mid(oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, 1, InStr(oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, "-") - 1))
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim Weight As Double
    
    Dim Time As Double
    Dim Hour As Long
    Dim Minute As Long
                        
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (PS_PP040_Validate("수정01") = False) Then
                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Trim(oDS_PS_PP040L.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                Else
                    If (pval.ColUID = "OrdMgNum") Then
                        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                        
                        Dim ProgBar01 As SAPbouiCOM.ProgressBar
                        Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("실행 중...", 100, False)

                        If oForm01.Items("OrdNum").Specific.VALUE = "" Then '//작지번호에 값이 없으면 작업지시가 불러오기전
                            Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                        Else '//작업지시가 선택된상태
                            If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "60" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "70" Then '//작업타입이 일반,조정, 설계
                                '//작지문서헤더번호가 일치하지 않으면
                                If oForm01.Items("PP030HNo").Specific.VALUE <> Mid(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, 1, InStr(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, "-") - 1) Then
                                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                                Else '//작지문서번호가 일치하면
                                    If oForm01.Items("BPLId").Specific.Selected.VALUE <> "1" Then
                                        '//신동사업부를 제외한 사업부만 체크
                                        For i = 1 To oMat01.RowCount
                                            '//현재 입력한 값이 이미 입력되어 있는경우
                                            If oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE And i <> pval.Row Then
                                                Call MDC_Com.MDC_GF_Message("이미 입력한 공정입니다.", "W")
                                                Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                                                GoTo Continue
                                            End If
    '                                        '//공정라인의 공정순서가 앞공정보다 높으면
    '                                        If Val(oMat01.Columns("Sequence").Cells(i).Specific.Value) >= MDC_PS_Common.GetValue("SELECT PS_PP030M.U_Sequence FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE CONVERT(NVARCHAR,PS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" & oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.Value & "'") Then
    '                                            Call MDC_Com.MDC_GF_Message("공정순서가 올바르지 않습니다.", "W")
    '                                            Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
    '                                            GoTo Continue
    '                                        End If
                                        Next
                                        
                                        '생산완료등록이 완료된 작번인지 체크_수량으로 비교(2012.08.27 송명규 추가)_S
                                        Dim WkCmDt As String
                                        Query01 = "EXEC PS_PP040_90 '" & Mid(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, 1, InStr(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, "-") - 1) & "'" 'oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE & "'"
                                        Call RecordSet01.DoQuery(Query01)
                                        WkCmDt = RecordSet01.Fields("WkCmDt").VALUE
                                        
                                        If RecordSet01.Fields("Return").VALUE = "1" Then '생산완료수량이 작업지시수량만큼 모두 등록이 되었다면
                                            If Sbo_Application.MessageBox("생산완료가 모두 등록된 작번(완료일자:" & WkCmDt & ")입니다. 계속 진행하시겠습니까?", "1", "예", "아니오") = "1" Then
                                                '계속 진행시에는 해당 작업지시문서번호 등록
                                            Else
'                                                Call MDC_Com.MDC_GF_Message("등록이 취소되었습니다.", "W")
                                                Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                                                GoTo Continue
                                            End If
                                        End If
                                        '생산완료등록이 완료된 작번인지 체크_수량으로 비교(2012.08.27 송명규 추가)_E
                                        
                                        '판매완료등록 체크_S(2015.07.14 송명규 추가)
                                        Dim OINV_Dt As String
                                        Query01 = "EXEC PS_PP040_91 '"
                                        Query01 = Query01 & Mid(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, 1, InStr(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, "-") - 1) & "','"
                                        Query01 = Query01 & oDS_PS_PP040H.GetValue("U_DocDate", 0) & "'"
                                        Call RecordSet01.DoQuery(Query01)
                                        OINV_Dt = RecordSet01.Fields("OINV_Dt").VALUE
                                        
                                        If RecordSet01.Fields("Return").VALUE = "1" Then '판매확정수량이 판매오더수량만큼 모두 등록이 되었다면
                                            Call Sbo_Application.MessageBox("판매완료(최종일자:" & OINV_Dt & ")된 작번입니다. 등록이 불가능합니다.", 1, "확인")
                                            Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                                            GoTo Continue
                                        End If
                                        '판매완료등록 체크_E(2015.07.14 송명규 추가)
                                        
                                                                                
                                    End If
                                    
                                    Query01 = "EXEC PS_PP040_01 '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "', '" & oForm01.Items("OrdType").Specific.Selected.VALUE & "'"
                                    RecordSet01.DoQuery Query01
                                    If RecordSet01.RecordCount = 0 Then
                                        Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                                    Else
                                        Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, RecordSet01.Fields("OrdMgNum").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_Sequence", pval.Row - 1, RecordSet01.Fields("Sequence").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_CpCode", pval.Row - 1, RecordSet01.Fields("CpCode").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_CpName", pval.Row - 1, RecordSet01.Fields("CpName").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_OrdGbn", pval.Row - 1, RecordSet01.Fields("OrdGbn").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_BPLId", pval.Row - 1, RecordSet01.Fields("BPLId").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_OrdNum", pval.Row - 1, RecordSet01.Fields("OrdNum").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_OrdSub1", pval.Row - 1, RecordSet01.Fields("OrdSub1").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_OrdSub2", pval.Row - 1, RecordSet01.Fields("OrdSub2").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_PP030HNo", pval.Row - 1, RecordSet01.Fields("PP030HNo").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_PP030MNo", pval.Row - 1, RecordSet01.Fields("PP030MNo").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_SelWt", pval.Row - 1, RecordSet01.Fields("SelWt").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_PSum", pval.Row - 1, RecordSet01.Fields("PSum").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_BQty", pval.Row - 1, RecordSet01.Fields("BQty").VALUE)
                                        Call oDS_PS_PP040L.setValue("U_PQty", pval.Row - 1, 0)
                                        Call oDS_PS_PP040L.setValue("U_PWeight", pval.Row - 1, 0)
                                        Call oDS_PS_PP040L.setValue("U_YQty", pval.Row - 1, 0)
                                        Call oDS_PS_PP040L.setValue("U_YWeight", pval.Row - 1, 0)
                                        Call oDS_PS_PP040L.setValue("U_NQty", pval.Row - 1, 0)
                                        Call oDS_PS_PP040L.setValue("U_NWeight", pval.Row - 1, 0)
                                        Call oDS_PS_PP040L.setValue("U_ScrapWt", pval.Row - 1, 0)
                                        Call oDS_PS_PP040L.setValue("U_WorkTime", pval.Row - 1, 0)
                                        Call oDS_PS_PP040L.setValue("U_LineId", pval.Row - 1, "")
                                        
                                        '//설비코드,명 Reset
                                        Call oDS_PS_PP040L.setValue("U_MachCode", pval.Row - 1, "")
                                        Call oDS_PS_PP040L.setValue("U_MachName", pval.Row - 1, "")
                                        '//불량코드테이블
                                        If oMat03.VisualRowCount = 0 Then
                                            Call PS_PP040_AddMatrixRow03(0, True)
                                        Else
                                            Call PS_PP040_AddMatrixRow03(oMat03.VisualRowCount)
                                        End If
                                        
                                        Call oDS_PS_PP040N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields("OrdMgNum").VALUE)
                                        Call oDS_PS_PP040N.setValue("U_CpCode", oMat03.VisualRowCount - 1, RecordSet01.Fields("CpCode").VALUE)
                                        Call oDS_PS_PP040N.setValue("U_CpName", oMat03.VisualRowCount - 1, RecordSet01.Fields("CpName").VALUE)
                                        Call oDS_PS_PP040N.setValue("U_OLineNum", oMat03.VisualRowCount - 1, pval.Row)
                                        
                                        
                                        
                                        '// 류영조
                                        If oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "60" Then
                                            Call oDS_PS_PP040H.setValue("U_BaseTime", 0, "1")
                                            oMat02.Columns("WorkCode").Cells(1).Specific.VALUE = "9999999"
'                                            oMat02.Columns("WorkName").Cells(1).Specific.Value = "조정"
'                                            Call oDS_PS_PP040M.setValue("U_WorkCode", 0, "9999999")
                                            Call oDS_PS_PP040M.setValue("U_WorkName", 0, "조정")
                                            oMat02.LoadFromDataSource
                                        Else
'                                            Call oDS_PS_PP040H.setValue("U_BaseTime", 0, "")
'                                            oMat02.Columns("WorkCode").Cells(1).Specific.Value = ""
'                                            oMat02.Columns("WorkName").Cells(1).Specific.Value = ""
                    '                        Call oDS_PS_PP040M.setValue("U_WorkCode", 0, "")
                    '                        Call oDS_PS_PP040M.setValue("U_WorkName", 0, "")
                                        End If
                                    End If
                                End If
                            ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT지원
                                '//올바른 공정코드인지 검사
                                If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP001L] WHERE U_CpCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'") = 0 Then
                                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                                Else
                                    For i = 1 To oMat01.RowCount
                                        '//현재 입력한 값이 이미 입력되어 있는경우
                                        If oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE And i <> pval.Row Then
                                            Call MDC_Com.MDC_GF_Message("이미 입력한 공정입니다.", "W")
                                            Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                                            GoTo Continue
                                        End If
                                    Next
                                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                                    Call oDS_PS_PP040L.setValue("U_CpCode", pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                                    Call oDS_PS_PP040L.setValue("U_CpName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                                    Call oDS_PS_PP040L.setValue("U_OrdGbn", pval.Row - 1, oForm01.Items("OrdGbn").Specific.Selected.VALUE)
                                    Call oDS_PS_PP040L.setValue("U_BPLId", pval.Row - 1, oForm01.Items("BPLId").Specific.Selected.VALUE)
                                    Call oDS_PS_PP040L.setValue("U_ItemCode", pval.Row - 1, "")
                                    Call oDS_PS_PP040L.setValue("U_ItemName", pval.Row - 1, "")
                                    '//PSMT지원은 품목코드 필요없음
'                                    Call oDS_PS_PP040L.setValue("U_ItemCode", pval.Row - 1, oForm01.Items("ItemCode").Specific.Value)
'                                    Call oDS_PS_PP040L.setValue("U_ItemName", pval.Row - 1, oForm01.Items("ItemName").Specific.Value)
                                    Call oDS_PS_PP040L.setValue("U_OrdNum", pval.Row - 1, oForm01.Items("OrdNum").Specific.VALUE)
                                    Call oDS_PS_PP040L.setValue("U_OrdSub1", pval.Row - 1, oForm01.Items("OrdSub1").Specific.VALUE)
                                    Call oDS_PS_PP040L.setValue("U_OrdSub2", pval.Row - 1, oForm01.Items("OrdSub2").Specific.VALUE)
                                    Call oDS_PS_PP040L.setValue("U_PP030HNo", pval.Row - 1, "")
                                    Call oDS_PS_PP040L.setValue("U_PP030MNo", pval.Row - 1, "")
                                    Call oDS_PS_PP040L.setValue("U_PSum", pval.Row - 1, 0)
                                    Call oDS_PS_PP040L.setValue("U_PQty", pval.Row - 1, 0)
                                    Call oDS_PS_PP040L.setValue("U_PWeight", pval.Row - 1, 0)
                                    Call oDS_PS_PP040L.setValue("U_YQty", pval.Row - 1, 0)
                                    Call oDS_PS_PP040L.setValue("U_YWeight", pval.Row - 1, 0)
                                    Call oDS_PS_PP040L.setValue("U_NQty", pval.Row - 1, 0)
                                    Call oDS_PS_PP040L.setValue("U_NWeight", pval.Row - 1, 0)
                                    Call oDS_PS_PP040L.setValue("U_ScrapWt", pval.Row - 1, 0)
                                    '//불량코드테이블
                                    If oMat03.VisualRowCount = 0 Then
                                        Call PS_PP040_AddMatrixRow03(0, True)
                                    Else
                                        If oDS_PS_PP040L.GetValue("U_OrdMgNum", pval.Row - 1) = oDS_PS_PP040N.GetValue("U_OrdMgNum", oMat03.VisualRowCount - 1) Then
                                        Else
                                            Call PS_PP040_AddMatrixRow03(oMat03.VisualRowCount)
                                        End If
                                    End If
                                    Call oDS_PS_PP040N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                                    Call oDS_PS_PP040N.setValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                                    Call oDS_PS_PP040N.setValue("U_CpName", oMat03.VisualRowCount - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                                End If
                            ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주
                            
                            ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적
                            
                            End If
Continue:
                            If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP040L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                                Call PS_PP040_AddMatrixRow01(pval.Row)
                            End If
                        End If
                        
                        ProgBar01.VALUE = 100
                        ProgBar01.Stop
                        Set ProgBar01 = Nothing
                        
                        Set RecordSet01 = Nothing
                    ElseIf pval.ColUID = "PQty" Then
                        If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0 Then
                            If Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "50" Or Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "60" Then
                                GoTo Skip_PQty
                            Else
                                Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP040L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                            End If
                        Else
Skip_PQty:
                            Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Call oDS_PS_PP040L.setValue("U_YQty", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Weight = Val(MDC_PS_Common.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" & oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE & "' AND U_CpCode = '" & oMat01.Columns("CpCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1)) / 1000
                            If Weight = 0 Then
                                Call oDS_PS_PP040L.setValue("U_PWeight", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                Call oDS_PS_PP040L.setValue("U_YWeight", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Else
                                Call oDS_PS_PP040L.setValue("U_PWeight", pval.Row - 1, Weight * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                Call oDS_PS_PP040L.setValue("U_YWeight", pval.Row - 1, Weight * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            End If
                            Call oDS_PS_PP040L.setValue("U_NQty", pval.Row - 1, 0)
                            Call oDS_PS_PP040L.setValue("U_NWeight", pval.Row - 1, 0)
                        End If
                    ElseIf pval.ColUID = "NQty" Then
                        If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0 Then
                            If Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "50" Or Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "60" Then
                                GoTo skip_Nqty
                            Else
                                Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP040L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                            End If
                        ElseIf Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) > Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) Then
                            If Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "50" Or Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "60" Then
                                GoTo skip_Nqty
                            Else
                                Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP040L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                            End If
                        Else
skip_Nqty:
                            Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Call oDS_PS_PP040L.setValue("U_YQty", pval.Row - 1, Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Weight = Val(MDC_PS_Common.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" & oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE & "' AND U_CpCode = '" & oMat01.Columns("CpCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1)) / 1000
                            If Weight = 0 Then
                                Call oDS_PS_PP040L.setValue("U_NWeight", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                Call oDS_PS_PP040L.setValue("U_YWeight", pval.Row - 1, (Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)))
                            Else
                                Call oDS_PS_PP040L.setValue("U_NWeight", pval.Row - 1, Weight * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                Call oDS_PS_PP040L.setValue("U_YWeight", pval.Row - 1, Weight * (Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)))
                            End If
                        End If
                        
                    ElseIf pval.ColUID = "WorkTime" Then '작업시간(공수)을 입력할 때
                        
                        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                        Dim ReturnValue As String
                        
                        If oForm01.Items("BPLId").Specific.Selected.VALUE <> "1" Then
                        
                            '적자 여부 확인 체크(2016.05.20 송명규 추가)_S
'                            Query01 = "EXEC PS_PP040_92 '"
'                            Query01 = Query01 & oMat01.Columns("PP030HNo").Cells(pval.Row).Specific.VALUE & "','"
'                            Query01 = Query01 & oMat01.Columns("PP030MNo").Cells(pval.Row).Specific.VALUE & "'"
'
'                            Call RecordSet01.DoQuery(Query01)
'
'                            ReturnValue = RecordSet01.Fields("ReturnValue").VALUE
'
'                            If ReturnValue <> "" Then '적자가 예상되는 작번은 메시지 출력
'                                If Sbo_Application.MessageBox(ReturnValue, "1", "예", "아니오") = "1" Then
'                                    '계속 진행시에는 해당 공수 등록
'                                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
'                                Else
'                                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
'    '                                Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
'    '                                GoTo Continue
'                                End If
'
'                            Else

                                Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))

'                            End If
                            '적자 여부 확인 체크(2016.05.20 송명규 추가)_S
                            
                        Else
                        
                            Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                        
                        End If
                        
                        Set RecordSet01 = Nothing
                        
                    ElseIf pval.ColUID = "BdwQty" Then '//기존도면매수
                        Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP040L.setValue("U_AdwQTy", pval.Row - 1, (Val(oMat01.Columns("DwRate").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100)
                        Call oDS_PS_PP040L.setValue("U_PQTy", pval.Row - 1, ((Val(oMat01.Columns("DwRate").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100) + Val(oMat01.Columns("NdwQTy").Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP040L.setValue("U_PWeight", pval.Row - 1, ((Val(oMat01.Columns("DwRate").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100) + Val(oMat01.Columns("NdwQTy").Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP040L.setValue("U_YQTy", pval.Row - 1, ((Val(oMat01.Columns("DwRate").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100) + Val(oMat01.Columns("NdwQTy").Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP040L.setValue("U_YWeight", pval.Row - 1, ((Val(oMat01.Columns("DwRate").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100) + Val(oMat01.Columns("NdwQTy").Cells(pval.Row).Specific.VALUE))
                    ElseIf pval.ColUID = "DwRate" Then '//도면 적용율
                        Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP040L.setValue("U_AdwQTy", pval.Row - 1, (Val(oMat01.Columns("BdwQty").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100)
                        Call oDS_PS_PP040L.setValue("U_PQTy", pval.Row - 1, ((Val(oMat01.Columns("BdwQty").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100) + Val(oMat01.Columns("NdwQTy").Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP040L.setValue("U_PWeight", pval.Row - 1, ((Val(oMat01.Columns("BdwQty").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100) + Val(oMat01.Columns("NdwQTy").Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP040L.setValue("U_YQTy", pval.Row - 1, ((Val(oMat01.Columns("BdwQty").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100) + Val(oMat01.Columns("NdwQTy").Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP040L.setValue("U_YWeight", pval.Row - 1, ((Val(oMat01.Columns("BdwQty").Cells(pval.Row).Specific.VALUE) * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)) / 100) + Val(oMat01.Columns("NdwQTy").Cells(pval.Row).Specific.VALUE))
                    ElseIf pval.ColUID = "NdwQTy" Then '//신규도면매수
                        Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP040L.setValue("U_PQTy", pval.Row - 1, (Val(oMat01.Columns("AdwQty").Cells(pval.Row).Specific.VALUE) + Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)))
                        Call oDS_PS_PP040L.setValue("U_PWeight", pval.Row - 1, (Val(oMat01.Columns("AdwQty").Cells(pval.Row).Specific.VALUE) + Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)))
                        Call oDS_PS_PP040L.setValue("U_YQTy", pval.Row - 1, (Val(oMat01.Columns("AdwQty").Cells(pval.Row).Specific.VALUE) + Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)))
                        Call oDS_PS_PP040L.setValue("U_YWeight", pval.Row - 1, (Val(oMat01.Columns("AdwQty").Cells(pval.Row).Specific.VALUE) + Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)))
                    ElseIf pval.ColUID = "MachCode" Then
                        Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP040L.setValue("U_MachName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                        'Call oDS_PS_PP040L.setValue("U_MachGrCd", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_MacdGrCd FROM [@PS_PP130H] WHERE U_MachCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                    ElseIf pval.ColUID = "CItemCod" Then
                        Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP040L.setValue("U_CItemNam", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_ItemNam2 FROM [@PS_PP005H] WHERE U_ItemCod1 = '" & oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE & "' and U_ItemCod2 = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                    ElseIf pval.ColUID = "SCpCode" Then '지원공정코드
                        Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP040L.setValue("U_SCpName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" & oMat01.Columns("SCpCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                    Else
                        Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    End If
                End If
            ElseIf (pval.ItemUID = "Mat02") Then
                If (pval.ColUID = "WorkCode") Then
                    '//기타작업
                    Call oDS_PS_PP040M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_PP040M.setValue("U_WorkName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                    If oMat02.RowCount = pval.Row And Trim(oDS_PS_PP040M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        Call PS_PP040_AddMatrixRow02(pval.Row)
                    End If
                ElseIf pval.ColUID = "NStart" Then
                    Call oDS_PS_PP040M.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    If Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE) = 0 Or Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) = 0 Then
                        Call oDS_PS_PP040M.setValue("U_NTime", pval.Row - 1, 0)
                        Call oDS_PS_PP040M.setValue("U_YTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE))
                        Call oDS_PS_PP040M.setValue("U_TTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE))
                    Else
                        If Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE) <= Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) Then
                            Time = Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) - Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE)
                        Else
                            Time = (2400 - Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE)) + Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE)
                        End If
                        Hour = Fix(Time / 100)
                        Minute = Time Mod 100
                        Time = Hour
                        If Minute > 0 Then
                            Time = Time + 0.5
                        End If
                        Call oDS_PS_PP040M.setValue("U_NTime", pval.Row - 1, Time)
                        Call oDS_PS_PP040M.setValue("U_YTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE) - Time)
                        Call oDS_PS_PP040M.setValue("U_TTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE) - Time)
                    End If
                ElseIf pval.ColUID = "NEnd" Then
                    Call oDS_PS_PP040M.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    If Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE) = 0 Or Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) = 0 Then
                        Call oDS_PS_PP040M.setValue("U_NTime", pval.Row - 1, 0)
                        Call oDS_PS_PP040M.setValue("U_YTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE))
                        Call oDS_PS_PP040M.setValue("U_TTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE))
                    Else
                        If Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE) <= Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) Then
                            Time = Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) - Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE)
                        Else
                            Time = (2400 - Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE)) + Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE)
                        End If
                        Hour = Fix(Time / 100)
                        Minute = Time Mod 100
                        Time = Hour
                        If Minute > 0 Then
                            Time = Time + 0.5
                        End If
                        Call oDS_PS_PP040M.setValue("U_NTime", pval.Row - 1, Time)
                        Call oDS_PS_PP040M.setValue("U_YTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE) - Time)
                        Call oDS_PS_PP040M.setValue("U_TTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE) - Time)
                    End If
                ElseIf pval.ColUID = "YTime" Then
                    Call oDS_PS_PP040M.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    Call oDS_PS_PP040M.setValue("U_TTime", pval.Row - 1, Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                Else
                    Call oDS_PS_PP040M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            ElseIf (pval.ItemUID = "Mat03") Then
                If (pval.ColUID = "FailCode") Then
                    Call oDS_PS_PP040N.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_PP040N.setValue("U_FailName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" & oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                Else
                    Call oDS_PS_PP040N.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_PP040H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "BaseTime") Then
                    Call oDS_PS_PP040H.setValue("U_" & pval.ItemUID, 0, Val(oForm01.Items(pval.ItemUID).Specific.VALUE))
                ElseIf (pval.ItemUID = "OrdMgNum") Then
                    Call oDS_PS_PP040H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        Call PS_PP040_OrderInfoLoad
                    End If
                ElseIf (pval.ItemUID = "ItemCode") Then
                    Call oDS_PS_PP040H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    Call PS_PP040_AddMatrixRow01(0, True)
                    oMat02.Clear
                    oMat02.FlushToDataSource
                    oMat02.LoadFromDataSource
                    Call PS_PP040_AddMatrixRow02(0, True)
                    oMat03.Clear
                    oMat03.FlushToDataSource
                    oMat03.LoadFromDataSource
                
                ElseIf (pval.ItemUID = "UseMCode") Then
                
                    Query01 = "EXEC PS_PP040_98 '" & oForm01.Items("UseMCode").Specific.VALUE
                  
                    Call RecordSet01.DoQuery(Query01)
    
                    oForm01.Items("UseMName").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                
                Else
                    Call oDS_PS_PP040H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oMat02.LoadFromDataSource
            oMat02.AutoResizeColumns
            oMat03.LoadFromDataSource
            oMat03.AutoResizeColumns
            oForm01.Update
            If pval.ItemUID = "Mat01" Then
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            ElseIf pval.ItemUID = "Mat02" Then
                oMat02.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            ElseIf pval.ItemUID = "Mat03" Then
                oMat03.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                oForm01.Items(pval.ItemUID).Click ct_Regular
            End If
        End If
    ElseIf pval.BeforeAction = False Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Call oForm01.Freeze(False)
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_PP040_FormItemEnabled
        If pval.ItemUID = "Mat01" Then
            Call PS_PP040_AddMatrixRow01(oMat01.VisualRowCount) '//UDO방식
        ElseIf pval.ItemUID = "Mat02" Then
            Call PS_PP040_AddMatrixRow02(oMat02.VisualRowCount) '//UDO방식
        End If
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP040_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Dim oDataTable01 As SAPbouiCOM.DataTable
'        If (pval.ItemUID = "ItemCode") Then
'            Dim oDataTable01 As SAPbouiCOM.DataTable
'            Set oDataTable01 = pval.SelectedObjects
'            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP040H", "U_CardCode,U_CardName")
'        End If
        If (pval.ItemUID = "ItemCode") Then
            If pval.SelectedObjects Is Nothing Then
            Else
                Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP040H", "U_ItemCode,U_ItemName")
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Call PS_PP040_AddMatrixRow01(0, True)
                oMat02.Clear
                oMat02.FlushToDataSource
                oMat02.LoadFromDataSource
                Call PS_PP040_AddMatrixRow02(0, True)
                oMat03.Clear
                oMat03.FlushToDataSource
                oMat03.LoadFromDataSource
            End If
        End If
'        If (pval.ItemUID = "Mat02") Then
'            If (pval.ColUID = "WorkCode") Then
'                If pval.SelectedObjects Is Nothing Then
'                Else
'                    Set oDataTable01 = pval.SelectedObjects
'                    Call oDS_PS_PP040M.setValue("U_WorkCode", pval.Row - 1, oDataTable01.Columns("empID").Cells(0).Value)
'                    Call oDS_PS_PP040M.setValue("U_WorkName", pval.Row - 1, oDataTable01.Columns("firstName").Cells(0).Value & oDataTable01.Columns("lastName").Cells(0).Value)
'                    If oMat02.RowCount = pval.Row And Trim(oDS_PS_PP040M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        Call PS_PP040_AddMatrixRow02(pval.Row)
'                    End If
'                    Set oDataTable01 = Nothing
'                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
'                    oMat02.LoadFromDataSource
'                    oMat02.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
'                End If
'            End If
'        End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Or pval.ItemUID = "Mat03" Then
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
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oMat01Row01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat02" Then
        If pval.Row > 0 Then
            oMat02Row02 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat03" Then
        If pval.Row > 0 Then
            oMat03Row03 = pval.Row
        End If
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
    Dim i, j As Long
    Dim Exist As Boolean
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            If oLastItemUID01 = "Mat01" Then
                If (PS_PP040_Validate("행삭제01") = False) Then
                    BubbleEvent = False
                    Exit Sub
                End If
Continue:
                For i = 1 To oMat03.RowCount
                    If oMat01.Columns("OrdMgNum").Cells(oLastColRow01).Specific.VALUE = oMat03.Columns("OrdMgNum").Cells(i).Specific.VALUE And oMat01.Columns("LineNum").Cells(oLastColRow01).Specific.VALUE = oMat03.Columns("OLineNum").Cells(i).Specific.VALUE Then
                    '//If oMat01.Columns("OrdMgNum").Cells(oLastColRow01).Specific.VALUE = oMat03.Columns("OrdMgNum").Cells(i).Specific.VALUE Then
                        oDS_PS_PP040N.RemoveRecord (i - 1)
                        oMat03.DeleteRow (i)
                        oMat03.FlushToDataSource
                        GoTo Continue
                    End If
                Next
            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            If oLastItemUID01 = "Mat01" Then
                For i = 1 To oMat01.VisualRowCount
                    oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                
                For i = 1 To oMat03.VisualRowCount
                    If oMat03.Columns("OLineNum").Cells(i).Specific.VALUE <> 1 Then
                        oMat03.Columns("OLineNum").Cells(i).Specific.VALUE = oMat03.Columns("OLineNum").Cells(i).Specific.VALUE - 1 '//i
                    End If
                Next i
                
                oMat01.FlushToDataSource
                Call oDS_PS_PP040L.RemoveRecord(oDS_PS_PP040L.Size - 1)
                oMat01.LoadFromDataSource
                If oMat01.RowCount = 0 Then
                    Call PS_PP040_AddMatrixRow01(0)
                Else
                    If Trim(oDS_PS_PP040L.GetValue("U_OrdMgNum", oMat01.RowCount - 1)) <> "" Then
                        Call PS_PP040_AddMatrixRow01(oMat01.RowCount)
                    End If
                End If
            ElseIf oLastItemUID01 = "Mat02" Then
                For i = 1 To oMat02.VisualRowCount
                    oMat02.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                oMat02.FlushToDataSource
                Call oDS_PS_PP040M.RemoveRecord(oDS_PS_PP040M.Size - 1)
                oMat02.LoadFromDataSource
                If oMat02.RowCount = 0 Then
                    Call PS_PP040_AddMatrixRow02(0)
                Else
                    If Trim(oDS_PS_PP040M.GetValue("U_WorkCode", oMat02.RowCount - 1)) <> "" Then
                        Call PS_PP040_AddMatrixRow02(oMat02.RowCount)
                    End If
                End If
            ElseIf oLastItemUID01 = "Mat03" Then
                For i = 1 To oMat03.VisualRowCount
                    oMat03.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                oMat03.FlushToDataSource
                If oDS_PS_PP040N.Size = 1 Then '//사이즈가 0일때는 행을 빼주면 oMat03.VisualRowCount 가 0 으로 변경되어서 문제가 생김
                Else
                    Call oDS_PS_PP040N.RemoveRecord(oDS_PS_PP040N.Size - 1)
                End If
                oMat03.LoadFromDataSource
                
                For i = 1 To oMat01.RowCount - 1 '//공정 테이블에는 있는데 불량 테이블에 존재하지 않는값이 있는경우 불량테이블에 값을 추가함
                    Exist = False
                    For j = 1 To oMat03.RowCount
                        If oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat03.Columns("OrdMgNum").Cells(j).Specific.VALUE And oMat01.Columns("LineNum").Cells(i).Specific.VALUE = oMat03.Columns("OLineNum").Cells(j).Specific.VALUE Then
                        '//If oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat03.Columns("OrdMgNum").Cells(j).Specific.VALUE Then
                            Exist = True
                        End If
                    Next
                    If Exist = False Then '//불량코드테이블에 값이 존재하지 않으면
                        If oMat03.VisualRowCount = 0 Then
                            Call PS_PP040_AddMatrixRow03(0, True)
                        Else
                            Call PS_PP040_AddMatrixRow03(oMat03.VisualRowCount)
                        End If
                        Call oDS_PS_PP040N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE)
                        Call oDS_PS_PP040N.setValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns("CpCode").Cells(i).Specific.VALUE)
                        Call oDS_PS_PP040N.setValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns("CpName").Cells(i).Specific.VALUE)
                        Call oDS_PS_PP040N.setValue("U_OLineNum", oMat03.VisualRowCount - 1, i)
                        oMat03.LoadFromDataSource
                        oMat03.AutoResizeColumns
                        oMat03.Columns("OLineNum").TitleObject.Sortable = True
                        Call oMat03.Columns("OLineNum").TitleObject.Sort(gst_Ascending)
'                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
'                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
                        oMat03.FlushToDataSource
                    End If
                Next
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RECORD_MOVE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RECORD_MOVE_Error
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim DocEntry As String
    Dim DocEntryNext As String
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE) '//원본문서
    DocEntryNext = Trim(oForm01.Items("DocEntry").Specific.VALUE) '//다음문서
    
    If pval.MenuUID = "1288" Then '//다음
        If pval.BeforeAction = True Then
            If oForm01.Mode = fm_ADD_MODE Then
                Sbo_Application.ActivateMenuItem ("1290")
                BubbleEvent = False
                Exit Sub
            ElseIf oForm01.Mode = fm_FIND_MODE Then
                If (oForm01.Items("DocEntry").Specific.VALUE = "") Then
                    Sbo_Application.ActivateMenuItem ("1290")
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
            If PS_PP040_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    ElseIf pval.MenuUID = "1289" Then '//이전
        If pval.BeforeAction = True Then
            If oForm01.Mode = fm_ADD_MODE Then
                Sbo_Application.ActivateMenuItem ("1291")
                BubbleEvent = False
                Exit Sub
            ElseIf oForm01.Mode = fm_FIND_MODE Then
                If (oForm01.Items("DocEntry").Specific.VALUE = "") Then
                    Sbo_Application.ActivateMenuItem ("1291")
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
            If PS_PP040_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    ElseIf pval.MenuUID = "1290" Then '//첫번째레코드로이동
        If pval.BeforeAction = True Then
            Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry DESC" '//가장마지막행을 부여
            RecordSet01.DoQuery Query01
            DocEntry = Val(RecordSet01.Fields(0).VALUE) '//원본문서
            DocEntryNext = Val(RecordSet01.Fields(0).VALUE) '//다음문서
            Set RecordSet01 = Nothing
            If PS_PP040_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    ElseIf pval.MenuUID = "1291" Then '//마지막문서로이동
        If pval.BeforeAction = True Then
            Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry ASC" '//가장 첫행을 부여
            RecordSet01.DoQuery Query01
            DocEntry = Val(RecordSet01.Fields(0).VALUE) '//원본문서
            DocEntryNext = Val(RecordSet01.Fields(0).VALUE) '//다음문서
            Set RecordSet01 = Nothing
            If PS_PP040_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    End If
    Set RecordSet01 = Nothing
    Exit Sub
Raise_EVENT_RECORD_MOVE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RECORD_MOVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP040_CreateItems() As Boolean
On Error GoTo PS_PP040_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP040H = oForm01.DataSources.DBDataSources("@PS_PP040H")
    Set oDS_PS_PP040L = oForm01.DataSources.DBDataSources("@PS_PP040L")
    Set oDS_PS_PP040M = oForm01.DataSources.DBDataSources("@PS_PP040M")
    Set oDS_PS_PP040N = oForm01.DataSources.DBDataSources("@PS_PP040N")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns

    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
    
    Set oMat03 = oForm01.Items("Mat03").Specific
    oMat03.SelectionMode = ms_NotSupported
    oMat03.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("Opt03", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt01").Specific.DataBind.SetBound(True, "", "Opt01")
    Call oForm01.Items("Opt02").Specific.DataBind.SetBound(True, "", "Opt02")
    Call oForm01.Items("Opt03").Specific.DataBind.SetBound(True, "", "Opt03")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt03")
'    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
'    Call oForm01.DataSources.UserDataSources.Add("WhsCode", dt_SHORT_TEXT, 100)
'    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
'    Call oForm01.Items("WhsCode").Specific.DataBind.SetBound(True, "", "WhsCode")
    
    Call oForm01.DataSources.UserDataSources.Add("CardType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("CardType").Specific.DataBind.SetBound(True, "", "CardType")
        
    Call oForm01.DataSources.UserDataSources.Add("EmpChk", dt_SHORT_TEXT, 10)
    Call oForm01.Items("EmpChk").Specific.DataBind.SetBound(True, "", "EmpChk")
    
    
    oDocType01 = "작업일보등록(작지)"
    If (oDocType01 = "작업일보등록(작지)") Then
        Call oForm01.Items("DocType").Specific.Select("10", psk_ByValue)
    ElseIf (oDocType01 = "작업일보등록(공정)") Then
        Call oForm01.Items("DocType").Specific.Select("20", psk_ByValue)
    End If
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP040_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP040_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP040_ComboBox_Setting()
On Error GoTo PS_PP040_ComboBox_Setting_Error

    Dim loopCount As Integer

    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP040", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PP040", "ItemCode")

    Dim sQry As String

     '점심시간 작업 사용하지 않아서 화면디자인에서 삭제했음. 2015/04/09
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "Mat02", "LTime", "Y", "Y")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "Mat02", "LTime", "N", "N")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("LTime"), "PS_PP040", "Mat02", "LTime")
    Call oForm01.Items("BPLId").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "10", "일반")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "20", "PSMT지원")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "30", "외주")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "40", "실적")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "50", "일반조정")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "60", "외주조정")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "70", "설계시간")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("OrdType").Specific, "PS_PP040", "OrdType")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "DocType", "", "10", "작지기준")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP040", "DocType", "", "20", "공정기준")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("DocType").Specific, "PS_PP040", "DocType")
    
    
    
    Call oForm01.Items("OrdGbn").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND CODE NOT IN('104','107') order by Code", "", False, False)
    
    
    
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")

    '거래처구분 콤보(2012.02.02 송명규 추가)
    Call oForm01.Items("CardType").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", False, False)
    Call oForm01.Items("CardType").Specific.Select(0, psk_Index)

    '작업구분코드(2014.04.15 송명규 수정)
    sQry = "           SELECT      U_Minor,"
    sQry = sQry & "                U_CdName"
    sQry = sQry & " FROM       [@PS_SY001L]"
    sQry = sQry & " WHERE      Code = 'P203'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    If oMat01.Columns("WorkCls").ValidValues.Count > 0 Then
    
        For loopCount = 0 To oMat01.Columns("WorkCls").ValidValues.Count - 1
            Call oMat01.Columns("WorkCls").ValidValues.Remove(0, psk_Index)
        Next
        
        Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("WorkCls"), sQry)
    Else
        Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("WorkCls"), sQry)
    End If
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP040_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP040_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP040_CF_ChooseFromList()
On Error GoTo PS_PP040_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

    Set oEdit = oForm01.Items("ItemCode").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = "4"
    oCFLCreationParams.uniqueID = "CFLITEMCODE"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

    Set oCons = oCFL.GetConditions()
    Set oCon = oCons.Add()
    oCon.Alias = "ItmsGrpCod"
    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    oCon.CondVal = "102"
    oCFL.SetConditions oCons

    oEdit.ChooseFromListUID = "CFLITEMCODE"
    oEdit.ChooseFromListAlias = "ItemCode"
    
'    Set oColumn = oMat02.Columns("WorkCode")
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_Employee
'    oCFLCreationParams.uniqueID = "CFLEMPLOYEE"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)
'
''    Set oCons = oCFL.GetConditions()
''    Set oCon = oCons.Add()
''    oCon.Alias = "CardType"
''    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
''    oCon.CondVal = "C"
''    oCFL.SetConditions oCons
'
'    oColumn.ChooseFromListUID = "CFLEMPLOYEE"
'    oColumn.ChooseFromListAlias = "empID"
'
'    Set oEdit = oForm01.Items("CARDNAME").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_BusinessPartner
'    oCFLCreationParams.uniqueID = "CFLCARDNAME"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLCARDNAME"
'    oEdit.ChooseFromListAlias = "CardName"
    Exit Sub
PS_PP040_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP040_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP040_FormItemEnabled()
On Error GoTo PS_PP040_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("OrdType").Enabled = True
        oForm01.Items("OrdMgNum").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("Button01").Enabled = True
        oForm01.Items("1").Enabled = True
        oForm01.Items("Mat01").Enabled = True
        oForm01.Items("Mat02").Enabled = True
        oForm01.Items("Mat03").Enabled = True
'        oMat02.Columns("NStart").Editable = False  '//비가동시작시간 사용안함
'        oMat02.Columns("NEnd").Editable = False    '//비가동종료시간 사용안함
        oMat02.Columns("NTime").Editable = True    '//비가동시간만 사용
        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_Index)
        Call oForm01.Items("OrdType").Specific.Select(0, psk_Index)
        If Trim(oOrdGbn) = "" Then
            Call oForm01.Items("OrdGbn").Specific.Select(0, psk_Index)
        Else
            Call oForm01.Items("OrdGbn").Specific.Select(oOrdGbn, psk_ByValue)
        End If
        'Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        
        Call PS_PP040_FormClear '//UDO방식
        If (oDocType01 = "작업일보등록(작지)") Then
            oDS_PS_PP040H.setValue "U_DocType", 0, "10"
'            Call oForm01.Items("DocType").Specific.Select("10", psk_ByValue)
        ElseIf (oDocType01 = "작업일보등록(공정)") Then
            Call oForm01.Items("DocType").Specific.Select("20", psk_ByValue)
        End If
        If Trim(oDocdate) = "" Then
            oForm01.Items("DocDate").Specific.VALUE = Format(Now - 1, "YYYYMMDD")
        Else
            oForm01.Items("DocDate").Specific.VALUE = oDocdate
        End If
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("OrdType").Enabled = True
        oForm01.Items("OrdMgNum").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("Button01").Enabled = True
        oForm01.Items("1").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        oForm01.Items("Mat02").Enabled = False
        oForm01.Items("Mat03").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        '//각모드에따른 아이템설정
        If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" & Trim(oDS_PS_PP040H.GetValue("DocEntry", 0)) & "'", 0, 1) = "Y" Then
            oForm01.Items("Focus").Click ct_Regular
            oForm01.Items("DocEntry").Enabled = False
            oForm01.Items("OrdType").Enabled = False
            oForm01.Items("OrdMgNum").Enabled = False
            oForm01.Items("DocDate").Enabled = False
            oForm01.Items("Button01").Enabled = False
            oForm01.Items("1").Enabled = False
            oForm01.Items("Mat01").Enabled = False
            oForm01.Items("Mat02").Enabled = False
            oForm01.Items("Mat03").Enabled = False
        Else
            If Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "10" Or Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "50" Or Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "60" Or Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "70" Then '//조정, 설계
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("DocEntry").Enabled = False
                oForm01.Items("OrdType").Enabled = False
                oForm01.Items("OrdMgNum").Enabled = True
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("Button01").Enabled = True
                oForm01.Items("1").Enabled = True
                oForm01.Items("Mat01").Enabled = True
                oForm01.Items("Mat02").Enabled = True
                oForm01.Items("Mat03").Enabled = True
            ElseIf Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "20" Then '//PSMT
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("DocEntry").Enabled = False
                oForm01.Items("OrdType").Enabled = False
                oForm01.Items("OrdMgNum").Enabled = True
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("Button01").Enabled = True
                oForm01.Items("1").Enabled = True
                oForm01.Items("Mat01").Enabled = True
                oForm01.Items("Mat02").Enabled = True
                oForm01.Items("Mat03").Enabled = True
            ElseIf Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "30" Then '//외주
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("DocEntry").Enabled = False
                oForm01.Items("OrdType").Enabled = False
                oForm01.Items("OrdMgNum").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("Button01").Enabled = False
                oForm01.Items("1").Enabled = False
                oForm01.Items("Mat01").Enabled = False
                oForm01.Items("Mat02").Enabled = False
                oForm01.Items("Mat03").Enabled = False
            ElseIf Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) = "40" Then '//실적
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("DocEntry").Enabled = False
                oForm01.Items("OrdType").Enabled = False
                oForm01.Items("OrdMgNum").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("Button01").Enabled = False
                oForm01.Items("1").Enabled = False
                oForm01.Items("Mat01").Enabled = False
                oForm01.Items("Mat02").Enabled = False
                oForm01.Items("Mat03").Enabled = False
            End If
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP040_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP040_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP040_AddMatrixRow01(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP040_AddMatrixRow01_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP040L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP040L.offset = oRow
    oDS_PS_PP040L.setValue "U_LineNum", oRow, oRow + 1
    oDS_PS_PP040L.setValue "U_WorkCls", oRow, "A" '작업구분을 기본으로 선택(2014.04.15 송명규 추가)
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP040_AddMatrixRow01_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP040_AddMatrixRow01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP040_AddMatrixRow02(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP040_AddMatrixRow02_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP040M.InsertRecord (oRow)
    End If
    oMat02.AddRow
    oDS_PS_PP040M.offset = oRow
    oDS_PS_PP040M.setValue "U_LineNum", oRow, oRow + 1
    oMat02.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP040_AddMatrixRow02_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP040_AddMatrixRow02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP040_AddMatrixRow03(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP040_AddMatrixRow03_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP040N.InsertRecord (oRow)
    End If
    oMat03.AddRow
    oDS_PS_PP040N.offset = oRow
    oDS_PS_PP040N.setValue "U_LineNum", oRow, oRow + 1
    oMat03.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP040_AddMatrixRow03_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP040_AddMatrixRow03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP040_FormClear()
On Error GoTo PS_PP040_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP040'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP040_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP040_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP040_EnableMenus()
On Error GoTo PS_PP040_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP040_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP040_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP040_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP040_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP040_FormItemEnabled
        Call PS_PP040_AddMatrixRow01(0, True) '//UDO방식일때
        Call PS_PP040_AddMatrixRow02(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_PP040_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP040_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP040_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_PP040_DataValidCheck() As Boolean
On Error GoTo PS_PP040_DataValidCheck_Error
    PS_PP040_DataValidCheck = False
    Dim i, j As Long
    Dim FailQty As Double
    
    
    
    If MDC_PS_Common.GetValue("select Count(*) from OFPR Where '" & oForm01.Items("DocDate").Specific.VALUE & "' between F_RefDate and T_RefDate And PeriodStat = 'Y'") > 0 Then
        Sbo_Application.SetStatusBarMessage "해당일자는 전기기간이 잠겼습니다. 일자를 확인바랍니다.", bmt_Short, True
        PS_PP040_DataValidCheck = False
        Exit Function
    End If
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP040_FormClear
    End If
    If oForm01.Items("OrdType").Specific.Selected.VALUE <> "10" And oForm01.Items("OrdType").Specific.Selected.VALUE <> "20" And oForm01.Items("OrdType").Specific.Selected.VALUE <> "50" And oForm01.Items("OrdType").Specific.Selected.VALUE <> "60" And oForm01.Items("OrdType").Specific.Selected.VALUE <> "70" Then
        Sbo_Application.SetStatusBarMessage "작업타입이 일반, PSMT지원, 조정, 설계가 아닙니다.", bmt_Short, True
        PS_PP040_DataValidCheck = False
        Exit Function
    End If
    
    If oForm01.Items("OrdNum").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "작지번호는 필수입니다.", bmt_Short, True
        oForm01.Items("OrdNum").Click ct_Regular
        PS_PP040_DataValidCheck = False
        Exit Function
    End If
    
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "공정정보 라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP040_DataValidCheck = False
        Exit Function
    End If
    If oMat02.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "작업자정보 라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP040_DataValidCheck = False
        Exit Function
    End If
    
    '마감상태 체크_S(2017.11.23 송명규 추가)
    If MDC_PS_Common.Check_Finish_Status(Trim(oForm01.Items("BPLId").Specific.VALUE), oForm01.Items("DocDate").Specific.VALUE, oForm01.TypeEx) = False Then
        Call Sbo_Application.SetStatusBarMessage("마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작업일보일자를 확인하고, 회계부서로 문의하세요.", bmt_Short, True)
        PS_PP040_DataValidCheck = False
        Exit Function
    End If
    '마감상태 체크_E(2017.11.23 송명규 추가)
    
' 작업자 1명이상 가능토록 수정 (이병각)
'    If oMat02.VisualRowCount > 2 Then '//한명이상 입력했을경우
'        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드
'            Sbo_Application.SetStatusBarMessage "작업자정보 한명만 입력할수 있습니다.", bmt_Short, True
'            PS_PP040_DataValidCheck = False
'            Exit Function
'        Else
'            '//휘팅,부품은 여러명 입력할수 있다.
'        End If
'    End If

    If oMat03.VisualRowCount = 0 Then
        Sbo_Application.SetStatusBarMessage "불량정보 라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP040_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "작지문서번호는 필수입니다.", bmt_Short, True
            oMat01.Columns("OrdMgNum").Cells(i).Click ct_Regular
            PS_PP040_DataValidCheck = False
            Exit Function
        End If
        If Trim(oForm01.Items("OrdType").Specific.VALUE) <> "50" And Trim(oForm01.Items("OrdType").Specific.VALUE) <> "60" Then
            If (Val(oMat01.Columns("PQty").Cells(i).Specific.VALUE) <= 0) Then
                Sbo_Application.SetStatusBarMessage "생산수량은 필수입니다.", bmt_Short, True
                oMat01.Columns("PQty").Cells(i).Click ct_Regular
                PS_PP040_DataValidCheck = False
                Exit Function
            End If
        End If
        If Trim(oForm01.Items("OrdType").Specific.VALUE) <> "50" And Trim(oForm01.Items("OrdType").Specific.VALUE) <> "60" And Trim(oForm01.Items("OrdType").Specific.VALUE) <> "70" Then
            If (Val(oMat01.Columns("WorkTime").Cells(i).Specific.VALUE) <= 0) Then
                Sbo_Application.SetStatusBarMessage "실동시간은 필수입니다.", bmt_Short, True
                oMat01.Columns("WorkTime").Cells(i).Click ct_Regular
                PS_PP040_DataValidCheck = False
                Exit Function
            End If
        End If
        
        '작업완료여부(2012.02.02. 송명규 추가)
        If Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "105" Or Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "106" Then '기계공구, 몰드일 경우만 작업완료여부 필수 체크
        
            If (oMat01.Columns("CompltYN").Cells(i).Specific.VALUE = "%") Then
                Sbo_Application.SetStatusBarMessage "작업구분이 기계공구, 몰드일경우는 작업완료여부가 필수입니다. 확인하십시오.", bmt_Short, True
                oMat01.Columns("CompltYN").Cells(i).Click ct_Regular
                PS_PP040_DataValidCheck = False
                Exit Function
            End If
        
        End If
        
        '//불량수량 검사
        FailQty = 0
        For j = 1 To oMat03.VisualRowCount
            '//불량코드를 입력했는지 check
            If Val(oMat03.Columns("FailQty").Cells(j).Specific.VALUE) <> 0 And oMat03.Columns("FailCode").Cells(j).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "불량수량이 입력되었을 때는 불량코드는 필수입니다.", bmt_Short, True
                oMat03.Columns("FailCode").Cells(i).Click ct_Regular
                PS_PP040_DataValidCheck = False
                Exit Function
            End If
            
            If Val(oMat03.Columns("FailQty").Cells(j).Specific.VALUE) = 0 And oMat03.Columns("FailCode").Cells(j).Specific.VALUE <> "" Then
                Sbo_Application.SetStatusBarMessage "불량코드를 확인하세요.", bmt_Short, True
                oMat03.Columns("FailCode").Cells(i).Click ct_Regular
                PS_PP040_DataValidCheck = False
                Exit Function
            End If
                
            If (oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat03.Columns("OrdMgNum").Cells(j).Specific.VALUE) And (oMat01.Columns("LineNum").Cells(i).Specific.VALUE = oMat03.Columns("OLineNum").Cells(j).Specific.VALUE) Then
                FailQty = FailQty + Val(oMat03.Columns("FailQty").Cells(j).Specific.VALUE)
            End If
        Next
        If Trim(oForm01.Items("OrdType").Specific.VALUE) <> "50" And Trim(oForm01.Items("OrdType").Specific.VALUE) <> "60" Then
            If Val(oMat01.Columns("NQty").Cells(i).Specific.VALUE) <> FailQty Then
                Sbo_Application.SetStatusBarMessage "공정리스트의 불량수량과 불량정보의 불량수량이 일치하지 않습니다.", bmt_Short, True
                PS_PP040_DataValidCheck = False
                Exit Function
            End If
        End If
        
        If Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "601" Or Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "111" Then
           'If oMat01.Columns("CpCode").Cells(i).Specific.VALUE = "CP80101" And Trim(oMat01.Columns("CItemCod").Cells(i).Specific.VALUE) = "" Then
           If oMat01.Columns("Sequence").Cells(i).Specific.VALUE = 1 And Trim(oMat01.Columns("CItemCod").Cells(i).Specific.VALUE) = "" Then
                Sbo_Application.SetStatusBarMessage "공정 사용 원재료코드가 없습니다. 사용 원재료를 선택해 주세요", bmt_Short, True
                PS_PP040_DataValidCheck = False
                Exit Function
            End If
        End If
    Next
    
    '비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_S
    For i = 1 To oMat02.VisualRowCount - 1
    
        If (oMat02.Columns("NCode").Cells(i).Specific.VALUE <> "") Then
            
            If oMat02.Columns("NTime").Cells(i).Specific.VALUE = "" Then
            
                Sbo_Application.SetStatusBarMessage "비가동코드가 입력되었을 때는 비가동시간은 필수입니다.", bmt_Short, True
                oMat02.Columns("NTime").Cells(i).Click ct_Regular
                PS_PP040_DataValidCheck = False
                Exit Function
            
            End If
        
        End If
        
        If (oMat02.Columns("NTime").Cells(i).Specific.VALUE <> "") Then
            
            If oMat02.Columns("NCode").Cells(i).Specific.VALUE = "" Then
            
                Sbo_Application.SetStatusBarMessage "비가동시간이 입력되었을 때는 비가동코드는 필수입니다.", bmt_Short, True
                oMat02.Columns("NCode").Cells(i).Click ct_Regular
                PS_PP040_DataValidCheck = False
                Exit Function
            
            End If
        
        End If
    
    Next
    '비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_E
    
    If (PS_PP040_Validate("검사01") = False) Then
        PS_PP040_DataValidCheck = False
        Exit Function
    End If
    
    Call oDS_PS_PP040L.RemoveRecord(oDS_PS_PP040L.Size - 1)
    Call oMat01.LoadFromDataSource
    Call oDS_PS_PP040M.RemoveRecord(oDS_PS_PP040M.Size - 1)
    Call oMat02.LoadFromDataSource
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP040_FormClear
    End If
    PS_PP040_DataValidCheck = True
    Exit Function
    
PS_PP040_DataValidCheck_Error:
    PS_PP040_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP040_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP040_MTX01()
On Error GoTo PS_PP040_MTX01_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Param01 = Trim(oForm01.Items("Param01").Specific.VALUE)
    Param02 = Trim(oForm01.Items("Param01").Specific.VALUE)
    Param03 = Trim(oForm01.Items("Param01").Specific.VALUE)
    Param04 = Trim(oForm01.Items("Param01").Specific.VALUE)
    
    Query01 = "SELECT 10"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP040_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_PP040L.InsertRecord (i)
        End If
        oDS_PS_PP040L.offset = i
        Call oDS_PS_PP040L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_PP040L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP040_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_PP040_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP040_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP040_SumWorkTime()
'******************************************************************************
'Function ID    : PS_PP040_SumWorkTime()
'해 당 모 듈    : 생산관리
'기        능    : 근무시간의 총합을 구함
'인        수    : 없음
'반   환   값   : 없음
'특 이 사 항    : 없음
'******************************************************************************
On Error GoTo PS_PP040_SumWorkTime_Error

    Dim loopCount As Integer
    Dim Total As Double
    
    For loopCount = 0 To oMat01.RowCount - 2
        Total = Total + CDbl(IIf(Trim(oMat01.Columns("WorkTime").Cells(loopCount + 1).Specific.VALUE) = "", 0, Trim(oMat01.Columns("WorkTime").Cells(loopCount + 1).Specific.VALUE)))
    Next
    
    oForm01.Items("Total").Specific.VALUE = Format(Total, "##0.#0")
    
    Exit Sub

PS_PP040_SumWorkTime_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP040_SumWorkTime_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP040_DI_API() As Boolean
On Error GoTo PS_PP040_DI_API_Error
    PS_PP040_DI_API = True
    Dim i, j As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim LineNumCount As Long
    Dim ResultDocNum As Long
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.VALUE, "&&&&-&&-&&")
    For i = 0 To ItemInformationCount - 1
        If ItemInformation(i).Check = True Then
            GoTo Continue_First
        End If
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
        oDIObject.Lines.BaseType = "22"
        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
        For j = i To UBound(ItemInformation)
            If ItemInformation(j).Check = True Then
                GoTo Continue_Second
            End If
            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
                GoTo Continue_Second
            End If
            '//같은것
            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.Add
            ItemInformation(j).PDN1No = LineNumCount
            ItemInformation(j).Check = True
Continue_Second:
        Next
        LineNumCount = LineNumCount + 1
Continue_First:
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To UBound(ItemInformation)
            Call oDS_PS_PP040L.setValue("U_OPDNNo", i, ResultDocNum)
            Call oDS_PS_PP040L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        Next
    Else
        GoTo PS_PP040_DI_API_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_PP040_DI_API_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_PP040_DI_API = False
    Set oDIObject = Nothing
    Exit Function
PS_PP040_DI_API_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP040_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP040_DI_API = False
End Function

Private Sub PS_PP040_FormResize()
On Error GoTo PS_PP040_FormResize_Error
    oForm01.Items("Mat02").Top = 170
    oForm01.Items("Mat02").Left = 7
    oForm01.Items("Mat02").Height = ((oForm01.Height - 170) / 3 * 1) - 20
    oForm01.Items("Mat02").Width = oForm01.Width / 2 - 14
    
    oForm01.Items("Mat03").Top = 170
    oForm01.Items("Mat03").Left = oForm01.Width / 2
    oForm01.Items("Mat03").Height = ((oForm01.Height - 170) / 3 * 1) - 20
    oForm01.Items("Mat03").Width = oForm01.Width / 2 - 14
    
    oForm01.Items("Mat01").Top = oForm01.Items("Mat03").Top + oForm01.Items("Mat03").Height + 40
    oForm01.Items("Mat01").Left = 7
    oForm01.Items("Mat01").Height = ((oForm01.Height - 170) / 3 * 2) - 80
    oForm01.Items("Mat01").Width = oForm01.Width - 21
    
    oForm01.Items("Opt01").Left = 10
    oForm01.Items("Opt02").Left = oForm01.Width / 2
    oForm01.Items("Opt03").Left = 10
    oForm01.Items("Opt03").Top = oForm01.Items("Mat03").Top + oForm01.Items("Mat03").Height + 20
    Exit Sub
PS_PP040_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP040_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP040_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_PP040_Validate_Error
    PS_PP040_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    Dim PrevDBCpQty As Long
    Dim PrevMATRIXCpQty As Long
    Dim CurrentDBCpQty As Long
    Dim CurrentMATRIXCpQty As Long
    Dim NextDBCpQty As Long
    Dim NextMATRIXCpQty As Long
    Dim PrevCpInfo As String
    Dim CurrentCpInfo As String
    Dim NextCpInfo As String
    
    Dim OrdMgNum As String
    Dim Exist As Boolean
    
    If (oForm01.Mode <> fm_ADD_MODE) Then
        If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
            Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
            PS_PP040_Validate = False
            GoTo PS_PP040_Validate_Exit
        End If
    End If
    
    If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "60" Then '//작업타입이 일반,조정인경우
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT지원인경우
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
        MDC_Com.MDC_GF_Message "해당작업타입은 변경이 불가능합니다.", "W"
        PS_PP040_Validate = False
        GoTo PS_PP040_Validate_Exit
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
        MDC_Com.MDC_GF_Message "해당작업타입은 변경이 불가능합니다.", "W"
        PS_PP040_Validate = False
        GoTo PS_PP040_Validate_Exit
    End If
    
    If ValidateType = "검사01" Then
        If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Then '//작업타입이 일반인경우
            For i = 1 To oMat01.VisualRowCount - 1 '//입력된 행에 대해
                If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" & oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE & "'", 0, 1) <= 0 Then
                    Call MDC_Com.MDC_GF_Message("작업지시문서가 존재하지 않습니다.", "W")
                    PS_PP040_Validate = False
                    GoTo PS_PP040_Validate_Exit
                End If
            Next
            
            If (oForm01.Mode <> fm_ADD_MODE) Then
                '//삭제된 행에 대한처리
                Query01 = "SELECT "
                Query01 = Query01 & " PS_PP040H.DocEntry,"
                Query01 = Query01 & " PS_PP040L.LineId,"
                Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,"
                Query01 = Query01 & " PS_PP040L.U_OrdGbn AS OrdGbn,"
                Query01 = Query01 & " PS_PP040L.U_PP030HNo AS PP030HNo,"
                Query01 = Query01 & " PS_PP040L.U_PP030MNo AS PP030MNo,"
                Query01 = Query01 & " PS_PP040L.U_OrdMgNum AS OrdMgNum "
                Query01 = Query01 & " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry "
                Query01 = Query01 & " WHERE PS_PP040L.DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
                RecordSet01.DoQuery Query01
                For i = 0 To RecordSet01.RecordCount - 1
                    Exist = False
                    For j = 1 To oMat01.VisualRowCount - 1 '//기존에 있는 행에대한처리
                        If (oMat01.Columns("LineId").Cells(j).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
                        Else
                            '//라인번호가 같고, 문서번호가 같으면 존재하는행
                            If Val(RecordSet01.Fields(0).VALUE) = Val(oForm01.Items("DocEntry").Specific.VALUE) And Val(RecordSet01.Fields(1).VALUE) = Val(oMat01.Columns("LineId").Cells(j).Specific.VALUE) Then
                                Exist = True
                            End If
                        End If
                    Next
                    If Exist = False Then '//삭제된 행중 수량관계를 알아본다.
                        If RecordSet01.Fields("OrdGbn").VALUE = "101" Then '//휘팅이면서
                            '//현재 공정이 실적공정이면..
                            If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1) = "Y" Then
                                '//휘팅벌크포장
    '                            PP040_CurrentPQty = 0
    '                            PP040_DBPQty = MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").Value & "' AND PS_PP040L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").Value & "'", 0, 1)
    '                            PP070_DBPQty = MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").Value & "' AND PS_PP070L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").Value & "'", 0, 1)
    '                            PP080_DBPQty = MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").Value & "' AND PS_PP070L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").Value & "'", 0, 1)
                                
                                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP070L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
                                    MDC_Com.MDC_GF_Message "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
                                    PS_PP040_Validate = False
                                    GoTo PS_PP040_Validate_Exit
                                End If
                                '//휘팅실적
                                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP080L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
                                    MDC_Com.MDC_GF_Message "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
                                    PS_PP040_Validate = False
                                    GoTo PS_PP040_Validate_Exit
                                End If
                            End If
                        End If
                        
                        If RecordSet01.Fields("OrdGbn").VALUE = "105" Or RecordSet01.Fields("OrdGbn").VALUE = "106" Then '//기계공구,몰드
                            '//그냥 입력가능
                        ElseIf RecordSet01.Fields("OrdGbn").VALUE = "101" Or RecordSet01.Fields("OrdGbn").VALUE = "102" Then '//휘팅,부품
                            '//삭제된 행에 대한 검사..
                            OrdMgNum = RecordSet01.Fields("OrdMgNum").VALUE '// DocEntry + '-' + LineId
                            CurrentCpInfo = OrdMgNum
                            
                            PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_02 '" & OrdMgNum & "'")
                            If PrevCpInfo = "" Then
                                '//해당공정이 첫공정이면 입력되어도 상관없다.
                            Else
                                PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & PrevCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040H.Canceled = 'N'")
                                '//재공이동 수량
                                PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" & PrevCpInfo & "' AND a.Canceled = 'N'")
                                PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" & PrevCpInfo & "' AND a.Canceled = 'N'")
                                
                                PrevMATRIXCpQty = 0
                                For j = 1 To oMat01.VisualRowCount - 1
                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = PrevCpInfo) Then
                                        PrevMATRIXCpQty = PrevMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
                                    End If
                                Next
                                CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040H.Canceled = 'N'")
                                CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" & CurrentCpInfo & "' AND a.Canceled = 'N'")
                                CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" & CurrentCpInfo & "' AND a.Canceled = 'N'")
                                
                                CurrentMATRIXCpQty = 0
                                For j = 1 To oMat01.VisualRowCount - 1
                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = CurrentCpInfo) Then
                                        CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
                                    End If
                                Next
                                If ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
                                    Sbo_Application.SetStatusBarMessage "삭제된 공정의 선행공정의 생산수량이 삭제된 공정의 생산수량을 미달합니다.", bmt_Short, True
                                    PS_PP040_Validate = False
                                    GoTo PS_PP040_Validate_Exit
                                End If
                            End If
    '                        If oForm01.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
    '                            NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & OrdMgNum & "'")
    '                            If NextCpInfo = "" Then
    '                                '//해당공정이 마지막공정이면 삭제되어도 상관없다.
    '                            Else
    '                                NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
    '                                NextMATRIXCpQty = 0
    '                                For j = 1 To oMat01.VisualRowCount - 1
    '                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = NextCpInfo) Then
    '                                        NextMATRIXCpQty = NextMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
    '                                    End If
    '                                Next
    '                                CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
    '                                CurrentMATRIXCpQty = 0
    '                                For j = 1 To oMat01.VisualRowCount - 1 '//현재공정은 삭제되었으므로.. 매트릭스에 존재하지 않는다.
    '                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = CurrentCpInfo) Then
    '                                        CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
    '                                    End If
    '                                Next
    '                                If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
    '                                    Sbo_Application.SetStatusBarMessage "삭제된 공정의 후행공정의 생산수량이 삭제된 공정의 생산수량을 초과합니다.", bmt_Short, True
    '                                    PS_PP040_Validate = False
    '                                    GoTo PS_PP040_Validate_Exit
    '                                End If
    '                            End If
    '                        End If
                        End If
                    End If
                    RecordSet01.MoveNext
                Next
            End If
            
            If (oForm01.Mode <> fm_ADD_MODE) Then
                For i = 1 To oMat01.VisualRowCount - 1
                    If (oMat01.Columns("LineId").Cells(i).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
                    Else
                        If oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE = "101" Then '//휘팅이면서
                            '//현재공정이 실적공정이면
                            If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "-" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1) = "Y" Then  '//현재 공정이 바렐 앞공정이면..
    '                            '//휘팅벌크포장,휘팅실적
                                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP070L.U_PP030MNo = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 _
                                Or (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP080L.U_PP030MNo = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 Then
                                    '//작업일보등록된문서중에 수정이 된문서를 구함
                                    Query01 = "SELECT "
                                    Query01 = Query01 & " PS_PP040L.U_OrdMgNum,"
                                    Query01 = Query01 & " PS_PP040L.U_Sequence,"
                                    Query01 = Query01 & " PS_PP040L.U_CpCode,"
                                    Query01 = Query01 & " PS_PP040L.U_ItemCode,"
                                    Query01 = Query01 & " PS_PP040L.U_PP030HNo,"
                                    Query01 = Query01 & " PS_PP040L.U_PP030MNo,"
                                    Query01 = Query01 & " PS_PP040L.U_PQty,"
                                    Query01 = Query01 & " PS_PP040L.U_NQty,"
                                    Query01 = Query01 & " PS_PP040L.U_ScrapWt,"
                                    Query01 = Query01 & " PS_PP040L.U_WorkTime"
                                    Query01 = Query01 & " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry"
                                    Query01 = Query01 & " WHERE PS_PP040H.DocEntry = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "'"
                                    Query01 = Query01 & " AND PS_PP040L.LineId = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'"
                                    Query01 = Query01 & " AND PS_PP040H.Canceled = 'N'"
                                    RecordSet01.DoQuery Query01
                                    If (RecordSet01.Fields(0).VALUE = oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE _
                                    And RecordSet01.Fields(1).VALUE = oMat01.Columns("Sequence").Cells(i).Specific.VALUE _
                                    And RecordSet01.Fields(2).VALUE = oMat01.Columns("CpCode").Cells(i).Specific.VALUE _
                                    And RecordSet01.Fields(3).VALUE = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE _
                                    And RecordSet01.Fields(4).VALUE = oMat01.Columns("PP030HNo").Cells(i).Specific.VALUE _
                                    And RecordSet01.Fields(5).VALUE = oMat01.Columns("PP030MNo").Cells(i).Specific.VALUE _
                                    And RecordSet01.Fields(6).VALUE = oMat01.Columns("PQty").Cells(i).Specific.VALUE _
                                    And RecordSet01.Fields(7).VALUE = oMat01.Columns("NQty").Cells(i).Specific.VALUE _
                                    And RecordSet01.Fields(8).VALUE = oMat01.Columns("ScrapWt").Cells(i).Specific.VALUE _
                                    And RecordSet01.Fields(9).VALUE = oMat01.Columns("WorkTime").Cells(i).Specific.VALUE _
                                    ) Then
                                    Else '//값이 변경된 행의경우
                                        MDC_Com.MDC_GF_Message "생산실적이 등록된 행은 수정할수 없습니다.", "W"
                                        PS_PP040_Validate = False
                                        GoTo PS_PP040_Validate_Exit
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            
            For i = 1 To oMat01.VisualRowCount - 1 '//입력된 모든행에 대해 입력가능성 검사
                If oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE = "105" Or oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE = "106" Then '//기계공구,몰드
                    '//그냥 입력가능
                ElseIf oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE = "101" Or oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE = "102" Then '//휘팅,부품
                    OrdMgNum = oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE
                    CurrentCpInfo = OrdMgNum
                    
                    PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_02 '" & OrdMgNum & "'")
                    If PrevCpInfo = "" Then
                        '//해당공정이 첫공정이면 입력되어도 상관없다.
                    Else
                        
                        'PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & PrevCpInfo & "' AND PS_PP040H.DocEntry <> '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND PS_PP040H.Canceled = 'N'")
                        PrevDBCpQty = MDC_PS_Common.GetValue("EXEC PS_PP040_07 '" & PrevCpInfo & "', '" & oForm01.Items("DocEntry").Specific.VALUE & "'")
                        '//재공 이동수량 반영
                        PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" & PrevCpInfo & "' AND a.Canceled = 'N'")
                        PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" & PrevCpInfo & "' AND a.Canceled = 'N'")
                        
                        PrevMATRIXCpQty = 0
                        For j = 1 To oMat01.VisualRowCount - 1
                            If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = PrevCpInfo) Then
                                PrevMATRIXCpQty = PrevMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
                            End If
                        Next
                        'CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND PS_PP040H.Canceled = 'N'")
                        CurrentDBCpQty = MDC_PS_Common.GetValue("EXEC PS_PP040_07 '" & CurrentCpInfo & "', '" & oForm01.Items("DocEntry").Specific.VALUE & "'")
                        CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" & CurrentCpInfo & "' AND a.Canceled = 'N'")
                        CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" & CurrentCpInfo & "' AND a.Canceled = 'N'")
                        
                        CurrentMATRIXCpQty = 0
                        For j = 1 To oMat01.VisualRowCount - 1
                            If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = CurrentCpInfo) Then
                                CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
                            End If
                        Next
                        '// 노대리님 요청 주석
                        If ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
                            Sbo_Application.SetStatusBarMessage "선행공정의 생산수량이 현공정의 생산수량에 미달 합니다.", bmt_Short, True
                            Call oMat01.SelectRow(i, True, False)
                            PS_PP040_Validate = False
                            GoTo PS_PP040_Validate_Exit
                        End If

                    End If
'                    If oForm01.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
'                        NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & OrdMgNum & "'")
'                        If NextCpInfo = "" Then
'                            '//해당공정이 마지막공정이면 삭제되어도 상관없다.
'                        Else
'                            NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP040H.DocEntry <> '" & oForm01.Items("DocEntry").Specific.Value & "' AND PS_PP040H.Canceled = 'N'")
'                            NextMATRIXCpQty = 0
'                            For j = 1 To oMat01.VisualRowCount - 1
'                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = NextCpInfo) Then
'                                    NextMATRIXCpQty = NextMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
'                                End If
'                            Next
'                            CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & oForm01.Items("DocEntry").Specific.Value & "' AND PS_PP040H.Canceled = 'N'")
'                            CurrentMATRIXCpQty = 0
'                            For j = 1 To oMat01.VisualRowCount - 1 '//현재공정은 삭제되었으므로.. 매트릭스에 존재하지 않는다.
'                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = CurrentCpInfo) Then
'                                    CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
'                                End If
'                            Next
'                            If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
'                                Sbo_Application.SetStatusBarMessage "후행공정의 생산수량이 현공정의 생산수량을 초과 합니다.", bmt_Short, True
'                                PS_PP040_Validate = False
'                                GoTo PS_PP040_Validate_Exit
'                            End If
'                        End If
'                    End If
                End If
            Next
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT지원인경우
            '//현재는 특별한 조건이 필요치 않음
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
            '//현재는 특별한 조건이 필요치 않음
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
            '//현재는 특별한 조건이 필요치 않음
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 조정인경우
            '//현재는 특별한 조건이 필요치 않음
        End If
    ElseIf ValidateType = "행삭제01" Then
        If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Then '//작업타입이 일반인경우
            '//행삭제전 행삭제가능여부검사
            If (oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
            Else
                If oMat01.Columns("OrdGbn").Cells(oMat01Row01).Specific.VALUE = "101" Then '//휘팅이면서
                    '//현재공정이 실적공정이면
                    If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "-" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1) = "Y" Then  '//현재 공정이 바렐 앞공정이면..
                        '//휘팅벌크포장
                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP070L.U_PP030MNo = '" & Val(oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
                            MDC_Com.MDC_GF_Message "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
                            PS_PP040_Validate = False
                            GoTo PS_PP040_Validate_Exit
                        End If
                        '//휘팅실적
                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP080L.U_PP030MNo = '" & Val(oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
                            MDC_Com.MDC_GF_Message "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
                            PS_PP040_Validate = False
                            GoTo PS_PP040_Validate_Exit
                        End If
                    End If
                
                ElseIf oMat01.Columns("OrdGbn").Cells(oMat01Row01).Specific.VALUE = "105" Or oMat01.Columns("OrdGbn").Cells(oMat01Row01).Specific.VALUE = "106" Then '//기계공구,몰드
                
                    '재고가 존재하면 행삭제 불가 기능 추가(2011.12.15 송명규 추가)
                    Dim QueryString As String
                    
                    QueryString = "                     SELECT      SUM(A.InQty) - SUM(A.OutQty) AS [StockQty]"
                    QueryString = QueryString & "  FROM       OINM AS A"
                    QueryString = QueryString & "                 INNER JOIN"
                    QueryString = QueryString & "                 OITM As B"
                    QueryString = QueryString & "                     ON A.ItemCode = B.ItemCode"
                    QueryString = QueryString & "  WHERE      B.U_ItmBsort IN ('105','106')"
                    QueryString = QueryString & "                 AND A.ItemCode = '" & oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.VALUE & "'"
                    QueryString = QueryString & "  GROUP BY  A.ItemCode"
                
                    If IIf((MDC_PS_Common.GetValue(QueryString, 0, 1)) = "", 0, (MDC_PS_Common.GetValue(QueryString, 0, 1))) > 0 Then
                    
                        MDC_Com.MDC_GF_Message "재고가 존재하는 작번입니다. 삭제할 수 없습니다.", "W"
                        PS_PP040_Validate = False
                        GoTo PS_PP040_Validate_Exit
                        
                    End If
                    
                End If
                
            End If
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 조정인경
        
        End If
    ElseIf ValidateType = "수정01" Then
        If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Then '//작업타입이 일반인경우
            '//수정전 수정가능여부검사
            If (oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE = "") Then '//새로추가된 행인경우, 수정하여도 무방하다
            Else
                If oMat01.Columns("OrdGbn").Cells(oMat01Row01).Specific.VALUE = "101" Then '//휘팅이면서
                    '//현재공정이 실적공정이면
                    If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "-" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1) = "Y" Then  '//현재 공정이 바렐 앞공정이면..
                        '//휘팅벌크포장
                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP070L.U_PP030MNo = '" & Val(oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
                            MDC_Com.MDC_GF_Message "수정된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
                            PS_PP040_Validate = False
                            GoTo PS_PP040_Validate_Exit
                        End If
                        '//휘팅실적
                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP080L.U_PP030MNo = '" & Val(oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
                            MDC_Com.MDC_GF_Message "수정된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
                            PS_PP040_Validate = False
                            GoTo PS_PP040_Validate_Exit
                        End If
                    End If
                End If
            End If
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 조정인경우
        End If
    ElseIf ValidateType = "취소" Then
        If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
            MDC_Com.MDC_GF_Message "이미취소된 문서 입니다. 취소할수 없습니다.", "W"
            PS_PP040_Validate = False
            GoTo PS_PP040_Validate_Exit
        End If
        If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Then '//작업타입이 일반인경우
            '//삭제된 행에 대한처리
            Query01 = "SELECT "
            Query01 = Query01 & " PS_PP040H.DocEntry,"
            Query01 = Query01 & " PS_PP040L.LineId,"
            Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,"
            Query01 = Query01 & " PS_PP040L.U_OrdGbn AS OrdGbn,"
            Query01 = Query01 & " PS_PP040L.U_PP030HNo AS PP030HNo,"
            Query01 = Query01 & " PS_PP040L.U_PP030MNo AS PP030MNo,"
            Query01 = Query01 & " PS_PP040L.U_OrdMgNum AS OrdMgNum "
            Query01 = Query01 & " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry "
            Query01 = Query01 & " WHERE PS_PP040L.DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
            RecordSet01.DoQuery Query01
            For i = 0 To RecordSet01.RecordCount - 1
                If RecordSet01.Fields("OrdGbn").VALUE = "101" Then '//휘팅이면서
                    '//현재공정이 실적포인트이면
                    If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1) = "Y" Then
                        '//휘팅벌크포장
                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP070L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
                            MDC_Com.MDC_GF_Message "생산실적 등록된 문서입니다. 적용할수 없습니다.", "W"
                            PS_PP040_Validate = False
                            GoTo PS_PP040_Validate_Exit
                        End If
                        '//휘팅실적
                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP080L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
                            MDC_Com.MDC_GF_Message "생산실적 등록된 문서입니다. 적용할수 없습니다.", "W"
                            PS_PP040_Validate = False
                            GoTo PS_PP040_Validate_Exit
                        End If
                    End If
                End If
                
                If RecordSet01.Fields("OrdGbn").VALUE = "105" Or RecordSet01.Fields("OrdGbn").VALUE = "106" Then '//기계공구,몰드
                    '//그냥 입력가능
                ElseIf RecordSet01.Fields("OrdGbn").VALUE = "101" Or RecordSet01.Fields("OrdGbn").VALUE = "102" Then '//휘팅,부품
                    '//삭제된 행에 대한 검사..
                    OrdMgNum = RecordSet01.Fields("OrdMgNum").VALUE '// DocEntry + '-' + LineId
                    CurrentCpInfo = OrdMgNum
                    
                    PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_02 '" & OrdMgNum & "'")
                    If PrevCpInfo = "" Then
                        '//해당공정이 첫공정이면 입력되어도 상관없다.
                    Else
                        PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & PrevCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040H.Canceled = 'N'")
                        PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" & PrevCpInfo & "' AND a.Canceled = 'N'")
                        PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" & PrevCpInfo & "' AND a.Canceled = 'N'")
                        
                        PrevMATRIXCpQty = 0
                        CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040H.Canceled = 'N'")
                        CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" & CurrentCpInfo & "' AND a.Canceled = 'N'")
                        CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" & CurrentCpInfo & "' AND a.Canceled = 'N'")
                        CurrentMATRIXCpQty = 0
                        If ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
                            Sbo_Application.SetStatusBarMessage "취소문서의 선행공정의 생산수량이 취소문서의 생산수량을 미달합니다.", bmt_Short, True
                            PS_PP040_Validate = False
                            GoTo PS_PP040_Validate_Exit
                        End If
                    End If
                    
'                    If oForm01.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
'                        NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & OrdMgNum & "'")
'                        If NextCpInfo = "" Then
'                            '//해당공정이 마지막공정이면 삭제되어도 상관없다.
'                        Else
'                            NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
'                            NextMATRIXCpQty = 0
'                            CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
'                            CurrentMATRIXCpQty = 0
'                            If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
'                                Sbo_Application.SetStatusBarMessage "취소문서의 후행공정의 생산수량이 취소문서의 생산수량을 초과합니다.", bmt_Short, True
'                                PS_PP040_Validate = False
'                                GoTo PS_PP040_Validate_Exit
'                            End If
'                        End If
'                    End If
                End If
                RecordSet01.MoveNext
            Next
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT인경우
        
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
            
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
        
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 조정인경우
        
        End If
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_PP040_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_PP040_Validate_Error:
    PS_PP040_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_PP040_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP040_OrderInfoLoad()
On Error GoTo PS_PP040_OrderInfoLoad_Error
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "60" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "70" Then  '//일반,조정, 설계
        If oForm01.Items("OrdMgNum").Specific.VALUE = "" Then
            Call MDC_Com.MDC_GF_Message("작업지시 관리번호를 입력하지 않습니다.", "W")
            GoTo PS_PP040_OrderInfoLoad_Exit
        Else
            Query01 = "SELECT "
            Query01 = Query01 & "U_OrdGbn,"
            Query01 = Query01 & "U_BPLId,"
            Query01 = Query01 & "U_ItemCode,"
            Query01 = Query01 & "U_ItemName,"
            Query01 = Query01 & "U_OrdNum,"
            Query01 = Query01 & "U_OrdSub1,"
            Query01 = Query01 & "U_OrdSub2,"
            Query01 = Query01 & "DocEntry"
            Query01 = Query01 & " FROM [@PS_PP030H]"
            Query01 = Query01 & " WHERE "
            Query01 = Query01 & " U_OrdNum + U_OrdSub1 + U_OrdSub2 = '" & oForm01.Items("OrdMgNum").Specific.VALUE & "'"
            Query01 = Query01 & " AND U_OrdGbn NOT IN('104','107') "
            Query01 = Query01 & " AND Canceled = 'N'"
            RecordSet01.DoQuery Query01
            If RecordSet01.RecordCount = 0 Then
                Call MDC_Com.MDC_GF_Message("작업지시 정보가 존재하지 않습니다.", "W")
                GoTo PS_PP040_OrderInfoLoad_Exit
            Else
                Call oForm01.Items("OrdGbn").Specific.Select(RecordSet01.Fields(0).VALUE, psk_ByValue)
                Call oForm01.Items("BPLId").Specific.Select(RecordSet01.Fields(1).VALUE, psk_ByValue)
                oForm01.Items("ItemCode").Specific.VALUE = RecordSet01.Fields(2).VALUE
                oForm01.Items("ItemName").Specific.VALUE = RecordSet01.Fields(3).VALUE
                oForm01.Items("OrdNum").Specific.VALUE = RecordSet01.Fields(4).VALUE
                oForm01.Items("OrdSub1").Specific.VALUE = RecordSet01.Fields(5).VALUE
                oForm01.Items("OrdSub2").Specific.VALUE = RecordSet01.Fields(6).VALUE
                oForm01.Items("PP030HNo").Specific.VALUE = RecordSet01.Fields(7).VALUE
'                '//매트릭스삭제
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
'                Call PS_PP040_AddMatrixRow01(0, True)
'                oMat02.Clear
'                oMat02.FlushToDataSource
'                oMat02.LoadFromDataSource
'                Call PS_PP040_AddMatrixRow02(0, True)
'                oMat03.Clear
'                oMat03.FlushToDataSource
'                oMat03.LoadFromDataSource
                oForm01.Update
            End If
        End If
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//PSMT
        If oForm01.Items("OrdMgNum").Specific.VALUE = "" Then
            Call MDC_Com.MDC_GF_Message("작업지시 관리번호를 입력하지 않습니다.", "W")
            GoTo PS_PP040_OrderInfoLoad_Exit
        Else
            oForm01.Items("OrdNum").Specific.VALUE = oForm01.Items("OrdMgNum").Specific.VALUE
            oForm01.Items("OrdSub1").Specific.VALUE = "000"
            oForm01.Items("OrdSub2").Specific.VALUE = "00"
            '//매트릭스삭제
            oMat01.Clear
            oMat01.FlushToDataSource
            oMat01.LoadFromDataSource
            Call PS_PP040_AddMatrixRow01(0, True)
            oMat02.Clear
            oMat02.FlushToDataSource
            oMat02.LoadFromDataSource
            Call PS_PP040_AddMatrixRow02(0, True)
            oMat03.Clear
            oMat03.FlushToDataSource
            oMat03.LoadFromDataSource
            oForm01.Update
        End If
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then
        Call MDC_Com.MDC_GF_Message("외주은 입력할수 없습니다.", "W")
        GoTo PS_PP040_OrderInfoLoad_Exit
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then
        Call MDC_Com.MDC_GF_Message("실적은 입력할수 없습니다.", "W")
        GoTo PS_PP040_OrderInfoLoad_Exit
    End If
    Set RecordSet01 = Nothing
    Exit Sub
PS_PP040_OrderInfoLoad_Exit:
    Set RecordSet01 = Nothing
    Exit Sub
PS_PP040_OrderInfoLoad_Error:
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_PP040_OrderInfoLoad_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Function PS_PP040_FindValidateDocument(ByVal ObjectType As String) As Boolean
On Error GoTo PS_PP040_FindValidateDocument_Error
    PS_PP040_FindValidateDocument = True
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim Query02 As String
    Dim RecordSet02 As SAPbobsCOM.Recordset
    
    Dim i As Long
    Dim DocEntry As String
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE) '//원본문서
    
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Query01 = " SELECT DocEntry"
    Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry = "
    Query01 = Query01 & DocEntry
    If (oDocType01 = "작업일보등록(작지)") Then
        Query01 = Query01 & " AND U_DocType = '10'"
    ElseIf (oDocType01 = "작업일보등록(공정)") Then
        Query01 = Query01 & " AND U_DocType = '20'"
    End If
    RecordSet01.DoQuery Query01
    If (RecordSet01.RecordCount = 0) Then
        If (oDocType01 = "작업일보등록(작지)") Then
            Sbo_Application.SetStatusBarMessage "작업일보등록(공정)문서 이거나 존재하지 않는 문서입니다.", bmt_Short, True
        ElseIf (oDocType01 = "작업일보등록(공정)") Then
            Sbo_Application.SetStatusBarMessage "작업일보등록(작지)문서 이거나 존재하지 않는 문서입니다.", bmt_Short, True
        End If
        PS_PP040_FindValidateDocument = False
        Set RecordSet01 = Nothing
        Set RecordSet02 = Nothing
        Exit Function
    End If
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Exit Function
PS_PP040_FindValidateDocument_Error:
    Sbo_Application.SetStatusBarMessage Err.Number & Err.Description, bmt_Short, True
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_PP040_FindValidateDocument = False
End Function

Public Function PS_PP040_DirectionValidateDocument(ByVal DocEntry As String, ByVal DocEntryNext As String, ByVal Direction As String, ByVal ObjectType As String) As Boolean
On Error GoTo PS_PP040_DirectionValidateDocument_Error
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim Query02 As String
    Dim RecordSet02 As SAPbobsCOM.Recordset
    
    Dim i As Long
    Dim MaxDocEntry As String
    Dim MinDocEntry As String
    Dim DoNext As Boolean
    Dim IsFirst As Boolean '//시작유무
    DoNext = True
    IsFirst = True
    
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Do While (DoNext = True)
        If (IsFirst <> True) Then
            If (DocEntry = DocEntryNext) Then '//문서전체를 경유하고도 유효값을 찾지못했다면
                Sbo_Application.SetStatusBarMessage "유효한문서가 존재하지 않습니다", bmt_Short, True
                PS_PP040_DirectionValidateDocument = False
                Set RecordSet01 = Nothing
                Set RecordSet02 = Nothing
                Exit Function
            End If
        End If
        If (Direction = "Next") Then
            Query01 = " SELECT TOP 1 DocEntry"
            Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry > "
            Query01 = Query01 & DocEntryNext
            If (oDocType01 = "작업일보등록(작지)") Then
                Query01 = Query01 & " AND U_DocType = '10'"
            ElseIf (oDocType01 = "작업일보등록(공정)") Then
                Query01 = Query01 & " AND U_DocType = '20'"
            End If
            Query01 = Query01 & " ORDER BY DocEntry ASC"
        ElseIf (Direction = "Prev") Then
            Query01 = " SELECT TOP 1 DocEntry"
            Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry < "
            Query01 = Query01 & DocEntryNext
            If (oDocType01 = "작업일보등록(작지)") Then
                Query01 = Query01 & " AND U_DocType = '10'"
            ElseIf (oDocType01 = "작업일보등록(공정)") Then
                Query01 = Query01 & " AND U_DocType = '20'"
            End If
            Query01 = Query01 & " ORDER BY DocEntry DESC"
        End If
        RecordSet01.DoQuery Query01
        If (RecordSet01.Fields(0).VALUE = 0) Then '//해당문서가 마지막문서라면
            If (Direction = "Next") Then
                Query02 = " SELECT TOP 1 DocEntry FROM [" & ObjectType & "]"
                If (oDocType01 = "작업일보등록(작지)") Then
                    Query02 = Query02 & " WHERE U_DocType = '10'"
                ElseIf (oDocType01 = "작업일보등록(공정)") Then
                    Query02 = Query02 & " WHERE U_DocType = '20'"
                End If
                Query02 = Query02 & " ORDER BY DocEntry ASC"
            ElseIf (Direction = "Prev") Then
                Query02 = " SELECT TOP 1 DocEntry FROM [" & ObjectType & "]"
                If (oDocType01 = "작업일보등록(작지)") Then
                    Query02 = Query02 & " WHERE U_DocType = '10'"
                ElseIf (oDocType01 = "작업일보등록(공정)") Then
                    Query02 = Query02 & " WHERE U_DocType = '20'"
                End If
                Query02 = Query02 & " ORDER BY DocEntry DESC"
            End If
            RecordSet02.DoQuery Query02
            If (RecordSet02.RecordCount = 0) Then '//문서가 아예 존재하지 않는다면
                Sbo_Application.SetStatusBarMessage "유효한문서가 존재하지 않습니다", bmt_Short, True
                Set RecordSet01 = Nothing
                Set RecordSet02 = Nothing
                PS_PP040_DirectionValidateDocument = False
                Exit Function
            Else
                If (Direction = "Next") Then
                    DocEntryNext = Val(RecordSet02.Fields(0).VALUE) - 1
                    Query01 = " SELECT TOP 1 DocEntry"
                    Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry > "
                    Query01 = Query01 & DocEntryNext
                    If (oDocType01 = "작업일보등록(작지)") Then
                        Query01 = Query01 & " AND U_DocType = '10'"
                    ElseIf (oDocType01 = "작업일보등록(공정)") Then
                        Query01 = Query01 & " AND U_DocType = '20'"
                    End If
                    Query01 = Query01 & " ORDER BY DocEntry ASC"
                    RecordSet01.DoQuery Query01
                ElseIf (Direction = "Prev") Then
                    DocEntryNext = Val(RecordSet02.Fields(0).VALUE) + 1
                    Query01 = " SELECT TOP 1 DocNum"
                    Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry < "
                    Query01 = Query01 & DocEntryNext
                    If (oDocType01 = "작업일보등록(작지)") Then
                        Query01 = Query01 & " AND U_DocType = '10'"
                    ElseIf (oDocType01 = "작업일보등록(공정)") Then
                        Query01 = Query01 & " AND U_DocType = '20'"
                    End If
                    Query01 = Query01 & " ORDER BY DocEntry DESC"
                    RecordSet01.DoQuery Query01
                End If
            End If
        End If
        If (oDocType01 = "작업일보등록(작지)") Then
            DoNext = False
            If (Direction = "Next") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) - 1
            ElseIf (Direction = "Prev") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) + 1
            End If
        ElseIf (oDocType01 = "작업일보등록(공정)") Then
            DoNext = False
            If (Direction = "Next") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) - 1
            ElseIf (Direction = "Prev") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) + 1
            End If
        End If
        IsFirst = False
    Loop
    If (DocEntry = DocEntryNext) Then '//다음문서가 유효하다면 그냥 넘어가고
        Call PS_PP040_FormItemEnabled '//UDO방식
    Else '//다음문서가 유효하지 않다면
        oForm01.Freeze True
        oForm01.Mode = fm_FIND_MODE
        Call PS_PP040_FormItemEnabled '//UDO방식
        If oForm01.Items("DocEntry").Enabled = True Then '//문서번호 필드가 입력이 가능하다면
            If (Direction = "Next") Then
                oForm01.Items("DocEntry").Specific.VALUE = Val(DocEntryNext + 1)
            ElseIf (Direction = "Prev") Then
                oForm01.Items("DocEntry").Specific.VALUE = Val(DocEntryNext - 1)
            End If
            oForm01.Items("1").Click ct_Regular
        End If
        oForm01.Freeze False
        PS_PP040_DirectionValidateDocument = False
        Set RecordSet01 = Nothing
        Set RecordSet02 = Nothing
        Exit Function
    End If
    PS_PP040_DirectionValidateDocument = True
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Exit Function
PS_PP040_DirectionValidateDocument_Error:
    Sbo_Application.SetStatusBarMessage Err.Number & Err.Description, bmt_Short, True
    PS_PP040_DirectionValidateDocument = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
End Function
