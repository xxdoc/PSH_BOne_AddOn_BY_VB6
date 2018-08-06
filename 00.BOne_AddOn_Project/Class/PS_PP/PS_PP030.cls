VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP030.cls
'//  Module         : 생산관리 > 작업지시등록
'//  Desc           : 기본
'//  FormType       : PS_PP030
'//  Create Date    : 2010.08.20
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Public oMat02          As SAPbouiCOM.Matrix
Public oMat03          As SAPbouiCOM.Matrix
Public oDS_PS_USERDS01 As SAPbouiCOM.DBDataSource '등록헤더
Public oDS_PS_PP030H   As SAPbouiCOM.DBDataSource '등록헤더
Public oDS_PS_PP030L   As SAPbouiCOM.DBDataSource '등록라인
Public oDS_PS_PP030M   As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMat01Row01    As Long
Private oMat02Row02    As Long
Private oMat03Row03    As Long

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    BatchNum As String
    Quantity As Long
    OPORNo   As Long
    POR1No   As Long
    Check    As Boolean
    OPDNNo   As Long
    PDN1No   As Long
End Type
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

Private oDocEntry01  As String
Private oSCardCod01  As String
Private oMark        As String
Private oFormMode01  As SAPbouiCOM.BoFormMode
Private oHasMatrix01 As Boolean

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP030.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP030_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP030_CreateItems
    Call PS_PP030_ComboBox_Setting
    Call PS_PP030_CF_ChooseFromList
    Call PS_PP030_EnableMenus
    Call PS_PP030_SetDocument(oFromDocEntry01)
    Call PS_PP030_FormResize
    Call Initialization
    
    oForm01.EnableMenu ("1283"), False '삭제
    oForm01.EnableMenu ("1287"), False '복제
    oForm01.EnableMenu ("1286"), True '닫기
    oForm01.EnableMenu ("1284"), True '취소
    oForm01.EnableMenu ("1293"), True '행삭제
    oForm01.EnableMenu ("1299"), False '행닫기
    
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

Sub Initialization()
On Error GoTo Initialization_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("SBPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
       
    '//아이디별 사번 세팅
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
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

        ' 공정금액 합계 추가 S
        Dim i&
        Dim Total As Currency

        Call Raise_EVENT_MATRIX_LOAD(FormUID, pval, BubbleEvent)
            
        For i = 0 To oMat03.VisualRowCount - 1
            Total = Total + oMat03.Columns("CpPrice").Cells(i + 1).Specific.VALUE
        Next i
        oForm01.Items("Total").Specific.VALUE = Total
        ' 공정금액 합계 추가 E
            
            
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
    Dim i As Long
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
                If oForm01.Mode = fm_OK_MODE Then
                    If (PS_PP030_Validate("취소") = False) Then
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
            
                If oForm01.Mode = fm_OK_MODE Then
                    If (PS_PP030_Validate("닫기") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If Sbo_Application.MessageBox("문서를 닫기(종료) 처리하겠습니까?", "1", "예", "아니오") <> "1" Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                Else
                    MDC_Com.MDC_GF_Message "현재 모드에서는 닫기(종료) 처리할 수 없습니다.", "W"
                    BubbleEvent = False
                    Exit Sub
                End If
                
            
            Case "1292": '행추가
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1292": '행추가
                If oLastItemUID01 = "Mat03" Then
                    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티인경우만
                        '//행추가가 가능검사
                        If (PS_PP030_Validate("행추가03") = False) Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        Call oMat03.AddRow(1, oMat03Row03 - 1)
                        For i = 1 To oMat03.VisualRowCount
                            oMat03.Columns("LineNum").Cells(i).Specific.VALUE = i
                            oMat03.Columns("Sequence").Cells(i).Specific.VALUE = i
                            
                            If oMat03Row03 = i Then '//새로추가된 행의 값 설정
                                Call oMat03.Columns("ReWorkYN").Cells(i).Specific.Select(0, psk_Index) '//PK/탈지일때 재작업여부 예
                                
                                Call oMat03.Columns("ResultYN").Cells(i).Specific.Select(1, psk_Index) '//실적여부 아니오
                                Call oMat03.Columns("ReportYN").Cells(i).Specific.Select(0, psk_Index) '//일보여부 예
                            End If
                        Next i
                        oMat03.FlushToDataSource
                        oMat03.LoadFromDataSource
                    Else
                        Call MDC_Com.MDC_GF_Message("멀티인 경우만 행추가 가능합니다.", "W")
                    End If
                End If
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_PP030_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_PP030_FormItemEnabled '//UDO방식
                Call PS_PP030_AddMatrixRow01(0, True) '//UDO방식
                Call PS_PP030_AddMatrixRow02(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_PP030_FormItemEnabled
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
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim oOrdGbn01 As String
    Dim oProcType01 As String
    
    Dim li_Cnt As Integer
    Dim li_LineId As Integer
    
    Dim lChildForm As Variant '팝업창 호출 용 변수(2012.04.12 송명규)
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_PP030_MTX01 '//조회
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                
                If PS_PP030_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oSCardCod01 = Trim(oForm01.Items("SCardCod").Specific.VALUE)
                
                oFormMode01 = oForm01.Mode
                '//멀티게이지 일괄생성기능구현 , 엔드베이링 추가 - 류영조
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then
                    If PS_PP030_AutoCreateMultiGage = False Then
                        Call PS_PP030_AddMatrixRow01(oMat02.VisualRowCount)
                        Call PS_PP030_AddMatrixRow02(oMat03.VisualRowCount)
                        BubbleEvent = False
                        Exit Sub
                    End If
                    oForm01.Mode = fm_FIND_MODE
                    Call PS_PP030_FormItemEnabled
                    Sbo_Application.ActivateMenuItem ("1282")
                    BubbleEvent = False
                    Exit Sub
                Else
                '//멀티게이지를 제외한 나머지 경우는 자동으로 입력
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_PP030_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                If oMat02.VisualRowCount = 0 Then
                    Call oMat02.Clear
                    Call oMat02.AddRow
                    Call oMat02.FlushToDataSource
                    Call oMat02.LoadFromDataSource
                End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        '표준공수조회, 품목별공수조회 버튼 클릭 시(2012.04.12 송명규)
        If pval.ItemUID = "btnWkSrch" Then '표준공수조회 버튼 클릭
        
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then
            
                Call MDC_Com.MDC_GF_Message("작업구분을 선택하십시오.", "W")
                Exit Sub
                
            Else
                
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then
            
                    Set lChildForm = New PS_PP033
                    Call lChildForm.LoadForm(Me)
                
                Else
                
                    Call MDC_Com.MDC_GF_Message("작업구분이 [제품_기계공구] 또는 [제품_몰드] 일 경우에만 사용이 가능합니다.", "W")
                    Exit Sub
                
                End If
                
            End If
        
        ElseIf pval.ItemUID = "btnItmSrch" Then '품목별공수조회 버튼 클릭
        
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then
            
                Call MDC_Com.MDC_GF_Message("작업구분을 선택하십시오.", "W")
                Exit Sub
                
            Else
                
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then
            
                    Set lChildForm = New PS_PP031
                    Call lChildForm.LoadForm
                
                Else
                
                    Call MDC_Com.MDC_GF_Message("작업구분이 [제품_기계공구] 또는 [제품_몰드] 일 경우에만 사용이 가능합니다.", "W")
                    Exit Sub
                
                End If
                
            End If
        
        End If
        '표준공수조회, 품목별공수조회 버튼 클릭 시(2012.04.12 송명규)
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    oOrdGbn01 = MDC_PS_Common.GetValue("SELECT U_OrdGbn FROM [@PS_PP030H] WHERE DocEntry = '" & oDocEntry01 & "'")
                    If (oOrdGbn01 = "105" Or oOrdGbn01 = "106") Then '//기계공구, 몰드
                        Query01 = "SELECT U_ProcType, DocEntry, LineId FROM [@PS_PP030L] WHERE DocEntry = '" & oDocEntry01 & "'"
                        RecordSet01.DoQuery Query01
                        For i = 0 To RecordSet01.RecordCount - 1
                            If (RecordSet01.Fields(0).VALUE = "10") Then
                                Call PS_PP030_PurchaseRequest(RecordSet01.Fields(1).VALUE, RecordSet01.Fields(2).VALUE)
                            End If
                            RecordSet01.MoveNext
                        Next
                    End If
                    Set RecordSet01 = Nothing
                    Call PS_PP030_FormItemEnabled
                    Call PS_PP030_AddMatrixRow01(0, True) '//UDO방식일때
                    Call PS_PP030_AddMatrixRow02(0, True) '//UDO방식일때
                    oForm01.Items("SCardCod").Specific.VALUE = oSCardCod01
                    oForm01.Items("Total").Specific.VALUE = 0 '공정금액 합계 초기화
                    
                    oForm01.Items("Button01").Click ct_Regular
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If pval.ActionSuccess = True Then
                    oForm01.Items("Total").Specific.VALUE = 0 '공정금액 합계 초기화
                End If
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    If (oFormMode01 = fm_UPDATE_MODE) Then
                        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                        oOrdGbn01 = MDC_PS_Common.GetValue("SELECT U_OrdGbn FROM [@PS_PP030H] WHERE DocEntry = '" & oDocEntry01 & "'")
                        If (oOrdGbn01 = "105" Or oOrdGbn01 = "106") Then '//기계공구, 몰드
                            Query01 = "SELECT U_ProcType, DocEntry, LineId FROM [@PS_PP030L] WHERE DocEntry = '" & oDocEntry01 & "'"
                            RecordSet01.DoQuery Query01
                            For i = 0 To RecordSet01.RecordCount - 1
                                If (RecordSet01.Fields(0).VALUE = "10") Then
                                    Call PS_PP030_PurchaseRequest(RecordSet01.Fields(1).VALUE, RecordSet01.Fields(2).VALUE)
                                End If
                                RecordSet01.MoveNext
                            Next
                        End If
                        If oOrdGbn01 = "104" Then
                            Query01 = "Update [@PS_PP030M] set VisOrder = U_Sequence - 1, LineId = U_Sequence, U_LineId = U_Sequence WHERE LineId <> U_Sequence And DocEntry = '" & oDocEntry01 & "'"
                            RecordSet01.DoQuery Query01
                            
                            Query01 = "SELECT Count(*), Min(LineId) FROM [@PS_PP030M] WHERE DocEntry = '" & oDocEntry01 & "' and U_CpCode = 'CP50107'"
                            RecordSet01.DoQuery Query01
                            
                            li_Cnt = RecordSet01.Fields(0).VALUE
                            li_LineId = RecordSet01.Fields(1).VALUE
                            
                            If (li_Cnt > 1) Then
                                Query01 = "Update [@PS_PP030M] set U_ResultYN = 'N' WHERE DocEntry = '" & oDocEntry01 & "' and LineId = '" & li_LineId & "'"
                                RecordSet01.DoQuery Query01
                            Else
                                Query01 = "Update [@PS_PP030M] set U_ResultYN = 'Y' WHERE DocEntry = '" & oDocEntry01 & "' and LineId = '" & li_LineId & "'"
                                RecordSet01.DoQuery Query01
                            End If
                            
                            
                        End If
                        oFormMode01 = fm_OK_MODE
                        oForm01.Mode = fm_FIND_MODE
                        Call PS_PP030_FormItemEnabled
                        oForm01.Items("DocEntry").Specific.VALUE = oDocEntry01
                        oForm01.Items("1").Click ct_Regular
                    End If
                    Call PS_PP030_FormItemEnabled
                End If
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//사용자값활성
        If oForm01.Mode = fm_FIND_MODE Then
            '//찾기모드는 입력가능하도록
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성 입력가능하도록
        Else
            Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성 입력은 안됨
        End If
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat02", "CntcCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat03", "CpBCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat03", "CpCode") '//사용자값활성
        If pval.ItemUID = "Mat02" Then
            If pval.ColUID = "ItemCode" Then
                If oMat02.Columns("InputGbn").Cells(pval.Row).Specific.Selected Is Nothing Then
                    Call MDC_Com.MDC_GF_Message("투입구분을 선택하세요", "W")
                    oMat02.Columns("InputGbn").Cells(pval.Row).Click ct_Regular
                    BubbleEvent = False
                    Exit Sub
                Else
                    If (PS_PP030_Validate("수정02") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    '[2011.2.14] 추가 Begin------------------------------------------------------------------------------------------------------------
                    '107010002(END BEARING #44),107010004(END BEARING #2) 일경우에는 투입자재를 직접 입력한다.
                    If Trim(oForm01.Items("ItemCode").Specific.VALUE) = "107010002" Or Trim(oForm01.Items("ItemCode").Specific.VALUE) = "107010004" Then
'                        BubbleEvent = False
                        Exit Sub
                    End If
                    '[2011.2.14] 추가 End--------------------------------------------------------------------------------------------------------------
                    
                    Dim OrdGbn As String
                    Dim InputGbn As String
                    OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.Selected.VALUE)
                    InputGbn = Trim(oMat02.Columns("InputGbn").Cells(pval.Row).Specific.Selected.VALUE)
                    Dim ChildForm01 As Variant
                    Set ChildForm01 = New PS_SM021
                    Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, OrdGbn, InputGbn, Trim(oDS_PS_PP030H.GetValue("U_BPLId", 0)))
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        ElseIf pval.ItemUID = "Mat03" Then
            If pval.ColUID = "FailCode" Then
                If oMat03.Columns("FailCode").Cells(pval.Row).Specific.VALUE = "" Then
                'If oForm01.Items("FailCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
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
            If (pval.ItemUID = "Mat02") Then
                If (PS_PP030_Validate("수정02") = False) Then
                    Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, Trim(oDS_PS_PP030L.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                Else
                    If (pval.ColUID = "특정컬럼") Then
                    Else
                        Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    End If
                End If
            ElseIf (pval.ItemUID = "Mat03") Then
                If (PS_PP030_Validate("수정03") = False) Then
                    Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, Trim(oDS_PS_PP030M.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                Else
                    If (pval.ColUID = "WorkGbn") Then
                        If oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = "10" Then
                           Call oDS_PS_PP030M.setValue("U_CpPrice", pval.Row - 1, MDC_PS_Common.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" & oMat03.Columns("CpCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1) * oMat03.Columns("StdHour").Cells(pval.Row).Specific.VALUE)
                        ElseIf oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = "20" Then
                            Call oDS_PS_PP030M.setValue("U_CpPrice", pval.Row - 1, MDC_PS_Common.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" & oMat03.Columns("CpCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1) * oMat03.Columns("StdHour").Cells(pval.Row).Specific.VALUE)
                        End If
                        
                        Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        
                    ElseIf (pval.ColUID = "특정컬럼") Then
                        Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        If oMat03.RowCount = pval.Row And Trim(oDS_PS_PP030M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                            PS_PP030_AddMatrixRow02 (pval.Row)
                        End If
                    Else
                        Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    End If
                End If
            Else
                If (pval.ItemUID = "OrdGbn") Then
                    Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    If oHasMatrix01 = True Then
                        If oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "104" Then '//작업구분이 멀티일때만
                            oForm01.Items("BasicGub").Enabled = True
                            oForm01.Items("MulGbn1").Enabled = True
                            oForm01.Items("MulGbn2").Enabled = True
                            oForm01.Items("MulGbn3").Enabled = True
                        Else
                            oForm01.Items("BasicGub").Enabled = False
                            oForm01.Items("MulGbn1").Enabled = False
                            oForm01.Items("MulGbn2").Enabled = False
                            oForm01.Items("MulGbn3").Enabled = False
                        End If
                        If (oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "107") Then '//엔드베어링일때
'                            oMat02.Columns("InputGbn").Editable = True
'                            Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("InputGbn"), "PS_PP030", "Mat02", "InputGbn2")
                            oMat02.Columns("InputGbn").Editable = True
                        Else
'                            oMat02.Columns("InputGbn").Editable = True
'                            Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("InputGbn"), "PS_PP030", "Mat02", "InputGbn")
                            oMat02.Columns("InputGbn").Editable = False
                        End If
                        If (oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "105" Or oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "106") Then
                            oMat02.Columns("Weight").Editable = True
                        Else
                            oMat02.Columns("Weight").Editable = False
                        End If
                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then  '//멀티,엔드베어링이면
'                            oMat03.Columns("CpBCode").Editable = False
'                            oMat03.Columns("CpCode").Editable = False
'                            oMat03.Columns("ResultYN").Editable = False
'                            oMat03.Columns("ReportYN").Editable = False
                        Else
'                            oMat03.Columns("CpBCode").Editable = True
'                            oMat03.Columns("CpCode").Editable = True
'                            oMat03.Columns("ResultYN").Editable = True
'                            oMat03.Columns("ReportYN").Editable = True
                        End If
                        oMat02.Clear
                        oMat02.FlushToDataSource
                        oMat02.LoadFromDataSource
                        Call PS_PP030_AddMatrixRow01(0, True)
                        oMat03.Clear
                        oMat03.FlushToDataSource
                        oMat03.LoadFromDataSource
                        Call PS_PP030_AddMatrixRow02(0, True)
                        Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
                        oForm01.Items("DocDate").Specific.VALUE = ""
                        oForm01.Items("DueDate").Specific.VALUE = ""
                        oForm01.Items("ItemCode").Specific.VALUE = "" '//공정리스트 매트릭스를 초기화 시킨다.
                        oForm01.Items("ItemName").Specific.VALUE = ""
                        oForm01.Items("JakMyung").Specific.VALUE = ""
                        oForm01.Items("JakSize").Specific.VALUE = ""
                        oForm01.Items("JakUnit").Specific.VALUE = ""
                        oForm01.Items("CntcCode").Specific.VALUE = ""
                        oForm01.Items("CntcName").Specific.VALUE = ""
                        oForm01.Items("OrdMgNum").Specific.VALUE = ""
                        oForm01.Items("ReqWt").Specific.VALUE = 0
                        oForm01.Items("SelWt").Specific.VALUE = 0
                    Else
                        '//그냥선택시
                        If oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "104" Or oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "107" Then '//작업구분이 멀티일때만
                            If oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "104" Then '//작업구분이 멀티일때만
                                oForm01.Items("BasicGub").Enabled = True
                                oForm01.Items("MulGbn1").Enabled = True
                                oForm01.Items("MulGbn2").Enabled = True
                                oForm01.Items("MulGbn3").Enabled = True
                            Else
                                oForm01.Items("BasicGub").Enabled = False
                                oForm01.Items("MulGbn1").Enabled = False
                                oForm01.Items("MulGbn2").Enabled = False
                                oForm01.Items("MulGbn3").Enabled = False
                            End If
                            If (oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "107") Then '//엔드베어링일때
'                                oMat02.Columns("InputGbn").Editable = True
'                                Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("InputGbn"), "PS_PP030", "Mat02", "InputGbn2")
                                oMat02.Columns("InputGbn").Editable = True
                            Else
'                                oMat02.Columns("InputGbn").Editable = True
'                                Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("InputGbn"), "PS_PP030", "Mat02", "InputGbn")
                                oMat02.Columns("InputGbn").Editable = False
                            End If
                            If (oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "105" Or oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "106") Then
                                oMat02.Columns("Weight").Editable = True
                            Else
                                oMat02.Columns("Weight").Editable = False
                            End If
                            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then  '//멀티,엔드베어링이면
'                                oMat03.Columns("CpBCode").Editable = False
'                                oMat03.Columns("CpCode").Editable = False
'                                oMat03.Columns("ResultYN").Editable = False
'                                oMat03.Columns("ReportYN").Editable = False
                            Else
'                                oMat03.Columns("CpBCode").Editable = True
'                                oMat03.Columns("CpCode").Editable = True
'                                oMat03.Columns("ResultYN").Editable = True
'                                oMat03.Columns("ReportYN").Editable = True
                            End If
                            oMat02.Clear
                            oMat02.FlushToDataSource
                            oMat02.LoadFromDataSource
                            Call PS_PP030_AddMatrixRow01(0, True)
                            oMat03.Clear
                            oMat03.FlushToDataSource
                            oMat03.LoadFromDataSource
                            Call PS_PP030_AddMatrixRow02(0, True)
                            'Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
                            Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
                            oForm01.Items("DocDate").Specific.VALUE = Format(Date, "YYYYMMDD")
                            oForm01.Items("DueDate").Specific.VALUE = Format(Date, "YYYYMMDD")
                            oForm01.Items("ItemCode").Specific.VALUE = "" '//공정리스트 매트릭스를 초기화 시킨다.
                            oForm01.Items("ItemName").Specific.VALUE = ""
                            oForm01.Items("JakMyung").Specific.VALUE = ""
                            oForm01.Items("JakSize").Specific.VALUE = ""
                            oForm01.Items("JakUnit").Specific.VALUE = ""
'                            oForm01.Items("CntcCode").Specific.Value = ""
'                            oForm01.Items("CntcName").Specific.Value = ""
                            oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
                            oForm01.Items("OrdMgNum").Specific.VALUE = Format(Date, "YYYYMMDD")
                            oForm01.Items("ReqWt").Specific.VALUE = 0
                            oForm01.Items("SelWt").Specific.VALUE = 0
                        ElseIf oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "선택" Then
                            '//아무행위도 하지 않음
                        Else '//멀티랑 엔드베어링일때
                            Call MDC_Com.MDC_GF_Message("멀티,엔드베어링작업만 선택할수 있습니다.", "W")
                            Call oForm01.Items("OrdGbn").Specific.Select(0, psk_Index)
                        End If
                    End If
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_PP030H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
                End If
            End If
            oMat02.LoadFromDataSource
            oMat03.LoadFromDataSource
            oMat02.AutoResizeColumns
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
    Dim True_False$
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Opt01" Then
            oForm01.Freeze True
            True_False = oMat02.Columns("Weight").Editable
            oForm01.Settings.MatrixUID = "Mat01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat02.Columns("Weight").Editable = True_False
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Opt02" Then
            oForm01.Freeze True
            True_False = oMat02.Columns("Weight").Editable
            oForm01.Settings.MatrixUID = "Mat02"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat02.Columns("Weight").Editable = True_False
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Opt03" Then
            oForm01.Freeze True
            True_False = oMat02.Columns("Weight").Editable
            oForm01.Settings.MatrixUID = "Mat03"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat02.Columns("Weight").Editable = True_False
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
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
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.Row = 0 Then
                
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource
                
            ElseIf pval.Row > 0 Then
                oHasMatrix01 = True
                oForm01.Freeze True
                oMat02.Clear
                oMat02.FlushToDataSource
                oMat02.LoadFromDataSource
                Call PS_PP030_AddMatrixRow01(0, True)
                '//아이템활성화하여 Validate 발생
                oForm01.Items("OrdGbn").Enabled = True
                oForm01.Items("BPLId").Enabled = True
                oForm01.Items("ItemCode").Enabled = True
                oForm01.Items("OrdMgNum").Enabled = True
                If oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Selected Is Nothing Then
                    Call oForm01.Items("OrdGbn").Specific.Select(0, psk_Index)
                    Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
                    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
                    oForm01.Items("DocDate").Specific.VALUE = ""
                    oForm01.Items("DueDate").Specific.VALUE = ""
                    oForm01.Items("BaseType").Specific.VALUE = ""
                    oForm01.Items("BaseNum").Specific.VALUE = ""
                    oForm01.Items("OrdMgNum").Specific.VALUE = ""
                    oForm01.Items("OrdNum").Specific.VALUE = ""
                    oForm01.Items("OrdSub1").Specific.VALUE = ""
                    oForm01.Items("OrdSub2").Specific.VALUE = ""
                    oForm01.Items("ItemCode").Specific.VALUE = ""
                    oForm01.Items("JakMyung").Specific.VALUE = ""
                    oForm01.Items("JakSize").Specific.VALUE = ""
                    oForm01.Items("JakUnit").Specific.VALUE = ""
                    oForm01.Items("ReqWt").Specific.VALUE = 0
                    oForm01.Items("SelWt").Specific.VALUE = 0
                    oForm01.Items("SjNum").Specific.VALUE = ""
                    oForm01.Items("SjLine").Specific.VALUE = ""
                    oForm01.Items("LotNo").Specific.VALUE = ""
                    oForm01.Items("SjPrice").Specific.VALUE = 0
                Else
                    Call oForm01.Items("OrdGbn").Specific.Select(oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Selected.VALUE, psk_ByValue)
                    Call oForm01.Items("BPLId").Specific.Select(oMat01.Columns("BPLId").Cells(pval.Row).Specific.Selected.VALUE, psk_ByValue)
                    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
                    oForm01.Items("ItemCode").Specific.VALUE = oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE
                    oForm01.Items("DocDate").Specific.VALUE = Format(Date, "YYYYMMDD")
                    oForm01.Items("DueDate").Specific.VALUE = oMat01.Columns("ReqDate").Cells(pval.Row).Specific.VALUE
                    oForm01.Items("BaseType").Specific.VALUE = oMat01.Columns("BaseType").Cells(pval.Row).Specific.VALUE
                    oForm01.Items("BaseNum").Specific.VALUE = oMat01.Columns("BaseNum").Cells(pval.Row).Specific.VALUE
                    oForm01.Items("OrdMgNum").Specific.VALUE = ""
                    '//기계공구,몰드일경우 작번이 생성되어있음
                    If oMat01.Columns("BaseType").Cells(pval.Row).Specific.VALUE = "작번요청" Then
                    '//If oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Selected.Value = "105" Or oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Selected.Value = "106" Then
'                        oForm01.Items("OrdNum").Specific.Value = oMat01.Columns("OrdNum").Cells(pval.Row).Specific.Value
'                        oForm01.Items("OrdSub1").Specific.Value = oMat01.Columns("OrdSub1").Cells(pval.Row).Specific.Value
'                        oForm01.Items("OrdSub2").Specific.Value = oMat01.Columns("OrdSub2").Cells(pval.Row).Specific.Value
                        Call oDS_PS_PP030H.setValue("U_OrdNum", 0, oMat01.Columns("OrdNum").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP030H.setValue("U_OrdSub1", 0, oMat01.Columns("OrdSub1").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP030H.setValue("U_OrdSub2", 0, oMat01.Columns("OrdSub2").Cells(pval.Row).Specific.VALUE)
                        
                        If oMat01.Columns("OrdSub1").Cells(pval.Row).Specific.VALUE = "00" Then
                            '// 메인작번일경우 작명과 규격에 품목명, 규격으로
                            Call oDS_PS_PP030H.setValue("U_JakMyung", 0, MDC_PS_Common.GetValue("SELECT FrgnName FROM [OITM] WHERE ItemCode = '" & oMat01.Columns("OrdNum").Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                            Call oDS_PS_PP030H.setValue("U_JakSize", 0, MDC_PS_Common.GetValue("SELECT U_Size FROM [OITM] WHERE ItemCode = '" & oMat01.Columns("OrdNum").Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                            Call oDS_PS_PP030H.setValue("U_JakUnit", 0, MDC_PS_Common.GetValue("SELECT salUnitMsr FROM [OITM] WHERE ItemCode = '" & oMat01.Columns("OrdNum").Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                        Else
                            '// 서브작번일경우 작명과 규격에 서브작번명, 규격으로
                            oForm01.Items("JakMyung").Specific.VALUE = oMat01.Columns("JakMyung").Cells(pval.Row).Specific.VALUE
                            oForm01.Items("JakSize").Specific.VALUE = oMat01.Columns("JakSize").Cells(pval.Row).Specific.VALUE
                            oForm01.Items("JakUnit").Specific.VALUE = oMat01.Columns("JakUnit").Cells(pval.Row).Specific.VALUE
                        End If
                            
                        
                    ElseIf oMat01.Columns("BaseType").Cells(pval.Row).Specific.VALUE = "생산요청" Then
                    '//Else '//생산요청번호
                        oForm01.Items("OrdMgNum").Specific.VALUE = Format(Now, "yyyyMMdd")
                        oForm01.Items("OrdNum").Specific.VALUE = Format(Now, "yyyyMMdd") & MDC_PS_Common.GetValue("EXEC PS_PP030_01 '" & oForm01.Items("OrdNum").Specific.VALUE & "'")
                        oForm01.Items("OrdSub1").Specific.VALUE = "00"
                        oForm01.Items("OrdSub2").Specific.VALUE = "000"
                        
                        oForm01.Items("JakMyung").Specific.VALUE = oMat01.Columns("JakMyung").Cells(pval.Row).Specific.VALUE
                        oForm01.Items("JakSize").Specific.VALUE = oMat01.Columns("JakSize").Cells(pval.Row).Specific.VALUE
                        oForm01.Items("JakUnit").Specific.VALUE = oMat01.Columns("JakUnit").Cells(pval.Row).Specific.VALUE
                    End If
                    
                    oForm01.Items("ReqWt").Specific.VALUE = oMat01.Columns("RemainWt").Cells(pval.Row).Specific.VALUE
                    oForm01.Items("SelWt").Specific.VALUE = oMat01.Columns("RemainWt").Cells(pval.Row).Specific.VALUE
                    oForm01.Items("SjNum").Specific.VALUE = oMat01.Columns("ORDRNum").Cells(pval.Row).Specific.VALUE
                    oForm01.Items("SjLine").Specific.VALUE = oMat01.Columns("RDR1Num").Cells(pval.Row).Specific.VALUE
                    oForm01.Items("LotNo").Specific.VALUE = MDC_PS_Common.GetValue("SELECT U_LotNo FROM [ORDR] WHERE DocEntry = '" & oMat01.Columns("ORDRNum").Cells(pval.Row).Specific.VALUE & "'", 0, 1)
                    oForm01.Items("SjPrice").Specific.VALUE = MDC_PS_Common.GetValue("SELECT LineTotal FROM [RDR1] WHERE DocEntry = '" & oMat01.Columns("ORDRNum").Cells(pval.Row).Specific.VALUE & "' AND LineNum = '" & oMat01.Columns("RDR1Num").Cells(pval.Row).Specific.VALUE & "'", 0, 1)
                End If
                oForm01.Items("OrdGbn").Enabled = False
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("ItemCode").Enabled = False
                If oMat01.Columns("BaseType").Cells(pval.Row).Specific.VALUE = "작번요청" Then
                    oForm01.Items("OrdMgNum").Enabled = False
                ElseIf oMat01.Columns("BaseType").Cells(pval.Row).Specific.VALUE = "생산요청" Then
                    oForm01.Items("OrdMgNum").Enabled = True
                End If
                oForm01.Freeze False
                oHasMatrix01 = False
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
    Dim i As Long
    Dim Exist As Boolean
    Dim sQry As String
    Dim TotalAmt As Long
    Dim ReqCod
    
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat02") Then
                If (PS_PP030_Validate("수정02") = False) Then
                    Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, Trim(oDS_PS_PP030L.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                Else
                    If (pval.ColUID = "ItemCode") Then
                        '//기타작업
                        Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        If oMat02.RowCount = pval.Row And Trim(oDS_PS_PP030L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                            PS_PP030_AddMatrixRow01 (pval.Row)
                        End If
                        '[2011.2.14] 추가 Begin------------------------------------------------------------------------------------------------------------
                        '107010002(END BEARING #44),107010004(END BEARING #2) 일경우에는 투입자재를 직접 입력한다.
                        If Trim(oForm01.Items("ItemCode").Specific.VALUE) = "107010002" Or Trim(oForm01.Items("ItemCode").Specific.VALUE) = "107010004" Then
                            oMat02.Columns("BatchNum").Editable = True
                        End If
                        '[2011.2.14] 추가 End--------------------------------------------------------------------------------------------------------------
                    '[2011.2.14] 추가 Begin------------------------------------------------------------------------------------------------------------
                    ElseIf pval.ColUID = "BatchNum" Then
                        Dim zQry As String
                        Dim oRecordset02 As SAPbobsCOM.Recordset
                        Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
                        oMat02.FlushToDataSource
                        
                        zQry = "EXEC [PS_PP030_06] '" & oMat02.Columns("ItemCode").Cells(pval.Row).Specific.VALUE & "', '" & oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"
                        oRecordset02.DoQuery zQry
                        
'                        oDS_PS_PP030L.setValue "U_ItemCode", oRow - 1, oRecordSet02.Fields(0).VALUE
                        oDS_PS_PP030L.setValue "U_ItemName", pval.Row - 1, oRecordset02.Fields("ItemName").VALUE
                        oDS_PS_PP030L.setValue "U_ItemGpCd", pval.Row - 1, oRecordset02.Fields("ItmsGrpCod").VALUE
                        oDS_PS_PP030L.setValue "U_Unit", pval.Row - 1, oRecordset02.Fields("InvntryUom").VALUE
                        oDS_PS_PP030L.setValue "U_Weight", pval.Row - 1, oRecordset02.Fields("Quantity").VALUE
                        oMat02.SetLineData pval.Row
                        
                        Set oRecordset02 = Nothing
                    '[2011.2.14] 추가 End--------------------------------------------------------------------------------------------------------------
                        
                    ElseIf pval.ColUID = "Weight" Then
                        If oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE < 0 Then
                            Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                        Else
                            Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        End If
                    ElseIf (pval.ColUID = "CntcCode") Then
                        Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP030L.setValue("U_CntcName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                    '//비고란 추가 안되는 것 수정 - 류영조
'                    ElseIf (pval.ColUID = "Comments") Then
'                        If Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "104" Or Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "107" Then
'                            Dim TotalQty As Double
'                            For i = 1 To oDS_PS_PP030L.Size - 1
'                                TotalQty = TotalQty + oDS_PS_PP030L.GetValue("U_Weight", i - 1)
'                            Next
'                            Call oDS_PS_PP030H.setValue("U_SelWt", 0, TotalQty)
'                        End If
                    ElseIf (pval.ColUID = "Comments") Then
                        If Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "104" Or Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "107" Then
                            '류영조
                            Dim TotalQty As Double
                            Dim useMkg As Currency
                            
    
                            If Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "107" Then
                                sQry = "Select IsNull(U_useMkg, 0) From [OITM] Where ItemCode = '" & Trim(oDS_PS_PP030H.GetValue("U_ItemCode", 0)) & "'"
                                oRecordSet01.DoQuery sQry
                                useMkg = oRecordSet01.Fields(0).VALUE / 1000
                                
                                For i = 1 To oDS_PS_PP030L.Size - 1
                                    TotalQty = TotalQty + oDS_PS_PP030L.GetValue("U_Weight", i - 1)
                                Next
                                If useMkg = 0 Then
                                    Call oDS_PS_PP030H.setValue("U_SelWt", 0, Round(TotalQty, 0))
                                    Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                                Else
                                    Call oDS_PS_PP030H.setValue("U_SelWt", 0, Round(TotalQty / useMkg, 0))
                                    Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                                End If
                            Else
                                For i = 1 To oDS_PS_PP030L.Size - 1
                                    TotalQty = TotalQty + oDS_PS_PP030L.GetValue("U_Weight", i - 1)
                                Next
                                Call oDS_PS_PP030H.setValue("U_SelWt", 0, TotalQty)
                                Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            End If
                            
                            Set oRecordSet01 = Nothing
                        End If
                        Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Else
                        Call oDS_PS_PP030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    End If
                End If
            ElseIf (pval.ItemUID = "Mat03") Then
                
                If pval.ColUID = "StdHour" Or pval.ColUID = "ReDate" Then
                    oMat03.FlushToDataSource '표준공수와 완료요구일은 수정이 가능해야 하므로 Flush 를 함
                    
                    If pval.ColUID = "StdHour" Then '표준공수 등록 시
                        '공정단가 계산_S
                        If oMat03.Columns("WorkGbn").Cells(pval.Row).Specific.VALUE = "10" Then
                            Call oDS_PS_PP030M.setValue("U_CpPrice", pval.Row - 1, MDC_PS_Common.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" & oMat03.Columns("CpCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1) * oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        ElseIf oMat03.Columns("WorkGbn").Cells(pval.Row).Specific.VALUE = "20" Then
                            Call oDS_PS_PP030M.setValue("U_CpPrice", pval.Row - 1, MDC_PS_Common.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" & oMat03.Columns("CpCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1) * oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        End If
                        Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        '공정단가 계산_E
                        
                        '합계 계산_S
                        For i = 0 To oMat03.VisualRowCount - 1
                        
'                            Call Sbo_Application.MessageBox(oDS_PS_PP030M.GetValue("U_CpPrice", i))
                            TotalAmt = TotalAmt + oDS_PS_PP030M.GetValue("U_CpPrice", i)
                            
                        Next
                        
                        oForm01.Items("Total").Specific.VALUE = TotalAmt
                        '합계 계산_E
                        
                    End If
            
                End If
                
                If pval.ColUID = "CpBCode" Or pval.ColUID = "CpCode" Then '작업일보가 등록된 작지 중에서 공정대분류와 공정중분류는 수정 불가
                    
                    If (PS_PP030_Validate("수정03") = False) Then
                        Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, Trim(oDS_PS_PP030M.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                    Else
                        
                        If (pval.ColUID = "CpBCode") Then
                            '//기타작업
                            Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            Call oDS_PS_PP030M.setValue("U_CpBName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT Name FROM [@PS_PP001H] WHERE Code = '" & oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                            If oMat03.RowCount = pval.Row And Trim(oDS_PS_PP030M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                                PS_PP030_AddMatrixRow02 (pval.Row)
                            End If
                        ElseIf (pval.ColUID = "CpCode") Then
                            Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            Call oDS_PS_PP030M.setValue("U_CpName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE Code = '" & oMat03.Columns("CpBCode").Cells(pval.Row).Specific.VALUE & "' AND U_CpCode = '" & oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                            Call oDS_PS_PP030M.setValue("U_StdHour", pval.Row - 1, 0)
                            Call oDS_PS_PP030M.setValue("U_CpPrice", pval.Row - 1, 0)
                            Call oDS_PS_PP030M.setValue("U_ResultYN", pval.Row - 1, "Y")
                            If oMat03.Columns("CpCode").Cells(pval.Row).Specific.VALUE = "CP50103" Or oMat03.Columns("CpCode").Cells(pval.Row).Specific.VALUE = "CP50106" Then
                                Call oDS_PS_PP030M.setValue("U_ReWorkYN", pval.Row - 1, "Y")
                                
                                'Call oMat03.Columns("ReWorkYN").Cells(pval.Row).Specific.Select(0, psk_Index) '//PK/탈지일때 재작업여부 예
                            Else
                                Call oDS_PS_PP030M.setValue("U_ReWorkYN", pval.Row - 1, "N")
                                'Call oMat03.Columns("ReWorkYN").Cells(pval.Row).Specific.Select(1, psk_Index) '//PK/탈지일때 재작업여부 아니오
                            End If
                            
    '                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티일때
    '                            Exist = False
    '                            For i = 1 To oMat03.RowCount - 1
    '                                If Trim(oDS_PS_PP030M.GetValue("U_CpCode", i - 1)) = "CP50106" Then '//탈지공정이 있으면
    '                                    Exist = True
    '                                End If
    '                            Next
    '                            If Exist = True Then
    '                                Call oForm01.Items("MulGbn1").Specific.Select("10", psk_ByValue)
    '                            ElseIf Exist = False Then
    '                                Call oForm01.Items("MulGbn1").Specific.Select("20", psk_ByValue)
    '                            End If
    '                        End If
                        ElseIf (pval.ColUID = "StdHour") Then '//표준공수 입력시
                            '//공정단가 표시
                            If oMat03.Columns("WorkGbn").Cells(pval.Row).Specific.VALUE = "10" Then
                                Call oDS_PS_PP030M.setValue("U_CpPrice", pval.Row - 1, MDC_PS_Common.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" & oMat03.Columns("CpCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1) * oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            ElseIf oMat03.Columns("WorkGbn").Cells(pval.Row).Specific.VALUE = "20" Then
                                Call oDS_PS_PP030M.setValue("U_CpPrice", pval.Row - 1, MDC_PS_Common.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" & oMat03.Columns("CpCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1) * oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            End If
                            Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        
                        Else
                            Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        End If
                    End If
                ElseIf (pval.ColUID = "FailCode") Then
                    If oMat03.Columns("ReWorkYN").Cells(pval.Row).Specific.VALUE = "Y" Then
                        Call oDS_PS_PP030M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP030M.setValue("U_FailName", pval.Row - 1, MDC_PS_Common.GetValue("Select U_SmalName From [@PS_PP003L] Where U_SmalCode = '" & oMat03.Columns("FailCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                    End If
                End If
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_PP030H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "OrdMgNum") Then
                    '//생산요청,작번요청
                    If (oForm01.Items("BaseType").Specific.VALUE = "작번요청") Then
                        Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oDS_PS_PP030H.GetValue("U_" & pval.ItemUID, 0))
                    Else '//생산요청이나, 기준문서타입이 없는경우
                        If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
                            Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                            Call oDS_PS_PP030H.setValue("U_OrdNum", 0, "")
                            Call oDS_PS_PP030H.setValue("U_OrdSub1", 0, "")
                            Call oDS_PS_PP030H.setValue("U_OrdSub2", 0, "")
                        Else
                            Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                            Call oDS_PS_PP030H.setValue("U_OrdNum", 0, oForm01.Items(pval.ItemUID).Specific.VALUE & MDC_PS_Common.GetValue("EXEC PS_PP030_01 '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
                            Call oDS_PS_PP030H.setValue("U_OrdSub1", 0, "00")
                            Call oDS_PS_PP030H.setValue("U_OrdSub2", 0, "000")
                        End If
                    End If
                ElseIf (pval.ItemUID = "ItemCode") Then
'                    If oForm01.Mode = fm_UPDATE_MODE Then
'                        Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                        Call oDS_PS_PP030H.setValue("U_ItemName", 0, MDC_GetData.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                    Else
                        Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                        Call oDS_PS_PP030H.setValue("U_ItemName", 0, MDC_GetData.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
                        If (oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104") Then '//멀티일경우
                            Call oDS_PS_PP030H.setValue("U_MulGbn1", 0, "")
                            Call oDS_PS_PP030H.setValue("U_MulGbn2", 0, "")
                            Call oDS_PS_PP030H.setValue("U_MulGbn3", 0, "")
                            Call oDS_PS_PP030H.setValue("U_MulGbn1", 0, MDC_PS_Common.GetValue("SELECT U_Jakup1 FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                            Call oDS_PS_PP030H.setValue("U_MulGbn2", 0, MDC_PS_Common.GetValue("SELECT U_Jakup2 FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                            Call oDS_PS_PP030H.setValue("U_MulGbn3", 0, MDC_PS_Common.GetValue("SELECT U_Jakup3 FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                        Else
                            Call oDS_PS_PP030H.setValue("U_MulGbn1", 0, "")
                            Call oDS_PS_PP030H.setValue("U_MulGbn2", 0, "")
                            Call oDS_PS_PP030H.setValue("U_MulGbn3", 0, "")
                        End If
                        Call PS_PP030_MTX03 '//공정리스트 처리
                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "102" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "111" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "601" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "602" Then  '//휘팅,부품이면
                            Call PS_PP030_MTX02 '//투입자재 처리
                        End If
                        '//멀티,엔드베어링의 경우 작명을 업데이트
                        If (oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107") Then
                            Call oDS_PS_PP030H.setValue("U_JakMyung", 0, oDS_PS_PP030H.GetValue("U_ItemName", 0))
                        End If
'                    End If
                ElseIf (pval.ItemUID = "SelWt") Then
                    Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    If Val(oForm01.Items(pval.ItemUID).Specific.VALUE) < 0 Then
                        Call oDS_PS_PP030H.setValue("U_SelWt", 0, 0)
                        Call MDC_Com.MDC_GF_Message("수,중량이 올바르지 않습니다.", "W")
                    End If
                    If oForm01.Items("BaseType").Specific.VALUE <> "" Then
                        If Val(oForm01.Items(pval.ItemUID).Specific.VALUE) > Val(oForm01.Items("ReqWt").Specific.VALUE) Then '//투입수량이 수주수량보다 크면
                            Call oDS_PS_PP030H.setValue("U_SelWt", 0, oForm01.Items("ReqWt").Specific.VALUE)
                            Call MDC_Com.MDC_GF_Message("수,중량이 올바르지 않습니다.", "W")
                        End If
                    End If
                 ElseIf (pval.ItemUID = "ReqCod") Then ' 요청자 추가 20180726 황영수
                    sQry = "SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & Trim(oForm01.Items("ReqCod").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("ReqNam").Specific.VALUE = oRecordSet01.Fields(0).VALUE
                 ElseIf (pval.ItemUID = "CntcCode") Then
                    Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_PP030H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                Else
                    If pval.ItemUID = "SItemCod" Or pval.ItemUID = "SCardCod" Then
                    Else
                        Call oDS_PS_PP030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    End If
                End If
            End If
            oMat02.LoadFromDataSource
            oMat03.LoadFromDataSource
            oMat02.AutoResizeColumns
            oMat02.AutoResizeColumns
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
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    Dim i As Long
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_PP030_FormItemEnabled
        If pval.ItemUID = "Mat01" Then
            oMat01.Clear
            oMat01.FlushToDataSource
            oMat01.LoadFromDataSource
        ElseIf pval.ItemUID = "Mat02" Then
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            For i = 1 To oMat02.VisualRowCount
                If i <= oMat02.VisualRowCount Then
                    If oMat02.Columns("InputGbn").Cells(i).Specific.VALUE = "" Then
                        oMat02.DeleteRow (i)
                        i = i - 1
                    End If
                End If
            Next
            For i = 1 To oMat02.VisualRowCount
                oMat02.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next
            oMat02.FlushToDataSource
            If oMat02.VisualRowCount = 0 Then
                Call PS_PP030_AddMatrixRow01(oMat02.VisualRowCount, True) '//UDO방식
            Else
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
                Call PS_PP030_AddMatrixRow01(oMat02.VisualRowCount) '//UDO방식
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
        ElseIf pval.ItemUID = "Mat03" Then
            Call PS_PP030_AddMatrixRow02(oMat03.VisualRowCount) '//UDO방식
            If Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "104" Then
                oMat03.Columns("Sequence").TitleObject.Sortable = True
                Call oMat03.Columns("Sequence").TitleObject.Sort(gst_Ascending)
                oMat03.Columns("Sequence").TitleObject.Sortable = False
                oMat03.FlushToDataSource
            End If
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
        Call PS_PP030_FormResize
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
        If (pval.ItemUID = "SItemCod") Then
            If pval.SelectedObjects Is Nothing Then
            Else
                Set oDataTable01 = pval.SelectedObjects
                oForm01.DataSources.UserDataSources("SItemCod").VALUE = oDataTable01.Columns("ItemCode").Cells(0).VALUE
                oForm01.DataSources.UserDataSources("SItemNam").VALUE = oDataTable01.Columns("ItemName").Cells(0).VALUE
                Set oDataTable01 = Nothing
            End If
        End If
        
        If (pval.ItemUID = "SCardCod") Then
            If pval.SelectedObjects Is Nothing Then
            Else
                Set oDataTable01 = pval.SelectedObjects
                oForm01.DataSources.UserDataSources("SCardCod").VALUE = oDataTable01.Columns("CardCode").Cells(0).VALUE
                oForm01.DataSources.UserDataSources("SCardNam").VALUE = oDataTable01.Columns("CardName").Cells(0).VALUE
                Set oDataTable01 = Nothing
            End If
        End If
'        If (pval.ItemUID = "CntcCode") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030H", "U_CntcCode,U_CntcName")
'        End If
'        If (pval.ItemUID = "Mat02") Then
'            If (pval.ColUID = "CntcCode") Then
'                If pval.SelectedObjects Is Nothing Then
'                Else
'                    Set oDataTable01 = pval.SelectedObjects
'                    Call oDS_PS_PP030L.setValue("U_CntcCode", pval.Row - 1, oDataTable01.Columns("empID").Cells(0).Value)
'                    Call oDS_PS_PP030L.setValue("U_CntcName", pval.Row - 1, oDataTable01.Columns("firstName").Cells(0).Value & oDataTable01.Columns("lastName").Cells(0).Value)
'                    Set oDataTable01 = Nothing
'                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
'                    oMat02.LoadFromDataSource
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
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            If oLastItemUID01 = "Mat02" Then
                If (PS_PP030_Validate("행삭제02") = False) Then
                    BubbleEvent = False
                    Exit Sub
                End If
            ElseIf oLastItemUID01 = "Mat03" Then
'                If oForm01.Items("OrdGbn").Specific.Selected.Value = "104" Then '//멀티일경우
'                    Call MDC_Com.MDC_GF_Message("멀티게이지는 공정을 변경할수 없습니다.", "W")
'                    BubbleEvent = False
'                    Exit Sub
'                ElseIf oForm01.Items("OrdGbn").Specific.Selected.Value = "107" Then '//엔드베어링일경우
'                    Call MDC_Com.MDC_GF_Message("엔드베어링은 공정을 변경할수 없습니다.", "W")
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                If oMat03.Columns("CpCode").Cells(oMat03Row03).Specific.Value = "CP30112" Then
'                    Call MDC_Com.MDC_GF_Message("바렐공정은 변경할수 없습니다.", "W")
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                If oMat03.Columns("CpCode").Cells(oMat03Row03).Specific.Value = "CP30114" Then
'                    Call MDC_Com.MDC_GF_Message("포장공정은 변경할수 없습니다.", "W")
'                    BubbleEvent = False
'                    Exit Sub
'                End If
                
                If (PS_PP030_Validate("행삭제03") = False) Then
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        ElseIf pval.BeforeAction = False Then
            If oLastItemUID01 = "Mat02" Then
                For i = 1 To oMat02.VisualRowCount
                    oMat02.Columns("LineNum").Cells(i).Specific.VALUE = i
                    Call oMat02.Columns("InputGbn").Cells(i).Specific.Select(0, psk_Index)
                Next i
                oMat02.FlushToDataSource
                Call oDS_PS_PP030L.RemoveRecord(oDS_PS_PP030L.Size - 1)
                oMat02.LoadFromDataSource
                If oMat02.RowCount = 0 Then
                    Call PS_PP030_AddMatrixRow01(0)
                Else
                    If Trim(oDS_PS_PP030L.GetValue("U_ItemCode", oMat02.RowCount - 1)) <> "" Then
                        Call PS_PP030_AddMatrixRow01(oMat02.RowCount)
                    End If
                End If
                If Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "104" Or Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "107" Then
                    '류영조
                    Dim TotalQty As Double
                    Dim useMkg As Currency
                    Dim oRecordSet01 As SAPbobsCOM.Recordset
                    Dim sQry As String
                    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

                    If Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "107" Then
                        sQry = "Select IsNull(U_useMkg, 0) From [OITM] Where ItemCode = '" & Trim(oDS_PS_PP030H.GetValue("U_ItemCode", 0)) & "'"
                        oRecordSet01.DoQuery sQry
                        useMkg = oRecordSet01.Fields(0).VALUE / 1000
                        
                        For i = 1 To oDS_PS_PP030L.Size - 1
                            TotalQty = TotalQty + oDS_PS_PP030L.GetValue("U_Weight", i - 1)
                        Next
                        If useMkg = 0 Then
                            Call oDS_PS_PP030H.setValue("U_SelWt", 0, TotalQty)
                        Else
                            Call oDS_PS_PP030H.setValue("U_SelWt", 0, TotalQty / useMkg)
                        End If
                    Else
                        For i = 1 To oDS_PS_PP030L.Size - 1
                            TotalQty = TotalQty + oDS_PS_PP030L.GetValue("U_Weight", i - 1)
                        Next
                        Call oDS_PS_PP030H.setValue("U_SelWt", 0, TotalQty)
                    End If
                    
                    Set oRecordSet01 = Nothing
                    
'                    Dim TotalQty As Double
'                    For i = 1 To oDS_PS_PP030L.Size - 1
'                        TotalQty = TotalQty + oDS_PS_PP030L.GetValue("U_Weight", i - 1)
'                    Next
'                    Call oDS_PS_PP030H.setValue("U_SelWt", 0, TotalQty)
                    oMat02.LoadFromDataSource
                    oForm01.Update
                End If
            ElseIf oLastItemUID01 = "Mat03" Then
                For i = 1 To oMat03.VisualRowCount
                    oMat03.Columns("LineNum").Cells(i).Specific.VALUE = i
                    oMat03.Columns("Sequence").Cells(i).Specific.VALUE = i
                Next i
                oMat03.FlushToDataSource
                Call oDS_PS_PP030M.RemoveRecord(oDS_PS_PP030M.Size - 1)
                oMat03.LoadFromDataSource
                If oMat03.RowCount = 0 Then
                    Call PS_PP030_AddMatrixRow02(0)
                Else
                    If Trim(oDS_PS_PP030M.GetValue("U_CpBCode", oMat03.RowCount - 1)) <> "" Then
                        Call PS_PP030_AddMatrixRow02(oMat03.RowCount)
                    End If
                End If
                
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티일때
'                    Call oForm01.Items("MulGbn1").Specific.Select("20", psk_ByValue)
'                    For i = 1 To oMat03.RowCount - 1
'                        If oMat03.Columns("CpCode").Cells(i).Specific.VALUE = "CP50106" Then '//탈지공정이 있으면
'                            Call oForm01.Items("MulGbn1").Specific.Select("10", psk_ByValue)
'                            Exit For
'                        End If
'                    Next
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP030_CreateItems() As Boolean
On Error GoTo PS_PP030_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_USERDS01 = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_PP030H = oForm01.DataSources.DBDataSources("@PS_PP030H")
    Set oDS_PS_PP030L = oForm01.DataSources.DBDataSources("@PS_PP030L")
    Set oDS_PS_PP030M = oForm01.DataSources.DBDataSources("@PS_PP030M")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    Set oMat03 = oForm01.Items("Mat03").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat02.SelectionMode = ms_NotSupported
    oMat03.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    oMat02.AutoResizeColumns
    oMat03.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("SBPLId", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("ItmBsort", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("ItmMsort", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("ReqType", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("SItemCod", dt_SHORT_TEXT, 20)
    Call oForm01.DataSources.UserDataSources.Add("SItemNam", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("SCardCod", dt_SHORT_TEXT, 20)
    Call oForm01.DataSources.UserDataSources.Add("SCardNam", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("ReqCod", dt_SHORT_TEXT, 20)
    Call oForm01.DataSources.UserDataSources.Add("ReqNam", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("Opt03", dt_SHORT_TEXT, 10)
    
    Call oForm01.Items("SBPLId").Specific.DataBind.SetBound(True, "", "SBPLId")
    Call oForm01.Items("ItmBsort").Specific.DataBind.SetBound(True, "", "ItmBsort")
    Call oForm01.Items("ItmMsort").Specific.DataBind.SetBound(True, "", "ItmMsort")
    Call oForm01.Items("ReqType").Specific.DataBind.SetBound(True, "", "ReqType")
    Call oForm01.Items("SItemCod").Specific.DataBind.SetBound(True, "", "SItemCod")
    Call oForm01.Items("SItemNam").Specific.DataBind.SetBound(True, "", "SItemNam")
    Call oForm01.Items("SCardCod").Specific.DataBind.SetBound(True, "", "SCardCod")
    Call oForm01.Items("SCardNam").Specific.DataBind.SetBound(True, "", "SCardNam")
    Call oForm01.Items("ReqCod").Specific.DataBind.SetBound(True, "", "ReqCod")
    Call oForm01.Items("ReqNam").Specific.DataBind.SetBound(True, "", "ReqNam")
    Call oForm01.Items("Opt01").Specific.DataBind.SetBound(True, "", "Opt01")
    Call oForm01.Items("Opt02").Specific.DataBind.SetBound(True, "", "Opt02")
    Call oForm01.Items("Opt03").Specific.DataBind.SetBound(True, "", "Opt03")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt03")
    
    Call oForm01.DataSources.UserDataSources.Add("Total", dt_SUM)
    Call oForm01.Items("Total").Specific.DataBind.SetBound(True, "", "Total")
'    oForm01.DataSources.UserDataSources.Item("DocDateFr").Value = 0
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP030_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP030_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP030_ComboBox_Setting()
On Error GoTo PS_PP030_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "ReqType", "", "10", "계획생산요청")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "ReqType", "", "20", "수주생산요청")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("ReqType").Specific, "PS_PP030", "ReqType", True)
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ReqType", "10", "계획생산요청")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ReqType", "20", "수주생산요청")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("ReqType"), "PS_PP030", "Mat01", "ReqType")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "BasicGub", "", "10", "통합")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "BasicGub", "", "20", "비통합")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("BasicGub").Specific, "PS_PP030", "BasicGub")
    Call oForm01.Items("BasicGub").Specific.Select("비통합", psk_ByValue)
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn1", "", "10", "탈지")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn1", "", "20", "비탈지")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("MulGbn1").Specific, "PS_PP030", "MulGbn1")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn2", "", "10", "시계")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn2", "", "20", "반시계")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("MulGbn2").Specific, "PS_PP030", "MulGbn2")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn3", "", "10", "배면")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "MulGbn3", "", "20", "상면")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("MulGbn3").Specific, "PS_PP030", "MulGbn3")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn", "10", "일반")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn", "20", "원재료")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn", "30", "스크랩")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn2", "20", "원재료")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn2", "30", "스크랩")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("InputGbn"), "PS_PP030", "Mat02", "InputGbn")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "ProcType", "10", "청구")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "ProcType", "20", "잔재")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat02", "ProcType", "30", "취소")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("ProcType"), "PS_PP030", "Mat02", "ProcType")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "WorkGbn", "10", "자가")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "WorkGbn", "20", "정밀")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "WorkGbn", "30", "외주")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns("WorkGbn"), "PS_PP030", "Mat03", "WorkGbn")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ResultYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ResultYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns("ResultYN"), "PS_PP030", "Mat03", "ResultYN")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReWorkYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReWorkYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns("ReWorkYN"), "PS_PP030", "Mat03", "ReWorkYN")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReportYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReportYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns("ReportYN"), "PS_PP030", "Mat03", "ReportYN")
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP030", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP030", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PP030", "ItemCode")
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, True)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", False, True)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmMsort").Specific, "SELECT PSH_ITMMSORT.U_Code, PSH_ITMMSORT.U_CodeName FROM [@PSH_ITMMSORT] PSH_ITMMSORT LEFT JOIN [@PSH_ITMBSORT] PSH_ITMBSORT ON PSH_ITMBSORT.Code = PSH_ITMMSORT.U_rCode WHERE PSH_ITMBSORT.U_PudYN = 'Y'", "", False, True)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] order by Code", "", False, True)
    Call oForm01.Items("OrdGbn").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", False, False)
    Call oForm01.Items("BPLId").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat02.Columns("ItemGpCd"), "SELECT ItmsGrpCod, ItmsGrpNam FROM [OITB]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat03.Columns("Unit"), "SELECT Code, Name FROM [@PSH_CPUOM]")
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP030_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP030_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP030_CF_ChooseFromList()
On Error GoTo PS_PP030_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

'    Set oEdit = oForm01.Items("CntcCode").Specific
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
'    oEdit.ChooseFromListUID = "CFLEMPLOYEE"
'    oEdit.ChooseFromListAlias = "empID"
'
'    Set oColumn = oMat02.Columns("CntcCode")
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_Employee
'    oCFLCreationParams.uniqueID = "CFLEMPLOYEE2"
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
'    oColumn.ChooseFromListUID = "CFLEMPLOYEE2"
'    oColumn.ChooseFromListAlias = "empID"
    
    Set oEdit = oForm01.Items("SCardCod").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = "2"
    oCFLCreationParams.uniqueID = "CFLSCARDCOD"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oEdit.ChooseFromListUID = "CFLSCARDCOD"
    oEdit.ChooseFromListAlias = "CardCode"
    
    Set oEdit = oForm01.Items("SItemCod").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = "4"
    oCFLCreationParams.uniqueID = "CFLSITEMCOD"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oEdit.ChooseFromListUID = "CFLSITEMCOD"
    oEdit.ChooseFromListAlias = "ItemCode"

    Exit Sub
PS_PP030_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP030_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP030_FormItemEnabled()
On Error GoTo PS_PP030_FormItemEnabled_Error

    Dim sQry01 As String
    Dim sQry02 As String

    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("OrdGbn").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("DocDate").Enabled = True '2011.12.05 송명규 수정(True에서 False로) False에서 True로 환원(2017.02.21 송명규)
        oForm01.Items("DueDate").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("BasicGub").Enabled = True
        oForm01.Items("MulGbn1").Enabled = False
        oForm01.Items("MulGbn2").Enabled = False
        oForm01.Items("MulGbn3").Enabled = False
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("OrdMgNum").Enabled = True
        oForm01.Items("SelWt").Enabled = True
        oForm01.Items("Mat01").Enabled = True
        oForm01.Items("Mat02").Enabled = True
        oForm01.Items("Mat03").Enabled = True
        oForm01.Items("Button01").Enabled = True
        oForm01.Items("1").Enabled = True
        
'        Call oForm01.Items("SBPLId").Specific.Select(0, psk_Index)
'        Call oForm01.Items("ItmBsort").Specific.Select(0, psk_Index)
'        Call oForm01.Items("ItmMsort").Specific.Select(0, psk_Index)
'        Call oForm01.Items("ReqType").Specific.Select(0, psk_Index)
        Call oForm01.Items("OrdGbn").Specific.Select(0, psk_Index)
        Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        oForm01.Items("DocDate").Specific.VALUE = Format(Date, "YYYYMMDD")
        oForm01.Items("DueDate").Specific.VALUE = Format(Date, "YYYYMMDD")
        oForm01.Items("SItemCod").Specific.VALUE = ""
        oForm01.Items("SCardCod").Specific.VALUE = ""
        oForm01.Items("OrdMgNum").Specific.VALUE = ""
        oForm01.Items("OrdNum").Specific.VALUE = ""
        oForm01.Items("OrdSub1").Specific.VALUE = ""
        oForm01.Items("OrdSub2").Specific.VALUE = ""
        oForm01.Items("OrdMgNum").Specific.VALUE = Format(Date, "YYYYMMDD")
        oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
        Call PS_PP030_FormClear '//UDO방식
        oMat02.Columns("BatchNum").Editable = False
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("OrdGbn").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("DocDate").Enabled = True '2011.12.05 송명규 수정(True에서 False로) False에서 True로 환원(2017.02.21 송명규)
        oForm01.Items("DueDate").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("BasicGub").Enabled = True
        oForm01.Items("MulGbn1").Enabled = False
        oForm01.Items("MulGbn2").Enabled = False
        oForm01.Items("MulGbn3").Enabled = False
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("OrdMgNum").Enabled = True
        oForm01.Items("SelWt").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        oForm01.Items("Mat02").Enabled = False
        oForm01.Items("Mat03").Enabled = False
        oForm01.Items("Button01").Enabled = True
        oForm01.Items("1").Enabled = True
        oMat02.Columns("BatchNum").Editable = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oMat02.Columns("BatchNum").Editable = False
        '//각모드에따른 아이템설정
        If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP030H] WHERE DocEntry = '" & Trim(oDS_PS_PP030H.GetValue("DocEntry", 0)) & "'", 0, 1) = "Y" Then
            oForm01.Items("Focus").Click ct_Regular
            oForm01.Items("OrdGbn").Enabled = False
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("DocDate").Enabled = False
            oForm01.Items("DueDate").Enabled = False
            oForm01.Items("ItemCode").Enabled = False
            oForm01.Items("CntcCode").Enabled = False
            oForm01.Items("MulGbn1").Enabled = False
            oForm01.Items("MulGbn2").Enabled = False
            oForm01.Items("MulGbn3").Enabled = False
            oForm01.Items("DocEntry").Enabled = False
            oForm01.Items("OrdMgNum").Enabled = False
            oForm01.Items("SelWt").Enabled = False
            oForm01.Items("Mat01").Enabled = False
            oForm01.Items("Mat02").Enabled = False
            oForm01.Items("Mat03").Enabled = False
            oForm01.Items("Button01").Enabled = False
            oForm01.Items("1").Enabled = False
        Else
            oForm01.Items("Focus").Click ct_Regular
            oForm01.Items("Mat01").Enabled = True
            oForm01.Items("Mat02").Enabled = True
            oForm01.Items("Mat03").Enabled = True
            oForm01.Items("Button01").Enabled = True
            oForm01.Items("1").Enabled = True
                        
            oForm01.Items("DocEntry").Enabled = False
            oForm01.Items("OrdGbn").Enabled = False
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("ItemCode").Enabled = False
            oForm01.Items("OrdMgNum").Enabled = False
            
            '실적(작업일보)문서가 없고 원가 상의 재공에 투입된 자료가 아니라면 아래 필드의 데이터는 수정(2017.02.21 송명규)
            '실적 자료 조회용 쿼리
            sQry01 = "          SELECT  COUNT(*)"
            sQry01 = sQry01 & " FROM    [@PS_PP040H] AS T0"
            sQry01 = sQry01 & "         INNER JOIN"
            sQry01 = sQry01 & "         [@PS_PP040L] AS T1"
            sQry01 = sQry01 & "             ON T0.DocEntry = T1.DocEntry"
            sQry01 = sQry01 & " WHERE   T0.Canceled = 'N'"
            sQry01 = sQry01 & "         AND T1.U_PP030HNo = " & Trim(oDS_PS_PP030H.GetValue("DocEntry", 0))
            
            '원가 자료 조회용 쿼리
            sQry02 = "          SELECT  COUNT(*)"
            sQry02 = sQry02 & " FROM    [@PS_CO130L] AS T0"
            sQry02 = sQry02 & " WHERE   T0.U_POEntry = " & Trim(oDS_PS_PP030H.GetValue("DocEntry", 0))
            
            'If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & Trim(oDS_PS_PP030H.GetValue("DocEntry", 0)) & "'") = 0 Then
            If MDC_PS_Common.GetValue(sQry01) = 0 Or MDC_PS_Common.GetValue(sQry02) = 0 Then '실적(작업일보) 및 원가계산된 자료가 없으면 아래 자료는 수정가능
                oForm01.Items("DocDate").Enabled = True '2011.12.05 송명규 수정(True에서 False로), False에서 True로 환원(2017.02.21 송명규)
                oForm01.Items("DueDate").Enabled = True
                oForm01.Items("CntcCode").Enabled = True
                oForm01.Items("SelWt").Enabled = True
                If Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "104" Then '//작업구분이 멀티일때만
                    oForm01.Items("BasicGub").Enabled = True
                    oForm01.Items("MulGbn1").Enabled = True
                    oForm01.Items("MulGbn2").Enabled = True
                    oForm01.Items("MulGbn3").Enabled = True
                Else
                    oForm01.Items("BasicGub").Enabled = False
                    oForm01.Items("MulGbn1").Enabled = False
                    oForm01.Items("MulGbn2").Enabled = False
                    oForm01.Items("MulGbn3").Enabled = False
                End If
            Else '//실적 등록 및 원가 계산된 자료가 하나라도 있으면
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("DueDate").Enabled = False
                oForm01.Items("CntcCode").Enabled = False
                oForm01.Items("SelWt").Enabled = False
                oForm01.Items("MulGbn1").Enabled = True
                oForm01.Items("MulGbn2").Enabled = False
                oForm01.Items("MulGbn3").Enabled = False

            End If
            
            If (Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "107") Then
                oMat02.Columns("InputGbn").Editable = True
            Else
                oMat02.Columns("InputGbn").Editable = False
            End If
            
            If (Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "105" Or Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "106") Then
                oMat02.Columns("Weight").Editable = True
            Else
                oMat02.Columns("Weight").Editable = False
            End If
            
'            If (Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "104" Or Trim(oDS_PS_PP030H.GetValue("U_OrdGbn", 0)) = "107") Then
'                oMat03.Columns("CpBCode").Editable = False
'                oMat03.Columns("CpCode").Editable = False
'                oMat03.Columns("ResultYN").Editable = False
'                oMat03.Columns("ReportYN").Editable = False
'            Else
'                oMat03.Columns("CpBCode").Editable = True
'                oMat03.Columns("CpCode").Editable = True
'                oMat03.Columns("ResultYN").Editable = True
'                oMat03.Columns("ReportYN").Editable = True
'            End If
        End If
    End If
'    oMat01.AutoResizeColumns
'    oMat02.AutoResizeColumns
'    oMat03.AutoResizeColumns
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP030_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP030_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP030_AddMatrixRow01(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP030_AddMatrixRow01_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP030L.InsertRecord (oRow)
    End If
    oMat02.AddRow
    oDS_PS_PP030L.Offset = oRow
    oDS_PS_PP030L.setValue "U_LineNum", oRow, oRow + 1
    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then '//엔드베어링은 투입구분,원재료,스크랩
        oDS_PS_PP030L.setValue "U_InputGbn", oRow, "20"
    Else '//나머지경우는 일반으로 선택
        oDS_PS_PP030L.setValue "U_InputGbn", oRow, "10"
    End If
    oDS_PS_PP030L.setValue "U_ProcType", oRow, "20"
    oMat02.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP030_AddMatrixRow01_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP030_AddMatrixRow01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_PP030_AddMatrixRow02(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP030_AddMatrixRow02_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP030M.InsertRecord (oRow)
    End If
    oMat03.AddRow
    oDS_PS_PP030M.Offset = oRow
    oDS_PS_PP030M.setValue "U_LineNum", oRow, oRow + 1
    oDS_PS_PP030M.setValue "U_Sequence", oRow, oRow + 1
    oDS_PS_PP030M.setValue "U_WorkGbn", oRow, "10"
    oDS_PS_PP030M.setValue "U_ReWorkYN", oRow, "N"
    oDS_PS_PP030M.setValue "U_ResultYN", oRow, "N"
    oDS_PS_PP030M.setValue "U_ReportYN", oRow, "Y"
    oMat03.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP030_AddMatrixRow02_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP030_AddMatrixRow02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP030_FormClear()
On Error GoTo PS_PP030_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP030'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    If (oForm01.Items("OrdGbn").Specific.Selected.VALUE <> "105" And oForm01.Items("OrdGbn").Specific.Selected.VALUE <> "106" And oForm01.Items("OrdGbn").Specific.Selected.VALUE <> "선택") Then
        If oForm01.Items("OrdMgNum").Specific.VALUE <> "" Then
            oForm01.Items("OrdNum").Specific.VALUE = oForm01.Items("OrdMgNum").Specific.VALUE & MDC_PS_Common.GetValue("EXEC PS_PP030_01 '" & oForm01.Items("OrdMgNum").Specific.VALUE & "'")
            oForm01.Items("OrdSub1").Specific.VALUE = "00"
            oForm01.Items("OrdSub2").Specific.VALUE = "000"
        End If
    End If
    Exit Sub
PS_PP030_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP030_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP030_EnableMenus()
On Error GoTo PS_PP030_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, True, True, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP030_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP030_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP030_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP030_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP030_FormItemEnabled
        Call PS_PP030_AddMatrixRow01(0, True) '//UDO방식일때
        Call PS_PP030_AddMatrixRow02(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_PP030_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP030_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP030_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP030_DataValidCheck() As Boolean
On Error GoTo PS_PP030_DataValidCheck_Error
    PS_PP030_DataValidCheck = False
    Dim i, j As Long
    Dim CP30112 As Boolean
    Dim CP30114 As Boolean
    
    Dim Lot104Exsits As Integer '멀티구전산 이전 lotno
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Dim sQry As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP030_FormClear
    End If
    
    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then
        Sbo_Application.SetStatusBarMessage "작지구분은 필수입니다.", bmt_Short, True
        oForm01.Items("OrdGbn").Click ct_Regular
        PS_PP030_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("BPLId").Specific.Selected.VALUE = "선택" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm01.Items("BPLId").Click ct_Regular
        PS_PP030_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("OrdNum").Specific.VALUE = "" Or oForm01.Items("OrdSub1").Specific.VALUE = "" Or oForm01.Items("OrdSub2").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "작지번호는 필수입니다.", bmt_Short, True
        oForm01.Items("OrdMgNum").Click ct_Regular
        PS_PP030_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "지시일자는 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_PP030_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("ItemCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "품목코드는 필수입니다.", bmt_Short, True
        oForm01.Items("ItemCode").Click ct_Regular
        PS_PP030_DataValidCheck = False
        Exit Function
    End If
    If Val(oForm01.Items("SelWt").Specific.VALUE) <= 0 Then
        Sbo_Application.SetStatusBarMessage "지시수,중량이 올바르지 않습니다.", bmt_Short, True
        oForm01.Items("SelWt").Click ct_Regular
        PS_PP030_DataValidCheck = False
        Exit Function
    End If
    
    '기계공구일경우 작번등록과 일자 비교_S(2017.02.21 송명규 추가)
    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then
    
        If PS_PP030_CheckDate = False Then
            Sbo_Application.SetStatusBarMessage "작업지시등록일은 작번등록일과 같거나 늦어야합니다. 확인하십시오.", bmt_Short, True
            PS_PP030_DataValidCheck = False
            Exit Function
        End If
    
    End If
    '기계공구일경우 작번등록과 일자 비교_E(2017.02.21 송명규 추가)
    
'    '마감상태 체크_S(2017.11.23 송명규 추가), 순환품인 경우 공정등록을 3달 후에도 등록하는 경우가 있어 주석 처리(2017.11.28 송명규)
'    If MDC_PS_Common.Check_Finish_Status(oForm01.Items("BPLId").Specific.VALUE, oForm01.Items("DocDate").Specific.VALUE) = False Then
'        Call Sbo_Application.SetStatusBarMessage("마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 지시일자를 확인하고, 회계부서로 문의하세요.", bmt_Short, True)
'        PS_PP030_DataValidCheck = False
'        Exit Function
'    End If
'    '마감상태 체크_E(2017.11.23 송명규 추가)
    
    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then '//휘팅일경우
'        CP30112 = False
'        CP30114 = False
'        For i = 1 To oMat03.VisualRowCount
'            If (oMat03.Columns("CpBCode").Cells(i).Specific.Value = "CP301") Then '//공정이 휘팅이면
'                If oMat03.Columns("CpCode").Cells(i).Specific.Value = "CP30112" Then '//바렐공정
'                    CP30112 = True
'                End If
'                If oMat03.Columns("CpCode").Cells(i).Specific.Value = "CP30114" Then '//포장공정
'                    CP30114 = True
'                End If
'            End If
'        Next
'        If CP30112 <> True Or CP30114 <> True Then
'            MDC_Com.MDC_GF_Message "휘팅공정은 바렐,포장공정은 필수 입니다.", "W"
'            PS_PP030_DataValidCheck = False
'            Exit Function
'        End If
    End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티는 투입자재라인이 존재해야함
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
        Query01 = "SELECT Count(*) FROM Z_DSMDFRY Where lotno = '" & oForm01.Items("OrdNum").Specific.VALUE & "'"
        Call RecordSet01.DoQuery(Query01)
        Lot104Exsits = RecordSet01.Fields(0).VALUE
        
        If Lot104Exsits = 0 Then
            If oMat02.VisualRowCount <= 1 Then
                Sbo_Application.SetStatusBarMessage "투입자재라인이 존재하지 않습니다.", bmt_Short, True
                PS_PP030_DataValidCheck = False
                Exit Function
            End If
        End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////매트릭스 행없이입력하기
    If oMat03.VisualRowCount <= 1 Then
        Sbo_Application.SetStatusBarMessage "공정리스트 라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP030_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat02.VisualRowCount - 1
        If (oMat02.Columns("InputGbn").Cells(i).Specific.Selected Is Nothing) Then
            Sbo_Application.SetStatusBarMessage "투입구분은 필수입니다.", bmt_Short, True
            oMat02.Columns("InputGbn").Cells(i).Click ct_Regular
            PS_PP030_DataValidCheck = False
            Exit Function
        End If
        If (oMat02.Columns("ItemCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
            oMat02.Columns("ItemCode").Cells(i).Click ct_Regular
            PS_PP030_DataValidCheck = False
            Exit Function
        End If
        If (oMat02.Columns("ItemGpCd").Cells(i).Specific.Selected Is Nothing) Then
            Sbo_Application.SetStatusBarMessage "품목그룹은 필수입니다.", bmt_Short, True
            PS_PP030_DataValidCheck = False
            Exit Function
        End If
        '//휘팅,부품,엔드베어링일경우
        If oForm01.Items("OrdGbn").Specific.Selected.VALUE <> "104" And oForm01.Items("OrdGbn").Specific.Selected.VALUE <> "105" And oForm01.Items("OrdGbn").Specific.Selected.VALUE <> "106" And oForm01.Items("OrdGbn").Specific.Selected.VALUE <> "107" Then
            If (oMat02.VisualRowCount > 2) Then
                Sbo_Application.SetStatusBarMessage "해당작지는 투입자재 한품목만 입력가능합니다.", bmt_Short, True
                PS_PP030_DataValidCheck = False
                Exit Function
            End If
        End If
        '//기계공구,몰드인경우
        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then
            If (oMat02.Columns("ProcType").Cells(i).Specific.Selected Is Nothing) Then
                Sbo_Application.SetStatusBarMessage "조달방식은 필수입니다.", bmt_Short, True
                oMat02.Columns("ProcType").Cells(i).Click ct_Regular
                PS_PP030_DataValidCheck = False
                Exit Function
            End If
            If Val(oMat02.Columns("Weight").Cells(i).Specific.VALUE) <= 0 Then
                Sbo_Application.SetStatusBarMessage "수,중량은 필수입니다.", bmt_Short, True
                oMat02.Columns("Weight").Cells(i).Click ct_Right
                PS_PP030_DataValidCheck = False
                Exit Function
            End If
        End If
        '//멀티,엔드베어링인경우
        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then
            If Val(oMat02.Columns("Weight").Cells(i).Specific.VALUE) <= 0 Then
                Sbo_Application.SetStatusBarMessage "수,중량은 필수입니다.", bmt_Short, True
                PS_PP030_DataValidCheck = False
                Exit Function
            End If
            If oMat02.Columns("BatchNum").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "배치번호는 필수입니다.", bmt_Short, True
                PS_PP030_DataValidCheck = False
                Exit Function
            End If
        End If
    Next
    
    For i = 1 To oMat03.VisualRowCount - 1
        If (oMat03.Columns("CpBCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "공정대분류는 필수입니다.", bmt_Short, True
            oMat03.Columns("CpBCode").Cells(i).Click ct_Regular
            PS_PP030_DataValidCheck = False
            Exit Function
        End If
        If (oMat03.Columns("CpCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "공정중분류는 필수입니다.", bmt_Short, True
            oMat03.Columns("CpCode").Cells(i).Click ct_Regular
            PS_PP030_DataValidCheck = False
            Exit Function
        End If
'        For j = i + 1 To oMat03.VisualRowCount - 1
'            If (oMat03.Columns("CpBCode").Cells(i).Specific.Value = oMat03.Columns("CpBCode").Cells(j).Specific.Value And oMat03.Columns("CpCode").Cells(i).Specific.Value = oMat03.Columns("CpCode").Cells(j).Specific.Value) Then
'                Sbo_Application.SetStatusBarMessage "중복된 공정이 존재합니다.", bmt_Short, True
'                oMat03.Columns("CpBCode").Cells(j).Click ct_Regular
'                PS_PP030_DataValidCheck = False
'                Exit Function
'            End If
'        Next
    Next
    
    If (PS_PP030_Validate("검사01") = False) Then
        PS_PP030_DataValidCheck = False
        Exit Function
    End If
    
    If (PS_PP030_Validate("검사02") = False) Then
        PS_PP030_DataValidCheck = False
        Exit Function
    End If
    
    If (PS_PP030_Validate("검사03") = False) Then
        PS_PP030_DataValidCheck = False
        Exit Function
    End If

    Call oDS_PS_PP030L.RemoveRecord(oDS_PS_PP030L.Size - 1)
    Call oDS_PS_PP030M.RemoveRecord(oDS_PS_PP030M.Size - 1)
    Call oMat02.LoadFromDataSource
    Call oMat03.LoadFromDataSource
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP030_FormClear
    End If
    PS_PP030_DataValidCheck = True
    Exit Function
PS_PP030_DataValidCheck_Error:
    PS_PP030_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP030_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP030_MTX01()
On Error GoTo PS_PP030_MTX01_Error
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
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim Param08 As String
    
    Param01 = Trim(oForm01.Items("SBPLId").Specific.Selected.VALUE)
    Param02 = Trim(oForm01.Items("ItmBsort").Specific.Selected.VALUE)
    Param03 = Trim(oForm01.Items("ItmMsort").Specific.Selected.VALUE)
    Param04 = Trim(oForm01.Items("ReqType").Specific.Selected.VALUE)
    Param05 = Trim(oForm01.Items("SItemCod").Specific.VALUE)
    Param06 = Trim(oForm01.Items("SCardCod").Specific.VALUE)
    Param07 = Trim(oForm01.Items("Mark").Specific.VALUE)
    Param08 = Trim(oForm01.Items("ReqCod").Specific.VALUE)
    
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    Query01 = "EXEC PS_PP030_02 '" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP030_MTX01_Exit
    End If
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_USERDS01.InsertRecord (i)
        End If
        oDS_PS_USERDS01.Offset = i
        Call oDS_PS_USERDS01.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_USERDS01.setValue("U_ColReg01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg02", i, RecordSet01.Fields(1).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg03", i, RecordSet01.Fields(2).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg04", i, RecordSet01.Fields(3).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg05", i, RecordSet01.Fields(4).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg06", i, RecordSet01.Fields(5).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColDt01", i, RecordSet01.Fields(6).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg07", i, RecordSet01.Fields(7).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg08", i, RecordSet01.Fields(8).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColQty01", i, RecordSet01.Fields(9).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColQty02", i, RecordSet01.Fields(10).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg09", i, RecordSet01.Fields(11).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg10", i, RecordSet01.Fields(12).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg11", i, RecordSet01.Fields(13).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg12", i, RecordSet01.Fields(14).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg13", i, RecordSet01.Fields(15).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg14", i, RecordSet01.Fields(16).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg15", i, RecordSet01.Fields(17).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg16", i, RecordSet01.Fields(18).VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg17", i, RecordSet01.Fields(19).VALUE)
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
PS_PP030_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_PP030_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP030_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP030_MTX02()
On Error GoTo PS_PP030_MTX02_Error
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
    Param01 = Trim(oForm01.Items("ItemCode").Specific.VALUE)
'    Param02 = Trim(oForm01.Items("Param01").Specific.VALUE)
'    Param03 = Trim(oForm01.Items("Param01").Specific.VALUE)
'    Param04 = Trim(oForm01.Items("Param01").Specific.VALUE)
'    If oForm01.Items("OrdGbn").Specific.Selected.Value = "104" Then '//멀티인경우
'        Query01 = "SELECT PS_PP001H.U_CpBCode, PS_PP001H.U_CpBName, PS_PP001L.U_CpCode, PS_PP001L.U_CpName FROM [@PS_PP001H] PS_PP001H LEFT JOIN [@PS_PP001L] PS_PP001L ON PS_PP001H.Code  =PS_PP001L.Code WHERE PS_PP001H.Code = 'CP501'"
'    ElseIf oForm01.Items("OrdGbn").Specific.Selected.Value = "107" Then '//엔드베어링인경우
'        Query01 = "SELECT PS_PP001H.U_CpBCode, PS_PP001H.U_CpBName, PS_PP001L.U_CpCode, PS_PP001L.U_CpName FROM [@PS_PP001H] PS_PP001H LEFT JOIN [@PS_PP001L] PS_PP001L ON PS_PP001H.Code  =PS_PP001L.Code WHERE PS_PP001H.Code = 'CP101'"
'    Else
        Query01 = "SELECT PS_PP005H.U_ItemCod2, PS_PP005H.U_ItemNam2, OITM.ItmsGrpCod FROM [@PS_PP005H] PS_PP005H LEFT JOIN [OITM] OITM ON PS_PP005H.U_ItemCod2 = OITM.ItemCode WHERE U_ItemCod1 = '" & Param01 & "'"
'    End If
    Call RecordSet01.DoQuery(Query01)
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call PS_PP030_AddMatrixRow01(0, True)
        GoTo PS_PP030_MTX02_Exit
    End If
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_PP030L.InsertRecord (i)
        End If
        oDS_PS_PP030L.Offset = i
        Call oDS_PS_PP030L.setValue("U_LineNum", i, i + 1)
        
        Call oDS_PS_PP030L.setValue("U_InputGbn", i, "10") '//투입구분 '//휘팅,부품의경우만 실행되므로 항상 10이다
        Call oDS_PS_PP030L.setValue("U_ItemCode", i, RecordSet01.Fields(0).VALUE) '//품목코드
        Call oDS_PS_PP030L.setValue("U_ItemName", i, RecordSet01.Fields(1).VALUE) '//품목이름
        Call oDS_PS_PP030L.setValue("U_ItemGpCd", i, RecordSet01.Fields(2).VALUE) '//품목그룹
        Call oDS_PS_PP030L.setValue("U_BatchNum", i, "") '//배치번호
        Call oDS_PS_PP030L.setValue("U_Weight", i, 0) '//중량
        Call oDS_PS_PP030L.setValue("U_DueDate", i, "")
        Call oDS_PS_PP030L.setValue("U_CntcCode", i, "")
        Call oDS_PS_PP030L.setValue("U_CntcName", i, "")
        Call oDS_PS_PP030L.setValue("U_ProcType", i, "20")
        Call oDS_PS_PP030L.setValue("U_Comments", i, "")
        Call oDS_PS_PP030L.setValue("U_LineId", i, "")
        If i = RecordSet01.RecordCount - 1 Then
            Call PS_PP030_AddMatrixRow01(i + 1) '//마지막행에 한줄추가
        End If
        RecordSet01.MoveNext
    Next
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    oForm01.Update
        
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP030_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP030_MTX02_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP030_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP030_MTX03()
On Error GoTo PS_PP030_MTX03_Error
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
    
    Dim ItemCode As String
    Dim BasicGub As String
    
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    
    BasicGub = Trim(oForm01.Items("BasicGub").Specific.VALUE)

    Query01 = "         EXEC [PS_PP030_07] '"
    Query01 = Query01 & ItemCode & "','"
    Query01 = Query01 & BasicGub & "'"

    Call RecordSet01.DoQuery(Query01)
    
    oMat03.Clear
    oMat03.FlushToDataSource
    oMat03.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        If Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "105" Or Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "106" Then
            oForm01.Items("Mat03").Enabled = True
        Else
            oForm01.Items("Mat03").Enabled = False
            '//휘팅,부품,멀티,엔베는 표준공정이 등록되지 않으면 진행불가능
        End If
        Call PS_PP030_AddMatrixRow02(0, True)
        '//GoTo PS_PP030_MTX03_Exit
    Else
        oForm01.Items("Mat03").Enabled = True
'        If oForm01.Items("OrdGbn").Specific.VALUE = "104" Then
'            Call oForm01.Items("MulGbn1").Specific.Select("10", psk_ByValue)
'        End If
    End If
    
    If Trim(oForm01.Items("OrdGbn").Specific.VALUE) <> "105" Then
        For i = 0 To RecordSet01.RecordCount - 1
            If i <> 0 Then
                oDS_PS_PP030M.InsertRecord (i)
            End If
            oDS_PS_PP030M.Offset = i
            Call oDS_PS_PP030M.setValue("U_LineNum", i, i + 1)
            Call oDS_PS_PP030M.setValue("U_Sequence", i, i + 1)
            Call oDS_PS_PP030M.setValue("U_CpBCode", i, RecordSet01.Fields(0).VALUE)
            Call oDS_PS_PP030M.setValue("U_CpBName", i, RecordSet01.Fields(1).VALUE)
            Call oDS_PS_PP030M.setValue("U_CpCode", i, RecordSet01.Fields(2).VALUE)
            Call oDS_PS_PP030M.setValue("U_CpName", i, RecordSet01.Fields(3).VALUE)
            Call oDS_PS_PP030M.setValue("U_Unit", i, RecordSet01.Fields(4).VALUE)
            Call oDS_PS_PP030M.setValue("U_ReWorkYN", i, "N")
            Call oDS_PS_PP030M.setValue("U_ResultYN", i, RecordSet01.Fields(5).VALUE)
            Call oDS_PS_PP030M.setValue("U_ReportYN", i, RecordSet01.Fields(6).VALUE)
            Call oDS_PS_PP030M.setValue("U_WorkGbn", i, "10")
            If i = RecordSet01.RecordCount - 1 Then
                Call PS_PP030_AddMatrixRow02(i + 1) '//마지막행에 한줄추가
            End If
            RecordSet01.MoveNext
        Next
    Else
        '//기계공구류는 검사공정을 기본적으로 입력
        oDS_PS_PP030M.Offset = 0
        Call oDS_PS_PP030M.setValue("U_LineNum", 0, 1)
        Call oDS_PS_PP030M.setValue("U_Sequence", 0, 1)
        Call oDS_PS_PP030M.setValue("U_CpBCode", 0, "CP204")
        Call oDS_PS_PP030M.setValue("U_CpBName", 0, "검사")
        Call oDS_PS_PP030M.setValue("U_CpCode", 0, "CP20402")
        Call oDS_PS_PP030M.setValue("U_CpName", 0, "최종검사")
        Call oDS_PS_PP030M.setValue("U_Unit", 0, "")
        Call oDS_PS_PP030M.setValue("U_ReWorkYN", 0, "N")
        Call oDS_PS_PP030M.setValue("U_ResultYN", 0, "N")
        Call oDS_PS_PP030M.setValue("U_ReportYN", 0, "N")
        Call oDS_PS_PP030M.setValue("U_WorkGbn", 0, "10")
        Call PS_PP030_AddMatrixRow02(1) '//마지막행에 한줄추가
    End If
    oMat03.LoadFromDataSource
    oMat03.AutoResizeColumns
    oForm01.Update
        
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP030_MTX03_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP030_MTX03_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP030_MTX03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP030_DI_API() As Boolean
On Error GoTo PS_PP030_DI_API_Error
    PS_PP030_DI_API = True
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
            Call oDS_PS_PP030L.setValue("U_OPDNNo", i, ResultDocNum)
            Call oDS_PS_PP030L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        Next
    Else
        GoTo PS_PP030_DI_API_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_PP030_DI_API_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_PP030_DI_API = False
    Set oDIObject = Nothing
    Exit Function
PS_PP030_DI_API_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP030_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP030_DI_API = False
End Function

Private Sub PS_PP030_FormResize()
On Error GoTo PS_PP030_FormResize_Error
    
    '생산요청, 작지리스트(Option)
    oForm01.Items("Opt01").Left = 10
    
    '생산요청, 작지리스트(Matrix)
    oForm01.Items("Mat01").Top = 58
    oForm01.Items("Mat01").Height = oForm01.Height / 2 - 120
    oForm01.Items("Mat01").Left = oForm01.Items("Opt01").Left
    oForm01.Items("Mat01").Width = oForm01.Width - 30
    
    '작업구분(Label)
    oForm01.Items("9").Top = oForm01.Items("Mat01").Height + oForm01.Items("Mat01").Top + 5
    oForm01.Items("9").Left = oForm01.Items("Opt01").Left
    '작업구분(TextBox)
    oForm01.Items("OrdGbn").Top = oForm01.Items("9").Top
    oForm01.Items("OrdGbn").Left = oForm01.Items("9").Left + oForm01.Items("9").Width
    oForm01.Items("BasicGub").Top = oForm01.Items("9").Top
    oForm01.Items("BasicGub").Left = oForm01.Items("9").Left + oForm01.Items("9").Width + oForm01.Items("9").Width
    '제품코드(Label)
    oForm01.Items("17").Top = oForm01.Items("9").Top + oForm01.Items("9").Height + 1
    oForm01.Items("17").Left = oForm01.Items("9").Left
    '제품코드(Link)
    oForm01.Items("1000001").Top = oForm01.Items("17").Top + 1
    oForm01.Items("1000001").Left = oForm01.Items("17").Left + oForm01.Items("17").Width - 15
    
    '제품코드(TextBox)
    oForm01.Items("ItemCode").Top = oForm01.Items("17").Top
    oForm01.Items("ItemCode").Left = oForm01.Items("OrdGbn").Left
    '제품명(TextBox)
    oForm01.Items("ItemName").Top = oForm01.Items("ItemCode").Top
    oForm01.Items("ItemName").Left = oForm01.Items("ItemCode").Left + oForm01.Items("ItemCode").Width
    
    '기준일자(Label)
    oForm01.Items("14").Top = oForm01.Items("17").Top + oForm01.Items("17").Height + 1
    oForm01.Items("14").Left = oForm01.Items("17").Left
    '기준일자(TextBox)
    oForm01.Items("OrdMgNum").Top = oForm01.Items("14").Top
    oForm01.Items("OrdMgNum").Left = oForm01.Items("ItemCode").Left

    '작업지시번호(Label)
    oForm01.Items("67").Top = oForm01.Items("14").Top + oForm01.Items("14").Height + 1
    oForm01.Items("67").Left = oForm01.Items("14").Left
    '작업지시번호(TextBox)
    oForm01.Items("OrdNum").Top = oForm01.Items("67").Top
    oForm01.Items("OrdNum").Left = oForm01.Items("OrdMgNum").Left
    '작업지시번호(Sub)(TextBox)
    oForm01.Items("OrdSub1").Top = oForm01.Items("67").Top
    oForm01.Items("OrdSub1").Left = oForm01.Items("OrdNum").Left + oForm01.Items("OrdNum").Width
    oForm01.Items("OrdSub2").Top = oForm01.Items("67").Top
    oForm01.Items("OrdSub2").Left = oForm01.Items("OrdSub1").Left + oForm01.Items("OrdSub1").Width
    
    '지시,완료일자(Label)
    oForm01.Items("18").Top = oForm01.Items("67").Top + oForm01.Items("67").Height + 1
    oForm01.Items("18").Left = oForm01.Items("67").Left
    '지시일자(TextBox)
    oForm01.Items("DocDate").Top = oForm01.Items("18").Top
    oForm01.Items("DocDate").Left = oForm01.Items("OrdNum").Left
    '완료일자(TextBox)
    oForm01.Items("DueDate").Top = oForm01.Items("18").Top
    oForm01.Items("DueDate").Left = oForm01.Items("DocDate").Left + oForm01.Items("DocDate").Width
    
    '담당자(Label)
    oForm01.Items("15").Top = oForm01.Items("18").Top + oForm01.Items("18").Height + 1
    oForm01.Items("15").Left = oForm01.Items("18").Left
    '담당자(TextBox)
    oForm01.Items("CntcCode").Top = oForm01.Items("15").Top
    oForm01.Items("CntcCode").Left = oForm01.Items("DocDate").Left
    '담당자명(TextBox)
    oForm01.Items("CntcName").Top = oForm01.Items("15").Top
    oForm01.Items("CntcName").Left = oForm01.Items("CntcCode").Left + oForm01.Items("CntcCode").Width
    
    '수주번호(Label)
    oForm01.Items("13").Top = oForm01.Items("15").Top + oForm01.Items("15").Height + 1
    oForm01.Items("13").Left = oForm01.Items("15").Left
    '수주번호(TextBox)
    oForm01.Items("SjNum").Top = oForm01.Items("13").Top
    oForm01.Items("SjNum").Left = oForm01.Items("CntcCode").Left
    '수주라인(TextBox)
    oForm01.Items("SjLine").Top = oForm01.Items("13").Top
    oForm01.Items("SjLine").Left = oForm01.Items("SjNum").Left + oForm01.Items("SjNum").Width
    
    '수주LOT번호(Label)
    oForm01.Items("39").Top = oForm01.Items("13").Top + oForm01.Items("13").Height + 1
    oForm01.Items("39").Left = oForm01.Items("13").Left
    '수주LOT번호(TextBox)
    oForm01.Items("LotNo").Top = oForm01.Items("39").Top
    oForm01.Items("LotNo").Left = oForm01.Items("SjNum").Left
    
    '멀티작업구분(Label)
    oForm01.Items("1000005").Top = oForm01.Items("39").Top + oForm01.Items("39").Height + 1
    oForm01.Items("1000005").Left = oForm01.Items("39").Left
    '멀티작업구분1(TextBox)
    oForm01.Items("MulGbn1").Top = oForm01.Items("1000005").Top
    oForm01.Items("MulGbn1").Left = oForm01.Items("LotNo").Left
    '멀티작업구분2(TextBox)
    oForm01.Items("MulGbn2").Top = oForm01.Items("1000005").Top
    oForm01.Items("MulGbn2").Left = oForm01.Items("MulGbn1").Left + oForm01.Items("MulGbn1").Width
    '멀티작업구분3(TextBox)
    oForm01.Items("MulGbn3").Top = oForm01.Items("1000005").Top
    oForm01.Items("MulGbn3").Left = oForm01.Items("MulGbn2").Left + oForm01.Items("MulGbn2").Width
    
    '기준문서구분(Label)
    oForm01.Items("63").Top = oForm01.Items("1000005").Top + oForm01.Items("1000005").Height + 1
    oForm01.Items("63").Left = oForm01.Items("1000005").Left
    '기준문서구분(TextBox)
    oForm01.Items("BaseType").Top = oForm01.Items("63").Top
    oForm01.Items("BaseType").Left = oForm01.Items("MulGbn1").Left
    
    '기준문서번호(Label)
    oForm01.Items("65").Top = oForm01.Items("63").Top
    oForm01.Items("65").Left = oForm01.Items("BaseType").Left + oForm01.Items("BaseType").Width
    '기준문서번호(TextBox)
    oForm01.Items("BaseNum").Top = oForm01.Items("65").Top
    oForm01.Items("BaseNum").Left = oForm01.Items("65").Left + oForm01.Items("65").Width
    
    '투입자재(Option)
    oForm01.Items("Opt02").Top = oForm01.Items("63").Top + oForm01.Items("63").Height + 15
    oForm01.Items("Opt02").Left = oForm01.Items("63").Left
    
    '투입자재(Matrix)
    oForm01.Items("Mat02").Top = oForm01.Items("Opt02").Top + oForm01.Items("Opt02").Height + 1
    oForm01.Items("Mat02").Left = oForm01.Items("63").Left
    oForm01.Items("Mat02").Width = oForm01.Width / 2 - 25
    oForm01.Items("Mat02").Height = oForm01.Height - oForm01.Items("Mat02").Top - 60
    
    '문서번호(Label)
    oForm01.Items("11").Top = oForm01.Items("9").Top
    oForm01.Items("11").Left = 320
    '문서번호(TextBox)
    oForm01.Items("DocEntry").Top = oForm01.Items("9").Top
    oForm01.Items("DocEntry").Left = oForm01.Items("11").Left + oForm01.Items("11").Width
    
    '사업장(Label)
    oForm01.Items("1000002").Top = oForm01.Items("14").Top
    oForm01.Items("1000002").Left = 255
    '사업장(TextBox)
    oForm01.Items("BPLId").Top = oForm01.Items("14").Top
    oForm01.Items("BPLId").Left = 335
    
    '작번이름(Label)
    oForm01.Items("70").Top = oForm01.Items("1000002").Top + oForm01.Items("1000002").Height + 1
    oForm01.Items("70").Left = oForm01.Items("11").Left
    '작번이름(TextBox)
    oForm01.Items("JakMyung").Top = oForm01.Items("70").Top
    oForm01.Items("JakMyung").Left = oForm01.Items("70").Left + oForm01.Items("70").Width
    
    '작번규격,단위(Label)
    oForm01.Items("72").Top = oForm01.Items("70").Top + oForm01.Items("70").Height + 1
    oForm01.Items("72").Left = oForm01.Items("70").Left
    '작번규격(TextBox)
    oForm01.Items("JakSize").Top = oForm01.Items("72").Top
    oForm01.Items("JakSize").Left = oForm01.Items("72").Left + oForm01.Items("72").Width
    '작번단위(TextBox)
    oForm01.Items("JakUnit").Top = oForm01.Items("72").Top
    oForm01.Items("JakUnit").Left = oForm01.Items("JakSize").Left + oForm01.Items("JakSize").Width
    
    '요청수,중량(Label)
    oForm01.Items("42").Top = oForm01.Items("72").Top + oForm01.Items("72").Height + 1
    oForm01.Items("42").Left = oForm01.Items("72").Left
    '요청수, 중량
    oForm01.Items("ReqWt").Top = oForm01.Items("42").Top
    oForm01.Items("ReqWt").Left = oForm01.Items("42").Left + oForm01.Items("42").Width
    
    '지시수,중량(Label)
    oForm01.Items("40").Top = oForm01.Items("42").Top + oForm01.Items("42").Height + 1
    oForm01.Items("40").Left = oForm01.Items("42").Left
    '지시수,중량
    oForm01.Items("SelWt").Top = oForm01.Items("40").Top
    oForm01.Items("SelWt").Left = oForm01.Items("40").Left + oForm01.Items("40").Width
    
    '수주금액(Label)
    oForm01.Items("38").Top = oForm01.Items("40").Top + oForm01.Items("40").Height + 1
    oForm01.Items("38").Left = oForm01.Items("40").Left
    '수주금액
    oForm01.Items("SjPrice").Top = oForm01.Items("38").Top
    oForm01.Items("SjPrice").Left = oForm01.Items("38").Left + oForm01.Items("38").Width
    
    '문서상태(Label)
    oForm01.Items("79").Top = oForm01.Items("38").Top + oForm01.Items("38").Height + 1
    oForm01.Items("79").Left = oForm01.Items("38").Left + 65
    '문서상태(TextBox)
    oForm01.Items("Status").Top = oForm01.Items("79").Top
    oForm01.Items("Status").Left = oForm01.Items("79").Left + oForm01.Items("79").Width
    
    '취소여부(Label)
    oForm01.Items("71").Top = oForm01.Items("79").Top + oForm01.Items("79").Height + 1
    oForm01.Items("71").Left = oForm01.Items("79").Left
    '취소여부(TextBox)
    oForm01.Items("Canceled").Top = oForm01.Items("71").Top
    oForm01.Items("Canceled").Left = oForm01.Items("71").Left + oForm01.Items("71").Width
    
    '공정리스트(Option)
    oForm01.Items("Opt03").Top = oForm01.Items("9").Top
    oForm01.Items("Opt03").Left = oForm01.Width / 2
    
    '표준공수조회(BUTTON)
    oForm01.Items("btnWkSrch").Top = oForm01.Items("9").Top - 2
    oForm01.Items("btnWkSrch").Left = oForm01.Items("Opt03").Left + oForm01.Items("Opt03").Width + 3
    
    '품목별공수조회(BUTTON)
    oForm01.Items("btnItmSrch").Top = oForm01.Items("btnWkSrch").Top
    oForm01.Items("btnItmSrch").Left = oForm01.Items("btnWkSrch").Left + oForm01.Items("btnWkSrch").Width + 3
    
    '공정금액합계(Label)
    oForm01.Items("77").Top = oForm01.Items("9").Top
    oForm01.Items("77").Left = oForm01.Items("btnItmSrch").Left + oForm01.Items("btnItmSrch").Width + 5
    
    '공정금액합계(TextBox)
    oForm01.Items("Total").Top = oForm01.Items("9").Top
    oForm01.Items("Total").Left = oForm01.Items("77").Left + oForm01.Items("77").Width
    
    '공정리스트(Matrix)
    oForm01.Items("Mat03").Left = oForm01.Items("Opt03").Left
    oForm01.Items("Mat03").Top = oForm01.Items("9").Top + 18
    oForm01.Items("Mat03").Height = oForm01.Height - oForm01.Items("Mat03").Top - 60
    oForm01.Items("Mat03").Width = oForm01.Width - oForm01.Items("Mat03").Left - 20
    
    Exit Sub
PS_PP030_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP030_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP030_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_PP030_Validate_Error
    PS_PP030_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim Exist As Boolean
    
    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP030H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
        PS_PP030_Validate = False
        GoTo PS_PP030_Validate_Exit
    End If
    
    If ValidateType = "검사01" Then

    ElseIf ValidateType = "검사02" Then '//투입자재 매트릭스에 대한 검사
        '//삭제된 행을 찾아서 삭제가능성 검사
        Query01 = "SELECT PS_PP030L.DocEntry,PS_PP030L.LineId,PS_PP030L.U_ProcType FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030L] PS_PP030L ON PS_PP030H.DocEntry = PS_PP030L.DocEntry WHERE PS_PP030H.Canceled = 'N' AND PS_PP030L.DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        RecordSet01.DoQuery Query01
        For i = 0 To RecordSet01.RecordCount - 1
            Exist = False
            For j = 1 To oMat02.RowCount - 1
                If (oMat02.Columns("LineId").Cells(j).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
                Else
                    '//라인번호가 같고, 문서번호가 같으면 존재하는행
                    If Val(RecordSet01.Fields(0).VALUE) = Val(oForm01.Items("DocEntry").Specific.VALUE) And Val(RecordSet01.Fields(1).VALUE) = Val(oMat02.Columns("LineId").Cells(j).Specific.VALUE) Then
                        Exist = True
                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드,기계공구
                            '//DB상에는 청구이고 매트릭스의 조달방법이 잔재로 변경된경우 수정할수 없다.
                            If RecordSet01.Fields(2).VALUE = "10" And oMat02.Columns("ProcType").Cells(j).Specific.Selected.VALUE <> "10" Then
                                Call MDC_Com.MDC_GF_Message("구매요청이 청구에서 잔재,취소로 변경되었습니다. 수정할수 없습니다.", "W")
                                PS_PP030_Validate = False
                                GoTo PS_PP030_Validate_Exit
                            End If
                        End If
                    End If
                End If
            Next
            If Exist = False Then '//삭제된 행중 구매요청에 아직 존재하면 수정할수 없다.
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드,기계공구
                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" & RecordSet01.Fields(0).VALUE & "' AND U_PP030LNo = '" & RecordSet01.Fields(1).VALUE & "'", 0, 1)) > 0 Then
                        MDC_Com.MDC_GF_Message "삭제된행이 구매요청문서 입니다. 적용할수 없습니다.", "W"
                        PS_PP030_Validate = False
                        GoTo PS_PP030_Validate_Exit
                    End If
                End If
                '//삭제된 행중에 멀티,엔드베어링중 작업일보에 등록된 행이면 수정할수 없다.
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then
                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & RecordSet01.Fields(0).VALUE & "'", 0, 1)) > 0 Then
                        MDC_Com.MDC_GF_Message "작업일보 등록된 행입니다. 수정할수 없습니다.", "W"
                        PS_PP030_Validate = False
                        GoTo PS_PP030_Validate_Exit
                    End If
                End If
                '//휘팅,부품은 삭제되는데 제약이 없다.
            End If
            RecordSet01.MoveNext
        Next
        
        For i = 1 To oMat02.RowCount - 1
            If (oMat02.Columns("LineId").Cells(i).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
            Else
                '//기존에 있던 행중에 멀티,엔드베어링중 작업일보에 등록된 행이면 수정할수 없다.
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then
                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'", 0, 1)) > 0 Then
                        Query01 = "SELECT "
                        Query01 = Query01 & " PS_PP030L.U_ItemCode,"
                        Query01 = Query01 & " PS_PP030L.U_ItemName,"
                        Query01 = Query01 & " PS_PP030L.U_ItemGpCd,"
                        Query01 = Query01 & " PS_PP030L.U_Weight,"
                        Query01 = Query01 & " PS_PP030H.U_BPLId,"
                        Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP030L.U_DueDate,112),"
                        Query01 = Query01 & " PS_PP030L.U_CntcCode,"
                        Query01 = Query01 & " PS_PP030L.U_CntcName,"
                        Query01 = Query01 & " PS_PP030L.U_ProcType,"
                        Query01 = Query01 & " PS_PP030L.U_Comments"
                        Query01 = Query01 & " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030L] PS_PP030L ON PS_PP030H.DocEntry = PS_PP030L.DocEntry"
                        Query01 = Query01 & " WHERE PS_PP030H.DocEntry = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'"
                        Query01 = Query01 & " AND PS_PP030L.LineId = '" & Trim(oMat02.Columns("LineId").Cells(i).Specific.VALUE) & "'"
                        Query01 = Query01 & " AND PS_PP030H.Canceled = 'N'"
                        RecordSet01.DoQuery Query01
                        If (RecordSet01.Fields(0).VALUE = oMat02.Columns("ItemCode").Cells(i).Specific.VALUE _
                        And RecordSet01.Fields(1).VALUE = oMat02.Columns("ItemName").Cells(i).Specific.VALUE _
                        And RecordSet01.Fields(2).VALUE = oMat02.Columns("ItemGpCd").Cells(i).Specific.Selected.VALUE _
                        And Val(RecordSet01.Fields(3).VALUE) = Val(oMat02.Columns("Weight").Cells(i).Specific.VALUE) _
                        And RecordSet01.Fields(4).VALUE = oForm01.Items("BPLId").Specific.Selected.VALUE _
                        And RecordSet01.Fields(5).VALUE = oMat02.Columns("DueDate").Cells(i).Specific.VALUE _
                        And RecordSet01.Fields(6).VALUE = oMat02.Columns("CntcCode").Cells(i).Specific.VALUE _
                        And RecordSet01.Fields(7).VALUE = oMat02.Columns("CntcName").Cells(i).Specific.VALUE _
                        And RecordSet01.Fields(8).VALUE = oMat02.Columns("ProcType").Cells(i).Specific.Selected.VALUE _
                        And RecordSet01.Fields(9).VALUE = oMat02.Columns("Comments").Cells(i).Specific.VALUE _
                        ) Then
                        Else '//값이 변경된 행의경우
                            MDC_Com.MDC_GF_Message "작업일보가 등록된 행은 수정할수 없습니다.", "W"
                            PS_PP030_Validate = False
                            GoTo PS_PP030_Validate_Exit
                        End If
                    End If
                End If
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드,기계공구
                    If oMat02.Columns("ProcType").Cells(i).Specific.Selected.VALUE = "20" Then '//잔재인 행은 제외
                    ElseIf oMat02.Columns("ProcType").Cells(i).Specific.Selected.VALUE = "30" Then '//취소인 행은 제외
                    Else '//청구인행에 대해
                        If (MDC_PS_Common.GetValue("SELECT U_OKYN FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND U_PP030LNo = '" & Trim(oMat02.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) = "Y" Then
                            '//결재가 완료된 값중
                            Query01 = "SELECT "
                            Query01 = Query01 & " PS_PP030L.U_ItemCode,"
                            Query01 = Query01 & " PS_PP030L.U_ItemName,"
                            Query01 = Query01 & " PS_PP030L.U_ItemGpCd,"
                            Query01 = Query01 & " Round(PS_PP030L.U_Weight,2),"
                            Query01 = Query01 & " PS_PP030H.U_BPLId,"
                            Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP030L.U_DueDate,112),"
                            Query01 = Query01 & " PS_PP030L.U_CntcCode,"
                            Query01 = Query01 & " PS_PP030L.U_CntcName,"
                            Query01 = Query01 & " PS_PP030L.U_ProcType,"
                            Query01 = Query01 & " PS_PP030L.U_Comments"
                            Query01 = Query01 & " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030L] PS_PP030L ON PS_PP030H.DocEntry = PS_PP030L.DocEntry"
                            Query01 = Query01 & " WHERE PS_PP030H.DocEntry = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'"
                            Query01 = Query01 & " AND PS_PP030L.LineId = '" & Trim(oMat02.Columns("LineId").Cells(i).Specific.VALUE) & "'"
                            Query01 = Query01 & " AND PS_PP030H.Canceled = 'N'"
                            RecordSet01.DoQuery Query01
                            
                            If (RecordSet01.Fields(0).VALUE = oMat02.Columns("ItemCode").Cells(i).Specific.VALUE _
                            And RecordSet01.Fields(1).VALUE = oMat02.Columns("ItemName").Cells(i).Specific.VALUE _
                            And RecordSet01.Fields(2).VALUE = oMat02.Columns("ItemGpCd").Cells(i).Specific.Selected.VALUE _
                            And Val(RecordSet01.Fields(3).VALUE) = Val(oMat02.Columns("Weight").Cells(i).Specific.VALUE) _
                            And RecordSet01.Fields(4).VALUE = oForm01.Items("BPLId").Specific.Selected.VALUE _
                            And RecordSet01.Fields(5).VALUE = oMat02.Columns("DueDate").Cells(i).Specific.VALUE _
                            And RecordSet01.Fields(6).VALUE = oMat02.Columns("CntcCode").Cells(i).Specific.VALUE _
                            And RecordSet01.Fields(7).VALUE = oMat02.Columns("CntcName").Cells(i).Specific.VALUE _
                            And RecordSet01.Fields(8).VALUE = oMat02.Columns("ProcType").Cells(i).Specific.Selected.VALUE _
                            And RecordSet01.Fields(9).VALUE = oMat02.Columns("Comments").Cells(i).Specific.VALUE _
                            ) Then
                            Else '//값이 변경된 행의경우
                                MDC_Com.MDC_GF_Message "구매요청 결재가 완료된 행은 수정할수 없습니다.", "W"
                                PS_PP030_Validate = False
                                GoTo PS_PP030_Validate_Exit
                            End If
                        End If
                    End If
                End If
            End If
        Next
    ElseIf ValidateType = "검사03" Then '//공정 매트릭스에 대한 검사
        '//삭제된 행을 찾아서 삭제가능성 검사
        Query01 = "SELECT PS_PP030M.DocEntry,PS_PP030M.LineId,PS_PP030M.U_Sequence, PS_PP030M.U_WorkGbn FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030M.DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        RecordSet01.DoQuery Query01
        For i = 0 To RecordSet01.RecordCount - 1
            Exist = False
            For j = 1 To oMat03.RowCount - 1
                If (oMat03.Columns("LineId").Cells(j).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
                Else
                    '//라인번호가 같고, 문서번호가 같으면 존재하는행,시퀀스도 같아야 한다. 행을 삭제할경우 시퀀스가 변경될수 있기때문에.
                    If Val(RecordSet01.Fields(0).VALUE) = Val(oForm01.Items("DocEntry").Specific.VALUE) And Val(RecordSet01.Fields(1).VALUE) = Val(oMat03.Columns("LineId").Cells(j).Specific.VALUE) And Val(RecordSet01.Fields(2).VALUE) = Val(oMat03.Columns("Sequence").Cells(j).Specific.VALUE) Then
                        Exist = True
'                        If oForm01.Items("OrdGbn").Specific.Selected.Value = "101" Then '//휘팅
'                            '//DB상에는 외주이고 매트릭스의 조달방법이 외주가 아닌경우 수정할수 없다.
'                            If RecordSet01.Fields(3).Value = "30" And oMat03.Columns("WorkGbn").Cells(j).Specific.Selected.Value <> "30" Then
'                                Call MDC_Com.MDC_GF_Message("작업구분이 외주에서 자가,정밀로 변경되었습니다. 수정할수 없습니다.", "W")
'                                PS_PP030_Validate = False
'                                GoTo PS_PP030_Validate_Exit
'                            End If
'                        End If
                        ' 검사조건은 필요가 없을 듯 한데 왜 넣었을까?주석처리(2017.12.07 송명규)
'                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구
'                            '//DB상에는 외주이고 매트릭스의 조달방법이 외주가 아닌경우 수정할수 없다.
'                            If RecordSet01.Fields(3).VALUE = "30" And oMat03.Columns("WorkGbn").Cells(j).Specific.Selected.VALUE <> "30" Then
'                                Call MDC_Com.MDC_GF_Message("작업구분이 외주에서 자가,정밀로 변경되었습니다. 수정할수 없습니다.", "W")
'                                PS_PP030_Validate = False
'                                GoTo PS_PP030_Validate_Exit
'                            End If
'                        End If
                        '위 검사조건은 필요가 없을 듯 한데 왜 넣었을까?주석처리(2017.12.07 송명규)
                    End If
                End If
            Next
            If Exist = False Then '//삭제된 행중 작업일보에 등록된행
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040L.U_PP030MNo = '" & RecordSet01.Fields(1).VALUE & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "삭제된행이 작업일보 등록된 행입니다. 적용할수 없습니다.", "W"
'                    PS_PP030_Validate = False
'                    GoTo PS_PP030_Validate_Exit
'                End If
                '//삭제된행중에 외주반출등록된행
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then
                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" & RecordSet01.Fields(0).VALUE & "' AND PS_MM130L.U_PP030MNo = '" & RecordSet01.Fields(1).VALUE & "'", 0, 1)) > 0 Then
                        MDC_Com.MDC_GF_Message "삭제된행이 외주반출 등록된 행입니다. 적용할수 없습니다.", "W"
                        PS_PP030_Validate = False
                        GoTo PS_PP030_Validate_Exit
                    End If
                End If
                '//삭제된행중에 외주등록된행
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구
                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE PS_MM005H.U_OrdType in ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" & RecordSet01.Fields(0).VALUE & "-" & RecordSet01.Fields(1).VALUE & "'", 0, 1)) > 0 Then
                        MDC_Com.MDC_GF_Message "삭제된행이 외주청구 등록된 행입니다. 적용할수 없습니다.", "W"
                        PS_PP030_Validate = False
                        GoTo PS_PP030_Validate_Exit
                    End If
                End If
            End If
            RecordSet01.MoveNext
        Next
        
        For i = 1 To oMat03.RowCount - 1
            If (oMat03.Columns("LineId").Cells(i).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
            Else
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP040L.U_PP030MNo = '" & Trim(oMat03.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    '//작업일보등록된문서중에 수정이 된문서를 구함
                    Query01 = "SELECT "
                    Query01 = Query01 & " PS_PP030M.U_CpBCode,"
                    Query01 = Query01 & " PS_PP030M.U_CpCode,"
                    Query01 = Query01 & " PS_PP030M.U_ResultYN,"
                    Query01 = Query01 & " PS_PP030M.U_ReportYN"
                    Query01 = Query01 & " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry"
                    Query01 = Query01 & " WHERE PS_PP030H.DocEntry = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'"
                    Query01 = Query01 & " AND PS_PP030M.LineId = '" & Trim(oMat03.Columns("LineId").Cells(i).Specific.VALUE) & "'"
                    Query01 = Query01 & " AND PS_PP030H.Canceled = 'N'"
                    RecordSet01.DoQuery Query01
                    If (RecordSet01.Fields(0).VALUE = oMat03.Columns("CpBCode").Cells(i).Specific.VALUE _
                    And RecordSet01.Fields(1).VALUE = oMat03.Columns("CpCode").Cells(i).Specific.VALUE _
                    And RecordSet01.Fields(2).VALUE = oMat03.Columns("ResultYN").Cells(i).Specific.Selected.VALUE _
                    And RecordSet01.Fields(3).VALUE = oMat03.Columns("ReportYN").Cells(i).Specific.Selected.VALUE _
                    ) Then
                    Else '//값이 변경된 행의경우
                        MDC_Com.MDC_GF_Message "작업일보가 등록된 행은 수정할수 없습니다.", "W"
                        Call oMat03.SelectRow(i, True, False)
                        PS_PP030_Validate = False
                        GoTo PS_PP030_Validate_Exit
                    End If
                End If
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then
                    If oMat03.Columns("WorkGbn").Cells(i).Specific.Selected.VALUE = "10" Then '//자가인 행은 제외
                    ElseIf oMat03.Columns("WorkGbn").Cells(i).Specific.Selected.VALUE = "20" Then '//정밀인 행은 제외
                    Else '//외주
                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_MM130L.U_PP030MNo = '" & Trim(oMat03.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 Then
                            '//외주반출등록된문서중에 수정이 된문서를 구함
                            Query01 = "SELECT "
                            Query01 = Query01 & " PS_PP030M.U_CpBCode,"
                            Query01 = Query01 & " PS_PP030M.U_CpCode,"
                            Query01 = Query01 & " PS_PP030M.U_ResultYN,"
                            Query01 = Query01 & " PS_PP030M.U_ReportYN,"
                            Query01 = Query01 & " PS_PP030M.U_WorkGbn"
                            Query01 = Query01 & " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry"
                            Query01 = Query01 & " WHERE PS_PP030H.DocEntry = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'"
                            Query01 = Query01 & " AND PS_PP030M.LineId = '" & Trim(oMat03.Columns("LineId").Cells(i).Specific.VALUE) & "'"
                            Query01 = Query01 & " AND PS_PP030H.Canceled = 'N'"
                            RecordSet01.DoQuery Query01
                            If (RecordSet01.Fields(0).VALUE = oMat03.Columns("CpBCode").Cells(i).Specific.VALUE _
                            And RecordSet01.Fields(1).VALUE = oMat03.Columns("CpCode").Cells(i).Specific.VALUE _
                            And RecordSet01.Fields(2).VALUE = oMat03.Columns("ResultYN").Cells(i).Specific.Selected.VALUE _
                            And RecordSet01.Fields(3).VALUE = oMat03.Columns("ReportYN").Cells(i).Specific.Selected.VALUE _
                            And RecordSet01.Fields(4).VALUE = oMat03.Columns("WorkGbn").Cells(i).Specific.Selected.VALUE _
                            ) Then
                            Else '//값이 변경된 행의경우
                                MDC_Com.MDC_GF_Message "외주반출이 등록된 행은 수정할수 없습니다.", "W"
                                Call oMat03.SelectRow(i, True, False)
                                PS_PP030_Validate = False
                                GoTo PS_PP030_Validate_Exit
                            End If
                        End If
                    End If
                End If
                If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구일대
                    If oMat03.Columns("WorkGbn").Cells(i).Specific.Selected.VALUE = "10" Then '//자가인 행은 제외
                    ElseIf oMat03.Columns("WorkGbn").Cells(i).Specific.Selected.VALUE = "20" Then '//정밀인 행은 제외
                    Else '//외주
                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "-" & Trim(oMat03.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 Then
                            '//외주청구등록된문서중에 수정이 된문서를 구함
                            Query01 = "SELECT "
                            Query01 = Query01 & " PS_PP030M.U_CpBCode,"
                            Query01 = Query01 & " PS_PP030M.U_CpCode,"
                            Query01 = Query01 & " PS_PP030M.U_ResultYN,"
                            Query01 = Query01 & " PS_PP030M.U_ReportYN,"
                            Query01 = Query01 & " PS_PP030M.U_WorkGbn"
                            Query01 = Query01 & " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry"
                            Query01 = Query01 & " WHERE PS_PP030H.DocEntry = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'"
                            Query01 = Query01 & " AND PS_PP030M.LineId = '" & Trim(oMat03.Columns("LineId").Cells(i).Specific.VALUE) & "'"
                            Query01 = Query01 & " AND PS_PP030H.Canceled = 'N'"
                            RecordSet01.DoQuery Query01
                            If (RecordSet01.Fields(0).VALUE = oMat03.Columns("CpBCode").Cells(i).Specific.VALUE _
                            And RecordSet01.Fields(1).VALUE = oMat03.Columns("CpCode").Cells(i).Specific.VALUE _
                            And RecordSet01.Fields(2).VALUE = oMat03.Columns("ResultYN").Cells(i).Specific.Selected.VALUE _
                            And RecordSet01.Fields(3).VALUE = oMat03.Columns("ReportYN").Cells(i).Specific.Selected.VALUE _
                            And RecordSet01.Fields(4).VALUE = oMat03.Columns("WorkGbn").Cells(i).Specific.Selected.VALUE _
                            ) Then
                            Else '//값이 변경된 행의경우
                                MDC_Com.MDC_GF_Message "외주청구가 등록된 행은 수정할수 없습니다.", "W"
                                Call oMat03.SelectRow(i, True, False)
                                PS_PP030_Validate = False
                                GoTo PS_PP030_Validate_Exit
                            End If
                        End If
                    End If
                End If
            End If
        Next
    '//모든값의 변경에 대해 검사하여 변경이 되었을시 수정가능검사를 하여 체크한다.
    ElseIf ValidateType = "수정02" Then
        If (oMat02.Columns("LineId").Cells(oMat02Row02).Specific.VALUE = "") Then '//새로추가된 행인경우, 수정하여도 무방하다
        Else
            '//삭제된 행중에 멀티,엔드베어링중 작업일보에 등록된 행이면 수정할수 없다.
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "작업일보 등록된 행입니다. 수정할수 없습니다.", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드,기계공구
                If (MDC_PS_Common.GetValue("SELECT U_OKYN FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND U_PP030LNo = '" & Trim(oMat02.Columns("LineId").Cells(oMat02Row02).Specific.VALUE) & "'", 0, 1)) = "Y" Then
                    MDC_Com.MDC_GF_Message "구매요청 결재가 진행된 행입니다. 수정할수 없습니다.", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
        End If
    ElseIf ValidateType = "행삭제02" Then
        '//행삭제전 행삭제가능여부검사
        If (oMat02.Columns("LineId").Cells(oMat02Row02).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
        Else
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "작업일보 등록된 행입니다. 삭제할수 없습니다.", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드,기계공구
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND U_PP030LNo = '" & Trim(oMat02.Columns("LineId").Cells(oMat02Row02).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "구매요청된행 입니다. 삭제할수 없습니다.", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
        End If
    '//모든값의 변경에 대해 검사하여 변경이 되었을시 수정가능검사를 하여 체크한다.
    ElseIf ValidateType = "수정03" Then
        If (oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE = "") Then '//새로추가된 행인경우, 수정하여도 무방하다

        Else
            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP040L.U_PP030MNo = '" & Trim(oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE) & "'", 0, 1)) > 0 Then
                MDC_Com.MDC_GF_Message "작업일보 등록된 행 입니다. 수정할수 없습니다.", "W"
                PS_PP030_Validate = False
                GoTo PS_PP030_Validate_Exit
            End If
            '//삭제된행중에 외주반출등록된행
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_MM130L.U_PP030MNo = '" & Trim(oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "외주반출 등록된 행입니다. 수정할수 없습니다.", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
            '//삭제된행중에 외주청구등록된행
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구일때
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "-" & Trim(oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "외주청구 등록된 행입니다. 수정할수 없습니다1.", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
        End If
    ElseIf ValidateType = "행삭제03" Then
        '//행삭제전 행삭제가능여부검사
        If (oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
        Else
            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP040L.U_PP030MNo = '" & Trim(oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE) & "'", 0, 1)) > 0 Then
                MDC_Com.MDC_GF_Message "작업일보 등록된 행 입니다. 삭제할수 없습니다.", "W"
                PS_PP030_Validate = False
                GoTo PS_PP030_Validate_Exit
            End If
            '//삭제된행중에 외주반출등록된행
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_MM130L.U_PP030MNo = '" & Trim(oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "외주반출 등록된 행입니다. 수정할수 없습니다.", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
            '//삭제된행중에 외주청구등록된행
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구일때
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "-" & Trim(oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "외주청구 등록된 행입니다. 수정할수 없습니다.2", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
        End If
    ElseIf ValidateType = "행추가03" Then
        '//행추가전 행추가가능여부검사
        If (oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
        Else
            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP040L.U_PP030MNo = '" & Trim(oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE) & "'", 0, 1)) > 0 Then
                MDC_Com.MDC_GF_Message "작업일보 등록된 행 입니다. 행추가할수 없습니다.", "W"
                PS_PP030_Validate = False
                GoTo PS_PP030_Validate_Exit
            End If
            '//삭제된행중에 외주반출등록된행
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_MM130L.U_PP030MNo = '" & Trim(oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "외주반출 등록된 행입니다. 행추가할수 없습니다.", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
            '//삭제된행중에 외주청구등록된행
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구일때
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40) AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "-" & Trim(oMat03.Columns("LineId").Cells(oMat03Row03).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "외주청구 등록된 행입니다. 행추가할수 없습니다.", "W"
                    PS_PP030_Validate = False
                    GoTo PS_PP030_Validate_Exit
                End If
            End If
        End If
    ElseIf ValidateType = "취소" Then
        '//취소가능유무검사
        If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP030H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
            MDC_Com.MDC_GF_Message "이미취소된 문서 입니다. 취소할수 없습니다.", "W"
            PS_PP030_Validate = False
            GoTo PS_PP030_Validate_Exit
        End If
        
        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드,기계공구
            If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND U_OKYN = 'Y'", 0, 1) > 0 Then
                MDC_Com.MDC_GF_Message "구매요청 결재가 승인되었습니다. 취소할수 없습니다.", "W"
                PS_PP030_Validate = False
                GoTo PS_PP030_Validate_Exit
            End If
        End If
                    
        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) > 0 Then
            MDC_Com.MDC_GF_Message "작업일보가 등록되었습니다. 취소할수 없습니다.", "W"
            PS_PP030_Validate = False
            GoTo PS_PP030_Validate_Exit
        End If
        
        '//삭제된행중에 외주반출등록된행
        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then
            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'", 0, 1)) > 0 Then
                MDC_Com.MDC_GF_Message "외주반출 등록된 행입니다. 취소할수 없습니다.", "W"
                PS_PP030_Validate = False
                GoTo PS_PP030_Validate_Exit
            End If
        End If
        
        '//삭제된행중에 외주청구등록된행
        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구일때
            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND U_PP030HNo = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "'", 0, 1)) > 0 Then
                MDC_Com.MDC_GF_Message "외주청구 등록된 행입니다. 취소할수 없습니다.", "W"
                PS_PP030_Validate = False
                GoTo PS_PP030_Validate_Exit
            End If
        End If
        
    ElseIf ValidateType = "닫기" Then
    
        '//닫기가능유무검사
        If MDC_PS_Common.GetValue("SELECT Status FROM [@PS_PP030H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "C" Then
            MDC_Com.MDC_GF_Message "이미 닫기(종료) 처리된 문서 입니다. 닫기(종료) 처리할 수 없습니다.", "W"
            PS_PP030_Validate = False
            GoTo PS_PP030_Validate_Exit
        End If
        
        
        '재고가 존재하면 닫기(종료) 불가 기능 추가(2012.01.11 송명규 추가)
        Dim QueryString As String
        
        QueryString = "                     SELECT      SUM(A.InQty) - SUM(A.OutQty) AS [StockQty]"
        QueryString = QueryString & "  FROM       OINM AS A"
        QueryString = QueryString & "                 INNER JOIN"
        QueryString = QueryString & "                 OITM As B"
        QueryString = QueryString & "                     ON A.ItemCode = B.ItemCode"
        QueryString = QueryString & "  WHERE      B.U_ItmBsort IN ('105','106')"
        QueryString = QueryString & "                 AND A.ItemCode = '" & oForm01.Items("ItemCode").Specific.VALUE & "'"
        QueryString = QueryString & "  GROUP BY  A.ItemCode"
    
        If IIf((MDC_PS_Common.GetValue(QueryString, 0, 1)) = "", 0, (MDC_PS_Common.GetValue(QueryString, 0, 1))) > 0 Then
        
            MDC_Com.MDC_GF_Message "재고가 존재하는 작업지시입니다. 닫기(종료) 처리할 수 없습니다.", "W"
            PS_PP030_Validate = False
            GoTo PS_PP030_Validate_Exit
            
        End If
        
'        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then '//몰드,기계공구
'            If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND U_OKYN = 'Y'", 0, 1) > 0 Then
'                MDC_Com.MDC_GF_Message "구매요청 결재가 승인되었습니다. 닫기(종료) 처리할 수 없습니다.", "W"
'                PS_PP030_Validate = False
'                GoTo PS_PP030_Validate_Exit
'            End If
'        End If
                    
'        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) > 0 Then
'            MDC_Com.MDC_GF_Message "작업일보가 등록되었습니다. 닫기(종료) 처리할 수 없습니다.", "W"
'            PS_PP030_Validate = False
'            GoTo PS_PP030_Validate_Exit
'        End If
        
'        '//삭제된행중에 외주반출등록된행
'        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "외주반출 등록된 행입니다. 닫기(종료) 처리할 수 없습니다.", "W"
'                PS_PP030_Validate = False
'                GoTo PS_PP030_Validate_Exit
'            End If
'        End If
'
'        '//삭제된행중에 외주청구등록된행
'        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Then '//기계공구일때
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "외주청구 등록된 행입니다. 닫기(종료) 처리할 수 없습니다.", "W"
'                PS_PP030_Validate = False
'                GoTo PS_PP030_Validate_Exit
'            End If
'        End If
    
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_PP030_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_PP030_Validate_Error:
    PS_PP030_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_PP030_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP030_PurchaseRequest(ByVal oDocEntry02 As Long, ByVal oLineId02 As Long)
On Error GoTo PS_PP030_PurchaseRequest_Error
    '//구매요청
    '//조달방식이 청구이면 [@PS_MM005H] 에 추가, 구매요청의 결재(OKYN) 값이 Y로 변경된 경우 수정불가, 작지에서는 청구행에 대해 행삭제불가
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim ItemName As String
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Query02 As String
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
        
    Query01 = "SELECT "
    Query01 = Query01 & "'" & DocEntry & "',"
    Query01 = Query01 & "'" & DocEntry & "',"
    Query01 = Query01 & " PS_PP030L.U_ItemCode,"
    Query01 = Query01 & " PS_PP030L.U_ItemName,"
    Query01 = Query01 & " PS_PP030L.U_Weight,"
    Query01 = Query01 & " PS_PP030L.U_Weight,"
    Query01 = Query01 & " 0,"
    Query01 = Query01 & " 0,"
    Query01 = Query01 & " PS_PP030H.U_BPLId,"
    Query01 = Query01 & "'" & DocEntry & "',"
    Query01 = Query01 & " CONVERT(NVARCHAR,GETDATE(),112),"
    Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP030L.U_DueDate,112),"
    Query01 = Query01 & " PS_PP030L.U_CntcCode,"
    Query01 = Query01 & " PS_PP030L.U_CntcName,"
    Query01 = Query01 & " (SELECT dept FROM [OHEM] WHERE empID = PS_PP030L.U_CntcCode),"
    Query01 = Query01 & " '',"
    Query01 = Query01 & " 'N',"
    Query01 = Query01 & " 'Y',"
    Query01 = Query01 & " '10',"
    Query01 = Query01 & " PS_PP030L.U_Comments,"
    Query01 = Query01 & " 'N',"
    Query01 = Query01 & " '',"
    Query01 = Query01 & " '10',"
    Query01 = Query01 & " '',"
    Query01 = Query01 & " 'O',"
    Query01 = Query01 & " PS_PP030H.DocEntry,"
    Query01 = Query01 & " PS_PP030L.LineId,"
    Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030L.LineId),"
    '// 청구시 필드 추가 - 류영조
    Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP030L.U_CGDate,112) As CGDate,"
    Query01 = Query01 & " PS_PP030H.U_OrdNum + '-' + PS_PP030H.U_OrdSub1 + '-' + PS_PP030H.U_OrdSub2 As OrdNum"
    ''''''''''''''''''''''''''''''''''''''''''
    Query01 = Query01 & " FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030L] PS_PP030L ON PS_PP030H.DocEntry = PS_PP030L.DocEntry"
    Query01 = Query01 & " WHERE PS_PP030H.DocEntry = '" & oDocEntry02 & "'"
    Query01 = Query01 & " AND PS_PP030L.LineId = '" & oLineId02 & "'"
    Query01 = Query01 & " AND PS_PP030H.Canceled = 'N'"
    RecordSet01.DoQuery Query01
    
    ItemName = MDC_PS_Common.Make_ItemName(Trim(RecordSet01.Fields(3).VALUE))
    
    DocEntry = MDC_PS_Common.GetValue("SELECT CASE WHEN ISNULL(MAX(CONVERT(INT,DocEntry)),0) = 0 THEN LEFT(CONVERT(NVARCHAR,'" & RecordSet01.Fields("CGDate").VALUE & "',112),6) + '0001' ELSE ISNULL(MAX(CONVERT(INT,DocEntry)),0)+1 END FROM [@PS_MM005H] WHERE LEFT(CONVERT(NVARCHAR,'" & RecordSet01.Fields("CGDate").VALUE & "',112),6) = LEFT(DocEntry,6)")
    
    '//구매요청이 취소되면 안되고 삭제되어야 한다.. 삭제하면서 작업지시도 동시삭제, 단 작업지시에 행이 1개만 존재한다면 삭제할수 없다.
    Query02 = " SELECT COUNT(*)"
    Query02 = Query02 & " FROM [@PS_MM005H]"
    Query02 = Query02 & " WHERE U_OrdType = '10' AND U_PP030HNo = '" & oDocEntry02 & "'"
    Query02 = Query02 & " AND U_PP030LNo = '" & oLineId02 & "'"
    'Query01 = Query01 & " AND Canceled = 'N'"
    RecordSet02.DoQuery Query02
    If RecordSet02.Fields(0).VALUE = 0 Then
        Query01 = "INSERT INTO [@PS_MM005H]"
        Query01 = Query01 & " ("
        Query01 = Query01 & " DocEntry,"
        Query01 = Query01 & " DocNum,"
        Query01 = Query01 & " U_ItemCode,"
        Query01 = Query01 & " U_ItemName,"
        Query01 = Query01 & " U_Qty,"
        Query01 = Query01 & " U_Weight,"
'        Query01 = Query01 & " U_Price,"
'        Query01 = Query01 & " U_LinTotal,"
        Query01 = Query01 & " U_BPLId,"
        Query01 = Query01 & " U_CgNum,"
        Query01 = Query01 & " U_DocDate,"
        Query01 = Query01 & " U_DueDate,"
        Query01 = Query01 & " U_CntcCode,"
        Query01 = Query01 & " U_CntcName,"
        Query01 = Query01 & " U_DeptCode,"
        Query01 = Query01 & " U_UseDept,"
        Query01 = Query01 & " U_Auto,"
        Query01 = Query01 & " U_QCYN,"
        'Query01 = Query01 & " U_ReType,"
''''''''        Query01 = Query01 & " U_Note,"
        Query01 = Query01 & " U_OKYN,"
        Query01 = Query01 & " U_OKDate,"
        Query01 = Query01 & " U_OrdType,"
        Query01 = Query01 & " U_ProcCode,"
        Query01 = Query01 & " U_Status,"
        '// 청구시 필드 추가 - 류영조
'        Query01 = Query01 & " U_DocDate,"
'        Query01 = Query01 & " U_DueDate,"
        Query01 = Query01 & " U_Comments,"
        Query01 = Query01 & " U_OrdNum,"
        '''''''''''''''''''''''''''''''''''
        Query01 = Query01 & " U_PP030HNo,"
        Query01 = Query01 & " U_PP030LNo,"
        Query01 = Query01 & " U_PP030DL,"
        Query01 = Query01 & " CreateDate" '생성일 추가(2014.02.24 송명규)
        Query01 = Query01 & " ) "
        Query01 = Query01 & "VALUES("
        Query01 = Query01 & "'" & DocEntry & "',"  '//DocEntry
        Query01 = Query01 & "'" & DocEntry & "',"  '//DocNum
        Query01 = Query01 & "'" & RecordSet01.Fields(2).VALUE & "'," '//ItemCode
        Query01 = Query01 & "'" & ItemName & "'," '//ItemName
        Query01 = Query01 & "" & RecordSet01.Fields(4).VALUE & "," '//Qty
        Query01 = Query01 & "" & RecordSet01.Fields(5).VALUE & "," '//Weight
'        Query01 = Query01 & "'" & RecordSet01.Fields(6).Value & "'," '//Price
'        Query01 = Query01 & "'" & RecordSet01.Fields(7).Value & "'," '//LineTotal
        Query01 = Query01 & "'" & RecordSet01.Fields(8).VALUE & "'," '//BPLId
        Query01 = Query01 & "'" & DocEntry & "',"  '//CgNum  'RecordSet01.Fields(9).Value
        Query01 = Query01 & "'" & RecordSet01.Fields("CGDate").VALUE & "'," '//DocDate
        Query01 = Query01 & "'" & RecordSet01.Fields(11).VALUE & "'," '//DueDate
        Query01 = Query01 & "'" & RecordSet01.Fields(12).VALUE & "'," '//CntcCode
        Query01 = Query01 & "'" & RecordSet01.Fields(13).VALUE & "'," '//CntcName
        Query01 = Query01 & "'" & RecordSet01.Fields(14).VALUE & "'," '//DeptCode
        Query01 = Query01 & "'" & RecordSet01.Fields(15).VALUE & "'," '//UseDept
        Query01 = Query01 & "'" & RecordSet01.Fields(16).VALUE & "'," '//Auto
        Query01 = Query01 & "'" & RecordSet01.Fields(17).VALUE & "',"  '//QCYN
        'Query01 = Query01 & "'" & RecordSet01.Fields(18).Value & "'," '//ReType
''''''        Query01 = Query01 & "'" & RecordSet01.Fields(19).Value & "'," '//Note
        Query01 = Query01 & "'" & RecordSet01.Fields(20).VALUE & "'," 'OKYN
        Query01 = Query01 & "'" & RecordSet01.Fields(21).VALUE & "'," 'OKDate
        Query01 = Query01 & "'" & RecordSet01.Fields(22).VALUE & "'," 'OrdType
        Query01 = Query01 & "'" & RecordSet01.Fields(23).VALUE & "'," 'ProcCode
        Query01 = Query01 & "'" & RecordSet01.Fields(24).VALUE & "'," 'Status
        '// 청구시 필드 추가 - 류영조
'        Query01 = Query01 & "'" & RecordSet01.Fields("CGDate").Value & "'," 'U_DocDate
'        Query01 = Query01 & "'" & RecordSet01.Fields("CGDate").Value & "'," 'U_DueDate
        Query01 = Query01 & "'" & RecordSet01.Fields(19).VALUE & "'," 'U_Comments
        Query01 = Query01 & "'" & RecordSet01.Fields("OrdNum").VALUE & "'," 'U_OrdNum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Query01 = Query01 & "'" & RecordSet01.Fields(25).VALUE & "'," 'PP030HNo
        Query01 = Query01 & "'" & RecordSet01.Fields(26).VALUE & "',"  'PP030LNo
        Query01 = Query01 & "'" & RecordSet01.Fields(27).VALUE & "',"  'PP030DL
        Query01 = Query01 & " GETDATE()" '생성일 추가(2014.02.24 송명규)
        Query01 = Query01 & ")"
        RecordSet01.DoQuery Query01
    Else
        Query01 = "UPDATE [@PS_MM005H] SET"
        Query01 = Query01 & " U_ItemCode = '" & RecordSet01.Fields(2).VALUE & "',"
        Query01 = Query01 & " U_ItemName = '" & ItemName & "',"
        Query01 = Query01 & " U_Qty = " & RecordSet01.Fields(4).VALUE & ","
        Query01 = Query01 & " U_Weight = " & RecordSet01.Fields(5).VALUE & ","
'        Query01 = Query01 & " U_Price = '" & RecordSet01.Fields(6).Value & "',"
'        Query01 = Query01 & " U_LinTotal = '" & RecordSet01.Fields(7).Value & "',"
        Query01 = Query01 & " U_BPLId = '" & RecordSet01.Fields(8).VALUE & "',"
        Query01 = Query01 & " U_DocDate = '" & RecordSet01.Fields("CGDate").VALUE & "',"
        Query01 = Query01 & " U_DueDate = '" & RecordSet01.Fields(11).VALUE & "',"
        Query01 = Query01 & " U_CntcCode = '" & RecordSet01.Fields(12).VALUE & "',"
        Query01 = Query01 & " U_CntcName = '" & RecordSet01.Fields(13).VALUE & "',"
        Query01 = Query01 & " U_DeptCode = '" & RecordSet01.Fields(14).VALUE & "',"
        Query01 = Query01 & " U_UseDept = '" & RecordSet01.Fields(15).VALUE & "',"
        Query01 = Query01 & " U_Auto = '" & RecordSet01.Fields(16).VALUE & "',"
        Query01 = Query01 & " U_QCYN = '" & RecordSet01.Fields(17).VALUE & "',"
'        Query01 = Query01 & " U_ReType = '" & RecordSet01.Fields(18).Value & "',"
        Query01 = Query01 & " U_Note = '" & RecordSet01.Fields(19).VALUE & "',"
'        Query01 = Query01 & " U_OKYN = '" & RecordSet01.Fields(20).Value & "',"
'        Query01 = Query01 & " U_OKDate = '" & RecordSet01.Fields(21).Value & "',"
        Query01 = Query01 & " U_OrdType = '" & RecordSet01.Fields(22).VALUE & "',"
        Query01 = Query01 & " U_ProcCode = '" & RecordSet01.Fields(23).VALUE & "',"
        '// 청구시 필드 추가 - 류영조
'        Query01 = Query01 & " U_DocDate = '" & RecordSet01.Fields("CGDate").Value & "',"
'        Query01 = Query01 & " U_DueDate = '" & RecordSet01.Fields("CGDate").Value & "',"
        Query01 = Query01 & " U_Comments = '" & RecordSet01.Fields(19).VALUE & "',"
        Query01 = Query01 & " UpdateDate = GETDATE()" '수정일 추가(2014.02.24 송명규)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Query01 = Query01 & " WHERE U_OrdType = '10' And U_PP030HNo = '" & RecordSet01.Fields(25).VALUE & "'"
        Query01 = Query01 & " AND U_PP030LNo = '" & RecordSet01.Fields(26).VALUE & "'"
        RecordSet01.DoQuery Query01
    End If
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Exit Sub
PS_PP030_PurchaseRequest_Error:
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_PP030_PurchaseRequest_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP030_AutoCreateMultiGage() As Boolean
On Error GoTo PS_PP030_AutoCreateMultiGage_Error
    PS_PP030_AutoCreateMultiGage = True
    Dim i, j, h, s As Long
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Dim RecordSet03 As SAPbobsCOM.Recordset
    Dim Query01 As String
    Dim Query02 As String
    Dim Query03 As String
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set RecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CurrentDocEntry As Long
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    For i = 1 To oMat02.VisualRowCount '//투입자재의 수량만큼
        Query01 = "SELECT AutoKey FROM [ONNM] WHERE ObjectCode = 'PS_PP030'"
        Call RecordSet01.DoQuery(Query01) '//PS_PP031의 최종문서번호
        CurrentDocEntry = RecordSet01.Fields(0).VALUE
        Query01 = "UPDATE [ONNM] SET AutoKey = AutoKey +1 WHERE ObjectCode = 'PS_PP030'"
        Call RecordSet01.DoQuery(Query01) '//문서번호증가, 이전문서번호는 현재 유저가 선점
        
        Query01 = "INSERT INTO [@PS_PP030H] ("
        Query01 = Query01 & "DocEntry,"
        Query01 = Query01 & "DocNum,"
        Query01 = Query01 & "Period,"
        Query01 = Query01 & "Instance,"
        Query01 = Query01 & "Series,"
        Query01 = Query01 & "Handwrtten,"
        Query01 = Query01 & "Canceled,"
        Query01 = Query01 & "Object,"
        Query01 = Query01 & "LogInst,"
        Query01 = Query01 & "UserSign,"
        Query01 = Query01 & "Transfered,"
        Query01 = Query01 & "Status,"
        Query01 = Query01 & "CreateDate,"
        Query01 = Query01 & "CreateTime,"
        Query01 = Query01 & "UpdateDate,"
        Query01 = Query01 & "UpdateTime,"
        Query01 = Query01 & "DataSource,"
        Query01 = Query01 & "U_BaseType,"
        Query01 = Query01 & "U_BaseNum,"
        Query01 = Query01 & "U_OrdGbn,"
        Query01 = Query01 & "U_DocDate,"
        Query01 = Query01 & "U_DueDate,"
        Query01 = Query01 & "U_ItemCode,"
        Query01 = Query01 & "U_ItemName,"
        Query01 = Query01 & "U_CntcCode,"
        Query01 = Query01 & "U_CntcName,"
        Query01 = Query01 & "U_SjNum,"
        Query01 = Query01 & "U_SjLine,"
        Query01 = Query01 & "U_OrdMgNum,"
        Query01 = Query01 & "U_OrdNum,"
        Query01 = Query01 & "U_OrdSub1,"
        Query01 = Query01 & "U_OrdSub2,"
        Query01 = Query01 & "U_JakMyung,"
        Query01 = Query01 & "U_ReqWt,"
        Query01 = Query01 & "U_SelWt,"
        Query01 = Query01 & "U_LotNo,"
        Query01 = Query01 & "U_SjPrice,"
        Query01 = Query01 & "U_MulGbn1,"
        Query01 = Query01 & "U_MulGbn2,"
        Query01 = Query01 & "U_MulGbn3,"
        Query01 = Query01 & "U_Comments,"
        Query01 = Query01 & "U_BPLId,"
        Query01 = Query01 & "U_BasicGub"
        Query01 = Query01 & ")"
        Query01 = Query01 & " VALUES("
        Query01 = Query01 & "'" & CurrentDocEntry & "'" & ","
        Query01 = Query01 & "'" & CurrentDocEntry & "'" & ","
        Query01 = Query01 & "'11'" & ","
        Query01 = Query01 & "'0'" & ","
        Query01 = Query01 & "'-1'" & ","
        Query01 = Query01 & "'N'" & ","
        Query01 = Query01 & "'N'" & ","
        Query01 = Query01 & "'PS_PP030'" & ","
        Query01 = Query01 & "NULL" & ","
        Query01 = Query01 & "'" & Sbo_Company.UserSignature & "'" & ","
        Query01 = Query01 & "'N'" & ","
        Query01 = Query01 & "'O'" & "," '//Status
        Query01 = Query01 & "CONVERT(NVARCHAR,GETDATE(),112)" & ","
        Query01 = Query01 & "SUBSTRING(CONVERT(NVARCHAR,GETDATE(),108),1,2) + SUBSTRING(CONVERT(NVARCHAR,GETDATE(),108),4,2)" & ","
        Query01 = Query01 & "NULL" & "," '//UpdateDate
        Query01 = Query01 & "NULL" & "," '//UpdateTime
        Query01 = Query01 & "'I'" & "," '//DataSource
        Query01 = Query01 & "NULL" & "," '//BaseType
        Query01 = Query01 & "NULL" & "," '//BaseNum
        Query01 = Query01 & "'" & oForm01.Items("OrdGbn").Specific.Selected.VALUE & "'" & ","
        If oForm01.Items("DocDate").Specific.VALUE = "" Then
            Query01 = Query01 & "NULL" & ","
        Else
            Query01 = Query01 & "'" & oForm01.Items("DocDate").Specific.VALUE & "'" & ","
        End If
        If oForm01.Items("DueDate").Specific.VALUE = "" Then
            Query01 = Query01 & "NULL" & ","
        Else
            Query01 = Query01 & "'" & oForm01.Items("DueDate").Specific.VALUE & "'" & ","
        End If
        Query01 = Query01 & "'" & oForm01.Items("ItemCode").Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oForm01.Items("ItemName").Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oForm01.Items("CntcCode").Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oForm01.Items("CntcName").Specific.VALUE & "'" & ","
        Query01 = Query01 & "NULL" & "," '//SjNum
        Query01 = Query01 & "NULL" & "," '//SjLine
        Query01 = Query01 & "'" & oForm01.Items("OrdMgNum").Specific.VALUE & "'" & "," '//신규작업지시번호를 조회
        Query01 = Query01 & "'" & oForm01.Items("OrdMgNum").Specific.VALUE & MDC_PS_Common.GetValue("EXEC PS_PP030_01 '" & oForm01.Items("OrdMgNum").Specific.VALUE & "'") & "'" & ","
        Query01 = Query01 & "'" & "00" & "'" & ","
        Query01 = Query01 & "'" & "000" & "'" & ","
        Query01 = Query01 & "'" & oForm01.Items("JakMyung").Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oForm01.Items("ReqWt").Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("Weight").Cells(i).Specific.VALUE & "'" & "," '//투입자재의 중량으로 입력되어야함
        Query01 = Query01 & "NULL" & "," '//LotNo
        Query01 = Query01 & "'" & oForm01.Items("SjPrice").Specific.VALUE & "'" & ","
        If oForm01.Items("MulGbn1").Specific.Selected Is Nothing Then
            Query01 = Query01 & "'" & "" & "'" & ","
        Else
            Query01 = Query01 & "'" & Trim(oForm01.Items("MulGbn1").Specific.Selected.VALUE) & "'" & ","
        End If
        If oForm01.Items("MulGbn2").Specific.Selected Is Nothing Then
            Query01 = Query01 & "'" & "" & "'" & ","
        Else
            Query01 = Query01 & "'" & Trim(oForm01.Items("MulGbn2").Specific.Selected.VALUE) & "'" & ","
        End If
        If oForm01.Items("MulGbn3").Specific.Selected Is Nothing Then
            Query01 = Query01 & "'" & "" & "'" & ","
        Else
            Query01 = Query01 & "'" & Trim(oForm01.Items("MulGbn3").Specific.Selected.VALUE) & "'" & ","
        End If
        Query01 = Query01 & "'" & Trim(oForm01.Items("Comments").Specific.VALUE) & "'" & ","
        Query01 = Query01 & "'" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'" & ","
        Query01 = Query01 & "'" & oForm01.Items("BasicGub").Specific.Selected.VALUE & "'"
        Query01 = Query01 & ")"
        Call RecordSet01.DoQuery(Query01)
        
        Query01 = "INSERT INTO [@PS_PP030L] ("
        Query01 = Query01 & "DocEntry,"
        Query01 = Query01 & "LineId,"
        Query01 = Query01 & "VisOrder,"
        Query01 = Query01 & "Object,"
        Query01 = Query01 & "LogInst,"
        Query01 = Query01 & "U_LineNum,"
        Query01 = Query01 & "U_InputGbn,"
        Query01 = Query01 & "U_ItemCode,"
        Query01 = Query01 & "U_ItemName,"
        Query01 = Query01 & "U_ItemGpCd,"
        Query01 = Query01 & "U_Weight,"
        Query01 = Query01 & "U_DueDate,"
        Query01 = Query01 & "U_CntcCode,"
        Query01 = Query01 & "U_CntcName,"
        Query01 = Query01 & "U_ProcType,"
        Query01 = Query01 & "U_Comments,"
        Query01 = Query01 & "U_BatchNum,"
        Query01 = Query01 & "U_LineId"
        Query01 = Query01 & ")"
        Query01 = Query01 & " VALUES("
        Query01 = Query01 & "'" & CurrentDocEntry & "'" & ","
        Query01 = Query01 & "'" & i & "'" & ","
        Query01 = Query01 & "'" & i - 1 & "'" & ","
        Query01 = Query01 & "'PS_PP030'" & ","
        Query01 = Query01 & "NULL" & ","
        Query01 = Query01 & "'" & i & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("InputGbn").Cells(i).Specific.Selected.VALUE & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("ItemCode").Cells(i).Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("ItemName").Cells(i).Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("ItemGpCd").Cells(i).Specific.Selected.VALUE & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("Weight").Cells(i).Specific.VALUE & "'" & ","
        If oMat02.Columns("DueDate").Cells(i).Specific.VALUE = "" Then
            Query01 = Query01 & "NULL" & ","
        Else
            Query01 = Query01 & "'" & oMat02.Columns("DueDate").Cells(i).Specific.VALUE & "'" & ","
        End If
        Query01 = Query01 & "'" & oMat02.Columns("CntcCode").Cells(i).Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("CntcName").Cells(i).Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("ProcType").Cells(i).Specific.Selected.VALUE & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("Comments").Cells(i).Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & oMat02.Columns("BatchNum").Cells(i).Specific.VALUE & "'" & ","
        Query01 = Query01 & "'" & i & "'"
        Query01 = Query01 & ")"
        Call RecordSet01.DoQuery(Query01)
        
        For j = 1 To oMat03.VisualRowCount
            Query01 = "INSERT INTO [@PS_PP030M] ("
            Query01 = Query01 & "DocEntry,"
            Query01 = Query01 & "LineId,"
            Query01 = Query01 & "VisOrder,"
            Query01 = Query01 & "Object,"
            Query01 = Query01 & "LogInst,"
            Query01 = Query01 & "U_LineNum,"
            Query01 = Query01 & "U_Sequence,"
            Query01 = Query01 & "U_CpBCode,"
            Query01 = Query01 & "U_CpBName,"
            Query01 = Query01 & "U_CpCode,"
            Query01 = Query01 & "U_CpName,"
            Query01 = Query01 & "U_StdHour,"
            Query01 = Query01 & "U_Unit,"
            Query01 = Query01 & "U_ReDate,"
            Query01 = Query01 & "U_WorkGbn,"
            Query01 = Query01 & "U_ReWorkYN,"
            Query01 = Query01 & "U_ResultYN,"
            Query01 = Query01 & "U_ReportYN,"
            Query01 = Query01 & "U_LineId"
            Query01 = Query01 & ")"
            Query01 = Query01 & " VALUES("
            Query01 = Query01 & "'" & CurrentDocEntry & "'" & ","
            Query01 = Query01 & "'" & j & "'" & ","
            Query01 = Query01 & "'" & j - 1 & "'" & ","
            Query01 = Query01 & "'PS_PP030'" & ","
            Query01 = Query01 & "NULL" & ","
            Query01 = Query01 & "'" & j & "'" & ","
            Query01 = Query01 & "'" & j & "'" & ","
            Query01 = Query01 & "'" & oMat03.Columns("CpBCode").Cells(j).Specific.VALUE & "'" & ","
            Query01 = Query01 & "'" & oMat03.Columns("CpBName").Cells(j).Specific.VALUE & "'" & ","
            Query01 = Query01 & "'" & oMat03.Columns("CpCode").Cells(j).Specific.VALUE & "'" & ","
            Query01 = Query01 & "'" & oMat03.Columns("CpName").Cells(j).Specific.VALUE & "'" & ","
            Query01 = Query01 & "'" & oMat03.Columns("StdHour").Cells(j).Specific.VALUE & "'" & ","
            Query01 = Query01 & "'" & oMat03.Columns("Unit").Cells(j).Specific.VALUE & "'" & ","
            If oMat03.Columns("ReDate").Cells(j).Specific.VALUE = "" Then
                Query01 = Query01 & "NULL" & ","
            Else
                Query01 = Query01 & "'" & oMat03.Columns("ReDate").Cells(j).Specific.VALUE & "'" & ","
            End If
            Query01 = Query01 & "'" & oMat03.Columns("WorkGbn").Cells(j).Specific.Selected.VALUE & "'" & ","
            Query01 = Query01 & "'" & oMat03.Columns("ReWorkYN").Cells(j).Specific.Selected.VALUE & "'" & ","
            Query01 = Query01 & "'" & oMat03.Columns("ResultYN").Cells(j).Specific.Selected.VALUE & "'" & ","
            Query01 = Query01 & "'" & oMat03.Columns("ReportYN").Cells(j).Specific.Selected.VALUE & "'" & ","
            Query01 = Query01 & "'" & j & "'"
            Query01 = Query01 & ")"
            Call RecordSet01.DoQuery(Query01)
        Next j
    Next i
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Set RecordSet03 = Nothing
    Exit Function
PS_PP030_AutoCreateMultiGage_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    PS_PP030_AutoCreateMultiGage = False
    Sbo_Application.SetStatusBarMessage "PS_PP030_AutoCreateMultiGage_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PS_PP030_CheckDate() As Boolean
'******************************************************************************
'Function ID : PS_PP030_CheckDate()
'해당모듈    : PS_PP030
'기능        : 선행프로세스와 일자 비교
'인수        : 없음
'반환값      : True-선행프로세스보다 일자가 같거나 느릴 경우, False-선행프로세스보다 일자가 빠를 경우
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP030_CheckDate_Error

    Dim Query01      As String
    Dim loopCount    As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BaseEntry    As String
    Dim BaseLine     As String
    Dim DocType      As String
    Dim CurDocDate   As String
    
    BaseEntry = Trim(oForm01.Items("BaseNum").Specific.VALUE)
    BaseLine = ""
    DocType = "PS_PP030"
    CurDocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)

    Query01 = "         EXEC PS_Z_CHECK_DATE '"
    Query01 = Query01 & BaseEntry & "','"
    Query01 = Query01 & BaseLine & "','"
    Query01 = Query01 & DocType & "','"
    Query01 = Query01 & CurDocDate & "'"

    Call oRecordSet01.DoQuery(Query01)

    If oRecordSet01.Fields("ReturnValue").VALUE = "False" Then
        PS_PP030_CheckDate = False
    Else
        PS_PP030_CheckDate = True
    End If
    
    Set oRecordSet01 = Nothing

    Exit Function
PS_PP030_CheckDate_Error:
    PS_PP030_CheckDate = False
    Call Sbo_Application.SetStatusBarMessage("PS_PP030_CheckDate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function
