VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_SD030.cls
'//  Module         : 판매관리 > 출하(선출)요청등록
'//  Desc           : 출하(선출)요청등록
'//  FormType       : PS_SD030
'//  Create Date    : 2010.08.20
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_SD030H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_SD030L As SAPbouiCOM.DBDataSource '등록라인

Private oDocType01 As String
Private oDoType01 As String
Public oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Public oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Public oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    Qty As Long '//수량
    Weight As Double '//중량
    Currency As String '//통화
    Price As Double '//단가
    LineTotal As Double '//총계
    WhsCode As String '//창고
    ORDRNum As Long '//판매오더문서
    RDR1Num As Long '//판매오더라인
    Check As Boolean
    ODLNNum As Long '//납품문서
    DLN1Num As Long '//납품라인
    ORDNNum As Long '//반품문서
    RDN1Num As Long '//반품라인
    SD030HNum As Long '//출하(선출)문서
    SD030LNum As Long '//출하(선출)라인
End Type
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String, Optional ByVal oFromDocType01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD030.srf")
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
    
    oFormUniqueID01 = "PS_SD030_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    oDocType01 = oFromDocType01
    oDoType01 = oFromDocType01
    Call PS_SD030_CreateItems
    Call PS_SD030_ComboBox_Setting
    Call PS_SD030_CF_ChooseFromList
    Call PS_SD030_EnableMenus
    Call PS_SD030_SetDocument(oFromDocEntry01)
    
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
                    If (PS_SD030_Validate("취소") = False) Then
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
                Call PS_SD030_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_SD030_FormItemEnabled '//UDO방식
                Call PS_SD030_AddMatrixRow(0, True) '//UDO방식
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
                    If (PS_SD030_FindValidateDocument("@PS_SD030H") = False) Then
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
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'                Call Sbo_Application.Menus.RemoveEx("MenuUID")
'        End If
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    Dim DocEntry As Long
    Dim i As Long
    If pval.BeforeAction = True Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_SD030_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_SD030_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                '//작지완료여부검사후 DI '//납품전기가능상황에서 납품(실적이 등록된것)
                If PS_SD030_ValidateDelivery = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                If PS_SD030_DI_API_01 = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                Call Sbo_Application.ActivateMenuItem("1281")
'                DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
'                oForm01.Mode = fm_FIND_MODE
'                oForm01.Items("DocEntry").Specific.Value = DocEntry
'                oForm01.Items("1").Click ct_Regular
            End If
        End If
        If pval.ItemUID = "Button04" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If oForm01.Items("ProgStat").Specific.Selected.VALUE = "3" Then '//납품상태이면
                    For i = 1 To oMat01.VisualRowCount - 1 '//AR송장처리된 문서존재유무검사
                        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [INV1] WHERE BaseType = '15' AND BaseEntry = '" & oMat01.Columns("ODLNNum").Cells(i).Specific.VALUE & "' AND BaseLine = '" & oMat01.Columns("DLN1Num").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
                            Call MDC_Com.MDC_GF_Message("AR송장처리된 문서가 존재합니다.", "W")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                    For i = 1 To oMat01.VisualRowCount - 1 '//반품처리된 문서존재유무검사
                        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [RDN1] WHERE BaseType = '15' AND BaseEntry = '" & oMat01.Columns("ODLNNum").Cells(i).Specific.VALUE & "' AND BaseLine = '" & oMat01.Columns("DLN1Num").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
                            Call MDC_Com.MDC_GF_Message("반품처리된 문서가 존재합니다.", "W")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                    If PS_SD030_DI_API_02 = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call Sbo_Application.ActivateMenuItem("1281")
                End If
'                DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
'                oForm01.Mode = fm_FIND_MODE
'                oForm01.Items("DocEntry").Specific.Value = DocEntry
'                oForm01.Items("1").Click ct_Regular
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_SD030_Print_Report01
            End If
        End If
        If pval.ItemUID = "Button03" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_SD030_Print_Report02
            End If
        End If
        If pval.ItemUID = "Button05" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_SD030_Print_Report03
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_SD030_FormItemEnabled
                    Call PS_SD030_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_SD030_FormItemEnabled
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "OrderNum") '//사용자값활성
        'Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "WhsCode") '//사용자값활성
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
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim ItemCode01 As String
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (PS_SD030_Validate("수정") = False) Then
                    Call oDS_PS_SD030L.setValue("U_" & pval.ColUID, pval.Row - 1, Trim(oDS_PS_SD030L.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                Else
                    If pval.ColUID = "OrderNum" Then
                        If oMat01.Columns("OrderNum").Cells(pval.Row).Specific.VALUE = "" Then
                            GoTo Raise_EVENT_VALIDATE_Exit
                        End If
                        For i = 1 To oMat01.RowCount
                            If pval.Row <> i Then '//현재 선택되어있는 행이 아니면
                                If (oMat01.Columns("OrderNum").Cells(pval.Row).Specific.VALUE = oMat01.Columns("OrderNum").Cells(i).Specific.VALUE) Then
                                    Call MDC_Com.MDC_GF_Message("동일한 수주가 존재합니다.", "W")
                                    oMat01.Columns("OrderNum").Cells(pval.Row).Specific.VALUE = ""
                                    GoTo Raise_EVENT_VALIDATE_Exit
                                End If
    '                            If (Mid(oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value, "-") - 1) <> _
    '                            Mid(oMat01.Columns("OrderNum").Cells(i).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(i).Specific.Value, "-") - 1)) Then
    '                                Call MDC_Com.MDC_GF_Message("동일하지않은 수주문서가 존재합니다.", "W")
    '                                oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value = ""
    '                                GoTo Raise_EVENT_VALIDATE_Exit
    '                            End If
                            End If
                        Next
                        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                        Query01 = "EXEC PS_SD030_01 '" & oMat01.Columns("OrderNum").Cells(pval.Row).Specific.VALUE & "','" & oForm01.Items("DocType").Specific.Selected.VALUE & "'"
                        Call RecordSet01.DoQuery(Query01)
                        For i = 0 To RecordSet01.RecordCount - 1
                            Call oDS_PS_SD030L.setValue("U_OrderNum", pval.Row - 1, RecordSet01.Fields("OrderNum").VALUE)
                            Call oDS_PS_SD030L.setValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                            Call oDS_PS_SD030L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
                            Call oDS_PS_SD030L.setValue("U_ItemGpCd", pval.Row - 1, RecordSet01.Fields("ItemGpCd").VALUE)
                            Call oDS_PS_SD030L.setValue("U_ItmBsort", pval.Row - 1, RecordSet01.Fields("ItmBsort").VALUE)
                            Call oDS_PS_SD030L.setValue("U_ItmMsort", pval.Row - 1, RecordSet01.Fields("ItmMsort").VALUE)
                            Call oDS_PS_SD030L.setValue("U_Unit1", pval.Row - 1, RecordSet01.Fields("Unit1").VALUE)
                            Call oDS_PS_SD030L.setValue("U_Size", pval.Row - 1, RecordSet01.Fields("Size").VALUE)
                            Call oDS_PS_SD030L.setValue("U_ItemType", pval.Row - 1, RecordSet01.Fields("ItemType").VALUE)
                            Call oDS_PS_SD030L.setValue("U_Quality", pval.Row - 1, RecordSet01.Fields("Quality").VALUE)
                            Call oDS_PS_SD030L.setValue("U_Mark", pval.Row - 1, RecordSet01.Fields("Mark").VALUE)
                            Call oDS_PS_SD030L.setValue("U_SbasUnit", pval.Row - 1, RecordSet01.Fields("SbasUnit").VALUE)
                            Call oDS_PS_SD030L.setValue("U_SjQty", pval.Row - 1, Val(RecordSet01.Fields("SjQty").VALUE))
                            Call oDS_PS_SD030L.setValue("U_SjWeight", pval.Row - 1, Val(RecordSet01.Fields("SjWeight").VALUE))
                            Call oDS_PS_SD030L.setValue("U_Qty", pval.Row - 1, Val(RecordSet01.Fields("Qty").VALUE))
                            Call oDS_PS_SD030L.setValue("U_UnWeight", pval.Row - 1, Val(RecordSet01.Fields("UnWeight").VALUE))
                            Call oDS_PS_SD030L.setValue("U_Weight", pval.Row - 1, Val(RecordSet01.Fields("Weight").VALUE))
                            Call oDS_PS_SD030L.setValue("U_Currency", pval.Row - 1, RecordSet01.Fields("Currency").VALUE)
                            Call oDS_PS_SD030L.setValue("U_Price", pval.Row - 1, Val(RecordSet01.Fields("Price").VALUE))
                            Call oDS_PS_SD030L.setValue("U_LinTotal", pval.Row - 1, Val(RecordSet01.Fields("LinTotal").VALUE))
                            Call oDS_PS_SD030L.setValue("U_WhsCode", pval.Row - 1, RecordSet01.Fields("WhsCode").VALUE)
                            Call oDS_PS_SD030L.setValue("U_WhsName", pval.Row - 1, RecordSet01.Fields("WhsName").VALUE)
                            Call oDS_PS_SD030L.setValue("U_Comments", pval.Row - 1, RecordSet01.Fields("Comments").VALUE)
                            Call oDS_PS_SD030L.setValue("U_TrType", pval.Row - 1, RecordSet01.Fields("TrType").VALUE)
                            Call oDS_PS_SD030L.setValue("U_ORDRNum", pval.Row - 1, RecordSet01.Fields("ORDRNum").VALUE)
                            Call oDS_PS_SD030L.setValue("U_RDR1Num", pval.Row - 1, RecordSet01.Fields("RDR1Num").VALUE)
                            Call oDS_PS_SD030L.setValue("U_Status", pval.Row - 1, RecordSet01.Fields("Status").VALUE)
                            Call oDS_PS_SD030L.setValue("U_LineId", pval.Row - 1, RecordSet01.Fields("LineId").VALUE)
                            RecordSet01.MoveNext
                        Next
                        If oMat01.RowCount = pval.Row And Trim(oDS_PS_SD030L.GetValue("U_OrderNum", pval.Row - 1)) <> "" Then
                            PS_SD030_AddMatrixRow (pval.Row)
                        End If
                        oMat01.LoadFromDataSource
                        oMat01.AutoResizeColumns
                        
                        If (oMat01.RowCount > 1) Then
                            oForm01.Items("Focus").Click ct_Regular
                            oForm01.Items("CardCode").Enabled = False
                            oForm01.Items("BPLId").Enabled = False
                            oForm01.Items("TrType").Enabled = False
                        Else
                            oForm01.Items("Focus").Click ct_Regular
                            oForm01.Items("CardCode").Enabled = True
                            oForm01.Items("BPLId").Enabled = True
                            oForm01.Items("TrType").Enabled = True
                        End If
                        oForm01.Update
                        Set RecordSet01 = Nothing
                    ElseIf pval.ColUID = "Qty" Then
                        If (Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0) Then
                            Call oDS_PS_SD030L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                            Call oDS_PS_SD030L.setValue("U_Weight", pval.Row - 1, 0)
                            Call oDS_PS_SD030L.setValue("U_LinTotal", pval.Row - 1, 0)
                        Else
                            ItemCode01 = Trim(oDS_PS_SD030L.GetValue("U_ItemCode", pval.Row - 1))
                            If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then '//EA자체품
                                Call oDS_PS_SD030L.setValue("U_Weight", pval.Row - 1, Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE))
                            ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                                Call oDS_PS_SD030L.setValue("U_Weight", pval.Row - 1, Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01)))
                            ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                                Call oDS_PS_SD030L.setValue("U_Weight", pval.Row - 1, (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE))
                            ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                                Call oDS_PS_SD030L.setValue("U_Weight", pval.Row - 1, Round(Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0))
                            ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG선택
                            End If
                            Call oDS_PS_SD030L.setValue("U_LinTotal", pval.Row - 1, Trim(oDS_PS_SD030L.GetValue("U_Weight", pval.Row - 1)) * Trim(oDS_PS_SD030L.GetValue("U_Price", pval.Row - 1)))
                            Call oDS_PS_SD030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        End If
                    ElseIf pval.ColUID = "Weight" Then
                        If (Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0) Then
                            Call oDS_PS_SD030L.setValue("U_Qty", pval.Row - 1, 0)
                            Call oDS_PS_SD030L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                            Call oDS_PS_SD030L.setValue("U_LinTotal", pval.Row - 1, 0)
                        Else
                            ItemCode01 = Trim(oDS_PS_SD030L.GetValue("U_ItemCode", pval.Row - 1))
                            If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then
                                Call oDS_PS_SD030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                                Call oDS_PS_SD030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                                Call oDS_PS_SD030L.setValue("U_Weight", pval.Row - 1, (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE))
                            ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                                Call oDS_PS_SD030L.setValue("U_Weight", pval.Row - 1, Round(Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0))
                            ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG선택
                                Call oDS_PS_SD030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            End If
                            Call oDS_PS_SD030L.setValue("U_LinTotal", pval.Row - 1, Trim(oDS_PS_SD030L.GetValue("U_Weight", pval.Row - 1)) * Trim(oDS_PS_SD030L.GetValue("U_Price", pval.Row - 1)))
                        End If
'                    ElseIf pval.ColUID = "WhsCode" Then
'                        Call oDS_PS_SD030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                        Call oDS_PS_SD030L.setValue("U_WhsName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT WhsName FROM [OWHS] WHERE WhsCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value & "'", 0, 1))
                    Else
                        Call oDS_PS_SD030L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    End If
                End If
                oMat01.LoadFromDataSource
                oMat01.AutoResizeColumns
                oForm01.Update
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_SD030H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "CntcCode") Then
                    Call oDS_PS_SD030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_SD030H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                Else
                    Call oDS_PS_SD030H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
            End If
        End If
    ElseIf pval.BeforeAction = False Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Exit:
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
        Call PS_SD030_FormItemEnabled
        Call PS_SD030_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    Dim oDataTable01 As SAPbouiCOM.DataTable
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_SD030H", "U_CardCode,U_CardName")
        End If
        If (pval.ItemUID = "DCardCod" Or pval.ItemUID = "DCardNam") Then
            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_SD030H", "U_DCardCod,U_DCardNam")
        End If
        If (pval.ItemUID = "Mat01") Then
            If (pval.ColUID = "WhsCode") Then
                If pval.SelectedObjects Is Nothing Then
                Else
                    Set oDataTable01 = pval.SelectedObjects
                    Call oDS_PS_SD030L.setValue("U_WhsCode", pval.Row - 1, oDataTable01.Columns("WhsCode").Cells(0).VALUE)
                    Call oDS_PS_SD030L.setValue("U_WhsName", pval.Row - 1, oDataTable01.Columns("WhsName").Cells(0).VALUE)
                    Set oDataTable01 = Nothing
                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
                    oMat01.LoadFromDataSource
                    oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                End If
            End If
        End If
        If (pval.ItemUID = "DCardCod" Or pval.ItemUID = "DCardNam") Then
            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_SD030H", "U_DCardCod,U_DCardNam")
        End If
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
            If (PS_SD030_Validate("행삭제") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_SD030L.RemoveRecord(oDS_PS_SD030L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_SD030_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_SD030L.GetValue("U_OrderNum", oMat01.RowCount - 1)) <> "" Then
                    Call PS_SD030_AddMatrixRow(oMat01.RowCount)
                End If
            End If
            If (oMat01.RowCount > 1) Then
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("TrType").Enabled = False
            Else
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("CardCode").Enabled = True
                oForm01.Items("BPLId").Enabled = True
                oForm01.Items("TrType").Enabled = True
            End If
            oForm01.Update
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
            If PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_SD030H") = False Then
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
            If PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_SD030H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    ElseIf pval.MenuUID = "1290" Then '//첫번째레코드로이동
        If pval.BeforeAction = True Then
            Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            Query01 = " SELECT TOP 1 DocEntry FROM [@PS_SD030H] ORDER BY DocEntry DESC" '//가장마지막행을 부여
            RecordSet01.DoQuery Query01
            DocEntry = Val(RecordSet01.Fields(0).VALUE) '//원본문서
            DocEntryNext = Val(RecordSet01.Fields(0).VALUE) '//다음문서
            Set RecordSet01 = Nothing
            If PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_SD030H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    ElseIf pval.MenuUID = "1291" Then '//마지막문서로이동
        If pval.BeforeAction = True Then
            Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            Query01 = " SELECT TOP 1 DocEntry FROM [@PS_SD030H] ORDER BY DocEntry ASC" '//가장 첫행을 부여
            RecordSet01.DoQuery Query01
            DocEntry = Val(RecordSet01.Fields(0).VALUE) '//원본문서
            DocEntryNext = Val(RecordSet01.Fields(0).VALUE) '//다음문서
            Set RecordSet01 = Nothing
            If PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_SD030H") = False Then
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


Private Function PS_SD030_CreateItems() As Boolean
On Error GoTo PS_SD030_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_SD030H = oForm01.DataSources.DBDataSources("@PS_SD030H")
    Set oDS_PS_SD030L = oForm01.DataSources.DBDataSources("@PS_SD030L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
'    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
'    Call oForm01.DataSources.UserDataSources.Add("WhsCode", dt_SHORT_TEXT, 100)
'    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
'    Call oForm01.Items("WhsCode").Specific.DataBind.SetBound(True, "", "WhsCode")
    
    If (oDocType01 = "출하요청") Then
        oForm01.Title = "출하요청[PS_SD030]"
        Call oForm01.Items("DocType").Specific.Select("1", psk_ByValue)
    ElseIf (oDocType01 = "선출요청") Then
        oForm01.Title = "선출요청[PS_SD031]"
        Call oForm01.Items("DocType").Specific.Select("2", psk_ByValue)
    End If
    
    '//2010.12.06 추가
    '담당자
    oDS_PS_SD030H.setValue "U_CntcCode", 0, User_MSTCOD()
    Call oDS_PS_SD030H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items("CntcCode").Specific.VALUE & "'", 0, 1))
    '//--------------
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_SD030_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD030_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SD030_ComboBox_Setting()
On Error GoTo PS_SD030_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "DocType", "", "1", "출하요청")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "DocType", "", "2", "선출요청")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("DocType").Specific, "PS_PS_SD030", "DocType")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "TrType", "", "1", "일반")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "TrType", "", "2", "임가공")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("TrType").Specific, "PS_PS_SD030", "TrType")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "1", "출하요청")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "2", "선출요청")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "3", "납품")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "4", "반품")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("ProgStat").Specific, "PS_PS_SD030", "ProgStat")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD030", "Mat01", "Status", "O", "미결")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD030", "Mat01", "Status", "C", "완료")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Status"), "PS_SD030", "Mat01", "Status")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD030", "Mat01", "TrType", "1", "일반")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD030", "Mat01", "TrType", "2", "군납")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("TrType"), "PS_SD030", "Mat01", "TrType")
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD030", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PS_SD030", "ItemCode")
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemGpCd"), "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmBsort"), "SELECT Code,Name FROM [@PSH_ITMBSORT]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemType"), "SELECT Code,Name FROM [@PSH_SHAPE]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Quality"), "SELECT Code,Name FROM [@PSH_QUALITY]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Mark"), "SELECT Code,Name FROM [@PSH_MARK]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("SbasUnit"), "SELECT Code,Name FROM [@PSH_UOMORG]")
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD030_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD030_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD030_CF_ChooseFromList()
On Error GoTo PS_SD030_CF_ChooseFromList_Error
'    '//ChooseFromList 설정
'    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
'    Dim oCons               As SAPbouiCOM.Conditions
'    Dim oCon                As SAPbouiCOM.Condition
'    Dim oCFL                As SAPbouiCOM.ChooseFromList
'    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
'    Dim oEdit               As SAPbouiCOM.EditText
'    Dim oColumn             As SAPbouiCOM.Column
'
'    Set oEdit = oForm01.Items("DCardCod").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_BusinessPartner
'    oCFLCreationParams.uniqueID = "CFLCARDCODE"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)
'
'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLCARDCODE"
'    oEdit.ChooseFromListAlias = "CardCode"

    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLs01             As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCons01             As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCon01              As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFL01              As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oCFLCreationParams01  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oEdit01             As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

    Set oEdit = oForm01.Items("DCardCod").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = lf_BusinessPartner
    oCFLCreationParams.uniqueID = "CFLCARDCODE"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

    oEdit.ChooseFromListUID = "CFLCARDCODE"
    oEdit.ChooseFromListAlias = "CardCode"
    
    Set oCons = oCFL.GetConditions()
    Set oCon = oCons.Add()
    oCon.Alias = "CardType"
    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    oCon.CondVal = "C"
    oCFL.SetConditions oCons
    
    
    Set oEdit01 = oForm01.Items("CardCode").Specific
    Set oCFLs01 = oForm01.ChooseFromLists
    Set oCFLCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
    
    oCFLCreationParams01.ObjectType = lf_BusinessPartner
    oCFLCreationParams01.uniqueID = "CFLCARD2CODE"
    oCFLCreationParams01.MultiSelection = False
    Set oCFL01 = oCFLs01.Add(oCFLCreationParams01)

    oEdit01.ChooseFromListUID = "CFLCARD2CODE"
    oEdit01.ChooseFromListAlias = "CardCode"
    
    Set oCons01 = oCFL01.GetConditions()
    Set oCon01 = oCons01.Add()
    oCon01.Alias = "CardType"
    oCon01.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    oCon01.CondVal = "C"
    oCFL01.SetConditions oCons01
    

    Set oColumn = oMat01.Columns("WhsCode")
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = lf_Warehouses
    oCFLCreationParams.uniqueID = "CFLWAREHOUSES"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oColumn.ChooseFromListUID = "CFLWAREHOUSES"
    oColumn.ChooseFromListAlias = "WhsCode"
    Exit Sub
PS_SD030_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD030_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD030_FormItemEnabled()
On Error GoTo PS_SD030_FormItemEnabled_Error
    Dim i As Long
    Call oForm01.Freeze(True)
    oForm01.Items("Focus").Click ct_Regular
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        oForm01.Items("TranCard").Enabled = True
        oForm01.Items("TranCode").Enabled = True
        oForm01.Items("Destin").Enabled = True
        oForm01.Items("TranCost").Enabled = True
        oForm01.Items("Comments").Enabled = True
        oForm01.Items("TrType").Enabled = True
        oForm01.Items("Mat01").Enabled = True
        oMat01.AutoResizeColumns
        Call PS_SD030_FormClear '//UDO방식
'        Call oForm01.Items("BPLId").Specific.Select("1", psk_ByValue)
        Call oForm01.Items("BPLId").Specific.Select(User_BPLId(), psk_ByValue) '//2010.12.06 추가
        If (oDoType01 = "출하요청") Then
            oForm01.Items("DocType").Enabled = True
            Call oForm01.Items("DocType").Specific.Select("1", psk_ByValue)
            oForm01.Items("CardCode").Click ct_Regular
            oForm01.Items("DocType").Enabled = False
            oForm01.Items("Button01").Visible = False
            oForm01.Items("Button04").Visible = False
            oForm01.Items("ProgStat").Enabled = True
            Call oForm01.Items("ProgStat").Specific.Select("1", psk_ByValue)
            oForm01.Items("CardCode").Click ct_Regular
            oForm01.Items("ProgStat").Enabled = False
            oMat01.Columns("ODLNNum").Visible = False
            oMat01.Columns("DLN1Num").Visible = False
            oForm01.Update
        ElseIf (oDoType01 = "선출요청") Then
            oForm01.Items("DocType").Enabled = True
            Call oForm01.Items("DocType").Specific.Select("2", psk_ByValue)
            oForm01.Items("CardCode").Click ct_Regular
            oForm01.Items("DocType").Enabled = False
            oForm01.Items("BPLId").Enabled = True
            Call oForm01.Items("BPLId").Specific.Select("2", psk_ByValue)
            oForm01.Items("CardCode").Click ct_Regular
            oForm01.Items("BPLId").Enabled = False
            Call oForm01.Items("TrType").Specific.Select("1", psk_ByValue)
            oForm01.Items("CardCode").Click ct_Regular
            oForm01.Items("TrType").Enabled = False
            oForm01.Items("Button01").Visible = True
            oForm01.Items("Button04").Visible = True
             oForm01.Items("ProgStat").Enabled = True
            Call oForm01.Items("ProgStat").Specific.Select("2", psk_ByValue)
            oForm01.Items("CardCode").Click ct_Regular
            oForm01.Items("ProgStat").Enabled = False
            oMat01.Columns("ODLNNum").Visible = True
            oMat01.Columns("DLN1Num").Visible = True
            oForm01.Items("Button03").Visible = False
        End If
        Call oForm01.Items("TrType").Specific.Select("1", psk_ByValue)
        '//2010.12.06 추가
        '담당자
        oDS_PS_SD030H.setValue "U_CntcCode", 0, User_MSTCOD()
        Call oDS_PS_SD030H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items("CntcCode").Specific.VALUE & "'", 0, 1))
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("Button01").Enabled = False
        oForm01.Items("Button04").Enabled = False
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        oForm01.Items("DocDate").Specific.String = Format(Now, "yyyymmdd")
        oForm01.Items("DueDate").Specific.String = Format(Now, "yyyymmdd")
        
        oForm01.Items("14").Visible = False
        oForm01.Items("16").Visible = False
        oForm01.Items("18").Visible = False
        oForm01.Items("20").Visible = False
        oForm01.Items("TranCard").Visible = False
        oForm01.Items("TranCode").Visible = False
        oForm01.Items("Destin").Visible = False
        oForm01.Items("TranCost").Visible = False
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        oForm01.Items("TranCard").Enabled = True
        oForm01.Items("TranCode").Enabled = True
        oForm01.Items("Destin").Enabled = True
        oForm01.Items("TranCost").Enabled = True
        oForm01.Items("Comments").Enabled = True
        oForm01.Items("TrType").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        oMat01.AutoResizeColumns
        Call oForm01.EnableMenu("1281", False)
        Call oForm01.EnableMenu("1282", True)
        oForm01.Items("Button01").Enabled = False
        oForm01.Items("Button04").Enabled = False
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        
        oForm01.Items("14").Visible = False
        oForm01.Items("16").Visible = False
        oForm01.Items("18").Visible = False
        oForm01.Items("20").Visible = False
        oForm01.Items("TranCard").Visible = False
        oForm01.Items("TranCode").Visible = False
        oForm01.Items("Destin").Visible = False
        oForm01.Items("TranCost").Visible = False
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        '//출하요청일때
        If MDC_PS_Common.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" & Trim(oDS_PS_SD030H.GetValue("DocEntry", 0)) & "'", 0, 1) = "1" Then
            Dim Enabled As Boolean
            Enabled = False
            For i = 0 To oDS_PS_SD030L.Size - 1
                '//매트릭스의 중량과 납품문서의 중량을 비교
                If Val(Trim(oDS_PS_SD030L.GetValue("U_Weight", i))) > Val(MDC_PS_Common.GetValue("SELECT SUM(U_Weight) FROM [@PS_SD040L] WHERE U_SD030H = '" & Trim(oDS_PS_SD030H.GetValue("DocEntry", 0)) & "' AND U_SD030L = '" & Trim(oDS_PS_SD030L.GetValue("U_LineId", i)) & "'", 0, 1)) Then
                    Enabled = True
                End If
            Next
            If Enabled = False Then '//문서가 수정불가능한경우
                oForm01.Items("DocEntry").Enabled = False
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("CntcCode").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("DueDate").Enabled = False
                oForm01.Items("TranCard").Enabled = False
                oForm01.Items("TranCode").Enabled = False
                oForm01.Items("Destin").Enabled = False
                oForm01.Items("TranCost").Enabled = False
                oForm01.Items("Comments").Enabled = True
                oForm01.Items("Mat01").Enabled = False
                oForm01.Items("TrType").Enabled = False
                oMat01.AutoResizeColumns
                Call oForm01.EnableMenu("1281", True)
                Call oForm01.EnableMenu("1282", False)
                oForm01.Items("Button01").Enabled = False
                oForm01.Items("Button04").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("DueDate").Enabled = False
            Else '//문서가 수정가능한경우
                oForm01.Items("DocEntry").Enabled = False
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("CntcCode").Enabled = True
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("DueDate").Enabled = True
                oForm01.Items("TranCard").Enabled = True
                oForm01.Items("TranCode").Enabled = True
                oForm01.Items("Destin").Enabled = True
                oForm01.Items("TranCost").Enabled = True
                oForm01.Items("Comments").Enabled = True
                oForm01.Items("TrType").Enabled = False
                oForm01.Items("Mat01").Enabled = True
                oMat01.AutoResizeColumns
                Call oForm01.EnableMenu("1281", True)
                Call oForm01.EnableMenu("1282", False)
                oForm01.Items("Button01").Enabled = False
                oForm01.Items("Button04").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("DueDate").Enabled = False
            End If
            oForm01.Items("14").Visible = False
            oForm01.Items("16").Visible = False
            oForm01.Items("18").Visible = False
            oForm01.Items("20").Visible = False
            oForm01.Items("TranCard").Visible = False
            oForm01.Items("TranCode").Visible = False
            oForm01.Items("Destin").Visible = False
            oForm01.Items("TranCost").Visible = False
        '//선출요청일때
        ElseIf MDC_PS_Common.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" & Trim(oDS_PS_SD030H.GetValue("DocEntry", 0)) & "'", 0, 1) = "2" Then
            '//납품일때
            If MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" & Trim(oDS_PS_SD030H.GetValue("DocEntry", 0)) & "'", 0, 1) = "3" Then
                oForm01.Items("DocEntry").Enabled = False
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("CntcCode").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("DueDate").Enabled = False
                oForm01.Items("TranCard").Enabled = False
                oForm01.Items("TranCode").Enabled = False
                oForm01.Items("Destin").Enabled = False
                oForm01.Items("TranCost").Enabled = False
                oForm01.Items("Comments").Enabled = True
                oForm01.Items("Mat01").Enabled = False
                oForm01.Items("TrType").Enabled = False
                oMat01.AutoResizeColumns
                Call oForm01.EnableMenu("1281", True)
                Call oForm01.EnableMenu("1282", False)
                oForm01.Items("Button01").Enabled = False
                oForm01.Items("Button04").Enabled = True
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("DueDate").Enabled = False
                
                oForm01.Items("14").Visible = True
                oForm01.Items("16").Visible = True
                oForm01.Items("18").Visible = True
                oForm01.Items("20").Visible = True
                oForm01.Items("TranCard").Visible = True
                oForm01.Items("TranCode").Visible = True
                oForm01.Items("Destin").Visible = True
                oForm01.Items("TranCost").Visible = True
            Else '//납품이 아닐때
                oForm01.Items("DocEntry").Enabled = False
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("CntcCode").Enabled = True
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("DueDate").Enabled = True
                oForm01.Items("TranCard").Enabled = True
                oForm01.Items("TranCode").Enabled = True
                oForm01.Items("Destin").Enabled = True
                oForm01.Items("TranCost").Enabled = True
                oForm01.Items("Comments").Enabled = True
                oForm01.Items("TrType").Enabled = False
                oForm01.Items("Mat01").Enabled = True
                oMat01.AutoResizeColumns
                Call oForm01.EnableMenu("1281", True)
                Call oForm01.EnableMenu("1282", False)
                oForm01.Items("Button01").Enabled = True
                oForm01.Items("Button04").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("DueDate").Enabled = False
                
                oForm01.Items("14").Visible = False
                oForm01.Items("16").Visible = False
                oForm01.Items("18").Visible = False
                oForm01.Items("20").Visible = False
                oForm01.Items("TranCard").Visible = False
                oForm01.Items("TranCode").Visible = False
                oForm01.Items("Destin").Visible = False
                oForm01.Items("TranCost").Visible = False
            End If
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD030_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD030_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD030_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD030_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_SD030L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD030L.Offset = oRow
    oDS_PS_SD030L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD030_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD030_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD030_FormClear()
On Error GoTo PS_SD030_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD030'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SD030_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD030_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD030_EnableMenus()
On Error GoTo PS_SD030_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1293", True)
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, True, True, True, True, True, True, False, False, False, False, True, False) '//메뉴설정
    Exit Sub
PS_SD030_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD030_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD030_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_SD030_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_SD030_FormItemEnabled
        Call PS_SD030_AddMatrixRow(0, True) '//UDO방식일때
    Else
        If (MDC_PS_Common.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" & oFromDocEntry01 & "'", 0, 1) = "1") Then
            oForm01.Title = "출하요청[PS_SD030]"
            Call oForm01.Items("DocType").Specific.Select("1", psk_ByValue)
        ElseIf (MDC_PS_Common.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" & oFromDocEntry01 & "'", 0, 1) = "2") Then
            oForm01.Title = "선출요청[PS_SD031]"
            Call oForm01.Items("DocType").Specific.Select("2", psk_ByValue)
        End If
        oForm01.Mode = fm_FIND_MODE
        Call PS_SD030_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_SD030_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD030_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD030_DataValidCheck() As Boolean
On Error GoTo PS_SD030_DataValidCheck_Error
    PS_SD030_DataValidCheck = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "전기일은 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_SD030_DataValidCheck = False
        Exit Function
    End If
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_SD030_DataValidCheck = False
        Exit Function
    End If
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("OrderNum").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "수주는 필수입니다.", bmt_Short, True
            oMat01.Columns("OrderNum").Cells(i).Click ct_Regular
            PS_SD030_DataValidCheck = False
            Exit Function
        End If
        If (Val(oMat01.Columns("Weight").Cells(i).Specific.VALUE) <= 0) Then
            Sbo_Application.SetStatusBarMessage "중량(수량)은 필수입니다.", bmt_Short, True
            oMat01.Columns("Weight").Cells(i).Click ct_Regular
            PS_SD030_DataValidCheck = False
            Exit Function
        End If
        For j = i + 1 To oMat01.VisualRowCount - 1
            If (oMat01.Columns("OrderNum").Cells(i).Specific.VALUE = oMat01.Columns("OrderNum").Cells(j).Specific.VALUE) Then
                Sbo_Application.SetStatusBarMessage "동일한 수주가 존재합니다.", bmt_Short, True
                oMat01.Columns("OrderNum").Cells(i).Click ct_Regular
                PS_SD030_DataValidCheck = False
                Exit Function
            End If
'            If (Mid(oMat01.Columns("OrderNum").Cells(i).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(i).Specific.Value, "-") - 1) <> _
'            Mid(oMat01.Columns("OrderNum").Cells(j).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(j).Specific.Value, "-") - 1)) Then
'                Sbo_Application.SetStatusBarMessage "동일하지않은 수주문서가 존재합니다.", bmt_Short, True
'                oMat01.Columns("OrderNum").Cells(i).Click ct_Regular
'                PS_SD030_DataValidCheck = False
'                Exit Function
'            End If
        Next
    Next
    
    If (PS_SD030_Validate("검사") = False) Then
        PS_SD030_DataValidCheck = False
        Exit Function
    End If
    
    Call oDS_PS_SD030L.RemoveRecord(oDS_PS_SD030L.Size - 1)
    Call oMat01.LoadFromDataSource
    If oForm01.Mode = fm_ADD_MODE Then
        Call PS_SD030_FormClear
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_SD030_DataValidCheck_Error:
    PS_SD030_DataValidCheck = False
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_SD030_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD030_MTX01()
On Error GoTo PS_SD030_MTX01_Error
    '//메트릭스에 데이터 로드
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
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_SD030_MTX01_Exit
    End If
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_SD030L.InsertRecord (i)
        End If
        oDS_PS_SD030L.Offset = i
        Call oDS_PS_SD030L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_SD030L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
    Exit Sub
PS_SD030_MTX01_Exit:
    Set RecordSet01 = Nothing
    Exit Sub
PS_SD030_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_SD030_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SD030_DI_API_01() As Boolean '//납품
On Error GoTo PS_SD030_DI_API_01_Error
    PS_SD030_DI_API_01 = True
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
    For i = 1 To oMat01.VisualRowCount - 1
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Qty = oMat01.Columns("Qty").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Weight = oMat01.Columns("Weight").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Currency = oMat01.Columns("Currency").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Price = oMat01.Columns("Price").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).LineTotal = oMat01.Columns("LinTotal").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).ORDRNum = Val(oMat01.Columns("ORDRNum").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).RDR1Num = Val(oMat01.Columns("RDR1Num").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).SD030HNum = Val(oForm01.Items("DocEntry").Specific.VALUE)
        ItemInformation(ItemInformationCount).SD030LNum = Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE)
        '//ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    oDIObject.UserFields("U_DCardCod").VALUE = Trim(oForm01.Items("DCardCod").Specific.VALUE)
    oDIObject.UserFields("U_DCardNam").VALUE = Trim(oForm01.Items("DCardNam").Specific.VALUE)
    oDIObject.UserFields("U_TradeType").VALUE = Trim(oForm01.Items("TrType").Specific.Selected.VALUE)
    If oForm01.Items("DocDate").Specific.VALUE <> "" Then
        oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    If oForm01.Items("DueDate").Specific.VALUE <> "" Then
        oDIObject.DocDueDate = Format(oForm01.Items("DueDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.UserFields("U_BaseType").VALUE = "PS_SD030" '//별로 의미가 없을듯..
        oDIObject.Lines.BaseType = 17
        oDIObject.Lines.BaseEntry = ItemInformation(i).ORDRNum
        oDIObject.Lines.BaseLine = ItemInformation(i).RDR1Num
        
        oDIObject.Lines.UserFields("U_Qty").VALUE = ItemInformation(i).Qty
        oDIObject.Lines.Quantity = ItemInformation(i).Weight
        oDIObject.Lines.Currency = ItemInformation(i).Currency
        oDIObject.Lines.Price = ItemInformation(i).Price
        oDIObject.Lines.LineTotal = ItemInformation(i).LineTotal
        ItemInformation(i).DLN1Num = LineNumCount
        LineNumCount = LineNumCount + 1
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey '//문서상태 납품으로 변경
        MDC_PS_Common.DoQuery ("UPDATE [@PS_SD030H] SET U_ProgStat = '3' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'")
        For i = 0 To ItemInformationCount - 1 '//납품,반품문서번호 업데이트
            MDC_PS_Common.DoQuery ("UPDATE [@PS_SD030L] SET U_ODLNNum = '" & ResultDocNum & "', U_DLN1Num = '" & ItemInformation(i).DLN1Num & "', U_ORDNNum = '', U_RDN1Num = '' WHERE DocEntry = '" & ItemInformation(i).SD030HNum & "' AND LineId = '" & ItemInformation(i).SD030LNum & "'")
        Next
    Else
        GoTo PS_SD030_DI_API_01_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    Set oDIObject = Nothing
    Exit Function
PS_SD030_DI_API_01_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_SD030_DI_API_01 = False
    Set oDIObject = Nothing
    Exit Function
PS_SD030_DI_API_01_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_SD030_DI_API_01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_SD030_DI_API_01 = False
    Set oDIObject = Nothing
End Function

Private Function PS_SD030_DI_API_02() As Boolean '//반품
On Error GoTo PS_SD030_DI_API_02_Error
    PS_SD030_DI_API_02 = True
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
    For i = 1 To oMat01.VisualRowCount - 1
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Qty = oMat01.Columns("Qty").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Weight = oMat01.Columns("Weight").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Currency = oMat01.Columns("Currency").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Price = oMat01.Columns("Price").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).LineTotal = oMat01.Columns("LinTotal").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).ORDRNum = Val(oMat01.Columns("ORDRNum").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).RDR1Num = Val(oMat01.Columns("RDR1Num").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).SD030HNum = Val(oForm01.Items("DocEntry").Specific.VALUE)
        ItemInformation(ItemInformationCount).SD030LNum = Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).ODLNNum = Val(oMat01.Columns("ODLNNum").Cells(i).Specific.VALUE) '//납품
        ItemInformation(ItemInformationCount).DLN1Num = Val(oMat01.Columns("DLN1Num").Cells(i).Specific.VALUE) '//납품
        '//ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oReturns)
    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    oDIObject.UserFields("U_DCardCod").VALUE = Trim(oForm01.Items("DCardCod").Specific.VALUE)
    oDIObject.UserFields("U_DCardNam").VALUE = Trim(oForm01.Items("DCardNam").Specific.VALUE)
    oDIObject.UserFields("U_TradeType").VALUE = Trim(oForm01.Items("TrType").Specific.Selected.VALUE)
    If oForm01.Items("DocDate").Specific.VALUE <> "" Then
        oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    If oForm01.Items("DueDate").Specific.VALUE <> "" Then
        oDIObject.DocDueDate = Format(oForm01.Items("DueDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.UserFields("U_BaseType").VALUE = "PS_SD030" '//별로 의미가 없을듯..
        oDIObject.Lines.BaseType = 15
        oDIObject.Lines.BaseEntry = ItemInformation(i).ODLNNum
        oDIObject.Lines.BaseLine = ItemInformation(i).DLN1Num
        
        oDIObject.Lines.UserFields("U_Qty").VALUE = ItemInformation(i).Qty
        oDIObject.Lines.Quantity = ItemInformation(i).Weight
        oDIObject.Lines.Currency = ItemInformation(i).Currency
        oDIObject.Lines.Price = ItemInformation(i).Price
        oDIObject.Lines.LineTotal = ItemInformation(i).LineTotal
        ItemInformation(i).RDN1Num = LineNumCount
        LineNumCount = LineNumCount + 1
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey '//문서상태 반품으로 변경
        MDC_PS_Common.DoQuery ("UPDATE [@PS_SD030H] SET U_ProgStat = '4' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'")
        For i = 0 To ItemInformationCount - 1 '//납품,반품문서번호 업데이트
            MDC_PS_Common.DoQuery ("UPDATE [@PS_SD030L] SET U_ORDNNum = '" & ResultDocNum & "', U_RDN1Num = '" & ItemInformation(i).RDN1Num & "' WHERE DocEntry = '" & ItemInformation(i).SD030HNum & "' AND LineId = '" & ItemInformation(i).SD030LNum & "'")
        Next
    Else
        GoTo PS_SD030_DI_API_02_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    Set oDIObject = Nothing
    Exit Function
PS_SD030_DI_API_02_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_SD030_DI_API_02 = False
    Set oDIObject = Nothing
    Exit Function
PS_SD030_DI_API_02_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_SD030_DI_API_02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_SD030_DI_API_02 = False
    Set oDIObject = Nothing
End Function

Public Function PS_SD030_FindValidateDocument(ByVal ObjectType As String) As Boolean
On Error GoTo PS_SD030_FindValidateDocument_Error
    PS_SD030_FindValidateDocument = True
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
    If (oDocType01 = "출하요청") Then
        Query01 = Query01 & " AND U_DocType = '1'"
    ElseIf (oDocType01 = "선출요청") Then
        Query01 = Query01 & " AND U_DocType = '2'"
    End If
    RecordSet01.DoQuery Query01
    If (RecordSet01.RecordCount = 0) Then
        If (oDocType01 = "출하요청") Then
            Sbo_Application.SetStatusBarMessage "선출요청문서 이거나 존재하지 않는 문서입니다.", bmt_Short, True
        ElseIf (oDocType01 = "선출요청") Then
            Sbo_Application.SetStatusBarMessage "출하요청문서 이거나 존재하지 않는 문서입니다.", bmt_Short, True
        End If
        PS_SD030_FindValidateDocument = False
        Set RecordSet01 = Nothing
        Set RecordSet02 = Nothing
        Exit Function
    End If
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Exit Function
PS_SD030_FindValidateDocument_Error:
    Sbo_Application.SetStatusBarMessage Err.Number & Err.Description, bmt_Short, True
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_SD030_FindValidateDocument = False
End Function

Public Function PS_SD030_DirectionValidateDocument(ByVal DocEntry As String, ByVal DocEntryNext As String, ByVal Direction As String, ByVal ObjectType As String) As Boolean
On Error GoTo PS_SD030_DirectionValidateDocument_Error
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
                PS_SD030_DirectionValidateDocument = False
                Set RecordSet01 = Nothing
                Set RecordSet02 = Nothing
                Exit Function
            End If
        End If
        If (Direction = "Next") Then
            Query01 = " SELECT TOP 1 DocEntry"
            Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry > "
            Query01 = Query01 & DocEntryNext
            If (oDocType01 = "출하요청") Then
                Query01 = Query01 & " AND U_DocType = '1'"
            ElseIf (oDocType01 = "선출요청") Then
                Query01 = Query01 & " AND U_DocType = '2'"
            End If
            Query01 = Query01 & " ORDER BY DocEntry ASC"
        ElseIf (Direction = "Prev") Then
            Query01 = " SELECT TOP 1 DocEntry"
            Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry < "
            Query01 = Query01 & DocEntryNext
            If (oDocType01 = "출하요청") Then
                Query01 = Query01 & " AND U_DocType = '1'"
            ElseIf (oDocType01 = "선출요청") Then
                Query01 = Query01 & " AND U_DocType = '2'"
            End If
            Query01 = Query01 & " ORDER BY DocEntry DESC"
        End If
        RecordSet01.DoQuery Query01
        If (RecordSet01.Fields(0).VALUE = 0) Then '//해당문서가 마지막문서라면
            If (Direction = "Next") Then
                Query02 = " SELECT TOP 1 DocEntry FROM [" & ObjectType & "]"
                If (oDocType01 = "출하요청") Then
                    Query02 = Query02 & " WHERE U_DocType = '1'"
                ElseIf (oDocType01 = "선출요청") Then
                    Query02 = Query02 & " WHERE U_DocType = '2'"
                End If
                Query02 = Query02 & " ORDER BY DocEntry ASC"
            ElseIf (Direction = "Prev") Then
                Query02 = " SELECT TOP 1 DocEntry FROM [" & ObjectType & "]"
                If (oDocType01 = "출하요청") Then
                    Query02 = Query02 & " WHERE U_DocType = '1'"
                ElseIf (oDocType01 = "선출요청") Then
                    Query02 = Query02 & " WHERE U_DocType = '2'"
                End If
                Query02 = Query02 & " ORDER BY DocEntry DESC"
            End If
            RecordSet02.DoQuery Query02
            If (RecordSet02.RecordCount = 0) Then '//문서가 아예 존재하지 않는다면
                Sbo_Application.SetStatusBarMessage "유효한문서가 존재하지 않습니다", bmt_Short, True
                Set RecordSet01 = Nothing
                Set RecordSet02 = Nothing
                PS_SD030_DirectionValidateDocument = False
                Exit Function
            Else
                If (Direction = "Next") Then
                    DocEntryNext = Val(RecordSet02.Fields(0).VALUE) - 1
                    Query01 = " SELECT TOP 1 DocEntry"
                    Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry > "
                    Query01 = Query01 & DocEntryNext
                    If (oDocType01 = "출하요청") Then
                        Query01 = Query01 & " AND U_DocType = '1'"
                    ElseIf (oDocType01 = "선출요청") Then
                        Query01 = Query01 & " AND U_DocType = '2'"
                    End If
                    Query01 = Query01 & " ORDER BY DocEntry ASC"
                    RecordSet01.DoQuery Query01
                ElseIf (Direction = "Prev") Then
                    DocEntryNext = Val(RecordSet02.Fields(0).VALUE) + 1
                    Query01 = " SELECT TOP 1 DocNum"
                    Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry < "
                    Query01 = Query01 & DocEntryNext
                    If (oDocType01 = "출하요청") Then
                        Query01 = Query01 & " AND U_DocType = '1'"
                    ElseIf (oDocType01 = "선출요청") Then
                        Query01 = Query01 & " AND U_DocType = '2'"
                    End If
                    Query01 = Query01 & " ORDER BY DocEntry DESC"
                    RecordSet01.DoQuery Query01
                End If
            End If
        End If
        If (oDocType01 = "출하요청") Then
            DoNext = False
            If (Direction = "Next") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) - 1
            ElseIf (Direction = "Prev") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) + 1
            End If
        ElseIf (oDocType01 = "선출요청") Then
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
        Call PS_SD030_FormItemEnabled '//UDO방식
    Else '//다음문서가 유효하지 않다면
        oForm01.Freeze True
        oForm01.Mode = fm_FIND_MODE
        Call PS_SD030_FormItemEnabled '//UDO방식
        If oForm01.Items("DocEntry").Enabled = True Then '//문서번호 필드가 입력이 가능하다면
            If (Direction = "Next") Then
                oForm01.Items("DocEntry").Specific.VALUE = Val(DocEntryNext + 1)
            ElseIf (Direction = "Prev") Then
                oForm01.Items("DocEntry").Specific.VALUE = Val(DocEntryNext - 1)
            End If
            oForm01.Items("1").Click ct_Regular
        End If
        oForm01.Freeze False
        PS_SD030_DirectionValidateDocument = False
        Set RecordSet01 = Nothing
        Set RecordSet02 = Nothing
        Exit Function
    End If
    PS_SD030_DirectionValidateDocument = True
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Exit Function
PS_SD030_DirectionValidateDocument_Error:
    Sbo_Application.SetStatusBarMessage Err.Number & Err.Description, bmt_Short, True
    PS_SD030_DirectionValidateDocument = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
End Function

Function PS_SD030_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_SD030_Validate_Error
    PS_SD030_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_SD030H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
        PS_SD030_Validate = False
        GoTo PS_SD030_Validate_Exit
    End If
    
    If ValidateType = "검사" Then
        For i = 1 To oMat01.VisualRowCount - 1 '//입력된 행에 대해
            If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [ORDR] ORDR LEFT JOIN [RDR1] RDR1 ON ORDR.DocEntry = RDR1.DocEntry WHERE CONVERT(NVARCHAR,ORDR.DocEntry) + '-' + CONVERT(NVARCHAR,RDR1.LineNum) = '" & oMat01.Columns("OrderNum").Cells(i).Specific.VALUE & "'", 0, 1) <= 0 Then
                Call MDC_Com.MDC_GF_Message("판매오더문서가 존재하지 않습니다.", "W")
                PS_SD030_Validate = False
                GoTo PS_SD030_Validate_Exit
            End If
        Next
        '//삭제된 행을 찾아서 삭제가능성 검사
        Dim Exist As Boolean
        Exist = False
        Query01 = "SELECT DocEntry,LineId FROM [@PS_SD030L] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        RecordSet01.DoQuery Query01
        For i = 0 To RecordSet01.RecordCount - 1
            Exist = False
            For j = 1 To oMat01.RowCount - 1
                '//라인번호가 같고, 품목코드가 같으면 존재하는행, LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
                If Val(RecordSet01.Fields(1).VALUE) = Val(oMat01.Columns("LineId").Cells(j).Specific.VALUE) And oMat01.Columns("LineId").Cells(j).Specific.VALUE <> "" Then
                    Exist = True
                End If
            Next
            If (Exist = False) Then '//삭제된 행중
                If (oForm01.Items("DocType").Specific.VALUE = "1") Then '//출하요청
                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" & Val(RecordSet01.Fields(0).VALUE) & "' AND PS_SD040L.U_SD030L = '" & Val(RecordSet01.Fields(1).VALUE) & "'", 0, 1)) > 0 Then
                        MDC_Com.MDC_GF_Message "삭제된행이 다른사용자에 의해 납품되었습니다. 적용할수 없습니다.", "W"
                        PS_SD030_Validate = False
                        GoTo PS_SD030_Validate_Exit
                    End If
                ElseIf (oForm01.Items("DocType").Specific.VALUE = "2") Then '//선출요청
                    If (MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "3") Then
                        MDC_Com.MDC_GF_Message "이미납품된 문서입니다. 삭제할수 없습니다.", "W"
                        PS_SD030_Validate = False
                        GoTo PS_SD030_Validate_Exit
                    End If
                End If
            End If
            RecordSet01.MoveNext
        Next
        '//수량가능성검사.
        For i = 1 To oMat01.VisualRowCount - 1
            If (oMat01.Columns("LineId").Cells(i).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
            Else
                '//매트릭스에 입력된 수량과 DB상에 존재하는 수량의 값비교
                If Val(oMat01.Columns("Weight").Cells(i).Specific.VALUE) < Val(MDC_PS_Common.GetValue("SELECT SUM(U_Weight) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND PS_SD040L.U_SD030L = '" & oMat01.Columns("LineId").Cells(i).Specific.VALUE & "'", 0, 1)) Then
                    MDC_Com.MDC_GF_Message "출하요청,선출요청 수량보다 작습니다.", "W"
                    PS_SD030_Validate = False
                    GoTo PS_SD030_Validate_Exit
                End If
                '//납품된 행이 있으면 값이 수정되어서는 안된다..
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_SD040L.U_SD030L = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    Query01 = "SELECT " '//품목코드는 변경되면 안된다.
                    Query01 = Query01 & " U_OrderNum, "
                    Query01 = Query01 & " U_ItemCode, "
                    Query01 = Query01 & " U_WhsCode "
                    Query01 = Query01 & " FROM [@PS_SD030L] PS_SD030L WHERE PS_SD030L.DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND PS_SD030L.LineId = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'"
                    RecordSet01.DoQuery Query01
                    If RecordSet01.Fields(0).VALUE = oMat01.Columns("OrderNum").Cells(i).Specific.VALUE And _
                    RecordSet01.Fields(1).VALUE = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE And _
                    RecordSet01.Fields(2).VALUE = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE _
                    Then
                    Else
                        MDC_Com.MDC_GF_Message "이미납품된 행입니다. 수정할수 없습니다.", "W"
                        PS_SD030_Validate = False
                        GoTo PS_SD030_Validate_Exit
                    End If
                End If
            End If
        Next
        '//
    ElseIf ValidateType = "수정" Then
        '//수정전 수정가능여부검사
'        If (oMat01.Columns("LineId").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 수정하여도 무방하다
'        Else
'            If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'                If (oForm01.Items("DocType").Specific.Value = "1") Then '//출하요청
'                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" & oForm01.Items("DocEntry").Specific.Value & "' AND PS_SD040L.U_SD030L = '" & Val(oMat01.Columns("LineId").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
'                        MDC_Com.MDC_GF_Message "납품된 행입니다. 수정할수 없습니다.", "W"
'                        PS_SD030_Validate = False
'                        GoTo PS_SD030_Validate_Exit
'                    End If
'                ElseIf (oForm01.Items("DocType").Specific.Value = "2") Then '//선출요청
'                    If (MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.Value & "'", 0, 1) = "3") Then
'                        MDC_Com.MDC_GF_Message "납품된 행입니다. 수정할수 없습니다.", "W"
'                        PS_SD030_Validate = False
'                        GoTo PS_SD030_Validate_Exit
'                    End If
'                End If
'            End If
'        End If
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
        If (oMat01.Columns("LineId").Cells(oLastColRow01).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
        Else
            If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
                If (oForm01.Items("DocType").Specific.VALUE = "1") Then '//출하요청
                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND PS_SD040L.U_SD030L = '" & Val(oMat01.Columns("LineId").Cells(oLastColRow01).Specific.VALUE) & "'", 0, 1)) > 0 Then
                        MDC_Com.MDC_GF_Message "납품된 행입니다. 삭제할수 없습니다.", "W"
                        PS_SD030_Validate = False
                        GoTo PS_SD030_Validate_Exit
                    End If
                ElseIf (oForm01.Items("DocType").Specific.VALUE = "2") Then '//선출요청
                    If (MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "3") Then
                        MDC_Com.MDC_GF_Message "납품된 행입니다. 삭제할수 없습니다.", "W"
                        PS_SD030_Validate = False
                        GoTo PS_SD030_Validate_Exit
                    End If
                End If
            End If
        End If
    ElseIf ValidateType = "취소" Then
        Query01 = "SELECT DocEntry,LineId,U_ItemCode FROM [@PS_SD030L] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        RecordSet01.DoQuery Query01
        For i = 0 To RecordSet01.RecordCount - 1
            If (oForm01.Items("DocType").Specific.VALUE = "1") Then '//출하요청
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" & Val(RecordSet01.Fields(0).VALUE) & "' AND PS_SD040L.U_SD030L = '" & Val(RecordSet01.Fields(1).VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "납품된 문서입니다. 적용할수 없습니다.", "W"
                    PS_SD030_Validate = False
                    GoTo PS_SD030_Validate_Exit
                End If
            ElseIf (oForm01.Items("DocType").Specific.VALUE = "2") Then '//선출요청
                If (MDC_PS_Common.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "3") Then
                    MDC_Com.MDC_GF_Message "납품된 문서입니다. 삭제할수 없습니다.", "W"
                    PS_SD030_Validate = False
                    GoTo PS_SD030_Validate_Exit
                End If
            End If
            RecordSet01.MoveNext
        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_SD030_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_SD030_Validate_Error:
    PS_SD030_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_SD030_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD030_Print_Report01()
On Error GoTo PS_SD030_Print_Report01_Error
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim i As Long
    
    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" And oDoType01 <> "선출요청" Then
        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
        Exit Sub
    End If
    
    If oForm01.Items("DocType").Specific.Selected.VALUE <> "2" Then
        Call MDC_Com.MDC_GF_Message("선출요청이 아닙니다.", "W")
        Exit Sub
    End If
    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" And oDoType01 <> "선출요청" Then
        Call MDC_Com.MDC_GF_Message("납품이 되지 않았습니다.", "W")
        Exit Sub
    End If
    If oForm01.Items("BPLId").Specific.Selected.VALUE <> "2" Then
        Call MDC_Com.MDC_GF_Message("사업장이 동래가 아닙니다.", "W")
        Exit Sub
    End If
    For i = 1 To oMat01.VisualRowCount - 1
        If oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.VALUE <> "105" And oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.VALUE <> "106" Then
            Call MDC_Com.MDC_GF_Message("품목이 기계공구,몰드가 아닙니다.", "W")
            Exit Sub
        End If
    Next
    
    Call ConnectODBC
    WinTitle = "[PS_PP540_10] 출하요청서"
    ReportName = "PS_PP540_10.rpt"
    sQry = "EXEC PS_PP540_10 '선출','" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)

    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_SD030_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD030_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD030_Print_Report02()
On Error GoTo PS_SD030_Print_Report02_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim i As Long
    Dim sQry01                      As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '//출하요청서
    If PS_SD030_ValidateCreditLine = False Then '//여신한도체크
        Exit Sub
    End If
    
    
    If oForm01.Items("BPLId").Specific.Selected.VALUE <> "1" And oForm01.Items("BPLId").Specific.Selected.VALUE <> "4" Then
        Call MDC_Com.MDC_GF_Message("사업장이 창원,서울이 아닙니다.", "W")
        Exit Sub
    End If
    
    '// 한도초과체크후 정상프린트가 되면 요청프린트유무에 'Y' Update
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    sQry01 = "Update [@PS_SD030H] Set U_PrtYn = 'Y' Where DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    Call oRecordSet01.DoQuery(sQry01)
    
'    For i = 1 To oMat01.VisualRowCount - 1
'        If oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.VALUE <> "101" Then
'            Call MDC_Com.MDC_GF_Message("품목이 휘팅이 아닙니다.", "W")
'            Exit Sub
'        End If
'    Next
    
    Call ConnectODBC
    WinTitle = "[PS_SD030_10] 레포트"
    ReportName = "PS_SD030_10.rpt"   '//임시출력
    sQry = "EXEC PS_SD030_10 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    sQry01 = "Select BPLName FROM [OBPL] WHERE BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "'"
    Call oRecordSet01.DoQuery(sQry01)
    gRpt_Formula(1) = "BPLName"
    gRpt_Formula_Value(1) = Trim(oRecordSet01.Fields(0).VALUE)

    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_SD030_Print_Report02_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD030_Print_Report02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PS_SD030_Print_Report03() '//임시 출하요청서 출력
On Error GoTo PS_SD030_Print_Report03_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim i As Long
    Dim sQry01                      As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    
    If oForm01.Items("BPLId").Specific.Selected.VALUE <> "1" And oForm01.Items("BPLId").Specific.Selected.VALUE <> "4" Then
        Call MDC_Com.MDC_GF_Message("사업장이 창원,서울이 아닙니다.", "W")
        Exit Sub
    End If
    
    
    Call ConnectODBC
    WinTitle = "[PS_SD030_20] 레포트"
    ReportName = "PS_SD030_20.rpt"
    sQry = "EXEC PS_SD030_10 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    sQry01 = "Select BPLName FROM [OBPL] WHERE BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "'"
    Call oRecordSet01.DoQuery(sQry01)
    gRpt_Formula(1) = "BPLName"
    gRpt_Formula_Value(1) = Trim(oRecordSet01.Fields(0).VALUE)

    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_SD030_Print_Report03_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD030_Print_Report03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SD030_ValidateDelivery() As Boolean
On Error GoTo PS_SD030_ValidateDelivery_Error
    PS_SD030_ValidateDelivery = True
    If PS_SD030_ValidateCreditLine = False Then
        PS_SD030_ValidateDelivery = False
        Exit Function
    End If
    
    '//생산완료 되었는지 확인
'    Dim i As Long
'
'    For i = 1 To oMat01.VisualRowCount - 1
'        mdc_ps_common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H
'    Next
    Exit Function
PS_SD030_ValidateDelivery_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD030_ValidateDelivery_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PS_SD030_ValidateCreditLine() As Boolean '//여신한도체크
On Error GoTo PS_SD030_ValidateCreditLine_Error
    PS_SD030_ValidateCreditLine = True
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    
    Dim i As Long
    Dim OCRDCreditLine As Currency '//고객여신한도
    Dim SD080CreditLine As Currency '//추가여신한도
    Dim OCRDBalance As Currency '//계정잔액
    Dim OCRDDNotesBal As Currency '//납품액
    Dim CurrentLineSum As Currency '//현재문서총계
    Dim OutPreP As Currency '출고예정금액
    '//If oMat01.Columns("WhsCode").Cells(1).Specific.VALUE = "104" And oMat01.Columns("ItmBsort").Cells(1).Specific.VALUE = "101" Then
    '//휘팅 서울출고
    If oMat01.Columns("WhsCode").Cells(1).Specific.VALUE = "101" And oMat01.Columns("ItmBsort").Cells(1).Specific.VALUE = "111" Then
        '//창원 분말
        If oDS_PS_SD030H.GetValue("U_PrtYn", 0) <> "Y" Then
            Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            Query01 = "EXEC PS_SD081_hando '" & Trim(oDS_PS_SD030H.GetValue("U_BPLId", 0)) & "', '" & oForm01.Items("CardCode").Specific.VALUE & "','" & oForm01.Items("DocDate").Specific.VALUE & "'"
            Call RecordSet01.DoQuery(Query01)
            
            If RecordSet01.RecordCount > 0 Then
                If RecordSet01.Fields("OverAmt").VALUE > 0 Then '//여신한도금액
                    Call MDC_Com.MDC_GF_Message("여신한도로 초과했습니다.", "W")
                    PS_SD030_ValidateCreditLine = False
                    Exit Function
                End If
            End If
        End If
        
        
        'EXEC [PS_SD081_hando] '4','12494','20110316'
        
'        OCRDCreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT CreditLine FROM [OCRD] WHERE CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        SD080CreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Sum(PS_SD080L.U_RequestP) FROM [@PS_SD080H] PS_SD080H LEFT JOIN [@PS_SD080L] PS_SD080L ON PS_SD080H.DocEntry = PS_SD080L.DocEntry WHERE PS_SD080H.Canceled = 'N' And PS_SD080H.U_OkYN = 'Y' AND PS_SD080H.U_DocDate = '" & oForm01.Items("DocDate").Specific.VALUE & "' AND PS_SD080L.U_CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
''        OCRDBalance = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Balance FROM [OCRD] WHERE CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        OCRDBalance = MDC_PS_Common.GetValue("SELECT ISNULL((Select Sum(Debit - Credit) from JDT1 where ShortName = '" & oForm01.Items("CardCode").Specific.VALUE & "' And Account = '11104010'),0)", 0, 1)
'        OCRDDNotesBal = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT DNotesBal FROM [OCRD] WHERE CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
''        OutPreP = MDC_PS_Common.GetValue("Select IsNull((Select Sum(LineTotal) From  [ORDR] a Inner Join [RDR1] b On a.DocEntry = b.DocEntry Where  b.LineStatus = 'O' And a.DocStatus = 'O' And a.CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        CurrentLineSum = 0
'        For i = 1 To oMat01.VisualRowCount - 1
'            CurrentLineSum = CurrentLineSum + Val(oMat01.Columns("LinTotal").Cells(i).Specific.VALUE)
'        Next
''        If ((OCRDCreditLine + SD080CreditLine) - (OCRDBalance + OCRDDNotesBal) < CurrentLineSum) Then
'        If ((OCRDCreditLine + SD080CreditLine) - (OCRDBalance + OCRDDNotesBal + OutPreP) < CurrentLineSum) Then
'            Call MDC_Com.MDC_GF_Message("여신한도가 부족합니다.", "W")
'            PS_SD030_ValidateCreditLine = False
'            Exit Function
'        End If
    End If
    Exit Function
PS_SD030_ValidateCreditLine_Error:
   PS_SD030_ValidateCreditLine = False
   Sbo_Application.SetStatusBarMessage "PS_SD030_ValidateCreditLine_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
