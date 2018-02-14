VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_CO170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_CO170.cls
'//  Module         : 재무관리 > 원가관리
'//  Desc           : 재공 완료품 대체관리등록
'//  FormType       : PS_CO170
'//  Create Date    : 2015.03.10
'//  Copyright  (c) Morning Data
'//  분말 제품을 통합관리하기 위해 등록 관리
'//  대체 등록된 제품은 제품 완료시 대체등록된 제품으로 작업일보(PP040) 포장공정으로 입력
'//  원가계산시 대체등록된 제품으로 원가계산을 하고 원가금액을 제품수불로 금액배부
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_CO170H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_CO170L As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_CO170.srf")
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
    
    oFormUniqueID01 = "PS_CO170_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "Code" '"DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    oDocType01 = oFromDocType01
    oDoType01 = oFromDocType01
    Call PS_CO170_CreateItems
    Call PS_CO170_ComboBox_Setting
    Call PS_CO170_EnableMenus
    
    Call PS_CO170_AddMatrixRow(0, True) '//UDO방식
    
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
                
            Case "1286": '닫기
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
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_CO170_FormItemEnabled '//UDO방식
                oForm01.Items("U_ItmBsort").Click ct_Regular
            Case "1282": '추가
                Call PS_CO170_FormItemEnabled '//UDO방식
                Call PS_CO170_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
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
'                    If (PS_CO170_FindValidateDocument("@PS_CO170H") = False) Then
'                        If Sbo_Application.Menus.Item("1281").Enabled = True Then '//찾기메뉴 활성화일때 수행
'                            Sbo_Application.ActivateMenuItem ("1281")
'                        Else
'                            Sbo_Application.SetStatusBarMessage "관리자에게 문의바랍니다", bmt_Short, True
'                        End If
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
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
                If PS_CO170_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                Dim ItmBsort$, ItmMsort$, Code$
                ItmBsort = Trim(oDS_PS_CO170H.GetValue("U_ItmBsort", 0))
                ItmMsort = Trim(oDS_PS_CO170H.GetValue("U_ItmMsort", 0))
                Code = ItmBsort & ItmMsort
                Call oDS_PS_CO170H.setValue("Code", 0, Code)
                Call oDS_PS_CO170H.setValue("Name", 0, Code)
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_CO170_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_CO170_FormItemEnabled
                    Call PS_CO170_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If pval.ActionSuccess = True Then
                End If
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_CO170_FormItemEnabled
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
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "U_ItmBsort" Then
                If oForm01.Items("U_ItmBsort").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
            If pval.ItemUID = "ItmMsort" Then
                If oForm01.Items("ItmMsort").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
            If pval.ItemUID = "Mat01" Then
                If pval.ColUID = "PO" Then
                    If oMat01.Columns("PO").Cells(pval.Row).Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                If pval.ColUID = "MPO" Then
                    If oMat01.Columns("MPO").Cells(pval.Row).Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
            
        End If
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "OrderNum") '//사용자값활성
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
            
                If pval.ColUID = "PO" Then
                    If oMat01.Columns("PO").Cells(pval.Row).Specific.VALUE = "" Then
                        GoTo Raise_EVENT_VALIDATE_Exit
                    End If
                    For i = 1 To oMat01.RowCount
                        If pval.Row <> i Then '//현재 선택되어있는 행이 아니면
                            If (oMat01.Columns("PO").Cells(pval.Row).Specific.VALUE = oMat01.Columns("PO").Cells(i).Specific.VALUE) Then
                                Call MDC_Com.MDC_GF_Message("동일한 항목이 존재합니다.", "W")
                                oMat01.Columns("PO").Cells(pval.Row).Specific.VALUE = ""
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
                    Query01 = "EXEC PS_CO170_01 '" & oMat01.Columns("PO").Cells(pval.Row).Specific.VALUE & "'"
                    Call RecordSet01.DoQuery(Query01)
                    For i = 0 To RecordSet01.RecordCount - 1
                        Call oDS_PS_CO170L.setValue("U_PO", pval.Row - 1, RecordSet01.Fields("PO").VALUE)
                        Call oDS_PS_CO170L.setValue("U_POEntry", pval.Row - 1, RecordSet01.Fields("POEntry").VALUE)
                        Call oDS_PS_CO170L.setValue("U_POLine", pval.Row - 1, RecordSet01.Fields("POLine").VALUE)
                        Call oDS_PS_CO170L.setValue("U_Sequence", pval.Row - 1, RecordSet01.Fields("Sequence").VALUE)
                        Call oDS_PS_CO170L.setValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                        Call oDS_PS_CO170L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
                        Call oDS_PS_CO170L.setValue("U_CpCode", pval.Row - 1, RecordSet01.Fields("CpCode").VALUE)
                        Call oDS_PS_CO170L.setValue("U_CpName", pval.Row - 1, RecordSet01.Fields("CpName").VALUE)
                        RecordSet01.MoveNext
                    Next
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_CO170L.GetValue("U_PO", pval.Row - 1)) <> "" Then
                        PS_CO170_AddMatrixRow (pval.Row)
                    End If
                    oMat01.LoadFromDataSource
                    oMat01.AutoResizeColumns
                    
                    
                    oForm01.Update
                    Set RecordSet01 = Nothing
                ElseIf pval.ColUID = "MPO" Then
                    If oMat01.Columns("MPO").Cells(pval.Row).Specific.VALUE = "" Then
                        GoTo Raise_EVENT_VALIDATE_Exit
                    End If
'                    For i = 1 To oMat01.RowCount
'                        If pval.Row <> i Then '//현재 선택되어있는 행이 아니면
'                            If (oMat01.Columns("MPO").Cells(pval.Row).Specific.VALUE = oMat01.Columns("MPO").Cells(i).Specific.VALUE) Then
'                                Call MDC_Com.MDC_GF_Message("동일한 항목이 존재합니다.", "W")
'                                oMat01.Columns("MPO").Cells(pval.Row).Specific.VALUE = ""
'                                GoTo Raise_EVENT_VALIDATE_Exit
'                            End If
''                            If (Mid(oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value, "-") - 1) <> _
''                            Mid(oMat01.Columns("OrderNum").Cells(i).Specific.Value, 1, InStr(oMat01.Columns("OrderNum").Cells(i).Specific.Value, "-") - 1)) Then
''                                Call MDC_Com.MDC_GF_Message("동일하지않은 수주문서가 존재합니다.", "W")
''                                oMat01.Columns("OrderNum").Cells(pval.Row).Specific.Value = ""
''                                GoTo Raise_EVENT_VALIDATE_Exit
''                            End If
'                        End If
'                    Next
                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    Query01 = "EXEC PS_CO170_01 '" & oMat01.Columns("MPO").Cells(pval.Row).Specific.VALUE & "'"
                    Call RecordSet01.DoQuery(Query01)
                    For i = 0 To RecordSet01.RecordCount - 1
                        Call oDS_PS_CO170L.setValue("U_MPO", pval.Row - 1, RecordSet01.Fields("PO").VALUE)
                        Call oDS_PS_CO170L.setValue("U_MPOEntry", pval.Row - 1, RecordSet01.Fields("POEntry").VALUE)
                        Call oDS_PS_CO170L.setValue("U_MPOLine", pval.Row - 1, RecordSet01.Fields("POLine").VALUE)
                        Call oDS_PS_CO170L.setValue("U_MSequenc", pval.Row - 1, RecordSet01.Fields("Sequence").VALUE)
                        Call oDS_PS_CO170L.setValue("U_MItemCod", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                        Call oDS_PS_CO170L.setValue("U_MItemNam", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
                        Call oDS_PS_CO170L.setValue("U_MCpCode", pval.Row - 1, RecordSet01.Fields("CpCode").VALUE)
                        Call oDS_PS_CO170L.setValue("U_MCpName", pval.Row - 1, RecordSet01.Fields("CpName").VALUE)
                        RecordSet01.MoveNext
                    Next
'                        If oMat01.RowCount = pval.Row And Trim(oDS_PS_CO170L.GetValue("U_MPO", pval.Row - 1)) <> "" Then
'                            PS_CO170_AddMatrixRow (pval.Row)
'                        End If
                    oMat01.LoadFromDataSource
                    oMat01.AutoResizeColumns
                    
                    
                    oForm01.Update
                    Set RecordSet01 = Nothing
                
                Else
                    Call oDS_PS_CO170L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If

                oMat01.LoadFromDataSource
                oMat01.AutoResizeColumns
                oForm01.Update
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                If (pval.ItemUID = "U_ItmBsort") Then
                    Call oDS_PS_CO170H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_CO170H.setValue("U_ItmBName", 0, MDC_PS_Common.GetValue("SELECT Name FROM [@PSH_ITMBSORT] WHERE Code = '" & oForm01.Items("U_ItmBsort").Specific.VALUE & "'", 0, 1))
                ElseIf (pval.ItemUID = "ItmMsort") Then
                    Call oDS_PS_CO170H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_CO170H.setValue("U_ItmMName", 0, MDC_PS_Common.GetValue("SELECT U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" & oForm01.Items("U_ItmBsort").Specific.VALUE & "' And U_Code = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                Else
                    Call oDS_PS_CO170H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
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
        Call PS_CO170_FormItemEnabled
        Call PS_CO170_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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
        If (pval.ItemUID = "U_ItmBsort" Or pval.ItemUID = "ItmBName") Then
            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_CO170H", "U_ItmBsort,U_ItmBName")
        End If
        If (pval.ItemUID = "ItmMsort" Or pval.ItemUID = "ItmMName") Then
            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_CO170H", "U_ItmMsort,U_ItmMName")
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
            
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_CO170L.RemoveRecord(oDS_PS_CO170L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_CO170_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_CO170L.GetValue("U_PO", oMat01.RowCount - 1)) <> "" Then
                    Call PS_CO170_AddMatrixRow(oMat01.RowCount)
                End If
            End If
            
            oForm01.Update
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Private Function PS_CO170_CreateItems() As Boolean
On Error GoTo PS_CO170_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_CO170H = oForm01.DataSources.DBDataSources("@PS_CO170H")
    Set oDS_PS_CO170L = oForm01.DataSources.DBDataSources("@PS_CO170L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
'    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
'    Call oForm01.DataSources.UserDataSources.Add("WhsCode", dt_SHORT_TEXT, 100)
'    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
'    Call oForm01.Items("WhsCode").Specific.DataBind.SetBound(True, "", "WhsCode")
    
    
    
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_CO170_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO170_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_CO170_ComboBox_Setting()
On Error GoTo PS_CO170_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_CO170_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO170_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Sub PS_CO170_FormItemEnabled()
On Error GoTo PS_CO170_FormItemEnabled_Error
    Dim i As Long
    Call oForm01.Freeze(True)
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("U_ItmBsort").Enabled = True
        oForm01.Items("ItmMsort").Enabled = True
        oForm01.Items("Mat01").Enabled = True
        oMat01.AutoResizeColumns
'        Call oForm01.Items("BPLId").Specific.Select("1", psk_ByValue)
        Call oForm01.Items("BPLId").Specific.Select(User_BPLId(), psk_ByValue) '//2010.12.06 추가
        
        
        
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("U_ItmBsort").Enabled = True
        oForm01.Items("ItmMsort").Enabled = True
        oForm01.Items("Comment").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        oMat01.AutoResizeColumns
        Call oForm01.EnableMenu("1281", False)
        Call oForm01.EnableMenu("1282", True)
        
        
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
                   
        
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("U_ItmBsort").Enabled = False
        oForm01.Items("ItmMsort").Enabled = False
        
        oMat01.AutoResizeColumns
        Call oForm01.EnableMenu("1281", True)
        Call oForm01.EnableMenu("1282", False)
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_CO170_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO170_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_CO170_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_CO170_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_CO170L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_CO170L.Offset = oRow
    oDS_PS_CO170L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_CO170_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_CO170_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Private Sub PS_CO170_EnableMenus()
On Error GoTo PS_CO170_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1293", True)
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, True, True, True, True, True, True, False, False, False, False, True, False) '//메뉴설정
    Exit Sub
PS_CO170_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_CO170_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Function PS_CO170_DataValidCheck() As Boolean
On Error GoTo PS_CO170_DataValidCheck_Error
    PS_CO170_DataValidCheck = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oForm01.Items("U_ItmBsort").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "품목대분류는 필수입니다.", bmt_Short, True
        oForm01.Items("U_ItmBsort").Click ct_Regular
        PS_CO170_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("ItmMsort").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "품목중분류는 필수입니다.", bmt_Short, True
        oForm01.Items("ItmMsort").Click ct_Regular
        PS_CO170_DataValidCheck = False
        Exit Function
    End If
    
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_CO170_DataValidCheck = False
        Exit Function
    End If
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("PO").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "작지문서라인은 필수입니다.", bmt_Short, True
            oMat01.Columns("PO").Cells(i).Click ct_Regular
            PS_CO170_DataValidCheck = False
            Exit Function
        End If
        If (Val(oMat01.Columns("MPO").Cells(i).Specific.VALUE) <= 0) Then
            Sbo_Application.SetStatusBarMessage "이동작지문서라인은 필수입니다.", bmt_Short, True
            oMat01.Columns("MPO").Cells(i).Click ct_Regular
            PS_CO170_DataValidCheck = False
            Exit Function
        End If
        
    Next
        
    Call oDS_PS_CO170L.RemoveRecord(oDS_PS_CO170L.Size - 1)
    Call oMat01.LoadFromDataSource
    If oForm01.Mode = fm_ADD_MODE Then

    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_CO170_DataValidCheck_Error:
    PS_CO170_DataValidCheck = False
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_CO170_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

