VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD091"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_SD091.cls
'//  Module         : 이동요청등록
'//  Desc           : SD
'//  FormType       : PS_SD091
'//  Create Date    : 2010.10.29
'//  Creator        : Dong sub Kim
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_SD091H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_SD091L As SAPbouiCOM.DBDataSource '등록라인

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
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD091.srf")
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
    
    oFormUniqueID01 = "PS_SD091_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_SD091_CreateItems
    Call PS_SD091_ComboBox_Setting
    Call PS_SD091_Initial_Setting
    Call PS_SD091_CF_ChooseFromList
    Call PS_SD091_EnableMenus
    Call PS_SD091_SetDocument(oFromDocEntry01)
    Call PS_SD091_FormResize
    
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
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
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
                Call PS_SD091_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_SD091_FormItemEnabled '//UDO방식
                Call PS_SD091_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_SD091_FormItemEnabled
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
        If pval.ItemUID = "PS_SD091" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_SD091_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_SD091_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        ElseIf pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_OK_MODE Then
                Call Print_StockTrans_Docu
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_SD091" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_SD091_FormItemEnabled
                    Call PS_SD091_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_SD091_FormItemEnabled
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
    Dim TempForm01 As Variant
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" And pval.ColUID = "ItemCode" And pval.CharPressed = 9 Then
            If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = "" Then
                Set TempForm01 = New PS_SM020
                Call TempForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                Call PS_SD091_AddMatrixRow(0, True)
                BubbleEvent = False
            End If
        End If
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "")
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OutWhCd", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "InWhCd", "") '//사용자값활성
    ElseIf pval.BeforeAction = False Then
    
    End If
    Call oForm01.Freeze(False)
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
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_SD091L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_SD091L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_SD091_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_SD091L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                End If
            Else
                If (pval.ItemUID = "BPLId") Then
                    If oForm01.Items("BPLId").Specific.VALUE = "1" Then
                        Call oDS_PS_SD091H.setValue("U_OutWhCd", 0, "104")
                        Call oDS_PS_SD091H.setValue("U_InWhCd", 0, "101")
                    ElseIf oForm01.Items("BPLId").Specific.VALUE = "4" Then
                        Call oDS_PS_SD091H.setValue("U_OutWhCd", 0, "101")
                        Call oDS_PS_SD091H.setValue("U_InWhCd", 0, "104")
                    End If
                End If
                
'                If (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_SD091H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_SD091H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_SD091H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_SD091H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                End If
            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
            oForm01.Update
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
    Dim sQry01 As String
    Dim sQry02 As String
    Dim sQry03 As String
    Dim sQry04 As String
    Dim sQry05 As String
    Dim sQry06 As String
    Dim sQry07 As String
    Dim sQry08 As String
    Dim sQry09 As String
    Dim sQry10 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Dim oRecordSet03 As SAPbobsCOM.Recordset
    Dim SumQty As Long, SumWeight As Currency
    Dim i As Integer
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
            ' 품목 이름 Query
                If (pval.ColUID = "ItemCode") Then
                
                    If (pval.Row = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Call PS_SD091_AddMatrixRow(pval.Row, False)
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                    End If
                    
                    oMat01.FlushToDataSource
                    
                    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
                    sQry02 = "Select ItemName, U_ItmBsort, U_ItmMsort, U_Unit1, U_Size, U_ItemType, U_Quality, U_Mark, U_CallSize, U_SbasUnit From [OITM] Where "
                    sQry02 = sQry02 & "ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) & "'"
                    oRecordset02.DoQuery sQry02
                    oDS_PS_SD091L.setValue "U_ItemName", pval.Row - 1, oRecordset02.Fields(0).VALUE
                    oDS_PS_SD091L.setValue "U_Unit1", pval.Row - 1, oRecordset02.Fields(3).VALUE
                    oDS_PS_SD091L.setValue "U_Size", pval.Row - 1, oRecordset02.Fields(4).VALUE
                    oDS_PS_SD091L.setValue "U_CallSize", pval.Row - 1, oRecordset02.Fields(8).VALUE
                    'oMat01.Columns("ItemName").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet02.Fields(0).VALUE)
                    'oMat01.Columns("Unit1").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet02.Fields(3).VALUE)
                    'oMat01.Columns("Size").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet02.Fields(4).VALUE)
            ' 품목 대분류
                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
                    sQry03 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oRecordset02.Fields(1).VALUE) & "'"
                    oRecordSet03.DoQuery sQry03
                    oDS_PS_SD091L.setValue "U_ItmBsort", pval.Row - 1, oRecordSet03.Fields(1).VALUE
                    
                    'Call oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).VALUE, psk_ByDescription)
                    Set oRecordSet03 = Nothing
            ' 품목 중분류
                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
                    sQry04 = "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_Code = '" & Trim(oRecordset02.Fields(2).VALUE) & "'"
                    oRecordSet03.DoQuery sQry04
                    oDS_PS_SD091L.setValue "U_ItmMsort", pval.Row - 1, oRecordSet03.Fields(1).VALUE
                    'Call oMat01.Columns("ItmMsort").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).VALUE, psk_ByDescription)
                    Set oRecordSet03 = Nothing
            ' 형태타입
                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
                    sQry05 = "SELECT Code, Name FROM [@PSH_SHAPE] WHERE Code = '" & Trim(oRecordset02.Fields(5).VALUE) & "'"
                    oRecordSet03.DoQuery sQry05
                    
                    oDS_PS_SD091L.setValue "U_ItemType", pval.Row - 1, oRecordSet03.Fields(1).VALUE
                    'Call oMat01.Columns("ItemType").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).VALUE, psk_ByDescription)
                    Set oRecordSet03 = Nothing
            ' 질별
                    
                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
                    sQry06 = "SELECT Code, Name FROM [@PSH_QUALITY] WHERE Code = '" & Trim(oRecordset02.Fields(6).VALUE) & "'"
                    oRecordSet03.DoQuery sQry06
                    
                    oDS_PS_SD091L.setValue "U_Quality", pval.Row - 1, oRecordSet03.Fields(1).VALUE
                    'Call oMat01.Columns("Quality").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).VALUE, psk_ByDescription)
                    Set oRecordSet03 = Nothing
            ' 인증기호
                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
                    sQry07 = "SELECT Code, Name FROM [@PSH_MARK] WHERE Code = '" & Trim(oRecordset02.Fields(7).VALUE) & "'"
                    oRecordSet03.DoQuery sQry07
                    
                    oDS_PS_SD091L.setValue "U_Mark", pval.Row - 1, oRecordSet03.Fields(1).VALUE
                    'Call oMat01.Columns("Mark").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).VALUE, psk_ByDescription)
                    Set oRecordSet03 = Nothing
            ' 판매기준단위
                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
                    sQry08 = "SELECT Code, Name FROM [@PSH_UOMORG] WHERE Code = '" & Trim(oRecordset02.Fields(9).VALUE) & "'"
                    oRecordSet03.DoQuery sQry08
                    
                    oDS_PS_SD091L.setValue "U_SbasUnit", pval.Row - 1, oRecordSet03.Fields(1).VALUE
                    'Call oMat01.Columns("SbasUnit").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).VALUE, psk_ByDescription)
                    Set oRecordSet03 = Nothing
                    
'                    oDS_PS_SD091L.setValue "U_Weight", pval.Row - 1, Val(oMat01.Columns("Weight").Cells(pval.Row).Specific.VALUE)
'                    oDS_PS_SD091L.setValue "U_Unweight", pval.Row - 1, Val(oMat01.Columns("Unweight").Cells(pval.Row).Specific.VALUE)
'                    oMat01.FlushToDataSource
'                    Call PS_SD091_AddMatrixRow(pval.Row, False)
                ElseIf (pval.ColUID = "Qty") Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE < 0 Then
                         Call oDS_PS_SD091L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                    Else
                         Call oDS_PS_SD091L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    End If
'                    Call oDS_PS_SD091L.setValue("U_" & pval.ColUID, pval.Row - 1, Round((oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE * Val(oMat01.Columns("Unweight").Cells(pval.Row).Specific.VALUE)) / 1000, 3))
'                    oMat01.Columns("Weight").Cells(pval.Row).Specific.VALUE = Round((Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) * _
'                                                                              Val(oMat01.Columns("Unweight").Cells(pval.Row).Specific.VALUE)) / 1000, 3)
                Else
                    Call oDS_PS_SD091L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
                
                    
                For i = 0 To oMat01.VisualRowCount - 1
                        If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                            SumQty = SumQty
                        Else
                            SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                        End If
                        SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                        
                 Next i
                 oForm01.Items("SumQty").Specific.VALUE = SumQty
                 oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                    
                
             ' 단중
'                If (pval.ColUID = "Unweight") Then
'                    oMat01.Columns("Weight").Cells(pval.Row).Specific.VALUE = Round((Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) * _
'                                                                              Val(oMat01.Columns("Unweight").Cells(pval.Row).Specific.VALUE)) / 1000, 3)
'                    oMat01.FlushToDataSource
''                    Call PS_SD091_AddMatrixRow(pval.Row, False)
'                End If
               
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_SD091H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "CntcCode") Then
                    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    Call oDS_PS_SD091H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
'                    Call oDS_PS_SD091H.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("lastName", "firstName", "[OHEM]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
                    sQry01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry01
                    oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                    Set oRecordSet01 = Nothing
                Else
                    Call oDS_PS_SD091H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
            End If
            
'            oMat01.FlushToDataSource
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oForm01.Update
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
Dim i As Integer
Dim SumQty As Long
Dim SumWeight As Currency

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_SD091_FormItemEnabled
        Call PS_SD091_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        
        For i = 0 To oMat01.VisualRowCount - 1
            If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                SumQty = SumQty
            Else
                SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
            End If
            SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
            
        Next i
        oForm01.Items("SumQty").Specific.VALUE = SumQty
        oForm01.Items("SumWeight").Specific.VALUE = SumWeight
        
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_SD091_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    
    Dim oDataTable01 As SAPbouiCOM.DataTable
    Dim oDataTable02 As SAPbouiCOM.DataTable
    
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        If (pval.ItemUID = "CardCode") Then
            Set oDataTable01 = pval.SelectedObjects
            If oDataTable01 Is Nothing Then
            Else
                Call oDS_PS_SD091H.setValue("U_CardCode", 0, oDataTable01.Columns("CardCode").Cells(0).VALUE)
                Call oDS_PS_SD091H.setValue("U_CardName", 0, oDataTable01.Columns("CardName").Cells(0).VALUE)
        ' // 찾기나 문서이동 버튼 클릭 시에 갱신으로 바뀌지 않음
                If oForm01.Mode <> fm_ADD_MODE Then oForm01.Mode = fm_UPDATE_MODE
            End If
        ElseIf (pval.ItemUID = "ShipTo") Then
            Set oDataTable02 = pval.SelectedObjects
            If oDataTable02 Is Nothing Then
            Else
                Call oDS_PS_SD091H.setValue("U_ShipTo", 0, oDataTable02.Columns("CardCode").Cells(0).VALUE)
                Call oDS_PS_SD091H.setValue("U_ShipNm", 0, oDataTable02.Columns("CardName").Cells(0).VALUE)
                If oForm01.Mode <> fm_ADD_MODE Then oForm01.Mode = fm_UPDATE_MODE
            End If
        End If
        
        oForm01.Update
    End If
    Set oDataTable01 = Nothing
    Set oDataTable02 = Nothing
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
    Dim SumQty As Long
    Dim SumWeight As Currency
    
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
'            If (PS_SD091_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_SD091L.RemoveRecord(oDS_PS_SD091L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_SD091_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_SD091L.GetValue("U_ItemCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_SD091_AddMatrixRow(oMat01.RowCount)
                End If
            End If
            
            For i = 0 To oMat01.VisualRowCount - 1
                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                    
            Next i
            oForm01.Items("SumQty").Specific.VALUE = SumQty
            oForm01.Items("SumWeight").Specific.VALUE = SumWeight
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_SD091_CreateItems() As Boolean
On Error GoTo PS_SD091_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_SD091H = oForm01.DataSources.DBDataSources("@PS_SD091H")
    Set oDS_PS_SD091L = oForm01.DataSources.DBDataSources("@PS_SD091L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    Call oForm01.DataSources.UserDataSources.Add("SumWeight", dt_QUANTITY)
    
    Call oForm01.Items("SumQty").Specific.DataBind.SetBound(True, "", "SumQty")
    Call oForm01.Items("SumWeight").Specific.DataBind.SetBound(True, "", "SumWeight")
    
    
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_SD091_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD091_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SD091_ComboBox_Setting()
On Error GoTo PS_SD091_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD091", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD091", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_SD091", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD091", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD091", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_SD091", "ItemCode")
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL Where BPLId = '1' Or BPLId = '4' order by BPLId", "1", False, False)
    
    ' 반출등록
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OutWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "104", False, True)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("InWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "101", False, True)
    
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] order by U_Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Mark"), "SELECT Code, Name FROM [@PSH_MARK] order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("SbasUnit"), "SELECT Code, Name  FROM [@PSH_UOMORG] order by Code")
    
'            ' 품목 대분류
'                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
'                    sQry03 = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oRecordSet02.Fields(1).Value) & "'"
'                    oRecordSet03.DoQuery sQry03
'                    Call oMat01.Columns("ItmBsort").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
'                    Set oRecordSet03 = Nothing
'            ' 품목 중분류
'                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
'                    sQry04 = "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_Code = '" & Trim(oRecordSet02.Fields(2).Value) & "'"
'                    oRecordSet03.DoQuery sQry04
'                    Call oMat01.Columns("ItmMsort").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
'                    Set oRecordSet03 = Nothing
'            ' 형태타입
'                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
'                    sQry05 = "SELECT Code, Name FROM [@PSH_SHAPE] WHERE Code = '" & Trim(oRecordSet02.Fields(5).Value) & "'"
'                    oRecordSet03.DoQuery sQry05
'                    Call oMat01.Columns("ItemType").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
'                    Set oRecordSet03 = Nothing
'            ' 질별
'                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
'                    sQry06 = "SELECT Code, Name FROM [@PSH_QUALITY] WHERE Code = '" & Trim(oRecordSet02.Fields(6).Value) & "'"
'                    oRecordSet03.DoQuery sQry06
'                    Call oMat01.Columns("Quality").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
'                    Set oRecordSet03 = Nothing
'            ' 인증기호
'                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
'                    sQry07 = "SELECT Code, Name FROM [@PSH_MARK] WHERE Code = '" & Trim(oRecordSet02.Fields(7).Value) & "'"
'                    oRecordSet03.DoQuery sQry07
'                    Call oMat01.Columns("Mark").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
'                    Set oRecordSet03 = Nothing
'            ' 판매기준단위
'                    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
'                    sQry08 = "SELECT Code, Name FROM [@PSH_UOMORG] WHERE Code = '" & Trim(oRecordSet02.Fields(9).Value) & "'"
'                    oRecordSet03.DoQuery sQry08
'                    Call oMat01.Columns("SbasUnit").Cells(pval.Row).Specific.Select(oRecordSet03.Fields(1).Value, psk_ByDescription)
'                    Set oRecordSet03 = Nothing
            
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD091_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD091_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD091_Initial_Setting()
On Error GoTo PS_SD091_Initial_Setting_Error
    ' 사업장
    Dim lcl_User_BPLId As String
    lcl_User_BPLId = MDC_PS_Common.User_BPLId()
    
    If lcl_User_BPLId = "1" Or lcl_User_BPLId = "4" Then '소속사업장이 창원이나 구로일 때만 콤보박스 Select (2011.09.21 송명규 추가)
        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    End If
 ' 인수자
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
    Exit Sub
PS_SD091_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD091_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD091_CF_ChooseFromList()
On Error GoTo PS_SD091_CF_ChooseFromList_Error
    Dim oCFL01                As SAPbouiCOM.ChooseFromList
    Dim oCFL02                As SAPbouiCOM.ChooseFromList
    Dim oCons                 As SAPbouiCOM.Conditions
    Dim oCon                  As SAPbouiCOM.Condition
    Dim oCFLs01               As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLs02               As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLCreationParams01  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oCFLCreationParams02  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit01               As SAPbouiCOM.EditText
    Dim oEdit02               As SAPbouiCOM.EditText
    
'    Set oEdit01 = oForm01.Items("CntcCode").Specific
'    Set oCFLs01 = oForm01.ChooseFromLists
'    Set oCFLCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams01.ObjectType = "171"
'    oCFLCreationParams01.uniqueID = "CFLCNTCCODE"
'    oCFLCreationParams01.MultiSelection = False
'    Set oCFL01 = oCFLs01.Add(oCFLCreationParams01)
'
'    oEdit01.ChooseFromListUID = "CFLCNTCCODE"
'    oEdit01.ChooseFromListAlias = "CntcCode"
'
'    Set oEdit01 = Nothing
'    Set oCFLs01 = Nothing
'    Set oCFLCreationParams01 = Nothing
     
    Set oEdit01 = oForm01.Items("CardCode").Specific
    Set oCFLs01 = oForm01.ChooseFromLists
    Set oCFLCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams01.ObjectType = lf_BusinessPartner
    oCFLCreationParams01.uniqueID = "CFLCARDCODE"
    oCFLCreationParams01.MultiSelection = False
    Set oCFL01 = oCFLs01.Add(oCFLCreationParams01)

    oEdit01.ChooseFromListUID = "CFLCARDCODE"
    oEdit01.ChooseFromListAlias = "CardCode"
    
    ' Choose from list 에 조건을 줄 경우
    ' choosefromlist가 화면에 나오면 서식세팅으로 원하는 필드값 추가 가능
    Set oCons = oCFL01.GetConditions()
    Set oCon = oCons.Add()
    oCon.Alias = "CardType"                                               ' Condition Field
    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL            ' Equal
    oCon.CondVal = "C"                                                   ' Condition Value
    oCFL01.SetConditions oCons
    
    Set oEdit02 = oForm01.Items("ShipTo").Specific
    Set oCFLs02 = oForm01.ChooseFromLists
    Set oCFLCreationParams02 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams02.ObjectType = lf_BusinessPartner
    oCFLCreationParams02.uniqueID = "CFLSHIPCODE"
    oCFLCreationParams02.MultiSelection = False
    Set oCFL02 = oCFLs02.Add(oCFLCreationParams02)

    oEdit02.ChooseFromListUID = "CFLSHIPCODE"
    oEdit02.ChooseFromListAlias = "CardCode"                ' 정의된 오브젝트만 사용 가능
    Exit Sub
PS_SD091_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD091_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD091_FormItemEnabled()
On Error GoTo PS_SD091_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        Call PS_SD091_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        Call oForm01.EnableMenu("1293", True)
             oForm01.Items("DocEntry").Enabled = False
             oForm01.Items("Btn01").Enabled = False
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
             oForm01.Items("DocEntry").Enabled = True
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        Call oForm01.EnableMenu("1293", True)
             oForm01.Items("Btn01").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
             oForm01.Items("DocEntry").Enabled = False
        Call oForm01.EnableMenu("1281", True)
        Call oForm01.EnableMenu("1282", True)
        Call oForm01.EnableMenu("1293", False)
             oForm01.Items("Btn01").Enabled = True
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD091_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD091_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD091_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD091_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_SD091L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD091L.Offset = oRow
    oDS_PS_SD091L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD091_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD091_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD091_FormClear()
On Error GoTo PS_SD091_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD091'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SD091_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD091_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD091_EnableMenus()
On Error GoTo PS_SD091_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    Call oForm01.EnableMenu("1293", True)         '// 행삭제
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    '//Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_SD091_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD091_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD091_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_SD091_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_SD091_FormItemEnabled
        Call PS_SD091_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_SD091_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_SD091_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD091_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_SD091_DataValidCheck() As Boolean
On Error GoTo PS_SD091_DataValidCheck_Error
    PS_SD091_DataValidCheck = False
    Dim i As Long
    If oForm01.Items("BPLId").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm01.Items("BPLId").Click ct_Regular
        PS_SD091_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("CntcCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "요청자는 필수입니다.", bmt_Short, True
        oForm01.Items("CntcCode").Click ct_Regular
        PS_SD091_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "요청일자는 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_SD091_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("OutWhCd").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "출고창고는 필수입니다.", bmt_Short, True
        oForm01.Items("OutWhCd").Click ct_Regular
        PS_SD091_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("InWhCd").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "입고창고는 필수입니다.", bmt_Short, True
        oForm01.Items("InWhCd").Click ct_Regular
        PS_SD091_DataValidCheck = False
        Exit Function
    End If
    If oMat01.VisualRowCount = 0 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_SD091_DataValidCheck = False
        Exit Function
    End If
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("ItemName").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
            PS_SD091_DataValidCheck = False
            Exit Function
        End If
        If (Val(oMat01.Columns("Qty").Cells(i).Specific.VALUE) <= 0) Then
            Sbo_Application.SetStatusBarMessage "수량은 필수입니다.", bmt_Short, True
            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
            PS_SD091_DataValidCheck = False
            Exit Function
        End If
        
    Next
    Call oDS_PS_SD091L.RemoveRecord(oDS_PS_SD091L.Size - 1)
    Call oMat01.LoadFromDataSource
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_SD091_FormClear
    End If
    PS_SD091_DataValidCheck = True
    Exit Function
PS_SD091_DataValidCheck_Error:
    PS_SD091_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_SD091_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD091_MTX01()
On Error GoTo PS_SD091_MTX01_Error
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
    Param01 = oForm01.Items("Param01").Specific.VALUE
    Param02 = oForm01.Items("Param01").Specific.VALUE
    Param03 = oForm01.Items("Param01").Specific.VALUE
    Param04 = oForm01.Items("Param01").Specific.VALUE
    
    Query01 = "SELECT 10"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_SD091_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_SD091L.InsertRecord (i)
        End If
        oDS_PS_SD091L.Offset = i
        Call oDS_PS_SD091L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_SD091L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_SD091_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_SD091_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD091_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_SD091_FormResize()
On Error GoTo PS_SD091_FormResize_Error
    
    Exit Sub
PS_SD091_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD091_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD091_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_SD091_Validate_Error
    PS_SD091_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "수정" Then
'        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
'        Dim Exist As Boolean
'        Exist = False
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            Exist = False
'            For j = 1 To oMat01.RowCount - 1
'                '//라인번호가 같고, 품목코드가 같으면 존재하는행 , LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
'                If Val(RecordSet01.Fields(1).Value) = Val(oMat01.Columns("U_LineNum").Cells(j).Specific.Value) And RecordSet01.Fields(2).Value = oMat01.Columns("1").Cells(j).Specific.Value And oMat01.Columns("U_LineNum").Cells(j).Specific.Value <> "" Then
'                    Exist = True
'                End If
'            Next
'            If (Exist = False) Then '//삭제된 행중
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "삭제된행이 다른사용자에 의해 출하,선출요청되었습니다. 적용할수 없습니다.", "W"
'                    PS_SD091_Validate = False
'                    GoTo PS_SD091_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
'        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
'            Else
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(oForm01.Items("8").Specific.Value) & "' AND U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할수 없습니다.", "W"
'                    PS_SD091_Validate = False
'                    GoTo PS_SD091_Validate_Exit
'                End If
'            End If
'        End If
    ElseIf ValidateType = "취소" Then
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
'                PS_SD091_Validate = False
'                GoTo PS_SD091_Validate_Exit
'            End If
'            RecordSet01.MoveNext
'        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_SD091_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_SD091_Validate_Error:
    PS_SD091_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_SD091_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub Print_StockTrans_Docu()
On Error GoTo Print_StockTrans_Docu_Error
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    
    Call ConnectODBC
    WinTitle = "[PS_SD045] 이동 요청서"
    ReportName = "PS_SD045.rpt"
    sQry = "EXEC PS_SD045_01 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    
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
    
Print_StockTrans_Docu_Error:
    
End Sub
