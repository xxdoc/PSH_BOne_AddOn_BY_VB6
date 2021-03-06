VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD477"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_SD477.cls
'//  Module         : 품목별판매집계현황
'//  Desc           :
'//  FormType       : PS_SD477
'//  Create Date    : 2010.01.25
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_SD477H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_SD477L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD477.srf")
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
    
    oFormUniqueID01 = "PS_SD477_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_SD477_CreateItems
    Call PS_SD477_ComboBox_Setting
    Call PS_SD477_Initial_Setting
    Call PS_SD477_CF_ChooseFromList
    Call PS_SD477_EnableMenus
    Call PS_SD477_SetDocument(oFromDocEntry01)
    Call PS_SD477_FormResize
    
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
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
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
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_SD477_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_SD477_Print_Report01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_SD477" Then
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "BPLId", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//사용자값활성
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
        If pval.ItemUID = "ItmBsort" Then

        End If
    
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
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "BPLId") Then
                oQuery01 = "SELECT BPLName, BPLId FROM [OBPL] WHERE BPLid = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("BPLName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf (pval.ItemUID = "CardCode") Then
                oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf (pval.ItemUID = "ItemCode") Then
                oQuery01 = "SELECT ItemName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf (pval.ItemUID = "CntcCode") Then
                oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
            oForm01.Items(pval.ItemUID).Click ct_Regular
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
        Call PS_SD477_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_SD477_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
    
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
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_SD477_CreateItems() As Boolean
On Error GoTo PS_SD477_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 100)
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 100)
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
     Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("Size", dt_SHORT_TEXT, 100)
'    Call oForm01.DataSources.UserDataSources.Add("CntcCode", dt_SHORT_TEXT, 100)
    
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    Call oForm01.Items("DocDateFr").Specific.DataBind.SetBound(True, "", "DocDateFr")
    Call oForm01.Items("DocDateTo").Specific.DataBind.SetBound(True, "", "DocDateTo")
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    Call oForm01.Items("Size").Specific.DataBind.SetBound(True, "", "Size")
    
    oForm01.Items("DocDateFr").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    
    oForm01.Items("DocDateTo").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_SD477_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD477_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SD477_ComboBox_Setting()
On Error GoTo PS_SD477_ComboBox_Setting_Error
     Call oForm01.Freeze(True)

        Call oForm01.Items("BPLId").Specific.ValidValues.Add("", "")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
        Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        
        Call oForm01.Items("ItemBsort").Specific.ValidValues.Add("", "")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", False, False)
        Call oForm01.Items("ItemBsort").Specific.Select(0, psk_Index)
        
'        Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD477", "SaleGbn", "", "", "")
'        Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD477", "SaleGbn", "", "N", "내수")
'        Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD477", "SaleGbn", "", "S", "수출")
'        Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD477", "SaleGbn", "", "L", "로컬")
'        Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("SaleGbn").Specific, "PS_SD477", "SaleGbn")
'        Call oForm01.Items("SaleGbn").Specific.Select(0, psk_Index)
        
'        Call oForm01.Items("ItmBsort").Specific.ValidValues.Add("", "")
'        Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", False, False)
'        Call oForm01.Items("ItmBsort").Specific.Select(0, psk_Index)
'
'        Call oForm01.Items("ItmMsort").Specific.ValidValues.Add("", "")
'        Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", "", False, False)
'        Call oForm01.Items("ItmMsort").Specific.Select(0, psk_Index)


    Call oForm01.Freeze(False)
    Exit Sub
    
PS_SD477_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD477_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD477_Initial_Setting()
On Error GoTo PS_SD477_Initial_Setting_Error
'  사업장
    Dim lcl_User_BPLId As String
    lcl_User_BPLId = MDC_PS_Common.User_BPLId()
    
    If lcl_User_BPLId = "1" Or lcl_User_BPLId = "4" Then '소속사업장이 창원이나 구로일 때만 콤보박스 Select (2011.09.21 송명규 추가)
        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    End If
 ' 인수자
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD()
    Exit Sub
PS_SD477_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD477_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_SD477_CF_ChooseFromList()
On Error GoTo PS_SD477_CF_ChooseFromList_Error
    Exit Sub
PS_SD477_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD477_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD477_FormItemEnabled()
On Error GoTo PS_SD477_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD477_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD477_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD477_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD477_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD477_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD477_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD477_FormClear()
On Error GoTo PS_SD477_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD477'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SD477_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD477_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD477_EnableMenus()
On Error GoTo PS_SD477_EnableMenus_Error
    Exit Sub
PS_SD477_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD477_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD477_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_SD477_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_SD477_FormItemEnabled
    Else
    End If
    Exit Sub
PS_SD477_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD477_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_SD477_DataValidCheck() As Boolean
On Error GoTo PS_SD477_DataValidCheck_Error
    PS_SD477_DataValidCheck = False
    Dim i As Long

    If oForm01.Items("BPLId").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm01.Items("BPLId").Click ct_Regular
        PS_SD477_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("DocDateFr").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "조회일자 Fr은 필수입니다.", bmt_Short, True
        oForm01.Items("DocDateFr").Click ct_Regular
        PS_SD477_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("DocDateTo").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "조회일자 To는 필수입니다.", bmt_Short, True
        oForm01.Items("DocDateTo").Click ct_Regular
        PS_SD477_DataValidCheck = False
        Exit Function
    End If
    PS_SD477_DataValidCheck = True
    Exit Function
PS_SD477_DataValidCheck_Error:
    PS_SD477_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_SD477_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PS_SD477_FormResize()
On Error GoTo PS_SD477_FormResize_Error
    
    Exit Sub
PS_SD477_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD477_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_SD477_Print_Report01()
On Error GoTo PS_SD477_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim sQry01                      As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Call ConnectODBC
    WinTitle = "품목별판매집계현황[PS_SD477]"
    ReportName = "PS_SD477_01.rpt"
    sQry = "EXEC PS_SD477_01 '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "', '" & oForm01.Items("ItemBsort").Specific.VALUE & "', '" & oForm01.Items("DocDateFr").Specific.VALUE & "', '" & oForm01.Items("DocDateTo").Specific.VALUE & "', "
    sQry = sQry + "'" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "', '" & Trim(oForm01.Items("Size").Specific.VALUE) & "'"

    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '// Formula 수식필드
    ReDim gRpt_Formula(3)
    ReDim gRpt_Formula_Value(3)
    
    gRpt_Formula(1) = "BPLId"
    If Trim(oForm01.Items("BPLId").Specific.VALUE) = "" Then
        gRpt_Formula_Value(1) = "전체"
    Else
        Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        sQry01 = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
        Call oRecordSet01.DoQuery(sQry01)
        gRpt_Formula_Value(1) = oRecordSet01.Fields(0).VALUE
        Set oRecordSet01 = Nothing
    End If
    
    gRpt_Formula(2) = "DocDateFr"
    gRpt_Formula_Value(2) = Format(oForm01.Items("DocDateFr").Specific.VALUE, "####-##-##")
    
    gRpt_Formula(3) = "DocDateTo"
    gRpt_Formula_Value(3) = Format(oForm01.Items("DocDateTo").Specific.VALUE, "####-##-##")
    
'    gRpt_Formula(4) = "ItmBsort"
'    If Trim(oForm01.Items("ItmBsort").Specific.Value) = "" Then
'        gRpt_Formula_Value(4) = "전체"
'    Else
'        Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'        sQry01 = "SELECT Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oForm01.Items("ItmBsort").Specific.Value) & "'"
'        Call oRecordSet01.DoQuery(sQry01)
'        gRpt_Formula_Value(4) = oRecordSet01.Fields(0).Value
'        Set oRecordSet01 = Nothing
'    End If
    
'    gRpt_Formula(5) = "ItmMsort"
'    If Trim(oForm01.Items("ItmMsort").Specific.Value) = "" Then
'        gRpt_Formula_Value(5) = "전체"
'    Else
'        Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'        sQry01 = "SELECT Name FROM [@PSH_ITMMSORT] WHERE U_Code = '" & Trim(oForm01.Items("ItmMsort").Specific.Value) & "'"
'        Call oRecordSet01.DoQuery(sQry01)
'        gRpt_Formula_Value(5) = oRecordSet01.Fields(0).Value
'        Set oRecordSet01 = Nothing
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_SD477_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD477_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub








