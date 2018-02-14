VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FX270"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_FX270.cls
'//  Module         : 레포트 > 고정자산 FX
'//  Desc           : 고정자산관리대장(개별)
'//  FormType       : PS_FX270
'//  Create Date    : 2017.08.07    HYS
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_FX270H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_FX270L As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FX270.srf")
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
    
    oFormUniqueID01 = "PS_FX270_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_FX270_CreateItems
    Call PS_FX270_ComboBox_Setting
    Call PS_FX270_Initial_Setting
    Call PS_FX270_CF_ChooseFromList
    Call PS_FX270_EnableMenus
    Call PS_FX270_SetDocument(oFromDocEntry01)
    Call PS_FX270_FormResize
    
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
                If PS_FX270_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_FX270_Print_Report01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_FX270" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "FixCode", "") '//사용자값활성
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
    
    Dim FixCode As String
    Dim SubCode As String
    Dim oQuery01 As String
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "BPLId") Then
                oQuery01 = "SELECT BPLName, BPLId FROM [OBPL] WHERE BPLid = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("BPLName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
          '  ElseIf (pval.ItemUID = "CardCode") Then
          '      oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
          '      oRecordSet01.DoQuery oQuery01
          '      oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
          '  ElseIf (pval.ItemUID = "ItemCode") Then
          '     oQuery01 = "SELECT ItemName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
          '      oRecordSet01.DoQuery oQuery01
          '      oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
          '  ElseIf (pval.ItemUID = "CntcCode") Then
          '      oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
          '      oRecordSet01.DoQuery oQuery01
          '      oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
            
            '//자산명 찿기
            If pval.ItemUID = "FixCode" Then
               FixCode = Left(oForm01.Items("FixCode").Specific.VALUE, 6)
               SubCode = Right(oForm01.Items("FixCode").Specific.VALUE, 3)
            
               sQry = "Select U_FixName From [@PS_FX005H] Where U_FixCode = '" & FixCode & "'"
               sQry = sQry & " and U_SubCode = '" & SubCode & "'"
               oRecordSet01.DoQuery sQry
               oForm01.Items("FixName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
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
        Call PS_FX270_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_FX270_FormResize
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


Private Function PS_FX270_CreateItems() As Boolean
On Error GoTo PS_FX270_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("ClasCode", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("FixCode", dt_SHORT_TEXT, 100)
    
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    Call oForm01.Items("ClasCode").Specific.DataBind.SetBound(True, "", "ClasCode")
    Call oForm01.Items("FixCode").Specific.DataBind.SetBound(True, "", "FixCode")
    
'    oForm01.Items("DocDateFr").Specific.VALUE = Format(Now, "YYYY") + "0101"
'
'    oForm01.Items("DocDateTo").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_FX270_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX270_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_FX270_ComboBox_Setting()
On Error GoTo PS_FX270_ComboBox_Setting_Error
     Call oForm01.Freeze(True)

        Call oForm01.Items("BPLId").Specific.ValidValues.Add("", "")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
        Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        
        '자산분류
        Call oForm01.Items("ClasCode").Specific.ValidValues.Add("%", "전체")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("ClasCode").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX001'", "", False, False)
        Call oForm01.Items("ClasCode").Specific.Select(0, psk_Index)
        
  
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_FX270_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX270_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX270_Initial_Setting()
On Error GoTo PS_FX270_Initial_Setting_Error
'  사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 ' 인수자
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD()
    Exit Sub
PS_FX270_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX270_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_FX270_CF_ChooseFromList()
On Error GoTo PS_FX270_CF_ChooseFromList_Error
    Exit Sub
PS_FX270_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX270_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX270_FormItemEnabled()
On Error GoTo PS_FX270_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_FX270_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX270_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX270_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_FX270_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    Call oForm01.Freeze(False)
    Exit Sub
PS_FX270_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_FX270_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_FX270_FormClear()
On Error GoTo PS_FX270_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_FX270'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_FX270_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX270_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX270_EnableMenus()
On Error GoTo PS_FX270_EnableMenus_Error
    Exit Sub
PS_FX270_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX270_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_FX270_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_FX270_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_FX270_FormItemEnabled

    Else

    End If
    Exit Sub
PS_FX270_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX270_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_FX270_DataValidCheck() As Boolean
On Error GoTo PS_FX270_DataValidCheck_Error
    PS_FX270_DataValidCheck = False
    Dim i As Long

    If oForm01.Items("BPLId").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm01.Items("BPLId").Click ct_Regular
        PS_FX270_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("FixCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "자산코드는 필수입니다.", bmt_Short, True
        oForm01.Items("FixCode").Click ct_Regular
        PS_FX270_DataValidCheck = False
        Exit Function
    End If
    
    PS_FX270_DataValidCheck = True
    Exit Function
PS_FX270_DataValidCheck_Error:
    PS_FX270_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_FX270_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PS_FX270_FormResize()
On Error GoTo PS_FX270_FormResize_Error
    
    Exit Sub
PS_FX270_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_FX270_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_FX270_Print_Report01()
On Error GoTo PS_FX270_Print_Report01_Error
    Dim ErrNum          As Integer
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim Sub_sQry                    As String
    Dim sQry01                      As String
    Dim BPLID                       As String
    Dim FixCode                     As String
    Dim HisType                     As String
    Dim oRecordSet        As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    
    '// 조회조건문
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    FixCode = Trim(oForm01.Items("FixCode").Specific.VALUE)
    
    '// 사진 6장 COPY
    Dim strid As String
    Dim fs As Object
    Dim P_Name1 As String
    Dim P_Name2 As String
    Dim P_Name3 As String
    Dim P_Name4 As String
    Dim P_Name5 As String
    Dim P_Name6 As String
    
'    P_Name1 = Left(FixCode, 6) + "_" + Right(FixCode, 3) + "_01"
'    P_Name2 = Left(FixCode, 6) + "_" + Right(FixCode, 3) + "_02"
'    P_Name3 = Left(FixCode, 6) + "_" + Right(FixCode, 3) + "_03"
'    P_Name4 = Left(FixCode, 6) + "_" + Right(FixCode, 3) + "_04"
'    P_Name5 = Left(FixCode, 6) + "_" + Right(FixCode, 3) + "_05"
'    P_Name6 = Left(FixCode, 6) + "_" + Right(FixCode, 3) + "_06"
    '사진은 PSHDB_IMG DB로 올림
    
'    Set fs = CreateObject("Scripting.FileSystemObject")
'
'    If fs.FileExists("\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name1 & ".bmp") = True Then
'
'        If fs.FileExists("\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC1.bmp") = True Then
'            Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC1.bmp"
'            Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC2.bmp"
'            Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC3.bmp"
'            Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC4.bmp"
'            Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC5.bmp"
'            Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC6.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name1 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC1.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name2 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC2.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name3 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC3.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name4 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC4.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name5 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC5.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name6 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC6.bmp"
'        Else
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name1 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC1.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name2 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC2.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name3 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC3.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name4 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC4.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name5 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC5.bmp"
'            fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\" & P_Name6 & ".bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC6.bmp"
'        End If
'    Else
'        Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC1.bmp"
'        Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC2.bmp"
'        Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC3.bmp"
'        Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC4.bmp"
'        Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC5.bmp"
'        Kill "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC6.bmp"
'        fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\NULL.bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC1.bmp"
'        fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\NULL.bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC2.bmp"
'        fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\NULL.bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC3.bmp"
'        fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\NULL.bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC4.bmp"
'        fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\NULL.bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC5.bmp"
'        fs.CopyFile "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\NULL.bmp", "\\191.1.1.220\ERP_Project\9999. Add_On\AddonSource\Asset_Pic\PIC6.bmp"
'    End If
'
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "고정자산관리대장(개별) [PS_FX270]"
    ReportName = "PS_FX270_01.rpt"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
    gRpt_Formula(1) = "BPLId"
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" & BPLID & "'"
    Call oRecordSet.DoQuery(sQry)
    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
    
    gRpt_Formula(2) = "Fix_Code"
    gRpt_Formula_Value(2) = FixCode
    
    
    '// SubReport
    ReDim gRpt_SRptSqry(2)
    ReDim gRpt_SRptName(2)
    
    ReDim gRpt_SFormula(2, 1)
    ReDim gRpt_SFormula_Value(2, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    gRpt_SFormula(2, 1) = ""
    gRpt_SFormula_Value(2, 1) = ""
    
    Sub_sQry = "EXEC [PS_FX270_02] '" & BPLID & "', '" & FixCode & "'"
    gRpt_SRptSqry(1) = Sub_sQry
    gRpt_SRptName(1) = "PS_FX270_01_SUB1"
    
    Sub_sQry = "EXEC [PS_FX270_03] '" & BPLID & "', '" & FixCode & "'"
    gRpt_SRptSqry(2) = Sub_sQry
    gRpt_SRptName(2) = "PS_FX270_01_SUB2"
    
    
    '/ Procedure 실행"
    sQry = "EXEC [PS_FX270_01] '" & BPLID & "', '" & FixCode & "'"
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PS_FX270_Print_Report01_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_FX270_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
     Sbo_Application.SetStatusBarMessage "PS_FX270_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub








