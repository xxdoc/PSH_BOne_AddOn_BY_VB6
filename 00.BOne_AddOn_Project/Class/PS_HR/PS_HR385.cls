VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR385"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File : PS_HR385.cls
'//  Module : 애드온>목표관리
'//  Desc : 사업부KPI목표출력
'//  FormType : PS_HR385
'//  Create Date : 2012.04.11
'//  Modifier : N.G.Y
'//  Modify Date :
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_HR385H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_HR385L As SAPbouiCOM.DBDataSource '등록라인
'Private oDS_PS_USERDS01 As SAPbouiCOM.DBDataSource    '등록라인

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
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR385.srf")
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
    
    oFormUniqueID01 = "PS_HR385_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_HR385_CreateItems
    Call PS_HR385_ComboBox_Setting
    Call PS_HR385_Initial_Setting
    Call PS_HR385_CF_ChooseFromList
    Call PS_HR385_EnableMenus
    Call PS_HR385_SetDocument(oFromDocEntry01)
    Call PS_HR385_FormResize
    
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
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_HR385_AddMatrixRow(oMat01.VisualRowCount)
                
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_HR385L.RemoveRecord(oDS_PS_HR385L.Size - 1)
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
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_HR385_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    Call PS_HR385_MTX01 '매트릭스에 데이터 로드
                End If
            End If
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_HR385_Print_Report01
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
        ElseIf pval.ItemUID = "Btn_Print" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
        
                Call PS_HR385_Print_Report01
            
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_HR385" Then
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
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '거래처코드 포맷서치 활성
        'Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '품목코드(작번) 포맷서치 활성
        
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
    
'        If pval.ItemUID = "Mat01" Then
'
'            If pval.Row = 0 Then
'
'                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
'                oMat01.FlushToDataSource
'
'            End If
'
'        End If
'
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
            If pval.ColUID = "DocEntry" Then
                Set oTempClass = New PS_MM070
                Call oTempClass.LoadForm(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
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
    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
        
'            If (pval.ItemUID = "CardCode") Then
'                oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
'                oRecordSet01.DoQuery oQuery01
'                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'            ElseIf (pval.ItemUID = "ItemCode") Then
'                oQuery01 = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
'                oRecordSet01.DoQuery oQuery01
'                oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'            ElseIf (pval.ItemUID = "CntcCode") Then
'                oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
'                oRecordSet01.DoQuery oQuery01
'                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'            End If
'
'            oForm01.Items(pval.ItemUID).Click ct_Regular
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
        Call PS_HR385_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_HR385_FormResize
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


Private Function PS_HR385_CreateItems() As Boolean
On Error GoTo PS_HR385_CreateItems_Error
    Call oForm01.Freeze(True)
'    Dim oQuery01 As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Set oDS_PS_HR385H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_HR385L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '매트릭스 초기화
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    '사업장_E
    
    
    '년월
    Call oForm01.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 200)
    Call oForm01.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")

    
    Call oForm01.Freeze(False)
    Exit Function
PS_HR385_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR385_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_HR385_ComboBox_Setting()
On Error GoTo PS_HR385_ComboBox_Setting_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Call oForm01.Freeze(True)
     
        '사업장 콤보박스 세팅_S
'        Call oForm01.Items("BPLId").Specific.ValidValues.Add("%", "전체")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
        Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        '사업장 콤보박스 세팅_E
        
        
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR385_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR385_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR385_Initial_Setting()
On Error GoTo PS_HR385_Initial_Setting_Error
    '사업장 사용자의 소속 사업장 선택
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '체크박스 설정
    'oForm01.Items("Chk01").Specific.Checked = True
    
    '년월 설정
    oForm01.Items("StdYear").Specific.VALUE = Format(Date, "YYYY")

 
    Exit Sub
PS_HR385_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR385_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_HR385_CF_ChooseFromList()
On Error GoTo PS_HR385_CF_ChooseFromList_Error
    Exit Sub
PS_HR385_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR385_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR385_FormItemEnabled()
On Error GoTo PS_HR385_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR385_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR385_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR385_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_HR385_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_HR385L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_HR385L.Offset = oRow
'    oDS_PS_HR385L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR385_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR385_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR385_FormClear()
On Error GoTo PS_HR385_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_HR385'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_HR385_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR385_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR385_EnableMenus()
On Error GoTo PS_HR385_EnableMenus_Error
    Exit Sub
PS_HR385_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR385_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR385_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_HR385_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_HR385_FormItemEnabled

    Else

    End If
    Exit Sub
PS_HR385_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR385_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_HR385_DataValidCheck() As Boolean
On Error GoTo PS_HR385_DataValidCheck_Error
    PS_HR385_DataValidCheck = False
    Dim i As Long

'    If oForm01.Items("DocDateTo").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "견적일자 To은 필수입니다.", bmt_Short, True
'        oForm01.Items("DocDateTo").Click ct_Regular
'        PS_HR385_DataValidCheck = False
'        Exit Function
'    End If

    PS_HR385_DataValidCheck = True
    Exit Function
PS_HR385_DataValidCheck_Error:
    PS_HR385_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_HR385_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_HR385_MTX01()
On Error GoTo PS_HR385_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId As String '사업장
    Dim StdYear As String '년월
    
    
    
    BPLId = oForm01.Items("BPLId").Specific.Selected.VALUE '사업장
    StdYear = oForm01.Items("StdYear").Specific.VALUE '년월
        
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_HR385_01 '" & BPLId & "','" & StdYear & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_HR385_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_HR385L.InsertRecord(loopCount)
        End If
        oDS_PS_HR385L.Offset = loopCount

        Call oDS_PS_HR385L.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_HR385L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("FETask1").VALUE) '사업부 중점실행과제
        Call oDS_PS_HR385L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("TeamNm").VALUE) '팀명
        Call oDS_PS_HR385L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("FETask2").VALUE) '팀중점실행과제
        Call oDS_PS_HR385L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("KPINm2").VALUE) '팀KPI명
        Call oDS_PS_HR385L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("RspNm").VALUE) '담당명
        Call oDS_PS_HR385L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("FETask3").VALUE) '담당중점실행과제
        Call oDS_PS_HR385L.setValue("U_ColReg07", loopCount, RecordSet01.Fields("KPINm3").VALUE) '담당KPI명
        Call oDS_PS_HR385L.setValue("U_ColReg08", loopCount, RecordSet01.Fields("FECts3").VALUE) '담당중점실행과제
        Call oDS_PS_HR385L.setValue("U_ColQty01", loopCount, RecordSet01.Fields("CrtYTrgt3").VALUE) '담당금년목표


        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next


    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_HR385_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_HR385_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR385_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_HR385_FormResize()
On Error GoTo PS_HR385_FormResize_Error
    
    Exit Sub
PS_HR385_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR385_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR385_Print_Report01()
On Error GoTo Print_Query_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String

    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim Sub_sQry        As String
    
'    Dim BPLId           As String
'    Dim CardCode       As String
'    Dim Pumtxt        As String


    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    '// 조회조건문
    
'    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
'    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
'    Pumtxt = Trim(oForm01.Items("Pumtxt").Specific.VALUE)
'
'    If Pumtxt = "" Then Pumtxt = "%"
'    If CardCode = "" Then CardCode = "%"
'
    
    Dim BPLId As String '사업장
    Dim StdYear As String '년월
    
    BPLId = oForm01.Items("BPLId").Specific.Selected.VALUE '사업장
    StdYear = oForm01.Items("StdYear").Specific.VALUE '년월
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_HR385] 레포트"
    ReportName = "PS_HR385_01.rpt"
 
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
    sQry = "EXEC PS_HR385_01 '" & BPLId & "','" & StdYear & "'"

    Call oRecordSet.DoQuery(sQry)
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
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
