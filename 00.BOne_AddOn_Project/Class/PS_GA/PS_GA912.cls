VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA912"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_GA912.cls
'//  Module         : GA
'//  Desc           : 전산기기반입등록
'//  FormType       : PS_GA912
'//  Create Date    : 2014.07.31
'//  Copyright  (c) Morning Data
'//********************************************************************************

'//사용방법
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_GA912
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseoVisitNo01 As String

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
Public Sub LoadForm(Optional ByVal oVisitNo As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA912.srf")
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
    
    oFormUniqueID01 = "PS_GA912_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
'    Set oBaseForm01 = oForm02
    oBaseoVisitNo01 = oVisitNo
    
    
    Call PS_GA912_CreateItems
    Call PS_GA912_ComboBox_Setting
    Call PS_GA912_CF_ChooseFromList
    Call PS_GA912_FormItemEnabled
    Call PS_GA912_EnableMenus
    '//Call PS_GA912_FormClear '//UDO방식일때
    '//Call PS_GA912_AddMatrixRow(0, True) '//UDO방식일때
    Call PS_GA912_MTX01
    oForm01.Items("VisitNo").Specific.VALUE = oVisitNo
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.EnableMenu ("1282"), True
    
    oForm01.Visible = True
    
    '최초 커서위치
    oForm01.ActiveItem = "Div"
    
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
                 
                 Call PS_GA912_Form_ini
            
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_GA912_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_GA912_FormItemEnabled '//UDO방식
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
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Then
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
'        If pval.ItemUID = "Btn_ret" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_GA912_MTX01
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
'        If pval.ItemUID = "Button02" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_GA912_SetBaseForm '//부모폼에입력
'                If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
'                    Call oForm01.Close
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
        
        If pval.ItemUID = "Btn_ret" Then
           Call PS_GA912_MTX01
        End If
               
        If pval.ItemUID = "Btn_save" Then
           Call PS_GA912_SAVE
        End If
                             
        If pval.ItemUID = "Btn_delete" Then
           Call PS_GA912_Delete
           Call PS_GA912_FormItemEnabled
        End If
        
        
        
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_GA912_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA912" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_GA912_FormItemEnabled
'                    Call PS_GA912_FormClear '//UDO방식일때
'                    Call PS_GA912_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_GA912_FormItemEnabled
'                End If
'            End If
'        End If
    End If
    Exit Sub
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    If pval.BeforeAction = True Then
   
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
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
        If pval.ItemUID = "Grid01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row > 0 Then
                    
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "Grid01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row >= 0 Then
                   PS_GA912_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Grid01" Then
            If pval.Row = -1 Then
'                oGrid01.Columns(pval.ColUID).TitleObject.Sortable = True
            Else
              
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
    Dim oQuery As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
    
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
        Call PS_GA912_FormItemEnabled
        '//Call PS_GA912_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
'        oForm01.Items("Mat01").Top = 70
'        oForm01.Items("Mat01").Height = (oForm01.Height / 2) - 70
'        oForm01.Items("Mat01").Left = 7
'        oForm01.Items("Mat01").Width = oForm01.Width - 21
'        oForm01.Items("Mat02").Top = (oForm01.Height / 2) + 10
'        oForm01.Items("Mat02").Height = (oForm01.Height / 2) - 75
'        oForm01.Items("Mat02").Left = 7
'        oForm01.Items("Mat02").Width = oForm01.Width - 21
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    Dim oDataTable01  As SAPbouiCOM.DataTable
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "ItemCode") Then
'            Set oDataTable01 = pval.SelectedObjects
'            If oDataTable01 Is Nothing Then
'            Else
'                oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
'                oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
'            End If
'        End If
'        oForm01.Update
    End If
    Set oDataTable01 = Nothing
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Then
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
        Set oGrid01 = Nothing
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
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
    '        For i = 1 To oMat01.VisualRowCount
    '            oMat01.Columns("COL01").Cells(i).Specific.Value = i
    '        Next i
    '        oMat01.FlushToDataSource
    '        Call oDS_PS_GA912L.RemoveRecord(oDS_PS_GA912L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_GA912_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_GA912_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_GA912_CreateItems() As Boolean
On Error GoTo PS_GA912_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    'oGrid01.SelectionMode = ms_NotSupported
    
    '순번
    Call oForm01.DataSources.UserDataSources.Add("LineNum", dt_SHORT_NUMBER)
    oForm01.Items("LineNum").Specific.DataBind.SetBound True, "", "LineNum"
    oForm01.Items("LineNum").Specific.VALUE = 0
    
    '물품명
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 50)
    oForm01.Items("ItemName").Specific.DataBind.SetBound True, "", "ItemName"
    
    '모델규격
    Call oForm01.DataSources.UserDataSources.Add("Size", dt_SHORT_TEXT, 50)
    oForm01.Items("Size").Specific.DataBind.SetBound True, "", "Size"
    
    '구역
    Call oForm01.DataSources.UserDataSources.Add("Place", dt_SHORT_TEXT, 20)
    oForm01.Items("Place").Specific.DataBind.SetBound True, "", "Place"
    
    '주변기기
    Call oForm01.DataSources.UserDataSources.Add("EtcNm", dt_SHORT_TEXT, 20)
    oForm01.Items("EtcNm").Specific.DataBind.SetBound True, "", "EtcNm"
    
    '출입시작일자
    Call oForm01.DataSources.UserDataSources.Add("DateFrom", dt_DATE)
    Call oForm01.Items("DateFrom").Specific.DataBind.SetBound(True, "", "DateFrom")
    oForm01.DataSources.UserDataSources.Item("DateFrom").VALUE = Format(Now, "YYYYMMDD")
    
    '출입종료일자
    Call oForm01.DataSources.UserDataSources.Add("DateTo", dt_DATE)
    Call oForm01.Items("DateTo").Specific.DataBind.SetBound(True, "", "DateTo")
    oForm01.DataSources.UserDataSources.Item("DateTo").VALUE = Format(Now, "YYYYMMDD")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function

PS_GA912_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA912_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_GA912_ComboBox_Setting()
On Error GoTo PS_GA912_ComboBox_Setting_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
        
    '분야
    oForm01.Items("Div").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Div").Specific, "SELECT b.U_Code, b.U_CodeNm FROM [@PS_GA050L] b Where b.Code = '24' and b.U_UseYN = 'Y'", "", False, False)
'    oForm01.Items("CarDiv").Specific.Select "0", psk_Index
        
    Call oForm01.Freeze(False)
        
    Exit Sub
PS_GA912_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA912_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA912_CF_ChooseFromList()
On Error GoTo PS_GA912_CF_ChooseFromList_Error
    '//ChooseFromList 설정
'    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
'    Dim oCons               As SAPbouiCOM.Conditions
'    Dim oCon                As SAPbouiCOM.Condition
'    Dim oCFL                As SAPbouiCOM.ChooseFromList
'    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
'    Dim oEdit               As SAPbouiCOM.EditText
'    Dim oColumn             As SAPbouiCOM.Column
'
'    Set oEdit = oForm01.Items("ItemCode").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = "4"
'    oCFLCreationParams.uniqueID = "CFLITEMCD"
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
'    oEdit.ChooseFromListUID = "CFLITEMCD"
'    oEdit.ChooseFromListAlias = "ItemCode"
    Exit Sub
PS_GA912_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA912_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA912_FormItemEnabled()
On Error GoTo PS_GA912_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseoVisitNo01 = "" Then
        
        Else
            oForm01.Items("VisitNo").Specific.VALUE = oBaseoVisitNo01
            
        End If
        '//각모드에따른 아이템설정
        '//Call PS_GA912_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA912_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA912_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA912_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA912_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_GA912L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_GA912L.Offset = oRow
'    oDS_PS_GA912L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA912_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA912_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA912_FormClear()
On Error GoTo PS_GA912_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_GA912'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_GA912_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA912_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA912_EnableMenus()
On Error GoTo PS_GA912_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_GA912_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA912_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_GA912_DataValidCheck() As Boolean
On Error GoTo PS_GA912_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_GA912_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_GA912_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_GA912_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_GA912_FormClear
    Exit Function
PS_GA912_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA912_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_GA912_MTX01()
On Error GoTo PS_GA912_MTX01_Error
    '//그리드에 데이터 로드 (조회)
    Call oForm01.Freeze(True)
    Dim i           As Long
    Dim sQry        As String
    Dim VisitNo     As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    VisitNo = oForm01.Items("VisitNo").Specific.VALUE
    
    sQry = " EXEC [PS_GA912_01]  '" & VisitNo & "'"

    oGrid01.DataTable.Clear

    Call oForm01.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    Call RecordSet01.DoQuery(sQry)
    If RecordSet01.RecordCount = 0 Then
       Call MDC_Com.MDC_GF_Message("자료가 존재하지 않습니다. ", "E")
       GoTo PS_GA912_MTX01_Exit
    End If

    oGrid01.AutoResizeColumns
    oForm01.Update

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA912_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA912_MTX01_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA912_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PS_GA912_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    
    '//그리드 자료를 head에 로드
  
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As Integer
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_GA912_MTX02_Error

    Call oForm01.Freeze(True)
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Param01 = oForm01.Items("VisitNo").Specific.VALUE
    Param02 = oGrid01.DataTable.Columns("순번").Cells(oRow).VALUE
    
    
    sQry = "EXEC PS_GA912_02 '" & Param01 & "', '" & Param02 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        Call PS_GA912_Form_ini
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_GA912_MTX02_Exit
    End If
    
    oForm01.DataSources.UserDataSources.Item("LineNum").VALUE = oRecordSet.Fields("LineNum").VALUE
    oForm01.DataSources.UserDataSources.Item("ItemName").VALUE = oRecordSet.Fields("ItemName").VALUE
    oForm01.DataSources.UserDataSources.Item("Size").VALUE = oRecordSet.Fields("Size").VALUE
    oForm01.DataSources.UserDataSources.Item("Place").VALUE = oRecordSet.Fields("Place").VALUE
    oForm01.DataSources.UserDataSources.Item("EtcNm").VALUE = oRecordSet.Fields("EtcNm").VALUE
    oForm01.DataSources.UserDataSources.Item("DateFrom").VALUE = oRecordSet.Fields("DateFrom").VALUE
    oForm01.DataSources.UserDataSources.Item("DateTo").VALUE = oRecordSet.Fields("DateTo").VALUE
    
    Set oCombo = oForm01.Items("Div").Specific
    oCombo.Select oRecordSet.Fields("Div").VALUE, psk_ByValue
    
    '//Key Disable
    oForm01.ActiveItem = "Div"
    
'    oForm01.Items("Div").Enabled = False
'    oForm01.Items("Grade").Enabled = False
'    oForm01.Items("Num").Enabled = False
'
    oForm01.Update
    
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA912_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_GA912_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA912_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PS_GA912_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim Seqncom As Integer
    Dim sQry    As String
    Dim VisitNo$, Div$, ItemName$, Size$, Place$, EtcNm$, DateFrom$, DateTo$
    Dim LineNum As Integer
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    On Error GoTo PS_GA912_SAVE_Error
    
    VisitNo = oForm01.Items("VisitNo").Specific.VALUE
    LineNum = oForm01.Items("LineNum").Specific.VALUE
    
    Div = oForm01.Items("Div").Specific.VALUE
    ItemName = oForm01.Items("ItemName").Specific.VALUE
    Size = oForm01.Items("Size").Specific.VALUE
    Place = oForm01.Items("Place").Specific.VALUE
    EtcNm = oForm01.Items("EtcNm").Specific.VALUE
    DateFrom = oForm01.Items("DateFrom").Specific.VALUE
    DateTo = oForm01.Items("DateTo").Specific.VALUE
    
    
    Call oForm01.Freeze(True)
    
    If Trim(VisitNo) = "" Then
        Call MDC_Com.MDC_GF_Message("문서번호에러. 확인바랍니다..", "E")
        GoTo PS_GA912_SAVE_Exit
    End If
    
    If LineNum > 0 Then
        '//갱신
        sQry = "Update [Z_PS_GA910M] set "
        sQry = sQry + "Div = '" & Div & "',"
        sQry = sQry + "ItemName = '" & ItemName & "',"
        sQry = sQry + "Size = '" & Size & "',"
        sQry = sQry + "Place = '" & Place & "',"
        sQry = sQry + "EtcNm = '" & EtcNm & "',"
        sQry = sQry + "DateFrom = '" & DateFrom & "',"
        sQry = sQry + "DateTo = '" & DateTo & "'"
        
        sQry = sQry + " Where VisitNo = '" & VisitNo & "' And LineNum = '" & LineNum & "'"
        
        oRecordSet.DoQuery sQry
        
    Else
        '//신규
        '순번 계산
        sQry = " Select Isnull(Max(LineNum),0) From [Z_PS_GA910M] Where VisitNo = '" & VisitNo & "'"
        oRecordSet.DoQuery sQry
        Seqncom = oRecordSet.Fields(0).VALUE
        Seqncom = Seqncom + 1
        '------------------------------------------------------
        
        sQry = "INSERT INTO [Z_PS_GA910M]"
        sQry = sQry & " ("
        sQry = sQry & "VisitNo,"
        sQry = sQry & "LineNum,"
        sQry = sQry & "Div,"
        sQry = sQry & "ItemName,"
        sQry = sQry & "Size,"
        sQry = sQry & "Place,"
        sQry = sQry & "EtcNm,"
        sQry = sQry & "DateFrom,"
        sQry = sQry & "DateTo"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("

        sQry = sQry & "'" & VisitNo & "',"
        sQry = sQry & "'" & Seqncom & "',"
        sQry = sQry & "'" & Div & "',"
        sQry = sQry & "'" & ItemName & "',"
        sQry = sQry & "'" & Size & "',"
        sQry = sQry & "'" & Place & "',"
        sQry = sQry & "'" & EtcNm & "',"
        sQry = sQry & "'" & DateFrom & "',"
        sQry = sQry & "'" & DateTo & "'"
        
        sQry = sQry & " ) "

        oRecordSet.DoQuery sQry
    End If
    
    Call PS_GA912_FormItemEnabled
    
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)

    Call PS_GA912_MTX01
    Call PS_GA912_Form_ini
    
    Exit Sub
    
PS_GA912_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_GA912_SAVE_Error:
    Call oForm01.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_GA912_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_GA912_Delete()
    '//선택된 자료 삭제
    Dim VisitNo As String
    Dim LineNum As Integer
    Dim i As Integer
    Dim Cnt As Integer
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry    As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_GA912_Delete_Error

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    VisitNo = oForm01.Items("VisitNo").Specific.VALUE
    LineNum = oForm01.Items("LineNum").Specific.VALUE
    
    If LineNum < 1 Then
        Call MDC_Com.MDC_GF_Message("순번을 확인바랍니다..", "E")
        GoTo PS_GA912_Delete_Exit
    End If
        
    sQry = " Select Count(*) From [Z_PS_GA910M] Where VisitNo = '" & VisitNo & "' And LineNum = '" & LineNum & "'"
    oRecordSet.DoQuery sQry
    
    Cnt = oRecordSet.Fields(0).VALUE
    If Cnt > 0 Then
        
        If Sbo_Application.MessageBox(" 선택한라인을 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
           sQry = "Delete From [Z_PS_GA910M] Where VisitNo = '" & VisitNo & "' And LineNum = '" & LineNum & "'"
           oRecordSet.DoQuery sQry
        End If
    Else
        Call MDC_Com.MDC_GF_Message("조회후 삭제하십시요.", "E")
        GoTo PS_GA912_Delete_Exit
    End If
    
    Call oForm01.Freeze(False)
        
    Call PS_GA912_MTX01
    Call PS_GA912_Form_ini
    
    Set oRecordSet = Nothing
    
    Exit Sub
PS_GA912_Delete_Exit:
    Set oRecordSet = Nothing
        
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA912_Delete_Error:
    Set oRecordSet = Nothing
   
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA912_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA912_Form_ini()
    Dim oCombo      As SAPbouiCOM.ComboBox

    '화면청소 ㅋ
    oForm01.DataSources.UserDataSources.Item("LineNum").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("ItemName").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("Size").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("Place").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("EtcNm").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("DateFrom").VALUE = Format(Now, "YYYYMMDD")
    oForm01.DataSources.UserDataSources.Item("DateTo").VALUE = Format(Now, "YYYYMMDD")
    Set oCombo = oForm01.Items("Div").Specific
    oCombo.Select 0, psk_Index
    
    oForm01.ActiveItem = "Div"
    oForm01.Update

End Sub







