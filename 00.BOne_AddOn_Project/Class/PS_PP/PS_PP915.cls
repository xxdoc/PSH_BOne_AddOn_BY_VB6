VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP915"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_PP915.cls
'//  Module : 생산관리 > 애드온 > 생산관리 > 장비가동시간관리
'//  Desc : 설비종합효율-기초정보조회
'//  FormType : PS_PP915
'//  Create Date(Start) : 2014.09.26
'//  Create Date(End) : 2014.10.30
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid
Public oGrid02 As SAPbouiCOM.Grid

Public oDS_PS_PP915A As SAPbouiCOM.DataTable
Public oDS_PS_PP915B As SAPbouiCOM.DataTable

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String
Public oBaseItmBsort01 As String

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
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String, Optional ByVal oItmBsort02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP915.srf")
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
    
    oFormUniqueID01 = "PS_PP915_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Set oBaseForm01 = oForm02
    oBaseItemUID01 = oItemUID02
    oBaseColUID01 = oColUID02
    oBaseColRow01 = oColRow02
    oBaseTradeType01 = oTradeType02
    oBaseItmBsort01 = oItmBsort02
    
    Call PS_PP915_CreateItems
    Call PS_PP915_ComboBox_Setting
'    Call PS_PP915_CF_ChooseFromList
'    Call PS_PP915_FormItemEnabled
'    Call PS_PP915_EnableMenus
    '//Call PS_PP915_FormClear '//UDO방식일때
    '//Call PS_PP915_AddMatrixRow(0, True) '//UDO방식일때
    
    oForm01.Items("FrYM").Specific.VALUE = Format(Now, "YYYYMM")
    oForm01.Items("ToYM").Specific.VALUE = Format(Now, "YYYYMM")
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    
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
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_PP915_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_PP915_FormItemEnabled '//UDO방식
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
    
        If pval.ItemUID = "BtnSearch" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_PP915_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_PP915_Print_Report01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        End If
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_PP915_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
    
        '폴더를 사용할 때는 필수 소스_S
        'Folder01이 선택되었을 때
        If pval.ItemUID = "Folder01" Then
            
            oForm01.PaneLevel = 1
            
        End If
        
        'Folder02가 선택되었을 때
        If pval.ItemUID = "Folder02" Then
            
            oForm01.PaneLevel = 2
            
        End If
        '폴더를 사용할 때는 필수 소스_E
    
        If pval.ItemUID = "PS_PP915" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP915_FormItemEnabled
'                    Call PS_PP915_FormClear '//UDO방식일때
'                    Call PS_PP915_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP915_FormItemEnabled
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "MachCode", "")
        
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
        Call PS_PP915_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
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
                If oGrid01.Rows.SelectedRows.Count > 0 Then
                
                    Call PS_PP915_GetDetail
                
'                    Call PS_PP915_SetBaseForm '//부모폼에입력
'                    If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
'                        Call oForm01.Close
'                    End If
                Else
                    BubbleEvent = False
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
       
    ElseIf pval.BeforeAction = False Then
        Call PS_PP915_FlushToItemValue(pval.ItemUID)
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
        Call PS_PP915_FormItemEnabled
        '//Call PS_PP915_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
        Call PS_PP915_FormResize
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
        'If (pval.ItemUID = "ItemCode") Then
         '   Set oDataTable01 = pval.SelectedObjects
        '    If oDataTable01 Is Nothing Then
        '    Else
              '  oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
         '     '  oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
         '   End If
       ' End If
        oForm01.Update
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
        Set oGrid02 = Nothing
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
    '        Call oDS_PS_PP915L.RemoveRecord(oDS_PS_PP915L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_PP915_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_PP915_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP915_CreateItems() As Boolean
On Error GoTo PS_PP915_CreateItems_Error

    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    'Dim C_Date   As Date
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    Set oGrid02 = oForm01.Items("Grid02").Specific
    
    Call oForm01.DataSources.DataTables.Add("PS_PP915A")
    Call oForm01.DataSources.DataTables.Add("PS_PP915B")
    
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("PS_PP915A")
    oGrid02.DataTable = oForm01.DataSources.DataTables.Item("PS_PP915B")
    
    Set oDS_PS_PP915A = oForm01.DataSources.DataTables.Item("PS_PP915A")
    Set oDS_PS_PP915B = oForm01.DataSources.DataTables.Item("PS_PP915B")
    
    'oGrid01.SelectionMode = ms_NotSupported
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")

    '작업구분
    Call oForm01.DataSources.UserDataSources.Add("OrdGbn", dt_SHORT_TEXT, 5)
    Call oForm01.Items("OrdGbn").Specific.DataBind.SetBound(True, "", "OrdGbn")
    
    '기준년월(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrYM", dt_SHORT_TEXT, 6)
    Call oForm01.Items("FrYM").Specific.DataBind.SetBound(True, "", "FrYM")
    
    '기준년월(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToYM", dt_SHORT_TEXT, 6)
    Call oForm01.Items("ToYM").Specific.DataBind.SetBound(True, "", "ToYM")
    
    '장비코드
    Call oForm01.DataSources.UserDataSources.Add("MachCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("MachCode").Specific.DataBind.SetBound(True, "", "MachCode")
    
    '장비명
    Call oForm01.DataSources.UserDataSources.Add("MachName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("MachName").Specific.DataBind.SetBound(True, "", "MachName")
    
'    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP915_CreateItems_Error:
'    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP915_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP915_ComboBox_Setting()
On Error GoTo PS_PP915_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", "1", False, False)
    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '작업구분_S
    sQry = "            SELECT      Code,"
    sQry = sQry & "                 Name"
    sQry = sQry & "  FROM       [@PSH_ITMBSORT]"
    sQry = sQry & "  WHERE      U_PudYN = 'Y'"
    sQry = sQry & "  ORDER BY  Code"
    Call oForm01.Items("OrdGbn").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdGbn").Specific, sQry, "%", False, False)
    '작업구분_E
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP915_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP915_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP915_CF_ChooseFromList()
On Error GoTo PS_PP915_CF_ChooseFromList_Error
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
PS_PP915_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP915_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP915_FormItemEnabled()
On Error GoTo PS_PP915_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseItmBsort01 = "" Then
        
        Else
        
        End If
        '//각모드에따른 아이템설정
        '//Call PS_PP915_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP915_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP915_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP915_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP915_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_PP915L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_PP915L.Offset = oRow
'    oDS_PS_PP915L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP915_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP915_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP915_DataValidCheck() As Boolean
On Error GoTo PS_PP915_DataValidCheck_Error
    Dim i As Long
    PS_PP915_DataValidCheck = True
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_PP915_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_PP915_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_PP915_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
'    Call PS_PP915_FormClear
    Exit Function
PS_PP915_DataValidCheck_Error:
    PS_PP915_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP915_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP915_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_PP915_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim TeamCode As String
    Dim RspCode As String
    
    Select Case oUID
        
        Case "MachCode"
        
            oForm01.Items("MachName").Specific.VALUE = MDC_GetData.Get_ReData("U_MachName", "U_MachCode", "[@PS_PP130H]", "'" & Trim(oForm01.Items("MachCode").Specific.VALUE) & "'") '장비명
            
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_PP915_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_PP915_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub


Private Sub PS_PP915_MTX01()
On Error GoTo PS_PP915_MTX01_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim Query02 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    '작업일보 조회_S
    Dim BPLId As String '사업장
    Dim OrdGbn As String '작업구분
    Dim FrYM As String '기준년월(시작)
    Dim ToYM As String '기준년월(종료)
    Dim MachCode As String '장비코드
    
    BPLId = Trim(oForm01.Items("BPLID").Specific.Selected.VALUE) '사업장
    OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.Selected.VALUE) '작업구분
    FrYM = Trim(oForm01.Items("FrYM").Specific.VALUE) '기준년월(시작)
    ToYM = Trim(oForm01.Items("ToYM").Specific.VALUE) '기준년월(종료)
    MachCode = Trim(oForm01.Items("MachCode").Specific.VALUE) '장비코드
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Query01 = "              EXEC PS_PP915_01 '"
    Query01 = Query01 & BPLId & "','"
    Query01 = Query01 & OrdGbn & "','"
    Query01 = Query01 & FrYM & "','"
    Query01 = Query01 & ToYM & "','"
    Query01 = Query01 & MachCode & "'"
        
    Call oGrid01.DataTable.Clear
    Call oDS_PS_PP915A.ExecuteQuery(Query01)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    oGrid01.Columns(12).RightJustified = True
    oGrid01.Columns(13).RightJustified = True
    oGrid01.Columns(14).RightJustified = True
    oGrid01.Columns(17).RightJustified = True
    oGrid01.Columns(20).RightJustified = True
    oGrid01.Columns(24).RightJustified = True
    oGrid01.Columns(27).RightJustified = True
    oGrid01.Columns(28).RightJustified = True
    oGrid01.Columns(29).RightJustified = True
    oGrid01.Columns(30).RightJustified = True
    oGrid01.Columns(31).RightJustified = True
    oGrid01.Columns(32).RightJustified = True
    oGrid01.Columns(33).RightJustified = True
    oGrid01.Columns(34).RightJustified = True
    oGrid01.Columns(35).RightJustified = True
    oGrid01.Columns(36).RightJustified = True
    oGrid01.Columns(37).RightJustified = True
    oGrid01.Columns(38).RightJustified = True
    oGrid01.Columns(39).RightJustified = True
    oGrid01.Columns(40).RightJustified = True
    oGrid01.Columns(41).RightJustified = True
    oGrid01.Columns(42).RightJustified = True
    oGrid01.Columns(43).RightJustified = True
    oGrid01.Columns(44).RightJustified = True
    oGrid01.Columns(45).RightJustified = True
    oGrid01.Columns(46).RightJustified = True
    oGrid01.Columns(47).RightJustified = True
    oGrid01.Columns(48).RightJustified = True
    oGrid01.Columns(49).RightJustified = True
    oGrid01.Columns(50).RightJustified = True
    oGrid01.Columns(51).RightJustified = True
    oGrid01.Columns(52).RightJustified = True
    oGrid01.Columns(53).RightJustified = True
    oGrid01.Columns(54).RightJustified = True
    oGrid01.Columns(55).RightJustified = True
    '작업일보 조회_E
    
    '장비가동시간_S
    Query02 = "              EXEC PS_PP915_02 '"
    Query02 = Query02 & BPLId & "','"
    Query02 = Query02 & OrdGbn & "','"
    Query02 = Query02 & FrYM & "','"
    Query02 = Query02 & ToYM & "','"
    Query02 = Query02 & MachCode & "'"

    Call oGrid02.DataTable.Clear
    Call oDS_PS_PP915B.ExecuteQuery(Query02)
'    oGrid05.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    oGrid02.Columns(8).RightJustified = True
    oGrid02.Columns(11).RightJustified = True
    oGrid02.Columns(14).RightJustified = True
    oGrid02.Columns(17).RightJustified = True
    oGrid02.Columns(20).RightJustified = True
    '장비가동시간_E
    
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid01.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP915_MTX01_Exit
    End If

    oGrid01.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP915_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP915_MTX01_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_PP915_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP915_DI_API() As Boolean
'On Error GoTo PS_PP915_DI_API_Error
'    PS_PP915_DI_API = True
'    Dim i, j As Long
'    Dim oDIObject As SAPbobsCOM.Documents
'    Dim RetVal As Long
'    Dim LineNumCount As Long
'    Dim ResultDocNum As Long
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Company.StartTransaction
'
'    ReDim ItemInformation(0)
'    ItemInformationCount = 0
'    For i = 1 To oMat01.VisualRowCount
'        ReDim Preserve ItemInformation(ItemInformationCount)
'        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).Check = False
'        ItemInformationCount = ItemInformationCount + 1
'    Next
'
'    LineNumCount = 0
'    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
'    oDIObject.BPL_IDAssignedToInvoice = oForm01.Items("BPLID").Specific.Selected.Value
'    oDIObject.CardCode = oForm01.Items("CardCode").Specific.Value
'    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.Value, "&&&&-&&-&&")
'    For i = 0 To UBound(ItemInformation)
'        If ItemInformation(i).Check = True Then
'            GoTo Continue_First
'        End If
'        If i <> 0 Then
'            oDIObject.Lines.Add
'        End If
'        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
'        oDIObject.Lines.WarehouseCode = oForm01.Items("WhsCode").Specific.Value
'        oDIObject.Lines.BaseType = "22"
'        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
'        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
'        For j = i To UBound(ItemInformation)
'            If ItemInformation(j).Check = True Then
'                GoTo Continue_Second
'            End If
'            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
'                GoTo Continue_Second
'            End If
'            '//같은것
'            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
'            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
'            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
'            oDIObject.Lines.BatchNumbers.Add
'            ItemInformation(j).PDN1No = LineNumCount
'            ItemInformation(j).Check = True
'Continue_Second:
'        Next
'        LineNumCount = LineNumCount + 1
'Continue_First:
'    Next
'    RetVal = oDIObject.Add
'    If RetVal = 0 Then
'        ResultDocNum = Sbo_Company.GetNewObjectKey
'        For i = 0 To UBound(ItemInformation)
'            Call oDS_PS_PP915L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_PP915L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
'    Else
'        GoTo PS_PP915_DI_API_Error
'    End If
'
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_Commit
'    End If
'    oMat01.LoadFromDataSource
'    oMat01.AutoResizeColumns
'
'    Set oDIObject = Nothing
'    Exit Function
'PS_PP915_DI_API_DI_Error:
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
'    PS_PP915_DI_API = False
'    Set oDIObject = Nothing
'    Exit Function
'PS_PP915_DI_API_Error:
'    Sbo_Application.SetStatusBarMessage "PS_PP915_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    PS_PP915_DI_API = False
End Function

Private Sub PS_PP915_SetBaseForm()
On Error GoTo PS_PP915_SetBaseForm_Error
    Dim i As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    If oBaseForm01 Is Nothing Then
    '//DoNothing
    Else
    
    End If
    Exit Sub
PS_PP915_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP915_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP915_FormResize()
On Error GoTo PS_PP915_FormResize_Error
    
    '그룹박스 크기 동적 할당
    oForm01.Items("GrpBox01").Height = oForm01.Items("Grid01").Height + 30
    oForm01.Items("GrpBox01").Width = oForm01.Items("Grid01").Width + 30
    
    Exit Sub
PS_PP915_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP915_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP915_Print_Report01()
On Error GoTo Print_Query_Error
    Dim DocNum As String
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim BPLId As String
    Dim StdYM As String
    Dim CntcCode As String
    
    BPLId = oForm01.Items("BPLID").Specific.VALUE
    StdYM = oForm01.Items("StdYM").Specific.VALUE
    CntcCode = MDC_PS_Common.User_MSTCOD
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP915] 레포트"
    
    ReportName = "PS_PP915_01.rpt"
    
    '쿼리
    sQry = "         EXEC PS_PP915_02 '"
    sQry = sQry & BPLId & "','"
    sQry = sQry & StdYM & "','"
    sQry = sQry & CntcCode & "'"
 
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

'    Call oRecordSet.DoQuery(sQry)
'
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_PP915_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub PS_PP915_GetDetail()
On Error GoTo PS_PP915_GetDetail_Error

    Dim loopCount1 As Integer
    
'    Dim BPLID As Integer
'    Dim CntcCode As String
'    Dim CntcName As String
'    Dim FrDt As String
'    Dim ToDt As String
    
    Dim MngNo As String
    
    Dim oTempClass As Variant

    For loopCount1 = 0 To oGrid01.Rows.Count - 1
        If oGrid01.Rows.IsSelected(loopCount1) = True Then
            
            '/ Procedure 실행
'             sQry = "EXEC PS_PP361_02 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "'"
''            sQry = "EXEC PS_PP361_02 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "','" & SYYYYMM & "','" & EYYYYMM & "'"
'            oRecordSet.DoQuery sQry
'
'            '// Procedure 실행(Grid 사용)
'            oForm01.DataSources.DataTables.Item("ZTEMP02").ExecuteQuery (sQry)
'            oGrid02.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP02")
    
'            SumPkQty = Trim(oGrid02.DataTable.GetValue(7, i))
'            SumPkWt = Trim(oGrid02.DataTable.GetValue(8, i))

            MngNo = Trim(oGrid01.DataTable.GetValue(5, loopCount1))
            
        End If
     Next

'    BPLID = oForm01.Items("BPLID").Specific.VALUE
'    FrDt = oForm01.Items("DocDatefr").Specific.VALUE
'    ToDt = oForm01.Items("DocDateto").Specific.VALUE

    Set oTempClass = New PS_GA160
    Call oTempClass.LoadForm(MngNo)

'    Call Sbo_Application.MessageBox(MngNo)
    
    Exit Sub

PS_PP915_GetDetail_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP915_GetDetail_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub












