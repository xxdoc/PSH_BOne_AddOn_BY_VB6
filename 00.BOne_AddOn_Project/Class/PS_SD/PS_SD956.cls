VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD956"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_SD956.cls
'//  Module             : 판매관리 > 애드온 > 판매리포트
'//  Desc               : 전산장비조회
'//  FormType           : PS_SD956
'//  Create Date(Start) : 2018.03.30
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid
Public oGrid02 As SAPbouiCOM.Grid
Public oGrid03 As SAPbouiCOM.Grid
Public oGrid04 As SAPbouiCOM.Grid

Public oDS_PS_SD956A As SAPbouiCOM.DataTable
Public oDS_PS_SD956B As SAPbouiCOM.DataTable
Public oDS_PS_SD956C As SAPbouiCOM.DataTable
Public oDS_PS_SD956D As SAPbouiCOM.DataTable

Public oBaseForm01      As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01   As String
Public oBaseColUID01    As String
Public oBaseColRow01    As Long
Public oBaseTradeType01 As String
Public oBaseItmBsort01  As String

Private oLastItemUID01  As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01   As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01   As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String, Optional ByVal oItmBsort02 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD956.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_SD956_" & GetTotalFormsCount
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
    
    Call PS_SD956_CreateItems
    Call PS_SD956_ComboBox_Setting
'    Call PS_SD956_CF_ChooseFromList
'    Call PS_SD956_FormItemEnabled
'    Call PS_SD956_EnableMenus
    '//Call PS_SD956_FormClear '//UDO방식일때
    '//Call PS_SD956_AddMatrixRow(0, True) '//UDO방식일때
    
'    oForm01.Items("FrYM").Specific.VALUE = Format(Now, "YYYYMM")
'    oForm01.Items("ToYM").Specific.VALUE = Format(Now, "YYYYMM")
    
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
                '//Call PS_SD956_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_SD956_FormItemEnabled '//UDO방식
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
    
        If pval.ItemUID = "BtnSrch" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_SD956_MTX01 '영업정보
                '생산정보
                '구매정보
                '회계정보
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        
'        ElseIf pval.ItemUID = "BtnPrt91" Then
'
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_SD956_Print_Report91
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
            
        End If
        
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
        
        'Folder03가 선택되었을 때
        If pval.ItemUID = "Folder03" Then
            
            oForm01.PaneLevel = 3
            
        End If
        
        'Folder04가 선택되었을 때
        If pval.ItemUID = "Folder04" Then
            
            oForm01.PaneLevel = 4
            
        End If
        '폴더를 사용할 때는 필수 소스_E
    
        If pval.ItemUID = "PS_SD956" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SD956_FormItemEnabled
'                    Call PS_SD956_FormClear '//UDO방식일때
'                    Call PS_SD956_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SD956_FormItemEnabled
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "")
        
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
        Call PS_SD956_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
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
                
                    Call PS_SD956_GetDetail
                
'                    Call PS_SD956_SetBaseForm '//부모폼에입력
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
        Call PS_SD956_FlushToItemValue(pval.ItemUID)
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
        Call PS_SD956_FormItemEnabled
        '//Call PS_SD956_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
        Call PS_SD956_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    
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
        Set oGrid03 = Nothing
        Set oGrid04 = Nothing
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
    '        Call oDS_PS_SD956L.RemoveRecord(oDS_PS_SD956L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_SD956_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_SD956_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_SD956_CreateItems() As Boolean
On Error GoTo PS_SD956_CreateItems_Error

    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    'Dim C_Date   As Date
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    Set oGrid02 = oForm01.Items("Grid02").Specific
    Set oGrid03 = oForm01.Items("Grid03").Specific
    Set oGrid04 = oForm01.Items("Grid04").Specific
    
    Call oForm01.DataSources.DataTables.Add("PS_SD956A")
    Call oForm01.DataSources.DataTables.Add("PS_SD956B")
    Call oForm01.DataSources.DataTables.Add("PS_SD956C")
    Call oForm01.DataSources.DataTables.Add("PS_SD956D")
    
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("PS_SD956A")
    oGrid02.DataTable = oForm01.DataSources.DataTables.Item("PS_SD956B")
    oGrid03.DataTable = oForm01.DataSources.DataTables.Item("PS_SD956C")
    oGrid04.DataTable = oForm01.DataSources.DataTables.Item("PS_SD956D")
    
    Set oDS_PS_SD956A = oForm01.DataSources.DataTables.Item("PS_SD956A")
    Set oDS_PS_SD956B = oForm01.DataSources.DataTables.Item("PS_SD956B")
    Set oDS_PS_SD956C = oForm01.DataSources.DataTables.Item("PS_SD956C")
    Set oDS_PS_SD956D = oForm01.DataSources.DataTables.Item("PS_SD956D")
    
    'oGrid01.SelectionMode = ms_NotSupported
    
    '수주일자(Fr)
    Call oForm01.DataSources.UserDataSources.Add("OrdFrDt", dt_DATE)
    Call oForm01.Items("OrdFrDt").Specific.DataBind.SetBound(True, "", "OrdFrDt")
    
    '수주일자(To)
    Call oForm01.DataSources.UserDataSources.Add("OrdToDt", dt_DATE)
    Call oForm01.Items("OrdToDt").Specific.DataBind.SetBound(True, "", "OrdToDt")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType", dt_SHORT_TEXT, 5)
    Call oForm01.Items("CardType").Specific.DataBind.SetBound(True, "", "CardType")
    
    '팀(거래처)
    Call oForm01.DataSources.UserDataSources.Add("CardTeam", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CardTeam").Specific.DataBind.SetBound(True, "", "CardTeam")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 50)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")

    '품명
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")

    '규격
    Call oForm01.DataSources.UserDataSources.Add("ItemSpec", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemSpec").Specific.DataBind.SetBound(True, "", "ItemSpec")
    
'    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_SD956_CreateItems_Error:
'    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD956_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SD956_ComboBox_Setting()
On Error GoTo PS_SD956_ComboBox_Setting_Error

    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim sQry         As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID        As String
    BPLID = MDC_PS_Common.User_BPLId
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    
    '거래처구분_S
    sQry = "        SELECT   U_Minor,"
    sQry = sQry & "          U_CdName"
    sQry = sQry & " FROM     [@PS_SY001L]"
    sQry = sQry & " WHERE    Code = 'C100'"
    sQry = sQry & " ORDER BY Code"
    Call oForm01.Items("CardType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType").Specific, sQry, "", False, False)
    Call oForm01.Items("CardType").Specific.Select(0, psk_Index)
    '거래처구분_E
    
    '거래처팀_S
    sQry = "        SELECT      U_Minor AS [Code],"
    sQry = sQry & "             U_CdName As [Name]"
    sQry = sQry & " FROM        [@PS_SY001L]"
    sQry = sQry & " WHERE       Code = 'I001'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Minor"
    Call oForm01.Items("CardTeam").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardTeam").Specific, sQry, "", False, False)
    Call oForm01.Items("CardTeam").Specific.Select(0, psk_Index)
    '거래처팀_E
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD956_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD956_CF_ChooseFromList()
On Error GoTo PS_SD956_CF_ChooseFromList_Error
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
PS_SD956_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD956_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD956_FormItemEnabled()
On Error GoTo PS_SD956_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseItmBsort01 = "" Then
        
        Else
        
        End If
        '//각모드에따른 아이템설정
        '//Call PS_SD956_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD956_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD956_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD956_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_SD956L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_SD956L.Offset = oRow
'    oDS_PS_SD956L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD956_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD956_DataValidCheck() As Boolean
On Error GoTo PS_SD956_DataValidCheck_Error
    Dim i As Long
    PS_SD956_DataValidCheck = True
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_SD956_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_SD956_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_SD956_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
'    Call PS_SD956_FormClear
    Exit Function
PS_SD956_DataValidCheck_Error:
    PS_SD956_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_SD956_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD956_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_SD956_FlushToItemValue_Error

    Dim i            As Integer
    Dim ErrNum       As Integer
    
    Select Case oUID
        
        Case "ItemCode"
        
            oForm01.Items("ItemName").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'") '품명
            oForm01.Items("ItemSpec").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'") '규격
       
    End Select

    Exit Sub
PS_SD956_FlushToItemValue_Error:

    Call MDC_Com.MDC_GF_Message("PS_SD956_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

Private Sub PS_SD956_MTX01()
On Error GoTo PS_SD956_MTX01_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    
    Dim i       As Long
    Dim Query   As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim OrdFrDt   As String '수주기간(Fr)
    Dim OrdToDt   As String '수주기간(To)
    Dim CardType  As String '거래처구분
    Dim CardTeam  As String '팀(거래처)
    Dim ItemCode  As String '작번
    
    OrdFrDt = Trim(oForm01.Items("OrdFrDt").Specific.VALUE) '수주기간(시작)
    OrdToDt = Trim(oForm01.Items("OrdToDt").Specific.VALUE) '수주기간(종료)
    CardType = Trim(oForm01.Items("CardType").Specific.VALUE) '거래처구분
    CardTeam = Trim(oForm01.Items("CardTeam").Specific.VALUE) '팀(거래처)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE) '작번
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Query = "       EXEC PS_SD956_01 '"
    Query = Query & OrdFrDt & "','"
    Query = Query & OrdToDt & "','"
    Query = Query & CardType & "','"
    Query = Query & CardTeam & "','"
    Query = Query & ItemCode & "'"
    
    Call oGrid01.DataTable.Clear
    Call oDS_PS_SD956A.ExecuteQuery(Query)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
    oGrid01.Columns(5).RightJustified = True
    oGrid01.Columns(6).RightJustified = True
    oGrid01.Columns(12).RightJustified = True
    oGrid01.Columns(14).RightJustified = True
    oGrid01.Columns(16).RightJustified = True
    oGrid01.Columns(17).RightJustified = True
    oGrid01.Columns(18).RightJustified = True
'    oGrid01.Columns(21).RightJustified = True
'    oGrid01.Columns(22).RightJustified = True
'    oGrid01.Columns(23).RightJustified = True
'    oGrid01.Columns(24).RightJustified = True
'    oGrid01.Columns(25).RightJustified = True
'    oGrid01.Columns(26).RightJustified = True
'    oGrid01.Columns(27).RightJustified = True
'    oGrid01.Columns(29).RightJustified = True
'    oGrid01.Columns(30).RightJustified = True
'    oGrid01.Columns(31).RightJustified = True
'    oGrid01.Columns(32).RightJustified = True
'    oGrid01.Columns(33).RightJustified = True
'
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
        GoTo PS_SD956_MTX01_Exit
    End If

    oGrid01.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_MTX01_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_SD956_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD956_MTX02()
On Error GoTo PS_SD956_MTX02_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    
    Dim i       As Long
    Dim Query   As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim OrdFrDt   As String '수주기간(Fr)
    Dim OrdToDt   As String '수주기간(To)
    Dim CardType  As String '거래처구분
    Dim CardTeam  As String '팀(거래처)
    Dim ItemCode  As String '작번
    
    OrdFrDt = Trim(oForm01.Items("OrdFrDt").Specific.VALUE) '수주기간(시작)
    OrdToDt = Trim(oForm01.Items("OrdToDt").Specific.VALUE) '수주기간(종료)
    CardType = Trim(oForm01.Items("CardType").Specific.VALUE) '거래처구분
    CardTeam = Trim(oForm01.Items("CardTeam").Specific.VALUE) '팀(거래처)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE) '작번
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Query = "       EXEC PS_SD956_02 '"
    Query = Query & OrdFrDt & "','"
    Query = Query & OrdToDt & "','"
    Query = Query & CardType & "','"
    Query = Query & CardTeam & "','"
    Query = Query & ItemCode & "'"
    
    Call oGrid02.DataTable.Clear
    Call oDS_PS_SD956B.ExecuteQuery(Query)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
'    oGrid01.Columns(14).RightJustified = True
'    oGrid01.Columns(15).RightJustified = True
'    oGrid01.Columns(16).RightJustified = True
'    oGrid01.Columns(17).RightJustified = True
'    oGrid01.Columns(18).RightJustified = True
'    oGrid01.Columns(19).RightJustified = True
'    oGrid01.Columns(20).RightJustified = True
'    oGrid01.Columns(21).RightJustified = True
'    oGrid01.Columns(22).RightJustified = True
'    oGrid01.Columns(23).RightJustified = True
'    oGrid01.Columns(24).RightJustified = True
'    oGrid01.Columns(25).RightJustified = True
'    oGrid01.Columns(26).RightJustified = True
'    oGrid01.Columns(27).RightJustified = True
'    oGrid01.Columns(29).RightJustified = True
'    oGrid01.Columns(30).RightJustified = True
'    oGrid01.Columns(31).RightJustified = True
'    oGrid01.Columns(32).RightJustified = True
'    oGrid01.Columns(33).RightJustified = True
'
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid02.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_SD956_MTX02_Exit
    End If

    oGrid02.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_MTX02_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_SD956_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD956_MTX03()
On Error GoTo PS_SD956_MTX03_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    
    Dim i       As Long
    Dim Query   As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim OrdFrDt   As String '수주기간(Fr)
    Dim OrdToDt   As String '수주기간(To)
    Dim CardType  As String '거래처구분
    Dim CardTeam  As String '팀(거래처)
    Dim ItemCode  As String '작번
    
    OrdFrDt = Trim(oForm01.Items("OrdFrDt").Specific.VALUE) '수주기간(시작)
    OrdToDt = Trim(oForm01.Items("OrdToDt").Specific.VALUE) '수주기간(종료)
    CardType = Trim(oForm01.Items("CardType").Specific.VALUE) '거래처구분
    CardTeam = Trim(oForm01.Items("CardTeam").Specific.VALUE) '팀(거래처)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE) '작번
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Query = "       EXEC PS_SD956_03 '"
    Query = Query & OrdFrDt & "','"
    Query = Query & OrdToDt & "','"
    Query = Query & CardType & "','"
    Query = Query & CardTeam & "','"
    Query = Query & ItemCode & "'"
    
    Call oGrid03.DataTable.Clear
    Call oDS_PS_SD956C.ExecuteQuery(Query)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
'    oGrid01.Columns(14).RightJustified = True
'    oGrid01.Columns(15).RightJustified = True
'    oGrid01.Columns(16).RightJustified = True
'    oGrid01.Columns(17).RightJustified = True
'    oGrid01.Columns(18).RightJustified = True
'    oGrid01.Columns(19).RightJustified = True
'    oGrid01.Columns(20).RightJustified = True
'    oGrid01.Columns(21).RightJustified = True
'    oGrid01.Columns(22).RightJustified = True
'    oGrid01.Columns(23).RightJustified = True
'    oGrid01.Columns(24).RightJustified = True
'    oGrid01.Columns(25).RightJustified = True
'    oGrid01.Columns(26).RightJustified = True
'    oGrid01.Columns(27).RightJustified = True
'    oGrid01.Columns(29).RightJustified = True
'    oGrid01.Columns(30).RightJustified = True
'    oGrid01.Columns(31).RightJustified = True
'    oGrid01.Columns(32).RightJustified = True
'    oGrid01.Columns(33).RightJustified = True
'
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid03.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_SD956_MTX03_Exit
    End If

    oGrid03.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_MTX03_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_MTX03_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_SD956_MTX03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD956_MTX04()
On Error GoTo PS_SD956_MTX04_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    
    Dim i       As Long
    Dim Query   As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim OrdFrDt   As String '수주기간(Fr)
    Dim OrdToDt   As String '수주기간(To)
    Dim CardType  As String '거래처구분
    Dim CardTeam  As String '팀(거래처)
    Dim ItemCode  As String '작번
    
    OrdFrDt = Trim(oForm01.Items("OrdFrDt").Specific.VALUE) '수주기간(시작)
    OrdToDt = Trim(oForm01.Items("OrdToDt").Specific.VALUE) '수주기간(종료)
    CardType = Trim(oForm01.Items("CardType").Specific.VALUE) '거래처구분
    CardTeam = Trim(oForm01.Items("CardTeam").Specific.VALUE) '팀(거래처)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE) '작번
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Query = "       EXEC PS_SD956_04 '"
    Query = Query & OrdFrDt & "','"
    Query = Query & OrdToDt & "','"
    Query = Query & CardType & "','"
    Query = Query & CardTeam & "','"
    Query = Query & ItemCode & "'"
    
    Call oGrid04.DataTable.Clear
    Call oDS_PS_SD956D.ExecuteQuery(Query)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
'    oGrid01.Columns(14).RightJustified = True
'    oGrid01.Columns(15).RightJustified = True
'    oGrid01.Columns(16).RightJustified = True
'    oGrid01.Columns(17).RightJustified = True
'    oGrid01.Columns(18).RightJustified = True
'    oGrid01.Columns(19).RightJustified = True
'    oGrid01.Columns(20).RightJustified = True
'    oGrid01.Columns(21).RightJustified = True
'    oGrid01.Columns(22).RightJustified = True
'    oGrid01.Columns(23).RightJustified = True
'    oGrid01.Columns(24).RightJustified = True
'    oGrid01.Columns(25).RightJustified = True
'    oGrid01.Columns(26).RightJustified = True
'    oGrid01.Columns(27).RightJustified = True
'    oGrid01.Columns(29).RightJustified = True
'    oGrid01.Columns(30).RightJustified = True
'    oGrid01.Columns(31).RightJustified = True
'    oGrid01.Columns(32).RightJustified = True
'    oGrid01.Columns(33).RightJustified = True
'
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid04.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_SD956_MTX04_Exit
    End If

    oGrid04.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_MTX04_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD956_MTX04_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_SD956_MTX04_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SD956_DI_API() As Boolean
'On Error GoTo PS_SD956_DI_API_Error
'    PS_SD956_DI_API = True
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
'    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
'    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
'    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.Value, "&&&&-&&-&&")
'    For i = 0 To UBound(ItemInformation)
'        If ItemInformation(i).Check = True Then
'            GoTo Continue_First
'        End If
'        If i <> 0 Then
'            oDIObject.Lines.Add
'        End If
'        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
'        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
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
'            Call oDS_PS_SD956L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_SD956L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
'    Else
'        GoTo PS_SD956_DI_API_Error
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
'PS_SD956_DI_API_DI_Error:
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
'    PS_SD956_DI_API = False
'    Set oDIObject = Nothing
'    Exit Function
'PS_SD956_DI_API_Error:
'    Sbo_Application.SetStatusBarMessage "PS_SD956_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    PS_SD956_DI_API = False
End Function

Private Sub PS_SD956_SetBaseForm()
On Error GoTo PS_SD956_SetBaseForm_Error
    Dim i As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    If oBaseForm01 Is Nothing Then
    '//DoNothing
    Else
    
    End If
    Exit Sub
PS_SD956_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD956_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD956_FormResize()
On Error GoTo PS_SD956_FormResize_Error
    
    '그룹박스 크기 동적 할당
    oForm01.Items("GrpBox01").Height = oForm01.Items("Grid01").Height + 40
    oForm01.Items("GrpBox01").Width = oForm01.Items("Grid01").Width + 30
    
    If oGrid01.Columns.Count > 0 Then
        Call oGrid01.AutoResizeColumns
    End If
    
    If oGrid02.Columns.Count > 0 Then
        oGrid02.AutoResizeColumns
    End If
    
    If oGrid03.Columns.Count > 0 Then
        oGrid03.AutoResizeColumns
    End If
    
    If oGrid04.Columns.Count > 0 Then
        oGrid04.AutoResizeColumns
    End If
    
    Exit Sub
PS_SD956_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD956_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD956_Print_Report01()
On Error GoTo Print_Query_Error

    Dim DocNum     As String
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String

    Dim i As Integer
    Dim ErrNum     As Integer
    Dim Sub_sQry   As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim BPLID    As String '사업장
    Dim TeamCode As String '팀(부서)
    Dim CntcCode As String '사용자
    Dim Ctgr     As String '전산장비분류
    Dim ModelNm  As String '모델명
    Dim LocCls   As String '위치구분
    Dim MngNo    As String '관리번호
    Dim RegFrDt  As String '등록기간(시작)
    Dim RegToDt  As String '등록기간(종료)
    Dim DisUseYN As String '폐기장비 포함 여부
    
    BPLID = Trim(oForm01.Items("BPLID01").Specific.Selected.VALUE) '사업장
    TeamCode = Trim(oForm01.Items("TeamCode01").Specific.VALUE) '팀(부서)
    CntcCode = Trim(oForm01.Items("CntcCode01").Specific.VALUE) '사용자
    Ctgr = Trim(oForm01.Items("Ctgr01").Specific.Selected.VALUE) '전산장비분류
    ModelNm = Trim(oForm01.Items("ModelNm01").Specific.VALUE) '모델명
    LocCls = Trim(oForm01.Items("LocCls01").Specific.Selected.VALUE) '위치구분
    MngNo = Trim(oForm01.Items("MngNo01").Specific.VALUE) '관리번호
    RegFrDt = Trim(oForm01.Items("RegFrDt01").Specific.VALUE) '등록기간(시작)
    RegToDt = Trim(oForm01.Items("RegToDt01").Specific.VALUE) '등록기간(종료)
    
    If oForm01.DataSources.UserDataSources.Item("DisUseYN01").VALUE = "Y" Then '폐기장비 포함 여부
        DisUseYN = "Y"
    Else
        DisUseYN = "N"
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_SD956] 레포트"
    
    '쿼리
    If Ctgr = "M" Then '모니터로 조회할 경우는 모니터만 단독으로 출력
    
        ReportName = "PS_SD956_06.rpt"
    
        sQry = "      EXEC PS_SD956_56 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & TeamCode & "','"
        sQry = sQry & CntcCode & "','"
        sQry = sQry & Ctgr & "','"
        sQry = sQry & ModelNm & "','"
        sQry = sQry & LocCls & "','"
        sQry = sQry & MngNo & "','"
        sQry = sQry & RegFrDt & "','"
        sQry = sQry & RegToDt & "','"
        sQry = sQry & DisUseYN & "'"
    
    Else
    
        ReportName = "PS_SD956_01.rpt"
    
        sQry = "      EXEC PS_SD956_51 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & TeamCode & "','"
        sQry = sQry & CntcCode & "','"
        sQry = sQry & Ctgr & "','"
        sQry = sQry & ModelNm & "','"
        sQry = sQry & LocCls & "','"
        sQry = sQry & MngNo & "','"
        sQry = sQry & RegFrDt & "','"
        sQry = sQry & RegToDt & "','"
        sQry = sQry & DisUseYN & "'"
    
    End If
 
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
        MDC_Com.MDC_GF_Message "PS_SD956_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub PS_SD956_Print_Report05()
On Error GoTo Print_Query_Error

    Dim DocNum     As String
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String

    Dim i          As Integer
    Dim ErrNum     As Integer
    Dim Sub_sQry   As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim BPLID    As String '사업장
    Dim CntcCode As String '사용자
    Dim Ctgr     As String '전산장비분류
    Dim ModelNm  As String '모델명
    Dim LocCls   As String '위치구분
    Dim MngNo    As String '관리번호
    Dim ChkFrDt  As String '점검기간(시작)
    Dim ChkToDt  As String '점검기간(종료)
    Dim ChkCls   As String '점검분류
    
    BPLID = Trim(oForm01.Items("BPLID05").Specific.Selected.VALUE) '사업장
    CntcCode = Trim(oForm01.Items("CntcCode05").Specific.VALUE) '사용자
    Ctgr = Trim(oForm01.Items("Ctgr05").Specific.Selected.VALUE) '전산장비분류
    ModelNm = Trim(oForm01.Items("ModelNm05").Specific.VALUE) '모델명
    LocCls = Trim(oForm01.Items("LocCls05").Specific.Selected.VALUE) '위치구분
    MngNo = Trim(oForm01.Items("MngNo05").Specific.VALUE) '관리번호
    ChkFrDt = Trim(oForm01.Items("ChkFrDt05").Specific.VALUE) '점검기간(시작)
    ChkToDt = Trim(oForm01.Items("ChkToDt05").Specific.VALUE) '점검기간(종료)
    ChkCls = Trim(oForm01.Items("ChkCls05").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_SD956] 레포트"
    
    ReportName = "PS_SD956_05.rpt"
    
    '쿼리
    sQry = "      EXEC PS_SD956_55 '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & CntcCode & "','"
    sQry = sQry & Ctgr & "','"
    sQry = sQry & ModelNm & "','"
    sQry = sQry & LocCls & "','"
    sQry = sQry & MngNo & "','"
    sQry = sQry & ChkFrDt & "','"
    sQry = sQry & ChkToDt & "','"
    sQry = sQry & ChkCls & "'"
 
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
        MDC_Com.MDC_GF_Message "PS_SD956_Print_Report05_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub PS_SD956_GetDetail()
On Error GoTo PS_SD956_GetDetail_Error

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

'    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
'    FrDt = oForm01.Items("DocDatefr").Specific.VALUE
'    ToDt = oForm01.Items("DocDateto").Specific.VALUE

    Set oTempClass = New PS_GA160
    Call oTempClass.LoadForm(MngNo)

'    Call Sbo_Application.MessageBox(MngNo)
    
    Exit Sub

PS_SD956_GetDetail_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD956_GetDetail_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub











