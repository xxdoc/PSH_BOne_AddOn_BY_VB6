VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM131"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_QM131.cls
'//  Module : 생산 관리 > 애드온 > 생산리포트
'//  Desc : 인원별부서별비가동내역조회
'//  FormType : PS_QM131
'//  Create Date(Start) : 2014.03.21
'//  Create Date(End) :
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************

'//사용방법
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_QM131
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form

Public oGrid11 As SAPbouiCOM.Grid
Public oGrid12 As SAPbouiCOM.Grid

Public oDS_PS_QM131A As SAPbouiCOM.DataTable
Public oDS_PS_QM131B As SAPbouiCOM.DataTable

Public oGrid21 As SAPbouiCOM.Grid
Public oGrid22 As SAPbouiCOM.Grid

Public oDS_PS_QM131C As SAPbouiCOM.DataTable
Public oDS_PS_QM131D As SAPbouiCOM.DataTable

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM131.srf")
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
    
    oFormUniqueID01 = "PS_QM131_" & GetTotalFormsCount
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
    
    Call PS_QM131_CreateItems
    Call PS_QM131_ComboBox_Setting
'    Call PS_QM131_CF_ChooseFromList
'    Call PS_QM131_FormItemEnabled
'    Call PS_QM131_EnableMenus
    '//Call PS_QM131_FormClear '//UDO방식일때
    '//Call PS_QM131_AddMatrixRow(0, True) '//UDO방식일때
    
    oForm01.Items("FrDt11").Specific.VALUE = Format(Now, "YYYYMM01")
    oForm01.Items("ToDt11").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oForm01.Items("FrDt21").Specific.VALUE = Format(Now - 60, "YYYYMM01")
    oForm01.Items("ToDt21").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oForm01.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    
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
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_QM131_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_QM131_FormItemEnabled '//UDO방식
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
    
        If pval.ItemUID = "BtnSrch11" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_QM131_MTX11
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        ElseIf pval.ItemUID = "BtnSrch21" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_QM131_MTX21
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
                        
        ElseIf pval.ItemUID = "BtnPrt21" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_QM131_Print_Report01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        End If
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_QM131_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM131" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        '폴더를 사용할 때는 필수 소스_S
        'Folder01이 선택되었을 때
        If pval.ItemUID = "Folder01" Then
            
            oForm01.PaneLevel = 1
            oForm01.DefButton = "BtnSrch11"
            
        End If
        
        'Folder02가 선택되었을 때
        If pval.ItemUID = "Folder02" Then
            
            oForm01.PaneLevel = 2
            oForm01.DefButton = "BtnSrch21"
            
        End If
        '폴더를 사용할 때는 필수 소스_E
        
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_QM131_FormItemEnabled
'                    Call PS_QM131_FormClear '//UDO방식일때
'                    Call PS_QM131_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_QM131_FormItemEnabled
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode21", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemName", "") '//사용자값활성
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
        Call PS_QM131_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
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
    
        If pval.ItemUID = "Grid11" Then
        
            If pval.Row = -1 Then
                oGrid11.Columns(pval.ColUID).TitleObject.Sortable = True
            Else
                If oGrid11.Rows.SelectedRows.Count > 0 Then
                
                    Call PS_QM131_MTX12
                
                End If
            End If
            
        ElseIf pval.ItemUID = "Grid21" Then
        
            If pval.Row = -1 Then
                oGrid21.Columns(pval.ColUID).TitleObject.Sortable = True
            Else
                If oGrid21.Rows.SelectedRows.Count > 0 Then
                
                    Call PS_QM131_MTX22
                
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
        Call PS_QM131_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
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
        Call PS_QM131_FormItemEnabled
        '//Call PS_QM131_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
    
        oForm01.Items("Grid11").Height = oForm01.Height - 170
        oForm01.Items("Grid11").Width = oForm01.Width / 2 + 430
        oForm01.Items("Grid12").Left = oForm01.Items("Grid11").Width + 40
        oForm01.Items("Grid12").Height = oForm01.Items("Grid11").Height
        oForm01.Items("Grid12").Width = oForm01.Width - oForm01.Items("Grid11").Width - 75
        
        oForm01.Items("Grid21").Height = oForm01.Items("Grid11").Height - 45
        oForm01.Items("Grid21").Width = oForm01.Items("Grid11").Width
        oForm01.Items("Grid22").Left = oForm01.Items("Grid12").Left
        oForm01.Items("Grid22").Height = oForm01.Items("Grid12").Height - 45
        oForm01.Items("Grid22").Width = oForm01.Items("Grid12").Width
    
        '그룹박스 크기 동적 할당
        oForm01.Items("GrpBox01").Height = oForm01.Items("Grid11").Height + 70
        oForm01.Items("GrpBox01").Width = oForm01.Items("Grid11").Width + oForm01.Items("Grid12").Width + 45
        
        If oGrid11.Columns.Count > 0 Then
            Call oGrid11.AutoResizeColumns
        End If
        
        If oGrid12.Columns.Count > 0 Then
            Call oGrid12.AutoResizeColumns
        End If
        
        If oGrid21.Columns.Count > 0 Then
            Call oGrid21.AutoResizeColumns
        End If
        
        If oGrid22.Columns.Count > 0 Then
            Call oGrid22.AutoResizeColumns
        End If
    
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
        Set oGrid11 = Nothing
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
    '        Call oDS_PS_QM131L.RemoveRecord(oDS_PS_QM131L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_QM131_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_QM131_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_QM131_CreateItems() As Boolean
On Error GoTo PS_QM131_CreateItems_Error
    
    Set oGrid11 = oForm01.Items("Grid11").Specific
    Set oGrid12 = oForm01.Items("Grid12").Specific
    Set oGrid21 = oForm01.Items("Grid21").Specific
    Set oGrid22 = oForm01.Items("Grid22").Specific
    
    oGrid11.SelectionMode = ms_Single
    oGrid12.SelectionMode = ms_Single
    oGrid21.SelectionMode = ms_Single
    oGrid22.SelectionMode = ms_Single
    
    Call oForm01.DataSources.DataTables.Add("PS_QM131A")
    Call oForm01.DataSources.DataTables.Add("PS_QM131B")
    Call oForm01.DataSources.DataTables.Add("PS_QM131C")
    Call oForm01.DataSources.DataTables.Add("PS_QM131D")
    
    oGrid11.DataTable = oForm01.DataSources.DataTables.Item("PS_QM131A")
    oGrid12.DataTable = oForm01.DataSources.DataTables.Item("PS_QM131B")
    oGrid21.DataTable = oForm01.DataSources.DataTables.Item("PS_QM131C")
    oGrid22.DataTable = oForm01.DataSources.DataTables.Item("PS_QM131D")
        
    Set oDS_PS_QM131A = oForm01.DataSources.DataTables.Item("PS_QM131A")
    Set oDS_PS_QM131B = oForm01.DataSources.DataTables.Item("PS_QM131B")
    Set oDS_PS_QM131C = oForm01.DataSources.DataTables.Item("PS_QM131C")
    Set oDS_PS_QM131D = oForm01.DataSources.DataTables.Item("PS_QM131D")
    
    '//////////출고일//////////
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID11", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID11").Specific.DataBind.SetBound(True, "", "BPLID11")
    
    '반출기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt11", dt_DATE)
    Call oForm01.Items("FrDt11").Specific.DataBind.SetBound(True, "", "FrDt11")
    
    '반출기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt11", dt_DATE)
    Call oForm01.Items("ToDt11").Specific.DataBind.SetBound(True, "", "ToDt11")
    '//////////출고일//////////
    
    '//////////생산완료일//////////
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID21", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID21").Specific.DataBind.SetBound(True, "", "BPLID21")
    
    '가입고기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt21", dt_DATE)
    Call oForm01.Items("FrDt21").Specific.DataBind.SetBound(True, "", "FrDt21")
    
    '가입고기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt21", dt_DATE)
    Call oForm01.Items("ToDt21").Specific.DataBind.SetBound(True, "", "ToDt21")
    
    '자체/외주
    Call oForm01.DataSources.UserDataSources.Add("InOut21", dt_SHORT_TEXT, 10)
    Call oForm01.Items("InOut21").Specific.DataBind.SetBound(True, "", "InOut21")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType21", dt_SHORT_TEXT, 10)
    Call oForm01.Items("CardType21").Specific.DataBind.SetBound(True, "", "CardType21")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType21", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItemType21").Specific.DataBind.SetBound(True, "", "ItemType21")
    
    '품목코드
    Call oForm01.DataSources.UserDataSources.Add("ItemCode21", dt_SHORT_TEXT, 50)
    Call oForm01.Items("ItemCode21").Specific.DataBind.SetBound(True, "", "ItemCode21")
    
    '품목명
    Call oForm01.DataSources.UserDataSources.Add("ItemName21", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName21").Specific.DataBind.SetBound(True, "", "ItemName21")
    '//////////생산완료일//////////
    
    Exit Function
PS_QM131_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM131_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_QM131_ComboBox_Setting()
On Error GoTo PS_QM131_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    
    '사업장(출고일)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID11").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", MDC_PS_Common.User_BPLId, False, False)
'    Call oForm01.Items("BPLID11").Specific.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '사업장(생산완료일)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID21").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", MDC_PS_Common.User_BPLId, False, False)
    
    '자체/외주(생산완료일)
    Call oForm01.Items("InOut21").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("InOut21").Specific.ValidValues.Add("IN", "자체")
    Call oForm01.Items("InOut21").Specific.ValidValues.Add("OUT", "외주")
    Call oForm01.Items("InOut21").Specific.Select(0, psk_Index)
    
    '거래처구분(생산완료일)
    sQry = "            SELECT  U_Minor,"
    sQry = sQry & "             U_CdName"
    sQry = sQry & "  FROM   [@PS_SY001L]"
    sQry = sQry & "  WHERE  Code = 'C100'"
    
    Call oForm01.Items("CardType21").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType21").Specific, sQry, "%", False, False)
'    Call oForm01.Items("HarfCls").Specific.Select(0, psk_Index)
    
    '품목구분(생산완료일)
    sQry = "            SELECT  U_Minor,"
    sQry = sQry & "             U_CdName"
    sQry = sQry & "  FROM   [@PS_SY001L]"
    sQry = sQry & "  WHERE  Code = 'S002'"
    
    Call oForm01.Items("ItemType21").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType21").Specific, sQry, "%", False, False)
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM131_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM131_CF_ChooseFromList()
On Error GoTo PS_QM131_CF_ChooseFromList_Error
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
PS_QM131_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM131_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM131_FormItemEnabled()
On Error GoTo PS_QM131_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseItmBsort01 = "" Then
        
        Else
        
        End If
        '//각모드에따른 아이템설정
        '//Call PS_QM131_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM131_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM131_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM131_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_QM131L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_QM131L.Offset = oRow
'    oDS_PS_QM131L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM131_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_QM131_DataValidCheck() As Boolean
On Error GoTo PS_QM131_DataValidCheck_Error
    Dim i As Long
    PS_QM131_DataValidCheck = True
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_QM131_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_QM131_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_QM131_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
'    Call PS_QM131_FormClear
    Exit Function
PS_QM131_DataValidCheck_Error:
    PS_QM131_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_QM131_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_QM131_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_QM131_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim TeamCode As String
    Dim RspCode As String
    
    Select Case oUID
        
        Case "ItemCode21"
        
            oForm01.Items("ItemName21").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items("ItemCode21").Specific.VALUE & "'")
        
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_QM131_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_QM131_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub


Private Sub PS_QM131_MTX11()
On Error GoTo PS_QM131_MTX11_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID As String
    Dim FrDt As String
    Dim ToDt As String
    
    BPLID = oForm01.Items("BPLID11").Specific.VALUE
    FrDt = oForm01.Items("FrDt11").Specific.VALUE
    ToDt = oForm01.Items("ToDt11").Specific.VALUE
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Query01 = "              EXEC PS_QM131_11 '"
    Query01 = Query01 & BPLID & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "'"

    Call oGrid11.DataTable.Clear
    Call oDS_PS_QM131A.ExecuteQuery(Query01)

    oGrid11.Columns(3).RightJustified = True
'    oGrid11.Columns(8).RightJustified = True
'    oGrid11.Columns(9).RightJustified = True
'    oGrid11.Columns(10).RightJustified = True
'    oGrid11.Columns(11).RightJustified = True
'    oGrid11.Columns(12).RightJustified = True
'    oGrid11.Columns(13).RightJustified = True
'    oGrid11.Columns(14).RightJustified = True
'    oGrid11.Columns(15).RightJustified = True
'    oGrid11.Columns(16).RightJustified = True
'    oGrid11.Columns(17).RightJustified = True
'    oGrid11.Columns(18).RightJustified = True
'    oGrid11.Columns(19).RightJustified = True
'    oGrid11.Columns(20).RightJustified = True
'    oGrid11.Columns(21).RightJustified = True
'    oGrid11.Columns(22).RightJustified = True
'    oGrid11.Columns(23).RightJustified = True
'    oGrid11.Columns(24).RightJustified = True
'    oGrid11.Columns(25).RightJustified = True
'    oGrid11.Columns(26).RightJustified = True
'    oGrid11.Columns(27).RightJustified = True
'    oGrid11.Columns(29).RightJustified = True
'    oGrid11.Columns(30).RightJustified = True
'    oGrid11.Columns(31).RightJustified = True
'    oGrid11.Columns(32).RightJustified = True
'    oGrid11.Columns(33).RightJustified = True
'
'    oGrid11.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid11.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid11.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid11.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid11.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid11.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid11.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid11.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid11.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강

    Call oGrid12.DataTable.Clear '상세 그리드도 클리어

    If oGrid11.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_QM131_MTX11_Exit
    End If
    
    oGrid11.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_MTX11_Exit:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_MTX11_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM131_MTX11_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM131_MTX12()
On Error GoTo PS_QM131_MTX12_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)

'    Call Sbo_Application.MessageBox(ItemCode)
    
    Dim loopCount1 As Long
    Dim Query01 As String

    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    'Dim BPLID As String '사업장
    Dim ItemCode As String '자재코드
    Dim FrDt As String
    Dim ToDt As String
    'Dim DocDate As String '사용일
    'Dim UsePrd As Integer '사용기간
    'Dim OptBtnValue As String

    FrDt = Trim(oForm01.Items("FrDt11").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt11").Specific.VALUE)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    For loopCount1 = 0 To oGrid11.Rows.Count - 1
        If oGrid11.Rows.IsSelected(loopCount1) = True Then
            
            ItemCode = Trim(oGrid11.DataTable.GetValue(0, loopCount1))
            'DocDate = Trim(Format(oGrid01.DataTable.GetValue(5, loopCount1), "YYYYMMDD"))
            
        End If
    Next
    
    Query01 = "              EXEC PS_QM131_12 '"
    Query01 = Query01 & ItemCode & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "'"
    
    Call oGrid12.DataTable.Clear
    Call oDS_PS_QM131B.ExecuteQuery(Query01)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

'    oGrid02.Columns(1).RightJustified = True
'    oGrid02.Columns(2).RightJustified = True
'    oGrid01.Columns(8).RightJustified = True
'    oGrid01.Columns(12).RightJustified = True
'    oGrid01.Columns(13).RightJustified = True
'    oGrid01.Columns(14).RightJustified = True
'    oGrid01.Columns(15).RightJustified = True

'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑

'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강

    Call RecordSet01.DoQuery(Query01)

    If oGrid12.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_QM131_MTX12_Exit
    End If

    oGrid12.AutoResizeColumns
    oForm01.Update

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_MTX12_Exit:
    Set RecordSet01 = Nothing

    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_MTX12_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_QM131_MTX12_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM131_MTX21()
On Error GoTo PS_QM131_MTX21_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID As String
    Dim FrDt As String
    Dim ToDt As String
    Dim InOut As String
    Dim CardType As String
    Dim ItemType As String
    Dim ItemCode As String
    
    BPLID = oForm01.Items("BPLID21").Specific.VALUE
    FrDt = oForm01.Items("FrDt21").Specific.VALUE
    ToDt = oForm01.Items("ToDt21").Specific.VALUE
    InOut = oForm01.Items("InOut21").Specific.VALUE
    CardType = oForm01.Items("CardType21").Specific.VALUE
    ItemType = oForm01.Items("ItemType21").Specific.VALUE
    ItemCode = oForm01.Items("ItemCode21").Specific.VALUE
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Query01 = "              EXEC PS_QM131_21 '"
    Query01 = Query01 & BPLID & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "','"
    Query01 = Query01 & InOut & "','"
    Query01 = Query01 & CardType & "','"
    Query01 = Query01 & ItemType & "','"
    Query01 = Query01 & ItemCode & "'"

    Call oGrid21.DataTable.Clear
    Call oDS_PS_QM131C.ExecuteQuery(Query01)

    oGrid21.Columns(7).RightJustified = True
    oGrid21.Columns(10).RightJustified = True
'    oGrid21.Columns(9).RightJustified = True
'    oGrid21.Columns(10).RightJustified = True
'    oGrid21.Columns(11).RightJustified = True
'    oGrid21.Columns(12).RightJustified = True
'    oGrid21.Columns(13).RightJustified = True
'    oGrid21.Columns(14).RightJustified = True
'    oGrid21.Columns(15).RightJustified = True
'    oGrid21.Columns(16).RightJustified = True
'    oGrid21.Columns(17).RightJustified = True
'    oGrid21.Columns(18).RightJustified = True
'    oGrid21.Columns(19).RightJustified = True
'    oGrid21.Columns(20).RightJustified = True
'    oGrid21.Columns(21).RightJustified = True
'    oGrid21.Columns(22).RightJustified = True
'    oGrid21.Columns(23).RightJustified = True
'    oGrid21.Columns(24).RightJustified = True
'    oGrid21.Columns(25).RightJustified = True
'    oGrid21.Columns(26).RightJustified = True
'    oGrid21.Columns(27).RightJustified = True
'    oGrid21.Columns(29).RightJustified = True
'    oGrid21.Columns(30).RightJustified = True
'    oGrid21.Columns(31).RightJustified = True
'    oGrid21.Columns(32).RightJustified = True
'    oGrid21.Columns(33).RightJustified = True
'
'    oGrid21.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid21.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid21.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid21.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid21.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid21.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid21.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid21.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid21.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
    
    Call oGrid22.DataTable.Clear '상세 그리드도 클리어

    If oGrid21.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_QM131_MTX21_Exit
    End If
    
    oGrid21.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_MTX21_Exit:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_MTX21_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM131_MTX21_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM131_MTX22()
On Error GoTo PS_QM131_MTX22_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)

'    Call Sbo_Application.MessageBox(ItemCode)
    
    Dim loopCount1 As Long
    Dim Query01 As String

    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    'Dim BPLID As String '사업장
    Dim ItemCode As String '자재코드
    Dim FrDt As String
    Dim ToDt As String
    'Dim DocDate As String '사용일
    'Dim UsePrd As Integer '사용기간
    'Dim OptBtnValue As String

    FrDt = Trim(oForm01.Items("FrDt21").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt21").Specific.VALUE)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    For loopCount1 = 0 To oGrid21.Rows.Count - 1
        If oGrid21.Rows.IsSelected(loopCount1) = True Then
            
            ItemCode = Trim(oGrid21.DataTable.GetValue(0, loopCount1))
            'DocDate = Trim(Format(oGrid01.DataTable.GetValue(5, loopCount1), "YYYYMMDD"))
            
        End If
    Next
    
    Query01 = "              EXEC PS_QM131_22 '"
    Query01 = Query01 & ItemCode & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "'"
    
    Call oGrid22.DataTable.Clear
    Call oDS_PS_QM131D.ExecuteQuery(Query01)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    oGrid22.Columns(4).RightJustified = True
    oGrid22.Columns(5).RightJustified = True
'    oGrid01.Columns(8).RightJustified = True
'    oGrid01.Columns(12).RightJustified = True
'    oGrid01.Columns(13).RightJustified = True
'    oGrid01.Columns(14).RightJustified = True
'    oGrid01.Columns(15).RightJustified = True

'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑

'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강

    Call RecordSet01.DoQuery(Query01)

    If oGrid22.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_QM131_MTX22_Exit
    End If

    oGrid22.AutoResizeColumns
    oForm01.Update

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_MTX22_Exit:
    Set RecordSet01 = Nothing

    Call oForm01.Freeze(False)
    Exit Sub
PS_QM131_MTX22_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_QM131_MTX22_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_QM131_DI_API() As Boolean
'On Error GoTo PS_QM131_DI_API_Error
'    PS_QM131_DI_API = True
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
'            Call oDS_PS_QM131L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_QM131L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
'    Else
'        GoTo PS_QM131_DI_API_Error
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
'PS_QM131_DI_API_DI_Error:
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
'    PS_QM131_DI_API = False
'    Set oDIObject = Nothing
'    Exit Function
'PS_QM131_DI_API_Error:
'    Sbo_Application.SetStatusBarMessage "PS_QM131_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    PS_QM131_DI_API = False
End Function

Private Sub PS_QM131_SetBaseForm()
On Error GoTo PS_QM131_SetBaseForm_Error
    Dim i As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    If oBaseForm01 Is Nothing Then
    '//DoNothing
    Else
    
    End If
    Exit Sub
PS_QM131_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM131_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM131_Print_Report01()
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

    Dim BPLID As String
    Dim FrDt As String
    Dim ToDt As String
    Dim InOut As String
    Dim CardType As String
    Dim ItemType As String
    Dim ItemCode As String
    
    BPLID = oForm01.Items("BPLID21").Specific.VALUE
    FrDt = oForm01.Items("FrDt21").Specific.VALUE
    ToDt = oForm01.Items("ToDt21").Specific.VALUE
    InOut = oForm01.Items("InOut21").Specific.VALUE
    CardType = oForm01.Items("CardType21").Specific.VALUE
    ItemType = oForm01.Items("ItemType21").Specific.VALUE
    ItemCode = oForm01.Items("ItemCode21").Specific.VALUE

    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_QM131] 레포트"

    ReportName = "PS_QM131_01.rpt"

    '쿼리
    sQry = "         EXEC PS_QM131_51 '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & FrDt & "','"
    sQry = sQry & ToDt & "','"
    sQry = sQry & InOut & "','"
    sQry = sQry & CardType & "','"
    sQry = sQry & ItemType & "','"
    sQry = sQry & ItemCode & "'"

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
        MDC_Com.MDC_GF_Message "PS_QM131_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub








