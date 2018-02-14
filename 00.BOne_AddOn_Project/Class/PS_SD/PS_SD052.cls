VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD052"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_SD052.cls
'//  Module             : 판매관리 > 애드온 > 기성매출관리
'//  Desc               : 작번별청구현황조회
'//  FormType           : PS_SD052
'//  Create Date(Start) : 2014.05.26
'//  Create Date(End)   : 2014.05.27
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************

'//사용방법
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_SD052
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD052.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_SD052_" & GetTotalFormsCount
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
    
    Call PS_SD052_CreateItems
    Call PS_SD052_ComboBox_Setting
    Call PS_SD052_CF_ChooseFromList
    Call PS_SD052_FormItemEnabled
    Call PS_SD052_EnableMenus
    '//Call PS_SD052_FormClear '//UDO방식일때
    '//Call PS_SD052_AddMatrixRow(0, True) '//UDO방식일때
    
    oForm01.Items("POType20YN").Specific.Checked = True
    
    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM01")
    oForm01.Items("ToDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    Call oForm01.Items("OrdNum").Click
    
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
                '//Call PS_SD052_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_SD052_FormItemEnabled '//UDO방식
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
                Call PS_SD052_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        ElseIf pval.ItemUID = "BtnPrint" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_SD052_Print_Report01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_SD052_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_SD052" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SD052_FormItemEnabled
'                    Call PS_SD052_FormClear '//UDO방식일때
'                    Call PS_SD052_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SD052_FormItemEnabled
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "SItemCode", "") '//사용자값활성
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
                    Call PS_SD052_SetBaseForm '//부모폼에입력
                    If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
                        Call oForm01.Close
                    End If
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
    
        If pval.ItemUID = "OrdNum" Then
        
            oForm01.Items("FrgnName").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
            
        ElseIf pval.ItemUID = "SItemCode" Then
        
            oForm01.Items("SItemName").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
        
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
        Call PS_SD052_FormItemEnabled
        '//Call PS_SD052_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
    
        Call PS_SD052_FormResize
        
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
    '        Call oDS_PS_SD052L.RemoveRecord(oDS_PS_SD052L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_SD052_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_SD052_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_SD052_CreateItems() As Boolean
On Error GoTo PS_SD052_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    'Dim C_Date   As Date
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    'oGrid01.SelectionMode = ms_NotSupported
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
    
    '기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    
    '기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    
    '품목코드(작번)
    Call oForm01.DataSources.UserDataSources.Add("OrdNum", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdNum").Specific.DataBind.SetBound(True, "", "OrdNum")
    
    '품목명
    Call oForm01.DataSources.UserDataSources.Add("FrgnName", dt_SHORT_TEXT, 200)
    Call oForm01.Items("FrgnName").Specific.DataBind.SetBound(True, "", "FrgnName")
    
    '품목코드(자재)
    Call oForm01.DataSources.UserDataSources.Add("SItemCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SItemCode").Specific.DataBind.SetBound(True, "", "SItemCode")
    
    '품목명(자재)
    Call oForm01.DataSources.UserDataSources.Add("SItemName", dt_SHORT_TEXT, 200)
    Call oForm01.Items("SItemName").Specific.DataBind.SetBound(True, "", "SItemName")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardType").Specific.DataBind.SetBound(True, "", "CardType")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemType").Specific.DataBind.SetBound(True, "", "ItemType")
    
    '품의구분
    Call oForm01.DataSources.UserDataSources.Add("POType", dt_SHORT_TEXT, 20)
    Call oForm01.Items("POType").Specific.DataBind.SetBound(True, "", "POType")
    
    '품의완료여부
    Call oForm01.DataSources.UserDataSources.Add("POYN", dt_SHORT_TEXT, 20)
    Call oForm01.Items("POYN").Specific.DataBind.SetBound(True, "", "POYN")
    
    '부자재품의 제외 여부
    Call oForm01.DataSources.UserDataSources.Add("POType20YN", dt_SHORT_TEXT, 1)
    Call oForm01.Items("POType20YN").Specific.DataBind.SetBound(True, "", "POType20YN")
    
    '실적금액계
    Call oForm01.DataSources.UserDataSources.Add("TResAmt", dt_SUM)
    Call oForm01.Items("TResAmt").Specific.DataBind.SetBound(True, "", "TResAmt")
    
    '품의금액계
    Call oForm01.DataSources.UserDataSources.Add("T030Amt", dt_SUM)
    Call oForm01.Items("T030Amt").Specific.DataBind.SetBound(True, "", "T030Amt")
    
    '검수금액계
    Call oForm01.DataSources.UserDataSources.Add("T070Amt", dt_SUM)
    Call oForm01.Items("T070Amt").Specific.DataBind.SetBound(True, "", "T070Amt")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_SD052_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD052_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SD052_ComboBox_Setting()
On Error GoTo PS_SD052_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '거래처구분
    sQry = "        SELECT      U_Minor AS [Code], "
    sQry = sQry & "             U_CdName AS [Name]"
    sQry = sQry & " FROM        [@PS_SY001L]"
    sQry = sQry & " WHERE       Code = 'C100'"
    Call oForm01.Items("CardType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType").Specific, sQry, "", False, False)
    Call oForm01.Items("CardType").Specific.Select("%", psk_ByValue)
    
    '품목구분
    sQry = "        SELECT      U_Minor AS [Code], "
    sQry = sQry & "             U_CdName AS [Name]"
    sQry = sQry & " FROM        [@PS_SY001L]"
    sQry = sQry & " WHERE       Code = 'S002'"
    Call oForm01.Items("ItemType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType").Specific, sQry, "", False, False)
    Call oForm01.Items("ItemType").Specific.Select("%", psk_ByValue)
    
    '품의구분
    sQry = "        SELECT      Code AS [Code],"
    sQry = sQry & "             Name AS [Name]"
    sQry = sQry & " FROM        [@PSH_ORDTYP]"
    sQry = sQry & " WHERE       Code IN ('10','20','30','40')" '4개 품의에 대해서만 조회
    sQry = sQry & " ORDER BY    Code"
    Call oForm01.Items("POType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("POType").Specific, sQry, "", False, False)
    Call oForm01.Items("POType").Specific.Select("%", psk_ByValue)
    
    '품의완료여부
    Call oForm01.Items("POYN").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("POYN").Specific.ValidValues.Add("Y", "품의완료")
    Call oForm01.Items("POYN").Specific.ValidValues.Add("N", "품의미완료")
    Call oForm01.Items("POYN").Specific.Select(0, psk_Index)
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD052_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD052_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD052_CF_ChooseFromList()
On Error GoTo PS_SD052_CF_ChooseFromList_Error
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
PS_SD052_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD052_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD052_FormItemEnabled()
On Error GoTo PS_SD052_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseItmBsort01 = "" Then
        
        Else
        
        End If
        '//각모드에따른 아이템설정
        '//Call PS_SD052_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD052_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD052_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD052_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD052_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_SD052L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_SD052L.Offset = oRow
'    oDS_PS_SD052L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD052_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD052_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD052_FormClear()
On Error GoTo PS_SD052_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD052'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SD052_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD052_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD052_EnableMenus()
On Error GoTo PS_SD052_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_SD052_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD052_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD052_DataValidCheck() As Boolean
On Error GoTo PS_SD052_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_SD052_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_SD052_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_SD052_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_SD052_FormClear
    Exit Function
PS_SD052_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD052_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD052_MTX01()
On Error GoTo PS_SD052_MTX01_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim loopCount01 As Long
    Dim Query01     As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID       As String '사업장
    Dim FrDt        As String '기간(Fr)
    Dim ToDt        As String '기간(To)
    Dim OrdNum      As String '품목코드(작번)
    Dim SItemCode   As String '품목코드(자재)
    Dim CardType    As String '거래처구분
    Dim ItemType    As String '품목구분
    Dim POType      As String '품의구분
    Dim POYN        As String '품의완료여부
    Dim POType20YN  As String '부자재품의 제외 여부
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    BPLID = oForm01.Items("BPLID").Specific.VALUE '사업장
    FrDt = oForm01.Items("FrDt").Specific.VALUE '기간(Fr)
    ToDt = oForm01.Items("ToDt").Specific.VALUE '기간(To)
    OrdNum = oForm01.Items("OrdNum").Specific.VALUE '품목코드(작번)
    SItemCode = oForm01.Items("SItemCode").Specific.VALUE '품목코드(자재)
    CardType = oForm01.Items("CardType").Specific.VALUE '거래처구분
    ItemType = oForm01.Items("ItemType").Specific.VALUE '품목구분
    POType = oForm01.Items("POType").Specific.VALUE '품의구분
    POYN = oForm01.Items("POYN").Specific.VALUE '품의완료여부
    POType20YN = IIf(oForm01.Items("POType20YN").Specific.Checked = True, "Y", "N") '부자재품의 제외 여부
    
    Query01 = "         EXEC PS_SD052_01 '"
    Query01 = Query01 & BPLID & "','" '사업장
    Query01 = Query01 & FrDt & "','" '기간(Fr)
    Query01 = Query01 & ToDt & "','" '기간(To)
    Query01 = Query01 & OrdNum & "','" '품목코드(작번)
    Query01 = Query01 & SItemCode & "','" '품목코드(자재)
    Query01 = Query01 & CardType & "','" '거래처구분
    Query01 = Query01 & ItemType & "','" '품목구분
    Query01 = Query01 & POType & "','" '품의구분
    Query01 = Query01 & POYN & "','" '품의완료여부
    Query01 = Query01 & POType20YN & "'" '부자재품의 제외 여부
    
    Call oGrid01.DataTable.Clear
    Call oForm01.DataSources.DataTables.Item("DataTable").ExecuteQuery(Query01)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
    oGrid01.Columns(29).RightJustified = True '요청수량
    oGrid01.Columns(30).RightJustified = True '실적단가
    oGrid01.Columns(31).RightJustified = True '실적금액
    oGrid01.Columns(34).RightJustified = True '품의수량/중량
    oGrid01.Columns(35).RightJustified = True '품의금액
    oGrid01.Columns(38).RightJustified = True '대비견적금액1
    oGrid01.Columns(40).RightJustified = True '대비견적금액2
    oGrid01.Columns(42).RightJustified = True '대비견적금액3
    oGrid01.Columns(44).RightJustified = True '가입고수량/중량
    oGrid01.Columns(45).RightJustified = True '가입고금액
    oGrid01.Columns(51).RightJustified = True '검수수량/중량
    oGrid01.Columns(52).RightJustified = True '검수금액
    oGrid01.Columns(59).RightJustified = True '수주금액
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '수주일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '요청일, 노랑
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '견적일, 노랑
'    oGrid01.Columns(14).BackColor = RGB(0, 210, 255) '차이(요청-견적), 하늘
'    oGrid01.Columns(15).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(18).BackColor = RGB(0, 210, 255) '차이(견적-품의), 하늘
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(20).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(21).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(22).BackColor = RGB(0, 210, 255) '차이(가입고-검수), 하늘
'    oGrid01.Columns(23).BackColor = RGB(255, 167, 167) '총소요일, 빨강
    
    Dim TResultAmt As Double '실적금액계
    Dim TMM030Amt As Double '품의금액계
    Dim TMM070Amt As Double '검수금액계
    
    For loopCount01 = 0 To oGrid01.Rows.Count - 1
        
        TResultAmt = TResultAmt + IIf(oGrid01.DataTable.GetValue(31, loopCount01) = "", 0, oGrid01.DataTable.GetValue(31, loopCount01))
        TMM030Amt = TMM030Amt + IIf(oGrid01.DataTable.GetValue(35, loopCount01) = "", 0, oGrid01.DataTable.GetValue(35, loopCount01))
        TMM070Amt = TMM070Amt + IIf(oGrid01.DataTable.GetValue(52, loopCount01) = "", 0, oGrid01.DataTable.GetValue(52, loopCount01))
    
    Next
    
    oForm01.Items("TResAmt").Specific.VALUE = TResultAmt
    oForm01.Items("T030Amt").Specific.VALUE = TMM030Amt
    oForm01.Items("T070Amt").Specific.VALUE = TMM070Amt

    If oGrid01.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_SD052_MTX01_Error
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
        
    oGrid01.AutoResizeColumns
    oForm01.Update
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub

PS_SD052_MTX01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD052_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SD052_DI_API() As Boolean
'On Error GoTo PS_SD052_DI_API_Error
'    PS_SD052_DI_API = True
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
'    oDIObject.BPL_IDAssignedToInvoice = oForm01.Items("BPLId").Specific.Selected.Value
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
'            Call oDS_PS_SD052L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_SD052L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
'    Else
'        GoTo PS_SD052_DI_API_Error
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
'PS_SD052_DI_API_DI_Error:
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
'    PS_SD052_DI_API = False
'    Set oDIObject = Nothing
'    Exit Function
'PS_SD052_DI_API_Error:
'    Sbo_Application.SetStatusBarMessage "PS_SD052_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    PS_SD052_DI_API = False
End Function

Private Sub PS_SD052_SetBaseForm()
On Error GoTo PS_SD052_SetBaseForm_Error
    Dim i As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    If oBaseForm01 Is Nothing Then
    '//DoNothing
    Else
    
    End If
    Exit Sub
PS_SD052_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD052_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD052_FormResize()
On Error GoTo PS_SD052_FormResize_Error
    
    Exit Sub
PS_SD052_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD052_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD052_Print_Report01()
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
            
    Call ConnectODBC
    
    Dim BPLID    As String
    Dim FrDt     As String
    Dim ToDt     As String
    Dim Purchase As String
    Dim OrderYN  As String
    Dim InputYN  As String
    Dim Class    As String
    
    BPLID = oForm01.Items("BPLId").Specific.VALUE
    FrDt = oForm01.Items("FrDt").Specific.VALUE
    ToDt = oForm01.Items("ToDt").Specific.VALUE
    Purchase = oForm01.Items("Purchase").Specific.VALUE
    OrderYN = oForm01.Items("OrderYN").Specific.VALUE
    InputYN = oForm01.Items("InputYN").Specific.VALUE
    Class = oForm01.Items("Class").Specific.VALUE
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_SD052] 레포트"
    
    If Purchase = "10" Then '원재료품의
    
        ReportName = "PS_SD052_10.rpt"
        
        '쿼리
        sQry = "      EXEC PS_SD052_10 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & OrderYN & "','"
        sQry = sQry & InputYN & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
    
    ElseIf Purchase = "20" Then '부재료품의
    
        ReportName = "PS_SD052_20.rpt"
    
        '쿼리
        sQry = "      EXEC PS_SD052_20 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & OrderYN & "','"
        sQry = sQry & InputYN & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
    
    ElseIf Purchase = "30" Then '가공비품의
    
        ReportName = "PS_SD052_30.rpt"
        
        '쿼리
        sQry = "      EXEC PS_SD052_30 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
    
    ElseIf Purchase = "40" Then ''외주제작품의
    
        ReportName = "PS_SD052_40.rpt"
        
        '쿼리
        sQry = "      EXEC PS_SD052_40 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
        
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
        MDC_Com.MDC_GF_Message "PS_SD052_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub




