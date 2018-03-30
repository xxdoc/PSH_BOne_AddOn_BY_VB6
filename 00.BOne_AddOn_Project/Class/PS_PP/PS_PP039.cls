VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP039"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_PP039.cls
'//  Module : 생산관리 > 애드온 > 생산관리
'//  Desc : 작업지시-공정추가등록,수정,삭제
'//  FormType : PS_PP039
'//  Create Date(Start) : 2014.12.09
'//  Create Date(End) : 2014.12.16
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************

'//사용방법
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_PP039
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oGrid01 As SAPbouiCOM.Grid
Private oDS_PS_PP039L As SAPbouiCOM.DBDataSource '등록라인

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String
Public oBaseItmBsort01 As String

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

''//사용자구조체
'Private Type ItemInformations
'    ItemCode As String
'    BatchNum As String
'    Quantity As Long
'    OPORNo As Long
'    POR1No As Long
'    Check As Boolean
'    OPDNNo As Long
'    PDN1No As Long
'End Type
'Private ItemInformation() As ItemInformations
'Private ItemInformationCount As Long
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String, Optional ByVal oItmBsort02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP039.srf")
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
    
    oFormUniqueID01 = "PS_PP039_" & GetTotalFormsCount
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
    
    Call PS_PP039_CreateItems
    Call PS_PP039_ComboBox_Setting
    Call PS_PP039_CF_ChooseFromList
    Call PS_PP039_FormItemEnabled
    Call PS_PP039_EnableMenus
    '//Call PS_PP039_FormClear '//UDO방식일때
    '//Call PS_PP039_AddMatrixRow(0, True) '//UDO방식일때
    
    oForm01.Items("FrDt").Specific.VALUE = Format(DateAdd("m", -2, Date), "YYYYMM01")
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
                '//Call PS_PP039_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_PP039_FormItemEnabled '//UDO방식
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
            
                If PS_PP039_CheckBeforeSearch() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                Call PS_PP039_MTX01
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        ElseIf pval.ItemUID = "BtnAdd" Then
            
            '필수 입력 사항 체크
            If PS_PP039_DataValidCheck = False Then
            
                BubbleEvent = False
                Exit Sub
                
            Else
                
                If PS_PP039_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                Call PS_PP039_MTX02(oForm01.Items("GridRow").Specific.VALUE)
                Call PS_PP039_AddMatrixRow(oMat01.RowCount)
                
            End If
        
        ElseIf pval.ItemUID = "BtnDel" Then
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
            
                If PS_PP039_DeleteData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                Call PS_PP039_MTX02(oForm01.Items("GridRow").Specific.VALUE)
                Call PS_PP039_AddMatrixRow(oMat01.RowCount)

            Else
            
            End If
            
'        ElseIf pval.ItemUID = "Mat01" And pval.ColUID = "Check" And pval.Row > 0 Then
'
'            If oMat01.RowCount >= pval.Row Then     '빈 Select 필드에 클릭했을 때 생기는 오류 수정을 위한 구문
'
'                If PS_PP039_CheckOKYN(pval.Row) = True Then
'
'                    Call MDC_Com.MDC_GF_Message("해당 자재는 구매요청승인이 이미 처리되어 선택할 수 없습니다", "E")
'
''                    Call oForm01.Freeze(True)
''                    Call oDS_PS_PP039L.setValue("U_ColReg01", pval.Row - 1, "N")
''                    Call oMat01.LoadFromDataSource
''                    Call oForm01.Freeze(False)
'
'                    oMat01.Columns("Check").Cells(pval.Row).Specific.Checked = False
'
'                    Exit Sub
'
'                End If
'
'            End If
        
        End If
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_PP039_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If

        
        
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP039" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP039_FormItemEnabled
'                    Call PS_PP039_FormClear '//UDO방식일때
'                    Call PS_PP039_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP039_FormItemEnabled
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
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.ColUID = "CpBCode" Then
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "CpBCode")
            
            ElseIf pval.ColUID = "CpCode" Then
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "CpCode")
            
            ElseIf pval.ColUID = "FailCode" Then
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "FailCode")
            
            End If
        
        Else
    
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '담당자
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum", "") '작번
        
        End If
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
    
        If pval.ItemUID = "Opt01" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Grid01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            Call oForm01.Freeze(False)
        End If
        If pval.ItemUID = "Opt02" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Mat01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            Call oForm01.Freeze(False)
        End If
    
        If pval.ItemUID = "Mat01" Then

            If pval.Row > 0 Then

                Call oMat01.SelectRow(pval.Row, True, False)
                
'                If pval.ColUID = "Check" Then
'
'                    If PS_PP039_CheckOKYN(pval.Row) = True Then
'
''                        Call oForm01.Freeze(True)
'
''                        Call oMat01.FlushToDataSource
''                        Call oMat01.LoadFromDataSource
'
'                        Call MDC_Com.MDC_GF_Message("해당 자재는 구매요청승인이 이미 처리되어 선택할 수 없습니다", "E")
''                        Call oDS_PS_PP039L.setValue("U_ColReg01", pval.Row - 1, False)
'                        oMat01.Columns("Check").Cells(pval.Row).Specific.Checked = False
'
''                        Call oMat01.LoadFromDataSource
'
''                        Call oForm01.Freeze(False)
'
'                        Exit Sub
'
'                    End If
'
'                End If

            End If
        
        End If
        
        If pval.ItemUID = "Grid01" Then
        
            If pval.Row >= 0 Then

                Call PS_PP039_MTX02(pval.Row)
                Call PS_PP039_AddMatrixRow(oMat01.RowCount)
                
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
'                If oGrid01.Rows.SelectedRows.Count > 0 Then
'                    Call PS_PP039_SetBaseForm '//부모폼에입력
'                    If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
'                        Call oForm01.Close
'                    End If
'                Else
'                    BubbleEvent = False
'                End If
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
    
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "Mat01") Then
            
                Call PS_PP039_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
        
            Else
        
                Call PS_PP039_FlushToItemValue(pval.ItemUID)
            
            End If
                
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
        Call PS_PP039_FormItemEnabled
        Call PS_PP039_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
    
        Call PS_PP039_FormResize
        
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
    '        Call oDS_PS_PP039L.RemoveRecord(oDS_PS_PP039L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_PP039_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_PP039_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP039_CreateItems() As Boolean
On Error GoTo PS_PP039_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    'Dim C_Date   As Date
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP039L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
'    oGrid01.SelectionMode = ms_NotSupported
'    oGrid01.au.AutoResizeColumns
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
    
    '작업구분
    Call oForm01.DataSources.UserDataSources.Add("OrdGbn", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdGbn").Specific.DataBind.SetBound(True, "", "OrdGbn")
        
    '작업지시일자(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    
    '작업지시일자(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    
    '담당자
    Call oForm01.DataSources.UserDataSources.Add("CntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcCode").Specific.DataBind.SetBound(True, "", "CntcCode")
    
    '담당자명
    Call oForm01.DataSources.UserDataSources.Add("CntcName", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcName").Specific.DataBind.SetBound(True, "", "CntcName")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum", dt_SHORT_TEXT, 11)
    Call oForm01.Items("OrdNum").Specific.DataBind.SetBound(True, "", "OrdNum")
    
    '서브작번1
    Call oForm01.DataSources.UserDataSources.Add("OrdSub1", dt_SHORT_TEXT, 2)
    Call oForm01.Items("OrdSub1").Specific.DataBind.SetBound(True, "", "OrdSub1")
    
    '서브작번2
    Call oForm01.DataSources.UserDataSources.Add("OrdSub2", dt_SHORT_TEXT, 3)
    Call oForm01.Items("OrdSub2").Specific.DataBind.SetBound(True, "", "OrdSub2")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("JakMyung", dt_SHORT_TEXT, 200)
    Call oForm01.Items("JakMyung").Specific.DataBind.SetBound(True, "", "JakMyung")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("JakSize", dt_SHORT_TEXT, 200)
    Call oForm01.Items("JakSize").Specific.DataBind.SetBound(True, "", "JakSize")
    
    '그리드에서 선택한 행의 작업지시 문서번호
    Call oForm01.DataSources.UserDataSources.Add("MainEntry", dt_SHORT_TEXT, 100)
    Call oForm01.Items("MainEntry").Specific.DataBind.SetBound(True, "", "MainEntry")
    
    '그리드에서 선택한 행 번호
    Call oForm01.DataSources.UserDataSources.Add("GridRow", dt_SHORT_TEXT, 100)
    Call oForm01.Items("GridRow").Specific.DataBind.SetBound(True, "", "GridRow")
    
    '그리드에서 선택한 작번(전체작번)
    Call oForm01.DataSources.UserDataSources.Add("FullOrdNum", dt_SHORT_TEXT, 100)
    Call oForm01.Items("FullOrdNum").Specific.DataBind.SetBound(True, "", "FullOrdNum")
    
    '공정금액 합계
    Call oForm01.DataSources.UserDataSources.Add("Total", dt_PRICE)
    Call oForm01.Items("Total").Specific.DataBind.SetBound(True, "", "Total")
    
    '//////////각 매트릭스 서식세팅 선택용 라디오버튼//////////S
    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt01").Specific.DataBind.SetBound(True, "", "Opt01")
    
    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt02").Specific.DataBind.SetBound(True, "", "Opt02")
    
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    '//////////각 매트릭스 서식세팅 선택용 라디오버튼//////////E
    
    '참조정보 관련 컨트롤 숨김_S
    'Haeder
    oForm01.Items("Static90").Visible = False
    oForm01.Items("Static91").Visible = False
    oForm01.Items("Static92").Visible = False
    oForm01.Items("MainEntry").Visible = False
    oForm01.Items("GridRow").Visible = False
    oForm01.Items("FullOrdNum").Visible = False
    
    'Line
'    oMat01.Columns("DocEntry").Visible = False '문서번호, 라인번호 히든 취소(2015.01.13 송명규)
'    oMat01.Columns("LineId").Visible = False
    oMat01.Columns("VisOrder").Visible = False
    oMat01.Columns("Object").Visible = False
    oMat01.Columns("LogInst").Visible = False
    oMat01.Columns("U_LineNum").Visible = False
    oMat01.AutoResizeColumns
    '참조정보 관련 컨트롤 숨김_E
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP039_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP039_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP039_ComboBox_Setting()
On Error GoTo PS_PP039_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '작업구분
    sQry = "           SELECT     Code AS [Code], "
    sQry = sQry & "               Name AS [Name]"
    sQry = sQry & " FROM      [@PSH_ITMBSORT]"
    sQry = sQry & " WHERE     U_PudYN = 'Y'"
    Call oForm01.Items("OrdGbn").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdGbn").Specific, sQry, "", False, False)
    Call oForm01.Items("OrdGbn").Specific.Select("105", psk_ByValue) '기계공구 기본 선택
    
    '//////////매트릭스//////////
'    '투입구분
'    Call oMat01.Columns("InputGbn").ValidValues.Add("10", "일반")
'    Call oMat01.Columns("InputGbn").ValidValues.Add("20", "원재료")
'    Call oMat01.Columns("InputGbn").ValidValues.Add("30", "스크랩")
'
'    '품목그룹
'    sQry = "           SELECT      ItmsGrpCod AS [Code], "
'    sQry = sQry & "                ItmsGrpNam AS [Name]"
'    sQry = sQry & " FROM       [OITB] a"
'
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemGpCd"), sQry)
'
'    '조달방식
'    Call oMat01.Columns("ProcType").ValidValues.Add("10", "청구")
'    Call oMat01.Columns("ProcType").ValidValues.Add("20", "잔재")
'    Call oMat01.Columns("ProcType").ValidValues.Add("30", "취소")

    '작업구분
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkGbn", "10", "자가")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkGbn", "20", "정밀")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkGbn", "30", "외주")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("WorkGbn"), "PS_PP039", "Mat01", "WorkGbn")
    
    '실적여부
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ResultYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ResultYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("ResultYN"), "PS_PP039", "Mat01", "ResultYN")
    
    '재작업여부
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ReWorkYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ReWorkYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("ReWorkYN"), "PS_PP039", "Mat01", "ReWorkYN")
    
    '일보여부
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ReportYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ReportYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("ReportYN"), "PS_PP039", "Mat01", "ReportYN")
    
    '작업여부
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("WorkYN"), "PS_PP039", "Mat01", "WorkYN")
    '//////////매트릭스//////////

    Call oForm01.Freeze(False)
    Exit Sub
PS_PP039_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP039_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP039_CF_ChooseFromList()
On Error GoTo PS_PP039_CF_ChooseFromList_Error
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
PS_PP039_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP039_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP039_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_PP039_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim sQry As String
    
    Dim OrdNum As String
    Dim OrdSub1 As String
    Dim OrdSub2 As String
    Dim TotalAmt As Double
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If oUID = "Mat01" Then
    
        If oCol = "CpBCode" Then
        
            Call oMat01.FlushToDataSource
            Call oDS_PS_PP039L.setValue("U_ColReg04", oRow - 1, MDC_PS_Common.GetValue("SELECT Name FROM [@PS_PP001H] WHERE Code = '" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'", 0, 1))
            Call oMat01.LoadFromDataSource
            
            If oMat01.RowCount = oRow And Trim(oDS_PS_PP039L.GetValue("U_ColReg02", oRow - 1)) <> "" Then
                Call PS_PP039_AddMatrixRow(oRow)
            End If
            
            Call oMat01.Columns.Item("CpBCode").Cells(oRow).Click '공정대분류 입력 후 TAB 조정
            
        ElseIf oCol = "CpCode" Then
            
            Call oMat01.FlushToDataSource
            Call oDS_PS_PP039L.setValue("U_ColReg06", oRow - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE Code = '" & oMat01.Columns("CpBCode").Cells(oRow).Specific.VALUE & "' AND U_CpCode = '" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'", 0, 1))
            
            '작업구분은 기본으로 "자가" 선택
            Call oDS_PS_PP039L.setValue("U_ColReg11", oRow - 1, "10")
            
            '공정선택 시 재작업여부 자동 선택
            If oMat01.Columns("CpCode").Cells(oRow).Specific.VALUE = "CP50103" Or oMat01.Columns("CpCode").Cells(oRow).Specific.VALUE = "CP50106" Then 'PK/탈지일때 재작업여부 "예"
            
                Call oDS_PS_PP039L.setValue("U_ColReg12", oRow - 1, "Y")
                
            Else
            
                Call oDS_PS_PP039L.setValue("U_ColReg12", oRow - 1, "N")
                
            End If
            
            Call oMat01.LoadFromDataSource
            
            Call oMat01.Columns.Item("CpCode").Cells(oRow).Click '공정중분류 입력 후 TAB 조정
            
        ElseIf oCol = "FailCode" Then
        
            Call oMat01.FlushToDataSource
            Call oDS_PS_PP039L.setValue("U_ColReg16", oRow - 1, MDC_PS_Common.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" & oMat01.Columns("FailCode").Cells(oRow).Specific.VALUE & "'", 0, 1))
            Call oMat01.LoadFromDataSource
            
            Call oMat01.Columns.Item("FailCode").Cells(oRow).Click '재작업사유 입력 후 TAB 조정
        
        End If
    
        If oCol = "StdHour" Or oCol = "ReDate" Then
        
            oMat01.FlushToDataSource '표준공수와 완료요구일은 수정이 가능해야 하므로 Flush 를 함
            
            If oCol = "StdHour" Then '표준공수 등록 시
                '공정단가 계산_S
                If oMat01.Columns("WorkGbn").Cells(oRow).Specific.VALUE = "10" Then
                    Call oDS_PS_PP039L.setValue("U_ColPrc01", oRow - 1, CDbl(MDC_PS_Common.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" & oMat01.Columns("CpCode").Cells(oRow).Specific.VALUE & "'", 0, 1)) * CDbl(oMat01.Columns(oCol).Cells(oRow).Specific.VALUE))
                ElseIf oMat01.Columns("WorkGbn").Cells(oRow).Specific.VALUE = "20" Then
                    Call oDS_PS_PP039L.setValue("U_ColPrc01", oRow - 1, MDC_PS_Common.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" & oMat01.Columns("CpCode").Cells(oRow).Specific.VALUE & "'", 0, 1) * oMat01.Columns(oCol).Cells(oRow).Specific.VALUE)
                End If
                Call oDS_PS_PP039L.setValue("U_ColQty01", oRow - 1, oMat01.Columns(oCol).Cells(oRow).Specific.VALUE)
                '공정단가 계산_E
                
                '합계 계산_S
                For loopCount = 0 To oMat01.VisualRowCount - 1
                
                    TotalAmt = TotalAmt + oDS_PS_PP039L.GetValue("U_ColPrc01", loopCount)
                    
                Next
                
                oForm01.Items("Total").Specific.VALUE = TotalAmt
                '합계 계산_E
            End If
    
            Call oMat01.LoadFromDataSource
    
            If oCol = "StdHour" Then

                Call oMat01.Columns.Item("StdHour").Cells(oRow).Click '표준공수 입력 후 TAB 조정

            ElseIf oCol = "ReDate" Then

                Call oMat01.Columns.Item("ReDate").Cells(oRow).Click '완료요구일 입력 후 TAB 조정

            End If
    
        End If
    
        Call oMat01.AutoResizeColumns
    
    ElseIf oUID = "CntcCode" Then
    
        oForm01.Items("CntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
    
    ElseIf oUID = "OrdNum" Or oUID = "OrdSub1" Or oUID = "OrdSub2" Then
        
        OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
        OrdSub1 = Trim(oForm01.Items("OrdSub1").Specific.VALUE)
        OrdSub2 = Trim(oForm01.Items("OrdSub2").Specific.VALUE)

        sQry = "           SELECT   CASE"
        sQry = sQry & "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakMyung"
        sQry = sQry & "             END AS [JakMyung],"
        sQry = sQry & "             CASE"
        sQry = sQry & "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakSize"
        sQry = sQry & "             END AS [JakSize]"
        sQry = sQry & " FROM     [@PS_PP020H] AS T0"
        sQry = sQry & " WHERE   T0.U_JakName = '" & OrdNum & "'"
        sQry = sQry & "             AND T0.U_SubNo1 = CASE WHEN '" & OrdSub1 & "' = '' THEN '00' ELSE '" & OrdSub1 & "' END"
        sQry = sQry & "             AND T0.U_SubNo2 = CASE WHEN '" & OrdSub2 & "' = '' THEN '000' ELSE '" & OrdSub2 & "' END"

        Call oRecordSet01.DoQuery(sQry)

        oForm01.Items("JakMyung").Specific.VALUE = oRecordSet01.Fields("JakMyung").VALUE
        oForm01.Items("JakSize").Specific.VALUE = oRecordSet01.Fields("JakSize").VALUE
    
    End If
    
    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_PP039_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("PS_PP039_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

Sub PS_PP039_FormItemEnabled()
On Error GoTo PS_PP039_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseItmBsort01 = "" Then
        
        Else
        
        End If
        '//각모드에따른 아이템설정
        '//Call PS_PP039_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP039_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP039_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP039_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP039_AddMatrixRow_Error

    Call oForm01.Freeze(True)
    
    If RowIserted = False Then '//행추가여부
        Call oDS_PS_PP039L.InsertRecord(oRow)
    End If
    
    Call oMat01.AddRow
    oDS_PS_PP039L.Offset = oRow
    Call oDS_PS_PP039L.setValue("U_LineNum", oRow, oRow + 1)
    Call oDS_PS_PP039L.setValue("U_ColReg02", oRow, oRow + 1)
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_PP039_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP039_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP039_FormClear()
On Error GoTo PS_PP039_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP039'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP039_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP039_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP039_EnableMenus()
On Error GoTo PS_PP039_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
'    Call MDC_GP_EnableMenus(oForm01, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP039_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP039_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP039_DataValidCheck() As Boolean
On Error GoTo PS_PP039_DataValidCheck_Error

'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_PP039_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_PP039_DataValidCheck = False
'        Exit Function
'    End If
    
    Dim loopCount As Long
    
    Dim CpBCode As String '공정대분류
    Dim CpCode As String '공정중분류
    
    Dim WorkGbn As String '작업구분
    Dim ReWorkYN As String '재작업여부
    Dim ResultYN As String '실적여부
    Dim ReportYN As String '일보여부
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
        If Trim(oDS_PS_PP039L.GetValue("U_ColReg01", loopCount)) = "Y" Then

            CpBCode = Trim(oDS_PS_PP039L.GetValue("U_ColReg03", loopCount)) '공정대분류
            CpCode = Trim(oDS_PS_PP039L.GetValue("U_ColReg05", loopCount)) '공정중분류
            
            WorkGbn = Trim(oDS_PS_PP039L.GetValue("U_ColReg11", loopCount)) '작업구분
            ReWorkYN = Trim(oDS_PS_PP039L.GetValue("U_ColReg12", loopCount)) '재작업여부
            ResultYN = Trim(oDS_PS_PP039L.GetValue("U_ColReg13", loopCount)) '실적여부
            ReportYN = Trim(oDS_PS_PP039L.GetValue("U_ColReg14", loopCount)) '일보여부

            If CpBCode = "" Then
            
                Call Sbo_Application.SetStatusBarMessage("공정대분류는 필수입니다.", bmt_Short, True)
                Call oMat01.Columns("CpBCode").Cells(loopCount + 1).Click(ct_Regular)
                PS_PP039_DataValidCheck = False
                Exit Function
                
            ElseIf CpCode = "" Then
            
                Call Sbo_Application.SetStatusBarMessage("공정중분류는 필수입니다.", bmt_Short, True)
                Call oMat01.Columns("CpCode").Cells(loopCount + 1).Click(ct_Regular)
                PS_PP039_DataValidCheck = False
                Exit Function
                          
            ElseIf WorkGbn = "" Then
                
                Call Sbo_Application.SetStatusBarMessage("작업구분은 필수입니다.", bmt_Short, True)
                Call oMat01.Columns("WorkGbn").Cells(loopCount + 1).Click(ct_Regular)
                PS_PP039_DataValidCheck = False
                Exit Function
            
            ElseIf ReWorkYN = "" Then
            
                Call Sbo_Application.SetStatusBarMessage("재작업여부는 필수입니다.", bmt_Short, True)
                Call oMat01.Columns("ReWorkYN").Cells(loopCount + 1).Click(ct_Regular)
                PS_PP039_DataValidCheck = False
                Exit Function
            
            ElseIf ResultYN = "" Then
            
                Call Sbo_Application.SetStatusBarMessage("실적여부는 필수입니다.", bmt_Short, True)
                Call oMat01.Columns("ResultYN").Cells(loopCount + 1).Click(ct_Regular)
                PS_PP039_DataValidCheck = False
                Exit Function
            
            ElseIf ReportYN = "" Then
                
                Call Sbo_Application.SetStatusBarMessage("일보여부는 필수입니다.", bmt_Short, True)
                Call oMat01.Columns("ReportYN").Cells(loopCount + 1).Click(ct_Regular)
                PS_PP039_DataValidCheck = False
                Exit Function
                
            End If

        End If

    Next
    
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
'    Call PS_PP039_FormClear

    PS_PP039_DataValidCheck = True

    Exit Function
PS_PP039_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP039_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP039_DataValidCheck = False
End Function

Private Sub PS_PP039_MTX01()
On Error GoTo PS_PP039_MTX01_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim loopCount01 As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String '사업장
    Dim OrdGbn As String '작업구분
    Dim FrDt As String '지시일자(Fr)
    Dim ToDt As String '지시일자(To)
    Dim CntcCode As String '담당자
    Dim OrdNum As String '작번
    Dim OrdSub1 As String '서브작번1
    Dim OrdSub2 As String '서브작번2
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.VALUE) '작업구분
    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE) '지시일자(Fr)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE) '지시일자(To)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE) '담당자
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE) '작번
    OrdSub1 = Trim(oForm01.Items("OrdSub1").Specific.VALUE) '서브작번1
    OrdSub2 = Trim(oForm01.Items("OrdSub2").Specific.VALUE) '서브작번2
    
    Query01 = "              EXEC PS_PP039_01 '"
    Query01 = Query01 & BPLID & "','" '사업장
    Query01 = Query01 & OrdGbn & "','" '작업구분
    Query01 = Query01 & FrDt & "','" '지시일자(Fr)
    Query01 = Query01 & ToDt & "','" '지시일자(To)
    Query01 = Query01 & CntcCode & "','" '담당자
    Query01 = Query01 & OrdNum & "','" '작번
    Query01 = Query01 & OrdSub1 & "','" '서브작번1
    Query01 = Query01 & OrdSub2 & "'" '서브작번2
    
    
    Call oGrid01.DataTable.Clear
    Call oForm01.DataSources.DataTables.Item("DataTable").ExecuteQuery(Query01)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
'    oGrid01.Columns(26).RightJustified = True '요청수량
'    oGrid01.Columns(27).RightJustified = True '실적단가
'    oGrid01.Columns(28).RightJustified = True '실적금액
'    oGrid01.Columns(31).RightJustified = True '품의수량/중량
'    oGrid01.Columns(32).RightJustified = True '품의금액
'    oGrid01.Columns(34).RightJustified = True '가입고수량/중량
'    oGrid01.Columns(35).RightJustified = True '가입고금액
'    oGrid01.Columns(37).RightJustified = True '검수수량/중량
'    oGrid01.Columns(38).RightJustified = True '검수금액

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

    If oGrid01.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP039_MTX01_Error
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
        
    oGrid01.AutoResizeColumns
    oForm01.Update
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub

PS_PP039_MTX01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP039_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_PP039_MTX02(ByVal pRow As Long)
'******************************************************************************
'Function ID : PS_PP039_MTX02()
'해당모듈 : PS_PP039
'기능 : 데이터 조회
'인수 : pRow(행번호)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP039_MTX02_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    Dim TotalAmt As Double
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    Dim FullOrdNum As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    DocEntry = Trim(oGrid01.DataTable.Columns.Item("문서번호").Cells(pRow).VALUE) '그리드에서 선택한 작업지시등록 문서번호
    FullOrdNum = Trim(oGrid01.DataTable.Columns.Item("작번").Cells(pRow).VALUE) & "-" & Trim(oGrid01.DataTable.Columns.Item("서브작번1").Cells(pRow).VALUE) & "-" & Trim(oGrid01.DataTable.Columns.Item("서브작번2").Cells(pRow).VALUE) '그리드에서 선택한 작번(전체작번)
    
    oForm01.Items("MainEntry").Specific.VALUE = DocEntry '그리드에서 선택한 행의 작업지시 문서번호 레이블에 저장
    oForm01.Items("GridRow").Specific.VALUE = pRow '그리드에서 선택한 행의 행번호
    oForm01.Items("FullOrdNum").Specific.VALUE = FullOrdNum '그리드에서 선택한 작번(전체작번)
    
    sQry = "         EXEC [PS_PP039_02] '"
    sQry = sQry & DocEntry & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_PP039L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

'    If (oRecordSet01.RecordCount = 0) Then
'
'        ErrNum = 1
''
''        oForm01.Mode = fm_ADD_MODE
'
''        Call PS_PP039_AddMatrixRow(0, True)
''        Call PS_PP039_LoadCaption
'
'        GoTo PS_PP039_MTX02_Error
'
'        Exit Sub
'    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP039L.Size Then
            oDS_PS_PP039L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_PP039L.Offset = i
        
        Call oDS_PS_PP039L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP039L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        
        Call oDS_PS_PP039L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("Sequence").VALUE)) '순번
        Call oDS_PS_PP039L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("CpBCode").VALUE)) '공정대분류
        Call oDS_PS_PP039L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("CpBName").VALUE)) '대분류명
        Call oDS_PS_PP039L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("CpCode").VALUE)) '공정중분류
        Call oDS_PS_PP039L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("CpName").VALUE)) '중분류명
        Call oDS_PS_PP039L.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("StdHour").VALUE)) '표준공수
        Call oDS_PS_PP039L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("Unit").VALUE)) '단위
        Call oDS_PS_PP039L.setValue("U_ColPrc01", i, Trim(oRecordSet01.Fields("CpPrice").VALUE)) '공정금액
'        Call oDS_PS_PP039L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("ReDate").VALUE), "YYYYMMDD")) '완료요구일
        Call oDS_PS_PP039L.setValue("U_ColDt01", i, Trim(oRecordSet01.Fields("ReDate").VALUE)) '완료요구일
        Call oDS_PS_PP039L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("WorkGbn").VALUE)) '작업구분
        Call oDS_PS_PP039L.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("ReWorkYN").VALUE)) '재작업여부
        Call oDS_PS_PP039L.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("ResultYN").VALUE)) '실적여부
        Call oDS_PS_PP039L.setValue("U_ColReg14", i, Trim(oRecordSet01.Fields("ReportYN").VALUE)) '일보여부
        Call oDS_PS_PP039L.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("FailCode").VALUE)) '재작업사유
        Call oDS_PS_PP039L.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("FailName").VALUE)) '재작업사유명
        Call oDS_PS_PP039L.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("WorkYN").VALUE)) '작업여부
        
        Call oDS_PS_PP039L.setValue("U_ColNum01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)) 'DocEntry
        Call oDS_PS_PP039L.setValue("U_ColNum02", i, Trim(oRecordSet01.Fields("LineId").VALUE)) 'LineId
        Call oDS_PS_PP039L.setValue("U_ColNum03", i, Trim(oRecordSet01.Fields("VisOrder").VALUE)) 'VisOrder
        Call oDS_PS_PP039L.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("Object").VALUE)) 'Object
        Call oDS_PS_PP039L.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("LogInst").VALUE)) 'LogInst
        Call oDS_PS_PP039L.setValue("U_ColNum04", i, Trim(oRecordSet01.Fields("LineNum").VALUE)) 'LineNum

        TotalAmt = TotalAmt + Trim(oRecordSet01.Fields("CpPrice").VALUE)

        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    '합계 계산_S
    oForm01.Items("Total").Specific.VALUE = TotalAmt
    '합계 계산_E
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP039_MTX02_Error:

    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP039_MTX02_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_PP039_AddData() As Boolean
'******************************************************************************
'Function ID : PS_PP039_AddData()
'해당모듈 : PS_PP039
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP039_AddData_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
 
    Dim Sequence As String '순번
    Dim CpBCode As String '공정대분류
    Dim CpBName As String '대분류명
    Dim CpCode As String '공정중분류
    Dim CpName As String '중분류명
    Dim StdHour As String '표준공수
    Dim Unit As String '단위
    Dim CpPrice As Double '공정금액
    Dim ReDate As String '완료요구일
    Dim WorkGbn As String '작업구분
    Dim ReWorkYN As String '재작업여부
    Dim ResultYN As String '실적여부
    Dim ReportYN As String '일보여부
    Dim FailCode As String '재작업사유
    Dim FailName As String '재작업사유명
    Dim WorkYN As String '작업여부
    
    Dim DocEntry As String 'DocEntry
    Dim LineId As String 'LineID
    Dim VisOrder As String 'VisOrder
    Dim Object As String 'Object
    Dim LogInst As String 'LogInst
    Dim LineNum As String 'U_LineNum
    
    Dim MainEntry As String '작업지시문서번호
    Dim BPLID As String '사업장코드
    Dim FullOrdNum As String '전체작번
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, False)
    
    MainEntry = Trim(oForm01.Items("MainEntry").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    FullOrdNum = Trim(oForm01.Items("FullOrdNum").Specific.VALUE)
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP039L.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            Sequence = Trim(oDS_PS_PP039L.GetValue("U_ColReg02", loopCount)) '순번
            CpBCode = Trim(oDS_PS_PP039L.GetValue("U_ColReg03", loopCount)) '공정대분류
            CpBName = Trim(oDS_PS_PP039L.GetValue("U_ColReg04", loopCount)) '대분류명
            CpCode = Trim(oDS_PS_PP039L.GetValue("U_ColReg05", loopCount)) '공정중분류
            CpName = Trim(oDS_PS_PP039L.GetValue("U_ColReg06", loopCount)) '중분류명
            StdHour = Trim(oDS_PS_PP039L.GetValue("U_ColQty01", loopCount)) '표준공수
            Unit = Trim(oDS_PS_PP039L.GetValue("U_ColReg08", loopCount)) '단위
            CpPrice = Trim(oDS_PS_PP039L.GetValue("U_ColPrc01", loopCount)) '공정금액
            ReDate = Trim(oDS_PS_PP039L.GetValue("U_ColDt01", loopCount)) '완료요구일
            WorkGbn = Trim(oDS_PS_PP039L.GetValue("U_ColReg11", loopCount)) '작업구분
            ReWorkYN = Trim(oDS_PS_PP039L.GetValue("U_ColReg12", loopCount)) '재작업여부
            ResultYN = Trim(oDS_PS_PP039L.GetValue("U_ColReg13", loopCount)) '실적여부
            ReportYN = Trim(oDS_PS_PP039L.GetValue("U_ColReg14", loopCount)) '일보여부
            FailCode = Trim(oDS_PS_PP039L.GetValue("U_ColReg15", loopCount)) '재작업사유
            FailName = Trim(oDS_PS_PP039L.GetValue("U_ColReg16", loopCount)) '재작업사유명
            WorkYN = Trim(oDS_PS_PP039L.GetValue("U_ColReg17", loopCount)) '작업여부
            
            DocEntry = Trim(oDS_PS_PP039L.GetValue("U_ColNum01", loopCount)) 'DocEntry
            LineId = Trim(oDS_PS_PP039L.GetValue("U_ColNum02", loopCount)) 'LineId
            VisOrder = Trim(oDS_PS_PP039L.GetValue("U_ColNum03", loopCount)) 'VisOrder
            Object = Trim(oDS_PS_PP039L.GetValue("U_ColReg18", loopCount)) 'Object
            LogInst = Trim(oDS_PS_PP039L.GetValue("U_ColReg19", loopCount)) 'LogInst
            LineNum = Trim(oDS_PS_PP039L.GetValue("U_ColNum04", loopCount)) 'U_LineNum

            sQry = "                EXEC [PS_PP039_03] "
            sQry = sQry & "'" & Sequence & "'," '순번
            sQry = sQry & "'" & CpBCode & "',"  '공정대분류
            sQry = sQry & "'" & CpBName & "',"  '대분류명
            sQry = sQry & "'" & CpCode & "',"  '공정중분류
            sQry = sQry & "'" & CpName & "',"  '중분류명
            sQry = sQry & "'" & StdHour & "',"  '표준공수
            sQry = sQry & "'" & Unit & "',"  '단위
            sQry = sQry & "'" & CpPrice & "',"  '공정금액
            sQry = sQry & "'" & ReDate & "',"  '완료요구일
            sQry = sQry & "'" & WorkGbn & "',"  '작업구분
            sQry = sQry & "'" & ReWorkYN & "',"  '재작업여부
            sQry = sQry & "'" & ResultYN & "',"  '실적여부
            sQry = sQry & "'" & ReportYN & "',"  '일보여부
            sQry = sQry & "'" & FailCode & "'," '재작업사유
            sQry = sQry & "'" & FailName & "'," '재작업사유명
            sQry = sQry & "'" & WorkYN & "'," '작업여부
            
            sQry = sQry & "'" & DocEntry & "',"  'DocEntry
            sQry = sQry & "'" & LineId & "',"  'LineID
            sQry = sQry & "'" & VisOrder & "',"  'VisOrder
            sQry = sQry & "'" & Object & "',"  'Object
            sQry = sQry & "'" & LogInst & "',"  'LogInst
            sQry = sQry & "'" & LineNum & "'," 'LineNum
            
            sQry = sQry & "'" & MainEntry & "'," '선택한 작업지시문서번호
            sQry = sQry & "'" & BPLID & "'," '사업장코드
            sQry = sQry & "'" & FullOrdNum & "'" '전체작번

            Call RecordSet01.DoQuery(sQry)
    
        End If
        
    Next

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("등록 완료!", "S")
    
    PS_PP039_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP039_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_PP039_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
    
    PS_PP039_AddData = False
    
End Function

Public Function PS_PP039_DeleteData() As Boolean
'******************************************************************************
'Function ID : PS_PP039_DeleteData()
'해당모듈 : PS_PP039
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP039_DeleteData_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim DocEntry As String '문서번호
    Dim LineId As String '라인번호
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oMat01.VisualRowCount = 0 Then
    
        ErrNum = 1
        GoTo PS_PP039_DeleteData_Error
    
    End If
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
        If Trim(oDS_PS_PP039L.GetValue("U_ColReg01", loopCount)) = "Y" Then '체크된 행만 'And Trim(oDS_PS_PP039L.GetValue("U_ColReg17", loopCount)) = "N" Then '선택된 행 중에 작업일보가 등록되지 않은 공정만 삭제
    
            DocEntry = Trim(oDS_PS_PP039L.GetValue("U_ColNum01", loopCount)) '문서번호
            LineId = Trim(oDS_PS_PP039L.GetValue("U_ColNum02", loopCount)) '라인번호
        
            sQry = "                 EXEC [PS_PP039_05] " '외주제작청구되었는지 체크
            sQry = sQry & "'" & DocEntry & "'," '문서번호
            sQry = sQry & "'" & LineId & "'"  '라인번호
            
            Call oRecordSet01.DoQuery(sQry)
            
            sQry = ""
            
            If oRecordSet01.Fields("CNT").VALUE > 0 Then
                ErrNum = 2
                GoTo PS_PP039_DeleteData_Error
            End If
            
            sQry = "                 EXEC [PS_PP039_06] " '작업일보가 등록되었는지 체크
            sQry = sQry & "'" & DocEntry & "'," '문서번호
            sQry = sQry & "'" & LineId & "'"  '라인번호
            
            Call oRecordSet01.DoQuery(sQry)
            
            sQry = ""
            
            If oRecordSet01.Fields("CNT").VALUE > 0 Then
                ErrNum = 3
                GoTo PS_PP039_DeleteData_Error
            End If
            
            sQry = "                 EXEC [PS_PP039_04] "
            sQry = sQry & "'" & DocEntry & "'," '문서번호
            sQry = sQry & "'" & LineId & "'"  '라인번호
            
            Call oRecordSet01.DoQuery(sQry)
    
        End If
    
    Next
    
    Set oRecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    PS_PP039_DeleteData = True
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP039_DeleteData_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    ElseIf ErrNum = 2 Then
        Call MDC_Com.MDC_GF_Message(loopCount + 1 & "행(" & DocEntry & "-" & LineId & ")은 외주제작청구가 등록되었습니다. 삭제할 수 없습니다. 삭제명령은 중단됩니다.", "W")
    ElseIf ErrNum = 3 Then
        Call MDC_Com.MDC_GF_Message(loopCount + 1 & "행(" & DocEntry & "-" & LineId & ")은 작업일보가 등록되었습니다. 삭제할 수 없습니다. 삭제명령은 중단됩니다.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP039_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
    
    PS_PP039_DeleteData = False
    
End Function


Private Function PS_PP039_DI_API() As Boolean
'On Error GoTo PS_PP039_DI_API_Error
'    PS_PP039_DI_API = True
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
'            Call oDS_PS_PP039L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_PP039L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
'    Else
'        GoTo PS_PP039_DI_API_Error
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
'PS_PP039_DI_API_DI_Error:
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
'    PS_PP039_DI_API = False
'    Set oDIObject = Nothing
'    Exit Function
'PS_PP039_DI_API_Error:
'    Sbo_Application.SetStatusBarMessage "PS_PP039_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    PS_PP039_DI_API = False
End Function

Private Sub PS_PP039_SetBaseForm()
On Error GoTo PS_PP039_SetBaseForm_Error
    Dim i As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    If oBaseForm01 Is Nothing Then
    '//DoNothing
    Else
    
    End If
    Exit Sub
PS_PP039_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP039_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP039_FormResize()
On Error GoTo PS_PP039_FormResize_Error
    
    oForm01.Items("Grid01").Height = oForm01.Height / 2 - 100
    oForm01.Items("Grid01").Width = oForm01.Width - 30
    
    If oGrid01.Columns.Count > 0 Then
        Call oGrid01.AutoResizeColumns
    End If
    
    oForm01.Items("Opt02").Top = oForm01.Height / 2 + 10
    
    oForm01.Items("Static93").Top = oForm01.Height / 2 + 10
    oForm01.Items("Static93").Left = oForm01.Items("Opt02").Width + 955
    oForm01.Items("Total").Top = oForm01.Items("Static93").Top
    oForm01.Items("Total").Left = oForm01.Items("Static93").Left + oForm01.Items("Static93").Width
    
    oForm01.Items("Static94").Left = oForm01.Items("Opt02").Width + 50
    oForm01.Items("Static94").Top = oForm01.Items("Static93").Top
    oForm01.Items("Static94").Left = oForm01.Items("Opt02").Width + 50
    
    oForm01.Items("Mat01").Top = oForm01.Items("Opt02").Top + 15
    oForm01.Items("Mat01").Height = oForm01.Items("Mat01").Top - 120
    oForm01.Items("Mat01").Width = oForm01.Width - 30
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_PP039_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP039_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP039_CheckBeforeSearch() As Boolean
'******************************************************************************
'Function ID : PS_PP039_CheckBeforeSearch()
'해당모듈 : PS_PP039
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP039_CheckBeforeSearch_Error

    Dim ErrNum As Integer
    ErrNum = 0
    
    Select Case True
        Case Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "%" '작업구분
            ErrNum = 1
            GoTo PS_PP039_CheckBeforeSearch_Error
    End Select

    PS_PP039_CheckBeforeSearch = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP039_CheckBeforeSearch_Error:
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회조건 작업구분은 필수선택 사항입니다. 확인하세요.", "E")
        Call oForm01.Items("OrdGbn").Click(ct_Regular)
    End If
    PS_PP039_CheckBeforeSearch = False
End Function

Private Function PS_PP039_CheckOKYN(ByVal pRow As Integer) As Boolean
'******************************************************************************
'Function ID : PS_PP039_CheckOKYN()
'해당모듈 : PS_PP039
'기능 : 구매요청승인 상태 확인
'인수 : 없음
'반환값 : True:구매요청승인Y, Fasle:구매요청승인N
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP039_CheckOKYN_Error

    Dim PP030DL As String
    Dim ErrNum As Integer
    Dim loopCount01 As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    ErrNum = 0
    
'    Call oMat01.FlushToDataSource
    PP030DL = oDS_PS_PP039L.GetValue("U_ColNum01", pRow - 1) & "-" & oDS_PS_PP039L.GetValue("U_ColNum02", pRow - 1)
    
    Query01 = "                SELECT     U_OKYN AS [OKYN]"
    Query01 = Query01 & " FROM      [@PS_MM005H] "
    Query01 = Query01 & " WHERE     U_PP030DL = '" & PP030DL & "'"
    Query01 = Query01 & "               AND U_OrdType = '10'" '원자재 구매요청만 조회

    Call RecordSet01.DoQuery(Query01)
    
    If RecordSet01.Fields("OKYN").VALUE = "Y" Then
        PS_PP039_CheckOKYN = True
    Else
        PS_PP039_CheckOKYN = False
    End If
    
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP039_CheckOKYN_Error:
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회조건 작업구분은 필수선택 사항입니다. 확인하세요.", "E")
        Call oForm01.Items("OrdGbn").Click(ct_Regular)
    End If
    PS_PP039_CheckOKYN = False
End Function

Private Sub PS_PP039_Print_Report01()
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
            
    Call ConnectODBC
    
    Dim BPLID As String
    Dim FrDt As String
    Dim ToDt As String
    Dim Purchase As String
    Dim OrderYN As String
    Dim InputYN As String
    Dim Class As String
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE)
    Purchase = Trim(oForm01.Items("Purchase").Specific.VALUE)
    OrderYN = Trim(oForm01.Items("OrderYN").Specific.VALUE)
    InputYN = Trim(oForm01.Items("InputYN").Specific.VALUE)
    Class = Trim(oForm01.Items("Class").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP039] 레포트"
    
    If Purchase = "10" Then '원재료품의
    
        ReportName = "PS_PP039_10.rpt"
        
        '쿼리
        sQry = "         EXEC PS_PP039_10 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & OrderYN & "','"
        sQry = sQry & InputYN & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
    
    ElseIf Purchase = "20" Then '부재료품의
    
        ReportName = "PS_PP039_20.rpt"
    
        '쿼리
        sQry = "         EXEC PS_PP039_20 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & OrderYN & "','"
        sQry = sQry & InputYN & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
    
    ElseIf Purchase = "30" Then '가공비품의
    
        ReportName = "PS_PP039_30.rpt"
        
        '쿼리
        sQry = "         EXEC PS_PP039_30 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
    
    ElseIf Purchase = "40" Then ''외주제작품의
    
        ReportName = "PS_PP039_40.rpt"
        
        '쿼리
        sQry = "         EXEC PS_PP039_40 '"
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
        MDC_Com.MDC_GF_Message "PS_PP039_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
