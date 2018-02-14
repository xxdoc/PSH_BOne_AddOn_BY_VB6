VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA920"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_GA920.cls
'//  Module         : GA
'//  Desc           : 출입자 입문처리
'//  FormType       : PS_GA920
'//  Create Date    : 2014.07.31
'//  Copyright  (c) Morning Data
'//********************************************************************************

'//사용방법
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_GA920
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid
Public oGrid02 As SAPbouiCOM.Grid
Public oGrid03 As SAPbouiCOM.Grid

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA920.srf")
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
    
    oFormUniqueID01 = "PS_GA920_" & GetTotalFormsCount
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
    
    Call PS_GA920_CreateItems
    Call PS_GA920_ComboBox_Setting
    Call PS_GA920_CF_ChooseFromList
    Call PS_GA920_FormItemEnabled
    Call PS_GA920_EnableMenus
    '//Call PS_GA920_FormClear '//UDO방식일때
    '//Call PS_GA920_AddMatrixRow(0, True) '//UDO방식일때
    
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.EnableMenu ("1282"), True
    
    oForm01.Visible = True
    
    '최초 커서위치
    oForm01.ActiveItem = "DocDate"
    
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
                oGrid01.DataTable.Clear
                oGrid02.DataTable.Clear
                oGrid03.DataTable.Clear
    
                Call PS_GA920_Form_ini
                
                oForm01.Items("Btn_ret").Click
            
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_GA920_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_GA920_FormItemEnabled '//UDO방식
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
    Dim TempForm01 As Variant
    
    If pval.BeforeAction = True Then
'        If pval.ItemUID = "Btn_ret" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_GA920_MTX01
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
'        If pval.ItemUID = "Button02" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_GA920_SetBaseForm '//부모폼에입력
'                If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
'                    Call oForm01.Close
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
        
        If pval.ItemUID = "Btn_ret" Then
           Call PS_GA920_MTX01
        End If
               
        If pval.ItemUID = "Btn01" Then
           Call PS_GA920_SAVE
        End If
                             
        If pval.ItemUID = "Btn_del" Then
           Call PS_GA920_Delete
           Call PS_GA920_FormItemEnabled
        End If
        
               
        
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_GA920_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA920" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_GA920_FormItemEnabled
'                    Call PS_GA920_FormClear '//UDO방식일때
'                    Call PS_GA920_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_GA920_FormItemEnabled
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
        If pval.CharPressed = 9 Then
                If pval.ItemUID = "CoRegNum" Then '//사업자번호
                    If oForm01.Items("CoRegNum").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                 ElseIf pval.ItemUID = "CustNm" Then '/업체명
                    If oForm01.Items("CustNm").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                 ElseIf pval.ItemUID = "MSTCOD" Then '/사번
                    If oForm01.Items("MSTCOD").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
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
        Call PS_GA920_FlushToItemValue(pval.ItemUID)
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
                   PS_GA920_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        ElseIf pval.ItemUID = "Grid02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row >= 0 Then
                   PS_GA920_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        ElseIf pval.ItemUID = "Grid03" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row >= 0 Then
                   PS_GA920_MTX03 pval.ItemUID, pval.Row, pval.ColUID
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
    
    Dim BPLId As String
    Dim Visitant As String
    Dim BirthDat As String
    
    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
          
          If pval.ItemChanged = True Then
            '사원명
            If pval.ItemUID = "Visitant" Or pval.ItemUID = "BirthDat" Then
                
                BPLId = oForm01.Items("BPLId").Specific.VALUE
                Visitant = oForm01.Items("Visitant").Specific.VALUE
                BirthDat = Format(oForm01.Items("BirthDat").Specific.VALUE, "YYYYMMDD")
                
                oGrid01.DataTable.Clear
    
                oQuery = " EXEC [PS_GA920_04]  '" & BPLId & "', '" & Visitant & "', '" & BirthDat & "'"
    
                oGrid01.DataTable.ExecuteQuery (oQuery)
               
            End If
            
          End If
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
        Call PS_GA920_FormItemEnabled
        '//Call PS_GA920_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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
    '        Call oDS_PS_GA920L.RemoveRecord(oDS_PS_GA920L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_GA920_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_GA920_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_GA920_CreateItems() As Boolean
On Error GoTo PS_GA920_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    Set oGrid02 = oForm01.Items("Grid02").Specific
    Set oGrid03 = oForm01.Items("Grid03").Specific
    'oGrid01.SelectionMode = ms_NotSupported
    oForm01.DataSources.DataTables.Add ("ZTEMP1")
    oForm01.DataSources.DataTables.Add ("ZTEMP2")
    oForm01.DataSources.DataTables.Add ("ZTEMP3")
    
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP1")
    oGrid02.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP2")
    oGrid03.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP3")
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 10)
    oForm01.Items("BPLId").Specific.DataBind.SetBound True, "", "BPLId"
    
    '접수일자
    Call oForm01.DataSources.UserDataSources.Add("DocDate", dt_DATE)
    Call oForm01.Items("DocDate").Specific.DataBind.SetBound(True, "", "DocDate")
    oForm01.DataSources.UserDataSources.Item("DocDate").VALUE = Format(Now, "YYYYMMDD")
    
    
    '관리순번
    Call oForm01.DataSources.UserDataSources.Add("LineNum", dt_SHORT_NUMBER, 10)
    oForm01.Items("LineNum").Specific.DataBind.SetBound True, "", "LineNum"
    
    '방문자명
    Call oForm01.DataSources.UserDataSources.Add("Visitant", dt_SHORT_TEXT, 10)
    oForm01.Items("Visitant").Specific.DataBind.SetBound True, "", "Visitant"
    
    '생년월일
    Call oForm01.DataSources.UserDataSources.Add("BirthDat", dt_DATE)
    Call oForm01.Items("BirthDat").Specific.DataBind.SetBound(True, "", "BirthDat")
    
    '사원명
    Call oForm01.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 10)
    oForm01.Items("FullName").Specific.DataBind.SetBound True, "", "FullName"
    
        
    '거래처명
    Call oForm01.DataSources.UserDataSources.Add("CustNm", dt_SHORT_TEXT, 50)
    oForm01.Items("CustNm").Specific.DataBind.SetBound True, "", "CustNm"
    
    '전화번호
    Call oForm01.DataSources.UserDataSources.Add("TelNo", dt_SHORT_TEXT, 20)
    oForm01.Items("TelNo").Specific.DataBind.SetBound True, "", "TelNo"
    
    
    '주소
    Call oForm01.DataSources.UserDataSources.Add("Addr", dt_SHORT_TEXT, 50)
    oForm01.Items("Addr").Specific.DataBind.SetBound True, "", "Addr"
    
    '차종
    Call oForm01.DataSources.UserDataSources.Add("CarDiv", dt_SHORT_TEXT, 10)
    oForm01.Items("CarDiv").Specific.DataBind.SetBound True, "", "CarDiv"
    
    '차량번호
    Call oForm01.DataSources.UserDataSources.Add("CarNo", dt_SHORT_TEXT, 10)
    oForm01.Items("CarNo").Specific.DataBind.SetBound True, "", "CarNo"
    
    '출입목적
    Call oForm01.DataSources.UserDataSources.Add("Purpose", dt_SHORT_TEXT, 50)
    oForm01.Items("Purpose").Specific.DataBind.SetBound True, "", "Purpose"
    
    
    
    '출입장소
    Call oForm01.DataSources.UserDataSources.Add("Location", dt_SHORT_TEXT, 10)
    oForm01.Items("Location").Specific.DataBind.SetBound True, "", "Location"
    
    '입문일자
    Call oForm01.DataSources.UserDataSources.Add("InDate", dt_DATE)
    Call oForm01.Items("InDate").Specific.DataBind.SetBound(True, "", "InDate")
    
    '입문시간
    Call oForm01.DataSources.UserDataSources.Add("InTime", dt_SHORT_NUMBER)
    Call oForm01.Items("InTime").Specific.DataBind.SetBound(True, "", "InTime")
    
    '입회자
    Call oForm01.DataSources.UserDataSources.Add("WPerson", dt_SHORT_TEXT, 10)
    oForm01.Items("WPerson").Specific.DataBind.SetBound True, "", "WPerson"
    
    '표찰번호
    Call oForm01.DataSources.UserDataSources.Add("IDNum", dt_SHORT_TEXT, 20)
    oForm01.Items("IDNum").Specific.DataBind.SetBound True, "", "IDNum"
    
    '비고
    Call oForm01.DataSources.UserDataSources.Add("Remark", dt_SHORT_TEXT, 50)
    oForm01.Items("Remark").Specific.DataBind.SetBound True, "", "Remark"
    
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function

PS_GA920_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA920_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_GA920_ComboBox_Setting()
On Error GoTo PS_GA920_ComboBox_Setting_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '사업장콤보박스세팅
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//차종
    oForm01.Items("CarDiv").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CarDiv").Specific, "SELECT b.U_Code, b.U_CodeNm FROM [@PS_GA050L] b Where b.Code = 'G002' and b.U_UseYN = 'Y'", "", False, False)
    
    '출입목적
    Call oForm01.Items("Purpose").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT      U_Code AS [Code],"
    sQry = sQry & "                U_CodeNm As Name"
    sQry = sQry & " FROM       [@PS_GA050L]"
    sQry = sQry & " WHERE      Code = 'G004'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Purpose").Specific, sQry, "1", False, False)
    Call oForm01.Items("Purpose").Specific.Select("%", psk_ByValue)
    
    
    '장소(Location)
    Call oForm01.Items("Location").Specific.ValidValues.Add("%", "선택")
    sQry = "           SELECT      U_Code AS [Code],"
    sQry = sQry & "                U_CodeNm As Name"
    sQry = sQry & " FROM       [@PS_GA050L]"
    sQry = sQry & " WHERE      Code = 'G003'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Location").Specific, sQry, "1", False, False)
    Call oForm01.Items("Location").Specific.Select("%", psk_ByValue)
    
    
    Call oForm01.Freeze(False)
    
    
    Exit Sub
PS_GA920_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA920_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA920_CF_ChooseFromList()
On Error GoTo PS_GA920_CF_ChooseFromList_Error
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
PS_GA920_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA920_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA920_FormItemEnabled()
On Error GoTo PS_GA920_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseItmBsort01 = "" Then
        
        Else
            Call oForm01.Items("ItmBsort").Specific.Select(oBaseItmBsort01, psk_ByValue)
            oForm01.Items("ItmBsort").Enabled = False
        End If
        '//각모드에따른 아이템설정
        '//Call PS_GA920_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA920_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA920_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA920_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA920_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_GA920L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_GA920L.Offset = oRow
'    oDS_PS_GA920L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA920_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA920_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA920_FormClear()
On Error GoTo PS_GA920_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_GA920'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_GA920_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA920_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA920_EnableMenus()
On Error GoTo PS_GA920_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_GA920_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA920_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_GA920_DataValidCheck() As Boolean
On Error GoTo PS_GA920_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_GA920_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_GA920_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_GA920_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_GA920_FormClear
    Exit Function
PS_GA920_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA920_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_GA920_MTX01()
On Error GoTo PS_GA920_MTX01_Error
    '//그리드에 데이터 로드 (조회)
    Call oForm01.Freeze(True)
    Dim i           As Long
    Dim sQry        As String
    Dim sQry1        As String
    Dim BPLId        As String
    Dim DocDate       As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    BPLId = oForm01.Items("BPLId").Specific.VALUE
    DocDate = oForm01.Items("DocDate").Specific.VALUE
    
    oGrid01.DataTable.Clear
    oGrid02.DataTable.Clear
    oGrid03.DataTable.Clear
    
    sQry = " EXEC [PS_GA920_01]  '" & BPLId & "', '" & DocDate & "'"
    sQry1 = " EXEC [PS_GA920_02]  '" & BPLId & "', '" & DocDate & "'"
    
    
'    Call oForm01.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
    oGrid03.DataTable.ExecuteQuery (sQry)
    oGrid02.DataTable.ExecuteQuery (sQry1)
    
'    Call RecordSet01.DoQuery(sQry)
'    If RecordSet01.RecordCount = 0 Then
'       Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
'       GoTo PS_GA920_MTX01_Exit
'    End If
        
        
    oGrid02.AutoResizeColumns
    oGrid03.AutoResizeColumns
    oForm01.Update
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA920_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA920_MTX01_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA920_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PS_GA920_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    
    '//그리드 자료를 head에 로드
  
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_GA920_MTX02_Error

    Call oForm01.Freeze(True)
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
'    Param01 = oGrid01.DataTable.Columns("사업장").Cells(oRow).VALUE
'    Param02 = oGrid01.DataTable.Columns("일자").Cells(oRow).VALUE
    
    
    
'    sQry = "EXEC PS_GA920_02 '" & Param01 & "', '" & Param02 & "'"
'
'    oGrid02.DataTable.Clear
'
''    Call oForm01.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry)
''    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
'
'    oGrid02.DataTable.ExecuteQuery (sQry)
'
'    Call oRecordSet.DoQuery(sQry)
'    If (oRecordSet.RecordCount = 0) Then
'        Call PS_GA920_Form_ini
'        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
'        GoTo PS_GA920_MTX02_Exit
'    Else
        Call PS_GA920_Form_ini
        
        If oUID = "Grid01" Then
            oForm01.DataSources.UserDataSources.Item("DocDate").VALUE = Format(Now, "YYYYMMDD")
            oForm01.DataSources.UserDataSources.Item("Visitant").VALUE = oGrid01.DataTable.Columns("성명").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("BirthDat").VALUE = oGrid01.DataTable.Columns("생년월일").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("FullName").VALUE = ""
            oForm01.DataSources.UserDataSources.Item("CustNm").VALUE = oGrid01.DataTable.Columns("상호").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("TelNo").VALUE = oGrid01.DataTable.Columns("전화번호").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("CarNo").VALUE = oGrid01.DataTable.Columns("차량번호").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("Addr").VALUE = oGrid01.DataTable.Columns("주소").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("Remark").VALUE = oGrid01.DataTable.Columns("비고").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("InDate").VALUE = Format(Now, "YYYYMMDD")
            oForm01.DataSources.UserDataSources.Item("InTime").VALUE = Format(Now, "HHMM")
            
            Set oCombo = oForm01.Items("CarDiv").Specific
            oCombo.Select oGrid01.DataTable.Columns("차종").Cells(oRow).VALUE, psk_ByValue
            
        ElseIf oUID = "Grid02" Then
            
            oForm01.DataSources.UserDataSources.Item("DocDate").VALUE = Format(Now, "YYYYMMDD")
            oForm01.DataSources.UserDataSources.Item("Visitant").VALUE = oGrid02.DataTable.Columns("방문자성명").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("BirthDat").VALUE = oGrid02.DataTable.Columns("생년월일").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("FullName").VALUE = oGrid02.DataTable.Columns("담당자").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("CustNm").VALUE = oGrid02.DataTable.Columns("업체명").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("Addr").VALUE = oGrid02.DataTable.Columns("주소").Cells(oRow).VALUE
            
            oForm01.DataSources.UserDataSources.Item("CarNo").VALUE = oGrid02.DataTable.Columns("차량번호").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("Purpose").VALUE = oGrid02.DataTable.Columns("출입목적").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("Remark").VALUE = oGrid02.DataTable.Columns("비고").Cells(oRow).VALUE
            oForm01.DataSources.UserDataSources.Item("InDate").VALUE = Format(Now, "YYYYMMDD")
            oForm01.DataSources.UserDataSources.Item("InTime").VALUE = Format(Now, "HHMM")
            
            
            Set oCombo = oForm01.Items("BPLId").Specific
            oCombo.Select oGrid02.DataTable.Columns("사업장").Cells(oRow).VALUE, psk_ByValue
            
            
            Set oCombo = oForm01.Items("Location").Specific
            oCombo.Select oGrid02.DataTable.Columns("방문구역").Cells(oRow).VALUE, psk_ByValue
            
            
            Set oCombo = oForm01.Items("CarDiv").Specific
            oCombo.Select oGrid02.DataTable.Columns("차종").Cells(oRow).VALUE, psk_ByValue
            
            Param01 = oForm01.Items("BPLId").Specific.VALUE
            Param02 = oGrid02.DataTable.Columns("방문자성명").Cells(oRow).VALUE
            Param03 = Format(oGrid02.DataTable.Columns("생년월일").Cells(oRow).VALUE, "YYYYMMDD")
            
            sQry = " EXEC [PS_GA920_04]  '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"
            
            oGrid01.DataTable.ExecuteQuery (sQry)
            
            
            
            
        End If
'    End If
    
    oForm01.Update
    
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA920_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_GA920_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA920_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA920_MTX03(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    
    '//그리드 자료를 head에 로드
  
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_GA920_MTX03_Error

    Call oForm01.Freeze(True)
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    Param01 = oGrid03.DataTable.Columns("사업장").Cells(oRow).VALUE
    Param02 = Format(oGrid03.DataTable.Columns("일자").Cells(oRow).VALUE, "YYYYMMDD")
    Param03 = oGrid03.DataTable.Columns("순번").Cells(oRow).VALUE
    
    sQry = "EXEC PS_GA920_03 '" & Param01 & "','" & Param02 & "', '" & Param03 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        Call PS_GA920_Form_ini
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_GA920_MTX03_Exit
    End If
    
    oForm01.DataSources.UserDataSources.Item("DocDate").VALUE = oRecordSet.Fields("DocDate").VALUE
    oForm01.DataSources.UserDataSources.Item("LineNum").VALUE = oRecordSet.Fields("LineNum").VALUE
    oForm01.DataSources.UserDataSources.Item("Visitant").VALUE = oRecordSet.Fields("Visitant").VALUE
    oForm01.DataSources.UserDataSources.Item("BirthDat").VALUE = oRecordSet.Fields("BirthDat").VALUE
    oForm01.DataSources.UserDataSources.Item("FullName").VALUE = oRecordSet.Fields("FullName").VALUE
    oForm01.DataSources.UserDataSources.Item("CustNm").VALUE = oRecordSet.Fields("CustNm").VALUE
    oForm01.DataSources.UserDataSources.Item("TelNo").VALUE = oRecordSet.Fields("TelNo").VALUE
    oForm01.DataSources.UserDataSources.Item("Addr").VALUE = oRecordSet.Fields("Addr").VALUE
    oForm01.DataSources.UserDataSources.Item("CarDiv").VALUE = oRecordSet.Fields("CarDiv").VALUE
    oForm01.DataSources.UserDataSources.Item("CarNo").VALUE = oRecordSet.Fields("CarNo").VALUE
    oForm01.DataSources.UserDataSources.Item("InDate").VALUE = oRecordSet.Fields("InDate").VALUE
    oForm01.DataSources.UserDataSources.Item("InTime").VALUE = oRecordSet.Fields("InTime").VALUE
    oForm01.DataSources.UserDataSources.Item("WPerson").VALUE = oRecordSet.Fields("WPerson").VALUE
    oForm01.DataSources.UserDataSources.Item("IDNum").VALUE = oRecordSet.Fields("IDNum").VALUE
    oForm01.DataSources.UserDataSources.Item("Purpose").VALUE = oRecordSet.Fields("Purpose").VALUE
    oForm01.DataSources.UserDataSources.Item("Remark").VALUE = oRecordSet.Fields("Remark").VALUE
    
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select oRecordSet.Fields("BPLId").VALUE, psk_ByValue
    Set oCombo = oForm01.Items("Location").Specific
    oCombo.Select oRecordSet.Fields("Location").VALUE, psk_ByValue
    
    sQry = "Select '성명' = Visitant, '생년월일' = BirthDat, '상호' = CustNm, '전화번호' = TelNo, '주소' = Addr, '비고' = remark "
    sQry = sQry & " From [Z_PS_GA901H] Where BPLId = '" & Param01 & "'"
    sQry = sQry & " And Visitant = '" & oRecordSet.Fields("Visitant").VALUE & "' And BirthDat = '" & oRecordSet.Fields("BirthDat").VALUE & "'"
    
    oGrid01.DataTable.ExecuteQuery (sQry)
    
    '//Key Disable
    oForm01.ActiveItem = "DocDate"
    
'    oForm01.Items("Div").Enabled = False
'    oForm01.Items("Grade").Enabled = False
'    oForm01.Items("Num").Enabled = False
'
    oForm01.Update
    
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA920_MTX03_Exit:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_GA920_MTX03_Error:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA920_MTX03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PS_GA920_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim Seqncom As Integer
    Dim sQry    As String
    Dim ym      As String
    Dim BPLId$, DocDate$, TelNo$, LineNum$, MSTCOD$, FULLNAME$, TeamCode$, RspCode$, CoRegNum$, CustNm$, Location$, Visitant$, BirthDat$
    Dim Addr$, Indate$, InTime$, WPerson$, IDNum$, Purpose$, Remark$, CarDiv$, CarNo$
    
    Dim TempForm01 As Variant
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    On Error GoTo PS_GA920_SAVE_Error
    
    BPLId = oForm01.Items("BPLId").Specific.VALUE
    DocDate = oForm01.Items("DocDate").Specific.VALUE
    LineNum = oForm01.Items("LineNum").Specific.VALUE
    Visitant = oForm01.Items("Visitant").Specific.VALUE
    BirthDat = oForm01.Items("BirthDat").Specific.VALUE
    CustNm = oForm01.Items("CustNm").Specific.VALUE
    TelNo = oForm01.Items("TelNo").Specific.VALUE
    FULLNAME = oForm01.Items("FullName").Specific.VALUE
    Location = oForm01.Items("Location").Specific.VALUE
    CarDiv = oForm01.Items("CarDiv").Specific.VALUE
    CarNo = oForm01.Items("CarNo").Specific.VALUE
    Addr = oForm01.Items("Addr").Specific.VALUE
    Indate = oForm01.Items("InDate").Specific.VALUE
    InTime = oForm01.Items("InTime").Specific.VALUE
    WPerson = oForm01.Items("WPerson").Specific.VALUE
    IDNum = oForm01.Items("IDNum").Specific.VALUE
    Purpose = oForm01.Items("Purpose").Specific.VALUE
    Remark = oForm01.Items("Remark").Specific.VALUE
    
    Call oForm01.Freeze(True)
    
    If Trim(Visitant) = "" Then
        Call MDC_Com.MDC_GF_Message("방문자성명이 없습니다. 확인바랍니다..", "E")
        oForm01.ActiveItem = "Visitant"
        GoTo PS_GA920_SAVE_Exit
    End If
    
    If Trim(BirthDat) = "" Then
        Call MDC_Com.MDC_GF_Message("방문자 생년월일이 없습니다. 확인바랍니다..", "E")
        oForm01.ActiveItem = "BirthDat"
        GoTo PS_GA920_SAVE_Exit
    End If
    
    sQry = " Select Count(*) From [Z_PS_GA901H] Where Visitant = '" & Visitant & "' and BirthDat = '" & BirthDat & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.Fields(0).VALUE <= 0 Then
        Set TempForm01 = New PS_GA901
        Call TempForm01.LoadForm(oForm01, Visitant, BirthDat, CustNm, TelNo, Addr)
        GoTo PS_GA920_SAVE_Exit
    Else
    
        If Trim(BPLId) = "" Then
            Call MDC_Com.MDC_GF_Message("사업장에러. 확인바랍니다..", "E")
            GoTo PS_GA920_SAVE_Exit
        End If
        
        If Trim(DocDate) = "" Then
            Call MDC_Com.MDC_GF_Message("등록일자에러. 확인바랍니다..", "E")
            oForm01.ActiveItem = "DocDate"
            GoTo PS_GA920_SAVE_Exit
        End If
        
        
        
        If Trim(Purpose) = "" Then
            Call MDC_Com.MDC_GF_Message("출입목적이 없습니다. 확인바랍니다..", "E")
            GoTo PS_GA920_SAVE_Exit
        End If
        
        If Trim(Indate) = "" Then
            Call MDC_Com.MDC_GF_Message("출입일자가 없습니다. 확인바랍니다..", "E")
            GoTo PS_GA920_SAVE_Exit
        End If
        
        
        If Trim(InTime) = "" Then
            Call MDC_Com.MDC_GF_Message("출입시간이 없습니다. 확인바랍니다..", "E")
            GoTo PS_GA920_SAVE_Exit
        End If
        
                
        If oForm01.Items("LineNum").Specific.VALUE > 0 Then
            
            '//갱신
            sQry = "Update [Z_PS_GA920H] set "
            sQry = sQry + "FullName = '" & FULLNAME & "',"
            sQry = sQry + "CustNm = '" & CustNm & "',"
            sQry = sQry + "Location = '" & Location & "',"
            sQry = sQry + "Visitant = '" & Visitant & "',"
            sQry = sQry + "BirthDat = '" & BirthDat & "',"
            sQry = sQry + "TelNo = '" & TelNo & "',"
            sQry = sQry + "Addr = '" & Addr & "',"
            sQry = sQry + "CarDiv = '" & CarDiv & "',"
            sQry = sQry + "CarNo = '" & CarNo & "',"
            sQry = sQry + "InDate = '" & Indate & "',"
            sQry = sQry + "InTime = '" & InTime & "',"
            sQry = sQry + "WPerson = '" & WPerson & "',"
            sQry = sQry + "IDNum = '" & IDNum & "',"
            sQry = sQry + "Purpose = '" & Purpose & "',"
            sQry = sQry + "Remark = '" & Remark & "'"
            
            sQry = sQry + " Where BPLId = '" & BPLId & "' And DocDate = '" & DocDate & "' And LineNum = '" & LineNum & "'"
            oRecordSet.DoQuery sQry
            
        Else
            '//신규
            '순번 계산
            If LineNum = 0 Then
                LineNum = 1
            Else
                sQry = " Select Max(LineNum) From [Z_PS_GA920H] Where BPLId = '" & BPLId & "' And DocDate = '" & DocDate & "'"
                oRecordSet.DoQuery sQry
                
                LineNum = oRecordSet.Fields(0).VALUE + 1
            End If
            
            '------------------------------------------------------
            
            sQry = "INSERT INTO [Z_PS_GA920H]"
            sQry = sQry & " ("
            sQry = sQry & "BPLId,"
            sQry = sQry & "DocDate,"
            sQry = sQry & "LineNum,"
            sQry = sQry & "Visitant,"
            sQry = sQry & "BirthDat,"
            sQry = sQry & "CustNm,"
            sQry = sQry & "TelNo,"
            sQry = sQry & "Addr,"
            sQry = sQry & "CarDiv,"
            sQry = sQry & "CarNo,"
            sQry = sQry & "Purpose,"
            sQry = sQry & "Location,"
            sQry = sQry & "FullName,"
            sQry = sQry & "InDate,"
            sQry = sQry & "InTime,"
            sQry = sQry & "WPerson,"
            sQry = sQry & "IDNum,"
            sQry = sQry & "Remark"
            sQry = sQry & " ) "
            sQry = sQry & "VALUES("
    
            sQry = sQry & "'" & BPLId & "',"
            sQry = sQry & "'" & DocDate & "',"
            sQry = sQry & LineNum & ","
            sQry = sQry & "'" & Visitant & "',"
            sQry = sQry & "'" & BirthDat & "',"
            sQry = sQry & "'" & CustNm & "',"
            sQry = sQry & "'" & TelNo & "',"
            sQry = sQry & "'" & Addr & "',"
            sQry = sQry & "'" & CarDiv & "',"
            sQry = sQry & "'" & CarNo & "',"
            sQry = sQry & "'" & Purpose & "',"
            sQry = sQry & "'" & Location & "',"
            sQry = sQry & "'" & FULLNAME & "',"
            sQry = sQry & "'" & Indate & "',"
            sQry = sQry & "'" & InTime & "',"
            sQry = sQry & "'" & WPerson & "',"
            sQry = sQry & "'" & IDNum & "',"
            sQry = sQry & "'" & Remark & "' )"
    
            oRecordSet.DoQuery sQry
        End If
    End If
    
    Call PS_GA920_FormItemEnabled
    
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)

    Call PS_GA920_MTX01
    Call PS_GA920_Form_ini
    
    oForm01.Items("DocDate").Specific.VALUE = DocDate
    Call PS_GA920_MTX01
    
    
    Exit Sub
    
PS_GA920_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_GA920_SAVE_Error:
    Call oForm01.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_GA920_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_GA920_Delete()
    '//선택된 자료 삭제
    Dim BPLId$, DocDate$, LineNum$
    Dim i As Integer
    Dim Cnt As Integer
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry    As String
    Dim oRecordSet As SAPbobsCOM.Recordset
        
    On Error GoTo PS_GA920_Delete_Error

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    BPLId = oForm01.Items("BPLId").Specific.VALUE
    DocDate = oForm01.Items("DocDate").Specific.VALUE
    LineNum = oForm01.Items("LineNum").Specific.VALUE
        
    sQry = " Select Count(*) From [Z_PS_GA920H] Where BPLId = '" & BPLId & "' And DocDate = '" & DocDate & "' And LineNum = '" & LineNum & "'"
    oRecordSet.DoQuery sQry
    
    Cnt = oRecordSet.Fields(0).VALUE
    If Cnt > 0 Then
        
        If Sbo_Application.MessageBox(" 선택한라인을 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
            sQry = "Delete From [Z_PS_GA920H] Where BPLId = '" & BPLId & "' And DocDate = '" & DocDate & "' And LineNum = '" & LineNum & "'"
            oRecordSet.DoQuery sQry
        End If
    Else
        Call MDC_Com.MDC_GF_Message("조회후 삭제하십시요.", "E")
        GoTo PS_GA920_Delete_Exit
    End If
    
    Call oForm01.Freeze(False)
        
    Call PS_GA920_MTX01
    Call PS_GA920_Form_ini
    
    Set oRecordSet = Nothing
    
    Exit Sub
PS_GA920_Delete_Exit:
    Set oRecordSet = Nothing
        
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA920_Delete_Error:
    Set oRecordSet = Nothing
   
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA920_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PS_GA920_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_GA920_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    Dim Qty As Integer
    Dim Calculate_Weight As Double
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId As String
    Dim TeamCode As String
    Dim RspCode As String
    
'    Select Case oUID

'        Case "BPLId"
'
'            'BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE)
'            BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
'
'            If oForm01.Items("TeamCode").Specific.ValidValues.Count > 0 Then
'                For i = oForm01.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
'                    Call oForm01.Items("TeamCode").Specific.ValidValues.Remove(i, psk_Index)
'                Next i
'            End If
'
'            '부서콤보세팅
'            Call oForm01.Items("TeamCode").Specific.ValidValues.Add("%", "전체")
'            sQry = "            SELECT      U_Code AS [Code],"
'            sQry = sQry & "                 U_CodeNm As [Name]"
'            sQry = sQry & "  FROM       [@PS_HR200L]"
'            sQry = sQry & "  WHERE      Code = '1'"
'            sQry = sQry & "                 AND U_UseYN = 'Y'"
'            sQry = sQry & "                 AND U_Char2 = '" & BPLId & "'"
'            sQry = sQry & "  ORDER BY  U_Seq"
'            Call MDC_SetMod.Set_ComboList(oForm01.Items("TeamCode").Specific, sQry, "", False, False)
'            Call oForm01.Items("TeamCode").Specific.Select(0, psk_Index)
'
'        Case "TeamCode"
'
'            'TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE)
'            TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE)
'
'            If oForm01.Items("RspCode").Specific.ValidValues.Count > 0 Then
'                For i = oForm01.Items("RspCode").Specific.ValidValues.Count - 1 To 0 Step -1
'                    Call oForm01.Items("RspCode").Specific.ValidValues.Remove(i, psk_Index)
'                Next i
'            End If
'
'            '담당콤보세팅
'            Call oForm01.Items("RspCode").Specific.ValidValues.Add("%", "전체")
'            sQry = "            SELECT      U_Code AS [Code],"
'            sQry = sQry & "                 U_CodeNm As [Name]"
'            sQry = sQry & "  FROM       [@PS_HR200L]"
'            sQry = sQry & "  WHERE      Code = '2'"
'            sQry = sQry & "                 AND U_UseYN = 'Y'"
'            sQry = sQry & "                 AND U_Char1 = '" & TeamCode & "'"
'            sQry = sQry & "  ORDER BY  U_Seq"
'            Call MDC_SetMod.Set_ComboList(oForm01.Items("RspCode").Specific, sQry, "", False, False)
'            Call oForm01.Items("RspCode").Specific.Select(0, psk_Index)
'
'
'    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_GA920_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("에러 메시지", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA920_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub
Private Sub PS_GA920_Form_ini()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim DocDate     As String
    
    DocDate = oForm01.Items("DocDate").Specific.VALUE
    
    '화면청소 ㅋ
    oForm01.DataSources.UserDataSources.Item("DocDate").VALUE = Format(Now, "YYYYMMDD")
    
    oForm01.DataSources.UserDataSources.Item("LineNum").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("FullName").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("CustNm").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("Location").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("Visitant").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("BirthDat").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("Addr").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("CarNo").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("InDate").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("InTime").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("WPerson").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("IDNum").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("Purpose").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("Remark").VALUE = ""
    
    Set oCombo = oForm01.Items("CarDiv").Specific
    oCombo.Select 0, psk_Index
    
    Set oCombo = oForm01.Items("Location").Specific
    oCombo.Select 0, psk_Index
    
    
    
    oForm01.ActiveItem = "DocDate"
    oForm01.Update

End Sub







