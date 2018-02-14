VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : S42.cls
'//  Module         : 배치번호선택
'//  Desc           :
'//  FormType       : 42
'//  Create Date    : 2010.08.22
'//  Modified Date  :
'//  Creator        :
'//  Copyright  (c) Morning Data
'//****************************************************************************

Option Explicit

Private oFormUniqueID01 As String
Private oForm01              As SAPbouiCOM.Form
Private oMat01              As SAPbouiCOM.Matrix
Private oMat02              As SAPbouiCOM.Matrix
Private oMat03              As SAPbouiCOM.Matrix
'Private oDS_ORDR As SAPbouiCOM.DBDataSource    '등록헤더
'Private oDS_RDR1 As SAPbouiCOM.DBDataSource    '등록라인

Private oLastItemUID01          As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01      As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMatTopRow01 As Long
Private oMatBottomLeftRow01 As Long
Private oMatBottomRightRow01 As Long

Public Sub LoadForm(ByVal oFormUniqueID As String)
    oFormUniqueID01 = oFormUniqueID
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    Set oMat01 = oForm01.Items("3").Specific
    Set oMat02 = oForm01.Items("4").Specific
    Set oMat03 = oForm01.Items("5").Specific
    oMatTopRow01 = 1
    oMatBottomLeftRow01 = 1
'    Set oDS_ORDR = oForm01.DataSources.DBDataSources("ORDR")
'    Set oDS_RDR1 = oForm01.DataSources.DBDataSources("RDR1")
    Call PS_S42_CreateItems
    Call PS_S42_FormItemEnabled
End Sub

Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Select Case pval.EventType
        Case et_FORM_LOAD:
            Call Raise_EVENT_FORM_LOAD(FormUID, pval, BubbleEvent)
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
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_S42_FormItemEnabled
            Case "1282": '추가
                Call PS_S42_FormItemEnabled
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_S42_FormItemEnabled
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
'    If pval.ItemUID = "Mat01" Then
'        If pval.Row > 0 Then
'            oLastItemUID01 = pval.ItemUID
'            oLastColUID01 = pval.ColUID
'            oLastColRow01 = pval.Row
'        End If
'    Else
'        oLastItemUID01 = pval.ItemUID
'        oLastColUID01 = ""
'        oLastColRow01 = 0
'    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_FORM_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_FORM_LOAD_Error
    If pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_FORM_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "48" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If (Val(oMat01.Columns("55").Cells(oMatTopRow01).Specific.Value) >= Val(oMat02.Columns("4").Cells(oMatBottomLeftRow01).Specific.Value)) Then
                Else
                    Call MDC_Com.MDC_GF_Message("선택수량이 필요총계를 초과하였습니다.", "W")
                    BubbleEvent = False
                    Exit Sub
                End If
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If (Val(oMat01.Columns("55").Cells(oMatTopRow01).Specific.Value) >= Val(oMat02.Columns("4").Cells(oMatBottomLeftRow01).Specific.Value)) Then
                Else
                    Call MDC_Com.MDC_GF_Message("선택수량이 필요총계를 초과하였습니다.", "W")
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_S42_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_S42_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "48" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_S42_FormItemEnabled
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_S42_FormItemEnabled
                End If
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "38", "U_SD030Num") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "38", "U_Unweight")
'        Call MDC_PS_Common.ActiveUserDefineValueAlways_Price(oForm01, pval, BubbleEvent, "38", "14")
'        Call MDC_PS_Common.ActiveUserDefineValueAlways_UnitWeight(oForm01, pval, BubbleEvent, "38", "11")
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
        If pval.ItemUID = "3" Then
            If pval.Row > 0 Then
                oMatTopRow01 = pval.Row
            End If
        End If
        If pval.ItemUID = "4" Then
            If pval.Row > 0 Then
                oMatBottomLeftRow01 = pval.Row
            End If
        End If
        If pval.ItemUID = "5" Then
            If pval.Row > 0 Then
                oMatBottomRightRow01 = pval.Row
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
            If (pval.ItemUID = "4") Then '//매트릭스
                If pval.ColUID = "4" Then
                    If (MDC_PS_Common.GetItem_ItmBsort(oMat01.Columns("1").Cells(oMatTopRow01).Specific.Value) = "104" _
                    Or MDC_PS_Common.GetItem_ItmBsort(oMat01.Columns("1").Cells(oMatTopRow01).Specific.Value) = "302") Then '//품목이 멀티일때
                        If (Val(oMat02.Columns("3").Cells(pval.Row).Specific.Value) <> Val(oMat02.Columns("4").Cells(pval.Row).Specific.Value)) Then
                            oMat02.Columns("4").Cells(pval.Row).Specific.Value = Val(oMat02.Columns("3").Cells(pval.Row).Specific.Value)
                        End If
                    ElseIf (MDC_PS_Common.GetItem_ItmBsort(oMat01.Columns("1").Cells(oMatTopRow01).Specific.Value) = "102") Then  '//품목이 부품일때,제품만해당
                        oMat02.Columns("4").Cells(pval.Row).Specific.Value = Round(oMat02.Columns("4").Cells(pval.Row).Specific.Value, 0)
                    End If
                End If
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
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "38" Then
            If pval.ColUID = "1" Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_42H", "U_CardCode,U_CardName")
'        End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
'    If pval.ItemUID = "Mat01" Then
'        If pval.Row > 0 Then
'            oLastItemUID01 = pval.ItemUID
'            oLastColUID01 = pval.ColUID
'            oLastColRow01 = pval.Row
'        End If
'    Else
'        oLastItemUID01 = pval.ItemUID
'        oLastColUID01 = ""
'        oLastColRow01 = 0
'    End If '
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
       
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
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
    '        For i = 1 To oMat01.VisualRowCount
    '            oMat01.Columns("COL01").Cells(i).Specific.Value = i
    '        Next i
    '        oMat01.FlushToDataSource
    '        Call oDS_ZYM30L.RemoveRecord(oDS_ZYM30L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_42_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_42L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_42_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S42_CreateItems()
On Error GoTo PS_S42_CreateItems_Error
    Call oForm01.Freeze(True)
'    Dim oNewITEM    As Object
'    Dim oItem       As Object
'    Dim oBTN As SAPbouiCOM.Button
'    'FIRM OFFER
'    Set oNewITEM = oForm01.Items.Add("Button01", it_BUTTON)
'    oNewITEM.AffectsFormMode = False
'
'    Set oItem = oForm01.Items("10000329")
'    oNewITEM.Left = oItem.Left
'    oNewITEM.Top = oItem.Top - 23
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = oItem.Width
'
'    Set oBTN = oNewITEM.Specific
'    oBTN.Caption = "FIRM OFFER"
'    oBTN.Type = bt_Caption
'    oForm01.Items("Button01").LinkTo = "10000329"
'    Set oItem = Nothing
'    Set oNewITEM = Nothing
    oMat01.Columns("0").Cells(1).Click ct_Regular
    If oMat02.VisualRowCount > 0 Then
        oMat02.Columns("0").Cells(1).Click ct_Regular
    End If
    oForm01.Items("16").Visible = False
    Call oForm01.Freeze(False)
    Exit Sub
PS_S42_CreateItems_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_S42_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S42_FormItemEnabled()
On Error GoTo PS_S42_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_S42_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_S42_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S42_Print_Report01()
On Error GoTo PS_S42_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    
    Call ConnectODBC
    DocNum = oForm01.Items("8").Specific.Value
    WinTitle = "[BEC_S42] FIRM OFFER"
    ReportName = "PS_S42_01.rpt"
    sQry = "EXEC PS_S42_01 '" & oForm01.Items("8").Specific.Value & "'"
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)

    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_S42_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_S42_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_S42_DataValidCheck() As Boolean
On Error GoTo PS_S42_DataValidCheck_Error
    PS_S42_DataValidCheck = True
    Dim i As Long
'    If oForm01.Items("4").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "고객은 필수입니다.", bmt_Short, True
'        oForm01.Items("4").Click ct_Regular
'        PS_S42_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount <= 1 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_S42_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount - 1
'        If (oMat01.Columns("1").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("1").Cells(i).Click ct_Regular
'            PS_S42_DataValidCheck = False
'            Exit Function
'        End If
'    Next
    Exit Function
PS_S42_DataValidCheck_Error:
    PS_S42_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_S42_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


