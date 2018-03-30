VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S142"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : S142.cls
'//  Module         : 구매오더
'//  Desc           :
'//  FormType       : 142
'//  Create Date    : 2010.08.22
'//  Modified Date  :
'//  Creator        :
'//  Copyright  (c) Morning Data
'//****************************************************************************

Option Explicit

Private oFormUniqueID01 As String
Private oForm01              As SAPbouiCOM.Form
Private oMat01              As SAPbouiCOM.Matrix
'Private oDS_OPOR As SAPbouiCOM.DBDataSource    '등록헤더
'Private oDS_POR1 As SAPbouiCOM.DBDataSource    '등록라인

Private oLastItemUID01          As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01      As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private oSetBackOrderFunction01 As Boolean

Public Sub LoadForm(ByVal oFormUniqueID As String)
    oFormUniqueID01 = oFormUniqueID
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    Set oMat01 = oForm01.Items("38").Specific
'    Set oDS_OPOR = oForm01.DataSources.DBDataSources("OPOR")
'    Set oDS_POR1 = oForm01.DataSources.DBDataSources("POR1")
    Call PS_S142_CreateItems
    Call PS_S142_FormItemEnabled
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
            If pval.BeforeAction = False Then
                If (oSetBackOrderFunction01 = True) Then
                    oSetBackOrderFunction01 = False
                    Call MDC_PS_Common.SBO_SetBackOrderFunction(oForm01)
                End If
            End If
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
                Call PS_S142_FormItemEnabled
            Case "1282": '추가
                Call PS_S142_FormItemEnabled
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_S142_FormItemEnabled
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
    If pval.ItemUID = "38" Then
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

Private Sub Raise_EVENT_FORM_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_FORM_LOAD_Error
    If pval.BeforeAction = False Then
        Call MDC_PS_Common.SBO_SetBackOrderFunction(oForm01)
    End If
    Exit Sub
Raise_EVENT_FORM_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_S142_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_S142_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If oForm01.Items("Combo01").Specific.Selected Is Nothing Then
                Else
                    Call PS_S142_Print_Report01
                End If
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_S142_FormItemEnabled
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_S142_FormItemEnabled
                End If
            End If
        End If
        If pval.ItemUID = "Btn01" Then
            Call PS_S142_Print_Report01
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
        If (pval.ItemUID = "38") Then
            If (pval.ColUID = "1") Then '//품목코드 변경시
                If pval.CharPressed = 9 Then
                    Dim TradeType As String
                    TradeType = Trim(oForm01.Items("TradeType").Specific.Selected.VALUE)
                    
                    Dim ChildForm01 As Variant
                    Set ChildForm01 = New PS_SM020
                    Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, oMat01.VisualRowCount, TradeType)
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        End If
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "38", "U_SD030Num") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "38", "U_Unweight")
'        Call MDC_PS_Common.ActiveUserDefineValueAlways_Price(oForm01, pval, BubbleEvent, "38", "14")
'        Call MDC_PS_Common.ActiveUserDefineValueAlways_UnitWeight(oForm01, pval, BubbleEvent, "38", "11")
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    If Err.Number = "-7008" Then
        MDC_Com.MDC_GF_Message "사용자정의필드가 활성화되어 있지 않습니다.", "W"
        BubbleEvent = False
        Exit Sub
    End If
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
            Call PS_S142_FormItemEnabled(True)
        End If
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
    
    ElseIf pval.BeforeAction = False Then
        If (pval.ItemUID = "10000330") Then
            If pval.ActionSuccess = True Then
                If oForm01.Mode = fm_ADD_MODE Then
                    oSetBackOrderFunction01 = True
                End If
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
            If (pval.ItemUID = "38") Then '//매트릭스
                Dim ItemCode01 As String
                ItemCode01 = oMat01.Columns("1").Cells(pval.Row).Specific.VALUE
                If (pval.ColUID = "U_Qty") Then '//수량필드 값변경시
                    If (Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) <= 0) Then
                        oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE = 0 '//수량
                        oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = 1 '//중량
                    Else
                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then '//EA자체품
                            oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                            If Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01)) = 0 Then
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = 1
                            Else
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01))
                            End If
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                            If (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) = 0 Then
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = 1
                            Else
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE)
                            End If
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                            If Round(Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0) = 0 Then
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = 1
                            Else
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = Round(Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0)
                            End If
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG입력
                        End If
                    End If
                ElseIf (pval.ColUID = "11") Then
                    If (Val(oMat01.Columns("11").Cells(pval.Row).Specific.VALUE) <= 0) Then
                        oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE = 0 '//수량
                        oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = 1 '//중량
                    Else
                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then '//EA자체품
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                            If (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) = 0 Then
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = 1
                            Else
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE)
                            End If
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                            If Round(Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0) = 0 Then
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = 1
                            Else
                                oMat01.Columns("11").Cells(pval.Row).Specific.VALUE = Round(Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0)
                            End If
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG입력
                        End If
                    End If
                ElseIf pval.ColUID = "1" Then
                    If oMat01.VisualRowCount > 1 Then
                        oForm01.Items("TradeType").Enabled = False
                    Else
                        oForm01.Items("TradeType").Enabled = True
                    End If
                End If
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If oLastItemUID01 = "4" Then
            oForm01.Items("Btn01").Enabled = True
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Call oForm01.Freeze(False)
    If Err.Number = "-7008 " Then
    Else
        Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
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
                If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [OITM] WHERE ItemCode = '" & oMat01.Columns("1").Cells(pval.Row).Specific.VALUE & "'", 0, 1) > 0 Then
                Else
                    BubbleEvent = False
                End If
            End If
        End If
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_SD380H", "U_CardCode,U_CardName")
'        End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "38" Then
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
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        If (oSetBackOrderFunction01 = True) Then
            oSetBackOrderFunction01 = False
            Call MDC_PS_Common.SBO_SetBackOrderFunction(oForm01)
        End If
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
            If oMat01.VisualRowCount > 1 Then
                oForm01.Items("TradeType").Enabled = False
            Else
                oForm01.Items("TradeType").Enabled = True
            End If
    '        For i = 1 To oMat01.VisualRowCount
    '            oMat01.Columns("COL01").Cells(i).Specific.Value = i
    '        Next i
    '        oMat01.FlushToDataSource
    '        Call oDS_ZYM30L.RemoveRecord(oDS_ZYM30L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_SD380_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SD380L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_SD380_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S142_CreateItems()
On Error GoTo PS_S142_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oNewITEM    As Object
    Dim oItem       As Object
    Dim oStatic01 As SAPbouiCOM.StaticText
    Dim oCombo01 As SAPbouiCOM.ComboBox
    
    Set oNewITEM = oForm01.Items.Add("TradeType", it_COMBO_BOX)
    Set oItem = oForm01.Items("2003")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    oNewITEM.DisplayDesc = True
    Set oCombo01 = oNewITEM.Specific
    Call oCombo01.DataBind.SetBound(True, "OPOR", "U_TradeType")
    Call oCombo01.ValidValues.Add("1", "일반")
    Call oCombo01.ValidValues.Add("2", "임가공")
    
    Set oNewITEM = oForm01.Items.Add("Static01", it_STATIC)
    Set oItem = oForm01.Items("2002")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    oNewITEM.LinkTo = "TradeType"
    Set oStatic01 = oNewITEM.Specific
    oStatic01.Caption = "거래형태"
    
    Set oItem = Nothing
    Set oNewITEM = Nothing
    
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oBTN As SAPbouiCOM.Button
    
    Set oNewITEM = oForm01.Items.Add("Button01", it_BUTTON)
    oNewITEM.AffectsFormMode = False

    Set oItem = oForm01.Items("10000329")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = oItem.Top - 23
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    oNewITEM.LinkTo = "Combo01"
    
    Set oBTN = oNewITEM.Specific
    oBTN.Caption = "출력"
    oBTN.Type = bt_Caption
    
    
    Set oNewITEM = oForm01.Items.Add("Combo01", it_COMBO_BOX)
    oNewITEM.AffectsFormMode = False

    Set oItem = oForm01.Items("10000330")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = oItem.Top - 21
    oNewITEM.Height = 16
    oNewITEM.Width = oItem.Width
    oNewITEM.LinkTo = "Button01"
    oNewITEM.DisplayDesc = True
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Combo01").Specific, "SELECT U_Minor,U_CdName FROM [@PS_SY001L] WHERE CODE = 'RPT01' AND U_RelCd = '142'", "01", False, False)
    
    Set oItem = Nothing
    Set oNewITEM = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_S142_CreateItems_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_S142_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S142_FormItemEnabled(Optional Status As Boolean)
On Error GoTo PS_S142_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("TradeType").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("TradeType").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("TradeType").Enabled = False
    End If
    If Status = True Then
        If oForm01.Mode = fm_UPDATE_MODE Then
            oForm01.Items("TradeType").Enabled = False
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_S142_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_S142_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S142_Print_Report01()
On Error GoTo PS_S142_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    
    Call ConnectODBC
    DocNum = oForm01.Items("8").Specific.VALUE
    WinTitle = "PS_S142_" & oForm01.Items("Combo01").Specific.Selected.VALUE & ""
    ReportName = "PS_S142_" & oForm01.Items("Combo01").Specific.Selected.VALUE & ".rpt"
    
    sQry = "EXEC PS_S142_" & oForm01.Items("Combo01").Specific.Selected.VALUE & " '" & oForm01.Items("8").Specific.VALUE & "'"
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
PS_S142_Print_Report01_Error:
    
End Sub



Function PS_S142_DataValidCheck() As Boolean
On Error GoTo PS_S142_DataValidCheck_Error
    PS_S142_DataValidCheck = True
    Dim i As Long
    If oForm01.Items("4").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "고객은 필수입니다.", bmt_Short, True
        oForm01.Items("4").Click ct_Regular
        PS_S142_DataValidCheck = False
        Exit Function
    ElseIf Trim(oForm01.Items("2001").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm01.Items("2001").Click ct_Regular
        PS_S142_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("TradeType").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "거래형태는 필수입니다.", bmt_Short, True
        oForm01.Items("TradeType").Click ct_Regular
        PS_S142_DataValidCheck = False
        Exit Function
    ElseIf Trim(oForm01.Items("2001").Specific.VALUE) <> "1" And oForm01.Items("TradeType").Specific.Selected.VALUE = "2" Then '//창원이 아닌데 임가공이 선택된경우 에러
        Sbo_Application.SetStatusBarMessage "창원사업장이 아닌경우 임가공거래가 불가능합니다.", bmt_Short, True
        PS_S142_DataValidCheck = False
        Exit Function
    End If
'    If oMat01.VisualRowCount <= 1 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_S142_DataValidCheck = False
'        Exit Function
'    End If
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("1").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
            oMat01.Columns("1").Cells(i).Click ct_Regular
            PS_S142_DataValidCheck = False
            Exit Function
        End If
'        If (Val(oMat01.Columns("U_Qty").Cells(i).Specific.Value) <= 0) Then
'            Sbo_Application.SetStatusBarMessage "수량은 필수입니다.", bmt_Short, True
'            oMat01.Columns("U_Qty").Cells(i).Click ct_Regular
'            PS_S142_DataValidCheck = False
'            Exit Function
'        End If
        If (Val(oMat01.Columns("11").Cells(i).Specific.VALUE) <= 0) Then
            Sbo_Application.SetStatusBarMessage "수량(중량)은 필수입니다.", bmt_Short, True
            oMat01.Columns("11").Cells(i).Click ct_Regular
            PS_S142_DataValidCheck = False
            Exit Function
        End If
        If (oMat01.Columns("14").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "단가는 필수입니다.", bmt_Short, True
            oMat01.Columns("14").Cells(i).Click ct_Regular
            PS_S142_DataValidCheck = False
            Exit Function
        End If
        If oForm01.Items("70").Specific.Selected.VALUE = "S" Or oForm01.Items("70").Specific.Selected.VALUE = "L" Then '//현지,시스템통화
            If "KRW" <> Right(oMat01.Columns("14").Cells(i).Specific.VALUE, 3) Then
                Sbo_Application.SetStatusBarMessage "헤더와 라인의 통화가 다릅니다.", bmt_Short, True
                PS_S142_DataValidCheck = False
                Exit Function
            End If
        End If
        If oForm01.Items("70").Specific.Selected.VALUE = "C" Then '//BP통화
            If oForm01.Items("63").Specific.VALUE <> Right(oMat01.Columns("14").Cells(i).Specific.VALUE, 3) Then '//DocCur 과 Price의 마지막3자리 비교
                Sbo_Application.SetStatusBarMessage "헤더와 라인의 통화가 다릅니다.", bmt_Short, True
                PS_S142_DataValidCheck = False
                Exit Function
            End If
        End If
        If oForm01.Items("TradeType").Specific.Selected.VALUE = "1" Then '//일반일때는
            If MDC_PS_Common.GetItem_TradeType(oMat01.Columns("1").Cells(i).Specific.VALUE) = "2" Then '//품목은 임가공
                Sbo_Application.SetStatusBarMessage "문서의 거래형태와 품목의 거래형태가 다릅니다.", bmt_Short, True
                oMat01.Columns("1").Cells(i).Click ct_Regular
                PS_S142_DataValidCheck = False
                Exit Function
            End If
        End If
        If oForm01.Items("TradeType").Specific.Selected.VALUE = "2" Then '//임가공일때는
            If MDC_PS_Common.GetItem_TradeType(oMat01.Columns("1").Cells(i).Specific.VALUE) = "1" Then '//품목은 일반
                Sbo_Application.SetStatusBarMessage "문서의 거래형태와 품목의 거래형태가 다릅니다.", bmt_Short, True
                oMat01.Columns("1").Cells(i).Click ct_Regular
                PS_S142_DataValidCheck = False
                Exit Function
            End If
        End If
    Next
    Exit Function
PS_S142_DataValidCheck_Error:
    PS_S142_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_S142_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

