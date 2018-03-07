VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S139"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : S139.cls
'//  Module         : 판매오더
'//  Desc           :
'//  FormType       : 139
'//  Create Date    : 2010.08.22
'//  Modified Date  :
'//  Creator        :
'//  Copyright  (c) Morning Data
'//****************************************************************************

Option Explicit

Private oFormUniqueID01 As String
Private oForm01         As SAPbouiCOM.Form
Private oMat01          As SAPbouiCOM.Matrix
'Private oDS_ORDR As SAPbouiCOM.DBDataSource    '등록헤더
'Private oDS_RDR1 As SAPbouiCOM.DBDataSource    '등록라인

Private oLastItemUID01  As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01   As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01   As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private oSetBackOrderFunction01 As Boolean

Public Sub LoadForm(ByVal oFormUniqueID As String)
    oFormUniqueID01 = oFormUniqueID
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    Set oMat01 = oForm01.Items("38").Specific
'    Set oDS_ORDR = oForm01.DataSources.DBDataSources("ORDR")
'    Set oDS_RDR1 = oForm01.DataSources.DBDataSources("RDR1")
    Call PS_S139_CreateItems
    Call PS_S139_FormItemEnabled
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
                If (PS_S139_Validate("취소") = False) Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1286": '닫기
            
                If (PS_S139_Validate("닫기") = False) Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "ORDR_Cancel":  '작번 닫기(종료)
            
                Call PS_S139_ORDR_Cancel(oLastColRow01)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_S139_FormItemEnabled
            Case "1282": '추가
                Call PS_S139_FormItemEnabled
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_S139_FormItemEnabled
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

        '작번 닫기(종료) 생성
        If pval.ItemUID = "38" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
            
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            Set oCreationPackage = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.uniqueID = "ORDR_Cancel"
            oCreationPackage.String = "작번 닫기(종료)"
            oCreationPackage.Enabled = True

            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(oCreationPackage)
            
        End If
        
    ElseIf pval.BeforeAction = False Then
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'                Call Sbo_Application.Menus.RemoveEx("MenuUID")
'        End If

        '작번 닫기(종료) 삭제
        If pval.ItemUID = "38" And pval.Row > 0 Then
            If oMat01.RowCount >= pval.Row Then
                Call Sbo_Application.Menus.RemoveEx("ORDR_Cancel")
            End If
        End If
        
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
                If PS_S139_DataValidCheck(oForm01.Mode) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_S139_DataValidCheck(oForm01.Mode) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_FIND_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If oForm01.Items("Combo01").Specific.Selected Is Nothing Then
                Else
                    Call PS_S139_Print_Report01
                End If
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_S139_FormItemEnabled
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_S139_FormItemEnabled
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
        If pval.ItemUID = "38" Then
            If (PS_S139_Validate("수정") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "DCardCod", "") '//사용자값활성
        If (pval.ItemUID = "38") Then
            If (pval.ColUID = "1") Then '//품목코드 변경시
                If pval.CharPressed = 9 Then
                    Dim TradeType As String
                    TradeType = oForm01.Items("TradeType").Specific.Selected.VALUE
                    
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
        If pval.ItemUID = "38" Then
            If (PS_S139_Validate("수정") = False) Then
                BubbleEvent = False
                GoTo Raise_EVENT_COMBO_SELECT_Exit
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        Sbo_Application.Forms.GetForm("-" & oForm01.Type, oForm01.TypeCount).Update
        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
            Call PS_S139_FormItemEnabled(True)
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_COMBO_SELECT_Exit:
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    If Err.Number = "-7008 " Then
    Else
        Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error
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
                    
                        If Trim(oForm01.Items("2001").Specific.VALUE) = "2" And oForm01.Items("81").Specific.VALUE <> "3" And oForm01.Items("81").Specific.VALUE <> "4" Then
                        
                            '동래공장이면서 문서상태가 취소,종료가 아닐경우에 작번을 변경하려고 할 때만 체크
                            '저장하려는 작번으로 등록된 작업지시가 있는지 체크_S(2012.04.09 송명규 수정)
                            Dim Query01 As String
                            Dim RecordSet01 As SAPbobsCOM.Recordset
                            Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
                            Query01 = "            SELECT   COUNT(*)"
                            Query01 = Query01 & "  FROM     [@PS_PP030H]"
                            Query01 = Query01 & "  WHERE    U_ItemCode = '" & oMat01.Columns("1").Cells(pval.Row).Specific.VALUE & "'"
                        
                            Call RecordSet01.DoQuery(Query01)
            
                            If RecordSet01.Fields(0).VALUE > 0 Then
            
                                Call MDC_Com.MDC_GF_Message("이미 작업지시가 등록된 작번입니다. 수정할 수 없습니다.", "W")
                                oMat01.Columns("1").Cells(pval.Row).Specific.VALUE = "" '품목코드 필드에 빈값 강제 입력
    
                            End If
                            '저장하려는 작번으로 등록된 작업지시가 있는지 체크_E(2012.04.09 송명규 수정)
                        
                            '//동래공장, 취소, 종료가 아닐때는 수정가능
                            oForm01.Items("TradeType").Enabled = True
                            
                        Else
                            oForm01.Items("TradeType").Enabled = False
                        End If
                    Else
                        oForm01.Items("TradeType").Enabled = True
                    End If
                End If
            ElseIf pval.ItemUID = "DCardCod" Then
                oForm01.Items("DCardNam").Specific.VALUE = MDC_PS_Common.GetValue("SELECT CardName FROM OCRD WHERE CardCode = '" & oForm01.Items("DCardCod").Specific.VALUE & "'", 0, 1)
            End If
        End If
        Sbo_Application.Forms.GetForm("-" & oForm01.Type, oForm01.TypeCount).Update
        
    ElseIf pval.BeforeAction = False Then
    
        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
            Call PS_S139_FormItemEnabled(True)
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
    End If '
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
            If (PS_S139_Validate("행삭제") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
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

Private Sub PS_S139_CreateItems()
On Error GoTo PS_S139_CreateItems_Error

    Call oForm01.Freeze(True)
    Dim oNewITEM  As Object
    Dim oItem     As Object
    Dim oStatic01 As SAPbouiCOM.StaticText
    Dim oCombo01  As SAPbouiCOM.ComboBox
    Dim oEdit01   As SAPbouiCOM.EditText
    
    '매트릭스 초기화
    Set oMat01 = oForm01.Items("38").Specific
    
    Set oNewITEM = oForm01.Items.Add("TradeType", it_COMBO_BOX)
    Set oItem = oForm01.Items("2003")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    oNewITEM.DisplayDesc = True
    Set oCombo01 = oNewITEM.Specific
    Call oCombo01.DataBind.SetBound(True, "ORDR", "U_TradeType")
    Call oCombo01.ValidValues.Add("1", "일반")
    Call oCombo01.ValidValues.Add("2", "임가공")
    Call oCombo01.ValidValues.Add("3", "선생산")
    
    Set oNewITEM = oForm01.Items.Add("Static01", it_STATIC)
    Set oItem = oForm01.Items("2002")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    oNewITEM.LinkTo = "TradeType"
    Set oStatic01 = oNewITEM.Specific
    oStatic01.Caption = "거래형태"
    
    Set oNewITEM = oForm01.Items.Add("DCardCod", it_EDIT)
    Set oItem = oForm01.Items("222")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    Set oEdit01 = oNewITEM.Specific
    Call oEdit01.DataBind.SetBound(True, "ORDR", "U_DCardCod")
    
    Set oNewITEM = oForm01.Items.Add("Static03", it_STATIC)
    Set oItem = oForm01.Items("230")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    oNewITEM.LinkTo = "DCardCod"
    Set oStatic01 = oNewITEM.Specific
    oStatic01.Caption = "납품처코드"
    
    Set oNewITEM = oForm01.Items.Add("DCardNam", it_EDIT)
    Set oItem = oForm01.Items("DCardCod")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    oNewITEM.Enabled = False
    Set oEdit01 = oNewITEM.Specific
    Call oEdit01.DataBind.SetBound(True, "ORDR", "U_DCardNam")
    
    Set oNewITEM = oForm01.Items.Add("Static04", it_STATIC)
    Set oItem = oForm01.Items("Static03")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    oNewITEM.LinkTo = "DCardNam"
    Set oStatic01 = oNewITEM.Specific
    oStatic01.Caption = "납품처명"
    
    Set oNewITEM = oForm01.Items.Add("LotNo", it_EDIT)
    Set oItem = oForm01.Items("DCardNam")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    Set oEdit01 = oNewITEM.Specific
    Call oEdit01.DataBind.SetBound(True, "ORDR", "U_LotNo")
    
    Set oNewITEM = oForm01.Items.Add("Static05", it_STATIC)
    Set oItem = oForm01.Items("Static04")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    oNewITEM.LinkTo = "LotNo"
    Set oStatic01 = oNewITEM.Specific
    oStatic01.Caption = "업체수주번호"
    
    Set oItem = Nothing
    Set oNewITEM = Nothing
    
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oBTN As SAPbouiCOM.Button
    'FIRM OFFER
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
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Combo01").Specific, "SELECT U_Minor,U_CdName FROM [@PS_SY001L] WHERE CODE = 'RPT01' AND U_RelCd = '139'", "01", False, False)
    
    Set oItem = Nothing
    Set oNewITEM = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_S139_CreateItems_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_S139_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S139_FormItemEnabled(Optional Status As Boolean)
On Error GoTo PS_S139_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("TradeType").Enabled = True
        oForm01.Items("DCardNam").Enabled = False
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("TradeType").Enabled = True
        oForm01.Items("DCardNam").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        If Trim(oForm01.Items("2001").Specific.VALUE) = "2" And oForm01.Items("81").Specific.VALUE <> "3" And oForm01.Items("81").Specific.VALUE <> "4" Then
            '//동래공장인 경우 수정가능토록 취소, 종료가 아닐경우
            oForm01.Items("TradeType").Enabled = True
        Else
            oForm01.Items("TradeType").Enabled = False
        End If
        oForm01.Items("DCardNam").Enabled = False
    End If
    If Status = True Then
        If oForm01.Mode = fm_UPDATE_MODE Then
            If Trim(oForm01.Items("2001").Specific.VALUE) = "2" And oForm01.Items("81").Specific.VALUE <> "3" And oForm01.Items("81").Specific.VALUE <> "4" Then
                oForm01.Items("TradeType").Enabled = True
            Else
                oForm01.Items("TradeType").Enabled = False
            End If
            oForm01.Items("DCardNam").Enabled = False
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_S139_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_S139_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S139_Print_Report01()
On Error GoTo PS_S139_Print_Report01_Error

    Dim DocNum      As String
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim sQry        As String
    
    Call ConnectODBC
    DocNum = oForm01.Items("8").Specific.VALUE
    WinTitle = "PS_S139_" & oForm01.Items("Combo01").Specific.Selected.VALUE & ""
    ReportName = "PS_S139_" & oForm01.Items("Combo01").Specific.Selected.VALUE & ".rpt"
    
    sQry = "EXEC PS_S139_" & oForm01.Items("Combo01").Specific.Selected.VALUE & " '" & oForm01.Items("8").Specific.VALUE & "'"
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
PS_S139_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_S139_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_S139_DataValidCheck(Optional FormMode&) As Boolean
On Error GoTo PS_S139_DataValidCheck_Error
    PS_S139_DataValidCheck = True
    Dim i, j As Long
    If oForm01.Items("4").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "고객은 필수입니다.", bmt_Short, True
        oForm01.Items("4").Click ct_Regular
        PS_S139_DataValidCheck = False
        Exit Function
    ElseIf Trim(oForm01.Items("2001").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm01.Items("2001").Click ct_Regular
        PS_S139_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("TradeType").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "거래형태는 필수입니다.", bmt_Short, True
        oForm01.Items("TradeType").Click ct_Regular
        PS_S139_DataValidCheck = False
        Exit Function
    ElseIf Trim(oForm01.Items("2001").Specific.VALUE) <> "1" And oForm01.Items("TradeType").Specific.Selected.VALUE = "2" Then '//창원이 아닌데 임가공이 선택된경우 에러
        Sbo_Application.SetStatusBarMessage "창원사업장이 아닌경우 임가공거래가 불가능합니다.", bmt_Short, True
        PS_S139_DataValidCheck = False
        Exit Function
    End If
    If oMat01.VisualRowCount <= 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_S139_DataValidCheck = False
        Exit Function
    End If
    
    Dim Query01     As String
    Dim Query02     As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CardCode As String
    Dim DocDate  As String
    Dim ItemCode As String
    Dim Text     As String
    Dim SAmt     As String
    Dim Amt      As Double
    Dim OverAmt  As Double
    
    CardCode = oForm01.Items("4").Specific.VALUE
    DocDate = oForm01.Items("10").Specific.VALUE
    
    Text = oForm01.Items("29").Specific.VALUE '수주총계
    
    SAmt = Trim(Replace(Text, "KRW", ""))
    
    Amt = Val(Replace(SAmt, ",", ""))
    
    ItemCode = oMat01.Columns("1").Cells(1).Specific.VALUE
    
    Query01 = "Select U_ItmBsort From OITM Where ItemCode = '" & ItemCode & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    If RecordSet01.Fields(0).VALUE = "111" Then '분말제품 일때
        Query02 = "EXEC [S139_hando] '" & CardCode & "', '" & DocDate & "'"
        Call RecordSet02.DoQuery(Query02)
        
        OverAmt = RecordSet02.Fields("OverAmt").VALUE
        
        If OverAmt - Amt < 0 Then
            Call MDC_Com.MDC_GF_Message("여신한도를 초과합니다. 확인바랍니다.", "W")
            PS_S139_DataValidCheck = False
            Exit Function
        End If
    End If
        
        
    
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("1").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
            oMat01.Columns("1").Cells(i).Click ct_Regular
            PS_S139_DataValidCheck = False
            Exit Function
        End If
'        If (Val(oMat01.Columns("U_Qty").Cells(i).Specific.Value) <= 0) Then
'            Sbo_Application.SetStatusBarMessage "수량은 필수입니다.", bmt_Short, True
'            oMat01.Columns("U_Qty").Cells(i).Click ct_Regular
'            PS_S139_DataValidCheck = False
'            Exit Function
'        End If
        If (Val(oMat01.Columns("11").Cells(i).Specific.VALUE) <= 0) Then
            Sbo_Application.SetStatusBarMessage "수량(중량)은 필수입니다.", bmt_Short, True
            oMat01.Columns("11").Cells(i).Click ct_Regular
            PS_S139_DataValidCheck = False
            Exit Function
        End If
        If (oMat01.Columns("14").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "단가는 필수입니다.", bmt_Short, True
            oMat01.Columns("14").Cells(i).Click ct_Regular
            PS_S139_DataValidCheck = False
            Exit Function
        End If
        If oForm01.Items("70").Specific.Selected.VALUE = "S" Or oForm01.Items("70").Specific.Selected.VALUE = "L" Then '//현지,시스템통화
            If "KRW" <> Right(oMat01.Columns("14").Cells(i).Specific.VALUE, 3) Then
                Sbo_Application.SetStatusBarMessage "헤더와 라인의 통화가 다릅니다.", bmt_Short, True
                PS_S139_DataValidCheck = False
                Exit Function
            End If
        End If
        If oForm01.Items("70").Specific.Selected.VALUE = "C" Then '//BP통화
            If oForm01.Items("63").Specific.VALUE <> Right(oMat01.Columns("14").Cells(i).Specific.VALUE, 3) Then '//DocCur 과 Price의 마지막3자리 비교
                Sbo_Application.SetStatusBarMessage "헤더와 라인의 통화가 다릅니다.", bmt_Short, True
                PS_S139_DataValidCheck = False
                Exit Function
            End If
        End If
        If oForm01.Items("TradeType").Specific.Selected.VALUE = "1" Then '//일반일때는
            If MDC_PS_Common.GetItem_TradeType(oMat01.Columns("1").Cells(i).Specific.VALUE) = "2" Then '//품목은 임가공
                Sbo_Application.SetStatusBarMessage "문서의 거래형태와 품목의 거래형태가 다릅니다.", bmt_Short, True
                oMat01.Columns("1").Cells(i).Click ct_Regular
                PS_S139_DataValidCheck = False
                Exit Function
            End If
        End If
        If oForm01.Items("TradeType").Specific.Selected.VALUE = "2" Then '//임가공일때는
            If MDC_PS_Common.GetItem_TradeType(oMat01.Columns("1").Cells(i).Specific.VALUE) = "1" Then '//품목은 일반
                Sbo_Application.SetStatusBarMessage "문서의 거래형태와 품목의 거래형태가 다릅니다.", bmt_Short, True
                oMat01.Columns("1").Cells(i).Click ct_Regular
                PS_S139_DataValidCheck = False
                Exit Function
            End If
        End If
    Next
    If FormMode = fm_UPDATE_MODE Then
        If (PS_S139_Validate("검사") = False) Then
            PS_S139_DataValidCheck = False
            Exit Function
        End If
    End If
    Exit Function
PS_S139_DataValidCheck_Error:
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_S139_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_S139_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Function PS_S139_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_S139_Validate_Error
    PS_S139_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If ValidateType = "검사" Then
        '//헤더
        '//출하요청등록된것이 존재하면
        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" & Val(oForm01.Items("8").Specific.VALUE) & "'", 0, 1)) > 0 Then
            If oForm01.Items("70").Specific.Selected.VALUE <> MDC_PS_Common.GetValue("SELECT CurSource FROM [ORDR] WHERE DocEntry ='" & Val(oForm01.Items("8").Specific.VALUE) & "'", 0, 1) Then
                '//폼의 현지통화와, DB의 현지통화가 다르면
                MDC_Com.MDC_GF_Message "출하,선출요청된 문서입니다. 통화가 변경되었습니다.", "W"
                PS_S139_Validate = False
                GoTo PS_S139_Validate_Exit
            End If
            If oForm01.Items("63").Specific.Selected.VALUE <> MDC_PS_Common.GetValue("SELECT DocCur FROM [ORDR] WHERE DocEntry ='" & Val(oForm01.Items("8").Specific.VALUE) & "'", 0, 1) Then
                '//폼의 현지통화와, DB의 현지통화가 다르면
                MDC_Com.MDC_GF_Message "출하,선출요청된 문서입니다. 통화가 변경되었습니다.", "W"
                PS_S139_Validate = False
                GoTo PS_S139_Validate_Exit
            End If
            If Val(oForm01.Items("2001").Specific.Selected.VALUE) <> Val(MDC_PS_Common.GetValue("SELECT BPLId FROM [ORDR] WHERE DocEntry ='" & Val(oForm01.Items("8").Specific.VALUE) & "'", 0, 1)) Then
                '//폼의 사업장과, DB의 사업장이 다르면
                MDC_Com.MDC_GF_Message "출하,선출요청된 문서입니다. 사업장이 변경되었습니다.", "W"
                PS_S139_Validate = False
                GoTo PS_S139_Validate_Exit
            End If
        End If
        '//라인
        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
        Dim Exist As Boolean
        Exist = False
        Query01 = "SELECT DocEntry,LineNum FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.VALUE & "'"
        RecordSet01.DoQuery Query01
        For i = 0 To RecordSet01.RecordCount - 1
            Exist = False
            For j = 1 To oMat01.RowCount - 1
                '//라인번호가 같고, 품목코드가 같으면 존재하는행 , LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
                If Val(RecordSet01.Fields(1).VALUE) = Val(oMat01.Columns("U_LineNum").Cells(j).Specific.VALUE) And oMat01.Columns("U_LineNum").Cells(j).Specific.VALUE <> "" Then
                    Exist = True
                End If
            Next
            If (Exist = False) Then '//삭제된 행중
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" & Val(RecordSet01.Fields(0).VALUE) & "' AND PS_SD030L.U_RDR1Num = '" & Val(RecordSet01.Fields(1).VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "삭제된행이 다른사용자에 의해 출하,선출요청되었습니다. 적용할 수 없습니다.", "W"
                    PS_S139_Validate = False
                    GoTo PS_S139_Validate_Exit
                End If
            End If
            RecordSet01.MoveNext
        Next
        '//수량가능성검사.
        For i = 1 To oMat01.VisualRowCount - 1
            If (oMat01.Columns("U_LineNum").Cells(i).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
            Else
                '//매트릭스에 입력된 수량과 DB상에 존재하는 수량의 값비교
                If Val(oMat01.Columns("11").Cells(i).Specific.VALUE) < Val(MDC_PS_Common.GetValue("SELECT SUM(U_Weight) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" & oForm01.Items("8").Specific.VALUE & "' AND PS_SD030L.U_RDR1Num = '" & oMat01.Columns("U_LineNum").Cells(i).Specific.VALUE & "'", 0, 1)) Then
                    MDC_Com.MDC_GF_Message i & "행의 수량이 출하요청,선출요청 수량보다 작습니다.", "W"
                    PS_S139_Validate = False
                    GoTo PS_S139_Validate_Exit
                End If
                '//이미,출하선출된 행이 있으면 값이 수정되어서는 안된다..
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" & Val(oForm01.Items("8").Specific.VALUE) & "' AND PS_SD030L.U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    Query01 = "SELECT " '//품목코드는 변경되면 안된다.
                    Query01 = Query01 & " ItemCode, "
                    Query01 = Query01 & " Price, "
                    Query01 = Query01 & " U_TrType "
                    Query01 = Query01 & " FROM [RDR1] RDR1 WHERE DocEntry = '" & oForm01.Items("8").Specific.VALUE & "' AND LineNum = '" & Val(oMat01.Columns("U_LineNum").Cells(i).Specific.VALUE) & "'"
                    RecordSet01.DoQuery Query01
                    If RecordSet01.Fields(0).VALUE = oMat01.Columns("1").Cells(i).Specific.VALUE And _
                    Val(RecordSet01.Fields(1).VALUE) = Val(Mid(oMat01.Columns("14").Cells(i).Specific.VALUE, 1, Len(oMat01.Columns("14").Cells(i).Specific.VALUE) - 4)) And _
                    RecordSet01.Fields(2).VALUE = oMat01.Columns("U_TrType").Cells(i).Specific.Selected.VALUE _
                    Then
                    Else
                        MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 수정할 수 없습니다.", "W"
                        PS_S139_Validate = False
                        GoTo PS_S139_Validate_Exit
                    End If
                End If
            
            End If
        Next
    ElseIf ValidateType = "수정" Then
'        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
'            Else
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" & Val(oForm01.Items("8").Specific.Value) & "' AND PS_SD030L.U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할수 없습니다.", "W"
'                    PS_S139_Validate = False
'                    GoTo PS_S139_Validate_Exit
'                End If
'            End If
'        End If

        

    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
            Else
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" & Val(oForm01.Items("8").Specific.VALUE) & "' AND PS_SD030L.U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할 수 없습니다.", "W"
                    PS_S139_Validate = False
                    GoTo PS_S139_Validate_Exit
                End If
                
                '작업지시가 존재하면 삭제 불가
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] WHERE Canceled = 'N' AND U_ItemCode = '" & oMat01.Columns("1").Cells(oLastColRow01).Specific.VALUE & "'", 0, 1)) > 0 Then
                    Call MDC_Com.MDC_GF_Message("해당 작번은 작업지시가 존재합니다. 삭제할 수 없습니다.", "W")
                    PS_S139_Validate = False
                    GoTo PS_S139_Validate_Exit
                End If
                
            End If
        End If
        
    ElseIf ValidateType = "취소" Then
    
        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.VALUE & "'"
        RecordSet01.DoQuery Query01
        For i = 0 To RecordSet01.RecordCount - 1
        
            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" & Val(RecordSet01.Fields(0).VALUE) & "' AND PS_SD030L.U_RDR1Num = '" & Val(RecordSet01.Fields(1).VALUE) & "'", 0, 1)) > 0 Then
                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할 수 없습니다.", "W"
                PS_S139_Validate = False
                GoTo PS_S139_Validate_Exit
            End If
            
'            '작업지시가 존재하면 취소 불가(작업지시등록 문서에 입력된 판매오더만 검사, 2011.12.26 송명규 추가)
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] WHERE Status = 'O' AND U_ItemCode = '" & oMat01.Columns("1").Cells(i + 1).Specific.VALUE & "' AND U_SjNum = '" & Val(RecordSet01.Fields(0).VALUE) & "' AND U_SjLine = '" & Val(RecordSet01.Fields(1).VALUE) & "'", 0, 1)) > 0 Then
'                Call MDC_Com.MDC_GF_Message(i + 1 & "행의 작번이 작업지시가 존재합니다. 문서를 취소할 수 없습니다.", "W")
'                PS_S139_Validate = False
'                GoTo PS_S139_Validate_Exit
'            End If
            
            RecordSet01.MoveNext
        Next
    
    ElseIf ValidateType = "닫기" Then
    
        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.VALUE & "'"
        RecordSet01.DoQuery Query01
        For i = 0 To RecordSet01.RecordCount - 1
        
            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Status = 'O' AND PS_SD030L.U_ORDRNum = '" & Val(RecordSet01.Fields(0).VALUE) & "' AND PS_SD030L.U_RDR1Num = '" & Val(RecordSet01.Fields(1).VALUE) & "'", 0, 1)) > 0 Then
                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할 수 없습니다.", "W"
                PS_S139_Validate = False
                GoTo PS_S139_Validate_Exit
            End If
            
            '작업지시가 존재하면 닫기 불가(작업지시등록 문서에 입력된 판매오더만 검사, 2011.12.26 송명규 추가)
            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] WHERE Status = 'O' AND U_ItemCode = '" & oMat01.Columns("1").Cells(i + 1).Specific.VALUE & "' AND U_SjNum = '" & Val(RecordSet01.Fields(0).VALUE) & "' AND U_SjLine = '" & Val(RecordSet01.Fields(1).VALUE) & "'", 0, 1)) > 0 Then
                Call MDC_Com.MDC_GF_Message(i + 1 & "행의 작번이 작업지시가 존재합니다. 문서를 닫기(종료) 처리할 수 없습니다.", "W")
                PS_S139_Validate = False
                GoTo PS_S139_Validate_Exit
            End If
            
            RecordSet01.MoveNext
        Next
    
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_S139_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_S139_Validate_Error:
    PS_S139_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_S139_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_S139_ORDR_Cancel(ByVal p_Row As Integer)
On Error GoTo PS_S139_ORDR_Cancel_Error

    Dim l_ErrNum     As Integer

    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim l_DocEntry As Integer '문서번호
    Dim l_CardCode As String '거래처코드
    Dim l_ItemCode As String '품목코드(작번)
    Dim l_WOEntry  As String '작업지시문서번호
    
    l_DocEntry = oForm01.Items("8").Specific.VALUE
    l_CardCode = oForm01.Items("4").Specific.VALUE
    l_ItemCode = oMat01.Columns("1").Cells(p_Row).Specific.VALUE
    
    '이미 등록된 작번 닫기(종료) 건인지 검사_S
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "        SELECT  'X' "
    sQry = sQry & " FROM    Z_ORDR_Cancel "
    sQry = sQry & " WHERE   DocEntry = " & l_DocEntry
    sQry = sQry & "         AND CardCode = '" & l_CardCode & "'"
    sQry = sQry & "         AND ItemCode = '" & l_ItemCode & "'"
    
    Call oRecordset02.DoQuery(sQry)
    
    If oRecordset02.RecordCount <> 0 Then
        l_ErrNum = 1
        GoTo PS_S139_ORDR_Cancel_Error
    End If
    '이미 등록된 작번 닫기(종료) 건인지 검사_E
    
    '작업지시가 존재하는지 검사_S
    Dim oRecordSet03 As SAPbobsCOM.Recordset
    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "        SELECT  DocEntry "
    sQry = sQry & " FROM    [@PS_PP030H] "
    sQry = sQry & " WHERE   U_ItemCode = '" & l_ItemCode & "'"
    sQry = sQry & "         AND U_SjNum = '" & l_DocEntry & "'"
    sQry = sQry & "         AND Status = 'O'" '닫기 되지 않은 작업지시가 존재하면
    sQry = sQry & "         AND U_OrdSub1 = '00'"
    sQry = sQry & "         AND U_OrdSub2 = '000'"
    
    Call oRecordSet03.DoQuery(sQry)
    
    If oRecordSet03.RecordCount <> 0 Then
        l_ErrNum = 2
        l_WOEntry = oRecordSet03.Fields(0).VALUE
        GoTo PS_S139_ORDR_Cancel_Error
    End If
    '작업지시가 존재하는지 검사_E
    
    If Sbo_Application.MessageBox("작번 : " & l_ItemCode & "를(을) 닫기(종료) 처리하시겠습니까?", "1", "예", "아니오") = "1" Then
        sQry = "INSERT INTO Z_ORDR_Cancel VALUES (" & l_DocEntry & ",'" & l_CardCode & "','" & l_ItemCode & "', GETDATE(), " & Sbo_Company.UserSignature & ")"
        Call oRecordSet01.DoQuery(sQry)
        MDC_Com.MDC_GF_Message "작번 닫기(종료) 처리가 완료되었습니다.", "S"
    Else
        
    End If

'    Call Sbo_Application.MessageBox("문서번호:" & l_DocEntry & ",거래처코드:" & l_CardCode & ",품목코드:" & l_ItemCode)

    Set oRecordSet01 = Nothing
    Set oRecordset02 = Nothing
    Set oRecordSet03 = Nothing
    Exit Sub
PS_S139_ORDR_Cancel_Error:
    Set oRecordSet01 = Nothing
    Set oRecordset02 = Nothing
    Set oRecordSet03 = Nothing
    If l_ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "이미 작번 닫기(종료) 처리가 등록된 작번입니다. 다시 확인 하십시오.", "E"
    ElseIf l_ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "작업지시가 존재하는 작번입니다. 작업지시를 먼저 닫기(종료)하십시오. 작업지시문서번호 : " & l_WOEntry, "E"
    Else
        Sbo_Application.SetStatusBarMessage "PS_S139_ORDR_Cancel_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub
