VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S720"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : S720.cls
'//  Module         : 출고
'//  Desc           :
'//  FormType       : 720
'//  Create Date    : 2010.08.22
'//  Modified Date  :
'//  Creator        :
'//  Copyright  (c) Morning Data
'//****************************************************************************

Option Explicit

Private oFormUniqueID01 As String
Private oForm01              As SAPbouiCOM.Form
Private oMat01              As SAPbouiCOM.Matrix
'Private oDS_OINV As SAPbouiCOM.DBDataSource    '등록헤더
'Private oDS_RDR1 As SAPbouiCOM.DBDataSource    '등록라인

Private oLastItemUID01          As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01      As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Public Sub LoadForm(ByVal oFormUniqueID As String)
    oFormUniqueID01 = oFormUniqueID
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    Set oMat01 = oForm01.Items("13").Specific
'    Set oDS_ORDR = oForm01.DataSources.DBDataSources("ORDR")
'    Set oDS_RDR1 = oForm01.DataSources.DBDataSources("RDR1")
    Call PS_S720_CreateItems
    Call PS_S720_FormItemEnabled
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
                Call PS_S720_FormItemEnabled
            Case "1282": '추가
                Call PS_S720_FormItemEnabled
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_S720_FormItemEnabled
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
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_S720_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_S720_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_S720_FormItemEnabled
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_S720_FormItemEnabled
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "13", "1") '//사용자값활성
        If (pval.ItemUID = "13") Then
            If (pval.ColUID = "1") Then '//품목코드 변경시
                If pval.CharPressed = 9 Then
                    Dim ChildForm01 As Variant
                    Set ChildForm01 = New PS_SM020
                    Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, oMat01.VisualRowCount, "")
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        End If
        
        If pval.CharPressed = 9 Then '//거래처코드
            If pval.ItemUID = "13" And pval.ColUID = "U_CardCode" Then
                If oMat01.Columns("U_CardCode").Cells(pval.Row).Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
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
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "13") Then '//매트릭스
                If pval.ColUID = "1" Then
                    '//품목코드
                    oMat01.Columns("15").Cells(pval.Row).Specific.VALUE = MDC_PS_Common.User_WhsCode("1") '//기본창고
                
                ElseIf pval.ColUID = "U_Qty" Then
                    If (Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) <= 0) Then
                        oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE = 0 '//수량
                        oMat01.Columns("9").Cells(pval.Row).Specific.VALUE = 1 '//중량
                    Else
                        Dim ItemCode01  As String
                        ItemCode01 = oMat01.Columns("1").Cells(pval.Row).Specific.VALUE
                        
                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then '//EA자체품
                            oMat01.Columns("9").Cells(pval.Row).Specific.VALUE = Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                            If Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01)) = 0 Then
                                oMat01.Columns("9").Cells(pval.Row).Specific.VALUE = 1
                            Else
                                oMat01.Columns("9").Cells(pval.Row).Specific.VALUE = Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01))
                            End If
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                            If (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) = 0 Then
                                oMat01.Columns("9").Cells(pval.Row).Specific.VALUE = 1
                            Else
                                oMat01.Columns("9").Cells(pval.Row).Specific.VALUE = (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE)
                            End If
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                            If Round(Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0) = 0 Then
                                oMat01.Columns("9").Cells(pval.Row).Specific.VALUE = 1
                            Else
                                oMat01.Columns("9").Cells(pval.Row).Specific.VALUE = Round(Val(oMat01.Columns("U_Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 2)
                            End If
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG입력
                        End If
                    End If

                End If
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemChanged = True Then
            If pval.ItemUID = "13" Then
                If (pval.ColUID = "U_CardCode") Then
                    oForm01.Freeze True
                    sQry = "select cardname from ocrd where cardcode = '" & Trim(oMat01.Columns("U_CardCode").Cells(pval.Row).Specific.VALUE) & "'"
                    oRecordSet.DoQuery sQry
                    oMat01.Columns("U_CardName").Editable = True
                    oMat01.Columns("U_CardName").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
                    oMat01.Columns("U_CardCode").Cells(pval.Row).Click ct_Regular
                    oMat01.Columns("U_CardName").Editable = False
                    oForm01.Freeze False
                    
                ElseIf pval.ColUID = "9" Then
                
                    '로그인한 사용자가 manager이거나 dev01일 경우만 총계를 계산하지 않고, AP송장의 총계를 조회(2012.01.05 송명규 추가)
                    If oForm01.Items("UserID").Specific.VALUE = "manager" Or oForm01.Items("UserID").Specific.VALUE = "dev01" Then

                        sQry = "EXEC PS_S720_01 '" & Trim(oMat01.Columns("1").Cells(pval.Row).Specific.VALUE) & "', '" & Trim(oForm01.Items("9").Specific.VALUE) & "'"
                        Call oRecordSet.DoQuery(sQry)
            
                        oMat01.Columns("14").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet.Fields("LineTotal").VALUE)
                                            
                    End If
                    '로그인한 사용자가 manager이거나 dev01일 경우만 총계를 계산하지 않고, AP송장의 총계를 조회(2012.01.05 송명규 추가)
                    
                End If
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_720H", "U_CardCode,U_CardName")
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
'            Call PS_720_AddMatrixRow(0)
'        Else
'            If Trim(oDS_720L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
'                Call PS_720_AddMatrixRow(oMat01.RowCount)
'            End If
'        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S720_CreateItems()
On Error GoTo PS_S720_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oNewITEM    As Object
    Dim oItem       As Object
    Dim oEdit01 As SAPbouiCOM.EditText
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

    '특정사용자만 출고 아이템 중량을 가져올 수 있도록 하기 위한 UserID TextBox 추가 (2011.10.01 송명규)
    Set oNewITEM = oForm01.Items.Add("UserID", it_EDIT)
    Set oItem = oForm01.Items("21")
    oNewITEM.Left = oItem.Left
    oNewITEM.Top = (oItem.Top + oItem.Height) + 1
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = oItem.Width
    Set oEdit01 = oNewITEM.Specific
    oNewITEM.Specific.VALUE = Sbo_Company.UserName '로그인한 사용자의 ID
    
    oForm01.Items("38").Click ct_Regular '증빙일로의 포커스 강제 이동
    
    oNewITEM.Visible = False 'UserID 텍스트박스 숨김
    '특정사용자만 출고 아이템 중량을 가져올 수 있도록 하기 위한 UserID TextBox 추가 (2011.10.01 송명규)
    
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_S720_CreateItems_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_S720_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S720_FormItemEnabled()
On Error GoTo PS_S720_FormItemEnabled_Error
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
PS_S720_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_S720_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_S720_Print_Report01()
On Error GoTo PS_S720_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    
    Call ConnectODBC
    DocNum = oForm01.Items("8").Specific.VALUE
    WinTitle = "[BEC_S720] FIRM OFFER"
    ReportName = "PS_S720_01.rpt"
    sQry = "EXEC PS_S720_01 '" & oForm01.Items("8").Specific.VALUE & "'"
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
PS_S720_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_S720_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_S720_DataValidCheck() As Boolean
On Error GoTo PS_S720_DataValidCheck_Error
    PS_S720_DataValidCheck = False
    Dim i As Long
    
    '마감상태 체크_S(2017.11.23 송명규 추가)
    If MDC_PS_Common.Check_Finish_Status(MDC_PS_Common.User_BPLId(), oForm01.Items("9").Specific.VALUE, oForm01.TypeEx) = False Then
        Call Sbo_Application.SetStatusBarMessage("마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기일을 확인하고, 회계부서로 문의하세요.", bmt_Short, True)
        PS_S720_DataValidCheck = False
        Exit Function
    End If
    '마감상태 체크_E(2017.11.23 송명규 추가)
    
'    If oForm01.Items("4").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "고객은 필수입니다.", bmt_Short, True
'        oForm01.Items("4").Click ct_Regular
'        PS_S720_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount <= 1 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_S720_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount - 1
'        If (oMat01.Columns("1").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("1").Cells(i).Click ct_Regular
'            PS_S720_DataValidCheck = False
'            Exit Function
'        End If
'    Next
    PS_S720_DataValidCheck = True
    Exit Function
PS_S720_DataValidCheck_Error:
    PS_S720_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_S720_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


