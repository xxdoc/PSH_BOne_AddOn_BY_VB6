VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SM030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_SM030.cls
'//  Module         : 재고관리 > 제품BOM 조회
'//  Desc           : 재고장
'//  FormType       : PS_SM030
'//  Create Date    : 2010.08.20
'//  Copyright  (c) Morning Data
'//********************************************************************************

'//사용방법
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_SM010
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)


Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix
Private oDS_PS_SM030H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_SM030L As SAPbouiCOM.DBDataSource '등록라인

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseItemCode01 As String

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMat01Row01 As Long
Private oMat02Row02 As Long

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
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oItemCode02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SM030.srf")
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
    
    oFormUniqueID01 = "PS_SM030_" & GetTotalFormsCount
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
    oBaseItemCode01 = oItemCode02
    
    Call PS_SM030_CreateItems
    Call PS_SM030_ComboBox_Setting
    Call PS_SM030_CF_ChooseFromList
    Call PS_SM030_FormItemEnabled
    Call PS_SM030_EnableMenus
    '//Call PS_SM030_FormClear '//UDO방식일때
    '//Call PS_SM030_AddMatrixRow(0, True) '//UDO방식일때
    
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
                '//Call PS_SM030_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_SM030_FormItemEnabled '//UDO방식
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
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_SM030_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_SM030_SetBaseForm '//부모폼에입력
                If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
                    Call oForm01.Close
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_SM030_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_SM030" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SM030_FormItemEnabled
'                    Call PS_SM030_FormClear '//UDO방식일때
'                    Call PS_SM030_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SM030_FormItemEnabled
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
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
        If pval.ItemUID = "ItmBsort" Then
            Dim i As Long
            For i = 0 To oForm01.Items("ItmMsort").Specific.ValidValues.Count - 1
                Call oForm01.Items("ItmMsort").Specific.ValidValues.Remove(0, psk_Index)
            Next
            Call oForm01.Items("ItmMsort").Specific.ValidValues.Add("선택", "선택")
            Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" & oForm01.Items("ItmBsort").Specific.Selected.VALUE & "' ORDER BY U_Code", "", False, False)
            If oForm01.Items("ItmMsort").Specific.ValidValues.Count > 0 Then
                Call oForm01.Items("ItmMsort").Specific.Select(0, psk_Index)
            End If
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
        If pval.ItemUID = "Mat01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row > 0 Then
                    Call oMat01.SelectRow(pval.Row, True, False)
                    oMat01Row01 = pval.Row
                    If GetItem_ManBtchNum(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) = "Y" Then '//배치를 사용하는품목
                        Call PS_SM030_MTX02
                    ElseIf GetItem_ManBtchNum(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) = "N" Then '//배치를 사용하지 않는품목
                        Call PS_SM030_MTX02
                    End If
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Mat02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row > 0 Then
                    Call oMat02.SelectRow(pval.Row, True, False)
                    oMat02Row02 = pval.Row
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Opt01" Then
            oForm01.Settings.MatrixUID = "Mat01"
            oForm01.Settings.Enabled = True
            oForm01.Settings.EnableRowFormat = True
        End If
        If pval.ItemUID = "Opt02" Then
            oForm01.Settings.MatrixUID = "Mat02"
            oForm01.Settings.Enabled = True
            oForm01.Settings.EnableRowFormat = True
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
        If pval.ItemUID = "Mat01" Then
            If pval.Row = 0 Then
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource
            End If
        End If
        If pval.ItemUID = "Mat02" Then
            If pval.Row = 0 Then
                oMat02.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat02.FlushToDataSource
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
        If pval.ItemUID = "Mat01" Then
            Dim TempForm01 As Variant
            If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.String <> "" Then
                Set TempForm01 = New PS_MM002
                Call TempForm01.LoadForm(MDC_PS_Common.User_BPLId, oMat01.Columns("ItemCode").Cells(pval.Row).Specific.String)
                BubbleEvent = False
            Else
                        
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    Dim SelQty As Double
    Dim rate As Double
    Dim i As Integer
                        
    Call oForm01.Freeze(True)
    Dim ItemCode01 As String
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If pval.ItemUID = "Mat01" Then
                If pval.ColUID = "SelQty" Then
                    '//기준수량 수정에따른 BOM수량 변동
                    If oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE <= 0 Then
                        oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE = 0
                        oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = 0
                        
                        For i = 1 To oMat02.RowCount
                            oMat02.Columns("StdQty").Cells(i).Specific.VALUE = oMat02.Columns("OStdQty").Cells(i).Specific.VALUE
                            oMat02.Columns("SelQty").Cells(i).Specific.VALUE = oMat02.Columns("OSelQty").Cells(i).Specific.VALUE
                            oMat02.Columns("SelWeight").Cells(i).Specific.VALUE = oMat02.Columns("OSelWeight").Cells(i).Specific.VALUE
                        Next
                    Else
'                        ItemCode01 = oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE
'                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then '//EA자체품
'                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
'                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
'                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01))
'                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
'                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
'                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
'                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Round(Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 2)
'                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG입력
'                        End If
                        
                        
                        SelQty = oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE
                        For i = 1 To oMat02.RowCount
                            If SelQty = 0 Then
                                rate = 1
                            Else
                                rate = SelQty / Val(oMat02.Columns("OStdQty").Cells(i).Specific.VALUE)
                            End If
                            
                            oMat02.Columns("StdQty").Cells(i).Specific.VALUE = SelQty
                            oMat02.Columns("SelQty").Cells(i).Specific.VALUE = Int((rate * oMat02.Columns("OSelQty").Cells(i).Specific.VALUE) * -1) * -1
                            'oMat02.Columns("SelWeight").Cells(i).Specific.VALUE = Round(rate * oMat02.Columns("OSelWeight").Cells(i).Specific.VALUE, 2)
                            oMat02.Columns("SelWeight").Cells(i).Specific.VALUE = Round(((Int((rate * oMat02.Columns("OSelQty").Cells(i).Specific.VALUE) * -1) * -1) / oMat02.Columns("OSelQty").Cells(i).Specific.VALUE) * oMat02.Columns("OSelWeight").Cells(i).Specific.VALUE, 2)
                        Next
                        
                    End If
                    oForm01.Update
'                ElseIf pval.ColUID = "SelWeight" Then
'                    If oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE <= 0 Then
'                        oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE = 0
'                        oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = 0
'                    Else
'                        ItemCode01 = oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE
'                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then '//EA자체품
'                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
'                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
'                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
'                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
'                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Round(Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 2)
'                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG입력
'                        End If
'                    End If
'                    oForm01.Update
                End If
            ElseIf pval.ItemUID = "Mat02" Then
            
            
                If pval.ColUID = "SelQty" Then
'                    If (MDC_PS_Common.GetItem_ManBtchNum(ItemCode01) = "Y") Then '//배치를 사용하는 품목의경우
                    If oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE <= 0 Then
                        oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE = 0
                        oMat02.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = 0
                    Else
                        ItemCode01 = oMat02.Columns("ItemCode").Cells(pval.Row).Specific.VALUE
                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then '//EA자체품
                            oMat02.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Val(oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                            oMat02.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Val(oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01))
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                            oMat02.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                          '  oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE = Round(Val(oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 2)
                            oMat02.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Round(Val(oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 2)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG입력
                        End If
                    End If
                    oForm01.Update
                End If
                If pval.ColUID = "StdQty" Then
'                    If (MDC_PS_Common.GetItem_ManBtchNum(ItemCode01) = "Y") Then '//배치를 사용하는 품목의경우
                    If oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE <= 0 Then
                        oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE = 0
                        oMat02.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = 0
                    Else
                        ItemCode01 = oMat02.Columns("ItemCode").Cells(pval.Row).Specific.VALUE
                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then '//EA자체품
                            oMat02.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Val(oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                            oMat02.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Val(oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01))
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                            oMat02.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                            oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE = Val(oMat02.Columns("StdQty").Cells(pval.Row).Specific.VALUE) / Val(oMat02.Columns("StdQty2").Cells(pval.Row).Specific.VALUE) * Val(oMat02.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG입력
                        End If
                    End If
                    oForm01.Update
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
        Call PS_SM030_FormItemEnabled
        '//Call PS_SM030_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
        oForm01.Items("Mat01").Top = 70
        oForm01.Items("Mat01").Height = (oForm01.Height / 2) - 70
        oForm01.Items("Mat01").Left = 7
        oForm01.Items("Mat01").Width = oForm01.Width - 21
        oForm01.Items("Mat02").Top = (oForm01.Height / 2) + 10
        oForm01.Items("Mat02").Height = (oForm01.Height / 2) - 75
        oForm01.Items("Mat02").Left = 7
        oForm01.Items("Mat02").Width = oForm01.Width - 21
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
'                oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'                oForm01.DataSources.UserDataSources("ItemName").Value = oDataTable01.Columns(1).Cells(0).Value
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
    '        Call oDS_PS_SM030L.RemoveRecord(oDS_PS_SM030L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_SM030_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_SM030_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_SM030_CreateItems() As Boolean
On Error GoTo PS_SM030_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_SM030H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_SM030L = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat02.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    oMat02.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 30)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    
    If oBaseItemCode01 <> "" Then
        oForm01.Items("ItemCode").Specific.VALUE = oBaseItemCode01
    End If
    
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    
    Call oForm01.DataSources.UserDataSources.Add("StockType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("StockType").Specific.DataBind.SetBound(True, "", "StockType")
    Call oForm01.DataSources.UserDataSources.Add("TradeType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("TradeType").Specific.DataBind.SetBound(True, "", "TradeType")
    
    Call oForm01.DataSources.UserDataSources.Add("ItemGpCd", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItemGpCd").Specific.DataBind.SetBound(True, "", "ItemGpCd")
    Call oForm01.DataSources.UserDataSources.Add("ItmBsort", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItmBsort").Specific.DataBind.SetBound(True, "", "ItmBsort")
    Call oForm01.DataSources.UserDataSources.Add("ItmMsort", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItmMsort").Specific.DataBind.SetBound(True, "", "ItmMsort")
    Call oForm01.DataSources.UserDataSources.Add("Size", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Size").Specific.DataBind.SetBound(True, "", "Size")
    Call oForm01.DataSources.UserDataSources.Add("ItemType", dt_SHORT_TEXT, 50)
    Call oForm01.Items("ItemType").Specific.DataBind.SetBound(True, "", "ItemType")
    Call oForm01.DataSources.UserDataSources.Add("Mark", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Mark").Specific.DataBind.SetBound(True, "", "Mark")
        
    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt01").Specific.DataBind.SetBound(True, "", "Opt01")
    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt02").Specific.DataBind.SetBound(True, "", "Opt02")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    
    Dim oChkBox As SAPbouiCOM.CheckBox
    
    Call oForm01.DataSources.UserDataSources.Add("Check01", dt_SHORT_TEXT, 1)
    Set oChkBox = oForm01.Items("ChkBox01").Specific
    oChkBox.ValOn = "Y"
    oChkBox.ValOff = "N"
    oChkBox.DataBind.SetBound True, , "Check01"
    oForm01.DataSources.UserDataSources("Check01").VALUE = "N" '미체크로 값을 주고 폼을 로드
    
    oForm01.Items("Mat01").Enabled = False
    oForm01.Items("Mat02").Enabled = False
        
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_SM030_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM030_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SM030_ComboBox_Setting()
On Error GoTo PS_SM030_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SM030", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SM030", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_SM030", "Mat01", "ItemCode")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM030", "StockType", "", "1", "재고있는품목")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM030", "StockType", "", "2", "전체")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("StockType").Specific, "PS_SM030", "StockType")
    Call oForm01.Items("StockType").Specific.Select("2", psk_ByValue)
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM030", "TradeType", "", "", "전체")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM030", "TradeType", "", "1", "일반")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SM030", "TradeType", "", "2", "임가공")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("TradeType").Specific, "PS_SM030", "TradeType")
    
    
    Call oForm01.Items("ItmBsort").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", False, False)
    Call oForm01.Items("ItmMsort").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", "", False, False)
    Call oForm01.Items("ItemType").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType").Specific, "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", "", False, False)
    Call oForm01.Items("Mark").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", "", False, False)
    Call oForm01.Items("ItemGpCd").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemGpCd").Specific, "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]", "", False, False)
    Call oForm01.Items("ItmBsort").Specific.Select(0, psk_Index)
    Call oForm01.Items("ItmMsort").Specific.Select(0, psk_Index)
    Call oForm01.Items("ItemType").Specific.Select(0, psk_Index)
    Call oForm01.Items("Mark").Specific.Select(0, psk_Index)
    Call oForm01.Items("ItemGpCd").Specific.Select(0, psk_Index)
    
    oForm01.Items("ItemCode").Click ct_Regular
    oForm01.Items("TradeType").Enabled = False
    
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM030_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM030_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SM030_CF_ChooseFromList()
On Error GoTo PS_SM030_CF_ChooseFromList_Error
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
PS_SM030_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM030_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SM030_FormItemEnabled()
On Error GoTo PS_SM030_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_SM030_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM030_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM030_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SM030_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SM030_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_SM030L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_SM030L.Offset = oRow
'    oDS_PS_SM030L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM030_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM030_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SM030_FormClear()
On Error GoTo PS_SM030_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SM030'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SM030_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM030_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SM030_EnableMenus()
On Error GoTo PS_SM030_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_SM030_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM030_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SM030_DataValidCheck() As Boolean
On Error GoTo PS_SM030_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_SM030_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_SM030_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_SM030_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_SM030_FormClear
    Exit Function
PS_SM030_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM030_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SM030_MTX01()
On Error GoTo PS_SM030_MTX01_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim Param08 As String
    Dim Param09 As String
    Dim Param10 As String
    Param01 = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    Param02 = Trim(oForm01.Items("StockType").Specific.Selected.VALUE)
    Param03 = Trim(oForm01.Items("TradeType").Specific.Selected.VALUE)
    Param04 = Trim(oForm01.Items("ItmBsort").Specific.Selected.VALUE)
    Param05 = Trim(oForm01.Items("ItmMsort").Specific.Selected.VALUE)
    Param06 = Trim(oForm01.Items("Size").Specific.VALUE)
    Param07 = Trim(oForm01.Items("ItemType").Specific.Selected.VALUE)
    Param08 = Trim(oForm01.Items("Mark").Specific.Selected.VALUE)
    Param09 = Trim(oForm01.Items("ItemName").Specific.VALUE)
    Param10 = Trim(oForm01.Items("ItemGpCd").Specific.Selected.VALUE)
    
    If oBaseForm01 Is Nothing Then
        Query01 = "EXEC PS_SM030_01 '" & Param01 & "','','','" & Param02 & "','','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "'"
    ElseIf oBaseForm01.Type = "149" Or oBaseForm01.Type = "139" Or oBaseForm01.Type = "140" Or oBaseForm01.Type = "180" Or oBaseForm01.Type = "133" Or oBaseForm01.Type = "179" Or oBaseForm01.Type = "60091" Then
        Query01 = "EXEC PS_SM030_01 '" & Param01 & "','Y','','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "'" '//판매Y,구매,재고타입(1:재고있는것만,2:전체),거래타입(1:일반,2:임가공)
    ElseIf oBaseForm01.Type = "142" Or oBaseForm01.Type = "143" Or oBaseForm01.Type = "182" Or oBaseForm01.Type = "141" Or oBaseForm01.Type = "181" Or oBaseForm01.Type = "60092" Then
        Query01 = "EXEC PS_SM030_01 '" & Param01 & "','','Y','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "'" '//판매,구매Y
    Else
        Query01 = "EXEC PS_SM030_01 '" & Param01 & "','','','" & Param02 & "','','" & Param04 & "','" & Param05 & "','" & Param06 & "','" & Param07 & "','" & Param08 & "','" & Param09 & "','" & Param10 & "'"
    End If
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        oForm01.Items("Mat01").Enabled = False
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
        GoTo PS_SM030_MTX01_Exit
    Else
        oForm01.Items("Mat01").Enabled = True
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_SM030H.InsertRecord (i)
        End If
        oDS_PS_SM030H.Offset = i
        Call oDS_PS_SM030H.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_SM030H.setValue("U_ColReg01", i, False)
        Call oDS_PS_SM030H.setValue("U_ColReg02", i, RecordSet01.Fields("ItemCode").VALUE)
        Call oDS_PS_SM030H.setValue("U_ColReg03", i, RecordSet01.Fields("ItemName").VALUE)
'        Call oDS_PS_SM030H.setValue("U_ColReg04", i, RecordSet01.Fields("CallSize").VALUE)
'        Call oDS_PS_SM030H.setValue("U_ColReg05", i, RecordSet01.Fields("Mark").VALUE)
        Call oDS_PS_SM030H.setValue("U_ColQty01", i, RecordSet01.Fields("OnHand").VALUE)
        Call oDS_PS_SM030H.setValue("U_ColQty02", i, RecordSet01.Fields("IsCommited").VALUE)
        Call oDS_PS_SM030H.setValue("U_ColQty03", i, RecordSet01.Fields("OnOrder").VALUE)
        Call oDS_PS_SM030H.setValue("U_ColQty04", i, RecordSet01.Fields("OnEnabled").VALUE)
        
        Call oDS_PS_SM030H.setValue("U_ColNum01", i, 0)
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM030_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM030_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM030_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SM030_MTX02()
On Error GoTo PS_SM030_MTX02_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    
    Dim SelQty As Double
    Dim rate As Double
    Dim StdQty As Double
    Dim StdQty2 As Double
    
    Param01 = oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.VALUE
    
    StdQty = oMat01.Columns("SelQty").Cells(oMat01Row01).Specific.VALUE
    
'    Param02 = oMat01.Columns("WhsCode").Cells(oMat01Row01).Specific.Value
'    Param03 = Trim(oForm01.Items("Param01").Specific.VALUE)
'    Param04 = Trim(oForm01.Items("Param01").Specific.VALUE)
    
    Query01 = "EXEC PS_SM030_02 '" & Param01 & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        oForm01.Items("Mat02").Enabled = False
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
        GoTo PS_SM030_MTX02_Exit
    Else
        oForm01.Items("Mat02").Enabled = True
    End If
    
    '//품목이 멀티이면 수량,중량필드 비활성화
    If MDC_PS_Common.GetItem_ItmBsort(Param01) = "104" _
    Or MDC_PS_Common.GetItem_ItmBsort(Param01) = "302" Then '//멀티
        oMat02.Columns("SelQty").Editable = False
        oMat02.Columns("SelWeight").Editable = False
    Else '//그외품목은 수량선택가능
        oMat02.Columns("SelQty").Editable = True
        oMat02.Columns("SelWeight").Editable = True
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_SM030L.InsertRecord (i)
        End If
        oDS_PS_SM030L.Offset = i
        Call oDS_PS_SM030L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_SM030L.setValue("U_ColReg01", i, "Y")
        Call oDS_PS_SM030L.setValue("U_ColReg02", i, RecordSet01.Fields("ItemCode").VALUE)
        Call oDS_PS_SM030L.setValue("U_ColReg03", i, RecordSet01.Fields("ItemName").VALUE)
        Call oDS_PS_SM030L.setValue("U_ColReg04", i, RecordSet01.Fields("Size").VALUE)
        Call oDS_PS_SM030L.setValue("U_ColReg05", i, RecordSet01.Fields("Ut").VALUE)
        Call oDS_PS_SM030L.setValue("U_ColReg06", i, RecordSet01.Fields("OutSize").VALUE)
        
        If StdQty <> 0 Then
            Call oDS_PS_SM030L.setValue("U_ColNum01", i, StdQty)
            rate = StdQty / RecordSet01.Fields("StdQty").VALUE
            
            Call oDS_PS_SM030L.setValue("U_ColNum02", i, Int((rate * RecordSet01.Fields("SelQty").VALUE) * -1) * -1)
            Call oDS_PS_SM030L.setValue("U_ColQty01", i, Round(((Int((rate * RecordSet01.Fields("SelQty").VALUE) * -1) * -1) / RecordSet01.Fields("SelQty").VALUE) * RecordSet01.Fields("SelWeight").VALUE, 2))
            
'            Call oDS_PS_SM030L.setValue("U_ColNum02", i, Round(RecordSet01.Fields("SelQty").VALUE * rate, 0))
'            Call oDS_PS_SM030L.setValue("U_ColQty01", i, Round(RecordSet01.Fields("SelWeight").VALUE * rate, 2))
        Else
           ' Call oDS_PS_SM030L.setValue("U_ColNum00", i, RecordSet01.Fields("StdQty2").VALUE)
            Call oDS_PS_SM030L.setValue("U_ColNum01", i, RecordSet01.Fields("StdQty").VALUE)
            Call oDS_PS_SM030L.setValue("U_ColNum02", i, RecordSet01.Fields("SelQty").VALUE)
            Call oDS_PS_SM030L.setValue("U_ColQty01", i, RecordSet01.Fields("SelWeight").VALUE)
        End If
       ' Call oDS_PS_SM030L.setValue("U_ColNum00", i, RecordSet01.Fields("StdQty2").VALUE)
        Call oDS_PS_SM030L.setValue("U_ColNum03", i, RecordSet01.Fields("StdQty").VALUE)
        Call oDS_PS_SM030L.setValue("U_ColNum04", i, RecordSet01.Fields("SelQty").VALUE)
        Call oDS_PS_SM030L.setValue("U_ColNum05", i, RecordSet01.Fields("StdQty").VALUE)
        Call oDS_PS_SM030L.setValue("U_ColQty02", i, RecordSet01.Fields("SelWeight").VALUE)
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM030_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SM030_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SM030_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_SM030_SetBaseForm()
On Error GoTo PS_SM030_SetBaseForm_Error
    Dim i As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    If oBaseForm01 Is Nothing Then
        '//DoNothing
    ElseIf oBaseForm01.Type = "133" Or oBaseForm01.Type = "139" Or oBaseForm01.Type = "140" Or oBaseForm01.Type = "141" Or oBaseForm01.Type = "142" Or oBaseForm01.Type = "143" Or oBaseForm01.Type = "149" Or oBaseForm01.Type = "179" Or oBaseForm01.Type = "180" Or oBaseForm01.Type = "181" Or oBaseForm01.Type = "182" Or oBaseForm01.Type = "60091" Or oBaseForm01.Type = "60092" Then '//호출한폼타입
        Set oBaseMat01 = oBaseForm01.Items("38").Specific
        For i = 1 To oMat01.RowCount
            If oMat01.Columns("CHK").Cells(i).Specific.Checked = True Then
                If (Val(oMat01.Columns("SelWeight").Cells(i).Specific.VALUE) <= 0) Then
                    '//중량이 선택되지 않은품목
                    ItemCode01 = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
                    oBaseMat01.Columns("1").Cells(oBaseColRow01).Specific.VALUE = ItemCode01 '//품목
                    oBaseMat01.Columns("1").Cells(oBaseColRow01 + 1).Click ct_Regular
                    oBaseColRow01 = oBaseColRow01 + 1
                Else
                    ItemCode01 = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
                    oBaseMat01.Columns("1").Cells(oBaseColRow01).Specific.VALUE = ItemCode01 '//품목
                    oBaseMat01.Columns("U_Qty").Cells(oBaseColRow01).Specific.VALUE = Val(oMat01.Columns("SelQty").Cells(i).Specific.VALUE) '//수량 '//수량을 변경하면 중량이 자동변경된다.
                    oBaseMat01.Columns("U_Unweight").Cells(oBaseColRow01).Specific.VALUE = Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) '//단중
                    oBaseMat01.Columns("11").Cells(oBaseColRow01).Specific.VALUE = Val(oMat01.Columns("SelWeight").Cells(i).Specific.VALUE) '//중량
                    oBaseMat01.Columns("14").Cells(oBaseColRow01).Specific.VALUE = MDC_PS_Common.GetValue("EXEC PS_SBO_GETPRICE '" & oBaseForm01.Items("4").Specific.VALUE & "','" & ItemCode01 & "'", 0, 1)
                    oBaseMat01.Columns("1").Cells(oBaseColRow01 + 1).Click ct_Regular
                    oBaseColRow01 = oBaseColRow01 + 1
                End If
            End If
        Next
        For i = 1 To oMat02.RowCount '//배치선택품목
            If oMat02.Columns("CHK").Cells(i).Specific.Checked = True Then
                If (Val(oMat02.Columns("SelWeight").Cells(i).Specific.VALUE) <= 0) Then
                    '//중량이 선택되지 않은품목
                    ItemCode01 = oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.VALUE
                    oBaseMat01.Columns("1").Cells(oBaseColRow01).Specific.VALUE = ItemCode01 '//품목
                    oBaseMat01.Columns("1").Cells(oBaseColRow01 + 1).Click ct_Regular
                    oBaseColRow01 = oBaseColRow01 + 1
                Else
                    ItemCode01 = oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.VALUE
                    oBaseMat01.Columns("1").Cells(oBaseColRow01).Specific.VALUE = ItemCode01 '//품목
                    oBaseMat01.Columns("U_Qty").Cells(oBaseColRow01).Specific.VALUE = Val(oMat02.Columns("SelQty").Cells(i).Specific.VALUE) '//수량 '//수량을 변경하면 중량이 자동변경된다.
                    oBaseMat01.Columns("U_Unweight").Cells(oBaseColRow01).Specific.VALUE = Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) '//단중
                    oBaseMat01.Columns("11").Cells(oBaseColRow01).Specific.VALUE = Val(oMat02.Columns("SelWeight").Cells(i).Specific.VALUE) '//중량
                    oBaseMat01.Columns("14").Cells(oBaseColRow01).Specific.VALUE = MDC_PS_Common.GetValue("EXEC PS_SBO_GETPRICE '" & oBaseForm01.Items("4").Specific.VALUE & "','" & ItemCode01 & "'", 0, 1)
'                    If oBaseForm01.Type = "140" Then
'                        oBaseMat01.Columns("U_BatchNum").Cells(oBaseColRow01).Specific.Value = oMat02.Columns("BatchNum").Cells(i).Specific.Value '//배치번호
'                    End If
                    oBaseMat01.Columns("1").Cells(oBaseColRow01 + 1).Click ct_Regular
                    oBaseColRow01 = oBaseColRow01 + 1
                End If
            End If
        Next
    ElseIf oBaseForm01.TypeEx = "720" Then
        Set oBaseMat01 = oBaseForm01.Items("13").Specific '//매트릭스
            For i = 1 To oMat01.RowCount '//품목선택품목
            If oMat01.Columns("CHK").Cells(i).Specific.Checked = True Then
                oBaseMat01.Columns("1").Cells(oBaseColRow01).Specific.VALUE = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE '//품목
                oBaseMat01.Columns("1").Cells(oBaseColRow01 + 1).Click ct_Regular
                oBaseColRow01 = oBaseColRow01 + 1
            End If
        Next
        For i = 1 To oMat02.RowCount '//배치선택품목
            If oMat02.Columns("CHK").Cells(i).Specific.Checked = True Then
                oBaseMat01.Columns("1").Cells(oBaseColRow01).Specific.VALUE = oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.VALUE '//품목
                oBaseMat01.Columns("1").Cells(oBaseColRow01 + 1).Click ct_Regular
                oBaseColRow01 = oBaseColRow01 + 1
            End If
        Next
    '//자재기타출고
    ElseIf oBaseForm01.TypeEx = "PS_MM090" Then
        Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//매트릭스

        For i = 1 To oMat02.RowCount '//배치선택품목
            If oMat02.Columns("CHK").Cells(i).Specific.Checked = True Then
                oBaseMat01.Columns("ItemName").Cells(oBaseColRow01).Specific.VALUE = oMat02.Columns("ItemName").Cells(i).Specific.VALUE '//품목
                oBaseMat01.Columns("Size").Cells(oBaseColRow01).Specific.VALUE = oMat02.Columns("Size").Cells(i).Specific.VALUE '//규격
                oBaseMat01.Columns("Unit").Cells(oBaseColRow01).Specific.VALUE = oMat02.Columns("Ut").Cells(i).Specific.VALUE '//단위
                oBaseMat01.Columns("Qty").Cells(oBaseColRow01).Specific.VALUE = Val(oMat02.Columns("SelQty").Cells(i).Specific.VALUE) '//수량
                oBaseMat01.Columns("Weight").Cells(oBaseColRow01).Specific.VALUE = (Val(oMat02.Columns("SelWeight").Cells(i).Specific.VALUE)) '//재적
                
                If i = 1 Then
                    oBaseForm01.Items("Comment1").Specific.VALUE = oMat02.Columns("OutSize").Cells(i).Specific.VALUE + " (" + oMat02.Columns("StdQty").Cells(i).Specific.VALUE + "EA)" '//수량
                End If

                oBaseColRow01 = oBaseColRow01 + 1
                
                oMat02.Columns("CHK").Cells(i).Specific.Checked = False
            End If
        Next
    
    '//자재청구등록
    ElseIf oBaseForm01.TypeEx = "PS_MM005" Then
        Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//매트릭스

        For i = 1 To oMat02.RowCount '//배치선택품목
            If oMat02.Columns("CHK").Cells(i).Specific.Checked = True Then
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oMat02.Columns("ItemCode").Cells(i).Specific.VALUE '//품목
                oBaseMat01.Columns("Qty").Cells(oBaseColRow01).Specific.VALUE = Val(oMat02.Columns("SelQty").Cells(i).Specific.VALUE) '//수량
                oBaseMat01.Columns("Weight").Cells(oBaseColRow01).Specific.VALUE = (Val(oMat02.Columns("SelWeight").Cells(i).Specific.VALUE)) '//재적
                
                If i = 1 Then
                    oBaseMat01.Columns("Comments").Cells(oBaseColRow01).Specific.VALUE = oMat02.Columns("OutSize").Cells(i).Specific.VALUE + " (" + oMat02.Columns("StdQty").Cells(i).Specific.VALUE + "EA)" '//수량
                End If
                
                oBaseColRow01 = oBaseColRow01 + 1
                oMat02.Columns("CHK").Cells(i).Specific.Checked = False
            End If
        Next
    '//포장사업부 외주반출등록
    ElseIf oBaseForm01.TypeEx = "PS_MM135" Then
        Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//매트릭스
        Dim Qty As Double
        Qty = oBaseForm01.Items("Qty").Specific.VALUE
        For i = 1 To oMat02.RowCount '//배치선택품목
            If oMat02.Columns("CHK").Cells(i).Specific.Checked = True Then
                oBaseMat01.Columns("Qty").Cells(oBaseColRow01).Specific.VALUE = Qty '//제품수량
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oBaseItemCode01 '//제품품목
                oBaseMat01.Columns("OutItmCd").Cells(oBaseColRow01).Specific.VALUE = oMat02.Columns("ItemCode").Cells(i).Specific.VALUE '//품목
                
                If Qty <= 0 Then
                    oBaseMat01.Columns("OutQty").Cells(oBaseColRow01).Specific.VALUE = Val(oMat02.Columns("SelQty").Cells(i).Specific.VALUE) '//수량
                    oBaseMat01.Columns("OutWt").Cells(oBaseColRow01).Specific.VALUE = (Val(oMat02.Columns("SelWeight").Cells(i).Specific.VALUE)) '//재적
                Else
                    oBaseMat01.Columns("OutQty").Cells(oBaseColRow01).Specific.VALUE = Round(Val(oMat02.Columns("SelQty").Cells(i).Specific.VALUE) / Val(oMat02.Columns("StdQty").Cells(i).Specific.VALUE) * Qty, 0) '//수량
                    oBaseMat01.Columns("OutWt").Cells(oBaseColRow01).Specific.VALUE = Round(Val(oMat02.Columns("SelWeight").Cells(i).Specific.VALUE) / Val(oMat02.Columns("StdQty").Cells(i).Specific.VALUE) * Qty, 2) '//재적
                End If
                             
                oBaseColRow01 = oBaseColRow01 + 1
                oMat02.Columns("CHK").Cells(i).Specific.Checked = False
            End If
        Next
        '포장 원자재 재고조사등록
    ElseIf oBaseForm01.TypeEx = "PS_MM097" Then
        Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific '//매트릭스

        For i = 1 To oMat02.RowCount '//배치선택품목
            If oMat02.Columns("CHK").Cells(i).Specific.Checked = True Then
                oBaseMat01.Columns("ItemCode").Cells(oBaseColRow01).Specific.VALUE = oMat02.Columns("ItemCode").Cells(i).Specific.VALUE '//품목

                oBaseColRow01 = oBaseColRow01 + 1
                oMat02.Columns("CHK").Cells(i).Specific.Checked = False
            End If
        Next
   
    End If
    Exit Sub
PS_SM030_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_SM030_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



