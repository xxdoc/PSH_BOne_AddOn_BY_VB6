VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM180.cls
'//  Module         : 재고관리 > 입고
'//  Desc           : 기본
'//  FormType       : PS_MM180
'//  Create Date    : 2010.08.20
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_MM180H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_MM180L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    BatchNum As String
    WhsCode As String
    Quantity As Double
    Check As Boolean
    OIGNNo As Long
    IGN1No As Long
    OIGENo As Long
    IGE1No As Long
    MM180HNum As Long
    MM180LNum As Long
End Type
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM180.srf")
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
    
    oFormUniqueID01 = "PS_MM180_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_MM180_CreateItems
    Call PS_MM180_ComboBox_Setting
    Call PS_MM180_CF_ChooseFromList
    Call PS_MM180_EnableMenus
    Call PS_MM180_SetDocument(oFromDocEntry01)
    Call PS_MM180_FormResize
    
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
    Dim i As Long
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
                If oForm01.Mode = fm_OK_MODE Then
                    If (PS_MM180_Validate("취소") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If PS_MM180_DI_API_02 = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                Else
                    MDC_Com.MDC_GF_Message "현재 모드에서는 취소할수 없습니다.", "W"
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            Case "1287": '복제
                If oForm01.Mode = fm_OK_MODE Then
                    '//복제가능
                Else
                    MDC_Com.MDC_GF_Message "현재 모드에서는 복제할수 없습니다.", "W"
                    BubbleEvent = False
                    Exit Sub
                End If
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
                Call Calc_SumQty
            Case "1281": '찾기
                Call PS_MM180_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_MM180_FormItemEnabled '//UDO방식
                Call PS_MM180_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_MM180_FormItemEnabled
            Case "1287": '복제
                PS_MM180_FormItemEnabled
                If (oForm01.Mode = fm_ADD_MODE) Then
                    Call PS_MM180_FormClear
                End If
                For i = 1 To oMat01.VisualRowCount
                    oMat01.Columns("LineId").Cells(i).Specific.VALUE = ""
                Next
                
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
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("처리 중...", 100, False)
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "PS_MM180" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_MM180_BatchCreate
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Call PS_MM180_AddMatrixRow(0, True)
                oForm01.Items("SumQty").Specific.VALUE = ""
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_MM180_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    '입고(애드온)문서의 생성이 성공하였을 때만 DI 실행(2011.12.01 송명규 수정)
                    Call PS_MM180_DI_API_01
                End If
                
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                
                If PS_MM180_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_MM180" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_MM180_FormItemEnabled
                    Call PS_MM180_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_MM180_FormItemEnabled
                End If
            End If
        End If
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Exit Sub
    
Raise_EVENT_ITEM_PRESSED_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    Dim ChildForm01 As Variant
    If pval.BeforeAction = True Then
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        If pval.ItemUID = "BItemCod" Then
            If pval.CharPressed = 9 Then
                If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
                    Set ChildForm01 = New PS_SM010
                    Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        End If
        
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "ItemCode" Then
                If pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = "" Then
                        Set ChildForm01 = New PS_SM010
                        Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                        BubbleEvent = False
                        Exit Sub
                    End If
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
'        If pval.ItemChanged = True Then
'            Call oForm01.Freeze(True)
'            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_MM180L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM180L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_MM180_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_MM180L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.Value)
'                End If
'            Else
'                If (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_MM180H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_MM180H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.Value)
'                    Call oDS_PS_MM180H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Selected.Value & "'"))
'                Else
'                    Call oDS_PS_MM180H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.Value)
'                End If
'            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
'            If pval.ItemUID = "Mat01" Then
'                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
'            Else
'                oForm01.Items(pval.ItemUID).Click ct_Regular
'            End If
'            Call oForm01.Freeze(False)
'        End If
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
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
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
    Dim i As Integer
    Dim SumQty As Currency
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "ItemCode") Then
                    '//기타작업
                    '//품목코드처리
                    Call oDS_PS_MM180L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_MM180L.setValue("U_ItemName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM180L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_MM180_AddMatrixRow (pval.Row)
                    End If
                ' 2011.01.13
                ElseIf (pval.ColUID = "Quantity") Then
                    For i = 1 To oMat01.VisualRowCount - 1
                        SumQty = SumQty + oMat01.Columns("Quantity").Cells(i).Specific.VALUE
                    Next
                    oForm01.Items("SumQty").Specific.VALUE = SumQty
                    Call oDS_PS_MM180L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                Else
                    Call oDS_PS_MM180L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_MM180H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "BItemCod") Then
                    oForm01.Items("BItemNam").Specific.VALUE = MDC_PS_Common.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1)
                ElseIf (pval.ItemUID = "BWhsCode") Then
                ElseIf (pval.ItemUID = "BBatchNm") Then
                ElseIf (pval.ItemUID = "BBatchSt") Then
                ElseIf (pval.ItemUID = "BBatchEd") Then
                ElseIf (pval.ItemUID = "BQuantity") Then
                Else
                    Call oDS_PS_MM180H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
                oForm01.Items(pval.ItemUID).Click ct_Regular
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oForm01.Update
            If pval.ItemUID = "Mat01" Then
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                oForm01.Items(pval.ItemUID).Click ct_Regular
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
    Dim i As Integer
    Dim SumQty As Currency
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
        For i = 1 To oMat01.VisualRowCount
            SumQty = SumQty + oMat01.Columns("Quantity").Cells(i).Specific.VALUE
        Next
        oForm01.Items("SumQty").Specific.VALUE = SumQty
        Call PS_MM180_FormItemEnabled
        Call PS_MM180_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_MM180_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    Dim oDataTable01 As SAPbouiCOM.DataTable
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        If (pval.ItemUID = "BItemCod") Then
            Set oDataTable01 = pval.SelectedObjects
            oForm01.DataSources.UserDataSources("BItemCod").VALUE = oDataTable01.Columns("ItemCode").Cells(0).VALUE
            oForm01.DataSources.UserDataSources("BItemNam").VALUE = oDataTable01.Columns("ItemName").Cells(0).VALUE
            Set oDataTable01 = Nothing
        End If
        If (pval.ItemUID = "BWhsCode") Then
            Set oDataTable01 = pval.SelectedObjects
            oForm01.DataSources.UserDataSources("BWhsCode").VALUE = oDataTable01.Columns("WhsCode").Cells(0).VALUE
            oForm01.DataSources.UserDataSources("BWhsName").VALUE = oDataTable01.Columns("WhsName").Cells(0).VALUE
            Set oDataTable01 = Nothing
        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_MM180H", "U_CardCode,U_CardName")
'        End If
        
        If (pval.ItemUID = "Mat01") Then
            If (pval.ColUID = "ItemCode") Then
                If pval.SelectedObjects Is Nothing Then
                Else
                    Set oDataTable01 = pval.SelectedObjects
                    Call oDS_PS_MM180L.setValue("U_ItemCode", pval.Row - 1, oDataTable01.Columns("ItemCode").Cells(0).VALUE)
                    Call oDS_PS_MM180L.setValue("U_ItemName", pval.Row - 1, oDataTable01.Columns("ItemName").Cells(0).VALUE)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM180L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        Call PS_MM180_AddMatrixRow(pval.Row)
                    End If
                    Set oDataTable01 = Nothing
                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
                    oMat01.LoadFromDataSource
                    oMat01.AutoResizeColumns
                    oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                End If
            End If
        End If
        If (pval.ItemUID = "Mat01") Then
            If (pval.ColUID = "WhsCode") Then
                If pval.SelectedObjects Is Nothing Then
                Else
                    Set oDataTable01 = pval.SelectedObjects
                    Call oDS_PS_MM180L.setValue("U_WhsCode", pval.Row - 1, oDataTable01.Columns("WhsCode").Cells(0).VALUE)
                    Call oDS_PS_MM180L.setValue("U_WhsName", pval.Row - 1, oDataTable01.Columns("WhsName").Cells(0).VALUE)
                    Set oDataTable01 = Nothing
                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
                    oMat01.LoadFromDataSource
                    oMat01.AutoResizeColumns
                    oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                End If
            End If
        End If
    End If
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
            If (PS_MM180_Validate("행삭제") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_MM180L.RemoveRecord(oDS_PS_MM180L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_MM180_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_MM180L.GetValue("U_ItemCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_MM180_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_MM180_CreateItems() As Boolean
On Error GoTo PS_MM180_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM180H = oForm01.DataSources.DBDataSources("@PS_MM180H")
    Set oDS_PS_MM180L = oForm01.DataSources.DBDataSources("@PS_MM180L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("BItemCod", dt_SHORT_TEXT, 30)
    Call oForm01.DataSources.UserDataSources.Add("BWhsCode", dt_SHORT_TEXT, 15)
    Call oForm01.DataSources.UserDataSources.Add("BItemNam", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("BWhsName", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("BBatchNm", dt_SHORT_TEXT, 32)
    Call oForm01.DataSources.UserDataSources.Add("BBatchSt", dt_SHORT_TEXT, 29)
    Call oForm01.DataSources.UserDataSources.Add("BBatchEd", dt_SHORT_TEXT, 29)
    Call oForm01.DataSources.UserDataSources.Add("BQuantity", dt_QUANTITY)
    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_QUANTITY, 100)
    
    Call oForm01.Items("BItemCod").Specific.DataBind.SetBound(True, "", "BItemCod")
    Call oForm01.Items("BWhsCode").Specific.DataBind.SetBound(True, "", "BWhsCode")
    Call oForm01.Items("BItemNam").Specific.DataBind.SetBound(True, "", "BItemNam")
    Call oForm01.Items("BWhsName").Specific.DataBind.SetBound(True, "", "BWhsName")
    Call oForm01.Items("BBatchNm").Specific.DataBind.SetBound(True, "", "BBatchNm")
    Call oForm01.Items("BBatchSt").Specific.DataBind.SetBound(True, "", "BBatchSt")
    Call oForm01.Items("BBatchEd").Specific.DataBind.SetBound(True, "", "BBatchEd")
    Call oForm01.Items("BQuantity").Specific.DataBind.SetBound(True, "", "BQuantity")
    Call oForm01.Items("SumQty").Specific.DataBind.SetBound(True, "", "SumQty")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_MM180_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM180_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_MM180_ComboBox_Setting()
On Error GoTo PS_MM180_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM180", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM180", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_MM180", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM180", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM180", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_MM180", "ItemCode")
    
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM180_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM180_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM180_CF_ChooseFromList()
On Error GoTo PS_MM180_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

'    Set oEdit = oForm01.Items("BItemCod").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = "4"
'    oCFLCreationParams.uniqueID = "CFLITEMCODE"
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
'    oEdit.ChooseFromListUID = "CFLITEMCODE"
'    oEdit.ChooseFromListAlias = "ItemCode"

    Set oEdit = oForm01.Items("BWhsCode").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = lf_Warehouses
    oCFLCreationParams.uniqueID = "CFLWHSCODE"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oEdit.ChooseFromListUID = "CFLWHSCODE"
    oEdit.ChooseFromListAlias = "WhsCode"
    
'    Set oColumn = oMat01.Columns("ItemCode")
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = "4"
'    oCFLCreationParams.uniqueID = "CFLITEMCODE2"
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
'    oColumn.ChooseFromListUID = "CFLITEMCODE2"
'    oColumn.ChooseFromListAlias = "ItemCode"

    Set oColumn = oMat01.Columns("WhsCode")
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = lf_Warehouses
    oCFLCreationParams.uniqueID = "CFLWHSCODE2"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oColumn.ChooseFromListUID = "CFLWHSCODE2"
    oColumn.ChooseFromListAlias = "WhsCode"
    
    Exit Sub
PS_MM180_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM180_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM180_FormItemEnabled()
On Error GoTo PS_MM180_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("Button01").Enabled = True
        oForm01.Items("Button02").Enabled = True
        oForm01.Items("Mat01").Enabled = True
        
        oForm01.Items("BItemCod").Specific.VALUE = ""
        oForm01.Items("BWhsCode").Specific.VALUE = ""
        oForm01.Items("BBatchNm").Specific.VALUE = ""
        oForm01.Items("BBatchSt").Specific.VALUE = ""
        oForm01.Items("BBatchEd").Specific.VALUE = ""
        oForm01.Items("BQuantity").Specific.VALUE = 0
        Call PS_MM180_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("Button01").Enabled = False
        oForm01.Items("Button02").Enabled = False
        oForm01.Items("Mat01").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("DocDate").Enabled = False
        oForm01.Items("Button01").Enabled = False
        oForm01.Items("Button02").Enabled = False
        oForm01.Items("Mat01").Enabled = False
    End If
    oMat01.AutoResizeColumns
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM180_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM180_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM180_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_MM180_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM180L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM180L.offset = oRow
    oDS_PS_MM180L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM180_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM180_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM180_FormClear()
On Error GoTo PS_MM180_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM180'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_MM180_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM180_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM180_EnableMenus()
On Error GoTo PS_MM180_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, True, False, True, False, False, False, False) '//메뉴설정
    Exit Sub
PS_MM180_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM180_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM180_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_MM180_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_MM180_FormItemEnabled
        Call PS_MM180_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_MM180_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_MM180_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM180_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_MM180_DataValidCheck() As Boolean
On Error GoTo PS_MM180_DataValidCheck_Error
    PS_MM180_DataValidCheck = False
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_MM180_FormClear
    End If
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "전기일은 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_MM180_DataValidCheck = False
        Exit Function
    End If
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_MM180_DataValidCheck = False
        Exit Function
    End If
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("ItemCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
            oMat01.Columns("ItemCode").Cells(i).Click ct_Regular
            PS_MM180_DataValidCheck = False
            Exit Function
        End If

        If (oMat01.Columns("WhsCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
            oMat01.Columns("WhsCode").Cells(i).Click ct_Regular
            PS_MM180_DataValidCheck = False
            Exit Function
        End If
        If (Val(oMat01.Columns("Quantity").Cells(i).Specific.VALUE) <= 0) Then
            Sbo_Application.SetStatusBarMessage "중량은 필수입니다.", bmt_Short, True
            oMat01.Columns("Quantity").Cells(i).Click ct_Regular
            PS_MM180_DataValidCheck = False
            Exit Function
        End If
        
        '//배치존재유무검사
        Dim BatchYN As String
        BatchYN = MDC_PS_Common.GetItem_ManBtchNum(oMat01.Columns("ItemCode").Cells(i).Specific.VALUE)
        
        If BatchYN = "Y" Then
            If (oMat01.Columns("BatchNum").Cells(i).Specific.VALUE = "") Then
                Sbo_Application.SetStatusBarMessage "배치는 필수입니다.", bmt_Short, True
                oMat01.Columns("BatchNum").Cells(i).Click ct_Regular
                PS_MM180_DataValidCheck = False
                Exit Function
            End If
        End If
        
        If BatchYN = "Y" Then
            If (UCase(oMat01.Columns("BatchNum").Cells(i).Specific.VALUE) <> oMat01.Columns("BatchNum").Cells(i).Specific.VALUE) Then
                Sbo_Application.SetStatusBarMessage "소문자로 입력하셨습니다.배치번호를 확인하십시오.", bmt_Short, True
                oMat01.Columns("BatchNum").Cells(i).Click ct_Regular
                PS_MM180_DataValidCheck = False
                Exit Function
            End If
        End If
                
        If BatchYN = "Y" Then
        
            '속도 문제로 쿼리 수정(2017.12.18 송명규), 조회 조건에서 품목코드 삭제, '18.01.02 복원
            'OIBT에서 ItemCode를 포함하면 502010128(PMC90R-0 1.8T*34.5W) 품목의 조회 시간이 1:30초를 초과하여 AddOn이 끊기는 현상 발생함
            'OIBT(View)가 아닌 OBTN(Table), OBTQ(Table)에서 직접 SELECT 하는 방식으로 쿼리 변경(2018.01.18 송명규)
            Query01 = "           SELECT    COUNT(*) "
            Query01 = Query01 + " FROM      [OBTN] AS T0"
            Query01 = Query01 + "           LEFT JOIN"
            Query01 = Query01 + "           [OBTQ] AS T1"
            Query01 = Query01 + "               ON T0.ItemCode = T1.ItemCode"
            Query01 = Query01 + "               AND T0.SysNumber = T1.SysNumber"
            Query01 = Query01 + " WHERE     T0.DistNumber = '" & oMat01.Columns("BatchNum").Cells(i).Specific.VALUE & "'"
            Query01 = Query01 + "           AND T0.ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & "'"
            Query01 = Query01 + "           AND T1.WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.VALUE & "'"
            Query01 = Query01 + "           AND T1.Quantity > 0"

            If MDC_PS_Common.GetValue(Query01) > 0 Then
                Sbo_Application.SetStatusBarMessage "해당품목의 가용배치가 존재합니다.", bmt_Short, True
                oMat01.Columns("ItemCode").Cells(i).Click ct_Regular
                PS_MM180_DataValidCheck = False
                Exit Function
            End If
        End If
        
        If BatchYN = "Y" Then
            '//속도가 느려 수정 2012.03.21 노근용
            Query01 = "select COUNT(*)"
            Query01 = Query01 + " from [@PS_PP040H] a Inner Join [@PS_PP040L] b On a.DocEntry = b.DocEntry And a.Canceled = 'N'"
            Query01 = Query01 + " Inner Join [@PS_PP030L] c On b.U_PP030HNo = c.DocEntry"
            Query01 = Query01 + " Where c.U_BatchNum = '" + oMat01.Columns("BatchNum").Cells(i).Specific.VALUE + "'"
            If MDC_PS_Common.GetValue(Query01) > 0 Then
                Sbo_Application.SetStatusBarMessage "이미 생산에 투입된 배치번호 입니다.", bmt_Short, True
                oMat01.Columns("ItemCode").Cells(i).Click ct_Regular
                PS_MM180_DataValidCheck = False
                Exit Function
            End If
'            Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            '//작업일보에 등록된 작업지시의 투입품, 멀티게이지의 경우만 해당된다.
            
'            Query01 = "SELECT U_BatchNum FROM [@PS_PP030L] WHERE DocEntry IN(SELECT U_PP030HNo FROM [@PS_PP040L] WHERE U_OrdGbn = '104') AND U_ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & "'"
'            RecordSet01.DoQuery Query01
'            For j = 0 To RecordSet01.RecordCount - 1
'                If RecordSet01.Fields(0).VALUE = oMat01.Columns("BatchNum").Cells(i).Specific.VALUE Then
'                    '//해당배치는 이미 생산에 투입된배치
'                    Sbo_Application.SetStatusBarMessage "이미 생산에 투입된 배치번호 입니다.", bmt_Short, True
'                    PS_MM180_DataValidCheck = False
'                    Set RecordSet01 = Nothing
'                    Exit Function
'                End If
'                RecordSet01.MoveNext
'            Next
'            Set RecordSet01 = Nothing
        End If
        
        '//품목,배치 중복검사
        If BatchYN = "Y" Then
            For j = i + 1 To oMat01.VisualRowCount - 1
                If (oMat01.Columns("ItemCode").Cells(i).Specific.VALUE = oMat01.Columns("ItemCode").Cells(j).Specific.VALUE) Then '//품목코드가 같고..
                    If (oMat01.Columns("BatchNum").Cells(i).Specific.VALUE = oMat01.Columns("BatchNum").Cells(j).Specific.VALUE) Then '//배치번호가 같으면..
                        Sbo_Application.SetStatusBarMessage "중복된 품목,배치가 존재합니다.", bmt_Short, True
                        oMat01.Columns("ItemCode").Cells(j).Click ct_Regular
                        PS_MM180_DataValidCheck = False
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
    Call oDS_PS_MM180L.RemoveRecord(oDS_PS_MM180L.Size - 1)
    Call oMat01.LoadFromDataSource
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_MM180_FormClear
    End If
    PS_MM180_DataValidCheck = True
    Exit Function
PS_MM180_DataValidCheck_Error:
    PS_MM180_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_MM180_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_MM180_MTX01()
On Error GoTo PS_MM180_MTX01_Error
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
    Param01 = oForm01.Items("Param01").Specific.VALUE
    Param02 = oForm01.Items("Param01").Specific.VALUE
    Param03 = oForm01.Items("Param01").Specific.VALUE
    Param04 = oForm01.Items("Param01").Specific.VALUE
    
    Query01 = "SELECT 10"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_MM180_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_MM180L.InsertRecord (i)
        End If
        oDS_PS_MM180L.offset = i
        Call oDS_PS_MM180L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_MM180L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_MM180_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_MM180_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM180_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_MM180_DI_API_01() As Boolean
On Error GoTo PS_MM180_DI_API_01_Error
    PS_MM180_DI_API_01 = True
    Dim i, j As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim LineNumCount As Long
    Dim ResultDocNum As Long
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenEntry)
    oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    oDIObject.UserFields("Comments").VALUE = "PS_MM180: 문서번호 " & oForm01.Items("DocEntry").Specific.VALUE & " 입고" '//2010/12/21 노대리님요청
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.Quantity = ItemInformation(i).Quantity
        If MDC_PS_Common.GetItem_ManBtchNum(ItemInformation(i).ItemCode) = "Y" Then
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(i).BatchNum
            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(i).Quantity
            oDIObject.Lines.BatchNumbers.Add
        End If
        ItemInformation(i).IGN1No = LineNumCount
        LineNumCount = LineNumCount + 1
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To UBound(ItemInformation)
            Call oDS_PS_MM180L.setValue("U_OIGNNum", i, ResultDocNum)
            Call oDS_PS_MM180L.setValue("U_IGN1Num", i, ItemInformation(i).IGN1No)
        Next
    Else
        GoTo PS_MM180_DI_API_01_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_MM180_DI_API_01_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_MM180_DI_API_01 = False
    Set oDIObject = Nothing
    Exit Function
PS_MM180_DI_API_01_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM180_DI_API_01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_MM180_DI_API_01 = False
End Function

Private Function PS_MM180_DI_API_02() As Boolean
On Error GoTo PS_MM180_DI_API_02_Error
    PS_MM180_DI_API_02 = True
    Dim i, j As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim LineNumCount As Long
    Dim ResultDocNum As Long
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount - 1
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).OIGNNo = oMat01.Columns("OIGNNum").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).MM180HNum = Val(oForm01.Items("DocEntry").Specific.VALUE)
        ItemInformation(ItemInformationCount).MM180LNum = Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenExit) '//기타출고
    oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    oDIObject.UserFields("Comments").VALUE = "PS_MM180: 문서번호 " & oForm01.Items("DocEntry").Specific.VALUE & " 입고취소" '//2010/12/21 노대리님요청
    oDIObject.UserFields("U_CancDoc").VALUE = ItemInformation(i - 2).OIGNNo 'ItemInformation(i - 2).MM180HNum
    oDIObject.UserFields("U_CtrlType").VALUE = "C" '//2010/12/21 노대리님요청
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.Quantity = ItemInformation(i).Quantity
        
        If MDC_PS_Common.GetItem_ManBtchNum(ItemInformation(i).ItemCode) = "Y" Then
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(i).BatchNum
            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(i).Quantity
            oDIObject.Lines.BatchNumbers.Add
        End If
        ItemInformation(i).IGE1No = LineNumCount
        LineNumCount = LineNumCount + 1
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To UBound(ItemInformation)
            MDC_PS_Common.DoQuery ("UPDATE [@PS_MM180L] SET U_OIGENum = '" & ResultDocNum & "', U_IGE1Num = '" & ItemInformation(i).IGE1No & "' WHERE DocEntry = '" & ItemInformation(i).MM180HNum & "' AND LineId = '" & ItemInformation(i).MM180LNum & "'")
        Next
    Else
        GoTo PS_MM180_DI_API_02_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_MM180_DI_API_02_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_MM180_DI_API_02 = False
    Set oDIObject = Nothing
    Exit Function
PS_MM180_DI_API_02_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM180_DI_API_02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_MM180_DI_API_02 = False
End Function

Private Sub PS_MM180_FormResize()
On Error GoTo PS_MM180_FormResize_Error
    
    Exit Sub
PS_MM180_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM180_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_MM180_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_MM180_Validate_Error
    PS_MM180_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_MM180H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
        PS_MM180_Validate = False
        GoTo PS_MM180_Validate_Exit
    End If
'
    If ValidateType = "수정" Then
'        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
'        Dim Exist As Boolean
'        Exist = False
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            Exist = False
'            For j = 1 To oMat01.RowCount - 1
'                '//라인번호가 같고, 품목코드가 같으면 존재하는행 , LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
'                If Val(RecordSet01.Fields(1).Value) = Val(oMat01.Columns("U_LineNum").Cells(j).Specific.Value) And RecordSet01.Fields(2).Value = oMat01.Columns("1").Cells(j).Specific.Value And oMat01.Columns("U_LineNum").Cells(j).Specific.Value <> "" Then
'                    Exist = True
'                End If
'            Next
'            If (Exist = False) Then '//삭제된 행중
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "삭제된행이 다른사용자에 의해 출하,선출요청되었습니다. 적용할수 없습니다.", "W"
'                    PS_MM180_Validate = False
'                    GoTo PS_MM180_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
            MDC_Com.MDC_GF_Message "해당모드에서는 행삭제를 할수 없습니다.", "W"
            PS_MM180_Validate = False
            GoTo PS_MM180_Validate_Exit
        End If
    ElseIf ValidateType = "취소" Then
        Query01 = "SELECT U_ItemCode AS ItemCode,U_BatchNum AS BatchNum,U_WhsCode AS WhsCode,U_Quantity AS Quantity FROM [@PS_MM180L] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        RecordSet01.DoQuery Query01
        For i = 0 To RecordSet01.RecordCount - 1
            If MDC_PS_Common.GetItem_ManBtchNum(RecordSet01.Fields(0).VALUE) = "Y" Then
                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [OIBT] WHERE ItemCode = '" & RecordSet01.Fields("ItemCode").VALUE & "' AND BatchNum = '" & RecordSet01.Fields("BatchNum").VALUE & "' AND WhsCode = '" & RecordSet01.Fields("WhsCode").VALUE & "' AND Quantity = " & RecordSet01.Fields("Quantity").VALUE & "", 0, 1)) <= 0 Then
                    MDC_Com.MDC_GF_Message "품목코드 : " & RecordSet01.Fields(0).VALUE & ", 배치번호 : " & RecordSet01.Fields(1).VALUE & " 의 재고가 존재하지 않습니다.", "W"
                    PS_MM180_Validate = False
                    GoTo PS_MM180_Validate_Exit
                End If
            End If
            RecordSet01.MoveNext
        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_MM180_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_MM180_Validate_Error:
    PS_MM180_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_MM180_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Public Sub PS_MM180_BatchCreate()
On Error GoTo PS_MM180_BatchCreate_Error
    If oForm01.Items("BItemCod").Specific.VALUE = "" Then
        Call MDC_Com.MDC_GF_Message("품목코드를 선택하지 않았습니다.", "W")
        oForm01.Items("BItemCod").Click ct_Regular
        Exit Sub
    End If
    If oForm01.Items("BWhsCode").Specific.VALUE = "" Then
        Call MDC_Com.MDC_GF_Message("창고코드를 선택하지 않았습니다.", "W")
        oForm01.Items("BWhsCode").Click ct_Regular
        Exit Sub
    End If
    If oForm01.Items("BBatchNm").Specific.VALUE = "" Then
        Call MDC_Com.MDC_GF_Message("배치코드를 선택하지 않았습니다.", "W")
        oForm01.Items("BBatchNm").Click ct_Regular
        Exit Sub
    End If
    If oForm01.Items("BBatchSt").Specific.VALUE = "" Then
        Call MDC_Com.MDC_GF_Message("배치시작번호를 선택하지 않았습니다.", "W")
        oForm01.Items("BBatchSt").Click ct_Regular
        Exit Sub
    End If
    If oForm01.Items("BBatchEd").Specific.VALUE = "" Then
        Call MDC_Com.MDC_GF_Message("배치끝번호를 선택하지 않았습니다.", "W")
        oForm01.Items("BBatchEd").Click ct_Regular
        Exit Sub
    End If
    If Val(oForm01.Items("BQuantity").Specific.VALUE) <= 0 Then
        Call MDC_Com.MDC_GF_Message("중량을 선택하지 않았습니다.", "W")
        oForm01.Items("BQuantity").Click ct_Regular
        Exit Sub
    End If
    If Len(Trim(oForm01.Items("BBatchNm").Specific.VALUE)) <> 8 Then
        Call MDC_Com.MDC_GF_Message("배치코드는 8자리 입니다. 확인하여 주십시오.", "W")
        oForm01.Items("BBatchNm").Click ct_Regular
        Exit Sub
    End If
    
    Dim i As Long
    Dim ValidBatch As Long
    Dim StartValue As Long
    Dim EndValue As Long
    Dim BatchNum As String
    Dim SumQty As Currency
    
    StartValue = Val(Right(oForm01.Items("BBatchSt").Specific.VALUE, 3))
    EndValue = Val(Right(oForm01.Items("BBatchEd").Specific.VALUE, 3))
    
    If StartValue = 0 And EndValue = 0 Then
        Call MDC_Com.MDC_GF_Message("배치번호의 범위가 올바르지 않습니다.", "W")
        Exit Sub
    End If
    ValidBatch = EndValue - StartValue '//앞에문서 - 뒤에문서
    If ValidBatch < 0 Then
        Call MDC_Com.MDC_GF_Message("배치번호의 범위가 올바르지 않습니다.", "W")
        Exit Sub
    End If
    
    oForm01.Freeze True
    ValidBatch = (ValidBatch / 10) + 1 '//10단위로 몇개생성가능한지 계산
    Dim MatrixRow As Long
    MatrixRow = oMat01.VisualRowCount
    For i = 1 To ValidBatch
        BatchNum = StartValue + (10 * (i - 1))
        If Val(BatchNum) < 100 Then
            BatchNum = "0" & BatchNum
        Else
            BatchNum = BatchNum
        End If
        Call oDS_PS_MM180L.setValue("U_ItemCode", MatrixRow - 1, oForm01.Items("BItemCod").Specific.VALUE)
        Call oDS_PS_MM180L.setValue("U_ItemName", MatrixRow - 1, MDC_PS_Common.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" & oForm01.Items("BItemCod").Specific.VALUE & "'", 0, 1))
        Call oDS_PS_MM180L.setValue("U_BatchNum", MatrixRow - 1, UCase(oForm01.Items("BBatchNm").Specific.VALUE & BatchNum))
        Call oDS_PS_MM180L.setValue("U_WhsCode", MatrixRow - 1, oForm01.Items("BWhsCode").Specific.VALUE)
        Call oDS_PS_MM180L.setValue("U_WhsName", MatrixRow - 1, MDC_PS_Common.GetValue("SELECT WhsName FROM [OWHS] WHERE WhsCode = '" & oForm01.Items("BWhsCode").Specific.VALUE & "'", 0, 1))
        Call oDS_PS_MM180L.setValue("U_Quantity", MatrixRow - 1, oForm01.Items("BQuantity").Specific.VALUE)
        Call PS_MM180_AddMatrixRow(MatrixRow)
        MatrixRow = MatrixRow + 1
    Next
    
    Call Calc_SumQty
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Freeze False
    Exit Sub
PS_MM180_BatchCreate_Error:
    oForm01.Freeze False
    Sbo_Application.SetStatusBarMessage "PS_MM180_BatchCreate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Calc_SumQty()
On Error GoTo Calc_SumQty_Error
    Dim i As Long
    Dim SumQty As Currency
    
    '//2011.01.19 추가
    For i = 1 To oMat01.VisualRowCount - 1
        SumQty = SumQty + oMat01.Columns("Quantity").Cells(i).Specific.VALUE
    Next
    oForm01.Items("SumQty").Specific.VALUE = SumQty
    
    Exit Sub
Calc_SumQty_Error:
    oForm01.Freeze False
    Sbo_Application.SetStatusBarMessage "Calc_SumQty_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
