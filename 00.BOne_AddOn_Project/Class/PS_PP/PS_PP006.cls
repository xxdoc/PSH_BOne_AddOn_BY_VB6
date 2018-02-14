VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP006.cls
'//  Module         : 재고  > 외주업체별 가공비단가등록
'//  Desc           : 기본
'//  FormType       : PS_PP006
'//  Create Date    : 2010.10.25
'//  Created By     : Kim Dong Sub
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
'Private oDS_PS_PP006H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP006H As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    LotNo As String
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
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP006.srf")
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
    
    oFormUniqueID01 = "PS_PP006_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP006_CreateItems
    Call PS_PP006_ComboBox_Setting
    Call PS_PP006_CF_ChooseFromList
    Call PS_PP006_EnableMenus
    Call PS_PP006_SetDocument(oFromDocEntry01)
    Call PS_PP006_FormResize
    
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
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_PP006_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call oMat01.Clear
                Call oMat01.FlushToDataSource
                Call oMat01.LoadFromDataSource
                Call PS_PP006_FormItemEnabled '//UDO방식
                Call PS_PP006_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_PP006_FormItemEnabled
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
    If pval.BeforeAction = True Then
        If pval.ItemUID = "PS_PP006" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "Btn01" Then      ' 조회
            Call PS_PP006_MTX01
        End If
                
        If pval.ItemUID = "1" Then      ' 저장
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_PP006_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_PP006_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            '//해야할일 작업
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP006" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "eCardCod") '//사용자값활성
        If pval.ColUID = "CpCode" Then
            If oMat01.Columns("CpCode").Cells(pval.Row).Specific.VALUE = "" Then
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "CpCode") '매트릭스의 품목중분류 포맷서치 적용
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
        If pval.ItemChanged = True Then
            Call oForm01.Freeze(True)
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "Unit") Then
                    '//기타작업
                    Call oDS_PS_PP006H.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP006H.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_PP006_AddMatrixRow (pval.Row)
                    End If
   
                Else
                    Call oDS_PS_PP006H.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
                
            Else
                If (pval.ItemUID = "CardCode") Then
                    Call oDS_PS_PP006H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "CardCode") Then
                    Call oDS_PS_PP006H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_PP006H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
                Else
                    Call oDS_PS_PP006H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oForm01.Update
            Call oForm01.Freeze(False)
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
'        If pval.ItemUID = "Mat01" Then
'            If pval.Row > 0 Then
'                Call oMat01.SelectRow(pval.Row, True, False)
'            End If
'        End If

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
Dim sQry01              As String
Dim oRecordSet01        As SAPbobsCOM.Recordset
Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If oForm01.Mode = fm_UPDATE_MODE Or oForm01.Mode = fm_ADD_MODE Then
                    If (pval.ColUID = "eCardCod") Then
                    
                        Call oDS_PS_PP006H.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP006H.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                            PS_PP006_AddMatrixRow (pval.Row)
                        End If
    
                        oMat01.LoadFromDataSource
                    
                        sQry01 = "SELECT CardName FROM [OCRD] WHERE CardCode = '" & Trim(oMat01.Columns("eCardCod").Cells(pval.Row).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry01
                        oMat01.Columns("eCardNam").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    
                    ElseIf (pval.ColUID = "ItemCode") Then
'                        oForm01.Mode = fm_UPDATE_MODE
                    ElseIf (pval.ColUID = "Cprice") Then
'                        oForm01.Mode = fm_UPDATE_MODE
                    ElseIf (pval.ColUID = "CtrDate") Then
'                        oForm01.Mode = fm_UPDATE_MODE
                    ElseIf (pval.ColUID = "CpCode") Then
                        Call oDS_PS_PP006H.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP006H.setValue("U_CpName", pval.Row - 1, MDC_GetData.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" & Trim(oDS_PS_PP006H.GetValue("U_CpCode", pval.Row - 1)) & "'"))
                        oMat01.LoadFromDataSource
                    End If
                End If
            ElseIf (pval.ItemUID = "CardCode") Then
                sQry01 = "SELECT CardName FROM [OCRD] WHERE CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry01
                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            Else
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
        Call PS_PP006_FormItemEnabled
        '//Call PS_PP006_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP006_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    Dim i As Integer
    
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Dim oDataTable01 As SAPbouiCOM.DataTable
        Set oDataTable01 = pval.SelectedObjects
        If (pval.ItemUID = "ItemCode") Then
            oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
            oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
            Set oDataTable01 = Nothing
        ElseIf (pval.ItemUID = "CardCode") Then
            oForm01.DataSources.UserDataSources("CardCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
            oForm01.DataSources.UserDataSources("CardName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
            Set oDataTable01 = Nothing
        End If
        
        If (pval.ItemUID = "Mat01") Then
            If (pval.ColUID = "eCardCod") Then
                oMat01.FlushToDataSource
                Set oDataTable01 = pval.SelectedObjects
'                oMat01.Columns("eCardCod").Cells(pval.Row).Specific.Value = oDataTable01.Columns("CardCode").Cells(0).Value
'                oMat01.Columns("eCardNam").Cells(pval.Row).Specific.Value = oDataTable01.Columns("CardName").Cells(0).Value
'                For i = 1 To oMat01.VisualRowCount
                Call oDS_PS_PP006H.setValue("U_eCardCod", pval.Row - 1, oDataTable01.Columns("CardCode").Cells(0).VALUE)
                Call oDS_PS_PP006H.setValue("U_eCardNam", pval.Row - 1, oDataTable01.Columns("CardName").Cells(0).VALUE)
'                    Call oDS_PS_PP006H.setValue("U_Cprice", i - 1, oMat01.Columns("Cprice").Cells(pval.Row).Specific.Value)
'                    Call oDS_PS_PP006H.setValue("U_CtrDate", i - 1, oMat01.Columns("CtrDate").Cells(pval.Row).Specific.Value)
'                Next i
                If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP006H.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                    Call PS_PP006_AddMatrixRow(pval.Row)
                End If
                
                Set oDataTable01 = Nothing
                
                oMat01.LoadFromDataSource
                
                
            ElseIf (pval.ColUID = "ItemCode") Then
                oMat01.FlushToDataSource
                Set oDataTable01 = pval.SelectedObjects
'                oMat01.Columns("ItemCode").Cells(pval.Row).Specific.Value = oDataTable01.Columns("ItemCode").Cells(0).Value
'                oMat01.Columns("ItemName").Cells(pval.Row).Specific.Value = oDataTable01.Columns("ItemName").Cells(0).Value
'                For i = 1 To oMat01.VisualRowCount
                Call oDS_PS_PP006H.setValue("U_ItemCode", pval.Row - 1, oDataTable01.Columns("ItemCode").Cells(0).VALUE)
                Call oDS_PS_PP006H.setValue("U_ItemName", pval.Row - 1, oDataTable01.Columns("ItemName").Cells(0).VALUE)
'                    Call oDS_PS_PP006H.setValue("U_Cprice", i - 1, oMat01.Columns("Cprice").Cells(pval.Row).Specific.Value)
'                    Call oDS_PS_PP006H.setValue("U_CtrDate", i - 1, oMat01.Columns("CtrDate").Cells(pval.Row).Specific.Value)
'                Next i
                
                If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP006H.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                    Call PS_PP006_AddMatrixRow(pval.Row)
                End If
                
                Set oDataTable01 = Nothing

                oMat01.LoadFromDataSource
                
                
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
'            If (PS_PP006_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_PP006H.RemoveRecord(oDS_PS_PP006H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_PP006_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_PP006H.GetValue("U_Code", oMat01.RowCount - 1)) <> "" Then
                    Call PS_PP006_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP006_CreateItems() As Boolean
On Error GoTo PS_PP006_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim optBtn       As SAPbouiCOM.OptionBtn
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Set oDS_PS_PP006H = oForm01.DataSources.DBDataSources("@PS_PP006H")
    Set oDS_PS_PP006H = oForm01.DataSources.DBDataSources("@PS_PP006H")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("Rad01", dt_SHORT_TEXT, 20)
    Call oForm01.DataSources.UserDataSources.Add("Rad02", dt_SHORT_TEXT, 20)
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("CardName", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("CtrDate", dt_DATE, 8)
    
    Call oForm01.Items("Rad01").Specific.DataBind.SetBound(True, "", "Rad01")
    Call oForm01.Items("Rad02").Specific.DataBind.SetBound(True, "", "Rad02")
    Call oForm01.Items("CardCode").Specific.DataBind.SetBound(True, "", "CardCode")
    Call oForm01.Items("CardName").Specific.DataBind.SetBound(True, "", "CardName")
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    Call oForm01.Items("CtrDate").Specific.DataBind.SetBound(True, "", "CtrDate")
    
    Set optBtn = oForm01.Items("Rad01").Specific
    optBtn.ValOn = "A"
    optBtn.ValOff = "0"
    optBtn.Selected = True

    Set optBtn = oForm01.Items("Rad02").Specific
    optBtn.ValOn = "B"
    optBtn.ValOff = "0"
    optBtn.GroupWith ("Rad01")
    
    Set optBtn = Nothing
    
    oForm01.Settings.MatrixUID = "Mat01"             ' 서식세팅
    oForm01.Settings.Enabled = True
    oForm01.Settings.EnableRowFormat = True
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP006_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP006_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP006_ComboBox_Setting()
On Error GoTo PS_PP006_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP006", "Mat01", "Unit", "EA", "수량")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP006", "Mat01", "Unit", "KG", "중량")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Unit"), "PS_PP006", "Mat01", "Unit")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP006", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP006", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PP006", "ItemCode")
    
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP006_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP006_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP006_CF_ChooseFromList()
On Error GoTo PS_PP006_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    ' 거래처 - 아이템
    Dim oCFLs01               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons01               As SAPbouiCOM.Conditions
    Dim oCon01                As SAPbouiCOM.Condition
    Dim oCFL01                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams01  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit01               As SAPbouiCOM.EditText
    Dim oColumn01             As SAPbouiCOM.Column

'    Set oEdit01 = oForm01.Items("CardCode").Specific
'    Set oCFLs01 = oForm01.ChooseFromLists
'    Set oCFLCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams01.ObjectType = lf_BusinessPartner
'    oCFLCreationParams01.uniqueID = "CFLCARDCODE01"
'    oCFLCreationParams01.MultiSelection = False
'    Set oCFL01 = oCFLs01.Add(oCFLCreationParams01)
'
'    oEdit01.ChooseFromListUID = "CFLCARDCODE01"
'    oEdit01.ChooseFromListAlias = "CardCode"
    
    ' 품목 - 아이템
    Dim oCFLs02               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons02               As SAPbouiCOM.Conditions
    Dim oCon02                As SAPbouiCOM.Condition
    Dim oCFL02                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams02  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit02               As SAPbouiCOM.EditText
    Dim oColumn02             As SAPbouiCOM.Column
    
    Set oEdit02 = oForm01.Items("ItemCode").Specific
    Set oCFLs02 = oForm01.ChooseFromLists
    Set oCFLCreationParams02 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
    
    oCFLCreationParams02.ObjectType = "4"
    oCFLCreationParams02.uniqueID = "CFLITEMCD02"
    oCFLCreationParams02.MultiSelection = False
    Set oCFL02 = oCFLs02.Add(oCFLCreationParams02)

    oEdit02.ChooseFromListUID = "CFLITEMCD02"
    oEdit02.ChooseFromListAlias = "ItemCode"
    
'    ' 거래처 - 매트릭스
'    Dim oCFLs03               As SAPbouiCOM.ChooseFromListCollection
'    Dim oCons03                As SAPbouiCOM.Conditions
'    Dim oCon03                As SAPbouiCOM.Condition
'    Dim oCFL03                As SAPbouiCOM.ChooseFromList
'    Dim oCFLCreationParams03  As SAPbouiCOM.ChooseFromListCreationParams
'    Dim oEdit03               As SAPbouiCOM.EditText
'    Dim oColumn03             As SAPbouiCOM.Column
'
'    Set oColumn03 = oMat01.Columns("eCardCod")
'    Set oCFLs03 = oForm01.ChooseFromLists
'    Set oCFLCreationParams03 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams03.ObjectType = lf_BusinessPartner
'    oCFLCreationParams03.uniqueID = "CFLCARDCODE03"
'    oCFLCreationParams03.MultiSelection = False
'    Set oCFL03 = oCFLs03.Add(oCFLCreationParams03)
'
'    oColumn03.ChooseFromListUID = "CFLCARDCODE03"
'    oColumn03.ChooseFromListAlias = "CardCode"
    
    ' 품목 - 매트릭스
    Dim oCFLs04               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons04               As SAPbouiCOM.Conditions
    Dim oCon04                As SAPbouiCOM.Condition
    Dim oCFL04                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams04  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit04               As SAPbouiCOM.EditText
    Dim oColumn04             As SAPbouiCOM.Column
    
    Set oColumn04 = oMat01.Columns("ItemCode")
    Set oCFLs04 = oForm01.ChooseFromLists
    Set oCFLCreationParams04 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
    
    oCFLCreationParams04.ObjectType = "4"
    oCFLCreationParams04.uniqueID = "CFLITEMCD04"
    oCFLCreationParams04.MultiSelection = False
    Set oCFL04 = oCFLs04.Add(oCFLCreationParams04)

    oColumn04.ChooseFromListUID = "CFLITEMCD04"
    oColumn04.ChooseFromListAlias = "ItemCode"
    
    
    Exit Sub
PS_PP006_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP006_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP006_FormItemEnabled()
On Error GoTo PS_PP006_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_PP006_FormClear '//UDO방식
'        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("Rad01").Enabled = True
        oForm01.Items("Rad02").Enabled = True
        oForm01.Mode = fm_ADD_MODE
'        oForm01.Items("Btn01").Specific.Value = "1"
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
'        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("Rad01").Enabled = True
        oForm01.Items("Rad02").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("Rad01").Enabled = True
        oForm01.Items("Rad02").Enabled = True
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP006_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP006_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP006_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP006_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP006H.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP006H.Offset = oRow
    oDS_PS_PP006H.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP006_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP006_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP006_FormClear()
On Error GoTo PS_PP006_FormClear_Error
'    Dim DocEntry As String
'    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP006'", "")
'    If DocEntry = 0 Then
'        oForm01.Items("DocEntry").Specific.Value = 1
'    Else
'        oForm01.Items("DocEntry").Specific.Value = DocEntry
'    End If
    Exit Sub
PS_PP006_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP006_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP006_EnableMenus()
On Error GoTo PS_PP006_EnableMenus_Error
'    //메뉴활성화
     Call oForm01.EnableMenu("1281", False)
     Call oForm01.EnableMenu("1282", True)
     Call oForm01.EnableMenu("1293", True)
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    '//Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP006_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP006_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP006_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP006_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP006_FormItemEnabled
        Call PS_PP006_AddMatrixRow(0, True)      '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP006_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP006_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP006_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP006_DataValidCheck() As Boolean
On Error GoTo PS_PP006_DataValidCheck_Error
    PS_PP006_DataValidCheck = False
    Dim i As Long
    Dim j As Long
    Dim Query01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    If oMat01.VisualRowCount = 0 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP006_DataValidCheck = False
        Exit Function
    End If
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        j = oMat01.VisualRowCount - 1
    ElseIf (oForm01.Mode = fm_UPDATE_MODE) Then
        j = oMat01.VisualRowCount - 1
    End If
    
    For i = 1 To j
        If (oMat01.Columns("eCardCod").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "외주 거래처 코드는 필수입니다.", bmt_Short, True
            oMat01.Columns("eCardNam").Cells(i).Click ct_Regular
            PS_PP006_DataValidCheck = False
            Exit Function
        ElseIf (oMat01.Columns("ItemCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품목 코드는 필수입니다.", bmt_Short, True
            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
            PS_PP006_DataValidCheck = False
            Exit Function
        ElseIf (oMat01.Columns("Cprice").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "계약 단가는 필수입니다.", bmt_Short, True
            oMat01.Columns("Cprice").Cells(i).Click ct_Regular
            PS_PP006_DataValidCheck = False
            Exit Function
        ElseIf (oMat01.Columns("CtrDate").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "계약 일자는 필수입니다.", bmt_Short, True
            oMat01.Columns("CtrDate").Cells(i).Click ct_Regular
            PS_PP006_DataValidCheck = False
            Exit Function
        End If
    Next
    
    Call oDS_PS_PP006H.RemoveRecord(oDS_PS_PP006H.Size - 1)
    Call oMat01.LoadFromDataSource
    
    ' 데이터 저장
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP006_SetBaseForm01
    End If
    
    ' 데이터 수정
    If (oForm01.Mode = fm_UPDATE_MODE) Then
        Call PS_PP006_SetBaseForm02
    End If
    
    PS_PP006_DataValidCheck = True
    Exit Function
PS_PP006_DataValidCheck_Error:
    PS_PP006_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP006_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP006_MTX01()
On Error GoTo PS_PP006_MTX01_Error
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
  
    Param01 = oForm01.Items("CardCode").Specific.VALUE
    Param02 = oForm01.Items("ItemCode").Specific.VALUE
    Param03 = oForm01.Items("CtrDate").Specific.VALUE
    
    If oForm01.DataSources.UserDataSources("Rad01").VALUE = "A" Then
        Param04 = "A"
    ElseIf oForm01.DataSources.UserDataSources("Rad01").VALUE = "B" Then
        Param04 = "B"
    End If
    
    Query01 = "EXEC PS_PP006_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "'"
    Call RecordSet01.DoQuery(Query01)

    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP006_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_PP006H.InsertRecord (i)
        End If
        oDS_PS_PP006H.Offset = i
        Call oDS_PS_PP006H.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP006H.setValue("Code", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_PP006H.setValue("U_eCardCod", i, RecordSet01.Fields(1).VALUE)
        Call oDS_PS_PP006H.setValue("U_eCardNam", i, RecordSet01.Fields(2).VALUE)
        Call oDS_PS_PP006H.setValue("U_ItemCode", i, RecordSet01.Fields(3).VALUE)
        Call oDS_PS_PP006H.setValue("U_ItemName", i, RecordSet01.Fields(4).VALUE)
        Call oDS_PS_PP006H.setValue("U_Unit", i, RecordSet01.Fields(5).VALUE)
        Call oDS_PS_PP006H.setValue("U_Cprice", i, RecordSet01.Fields(6).VALUE)
        Call oDS_PS_PP006H.setValue("U_CtrDate", i, RecordSet01.Fields(7).VALUE)
        Call oDS_PS_PP006H.setValue("U_CpCode", i, RecordSet01.Fields(8).VALUE)
        Call oDS_PS_PP006H.setValue("U_CpName", i, RecordSet01.Fields(9).VALUE)
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    
    oForm01.Mode = fm_UPDATE_MODE
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update      ' Reloads the form from the application.
                        ' Use this method to update an object from the database.
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Call PS_PP006_AddMatrixRow(oMat01.VisualRowCount, False)      '//UDO방식일때
    
    Exit Sub
PS_PP006_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_PP006_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP006_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_PP006_FormResize()
On Error GoTo PS_PP006_FormResize_Error

    Exit Sub
PS_PP006_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP006_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP006_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_PP006_Validate_Error
    PS_PP006_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
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
'                    PS_PP006_Validate = False
'                    GoTo PS_PP006_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
'        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
'            Else
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(oForm01.Items("8").Specific.Value) & "' AND U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할수 없습니다.", "W"
'                    PS_PP006_Validate = False
'                    GoTo PS_PP006_Validate_Exit
'                End If
'            End If
'        End If
    ElseIf ValidateType = "취소" Then
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
'                PS_PP006_Validate = False
'                GoTo PS_PP006_Validate_Exit
'            End If
'            RecordSet01.MoveNext
'        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_PP006_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_PP006_Validate_Error:
    PS_PP006_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_PP006_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP006_SetBaseForm01()
On Error GoTo PS_PP006_SetBaseForm01_Error
    Dim i As Long
    Dim ItemCode01 As String

    Dim Query01 As String
    Dim Query02 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim RecordSet02 As SAPbobsCOM.Recordset
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim Param08 As Double
    Dim Param09 As String
    Dim Param10 As String
    Dim Param11 As String
    
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
 
        For i = 1 To oMat01.RowCount '//품목선택품목
            Query01 = "SELECT ISNULL(MAX(Convert(Numeric, Code)), 0) + 1 FROM [@PS_PP006H]"
            Call RecordSet01.DoQuery(Query01)
            
            Param01 = Trim(RecordSet01.Fields(0).VALUE)
            Param02 = Trim(RecordSet01.Fields(0).VALUE)
            Param03 = oMat01.Columns("eCardCod").Cells(i).Specific.VALUE
            Param04 = oMat01.Columns("eCardNam").Cells(i).Specific.VALUE
            Param05 = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
            Param06 = oMat01.Columns("ItemName").Cells(i).Specific.VALUE
            Param07 = oMat01.Columns("Unit").Cells(i).Specific.VALUE
            Param08 = oMat01.Columns("Cprice").Cells(i).Specific.VALUE
            Param09 = oMat01.Columns("CtrDate").Cells(i).Specific.VALUE
            Param10 = oMat01.Columns("CpCode").Cells(i).Specific.VALUE
            Param11 = oMat01.Columns("CpName").Cells(i).Specific.VALUE
            Query02 = "EXEC PS_PP006_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "', '" & Param08 & "', '" & Param09 & "', '" & Param10 & "', '" & Param11 & "'"
            Call RecordSet02.DoQuery(Query02)
        Next
        Call MDC_Com.MDC_GF_Message("데이터를 추가하였습니다.", "S")
        Call oDS_PS_PP006H.RemoveRecord(oDS_PS_PP006H.Size - 1)
        Call oMat01.FlushToDataSource
        Call oMat01.LoadFromDataSource
        Call oMat01.Clear
        Call PS_PP006_AddMatrixRow(0, True)      '//UDO방식일때
        
    Exit Sub
PS_PP006_SetBaseForm01_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP006_SetBaseForm01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP006_SetBaseForm02()
On Error GoTo PS_PP006_SetBaseForm02_Error
    Dim i As Long
    Dim ItemCode01 As String

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As Double
    Dim Param08 As String
    Dim Param09 As String
    Dim Param10 As String
    
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'    Query01 = "SELECT MAX(Code) + 1 FROM [@PS_PP006H]"
'    Call RecordSet01.DoQuery(Query01)
 
        For i = 1 To oMat01.RowCount '//품목선택품목
            Param01 = oMat01.Columns("Code").Cells(i).Specific.VALUE
            Param02 = oMat01.Columns("eCardCod").Cells(i).Specific.VALUE
            Param03 = oMat01.Columns("eCardNam").Cells(i).Specific.VALUE
            Param04 = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
            Param05 = oMat01.Columns("ItemName").Cells(i).Specific.VALUE
            Param06 = oMat01.Columns("Unit").Cells(i).Specific.VALUE
            Param07 = oMat01.Columns("Cprice").Cells(i).Specific.VALUE
            Param08 = oMat01.Columns("CtrDate").Cells(i).Specific.VALUE
            Param09 = oMat01.Columns("CpCode").Cells(i).Specific.VALUE
            Param10 = oMat01.Columns("CpName").Cells(i).Specific.VALUE
            Query01 = "EXEC PS_PP006_03 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', "
            Query01 = Query01 & "'" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "', '" & Param08 & "', '" & Param09 & "', '" & Param10 & "'"
            Call RecordSet01.DoQuery(Query01)
            Call MDC_Com.MDC_GF_Message("데이터를 수정 하였습니다.", "S")
        Next
        oMat01.FlushToDataSource
        oMat01.LoadFromDataSource
    Exit Sub
PS_PP006_SetBaseForm02_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP006_SetBaseForm02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

