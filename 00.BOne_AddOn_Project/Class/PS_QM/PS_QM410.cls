VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_QM410.cls
'//  Module : 품질관리 > 기계공구품질부적합관리
'//  Desc : 품질검사등록
'//  FormType : PS_QM410
'//  Create Date : 2014.01.13
'//  Modified Date :
'//  Creator : Song Myounggyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM410H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_QM410L As SAPbouiCOM.DBDataSource '등록라인

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

Private oDocType01 As String
Private oDocEntry01 As String
Private oFormMode01 As SAPbouiCOM.BoFormMode

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM410.srf")
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
    
    oFormUniqueID01 = "PS_QM410_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM410_CreateItems
    Call PS_QM410_ComboBox_Setting
    Call PS_QM410_CF_ChooseFromList
    Call PS_QM410_EnableMenus
    Call PS_QM410_SetDocument(oFromDocEntry01)
    Call PS_QM410_FormResize
    
    oForm01.Update
    oForm01.Freeze False
    
    Call oForm01.Items("DrawNo").Click
    
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

    Dim i As Integer
    Dim DocEntry As Integer
    Dim ChkSeq As Integer '검사순번

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
                Call PS_QM410_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_QM410_FormItemEnabled '//UDO방식
                Call PS_QM410_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_QM410_FormItemEnabled
                
            Case "1287" '복제
            
                oForm01.Freeze True
                
                DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM410'", "")
                ChkSeq = MDC_GetData.Get_ReData("MAX(U_ChkSeq)", "U_ItemCode", "[@PS_QM410H]", "'" & oForm01.Items("ItemCode").Specific.VALUE & "'", "AND Status = 'O'")
'                Call oDS_PS_QM400H.setValue("DocEntry", 0, DocEntry)
                oForm01.Items("DocEntry").Specific.VALUE = DocEntry
                oForm01.Items("ChkSeq").Specific.VALUE = ChkSeq + 1
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    Call oDS_PS_QM410H.setValue("DocEntry", i, DocEntry)
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
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
        If pval.ItemUID = "PS_QM410" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
            
'                If PS_QM410_DataExistCheck("A") = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                If PS_QM410_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                
'                If PS_QM410_DataExistCheck("U") = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                If PS_QM410_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        
        ElseIf pval.ItemUID = "BtnOpen" Then
        
            If oForm01.Items("DrawNo").Specific.VALUE = "" Then
                Call MDC_Com.MDC_GF_Message("도면번호를 입력하십시오.", "E")
                BubbleEvent = False
                Exit Sub
            End If
        
            Call PS_QM410_MTX01
                    
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM410" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_QM410_FormItemEnabled
                    Call PS_QM410_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_QM410_FormItemEnabled
                End If
            End If
        ElseIf pval.ItemUID = "Mat01" Then
        
            If pval.ColUID = "Check" Then '선택 클릭시 매트릭스 플러시(2015.07.14), 하지 않을 경우 갱신 시 적용 안됨
                Call oMat01.FlushToDataSource
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
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.ColUID = "BatchNum" Then
            
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "BatchNum")
            
            End If
        
        Else
        
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "DrawNo", "") '도면번호
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '납품처
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '작번
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "MM050HNo", "") '가입고번호
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '사번
        
        End If
    
'        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "BatchNum")
    
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
'        If pval.ItemChanged = True Then
'            Call oForm01.Freeze(True)
'            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_QM410L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM410L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_QM410_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_QM410L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.Value)
'                End If
'            Else
'                If (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_QM410H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_QM410H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.Value)
'                    Call oDS_PS_QM410H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Selected.Value & "'"))
'                Else
'                    Call oDS_PS_QM410H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.Value)
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

        If pval.ItemChanged = True Then
        
            Call PS_QM410_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
        
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
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat01.SelectRow(pval.Row, True, False)
            End If
        Else
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = ""
            oLastColRow01 = 0
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
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "OrdNum" Then
                Set oTempClass = New PS_PP030
                Call oTempClass.LoadForm(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
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

    If pval.BeforeAction = True Then
    
    
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemChanged = True Then
        
            Call PS_QM410_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
        
        End If

    End If
    
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_QM410_FormItemEnabled
        Call PS_QM410_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_QM410_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "ItemCode") Then
'            Dim oDataTable01 As SAPbouiCOM.DataTable
'            Set oDataTable01 = pval.SelectedObjects
'            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_QM410H", "U_CardCode,U_CardName")
'        End If
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
'            If (PS_QM410_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_QM410L.RemoveRecord(oDS_PS_QM410L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_QM410_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_QM410L.GetValue("U_ChkPnt", oMat01.RowCount - 1)) <> "" Then
                    Call PS_QM410_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_QM410_CreateItems() As Boolean
On Error GoTo PS_QM410_CreateItems_Error
   
    Set oDS_PS_QM410H = oForm01.DataSources.DBDataSources("@PS_QM410H")
    Set oDS_PS_QM410L = oForm01.DataSources.DBDataSources("@PS_QM410L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Exit Function
PS_QM410_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM410_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_QM410_ComboBox_Setting()
On Error GoTo PS_QM410_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
'    Set oCombo = oForm01.Items("BPLId").Specific
'    sQry = "SELECT BPLId, BPLName From [OBPL]"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop

    '자체/수입
    Call oForm01.Items("Class").Specific.ValidValues.Add("%", "선택")
    Call oForm01.Items("Class").Specific.ValidValues.Add("1", "수입품검사")
    Call oForm01.Items("Class").Specific.ValidValues.Add("2", "자체품검사")
    Call oForm01.Items("Class").Specific.Select(0, psk_Index)
    
    '검사종류
    sQry = "            SELECT  U_Minor AS [Code],"
    sQry = sQry & "             U_CdName AS [Name]"
    sQry = sQry & "  FROM   [@PS_SY001L]"
    sQry = sQry & "  WHERE  Code = 'Q014'"
    
    Call oForm01.Items("ChkCls").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ChkCls").Specific, sQry, "", False, False)
    Call oForm01.Items("ChkCls").Specific.Select(0, psk_Index)
    
    '//////////매트릭스//////////S
    '검사항목
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName + '[' + ISNULL(U_RelCd,'') + ']' AS [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'Q011'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ChkPnt"), sQry)
    
    '부호(최소)
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName AS [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'Q013'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("MinMark"), sQry)
    
    '부호(최대)
    sQry = "            SELECT      U_Minor AS [Code],"
    sQry = sQry & "                 U_CdName AS [Name]"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'Q013'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("MaxMark"), sQry)
    
    '//////////매트릭스//////////E
    
    
    
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
        

    Exit Sub
PS_QM410_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM410_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM410_CF_ChooseFromList()
On Error GoTo PS_QM410_CF_ChooseFromList_Error
'    '//ChooseFromList 설정
'    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
'    Dim oCons               As SAPbouiCOM.Conditions
'    Dim oCon                As SAPbouiCOM.Condition
'    Dim oCFL                As SAPbouiCOM.ChooseFromList
'    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
'    Dim oEdit               As SAPbouiCOM.EditText
'    Dim oColumn             As SAPbouiCOM.Column
'
'    Set oEdit = oForm01.Items("CARDCODE").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_BusinessPartner
'    oCFLCreationParams.uniqueID = "CFLCARDCODE"
'    oCFLCreationParams.MultiSelection = False
''    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLCARDCODE"
'    oEdit.ChooseFromListAlias = "CardCode"
'
'    Set oEdit = oForm01.Items("CARDNAME").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_BusinessPartner
'    oCFLCreationParams.uniqueID = "CFLCARDNAME"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLCARDNAME"
'    oEdit.ChooseFromListAlias = "CardName"
    Exit Sub
PS_QM410_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM410_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM410_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_QM410_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer

    Dim StdPnt As String '기준치수
    Dim MinPnt As String '공차(최소)
    Dim MaxPnt As String '공차(최대)
    Dim InputData As String '입력한 측정값
    Dim CurCol As String '현재 컬럼명
    Dim MinMark As String
    Dim MaxMark As String
    Dim ChkPnt As String '검사항목
        
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String
    Dim TeamCode As String
    Dim ChkSeq As String
    Dim sQry As String
'    Dim RspCode As String
    
    Select Case oUID
        
        Case "Mat01"
        
            Call oForm01.Freeze(True)
        
            Call oMat01.FlushToDataSource
        
            If oMat01.RowCount = oRow And Trim(oDS_PS_QM410L.GetValue("U_ChkPnt", oRow - 1)) <> "" Then 'And Trim(oDS_PS_QM410L.GetValue("U_StdPnt", oRow - 1)) <> "" And Trim(oDS_PS_QM410L.GetValue("U_MinPnt", oRow - 1)) <> "" And Trim(oDS_PS_QM410L.GetValue("U_MaxPnt", oRow - 1)) <> "" Then '기준치수가 입력되었을 때 행 추가
                Call PS_QM410_AddMatrixRow(oRow) '행 추가
            End If
            
            '입력값의 범위 체크
            StdPnt = Trim(oDS_PS_QM410L.GetValue("U_StdPnt", oRow - 1)) '기준치
            MinMark = Trim(oDS_PS_QM410L.GetValue("U_MinMark", oRow - 1)) '부호(최소)
            MinPnt = Trim(oDS_PS_QM410L.GetValue("U_MinPnt", oRow - 1)) '공차(최소)
            MaxMark = Trim(oDS_PS_QM410L.GetValue("U_MaxMark", oRow - 1)) '부호(최대)
            MaxPnt = Trim(oDS_PS_QM410L.GetValue("U_MaxPnt", oRow - 1)) '공차(최대)
            InputData = Trim(oDS_PS_QM410L.GetValue("U_" & oCol, oRow - 1)) '측정치
            ChkPnt = Trim(oDS_PS_QM410L.GetValue("U_ChkPnt", oRow - 1)) '검사항목
            
            If oCol = "Data01" Or oCol = "Data02" Or oCol = "Data03" Or oCol = "Data04" Or oCol = "Data05" Or oCol = "Data06" Or oCol = "Data07" Or oCol = "Data08" Or oCol = "Data09" Or oCol = "Data10" Or oCol = "Data11" Or oCol = "Data12" Or oCol = "Data13" Or oCol = "Data14" Or oCol = "Data15" Then
            
                If InputData <> "" And ChkPnt < 90 Then '입력 값이 존재하고 제목(또는 빈행)이 아닐 때만 체크
                
                    If (PS_QM410_CheckInputData(StdPnt, MinPnt, MaxPnt, InputData, MinMark, MaxMark)) = False Then
                            
                        '메시지박스 예, 아니오
                        If Sbo_Application.MessageBox("입력한 측정치는 불량입니다. 등록하시겠습니까?", "1", "예", "아니오") = "1" Then
                        Else
    '                        Call oDS_PS_QM410L.setValue("U_" & oCol, oRow - 1, "") '빈값 세팅
    '                        Call oMat01.LoadFromDataSource
                            oMat01.Columns(oCol).Cells(oRow).Specific.VALUE = "" '빈값 세팅
                        End If
                            
                    End If
                
                End If
                
            End If
            
            oMat01.AutoResizeColumns

            Call oForm01.Freeze(False)

        Case "CardCode"
        
            Call oDS_PS_QM410H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oDS_PS_QM410H.GetValue("U_CardCode", 0) & "'")) '납품처
            
        Case "CntcCode"
        
            Call oDS_PS_QM410H.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oDS_PS_QM410H.GetValue("U_CntcCode", 0) & "'")) '검사자성명
            
        Case "ItemCode"
        
            Call oDS_PS_QM410H.setValue("U_ItemName", 0, MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oDS_PS_QM410H.GetValue("U_ItemCode", 0) & "'")) '작번
            
            '검사순번 조회
            sQry = "           SELECT      ISNULL(MAX(U_ChkSeq), 0) AS [ChkSeq]"
            sQry = sQry & " FROM       [@PS_QM410H]"
            sQry = sQry & " WHERE     U_ItemCode = '" & Trim(oDS_PS_QM410H.GetValue("U_ItemCode", 0)) & "'"
            sQry = sQry & "                AND Status = 'O'"
            
            Call oRecordSet01.DoQuery(sQry)
            
            If oRecordSet01.Fields("ChkSeq").VALUE = 0 Then
                
                ChkSeq = 1
            
            Else
            
                ChkSeq = oRecordSet01.Fields("ChkSeq").VALUE + 1
            
            End If
        
            Call oDS_PS_QM410H.setValue("U_ChkSeq", 0, ChkSeq)
            '검사순번 조회
            
            '납품처 자동등록
            sQry = "           SELECT       T0.CardCode,"
            sQry = sQry & "                 T0.CardName"
            sQry = sQry & " FROM        ORDR AS T0"
            sQry = sQry & "                 INNER JOIN"
            sQry = sQry & "                 RDR1 AS T1"
            sQry = sQry & "                     ON T0.DocEntry = T1.DocEntry"
            sQry = sQry & " WHERE       T0.Canceled = 'N'"
            sQry = sQry & "                 AND T1.ItemCode = '" & Trim(oDS_PS_QM410H.GetValue("U_ItemCode", 0)) & "'"
            
            Call oRecordSet01.DoQuery(sQry)
            
            Call oDS_PS_QM410H.setValue("U_CardCode", 0, oRecordSet01.Fields("CardCode").VALUE) '거래처코드
            Call oDS_PS_QM410H.setValue("U_CardName", 0, oRecordSet01.Fields("CardName").VALUE) '거래처명
            '납품처 자동등록
        
        Case "TotalQty"
        
            Call PS_QM410_CalculateQty(oDS_PS_QM410H.GetValue("U_TotalQty", 0), oDS_PS_QM410H.GetValue("U_PassQty", 0))
        
        Case "PassQty"
        
            Call PS_QM410_CalculateQty(oDS_PS_QM410H.GetValue("U_TotalQty", 0), oDS_PS_QM410H.GetValue("U_PassQty", 0))
        
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_QM410_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
'        Call MDC_Com.MDC_GF_Message("해당 측정치는 불량 측정치입니다.", "E")
'    ElseIf ErrNum = 2 Then
'        Call MDC_Com.MDC_GF_Message("시각(시)는 24미만의 값만 입력이 가능합니다.", "E")
'    ElseIf ErrNum = 3 Then
'        Call MDC_Com.MDC_GF_Message("시각(분)은 60미만의 값만 입력이 가능합니다.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_QM410_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Sub PS_QM410_FormItemEnabled()
On Error GoTo PS_QM410_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
'
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        Call PS_QM410_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        
'        Call oForm01.Items("Class").Specific.Select(0, psk_Index)
        Call oForm01.Items("DrawNo").Click
                
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
'        Call oForm01.Items("Class").Specific.Select(0, psk_Index)
        Call oForm01.Items("DrawNo").Click
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        
    End If
    Call oMat01.AutoResizeColumns
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM410_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM410_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM410_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM410_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM410L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_QM410L.Offset = oRow
    Call oDS_PS_QM410L.setValue("U_LineNum", oRow, oRow + 1)
    Call oDS_PS_QM410L.setValue("U_Check", oRow, "Y")
    
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM410_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM410_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM410_FormClear()
On Error GoTo PS_QM410_FormClear_Error

    Dim DocEntry As String
    
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM410'", "")
    
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    
    Call oDS_PS_QM410H.setValue("U_ChkDate", 0, Format(Now, "YYYYMMDD"))
    
    Exit Sub
PS_QM410_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM410_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM410_EnableMenus()
On Error GoTo PS_QM410_EnableMenus_Error
    '//메뉴활성화
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), True        '// 닫기
    oForm01.EnableMenu ("1287"), True        '// 복제
    oForm01.EnableMenu ("1285"), True        '// 복원
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True
    Exit Sub
PS_QM410_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM410_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM410_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM410_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM410_FormItemEnabled
        Call PS_QM410_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_QM410_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_QM410_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM410_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_QM410_DataValidCheck() As Boolean
On Error GoTo PS_QM410_DataValidCheck_Error
    PS_QM410_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_QM410_FormClear
    End If
    
    '합격수량, 전체수량 체크
    If CDbl(oDS_PS_QM410H.GetValue("U_TotalQty", 0)) < CDbl(oDS_PS_QM410H.GetValue("U_PassQty", 0)) Then
        Call Sbo_Application.SetStatusBarMessage("합격수량이 전체수량보다 많습니다. 확인하세요.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    '도면번호 미입력 시
    If oForm01.Items("DrawNo").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("도면번호가 입력되지 않았습니다.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    '자체/수입
    If oForm01.Items("Class").Specific.VALUE = "%" Then
        Call Sbo_Application.SetStatusBarMessage("자체/수입이 선택되지 않았습니다.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    '납품처
    If oForm01.Items("CardCode").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("납품처가 입력되지 않았습니다.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    '검사년월일
    If oForm01.Items("ChkDate").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("검사일자가 입력되지 않았습니다.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    '전체수량
    If oForm01.Items("TotalQty").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("전체수량이 입력되지 않았습니다.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    '합격수량
    If oForm01.Items("PassQty").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("합격수량이 입력되지 않았습니다.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    '불합격수량
    If oForm01.Items("BadQty").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("불합격수량이 입력되지 않았습니다.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    '검사자
    If oForm01.Items("CntcCode").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("검사자정보가 입력되지 않았습니다.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    '수입품검사일 경우
    If oForm01.Items("Class").Specific.VALUE = "1" Then
        '가입고번호가 없으면
        If oForm01.Items("MM050HNo").Specific.VALUE = "" Then
            '선검사사유 필수
            If oForm01.Items("PreChkNt").Specific.VALUE = "" Then
                Call Sbo_Application.SetStatusBarMessage("[가입고번호]를 입력하지 않을 경우 [선검사사유]는 필수입니다.", bmt_Short, True)
                Call oForm01.Items("PreChkNt").Click
                PS_QM410_DataValidCheck = False
                Exit Function
            End If
        End If
    End If
    
    '검사종류
        If oForm01.Items("ChkCls").Specific.VALUE = "%" Then
        Call Sbo_Application.SetStatusBarMessage("검사종류가 선택되지 않았습니다.", bmt_Short, True)
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
   
    '라인정보 미입력 시
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM410_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        
        If oMat01.Columns("ChkPnt").Cells(i).Specific.Selected.VALUE < 90 Then '제목 및 빈행은 체크하지 않음
        
            If (oMat01.Columns("ChkPnt").Cells(i).Specific.Selected.VALUE = "") Then
                Sbo_Application.SetStatusBarMessage "검사항목은 필수입니다.", bmt_Short, True
                oMat01.Columns("ChkPnt").Cells(i).Click ct_Regular
                PS_QM410_DataValidCheck = False
                Exit Function
            End If
            
            If (oMat01.Columns("StdPnt").Cells(i).Specific.VALUE = "") Then
                Sbo_Application.SetStatusBarMessage "기준치수는 필수입니다.", bmt_Short, True
                oMat01.Columns("StdPnt").Cells(i).Click ct_Regular
                PS_QM410_DataValidCheck = False
                Exit Function
            End If
            
            If (oMat01.Columns("MinPnt").Cells(i).Specific.VALUE = "") Then
                Sbo_Application.SetStatusBarMessage "최소공차는 필수입니다.", bmt_Short, True
                oMat01.Columns("MinPnt").Cells(i).Click ct_Regular
                PS_QM410_DataValidCheck = False
                Exit Function
            End If
            
            If (oMat01.Columns("MaxPnt").Cells(i).Specific.VALUE = "") Then
                Sbo_Application.SetStatusBarMessage "최대공차는 필수입니다.", bmt_Short, True
                oMat01.Columns("MaxPnt").Cells(i).Click ct_Regular
                PS_QM410_DataValidCheck = False
                Exit Function
            End If
        End If

    Next
    
    Call oDS_PS_QM410L.RemoveRecord(oDS_PS_QM410L.Size - 1)
    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_QM410_FormClear
    End If
    
    PS_QM410_DataValidCheck = True
    Exit Function
PS_QM410_DataValidCheck_Error:
    PS_QM410_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_QM410_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_QM410_MTX01()
On Error GoTo PS_QM410_MTX01_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DrawNo As String
    Dim ReviNo As String

    DrawNo = oForm01.Items("DrawNo").Specific.VALUE
    ReviNo = oForm01.Items("ReviNo").Specific.VALUE
    
    Query01 = "              EXEC PS_QM410_01 '"
    Query01 = Query01 & DrawNo & "', '"
    Query01 = Query01 & ReviNo & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("검사기준이 존재하지 않습니다.", "E")
        GoTo PS_QM410_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_QM410L.InsertRecord (i)
        End If
        oDS_PS_QM410L.Offset = i
        Call oDS_PS_QM410L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_QM410L.setValue("U_Check", i, "Y") '선택
        Call oDS_PS_QM410L.setValue("U_ChkPnt", i, RecordSet01.Fields("ChkPnt").VALUE) '검사항목
        Call oDS_PS_QM410L.setValue("U_StdPnt", i, RecordSet01.Fields("StdPnt").VALUE) '기준치수
        Call oDS_PS_QM410L.setValue("U_MinMark", i, RecordSet01.Fields("MinMark").VALUE) '부호(최소)
        Call oDS_PS_QM410L.setValue("U_MinPnt", i, RecordSet01.Fields("MinPnt").VALUE) '공차(최소)
        Call oDS_PS_QM410L.setValue("U_MaxMark", i, RecordSet01.Fields("MaxMark").VALUE) '부호(최대)
        Call oDS_PS_QM410L.setValue("U_MaxPnt", i, RecordSet01.Fields("MaxPnt").VALUE) '공차(최대)
        Call oDS_PS_QM410L.setValue("U_Data01", i, RecordSet01.Fields("Data01").VALUE) 'Data01
        Call oDS_PS_QM410L.setValue("U_Data02", i, RecordSet01.Fields("Data02").VALUE) 'Data02
        Call oDS_PS_QM410L.setValue("U_Data03", i, RecordSet01.Fields("Data03").VALUE) 'Data03
        Call oDS_PS_QM410L.setValue("U_Data04", i, RecordSet01.Fields("Data04").VALUE) 'Data04
        Call oDS_PS_QM410L.setValue("U_Data05", i, RecordSet01.Fields("Data05").VALUE) 'Data05
        Call oDS_PS_QM410L.setValue("U_Data06", i, RecordSet01.Fields("Data06").VALUE) 'Data06
        Call oDS_PS_QM410L.setValue("U_Data07", i, RecordSet01.Fields("Data07").VALUE) 'Data07
        Call oDS_PS_QM410L.setValue("U_Data08", i, RecordSet01.Fields("Data08").VALUE) 'Data08
        Call oDS_PS_QM410L.setValue("U_Data09", i, RecordSet01.Fields("Data09").VALUE) 'Data09
        Call oDS_PS_QM410L.setValue("U_Data10", i, RecordSet01.Fields("Data10").VALUE) 'Data10
        Call oDS_PS_QM410L.setValue("U_Data11", i, RecordSet01.Fields("Data11").VALUE) 'Data11
        Call oDS_PS_QM410L.setValue("U_Data12", i, RecordSet01.Fields("Data12").VALUE) 'Data12
        Call oDS_PS_QM410L.setValue("U_Data13", i, RecordSet01.Fields("Data13").VALUE) 'Data13
        Call oDS_PS_QM410L.setValue("U_Data14", i, RecordSet01.Fields("Data14").VALUE) 'Data14
        Call oDS_PS_QM410L.setValue("U_Data15", i, RecordSet01.Fields("Data15").VALUE) 'Data15
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    Call PS_QM410_AddMatrixRow(oMat01.VisualRowCount) '행 추가
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM410_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_QM410_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM410_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_QM410_DI_API() As Boolean
On Error GoTo PS_QM410_DI_API_Error
    PS_QM410_DI_API = True
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
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.VALUE, "&&&&-&&-&&")
    For i = 0 To ItemInformationCount - 1
        If ItemInformation(i).Check = True Then
            GoTo Continue_First
        End If
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
        oDIObject.Lines.BaseType = "22"
        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
        For j = i To UBound(ItemInformation)
            If ItemInformation(j).Check = True Then
                GoTo Continue_Second
            End If
            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
                GoTo Continue_Second
            End If
            '//같은것
            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.Add
            ItemInformation(j).PDN1No = LineNumCount
            ItemInformation(j).Check = True
Continue_Second:
        Next
        LineNumCount = LineNumCount + 1
Continue_First:
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To UBound(ItemInformation)
            Call oDS_PS_QM410L.setValue("U_OPDNNo", i, ResultDocNum)
            Call oDS_PS_QM410L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        Next
    Else
        GoTo PS_QM410_DI_API_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_QM410_DI_API_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_QM410_DI_API = False
    Set oDIObject = Nothing
    Exit Function
PS_QM410_DI_API_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM410_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_QM410_DI_API = False
End Function

Private Sub PS_QM410_FormResize()
On Error GoTo PS_QM410_FormResize_Error
    
    Exit Sub
PS_QM410_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM410_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_QM410_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_QM410_Validate_Error
    PS_QM410_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_QM410H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
'        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
'        PS_QM410_Validate = False
'        GoTo PS_QM410_Validate_Exit
'    End If
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
'                    PS_QM410_Validate = False
'                    GoTo PS_QM410_Validate_Exit
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
'                    PS_QM410_Validate = False
'                    GoTo PS_QM410_Validate_Exit
'                End If
'            End If
'        End If
    ElseIf ValidateType = "취소" Then
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
'                PS_QM410_Validate = False
'                GoTo PS_QM410_Validate_Exit
'            End If
'            RecordSet01.MoveNext
'        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_QM410_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_QM410_Validate_Error:
    PS_QM410_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_QM410_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PS_QM410_DataExistCheck(ByVal pMode As String) As Boolean
'******************************************************************************
'Function ID : PS_QM410_DataExistCheck()
'해당모듈 : PS_QM410
'기능 : 동일 도번의 2중 등록을 방지하기 위한 체크
'인수 : pMode : 문서의 모드, U:갱신, A:추가
'반환값 : True : 동일한 데이터 없음, False : 동일한 데이터 존재함
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM410_DataExistCheck_Error

    Dim i As Long
    Dim Query01 As String
    Dim ErrNum As String
    Dim DocEntry As Integer
    Dim CurDocEntry As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DrawNo As String
    
    DrawNo = oForm01.Items("DrawNo").Specific.VALUE
    
    If pMode = "U" Then
        CurDocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    End If
    
    Query01 = "                SELECT   DocEntry"
    Query01 = Query01 & " FROM     [@PS_QM410H]"
    Query01 = Query01 & " WHERE   U_DrawNo = '" & DrawNo & "'"
    If pMode = "U" Then
        Query01 = Query01 & "         AND DocEntry <> " & CurDocEntry
    End If
    Call RecordSet01.DoQuery(Query01)
    
    If RecordSet01.RecordCount <> 0 Then
        ErrNum = 1
        DocEntry = Trim(RecordSet01.Fields("DocEntry").VALUE)
        GoTo PS_QM410_DataExistCheck_Error
    End If
    
    PS_QM410_DataExistCheck = True
    Set RecordSet01 = Nothing
    
    Exit Function
PS_QM410_DataExistCheck_Error:
    
    Set RecordSet01 = Nothing
    PS_QM410_DataExistCheck = False
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("동일 도번으로 등록된 데이터가 존재합니다. 확인하세요. 문서번호 : " & DocEntry, "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_QM410_DataExistCheck_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
    
End Function

Private Sub PS_QM410_CalculateQty(ByVal pTotalQty As Double, ByVal pPassQty As Double)
'******************************************************************************
'Function ID : PS_QM410_CalculateQty()
'해당모듈 : PS_QM410
'기능 : 불량수량 계산
'인수 : pTotalQty : 전체수량, pPassQty : 합격수량
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM410_CalculateQty_Error

    Dim ErrNum As String
    
    If pTotalQty < pPassQty Then
    
        ErrNum = "1"
        GoTo PS_QM410_CalculateQty_Error
    
    End If
    
    Call oDS_PS_QM410H.setValue("U_BadQty", 0, pTotalQty - pPassQty)
    
    Exit Sub
PS_QM410_CalculateQty_Error:

    If ErrNum = "1" Then
        Call MDC_Com.MDC_GF_Message("합격수량이 전체수량보다 큽니다. 확인하세요.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_QM410_CalculateQty_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
    
End Sub

Private Function PS_QM410_CheckInputData(ByVal pStdPnt As String, ByVal pMinPnt As String, ByVal pMaxPnt As String, ByVal pInputData As String, ByVal pMinMark As String, ByVal pMaxMark As String)
'******************************************************************************
'Function ID : PS_QM410_CheckInputData()
'해당모듈 : PS_QM410
'기능 : 입력한 측정치수의 합격불합격 여부 체크
'인수 : pStdPnt : 기준치, pMinPnt : 공차(최소), pMaxPnt : 공차(최대), pInputData : 입력한 측정치, pMinMark : 부호(최소), pMaxMark : 부호(최대)
'반환값 : True : 합격, False : 불합격
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM410_CalculateQty_Error

    Dim ErrNum As String
    Dim MinData As String '최소기준치
    Dim MaxData As String '최대기준치
'    Dim MinMark As String '부호(최소)
'    Dim MaxMark As String '부호(최대)
    
    '최소기준치 계산
    If pMinMark = "03" Then '03 : +
        MinData = CDbl(pStdPnt) + CDbl(pMinPnt)
    ElseIf pMinMark = "02" Then '02 : -
        MinData = CDbl(pStdPnt) - CDbl(pMinPnt)
    Else
        MinData = CDbl(pStdPnt)
    End If

    '최대기준치 계산
    If pMaxMark = "03" Then '03 : +
        MaxData = CDbl(pStdPnt) + CDbl(pMaxPnt)
    ElseIf pMaxMark = "02" Then '02 : -
        MaxData = CDbl(pStdPnt) - CDbl(pMaxPnt)
    Else
        MaxData = CDbl(pStdPnt)
    End If

    If pInputData < MinData Or pInputData > MaxData Then
        PS_QM410_CheckInputData = False
    Else
        PS_QM410_CheckInputData = True
    End If
    
    Exit Function
PS_QM410_CalculateQty_Error:
    PS_QM410_CheckInputData = False
    If ErrNum = "1" Then
        
    Else
        Call MDC_Com.MDC_GF_Message("PS_QM410_CalculateQty_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
    
End Function
