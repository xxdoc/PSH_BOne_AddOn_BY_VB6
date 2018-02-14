VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP004.cls
'//  Module         : 생산관리 > 품목별표준공정등록
'//  Desc           : 기본
'//  FormType       : PS_PP004
'//  Create Date    : 2010.08.20
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix
Public oMat03 As SAPbouiCOM.Matrix
Private oDS_PS_USERDS01 As SAPbouiCOM.DBDataSource '품목
Private oDS_PS_USERDS02 As SAPbouiCOM.DBDataSource '공정
Private oDS_PS_USERDS03 As SAPbouiCOM.DBDataSource '공정추가

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMat01Row01 As Long
Private oMat02Row02 As Long
Private oMat03Row03 As Long

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP004.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP004_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP004_CreateItems
    Call PS_PP004_ComboBox_Setting
    Call PS_PP004_CF_ChooseFromList
    Call PS_PP004_EnableMenus
    Call PS_PP004_SetDocument(oFromDocEntry01)
    Call PS_PP004_FormResize
    
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
                '//Call PS_PP004_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_PP004_FormItemEnabled '//UDO방식
                '//Call PS_PP004_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_PP004_FormItemEnabled
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
        If pval.ItemUID = "Mat03" Then
            Call oForm01.EnableMenu("1293", True)
        Else
            Call oForm01.EnableMenu("1293", False)
        End If
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
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Or pval.ItemUID = "Mat03" Then
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
        If pval.ItemUID = "PS_PP004" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_PP004_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                Call PS_PP004_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If (PS_PP004_Process_Register = False) Then
                    BubbleEvent = False
                    Exit Sub
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button03" Then '//위로
            Dim Sequence As Long
            Dim ItemCode As String
            Dim ItemName As String
            Dim CpBCode As String
            Dim CpBName As String
            Dim CpCode As String
            Dim CpName As String
            Dim CpUnWt As Double
            Dim ResultYN As String
            Dim ReportYN As String
            Dim DayProYN As String
            Dim CurrentMat03Row As Long
            If oForm01.Mode = fm_ADD_MODE Then
                If oMat03.GetNextSelectedRow(0, ot_RowOrder) = -1 Then '//행이선택이 안되었을때
                    Call MDC_Com.MDC_GF_Message("행이선택되지 않았습니다.", "W")
                ElseIf oMat03.GetNextSelectedRow(0, ot_RowOrder) = 1 Then '//첫번째행일때
                    '//아무행위안함
                Else
                    CurrentMat03Row = oMat03Row03
                    Sequence = oMat03.Columns("Sequence").Cells(CurrentMat03Row).Specific.VALUE
                    ItemCode = oMat03.Columns("ItemCode").Cells(CurrentMat03Row).Specific.VALUE
                    ItemName = oMat03.Columns("ItemName").Cells(CurrentMat03Row).Specific.VALUE
                    CpCode = oMat03.Columns("CpCode").Cells(CurrentMat03Row).Specific.VALUE
                    CpName = oMat03.Columns("CpName").Cells(CurrentMat03Row).Specific.VALUE
                    CpBCode = oMat03.Columns("CpBCode").Cells(CurrentMat03Row).Specific.VALUE
                    CpBName = oMat03.Columns("CpBName").Cells(CurrentMat03Row).Specific.VALUE
                    CpUnWt = oMat03.Columns("CpUnWt").Cells(CurrentMat03Row).Specific.VALUE
                    ResultYN = oMat03.Columns("ResultYN").Cells(CurrentMat03Row).Specific.Selected.VALUE
                    ReportYN = oMat03.Columns("ReportYN").Cells(CurrentMat03Row).Specific.Selected.VALUE
                    DayProYN = oMat03.Columns("DayProYN").Cells(CurrentMat03Row).Specific.Selected.VALUE
                    oMat03.Columns("Sequence").Cells(CurrentMat03Row).Specific.VALUE = Val(oMat03.Columns("Sequence").Cells(CurrentMat03Row - 1).Specific.VALUE) + 1
                    oMat03.Columns("Sequence").Cells(CurrentMat03Row - 1).Specific.VALUE = Sequence - 1
                    oMat03.Columns("ItemCode").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("ItemCode").Cells(CurrentMat03Row - 1).Specific.VALUE
                    oMat03.Columns("ItemCode").Cells(CurrentMat03Row - 1).Specific.VALUE = ItemCode
                    oMat03.Columns("ItemName").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("ItemName").Cells(CurrentMat03Row - 1).Specific.VALUE
                    oMat03.Columns("ItemName").Cells(CurrentMat03Row - 1).Specific.VALUE = ItemName
                    oMat03.Columns("CpBCode").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpBCode").Cells(CurrentMat03Row - 1).Specific.VALUE
                    oMat03.Columns("CpBCode").Cells(CurrentMat03Row - 1).Specific.VALUE = CpBCode
                    oMat03.Columns("CpBName").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpBName").Cells(CurrentMat03Row - 1).Specific.VALUE
                    oMat03.Columns("CpBName").Cells(CurrentMat03Row - 1).Specific.VALUE = CpBName
                    oMat03.Columns("CpCode").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpCode").Cells(CurrentMat03Row - 1).Specific.VALUE
                    oMat03.Columns("CpCode").Cells(CurrentMat03Row - 1).Specific.VALUE = CpCode
                    oMat03.Columns("CpName").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpName").Cells(CurrentMat03Row - 1).Specific.VALUE
                    oMat03.Columns("CpName").Cells(CurrentMat03Row - 1).Specific.VALUE = CpName
                    oMat03.Columns("CpUnWt").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpUnWt").Cells(CurrentMat03Row - 1).Specific.VALUE
                    oMat03.Columns("CpUnWt").Cells(CurrentMat03Row - 1).Specific.VALUE = CpUnWt
                    Call oMat03.Columns("ResultYN").Cells(CurrentMat03Row).Specific.Select(oMat03.Columns("ResultYN").Cells(CurrentMat03Row - 1).Specific.Selected.VALUE, psk_ByValue)
                    Call oMat03.Columns("ResultYN").Cells(CurrentMat03Row - 1).Specific.Select(ResultYN, psk_ByValue)
                    Call oMat03.Columns("ReportYN").Cells(CurrentMat03Row).Specific.Select(oMat03.Columns("ReportYN").Cells(CurrentMat03Row - 1).Specific.Selected.VALUE, psk_ByValue)
                    Call oMat03.Columns("ReportYN").Cells(CurrentMat03Row - 1).Specific.Select(ReportYN, psk_ByValue)
                    Call oMat03.Columns("DayProYN").Cells(CurrentMat03Row).Specific.Select(oMat03.Columns("DayProYN").Cells(CurrentMat03Row - 1).Specific.Selected.VALUE, psk_ByValue)
                    Call oMat03.Columns("DayProYN").Cells(CurrentMat03Row - 1).Specific.Select(DayProYN, psk_ByValue)
                    Call oMat03.SelectRow(CurrentMat03Row - 1, True, False)
                    oMat03Row03 = CurrentMat03Row - 1
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button04" Then '//아래로
            If oForm01.Mode = fm_ADD_MODE Then
                If oMat03.GetNextSelectedRow(0, ot_RowOrder) = -1 Then
                    Call MDC_Com.MDC_GF_Message("행이선택되지 않았습니다.", "W")
                ElseIf oMat03.GetNextSelectedRow(0, ot_RowOrder) = oMat03.VisualRowCount Then '//마지막행일때
                    '//아무행위안함
                Else
                    CurrentMat03Row = oMat03Row03
                    Sequence = oMat03.Columns("Sequence").Cells(CurrentMat03Row).Specific.VALUE
                    ItemCode = oMat03.Columns("ItemCode").Cells(CurrentMat03Row).Specific.VALUE
                    ItemName = oMat03.Columns("ItemName").Cells(CurrentMat03Row).Specific.VALUE
                    CpBCode = oMat03.Columns("CpBCode").Cells(CurrentMat03Row).Specific.VALUE
                    CpBName = oMat03.Columns("CpBName").Cells(CurrentMat03Row).Specific.VALUE
                    CpCode = oMat03.Columns("CpCode").Cells(CurrentMat03Row).Specific.VALUE
                    CpName = oMat03.Columns("CpName").Cells(CurrentMat03Row).Specific.VALUE
                    CpUnWt = oMat03.Columns("CpUnWt").Cells(CurrentMat03Row).Specific.VALUE
                    ResultYN = oMat03.Columns("ResultYN").Cells(CurrentMat03Row).Specific.Selected.VALUE
                    ReportYN = oMat03.Columns("ReportYN").Cells(CurrentMat03Row).Specific.Selected.VALUE
                    DayProYN = oMat03.Columns("DayProYN").Cells(CurrentMat03Row).Specific.Selected.VALUE
                    oMat03.Columns("Sequence").Cells(CurrentMat03Row).Specific.VALUE = Val(oMat03.Columns("Sequence").Cells(CurrentMat03Row + 1).Specific.VALUE) - 1
                    oMat03.Columns("Sequence").Cells(CurrentMat03Row + 1).Specific.VALUE = Sequence + 1
                    oMat03.Columns("ItemCode").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("ItemCode").Cells(CurrentMat03Row + 1).Specific.VALUE
                    oMat03.Columns("ItemCode").Cells(CurrentMat03Row + 1).Specific.VALUE = ItemCode
                    oMat03.Columns("ItemName").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("ItemName").Cells(CurrentMat03Row + 1).Specific.VALUE
                    oMat03.Columns("ItemName").Cells(CurrentMat03Row + 1).Specific.VALUE = ItemName
                    oMat03.Columns("CpBCode").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpBCode").Cells(CurrentMat03Row + 1).Specific.VALUE
                    oMat03.Columns("CpBCode").Cells(CurrentMat03Row + 1).Specific.VALUE = CpBCode
                    oMat03.Columns("CpBName").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpBName").Cells(CurrentMat03Row + 1).Specific.VALUE
                    oMat03.Columns("CpBName").Cells(CurrentMat03Row + 1).Specific.VALUE = CpBName
                    oMat03.Columns("CpCode").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpCode").Cells(CurrentMat03Row + 1).Specific.VALUE
                    oMat03.Columns("CpCode").Cells(CurrentMat03Row + 1).Specific.VALUE = CpCode
                    oMat03.Columns("CpName").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpName").Cells(CurrentMat03Row + 1).Specific.VALUE
                    oMat03.Columns("CpName").Cells(CurrentMat03Row + 1).Specific.VALUE = CpName
                    oMat03.Columns("CpUnWt").Cells(CurrentMat03Row).Specific.VALUE = oMat03.Columns("CpUnWt").Cells(CurrentMat03Row + 1).Specific.VALUE
                    oMat03.Columns("CpUnWt").Cells(CurrentMat03Row + 1).Specific.VALUE = CpUnWt
                    Call oMat03.Columns("ResultYN").Cells(CurrentMat03Row).Specific.Select(oMat03.Columns("ResultYN").Cells(CurrentMat03Row + 1).Specific.Selected.VALUE, psk_ByValue)
                    Call oMat03.Columns("ResultYN").Cells(CurrentMat03Row + 1).Specific.Select(ResultYN, psk_ByValue)
                    Call oMat03.Columns("ReportYN").Cells(CurrentMat03Row).Specific.Select(oMat03.Columns("ReportYN").Cells(CurrentMat03Row + 1).Specific.Selected.VALUE, psk_ByValue)
                    Call oMat03.Columns("ReportYN").Cells(CurrentMat03Row + 1).Specific.Select(ReportYN, psk_ByValue)
                    Call oMat03.Columns("DayProYN").Cells(CurrentMat03Row).Specific.Select(oMat03.Columns("DayProYN").Cells(CurrentMat03Row + 1).Specific.Selected.VALUE, psk_ByValue)
                    Call oMat03.Columns("DayProYN").Cells(CurrentMat03Row + 1).Specific.Select(DayProYN, psk_ByValue)
                    Call oMat03.SelectRow(CurrentMat03Row + 1, True, False)
                    oMat03Row03 = CurrentMat03Row + 1
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_PP004_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'                If PS_PP004_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP004" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button03" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button04" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP004_FormItemEnabled
'                    Call PS_PP004_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP004_FormItemEnabled
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
                If (pval.ColUID = "ItemCode") Then
                    
                End If
            Else
                If (pval.ItemUID = "ItmBsort") Then
                    Call MDC_SetMod.Remove_ComboList(oForm01.Items("ItmMsort").Specific)
                    oForm01.Items("ItmMsort").Specific.ValidValues.Add "", ""
                    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" & oForm01.Items(pval.ItemUID).Specific.Selected.VALUE & "' order by CONVERT(INT,U_CODE)", "", False, False)
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
        If pval.ItemUID = "Opt01" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Opt02" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat02"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Opt03" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat03"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)
                oMat01Row01 = pval.Row
                Call PS_PP004_MTX02
            End If
        End If
        If pval.ItemUID = "Mat02" Then
            If pval.Row > 0 Then
                Call oMat02.SelectRow(pval.Row, True, False)
                oMat02Row02 = pval.Row
            End If
        End If
        If pval.ItemUID = "Mat03" Then
            If pval.Row > 0 Then
                Call oMat03.SelectRow(pval.Row, True, False)
                oMat03Row03 = pval.Row
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
    Dim i As Long
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat02" Then
            If pval.Row > 0 Then
                Dim A As Long
                If (oMat01Row01 = 0) Then
                    Call MDC_Com.MDC_GF_Message("품목이 선택되지 않았습니다.", "W")
                    BubbleEvent = False
                    Exit Sub
                Else
                    For i = 1 To oMat03.VisualRowCount
                        If (oMat02.Columns("CpBCode").Cells(pval.Row).Specific.VALUE = oMat03.Columns("CpBCode").Cells(i).Specific.VALUE And oMat02.Columns("CpCode").Cells(pval.Row).Specific.VALUE = oMat03.Columns("CpCode").Cells(i).Specific.VALUE) Then
                            '//공정매트릭스와, 품목별공정결과 매트릭스에 똑같은 공정이 존재할경우
                            Call MDC_Com.MDC_GF_Message("같은 공정이 존재합니다.", "W")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                    '//멀티,엔드베어링
                    If oMat01.Columns("ItmBsort").Cells(oMat01Row01).Specific.VALUE = "104" Or oMat01.Columns("ItmBsort").Cells(oMat01Row01).Specific.VALUE = "107" Then
                        For i = 1 To oMat02.VisualRowCount
                            oMat03.AddRow
                            oMat03.Columns("LineNum").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat03.VisualRowCount
                            oMat03.Columns("Sequence").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat03.VisualRowCount
                            oMat03.Columns("ItemCode").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.VALUE
                            oMat03.Columns("ItemName").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat01.Columns("ItemName").Cells(oMat01Row01).Specific.VALUE
                            oMat03.Columns("CpBCode").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat02.Columns("CpBCode").Cells(i).Specific.VALUE
                            oMat03.Columns("CpBName").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat02.Columns("CpBName").Cells(i).Specific.VALUE
                            oMat03.Columns("CpCode").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat02.Columns("CpCode").Cells(i).Specific.VALUE
                            oMat03.Columns("CpName").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat02.Columns("CpName").Cells(i).Specific.VALUE
                            oMat03.Columns("CpUnWt").Cells(oMat03.VisualRowCount).Specific.VALUE = 0
                            If i = oMat02.VisualRowCount Then
                                Call oMat03.Columns("ResultYN").Cells(oMat03.VisualRowCount).Specific.Select("Y", psk_ByValue)
                            Else
                                Call oMat03.Columns("ResultYN").Cells(oMat03.VisualRowCount).Specific.Select("N", psk_ByValue)
                            End If
                            Call oMat03.Columns("ReportYN").Cells(oMat03.VisualRowCount).Specific.Select("Y", psk_ByValue)
                            Call oMat03.Columns("DayProYN").Cells(oMat03.VisualRowCount).Specific.Select("N", psk_ByValue)
                        Next
                    Else
                        oMat03.AddRow
                        oMat03.Columns("LineNum").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat03.VisualRowCount
                        oMat03.Columns("Sequence").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat03.VisualRowCount
                        oMat03.Columns("ItemCode").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.VALUE
                        oMat03.Columns("ItemName").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat01.Columns("ItemName").Cells(oMat01Row01).Specific.VALUE
                        oMat03.Columns("CpBCode").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat02.Columns("CpBCode").Cells(oMat02Row02).Specific.VALUE
                        oMat03.Columns("CpBName").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat02.Columns("CpBName").Cells(oMat02Row02).Specific.VALUE
                        oMat03.Columns("CpCode").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat02.Columns("CpCode").Cells(oMat02Row02).Specific.VALUE
                        oMat03.Columns("CpName").Cells(oMat03.VisualRowCount).Specific.VALUE = oMat02.Columns("CpName").Cells(oMat02Row02).Specific.VALUE
                        oMat03.Columns("CpUnWt").Cells(oMat03.VisualRowCount).Specific.VALUE = 0
                        Call oMat03.Columns("ResultYN").Cells(oMat03.VisualRowCount).Specific.Select("N", psk_ByValue)
                        Call oMat03.Columns("ReportYN").Cells(oMat03.VisualRowCount).Specific.Select("Y", psk_ByValue)
                        Call oMat03.Columns("DayProYN").Cells(oMat03.VisualRowCount).Specific.Select("N", psk_ByValue)
                    End If
                    
                End If
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
            If (pval.ItemUID = "Mat03") Then
                If (pval.ColUID = "StdTime") Then
                    oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = Left(Val(oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE), 4)
'                    '//기타작업
'                    Call oDS_PS_PP004L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP004L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_PP004_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_PP004L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
                End If
            Else
'                If (pval.ItemUID = "DocEntry") Then
'                    Call oDS_PS_PP004H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_PP004H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_PP004H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_PP004H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                End If
            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
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
        Call PS_PP004_FormItemEnabled
        '//Call PS_PP004_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP004_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        If (pval.ItemUID = "ItemCode") Then
            Dim oDataTable01 As SAPbouiCOM.DataTable
            Set oDataTable01 = pval.SelectedObjects
            oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns("ItemCode").Cells(0).VALUE
            oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns("ItemName").Cells(0).VALUE
            Set oDataTable01 = Nothing
        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP004H", "U_CardCode,U_CardName")
'        End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Or pval.ItemUID = "Mat03" Then
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
            If oMat03.Columns("CpCode").Cells(oMat03Row03).Specific.VALUE = "CP30112" Then '//바렐공정
                MDC_Com.MDC_GF_Message "바렐공정은 행삭제 할수 없습니다.", "W"
                BubbleEvent = False
                Exit Sub
            End If
            If oMat03.Columns("CpCode").Cells(oMat03Row03).Specific.VALUE = "CP30114" Then '//포장공정
                MDC_Com.MDC_GF_Message "포장공정은 행삭제 할수 없습니다.", "W"
                BubbleEvent = False
                Exit Sub
            End If
        ElseIf pval.BeforeAction = False Then
            If oMat01.Columns("ItmBsort").Cells(oMat01Row01).Specific.VALUE = "104" Or oMat01.Columns("ItmBsort").Cells(oMat01Row01).Specific.VALUE = "107" Then
                 '행별 삭제를 위해 주석 처리(2017.07.14 송명규)
'                oMat03.Clear
'                oMat03.FlushToDataSource
'                oMat03.LoadFromDataSource
                 '행별 삭제를 위해 주석 처리(2017.07.14 송명규)
            Else
                For i = 1 To oMat03.VisualRowCount
                    oMat03.Columns("LineNum").Cells(i).Specific.VALUE = i
                    oMat03.Columns("Sequence").Cells(i).Specific.VALUE = i '//순번
                Next i
            End If
        End If
    Else
        BubbleEvent = False
        Exit Sub
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP004_CreateItems() As Boolean
On Error GoTo PS_PP004_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_USERDS01 = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_USERDS02 = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    Set oDS_PS_USERDS03 = oForm01.DataSources.DBDataSources("@PS_USERDS03")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    Set oMat03 = oForm01.Items("Mat03").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat02.SelectionMode = ms_NotSupported
    oMat03.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    oMat02.AutoResizeColumns
    oMat03.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("ItmBsort", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("ItmMsort", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 20)
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("Opt03", dt_SHORT_TEXT, 100)
    
    Call oForm01.Items("ItmBsort").Specific.DataBind.SetBound(True, "", "ItmBsort")
    Call oForm01.Items("ItmMsort").Specific.DataBind.SetBound(True, "", "ItmMsort")
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    Call oForm01.Items("Opt01").Specific.DataBind.SetBound(True, "", "Opt01")
    Call oForm01.Items("Opt02").Specific.DataBind.SetBound(True, "", "Opt02")
    Call oForm01.Items("Opt03").Specific.DataBind.SetBound(True, "", "Opt03")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt03")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP004_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP004_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP004_ComboBox_Setting()
On Error GoTo PS_PP004_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP004", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP004", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP004", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP004", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP004", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PP004", "ItemCode")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP004", "Mat03", "ResultYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP004", "Mat03", "ResultYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns("ResultYN"), "PS_PP004", "Mat03", "ResultYN")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP004", "Mat03", "ReportYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP004", "Mat03", "ReportYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns("ReportYN"), "PS_PP004", "Mat03", "ReportYN")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP004", "Mat03", "DayProYN", "Y", "예")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP004", "Mat03", "DayProYN", "N", "아니오")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat03.Columns("DayProYN"), "PS_PP004", "Mat03", "DayProYN")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", False, False)
    
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmsGrpCod"), "SELECT ItmsGrpCod, ItmsGrpNam FROM OITB order by ItmsGrpCod")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmMsort"), "SELECT Code, Name FROM [@PSH_ITMMSORT] order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Mark"), "SELECT Code, Name FROM [@PSH_MARK] order by Code")
    
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat02.Columns("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code")
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP004_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP004_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP004_CF_ChooseFromList()
On Error GoTo PS_PP004_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

    Set oEdit = oForm01.Items("ItemCode").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = lf_Items
    oCFLCreationParams.uniqueID = "CFLITEMCODE"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oEdit.ChooseFromListUID = "CFLITEMCODE"
    oEdit.ChooseFromListAlias = "ItemCode"

    Exit Sub
PS_PP004_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP004_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP004_FormItemEnabled()
On Error GoTo PS_PP004_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_PP004_FormClear '//UDO방식
'        Call oForm01.EnableMenu("1281", True) '//찾기
'        Call oForm01.EnableMenu("1282", False) '//추가
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
'        Call oForm01.EnableMenu("1281", False) '//찾기
'        Call oForm01.EnableMenu("1282", True) '//추가
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP004_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP004_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP004_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP004_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_PP004L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_PP004L.Offset = oRow
'    oDS_PS_PP004L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP004_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP004_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP004_FormClear()
On Error GoTo PS_PP004_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP004'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP004_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP004_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP004_EnableMenus()
On Error GoTo PS_PP004_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP004_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP004_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP004_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP004_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP004_FormItemEnabled
        '//Call PS_PP004_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP004_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP004_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP004_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_PP004_DataValidCheck() As Boolean
On Error GoTo PS_PP004_DataValidCheck_Error
    PS_PP004_DataValidCheck = False
    Dim i As Long
    If oForm01.Items("ItmBsort").Specific.Selected Is Nothing Then
        Sbo_Application.SetStatusBarMessage "대분류는 필수입니다.", bmt_Short, True
        oForm01.Items("ItmBsort").Click ct_Regular
        PS_PP004_DataValidCheck = False
        Exit Function
    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_PP004_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_PP004_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_PS_PP004L.RemoveRecord(oDS_PS_PP004L.Size - 1)
'    Call oMat01.LoadFromDataSource
'    If (oForm01.Mode = fm_ADD_MODE) Then
'        Call PS_PP004_FormClear
'    End If
    PS_PP004_DataValidCheck = True
    Exit Function
PS_PP004_DataValidCheck_Error:
    PS_PP004_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP004_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP004_MTX01()
On Error GoTo PS_PP004_MTX01_Error
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
    If oForm01.Items("ItmBsort").Specific.Selected Is Nothing Then
    Else
        Param01 = oForm01.Items("ItmBsort").Specific.Selected.VALUE '//대분류
    End If
    If oForm01.Items("ItmMsort").Specific.Selected Is Nothing Then '//중분류
    Else
        Param02 = oForm01.Items("ItmMsort").Specific.Selected.VALUE
    End If
    Param03 = oForm01.Items("ItemCode").Specific.VALUE '//품목코드
    'Param04 = oForm01.Items("Param01").Specific.Value
    
    Query01 = "EXEC PS_PP004_01 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    oMat03.Clear
    oMat03.FlushToDataSource
    oMat03.LoadFromDataSource
    
    oMat01Row01 = 0
    oMat02Row02 = 0
    oMat03Row03 = 0
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP004_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    '//멀티,엔드베어링의경우
    If RecordSet01.Fields("ItmBsort").VALUE = "107" Or RecordSet01.Fields("ItmBsort").VALUE = "104" Then
        oForm01.Items("Button03").Enabled = False
        oForm01.Items("Button04").Enabled = False
        oMat03.Columns("ResultYN").Editable = False
        oMat03.Columns("ReportYN").Editable = False
        oMat03.Columns("DayProYN").Editable = False
    Else
        oForm01.Items("Button03").Enabled = True
        oForm01.Items("Button04").Enabled = True
        oMat03.Columns("ResultYN").Editable = True
        oMat03.Columns("ReportYN").Editable = True
        oMat03.Columns("DayProYN").Editable = True
    End If
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_USERDS01.InsertRecord (i)
        End If
        oDS_PS_USERDS01.offset = i
        Call oDS_PS_USERDS01.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_USERDS01.setValue("U_ColReg01", i, RecordSet01.Fields("ItemCode").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg02", i, RecordSet01.Fields("ItemName").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg03", i, RecordSet01.Fields("ItmsGrpCod").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg04", i, RecordSet01.Fields("ItmBsort").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg05", i, RecordSet01.Fields("ItmMsort").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg06", i, RecordSet01.Fields("Unit1").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg07", i, RecordSet01.Fields("Size").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg08", i, RecordSet01.Fields("ItemType").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColQty01", i, RecordSet01.Fields("UnWeight").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg09", i, RecordSet01.Fields("Quality").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg10", i, RecordSet01.Fields("Mark").VALUE)
        Call oDS_PS_USERDS01.setValue("U_ColReg11", i, RecordSet01.Fields("CallSize").VALUE)
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    Query01 = "EXEC PS_PP004_02 '" & Param01 & "'"
    Call RecordSet01.DoQuery(Query01)
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP004_MTX01_Exit
    End If
    
    Set ProgressBar01 = Nothing
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_USERDS02.InsertRecord (i)
        End If
        oDS_PS_USERDS02.offset = i
        Call oDS_PS_USERDS02.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_USERDS02.setValue("U_ColReg01", i, RecordSet01.Fields("CpBCode").VALUE)
        Call oDS_PS_USERDS02.setValue("U_ColReg02", i, RecordSet01.Fields("CpBName").VALUE)
        Call oDS_PS_USERDS02.setValue("U_ColReg03", i, RecordSet01.Fields("CpCode").VALUE)
        Call oDS_PS_USERDS02.setValue("U_ColReg04", i, RecordSet01.Fields("CpName").VALUE)
        Call oDS_PS_USERDS02.setValue("U_ColReg05", i, RecordSet01.Fields("ItmBsort").VALUE)
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
PS_PP004_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_PP004_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP004_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP004_MTX02()
On Error GoTo PS_PP004_MTX02_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Param01 = oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.VALUE
'    Param02 = oForm01.Items("Param01").Specific.Value
'    Param03 = oForm01.Items("Param01").Specific.Value
'    Param04 = oForm01.Items("Param01").Specific.Value
    
    Query01 = "EXEC PS_PP004_03 '" & Param01 & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat03.Clear
    oMat03.FlushToDataSource
    oMat03.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP004_MTX02_Exit
    End If
    
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
        
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_USERDS03.InsertRecord (i)
        End If
        oDS_PS_USERDS03.offset = i
        Call oDS_PS_USERDS03.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_USERDS03.setValue("U_ColNum01", i, RecordSet01.Fields("Sequence").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg01", i, RecordSet01.Fields("ItemCode").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg02", i, RecordSet01.Fields("ItemName").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg03", i, RecordSet01.Fields("CpBCode").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg04", i, RecordSet01.Fields("CpBName").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg05", i, RecordSet01.Fields("CpCode").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg06", i, RecordSet01.Fields("CpName").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColQty01", i, RecordSet01.Fields("CpUnWt").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg07", i, RecordSet01.Fields("ResultYN").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg08", i, RecordSet01.Fields("ReportYN").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg10", i, RecordSet01.Fields("DayProYN").VALUE)
        Call oDS_PS_USERDS03.setValue("U_ColReg09", i, RecordSet01.Fields("StdTime").VALUE)
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat03.LoadFromDataSource
    oMat03.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP004_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_PP004_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP004_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP004_DI_API() As Boolean
On Error GoTo PS_PP004_DI_API_Error
    PS_PP004_DI_API = True
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
        ItemInformation(ItemInformationCount).LotNo = oMat01.Columns("LotNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = oForm01.Items("BPLId").Specific.Selected.VALUE
    oDIObject.CardCode = oForm01.Items("CardCode").Specific.VALUE
    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.VALUE, "&&&&-&&-&&")
    For i = 0 To ItemInformationCount - 1
        If ItemInformation(i).Check = True Then
            GoTo Continue_First
        End If
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = oForm01.Items("WhsCode").Specific.VALUE
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
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).LotNo
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
'            Call oDS_PS_PP004L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_PP004L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        Next
    Else
        GoTo PS_PP004_DI_API_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_PP004_DI_API_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_PP004_DI_API = False
    Set oDIObject = Nothing
    Exit Function
PS_PP004_DI_API_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP004_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP004_DI_API = False
End Function


Private Function PS_PP004_Process_Register() As Boolean
On Error GoTo PS_PP004_Process_Register_Error
    PS_PP004_Process_Register = True
    Dim i As Long
    Dim Code As Long
    Dim Count As Long
    
    Dim CP30112 As Boolean
    Dim CP30114 As Boolean
    For i = 1 To oMat03.VisualRowCount
        If (oMat03.Columns("CpBCode").Cells(i).Specific.VALUE = "CP301") Then '//공정이 휘팅이면
            If oMat03.Columns("CpCode").Cells(i).Specific.VALUE = "CP30112" Then '//바렐공정
                CP30112 = True
            End If
            If oMat03.Columns("CpCode").Cells(i).Specific.VALUE = "CP30114" Then '//포장공정
                CP30114 = True
            End If
        End If
    Next
    For i = 1 To oMat03.VisualRowCount
        If (oMat03.Columns("CpBCode").Cells(i).Specific.VALUE = "CP301") Then '//공정이 휘팅이면
            If CP30112 <> True Or CP30114 <> True Then
                MDC_Com.MDC_GF_Message "휘팅공정은 바렐,포장공정은 필수 입니다.", "W"
                PS_PP004_Process_Register = False
                Exit Function
            End If
        End If
    Next
    
    Count = 0
    
    For i = 1 To oMat03.VisualRowCount
        If (oMat03.Columns("DayProYN").Cells(i).Specific.VALUE = "Y") Then '//일생산실적
            Count = Count + 1
        End If
    Next
    
    If Count > 1 Then
        Sbo_Application.MessageBox ("일생산실적 포인트는 하나만 가능합니다. 저장되지 않았습니다.")
        PS_PP004_Process_Register = False
        Exit Function
    End If
    
    MDC_PS_Common.DoQuery ("DELETE [@PS_PP004H] WHERE U_ItemCode = '" & oMat01.Columns("ItemCode").Cells(oMat01Row01).Specific.VALUE & "'")
    For i = 1 To oMat03.VisualRowCount
        Code = MDC_PS_Common.GetValue("SELECT ISNULL(MAX(CONVERT(INT,Code)),0)+1 FROM [@PS_PP004H]")
        MDC_PS_Common.DoQuery ("INSERT INTO [@PS_PP004H] VALUES('" & Code & "','" & Code & "','" & oMat03.Columns("Sequence").Cells(i).Specific.VALUE & "','" & oMat03.Columns("ItemCode").Cells(i).Specific.VALUE & "','" & oMat03.Columns("ItemName").Cells(i).Specific.VALUE & "','" & oMat03.Columns("CpBCode").Cells(i).Specific.VALUE & "','" & oMat03.Columns("CpBName").Cells(i).Specific.VALUE & "','" & oMat03.Columns("CpCode").Cells(i).Specific.VALUE & "','" & oMat03.Columns("CpName").Cells(i).Specific.VALUE & "','" & oMat03.Columns("CpUnWt").Cells(i).Specific.VALUE & "','" & oMat03.Columns("StdTime").Cells(i).Specific.VALUE & "','" & oMat03.Columns("ResultYN").Cells(i).Specific.Selected.VALUE & "','" & oMat03.Columns("ReportYN").Cells(i).Specific.Selected.VALUE & "','" & oMat03.Columns("DayProYN").Cells(i).Specific.Selected.VALUE & "')")
    Next
    Call MDC_Com.MDC_GF_Message("데이터가 " & oMat03.VisualRowCount & " 개 저장되었습니다.", "W")
    PS_PP004_Process_Register = True
    Exit Function
PS_PP004_Process_Register_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP004_Process_Register_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP004_FormResize()
On Error GoTo PS_PP004_FormResize_Error
    oForm01.Items("Mat01").Top = 90
    oForm01.Items("Mat01").Height = (oForm01.ClientHeight) - 120
    oForm01.Items("Mat01").Left = 7
    oForm01.Items("Mat01").Width = oForm01.ClientWidth / 2 - 10
    
    oForm01.Items("Mat02").Top = 90
    oForm01.Items("Mat02").Height = ((oForm01.ClientHeight - 120) / 2) - 15
    oForm01.Items("Mat02").Left = oForm01.ClientWidth / 2
    oForm01.Items("Mat02").Width = oForm01.ClientWidth / 2 - 10
    
    oForm01.Items("Mat03").Top = 90 + ((oForm01.ClientHeight - 120) / 2) + 15
    oForm01.Items("Mat03").Height = ((oForm01.ClientHeight - 120) / 2) - 15
    oForm01.Items("Mat03").Left = oForm01.ClientWidth / 2
    oForm01.Items("Mat03").Width = oForm01.ClientWidth / 2 - 10
    
    oForm01.Items("Button02").Top = oForm01.Items("Mat03").Top - 25
    oForm01.Items("Button03").Top = oForm01.Items("Mat03").Top - 25
    oForm01.Items("Button04").Top = oForm01.Items("Mat03").Top - 25
    
    oForm01.Items("Button02").Left = (oForm01.Items("Mat03").Left + oForm01.Items("Mat03").Width) - 100
    oForm01.Items("Button03").Left = (oForm01.Items("Mat03").Left + oForm01.Items("Mat03").Width) - 202
    oForm01.Items("Button04").Left = (oForm01.Items("Mat03").Left + oForm01.Items("Mat03").Width) - 304
    
    oForm01.Items("Opt03").Top = oForm01.Items("Mat03").Top - 20
    oForm01.Items("Opt03").Left = oForm01.Items("Mat03").Left
    
    oForm01.Items("Opt02").Left = oForm01.Items("Mat02").Left
    Exit Sub
PS_PP004_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP004_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
