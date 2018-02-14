VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S146"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : S146.cls
'//  Module         :
'//  Description    : 입금-지급수단
'//  FormType       : 146
'//  Create Date    : 2010.11.24
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix

Private oMat01Row           As Long

Public Sub LoadForm(ByVal oFormUniqueID As String)
    oFormUniqueID01 = oFormUniqueID
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
'    Set oMat01 = oForm01.Items("76").Specific
    Call CreateItems
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
        
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_FORM_LOAD
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
'                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1285": '복원
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1285": '복원
            Case "1287": '복제
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        If eventInfo.ItemUID = "76" Then
            If eventInfo.Row > 0 Then
                oMat01Row = eventInfo.Row
            End If
        End If
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormDataEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_FormDataEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    Dim stdItem         As Object
    Dim newItem         As Object
    Dim oEdit01         As SAPbouiCOM.EditText
    Dim oEdit02         As SAPbouiCOM.EditText
    Dim oEdit03         As SAPbouiCOM.EditText
    Dim optBtn          As SAPbouiCOM.OptionBtn
    Dim oCombo01        As SAPbouiCOM.ComboBox
    Dim oCombo02        As SAPbouiCOM.ComboBox
    Dim oStatic01       As SAPbouiCOM.StaticText
    Dim oStatic02       As SAPbouiCOM.StaticText
    Dim oStatic03       As SAPbouiCOM.StaticText
    Dim oStatic04       As SAPbouiCOM.StaticText
    Dim oChkBox         As SAPbouiCOM.CheckBox
    
    Dim oRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)

    oForm01.Freeze True

'    '//전자어음-CheckBox
'    Set newItem = oForm01.Items.Add("ChkBox", it_CHECK_BOX)
'    Set stdItem = oForm01.Items("114")
'    newItem.Left = stdItem.Left
'    newItem.Top = stdItem.Top + 22
'    newItem.Height = stdItem.Height
'    newItem.Width = stdItem.Width - 50
'    newItem.FromPane = 5
'    newItem.ToPane = 5
''    newItem.LinkTo = "TradeType"
'    Set oChkBox = newItem.Specific
'    Call oChkBox.DataBind.SetBound(True, "OBOE", "U_EBill")
'    oChkBox.ValOn = "Y"
'    oChkBox.ValOff = "N"
'    oChkBox.Caption = "전자어음"

    '//전자어음-CheckBox
    Set newItem = oForm01.Items.Add("Combo01", it_COMBO_BOX)
    Set stdItem = oForm01.Items("118")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top - 17
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 5
    newItem.ToPane = 5
    newItem.DisplayDesc = True
    Set oCombo01 = newItem.Specific
    Call oCombo01.DataBind.SetBound(True, "OBOE", "U_EBill")
'    oCombo01.ValidValues.Add "1", "약속"
'    oCombo01.ValidValues.Add "2", "전자"
'    oCombo01.ValidValues.Add "3", "당좌"
'    oCombo01.ValidValues.Add "4", "가계"
'    Call oCombo01.Select(0, psk_Index)

    Set newItem = oForm01.Items.Add("Static04", it_STATIC)
    Set stdItem = oForm01.Items("117")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top - 17
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 5
    newItem.ToPane = 5
    newItem.LinkTo = "Combo01"
    Set oStatic04 = newItem.Specific
    oStatic04.Caption = "어음구분"
    
    
    '//발행일
    Set newItem = oForm01.Items.Add("Edit01", it_EDIT)
    Set stdItem = oForm01.Items("159")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top + 18
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 5
    newItem.ToPane = 5
'    newItem.LinkTo = "Static01"
    Set oEdit01 = newItem.Specific
    Call oEdit01.DataBind.SetBound(True, "OBOE", "U_IssueDat")

    Set newItem = oForm01.Items.Add("Static01", it_STATIC)
    Set stdItem = oForm01.Items("158")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top + 18
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 5
    newItem.ToPane = 5
    newItem.LinkTo = "Edit01"
    Set oStatic01 = newItem.Specific
    oStatic01.Caption = "발행일"
    
    '//발행인
    Set newItem = oForm01.Items.Add("Edit02", it_EDIT)
    Set stdItem = oForm01.Items("159")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top + 68
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 5
    newItem.ToPane = 5
'    newItem.LinkTo = "Static01"
    Set oEdit02 = newItem.Specific
    Call oEdit02.DataBind.SetBound(True, "OBOE", "U_IssuePsn")
    
    Set newItem = oForm01.Items.Add("Static02", it_STATIC)
    Set stdItem = oForm01.Items("158")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top + 68
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 5
    newItem.ToPane = 5
    newItem.LinkTo = "Edit02"
    Set oStatic02 = newItem.Specific
    oStatic02.Caption = "발행인"
    
    '//지급지
    Set newItem = oForm01.Items.Add("Edit03", it_EDIT)
    Set stdItem = oForm01.Items("159")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top + 86
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 5
    newItem.ToPane = 5
'    newItem.LinkTo = "Static01"
    Set oEdit03 = newItem.Specific
    Call oEdit03.DataBind.SetBound(True, "OBOE", "U_PayBank")

    Set newItem = oForm01.Items.Add("Static03", it_STATIC)
    Set stdItem = oForm01.Items("158")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top + 86
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 5
    newItem.ToPane = 5
    newItem.LinkTo = "Edit03"
    Set oStatic03 = newItem.Specific
    oStatic03.Caption = "지급지"
            
    oForm01.Freeze False
    
    
    Set oStatic04 = Nothing
    Set oCombo01 = Nothing
    Set oStatic01 = Nothing
    Set oEdit01 = Nothing
    Set oStatic02 = Nothing
    Set oEdit02 = Nothing
    Set oStatic03 = Nothing
    Set oEdit03 = Nothing
    Set oRecordset = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    Set oStatic04 = Nothing
    Set oCombo01 = Nothing
    Set oStatic01 = Nothing
    Set oEdit01 = Nothing
    Set oStatic02 = Nothing
    Set oEdit02 = Nothing
    Set oStatic03 = Nothing
    Set oEdit03 = Nothing
    Set oRecordset = Nothing
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

