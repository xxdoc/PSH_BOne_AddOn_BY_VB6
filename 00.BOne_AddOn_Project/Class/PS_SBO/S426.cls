VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S426"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : S426.cls
'//  Module         :
'//  Description    : 자금관리-지급
'//  FormType       : 426
'//  Create Date    : 2011.01.26
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
    oForm01.Items("5").Click ct_Regular
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "5" Or pval.ItemUID = "32" Then
                        oForm01.Freeze True
                        Call oForm01.Items("BPLId").Specific.Select(User_BPLId(), psk_ByValue)
                        
                        sQry = "Select VATRegNum From OCRD Where CardCode = '" + oForm01.Items("5").Specific.VALUE + "'"
                        oRecordSet01.DoQuery sQry
                        oForm01.Items("26").Specific.VALUE = oRecordSet01(0).VALUE
                        
                        oForm01.Items("5").Click ct_Regular
                        oForm01.Freeze False
                    End If
                End If
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
                oForm01.Freeze True
                Call oForm01.Items("BPLId").Specific.Select(User_BPLId(), psk_ByValue)
                oForm01.Items("5").Click ct_Regular
                oForm01.Freeze False
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
    Dim oCombo01        As SAPbouiCOM.ComboBox
    Dim oCombo02        As SAPbouiCOM.ComboBox
    Dim oStatic01       As SAPbouiCOM.StaticText
    Dim oStatic02       As SAPbouiCOM.StaticText
    Dim oChkBox         As SAPbouiCOM.CheckBox
    Dim optBtn          As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    oForm01.Freeze True

    '//사업장-ComboBox
    Set newItem = oForm01.Items.Add("BPLId", it_COMBO_BOX)
    Set stdItem = oForm01.Items("95")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top + 16
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.DisplayDesc = True
    Set oCombo01 = newItem.Specific
    Call oCombo01.DataBind.SetBound(True, "OVPM", "U_BPLId")

    sQry = "select BPLId, BPLName from [OBPL] order by BPLId"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo01.ValidValues.Add Trim(oRecordSet.Fields(0).VALUE), Trim(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    Call oCombo01.Select(User_BPLId(), psk_ByValue)
    
    Set newItem = oForm01.Items.Add("Static01", it_STATIC)
    Set stdItem = oForm01.Items("96")
    newItem.Left = stdItem.Left
    newItem.Top = stdItem.Top + 16
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.LinkTo = "BPLId"
    Set oStatic01 = newItem.Specific
    oStatic01.Caption = "사업장"
        
    oForm01.Freeze False
    
    Set oRecordSet = Nothing
    Set oStatic01 = Nothing
    Set oCombo01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    Set oStatic01 = Nothing
    Set oCombo01 = Nothing
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
