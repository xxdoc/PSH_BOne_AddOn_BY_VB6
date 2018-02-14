VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S804"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : S804.cls
'//  Module         : FI
'//  Description    : 계정과목표
'//  FormType       : S804
'//  Create Date    : 2010.11.19
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Private oFormUniqueID01 As String
Private oForm01              As SAPbouiCOM.Form
'Private oMat01              As SAPbouiCOM.Matrix
Private oDS_OACT As SAPbouiCOM.DBDataSource    '등록헤더
'Private oDS_RDR1 As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID          As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID      As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Public Sub LoadForm(ByVal oFormUniqueID As String)
    oFormUniqueID01 = oFormUniqueID
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    Call CreateItems
'    Call FormItemEnabled
End Sub

Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
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
                oLast_Item_UID = pval.ItemUID
                If oLast_Item_UID = "38" Then
                    If pval.Row > 0 Then
                        oLast_Item_UID = pval.ItemUID
                        oLast_Col_UID = pval.ColUID
                        oLast_Col_Row = pval.Row
                    End If
                Else
                    oLast_Item_UID = pval.ItemUID
                    oLast_Col_UID = ""
                    oLast_Col_Row = 0
                End If
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
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
                oLast_Item_UID = pval.ItemUID
                If oLast_Item_UID = "38" Then
                    If pval.Row > 0 Then
                        oLast_Item_UID = pval.ItemUID
                        oLast_Col_UID = pval.ColUID
                        oLast_Col_Row = pval.Row
                    End If
                Else
                    oLast_Item_UID = pval.ItemUID
                    oLast_Col_UID = ""
                    oLast_Col_Row = 0
                End If
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
'                Set oMat01 = Nothing
        End Select
    End If
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
            Case "1293": '행닫기
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행닫기
            Case "1281": '찾기
            Case "1282": '추가
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//작업
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    Dim oNewITEM    As Object
    Dim oItem       As Object
    Dim oBTN As SAPbouiCOM.Button
    
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oItem = oForm01.Items("2000")
    '--------------------------------------------------------------------------------------------
    Set oNewITEM = oForm01.Items.Add("Text", it_STATIC)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23
    oNewITEM.Left = oItem.Left
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 60
    oNewITEM.Specific.Caption = "분개전표"
    '--------------------------------------------------------------------------------------------
    Set oNewITEM = oForm01.Items.Add("RptCre01", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23
    oNewITEM.Left = oItem.Left + 60 + 40 + 103
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105
    
    Set oNewITEM = oForm01.Items.Add("RptDeb01", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23
    oNewITEM.Left = oItem.Left + 60 + 33
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105

    Set oNewITEM = oForm01.Items.Add("Text01", it_STATIC)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23
    oNewITEM.Left = oItem.Left + 54
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 40
    oNewITEM.LinkTo = "RptCre01"
    oNewITEM.Specific.Caption = "항목1"

    oForm01.Items("RptCre01").Specific.DataBind.SetBound True, "OACT", "U_RptCre01"
    oForm01.Items("RptDeb01").Specific.DataBind.SetBound True, "OACT", "U_RptDeb01"
    '--------------------------------------------------------------------------------------------
    Set oNewITEM = oForm01.Items.Add("RptCre02", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16
    oNewITEM.Left = oItem.Left + 60 + 40 + 103
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105
    
    Set oNewITEM = oForm01.Items.Add("RptDeb02", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16
    oNewITEM.Left = oItem.Left + 60 + 33
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105

    Set oNewITEM = oForm01.Items.Add("Text02", it_STATIC)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16
    oNewITEM.Left = oItem.Left + 54
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 40
    oNewITEM.LinkTo = "RptCre02"
    oNewITEM.Specific.Caption = "항목2"

    oForm01.Items("RptCre02").Specific.DataBind.SetBound True, "OACT", "U_RptCre02"
    oForm01.Items("RptDeb02").Specific.DataBind.SetBound True, "OACT", "U_RptDeb02"
    '--------------------------------------------------------------------------------------------
    Set oNewITEM = oForm01.Items.Add("RptCre03", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16
    oNewITEM.Left = oItem.Left + 60 + 40 + 103
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105
    
    Set oNewITEM = oForm01.Items.Add("RptDeb03", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16
    oNewITEM.Left = oItem.Left + 60 + 33
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105

    Set oNewITEM = oForm01.Items.Add("Text03", it_STATIC)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16
    oNewITEM.Left = oItem.Left + 54
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 40
    oNewITEM.LinkTo = "RptCre03"
    oNewITEM.Specific.Caption = "항목3"

    oForm01.Items("RptCre03").Specific.DataBind.SetBound True, "OACT", "U_RptCre03"
    oForm01.Items("RptDeb03").Specific.DataBind.SetBound True, "OACT", "U_RptDeb03"
    '--------------------------------------------------------------------------------------------
    Set oNewITEM = oForm01.Items.Add("RptCre04", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16
    oNewITEM.Left = oItem.Left + 60 + 40 + 103
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105
    
    Set oNewITEM = oForm01.Items.Add("RptDeb04", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16
    oNewITEM.Left = oItem.Left + 60 + 33
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105

    Set oNewITEM = oForm01.Items.Add("Text04", it_STATIC)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16
    oNewITEM.Left = oItem.Left + 54
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 40
    oNewITEM.LinkTo = "RptCre04"
    oNewITEM.Specific.Caption = "항목4"

    oForm01.Items("RptCre04").Specific.DataBind.SetBound True, "OACT", "U_RptCre04"
    oForm01.Items("RptDeb04").Specific.DataBind.SetBound True, "OACT", "U_RptDeb04"
    '--------------------------------------------------------------------------------------------
    Set oNewITEM = oForm01.Items.Add("RptCre05", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16
    oNewITEM.Left = oItem.Left + 60 + 40 + 103
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105
    
    Set oNewITEM = oForm01.Items.Add("RptDeb05", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16
    oNewITEM.Left = oItem.Left + 60 + 33
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105

    Set oNewITEM = oForm01.Items.Add("Text05", it_STATIC)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16
    oNewITEM.Left = oItem.Left + 54
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 40
    oNewITEM.LinkTo = "RptCre05"
    oNewITEM.Specific.Caption = "항목5"

    oForm01.Items("RptCre05").Specific.DataBind.SetBound True, "OACT", "U_RptCre05"
    oForm01.Items("RptDeb05").Specific.DataBind.SetBound True, "OACT", "U_RptDeb05"
    '--------------------------------------------------------------------------------------------
    Set oNewITEM = oForm01.Items.Add("RptCre06", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16 + 16
    oNewITEM.Left = oItem.Left + 60 + 40 + 103
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105
    
    Set oNewITEM = oForm01.Items.Add("RptDeb06", it_COMBO_BOX)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16 + 16
    oNewITEM.Left = oItem.Left + 60 + 33
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 105

    Set oNewITEM = oForm01.Items.Add("Text06", it_STATIC)
'    oNewITEM.AffectsFormMode = False
    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16 + 16
    oNewITEM.Left = oItem.Left + 54
    oNewITEM.Height = oItem.Height
    oNewITEM.Width = 40
    oNewITEM.LinkTo = "RptCre06"
    oNewITEM.Specific.Caption = "항목6"

    oForm01.Items("RptCre06").Specific.DataBind.SetBound True, "OACT", "U_RptCre06"
    oForm01.Items("RptDeb06").Specific.DataBind.SetBound True, "OACT", "U_RptDeb06"
    '--------------------------------------------------------------------------------------------
    '//Combo
    sQry = "select U_Minor, U_CdName from [@PS_SY001L] Where Code = 'F001' Order by Convert(Int, U_LineNum)"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oForm01.Items("RptCre01").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptDeb01").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptCre02").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptDeb02").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptCre03").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptDeb03").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptCre04").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptDeb04").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptCre05").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptDeb05").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptCre06").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oForm01.Items("RptDeb06").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oRecordSet01.MoveNext
    Loop
    '--------------------------------------------------------------------------------------------
'    Set oNewITEM = oForm01.Items.Add("Text", it_STATIC)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23
'    oNewITEM.Left = oItem.Left
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 60
'    oNewITEM.Specific.Caption = "분개전표"
'
'    Set oNewITEM = oForm01.Items.Add("RptItm01", it_COMBO_BOX)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23
'    oNewITEM.Left = oItem.Left + 60 + 40 + 82
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 110
'
'    '//Combo
'    Set oCombo = oForm01.Items("RptItm01").Specific
'    sQry = "select U_Minor, U_CdName from [@PS_SY001L] Where Code = 'F001' Order by U_LineNum"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
'
'    Set oNewITEM = oForm01.Items.Add("Rpttxt01", it_EDIT)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23
'    oNewITEM.Left = oItem.Left + 60 + 40
'    oNewITEM.Height = oItem.Height
'    oNewITEM.LinkTo = "RptItm01"
'    oNewITEM.Width = 80
'
'    Set oNewITEM = oForm01.Items.Add("Text01", it_STATIC)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23
'    oNewITEM.Left = oItem.Left + 60
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 40
'    oNewITEM.LinkTo = "Rpttxt01"
'    oNewITEM.Specific.Caption = "항목 1"
'
'    oForm01.Items("RptItm01").Specific.DataBind.SetBound True, "OACT", "U_RptItm01"
'    oForm01.Items("Rpttxt01").Specific.DataBind.SetBound True, "OACT", "U_Rpttxt01"
'
'    Set oNewITEM = oForm01.Items.Add("RptItm02", it_COMBO_BOX)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40 + 82
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 110
'
'    '//Combo
'    Set oCombo = oForm01.Items("RptItm02").Specific
'    sQry = "select U_Minor, U_CdName from [@PS_SY001L] Where Code = 'F001' Order by U_LineNum"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
'
'    Set oNewITEM = oForm01.Items.Add("Rpttxt02", it_EDIT)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40
'    oNewITEM.Height = oItem.Height
'    oNewITEM.LinkTo = "RptItm02"
'    oNewITEM.Width = 80
'
'    Set oNewITEM = oForm01.Items.Add("Text02", it_STATIC)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16
'    oNewITEM.Left = oItem.Left + 60
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 40
'    oNewITEM.LinkTo = "Rpttxt02"
'    oNewITEM.Specific.Caption = "항목 2"
'
'    oForm01.Items("RptItm02").Specific.DataBind.SetBound True, "OACT", "U_RptItm02"
'    oForm01.Items("Rpttxt02").Specific.DataBind.SetBound True, "OACT", "U_Rpttxt02"
'
'    Set oNewITEM = oForm01.Items.Add("RptItm03", it_COMBO_BOX)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40 + 82
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 110
'
'    '//Combo
'    Set oCombo = oForm01.Items("RptItm03").Specific
'    sQry = "select U_Minor, U_CdName from [@PS_SY001L] Where Code = 'F001' Order by U_LineNum"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
'
'    Set oNewITEM = oForm01.Items.Add("Rpttxt03", it_EDIT)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40
'    oNewITEM.Height = oItem.Height
'    oNewITEM.LinkTo = "RptItm03"
'    oNewITEM.Width = 80
'
'    Set oNewITEM = oForm01.Items.Add("Text03", it_STATIC)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 40
'    oNewITEM.LinkTo = "Rpttxt03"
'    oNewITEM.Specific.Caption = "항목 3"
'
'    oForm01.Items("RptItm03").Specific.DataBind.SetBound True, "OACT", "U_RptItm03"
'    oForm01.Items("Rpttxt03").Specific.DataBind.SetBound True, "OACT", "U_Rpttxt03"
'
'    Set oNewITEM = oForm01.Items.Add("RptItm04", it_COMBO_BOX)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40 + 82
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 110
'
'    '//Combo
'    Set oCombo = oForm01.Items("RptItm04").Specific
'    sQry = "select U_Minor, U_CdName from [@PS_SY001L] Where Code = 'F001' Order by U_LineNum"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
'
'    Set oNewITEM = oForm01.Items.Add("Rpttxt04", it_EDIT)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40
'    oNewITEM.Height = oItem.Height
'    oNewITEM.LinkTo = "RptItm04"
'    oNewITEM.Width = 80
'
'    Set oNewITEM = oForm01.Items.Add("Text04", it_STATIC)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 40
'    oNewITEM.LinkTo = "Rpttxt04"
'    oNewITEM.Specific.Caption = "항목 4"
'
'    oForm01.Items("RptItm04").Specific.DataBind.SetBound True, "OACT", "U_RptItm04"
'    oForm01.Items("Rpttxt04").Specific.DataBind.SetBound True, "OACT", "U_Rpttxt04"
'
'    Set oNewITEM = oForm01.Items.Add("RptItm05", it_COMBO_BOX)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40 + 82
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 110
'
'    '//Combo
'    Set oCombo = oForm01.Items("RptItm05").Specific
'    sQry = "select U_Minor, U_CdName from [@PS_SY001L] Where Code = 'F001' Order by U_LineNum"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
'
'    Set oNewITEM = oForm01.Items.Add("Rpttxt05", it_EDIT)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40
'    oNewITEM.Height = oItem.Height
'    oNewITEM.LinkTo = "RptItm05"
'    oNewITEM.Width = 80
'
'    Set oNewITEM = oForm01.Items.Add("Text05", it_STATIC)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 40
'    oNewITEM.LinkTo = "Rpttxt05"
'    oNewITEM.Specific.Caption = "항목 5"
'
'    oForm01.Items("RptItm05").Specific.DataBind.SetBound True, "OACT", "U_RptItm05"
'    oForm01.Items("Rpttxt05").Specific.DataBind.SetBound True, "OACT", "U_Rpttxt05"
'
'    Set oNewITEM = oForm01.Items.Add("RptItm06", it_COMBO_BOX)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40 + 82
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 110
'
'    '//Combo
'    Set oCombo = oForm01.Items("RptItm06").Specific
'    sQry = "select U_Minor, U_CdName from [@PS_SY001L] Where Code = 'F001' Order by U_LineNum"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
'
'    Set oNewITEM = oForm01.Items.Add("Rpttxt06", it_EDIT)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60 + 40
'    oNewITEM.Height = oItem.Height
'    oNewITEM.LinkTo = "RptItm06"
'    oNewITEM.Width = 80
'
'    Set oNewITEM = oForm01.Items.Add("Text06", it_STATIC)
''    oNewITEM.AffectsFormMode = False
'    oNewITEM.Top = oItem.Top + 23 + 16 + 16 + 16 + 16 + 16
'    oNewITEM.Left = oItem.Left + 60
'    oNewITEM.Height = oItem.Height
'    oNewITEM.Width = 40
'    oNewITEM.LinkTo = "Rpttxt06"
'    oNewITEM.Specific.Caption = "항목 6"
'
'    oForm01.Items("RptItm06").Specific.DataBind.SetBound True, "OACT", "U_RptItm06"
'    oForm01.Items("Rpttxt06").Specific.DataBind.SetBound True, "OACT", "U_Rpttxt06"
    
    Set oItem = Nothing
    Set oNewITEM = Nothing
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
CreateItems_Error:
    Set oItem = Nothing
    Set oNewITEM = Nothing
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
