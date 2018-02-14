VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SM60100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//****************************************************************************
'//  File           : SM60100.cls
'//  Module         : 인사관리 > 사용자 정의 필드
'//  Desc           : OHEM
'//****************************************************************************

Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form
Public oMat1 As SAPbouiCOM.Matrix
Public oGrid1 As SAPbouiCOM.Grid

Private sFromPane       As Long
Private sToPane         As Long
Private sInit           As Boolean
Private sTable          As String
Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(ByVal oFormUid As String)

    Dim sQry                As String
    Dim oItem               As SAPbouiCOM.Item
    Dim oItem01             As SAPbouiCOM.Item
    Dim oCombo              As SAPbouiCOM.ComboBox
    Dim oButton             As SAPbouiCOM.Button
    
    oFormUniqueID = oFormUid
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    Call SM60100_CreateItems
    
    Call SM60100_FormItemEnabled
End Sub

Private Sub SM60100_CreateItems()

    Dim sQry                As String
    Dim oItem               As SAPbouiCOM.Item
    Dim oItem01             As SAPbouiCOM.Item
    Dim oCombo              As SAPbouiCOM.ComboBox
    Dim oFolder             As SAPbouiCOM.Folder
    Dim oEdit               As SAPbouiCOM.EditText
    
    Dim oUserField          As SAPbobsCOM.UserFieldsMD
    On Error GoTo Error_Handle
    
    oForm.Freeze True

    Set oItem = Nothing
    
    oForm.Freeze False
    
    Exit Sub

Error_Handle:
    oForm.Freeze False
    Sbo_Application.SetStatusBarMessage "CreateItems Error:" & Err.Description, bmt_Short, True
End Sub


Private Sub SM60100_FormItemEnabled()

    On Error GoTo SM60100_FormItemEnabled_Error
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
    
       
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        
       
    ElseIf (oForm.Mode = fm_OK_MODE) Then
      
    End If
    Call oForm.Freeze(False)
    Exit Sub
SM60100_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "SM60100_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

    On Error GoTo Raise_FormItemEvent_Error
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            If pval.Before_Action = False Then
               
            End If
            
           
        Case et_KEY_DOWN: '//2
            
        Case et_GOT_FOCUS: '//3
            '// 종료, 취소 막음
            
        Case et_LOST_FOCUS: '//4

        Case et_COMBO_SELECT: '//5
            If pval.Before_Action = False Then
                
            End If
        Case et_CLICK: '//6
            '// 종료, 취소 막음
            
        Case et_DOUBLE_CLICK: '//7

        Case et_MATRIX_LINK_PRESSED '//8

        Case et_MATRIX_COLLAPSE_PRESSED '//9

        Case et_VALIDATE: '//10
            
        Case et_MATRIX_LOAD: '//11
        
        Case et_DATASOURCE_LOAD '//12

        Case et_FORM_LOAD: '//16

        Case et_FORM_UNLOAD: '//17
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
            End If
        Case et_FORM_ACTIVATE: '//18

        Case et_FORM_DEACTIVATE: '//19

        Case et_FORM_CLOSE '//20

        Case et_FORM_RESIZE '//21
            
        Case et_FORM_KEY_DOWN '//22
            
        Case et_FORM_MENU_HILIGHT '//23
            
        Case et_CHOOSE_FROM_LIST '//27
            
        Case et_PICKER_CLICKED '//37
            
        Case et_GRID_SORT '//38
            
        Case et_Drag '//39
            
    End Select
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_FormMenuEvent_Error
    
    If (pval.BeforeAction = True) Then
        Select Case pval.MenuUID
            Case "1284":
            Case "1286":
            Case "1293":
                
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1284":
            Case "1286":
            Case "1293":
                
            Case "1281":
                Call SM60100_FormItemEnabled
            Case "1282":
                
                oForm.Freeze False
            Case "1288", "1289", "1290", "1291":
                Call SM60100_FormItemEnabled
        End Select
    End If
    
    Exit Sub
    
Raise_FormMenuEvent_Error:
    oForm.Freeze False
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_FormDataEvent_Error
    
    If (BusinessObjectInfo.BeforeAction = True) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
                If oForm.Items.Item("SOTYPE").Specific.Value = "2" And sInit = False Then
                    Call SM60100_CreateItems
                End If
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

    Set oForm = Sbo_Application.Forms.Item(pval.FormUID)
    
    If pval.BeforeAction Then
        Select Case pval.EventType
        Case et_RIGHT_CLICK
'            If Left$(pval.ItemUID, 3) = "CMB" Or Left$(pval.ItemUID, 3) = "EDT" Then
'                oForm.ActiveItem = "16"
'                BubbleEvent = False
'                Exit Sub
'            ElseIf pval.ItemUID = "38" Then
'                If GF_Nz(GF_DLookup("SUM(U_INGQTY)", "POR1", " DOCENTRY = " & Trim$(oForm.Items("8").Specific.String))) > 0 Then
'                    oForm.ActiveItem = "16"
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'            End If
        End Select
    Else
        Select Case pval.EventType
        Case et_RIGHT_CLICK
'            If Left$(pval.ItemUID, 3) = "CMB" Or Left$(pval.ItemUID, 3) = "EDT" Then
'                BubbleEvent = False
'                Exit Sub
'            End If
        End Select
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub SM60100_Print_Report01()

    Dim DOCNUM As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    
    On Error GoTo SM60100_Print_Report01_Error

    If MDC_GetData.Get_ReData("COUNT(CardCode)", "DocEntry", "[ORDR]", "'" & oForm.Items.Item("8").Specific.Value & "'") > 0 Then
    
        If ConnectODBC = False Then
            GoTo SM60100_Print_Report01_Error
        End If
    
        '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WinTitle = "[SM60100] : ORDER SHEET"
        ReportName = "SM60100_1.RPT"
        sQry = " Exec SM60100_1 " & "'" & Trim(oForm.Items("8").Specific.Value) & "'"

        '/ Formula 수식필드***************************************************/
        ReDim gRpt_Formula(8)
        ReDim gRpt_Formula_Value(8)
        
        '/ SubReport /
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        
        If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
            Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
        End If
    Else
        Sbo_Application.SetStatusBarMessage "저장된 판매오더 문서가 없습니다.", bmt_Short, True
    End If
    
    
    
    Exit Sub
SM60100_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "SM60100_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function SM60100_DataValidCheck() As Boolean
    
    On Error GoTo SM60100_DataValidCheck_Error
    
    SM60100_DataValidCheck = False
'
'    '// 만약 내수일 경우에는 SKIP
'    If Not oForm.Items("SOTYPE").Specific.Selected Is Nothing Then
'        If Trim$(oForm.Items("SOTYPE").Specific.Selected.Value) = "2" Then
'            If oForm.Items("SETTLE").Specific.Selected Is Nothing Then
'                Sbo_Application.SetStatusBarMessage "결제조건은 필수 입력 항목입니다.", bmt_Short, True
'                SM60100_DataValidCheck = False
'                Exit Function
'            End If
'        End If
'    End If

    SM60100_DataValidCheck = True
    
    Exit Function

SM60100_DataValidCheck_Error:
    SM60100_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "SM60100_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function







