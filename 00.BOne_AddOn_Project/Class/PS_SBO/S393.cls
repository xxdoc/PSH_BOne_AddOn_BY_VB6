VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S393"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : S393.cls
'//  Module         :
'//  Description    : 분개장문서
'//  FormType       : 393
'//  Create Date    : 2010.11.26
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
    Set oMat01 = oForm01.Items("76").Specific
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
                If pval.ItemUID = "76" Then
                    If pval.Row Then
                        oMat01Row = pval.Row
                    End If
                End If
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
                If pval.ItemUID = "Btn01" Then
                    Call Print_Report01
                ElseIf pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1291")
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "BPLId02" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    '//라인
                    If pval.ItemUID = "76" And (pval.ColUID = "U_BPLId") Then
                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                    End If
                End If
            Case et_CLICK: '//6
                If pval.ItemUID = "76" And pval.ColUID = "U_BPLId" Then
                    If oMat01.VisualRowCount > 1 And oMat01.Columns("U_BPLId").Cells(pval.Row).Specific.Value <> "" Then
                        oForm01.Freeze True
                        oForm01.Items("BPLId02").Specific.Select oMat01.Columns("U_BPLId").Cells(pval.Row).Specific.Value
                        oForm01.Freeze False
                    End If
                End If
            
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
                Call Form_Resize(FormUID, pval, BubbleEvent)
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
                Call oForm01.Items("BPLId01").Specific.Select(User_BPLId(), psk_ByValue)
                oForm01.Items("6").Click ct_Regular
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
    Dim oEdit           As SAPbouiCOM.EditText
    Dim optBtn          As SAPbouiCOM.OptionBtn
    Dim oCombo01        As SAPbouiCOM.ComboBox
    Dim oCombo02        As SAPbouiCOM.ComboBox
    Dim oStatic         As SAPbouiCOM.StaticText
        
    Dim oRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)

    oForm01.Freeze True
    
    '//회계전표 버튼
    Set newItem = oForm01.Items.Add("Btn01", it_BUTTON)
    Set stdItem = oForm01.Items("2")
    newItem.Left = stdItem.Left + 90
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = 100
    newItem.Specific.Caption = "회계 전표"
    
    '//전결 라디오 버튼
    Set newItem = oForm01.Items.Add("RBtn01", it_OPTION_BUTTON)
    newItem.Left = stdItem.Left + 200
    newItem.Top = stdItem.Top - 8
    newItem.Height = stdItem.Height
    newItem.Width = 50
    newItem.Specific.Caption = "담당"
    
    '//전결 라디오 버튼
    Set newItem = oForm01.Items.Add("RBtn02", it_OPTION_BUTTON)
    newItem.Left = stdItem.Left + 200
    newItem.Top = stdItem.Top + 10
    newItem.Height = stdItem.Height
    newItem.Width = 50
    newItem.Specific.Caption = "차장"
    
     '//전결 라디오 버튼
    Set newItem = oForm01.Items.Add("RBtn03", it_OPTION_BUTTON)
    newItem.Left = stdItem.Left + 250
    newItem.Top = stdItem.Top - 8
    newItem.Height = stdItem.Height
    newItem.Width = 70
    newItem.Specific.Caption = "팀장"
    
    '//전결 라디오 버튼
    Set newItem = oForm01.Items.Add("RBtn04", it_OPTION_BUTTON)
    newItem.Left = stdItem.Left + 250
    newItem.Top = stdItem.Top + 10
    newItem.Height = stdItem.Height
    newItem.Width = 70
    newItem.Specific.Caption = "사업부장"
    
     '//전결 라디오 버튼
    Set newItem = oForm01.Items.Add("RBtn05", it_OPTION_BUTTON)
    newItem.Left = stdItem.Left + 325
    newItem.Top = stdItem.Top - 8
    newItem.Height = stdItem.Height
    newItem.Width = 70
    newItem.Specific.Caption = "전무"
    
    '//전결 라디오 버튼
    Set newItem = oForm01.Items.Add("RBtn06", it_OPTION_BUTTON)
    newItem.Left = stdItem.Left + 325
    newItem.Top = stdItem.Top + 10
    newItem.Height = stdItem.Height
    newItem.Width = 70
    newItem.Specific.Caption = "사장"
    
    '//라디오버튼
    oForm01.DataSources.UserDataSources.Add "RadioBtn01", dt_SHORT_TEXT, 1
    Set optBtn = oForm01.Items("RBtn01").Specific
    optBtn.ValOn = "1"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn01"
    optBtn.Selected = True
    
    Set optBtn = oForm01.Items("RBtn02").Specific
    optBtn.ValOn = "2"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn01"
    optBtn.GroupWith ("RBtn01")
    
    Set optBtn = oForm01.Items("RBtn03").Specific
    optBtn.ValOn = "3"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn01"
    optBtn.GroupWith ("RBtn01")
    
    Set optBtn = oForm01.Items("RBtn04").Specific
    optBtn.ValOn = "4"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn01"
    optBtn.GroupWith ("RBtn01")
    
    Set optBtn = oForm01.Items("RBtn05").Specific
    optBtn.ValOn = "5"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn01"
    optBtn.GroupWith ("RBtn01")
    
    Set optBtn = oForm01.Items("RBtn06").Specific
    optBtn.ValOn = "6"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "RadioBtn01"
    optBtn.GroupWith ("RBtn01")
    
    oForm01.DataSources.UserDataSources("RadioBtn01").Value = "0"
    
    '//사업장-ComboBox
    Set newItem = oForm01.Items.Add("Static01", it_STATIC)
    Set stdItem = oForm01.Items("2006")
    newItem.Left = stdItem.Left + 93
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
'    newItem.LinkTo = "TradeType"
    Set oStatic = newItem.Specific
    oStatic.Caption = "사업장"
    
    Set newItem = oForm01.Items.Add("BPLId01", it_COMBO_BOX)
    Set stdItem = oForm01.Items("2007")
    newItem.Left = stdItem.Left + 93
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width + 40
    newItem.DisplayDesc = True
    Set oCombo01 = newItem.Specific
    Call oCombo01.DataBind.SetBound(True, "OBTF", "U_BPLId")

    sQry = "select BPLId, BPLName from [OBPL] order by BPLId"
    oRecordset.DoQuery sQry
    Do Until oRecordset.EOF
        oCombo01.ValidValues.Add Trim(oRecordset.Fields(0).Value), Trim(oRecordset.Fields(1).Value)
        oRecordset.MoveNext
    Loop
    Call oCombo01.Select(User_BPLId(), psk_ByValue)

    '//사업장-ComboBox
    Set newItem = oForm01.Items.Add("Static02", it_STATIC)
    Set stdItem = oForm01.Items("2001")
    newItem.Left = stdItem.Left + 161
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 2
    newItem.ToPane = 2
'    newItem.LinkTo = "TradeType"
    Set oStatic = newItem.Specific
    oStatic.Caption = "사업장"
    
    Set newItem = oForm01.Items.Add("BPLId02", it_COMBO_BOX)
    Set stdItem = oForm01.Items("2000")
    newItem.Left = stdItem.Left + 161
    newItem.Top = stdItem.Top
    newItem.Height = stdItem.Height
    newItem.Width = stdItem.Width
    newItem.FromPane = 2
    newItem.ToPane = 2
    newItem.DisplayDesc = True
    Set oCombo02 = newItem.Specific
    Call oCombo02.DataBind.SetBound(True, "BTF1", "U_BPLId")
    
'    sQry = "select BPLId, BPLName from [OBPL] order by BPLId"
'    oRecordset.DoQuery sQry
'    Do Until oRecordset.EOF
'        oCombo02.ValidValues.Add Trim(oRecordset.Fields(0).Value), Trim(oRecordset.Fields(1).Value)
'        oRecordset.MoveNext
'    Loop

    oForm01.Freeze False
    
    Set oCombo01 = Nothing
    Set oCombo02 = Nothing
    Set oRecordset = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    Set oCombo01 = Nothing
    Set oCombo02 = Nothing
    Set oRecordset = Nothing
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, j&, DocNum$, LineId$, Counts$, Cnt$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Dim oRecordset    As SAPbobsCOM.Recordset
        
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "BPLId02"
            oForm01.Freeze True
            For i = 1 To oMat01.VisualRowCount
                If oMat01Row = i Then
                    If oMat01.Columns("1").Cells(i).Specific.Value <> "" Then
                        oMat01.Columns("U_BPLId").Cells(i).Specific.Select Trim(oForm01.Items("BPLId02").Specific.Selected.Value)
                    End If
                End If
            Next i
            oForm01.Freeze False
        
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "76" Then
        Select Case oCol
            Case "U_BPLId"
                oForm01.Freeze True
                oForm01.Items("BPLId02").Specific.Select oMat01.Columns("U_BPLId").Cells(oRow).Specific.Value
                oForm01.Freeze False
        End Select
    End If
     
    Set oRecordset = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub Form_Resize(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Form_Resize_Error
    
     oForm01.Items("Static01").Left = oForm01.Items("2006").Left + 93
     oForm01.Items("BPLId01").Left = oForm01.Items("2007").Left + 93
     oForm01.Items("Static02").Top = oForm01.Items("2001").Top
     oForm01.Items("Static02").Left = oForm01.Items("2001").Left + 161
     oForm01.Items("BPLId02").Top = oForm01.Items("2000").Top
     oForm01.Items("BPLId02").Left = oForm01.Items("2000").Left + 161

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Form_Resize_Error:
    Sbo_Application.SetStatusBarMessage "Form_Resize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Print_Report01()
On Error GoTo Print_Report01_Error
    Dim i&
    Dim ErrNum&
    Dim TransId                 As String
    Dim WinTitle                As String
    Dim ReportName              As String
    Dim sQry                    As String
    Dim oText(1)                As String
    Dim oRecordSet01            As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call ConnectODBC
    
    TransId = Trim(oForm01.Items("5").Specific.Value)
    
    If TransId = "" Then
        ErrNum = 1
        GoTo Print_Report01_Error
    End If
    
    WinTitle = "회계전표 [PS_FI010]"
    ReportName = "PS_FI010_01.rpt"
       
    '//Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    oText(1) = oForm01.DataSources.UserDataSources("RadioBtn01").Value
    
    For i = 1 To 1
        If Len("" & i & "") = 1 Then
            gRpt_Formula(i) = "F0" & i & ""
        Else
            gRpt_Formula(i) = "F" & i & ""
        End If
        gRpt_Formula_Value(i) = oText(i)
    Next i
    
    '//SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '//조회조건문
    sQry = "EXEC [PS_FI010_03] '" & TransId & "'"
'    oRecordSet01.DoQuery sQry
'    If oRecordSet01.RecordCount = 0 Then
'        ErrNum = 2
'        GoTo Print_Report01_Error
'    End If
    
    '//CR Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report01_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

