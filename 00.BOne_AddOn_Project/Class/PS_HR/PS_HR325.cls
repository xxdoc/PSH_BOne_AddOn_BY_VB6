VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR325"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_HR325.cls
'//  Module         : HR
'//  Desc           : KPI목표등록 마감
'//  FormType       : PS_HR325
'//  Create Date    : 2012.2.15
'//  Creator        : N.G.Y
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_HR325L As SAPbouiCOM.DBDataSource '등록라인

Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String  '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long    '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMat01Row01 As Long
Private oText(4)    As String

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR325.srf")
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
    
    oFormUniqueID01 = "PS_HR325_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
        
    oForm01.Freeze True
    Call PS_HR325_CreateItems
    Call PS_HR325_ComboBox_Setting
    Call PS_HR325_Initial_Setting
    Call PS_HR325_EnableMenus
    
    
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
            
        Case et_COMBO_SELECT: '//5
            Call Raise_EVENT_COMBO_SELECT(FormUID, pval, BubbleEvent)
        Case et_CLICK: '//6
            Call Raise_EVENT_CLICK(FormUID, pval, BubbleEvent)
        Case et_DOUBLE_CLICK: '//7
            Call Raise_EVENT_DOUBLE_CLICK(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LINK_PRESSED '//8
            
        Case et_VALIDATE: '//10
            
        Case et_MATRIX_LOAD: '//11
            
        Case et_FORM_ACTIVATE: '//18
            '//et_FORM_ACTIVATE
        Case et_FORM_DEACTIVATE: '//19
            '//et_FORM_DEACTIVATE
        Case et_FORM_RESIZE '//20
            
        Case et_CHOOSE_FROM_LIST '//27
            
        Case et_GOT_FOCUS: '//3
            
        Case et_LOST_FOCUS: '//4
            '//et_LOST_FOCUS
        Case et_FORM_UNLOAD: '//17
           
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
                '//Call PS_HR325_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_HR325_FormItemEnabled '//UDO방식
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
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_HR325_MTX01 '//조회
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Btn02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_HR325_Save '//저장
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        
    ElseIf pval.BeforeAction = False Then
        
    End If

    Exit Sub
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "Mat01" And pval.Row = "0" And pval.ColUID = "FinYN" Then
            Dim Chk$
            Dim i As Long
            oForm01.Freeze True
            oMat01.FlushToDataSource
            If Trim(oDS_PS_HR325L.GetValue("U_ColReg05", 0)) = "" Or Trim(oDS_PS_HR325L.GetValue("U_ColReg05", 0)) = "N" Then
                Chk = "Y"
            ElseIf Trim(oDS_PS_HR325L.GetValue("U_ColReg05", 0)) = "Y" Then
                Chk = "N"
            End If
            For i = 0 To oMat01.VisualRowCount - 1
                oDS_PS_HR325L.setValue "U_ColReg05", i, Chk
            Next i
            oMat01.LoadFromDataSource
            oForm01.Freeze False
        End If
    End If
    Exit Sub
Raise_EVENT_DOUBLE_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_DOUBLE_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_HR325_CreateItems() As Boolean
On Error GoTo PS_HR325_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_HR325L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    
    Call oForm01.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 100)
    Call oForm01.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")
    
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 20)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    
    Call oForm01.DataSources.UserDataSources.Add("Team", dt_SHORT_TEXT, 20)
    Call oForm01.Items("Team").Specific.DataBind.SetBound(True, "", "Team")
    
    Call oForm01.DataSources.UserDataSources.Add("Rsp", dt_SHORT_TEXT, 20)
    Call oForm01.Items("Rsp").Specific.DataBind.SetBound(True, "", "Rsp")
    
    Call oForm01.DataSources.UserDataSources.Add("FinStatus", dt_SHORT_TEXT, 20)
    Call oForm01.Items("FinStatus").Specific.DataBind.SetBound(True, "", "FinStatus")
    
    
    oForm01.Items("Mat01").Enabled = True
    oForm01.Items("StdYear").Specific.VALUE = Format(Now, "YYYY")
    
    oMat01.Columns("DocTypeCd").Visible = False '문서타입 Hidden
    oMat01.Columns("DocEntry").Visible = False '문서번호 Hidden
    
        
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_HR325_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR325_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_HR325_ComboBox_Setting()
On Error GoTo PS_HR325_ComboBox_Setting_Error
    Dim i As Long
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry As String
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    ' 사업부
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)

    ' 마감상태
    Set oCombo = oForm01.Items("FinStatus").Specific
    oCombo.ValidValues.Add "%", "전체"
    oCombo.ValidValues.Add "N", "진행"
    oCombo.ValidValues.Add "Y", "마감"
    oCombo.Select "0", psk_Index
    
    ' 팀 //테이블 select 부분 추가 해야함
    'Call MDC_SetMod.Set_ComboList(oForm01.Items("Team").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y'", "101", False, False)
    ' 담당 //테이블 select 부분 추가 해야함
    'Call MDC_SetMod.Set_ComboList(oForm01.Items("Rsp").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y'", "101", False, False)
    
    oMat01.Columns("FinYN").ValidValues.Add "N", "진행"
    oMat01.Columns("FinYN").ValidValues.Add "Y", "마감"
    
    '매트릭스-사업부콤보_S
    sQry = "            SELECT      BPLId,"
    sQry = sQry & "                 BPLName"
    sQry = sQry & "  FROM       OBPL"
    sQry = sQry & "  ORDER BY  BPLID"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), sQry)
    '매트릭스-사업부콤보_E
    
    '매트릭스-팀콤보_S
    sQry = "           SELECT      b.U_Code, "
    sQry = sQry & "                b.U_CodeNm"
    sQry = sQry & " FROM       [@PS_HR200H] a"
    sQry = sQry & "                Inner Join"
    sQry = sQry & "                [@PS_HR200L] b"
    sQry = sQry & "                    On a.Code = b.Code"
    sQry = sQry & "                    And a.Name = '부서'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Team"), sQry)
    '매트릭스-팀콤보_S
    
    '매트릭스-담당콤보_S
    sQry = "           SELECT      b.U_Code, "
    sQry = sQry & "                b.U_CodeNm"
    sQry = sQry & " FROM       [@PS_HR200H] a"
    sQry = sQry & "                Inner Join"
    sQry = sQry & "                [@PS_HR200L] b"
    sQry = sQry & "                    On a.Code = b.Code"
    sQry = sQry & "                    And a.Name = '담당'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Rsp"), sQry)
    '매트릭스-담당콤보_S
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR325_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR325_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR325_Initial_Setting()
On Error GoTo PS_HR325_Initial_Setting_Error
 ' 사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

    Exit Sub
PS_HR325_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR325_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR325_EnableMenus()
On Error GoTo PS_HR325_EnableMenus_Error
    '//메뉴활성화
    Call oForm01.EnableMenu("1281", False)
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_HR325_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR325_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR325_MTX01()
On Error GoTo PS_HR325_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    
    Dim Temp01 As String
    Dim Temp02 As String
    Dim Temp03 As String
    
    Param01 = Trim(oForm01.Items("StdYear").Specific.VALUE)
    Param02 = Trim(oForm01.Items("FinStatus").Specific.VALUE)
    Param03 = Trim(oForm01.Items("BPLId").Specific.VALUE)
    Param04 = Trim(oForm01.Items("Team").Specific.VALUE)
    Param05 = Trim(oForm01.Items("Rsp").Specific.VALUE)
    
        
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_HR325_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
'        oForm01.Items("Mat01").Enabled = False
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
        GoTo PS_HR325_MTX01_Exit
'    Else
'        oForm01.Items("Mat01").Enabled = True
    End If
    
    
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_HR325L.InsertRecord (i)
        End If
        oDS_PS_HR325L.Offset = i
        Call oDS_PS_HR325L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_HR325L.setValue("U_ColReg01", i, RecordSet01.Fields("BPLId").VALUE)
        Call oDS_PS_HR325L.setValue("U_ColReg02", i, RecordSet01.Fields("Team").VALUE)
        Call oDS_PS_HR325L.setValue("U_ColReg03", i, RecordSet01.Fields("Rsp").VALUE)
        Call oDS_PS_HR325L.setValue("U_ColReg04", i, RecordSet01.Fields("EmpNm").VALUE)
        Call oDS_PS_HR325L.setValue("U_ColReg05", i, RecordSet01.Fields("FinYN").VALUE)
        Call oDS_PS_HR325L.setValue("U_ColReg06", i, RecordSet01.Fields("DocEntry").VALUE)
        Call oDS_PS_HR325L.setValue("U_ColReg07", i, RecordSet01.Fields("DocType").VALUE)
        Call oDS_PS_HR325L.setValue("U_ColReg08", i, RecordSet01.Fields("DocTypeCd").VALUE) '문서타입코드
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    Call MDC_Com.MDC_GF_Message("조회를 성공하였습니다.", "S")
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR325_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR325_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR325_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR325_Save()
On Error GoTo PS_HR325_Save_Error

    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.RowCount - 1
        
        Param01 = Trim(oDS_PS_HR325L.GetValue("U_ColReg06", i)) '문서번호
        Param02 = Trim(oDS_PS_HR325L.GetValue("U_ColReg08", i)) 'Object 명(테이블명)
        Param03 = Trim(oDS_PS_HR325L.GetValue("U_ColReg05", i)) '마감여부
        
        Query01 = "                UPDATE   [@" & Param02 & "H]"
        Query01 = Query01 & " SET         U_FinYn = '" & Param03 & "'"
        Query01 = Query01 & " WHERE    DocEntry = " & Param01
        
        Call RecordSet01.DoQuery(Query01)
        
    Next

    Call Sbo_Application.MessageBox("마감처리를 완료하였습니다.")
'    Call Sbo_Application.SetStatusBarMessage("마감처리를 완료하였습니다.", bmt_Short, False)
    
    Set RecordSet01 = Nothing
    Exit Sub
    
PS_HR325_Save_Error:
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_HR325_Save_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    Dim sQry As String
    
    Dim sCount As Integer, sSeq As Integer
    
    Dim BPLID As String
    Dim Team As String
    Dim i As Long
    
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemUID = "BPLId" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_HR325L.Clear
'                    If oForm01.Mode = fm_ADD_MODE Then
'                        Call Add_MatrixRow(0, False)
'                    End If
                                        
                    oForm01.Freeze False
                End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "BPLId" Then
            oForm01.Freeze True
            
            '//해당사업장의 팀을 Reset
            sCount = oForm01.Items("Team").Specific.ValidValues.Count
            sSeq = sCount
            For i = 1 To sCount
                oForm01.Items("Team").Specific.ValidValues.Remove sSeq - 1, psk_Index
                sSeq = sSeq - 1
            Next i
            
            BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
            
            '//해당사업장의 사업팀을 Set
            sQry = "SELECT b.U_Code, b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code And a.Name = '부서' And b.U_Char2 = '" + BPLID + "' order by b.U_Seq"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("Team").Specific.ValidValues.Add "-", "선택"
            
            Do Until oRecordSet01.EOF
                
                oForm01.Items("Team").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                oRecordSet01.MoveNext
            Loop
            
            oForm01.Items("Team").Specific.Select 0, psk_Index
            
                                
            
            oForm01.Freeze False
        ElseIf pval.ItemUID = "Team" Then
            oForm01.Freeze True
            
            '//해당팀의 담당을 Reset
            sCount = oForm01.Items("Rsp").Specific.ValidValues.Count
            sSeq = sCount
            For i = 1 To sCount
                oForm01.Items("Rsp").Specific.ValidValues.Remove sSeq - 1, psk_Index
                sSeq = sSeq - 1
            Next i
            
            Team = oForm01.Items("Team").Specific.VALUE
            
            '//해당사업장의 사업팀을 Set
            sQry = "SELECT b.U_Code, b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code And a.Name = '담당' And b.U_Char1 = '" + Team + "' order by b.U_Seq"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("Rsp").Specific.ValidValues.Add "-", "선택"
            
            Do Until oRecordSet01.EOF
                
                oForm01.Items("Rsp").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                oRecordSet01.MoveNext
            Loop
            
            oForm01.Items("Rsp").Specific.Select 0, psk_Index
            
            oForm01.Freeze False
        End If
                
        
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
