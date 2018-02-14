VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY683"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PH_PY683.cls
'//  Module : 인사관리 > 근태관리 > 근태리포트
'//  Desc : 교대근무인정현황
'//  FormType : PH_PY683
'//  Create Date(Start) : 2016.01.28
'//  Creator : NGY
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml As String
    Dim oXmlDoc As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY683.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY683_" & GetTotalFormsCount
    Call AddForms(Me, oFormUniqueID01, "PH_PY683") '//폼추가
    Call Sbo_Application.LoadBatchActions(oXmlDoc.xml)
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    '//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    Call oForm.Freeze(True)
    Call PH_PY683_CreateItems
    Call PH_PY683_ComboBox_Setting
    Call PH_PY683_CF_ChooseFromList
    Call PH_PY683_EnableMenus
    Call PH_PY683_SetDocument(oFromDocEntry01)
    Call PH_PY683_FormResize
    
    Call oForm.EnableMenu("1283", False) '삭제
    Call oForm.EnableMenu("1286", False) '닫기
    Call oForm.EnableMenu("1287", False) '복제
    Call oForm.EnableMenu("1285", False) '복원
    Call oForm.EnableMenu("1284", False) '취소
    Call oForm.EnableMenu("1293", False) '행삭제
    Call oForm.EnableMenu("1281", False)
    Call oForm.EnableMenu("1282", True)
  
    
    
    Call oForm.Update
    Call oForm.Freeze(False)
    
    oForm.Visible = True
    Set oXmlDoc = Nothing

    Exit Sub
LoadForm_Error:
    Call oForm.Update
    Call oForm.Freeze(False)
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Call Sbo_Application.SetStatusBarMessage("Form_Load Error:" & Err.Description, bmt_Short, True)
End Sub

Public Sub PH_PY683_MTX01()
'******************************************************************************
'Function ID : PH_PY683_MTX01()
'해당모듈 : PH_PY683
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY683_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String
    Dim YM As String
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    YM = Trim(oForm.Items("YM").Specific.VALUE)
    
    Call oForm.Freeze(True)
    
    sQry = "EXEC [PH_PY683_01] '" & CLTCOD & "', '" & YM & "'"
    
    oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable")
    Call oGrid01.DataTable.Clear
    Call oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry)
    
    If oGrid01.Rows.Count = 0 Then
        ErrNum = 1
        GoTo PH_PY683_MTX01_Error
    End If
        
    Call oGrid01.AutoResizeColumns
    Call oForm.Update
    
    Call oForm.Freeze(False)
  
'    Set ProgBar01 = Nothing
'    Set RecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY683_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
'    Set ProgBar01 = Nothing
'    Set RecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY683_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Function PH_PY683_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PH_PY683_HeaderSpaceLineDel()
'해당모듈 : PH_PY683
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY683_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm.Items("YM").Specific.VALUE) = ""
            ErrNum = 1
            GoTo PH_PY683_HeaderSpaceLineDel_Error

    End Select

    PH_PY683_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY683_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "기준년월는 필수조회 조건입니다. 확인하세요.", "E"
        Call oForm.Items("YM").CLICK(ct_Regular)
    End If
    PH_PY683_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PH_PY683_MatrixSpaceLineDel() As Boolean
On Error GoTo PH_PY683_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PH_PY683_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY683_MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 사원코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 시간이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 등록일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 비가동코드가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PH_PY683_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PH_PY683_MatrixSpaceLineDel = False
End Function

Private Sub PH_PY683_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY683_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String
    Dim TeamCode As String
    Dim RspCode As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    
    Call oForm.Freeze(True)
    
    Select Case oUID
            
        
        
    End Select

    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PH_PY683_FlushToItemValue_Error:

    Call oForm.Freeze(False)
'    Call ProgBar01.Stop
'    ProgBar01.VALUE = 100
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing

    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("PH_PY683_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If

End Sub

'/폼의 아이템 사용지정
Sub PH_PY683_FormItemEnabled()
On Error GoTo PH_PY683_FormItemEnabled_Error
    If (oForm.Mode = fm_ADD_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_FIND_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
'        oForm.Items("ItemCode").Enabled = True

    ElseIf (oForm.Mode = fm_OK_MODE) Then
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'        Call CLTCOD_Select(oForm, "SCLTCOD")
        
    End If
    
    Exit Sub
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY683_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY683_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/아이템 변경 이벤트
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error

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
Raise_FormItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormMenuEvent_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
'                Call PH_PY683_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
'                oForm.Mode = fm_ADD_MODE
'                BubbleEvent = False
'                Call PH_PY683_LoadCaption
                
                'oForm.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PH_PY683_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PH_PY683A.Clear

'                Call PH_PY683_LoadCaption
'                Call PH_PY683_FormItemEnabled
                '//Call PH_PY683_FormItemEnabled '//UDO방식
                '//Call PH_PY683_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY683_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
        End Select
    End If
    Exit Sub
Raise_FormMenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    
        If pval.ItemUID = "PH_PY683" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
            End If
        End If
            
        If pval.ItemUID = "BtnSearch" Then '/조회
            
            If PH_PY683_HeaderSpaceLineDel = False Then
                BubbleEvent = False
                Exit Sub
            End If
        
            Call PH_PY683_MTX01
        
'        ElseIf pval.ItemUID = "BtnPrint" Then
'
'            If PH_PY683_HeaderSpaceLineDel = False Then
'                BubbleEvent = False
'                Exit Sub
'            End If
'
'            Call PH_PY683_Print_Report01
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PH_PY683" Then
            If oForm.Mode = fm_ADD_MODE Then
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            ElseIf oForm.Mode = fm_OK_MODE Then
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
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD", "") '사번
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        
        Call PH_PY683_FlushToItemValue(pval.ItemUID)
        
    End If

    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

    Call oForm.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
            
            Call PH_PY683_FlushToItemValue(pval.ItemUID)
            
        End If
        
    ElseIf pval.BeforeAction = False Then

    End If
    
    Call oForm.Freeze(False)
    
    Exit Sub
    
Raise_EVENT_VALIDATE_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PH_PY683_FormItemEnabled
        '//Call PH_PY683_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PH_PY683_FormResize
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
'            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY683A", "U_CardCode,U_CardName")
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
        Set oForm = Nothing
        Set oGrid01 = Nothing
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

        ElseIf pval.BeforeAction = False Then
            
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY683_CreateItems() As Boolean
On Error GoTo PH_PY683_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm.Items("Grid01").Specific
    Call oForm.DataSources.DataTables.Add("DataTable")
    oGrid01.SelectionMode = ms_Auto

    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    
    '기준년월
    Call oForm.DataSources.UserDataSources.Add("YM", dt_SHORT_TEXT, 6)
    Call oForm.Items("YM").Specific.DataBind.SetBound(True, "", "YM")
    
    oForm.Items("YM").Specific.String = Format(Now, "YYYYMM")
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY683_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY683_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PH_PY683_ComboBox_Setting()
On Error GoTo PH_PY683_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)

    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY683_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY683_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY683_CF_ChooseFromList()
On Error GoTo PH_PY683_CF_ChooseFromList_Error

    Exit Sub
PH_PY683_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY683_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY683_EnableMenus()
On Error GoTo PH_PY683_EnableMenus_Error
    
    Exit Sub
PH_PY683_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY683_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY683_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY683_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PH_PY683_FormItemEnabled
        '//Call PH_PY683_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm.Mode = fm_FIND_MODE
'        Call PH_PY683_FormItemEnabled
'        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm.Items("1").Click ct_Regular
    End If
    Exit Sub
PH_PY683_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY683_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY683_FormResize()
On Error GoTo PH_PY683_FormResize_Error
    
    Exit Sub
PH_PY683_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY683_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'Private Sub PH_PY683_Print_Report01()
'    On Error GoTo PH_PY683_Print_Report01_Error
'
'    Dim WinTitle As String
'    Dim ReportName As String
'    Dim sQry As String
'
'    Dim CLTCOD As String '사업장
'    Dim TeamCode As String '팀
'    Dim RspCode As String '담당
'    Dim ClsCode As String '반
'    Dim MSTCOD As String '사원번호
'    Dim yyyy As String '기준년도
'
'    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
'    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '팀
'    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE) '담당
'    ClsCode = Trim(oForm.Items("ClsCode").Specific.VALUE) '반
'    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사원번호
'    yyyy = Trim(oForm.Items("YYYY").Specific.VALUE) '기준년도
'
'    Dim ProgBar01 As SAPbouiCOM.ProgressBar
'    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
'
'     '/ ODBC 연결 체크
'    If ConnectODBC = False Then
'        GoTo PH_PY683_Print_Report01_Error
'    End If
'
'    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'
'    WinTitle = "[PH_PY683] 비근무일수 현황"
'
'    ReportName = "PH_PY683_01.rpt"
'
'    '// Formula 수식필드
'    ReDim gRpt_Formula(1)
'    ReDim gRpt_Formula_Value(1)
'
'    '// SubReport
'    ReDim gRpt_SRptSqry(1)
'    ReDim gRpt_SRptName(1)
'
'    ReDim gRpt_SFormula(1, 1)
'    ReDim gRpt_SFormula_Value(1, 1)
'
'    gRpt_SFormula(1, 1) = ""
'    gRpt_SFormula_Value(1, 1) = ""
'
'    '/ Procedure 실행"
'    sQry = "                EXEC [PH_PY683_02] "
'    sQry = sQry & "'" & CLTCOD & "'," '사업장
'    sQry = sQry & "'" & TeamCode & "',"  '팀
'    sQry = sQry & "'" & RspCode & "',"  '담당
'    sQry = sQry & "'" & ClsCode & "',"  '반
'    sQry = sQry & "'" & MSTCOD & "',"  '사번
'    sQry = sQry & "'" & yyyy & "'"  '기준년도
'
'    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V", "", 2) = False Then
'        GoTo PH_PY683_Print_Report01_Error
'    End If
'
'    ProgBar01.VALUE = 100
'    ProgBar01.Stop
'    Set ProgBar01 = Nothing
'
'    Exit Sub
'
'PH_PY683_Print_Report01_Error:
'
'    ProgBar01.VALUE = 100
'    ProgBar01.Stop
'    Set ProgBar01 = Nothing
'
'    Call MDC_Com.MDC_GF_Message("Print_Query_Error:" & Err.Number & " - " & Err.Description, "E")
'
'End Sub











