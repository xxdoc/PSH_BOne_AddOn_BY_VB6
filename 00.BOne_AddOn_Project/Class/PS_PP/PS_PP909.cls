VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP909"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_PP909.cls
'//  Module : 생산관리 > 애드온 > 생산리포트
'//  Desc : 설계 실동공수 등록현황(기계)
'//  FormType : PS_PP909
'//  Create Date(Start) : 2012.08.06
'//  Create Date(End) : 2012.08.06
'//  Modified Date :
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form

Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix

Private oDS_PS_PP909H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP909L As SAPbouiCOM.DBDataSource '라인(작업자정보)
Private oDS_PS_PP909M As SAPbouiCOM.DBDataSource '라인(세부정보)

Public oGrid01 As SAPbouiCOM.Grid

Public oDS_PS_PP909N As SAPbouiCOM.DataTable

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
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP909.srf")
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
    
    oFormUniqueID01 = "PS_PP909_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP909_CreateItems
    Call PS_PP909_ComboBox_Setting
    Call PS_PP909_Initial_Setting
    Call PS_PP909_CF_ChooseFromList
    Call PS_PP909_EnableMenus
    Call PS_PP909_SetDocument(oFromDocEntry01)
    Call PS_PP909_FormResize
    
    oForm01.Items("Folder02").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    
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
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_PP909_AddMatrixRow(oMat01.VisualRowCount)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_PP909L.RemoveRecord(oDS_PS_PP909L.Size - 1)
                Call oMat01.LoadFromDataSource
                Call oForm01.Freeze(False)
            
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
    
        If pval.ItemUID = "BtnSrch01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_PP909_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    
                    '상세 정보 초기화
                    oMat02.Clear
                    oDS_PS_PP909M.Clear
                    
                    Call PS_PP909_MTX01 '매트릭스에 데이터 로드
                    
                End If
            End If
            
        ElseIf pval.ItemUID = "BtnSrch02" Then
        
            Call PS_PP909_MTX03 '데이터 로드
            
        ElseIf pval.ItemUID = "Btn_Print" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
        
                Call PS_PP909_Print_Report01
            
            End If
        
        ElseIf pval.ItemUID = "BtnPrt02" Then
        
            Call PS_PP909_Print_Report02
        
        End If
    ElseIf pval.BeforeAction = False Then
    
        '폴더를 사용할 때는 필수 소스_S
        'Folder01이 선택되었을 때
        If pval.ItemUID = "Folder01" Then
            
            oForm01.PaneLevel = 1
            oForm01.DefButton = "BtnSrch01"
            
        End If
        
        'Folder02가 선택되었을 때
        If pval.ItemUID = "Folder02" Then
            
            oForm01.PaneLevel = 2
            oForm01.DefButton = "BtnSrch02"
            
        End If
    
        If pval.ItemUID = "PS_PP909" Then
        
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
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '사원코드 포맷서치 활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode02", "") '사원코드 포맷서치 활성
        
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
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                
                Call oMat01.SelectRow(pval.Row, True, False)
                
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
            End If
            
        ElseIf pval.ItemUID = "Mat02" Then
            If pval.Row > 0 Then
                
                Call oMat02.SelectRow(pval.Row, True, False)
                
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
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
    
        If pval.ItemUID = "Mat01" Then '작업자 정보

            If pval.Row = 0 Then

                '정렬
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource

            Else

                Call PS_PP909_MTX02(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE)

            End If

        ElseIf pval.ItemUID = "Mat02" Then '세부 정보

            If pval.Row = 0 Then
            
                '정렬
                oMat02.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat02.FlushToDataSource
            
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
    
        If pval.ItemUID = "Mat01" Then
    
            Dim PP040 As Object
            Set PP040 = New PS_PP040
    
            Call PP040.LoadForm(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE)
    
            Set PP040 = Nothing
        
        ElseIf pval.ItemUID = "Mat02" Then
        
            Dim PP030 As Object
            Set PP030 = New PS_PP030
    
            Call PP030.LoadForm(oMat02.Columns("PP030HNo").Cells(pval.Row).Specific.VALUE)
    
            Set PP030 = Nothing
        
        End If
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Set PP030 = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "CntcCode") Then
            
                oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
            ElseIf (pval.ItemUID = "CntcCode02") Then
            
                oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("CntcName02").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
            End If
            
            oForm01.Items(pval.ItemUID).Click ct_Regular
        End If
    ElseIf pval.BeforeAction = False Then

    End If
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_PP909_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP909_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
    
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
    ElseIf pval.ItemUID = "Mat02" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat03" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat04" Then
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
        Set oMat02 = Nothing
    End If
    Exit Sub
Raise_EVENT_FORM_UNLOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_UNLOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP909_CreateItems() As Boolean
On Error GoTo PS_PP909_CreateItems_Error
    Call oForm01.Freeze(True)
'    Dim oQuery01 As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

'    Set oDS_PS_PP909H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_PP909L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_PP909M = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    
    '매트릭스 초기화
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    
    Call oForm01.DataSources.DataTables.Add("PS_PP909N")
    
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("PS_PP909N")
    
    Set oDS_PS_PP909N = oForm01.DataSources.DataTables.Item("PS_PP909N")
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    '사업장_E
    
    '작업자코드_S
    Call oForm01.DataSources.UserDataSources.Add("CntcCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CntcCode").Specific.DataBind.SetBound(True, "", "CntcCode")
    '작업자코드_E
    
    '작업자성명_S
    Call oForm01.DataSources.UserDataSources.Add("CntcName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CntcName").Specific.DataBind.SetBound(True, "", "CntcName")
    '작업자성명_E
    
    '작업일자Fr_S
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    '작업일자Fr_E
    
    '작업일자To_S
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    '작업일자To_E
    
    
    '사업장2_S
    Call oForm01.DataSources.UserDataSources.Add("BPLID02", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID02").Specific.DataBind.SetBound(True, "", "BPLID02")
    '사업장2_E
    
    '작업자코드2_S
    Call oForm01.DataSources.UserDataSources.Add("CntcCode02", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CntcCode02").Specific.DataBind.SetBound(True, "", "CntcCode02")
    '작업자코드2_E
    
    '작업자성명2_S
    Call oForm01.DataSources.UserDataSources.Add("CntcName02", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CntcName02").Specific.DataBind.SetBound(True, "", "CntcName02")
    '작업자성명2_E
    
    '작업일자Fr2_S
    Call oForm01.DataSources.UserDataSources.Add("FrDt02", dt_DATE)
    Call oForm01.Items("FrDt02").Specific.DataBind.SetBound(True, "", "FrDt02")
    '작업일자Fr2_E
    
    '작업일자To2_S
    Call oForm01.DataSources.UserDataSources.Add("ToDt02", dt_DATE)
    Call oForm01.Items("ToDt02").Specific.DataBind.SetBound(True, "", "ToDt02")
    '작업일자To2_E
        
    Call oForm01.Freeze(False)
    Exit Function
PS_PP909_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP909_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP909_ComboBox_Setting()
On Error GoTo PS_PP909_ComboBox_Setting_Error
    
    Call oForm01.Freeze(True)

    '사업장_S
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName From [OBPL] order by 1", MDC_PS_Common.User_BPLId(), False, False)
    '사업장_E
    
    '사업장02_S
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID02").Specific, "SELECT BPLId, BPLName From [OBPL] order by 1", MDC_PS_Common.User_BPLId(), False, False)
    '사업장02_E

    Call oForm01.Freeze(False)
    Exit Sub
PS_PP909_ComboBox_Setting_Error:
'    Set lRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP909_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP909_Initial_Setting()
On Error GoTo PS_PP909_Initial_Setting_Error

    '사업장 사용자의 소속 사업장 선택
'    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '기준일자
'    oForm01.Items("StdDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    '날짜 설정
    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYYMMDD")
    
    oForm01.Items("FrDt02").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    oForm01.Items("ToDt02").Specific.VALUE = Format(Date, "YYYYMMDD")
    
    '날짜 빈칸으로 설정
'    oForm01.Items("ToDt").Specific.VALUE = ""
'    oForm01.Items("FrDt").Specific.VALUE = ""
 
    Exit Sub
PS_PP909_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP909_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_PP909_CF_ChooseFromList()
On Error GoTo PS_PP909_CF_ChooseFromList_Error
    Exit Sub
PS_PP909_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP909_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP909_FormItemEnabled()
On Error GoTo PS_PP909_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP909_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP909_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP909_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP909_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP909L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP909L.Offset = oRow
'    oDS_PS_PP909L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP909_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP909_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP909_FormClear()
On Error GoTo PS_PP909_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP909'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP909_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP909_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP909_EnableMenus()
On Error GoTo PS_PP909_EnableMenus_Error
    Exit Sub
PS_PP909_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP909_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP909_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP909_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP909_FormItemEnabled

    Else

    End If
    Exit Sub
PS_PP909_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP909_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP909_DataValidCheck() As Boolean
On Error GoTo PS_PP909_DataValidCheck_Error
    PS_PP909_DataValidCheck = False
    Dim i As Long

'    If oForm01.Items("DocDateTo").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "견적일자 To은 필수입니다.", bmt_Short, True
'        oForm01.Items("DocDateTo").Click ct_Regular
'        PS_PP909_DataValidCheck = False
'        Exit Function
'    End If

    PS_PP909_DataValidCheck = True
    Exit Function
PS_PP909_DataValidCheck_Error:
    PS_PP909_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP909_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP909_MTX01()
On Error GoTo PS_PP909_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String '사업장
    Dim CntcCode As String '작업자코드
    Dim FrDt As String '시작일자
    Dim ToDt As String '종료일자
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP909_01 '" & BPLID & "','" & FrDt & "','" & ToDt & "','" & CntcCode & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_PP909_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP909L.InsertRecord(loopCount)
        End If
        oDS_PS_PP909L.Offset = loopCount
        
        Call oDS_PS_PP909L.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP909L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("DocEntry").VALUE) '작업일보문서번호
        Call oDS_PS_PP909L.setValue("U_ColDt01", loopCount, RecordSet01.Fields("DocDate").VALUE) '작업일보등록일자
        Call oDS_PS_PP909L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("CntcCode").VALUE) '사번
        Call oDS_PS_PP909L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("CntcName").VALUE) '성명
        Call oDS_PS_PP909L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("YTime").VALUE) '작업시간
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP909_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP909_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP909_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP909_MTX02(ByVal prmDocEntry As String)
On Error GoTo PS_PP909_MTX02_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP909_02 '" & prmDocEntry & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat02.Clear
        GoTo PS_PP909_MTX02_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP909M.InsertRecord(loopCount)
        End If
        oDS_PS_PP909M.Offset = loopCount
        
        Call oDS_PS_PP909M.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP909M.setValue("U_ColReg01", loopCount, RecordSet01.Fields("PP030HNo").VALUE) '작업지시문서번호
        Call oDS_PS_PP909M.setValue("U_ColReg02", loopCount, RecordSet01.Fields("PP030MNo").VALUE) '작업지시라인번호
        Call oDS_PS_PP909M.setValue("U_ColReg03", loopCount, RecordSet01.Fields("OrdNum").VALUE) '메인작번
        Call oDS_PS_PP909M.setValue("U_ColReg04", loopCount, RecordSet01.Fields("OrdSub1").VALUE) '서브작번1
        Call oDS_PS_PP909M.setValue("U_ColReg05", loopCount, RecordSet01.Fields("OrdSub2").VALUE) '서브작번2
        Call oDS_PS_PP909M.setValue("U_ColReg06", loopCount, RecordSet01.Fields("WorkTime").VALUE) '공정정보작업시간
        Call oDS_PS_PP909M.setValue("U_ColReg07", loopCount, RecordSet01.Fields("BdwQty").VALUE) '기존도면매수
        Call oDS_PS_PP909M.setValue("U_ColReg08", loopCount, RecordSet01.Fields("DwRate").VALUE) '적용율
        Call oDS_PS_PP909M.setValue("U_ColReg09", loopCount, RecordSet01.Fields("AdwQty").VALUE) '적용매수
        Call oDS_PS_PP909M.setValue("U_ColReg10", loopCount, RecordSet01.Fields("NdwQty").VALUE) '신규도면매수
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP909_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP909_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP909_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP909_MTX03()
On Error GoTo PS_PP909_MTX03_Error
        
    Call oForm01.Freeze(True)
    
    Dim i As Long
    Dim Query01 As String
    Dim ErrNum As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID As String
    Dim CntcCode As String
    Dim FrDt As String
    Dim ToDt As String
    
    BPLID = oForm01.Items("BPLID02").Specific.VALUE
    CntcCode = oForm01.Items("CntcCode02").Specific.VALUE
    FrDt = oForm01.Items("FrDt02").Specific.VALUE
    ToDt = oForm01.Items("ToDt02").Specific.VALUE
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Query01 = "              EXEC PS_PP909_03 '"
    Query01 = Query01 & BPLID & "','"
    Query01 = Query01 & CntcCode & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "'"

    Call oGrid01.DataTable.Clear
    Call oDS_PS_PP909N.ExecuteQuery(Query01)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    oGrid01.Columns(8).RightJustified = True
'    oGrid01.Columns(4).RightJustified = True
'    oGrid01.Columns(5).RightJustified = True
'    oGrid01.Columns(6).RightJustified = True
'    oGrid01.Columns(7).RightJustified = True
'    oGrid01.Columns(8).RightJustified = True
'    oGrid01.Columns(9).RightJustified = True
'    oGrid01.Columns(10).RightJustified = True
'    oGrid01.Columns(11).RightJustified = True
'    oGrid01.Columns(12).RightJustified = True
'    oGrid01.Columns(13).RightJustified = True
'    oGrid01.Columns(14).RightJustified = True
'    oGrid01.Columns(15).RightJustified = True
'    oGrid01.Columns(16).RightJustified = True
'    oGrid01.Columns(17).RightJustified = True
'    oGrid01.Columns(18).RightJustified = True
'    oGrid01.Columns(19).RightJustified = True
'    oGrid01.Columns(20).RightJustified = True
'    oGrid01.Columns(21).RightJustified = True
'    oGrid01.Columns(22).RightJustified = True
'    oGrid01.Columns(23).RightJustified = True
'    oGrid01.Columns(24).RightJustified = True
'    oGrid01.Columns(25).RightJustified = True
'    oGrid01.Columns(26).RightJustified = True
'    oGrid01.Columns(27).RightJustified = True
'    oGrid01.Columns(28).RightJustified = True
'    oGrid01.Columns(29).RightJustified = True
'    oGrid01.Columns(30).RightJustified = True
'    oGrid01.Columns(31).RightJustified = True
'    oGrid01.Columns(32).RightJustified = True
'    oGrid01.Columns(33).RightJustified = True
'    oGrid01.Columns(34).RightJustified = True
    
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid01.Rows.Count = 1 Then
        ErrNum = 1
        GoTo PS_PP909_MTX03_Error:
    End If
        
    oGrid01.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP909_MTX03_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP909_MTX03_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    If ErrNum = 1 Then
    
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
    
    Else

        Sbo_Application.SetStatusBarMessage "PS_PP909_MTX03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
    End If
    
    Set RecordSet01 = Nothing
    
    Call oForm01.Freeze(False)
    
End Sub

Private Sub PS_PP909_FormResize()
On Error GoTo PS_PP909_FormResize_Error
    
    '그룹박스 크기 동적 할당
    oForm01.Items("GrpBox01").Height = oForm01.Items("Grid01").Height + 90
    oForm01.Items("GrpBox01").Width = oForm01.Items("Grid01").Width + 30
    
    oForm01.Items("Mat01").Width = oForm01.Width / 2 - 90
    oForm01.Items("Mat01").Height = oForm01.Height - oForm01.Items("Mat01").Top - (oForm01.Height - oForm01.Items("2").Top) - 20
    
    oForm01.Items("Mat02").Left = oForm01.Width / 2 - 40
    oForm01.Items("Mat02").Width = oForm01.Width / 2 + 5
    oForm01.Items("Mat02").Height = oForm01.Items("Mat01").Height
    
    oForm01.Items("Static06").Left = oForm01.Width / 2 - 40
    
    oMat01.AutoResizeColumns
    oMat02.AutoResizeColumns
    
    Exit Sub
PS_PP909_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP909_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP909_Print_Report01()
On Error GoTo Print_Query_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String

    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim Sub_sQry        As String
    
'    Dim BPLId           As String
'    Dim CardCode       As String
'    Dim Pumtxt        As String


    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    '// 조회조건문
    
'    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
'    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
'    Pumtxt = Trim(oForm01.Items("Pumtxt").Specific.VALUE)
'
'    If Pumtxt = "" Then Pumtxt = "%"
'    If CardCode = "" Then CardCode = "%"
'
    
    Dim BPLID As String '사업장
    Dim ItemClass As String '품목구분
    Dim TradeType As String '거래형태
    Dim FrDt As String '납기일시작
    Dim ToDt As String '납기일종료
    Dim CardCode As String '거래처
    Dim ItemCode As String '품목코드(작번)
    Dim DocStatus As String '문서상태
    Dim Chk01 As String '미출고
    Dim Chk02 As String '미납품
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    ItemClass = IIf(oForm01.Items("ItemClass").Specific.Selected.VALUE = "%", "", oForm01.Items("ItemClass").Specific.Selected.VALUE)
    TradeType = IIf(oForm01.Items("TradeType").Specific.Selected.VALUE = "%", "", oForm01.Items("TradeType").Specific.Selected.VALUE)
    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    DocStatus = IIf(oForm01.Items("DocStatus").Specific.Selected.VALUE = "%", "", oForm01.Items("DocStatus").Specific.Selected.VALUE)
    Chk01 = IIf(oForm01.Items("Chk01").Specific.Checked = True, "1", "0")
    Chk02 = IIf(oForm01.Items("Chk02").Specific.Checked = True, "1", "0")
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP909] 레포트"
    ReportName = "PS_PP909.rpt"
 
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    sQry = "EXEC PS_PP909_01 '" & BPLID & "','" & ItemClass & "','" & TradeType & "','" & FrDt & "','" & ToDt & "','" & CardCode & "','" & ItemCode & "','" & DocStatus & "','" & Chk01 & "','" & Chk02 & "'"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub PS_PP909_Print_Report02()
On Error GoTo Print_Query_Error

    Dim DocNum As String
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim BPLID As String '사업장
    Dim CntcCode As String
    Dim FrDt As String
    Dim ToDt As String
    
    BPLID = oForm01.Items("BPLID02").Specific.VALUE
    CntcCode = oForm01.Items("CntcCode02").Specific.VALUE
    FrDt = oForm01.Items("FrDt02").Specific.VALUE
    ToDt = oForm01.Items("ToDt02").Specific.VALUE
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP909] 레포트"
    
    ReportName = "PS_PP909_02.rpt"
    
    sQry = "         EXEC PS_PP909_53 '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & CntcCode & "','"
    sQry = sQry & FrDt & "','"
    sQry = sQry & ToDt & "'"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
'    Call oRecordSet.DoQuery(sQry)
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_PP909_Print_Report02_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
End Sub




