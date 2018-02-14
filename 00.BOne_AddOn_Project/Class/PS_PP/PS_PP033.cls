VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP033"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_PP033.cls
'//  Module : 생산관리 > 애드온 > 생산관리
'//  Desc : 표준공수조회
'//  FormType : PS_PP033
'//  Create Date(Start) : 2012.04.12
'//  Create Date(End) :
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form

Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix

Private oDS_PS_PP033H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP033L As SAPbouiCOM.DBDataSource '라인(품목)
Private oDS_PS_PP033M As SAPbouiCOM.DBDataSource '라인(공수)

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'Private gMotherForm As SAPbouiCOM.Form '마더 폼(작업지시등록)을 저장할 Form객체
'Private gMotherDataSource As SAPbouiCOM.DBDataSource '마더 폼의 데이터소스(공정테이블)를 저장할 객체
Private gPS_PP030 As PS_PP030 '작업지시등록 클래스

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
Public Sub LoadForm(Optional ByVal prmPS_PP030 As PS_PP030, Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP033.srf")
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
    
    oFormUniqueID01 = "PS_PP033_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
    
'    Set gMotherForm = pMotherForm '마더폼 할당
'    Set gMotherDataSource = pMotherDataSource '마더폼의 데이터소스(공정테이블) 할당
    Set gPS_PP030 = prmPS_PP030
        
    oForm01.Freeze True
    Call PS_PP033_CreateItems
    Call PS_PP033_ComboBox_Setting
    Call PS_PP033_Initial_Setting
    Call PS_PP033_CF_ChooseFromList
    Call PS_PP033_EnableMenus
    Call PS_PP033_SetDocument(oFromDocEntry01)
    Call PS_PP033_FormResize
    
    oForm01.Update
    oForm01.Freeze False
    
'    oForm01.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    
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
                Call PS_PP033_AddMatrixRow(oMat01.VisualRowCount)
            
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
                Call oDS_PS_PP033L.RemoveRecord(oDS_PS_PP033L.Size - 1)
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
    ElseIf pval.ItemUID = "Mat02" Then
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
    
        If pval.ItemUID = "btnSearch" Then
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_PP033_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    
                    '공수 정보 초기화
                    oMat02.Clear
                    oDS_PS_PP033M.Clear
                    
                    Call PS_PP033_MTX01 '매트릭스에 데이터 로드
                    
                End If
            End If

        ElseIf pval.ItemUID = "btnConfirm" Then '확인버튼 클릭시
        
            If oForm01.Mode = fm_ADD_MODE Then

                '표준공수를 작업지시에 입력
                Call PS_PP033_PutCheckedData
            
            End If
        
        End If
    ElseIf pval.BeforeAction = False Then
    
'        '폴더를 사용할 때는 필수 소스_S
'        'Folder01이 선택되었을 때
'        If pval.ItemUID = "Folder01" Then
'
'            oForm01.PaneLevel = 1
'
'        End If
'
'        'Folder02가 선택되었을 때
'        If pval.ItemUID = "Folder02" Then
'
'            oForm01.PaneLevel = 2
'
'        End If
'
'        'Folder03이 선택되었을 때
'        If pval.ItemUID = "Folder03" Then
'
'            oForm01.PaneLevel = 3
'
'        End If
'        '폴더를 사용할 때는 필수 소스_E
    
        If pval.ItemUID = "PS_PP033" Then
        
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
        
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '거래처코드 포맷서치 활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '품목코드(작번) 포맷서치 활성
        
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
    
        If pval.ItemUID = "Mat01" Then '품목

            If pval.Row = 0 Then

                '정렬
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource

            Else

                Call PS_PP033_MTX02(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE)

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
    
'        If pval.ItemUID = "Mat02" Then
'
'            Dim PP030 As Object
'            Set PP030 = New PS_PP030
'
'            Call PP030.LoadForm(oMat02.Columns("WOEntry").Cells(pval.Row).Specific.VALUE)
'
'            Set PP030 = Nothing
'
'        ElseIf pval.ItemUID = "Mat04" Then
'
'            Dim PP040 As Object
'            Set PP040 = New PS_PP040
'
'            Call PP040.LoadForm(oMat04.Columns("WREntry").Cells(pval.Row).Specific.VALUE)
'
'            Set PP040 = Nothing
'
'        End If
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
'    Set PP030 = Nothing
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
        
'            If (pval.ItemUID = "CardCode") Then
'                oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
'                oRecordSet01.DoQuery oQuery01
'                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'            ElseIf (pval.ItemUID = "ItemCode") Then
'                oQuery01 = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
'                oRecordSet01.DoQuery oQuery01
'                oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'            ElseIf (pval.ItemUID = "CntcCode") Then
'                oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
'                oRecordSet01.DoQuery oQuery01
'                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'            End If
            
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
        Call PS_PP033_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP033_FormResize
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

Private Function PS_PP033_CreateItems() As Boolean
On Error GoTo PS_PP033_CreateItems_Error
    Call oForm01.Freeze(True)
'    Dim oQuery01 As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

'    Set oDS_PS_PP033H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_PP033L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_PP033M = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    
    '매트릭스 초기화
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
    
    '품명_S
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    '품명_E
    
    '규격_S
    Call oForm01.DataSources.UserDataSources.Add("Spec", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Spec").Specific.DataBind.SetBound(True, "", "Spec")
    '규격_E
    
    '표시유지_S
    Call oForm01.DataSources.UserDataSources.Add("Chk01", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Chk01").Specific.DataBind.SetBound(True, "", "Chk01")
    '표시유지_E
    
    '공정금액 합계_S
    Call oForm01.DataSources.UserDataSources.Add("CpPriceSum", dt_SUM)
    Call oForm01.Items("CpPriceSum").Specific.DataBind.SetBound(True, "", "CpPriceSum")
    '공정금액 합계_E
    
    
    '문서번호 Hidden
    oMat01.Columns("DocEntry").Visible = False
    
    Call oForm01.Freeze(False)
    Exit Function
PS_PP033_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP033_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP033_ComboBox_Setting()
On Error GoTo PS_PP033_ComboBox_Setting_Error
    
    Call oForm01.Freeze(True)
        
'        '품목대분류 콤보박스 세팅_S
'        Call oForm01.Items("ItmBsort").Specific.ValidValues.Add("", "전체")
'        Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN ('105','106') order by Code", "", False, False)
'        Call oForm01.Items("ItmBsort").Specific.Select(0, psk_Index)
'        '품목대분류 콤보박스 세팅_E
'
'        '품목구분 콤보박스 세팅_S
'        Call oForm01.Items("ItemClass").Specific.ValidValues.Add("", "전체")
'        Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemClass").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", False, False)
'        Call oForm01.Items("ItemClass").Specific.Select(0, psk_Index)
'        '품목구분 콤보박스 세팅_E
'
'        '거래처구분_S
'        Call oForm01.Items("CardType").Specific.ValidValues.Add("", "전체")
'        Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", False, False)
'        Call oForm01.Items("CardType").Specific.Select(0, psk_Index)
'        '거래처구분_E

    Call oForm01.Freeze(False)
    Exit Sub
PS_PP033_ComboBox_Setting_Error:
'    Set lRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP033_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP033_Initial_Setting()
On Error GoTo PS_PP033_Initial_Setting_Error

    '사업장 사용자의 소속 사업장 선택
'    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '기준일자
'    oForm01.Items("StdDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    

    
    '날짜 설정
'    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYYMMDD")
'    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    '날짜 빈칸으로 설정
'    oForm01.Items("ToDt").Specific.VALUE = ""
'    oForm01.Items("FrDt").Specific.VALUE = ""
 
    Exit Sub
PS_PP033_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP033_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_PP033_CF_ChooseFromList()
On Error GoTo PS_PP033_CF_ChooseFromList_Error
    Exit Sub
PS_PP033_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP033_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP033_FormItemEnabled()
On Error GoTo PS_PP033_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP033_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP033_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP033_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP033_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP033L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP033L.Offset = oRow
'    oDS_PS_PP033L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP033_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP033_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP033_FormClear()
On Error GoTo PS_PP033_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP033'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP033_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP033_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP033_EnableMenus()
On Error GoTo PS_PP033_EnableMenus_Error
    Exit Sub
PS_PP033_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP033_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP033_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP033_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP033_FormItemEnabled

    Else

    End If
    Exit Sub
PS_PP033_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP033_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_PP033_DataValidCheck() As Boolean
On Error GoTo PS_PP033_DataValidCheck_Error
    PS_PP033_DataValidCheck = False
    Dim i As Long

'    If oForm01.Items("DocDateTo").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "견적일자 To은 필수입니다.", bmt_Short, True
'        oForm01.Items("DocDateTo").Click ct_Regular
'        PS_PP033_DataValidCheck = False
'        Exit Function
'    End If

    PS_PP033_DataValidCheck = True
    Exit Function
PS_PP033_DataValidCheck_Error:
    PS_PP033_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP033_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP033_MTX01()
On Error GoTo PS_PP033_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ItemName As String '품목구분
    Dim SPEC As String '규격
    
    ItemName = oForm01.Items("ItemName").Specific.VALUE
    SPEC = oForm01.Items("Spec").Specific.VALUE
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP033_01 '" & ItemName & "','" & SPEC & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_PP033_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP033L.InsertRecord(loopCount)
        End If
        oDS_PS_PP033L.Offset = loopCount
        
        Call oDS_PS_PP033L.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP033L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("ItemName").VALUE) '품목명
        Call oDS_PS_PP033L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("Spec").VALUE) '규격
        Call oDS_PS_PP033L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("DocEntry").VALUE) '문서번호
        
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
    
PS_PP033_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP033_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP033_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP033_MTX02(ByVal prmDocEntry As String)
On Error GoTo PS_PP033_MTX02_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim CpPriceSum As Currency '공정금액 합계
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP033_02 '" & prmDocEntry & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat02.Clear
        GoTo PS_PP033_MTX02_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP033M.InsertRecord(loopCount)
        End If
        oDS_PS_PP033M.Offset = loopCount
        
        Call oDS_PS_PP033M.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP033M.setValue("U_ColReg01", loopCount, RecordSet01.Fields("Select").VALUE) '선택
        Call oDS_PS_PP033M.setValue("U_ColReg02", loopCount, RecordSet01.Fields("CpBCode").VALUE) '공정대분류
        Call oDS_PS_PP033M.setValue("U_ColReg03", loopCount, RecordSet01.Fields("CpBName").VALUE) '대분류명
        Call oDS_PS_PP033M.setValue("U_ColReg04", loopCount, RecordSet01.Fields("CpCode").VALUE) '공정중분류
        Call oDS_PS_PP033M.setValue("U_ColReg05", loopCount, RecordSet01.Fields("CpName").VALUE) '중분류명
        Call oDS_PS_PP033M.setValue("U_ColReg06", loopCount, RecordSet01.Fields("StdHour").VALUE) '표준공수
        Call oDS_PS_PP033M.setValue("U_ColSum01", loopCount, RecordSet01.Fields("Price").VALUE) '공정단가
        Call oDS_PS_PP033M.setValue("U_ColSum02", loopCount, RecordSet01.Fields("CpPrice").VALUE) '공정금액
        
        CpPriceSum = CpPriceSum + RecordSet01.Fields("CpPrice").VALUE
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    
    oForm01.Items("CpPriceSum").Specific.VALUE = CpPriceSum
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP033_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP033_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP033_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP033_FormResize()
On Error GoTo PS_PP033_FormResize_Error

    oForm01.Items("Mat01").Width = oForm01.Width / 2 - 80
    oForm01.Items("Mat01").Height = oForm01.Height - oForm01.Items("Mat01").Top - (oForm01.Height - oForm01.Items("2").Top) - 10

    oForm01.Items("Mat02").Left = oForm01.Width / 2 - 60
    oForm01.Items("Mat02").Width = oForm01.Width / 2 + 35
    oForm01.Items("Mat02").Height = oForm01.Items("Mat01").Height
    
    oForm01.Items("Static04").Left = oForm01.Items("Mat02").Left
    
'
'    oForm01.Items("Mat03").Left = oForm01.Width / 2 + 10
'    oForm01.Items("Mat03").Width = oForm01.Width / 2 - 40
'    oForm01.Items("Mat03").Height = oForm01.Items("Mat01").Height - 20
'
'    oForm01.Items("Mat04").Left = oForm01.Width / 2 + 10
'    oForm01.Items("Mat04").Width = oForm01.Width / 2 - 40
'    oForm01.Items("Mat04").Height = oForm01.Items("Mat01").Height - 20
'
'    oForm01.Items("Rec01").Left = oForm01.Width / 2
'    oForm01.Items("Rec01").Height = oForm01.Items("Mat02").Height + 20
'    oForm01.Items("Rec01").Width = oForm01.Items("Mat02").Width + 20
'
'    oForm01.Items("Folder01").Width = 100
'    oForm01.Items("Folder01").Left = oForm01.Items("Rec01").Left
'
'    oForm01.Items("Folder02").Width = 100
'    oForm01.Items("Folder02").Left = oForm01.Items("Folder01").Left + oForm01.Items("Folder02").Width
'
'    oForm01.Items("Folder03").Width = 100
'    oForm01.Items("Folder03").Left = oForm01.Items("Folder02").Left + oForm01.Items("Folder03").Width
    
    
    oMat01.AutoResizeColumns
    oMat02.AutoResizeColumns
    
    Exit Sub
PS_PP033_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP033_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP033_Print_Report01()
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
    
    Dim BPLId As String '사업장
    Dim ItemClass As String '품목구분
    Dim TradeType As String '거래형태
    Dim FrDt As String '납기일시작
    Dim ToDt As String '납기일종료
    Dim CardCode As String '거래처
    Dim ItemCode As String '품목코드(작번)
    Dim DocStatus As String '문서상태
    Dim Chk01 As String '미출고
    Dim Chk02 As String '미납품
    
    BPLId = oForm01.Items("BPLId").Specific.Selected.VALUE
    ItemClass = IIf(oForm01.Items("ItemClass").Specific.Selected.VALUE = "%", "", oForm01.Items("ItemClass").Specific.Selected.VALUE)
    TradeType = IIf(oForm01.Items("TradeType").Specific.Selected.VALUE = "%", "", oForm01.Items("TradeType").Specific.Selected.VALUE)
    FrDt = oForm01.Items("FrDt").Specific.VALUE
    ToDt = oForm01.Items("ToDt").Specific.VALUE
    CardCode = oForm01.Items("CardCode").Specific.VALUE
    ItemCode = oForm01.Items("ItemCode").Specific.VALUE
    DocStatus = IIf(oForm01.Items("DocStatus").Specific.Selected.VALUE = "%", "", oForm01.Items("DocStatus").Specific.Selected.VALUE)
    Chk01 = IIf(oForm01.Items("Chk01").Specific.Checked = True, "1", "0")
    Chk02 = IIf(oForm01.Items("Chk02").Specific.Checked = True, "1", "0")
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP033] 레포트"
    ReportName = "PS_PP033.rpt"
 
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
    sQry = "EXEC PS_PP033_01 '" & BPLId & "','" & ItemClass & "','" & TradeType & "','" & FrDt & "','" & ToDt & "','" & CardCode & "','" & ItemCode & "','" & DocStatus & "','" & Chk01 & "','" & Chk02 & "'"

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

Private Sub PS_PP033_PutCheckedData()
'******************************************************************************
'Function ID : PS_PP033_PutCheckedData()
'해당모듈 : PS_PP033
'기능 : 표준공수를 작업지시등록 화면에 추가
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP033_PutCheckedData_Error

    Dim loopCount As Integer
    Dim rowIndex As Integer
    
    rowIndex = 0
    
    oMat02.FlushToDataSource

    Call gPS_PP030.oForm01.Freeze(True)
    Call gPS_PP030.oMat03.Clear
    Call gPS_PP030.oDS_PS_PP030M.Clear
    Call gPS_PP030.oMat03.FlushToDataSource
    
    For loopCount = 0 To oMat02.VisualRowCount - 1

        If Trim(oDS_PS_PP033M.GetValue("U_ColReg01", loopCount)) = "Y" Then '선택된 행만 공수 등록
        
            If rowIndex <> 0 Then
               Call gPS_PP030.oDS_PS_PP030M.InsertRecord(rowIndex)
            End If

            gPS_PP030.oDS_PS_PP030M.Offset = rowIndex

            Call gPS_PP030.oDS_PS_PP030M.setValue("U_LineNum", rowIndex, rowIndex + 1) '라인번호
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_Sequence", rowIndex, rowIndex + 1) '순번
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_CpBCode", rowIndex, oMat02.Columns("CpBCode").Cells(loopCount + 1).Specific.VALUE) '공정대분류
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_CpBName", rowIndex, oMat02.Columns("CpBName").Cells(loopCount + 1).Specific.VALUE) '대분류명
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_CpCode", rowIndex, oMat02.Columns("CpCode").Cells(loopCount + 1).Specific.VALUE) '공정중분류
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_CpName", rowIndex, oMat02.Columns("CpName").Cells(loopCount + 1).Specific.VALUE) '중분류명
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_StdHour", rowIndex, oMat02.Columns("StdHour").Cells(loopCount + 1).Specific.VALUE) '표준공수
            
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_CpPrice", rowIndex, oMat02.Columns("CpPrice").Cells(loopCount + 1).Specific.VALUE) '공정금액
            
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_WorkGbn", rowIndex, "10") '작업구분
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_ReWorkYN", rowIndex, "N") '재작업여부
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_ResultYN", rowIndex, "N") '실적여부
            Call gPS_PP030.oDS_PS_PP030M.setValue("U_ReportYN", rowIndex, "Y") '일보여부

            rowIndex = rowIndex + 1

        End If

    Next

    Call gPS_PP030.oMat03.LoadFromDataSource
    Call gPS_PP030.PS_PP030_AddMatrixRow02(gPS_PP030.oMat03.VisualRowCount, False)
    Call gPS_PP030.oMat03.AutoResizeColumns
    
    gPS_PP030.oForm01.Items("Total").Specific.VALUE = oForm01.Items("CpPriceSum").Specific.VALUE '공정금액 합계
    
    Call gPS_PP030.oForm01.Freeze(False)
    
    If oForm01.Items("Chk01").Specific.Checked = False Then
        Call oForm01.Close
    End If
    
Exit Sub
PS_PP033_PutCheckedData_Error:
    Call Sbo_Application.SetStatusBarMessage("PS_PP033_PutCheckedData_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub




