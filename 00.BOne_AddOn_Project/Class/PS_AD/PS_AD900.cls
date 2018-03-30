VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_AD900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_AD900.cls
'//  Module : 운영관리 > 프로그램 수정 및 추가 요청
'//  Desc : 프로그램 수정 및 추가 요청 현황
'//  FormType : PS_AD900
'//  Create Date(Start) : 2011.12.30
'//  Create Date(End) : 2011.12.30
'//  Modified Date :
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix
Private oDS_PS_AD900H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_AD900L As SAPbouiCOM.DBDataSource '등록라인
Private oDS_PS_USERDS02 As SAPbouiCOM.DBDataSource

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_AD900.srf")
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
    
    oFormUniqueID01 = "PS_AD900_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_AD900_CreateItems
    Call PS_AD900_ComboBox_Setting
    Call PS_AD900_Initial_Setting
    Call PS_AD900_CF_ChooseFromList
    Call PS_AD900_EnableMenus
    Call PS_AD900_SetDocument(oFromDocEntry01)
    Call PS_AD900_FormResize
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    oForm01.Items("Folder02").Enabled = False '건수 Tab Enabled = False
    
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
                Call PS_AD900_AddMatrixRow(oMat01.VisualRowCount)
                Call PS_AD900_AddMatrixRow2(oMat02.VisualRowCount)
            
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
                Call oDS_PS_AD900L.RemoveRecord(oDS_PS_AD900L.Size - 1)
                Call oMat01.LoadFromDataSource
                
                Call oDS_PS_USERDS02.RemoveRecord(oDS_PS_USERDS02.Size - 1)
                Call oMat02.LoadFromDataSource
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
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_AD900_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    Call PS_AD900_MTX01 '매트릭스에 데이터 로드
                    oForm01.Items("Folder02").Enabled = True
                End If
            End If
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_AD900_Print_Report01
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
        ElseIf pval.ItemUID = "BtnPrint" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
        
                Call PS_AD900_Print_Report01
            
            End If
        ElseIf pval.ItemUID = "Folder02" Then
        
            Call PS_AD900_MTX02 '매트릭스에 데이터 로드
                
        End If
    ElseIf pval.BeforeAction = False Then
        
        '폴더를 사용할 때는 필수 소스_S
        'Folder01이 선택되었을 때
        If pval.ItemUID = "Folder01" Then
            
            oForm01.PaneLevel = 1
            
        End If
        
        'Folder02가 선택되었을 때
        If pval.ItemUID = "Folder02" Then
            
            oForm01.PaneLevel = 2
            
        End If
        '폴더를 사용할 때는 필수 소스_E
    
        If pval.ItemUID = "PS_AD900" Then
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
        
        If pval.ItemUID = "ReqEmpCd" Then
        
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ReqEmpCd", "") '요청자 포맷서치 설정
        
        End If
        
        
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
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row = 0 Then
            
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource
            
            End If
        ElseIf pval.ItemUID = "Mat02" Then
        
            If pval.Row = 0 Then
            
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
    
        If pval.ItemUID = "Mat01" And pval.ColUID = "AD100_Doc" Then
    
            Dim AD100 As Object
            Set AD100 = New PS_AD100
    
            Call AD100.LoadForm(oMat01.Columns("AD100_Doc").Cells(pval.Row).Specific.VALUE)
    
            Set AD100 = Nothing
            
        ElseIf pval.ItemUID = "Mat01" And pval.ColUID = "AD101_Doc" Then
            
            Dim AD101 As Object
            Set AD101 = New PS_AD101
    
            Call AD101.LoadForm(oMat01.Columns("AD101_Doc").Cells(pval.Row).Specific.VALUE)
    
            Set AD101 = Nothing
            
        ElseIf pval.ItemUID = "Mat01" And pval.ColUID = "AD102_Doc" Then
            
            Dim AD102 As Object
            Set AD102 = New PS_AD102
    
            Call AD102.LoadForm(oMat01.Columns("AD102_Doc").Cells(pval.Row).Specific.VALUE)
    
            Set AD102 = Nothing
            
        End If
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
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
            
            If pval.ItemUID = "ReqEmpCd" Then '요청자
            
                oQuery01 = "SELECT U_FULLNAME FROM OHEM WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                Call oRecordSet01.DoQuery(oQuery01)
                oForm01.Items("ReqEmpNm").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
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
        Call PS_AD900_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_AD900_FormResize
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
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_AD900_CreateItems() As Boolean
On Error GoTo PS_AD900_CreateItems_Error
    Call oForm01.Freeze(True)
'    Dim oQuery01 As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Set oDS_PS_AD900H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_AD900L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_USERDS02 = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    
    '매트릭스 초기화
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    '사업장_E
    
    '요청자_S
    Call oForm01.DataSources.UserDataSources.Add("ReqEmpCd", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ReqEmpCd").Specific.DataBind.SetBound(True, "", "ReqEmpCd")
    '요청자_E
    
    '요청자성명_S
    Call oForm01.DataSources.UserDataSources.Add("ReqEmpNm", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ReqEmpNm").Specific.DataBind.SetBound(True, "", "ReqEmpNm")
    '요청자성명_E
    
    '요청일 시작_S
    Call oForm01.DataSources.UserDataSources.Add("ReqFrDt", dt_DATE)
    Call oForm01.Items("ReqFrDt").Specific.DataBind.SetBound(True, "", "ReqFrDt")
    '요청일 시작_E
    
    '요청일 종료_S
    Call oForm01.DataSources.UserDataSources.Add("ReqToDt", dt_DATE)
    Call oForm01.Items("ReqToDt").Specific.DataBind.SetBound(True, "", "ReqToDt")
    '요청일 종료_E

    '완료요구일 시작_S
    Call oForm01.DataSources.UserDataSources.Add("FinReqFrDt", dt_DATE)
    Call oForm01.Items("FinReqFrDt").Specific.DataBind.SetBound(True, "", "FinReqFrDt")
    '완료요구일 시작_E

    '완료요구일 종료_S
    Call oForm01.DataSources.UserDataSources.Add("FinReqToDt", dt_DATE)
    Call oForm01.Items("FinReqToDt").Specific.DataBind.SetBound(True, "", "FinReqToDt")
    '완료요구일 종료_E
    
    '프로그램구분_S
    Call oForm01.DataSources.UserDataSources.Add("PGType", dt_SHORT_TEXT, 100)
    Call oForm01.Items("PGType").Specific.DataBind.SetBound(True, "", "PGType")
    '프로그램구분_E
    
    '요청구분
    Call oForm01.DataSources.UserDataSources.Add("ReqType", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ReqType").Specific.DataBind.SetBound(True, "", "ReqType")
    '요청구분
    
    '완료일 시작_S
    Call oForm01.DataSources.UserDataSources.Add("FinFrDt", dt_DATE)
    Call oForm01.Items("FinFrDt").Specific.DataBind.SetBound(True, "", "FinFrDt")
    '완료일 시작_E

    '완료일 종료_S
    Call oForm01.DataSources.UserDataSources.Add("FinToDt", dt_DATE)
    Call oForm01.Items("FinToDt").Specific.DataBind.SetBound(True, "", "FinToDt")
    '완료일 종료_E
    
    '접수상태
    Call oForm01.DataSources.UserDataSources.Add("AcStatus", dt_SHORT_TEXT, 100)
    Call oForm01.Items("AcStatus").Specific.DataBind.SetBound(True, "", "AcStatus")
    '접수상태
    
    Call oForm01.Freeze(False)
    Exit Function
PS_AD900_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_AD900_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_AD900_ComboBox_Setting()
On Error GoTo PS_AD900_ComboBox_Setting_Error
    
    Dim sQry As String
    
    Call oForm01.Freeze(True)
     
        '사업장 콤보박스 세팅_S
        Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
        '사업장 콤보박스 세팅_E
        
        '프로그램구분 세팅_S
        Call oForm01.Items("PGType").Specific.ValidValues.Add("%", "전체")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("PGType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'A001'", "", False, False)
        Call oForm01.Items("PGType").Specific.Select(0, psk_Index)
        '프로그램구분 세팅_E
        
        '요청구분 세팅_S
        Call oForm01.Items("ReqType").Specific.ValidValues.Add("%", "전체")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("ReqType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'A002'", "", False, False)
        Call oForm01.Items("ReqType").Specific.Select(0, psk_Index)
        '요청구분 세팅_E
        
        '접수상태_S
        Call oForm01.Items("AcStatus").Specific.ValidValues.Add("%", "전체")
        Call oForm01.Items("AcStatus").Specific.ValidValues.Add("A", "미접수")
        Call oForm01.Items("AcStatus").Specific.ValidValues.Add("B", "접수")
        Call oForm01.Items("AcStatus").Specific.ValidValues.Add("C", "반려")
        Call oForm01.Items("AcStatus").Specific.ValidValues.Add("D", "완료")
        Call oForm01.Items("AcStatus").Specific.Select(0, psk_Index) '기본으로 전체 세팅
        '접수상태_E
        
'        '품목구분 콤보박스 세팅_S
'        Call oForm01.Items("ItemClass").Specific.ValidValues.Add("%", "전체")
'        Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemClass").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", False, False)
'        Call oForm01.Items("ItemClass").Specific.Select(0, psk_Index)
'        '품목구분 콤보박스 세팅_E
'
'        '거래형태 콤보박스 세팅_S
'        Call oForm01.Items("TradeType").Specific.ValidValues.Add("%", "전체")
'        Call MDC_SetMod.Set_ComboList(oForm01.Items("TradeType").Specific, "SELECT  U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'T001' ORDER BY U_Minor", "", False, False)
'        Call oForm01.Items("TradeType").Specific.Select(0, psk_Index)
'        '거래형태 콤보박스 세팅_E
'
'        '수주문서상태 세팅_S
'        Call oForm01.Items("DocStatus").Specific.ValidValues.Add("%", "전체")
'        Call oForm01.Items("DocStatus").Specific.ValidValues.Add("O", "미결")
'        Call oForm01.Items("DocStatus").Specific.ValidValues.Add("C", "종료")
'        Call oForm01.Items("DocStatus").Specific.Select(0, psk_Index) '기본으로 전체 세팅
'        '수주문서상태 세팅_E
'
        '매트릭스 컬럼의 콤보박스 설정_S
'        Dim oColumn As SAPbouiCOM.Column
'        Dim lRecordSet As SAPbobsCOM.Recordset '레코드셋
'        Dim lQuery As String '쿼리스트링
'
'        Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
'
'        '거래형태 콤보박스
'        Set oColumn = oMat01.Columns("TradeType")
'        lQuery = "SELECT  U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'T001' ORDER BY U_Minor"
'        Call lRecordSet.DoQuery(lQuery)
'
'        Do Until lRecordSet.EOF
'            oColumn.ValidValues.Add lRecordSet.Fields(0).VALUE, lRecordSet.Fields(1).VALUE
'            lRecordSet.MoveNext
'        Loop
'        Set lRecordSet = Nothing
'        '거래형태 콤보박스
'
'        '문서상태
'        Set oColumn = oMat01.Columns("DocStatus")
'        Call oColumn.ValidValues.Add("O", "미결")
'        Call oColumn.ValidValues.Add("C", "종료")
'        '문서상태

        '난이도
        sQry = "           SELECT      U_Minor, "
        sQry = sQry & "                U_CdName "
        sQry = sQry & " FROM       [@PS_SY001L] "
        sQry = sQry & " WHERE      Code = 'A004'"
        sQry = sQry & "                AND ISNULL(U_UseYN, 'Y') = 'Y'"
        sQry = sQry & " ORDER BY  U_Seq"
        
        Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("LvlDiff"), sQry)

        '매트릭스 컬럼의 콤보박스 설정_E
        
        
    Call oForm01.Freeze(False)
    Exit Sub
PS_AD900_ComboBox_Setting_Error:
'    Set lRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_AD900_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_AD900_Initial_Setting()
On Error GoTo PS_AD900_Initial_Setting_Error
    '사업장 사용자의 소속 사업장 선택
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '날짜 설정
'    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYYMMDD")
'    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    '날짜 빈칸으로 설정
    
    Exit Sub
PS_AD900_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_AD900_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_AD900_CF_ChooseFromList()
On Error GoTo PS_AD900_CF_ChooseFromList_Error
    Exit Sub
PS_AD900_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_AD900_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_AD900_FormItemEnabled()
On Error GoTo PS_AD900_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_AD900_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_AD900_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_AD900_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_AD900_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_AD900L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_AD900L.Offset = oRow
'    oDS_PS_AD900L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_AD900_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_AD900_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_AD900_AddMatrixRow2(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_AD900_AddMatrixRow2_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_USERDS02.InsertRecord (oRow)
    End If
    oMat02.AddRow
    oDS_PS_USERDS02.Offset = oRow
'    oDS_PS_USERDS02.setValue "U_LineNum", oRow, oRow + 1
    oMat02.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_AD900_AddMatrixRow2_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_AD900_AddMatrixRow2_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_AD900_FormClear()
On Error GoTo PS_AD900_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_AD900'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_AD900_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_AD900_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_AD900_EnableMenus()
On Error GoTo PS_AD900_EnableMenus_Error
    Exit Sub
PS_AD900_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_AD900_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_AD900_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_AD900_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_AD900_FormItemEnabled

    Else

    End If
    Exit Sub
PS_AD900_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_AD900_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_AD900_DataValidCheck() As Boolean
On Error GoTo PS_AD900_DataValidCheck_Error
    PS_AD900_DataValidCheck = False
    Dim i As Long

'    If oForm01.Items("DocDateTo").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "견적일자 To은 필수입니다.", bmt_Short, True
'        oForm01.Items("DocDateTo").Click ct_Regular
'        PS_AD900_DataValidCheck = False
'        Exit Function
'    End If

    PS_AD900_DataValidCheck = True
    Exit Function
PS_AD900_DataValidCheck_Error:
    PS_AD900_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_AD900_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_AD900_MTX01()
On Error GoTo PS_AD900_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String '사업장
    Dim ReqEmpCd As String '요청자
    Dim ReqFrDt As String '요청일시작
    Dim ReqToDt As String '요청일종료
    Dim FinReqFrDt As String '완료요구일시작
    Dim FinReqToDt As String '완료요구일종료
    Dim PGType As String '프로그램구분
    Dim ReqType As String '요청구분
    
    Dim FinFrDt As String '완료일시작
    Dim FinToDt As String '완료일종료
    Dim AcStatus As String '접수상태
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    ReqEmpCd = oForm01.Items("ReqEmpCd").Specific.VALUE
    ReqFrDt = oForm01.Items("ReqFrDt").Specific.VALUE
    ReqToDt = oForm01.Items("ReqToDt").Specific.VALUE
    FinReqFrDt = oForm01.Items("FinReqFrDt").Specific.VALUE
    FinReqToDt = oForm01.Items("FinReqToDt").Specific.VALUE
    PGType = IIf(oForm01.Items("PGType").Specific.Selected.VALUE = "%", "", oForm01.Items("PGType").Specific.Selected.VALUE)
    ReqType = IIf(oForm01.Items("ReqType").Specific.Selected.VALUE = "%", "", oForm01.Items("ReqType").Specific.Selected.VALUE)
    
    FinFrDt = oForm01.Items("FinFrDt").Specific.VALUE
    FinToDt = oForm01.Items("FinToDt").Specific.VALUE
    AcStatus = IIf(oForm01.Items("AcStatus").Specific.Selected.VALUE = "%", "", oForm01.Items("AcStatus").Specific.Selected.VALUE)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_AD900_01 '" & BPLID & "','" & ReqEmpCd & "','" & ReqFrDt & "','" & ReqToDt & "','" & FinReqFrDt & "','" & FinReqToDt & "','" & PGType & "','" & ReqType & "','" & FinFrDt & "','" & FinToDt & "','" & AcStatus & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_AD900_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_AD900L.InsertRecord(loopCount)
        End If
        oDS_PS_AD900L.Offset = loopCount
        
        Call oDS_PS_AD900L.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_AD900L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("ReqCd").VALUE) '요청자사번
        Call oDS_PS_AD900L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("ReqNm").VALUE) '요청자성명
        Call oDS_PS_AD900L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("PGType").VALUE) '프로그램구분
        Call oDS_PS_AD900L.setValue("U_ColReg13", loopCount, RecordSet01.Fields("MainMenu").VALUE) '주메뉴
        Call oDS_PS_AD900L.setValue("U_ColReg14", loopCount, RecordSet01.Fields("MainMenuNM").VALUE) '주메뉴명
        Call oDS_PS_AD900L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("PGCd").VALUE) '프로그램코드
        Call oDS_PS_AD900L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("PGNm").VALUE) '프로그램명
        Call oDS_PS_AD900L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("ReqTypeNm").VALUE) '요청구분
        Call oDS_PS_AD900L.setValue("U_ColTxt01", loopCount, RecordSet01.Fields("ReqNote").VALUE) '요청사항(길이 문제로 U_ColTxt01로 수정, 2012.04.17 송명규)
        Call oDS_PS_AD900L.setValue("U_ColReg19", loopCount, RecordSet01.Fields("LvlDiff").VALUE) '난이도(2014.03.14 송명규 추가)
        Call oDS_PS_AD900L.setValue("U_ColReg18", loopCount, RecordSet01.Fields("GWNo").VALUE) '기안문서번호(2014.03.13 송명규 추가)
        Call oDS_PS_AD900L.setValue("U_ColReg08", loopCount, RecordSet01.Fields("AcStatus").VALUE) '접수상태
        Call oDS_PS_AD900L.setValue("U_ColReg09", loopCount, RecordSet01.Fields("ReqDocDate").VALUE) '요청일자
        Call oDS_PS_AD900L.setValue("U_ColReg10", loopCount, RecordSet01.Fields("AcDocDate").VALUE) '접수일자
        Call oDS_PS_AD900L.setValue("U_ColReg11", loopCount, RecordSet01.Fields("FinReqDate").VALUE) '완료요구일자
        Call oDS_PS_AD900L.setValue("U_ColReg20", loopCount, RecordSet01.Fields("CltDate").VALUE) '완료예정일자(2014.03.14 송명규 추가)
        Call oDS_PS_AD900L.setValue("U_ColReg12", loopCount, RecordSet01.Fields("FinDocDate").VALUE) '완료일자
        Call oDS_PS_AD900L.setValue("U_ColReg15", loopCount, RecordSet01.Fields("AD100_DocEntry").VALUE) '등록문서번호
        Call oDS_PS_AD900L.setValue("U_ColReg21", loopCount, RecordSet01.Fields("AD100_Line").VALUE) '등록문서라인
        Call oDS_PS_AD900L.setValue("U_ColReg16", loopCount, RecordSet01.Fields("AD101_DocEntry").VALUE) '접수문서번호
        Call oDS_PS_AD900L.setValue("U_ColReg22", loopCount, RecordSet01.Fields("AD101_Line").VALUE) '접수문서라인
        Call oDS_PS_AD900L.setValue("U_ColReg17", loopCount, RecordSet01.Fields("AD102_DocEntry").VALUE) '완료문서번호
        Call oDS_PS_AD900L.setValue("U_ColReg23", loopCount, RecordSet01.Fields("AD102_Line").VALUE) '완료문서라인
        
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
    
PS_AD900_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_AD900_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_AD900_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_AD900_MTX02()
On Error GoTo PS_AD900_MTX02_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String '사업장
    Dim ReqEmpCd As String '요청자
    Dim ReqFrDt As String '요청일시작
    Dim ReqToDt As String '요청일종료
    Dim FinReqFrDt As String '완료요구일시작
    Dim FinReqToDt As String '완료요구일종료
    Dim PGType As String '프로그램구분
    Dim ReqType As String '요청구분
    
    Dim FinFrDt As String '완료일시작
    Dim FinToDt As String '완료일종료
    Dim AcStatus As String '접수상태
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    ReqEmpCd = oForm01.Items("ReqEmpCd").Specific.VALUE
    ReqFrDt = oForm01.Items("ReqFrDt").Specific.VALUE
    ReqToDt = oForm01.Items("ReqToDt").Specific.VALUE
    FinReqFrDt = oForm01.Items("FinReqFrDt").Specific.VALUE
    FinReqToDt = oForm01.Items("FinReqToDt").Specific.VALUE
    PGType = IIf(oForm01.Items("PGType").Specific.Selected.VALUE = "%", "", oForm01.Items("PGType").Specific.Selected.VALUE)
    ReqType = IIf(oForm01.Items("ReqType").Specific.Selected.VALUE = "%", "", oForm01.Items("ReqType").Specific.Selected.VALUE)
    
    FinFrDt = oForm01.Items("FinFrDt").Specific.VALUE
    FinToDt = oForm01.Items("FinToDt").Specific.VALUE
    AcStatus = IIf(oForm01.Items("AcStatus").Specific.Selected.VALUE = "%", "", oForm01.Items("AcStatus").Specific.Selected.VALUE)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_AD900_02 '" & BPLID & "','" & ReqEmpCd & "','" & ReqFrDt & "','" & ReqToDt & "','" & FinReqFrDt & "','" & FinReqToDt & "','" & PGType & "','" & ReqType & "','" & FinFrDt & "','" & FinToDt & "','" & AcStatus & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat02.Clear
        GoTo PS_AD900_MTX02_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_USERDS02.InsertRecord(loopCount)
        End If
        oDS_PS_USERDS02.Offset = loopCount
        
        Call oDS_PS_USERDS02.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_USERDS02.setValue("U_ColReg01", loopCount, RecordSet01.Fields("MainMenu").VALUE) '모듈
        Call oDS_PS_USERDS02.setValue("U_ColReg02", loopCount, RecordSet01.Fields("MainMenuNm").VALUE) '모듈명
        Call oDS_PS_USERDS02.setValue("U_ColReg03", loopCount, RecordSet01.Fields("ClassNm").VALUE) '구분
        Call oDS_PS_USERDS02.setValue("U_ColReg04", loopCount, RecordSet01.Fields("M01").VALUE) '1월
        Call oDS_PS_USERDS02.setValue("U_ColReg05", loopCount, RecordSet01.Fields("M02").VALUE) '2월
        Call oDS_PS_USERDS02.setValue("U_ColReg06", loopCount, RecordSet01.Fields("M03").VALUE) '3월
        Call oDS_PS_USERDS02.setValue("U_ColReg07", loopCount, RecordSet01.Fields("M04").VALUE) '4월
        Call oDS_PS_USERDS02.setValue("U_ColReg08", loopCount, RecordSet01.Fields("M05").VALUE) '5월
        Call oDS_PS_USERDS02.setValue("U_ColReg09", loopCount, RecordSet01.Fields("M06").VALUE) '6월
        Call oDS_PS_USERDS02.setValue("U_ColReg10", loopCount, RecordSet01.Fields("M07").VALUE) '7월
        Call oDS_PS_USERDS02.setValue("U_ColReg11", loopCount, RecordSet01.Fields("M08").VALUE) '8월
        Call oDS_PS_USERDS02.setValue("U_ColReg12", loopCount, RecordSet01.Fields("M09").VALUE) '9월
        Call oDS_PS_USERDS02.setValue("U_ColReg13", loopCount, RecordSet01.Fields("M10").VALUE) '10월
        Call oDS_PS_USERDS02.setValue("U_ColReg14", loopCount, RecordSet01.Fields("M11").VALUE) '11월
        Call oDS_PS_USERDS02.setValue("U_ColReg15", loopCount, RecordSet01.Fields("M12").VALUE) '12월
        
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
    
PS_AD900_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_AD900_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_AD900_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_AD900_FormResize()
On Error GoTo PS_AD900_FormResize_Error
    
    '그룹박스 크기 동적 할당
    oForm01.Items("GrpBox01").Height = oForm01.Items("Mat01").Height + 20
    oForm01.Items("GrpBox01").Width = oForm01.Items("Mat01").Width + 25
    
    Exit Sub
PS_AD900_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_AD900_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_AD900_Print_Report01()
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

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", oRecordSet.RecordCount, False)

    Call ConnectODBC
    
    '// 조회조건문
    Dim BPLID As String '사업장
    Dim ReqEmpCd As String '요청자
    Dim ReqFrDt As String '요청일시작
    Dim ReqToDt As String '요청일종료
    Dim FinReqFrDt As String '완료요구일시작
    Dim FinReqToDt As String '완료요구일종료
    Dim PGType As String '프로그램구분
    Dim ReqType As String '요청구분
    
    Dim FinFrDt As String '완료일시작
    Dim FinToDt As String '완료일종료
    Dim AcStatus As String '접수상태
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    ReqEmpCd = oForm01.Items("ReqEmpCd").Specific.VALUE
    ReqFrDt = oForm01.Items("ReqFrDt").Specific.VALUE
    ReqToDt = oForm01.Items("ReqToDt").Specific.VALUE
    FinReqFrDt = oForm01.Items("FinReqFrDt").Specific.VALUE
    FinReqToDt = oForm01.Items("FinReqToDt").Specific.VALUE
    PGType = IIf(oForm01.Items("PGType").Specific.Selected.VALUE = "%", "", oForm01.Items("PGType").Specific.Selected.VALUE)
    ReqType = IIf(oForm01.Items("ReqType").Specific.Selected.VALUE = "%", "", oForm01.Items("ReqType").Specific.Selected.VALUE)
    
    FinFrDt = oForm01.Items("FinFrDt").Specific.VALUE
    FinToDt = oForm01.Items("FinToDt").Specific.VALUE
    AcStatus = IIf(oForm01.Items("AcStatus").Specific.Selected.VALUE = "%", "", oForm01.Items("AcStatus").Specific.Selected.VALUE)
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_AD900] 레포트"
    ReportName = "PS_AD900_01.rpt"
 
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
    sQry = "EXEC PS_AD900_01 '" & BPLID & "','" & ReqEmpCd & "','" & ReqFrDt & "','" & ReqToDt & "','" & FinReqFrDt & "','" & FinReqToDt & "','" & PGType & "','" & ReqType & "','" & FinFrDt & "','" & FinToDt & "','" & AcStatus & "'"

    oRecordSet.DoQuery sQry
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
    Set ProgBar01 = Nothing
    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub


