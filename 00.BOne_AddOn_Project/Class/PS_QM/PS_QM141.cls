VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM141"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File : PS_QM141.cls
'//  Module : 품질관리
'//  Desc : 실패비용현황(기계,공구)
'//  FormType : PS_QM141
'//  Create Date : 2013.08.29
'//  Modifier : Song Myounggyu
'//  Modify Date :
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM141H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_QM141L As SAPbouiCOM.DBDataSource '등록라인
'Private oDS_PS_USERDS01 As SAPbouiCOM.DBDataSource    '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM141.srf")
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
    
    oFormUniqueID01 = "PS_QM141_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM141_CreateItems
    Call PS_QM141_ComboBox_Setting
    Call PS_QM141_Initial_Setting
    Call PS_QM141_CF_ChooseFromList
    Call PS_QM141_EnableMenus
    Call PS_QM141_SetDocument(oFromDocEntry01)
    Call PS_QM141_FormResize
    
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
                Call PS_QM141_AddMatrixRow(oMat01.VisualRowCount)
            
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
                Call oDS_PS_QM141L.RemoveRecord(oDS_PS_QM141L.Size - 1)
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
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_QM141_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    Call PS_QM141_MTX01 '매트릭스에 데이터 로드
                End If
            End If
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_QM141_Print_Report01
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
        ElseIf pval.ItemUID = "Btn_Print" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
        
                Call PS_QM141_Print_Report01
            
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM141" Then
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
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '거래처코드 포맷서치 활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '품목코드(작번) 포맷서치 활성
        
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
        
            If (pval.ItemUID = "CardCode") Then
                oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf (pval.ItemUID = "ItemCode") Then
                oQuery01 = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf (pval.ItemUID = "CntcCode") Then
                oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
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
        Call PS_QM141_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_QM141_FormResize
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


Private Function PS_QM141_CreateItems() As Boolean
On Error GoTo PS_QM141_CreateItems_Error
    Call oForm01.Freeze(True)
'    Dim oQuery01 As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'
    Set oDS_PS_QM141L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '매트릭스 초기화
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    '사업장_E
    
    '기간 시작_S
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_DATE)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    '기간 시작_E

    '기간 종료_S
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    '기간 종료_E
    
    Call oForm01.Freeze(False)
    Exit Function
PS_QM141_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM141_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_QM141_ComboBox_Setting()
On Error GoTo PS_QM141_ComboBox_Setting_Error
    
    Call oForm01.Freeze(True)
     
        '사업장 콤보박스 세팅_S
        Call oForm01.Items("BPLId").Specific.ValidValues.Add("%", "전체")
        Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
        Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        '사업장 콤보박스 세팅_S
        
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
'        '매트릭스 컬럼의 콤보박스 설정_S
'        Dim oColumn As SAPbouiCOM.Column
'        Dim lRecordSet As SAPbobsCOM.Recordset '레코드셋
'        Dim lQuery As String '쿼리스트링

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
'        '매트릭스 컬럼의 콤보박스 설정_E
        
        
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM141_ComboBox_Setting_Error:
'    Set lRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM141_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM141_Initial_Setting()
On Error GoTo PS_QM141_Initial_Setting_Error
    '사업장 사용자의 소속 사업장 선택
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 
'    '체크박스 설정
'    oForm01.Items("Chk01").Specific.Checked = False '기본으로 체크해제(2011.11.02 송명규 수정)
'    oForm01.Items("Chk02").Specific.Checked = True
    
    '날짜 설정
    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYYMMDD")
    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM") + "01"
'    '날짜 빈칸으로 설정
'    oForm01.Items("ToDt").Specific.VALUE = ""
'    oForm01.Items("FrDt").Specific.VALUE = ""
 
    Exit Sub
PS_QM141_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM141_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_QM141_CF_ChooseFromList()
On Error GoTo PS_QM141_CF_ChooseFromList_Error
    Exit Sub
PS_QM141_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM141_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM141_FormItemEnabled()
On Error GoTo PS_QM141_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM141_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM141_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM141_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM141_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM141L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_QM141L.offset = oRow
'    oDS_PS_QM141L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM141_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM141_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM141_FormClear()
On Error GoTo PS_QM141_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM141'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_QM141_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM141_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM141_EnableMenus()
On Error GoTo PS_QM141_EnableMenus_Error
    Exit Sub
PS_QM141_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM141_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM141_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM141_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM141_FormItemEnabled

    Else

    End If
    Exit Sub
PS_QM141_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM141_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_QM141_DataValidCheck() As Boolean
On Error GoTo PS_QM141_DataValidCheck_Error
    PS_QM141_DataValidCheck = False
    Dim i As Long

'    If oForm01.Items("DocDateTo").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "견적일자 To은 필수입니다.", bmt_Short, True
'        oForm01.Items("DocDateTo").Click ct_Regular
'        PS_QM141_DataValidCheck = False
'        Exit Function
'    End If

    PS_QM141_DataValidCheck = True
    Exit Function
PS_QM141_DataValidCheck_Error:
    PS_QM141_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_QM141_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_QM141_MTX01()
On Error GoTo PS_QM141_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String '사업장
    Dim FrDt As String '기간시작
    Dim ToDt As String '기간종료
    
    BPLID = oForm01.Items("BPLId").Specific.Selected.VALUE
    FrDt = oForm01.Items("FrDt").Specific.VALUE
    ToDt = oForm01.Items("ToDt").Specific.VALUE
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_QM141_01 '" & BPLID & "','" & FrDt & "','" & ToDt & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_QM141_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_QM141L.InsertRecord(loopCount)
        End If
        oDS_PS_QM141L.offset = loopCount
        
        Call oDS_PS_QM141L.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_QM141L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("Class").VALUE) '구분
        Call oDS_PS_QM141L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("PreYear").VALUE) '전년실적
        Call oDS_PS_QM141L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("Mon01").VALUE) '1월
        Call oDS_PS_QM141L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("Mon02").VALUE) '2월
        Call oDS_PS_QM141L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("Mon03").VALUE) '3월
        Call oDS_PS_QM141L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("Mon04").VALUE) '4월
        Call oDS_PS_QM141L.setValue("U_ColReg07", loopCount, RecordSet01.Fields("Mon05").VALUE) '5월
        Call oDS_PS_QM141L.setValue("U_ColReg08", loopCount, RecordSet01.Fields("Mon06").VALUE) '6월
        Call oDS_PS_QM141L.setValue("U_ColReg09", loopCount, RecordSet01.Fields("Mon07").VALUE) '7월
        Call oDS_PS_QM141L.setValue("U_ColReg10", loopCount, RecordSet01.Fields("Mon08").VALUE) '8월
        Call oDS_PS_QM141L.setValue("U_ColReg11", loopCount, RecordSet01.Fields("Mon09").VALUE) '9월
        Call oDS_PS_QM141L.setValue("U_ColReg12", loopCount, RecordSet01.Fields("Mon10").VALUE) '10월
        Call oDS_PS_QM141L.setValue("U_ColReg13", loopCount, RecordSet01.Fields("Mon11").VALUE) '11월
        Call oDS_PS_QM141L.setValue("U_ColReg14", loopCount, RecordSet01.Fields("Mon12").VALUE) '12월
        Call oDS_PS_QM141L.setValue("U_ColReg15", loopCount, RecordSet01.Fields("Total").VALUE) '계
        
'        Call oDS_PS_QM141L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("SO_No").VALUE) '오더번호
'        Call oDS_PS_QM141L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("LotNo").VALUE) '주문번호
'        Call oDS_PS_QM141L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("TradeType").VALUE) '거래형태
'        Call oDS_PS_QM141L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("CardCode").VALUE) '거래처코드
'        Call oDS_PS_QM141L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("CardName").VALUE) '거래처명
'        Call oDS_PS_QM141L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("ItemCode").VALUE) '품목코드(작번)
'        Call oDS_PS_QM141L.setValue("U_ColReg07", loopCount, RecordSet01.Fields("ItemName").VALUE) '품명
'        Call oDS_PS_QM141L.setValue("U_ColReg08", loopCount, RecordSet01.Fields("Spec").VALUE) '규격
'        Call oDS_PS_QM141L.setValue("U_ColReg11", loopCount, RecordSet01.Fields("Unit").VALUE) '단위
'        Call oDS_PS_QM141L.setValue("U_ColDt01", loopCount, Format(RecordSet01.Fields("DocDate").VALUE, "yyyyMMdd")) '수주일
'        Call oDS_PS_QM141L.setValue("U_ColDt02", loopCount, Format(RecordSet01.Fields("DueDate").VALUE, "yyyyMMdd")) '납기일
'        Call oDS_PS_QM141L.setValue("U_ColQty01", loopCount, RecordSet01.Fields("UP_Qty").VALUE) '미납수량
'        Call oDS_PS_QM141L.setValue("U_ColSum01", loopCount, RecordSet01.Fields("UP_Amt").VALUE) '미납금액
'        Call oDS_PS_QM141L.setValue("U_ColQty02", loopCount, RecordSet01.Fields("SO_Qty").VALUE) '수주수량
'        Call oDS_PS_QM141L.setValue("U_ColSum02", loopCount, RecordSet01.Fields("SO_Amt").VALUE) '수주금액
'        Call oDS_PS_QM141L.setValue("U_ColReg09", loopCount, RecordSet01.Fields("Req_No").VALUE) '생산의뢰번호
'        Call oDS_PS_QM141L.setValue("U_ColQty03", loopCount, RecordSet01.Fields("Req_Qty").VALUE) '생산의뢰수량
'        Call oDS_PS_QM141L.setValue("U_ColQty04", loopCount, RecordSet01.Fields("Deli_Qty").VALUE) '출고수량
'        Call oDS_PS_QM141L.setValue("U_ColSum03", loopCount, RecordSet01.Fields("Deli_Amt").VALUE) '출고금액
'        Call oDS_PS_QM141L.setValue("U_ColQty05", loopCount, RecordSet01.Fields("AR_Qty").VALUE) '납품수량
'        Call oDS_PS_QM141L.setValue("U_ColSum04", loopCount, RecordSet01.Fields("AR_Amt").VALUE) '납품금액
'        Call oDS_PS_QM141L.setValue("U_ColReg10", loopCount, RecordSet01.Fields("DocStatus").VALUE) '문서상태
        
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
    
PS_QM141_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_QM141_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM141_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_QM141_FormResize()
On Error GoTo PS_QM141_FormResize_Error
    
    Exit Sub
PS_QM141_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM141_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM141_Print_Report01()
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
    
    BPLID = oForm01.Items("BPLId").Specific.Selected.VALUE
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
    WinTitle = "[PS_QM141] 레포트"
    ReportName = "PS_QM141.rpt"
 
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
    sQry = "EXEC PS_QM141_01 '" & BPLID & "','" & ItemClass & "','" & TradeType & "','" & FrDt & "','" & ToDt & "','" & CardCode & "','" & ItemCode & "','" & DocStatus & "','" & Chk01 & "','" & Chk02 & "'"

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


