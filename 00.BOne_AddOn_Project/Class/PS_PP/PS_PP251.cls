VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP251"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_PP251.cls
'//  Module             : 생산 관리 > 애드온 > 생산관리
'//  Desc               : 월별 가동률 관리(기계)
'//  FormType           : PS_PP251
'//  Create Date(Start) : 2018.01.17
'//  Create Date(End)   :
'//  Modified Date      :
'//  Creator            : Song Myoung gyu
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form

Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix
Public oMat03 As SAPbouiCOM.Matrix

'Private oDS_PS_PP251H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP251L  As SAPbouiCOM.DBDataSource '라인(작업일보)
Private oDS_PS_PP251M  As SAPbouiCOM.DBDataSource '라인(구매요청)
Private oDS_PS_PP251N  As SAPbouiCOM.DBDataSource '라인(공용)

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

''//사용자구조체
'Private Type ItemInformations
'    ItemCode As String
'    BatchNum As String
'    Quantity As Long
'    OPORNo As Long
'    POR1No As Long
'    Check As Boolean
'    OPDNNo As Long
'    PDN1No As Long
'End Type
'Private ItemInformation() As ItemInformations
'Private ItemInformationCount As Long
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP251.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP251_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP251_CreateItems
    Call PS_PP251_ComboBox_Setting
    Call PS_PP251_Initial_Setting
    Call PS_PP251_CF_ChooseFromList
    Call PS_PP251_EnableMenus
    Call PS_PP251_SetDocument(oFromDocEntry01)
    Call PS_PP251_FormResize
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    
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
                Call PS_PP251_AddMatrixRow(oMat01.VisualRowCount, oMat01, oDS_PS_PP251L) '작업일보
                Call PS_PP251_AddMatrixRow(oMat02.VisualRowCount, oMat02, oDS_PS_PP251M) '구매요청
                Call PS_PP251_AddMatrixRow(oMat03.VisualRowCount, oMat03, oDS_PS_PP251N) '공용
            
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
                '작업일보
                Call oDS_PS_PP251L.RemoveRecord(oDS_PS_PP251L.Size - 1)
                Call oMat01.LoadFromDataSource
                
                '구매요청
                Call oDS_PS_PP251M.RemoveRecord(oDS_PS_PP251M.Size - 1)
                Call oMat02.LoadFromDataSource
                
                '공용
                Call oDS_PS_PP251N.RemoveRecord(oDS_PS_PP251N.Size - 1)
                Call oMat03.LoadFromDataSource
                
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
    ElseIf pval.ItemUID = "Mat03" Then
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
            Call PS_PP251_MTX01 '매트릭스에 데이터 로드
        ElseIf pval.ItemUID = "BtnSrch02" Then
            Call PS_PP251_MTX02
        ElseIf pval.ItemUID = "BtnSrch03" Then
            Call PS_PP251_MTX03
            
        ElseIf pval.ItemUID = "BtnSave01" Then
            Call PS_PP251_SaveData01
        ElseIf pval.ItemUID = "BtnSave02" Then
            Call PS_PP251_SaveData02
        ElseIf pval.ItemUID = "BtnSave03" Then
            Call PS_PP251_SaveData03
        
        ElseIf pval.ItemUID = "BtnCheck01" Then

            Call PS_PP251_CheckAll01

        ElseIf pval.ItemUID = "BtnDel02" Then
            Call PS_PP251_DeleteData02
'        ElseIf pval.ItemUID = "BtnPrint" Then
'
'            If oForm01.Mode = fm_ADD_MODE Then
'
'                Call PS_PP251_Print_Report(pval.ItemUID)
'
'            End If
            
        End If
        
    ElseIf pval.BeforeAction = False Then
    
        '폴더를 사용할 때는 필수 소스_S
        'Folder01이 선택되었을 때
        If pval.ItemUID = "Folder01" Then
            Call oForm01.Freeze(True)
            oForm01.PaneLevel = 1
            oForm01.DefButton = "BtnSrch01"
            oForm01.Settings.MatrixUID = "Mat01"
            oMat01.AutoResizeColumns
            Call oForm01.Freeze(False)
        End If
        
'        'Folder02가 선택되었을 때
'        If pval.ItemUID = "Folder02" Then
'            Call oForm01.Freeze(True)
'            oForm01.PaneLevel = 2
'            oForm01.DefButton = "BtnSrch02"
'            oForm01.Settings.MatrixUID = "Mat02"
'            oMat02.AutoResizeColumns
'            Call oForm01.Freeze(False)
'        End If
'
'        'Folder03이 선택되었을 때
'        If pval.ItemUID = "Folder03" Then
'            Call oForm01.Freeze(True)
'            oForm01.PaneLevel = 3
'            oForm01.DefButton = "BtnSrch03"
'            oForm01.Settings.MatrixUID = "Mat03"
'            oMat03.AutoResizeColumns
'            Call oForm01.Freeze(False)
'        End If
'        '폴더를 사용할 때는 필수 소스_E
    
        If pval.ItemUID = "PS_PP251" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        End If
        
    End If
    
    Exit Sub
Raise_EVENT_ITEM_PRESSED_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    If pval.BeforeAction = True Then
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum01", "") '작업일보-작번
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CpCode01", "") '작업일보-공정
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "WorkCode01", "") '작업일보-작업자
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum02", "") '구매요청-작번
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode02", "") '구매요청-청구자
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode03", "") '구매요청-사번
        
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
        ElseIf pval.ItemUID = "Mat03" Then
            If pval.Row > 0 Then
                
                Call oMat03.SelectRow(pval.Row, True, False)
                
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
'    Set PP030 = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    Call oForm01.Freeze(True)
    
    Dim oQuery01     As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "OrdNum01") Then '작업일보-작번
                
                oForm01.Items("OrdName01").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "OITM", "'" & oForm01.Items("OrdNum01").Specific.VALUE & "'") '작업일보-품명
                oForm01.Items("OrdSpec01").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "OITM", "'" & oForm01.Items("OrdNum01").Specific.VALUE & "'")  '작업일보-규격
            
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
        Call PS_PP251_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP251_FormResize
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
        Set oMat03 = Nothing
        
        'Set oDS_PS_PP251H = Nothing
        Set oDS_PS_PP251L = Nothing
        Set oDS_PS_PP251M = Nothing
        Set oDS_PS_PP251N = Nothing
        
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

Private Function PS_PP251_CreateItems() As Boolean
On Error GoTo PS_PP251_CreateItems_Error

    Call oForm01.Freeze(True)
    
'    Dim oQuery01 As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Set oDS_PS_PP251L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
'    Set oDS_PS_PP251M = oForm01.DataSources.DBDataSources("@PS_USERDS02")
'    Set oDS_PS_PP251N = oForm01.DataSources.DBDataSources("@PS_USERDS03")
    
    '매트릭스 초기화
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
'    Set oMat02 = oForm01.Items("Mat02").Specific
'    oMat02.SelectionMode = ms_NotSupported
'    oMat02.AutoResizeColumns
'
'    Set oMat03 = oForm01.Items("Mat03").Specific
'    oMat03.SelectionMode = ms_NotSupported
'    oMat03.AutoResizeColumns
    
    '//////////가동률 관리//////////S
    '기준년월
    Call oForm01.DataSources.UserDataSources.Add("StdYM01", dt_SHORT_TEXT, 6)
    Call oForm01.Items("StdYM01").Specific.DataBind.SetBound(True, "", "StdYM01")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType01", dt_SHORT_TEXT, 5)
    Call oForm01.Items("CardType01").Specific.DataBind.SetBound(True, "", "CardType01")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType01", dt_SHORT_TEXT, 5)
    Call oForm01.Items("ItemType01").Specific.DataBind.SetBound(True, "", "ItemType01")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum01", dt_SHORT_TEXT, 11)
    Call oForm01.Items("OrdNum01").Specific.DataBind.SetBound(True, "", "OrdNum01")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("OrdName01", dt_SHORT_TEXT, 200)
    Call oForm01.Items("OrdName01").Specific.DataBind.SetBound(True, "", "OrdName01")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("OrdSpec01", dt_SHORT_TEXT, 200)
    Call oForm01.Items("OrdSpec01").Specific.DataBind.SetBound(True, "", "OrdSpec01")
    
    '최종등록일
    
    '//////////가동률 관리//////////E
    
'    '//////////구매요청//////////S
'    '청구기간(Fr)
'    Call oForm01.DataSources.UserDataSources.Add("FrDt02", dt_DATE)
'    Call oForm01.Items("FrDt02").Specific.DataBind.SetBound(True, "", "FrDt02")
'    oForm01.DataSources.UserDataSources.Item("FrDt02").VALUE = Format(Date, "YYYYMM01") 'Format(DateAdd("m", -1, Date), "YYYYMM01") '전월1일
'
'    '청구기간(To)
'    Call oForm01.DataSources.UserDataSources.Add("ToDt02", dt_DATE)
'    Call oForm01.Items("ToDt02").Specific.DataBind.SetBound(True, "", "ToDt02")
'    oForm01.DataSources.UserDataSources.Item("ToDt02").VALUE = Format(Date, "YYYYMMDD") 'Replace(DateSerial(Year(Date), Month(Date), 1 - 1), "-", "") '전월말일
'
'    '작번
'    Call oForm01.DataSources.UserDataSources.Add("OrdNum02", dt_SHORT_TEXT, 11)
'    Call oForm01.Items("OrdNum02").Specific.DataBind.SetBound(True, "", "OrdNum02")
'
'    '작명
'    Call oForm01.DataSources.UserDataSources.Add("OrdName02", dt_SHORT_TEXT, 200)
'    Call oForm01.Items("OrdName02").Specific.DataBind.SetBound(True, "", "OrdName02")
'
'    '규격
'    Call oForm01.DataSources.UserDataSources.Add("OrdSpec02", dt_SHORT_TEXT, 200)
'    Call oForm01.Items("OrdSpec02").Specific.DataBind.SetBound(True, "", "OrdSpec02")
'
'    '청구자
'    Call oForm01.DataSources.UserDataSources.Add("CntcCode02", dt_SHORT_TEXT, 20)
'    Call oForm01.Items("CntcCode02").Specific.DataBind.SetBound(True, "", "CntcCode02")
'
'    '청구자성명
'    Call oForm01.DataSources.UserDataSources.Add("CntcName02", dt_SHORT_TEXT, 50)
'    Call oForm01.Items("CntcName02").Specific.DataBind.SetBound(True, "", "CntcName02")
'
'    '품의구분
'    Call oForm01.DataSources.UserDataSources.Add("POType02", dt_SHORT_TEXT, 5)
'    Call oForm01.Items("POType02").Specific.DataBind.SetBound(True, "", "POType02")
'
'    '청구사유
'    Call oForm01.DataSources.UserDataSources.Add("OrdCls02", dt_SHORT_TEXT, 5)
'    Call oForm01.Items("OrdCls02").Specific.DataBind.SetBound(True, "", "OrdCls02")
'
'    '연간품제외
'    Call oForm01.DataSources.UserDataSources.Add("YearPDYN02", dt_SHORT_TEXT, 1)
'    Call oForm01.Items("YearPDYN02").Specific.DataBind.SetBound(True, "", "YearPDYN02")
'    oForm01.Items("YearPDYN02").Specific.Checked = False
'    '//////////구매요청//////////E
'
'    '//////////공용//////////S
'    '공용기간(Fr)
'    Call oForm01.DataSources.UserDataSources.Add("FrDt03", dt_DATE)
'    Call oForm01.Items("FrDt03").Specific.DataBind.SetBound(True, "", "FrDt03")
'    oForm01.DataSources.UserDataSources.Item("FrDt03").VALUE = Format(Date, "YYYYMM01") 'Format(DateAdd("m", -1, Date), "YYYYMM01") '전월1일
'
'    '공용기간(To)
'    Call oForm01.DataSources.UserDataSources.Add("ToDt03", dt_DATE)
'    Call oForm01.Items("ToDt03").Specific.DataBind.SetBound(True, "", "ToDt03")
'    oForm01.DataSources.UserDataSources.Item("ToDt03").VALUE = Format(Date, "YYYYMMDD") 'Replace(DateSerial(Year(Date), Month(Date), 1 - 1), "-", "") '전월말일
'
'    '사번
'    Call oForm01.DataSources.UserDataSources.Add("CntcCode03", dt_SHORT_TEXT, 20)
'    Call oForm01.Items("CntcCode03").Specific.DataBind.SetBound(True, "", "CntcCode03")
'
'    '성명
'    Call oForm01.DataSources.UserDataSources.Add("CntcName03", dt_SHORT_TEXT, 50)
'    Call oForm01.Items("CntcName03").Specific.DataBind.SetBound(True, "", "CntcName03")
'
'    '목적구분
'    Call oForm01.DataSources.UserDataSources.Add("ObjCls03", dt_SHORT_TEXT, 5)
'    Call oForm01.Items("ObjCls03").Specific.DataBind.SetBound(True, "", "ObjCls03")
'
'    '목적
'    Call oForm01.DataSources.UserDataSources.Add("Object03", dt_SHORT_TEXT, 100)
'    Call oForm01.Items("Object03").Specific.DataBind.SetBound(True, "", "Object03")
'
'    '작번
'    Call oForm01.DataSources.UserDataSources.Add("OrdNum03", dt_SHORT_TEXT, 20)
'    Call oForm01.Items("OrdNum03").Specific.DataBind.SetBound(True, "", "OrdNum03")
'    '//////////공용//////////E
    
    Call oForm01.Freeze(False)
    Exit Function
PS_PP251_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP251_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP251_ComboBox_Setting()
On Error GoTo PS_PP251_ComboBox_Setting_Error
    
    Dim sQry As String
    
    Call oForm01.Freeze(True)

    '/////////가동률 관리//////////S
    '거래처구분
    sQry = "        SELECT      U_Minor,"
    sQry = sQry & "             U_CdName"
    sQry = sQry & " FROM        [@PS_SY001L]"
    sQry = sQry & " WHERE       Code = 'C100'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call oForm01.Items("CardType01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType01").Specific, sQry, "%", False, False)
    
    '품목구분
    sQry = "        SELECT      U_Minor,"
    sQry = sQry & "             U_CdName"
    sQry = sQry & " FROM        [@PS_SY001L]"
    sQry = sQry & " WHERE       Code = 'S002'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    Call oForm01.Items("ItemType01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType01").Specific, sQry, "%", False, False)
    '//////////가동률 관리//////////E
    
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_PP251_ComboBox_Setting_Error:
'    Set lRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP251_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP251_Initial_Setting()
On Error GoTo PS_PP251_Initial_Setting_Error

    '사업장 사용자의 소속 사업장 선택
    'Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '기준일자
    oForm01.Items("StdYM01").Specific.VALUE = Format(Now, "YYYYMM")
    
    '날짜 설정
'    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYYMMDD")
'    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    '날짜 빈칸으로 설정
'    oForm01.Items("ToDt").Specific.VALUE = ""
'    oForm01.Items("FrDt").Specific.VALUE = ""
 
    Exit Sub
PS_PP251_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP251_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP251_CF_ChooseFromList()
On Error GoTo PS_PP251_CF_ChooseFromList_Error
    Exit Sub
PS_PP251_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP251_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP251_FormItemEnabled()
On Error GoTo PS_PP251_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP251_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP251_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP251_AddMatrixRow(ByVal oRow As Long, ByVal prmMat As SAPbouiCOM.Matrix, ByVal prmDataSource As SAPbouiCOM.DBDataSource, Optional RowIserted As Boolean)
On Error GoTo PS_PP251_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        prmDataSource.InsertRecord (oRow)
    End If
    prmMat.AddRow
    prmDataSource.Offset = oRow
'    oDS_PS_PP251L.setValue "U_LineNum", oRow, oRow + 1
    prmMat.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP251_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP251_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP251_FormClear()
On Error GoTo PS_PP251_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP251'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP251_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP251_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP251_EnableMenus()
On Error GoTo PS_PP251_EnableMenus_Error
    Exit Sub
PS_PP251_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP251_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP251_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP251_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP251_FormItemEnabled

    Else

    End If
    Exit Sub
PS_PP251_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP251_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP251_DataValidCheck() As Boolean
On Error GoTo PS_PP251_DataValidCheck_Error
    PS_PP251_DataValidCheck = False
    Dim i As Long

'    If oForm01.Items("DocDateTo").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "견적일자 To은 필수입니다.", bmt_Short, True
'        oForm01.Items("DocDateTo").Click ct_Regular
'        PS_PP251_DataValidCheck = False
'        Exit Function
'    End If

    PS_PP251_DataValidCheck = True
    Exit Function
PS_PP251_DataValidCheck_Error:
    PS_PP251_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP251_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP251_MTX01()
'******************************************************************************
'Function ID : PS_PP251_MTX01()
'해당모듈    : PS_PP251
'기능        : 가동률 관리 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP251_MTX01_Error
    
    Dim loopCount   As Long
    Dim Query01     As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim StdYM    As String '기준년월
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim OrdNum   As String '작번
    
    StdYM = Trim(oForm01.Items("StdYM01").Specific.VALUE) '기준년월
    CardType = Trim(oForm01.Items("CardType01").Specific.VALUE) '거래처구분
    ItemType = Trim(oForm01.Items("ItemType01").Specific.VALUE) '품목구분
    OrdNum = Trim(oForm01.Items("OrdNum01").Specific.VALUE) '작번
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "         EXEC PS_PP251_01 '"
    Query01 = Query01 & StdYM & "','"
    Query01 = Query01 & CardType & "','"
    Query01 = Query01 & ItemType & "','"
    Query01 = Query01 & OrdNum & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_PP251_MTX01_Exit
    End If
    
'    oForm01.Items("RegDate").Specific.VALUE = RecordSet01.Fields("RegDate").VALUE
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP251L.InsertRecord(loopCount)
        End If
        oDS_PS_PP251L.Offset = loopCount
        
        Call oDS_PS_PP251L.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP251L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("Select").VALUE) '선택
        Call oDS_PS_PP251L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("ItemCode").VALUE) '품목코드
        Call oDS_PS_PP251L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("CardName").VALUE) '거래처명
        Call oDS_PS_PP251L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("InOut").VALUE) '자체 / 외주
        Call oDS_PS_PP251L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("GrpName").VALUE) '계열사명
        Call oDS_PS_PP251L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("ItemName").VALUE) '품명
        Call oDS_PS_PP251L.setValue("U_ColReg07", loopCount, RecordSet01.Fields("ItemSpec").VALUE) '규격
        Call oDS_PS_PP251L.setValue("U_ColReg08", loopCount, RecordSet01.Fields("ItemUnit").VALUE) '단위
        Call oDS_PS_PP251L.setValue("U_ColSum01", loopCount, RecordSet01.Fields("WOQty").VALUE) '지시수량
        Call oDS_PS_PP251L.setValue("U_ColSum02", loopCount, RecordSet01.Fields("WorkQty").VALUE) '생산수량
        Call oDS_PS_PP251L.setValue("U_ColSum03", loopCount, RecordSet01.Fields("OrdQty").VALUE) '수주수량
        Call oDS_PS_PP251L.setValue("U_ColSum04", loopCount, RecordSet01.Fields("OrdAmt").VALUE) '수주금액
        Call oDS_PS_PP251L.setValue("U_ColQty01", loopCount, RecordSet01.Fields("StdTime").VALUE) '표준공수
        Call oDS_PS_PP251L.setValue("U_ColQty02", loopCount, RecordSet01.Fields("WorkTime").VALUE) '실동공수(전월까지)
        Call oDS_PS_PP251L.setValue("U_ColQty03", loopCount, RecordSet01.Fields("AddTime").VALUE) '추가예상공수
        Call oDS_PS_PP251L.setValue("U_ColQty04", loopCount, RecordSet01.Fields("TotTime").VALUE) '공수합계
        Call oDS_PS_PP251L.setValue("U_ColQty05", loopCount, RecordSet01.Fields("Plan1st").VALUE) '1주계획
        Call oDS_PS_PP251L.setValue("U_ColQty06", loopCount, RecordSet01.Fields("Rslt1st").VALUE) '1주실적
        Call oDS_PS_PP251L.setValue("U_ColQty07", loopCount, RecordSet01.Fields("Plan2nd").VALUE) '2주계획
        Call oDS_PS_PP251L.setValue("U_ColQty08", loopCount, RecordSet01.Fields("Rslt2nd").VALUE) '2주실적
        Call oDS_PS_PP251L.setValue("U_ColQty09", loopCount, RecordSet01.Fields("Plan3rd").VALUE) '3주계획
        Call oDS_PS_PP251L.setValue("U_ColQty10", loopCount, RecordSet01.Fields("Rslt3rd").VALUE) '3주실적
        Call oDS_PS_PP251L.setValue("U_ColQty11", loopCount, RecordSet01.Fields("Plan4th").VALUE) '4주계획
        Call oDS_PS_PP251L.setValue("U_ColQty12", loopCount, RecordSet01.Fields("Rslt4th").VALUE) '4주실적
        Call oDS_PS_PP251L.setValue("U_ColQty13", loopCount, RecordSet01.Fields("Plan5th").VALUE) '5주계획
        Call oDS_PS_PP251L.setValue("U_ColQty14", loopCount, RecordSet01.Fields("Rslt5th").VALUE) '5주실적
        Call oDS_PS_PP251L.setValue("U_ColQty15", loopCount, RecordSet01.Fields("PlanNMth").VALUE) '차월계획
        Call oDS_PS_PP251L.setValue("U_ColQty16", loopCount, RecordSet01.Fields("PlanNNMth").VALUE) '차차월계획
        Call oDS_PS_PP251L.setValue("U_ColQty17", loopCount, RecordSet01.Fields("PlanLMth").VALUE) '이후계획
        Call oDS_PS_PP251L.setValue("U_ColDt01", loopCount, Format(RecordSet01.Fields("RegDate").VALUE, "YYYYMMDD")) '최종등록일
        
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
    
PS_PP251_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP251_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP251_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP251_MTX02()
'******************************************************************************
'Function ID : PS_PP251_MTX02()
'해당모듈    : PS_PP251
'기능        : 구매요청
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP251_MTX02_Error
    
    Dim loopCount   As Long
    Dim Query01     As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim FrDt     As String '요청일(시작)
    Dim ToDt     As String '요청일(종료)
    Dim OrdNum   As String '작번
    Dim CntcCode As String '청구자사번
    Dim POType   As String '품의구분
    Dim OrdCls   As String '청구사유
    Dim YearPdYN As String '연간품제외
    
    FrDt = Trim(oForm01.Items("FrDt02").Specific.VALUE) '요청일(시작)
    ToDt = Trim(oForm01.Items("ToDt02").Specific.VALUE) '요청일(종료)
    OrdNum = Trim(oForm01.Items("OrdNum02").Specific.VALUE) '작번
    CntcCode = Trim(oForm01.Items("CntcCode02").Specific.VALUE) '청구자사번
    POType = Trim(oForm01.Items("POType02").Specific.VALUE) '품의구분
    OrdCls = Trim(oForm01.Items("OrdCls02").Specific.VALUE) '청구사유
    
    If oForm01.DataSources.UserDataSources.Item("YearPDYN02").VALUE = "Y" Then '연간품여부
        YearPdYN = "Y"
    Else
        YearPdYN = "N"
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "         EXEC PS_PP251_11 '"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "','"
    Query01 = Query01 & OrdNum & "','"
    Query01 = Query01 & CntcCode & "','"
    Query01 = Query01 & POType & "','"
    Query01 = Query01 & OrdCls & "','"
    Query01 = Query01 & YearPdYN & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat02.Clear
        GoTo PS_PP251_MTX02_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP251M.InsertRecord(loopCount)
        End If
        oDS_PS_PP251M.Offset = loopCount
        
        Call oDS_PS_PP251M.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP251M.setValue("U_ColReg01", loopCount, RecordSet01.Fields("Select").VALUE) '선택
        Call oDS_PS_PP251M.setValue("U_ColReg02", loopCount, RecordSet01.Fields("MM005Entry").VALUE) '청구번호
        Call oDS_PS_PP251M.setValue("U_ColReg03", loopCount, RecordSet01.Fields("MM005RCode").VALUE) '청구사유
        Call oDS_PS_PP251M.setValue("U_ColReg04", loopCount, RecordSet01.Fields("MM005Reasn").VALUE) '세부사항
        Call oDS_PS_PP251M.setValue("U_ColReg05", loopCount, RecordSet01.Fields("PP030Entry").VALUE) '작지문서번호
        Call oDS_PS_PP251M.setValue("U_ColReg06", loopCount, RecordSet01.Fields("PP030Date").VALUE) '작지등록일
        Call oDS_PS_PP251M.setValue("U_ColReg07", loopCount, RecordSet01.Fields("MM005Date").VALUE) '구매요청일
        Call oDS_PS_PP251M.setValue("U_ColReg08", loopCount, RecordSet01.Fields("POTypeNM").VALUE) '품의구분
        Call oDS_PS_PP251M.setValue("U_ColReg09", loopCount, RecordSet01.Fields("PP080Date").VALUE) '생산완료일
        Call oDS_PS_PP251M.setValue("U_ColReg10", loopCount, RecordSet01.Fields("CreateDate").VALUE) '등록일(시스템)
        Call oDS_PS_PP251M.setValue("U_ColSum01", loopCount, RecordSet01.Fields("MM005SDiff").VALUE) '시스템일-청구일
        Call oDS_PS_PP251M.setValue("U_ColSum02", loopCount, RecordSet01.Fields("PP030Diff").VALUE) '작지등록일-청구일
        Call oDS_PS_PP251M.setValue("U_ColSum03", loopCount, RecordSet01.Fields("PP080Diff").VALUE) '생산완료일-청구일
        Call oDS_PS_PP251M.setValue("U_ColReg14", loopCount, RecordSet01.Fields("ItemCode").VALUE) '품목코드
        Call oDS_PS_PP251M.setValue("U_ColReg15", loopCount, RecordSet01.Fields("ItemName").VALUE) '품명
        Call oDS_PS_PP251M.setValue("U_ColReg16", loopCount, RecordSet01.Fields("OrdNum").VALUE) '작번
        Call oDS_PS_PP251M.setValue("U_ColReg17", loopCount, RecordSet01.Fields("OrdName").VALUE) '작명
        Call oDS_PS_PP251M.setValue("U_ColReg18", loopCount, RecordSet01.Fields("OrdSpec").VALUE) '규격
        Call oDS_PS_PP251M.setValue("U_ColReg19", loopCount, RecordSet01.Fields("CntcCode").VALUE) '청구자사번
        Call oDS_PS_PP251M.setValue("U_ColReg20", loopCount, RecordSet01.Fields("CntcName").VALUE) '청구자성명
        Call oDS_PS_PP251M.setValue("U_ColSum04", loopCount, RecordSet01.Fields("MM030Amt").VALUE) '품의금액
        Call oDS_PS_PP251M.setValue("U_ColReg22", loopCount, RecordSet01.Fields("MM030Date").VALUE) '품의일자
        Call oDS_PS_PP251M.setValue("U_ColReg23", loopCount, RecordSet01.Fields("YearPDYN").VALUE) '연간품여부
        
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
    
PS_PP251_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP251_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP251_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP251_MTX03()
'******************************************************************************
'Function ID : PS_PP251_MTX03()
'해당모듈    : PS_PP251
'기능        : 공용조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP251_MTX03_Error
    
    Dim loopCount   As Long
    Dim Query01     As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim FrDt     As String '공용시작일(시작)
    Dim ToDt     As String '공용시작일(종료)
    Dim CntcCode As String '사번
    Dim ObjCls   As String '목적구분
    Dim Object   As String '목적
    Dim OrdNum   As String '작번
    
    FrDt = Trim(oForm01.Items("FrDt03").Specific.VALUE) '공용시작일(시작)
    ToDt = Trim(oForm01.Items("ToDt03").Specific.VALUE) '공용시작일(종료)
    CntcCode = Trim(oForm01.Items("CntcCode03").Specific.VALUE) '사번
    ObjCls = Trim(oForm01.Items("ObjCls03").Specific.VALUE) '목적구분
    Object = Trim(oForm01.Items("Object03").Specific.VALUE) '목적
    OrdNum = Trim(oForm01.Items("OrdNum03").Specific.VALUE) '작번
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "         EXEC PS_PP251_21 '"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "','"
    Query01 = Query01 & CntcCode & "','"
    Query01 = Query01 & ObjCls & "','"
    Query01 = Query01 & Object & "','"
    Query01 = Query01 & OrdNum & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    oMat03.Clear
    oMat03.FlushToDataSource
    oMat03.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat03.Clear
        GoTo PS_PP251_MTX03_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP251N.InsertRecord(loopCount)
        End If
        oDS_PS_PP251N.Offset = loopCount
        
        Call oDS_PS_PP251N.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP251N.setValue("U_ColReg01", loopCount, RecordSet01.Fields("Select").VALUE) '선택
        Call oDS_PS_PP251N.setValue("U_ColReg02", loopCount, RecordSet01.Fields("DocEntry").VALUE) '관리번호
        Call oDS_PS_PP251N.setValue("U_ColReg03", loopCount, RecordSet01.Fields("MSTCOD").VALUE) '사원번호
        Call oDS_PS_PP251N.setValue("U_ColReg04", loopCount, RecordSet01.Fields("MSTNAM").VALUE) '사원성명
        Call oDS_PS_PP251N.setValue("U_ColReg05", loopCount, RecordSet01.Fields("FrDate").VALUE) '출발일자
        Call oDS_PS_PP251N.setValue("U_ColReg06", loopCount, RecordSet01.Fields("FrHour").VALUE) '출발시각
        Call oDS_PS_PP251N.setValue("U_ColReg07", loopCount, RecordSet01.Fields("ToDate").VALUE) '도착일자
        Call oDS_PS_PP251N.setValue("U_ColReg08", loopCount, RecordSet01.Fields("ToHour").VALUE) '도착시각
        Call oDS_PS_PP251N.setValue("U_ColReg09", loopCount, RecordSet01.Fields("OrdNum").VALUE) '작번
        Call oDS_PS_PP251N.setValue("U_ColReg10", loopCount, RecordSet01.Fields("ObjCls").VALUE) '목적구분
        Call oDS_PS_PP251N.setValue("U_ColReg11", loopCount, RecordSet01.Fields("Object").VALUE) '목적내용
        Call oDS_PS_PP251N.setValue("U_ColReg12", loopCount, RecordSet01.Fields("Dest2").VALUE) '목적지
        Call oDS_PS_PP251N.setValue("U_ColSum01", loopCount, RecordSet01.Fields("TransExp").VALUE) '교통비
        Call oDS_PS_PP251N.setValue("U_ColSum02", loopCount, RecordSet01.Fields("DayExp").VALUE) '일비
        Call oDS_PS_PP251N.setValue("U_ColSum03", loopCount, RecordSet01.Fields("FoodExp").VALUE) '식비
        Call oDS_PS_PP251N.setValue("U_ColSum04", loopCount, RecordSet01.Fields("ParkExp").VALUE) '주차비
        Call oDS_PS_PP251N.setValue("U_ColSum05", loopCount, RecordSet01.Fields("TollExp").VALUE) '도로비
        Call oDS_PS_PP251N.setValue("U_ColSum06", loopCount, RecordSet01.Fields("TotalExp").VALUE) '계
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat03.LoadFromDataSource
    oMat03.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP251_MTX03_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP251_MTX03_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP251_MTX03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP251_FormResize()
On Error GoTo PS_PP251_FormResize_Error
    
    '그룹박스 크기 동적 할당
    oForm01.Items("GrpBox01").Height = oForm01.Items("Mat01").Height + 120
    oForm01.Items("GrpBox01").Width = oForm01.Items("Mat01").Width + 20
    
    oMat01.AutoResizeColumns
    
    Exit Sub
    
PS_PP251_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP251_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Function PS_PP251_SaveData01() As Boolean
'******************************************************************************
'Function ID : PS_PP251_SaveData01()
'해당모듈    : PS_PP251
'기능        : 가동률 관리의 자료 저장
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP251_SaveData01_Error

    Dim loopCount As Integer
    Dim sQry      As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim StdYM           As String '기준년월
        Dim OrdNum          As String '품목코드
        Dim AddTime         As Double '추가예상공수
        Dim Plan1st         As Double '1주차계획
        Dim Plan2nd     As Double '2주차계획
        Dim Plan3rd         As Double '3주차계획
        Dim Plan4th     As Double '4주차계획
        Dim Plan5th         As Double '5주차계획
        Dim PlanNMth    As Double '차월계획
        Dim PlanNNMth   As Double '차차월계획
        Dim PlanLMth    As Double '이후계획
        Dim RegDate         As String '최종등록일

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)
    
    StdYM = oForm01.DataSources.UserDataSources.Item("StdYM01").VALUE

    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP251L.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            OrdNum = Trim(oDS_PS_PP251L.GetValue("U_ColReg02", loopCount)) '품목코드
            AddTime = Trim(oDS_PS_PP251L.GetValue("U_ColQty03", loopCount)) '추가예상공수
            Plan1st = Trim(oDS_PS_PP251L.GetValue("U_ColQty05", loopCount)) '1주차계획
            Plan2nd = Trim(oDS_PS_PP251L.GetValue("U_ColQty07", loopCount)) '2주차계획
            Plan3rd = Trim(oDS_PS_PP251L.GetValue("U_ColQty09", loopCount)) '3주차계획
            Plan4th = Trim(oDS_PS_PP251L.GetValue("U_ColQty11", loopCount)) '4주차계획
            Plan5th = Trim(oDS_PS_PP251L.GetValue("U_ColQty13", loopCount)) '5주차계획
            PlanNMth = Trim(oDS_PS_PP251L.GetValue("U_ColQty15", loopCount)) '차월계획
            PlanNNMth = Trim(oDS_PS_PP251L.GetValue("U_ColQty16", loopCount)) '차차월계획
            PlanLMth = Trim(oDS_PS_PP251L.GetValue("U_ColQty17", loopCount)) '이후계획
            RegDate = Trim(oDS_PS_PP251L.GetValue("U_ColDt01", loopCount)) '최종등록일

            sQry = "      EXEC [PS_PP251_02] '"
            sQry = sQry & StdYM & "','" '기준년월
            sQry = sQry & OrdNum & "','" '품목코드
            sQry = sQry & AddTime & "','" '추가예상공수
            sQry = sQry & Plan1st & "','" '1주차계획
            sQry = sQry & Plan2nd & "','" '2주차계획
            sQry = sQry & Plan3rd & "','" '3주차계획
            sQry = sQry & Plan4th & "','" '4주차계획
            sQry = sQry & Plan5th & "','" '5주차계획
            sQry = sQry & PlanNMth & "','" '차월계획
            sQry = sQry & PlanNNMth & "','" '차차월계획
            sQry = sQry & PlanLMth & "','" '이후계획
            sQry = sQry & RegDate & "'" '최종등록일
            
            Call RecordSet01.DoQuery(sQry)
            
            ProgBar01.VALUE = ProgBar01.VALUE + 1
            ProgBar01.Text = ProgBar01.VALUE & "/" & oMat01.VisualRowCount - 1 & "건 저장중..."
            
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    MDC_Com.MDC_GF_Message "저장 완료!", "S"
    
    Set RecordSet01 = Nothing
    
    PS_PP251_SaveData01 = True

    Exit Function

PS_PP251_SaveData01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PS_PP251_SaveData01 = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP251_SaveData01_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_PP251_SaveData02() As Boolean
'******************************************************************************
'Function ID : PS_PP251_SaveData02()
'해당모듈    : PS_PP251
'기능        : 구매요청의 청구사유, 세부사항 업데이트
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP251_SaveData02_Error

    Dim loopCount As Integer
    Dim sQry      As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim MM005Entry  As String '청구번호
    Dim MM005RCode  As String '청구사유
    Dim MM005Reasn  As String '세부사항
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)
    
    Call oMat02.FlushToDataSource
    For loopCount = 0 To oMat02.VisualRowCount - 1
        If Trim(oDS_PS_PP251M.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            MM005Entry = Trim(oDS_PS_PP251M.GetValue("U_ColReg02", loopCount)) '청구번호
            MM005RCode = Trim(oDS_PS_PP251M.GetValue("U_ColReg03", loopCount)) '청구사유
            MM005Reasn = Trim(oDS_PS_PP251M.GetValue("U_ColReg04", loopCount)) '세부상태
            
            sQry = "      EXEC [PS_PP251_12] '"
            sQry = sQry & MM005Entry & "','" '청구번호
            sQry = sQry & MM005RCode & "','" '청구사유
            sQry = sQry & MM005Reasn & "'" '세부사항
            
            Call RecordSet01.DoQuery(sQry)
            
            ProgBar01.VALUE = ProgBar01.VALUE + 1
            ProgBar01.Text = ProgBar01.VALUE & "/" & oMat01.VisualRowCount - 1 & "건 저장중..."
            
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    MDC_Com.MDC_GF_Message "저장 완료!", "S"
    
    Set RecordSet01 = Nothing
    
    PS_PP251_SaveData02 = True

    Exit Function

PS_PP251_SaveData02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PS_PP251_SaveData02 = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP251_SaveData02_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_PP251_SaveData03() As Boolean
'******************************************************************************
'Function ID : PS_PP251_SaveData03()
'해당모듈    : PS_PP251
'기능        : 공용 목적구분 업데이트
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP251_SaveData03_Error

    Dim loopCount As Integer
    Dim sQry      As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry  As String '공용관리번호
    Dim OrdNum    As String '작번
    Dim ObjCls    As String '공용목적구분
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, False)
    
    Call oMat03.FlushToDataSource
    For loopCount = 0 To oMat03.VisualRowCount - 1
        If Trim(oDS_PS_PP251N.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            DocEntry = Trim(oDS_PS_PP251N.GetValue("U_ColReg02", loopCount)) '공용관리번호
            OrdNum = Trim(oDS_PS_PP251N.GetValue("U_ColReg09", loopCount)) '작번
            ObjCls = Trim(oDS_PS_PP251N.GetValue("U_ColReg10", loopCount)) '공용목적구분
            
            sQry = "      EXEC [PS_PP251_22] '"
            sQry = sQry & DocEntry & "','" '공용관리번호
            sQry = sQry & OrdNum & "','" '작번
            sQry = sQry & ObjCls & "'" '공용목적구분
            
            Call RecordSet01.DoQuery(sQry)
            
            ProgBar01.VALUE = ProgBar01.VALUE + 1
            ProgBar01.Text = ProgBar01.VALUE & "/" & oMat01.VisualRowCount - 1 & "건 저장중..."
            
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    MDC_Com.MDC_GF_Message "저장 완료!", "S"
    
    Set RecordSet01 = Nothing
    
    PS_PP251_SaveData03 = True

    Exit Function

PS_PP251_SaveData03_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    PS_PP251_SaveData03 = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_PP251_SaveData03_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Sub PS_PP251_DeleteData02()
'******************************************************************************
'Function ID : PS_PP251_DeleteData02()
'해당모듈    : PS_PP251
'기능        : 청구사유삭제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP251_DeleteData02_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim MM005Entry As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("삭제 중...", 100, False)
    
    Call oMat02.FlushToDataSource
    
    For loopCount = 0 To oMat02.VisualRowCount - 1
        If Trim(oDS_PS_PP251M.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            MM005Entry = Trim(oDS_PS_PP251M.GetValue("U_ColReg02", loopCount))
            
            sQry = "      EXEC [PS_PP251_13] '"
            sQry = sQry & MM005Entry & "'" '청구번호
            
            Call oRecordSet01.DoQuery(sQry)
        End If
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Exit Sub

PS_PP251_DeleteData02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP251_DeleteData02_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
    
End Sub

Private Sub PS_PP251_CheckAll01()
'******************************************************************************
'Function ID : PS_PP251_CheckAll01()
'해당모듈    : PS_PP251
'기능        : 체크박스 전체선택, 해제
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP251_CheckAll01_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm01.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP251L.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PS_PP251L.Offset = loopCount
        If CheckType = "N" Then
            Call oDS_PS_PP251L.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PS_PP251L.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)

    Exit Sub
    
PS_PP251_CheckAll01_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_PP251_CheckAll01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Sub PS_PP251_Print_Report(ByVal prmItemUID As String)
On Error GoTo PS_PP251_Print_Report_Error

    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
   
    Call ConnectODBC
   
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    
    WinTitle = "[PS_PP251] 레포트"
    
    Dim FrDt As String '작번등록년월(시작)
    Dim ToDt As String '작번등록년월(종료)
    Dim FrgnName As String '품명
    Dim CardType As String '거래처구분
    Dim SPEC As String '규격
    Dim ItemClass As String '품목구분
    Dim WCYN As String '생산완료여부
    Dim CardCode As String '거래처
    Dim ItemCode As String '품목(작번)
    
    Dim OrdNum As String 'Main작번
    Dim OrdSub As String 'Sub작번
    Dim CpCode As String '공정코드
    
    If prmItemUID = "Btn01" Then '작업일보
    
        FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE)
        ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE)
        FrgnName = IIf(oForm01.Items("FrgnName").Specific.VALUE = "", "%", oForm01.Items("FrgnName").Specific.VALUE)
        CardType = Trim(oForm01.Items("CardType").Specific.Selected.VALUE)
        SPEC = IIf(oForm01.Items("Spec").Specific.VALUE = "", "%", oForm01.Items("Spec").Specific.VALUE)
        ItemClass = Trim(oForm01.Items("ItemClass").Specific.Selected.VALUE)
        WCYN = IIf(oForm01.Items("WCYN").Specific.Selected.VALUE = "", "%", oForm01.Items("WCYN").Specific.Selected.VALUE)
        CardCode = IIf(oForm01.Items("CardCode").Specific.VALUE = "", "%", oForm01.Items("CardCode").Specific.VALUE)
        ItemCode = IIf(oForm01.Items("ItemCode").Specific.VALUE = "", "%", oForm01.Items("ItemCode").Specific.VALUE)
       
        ReportName = "PS_PP251_01.rpt"
    
        Query01 = "EXEC PS_PP251_01 '" & FrDt & "','" & ToDt & "','" & FrgnName & "','" & CardType & "','" & SPEC & "','" & ItemClass & "','" & WCYN & "','" & CardCode & "','" & ItemCode & "'"
       
    ElseIf prmItemUID = "Btn02" Then '구매요청
    
        OrdNum = oMat02.Columns("OrdNum").Cells(1).Specific.VALUE
    
        ReportName = "PS_PP251_02.rpt"
    
        Query01 = "EXEC PS_PP251_02 '" & OrdNum & "'"
    
    ElseIf prmItemUID = "Btn03" Then '구매요청
    
        OrdNum = Trim(oForm01.Items("MainJak").Specific.VALUE)
        OrdSub = Trim(oForm01.Items("SubJak").Specific.VALUE)
        
        ReportName = "PS_PP251_03.rpt"
    
        Query01 = "EXEC PS_PP251_03 '" & OrdNum & "','" & OrdSub & "'"
    
    End If

    '// Formula 수식필드
    ReDim gRpt_Formula(0)
    ReDim gRpt_Formula_Value(0)
       
    '// SubReport
    ReDim gRpt_SRptSqry(0)
    ReDim gRpt_SRptName(0)
    
    ReDim gRpt_SFormula(0, 0)
    ReDim gRpt_SFormula_Value(0, 0)
    
    gRpt_SFormula(0, 0) = ""
    gRpt_SFormula_Value(0, 0) = ""
    
    Call RecordSet01.DoQuery(Query01)

'    If RecordSet01.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PS_PP251_Print_Report_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", Query01, "", "N", "V") = False Then
    End If
    
    Set RecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP251_Print_Report_Error:
    Set RecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_PP251_Print_Report:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub



