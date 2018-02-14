VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD035"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_SD035.cls
'//  Module : 판매관관리>납품관리
'//  Desc : 출하리스트 등록 및 조회
'//  FormType : PS_SD035
'//  Create Date : 2014.04.22
'//  Modified Date : 2014.04.24
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_SD035H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_SD035L As SAPbouiCOM.DBDataSource '등록라인

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

Private oDocType01 As String
Private oDocEntry01 As String
Private oFormMode01 As SAPbouiCOM.BoFormMode

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD035.srf")
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
    
    oFormUniqueID01 = "PS_SD035_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
'    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_SD035_CreateItems
    Call PS_SD035_ComboBox_Setting
    Call PS_SD035_Initial_Setting
    Call PS_SD035_CF_ChooseFromList
    Call PS_SD035_EnableMenus
    Call PS_SD035_SetDocument(oFromDocEntry01)
    Call PS_SD035_FormResize
    
    oForm01.EnableMenu ("1283"), False         '// 삭제
    oForm01.EnableMenu ("1287"), False         '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
    oForm01.ActiveItem = "ORDR_FrDt"
    
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

    Dim i As Integer

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_SD035_AddMatrixRow(oMat01.VisualRowCount)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_SD035_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_SD035_FormItemEnabled '//UDO방식
                Call PS_SD035_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_SD035_FormItemEnabled
            Case "1287"
            
                oForm01.Freeze True
                Call oDS_PS_SD035H.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat01.VisualRowCount - 1
                    Call oMat01.FlushToDataSource
                    Call oDS_PS_SD035H.setValue("DocEntry", i, "")
                    Call oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_SD035L.RemoveRecord(oDS_PS_SD035L.Size - 1)
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
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
'            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
'            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
'            MenuCreationParams01.uniqueID = "MenuUID"
'            MenuCreationParams01.String = "메뉴명"
'            MenuCreationParams01.Enabled = True
'            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
'        End If
    ElseIf pval.BeforeAction = False Then
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'                Call Sbo_Application.Menus.RemoveEx("MenuUID")
'        End If
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
        If pval.ItemUID = "PS_SD035" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnSrch01" Then
            Call PS_SD035_MTX01 '대상조회
        End If
        
        If pval.ItemUID = "BtnSrch02" Then
            Call PS_SD035_MTX02 '출하리스트 조회
        End If
        
        If pval.ItemUID = "BtnSave" Then '저장
            Call PS_SD035_SaveData
        End If
        
        If pval.ItemUID = "BtnDelete" Then '삭제
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
                Call PS_SD035_DeleteData
                Call PS_SD035_MTX02
            Else
            
            End If
        End If
        
        If pval.ItemUID = "BtnPrint" Then '인쇄
            Call PS_SD035_Print_Report01
        End If
        
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_SD035" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SD035_FormItemEnabled
'                    Call PS_SD035_AddMatrixRow(0, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_SD035_FormItemEnabled
'                End If
'            End If
'        End If

    End If
    Exit Sub
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
        
        ElseIf pval.ItemUID = "CardCode" Then '수주처(대상조회)

            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "")
                
            End If

        ElseIf pval.ItemUID = "SCardCode" Then '수주처(출하리스트)
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "SCardCode", "")
                
            End If
            
        ElseIf pval.ItemUID = "ItemCode" Then '작번(대상조회)

            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "")
                
            End If

        ElseIf pval.ItemUID = "SItemCode" Then '작번(출하리스트)
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "SItemCode", "")
                
            End If
        
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
    
        If pval.ItemChanged = True Then
            
        End If
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
'        If pval.ItemUID = "Mat01" Then
'            If pval.Row > 0 Then
'                Call oMat01.SelectRow(pval.Row, True, False)
'            End If
'        End If
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat01.SelectRow(pval.Row, True, False)
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
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            
            If pval.ColUID = "DocEntry" Then
            
                Set oTempClass = New PS_QM300
                
                Call oTempClass.LoadForm(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE)
            
            End If
            
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Set oTempClass = Nothing
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Set oTempClass = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim PS_PH030() As String

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then

            If (pval.ItemUID = "Mat01") Then
            
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                oMat01.AutoResizeColumns
                
            ElseIf (pval.ItemUID = "CardCode") Then '수주처(대상조회)
            
                oForm01.Items("CardName").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")

            ElseIf (pval.ItemUID = "SCardCode") Then '수주처(출하리스트)
            
                oForm01.Items("SCardName").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
                
            ElseIf (pval.ItemUID = "ItemCode") Then '작번(대상조회)
            
                oForm01.Items("ItemName").Specific.VALUE = MDC_GetData.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")

            ElseIf (pval.ItemUID = "SItemCode") Then '작번(출하리스트)
            
                oForm01.Items("SItemName").Specific.VALUE = MDC_GetData.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
                
            End If

        End If
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_SD035_FormItemEnabled
'        Call PS_SD035_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        Call oMat01.AutoResizeColumns
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_SD035_FormResize
        Call oMat01.AutoResizeColumns
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
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_SD035L.RemoveRecord(oDS_PS_SD035L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_SD035_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_SD035L.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_SD035_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_SD035_CreateItems() As Boolean
On Error GoTo PS_SD035_CreateItems_Error
   
'    Set oDS_PS_SD035H = oForm01.DataSources.DBDataSources("@PS_SD035H")
    Set oDS_PS_SD035L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '//////////대상조회//////////
    '수주일자(시작)
    Call oForm01.DataSources.UserDataSources.Add("ORDR_FrDt", dt_DATE)
    Call oForm01.Items("ORDR_FrDt").Specific.DataBind.SetBound(True, "", "ORDR_FrDt")
    
    '수주일자(종료)
    Call oForm01.DataSources.UserDataSources.Add("ORDR_ToDt", dt_DATE)
    Call oForm01.Items("ORDR_ToDt").Specific.DataBind.SetBound(True, "", "ORDR_ToDt")
    
    '생산완료일자(시작)
    Call oForm01.DataSources.UserDataSources.Add("PP080_FrDt", dt_DATE)
    Call oForm01.Items("PP080_FrDt").Specific.DataBind.SetBound(True, "", "PP080_FrDt")
    
    '생산완료일자(종료)
    Call oForm01.DataSources.UserDataSources.Add("PP080_ToDt", dt_DATE)
    Call oForm01.Items("PP080_ToDt").Specific.DataBind.SetBound(True, "", "PP080_ToDt")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("CardType").Specific.DataBind.SetBound(True, "", "CardType")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItemType").Specific.DataBind.SetBound(True, "", "ItemType")
    
    '출고구분(수기,정상)
    Call oForm01.DataSources.UserDataSources.Add("OutCls", dt_SHORT_TEXT, 10)
    Call oForm01.Items("OutCls").Specific.DataBind.SetBound(True, "", "OutCls")
    
    '수주처
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CardCode").Specific.DataBind.SetBound(True, "", "CardCode")
    
    '수주처명
    Call oForm01.DataSources.UserDataSources.Add("CardName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CardName").Specific.DataBind.SetBound(True, "", "CardName")
    
    '품목코드(작번)
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 50)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    
    '품목명
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    '//////////대상조회//////////
    
    '//////////출하리스트조회//////////
    '출하일자(시작)
    Call oForm01.DataSources.UserDataSources.Add("Out_FrDt", dt_DATE)
    Call oForm01.Items("Out_FrDt").Specific.DataBind.SetBound(True, "", "Out_FrDt")
    
    '출하일자(종료)
    Call oForm01.DataSources.UserDataSources.Add("Out_ToDt", dt_DATE)
    Call oForm01.Items("Out_ToDt").Specific.DataBind.SetBound(True, "", "Out_ToDt")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("SCardType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SCardType").Specific.DataBind.SetBound(True, "", "SCardType")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("SItemType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SItemType").Specific.DataBind.SetBound(True, "", "SItemType")
    
    '출고구분(수기,정상)
    Call oForm01.DataSources.UserDataSources.Add("SOutCls", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SOutCls").Specific.DataBind.SetBound(True, "", "SOutCls")
    
    '수주처
    Call oForm01.DataSources.UserDataSources.Add("SCardCode", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SCardCode").Specific.DataBind.SetBound(True, "", "SCardCode")
    
    '수주처명
    Call oForm01.DataSources.UserDataSources.Add("SCardName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("SCardName").Specific.DataBind.SetBound(True, "", "SCardName")
    
    '품목코드(작번)
    Call oForm01.DataSources.UserDataSources.Add("SItemCode", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SItemCode").Specific.DataBind.SetBound(True, "", "SItemCode")
    
    '품목명
    Call oForm01.DataSources.UserDataSources.Add("SItemName", dt_SHORT_TEXT, 200)
    Call oForm01.Items("SItemName").Specific.DataBind.SetBound(True, "", "SItemName")
    '//////////출하리스트조회//////////
    
    '문서번호, 라인번호 Hidden 처리
'    oMat01.Columns("DocEntry").Visible = False
'    oMat01.Columns("LineId").Visible = False
    
    Exit Function
PS_SD035_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD035_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SD035_ComboBox_Setting()
On Error GoTo PS_SD035_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '거래처구분(대상조회)
    sQry = "            SELECT      U_Minor,"
    sQry = sQry & "                 U_CdName"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'C100'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    
    Call oForm01.Items("CardType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType").Specific, sQry, "", False, False)
    Call oForm01.Items("CardType").Specific.Select(0, psk_Index)
    
    '거래처구분(출하리스트조회)
    Call oForm01.Items("SCardType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SCardType").Specific, sQry, "", False, False)
    Call oForm01.Items("SCardType").Specific.Select(0, psk_Index)
    
    '품목구분(대상조회)
    sQry = "            SELECT      U_Minor,"
    sQry = sQry & "                 U_CdName"
    sQry = sQry & "  FROM       [@PS_SY001L]"
    sQry = sQry & "  WHERE      Code = 'S002'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    
    Call oForm01.Items("ItemType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType").Specific, sQry, "", False, False)
    Call oForm01.Items("ItemType").Specific.Select(0, psk_Index)
    
    '품목구분(출하리스트조회)
    Call oForm01.Items("SItemType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SItemType").Specific, sQry, "", False, False)
    Call oForm01.Items("SItemType").Specific.Select(0, psk_Index)
    
    '출고구분(대상조회)
    Call oForm01.Items("OutCls").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("OutCls").Specific.ValidValues.Add("Y", "정상출고")
    Call oForm01.Items("OutCls").Specific.ValidValues.Add("N", "수기출고")
    Call oForm01.Items("OutCls").Specific.Select(0, psk_Index)
    
    '출고구분(출하리스트조회)
    Call oForm01.Items("SOutCls").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("SOutCls").Specific.ValidValues.Add("Y", "정상출고")
    Call oForm01.Items("SOutCls").Specific.ValidValues.Add("N", "수기출고")
    Call oForm01.Items("SOutCls").Specific.Select(0, psk_Index)
    
    Set oRecordSet01 = Nothing

    Exit Sub
PS_SD035_ComboBox_Setting_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD035_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD035_CF_ChooseFromList()
On Error GoTo PS_SD035_CF_ChooseFromList_Error
    Exit Sub
PS_SD035_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD035_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD035_FormItemEnabled()
On Error GoTo PS_SD035_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
'
        'oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
'        Call PS_SD035_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD035_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD035_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD035_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD035_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_SD035L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD035L.Offset = oRow
    oDS_PS_SD035L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD035_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD035_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD035_FormClear()
On Error GoTo PS_SD035_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD035'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SD035_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD035_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD035_Initial_Setting()
On Error GoTo PS_SD035_Initial_Setting_Error

    ' 사업장
'    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

    '출하일자
    oForm01.Items("Out_FrDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    oForm01.Items("Out_ToDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    Exit Sub
PS_SD035_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD035_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD035_EnableMenus()
On Error GoTo PS_SD035_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_SD035_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD035_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD035_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_SD035_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_SD035_FormItemEnabled
'        Call PS_SD035_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_SD035_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_SD035_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD035_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_SD035_DataValidCheck() As Boolean
On Error GoTo PS_SD035_DataValidCheck_Error
    PS_SD035_DataValidCheck = False
    Dim i As Long
'    If (oForm01.Mode = fm_ADD_MODE) Then
'        Call PS_SD035_FormClear
'    End If
    
    '검사자 미입력시
    If oForm01.Items("ChkWkCd").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("검사자가 입력되지 않았습니다.", bmt_Short, True)
        PS_SD035_DataValidCheck = False
        Exit Function
    End If
    
    '등록일자 미입력시
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("등록일자가 입력되지 않았습니다.", bmt_Short, True)
        PS_SD035_DataValidCheck = False
        Exit Function
    End If
    
    '라인정보 미입력 시
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_SD035_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        
'        '사번
'        If (oMat01.Columns("CntcCode").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "사번은 필수입니다.", bmt_Short, True
'            oMat01.Columns("CntcCode").Cells(i).Click ct_Regular
'            PS_SD035_DataValidCheck = False
'            Exit Function
'        End If
'
'        '검진구분
'        If (oMat01.Columns("ChkCls").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "검진구분은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ChkCls").Cells(i).Click ct_Regular
'            PS_SD035_DataValidCheck = False
'            Exit Function
'        End If
        
    Next
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_SD035L.RemoveRecord(oDS_PS_SD035L.Size - 1)
    Call oMat01.LoadFromDataSource

'    If (oForm01.Mode = fm_ADD_MODE) Then
'        Call PS_SD035_FormClear
'    End If
    
    PS_SD035_DataValidCheck = True
    Exit Function
PS_SD035_DataValidCheck_Error:
    PS_SD035_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_SD035_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD035_MTX01()
On Error GoTo PS_SD035_MTX01_Error

    '//메트릭스에 데이터 로드
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ORDR_FrDt As String '수주일자(시작)
    Dim ORDR_ToDt As String '수주일자(종료)
    Dim PP080_FrDt As String '생산완료일자(시작)
    Dim PP080_ToDt As String '생산완료일자(종료)
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim OutCls As String '출고구분
    Dim CardCode As String '수주처
    Dim ItemCode As String '작번
    
    ORDR_FrDt = Trim(oForm01.Items("ORDR_FrDt").Specific.VALUE)  '수주일자(시작)
    ORDR_ToDt = Trim(oForm01.Items("ORDR_ToDt").Specific.VALUE) '수주일자(종료)
    PP080_FrDt = Trim(oForm01.Items("PP080_FrDt").Specific.VALUE) '생산완료일자(시작)
    PP080_ToDt = Trim(oForm01.Items("PP080_ToDt").Specific.VALUE) '생산완료일자(종료)
    CardType = Trim(oForm01.Items("CardType").Specific.VALUE) '거래처구분
    ItemType = Trim(oForm01.Items("ItemType").Specific.VALUE) '품목구분
    OutCls = Trim(oForm01.Items("OutCls").Specific.VALUE) '출고구분
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE) '수주처
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE) '수주처
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_SD035_01 '"
    Query01 = Query01 & ORDR_FrDt & "','"
    Query01 = Query01 & ORDR_ToDt & "','"
    Query01 = Query01 & PP080_FrDt & "','"
    Query01 = Query01 & PP080_ToDt & "','"
    Query01 = Query01 & CardType & "','"
    Query01 = Query01 & ItemType & "','"
    Query01 = Query01 & OutCls & "','"
    Query01 = Query01 & CardCode & "','"
    Query01 = Query01 & ItemCode & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_SD035_MTX01_Exit
    End If
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_SD035L.InsertRecord (i)
        End If
        oDS_PS_SD035L.Offset = i
        Call oDS_PS_SD035L.setValue("U_LineNum", i, i + 1) '라인번호
        Call oDS_PS_SD035L.setValue("U_ColReg01", i, RecordSet01.Fields("Check").VALUE) '선택
        Call oDS_PS_SD035L.setValue("U_ColReg02", i, RecordSet01.Fields("ItemCode").VALUE) '작번(품목코드)
        Call oDS_PS_SD035L.setValue("U_ColReg03", i, RecordSet01.Fields("ItemName").VALUE) '품명
        Call oDS_PS_SD035L.setValue("U_ColReg04", i, RecordSet01.Fields("CardType").VALUE) '거래처구분
        Call oDS_PS_SD035L.setValue("U_ColReg05", i, RecordSet01.Fields("ItemType").VALUE) '품목구분
        Call oDS_PS_SD035L.setValue("U_ColReg06", i, RecordSet01.Fields("CardCode").VALUE) '수주처코드
        Call oDS_PS_SD035L.setValue("U_ColReg07", i, RecordSet01.Fields("CardName").VALUE) '수주처
        Call oDS_PS_SD035L.setValue("U_ColReg18", i, RecordSet01.Fields("ORDR_LotNo").VALUE) '수주주문번호
        Call oDS_PS_SD035L.setValue("U_ColReg08", i, RecordSet01.Fields("ORDR_Date").VALUE) '수주일자
        Call oDS_PS_SD035L.setValue("U_ColQty01", i, RecordSet01.Fields("ORDR_Qty").VALUE) '수주수량
        Call oDS_PS_SD035L.setValue("U_ColReg10", i, RecordSet01.Fields("PP080_Date").VALUE) '생산완료일자
        Call oDS_PS_SD035L.setValue("U_ColQty02", i, RecordSet01.Fields("PP080_Qty").VALUE) '생산완료수량
        Call oDS_PS_SD035L.setValue("U_ColReg12", i, RecordSet01.Fields("PP080_YN").VALUE) '생산완료여부
        Call oDS_PS_SD035L.setValue("U_ColReg13", i, RecordSet01.Fields("QM4XX_Date").VALUE) '검사일자(최종)
        Call oDS_PS_SD035L.setValue("U_ColQty03", i, RecordSet01.Fields("QM4XX_Qty").VALUE) '검사수량(최종)
        Call oDS_PS_SD035L.setValue("U_ColReg15", i, RecordSet01.Fields("QM4XX_YN").VALUE) '검사여부
        Call oDS_PS_SD035L.setValue("U_ColReg16", i, RecordSet01.Fields("SD040_Date").VALUE) '출하일자(최종)
        Call oDS_PS_SD035L.setValue("U_ColQty04", i, RecordSet01.Fields("SD040_Qty").VALUE) '출하수량(최종)
        Call oDS_PS_SD035L.setValue("U_ColDt01", i, RecordSet01.Fields("Out_Date").VALUE) '출하일자(신규)
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD035_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_SD035_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD035_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD035_MTX02()
On Error GoTo PS_SD035_MTX02_Error

    '//메트릭스에 데이터 로드
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Out_FrDt As String '출하일자(시작)
    Dim Out_ToDt As String '출하일자(종료)
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim OutCls As String '출고구분
    Dim CardCode As String '수주처
    Dim ItemCode As String '작번
    
    Out_FrDt = Trim(oForm01.Items("Out_FrDt").Specific.VALUE) '출하일자(시작)
    Out_ToDt = Trim(oForm01.Items("Out_ToDt").Specific.VALUE) '출하일자(종료)
    CardType = Trim(oForm01.Items("SCardType").Specific.VALUE) '거래처구분
    ItemType = Trim(oForm01.Items("SItemType").Specific.VALUE) '품목구분
    OutCls = Trim(oForm01.Items("SOutCls").Specific.VALUE) '출고구분
    CardCode = Trim(oForm01.Items("SCardCode").Specific.VALUE) '수주처
    ItemCode = Trim(oForm01.Items("SItemCode").Specific.VALUE) '작번
    
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    Query01 = "              EXEC PS_SD035_03 '"
    Query01 = Query01 & Out_FrDt & "','"
    Query01 = Query01 & Out_ToDt & "','"
    Query01 = Query01 & CardType & "','"
    Query01 = Query01 & ItemType & "','"
    Query01 = Query01 & OutCls & "','"
    Query01 = Query01 & CardCode & "','"
    Query01 = Query01 & ItemCode & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_SD035_MTX02_Exit
    End If
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_SD035L.InsertRecord (i)
        End If
        oDS_PS_SD035L.Offset = i
        Call oDS_PS_SD035L.setValue("U_LineNum", i, i + 1) '라인번호
        Call oDS_PS_SD035L.setValue("U_ColReg01", i, RecordSet01.Fields("Check").VALUE) '선택
        Call oDS_PS_SD035L.setValue("U_ColReg02", i, RecordSet01.Fields("ItemCode").VALUE) '작번(품목코드)
        Call oDS_PS_SD035L.setValue("U_ColReg03", i, RecordSet01.Fields("ItemName").VALUE) '품명
        Call oDS_PS_SD035L.setValue("U_ColReg04", i, RecordSet01.Fields("CardType").VALUE) '거래처구분
        Call oDS_PS_SD035L.setValue("U_ColReg05", i, RecordSet01.Fields("ItemType").VALUE) '품목구분
        Call oDS_PS_SD035L.setValue("U_ColReg06", i, RecordSet01.Fields("CardCode").VALUE) '수주처코드
        Call oDS_PS_SD035L.setValue("U_ColReg07", i, RecordSet01.Fields("CardName").VALUE) '수주처
        Call oDS_PS_SD035L.setValue("U_ColReg18", i, RecordSet01.Fields("ORDR_LotNo").VALUE) '수주주문번호
        Call oDS_PS_SD035L.setValue("U_ColReg08", i, RecordSet01.Fields("ORDR_Date").VALUE) '수주일자
        Call oDS_PS_SD035L.setValue("U_ColQty01", i, RecordSet01.Fields("ORDR_Qty").VALUE) '수주수량
        Call oDS_PS_SD035L.setValue("U_ColReg10", i, RecordSet01.Fields("PP080_Date").VALUE) '생산완료일자
        Call oDS_PS_SD035L.setValue("U_ColQty02", i, RecordSet01.Fields("PP080_Qty").VALUE) '생산완료수량
        Call oDS_PS_SD035L.setValue("U_ColReg12", i, RecordSet01.Fields("PP080_YN").VALUE) '생산완료여부
        Call oDS_PS_SD035L.setValue("U_ColReg13", i, RecordSet01.Fields("QM4XX_Date").VALUE) '검사일자(최종)
        Call oDS_PS_SD035L.setValue("U_ColQty03", i, RecordSet01.Fields("QM4XX_Qty").VALUE) '검사수량(최종)
        Call oDS_PS_SD035L.setValue("U_ColReg15", i, RecordSet01.Fields("QM4XX_YN").VALUE) '검사여부
        Call oDS_PS_SD035L.setValue("U_ColReg16", i, RecordSet01.Fields("SD040_Date").VALUE) '출하일자(최종)
        Call oDS_PS_SD035L.setValue("U_ColQty04", i, RecordSet01.Fields("SD040_Qty").VALUE) '출하수량(최종)
        Call oDS_PS_SD035L.setValue("U_ColDt01", i, Replace(RecordSet01.Fields("Out_Date").VALUE, ".", "")) '출하일자(신규)
        Call oDS_PS_SD035L.setValue("U_ColReg19", i, RecordSet01.Fields("DocEntry").VALUE) '문서번호(Key)
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD035_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_SD035_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD035_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_SD035_SaveData()
'******************************************************************************
'Function ID : PS_SD035_SaveData()
'해당모듈 : PS_SD035
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD035_SaveData_Error

    Dim loopCount As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    Dim ItemCode As String
    Dim ItemName As String
    Dim ItemType As String
    Dim CardType As String
    Dim CardCode As String
    Dim CardName As String
    Dim ORDR_LotNo As String
    Dim ORDR_Date As String
    Dim ORDR_Qty As Double
    Dim PP080_Date As String
    Dim PP080_Qty As Double
    Dim PP080_YN As String
    Dim QM4XX_Date As String
    Dim QM4XX_Qty As Double
    Dim QM4XX_YN As String
    Dim SD040_Date As String
    Dim SD040_Qty As Double
    Dim Out_Date As String

    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
        If Trim(oDS_PS_SD035L.GetValue("U_ColReg01", loopCount)) = "Y" Then
    
'            Call Sbo_Application.MessageBox(Trim(oDS_PS_SD035L.GetValue("U_ColReg02", loopCount)))
    
            ItemCode = Trim(oDS_PS_SD035L.GetValue("U_ColReg02", loopCount)) '품목코드(작번)
            ItemName = Trim(oDS_PS_SD035L.GetValue("U_ColReg03", loopCount)) '품명
            ItemType = Trim(oDS_PS_SD035L.GetValue("U_ColReg05", loopCount)) '품목구분
            CardType = Trim(oDS_PS_SD035L.GetValue("U_ColReg04", loopCount)) '거래처구분
            CardCode = Trim(oDS_PS_SD035L.GetValue("U_ColReg06", loopCount)) '수주처
            CardName = Trim(oDS_PS_SD035L.GetValue("U_ColReg07", loopCount)) '수주처명
            ORDR_LotNo = Trim(oDS_PS_SD035L.GetValue("U_ColReg18", loopCount)) '수주주문번호
            ORDR_Date = Trim(oDS_PS_SD035L.GetValue("U_ColReg08", loopCount)) '수주일자
            ORDR_Qty = CDbl(Trim(oDS_PS_SD035L.GetValue("U_ColQty01", loopCount))) '수주수량
            PP080_Date = Trim(oDS_PS_SD035L.GetValue("U_ColReg10", loopCount)) '생산완료일자(최종)
            PP080_Qty = CDbl(Trim(oDS_PS_SD035L.GetValue("U_ColQty02", loopCount))) '생산완료수량(누계)
            PP080_YN = Trim(oDS_PS_SD035L.GetValue("U_ColReg12", loopCount)) '생산완료여부
            QM4XX_Date = Trim(oDS_PS_SD035L.GetValue("U_ColReg13", loopCount)) '검사일자(최종)
            QM4XX_Qty = CDbl(Trim(oDS_PS_SD035L.GetValue("U_ColQty03", loopCount))) '검사수량(누계)
            QM4XX_YN = Trim(oDS_PS_SD035L.GetValue("U_ColReg15", loopCount)) '검사여부
            SD040_Date = Trim(oDS_PS_SD035L.GetValue("U_ColReg16", loopCount)) '출고일자(최종)
            SD040_Qty = CDbl(Trim(oDS_PS_SD035L.GetValue("U_ColQty04", loopCount))) '출고수량(누계)
            Out_Date = Left(Trim(oDS_PS_SD035L.GetValue("U_ColDt01", loopCount)), 4) & "." & Right(Left(Trim(oDS_PS_SD035L.GetValue("U_ColDt01", loopCount)), 6), 2) & "." & Right(Trim(oDS_PS_SD035L.GetValue("U_ColDt01", loopCount)), 2) '출고일자(신규)
            DocEntry = Trim(oDS_PS_SD035L.GetValue("U_ColReg19", loopCount)) '관리번호
            
            If DocEntry = "" Then
            
                'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
                sQry = "SELECT ISNULL(MAX(CONVERT(INT,DocEntry)), 0) FROM [Z_PS_SD035_01]"
                Call RecordSet01.DoQuery(sQry)
        
                If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
                    DocEntry = 1
                Else
                    DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
                End If
                
                sQry = "                EXEC [PS_SD035_02] "
                sQry = sQry & "'" & DocEntry & "'," '관리번호
                sQry = sQry & "'" & ItemCode & "',"  '품목코드
                sQry = sQry & "'" & ItemName & "',"  '품명
                sQry = sQry & "'" & ItemType & "',"  '품목구분
                sQry = sQry & "'" & CardType & "',"  '거래처구분
                sQry = sQry & "'" & CardCode & "',"  '수주처코드
                sQry = sQry & "'" & CardName & "',"  '수주처명
                sQry = sQry & "'" & ORDR_LotNo & "',"  '수주주문번호
                sQry = sQry & "'" & ORDR_Date & "',"  '수주일자
                sQry = sQry & "'" & ORDR_Qty & "',"  '수주수량
                sQry = sQry & "'" & PP080_Date & "',"  '생산완료일자(최종)
                sQry = sQry & "'" & PP080_Qty & "',"  '생산완료수량(누계)
                sQry = sQry & "'" & PP080_YN & "',"  '생산완료여부
                sQry = sQry & "'" & QM4XX_Date & "',"  '검사일자(최종)
                sQry = sQry & "'" & QM4XX_Qty & "',"  '검사수량(누계)
                sQry = sQry & "'" & QM4XX_YN & "',"  '검사여부
                sQry = sQry & "'" & SD040_Date & "',"  '출고일자(최종)
                sQry = sQry & "'" & SD040_Qty & "',"  '출고일자(누계)
                sQry = sQry & "'" & Out_Date & "'"  '출고일자
                
            Else
            
                sQry = "                EXEC [PS_SD035_04]"
                sQry = sQry & "'" & DocEntry & "'," '관리번호
                sQry = sQry & "'" & Out_Date & "'" '출고일자
                
            End If
            
            Call RecordSet02.DoQuery(sQry)

        End If

    Next

    Call MDC_Com.MDC_GF_Message("등록 완료!", "S")
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD035_SaveData_Error:
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PS_SD035_SaveData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_SD035_DeleteData()
'******************************************************************************
'Function ID : PS_SD035_DeleteData()
'해당모듈 : PS_SD035
'기능 : 데이터 DELETE
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD035_DeleteData_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim DocEntry As String
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
        If Trim(oDS_PS_SD035L.GetValue("U_ColReg01", loopCount)) = "Y" Then
    
            DocEntry = Trim(oDS_PS_SD035L.GetValue("U_ColReg19", loopCount)) '관리번호
            
            If DocEntry = "" Then
            
                ErrNum = "1"
                GoTo PS_SD035_DeleteData_Error
                
            Else
            
                sQry = "                EXEC [PS_SD035_05]"
                sQry = sQry & "'" & DocEntry & "'" '관리번호
                
            End If
            
            Call RecordSet01.DoQuery(sQry)

        End If

    Next

    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    Set RecordSet01 = Nothing
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD035_DeleteData_Error:
    Set RecordSet01 = Nothing

    If ErrNum = "1" Then
        Call MDC_Com.MDC_GF_Message("관리번호가 존재하지 않아 삭제할 수 없습니다. 출하리스트 조회를 이용하십시오.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_SD035_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Private Sub PS_SD035_FormResize()
On Error GoTo PS_SD035_FormResize_Error
    
    Exit Sub
PS_SD035_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD035_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD035_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_SD035_Validate_Error
    PS_SD035_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "수정" Then
        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
    ElseIf ValidateType = "취소" Then
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_SD035_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_SD035_Validate_Error:
    PS_SD035_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_SD035_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PS_SD035_Print_Report01()
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
            
    Call ConnectODBC
    
    Dim Out_FrDt As String '출하일자(시작)
    Dim Out_ToDt As String '출하일자(종료)
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim OutCls As String '출고구분
    Dim CardCode As String '수주처
    
    Out_FrDt = Trim(oForm01.Items("Out_FrDt").Specific.VALUE) '출하일자(시작)
    Out_ToDt = Trim(oForm01.Items("Out_ToDt").Specific.VALUE) '출하일자(종료)
    CardType = Trim(oForm01.Items("SCardType").Specific.VALUE) '거래처구분
    ItemType = Trim(oForm01.Items("SItemType").Specific.VALUE) '품목구분
    OutCls = Trim(oForm01.Items("SOutCls").Specific.VALUE) '출고구분
    CardCode = Trim(oForm01.Items("SCardCode").Specific.VALUE) '수주처
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", 100, False)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_SD035] 출하리스트"
    ReportName = "PS_SD035_01.rpt"
 
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
    
    sQry = "         EXEC PS_SD035_06 '"
    sQry = sQry & Out_FrDt & "','"
    sQry = sQry & Out_ToDt & "','"
    sQry = sQry & CardType & "','"
    sQry = sQry & ItemType & "','"
    sQry = sQry & OutCls & "','"
    sQry = sQry & CardCode & "'"
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        GoTo Print_Query_Error:
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_SD035_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub






