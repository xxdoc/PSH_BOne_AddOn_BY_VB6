VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP552"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_PP552.cls
'//  Module : 생산관리 > 애드온 > 공정관리
'//  Desc : 공정상태 관리
'//  FormType : PS_PP552
'//  Create Date(Start) : 2015.03.30
'//  Create Date(End) :
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************

'//사용방법
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_PP552
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix
Public oMat03 As SAPbouiCOM.Matrix

Private oDS_PS_PP552L As SAPbouiCOM.DBDataSource '등록라인
Private oDS_PS_PP552M As SAPbouiCOM.DBDataSource '등록라인
Private oDS_PS_PP552N As SAPbouiCOM.DBDataSource '등록라인

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String
Public oBaseItmBsort01 As String

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String, Optional ByVal oItmBsort02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP552.srf")
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
    
    oFormUniqueID01 = "PS_PP552_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Set oBaseForm01 = oForm02
    oBaseItemUID01 = oItemUID02
    oBaseColUID01 = oColUID02
    oBaseColRow01 = oColRow02
    oBaseTradeType01 = oTradeType02
    oBaseItmBsort01 = oItmBsort02
    
    Call PS_PP552_CreateItems
    Call PS_PP552_ComboBox_Setting
    Call PS_PP552_CF_ChooseFromList
    Call PS_PP552_FormItemEnabled
    Call PS_PP552_EnableMenus
    '//Call PS_PP552_FormClear '//UDO방식일때
    '//Call PS_PP552_AddMatrixRow(0, True) '//UDO방식일때
    
    oForm01.Items("FrDt01").Specific.VALUE = Format(DateAdd("m", -2, Date), "YYYYMM01")
    oForm01.Items("ToDt01").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oForm01.Items("FrDt02").Specific.VALUE = Format(DateAdd("m", -2, Date), "YYYYMM01")
    oForm01.Items("ToDt02").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oForm01.Items("FrDt03").Specific.VALUE = Format(DateAdd("m", -2, Date), "YYYYMM01")
    oForm01.Items("ToDt03").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oForm01.Items("CpStatus03").Enabled = False '계획공수 등록 시는 공정상태가 대기인 자료만 조회하기 위한 콤보 선택 불가 조치
    
    Call oForm01.Items("Folder01").Click
    
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
                '//Call PS_PP552_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_PP552_FormItemEnabled '//UDO방식
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
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Then
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
            
                If PS_PP552_CheckBeforeSearch() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                Call PS_PP552_MTX01
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        ElseIf pval.ItemUID = "BtnDel01" Then
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
            
                If PS_PP552_DeleteData01() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                Call PS_PP552_MTX01
                
            Else
            
            End If
            
        ElseIf pval.ItemUID = "BtnSrch02" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_PP552_CheckBeforeSearch() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                Call PS_PP552_MTX02
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        ElseIf pval.ItemUID = "BtnSave02" Then
            
            If PS_PP552_SaveData02() = False Then
                BubbleEvent = False
                Exit Sub
            End If
        
            Call PS_PP552_MTX02
        
        ElseIf pval.ItemUID = "BtnSrch03" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_PP552_CheckBeforeSearch() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                Call PS_PP552_MTX03
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        ElseIf pval.ItemUID = "BtnSave03" Then
            
            If PS_PP552_SaveData03() = False Then
                BubbleEvent = False
                Exit Sub
            End If
        
            Call PS_PP552_MTX03
            
        ElseIf pval.ItemUID = "Mat01" And pval.ColUID = "Check" And pval.Row > 0 Then
        
            If oMat01.RowCount >= pval.Row Then     '빈 Select 필드에 클릭했을 때 생기는 오류 수정을 위한 구문
                
                If PS_PP552_CheckAfterStatus(pval.Row) = True Then

                    Call MDC_Com.MDC_GF_Message("선택한 작번은 이후 공정상태가 존재합니다. 선택할 수 없습니다.", "E")
                    
'                    Call oForm01.Freeze(True)
'                    Call oDS_PS_PP552L.setValue("U_ColReg01", pval.Row - 1, "N")
'                    Call oMat01.LoadFromDataSource
'                    Call oForm01.Freeze(False)
                    
                    oMat01.Columns("Check").Cells(pval.Row).Specific.Checked = False

                    Exit Sub

                End If
            
            End If
        
        End If
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_PP552_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If

        
        
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP552" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
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
        
        'Folder03가 선택되었을 때
        If pval.ItemUID = "Folder03" Then
            
            oForm01.PaneLevel = 3
            oForm01.DefButton = "BtnSrch03"
            
        End If
        '폴더를 사용할 때는 필수 소스_E
        
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP552_FormItemEnabled
'                    Call PS_PP552_FormClear '//UDO방식일때
'                    Call PS_PP552_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP552_FormItemEnabled
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
        
'            If pval.ColUID = "ItemCode" Then
'
'                Dim OrdGbn As String
'                Dim InputGbn As String
'
'                OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.Selected.VALUE)
'                InputGbn = oMat01.Columns("InputGbn").Cells(pval.Row).Specific.Selected.VALUE
'
'                Dim ChildForm01 As Variant
'
'                Set ChildForm01 = New PS_SM021
'                Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, OrdGbn, InputGbn, Trim(oForm01.Items("BPLID").Specific.VALUE))
'                BubbleEvent = False
'
'            ElseIf pval.ColUID = "CntcCode" Then
'
'                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "CntcCode") '사번 포맷서치설정
'
'            End If
        
        Else
    
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CpCode01", "") '공정
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode01", "") '등록자
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum01", "") '작번
            
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CpCode02", "") '공정
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode02", "") '등록자
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum02", "") '작번
            
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CpCode03", "") '공정
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode03", "") '등록자
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum03", "") '작번
        
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
    
        If pval.ItemUID = "Opt01" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Grid01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            Call oForm01.Freeze(False)
        End If
        If pval.ItemUID = "Opt02" Then
            Call oForm01.Freeze(True)
            oForm01.Settings.MatrixUID = "Mat01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            Call oForm01.Freeze(False)
        End If
    
        If pval.ItemUID = "Mat01" Then

            If pval.Row > 0 Then

                Call oMat01.SelectRow(pval.Row, True, False)
                
            End If
        
        End If
        
        If pval.ItemUID = "Mat02" Then

            If pval.Row > 0 Then

                Call oMat02.SelectRow(pval.Row, True, False)
                
            End If
        
        End If
        
        If pval.ItemUID = "Mat03" Then

            If pval.Row > 0 Then

                Call oMat03.SelectRow(pval.Row, True, False)
                
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
    
        If pval.ItemUID = "Grid01" Then
            If pval.Row = -1 Then
'                oGrid01.Columns(pval.ColUID).TitleObject.Sortable = True
                
            Else
'                If oGrid01.Rows.SelectedRows.Count > 0 Then
'                    Call PS_PP552_SetBaseForm '//부모폼에입력
'                    If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
'                        Call oForm01.Close
'                    End If
'                Else
'                    BubbleEvent = False
'                End If
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
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "Mat01") Then
            
                Call PS_PP552_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
        
            Else
        
                Call PS_PP552_FlushToItemValue(pval.ItemUID)
            
            End If
                
        End If
       
    ElseIf pval.BeforeAction = False Then

    End If
    
    Call oForm01.Freeze(False)
    
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_PP552_FormItemEnabled
        Call PS_PP552_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
    
        Call PS_PP552_FormResize
        
'        oForm01.Items("Mat01").Top = 70
'        oForm01.Items("Mat01").Height = (oForm01.Height / 2) - 70
'        oForm01.Items("Mat01").Left = 7
'        oForm01.Items("Mat01").Width = oForm01.Width - 21
'        oForm01.Items("Mat02").Top = (oForm01.Height / 2) + 10
'        oForm01.Items("Mat02").Height = (oForm01.Height / 2) - 75
'        oForm01.Items("Mat02").Left = 7
'        oForm01.Items("Mat02").Width = oForm01.Width - 21
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    Dim oDataTable01  As SAPbouiCOM.DataTable
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        'If (pval.ItemUID = "ItemCode") Then
         '   Set oDataTable01 = pval.SelectedObjects
        '    If oDataTable01 Is Nothing Then
        '    Else
              '  oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
         '     '  oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
         '   End If
       ' End If
        oForm01.Update
    End If
    Set oDataTable01 = Nothing
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
        Set oMat02 = Nothing
        Set oMat03 = Nothing
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
    '        For i = 1 To oMat01.VisualRowCount
    '            oMat01.Columns("COL01").Cells(i).Specific.Value = i
    '        Next i
    '        oMat01.FlushToDataSource
    '        Call oDS_PS_PP552L.RemoveRecord(oDS_PS_PP552L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_PP552_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_PP552_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP552_CreateItems() As Boolean
On Error GoTo PS_PP552_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    'Dim C_Date   As Date
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP552L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_PP552M = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    Set oDS_PS_PP552N = oForm01.DataSources.DBDataSources("@PS_USERDS03")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
    
    Set oMat03 = oForm01.Items("Mat03").Specific
    oMat03.SelectionMode = ms_NotSupported
    oMat03.AutoResizeColumns
    
    '//////////공정상태 삭제//////////_S
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID01", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID01").Specific.DataBind.SetBound(True, "", "BPLID01")
    
    '기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt01", dt_DATE)
    Call oForm01.Items("FrDt01").Specific.DataBind.SetBound(True, "", "FrDt01")
    
    '기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt01", dt_DATE)
    Call oForm01.Items("ToDt01").Specific.DataBind.SetBound(True, "", "ToDt01")
    
    '공정코드
    Call oForm01.DataSources.UserDataSources.Add("CpCode01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CpCode01").Specific.DataBind.SetBound(True, "", "CpCode01")
    
    '공정명
    Call oForm01.DataSources.UserDataSources.Add("CpName01", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CpName01").Specific.DataBind.SetBound(True, "", "CpName01")
    
    '등록자사번
    Call oForm01.DataSources.UserDataSources.Add("CntcCode01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcCode01").Specific.DataBind.SetBound(True, "", "CntcCode01")
    
    '등록자성명
    Call oForm01.DataSources.UserDataSources.Add("CntcName01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcName01").Specific.DataBind.SetBound(True, "", "CntcName01")
    
    '작업구분
    Call oForm01.DataSources.UserDataSources.Add("WorkGbn01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("WorkGbn01").Specific.DataBind.SetBound(True, "", "WorkGbn01")

    '공정상태
    Call oForm01.DataSources.UserDataSources.Add("CpStatus01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CpStatus01").Specific.DataBind.SetBound(True, "", "CpStatus01")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardType01").Specific.DataBind.SetBound(True, "", "CardType01")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemType01").Specific.DataBind.SetBound(True, "", "ItemType01")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum01", dt_SHORT_TEXT, 11)
    Call oForm01.Items("OrdNum01").Specific.DataBind.SetBound(True, "", "OrdNum01")
    
    '서브작번1
    Call oForm01.DataSources.UserDataSources.Add("OrdSub101", dt_SHORT_TEXT, 2)
    Call oForm01.Items("OrdSub101").Specific.DataBind.SetBound(True, "", "OrdSub101")
    
    '서브작번2
    Call oForm01.DataSources.UserDataSources.Add("OrdSub201", dt_SHORT_TEXT, 3)
    Call oForm01.Items("OrdSub201").Specific.DataBind.SetBound(True, "", "OrdSub201")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("ItemName01", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemName01").Specific.DataBind.SetBound(True, "", "ItemName01")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("ItemSpec01", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemSpec01").Specific.DataBind.SetBound(True, "", "ItemSpec01")
    '//////////공정상태 삭제//////////_E
    
    '//////////공정상태 비고 등록//////////_S
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID02", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID02").Specific.DataBind.SetBound(True, "", "BPLID02")
    
    '기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt02", dt_DATE)
    Call oForm01.Items("FrDt02").Specific.DataBind.SetBound(True, "", "FrDt02")
    
    '기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt02", dt_DATE)
    Call oForm01.Items("ToDt02").Specific.DataBind.SetBound(True, "", "ToDt02")
    
    '공정코드
    Call oForm01.DataSources.UserDataSources.Add("CpCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CpCode02").Specific.DataBind.SetBound(True, "", "CpCode02")
    
    '공정명
    Call oForm01.DataSources.UserDataSources.Add("CpName02", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CpName02").Specific.DataBind.SetBound(True, "", "CpName02")
    
    '등록자사번
    Call oForm01.DataSources.UserDataSources.Add("CntcCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcCode02").Specific.DataBind.SetBound(True, "", "CntcCode02")
    
    '등록자성명
    Call oForm01.DataSources.UserDataSources.Add("CntcName02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcName02").Specific.DataBind.SetBound(True, "", "CntcName02")
    
    '작업구분
    Call oForm01.DataSources.UserDataSources.Add("WorkGbn02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("WorkGbn02").Specific.DataBind.SetBound(True, "", "WorkGbn02")

    '공정상태
    Call oForm01.DataSources.UserDataSources.Add("CpStatus02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CpStatus02").Specific.DataBind.SetBound(True, "", "CpStatus02")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardType02").Specific.DataBind.SetBound(True, "", "CardType02")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemType02").Specific.DataBind.SetBound(True, "", "ItemType02")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum02", dt_SHORT_TEXT, 11)
    Call oForm01.Items("OrdNum02").Specific.DataBind.SetBound(True, "", "OrdNum02")
    
    '서브작번1
    Call oForm01.DataSources.UserDataSources.Add("OrdSub102", dt_SHORT_TEXT, 2)
    Call oForm01.Items("OrdSub102").Specific.DataBind.SetBound(True, "", "OrdSub102")
    
    '서브작번2
    Call oForm01.DataSources.UserDataSources.Add("OrdSub202", dt_SHORT_TEXT, 3)
    Call oForm01.Items("OrdSub202").Specific.DataBind.SetBound(True, "", "OrdSub202")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("ItemName02", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemName02").Specific.DataBind.SetBound(True, "", "ItemName02")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("ItemSpec02", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemSpec02").Specific.DataBind.SetBound(True, "", "ItemSpec02")
    '//////////공정상태 비고 등록//////////_E
    
    '//////////공정상태 계획공수 등록//////////_S
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID03", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID03").Specific.DataBind.SetBound(True, "", "BPLID03")
    
    '기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt03", dt_DATE)
    Call oForm01.Items("FrDt03").Specific.DataBind.SetBound(True, "", "FrDt03")
    
    '기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt03", dt_DATE)
    Call oForm01.Items("ToDt03").Specific.DataBind.SetBound(True, "", "ToDt03")
    
    '공정코드
    Call oForm01.DataSources.UserDataSources.Add("CpCode03", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CpCode03").Specific.DataBind.SetBound(True, "", "CpCode03")
    
    '공정명
    Call oForm01.DataSources.UserDataSources.Add("CpName03", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CpName03").Specific.DataBind.SetBound(True, "", "CpName03")
    
    '등록자사번
    Call oForm01.DataSources.UserDataSources.Add("CntcCode03", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcCode03").Specific.DataBind.SetBound(True, "", "CntcCode03")
    
    '등록자성명
    Call oForm01.DataSources.UserDataSources.Add("CntcName03", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcName03").Specific.DataBind.SetBound(True, "", "CntcName03")
    
    '작업구분
    Call oForm01.DataSources.UserDataSources.Add("WorkGbn03", dt_SHORT_TEXT, 20)
    Call oForm01.Items("WorkGbn03").Specific.DataBind.SetBound(True, "", "WorkGbn03")

    '공정상태
    Call oForm01.DataSources.UserDataSources.Add("CpStatus03", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CpStatus03").Specific.DataBind.SetBound(True, "", "CpStatus03")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType03", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardType03").Specific.DataBind.SetBound(True, "", "CardType03")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType03", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemType03").Specific.DataBind.SetBound(True, "", "ItemType03")
    
    '작번
    Call oForm01.DataSources.UserDataSources.Add("OrdNum03", dt_SHORT_TEXT, 11)
    Call oForm01.Items("OrdNum03").Specific.DataBind.SetBound(True, "", "OrdNum03")
    
    '서브작번1
    Call oForm01.DataSources.UserDataSources.Add("OrdSub103", dt_SHORT_TEXT, 2)
    Call oForm01.Items("OrdSub103").Specific.DataBind.SetBound(True, "", "OrdSub103")
    
    '서브작번2
    Call oForm01.DataSources.UserDataSources.Add("OrdSub203", dt_SHORT_TEXT, 3)
    Call oForm01.Items("OrdSub203").Specific.DataBind.SetBound(True, "", "OrdSub203")
    
    '품명
    Call oForm01.DataSources.UserDataSources.Add("ItemName03", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemName03").Specific.DataBind.SetBound(True, "", "ItemName03")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("ItemSpec03", dt_SHORT_TEXT, 200)
    Call oForm01.Items("ItemSpec03").Specific.DataBind.SetBound(True, "", "ItemSpec03")
    '//////////공정상태 계획공수 등록//////////_E
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP552_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP552_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP552_ComboBox_Setting()
On Error GoTo PS_PP552_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    
    BPLID = MDC_PS_Common.User_BPLId
    
    '//////////공정상태 삭제//////////_S
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID01").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, False, False)
    
    '작업구분
    sQry = "           SELECT     Code AS [Code], "
    sQry = sQry & "               Name AS [Name]"
    sQry = sQry & " FROM      [@PSH_ITMBSORT]"
    sQry = sQry & " WHERE     U_PudYN = 'Y'"
    Call oForm01.Items("WorkGbn01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("WorkGbn01").Specific, sQry, "%", False, False)
    
    '공정상태
    Call oForm01.Items("CpStatus01").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("CpStatus01").Specific.ValidValues.Add("10", "대기")
    Call oForm01.Items("CpStatus01").Specific.ValidValues.Add("20", "시작")
    Call oForm01.Items("CpStatus01").Specific.ValidValues.Add("30", "완료")
    Call oForm01.Items("CpStatus01").Specific.Select(0, psk_Index)
    
    '거래처구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'C100'"
    Call oForm01.Items("CardType01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType01").Specific, sQry, "%", False, False)
    
    '품목구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'S002'"
    Call oForm01.Items("ItemType01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType01").Specific, sQry, "%", False, False)
    
    '//////////매트릭스//////////
    '공정상태
    Call oMat01.Columns("CpStatus").ValidValues.Add("10", "대기")
    Call oMat01.Columns("CpStatus").ValidValues.Add("20", "시작")
    Call oMat01.Columns("CpStatus").ValidValues.Add("30", "완료")
    '//////////매트릭스//////////
    '//////////공정상태 삭제//////////_E
    
    '//////////공정상태 비고 등록//////////_S
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID02").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, False, False)
    
    '작업구분
    sQry = "           SELECT     Code AS [Code], "
    sQry = sQry & "               Name AS [Name]"
    sQry = sQry & " FROM      [@PSH_ITMBSORT]"
    sQry = sQry & " WHERE     U_PudYN = 'Y'"
    Call oForm01.Items("WorkGbn02").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("WorkGbn02").Specific, sQry, "%", False, False)
    
    '공정상태
    Call oForm01.Items("CpStatus02").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("CpStatus02").Specific.ValidValues.Add("10", "대기")
    Call oForm01.Items("CpStatus02").Specific.ValidValues.Add("20", "시작")
    Call oForm01.Items("CpStatus02").Specific.ValidValues.Add("30", "완료")
    Call oForm01.Items("CpStatus02").Specific.Select(0, psk_Index)
    
    '거래처구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'C100'"
    Call oForm01.Items("CardType02").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType02").Specific, sQry, "%", False, False)
    
    '품목구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'S002'"
    Call oForm01.Items("ItemType02").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType02").Specific, sQry, "%", False, False)
    
    '//////////매트릭스//////////
    '공정상태
    Call oMat02.Columns("CpStatus").ValidValues.Add("10", "대기")
    Call oMat02.Columns("CpStatus").ValidValues.Add("20", "시작")
    Call oMat02.Columns("CpStatus").ValidValues.Add("30", "완료")
    '//////////매트릭스//////////
    '//////////공정상태 비고 등록//////////_E

    '//////////공정상태 계획공수 등록//////////_S
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID03").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, False, False)
    
    '작업구분
    sQry = "           SELECT     Code AS [Code], "
    sQry = sQry & "               Name AS [Name]"
    sQry = sQry & " FROM      [@PSH_ITMBSORT]"
    sQry = sQry & " WHERE     U_PudYN = 'Y'"
    Call oForm01.Items("WorkGbn03").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("WorkGbn03").Specific, sQry, "%", False, False)
    
    '공정상태
    Call oForm01.Items("CpStatus03").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("CpStatus03").Specific.ValidValues.Add("10", "대기")
    Call oForm01.Items("CpStatus03").Specific.ValidValues.Add("20", "시작")
    Call oForm01.Items("CpStatus03").Specific.ValidValues.Add("30", "완료")
    Call oForm01.Items("CpStatus03").Specific.Select(1, psk_Index)
    
    '거래처구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'C100'"
    Call oForm01.Items("CardType03").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType03").Specific, sQry, "%", False, False)
    
    '품목구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'S002'"
    Call oForm01.Items("ItemType03").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType03").Specific, sQry, "%", False, False)
    
    '//////////매트릭스//////////
    '공정상태
    Call oMat03.Columns("CpStatus").ValidValues.Add("10", "대기")
    Call oMat03.Columns("CpStatus").ValidValues.Add("20", "시작")
    Call oMat03.Columns("CpStatus").ValidValues.Add("30", "완료")
    '//////////매트릭스//////////
    '//////////공정상태 계획공수 등록//////////_E

    Call oForm01.Freeze(False)
    Exit Sub
PS_PP552_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP552_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP552_CF_ChooseFromList()
On Error GoTo PS_PP552_CF_ChooseFromList_Error
    '//ChooseFromList 설정
'    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
'    Dim oCons               As SAPbouiCOM.Conditions
'    Dim oCon                As SAPbouiCOM.Condition
'    Dim oCFL                As SAPbouiCOM.ChooseFromList
'    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
'    Dim oEdit               As SAPbouiCOM.EditText
'    Dim oColumn             As SAPbouiCOM.Column
'
'    Set oEdit = oForm01.Items("ItemCode").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = "4"
'    oCFLCreationParams.uniqueID = "CFLITEMCD"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)
'
''    Set oCons = oCFL.GetConditions()
''    Set oCon = oCons.Add()
''    oCon.Alias = "CardType"
''    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
''    oCon.CondVal = "C"
''    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLITEMCD"
'    oEdit.ChooseFromListAlias = "ItemCode"
    Exit Sub
PS_PP552_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP552_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP552_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_PP552_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim sQry As String
    
    Dim OrdNum As String
    Dim OrdSub1 As String
    Dim OrdSub2 As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Select Case oUID
'
'        Case "Mat01"
'
'            If oCol = "RWorkTime" Then
'
''                Call oMat01.FlushToDataSource
''
''                RWorkTime = oDS_PS_PP552L.GetValue("U_ColQty01", oRow - 1)
''                Price = oDS_PS_PP552L.GetValue("U_ColSum01", oRow - 1)
''                Amount = RWorkTime * Price
''                Call oDS_PS_PP552L.setValue("U_ColSum02", oRow - 1, Amount)
''
''                For loopCount = 0 To oMat01.VisualRowCount - 1
''
''                    totalAmt = totalAmt + oDS_PS_PP552L.GetValue("U_ColSum02", loopCount)
''
''                Next
''
''                oForm01.Items("Total").Specific.VALUE = totalAmt
''
''                Call oMat01.LoadFromDataSource
''
'            End If
''
'            Call oMat01.AutoResizeColumns
'
'        Case "OrdNum"
'
'            oForm01.Items("JakMyung").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
'            oForm01.Items("JakSize").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
'
'    End Select

    If oUID = "Mat01" Then
    
'        If oCol = "ItemCode" Then
'
'            Call oDS_PS_PP552L.setValue("U_ColReg03", oRow - 1, oMat01.Columns(oCol).Cells(oRow).Specific.VALUE)
'            If oMat01.RowCount = oRow And Trim(oDS_PS_PP552L.GetValue("U_ColReg03", oRow - 1)) <> "" Then
'                Call PS_PP552_AddMatrixRow(oRow)
'            End If
'
'        ElseIf oCol = "CntcCode" Then
'
'            Call oMat01.FlushToDataSource
'            Call oDS_PS_PP552L.setValue("U_ColReg11", oRow - 1, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'"))
'            Call oMat01.LoadFromDataSource
'
''            oMat01.Columns("CntcName").Cells(oRow).Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'")
'
'        End If
'
        Call oMat01.AutoResizeColumns
    
    ElseIf oUID = "Mat02" Then
    
        Call oMat02.AutoResizeColumns
    
    ElseIf oUID = "Mat03" Then
    
        Call oMat03.AutoResizeColumns
    
    ElseIf oUID = "CntcCode01" Then
    
        oForm01.Items("CntcName01").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
        
    ElseIf oUID = "CntcCode02" Then
    
        oForm01.Items("CntcName02").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
        
    ElseIf oUID = "CntcCode03" Then
    
        oForm01.Items("CntcName03").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
        
    ElseIf oUID = "CpCode01" Then
    
        oForm01.Items("CpName01").Specific.VALUE = MDC_GetData.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
        
    ElseIf oUID = "CpCode02" Then
    
        oForm01.Items("CpName02").Specific.VALUE = MDC_GetData.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
        
    ElseIf oUID = "CpCode03" Then
    
        oForm01.Items("CpName03").Specific.VALUE = MDC_GetData.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" & oForm01.Items(oUID).Specific.VALUE & "'")
        
    ElseIf oUID = "OrdNum01" Or oUID = "OrdSub101" Or oUID = "OrdSub201" Then
        
        OrdNum = Trim(oForm01.Items("OrdNum01").Specific.VALUE)
        OrdSub1 = Trim(oForm01.Items("OrdSub101").Specific.VALUE)
        OrdSub2 = Trim(oForm01.Items("OrdSub201").Specific.VALUE)

        sQry = "           SELECT   CASE"
        sQry = sQry & "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakMyung"
        sQry = sQry & "             END AS [ItemName],"
        sQry = sQry & "             CASE"
        sQry = sQry & "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakSize"
        sQry = sQry & "             END AS [ItemSpec]"
        sQry = sQry & " FROM     [@PS_PP020H] AS T0"
        sQry = sQry & " WHERE   T0.U_JakName = '" & OrdNum & "'"
        sQry = sQry & "             AND T0.U_SubNo1 = CASE WHEN '" & OrdSub1 & "' = '' THEN '00' ELSE '" & OrdSub1 & "' END"
        sQry = sQry & "             AND T0.U_SubNo2 = CASE WHEN '" & OrdSub2 & "' = '' THEN '000' ELSE '" & OrdSub2 & "' END"

        Call oRecordSet01.DoQuery(sQry)

        oForm01.Items("ItemName01").Specific.VALUE = oRecordSet01.Fields("ItemName").VALUE
        oForm01.Items("ItemSpec01").Specific.VALUE = oRecordSet01.Fields("ItemSpec").VALUE
    
    ElseIf oUID = "OrdNum02" Or oUID = "OrdSub102" Or oUID = "OrdSub202" Then
        
        OrdNum = Trim(oForm01.Items("OrdNum02").Specific.VALUE)
        OrdSub1 = Trim(oForm01.Items("OrdSub102").Specific.VALUE)
        OrdSub2 = Trim(oForm01.Items("OrdSub202").Specific.VALUE)

        sQry = "           SELECT   CASE"
        sQry = sQry & "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakMyung"
        sQry = sQry & "             END AS [ItemName],"
        sQry = sQry & "             CASE"
        sQry = sQry & "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakSize"
        sQry = sQry & "             END AS [ItemSpec]"
        sQry = sQry & " FROM     [@PS_PP020H] AS T0"
        sQry = sQry & " WHERE   T0.U_JakName = '" & OrdNum & "'"
        sQry = sQry & "             AND T0.U_SubNo1 = CASE WHEN '" & OrdSub1 & "' = '' THEN '00' ELSE '" & OrdSub1 & "' END"
        sQry = sQry & "             AND T0.U_SubNo2 = CASE WHEN '" & OrdSub2 & "' = '' THEN '000' ELSE '" & OrdSub2 & "' END"

        Call oRecordSet01.DoQuery(sQry)

        oForm01.Items("ItemName02").Specific.VALUE = oRecordSet01.Fields("ItemName").VALUE
        oForm01.Items("ItemSpec02").Specific.VALUE = oRecordSet01.Fields("ItemSpec").VALUE
    
    ElseIf oUID = "OrdNum03" Or oUID = "OrdSub103" Or oUID = "OrdSub203" Then
        
        OrdNum = Trim(oForm01.Items("OrdNum03").Specific.VALUE)
        OrdSub1 = Trim(oForm01.Items("OrdSub103").Specific.VALUE)
        OrdSub2 = Trim(oForm01.Items("OrdSub203").Specific.VALUE)

        sQry = "           SELECT   CASE"
        sQry = sQry & "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakMyung"
        sQry = sQry & "             END AS [ItemName],"
        sQry = sQry & "             CASE"
        sQry = sQry & "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        sQry = sQry & "                 ELSE T0.U_JakSize"
        sQry = sQry & "             END AS [ItemSpec]"
        sQry = sQry & " FROM     [@PS_PP020H] AS T0"
        sQry = sQry & " WHERE   T0.U_JakName = '" & OrdNum & "'"
        sQry = sQry & "             AND T0.U_SubNo1 = CASE WHEN '" & OrdSub1 & "' = '' THEN '00' ELSE '" & OrdSub1 & "' END"
        sQry = sQry & "             AND T0.U_SubNo2 = CASE WHEN '" & OrdSub2 & "' = '' THEN '000' ELSE '" & OrdSub2 & "' END"

        Call oRecordSet01.DoQuery(sQry)

        oForm01.Items("ItemName03").Specific.VALUE = oRecordSet01.Fields("ItemName").VALUE
        oForm01.Items("ItemSpec03").Specific.VALUE = oRecordSet01.Fields("ItemSpec").VALUE
    
    End If
    
    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_PP552_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("PS_PP552_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

Sub PS_PP552_FormItemEnabled()
On Error GoTo PS_PP552_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseItmBsort01 = "" Then
        
        Else
        
        End If
        '//각모드에따른 아이템설정
        '//Call PS_PP552_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP552_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP552_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP552_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP552_AddMatrixRow_Error

    Call oForm01.Freeze(True)
    
    If RowIserted = False Then '//행추가여부
        Call oDS_PS_PP552L.InsertRecord(oRow)
    End If
    
    Call oMat01.AddRow
    oDS_PS_PP552L.Offset = oRow
    Call oDS_PS_PP552L.setValue("U_LineNum", oRow, oRow + 1)
    Call oDS_PS_PP552L.setValue("U_ColReg02", oRow, "10") '투입구분
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_PP552_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP552_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP552_FormClear()
On Error GoTo PS_PP552_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP552'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP552_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP552_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP552_EnableMenus()
On Error GoTo PS_PP552_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
'    Call MDC_GP_EnableMenus(oForm01, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP552_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP552_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP552_DataValidCheck() As Boolean
On Error GoTo PS_PP552_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_PP552_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_PP552_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_PP552_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_PP552_FormClear
    Exit Function
PS_PP552_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP552_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Public Sub PS_PP552_MTX01()
'******************************************************************************
'Function ID : PS_PP552_MTX01()
'해당모듈 : PS_PP552
'기능 : 공정상태 삭제 대상 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP552_MTX01_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID As String '사업장
    Dim FrDt As String '기간(시작)
    Dim ToDt As String '기간(종료)
    Dim CpCode As String '공정
    Dim WorkerCode As String '등록자
    Dim WorkGbn As String '작업구분
    Dim CpStatus As String '공정상태
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim MainOrdNum As String '메인작번
    Dim SubOrdNum01 As String '서브작번1
    Dim SubOrdNum02 As String '서브작번2

    BPLID = Trim(oForm01.Items("BPLID01").Specific.Selected.VALUE) '사업장
    FrDt = Trim(oForm01.Items("FrDt01").Specific.VALUE) '기간(시작)
    ToDt = Trim(oForm01.Items("ToDt01").Specific.VALUE) '기간(종료)
    CpCode = Trim(oForm01.Items("CpCode01").Specific.VALUE) '공정
    WorkerCode = Trim(oForm01.Items("CntcCode01").Specific.VALUE) '등록자
    WorkGbn = Trim(oForm01.Items("WorkGbn01").Specific.VALUE) '작업구분
    CpStatus = Trim(oForm01.Items("CpStatus01").Specific.VALUE) '공정상태
    CardType = Trim(oForm01.Items("CardType01").Specific.VALUE) '거래처구분
    ItemType = Trim(oForm01.Items("ItemType01").Specific.VALUE) '품목구분
    MainOrdNum = Trim(oForm01.Items("OrdNum01").Specific.VALUE) '메인작번
    SubOrdNum01 = Trim(oForm01.Items("OrdSub101").Specific.VALUE) '서브작번1
    SubOrdNum02 = Trim(oForm01.Items("OrdSub201").Specific.VALUE)  '서브작번2

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)

    Call oForm01.Freeze(True)

    sQry = "         EXEC [PS_PP552_01] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "','" '기간(종료)
    sQry = sQry & CpCode & "','" '공정
    sQry = sQry & WorkerCode & "','" '등록자
    sQry = sQry & WorkGbn & "','" '작업구분
    sQry = sQry & CpStatus & "','" '공정상태
    sQry = sQry & CardType & "','" '거래처구분
    sQry = sQry & ItemType & "','" '품목구분
    sQry = sQry & MainOrdNum & "','" '메인작번
    sQry = sQry & SubOrdNum01 & "','" '서브작번1
    sQry = sQry & SubOrdNum02 & "'" '서브작번2

    Call oRecordSet01.DoQuery(sQry)

    oMat01.Clear
    oDS_PS_PP552L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then

        ErrNum = 1

        GoTo PS_PP552_MTX01_Error

        Exit Sub
    End If

    For loopCount = 0 To oRecordSet01.RecordCount - 1
        If loopCount + 1 > oDS_PS_PP552L.Size Then
            Call oDS_PS_PP552L.InsertRecord(loopCount)
        End If

        oMat01.AddRow
        oDS_PS_PP552L.Offset = loopCount

        Call oDS_PS_PP552L.setValue("U_LineNum", loopCount, loopCount + 1)
        Call oDS_PS_PP552L.setValue("U_ColReg01", loopCount, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PS_PP552L.setValue("U_ColReg02", loopCount, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작번
        Call oDS_PS_PP552L.setValue("U_ColReg03", loopCount, Trim(oRecordSet01.Fields("OrdSub1").VALUE)) '서브작번1
        Call oDS_PS_PP552L.setValue("U_ColReg04", loopCount, Trim(oRecordSet01.Fields("OrdSub2").VALUE)) '서브작번2
        Call oDS_PS_PP552L.setValue("U_ColQty01", loopCount, Trim(oRecordSet01.Fields("WkOrdQty").VALUE)) '작업지시수량
        Call oDS_PS_PP552L.setValue("U_ColReg05", loopCount, Trim(oRecordSet01.Fields("OrdName").VALUE)) '작명
        Call oDS_PS_PP552L.setValue("U_ColReg06", loopCount, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명
        Call oDS_PS_PP552L.setValue("U_ColReg07", loopCount, Trim(oRecordSet01.Fields("ItemSpec").VALUE)) '규격
        Call oDS_PS_PP552L.setValue("U_ColReg08", loopCount, Trim(oRecordSet01.Fields("CpCode").VALUE)) '공정코드
        Call oDS_PS_PP552L.setValue("U_ColReg09", loopCount, Trim(oRecordSet01.Fields("CpName").VALUE)) '공정명
        Call oDS_PS_PP552L.setValue("U_ColReg10", loopCount, Trim(oRecordSet01.Fields("CpCount").VALUE)) '공정횟수
        Call oDS_PS_PP552L.setValue("U_ColReg11", loopCount, Trim(oRecordSet01.Fields("CpStatus").VALUE)) '공정상태
        Call oDS_PS_PP552L.setValue("U_ColDt01", loopCount, Format(Trim(oRecordSet01.Fields("RegDate").VALUE), "YYYYMMDD")) '등록일
        Call oDS_PS_PP552L.setValue("U_ColReg12", loopCount, Trim(oRecordSet01.Fields("WorkerCode").VALUE)) '등록자사번
        Call oDS_PS_PP552L.setValue("U_ColReg13", loopCount, Trim(oRecordSet01.Fields("WorkerName").VALUE)) '등록자성명
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"

    Next

    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop

    Call oForm01.Freeze(False)

    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP552_MTX01_Error:

    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing

    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP552_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP552_MTX02()
'******************************************************************************
'Function ID : PS_PP552_MTX02()
'해당모듈 : PS_PP552
'기능 : 공정상태 비고 등록 대상 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP552_MTX02_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID As String '사업장
    Dim FrDt As String '기간(시작)
    Dim ToDt As String '기간(종료)
    Dim CpCode As String '공정
    Dim WorkerCode As String '등록자
    Dim WorkGbn As String '작업구분
    Dim CpStatus As String '공정상태
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim MainOrdNum As String '메인작번
    Dim SubOrdNum01 As String '서브작번1
    Dim SubOrdNum02 As String '서브작번2

    BPLID = Trim(oForm01.Items("BPLID02").Specific.Selected.VALUE) '사업장
    FrDt = Trim(oForm01.Items("FrDt02").Specific.VALUE) '기간(시작)
    ToDt = Trim(oForm01.Items("ToDt02").Specific.VALUE) '기간(종료)
    CpCode = Trim(oForm01.Items("CpCode02").Specific.VALUE) '공정
    WorkerCode = Trim(oForm01.Items("CntcCode02").Specific.VALUE) '등록자
    WorkGbn = Trim(oForm01.Items("WorkGbn02").Specific.VALUE) '작업구분
    CpStatus = Trim(oForm01.Items("CpStatus02").Specific.VALUE) '공정상태
    CardType = Trim(oForm01.Items("CardType02").Specific.VALUE) '거래처구분
    ItemType = Trim(oForm01.Items("ItemType02").Specific.VALUE) '품목구분
    MainOrdNum = Trim(oForm01.Items("OrdNum02").Specific.VALUE) '메인작번
    SubOrdNum01 = Trim(oForm01.Items("OrdSub102").Specific.VALUE) '서브작번1
    SubOrdNum02 = Trim(oForm01.Items("OrdSub202").Specific.VALUE)  '서브작번2

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)

    Call oForm01.Freeze(True)

    sQry = "         EXEC [PS_PP552_02] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "','" '기간(종료)
    sQry = sQry & CpCode & "','" '공정
    sQry = sQry & WorkerCode & "','" '등록자
    sQry = sQry & WorkGbn & "','" '작업구분
    sQry = sQry & CpStatus & "','" '공정상태
    sQry = sQry & CardType & "','" '거래처구분
    sQry = sQry & ItemType & "','" '품목구분
    sQry = sQry & MainOrdNum & "','" '메인작번
    sQry = sQry & SubOrdNum01 & "','" '서브작번1
    sQry = sQry & SubOrdNum02 & "'" '서브작번2

    Call oRecordSet01.DoQuery(sQry)

    oMat02.Clear
    oDS_PS_PP552M.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then

        ErrNum = 1

        GoTo PS_PP552_MTX02_Error

        Exit Sub
    End If

    For loopCount = 0 To oRecordSet01.RecordCount - 1
        If loopCount + 1 > oDS_PS_PP552M.Size Then
            Call oDS_PS_PP552M.InsertRecord(loopCount)
        End If

        oMat02.AddRow
        oDS_PS_PP552M.Offset = loopCount

        Call oDS_PS_PP552M.setValue("U_LineNum", loopCount, loopCount + 1)
        Call oDS_PS_PP552M.setValue("U_ColReg01", loopCount, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PS_PP552M.setValue("U_ColReg02", loopCount, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작번
        Call oDS_PS_PP552M.setValue("U_ColReg03", loopCount, Trim(oRecordSet01.Fields("OrdSub1").VALUE)) '서브작번1
        Call oDS_PS_PP552M.setValue("U_ColReg04", loopCount, Trim(oRecordSet01.Fields("OrdSub2").VALUE)) '서브작번2
        Call oDS_PS_PP552M.setValue("U_ColQty01", loopCount, Trim(oRecordSet01.Fields("WkOrdQty").VALUE)) '작업지시수량
        Call oDS_PS_PP552M.setValue("U_ColReg05", loopCount, Trim(oRecordSet01.Fields("OrdName").VALUE)) '작명
        Call oDS_PS_PP552M.setValue("U_ColReg06", loopCount, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명
        Call oDS_PS_PP552M.setValue("U_ColReg07", loopCount, Trim(oRecordSet01.Fields("ItemSpec").VALUE)) '규격
        Call oDS_PS_PP552M.setValue("U_ColReg08", loopCount, Trim(oRecordSet01.Fields("CpCode").VALUE)) '공정코드
        Call oDS_PS_PP552M.setValue("U_ColReg09", loopCount, Trim(oRecordSet01.Fields("CpName").VALUE)) '공정명
        Call oDS_PS_PP552M.setValue("U_ColReg10", loopCount, Trim(oRecordSet01.Fields("CpCount").VALUE)) '공정횟수
        Call oDS_PS_PP552M.setValue("U_ColReg11", loopCount, Trim(oRecordSet01.Fields("CpStatus").VALUE)) '공정상태
        Call oDS_PS_PP552M.setValue("U_ColDt01", loopCount, Format(Trim(oRecordSet01.Fields("RegDate").VALUE), "YYYYMMDD")) '등록일
        Call oDS_PS_PP552M.setValue("U_ColReg12", loopCount, Trim(oRecordSet01.Fields("Comment").VALUE)) '비고
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"

    Next

    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    ProgBar01.Stop

    Call oForm01.Freeze(False)

    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP552_MTX02_Error:

    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing

    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP552_MTX02_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_PP552_MTX03()
'******************************************************************************
'Function ID : PS_PP552_MTX03()
'해당모듈 : PS_PP552
'기능 : 공정상태 계획공수 등록 대상 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP552_MTX03_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID As String '사업장
    Dim FrDt As String '기간(시작)
    Dim ToDt As String '기간(종료)
    Dim CpCode As String '공정
    Dim WorkerCode As String '등록자
    Dim WorkGbn As String '작업구분
    Dim CpStatus As String '공정상태
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    Dim MainOrdNum As String '메인작번
    Dim SubOrdNum01 As String '서브작번1
    Dim SubOrdNum02 As String '서브작번2

    BPLID = Trim(oForm01.Items("BPLID03").Specific.Selected.VALUE) '사업장
    FrDt = Trim(oForm01.Items("FrDt03").Specific.VALUE) '기간(시작)
    ToDt = Trim(oForm01.Items("ToDt03").Specific.VALUE) '기간(종료)
    CpCode = Trim(oForm01.Items("CpCode03").Specific.VALUE) '공정
    WorkerCode = Trim(oForm01.Items("CntcCode03").Specific.VALUE) '등록자
    WorkGbn = Trim(oForm01.Items("WorkGbn03").Specific.VALUE) '작업구분
    CpStatus = Trim(oForm01.Items("CpStatus03").Specific.VALUE) '공정상태
    CardType = Trim(oForm01.Items("CardType03").Specific.VALUE) '거래처구분
    ItemType = Trim(oForm01.Items("ItemType03").Specific.VALUE) '품목구분
    MainOrdNum = Trim(oForm01.Items("OrdNum03").Specific.VALUE) '메인작번
    SubOrdNum01 = Trim(oForm01.Items("OrdSub103").Specific.VALUE) '서브작번1
    SubOrdNum02 = Trim(oForm01.Items("OrdSub203").Specific.VALUE)  '서브작번2

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)

    Call oForm01.Freeze(True)

    sQry = "         EXEC [PS_PP552_03] '"
    sQry = sQry & BPLID & "','" '사업장
    sQry = sQry & FrDt & "','" '기간(시작)
    sQry = sQry & ToDt & "','" '기간(종료)
    sQry = sQry & CpCode & "','" '공정
    sQry = sQry & WorkerCode & "','" '등록자
    sQry = sQry & WorkGbn & "','" '작업구분
    sQry = sQry & CpStatus & "','" '공정상태
    sQry = sQry & CardType & "','" '거래처구분
    sQry = sQry & ItemType & "','" '품목구분
    sQry = sQry & MainOrdNum & "','" '메인작번
    sQry = sQry & SubOrdNum01 & "','" '서브작번1
    sQry = sQry & SubOrdNum02 & "'" '서브작번2

    Call oRecordSet01.DoQuery(sQry)

    oMat03.Clear
    oDS_PS_PP552N.Clear
    oMat03.FlushToDataSource
    oMat03.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then

        ErrNum = 1

        GoTo PS_PP552_MTX03_Error

        Exit Sub
    End If

    For loopCount = 0 To oRecordSet01.RecordCount - 1
        If loopCount + 1 > oDS_PS_PP552N.Size Then
            Call oDS_PS_PP552N.InsertRecord(loopCount)
        End If

        oMat03.AddRow
        oDS_PS_PP552N.Offset = loopCount

        Call oDS_PS_PP552N.setValue("U_LineNum", loopCount, loopCount + 1)
        Call oDS_PS_PP552N.setValue("U_ColReg01", loopCount, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PS_PP552N.setValue("U_ColReg02", loopCount, Trim(oRecordSet01.Fields("OrdNum").VALUE)) '작번
        Call oDS_PS_PP552N.setValue("U_ColReg03", loopCount, Trim(oRecordSet01.Fields("OrdSub1").VALUE)) '서브작번1
        Call oDS_PS_PP552N.setValue("U_ColReg04", loopCount, Trim(oRecordSet01.Fields("OrdSub2").VALUE)) '서브작번2
        Call oDS_PS_PP552N.setValue("U_ColQty02", loopCount, Trim(oRecordSet01.Fields("WkOrdQty").VALUE)) '작업지시수량
        Call oDS_PS_PP552N.setValue("U_ColReg05", loopCount, Trim(oRecordSet01.Fields("OrdName").VALUE)) '작명
        Call oDS_PS_PP552N.setValue("U_ColReg06", loopCount, Trim(oRecordSet01.Fields("ItemName").VALUE)) '품명
        Call oDS_PS_PP552N.setValue("U_ColReg07", loopCount, Trim(oRecordSet01.Fields("ItemSpec").VALUE)) '규격
        Call oDS_PS_PP552N.setValue("U_ColReg08", loopCount, Trim(oRecordSet01.Fields("CpCode").VALUE)) '공정코드
        Call oDS_PS_PP552N.setValue("U_ColReg09", loopCount, Trim(oRecordSet01.Fields("CpName").VALUE)) '공정명
        Call oDS_PS_PP552N.setValue("U_ColReg10", loopCount, Trim(oRecordSet01.Fields("CpCount").VALUE)) '공정횟수
        Call oDS_PS_PP552N.setValue("U_ColReg11", loopCount, Trim(oRecordSet01.Fields("CpStatus").VALUE)) '공정상태
        Call oDS_PS_PP552N.setValue("U_ColDt01", loopCount, Format(Trim(oRecordSet01.Fields("RegDate").VALUE), "YYYYMMDD")) '등록일
        Call oDS_PS_PP552N.setValue("U_ColQty01", loopCount, Trim(oRecordSet01.Fields("PlanHour").VALUE)) '계획공수
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"

    Next

    oMat03.LoadFromDataSource
    oMat03.AutoResizeColumns
    ProgBar01.Stop

    Call oForm01.Freeze(False)

    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP552_MTX03_Error:

    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing

    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP552_MTX03_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_PP552_SaveData02() As Boolean
'******************************************************************************
'Function ID : PS_PP552_SaveData02()
'해당모듈 : PS_PP552
'기능 : 비고 등록
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP552_SaveData02_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
 
    Dim MainOrdNum As String '작번
    Dim SubOrdNum01 As String '서브작번1
    Dim SubOrdNum02 As String '서브작번2
    Dim CpCode As String '공정코드
    Dim CpCount As String '공정횟수
    Dim CpStatus As String '공정상태
    
    Dim Comment As String '비고
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, False)
    
    Call oMat02.FlushToDataSource
    For loopCount = 0 To oMat02.VisualRowCount - 1
        If Trim(oDS_PS_PP552M.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            MainOrdNum = Trim(oDS_PS_PP552M.GetValue("U_ColReg02", loopCount)) '작번
            SubOrdNum01 = Trim(oDS_PS_PP552M.GetValue("U_ColReg03", loopCount)) '서브작번1
            SubOrdNum02 = Trim(oDS_PS_PP552M.GetValue("U_ColReg04", loopCount)) '서브작번2
            CpCode = Trim(oDS_PS_PP552M.GetValue("U_ColReg08", loopCount)) '공정코드
            CpCount = Trim(oDS_PS_PP552M.GetValue("U_ColReg10", loopCount)) - 1 '공정횟수
            CpStatus = Trim(oDS_PS_PP552M.GetValue("U_ColReg11", loopCount)) '공정상태
            Comment = Trim(oDS_PS_PP552M.GetValue("U_ColReg12", loopCount)) '비고
            
            sQry = "         EXEC [PS_PP552_12] '"
            sQry = sQry & MainOrdNum & "','" '작번
            sQry = sQry & SubOrdNum01 & "','" '서브작번1
            sQry = sQry & SubOrdNum02 & "','" '서브작번2
            sQry = sQry & CpCode & "','" '공정코드
            sQry = sQry & CpCount & "','" '공정횟수
            sQry = sQry & CpStatus & "','"  '공정상태
            sQry = sQry & Comment & "'"  '비고

            Call RecordSet01.DoQuery(sQry)
    
        End If
        
    Next

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("등록 완료!", "S")
    
    PS_PP552_SaveData02 = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP552_SaveData02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_PP552_SaveData02_Error:" & Err.Number & " - " & Err.Description, "E")
    
    PS_PP552_SaveData02 = False
    
End Function

Public Function PS_PP552_SaveData03() As Boolean
'******************************************************************************
'Function ID : PS_PP552_SaveData03()
'해당모듈 : PS_PP552
'기능 : 계획공수 등록
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP552_SaveData03_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
 
    Dim MainOrdNum As String '작번
    Dim SubOrdNum01 As String '서브작번1
    Dim SubOrdNum02 As String '서브작번2
    Dim CpCode As String '공정코드
    Dim CpCount As String '공정횟수
    Dim CpStatus As String '공정상태
    
    Dim PlanHour As Double '비고
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, False)
    
    Call oMat03.FlushToDataSource
    For loopCount = 0 To oMat03.VisualRowCount - 1
        If Trim(oDS_PS_PP552N.GetValue("U_ColReg01", loopCount)) = "Y" Then
            
            MainOrdNum = Trim(oDS_PS_PP552N.GetValue("U_ColReg02", loopCount)) '작번
            SubOrdNum01 = Trim(oDS_PS_PP552N.GetValue("U_ColReg03", loopCount)) '서브작번1
            SubOrdNum02 = Trim(oDS_PS_PP552N.GetValue("U_ColReg04", loopCount)) '서브작번2
            CpCode = Trim(oDS_PS_PP552N.GetValue("U_ColReg08", loopCount)) '공정코드
            CpCount = Trim(oDS_PS_PP552N.GetValue("U_ColReg10", loopCount)) - 1 '공정횟수
            CpStatus = Trim(oDS_PS_PP552N.GetValue("U_ColReg11", loopCount)) '공정상태
            PlanHour = Trim(oDS_PS_PP552N.GetValue("U_ColQty01", loopCount)) '비고
            
            sQry = "         EXEC [PS_PP552_13] '"
            sQry = sQry & MainOrdNum & "','" '작번
            sQry = sQry & SubOrdNum01 & "','" '서브작번1
            sQry = sQry & SubOrdNum02 & "','" '서브작번2
            sQry = sQry & CpCode & "','" '공정코드
            sQry = sQry & CpCount & "','" '공정횟수
            sQry = sQry & CpStatus & "','" '공정상태
            sQry = sQry & PlanHour & "'"  '계획공수

            Call RecordSet01.DoQuery(sQry)
    
        End If
        
    Next

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("등록 완료!", "S")
    
    PS_PP552_SaveData03 = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP552_SaveData03_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_PP552_SaveData03_Error:" & Err.Number & " - " & Err.Description, "E")
    
    PS_PP552_SaveData03 = False
    
End Function

Public Function PS_PP552_DeleteData01() As Boolean
'******************************************************************************
'Function ID : PS_PP552_DeleteData01()
'해당모듈 : PS_PP552
'기능 : 공정정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP552_DeleteData01_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim MainOrdNum As String '작번
    Dim SubOrdNum01 As String '서브작번1
    Dim SubOrdNum02 As String '서브작번2
    Dim CpCode As String '공정코드
    Dim CpCount As String '공정횟수
    Dim CpStatus As String '공정상태
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oMat01.VisualRowCount = 0 Then
    
        ErrNum = 1
        GoTo PS_PP552_DeleteData01_Error
    
    End If
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
        If Trim(oDS_PS_PP552L.GetValue("U_ColReg01", loopCount)) = "Y" Then
    
            MainOrdNum = Trim(oDS_PS_PP552L.GetValue("U_ColReg02", loopCount)) '작번
            SubOrdNum01 = Trim(oDS_PS_PP552L.GetValue("U_ColReg03", loopCount)) '서브작번1
            SubOrdNum02 = Trim(oDS_PS_PP552L.GetValue("U_ColReg04", loopCount)) '서브작번2
            CpCode = Trim(oDS_PS_PP552L.GetValue("U_ColReg08", loopCount)) '공정코드
            CpCount = Trim(oDS_PS_PP552L.GetValue("U_ColReg10", loopCount)) - 1 '공정횟수
            CpStatus = Trim(oDS_PS_PP552L.GetValue("U_ColReg11", loopCount)) '공정상태
            
            sQry = "                EXEC [PS_PP552_31] "
            sQry = sQry & "'" & MainOrdNum & "'," '작번
            sQry = sQry & "'" & SubOrdNum01 & "'," '서브작번1
            sQry = sQry & "'" & SubOrdNum02 & "'," '서브작번2
            sQry = sQry & "'" & CpCode & "'," '공정코드
            sQry = sQry & "'" & CpCount & "'," '공정횟수
            sQry = sQry & "'" & CpStatus & "'"  '공정상태
            
            Call oRecordSet01.DoQuery(sQry)
    
        End If
    
    Next
    
    Set oRecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
    PS_PP552_DeleteData01 = True
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP552_DeleteData01_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP552_DeleteData01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
    
    PS_PP552_DeleteData01 = False
    
End Function

Private Function PS_PP552_DI_API() As Boolean
'On Error GoTo PS_PP552_DI_API_Error
'    PS_PP552_DI_API = True
'    Dim i, j As Long
'    Dim oDIObject As SAPbobsCOM.Documents
'    Dim RetVal As Long
'    Dim LineNumCount As Long
'    Dim ResultDocNum As Long
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Company.StartTransaction
'
'    ReDim ItemInformation(0)
'    ItemInformationCount = 0
'    For i = 1 To oMat01.VisualRowCount
'        ReDim Preserve ItemInformation(ItemInformationCount)
'        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.Value
'        ItemInformation(ItemInformationCount).Check = False
'        ItemInformationCount = ItemInformationCount + 1
'    Next
'
'    LineNumCount = 0
'    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
'    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
'    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
'    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.Value, "&&&&-&&-&&")
'    For i = 0 To UBound(ItemInformation)
'        If ItemInformation(i).Check = True Then
'            GoTo Continue_First
'        End If
'        If i <> 0 Then
'            oDIObject.Lines.Add
'        End If
'        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
'        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
'        oDIObject.Lines.BaseType = "22"
'        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
'        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
'        For j = i To UBound(ItemInformation)
'            If ItemInformation(j).Check = True Then
'                GoTo Continue_Second
'            End If
'            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
'                GoTo Continue_Second
'            End If
'            '//같은것
'            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
'            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
'            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
'            oDIObject.Lines.BatchNumbers.Add
'            ItemInformation(j).PDN1No = LineNumCount
'            ItemInformation(j).Check = True
'Continue_Second:
'        Next
'        LineNumCount = LineNumCount + 1
'Continue_First:
'    Next
'    RetVal = oDIObject.Add
'    If RetVal = 0 Then
'        ResultDocNum = Sbo_Company.GetNewObjectKey
'        For i = 0 To UBound(ItemInformation)
'            Call oDS_PS_PP552L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_PP552L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
'    Else
'        GoTo PS_PP552_DI_API_Error
'    End If
'
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_Commit
'    End If
'    oMat01.LoadFromDataSource
'    oMat01.AutoResizeColumns
'
'    Set oDIObject = Nothing
'    Exit Function
'PS_PP552_DI_API_DI_Error:
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
'    PS_PP552_DI_API = False
'    Set oDIObject = Nothing
'    Exit Function
'PS_PP552_DI_API_Error:
'    Sbo_Application.SetStatusBarMessage "PS_PP552_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    PS_PP552_DI_API = False
End Function

Private Sub PS_PP552_SetBaseForm()
On Error GoTo PS_PP552_SetBaseForm_Error
    Dim i As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As Matrix
    If oBaseForm01 Is Nothing Then
    '//DoNothing
    Else
    
    End If
    Exit Sub
PS_PP552_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP552_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP552_FormResize()
On Error GoTo PS_PP552_FormResize_Error

    '그룹박스 크기 동적 할당
    oForm01.Items("GrpBox01").Height = oForm01.Items("Mat01").Height + 170
    oForm01.Items("GrpBox01").Width = oForm01.Items("Mat01").Width + 30
    
    Call oMat01.AutoResizeColumns
    Call oMat02.AutoResizeColumns
    Call oMat03.AutoResizeColumns
    
    Exit Sub
PS_PP552_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP552_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP552_CheckBeforeSearch() As Boolean
'******************************************************************************
'Function ID : PS_PP552_CheckBeforeSearch()
'해당모듈 : PS_PP552
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP552_CheckBeforeSearch_Error

    Dim ErrNum As Integer
    ErrNum = 0
    
'    Select Case True
'        Case Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "%" '작업구분
'            ErrNum = 1
'            GoTo PS_PP552_CheckBeforeSearch_Error
'    End Select

    PS_PP552_CheckBeforeSearch = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP552_CheckBeforeSearch_Error:
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회조건 작업구분은 필수선택 사항입니다. 확인하세요.", "E")
        Call oForm01.Items("OrdGbn").Click(ct_Regular)
    End If
    PS_PP552_CheckBeforeSearch = False
End Function

Private Function PS_PP552_CheckAfterStatus(ByVal pRow As Integer) As Boolean
'******************************************************************************
'Function ID : PS_PP552_CheckAfterStatus()
'해당모듈 : PS_PP552
'기능 : 이후 공정상태가 존재하는지 확인
'인수 : 없음
'반환값 : True:존재, Fasle:미존재
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP552_CheckAfterStatus_Error

    Dim ErrNum As Integer
    Dim loopCount As Long
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim MainOrdNum As String '작번
    Dim SubOrdNum01 As String '서브작번1
    Dim SubOrdNum02 As String '서브작번2
    Dim CpCode As String '공정코드
    Dim CpCount As String '공정횟수
    Dim CpStatus As String '공정상태
    
    MainOrdNum = Trim(oDS_PS_PP552L.GetValue("U_ColReg02", pRow - 1)) '작번
    SubOrdNum01 = Trim(oDS_PS_PP552L.GetValue("U_ColReg03", pRow - 1)) '서브작번1
    SubOrdNum02 = Trim(oDS_PS_PP552L.GetValue("U_ColReg04", pRow - 1)) '서브작번2
    CpCode = Trim(oDS_PS_PP552L.GetValue("U_ColReg08", pRow - 1)) '공정코드
    CpCount = Trim(oDS_PS_PP552L.GetValue("U_ColReg10", pRow - 1)) - 1 '공정횟수
    CpStatus = Trim(oDS_PS_PP552L.GetValue("U_ColReg11", pRow - 1)) '공정상태
    
    'ErrNum = 0
    
'    Call oMat01.FlushToDataSource
    
    sQry = "         EXEC [PS_PP552_91] '"
    sQry = sQry & MainOrdNum & "','" '작번
    sQry = sQry & SubOrdNum01 & "','" '서브작번1
    sQry = sQry & SubOrdNum02 & "','" '서브작번2
    sQry = sQry & CpCode & "','" '공정코드
    sQry = sQry & CpCount & "','" '공정횟수
    sQry = sQry & CpStatus & "'" '공정상태

    Call RecordSet01.DoQuery(sQry)
    
    If RecordSet01.Fields("AfterStatusYN").VALUE = "Y" Then
        PS_PP552_CheckAfterStatus = True
    Else
        PS_PP552_CheckAfterStatus = False
    End If
    
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP552_CheckAfterStatus_Error:

    If ErrNum = 1 Then
        
    End If
    
    PS_PP552_CheckAfterStatus = False
End Function

Private Sub PS_PP552_Print_Report01()
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
    
    Dim BPLID As String
    Dim FrDt As String
    Dim ToDt As String
    Dim Purchase As String
    Dim OrderYN As String
    Dim InputYN As String
    Dim Class As String
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    FrDt = Trim(oForm01.Items("FrDt").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt").Specific.VALUE)
    Purchase = Trim(oForm01.Items("Purchase").Specific.VALUE)
    OrderYN = Trim(oForm01.Items("OrderYN").Specific.VALUE)
    InputYN = Trim(oForm01.Items("InputYN").Specific.VALUE)
    Class = Trim(oForm01.Items("Class").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP552] 레포트"
    
    If Purchase = "10" Then '원재료품의
    
        ReportName = "PS_PP552_10.rpt"
        
        '쿼리
        sQry = "         EXEC PS_PP552_10 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & OrderYN & "','"
        sQry = sQry & InputYN & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
    
    ElseIf Purchase = "20" Then '부재료품의
    
        ReportName = "PS_PP552_20.rpt"
    
        '쿼리
        sQry = "         EXEC PS_PP552_20 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & OrderYN & "','"
        sQry = sQry & InputYN & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
    
    ElseIf Purchase = "30" Then '가공비품의
    
        ReportName = "PS_PP552_30.rpt"
        
        '쿼리
        sQry = "         EXEC PS_PP552_30 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
    
    ElseIf Purchase = "40" Then ''외주제작품의
    
        ReportName = "PS_PP552_40.rpt"
        
        '쿼리
        sQry = "         EXEC PS_PP552_40 '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & FrDt & "','"
        sQry = sQry & ToDt & "','"
        sQry = sQry & Purchase & "','"
        sQry = sQry & Class & "','"
        sQry = sQry & 2 & "'" '출력은 2
        
    End If
    
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
    


    Call oRecordSet.DoQuery(sQry)
    
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
        MDC_Com.MDC_GF_Message "PS_PP552_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub








