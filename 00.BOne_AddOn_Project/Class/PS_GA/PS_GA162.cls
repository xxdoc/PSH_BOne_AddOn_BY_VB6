VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA162"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_GA162.cls
'//  Module : 총무관리>전산장비관리>
'//  Desc : 위치이력
'//  FormType : PS_GA162
'//  Create Date : 2012.10.26
'//  Modified Date :
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_GA162H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_GA162L As SAPbouiCOM.DBDataSource '등록라인

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
Public Sub LoadForm(ByVal prmSeq As Integer)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA162.srf")
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
    
    oFormUniqueID01 = "PS_GA162_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_GA162_CreateItems
    Call PS_GA162_ComboBox_Setting
    Call PS_GA162_Initial_Setting(prmSeq)
    Call PS_GA162_CF_ChooseFromList
    Call PS_GA162_EnableMenus
    Call PS_GA162_SetDocument(prmSeq)
    Call PS_GA162_FormResize
    
    oForm01.EnableMenu ("1283"), False         '// 삭제
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), True        '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    
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
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_GA162_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_GA162_FormItemEnabled '//UDO방식
                Call PS_GA162_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_GA162_FormItemEnabled
            Case "1287"
            
                oForm01.Freeze True
                Call oDS_PS_GA162H.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat01.VisualRowCount - 1
                    Call oMat01.FlushToDataSource
                    Call oDS_PS_GA162H.setValue("DocEntry", i, "")
                    Call oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
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
    
        If pval.ItemUID = "PS_GA162" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_GA162_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_GA162_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        End If
        
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "PS_GA162" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_GA162_FormItemEnabled
                    Call PS_GA162_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_GA162_FormItemEnabled
                End If
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
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.ColUID = "CntcCode" Then
            
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "CntcCode") '사번 포맷서치설정
            
            ElseIf pval.ColUID = "CpCode" Then
            
                If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                    Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "CpCode") '공정 포맷서치설정
            
                End If
                
            ElseIf pval.ColUID = "ToxCode" Then
            
                If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
                
                    Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "ToxCode") '유해인자 포맷서치설정
                
                End If
                
            End If
        
'        ElseIf pval.ItemUID = "RegCd" Then
'
'            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RegCd", "") '등록자 포맷서치 설정
'
'        ElseIf pval.ItemUID = "ReqCd" Then
'
'            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ReqCd", "") '요청자 포맷서치 설정
        
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
            
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then

            If (pval.ItemUID = "Mat01") Then
            
                If pval.ColUID = "DocDate" Then
                
                    oMat01.FlushToDataSource
                    Call oDS_PS_GA162L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    Call oDS_PS_GA162L.setValue("U_CpName", pval.Row - 1, MDC_GetData.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'")) '공정명
                    oMat01.LoadFromDataSource
                    
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_GA162L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_GA162_AddMatrixRow (pval.Row)
                    End If
                    oMat01.LoadFromDataSource
            
                ElseIf pval.ColUID = "CntcCode" Then
                    oMat01.FlushToDataSource
                    
                    Call oDS_PS_GA162L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_GA162L.setValue("U_CntcName", pval.Row - 1, MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'")) '성명
'                    Call oDS_PS_GA162L.setValue("U_TeamCode", pval.Row - 1, MDC_GetData.Get_ReData("U_TeamCode", "U_MSTCOD", "[OHEM]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'")) '소속팀
'                    Call oDS_PS_GA162L.setValue("U_RspCode", pval.Row - 1, MDC_GetData.Get_ReData("U_RspCode", "U_MSTCOD", "[OHEM]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'")) '소속담당
                    oMat01.LoadFromDataSource
                    
                Else
                
                End If
                
'                Call oMat01.Columns(pval.ColUID).Cells(pval.Row).Click(ct_Regular)
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                
            End If
        End If
        
    ElseIf pval.BeforeAction = False Then
        Call oMat01.AutoResizeColumns
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
        Call PS_GA162_FormItemEnabled
        Call PS_GA162_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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
        Call PS_GA162_FormResize
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
            Call oDS_PS_GA162L.RemoveRecord(oDS_PS_GA162L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_GA162_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_GA162L.GetValue("U_DocDate", oMat01.RowCount - 1)) <> "" Then
                    Call PS_GA162_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_GA162_CreateItems() As Boolean
On Error GoTo PS_GA162_CreateItems_Error
   
    Set oDS_PS_GA162H = oForm01.DataSources.DBDataSources("@PS_GA162H")
    Set oDS_PS_GA162L = oForm01.DataSources.DBDataSources("@PS_GA162L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Exit Function
PS_GA162_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA162_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_GA162_ComboBox_Setting()
On Error GoTo PS_GA162_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_GA162", "Mat01", "UseYN", "Y", "Y")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_GA162", "Mat01", "UseYN", "N", "N")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("UseYN"), "PS_GA162", "Mat01", "UseYN")

'    '사업부
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
'
'    '반기구분_S
'    sQry = "            SELECT  U_Code,"
'    sQry = sQry & "             U_CodeNm"
'    sQry = sQry & "  FROM   [@PS_GA050L]"
'    sQry = sQry & "  WHERE  Code = '6'"
'    sQry = sQry & "             AND U_UseYN = 'Y'"
'
'    Call oForm01.Items("HarfCls").Specific.ValidValues.Add("%", "선택")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("HarfCls").Specific, sQry, "", False, False)
'    Call oForm01.Items("HarfCls").Specific.Select(0, psk_Index)
'    '반기구분_E
'
'    '매트릭스-소속팀_S
'    sQry = "           SELECT      b.U_Code, "
'    sQry = sQry & "                b.U_CodeNm"
'    sQry = sQry & " FROM       [@PS_HR200H] a"
'    sQry = sQry & "                INNER JOIN"
'    sQry = sQry & "                [@PS_HR200L] b"
'    sQry = sQry & "                    On a.Code = b.Code"
'    sQry = sQry & "                    And a.Name = '부서'"
'
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("TeamCode"), sQry)
'    '매트릭스-소속팀_E
'
'    '매트릭스-소속담당콤보_S
'    sQry = "           SELECT      b.U_Code, "
'    sQry = sQry & "                b.U_CodeNm"
'    sQry = sQry & " FROM       [@PS_HR200H] a"
'    sQry = sQry & "                INNER JOIN"
'    sQry = sQry & "                [@PS_HR200L] b"
'    sQry = sQry & "                    On a.Code = b.Code"
'    sQry = sQry & "                    And a.Name = '담당'"
'
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("RspCode"), sQry)
'    '매트릭스-소속담당콤보_E
'
    '매트릭스-위치구분_S
'    sQry = "           SELECT      T0.U_Code AS [Code],"
'    sQry = sQry & "                T0.U_CodeNm AS [Name]"
'    sQry = sQry & " FROM       ("
'    sQry = sQry & "                    SELECT      '%' AS [U_Code],"
'    sQry = sQry & "                                   '선택' AS [U_CodeNm],"
'    sQry = sQry & "                                   0 AS [U_Seq]"
'    sQry = sQry & "                    UNION ALL"
'    sQry = sQry & "                    SELECT      b.U_Code,"
'    sQry = sQry & "                                   b.U_CodeNm,"
'    sQry = sQry & "                                   b.U_Seq"
'    sQry = sQry & "                    FROM       [@PS_GA050L] b"
'    sQry = sQry & "                    WHERE      b.Code = '14'"
'    sQry = sQry & "                                   AND U_UseYN = 'Y'"
'    sQry = sQry & "                ) AS T0"
'    sQry = sQry & " ORDER BY T0.U_Seq"

    sQry = "            SELECT      b.U_Code,"
    sQry = sQry & "                 b.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_GA050L] b"
    sQry = sQry & "  WHERE      b.Code = '14'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  b.U_Seq"

    Call oMat01.Columns("LocCls").ValidValues.Add("", "")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("LocCls"), sQry)
    '매트릭스-위치구분_E

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
PS_GA162_ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA162_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA162_CF_ChooseFromList()
On Error GoTo PS_GA162_CF_ChooseFromList_Error
    Exit Sub
PS_GA162_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA162_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA162_FormItemEnabled()
On Error GoTo PS_GA162_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
'
        'oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        oForm01.Items("Seq").Enabled = False
        
        Call PS_GA162_FormClear '//UDO방식
'        oForm01.Mode = fm_OK_MODE '추가 후 모드를 확인 모드로 강제 지정
        
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
'        oForm01.Items("DocEntry").Specific.VALUE = ""
'        oForm01.Items("DocEntry").Enabled = True

        oForm01.Items("Seq").Specific.VALUE = ""
        oForm01.Items("Seq").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("Seq").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA162_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA162_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA162_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA162_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_GA162L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_GA162L.Offset = oRow
    oDS_PS_GA162L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA162_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA162_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA162_FormClear()
On Error GoTo PS_GA162_FormClear_Error

    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_GA162'", "")
    
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    
    Exit Sub
PS_GA162_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA162_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA162_Initial_Setting(ByVal prmSeq As Integer)
On Error GoTo PS_GA162_Initial_Setting_Error
    ' 사업장
'    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

    '검진년도
    oForm01.Items("Seq").Specific.VALUE = prmSeq

    Exit Sub
PS_GA162_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA162_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA162_EnableMenus()
On Error GoTo PS_GA162_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_GA162_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA162_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA162_SetDocument(ByVal prmSeq As String)
On Error GoTo PS_GA162_SetDocument_Error
    
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT COUNT(*) FROM [@PS_GA162H] WHERE U_Seq = " & prmSeq
    Call oRecordSet01.DoQuery(sQry)

    If (oRecordSet01.Fields(0).VALUE = 0) Then '데이터가 이미 존재하지 않으면 AddMode
        Call PS_GA162_FormItemEnabled
        Call PS_GA162_AddMatrixRow(0, True) '//UDO방식일때
    Else '존재하면 FindMode
        oForm01.Mode = fm_FIND_MODE
        Call PS_GA162_FormItemEnabled
        oForm01.Items("Seq").Specific.VALUE = prmSeq
        oForm01.Items("1").Click ct_Regular
    End If
    
    Set oRecordSet01 = Nothing
    
    Exit Sub
    
PS_GA162_SetDocument_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_GA162_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_GA162_DataValidCheck() As Boolean
On Error GoTo PS_GA162_DataValidCheck_Error
    PS_GA162_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_GA162_FormClear
    End If
    
'    '검진년도 미입력 시
'    If oForm01.Items("ChkYear").Specific.VALUE = "" Then
'        Call Sbo_Application.SetStatusBarMessage("검진년도가 입력되지 않았습니다.", bmt_Short, True)
'        PS_GA162_DataValidCheck = False
'        Exit Function
'    End If
'
'    '반기구분 미선택시
'    If oForm01.Items("HarfCls").Specific.Selected.VALUE = "%" Then
'        Call Sbo_Application.SetStatusBarMessage("반기구분이 선택되지 않았습니다.", bmt_Short, True)
'        PS_GA162_DataValidCheck = False
'        Exit Function
'    End If
    
    '라인정보 미입력 시
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_GA162_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        
        '일자
        If (oMat01.Columns("DocDate").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "일자는 필수입니다.", bmt_Short, True
            oMat01.Columns("DocDate").Cells(i).Click ct_Regular
            PS_GA162_DataValidCheck = False
            Exit Function
        End If
        
        '위치구분
        If (oMat01.Columns("LocCls").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "위치구분은 필수입니다.", bmt_Short, True
            oMat01.Columns("LocCls").Cells(i).Click ct_Regular
            PS_GA162_DataValidCheck = False
            Exit Function
        End If

        '주위치
        If (oMat01.Columns("MainLoc").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "주위치는 필수입니다.", bmt_Short, True
            oMat01.Columns("MainLoc").Cells(i).Click ct_Regular
            PS_GA162_DataValidCheck = False
            Exit Function
        End If

    Next
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_GA162L.RemoveRecord(oDS_PS_GA162L.Size - 1)
    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_GA162_FormClear
    End If
    
    PS_GA162_DataValidCheck = True
    Exit Function
PS_GA162_DataValidCheck_Error:
    PS_GA162_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_GA162_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PS_GA162_MTX01()
On Error GoTo PS_GA162_MTX01_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Param01 = oForm01.Items("Param01").Specific.VALUE
    Param02 = oForm01.Items("Param01").Specific.VALUE
    Param03 = oForm01.Items("Param01").Specific.VALUE
    Param04 = oForm01.Items("Param01").Specific.VALUE
    
    Query01 = "SELECT 10"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_GA162_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_GA162L.InsertRecord (i)
        End If
        oDS_PS_GA162L.Offset = i
        Call oDS_PS_GA162L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_GA162L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_GA162_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_GA162_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA162_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_GA162_FormResize()
On Error GoTo PS_GA162_FormResize_Error
    
    Exit Sub
PS_GA162_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA162_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_GA162_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_GA162_Validate_Error
    PS_GA162_Validate = True
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
PS_GA162_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_GA162_Validate_Error:
    PS_GA162_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_GA162_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function












