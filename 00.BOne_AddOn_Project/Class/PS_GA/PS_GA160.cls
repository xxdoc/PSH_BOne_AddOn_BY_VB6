VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_GA160.cls
'//  Module : 총무관리>전산장비관리>
'//  Desc : 전산장비등록
'//  FormType : PS_GA160
'//  Create Date : 2012.10.25
'//  Modified Date : 2012.10.30
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Private QREncode As New cQREncode

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_GA160H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_GA160L As SAPbouiCOM.DBDataSource '등록라인

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
Public Sub LoadForm(Optional ByVal prmManageNo As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA160.srf")
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
    
    oFormUniqueID01 = "PS_GA160_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_GA160_CreateItems
    Call PS_GA160_ComboBox_Setting
    Call PS_GA160_CF_ChooseFromList
    Call PS_GA160_EnableMenus
    Call PS_GA160_SetDocument(prmManageNo)
    Call PS_GA160_FormResize
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), False        '// 취소
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
                Call PS_GA160_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_GA160_FormItemEnabled '//UDO방식
                Call PS_GA160_Initial_Setting
'                Call PS_GA160_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_GA160_FormItemEnabled
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

    Dim SavedPointer As Integer

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "PS_GA160" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_GA160_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = oForm01.Items("Code").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_GA160_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = oForm01.Items("Code").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        '사용자이력
        If pval.ItemUID = "btnUser" Then
            
            Call PS_GA161_Open
            
        End If
        
        '위치이력
        If pval.ItemUID = "btnLoc" Then
        
            Call PS_GA162_Open
        
        End If
        
        '점검이력
        If pval.ItemUID = "btnChk" Then
        
            Call PS_GA163_Open
        
        End If
        
        'HW이력
        If pval.ItemUID = "btnHW" Then
        
            Call PS_GA164_Open
        
        End If
        
        'SW이력
        If pval.ItemUID = "btnSW" Then
        
            Call PS_GA165_Open
        
        End If
        
        'IP이력
        If pval.ItemUID = "btnIP" Then
        
            Call PS_GA166_Open
        
        End If
        
'        If pval.ItemUID = "BtnTest" Then
'
'            With QREncode.QREncode(New_c.Crypt.VBStringToUTF8(Me.oForm01.Items("Comment").Specific.VALUE)) 'ImageSurface is the default
''              .WriteContentToPngFile App.Path & "\QR.png"
'              .WriteContentToJpgFile App.Path & "\QR.jpg", 95
'            End With
'
'        End If
        
        
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA160" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_GA160_FormItemEnabled
'                    Call PS_GA160_AddMatrixRow(0, True) '//UDO방식일때
                    Call PS_GA160_Initial_Setting
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_GA160_FormItemEnabled
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
                If pval.ColUID = "Code" Then
                    oMat01.FlushToDataSource
                    
                    Call oDS_PS_GA160L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_GA160L.setValue("U_Seq", pval.Row - 1, pval.Row)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_GA160L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_GA160_AddMatrixRow (pval.Row)
                    End If
                    oMat01.LoadFromDataSource
                End If
                Call oMat01.Columns("UseYN").Cells(pval.Row).Specific.Select(0, psk_Index) '기본으로 'Y' 세팅
                
'                Call oMat01.Columns(pval.ColUID).Cells(pval.Row).Click(ct_Regular)
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                
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
        Call PS_GA160_FormItemEnabled
        Call PS_GA160_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_GA160_FormResize
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
            Call oDS_PS_GA160L.RemoveRecord(oDS_PS_GA160L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_GA160_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_GA160L.GetValue("U_Code", oMat01.RowCount - 1)) <> "" Then
                    Call PS_GA160_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_GA160_CreateItems() As Boolean
On Error GoTo PS_GA160_CreateItems_Error
   
    Set oDS_PS_GA160H = oForm01.DataSources.DBDataSources("@PS_GA160H")
'    Set oDS_PS_GA160L = oForm01.DataSources.DBDataSources("@PS_GA160L")
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
'    oMat01.SelectionMode = ms_NotSupported
'    oMat01.AutoResizeColumns
    
    Exit Function
PS_GA160_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA160_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_GA160_ComboBox_Setting()
On Error GoTo PS_GA160_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    'Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    'Dim oRecordSet01 As SAPbobsCOM.Recordset
        
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_GA160", "Mat01", "UseYN", "Y", "Y")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_GA160", "Mat01", "UseYN", "N", "N")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("UseYN"), "PS_GA160", "Mat01", "UseYN")
    
    '분류_S
    sQry = "            SELECT      U_Code,"
    sQry = sQry & "                 U_CodeNm"
    sQry = sQry & "  FROM       [@PS_GA050L]"
    sQry = sQry & "  WHERE      Code = '12'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call oForm01.Items("Ctgr").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Ctgr").Specific, sQry, "%", False, False)
    '분류_E
    
    '사업장_S
    sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId"
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, sQry, MDC_PS_Common.User_BPLId, False, False)
    '사업장_E
    
    '제조사_S
    sQry = "            SELECT      U_Code,"
    sQry = sQry & "                 U_CodeNm"
    sQry = sQry & "  FROM       [@PS_GA050L]"
    sQry = sQry & "  WHERE      Code = '13'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call oForm01.Items("Maker").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Maker").Specific, sQry, "%", False, False)
    '제조사_E
    
    'OS_S
    sQry = "            SELECT      U_Code,"
    sQry = sQry & "                 U_CodeNm"
    sQry = sQry & "  FROM       [@PS_GA050L]"
    sQry = sQry & "  WHERE      Code = '16'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call oForm01.Items("OS").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OS").Specific, sQry, "%", False, False)
    'OS_E
    
    '구입여부_S
    sQry = "            SELECT      U_Code,"
    sQry = sQry & "                 U_CodeNm"
    sQry = sQry & "  FROM       [@PS_GA050L]"
    sQry = sQry & "  WHERE      Code = '17'"
    sQry = sQry & "                 AND U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  U_Seq"
    Call oForm01.Items("PchsYN").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("PchsYN").Specific, sQry, "%", False, False)
    '구입여부_E

'    Set oCombo = Nothing
'    Set oRecordSet01 = Nothing

    Exit Sub
PS_GA160_ComboBox_Setting_Error:
'    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA160_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA160_CF_ChooseFromList()
On Error GoTo PS_GA160_CF_ChooseFromList_Error
    Exit Sub
PS_GA160_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA160_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA160_FormItemEnabled()
On Error GoTo PS_GA160_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
'
        oForm01.Items("Code").Enabled = False
        oForm01.Items("MngNo").Enabled = True
'        oForm01.Items("Mat01").Enabled = True
        Call PS_GA160_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("Code").Specific.VALUE = ""
        oForm01.Items("Code").Enabled = True
        oForm01.Items("MngNo").Enabled = True
'        oForm01.Items("Mat01").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("Code").Enabled = False
        oForm01.Items("MngNo").Enabled = False
'        oForm01.Items("Mat01").Enabled = True
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA160_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA160_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA160_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA160_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_GA160L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_GA160L.Offset = oRow
    oDS_PS_GA160L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_GA160_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA160_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA160_FormClear()
On Error GoTo PS_GA160_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_GA160'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
        oForm01.Items("Code").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
        oForm01.Items("Code").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_GA160_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA160_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA160_EnableMenus()
On Error GoTo PS_GA160_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_GA160_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA160_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA160_Initial_Setting()
On Error GoTo PS_GA160_Initial_Setting_Error

    '전산장비분류
    Call oForm01.Items("Ctgr").Specific.Select("%", psk_ByValue)
    
    '사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

    '제조사
    Call oForm01.Items("Maker").Specific.Select("%", psk_ByValue)
    
    'OS
    Call oForm01.Items("OS").Specific.Select("%", psk_ByValue)
    
    '구입여부
    Call oForm01.Items("PchsYN").Specific.Select("%", psk_ByValue)

    Exit Sub
PS_GA160_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA160_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA160_SetDocument(ByVal prmManageNo As String)
On Error GoTo PS_GA160_SetDocument_Error
    If (prmManageNo = "") Then
        Call PS_GA160_FormItemEnabled
'        Call PS_GA160_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_GA160_FormItemEnabled
        oForm01.Items("MngNo").Specific.VALUE = prmManageNo '관리번호
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_GA160_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA160_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_GA160_DataValidCheck() As Boolean
On Error GoTo PS_GA160_DataValidCheck_Error
    PS_GA160_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_GA160_FormClear
    End If
    
    '분류 미선택 시
    If oForm01.Items("Ctgr").Specific.VALUE = "%" Then
        Call Sbo_Application.SetStatusBarMessage("분류가 선택되지 않았습니다.", bmt_Short, True)
        PS_GA160_DataValidCheck = False
        Exit Function
    End If
    
    '사업장 미선택 시
    If oForm01.Items("BPLId").Specific.VALUE = "%" Then
        Call Sbo_Application.SetStatusBarMessage("사업장이 선택되지 않았습니다.", bmt_Short, True)
        PS_GA160_DataValidCheck = False
        Exit Function
    End If
    
    '관리번호 미입력 시
    If oForm01.Items("MngNo").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("관리번호가 입력되지 않았습니다.", bmt_Short, True)
        PS_GA160_DataValidCheck = False
        Exit Function
    End If
    
    '모델명 미입력 시
    If oForm01.Items("ModelNm").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("모델명이 입력되지 않았습니다.", bmt_Short, True)
        PS_GA160_DataValidCheck = False
        Exit Function
    End If
    
    '제조사 미선택 시
    If oForm01.Items("Maker").Specific.VALUE = "%" Then
        Call Sbo_Application.SetStatusBarMessage("제조사가 선택되지 않았습니다.", bmt_Short, True)
        PS_GA160_DataValidCheck = False
        Exit Function
    End If
    
    '구입여부 미선택 시
    If oForm01.Items("PchsYN").Specific.VALUE = "%" Then
        Call Sbo_Application.SetStatusBarMessage("구입여부가 선택되지 않았습니다.", bmt_Short, True)
        PS_GA160_DataValidCheck = False
        Exit Function
    End If
    
'    '라인정보 미입력 시
'    If oMat01.VisualRowCount = 1 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_GA160_DataValidCheck = False
'        Exit Function
'    End If
'
'    For i = 1 To oMat01.VisualRowCount - 1
'        If (oMat01.Columns("Code").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "코드는 필수입니다.", bmt_Short, True
'            oMat01.Columns("Code").Cells(i).Click ct_Regular
'            PS_GA160_DataValidCheck = False
'            Exit Function
'        End If
'
'        If (oMat01.Columns("CodeNm").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "코드명은 필수입니다.", bmt_Short, True
'            oMat01.Columns("CodeNm").Cells(i).Click ct_Regular
'            PS_GA160_DataValidCheck = False
'            Exit Function
'        End If
'    Next

'    Call oMat01.FlushToDataSource
'    Call oDS_PS_GA160L.RemoveRecord(oDS_PS_GA160L.Size - 1)
'    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_GA160_FormClear
    End If
    
    PS_GA160_DataValidCheck = True
    Exit Function
PS_GA160_DataValidCheck_Error:
    PS_GA160_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_GA160_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PS_GA160_MTX01()
On Error GoTo PS_GA160_MTX01_Error
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
    Param01 = Trim(oForm01.Items("Param01").Specific.VALUE)
    Param02 = Trim(oForm01.Items("Param01").Specific.VALUE)
    Param03 = Trim(oForm01.Items("Param01").Specific.VALUE)
    Param04 = Trim(oForm01.Items("Param01").Specific.VALUE)
    
    Query01 = "SELECT 10"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_GA160_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_GA160L.InsertRecord (i)
        End If
        oDS_PS_GA160L.Offset = i
        Call oDS_PS_GA160L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_GA160L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_GA160_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_GA160_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA160_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_GA160_FormResize()
On Error GoTo PS_GA160_FormResize_Error
    
    Exit Sub
PS_GA160_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA160_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_GA160_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_GA160_Validate_Error
    PS_GA160_Validate = True
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
PS_GA160_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_GA160_Validate_Error:
    PS_GA160_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_GA160_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_GA161_Open()
'******************************************************************************
'Function ID : PS_GA161_Open()
'해당모듈 : PS_GA160
'기능 : 사용자이력 창 호출
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA161_Open_Error

    Dim oTempClass As Variant
    Dim Seq As Integer

    Set oTempClass = New PS_GA161
    Seq = oForm01.Items("Code").Specific.VALUE
    
    Call oTempClass.LoadForm(Seq)

    Exit Sub

PS_GA161_Open_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA161_Open_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA162_Open()
'******************************************************************************
'Function ID : PS_GA162_Open()
'해당모듈 : PS_GA160
'기능 : 위치이력 창 호출
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA162_Open_Error

    Dim oTempClass As Variant
    Dim Seq As Integer

    Set oTempClass = New PS_GA162
    Seq = oForm01.Items("Code").Specific.VALUE
    
    Call oTempClass.LoadForm(Seq)

    Exit Sub

PS_GA162_Open_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA162_Open_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA163_Open()
'******************************************************************************
'Function ID : PS_GA163_Open()
'해당모듈 : PS_GA160
'기능 : 점검이력 창 호출
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA163_Open_Error

    Dim oTempClass As Variant
    Dim Seq As Integer

    Set oTempClass = New PS_GA163
    Seq = oForm01.Items("Code").Specific.VALUE
    
    Call oTempClass.LoadForm(Seq)

    Exit Sub

PS_GA163_Open_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA163_Open_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA164_Open()
'******************************************************************************
'Function ID : PS_GA164_Open()
'해당모듈 : PS_GA160
'기능 : HW이력 창 호출
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA164_Open_Error

    Dim oTempClass As Variant
    Dim Seq As Integer

    Set oTempClass = New PS_GA164
    Seq = oForm01.Items("Code").Specific.VALUE
    
    Call oTempClass.LoadForm(Seq)

    Exit Sub

PS_GA164_Open_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA164_Open_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA165_Open()
'******************************************************************************
'Function ID : PS_GA165_Open()
'해당모듈 : PS_GA160
'기능 : SW이력 창 호출
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA165_Open_Error

    Dim oTempClass As Variant
    Dim Seq As Integer

    Set oTempClass = New PS_GA165
    Seq = oForm01.Items("Code").Specific.VALUE
    
    Call oTempClass.LoadForm(Seq)

    Exit Sub

PS_GA165_Open_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA165_Open_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA166_Open()
'******************************************************************************
'Function ID : PS_GA166_Open()
'해당모듈 : PS_GA160
'기능 : IP이력 창 호출
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA166_Open_Error

    Dim oTempClass As Variant
    Dim Seq As Integer

    Set oTempClass = New PS_GA166
    Seq = oForm01.Items("Code").Specific.VALUE
    
    Call oTempClass.LoadForm(Seq)

    Exit Sub

PS_GA166_Open_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA166_Open_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
