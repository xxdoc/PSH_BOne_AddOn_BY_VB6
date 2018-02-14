VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM152"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_QM152.cls
'//  Module         : QM
'//  Desc           : 제안평가(QC)
'//  FormType       : PS_QM152
'//  Create Date    : 2014.03.14
'//  Copyright  (c) Morning Data
'//********************************************************************************

'//사용방법
'//Dim ChildForm01 As Variant
'//Set ChildForm01 = New PS_QM152
'//Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)

Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String
Public oBaseItmBsort01 As String

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
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String, Optional ByVal oItmBsort02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM152.srf")
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
    
    oFormUniqueID01 = "PS_QM152_" & GetTotalFormsCount
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
    
    Call PS_QM152_CreateItems
    Call PS_QM152_ComboBox_Setting
    Call PS_QM152_CF_ChooseFromList
    Call PS_QM152_FormItemEnabled
    Call PS_QM152_EnableMenus
    '//Call PS_QM152_FormClear '//UDO방식일때
    '//Call PS_QM152_AddMatrixRow(0, True) '//UDO방식일때
    
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.EnableMenu ("1282"), True
    
    oForm01.Visible = True
    
    '최초 커서위치
    oForm01.ActiveItem = "ym"
    
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
                 
                 Call PS_QM152_Form_ini
            
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_QM152_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_QM152_FormItemEnabled '//UDO방식
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
'        If pval.ItemUID = "Btn_ret" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_QM152_MTX01
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
'        If pval.ItemUID = "Button02" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                Call PS_QM152_SetBaseForm '//부모폼에입력
'                If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
'                    Call oForm01.Close
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
        
        If pval.ItemUID = "Btn_ret" Then
           Call PS_QM152_MTX01
        End If
               
        If pval.ItemUID = "Btn_save" Then
           Call PS_QM152_SAVE
        End If
                             
        
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_QM152_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM152" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_QM152_FormItemEnabled
'                    Call PS_QM152_FormClear '//UDO방식일때
'                    Call PS_QM152_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_QM152_FormItemEnabled
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
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "proposalno" Then
                If oForm01.Items("proposalno").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
    
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
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
        If pval.ItemUID = "Grid01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row > 0 Then
                    
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "Grid01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.Row >= 0 Then
                   PS_QM152_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
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
    Dim Mark       As Double
    Dim mark_1     As Double
    Dim mark_c     As Double
    Dim saup       As String
    Dim proposalno As String
    Dim oCombo     As SAPbouiCOM.ComboBox
    Dim oQuery     As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
          
          If pval.ItemChanged = True Then
            '제안번호
            If pval.ItemUID = "proposalno" Then
               saup = oForm01.Items("saup").Specific.VALUE
               proposalno = oForm01.Items("proposalno").Specific.VALUE
               oQuery = "EXEC PS_QM152_02 '" & saup & "', '" & proposalno & "'"
               oRecordSet.DoQuery oQuery
               
                oForm01.DataSources.UserDataSources.Item("ymd").VALUE = oRecordSet.Fields("ymd").VALUE
                oForm01.DataSources.UserDataSources.Item("sabun").VALUE = oRecordSet.Fields("sabun").VALUE
                oForm01.DataSources.UserDataSources.Item("kname").VALUE = oRecordSet.Fields("kname").VALUE
                oForm01.DataSources.UserDataSources.Item("proposalno").VALUE = oRecordSet.Fields("proposalno").VALUE
                oForm01.DataSources.UserDataSources.Item("buseonm").VALUE = oRecordSet.Fields("buseonm").VALUE
                oForm01.DataSources.UserDataSources.Item("sectnm").VALUE = oRecordSet.Fields("sectnm").VALUE
                oForm01.DataSources.UserDataSources.Item("staffnm").VALUE = oRecordSet.Fields("staffnm").VALUE
                oForm01.DataSources.UserDataSources.Item("wrkgrdnm").VALUE = oRecordSet.Fields("wrkgrdnm").VALUE
                oForm01.DataSources.UserDataSources.Item("title").VALUE = oRecordSet.Fields("title").VALUE
                oForm01.DataSources.UserDataSources.Item("sbuseo").VALUE = oRecordSet.Fields("sbuseo").VALUE
                oForm01.DataSources.UserDataSources.Item("ssect").VALUE = oRecordSet.Fields("ssect").VALUE
                oForm01.DataSources.UserDataSources.Item("sbuseonm").VALUE = oRecordSet.Fields("sbuseonm").VALUE
                oForm01.DataSources.UserDataSources.Item("ssectnm").VALUE = oRecordSet.Fields("ssectnm").VALUE
                oForm01.DataSources.UserDataSources.Item("isymd").VALUE = oRecordSet.Fields("isymd").VALUE
                oForm01.DataSources.UserDataSources.Item("effectamt").VALUE = oRecordSet.Fields("effectamt").VALUE
                oForm01.DataSources.UserDataSources.Item("mark").VALUE = oRecordSet.Fields("mark").VALUE
                oForm01.DataSources.UserDataSources.Item("grade").VALUE = oRecordSet.Fields("grade").VALUE
                oForm01.DataSources.UserDataSources.Item("par").VALUE = oRecordSet.Fields("par").VALUE
                oForm01.DataSources.UserDataSources.Item("prizeamt").VALUE = oRecordSet.Fields("prizeamt").VALUE
                
                oForm01.DataSources.UserDataSources.Item("mark_1").VALUE = oRecordSet.Fields("mark_1").VALUE
                oForm01.DataSources.UserDataSources.Item("grade_1").VALUE = oRecordSet.Fields("grade_1").VALUE
                oForm01.DataSources.UserDataSources.Item("par_1").VALUE = oRecordSet.Fields("par_1").VALUE
                oForm01.DataSources.UserDataSources.Item("prizeamt_1").VALUE = oRecordSet.Fields("prizeamt_1").VALUE
                
                oForm01.DataSources.UserDataSources.Item("mark_a").VALUE = oRecordSet.Fields("mark_a").VALUE
                oForm01.DataSources.UserDataSources.Item("grade_a").VALUE = oRecordSet.Fields("grade_a").VALUE
                oForm01.DataSources.UserDataSources.Item("par_a").VALUE = oRecordSet.Fields("par_a").VALUE
                oForm01.DataSources.UserDataSources.Item("prizeamt_a").VALUE = oRecordSet.Fields("prizeamt_a").VALUE
                
                Set oCombo = oForm01.Items("saup").Specific
                oCombo.Select oRecordSet.Fields("saup").VALUE, psk_ByValue
                Set oCombo = oForm01.Items("div").Specific
                oCombo.Select oRecordSet.Fields("div").VALUE, psk_ByValue
                Set oCombo = oForm01.Items("field").Specific
                oCombo.Select oRecordSet.Fields("field").VALUE, psk_ByValue
                
                Set oCombo = oForm01.Items("adoptdiv").Specific
                oCombo.Select oRecordSet.Fields("adoptdiv").VALUE, psk_ByValue
                Set oCombo = oForm01.Items("isdiv").Specific
                oCombo.Select oRecordSet.Fields("isdiv").VALUE, psk_ByValue
                
                '//커서위치
                oForm01.ActiveItem = "adoptdiv"
    
    
            End If
            
            '점수QC
            If pval.ItemUID = "mark_1" Then
               oQuery = "SELECT Grade, Par, PrizeAmt FROM [ZPS_QM150] WHERE '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "' BETWEEN MarkMin AND MarkMax AND BPLID = '" & oForm01.Items("saup").Specific.VALUE & "'"
               oRecordSet.DoQuery oQuery
               oForm01.Items("grade_1").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
               oForm01.Items("par_1").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)
               oForm01.Items("prizeamt_1").Specific.VALUE = Trim(oRecordSet.Fields(2).VALUE)
               
               Mark = oForm01.Items("mark").Specific.VALUE
               mark_1 = oForm01.Items("mark_1").Specific.VALUE
               mark_c = Round((Mark + mark_1) / 2, 0)
                                            
               oForm01.Items("mark_a").Specific.VALUE = mark_c
               oQuery = "SELECT Grade, Par, PrizeAmt FROM [ZPS_QM150] WHERE '" & mark_c & "' BETWEEN MarkMin AND MarkMax AND BPLID = '" & oForm01.Items("saup").Specific.VALUE & "'"
               oRecordSet.DoQuery oQuery
               oForm01.Items("grade_a").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
               oForm01.Items("par_a").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)
               oForm01.Items("prizeamt_a").Specific.VALUE = Trim(oRecordSet.Fields(2).VALUE)
               
            End If
            
          End If
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
        Call PS_QM152_FormItemEnabled
        '//Call PS_QM152_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
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
'        If (pval.ItemUID = "ItemCode") Then
'            Set oDataTable01 = pval.SelectedObjects
'            If oDataTable01 Is Nothing Then
'            Else
'                oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
'                oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
'            End If
'        End If
'        oForm01.Update
    End If
    Set oDataTable01 = Nothing
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
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
Raise_EVENT_GOT_FOCUS_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_GOT_FOCUS_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_FORM_UNLOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_FORM_UNLOAD_Error
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
        RemoveForms oFormUniqueID01
        Set oForm01 = Nothing
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
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
    '        For i = 1 To oMat01.VisualRowCount
    '            oMat01.Columns("COL01").Cells(i).Specific.Value = i
    '        Next i
    '        oMat01.FlushToDataSource
    '        Call oDS_PS_QM152L.RemoveRecord(oDS_PS_QM152L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_QM152_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_QM152_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_QM152_CreateItems() As Boolean
On Error GoTo PS_QM152_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    'oGrid01.SelectionMode = ms_NotSupported
    
    '접수일자
    Call oForm01.DataSources.UserDataSources.Add("ymd", dt_DATE)
    Call oForm01.Items("ymd").Specific.DataBind.SetBound(True, "", "ymd")
    oForm01.DataSources.UserDataSources.Item("ymd").VALUE = Format(Now, "YYYYMMDD")
    
    '사원번호
    Call oForm01.DataSources.UserDataSources.Add("sabun", dt_SHORT_TEXT, 10)
    oForm01.Items("sabun").Specific.DataBind.SetBound True, "", "sabun"
    Call oForm01.DataSources.UserDataSources.Add("kname", dt_SHORT_TEXT, 10)
    oForm01.Items("kname").Specific.DataBind.SetBound True, "", "kname"
    
    '제안번호
    Call oForm01.DataSources.UserDataSources.Add("proposalno", dt_SHORT_TEXT, 10)
    oForm01.Items("proposalno").Specific.DataBind.SetBound True, "", "proposalno"
    
    '소속
    Call oForm01.DataSources.UserDataSources.Add("buseonm", dt_SHORT_TEXT, 10)
    oForm01.Items("buseonm").Specific.DataBind.SetBound True, "", "buseonm"
    Call oForm01.DataSources.UserDataSources.Add("sectnm", dt_SHORT_TEXT, 10)
    oForm01.Items("sectnm").Specific.DataBind.SetBound True, "", "sectnm"
    Call oForm01.DataSources.UserDataSources.Add("staffnm", dt_SHORT_TEXT, 10)
    oForm01.Items("staffnm").Specific.DataBind.SetBound True, "", "staffnm"
    Call oForm01.DataSources.UserDataSources.Add("wrkgrdnm", dt_SHORT_TEXT, 10)
    oForm01.Items("wrkgrdnm").Specific.DataBind.SetBound True, "", "wrkgrdnm"
    
    '제목
    Call oForm01.DataSources.UserDataSources.Add("title", dt_SHORT_TEXT, 100)
    oForm01.Items("title").Specific.DataBind.SetBound True, "", "title"
    
    '실시부서,담당
    Call oForm01.DataSources.UserDataSources.Add("sbuseo", dt_SHORT_TEXT, 10)
    oForm01.Items("sbuseo").Specific.DataBind.SetBound True, "", "sbuseo"
    Call oForm01.DataSources.UserDataSources.Add("ssect", dt_SHORT_TEXT, 10)
    oForm01.Items("ssect").Specific.DataBind.SetBound True, "", "ssect"
    
    Call oForm01.DataSources.UserDataSources.Add("sbuseonm", dt_SHORT_TEXT, 10)
    oForm01.Items("sbuseonm").Specific.DataBind.SetBound True, "", "sbuseonm"
    Call oForm01.DataSources.UserDataSources.Add("ssectnm", dt_SHORT_TEXT, 10)
    oForm01.Items("ssectnm").Specific.DataBind.SetBound True, "", "ssectnm"
    
    '시행일자
    Call oForm01.DataSources.UserDataSources.Add("isymd", dt_DATE)
    Call oForm01.Items("isymd").Specific.DataBind.SetBound(True, "", "isymd")
    
    '효과금액
    Call oForm01.DataSources.UserDataSources.Add("effectamt", dt_SUM)
    oForm01.Items("effectamt").Specific.DataBind.SetBound True, "", "effectamt"
    
    '점수
    Call oForm01.DataSources.UserDataSources.Add("mark", dt_PRICE)
    oForm01.Items("mark").Specific.DataBind.SetBound True, "", "mark"
    
    '등급
    Call oForm01.DataSources.UserDataSources.Add("grade", dt_SHORT_TEXT, 10)
    oForm01.Items("grade").Specific.DataBind.SetBound True, "", "grade"
    
    '평가점
    Call oForm01.DataSources.UserDataSources.Add("par", dt_PRICE)
    oForm01.Items("par").Specific.DataBind.SetBound True, "", "par"
    
    '시상액
    Call oForm01.DataSources.UserDataSources.Add("prizeamt", dt_SUM)
    oForm01.Items("prizeamt").Specific.DataBind.SetBound True, "", "prizeamt"
    
    Call oForm01.DataSources.UserDataSources.Add("mark_1", dt_PRICE)
    oForm01.Items("mark_1").Specific.DataBind.SetBound True, "", "mark_1"
    
    Call oForm01.DataSources.UserDataSources.Add("grade_1", dt_SHORT_TEXT, 10)
    oForm01.Items("grade_1").Specific.DataBind.SetBound True, "", "grade_1"
    
    Call oForm01.DataSources.UserDataSources.Add("par_1", dt_PRICE)
    oForm01.Items("par_1").Specific.DataBind.SetBound True, "", "par_1"
    
    Call oForm01.DataSources.UserDataSources.Add("prizeamt_1", dt_SUM)
    oForm01.Items("prizeamt_1").Specific.DataBind.SetBound True, "", "prizeamt_1"
    
    Call oForm01.DataSources.UserDataSources.Add("mark_a", dt_PRICE)
    oForm01.Items("mark_a").Specific.DataBind.SetBound True, "", "mark_a"
    
    Call oForm01.DataSources.UserDataSources.Add("grade_a", dt_SHORT_TEXT, 10)
    oForm01.Items("grade_a").Specific.DataBind.SetBound True, "", "grade_a"
    
    Call oForm01.DataSources.UserDataSources.Add("par_a", dt_PRICE)
    oForm01.Items("par_a").Specific.DataBind.SetBound True, "", "par_a"
    
    Call oForm01.DataSources.UserDataSources.Add("prizeamt_a", dt_SUM)
    oForm01.Items("prizeamt_a").Specific.DataBind.SetBound True, "", "prizeamt_a"
    
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function

PS_QM152_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM152_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_QM152_ComboBox_Setting()
On Error GoTo PS_QM152_ComboBox_Setting_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '사업장콤보박스세팅
    Call MDC_SetMod.Set_ComboList(oForm01.Items("saup").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", False, False)
    Set oCombo = oForm01.Items("saup").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '콤보에 기본값설정
    '제안구분
    Call oForm01.Items("div").Specific.ValidValues.Add("0", "정식")
    Call oForm01.Items("div").Specific.ValidValues.Add("1", "약식")
    Call oForm01.Items("div").Specific.ValidValues.Add("2", "등외")
    oForm01.Items("div").Specific.Select "0", psk_Index
    
    '분야
    'oForm01.Items("field").Specific.ValidValues.Add "", ""
    Call MDC_SetMod.Set_ComboList(oForm01.Items("field").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q015' order by U_Minor", "", False, False)
    oForm01.Items("field").Specific.Select "0", psk_Index
    
    '채택구분
    Call oForm01.Items("adoptdiv").Specific.ValidValues.Add("Y", "채택")
    Call oForm01.Items("adoptdiv").Specific.ValidValues.Add("N", "불채택")
    oForm01.Items("adoptdiv").Specific.Select "0", psk_Index
    
    '실시구분
    Call oForm01.Items("isdiv").Specific.ValidValues.Add("N", "미실시")
    Call oForm01.Items("isdiv").Specific.ValidValues.Add("Y", "실시")
    Call oForm01.Items("isdiv").Specific.ValidValues.Add("G", "검토")
    oForm01.Items("isdiv").Specific.Select "0", psk_Index
    
    Call oForm01.Freeze(False)
    
    
    
    Exit Sub
PS_QM152_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM152_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM152_CF_ChooseFromList()
On Error GoTo PS_QM152_CF_ChooseFromList_Error
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
PS_QM152_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM152_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM152_FormItemEnabled()
On Error GoTo PS_QM152_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        If oBaseItmBsort01 = "" Then
        
        Else
            Call oForm01.Items("ItmBsort").Specific.Select(oBaseItmBsort01, psk_ByValue)
            oForm01.Items("ItmBsort").Enabled = False
        End If
        '//각모드에따른 아이템설정
        '//Call PS_QM152_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM152_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM152_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM152_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM152_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_QM152L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_QM152L.Offset = oRow
'    oDS_PS_QM152L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM152_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM152_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM152_FormClear()
On Error GoTo PS_QM152_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM152'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_QM152_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM152_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM152_EnableMenus()
On Error GoTo PS_QM152_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_QM152_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM152_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_QM152_DataValidCheck() As Boolean
On Error GoTo PS_QM152_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_QM152_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_QM152_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_QM152_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_QM152_FormClear
    Exit Function
PS_QM152_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM152_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_QM152_MTX01()
On Error GoTo PS_QM152_MTX01_Error
    '//그리드에 데이터 로드 (조회)
    Call oForm01.Freeze(True)
    Dim i           As Long
    Dim sQry        As String
    Dim saup        As String
    Dim YM          As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    saup = oForm01.Items("saup").Specific.VALUE
    YM = oForm01.Items("ym").Specific.VALUE
    
    sQry = " EXEC [PS_QM152_01]  '" & saup & "', '" & YM & "'"
    
    oGrid01.DataTable.Clear
    
    Call oForm01.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")
    
    Call RecordSet01.DoQuery(sQry)
    If RecordSet01.RecordCount = 0 Then
       Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다. 사업장,년월을 확인 하세요", "E")
       GoTo PS_QM152_MTX01_Exit
    End If
        
    oGrid01.AutoResizeColumns
    oForm01.Update
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM152_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM152_MTX01_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM152_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PS_QM152_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    
    '//그리드 자료를 head에 로드
  
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_QM152_MTX02_Error

    Call oForm01.Freeze(True)
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Param01 = oGrid01.DataTable.Columns("사업장").Cells(oRow).VALUE
    Param02 = oGrid01.DataTable.Columns("제안번호").Cells(oRow).VALUE
    
    
    sQry = "EXEC PS_QM152_02 '" & Param01 & "', '" & Param02 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        Call PS_QM152_Form_ini
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_QM152_MTX02_Exit
    End If
    
    oForm01.DataSources.UserDataSources.Item("ymd").VALUE = oRecordSet.Fields("ymd").VALUE
    oForm01.DataSources.UserDataSources.Item("sabun").VALUE = oRecordSet.Fields("sabun").VALUE
    oForm01.DataSources.UserDataSources.Item("kname").VALUE = oRecordSet.Fields("kname").VALUE
    oForm01.DataSources.UserDataSources.Item("proposalno").VALUE = oRecordSet.Fields("proposalno").VALUE
    oForm01.DataSources.UserDataSources.Item("buseonm").VALUE = oRecordSet.Fields("buseonm").VALUE
    oForm01.DataSources.UserDataSources.Item("sectnm").VALUE = oRecordSet.Fields("sectnm").VALUE
    oForm01.DataSources.UserDataSources.Item("staffnm").VALUE = oRecordSet.Fields("staffnm").VALUE
    oForm01.DataSources.UserDataSources.Item("wrkgrdnm").VALUE = oRecordSet.Fields("wrkgrdnm").VALUE
    oForm01.DataSources.UserDataSources.Item("title").VALUE = oRecordSet.Fields("title").VALUE
    oForm01.DataSources.UserDataSources.Item("sbuseo").VALUE = oRecordSet.Fields("sbuseo").VALUE
    oForm01.DataSources.UserDataSources.Item("ssect").VALUE = oRecordSet.Fields("ssect").VALUE
    oForm01.DataSources.UserDataSources.Item("sbuseonm").VALUE = oRecordSet.Fields("sbuseonm").VALUE
    oForm01.DataSources.UserDataSources.Item("ssectnm").VALUE = oRecordSet.Fields("ssectnm").VALUE
    oForm01.DataSources.UserDataSources.Item("isymd").VALUE = oRecordSet.Fields("isymd").VALUE
    oForm01.DataSources.UserDataSources.Item("effectamt").VALUE = oRecordSet.Fields("effectamt").VALUE
    oForm01.DataSources.UserDataSources.Item("mark").VALUE = oRecordSet.Fields("mark").VALUE
    oForm01.DataSources.UserDataSources.Item("grade").VALUE = oRecordSet.Fields("grade").VALUE
    oForm01.DataSources.UserDataSources.Item("par").VALUE = oRecordSet.Fields("par").VALUE
    oForm01.DataSources.UserDataSources.Item("prizeamt").VALUE = oRecordSet.Fields("prizeamt").VALUE
    
    oForm01.DataSources.UserDataSources.Item("mark_1").VALUE = oRecordSet.Fields("mark_1").VALUE
    oForm01.DataSources.UserDataSources.Item("grade_1").VALUE = oRecordSet.Fields("grade_1").VALUE
    oForm01.DataSources.UserDataSources.Item("par_1").VALUE = oRecordSet.Fields("par_1").VALUE
    oForm01.DataSources.UserDataSources.Item("prizeamt_1").VALUE = oRecordSet.Fields("prizeamt_1").VALUE
    
    oForm01.DataSources.UserDataSources.Item("mark_a").VALUE = oRecordSet.Fields("mark_a").VALUE
    oForm01.DataSources.UserDataSources.Item("grade_a").VALUE = oRecordSet.Fields("grade_a").VALUE
    oForm01.DataSources.UserDataSources.Item("par_a").VALUE = oRecordSet.Fields("par_a").VALUE
    oForm01.DataSources.UserDataSources.Item("prizeamt_a").VALUE = oRecordSet.Fields("prizeamt_a").VALUE

    Set oCombo = oForm01.Items("saup").Specific
    oCombo.Select oRecordSet.Fields("saup").VALUE, psk_ByValue
    Set oCombo = oForm01.Items("div").Specific
    oCombo.Select oRecordSet.Fields("div").VALUE, psk_ByValue
    Set oCombo = oForm01.Items("field").Specific
    oCombo.Select oRecordSet.Fields("field").VALUE, psk_ByValue
    
    Set oCombo = oForm01.Items("adoptdiv").Specific
    oCombo.Select oRecordSet.Fields("adoptdiv").VALUE, psk_ByValue
    Set oCombo = oForm01.Items("isdiv").Specific
    oCombo.Select oRecordSet.Fields("isdiv").VALUE, psk_ByValue
    
    '//Key Disable
    oForm01.ActiveItem = "adoptdiv"
    
    oForm01.Items("proposalno").Enabled = False
'    oForm01.Items("Grade").Enabled = False
'    oForm01.Items("Num").Enabled = False
'
    oForm01.Update
    
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM152_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_QM152_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM152_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PS_QM152_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim Seqncom As Integer
    Dim sQry    As String
    Dim YM      As String
    Dim saup$, proposalno$, adoptdiv$, isdiv$, grade_1$, grade_a$
    Dim mark_1, Par_1, PrizeAmt_1, mark_a, Par_a, PrizeAmt_a  As Double
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    On Error GoTo PS_QM152_SAVE_Error
    
    saup = oForm01.Items("saup").Specific.VALUE
    proposalno = oForm01.Items("proposalno").Specific.VALUE
    adoptdiv = oForm01.Items("adoptdiv").Specific.VALUE
    isdiv = oForm01.Items("isdiv").Specific.VALUE
    mark_1 = oForm01.Items("mark_1").Specific.VALUE
    grade_1 = oForm01.Items("grade_1").Specific.VALUE
    Par_1 = oForm01.Items("par_1").Specific.VALUE
    PrizeAmt_1 = oForm01.Items("prizeamt_1").Specific.VALUE
    mark_a = oForm01.Items("mark_a").Specific.VALUE
    grade_a = oForm01.Items("grade_a").Specific.VALUE
    Par_a = oForm01.Items("par_a").Specific.VALUE
    PrizeAmt_a = oForm01.Items("prizeamt_a").Specific.VALUE
        
    Call oForm01.Freeze(True)
    
    '//갱신
    sQry = "Update [ZPS_QM151] set "
    sQry = sQry + "adoptdiv = '" & adoptdiv & "',"
    sQry = sQry + "isdiv = '" & isdiv & "',"
    sQry = sQry + "mark_1 = '" & mark_1 & "',"
    sQry = sQry + "grade_1 = '" & grade_1 & "',"
    sQry = sQry + "par_1 = '" & Par_1 & "',"
    sQry = sQry + "prizeamt_1 = '" & PrizeAmt_1 & "',"
    sQry = sQry + "mark_a = '" & mark_a & "',"
    sQry = sQry + "grade_a = '" & grade_a & "',"
    sQry = sQry + "par_a = '" & Par_a & "',"
    sQry = sQry + "prizeamt_a = '" & PrizeAmt_a & "'"
    
    sQry = sQry + " Where saup = '" & saup & "' And proposalno = '" & proposalno & "'"
    
    oRecordSet.DoQuery sQry
    
    
    Call PS_QM152_FormItemEnabled
    
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)

    Call PS_QM152_MTX01
    Call PS_QM152_Form_ini
    
    Exit Sub
    
PS_QM152_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_QM152_SAVE_Error:
    Call oForm01.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_QM152_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM152_Form_ini()
    Dim oCombo      As SAPbouiCOM.ComboBox

    '화면청소 ㅋ
    oForm01.DataSources.UserDataSources.Item("ymd").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("sabun").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("kname").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("proposalno").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("buseonm").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("sectnm").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("staffnm").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("wrkgrdnm").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("sbuseo").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("ssect").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("sbuseonm").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("ssectnm").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("isymd").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("effectamt").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("mark").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("grade").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("par").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("prizeamt").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("mark_1").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("grade_1").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("par_1").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("prizeamt_1").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("mark_a").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("grade_a").VALUE = ""
    oForm01.DataSources.UserDataSources.Item("par_a").VALUE = 0
    oForm01.DataSources.UserDataSources.Item("prizeamt_a").VALUE = 0
    Set oCombo = oForm01.Items("div").Specific
    oCombo.Select 0, psk_Index
    Set oCombo = oForm01.Items("field").Specific
    oCombo.Select 0, psk_Index
    oForm01.Items("title").Specific.VALUE = ""
    Set oCombo = oForm01.Items("adoptdiv").Specific
    oCombo.Select 0, psk_Index
    Set oCombo = oForm01.Items("isdiv").Specific
    oCombo.Select 0, psk_Index
    
    oForm01.ActiveItem = "proposalno"
    oForm01.Update

End Sub





