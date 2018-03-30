VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP084"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP084.cls
'//  Module         : 생산관리 > 분말생산완료등록
'//  Desc           : 분말생산완료등록
'//  FormType       : PS_PP084
'//  Create Date    : 2014.09.14
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix

Private oDS_PS_PP084H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP084L As SAPbouiCOM.DBDataSource '등록라인
Private oDS_PS_PP0841L As SAPbouiCOM.DBDataSource '

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    OrdGbn As String '//작지구분
    PP030HNo As String
    PP030MNo As String
    ItemCode As String
    BatchNum As String
    Quantity As Double
    WhsCode As String
    LineNum As Long
    'AcctCode As String
End Type
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

Private oDocEntry As String
Private oStatus   As String
Private oCanceled   As String

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP084.srf")
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
    
    oFormUniqueID01 = "PS_PP084_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP084_CreateItems
    Call PS_PP084_ComboBox_Setting
    Call PS_PP084_CF_ChooseFromList
    Call PS_PP084_EnableMenus
    Call PS_PP084_SetDocument(oFromDocEntry01)
    Call PS_PP084_FormResize
    
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
    Dim i&, RowCounter&
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
'                If oForm01.Mode = fm_OK_MODE Then
                    For i = 1 To oMat01.VisualRowCount - 1
                        If oMat01.Columns("Check").Cells(i).Specific.Checked = True And Trim(oMat01.Columns("OIGENum").Cells(i).Specific.VALUE) = "" Then
                            RowCounter = RowCounter + 1
                        End If
                    Next
                    If RowCounter = 0 Then
                        MDC_Com.MDC_GF_Message "취소할 항목을 선택해주세요.", "W"
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If (PS_PP084_Validate("취소") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If Sbo_Application.MessageBox("정말로 취소하시겠습니까?", "1", "예", "아니오") <> "1" Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    If PS_PP084_DI_API03 = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
'                    If PS_PP084_DI_API02 = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
                    oDocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
'                Else
'                    MDC_Com.MDC_GF_Message "현재 모드에서는 취소할수 없습니다.", "W"
'                    BubbleEvent = False
'                    Exit Sub
'                End If
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
                Dim sQry            As String
                Dim oRecordSet01      As SAPbobsCOM.Recordset
                Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                
                sQry = "Select Min(IsNULL(U_OIGENum, '')) From [@PS_PP080L] where DocEntry = '" & oDocEntry & "'"
                oRecordSet01.DoQuery sQry
                
                If Trim(oRecordSet01.Fields(0).VALUE) = "" Then
                    oStatus = "O"
                    oCanceled = "N"
                Else
                    oStatus = "C"
                    oCanceled = "Y"
                End If
                
                Call MDC_PS_Common.DoQuery("UPDATE [@PS_PP080H] SET Status = '" & oStatus & "', Canceled = '" & oCanceled & "' WHERE DocEntry = '" & oDocEntry & "'")
                
                oForm01.Mode = fm_FIND_MODE
                Call PS_PP084_FormItemEnabled
                oForm01.Items("DocEntry").Specific.VALUE = oDocEntry
                oForm01.Items("1").Click ct_Regular
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_PP084_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_PP084_FormItemEnabled '//UDO방식
                Call PS_PP084_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_PP084_FormItemEnabled
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
    Dim BoxCnt As Integer
    Dim i As Integer
    Dim PP030No As String
                
    If pval.BeforeAction = True Then
        Dim oTempClass As Variant
        If pval.ItemUID = "Link01" Then
            Set oTempClass = New PS_PP040
            Call oTempClass.LoadForm(oForm01.Items("PP040No").Specific.VALUE)
            BubbleEvent = False
            Exit Sub
        End If
        If pval.ItemUID = "PS_PP084" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Btn1" Then
            If oForm01.Items("OrdGbn").Specific.VALUE = "601" Then
                If oForm01.Mode = fm_ADD_MODE Then
                    PP030No = oForm01.Items("PP030No").Specific.VALUE
                    BoxCnt = oForm01.Items("BoxCnt").Specific.VALUE
                    For i = 1 To BoxCnt
                        oMat01.Columns("PP030No").Cells(i).Specific.VALUE = PP030No
                        If i <> BoxCnt Then
                            Call PS_PP084_AddMatrixRow(oMat01.VisualRowCount)
                        End If
                    Next
                    
                ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                ElseIf oForm01.Mode = fm_OK_MODE Then
                End If
            End If
        End If
        
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_PP084_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                If PS_PP084_DI_API01 = False Then
                    Call PS_PP084_AddMatrixRow(oMat01.VisualRowCount)
                    BubbleEvent = False
                    Exit Sub
                Else

                End If
                oDocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_PP084_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP084" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call MDC_PS_Common.DoQuery("EXEC PS_PP084_03 '" & oDocEntry & "'")
                    Call PS_PP084_FormItemEnabled
                    Call PS_PP084_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_PP084_FormItemEnabled
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "PP030No", "") '//사용자값활성
        If pval.ItemUID = "PP030No" Then
            If oForm01.Items("BPLId").Specific.Selected.VALUE = "선택" Then
                Call MDC_Com.MDC_GF_Message("사업장은 필수입니다.", "W")
                BubbleEvent = False
                Exit Sub
            ElseIf oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then
                Call MDC_Com.MDC_GF_Message("작업구분은 필수입니다.", "W")
                BubbleEvent = False
                Exit Sub
            End If
        End If
'        If pval.ColUID = "PP030No" Then
'            If oForm01.Items("BPLId").Specific.Selected.VALUE = "선택" Then
'                Call MDC_Com.MDC_GF_Message("사업장은 필수입니다.", "W")
'                BubbleEvent = False
'                Exit Sub
'            ElseIf oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then
'                Call MDC_Com.MDC_GF_Message("작업구분은 필수입니다.", "W")
'                BubbleEvent = False
'                Exit Sub
'            Else
'            '/////////////////////////////////////////////////////////
'               If pval.ColUID = "PP030No" And oForm01.Items("BPLId").Specific.Selected.VALUE = "1" And oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Then
'                    '//창원 휘팅일때
'                    Dim ChildForm01 As Variant
'                    Set ChildForm01 = New PS_PP071 '작지조회
'                    Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, Trim(oForm01.Items("BPLId").Specific.Selected.VALUE), Trim(oForm01.Items("OrdGbn").Specific.Selected.VALUE))
'                    Set ChildForm01 = Nothing
'                Else
'                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "PP030No") '//사용자값활성
'                End If
'            '///////////////////////////////////////////////////////
'
''                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "PP030No") '//사용자값활성
'            End If
'        End If
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
            Call oForm01.Freeze(True)
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "특정컬럼") Then
                    '//기타작업
                    Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_PP084_AddMatrixRow (pval.Row)
                    End If
                Else
                    Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                End If
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_PP084H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                ElseIf (pval.ItemUID = "BPLId") Then
                    Call oDS_PS_PP084H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    Call PS_PP084_AddMatrixRow(0, True)
                ElseIf (pval.ItemUID = "OrdGbn") Then
                    Call oDS_PS_PP084H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    Call PS_PP084_AddMatrixRow(0, True)
                    
                    If oForm01.Items(pval.ItemUID).Specific.Selected.VALUE = "104" Then '//멀티의경우
                        oMat01.Columns("BWeight").Visible = False
                        oMat01.Columns("PWeight").Visible = False
                        oMat01.Columns("YWeight").Visible = False
                        oMat01.Columns("NWeight").Visible = False
                    Else '//그외의경우
                        oMat01.Columns("BWeight").Visible = True
                        oMat01.Columns("PWeight").Visible = True
                        oMat01.Columns("YWeight").Visible = True
                        oMat01.Columns("NWeight").Visible = True
                    End If
                    If oForm01.Items("OrdGbn").Specific.VALUE = "111" Then
                        oForm01.Items("PP030No").Enabled = False
                        oForm01.Items("PQty").Enabled = False
                        oForm01.Items("BoxCnt").Enabled = False
                        oForm01.Items("Btn1").Enabled = False
                        
                        Call LoadData '분말포장대기 자료 SELECT
                    ElseIf oForm01.Items("OrdGbn").Specific.VALUE = "601" Then
                        oForm01.Items("BoxCnt").Specific.VALUE = 1
                        
                        oForm01.Items("PP030No").Enabled = True
                        oForm01.Items("PQty").Enabled = True
                        oForm01.Items("BoxCnt").Enabled = True
                        oForm01.Items("Btn1").Enabled = True

                    Else
                        oForm01.Items("PP030No").Enabled = False
                        oForm01.Items("PQty").Enabled = False
                        oForm01.Items("BoxCnt").Enabled = False
                        oForm01.Items("Btn1").Enabled = False
                        
                        oMat02.Clear
                        oMat02.FlushToDataSource
                        oMat02.LoadFromDataSource
                    End If
                Else
                    Call oDS_PS_PP084H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oForm01.Update
            If pval.ItemUID = "Mat01" Then
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Collapsed
            Else
                
            End If
            Call oForm01.Freeze(False)
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
Dim Query01 As String
Dim RecordSet01 As SAPbobsCOM.Recordset

    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)
                If oMat01.Columns("OrdGbn").Cells(pval.Row).Specific.VALUE = "105" Or oMat01.Columns("OrdGbn").Cells(pval.Row).Specific.VALUE = "106" Then
                    
                    Dim ProgBar01 As SAPbouiCOM.ProgressBar
                    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
                    
                                       
    
                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    
                    
                    ProgBar01.VALUE = 100
                    ProgBar01.Stop
                    Set ProgBar01 = Nothing

                End If
            End If
        End If
        If pval.ItemUID = "Mat02" Then
            If pval.Row > 0 Then
                Call oMat02.SelectRow(pval.Row, True, False)
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    Dim i&, j&
    
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If pval.ItemUID = "Mat01" And pval.Row = 0 And pval.ColUID = "Check" Then
            Dim Check$
            oForm01.Freeze True
            oMat01.FlushToDataSource
            If Trim(oDS_PS_PP084L.GetValue("U_Check", 0)) = "" Or Trim(oDS_PS_PP084L.GetValue("U_Check", 0)) = "N" Then
                Check = "Y"
            ElseIf Trim(oDS_PS_PP084L.GetValue("U_Check", 0)) = "Y" Then
                Check = "N"
            End If
            For i = 0 To oMat01.VisualRowCount - 2
                If Trim(oDS_PS_PP084L.GetValue("U_OIGENum", 0)) <> "" Then
                    oDS_PS_PP084L.setValue "U_Check", i, "Y"
                Else
                    oDS_PS_PP084L.setValue "U_Check", i, Check
                End If
            Next i
            oMat01.LoadFromDataSource
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Mat02" And pval.Row <> "0" And oForm01.Mode = fm_ADD_MODE And Trim(oDS_PS_PP0841L.GetValue("U_ColReg04", pval.Row - 1)) <> "" Then
            oForm01.Freeze True

            If oMat01.VisualRowCount = 0 Then
                oDS_PS_PP084L.Clear
            End If


            j = 0
            For i = 0 To oMat01.VisualRowCount - 1
'                        oMat02.LoadFromDataSource
                If Trim(oDS_PS_PP084L.GetValue("U_PP030No", i)) = Trim(oDS_PS_PP0841L.GetValue("U_ColReg01", pval.Row - 1)) And _
                   Trim(oDS_PS_PP084L.GetValue("U_BatchNum", i)) = Trim(oDS_PS_PP0841L.GetValue("U_ColReg04", pval.Row - 1)) Then
                    MDC_Com.MDC_GF_Message "같은 행을 두번 선택할 수 없습니다. 확인하세요." & Err.Number & " - " & Err.Description, "W"
                    j = 1
                End If
            Next

            If j = 0 Then
                oMat01.Columns("BatchNum").Cells(oMat01.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP0841L.GetValue("U_ColReg04", pval.Row - 1))
                oMat01.Columns("PQty").Cells(oMat01.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP0841L.GetValue("U_ColQty01", pval.Row - 1))
                oMat01.Columns("PP030No").Cells(oMat01.VisualRowCount).Specific.VALUE = Trim(oDS_PS_PP0841L.GetValue("U_ColReg01", pval.Row - 1))
                
                
'                Call oDS_PS_PP084L.setValue("U_PP030No", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColReg01", pval.Row - 1)))
'                Call oDS_PS_PP084L.setValue("U_OrdGbn", oMat01.VisualRowCount - 1, oForm01.Items("OrdGbn").Specific.VALUE)
'                Call oDS_PS_PP084L.setValue("U_OrdNum", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColReg05", pval.Row - 1)))
'                Call oDS_PS_PP084L.setValue("U_OrdSub1", oMat01.VisualRowCount - 1, "00")
'                Call oDS_PS_PP084L.setValue("U_OrdSub2", oMat01.VisualRowCount - 1, "000")
'                Call oDS_PS_PP084L.setValue("U_ORDRNo", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColReg06", pval.Row - 1)))
'                Call oDS_PS_PP084L.setValue("U_RDR1No", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColReg07", pval.Row - 1)))
'                Call oDS_PS_PP084L.setValue("U_BPLId", oMat01.VisualRowCount - 1, oForm01.Items("BPLId").Specific.VALUE)
'
'                Call oDS_PS_PP084L.setValue("U_PP030HNo", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColReg01", pval.Row - 1)))
'                Call oDS_PS_PP084L.setValue("U_PP030MNo", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColReg09", pval.Row - 1)))
'
'                Call oDS_PS_PP084L.setValue("U_ItemCode", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColReg04", pval.Row - 1)))
'                Call oDS_PS_PP084L.setValue("U_ItemName", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColReg02", pval.Row - 1)))
'                Call oDS_PS_PP084L.setValue("U_WhsCode", oMat01.VisualRowCount - 1, "101")
'                Call oDS_PS_PP084L.setValue("U_WhsName", oMat01.VisualRowCount - 1, "창원")
'                Call oDS_PS_PP084L.setValue("U_BatchNum", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColReg03", pval.Row - 1)))
'
'                If Trim(oDS_PS_PP0841L.GetValue("U_ColQty01", pval.Row - 1)) > Trim(oDS_PS_PP0841L.GetValue("U_ColQty02", pval.Row - 1)) Then
'                    Call oDS_PS_PP0841L.setValue("U_BQty", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColQty02", pval.Row - 1)))
'                    Call oDS_PS_PP084L.setValue("U_BWeight", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColQty02", pval.Row - 1)))
'                Else
'                    Call oDS_PS_PP084L.setValue("U_BQty", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColQty01", pval.Row - 1)))
'                    Call oDS_PS_PP084L.setValue("U_BWeight", oMat01.VisualRowCount - 1, Trim(oDS_PS_PP0841L.GetValue("U_ColQty01", pval.Row - 1)))
'                End If
'
'
'
'
'                Call oDS_PS_PP084L.setValue("U_LineId", oMat01.VisualRowCount - 1, oMat01.VisualRowCount)
'
                
'                Call PS_PP084_AddMatrixRow(oMat01.VisualRowCount, False)
                
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                j = 0
            End If
                    
            oForm01.Freeze False
            BubbleEvent = False
        End If
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
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim Weight As Double
    
    Dim SumQty As Currency
    
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "PP030No") Then
'                    For i = 1 To oMat01.RowCount
'                        '//현재 입력한 값이 이미 입력되어 있는경우
'                        If oMat01.Columns("PP030No").Cells(i).Specific.VALUE = oMat01.Columns("PP030No").Cells(pval.Row).Specific.VALUE And i <> pval.Row Then
'                            Call MDC_Com.MDC_GF_Message("이미 입력한 작업지시문서입니다.", "W")
'                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
'                            GoTo Continue
'                        End If
'                    Next
                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    
                    Query01 = "EXEC PS_PP084_02 '" & oMat01.Columns("PP030No").Cells(pval.Row).Specific.VALUE & "','" & oForm01.Items("OrdGbn").Specific.Selected.VALUE & "'"
                    
                    Call RecordSet01.DoQuery(Query01)
                    If RecordSet01.RecordCount = 0 Then
                        Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                    Else
                        Call oDS_PS_PP084L.setValue("U_PP030No", pval.Row - 1, RecordSet01.Fields("PP030No").VALUE)
                        Call oDS_PS_PP084L.setValue("U_OrdGbn", pval.Row - 1, RecordSet01.Fields("OrdGbn").VALUE)
                        Call oDS_PS_PP084L.setValue("U_OrdNum", pval.Row - 1, RecordSet01.Fields("OrdNum").VALUE)
                        Call oDS_PS_PP084L.setValue("U_OrdSub1", pval.Row - 1, RecordSet01.Fields("OrdSub1").VALUE)
                        Call oDS_PS_PP084L.setValue("U_OrdSub2", pval.Row - 1, RecordSet01.Fields("OrdSub2").VALUE)
                        Call oDS_PS_PP084L.setValue("U_PP030HNo", pval.Row - 1, RecordSet01.Fields("PP030HNo").VALUE)
                        Call oDS_PS_PP084L.setValue("U_PP030MNo", pval.Row - 1, RecordSet01.Fields("PP030MNo").VALUE)
                        Call oDS_PS_PP084L.setValue("U_ORDRNo", pval.Row - 1, RecordSet01.Fields("ORDRNo").VALUE)
                        Call oDS_PS_PP084L.setValue("U_RDR1No", pval.Row - 1, RecordSet01.Fields("RDR1No").VALUE)
                        Call oDS_PS_PP084L.setValue("U_BPLId", pval.Row - 1, RecordSet01.Fields("BPLId").VALUE)
                        Call oDS_PS_PP084L.setValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                        Call oDS_PS_PP084L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
                        Call oDS_PS_PP084L.setValue("U_CpCode", pval.Row - 1, RecordSet01.Fields("CpCode").VALUE)
                        Call oDS_PS_PP084L.setValue("U_CpName", pval.Row - 1, RecordSet01.Fields("CpName").VALUE)
                        Call oDS_PS_PP084L.setValue("U_BQty", pval.Row - 1, Val(RecordSet01.Fields("BQty").VALUE))
                        Call oDS_PS_PP084L.setValue("U_BWeight", pval.Row - 1, Val(RecordSet01.Fields("BWeight").VALUE))
'                        Call oDS_PS_PP084L.setValue("U_PQty", pval.Row - 1, Val(RecordSet01.Fields("PQty").VALUE))
                        Call oDS_PS_PP084L.setValue("U_PWeight", pval.Row - 1, Val(RecordSet01.Fields("PWeight").VALUE))
'                        Call oDS_PS_PP084L.setValue("U_YQty", pval.Row - 1, Val(RecordSet01.Fields("YQty").VALUE))
                        Call oDS_PS_PP084L.setValue("U_YWeight", pval.Row - 1, Val(RecordSet01.Fields("YWeight").VALUE))
                        Call oDS_PS_PP084L.setValue("U_NQty", pval.Row - 1, Val(RecordSet01.Fields("NQty").VALUE))
                        Call oDS_PS_PP084L.setValue("U_NWeight", pval.Row - 1, Val(RecordSet01.Fields("NWeight").VALUE))
                        Call oDS_PS_PP084L.setValue("U_WhsCode", pval.Row - 1, RecordSet01.Fields("WhsCode").VALUE)
                        Call oDS_PS_PP084L.setValue("U_WhsName", pval.Row - 1, RecordSet01.Fields("WhsName").VALUE)
'                        Call oDS_PS_PP084L.setValue("U_BatchNum", pval.Row - 1, RecordSet01.Fields("BatchNum").VALUE)
                        Call oDS_PS_PP084L.setValue("U_LineId", pval.Row - 1, RecordSet01.Fields("LineId").VALUE)
                        
                        oMat01.LoadFromDataSource
                        
                        '//합격수량 sum
                        For i = 0 To oMat01.VisualRowCount - 1
                            SumQty = SumQty + oMat01.Columns("YQty").Cells(i + 1).Specific.VALUE
                        Next
                            
                        oForm01.Items("SumQty").Specific.VALUE = SumQty
                        
                    End If
Continue:
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_PP084_AddMatrixRow (pval.Row)
                    End If
                    Set RecordSet01 = Nothing
                ElseIf pval.ColUID = "PQty" Then
                    If oMat01.Columns("OrdGbn").Cells(pval.Row).Specific.VALUE = "104" Then '//멀티이면
                        If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0 Then
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                        ElseIf Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <> Val(oMat01.Columns("YQty").Cells(pval.Row).Specific.VALUE) Then
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                        Else
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Call oDS_PS_PP084L.setValue("U_YQty", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Call oDS_PS_PP084L.setValue("U_NQty", pval.Row - 1, 0)
                        End If
                    Else '//엔트베어링,휘팅,부품,기계,몰드
                        If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0 Then
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                        Else
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Call oDS_PS_PP084L.setValue("U_YQty", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            
                            
                            If oMat01.Columns("OrdGbn").Cells(pval.Row).Specific.VALUE = "101" Then
                                Weight = Val(MDC_PS_Common.GetValue("SELECT U_UnWeight  FROM [OITM] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1)) / 1000
                            Else
                                Weight = 0
                            End If
                            If Weight = 0 Then
                                Call oDS_PS_PP084L.setValue("U_PWeight", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                Call oDS_PS_PP084L.setValue("U_YWeight", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Else
                                Call oDS_PS_PP084L.setValue("U_PWeight", pval.Row - 1, Weight * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                Call oDS_PS_PP084L.setValue("U_YWeight", pval.Row - 1, Weight * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            End If
                            Call oDS_PS_PP084L.setValue("U_NQty", pval.Row - 1, 0)
                            Call oDS_PS_PP084L.setValue("U_NWeight", pval.Row - 1, 0)
                        End If
                    End If
                    
                        oMat01.LoadFromDataSource
                        
                        '//합격수량 sum
                        For i = 0 To oMat01.VisualRowCount - 1
                            SumQty = SumQty + oMat01.Columns("YQty").Cells(i + 1).Specific.VALUE
                        Next
                            
                        oForm01.Items("SumQty").Specific.VALUE = SumQty
                ElseIf pval.ColUID = "YQty" Then
                    If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) > Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) Then
                        Call MDC_Com.MDC_GF_Message("합격중량이 생산중량보다 클수 없습니다. 확인바랍니다.", "E")
                        Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                    Else
                        Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP084L.setValue("U_YWeight", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    End If
                    
                ElseIf pval.ColUID = "NQty" Then
                    If oMat01.Columns("OrdGbn").Cells(pval.Row).Specific.VALUE = "104" Then '//멀티이면
                        If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) > 0 Then
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                        ElseIf Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) > Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) Then
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                        Else
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Call oDS_PS_PP084L.setValue("U_YQty", pval.Row - 1, Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                        End If
                    Else '//엔트베어링,휘팅,부품,기계,몰드 //불량수량은 합격수량에 영향을 안준다.
                        If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0 Then
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                        ElseIf Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) > Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) Then
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP084L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                        Else
                            Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            'Call oDS_PS_PP084L.setValue("U_YQty", pval.Row - 1, Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            If oMat01.Columns("OrdGbn").Cells(pval.Row).Specific.VALUE = "101" Then '//휘팅이면
                                Weight = Val(MDC_PS_Common.GetValue("SELECT U_UnWeight  FROM [OITM] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE & "'", 0, 1)) / 1000
                            Else
                                Weight = 0
                            End If
                            If Weight = 0 Then
                                Call oDS_PS_PP084L.setValue("U_NWeight", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                'Call oDS_PS_PP084L.setValue("U_YWeight", pval.Row - 1, (Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)))
                            Else
                                Call oDS_PS_PP084L.setValue("U_NWeight", pval.Row - 1, Weight * Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                                'Call oDS_PS_PP084L.setValue("U_YWeight", pval.Row - 1, Weight * (Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)))
                            End If
                        End If
                    End If
                    
                    oMat01.LoadFromDataSource
                        
                    '//합격수량 sum
                    For i = 0 To oMat01.VisualRowCount - 1
                        SumQty = SumQty + oMat01.Columns("YQty").Cells(i + 1).Specific.VALUE
                    Next
                        
                    oForm01.Items("SumQty").Specific.VALUE = SumQty
                Else
                    Call oDS_PS_PP084L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_PP084H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "CardCode") Then
                    Call oDS_PS_PP084H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_PP084H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
                ElseIf (pval.ItemUID = "CntcCode") Then
                    Call oDS_PS_PP084H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_PP084H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                Else
                    Call oDS_PS_PP084H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oForm01.Update
            If pval.ItemUID = "Mat01" Then
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                oForm01.Items(pval.ItemUID).Click ct_Regular
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
    Dim i As Integer
    Dim SumQty As Currency
    
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        For i = 0 To oMat01.VisualRowCount - 1
                    If oMat01.Columns("PQty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat01.Columns("PQty").Cells(i + 1).Specific.VALUE
                    End If
         
        Next i
        oForm01.Items("SumQty").Specific.VALUE = SumQty
         
        Call PS_PP084_FormItemEnabled
        Call PS_PP084_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP084_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Dim oDataTable01 As SAPbouiCOM.DataTable
'        If (pval.ItemUID = "ItemCode") Then
'            Dim oDataTable01 As SAPbouiCOM.DataTable
'            Set oDataTable01 = pval.SelectedObjects
'            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CntcCode") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP080H", "U_CntcCode,U_CntcName")
'        End If
        If (pval.ItemUID = "Mat01") Then
            If (pval.ColUID = "WhsCode") Then
                If pval.SelectedObjects Is Nothing Then
                Else
                    Set oDataTable01 = pval.SelectedObjects
                    Call oDS_PS_PP084L.setValue("U_WhsCode", pval.Row - 1, oDataTable01.Columns("WhsCode").Cells(0).VALUE)
                    Call oDS_PS_PP084L.setValue("U_WhsName", pval.Row - 1, oDataTable01.Columns("WhsName").Cells(0).VALUE)
                    Set oDataTable01 = Nothing
                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
                    oMat01.LoadFromDataSource
                    oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                End If
            End If
        End If
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
    
    Dim SumQty As Currency
    
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            If (PS_PP084_Validate("행삭제") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_PP084L.RemoveRecord(oDS_PS_PP084L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_PP084_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_PP084L.GetValue("U_PP030No", oMat01.RowCount - 1)) <> "" Then
                    Call PS_PP084_AddMatrixRow(oMat01.RowCount)
                End If
                
                
                '//합격수량 sum
                For i = 0 To oMat01.VisualRowCount - 1
                    SumQty = SumQty + oMat01.Columns("YQty").Cells(i + 1).Specific.VALUE
                Next
                    
                oForm01.Items("SumQty").Specific.VALUE = SumQty
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP084_CreateItems() As Boolean
On Error GoTo PS_PP084_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP084H = oForm01.DataSources.DBDataSources("@PS_PP080H")
    Set oDS_PS_PP084L = oForm01.DataSources.DBDataSources("@PS_PP080L")
    Set oDS_PS_PP0841L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
    
    '//합계수량 sum 해서 보여줌 -선언
    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    oForm01.Items("SumQty").Specific.DataBind.SetBound True, "", "SumQty"
    
    oMat01.Columns("OrdSub1").Visible = False
    oMat01.Columns("OrdSub2").Visible = False
    oMat01.Columns("ORDRNo").Visible = False
    oMat01.Columns("RDR1No").Visible = False
    
'    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
'    Call oForm01.DataSources.UserDataSources.Add("WhsCode", dt_SHORT_TEXT, 100)
'    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
'    Call oForm01.Items("WhsCode").Specific.DataBind.SetBound(True, "", "WhsCode")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP084_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP084_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP084_ComboBox_Setting()
On Error GoTo PS_PP084_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP084", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP084", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP084", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP084", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP084", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PP084", "ItemCode")
    
    Call oForm01.Items("OrdGbn").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code In ('111','601') And U_PudYN = 'Y' order by Code", "", False, False)
    Call oForm01.Items("BPLId").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP084_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP084_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP084_CF_ChooseFromList()
On Error GoTo PS_PP084_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

'    Set oEdit = oForm01.Items("CntcCode").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_Employee
'    oCFLCreationParams.uniqueID = "CFLEMPLOYEE"
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
'    oEdit.ChooseFromListUID = "CFLEMPLOYEE"
'    oEdit.ChooseFromListAlias = "empID"

    Set oColumn = oMat01.Columns("WhsCode")
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = lf_Warehouses
    oCFLCreationParams.uniqueID = "CFLWAREHOUSES"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oColumn.ChooseFromListUID = "CFLWAREHOUSES"
    oColumn.ChooseFromListAlias = "WhsCode"
    Exit Sub
PS_PP084_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP084_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP084_FormItemEnabled()
On Error GoTo PS_PP084_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        Call PS_PP084_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        '//Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        Call oForm01.Items("OrdGbn").Specific.Select(0, psk_Index)
        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
        oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("OrdGbn").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
'        oForm01.Items("Mat01").Enabled = True
        oMat01.Columns("PP030No").Editable = True
        oMat01.Columns("PQty").Editable = True
        oMat01.Columns("NQty").Editable = True
        oMat01.Columns("WhsCode").Editable = True
        oMat01.Columns("Check").Editable = False
        
        '//수량 Sum
        oForm01.Items("SumQty").Specific.VALUE = 0
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("OrdGbn").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oMat01.Columns("PP030No").Editable = False
        oMat01.Columns("PQty").Editable = False
        oMat01.Columns("NQty").Editable = False
        oMat01.Columns("WhsCode").Editable = False
        oMat01.Columns("Check").Editable = False
        
        '//수량 Sum
        oForm01.Items("SumQty").Specific.VALUE = 0
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        '//각모드에따른 아이템설정
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("OrdGbn").Enabled = False
        oForm01.Items("CntcCode").Enabled = False
        oForm01.Items("DocDate").Enabled = False
'        oForm01.Items("Mat01").Enabled = False
        
        oMat01.Columns("PP030No").Editable = False
        oMat01.Columns("PQty").Editable = False
        oMat01.Columns("NQty").Editable = False
        oMat01.Columns("WhsCode").Editable = False
        If Trim(oDS_PS_PP084H.GetValue("CanCeled", 0)) = "Y" Then
            oMat01.Columns("Check").Editable = False
        Else
            oMat01.Columns("Check").Editable = True
        End If
        If Trim(oDS_PS_PP084H.GetValue("U_OrdGbn", 0)) = "104" Then '//멀티의경우
            oMat01.Columns("BWeight").Visible = False
            oMat01.Columns("PWeight").Visible = False
            oMat01.Columns("YWeight").Visible = False
            oMat01.Columns("NWeight").Visible = False
        Else '//그외의경우
            oMat01.Columns("BWeight").Visible = True
            oMat01.Columns("PWeight").Visible = True
            oMat01.Columns("YWeight").Visible = True
            oMat01.Columns("NWeight").Visible = True
        End If
    End If
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
'
    
    oMat01.AutoResizeColumns
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP084_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP084_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP084_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP084_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP084L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP084L.Offset = oRow
    oDS_PS_PP084L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP084_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP084_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP084_FormClear()
On Error GoTo PS_PP084_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP080'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP084_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP084_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP084_EnableMenus()
On Error GoTo PS_PP084_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP084_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP084_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP084_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP084_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP084_FormItemEnabled
        Call PS_PP084_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP084_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP084_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP084_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_PP084_DataValidCheck() As Boolean
On Error GoTo PS_PP084_DataValidCheck_Error
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim RDR1Qty         As Currency
    Dim PP080LQty       As Currency
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    PS_PP084_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP084_FormClear
    End If
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "작성일은 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_PP084_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("CntcCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "담당자는 필수입니다.", bmt_Short, True
        oForm01.Items("CntcCode").Click ct_Regular
        PS_PP084_DataValidCheck = False
        Exit Function
    End If
    If oMat01.VisualRowCount <= 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP084_DataValidCheck = False
        Exit Function
    End If
    
    '마감상태 체크_S(2017.11.23 송명규 추가)
    If MDC_PS_Common.Check_Finish_Status(Trim(oForm01.Items("BPLId").Specific.VALUE), oForm01.Items("DocDate").Specific.VALUE, oForm01.TypeEx) = False Then
        Call Sbo_Application.SetStatusBarMessage("마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 완료일자를 확인하고, 회계부서로 문의하세요.", bmt_Short, True)
        PS_PP084_DataValidCheck = False
        Exit Function
    End If
    '마감상태 체크_E(2017.11.23 송명규 추가)
    
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("PP030No").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "작업지시문서는 필수입니다.", bmt_Short, True
            oMat01.Columns("PP030No").Cells(i).Click ct_Regular
            PS_PP084_DataValidCheck = False
            Exit Function
        End If
        If (oMat01.Columns("WhsCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "창고코드는 필수입니다.", bmt_Short, True
            oMat01.Columns("WhsCode").Cells(i).Click ct_Regular
            PS_PP084_DataValidCheck = False
            Exit Function
        End If
        If (oMat01.Columns("ItemCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품목코드는 필수입니다.", bmt_Short, True
            oMat01.Columns("ItemCode").Cells(i).Click ct_Regular
            PS_PP084_DataValidCheck = False
            Exit Function
        End If
        If (Val(oMat01.Columns("PQty").Cells(i).Specific.VALUE) <= 0) Then
            Sbo_Application.SetStatusBarMessage "생산수량은 필수입니다.", bmt_Short, True
            oMat01.Columns("PQty").Cells(i).Click ct_Regular
            PS_PP084_DataValidCheck = False
            Exit Function
        End If
        '//기계몰드는 수주수량보다 생산수량이 많을 수 없다.
        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then
        
'            If MDC_PS_Common.User_MSTCOD = "12050802" Or MDC_PS_Common.User_MSTCOD = "1178504" Then '접속한 사용자가 노근용과장과 송명규대리일 때는 수주잔량과 생산완료량 비교 안함(2011.12.24)
'
'            Else
        
                If Val(oMat01.Columns("BQty").Cells(i).Specific.VALUE) < Val(oMat01.Columns("YQty").Cells(i).Specific.VALUE) Then
                    Sbo_Application.SetStatusBarMessage "수주잔량수량보다 생산수량이 많습니다. 확인바랍니다..", bmt_Short, True
                    oMat01.Columns("YQty").Cells(i).Click ct_Regular
                    PS_PP084_DataValidCheck = False
                    Exit Function
                End If
'            End If
            
            '기계 부품일 경우 수주량 완료량 비교해서 막기 - 류영조
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" And oForm01.Items("BPLId").Specific.Selected.VALUE = "2" Then
                sQry = "Select U_ItmMSort From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                
                If Trim(oRecordSet01.Fields(0).VALUE) = "10502" Or Trim(oRecordSet01.Fields(0).VALUE) = "10503" Or Trim(oRecordSet01.Fields(0).VALUE) = "10504" Then
                    sQry = "Select Quantity From RDR1 Where DocEntry = '" & Trim(oMat01.Columns("ORDRNo").Cells(i).Specific.VALUE) & "' And "
                    sQry = sQry & "LineNum = '" & Trim(oMat01.Columns("RDR1No").Cells(i).Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    RDR1Qty = oRecordSet01.Fields(0).VALUE
                    
                    sQry = "Select ISNULL(Sum(a.U_PQty),0) From [@PS_PP080L] a Inner Join [@PS_PP080H] b On a.DocEntry = b.DocEntry "
                    sQry = sQry & "Where a.U_ORDRNo = '" & Trim(oMat01.Columns("ORDRNo").Cells(i).Specific.VALUE) & "' And "
                    sQry = sQry & "a.U_RDR1No = '" & Trim(oMat01.Columns("RDR1No").Cells(i).Specific.VALUE) & "' And "
'                    sQry = sQry & "b.Canceled = 'N'"
                    sQry = sQry & "ISNULL(a.U_Check, 'N') = 'N'" '문서전체의 취소여부가 아닌 취소처리를 위한 체크박스의 값을 기준으로 조회(2011.12.28 송명규)
                    oRecordSet01.DoQuery sQry
                    PP080LQty = oRecordSet01.Fields(0).VALUE + oMat01.Columns("PQty").Cells(i).Specific.VALUE
                    
                    If RDR1Qty = PP080LQty Then '수주량보다 완료량이 적은 경우
                    
                        '검수입고(원재료품의, 외주제작품의, 가공비품의)가 등록 되지 않으면 생산완료 등록 불가(2012.01.12 송명규 수정)
                        sQry = "EXEC [PS_PP084_09] '" & Trim(oMat01.Columns("PP030HNo").Cells(i).Specific.VALUE) & "'"
                        oRecordSet01.DoQuery sQry
                        If oRecordSet01.Fields(0).VALUE <> 0 Then
                        
                            If oRecordSet01.Fields(1).VALUE = "10" Then
                            
                                Sbo_Application.SetStatusBarMessage "" & i & "번 라인 : 원재료품의가 모두 검수입고 되지 않았습니다. 확인해주세요.", bmt_Short, True
                                PS_PP084_DataValidCheck = False
                                Exit Function
                            
                            ElseIf oRecordSet01.Fields(1).VALUE = "30" Then
                            
                                Sbo_Application.SetStatusBarMessage "" & i & "번 라인 : 가공비품의가 모두 검수입고 되지 않았습니다. 확인해주세요.", bmt_Short, True
                                PS_PP084_DataValidCheck = False
                                Exit Function
                            
                            ElseIf oRecordSet01.Fields(1).VALUE = "40" Then
                            
                                Sbo_Application.SetStatusBarMessage "" & i & "번 라인 : 외주제작품의가 모두 검수입고 되지 않았습니다. 확인해주세요.", bmt_Short, True
                                PS_PP084_DataValidCheck = False
                                Exit Function
                            
                            End If
                            
                        End If
                        '검수입고(원재료품의, 외주제작품의, 가공비품의)가 등록 되지 않으면 생산완료 등록 불가(2012.01.12 송명규 수정)
                        
                    ElseIf RDR1Qty > PP080LQty Then '수주량보다 완료량이 적은 경우 무조건 완료를 잡을 수 있게 한다.
                    ElseIf RDR1Qty < PP080LQty Then '수주량보다 완료량이 많은 경우 무조건 완료를 잡을 수 없게 한다.
                    
'                        If MDC_PS_Common.User_MSTCOD = "12050802" Or MDC_PS_Common.User_MSTCOD = "1178504" Then '접속한 사용자가 노근용과장과 송명규대리일 때는 수주량과 생산완료량 비교 안함(2011.12.24)
'
'                        Else
                        
                            If Trim(oMat01.Columns("ORDRNo").Cells(i).Specific.VALUE) = 0 Or Trim(oMat01.Columns("ORDRNo").Cells(i).Specific.VALUE) = "" Then
                            Else
                                Sbo_Application.SetStatusBarMessage "수주수량보다 생산완료수량이 더 많습니다. 확인해주세요.", bmt_Short, True
                                PS_PP084_DataValidCheck = False
                                Exit Function
                            End If
                            
'                        End If
                        
                    End If
                End If
            End If
        End If
        
        '//부품,멀티인경우
        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "102" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then
            If oMat01.Columns("BatchNum").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "부품,멀티작업은 배치번호가 필수입니다.", bmt_Short, True
                PS_PP084_DataValidCheck = False
                Exit Function
            End If
        End If
    Next
    
    If (PS_PP084_Validate("검사01") = False) Then
        PS_PP084_DataValidCheck = False
        Exit Function
    End If
    
    Call oDS_PS_PP084L.RemoveRecord(oDS_PS_PP084L.Size - 1)
    Call oMat01.LoadFromDataSource
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP084_FormClear
    End If
    PS_PP084_DataValidCheck = True
    Set oRecordSet01 = Nothing
    Exit Function
PS_PP084_DataValidCheck_Error:
    Set oRecordSet01 = Nothing
    PS_PP084_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP084_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP084_MTX01()
On Error GoTo PS_PP084_MTX01_Error
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
        GoTo PS_PP084_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_PP084L.InsertRecord (i)
        End If
        oDS_PS_PP084L.Offset = i
        Call oDS_PS_PP084L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_PP084L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_PP084_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_PP084_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP084_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP084_DI_API01() As Boolean
On Error GoTo PS_PP084_DI_API01_Error
    PS_PP084_DI_API01 = True
    Dim i, j As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim ResultDocNum As Long
    Dim BaseLot As String
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).OrdGbn = oMat01.Columns("OrdGbn").Cells(i).Specific.Selected.VALUE
        ItemInformation(ItemInformationCount).PP030HNo = oMat01.Columns("PP030HNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).PP030MNo = oMat01.Columns("PP030MNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("YQty").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.VALUE
        ItemInformationCount = ItemInformationCount + 1
    Next
        
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenEntry)
    oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.Quantity = ItemInformation(i).Quantity
        If (ItemInformation(i).OrdGbn = "102" Or ItemInformation(i).OrdGbn = "104" Or ItemInformation(i).OrdGbn = "111") Then  '//부품,멀티인경우 배치를 선택
            If MDC_PS_Common.GetItem_ManBtchNum(ItemInformation(i).ItemCode) = "Y" Then '//배치사용품목이면
                oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(i).BatchNum
                oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(i).Quantity
                oDIObject.Lines.BatchNumbers.Add
            End If
        End If
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        oForm01.Items("OIGNNo").Specific.VALUE = ResultDocNum
        Call oDS_PS_PP084H.setValue("U_OIGNNo", 0, ResultDocNum)
        For i = 0 To ItemInformationCount - 1
            If (ItemInformation(i).OrdGbn = "111") Then '//분말일 경우 포장지시 Table에 생산입력 Sign Update
                MDC_PS_Common.DoQuery ("UPDATE [Z_PACKING_PD] SET PP080YN = 'Y' WHERE BatchNum = '" & ItemInformation(i).BatchNum & "'")
            End If
            
        Next
    Else
        GoTo PS_PP084_DI_API01_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_PP084_DI_API01_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_PP084_DI_API01 = False
    Set oDIObject = Nothing
    Exit Function
PS_PP084_DI_API01_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_PP084_DI_API01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP084_DI_API01 = False
End Function

Private Function PS_PP084_DI_API03() As Boolean
On Error GoTo PS_PP084_DI_API03_Error
    PS_PP084_DI_API03 = True
    
    Dim i, j, K As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim ResultDocNum As Long
    Dim BaseLot As String
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount - 1
        If Trim(oMat01.Columns("OIGENum").Cells(i).Specific.VALUE) = "" Then
            If oMat01.Columns("Check").Cells(i).Specific.Checked = True Then
                ReDim Preserve ItemInformation(ItemInformationCount)
                ItemInformation(ItemInformationCount).OrdGbn = oMat01.Columns("OrdGbn").Cells(i).Specific.Selected.VALUE
                ItemInformation(ItemInformationCount).PP030HNo = oMat01.Columns("PP030HNo").Cells(i).Specific.VALUE
                ItemInformation(ItemInformationCount).PP030MNo = oMat01.Columns("PP030MNo").Cells(i).Specific.VALUE
                ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
                ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("YQty").Cells(i).Specific.VALUE
                ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
                ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.VALUE
                ItemInformation(ItemInformationCount).LineNum = oMat01.Columns("LineNum").Cells(i).Specific.VALUE
                ItemInformationCount = ItemInformationCount + 1
            End If
'            k = k + 1
        End If
    Next i
    
'    If k = oMat01.VisualRowCount - 1 Then
'        oStatus = "C"
'        oCanceled = "Y"
'    Else
'        oStatus = "O"
'        oCanceled = "N"
'    End If
    
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenExit)
    oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    oDIObject.UserFields("U_CancDoc").VALUE = Trim(oForm01.Items("OIGNNo").Specific.VALUE)
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.Quantity = ItemInformation(i).Quantity
        If (ItemInformation(i).OrdGbn = "102" Or ItemInformation(i).OrdGbn = "104" Or ItemInformation(i).OrdGbn = "111") Then '//부품,멀티인경우 배치를 선택
            If MDC_PS_Common.GetItem_ManBtchNum(ItemInformation(i).ItemCode) = "Y" Then '//배치사용품목이면
                oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(i).BatchNum
                oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(i).Quantity
                oDIObject.Lines.BatchNumbers.Add
            End If
        End If
    Next

    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To ItemInformationCount - 1
            Call MDC_PS_Common.DoQuery("UPDATE [@PS_PP080L] SET U_OIGENum = '" & ResultDocNum & "', U_IGE1Num = '" & i & "', U_Check = 'Y' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "' And LineId = '" & ItemInformation(i).LineNum & "'")
            
            '//부품 실적추가분 취소처리 => 수량을 0으로 처리
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "111" Then
                Call MDC_PS_Common.DoQuery("UPDATE [@PS_PP040L] SET U_PQty = 0, U_PWeight = 0, U_YQty = 0, U_YWeight = 0 WHERE DocEntry = '" & oForm01.Items("PP040No").Specific.VALUE & "' And LineId = '" & ItemInformation(i).LineNum & "'")
                Call MDC_PS_Common.DoQuery("UPDATE [Z_PACKING_PD] SET PP080YN = 'N' WHERE BatchNum = '" & ItemInformation(i).BatchNum & "'")
            End If
            
            '//분말재공 실적추가분 취소처리 => 수량을 0으로 처리
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "601" Then
                Call MDC_PS_Common.DoQuery("UPDATE [@PS_PP040L] SET U_PQty = 0, U_PWeight = 0, U_YQty = 0, U_YWeight = 0 WHERE DocEntry = '" & oForm01.Items("PP040No").Specific.VALUE & "' And LineId = '" & ItemInformation(i).LineNum & "'")
            End If
            
'            oMat01.Columns("OIGENum").Cells(ItemInformation(i).LineNum).Specific.VALUE = ResultDocNum
'            oMat01.Columns("IGE1Num").Cells(ItemInformation(i).LineNum).Specific.VALUE = i
        Next i


    Else
        GoTo PS_PP084_DI_API03_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Set oDIObject = Nothing
    Exit Function
PS_PP084_DI_API03_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_PP084_DI_API03 = False
    Set oDIObject = Nothing
    Exit Function
PS_PP084_DI_API03_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_PP084_DI_API03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP084_DI_API03 = False
End Function

Private Sub PS_PP084_FormResize()
On Error GoTo PS_PP084_FormResize_Error
    
    Exit Sub
PS_PP084_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP084_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP084_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_PP084_Validate_Error
    PS_PP084_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP080H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
'        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
'        PS_PP084_Validate = False
'        GoTo PS_PP084_Validate_Exit
'    End If
    
    If ValidateType = "검사01" Then
        For i = 1 To oMat01.VisualRowCount - 1 '//입력된 행에 대해
            If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "102" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "107" Then
                If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" & oMat01.Columns("PP030No").Cells(i).Specific.VALUE & "'", 0, 1) <= 0 Then
                    Call MDC_Com.MDC_GF_Message("작업지시문서가 존재하지 않습니다.", "W")
                    PS_PP084_Validate = False
                    GoTo PS_PP084_Validate_Exit
                End If
            ElseIf oForm01.Items("OrdGbn").Specific.Selected.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "106" Then
                If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND PS_PP030H.DocEntry = '" & oMat01.Columns("PP030No").Cells(i).Specific.VALUE & "'", 0, 1) <= 0 Then
                    Call MDC_Com.MDC_GF_Message("작업지시문서가 존재하지 않습니다.", "W")
                    PS_PP084_Validate = False
                    GoTo PS_PP084_Validate_Exit
                End If
            End If
        Next
    ElseIf ValidateType = "수정" Then
'        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
'        Dim Exist As Boolean
'        Exist = False
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            Exist = False
'            For j = 1 To oMat01.RowCount - 1
'                '//라인번호가 같고, 품목코드가 같으면 존재하는행 , LineNum에 값이 존재하는지 확인필요(행삭제된행인경우 LineNum이 존재하지않음)
'                If Val(RecordSet01.Fields(1).Value) = Val(oMat01.Columns("U_LineNum").Cells(j).Specific.Value) And RecordSet01.Fields(2).Value = oMat01.Columns("1").Cells(j).Specific.Value And oMat01.Columns("U_LineNum").Cells(j).Specific.Value <> "" Then
'                    Exist = True
'                End If
'            Next
'            If (Exist = False) Then '//삭제된 행중
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "삭제된행이 다른사용자에 의해 출하,선출요청되었습니다. 적용할수 없습니다.", "W"
'                    PS_PP084_Validate = False
'                    GoTo PS_PP084_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        If oForm01.Mode <> fm_ADD_MODE Then
            MDC_Com.MDC_GF_Message "현재모드는 행삭제가 불가능합니다.", "W"
            PS_PP084_Validate = False
            GoTo PS_PP084_Validate_Exit
        End If
        '//행삭제전 행삭제가능여부검사
'        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
'            Else
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(oForm01.Items("8").Specific.Value) & "' AND U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할수 없습니다.", "W"
'                    PS_PP084_Validate = False
'                    GoTo PS_PP084_Validate_Exit
'                End If
'            End If
'        End If
    ElseIf ValidateType = "취소" Then
        '//취소할때 조건..
        '//재고가 모두 존재하는가..
'        Dim OIGNNo As String
'        Dim OrdGbn As String
'        OIGNNo = Trim(oForm01.Items("OIGNNo").Specific.VALUE)
'        OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.Selected.VALUE)
'        Query01 = "EXEC PS_PP084_04 '" & OrdGbn & "','" & OIGNNo & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            Call MDC_Com.MDC_GF_Message("품목코드: " & RecordSet01.Fields(0).VALUE & " 의 재고가 부족합니다.", "W")
'            PS_PP084_Validate = False
'            GoTo PS_PP084_Validate_Exit
'        Next
'''''''        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'''''''        RecordSet01.DoQuery Query01
'''''''        For i = 0 To RecordSet01.RecordCount - 1
'''''''            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'''''''                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
'''''''                PS_PP084_Validate = False
'''''''                GoTo PS_PP084_Validate_Exit
'''''''            End If
'''''''            RecordSet01.MoveNext
'''''''        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_PP084_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_PP084_Validate_Error:
    PS_PP084_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_PP084_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID$, OrdGbn$
        
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.VALUE)
    
    
    
    sQry = "EXEC [PS_PP084_04] '" & BPLID & "','" & OrdGbn & "'"
    oRecordSet01.DoQuery sQry
    
    oMat02.Clear
    oDS_PS_PP0841L.Clear
    
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP0841L.Size Then
            oDS_PS_PP0841L.InsertRecord (i)
        End If
        
        oMat02.AddRow
        oDS_PS_PP0841L.Offset = i
        oDS_PS_PP0841L.setValue "U_ColNum01", i, i + 1
        oDS_PS_PP0841L.setValue "U_ColDt01", i, Trim(oRecordSet01.Fields("DocDate").VALUE)
        oDS_PS_PP0841L.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("PP030No").VALUE)
        oDS_PS_PP0841L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)
        oDS_PS_PP0841L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("ItemName").VALUE)
        oDS_PS_PP0841L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("BatchNum").VALUE)
        oDS_PS_PP0841L.setValue "U_ColQty01", i, Trim(oRecordSet01.Fields("Quantity").VALUE)
        oDS_PS_PP0841L.setValue "U_ColQty02", i, Trim(oRecordSet01.Fields("BoxKg").VALUE)
        oDS_PS_PP0841L.setValue "U_ColNum02", i, Trim(oRecordSet01.Fields("BoxCnt").VALUE)
        oDS_PS_PP0841L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("InspNo").VALUE)

        oRecordSet01.MoveNext
        
    Next
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    
    oForm01.Freeze False
  
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
