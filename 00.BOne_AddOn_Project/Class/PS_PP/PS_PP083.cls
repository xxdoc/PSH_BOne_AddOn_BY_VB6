VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP083"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP083.cls
'//  Module         : 생산관리 > 생산완료등록
'//  Desc           : 생산완료등록
'//  FormType       : PS_PP083
'//  Create Date    : 2013.08.15  N.G.Y
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_PP083H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP083L As SAPbouiCOM.DBDataSource '등록라인

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
    Quantity As Long
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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP083.srf")
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
    
    oFormUniqueID01 = "PS_PP083_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP083_CreateItems
    Call PS_PP083_ComboBox_Setting
    Call PS_PP083_CF_ChooseFromList
    Call PS_PP083_EnableMenus
    Call PS_PP083_SetDocument(oFromDocEntry01)
    Call PS_PP083_FormResize
    
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
                    If (PS_PP083_Validate("취소") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If Sbo_Application.MessageBox("정말로 취소하시겠습니까?", "1", "예", "아니오") <> "1" Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    If PS_PP083_DI_API03 = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
'                    If PS_PP083_DI_API02 = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
                    oDocEntry = oForm01.Items("DocEntry").Specific.VALUE
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
                
                sQry = "Select Min(IsNULL(U_OIGENum, '')) From [@PS_PP083L] where DocEntry = '" & oDocEntry & "'"
                oRecordSet01.DoQuery sQry
                
                If Trim(oRecordSet01.Fields(0).VALUE) = "" Then
                    oStatus = "O"
                    oCanceled = "N"
                Else
                    oStatus = "C"
                    oCanceled = "Y"
                End If
                
                Call MDC_PS_Common.DoQuery("UPDATE [@PS_PP083H] SET Status = '" & oStatus & "', Canceled = '" & oCanceled & "' WHERE DocEntry = '" & oDocEntry & "'")
                
                oForm01.Mode = fm_FIND_MODE
                Call PS_PP083_FormItemEnabled
                oForm01.Items("DocEntry").Specific.VALUE = oDocEntry
                oForm01.Items("1").Click ct_Regular
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_PP083_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_PP083_FormItemEnabled '//UDO방식
                Call PS_PP083_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_PP083_FormItemEnabled
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
        Dim oTempClass As Variant
        
        If pval.ItemUID = "PS_PP083" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_PP083_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                If PS_PP083_DI_API01 = False Then
                    Call PS_PP083_AddMatrixRow(oMat01.VisualRowCount)
                    BubbleEvent = False
                    Exit Sub
                End If
                oDocEntry = oForm01.Items("DocEntry").Specific.VALUE
               
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_PP083_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//정상적으로 입고문서가 생성이 안되었을때 임시시작
                If oForm01.Items("OIGNNo").Specific.VALUE = "" Then
                    If PS_PP083_DI_API01 = False Then
                        Call PS_PP083_AddMatrixRow(oMat01.VisualRowCount)
                        BubbleEvent = False
                        Exit Sub
                    End If
                    oDocEntry = oForm01.Items("DocEntry").Specific.VALUE
                End If
                '//정상적으로 입고문서가 생성이 안되었을때 임시종료
                
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP083" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_PP083_FormItemEnabled
                    Call PS_PP083_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_PP083_FormItemEnabled
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
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "CntcCode" Then
            If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//사용자값활성
            End If
        ElseIf pval.ItemUID = "CardCode" Then
            If oForm01.Items("CardCode").Specific.VALUE = "" Then
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '//사용자값활성
            End If
        End If
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
        If pval.ColUID = "ItemCode" Then
            If oForm01.Items("BPLId").Specific.Selected.VALUE = "선택" Then
                Call MDC_Com.MDC_GF_Message("사업장은 필수입니다.", "W")
                BubbleEvent = False
                Exit Sub
            Else
                
            If oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE = "" Then
                Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                BubbleEvent = False
            End If
            '/////////////////////////////////////////////////////////
            'Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
            '///////////////////////////////////////////////////////
        
'                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "PP030No") '//사용자값활성
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
            Call oForm01.Freeze(True)
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "특정컬럼") Then
                    '//기타작업
                    Call oDS_PS_PP083L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP083L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_PP083_AddMatrixRow (pval.Row)
                    End If
                Else
                    Call oDS_PS_PP083L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                End If
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_PP083H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                ElseIf (pval.ItemUID = "BPLId") Then
                    Call oDS_PS_PP083H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    oMat01.Clear
                    oMat01.FlushToDataSource
                    oMat01.LoadFromDataSource
                    Call PS_PP083_AddMatrixRow(0, True)
                
                Else
                    Call oDS_PS_PP083H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
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
Dim li_row As Integer

    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)

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
    Dim i&
    
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If pval.ItemUID = "Mat01" And pval.Row = 0 And pval.ColUID = "Check" Then
            Dim Check$
            oForm01.Freeze True
            oMat01.FlushToDataSource
            If Trim(oDS_PS_PP083L.GetValue("U_Check", 0)) = "" Or Trim(oDS_PS_PP083L.GetValue("U_Check", 0)) = "N" Then
                Check = "Y"
            ElseIf Trim(oDS_PS_PP083L.GetValue("U_Check", 0)) = "Y" Then
                Check = "N"
            End If
            For i = 0 To oMat01.VisualRowCount - 2
                If Trim(oDS_PS_PP083L.GetValue("U_OIGENum", 0)) <> "" Then
                    oDS_PS_PP083L.setValue "U_Check", i, "Y"
                Else
                    oDS_PS_PP083L.setValue "U_Check", i, Check
                End If
            Next i
            oMat01.LoadFromDataSource
            oForm01.Freeze False
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
    
    Dim BPLID  As String
    
    Dim SumQty As Currency
    
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "ItemCode") Then
                    
                    BPLID = oForm01.Items("BPLId").Specific.VALUE
                    
                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    Query01 = "Select ItemName, U_ItmBsort, WhsCode = '10' + '" & BPLID & "', "
                    Query01 = Query01 + " WhsName = (Select WhsName From OWHS Where WhsCode = '10' + '" & BPLID & "')"
                    Query01 = Query01 + " From OITM Where ItemCode = '" & oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE & "'"
                    Call RecordSet01.DoQuery(Query01)
                    If RecordSet01.RecordCount = 0 Then
                        Call oDS_PS_PP083L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                    Else
                        Call oDS_PS_PP083L.setValue("U_ItemCode", pval.Row - 1, oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP083L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
                        Call oDS_PS_PP083L.setValue("U_OrdGbn", pval.Row - 1, RecordSet01.Fields("U_ItmBsort").VALUE)
                        
                        Call oDS_PS_PP083L.setValue("U_WhsCode", pval.Row - 1, RecordSet01.Fields("WhsCode").VALUE)
                        Call oDS_PS_PP083L.setValue("U_WhsName", pval.Row - 1, RecordSet01.Fields("WhsName").VALUE)
                        
                        oMat01.LoadFromDataSource
                        
''                        //합격수량 sum
                        For i = 0 To oMat01.VisualRowCount - 1
                            SumQty = SumQty + oMat01.Columns("YQty").Cells(i + 1).Specific.VALUE
                        Next

                        oForm01.Items("SumQty").Specific.VALUE = SumQty
                        
                    End If
Continue:
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP083L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_PP083_AddMatrixRow (pval.Row)
                    End If
                    Set RecordSet01 = Nothing
                ElseIf pval.ColUID = "YQty" Then
                    If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) < 0 Then
                        Call oDS_PS_PP083L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP083L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                    Else
                        Call oDS_PS_PP083L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                        Call oDS_PS_PP083L.setValue("U_YWeight", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    End If
                    
                    oMat01.LoadFromDataSource
                    '//합격수량 sum
                    For i = 0 To oMat01.VisualRowCount - 1
                        SumQty = SumQty + oMat01.Columns("YQty").Cells(i + 1).Specific.VALUE
                    Next
                        
                    oForm01.Items("SumQty").Specific.VALUE = SumQty
                
                Else
                    Call oDS_PS_PP083L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_PP083H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "CntcCode") Then
                    Call oDS_PS_PP083H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_PP083H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                ElseIf (pval.ItemUID = "CardCode") Then
                    Call oDS_PS_PP083H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_PP083H.setValue("U_CardName", 0, MDC_PS_Common.GetValue("SELECT CardName FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                Else
                    Call oDS_PS_PP083H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
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
                    If oMat01.Columns("YQty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat01.Columns("YQty").Cells(i + 1).Specific.VALUE
                    End If
         
        Next i
        oForm01.Items("SumQty").Specific.VALUE = SumQty
         
        Call PS_PP083_FormItemEnabled
        Call PS_PP083_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP083_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP083H", "U_CntcCode,U_CntcName")
'        End If
        If (pval.ItemUID = "Mat01") Then
            If (pval.ColUID = "WhsCode") Then
                If pval.SelectedObjects Is Nothing Then
                Else
                    Set oDataTable01 = pval.SelectedObjects
                    Call oDS_PS_PP083L.setValue("U_WhsCode", pval.Row - 1, oDataTable01.Columns("WhsCode").Cells(0).VALUE)
                    Call oDS_PS_PP083L.setValue("U_WhsName", pval.Row - 1, oDataTable01.Columns("WhsName").Cells(0).VALUE)
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
            If (PS_PP083_Validate("행삭제") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_PP083L.RemoveRecord(oDS_PS_PP083L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_PP083_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_PP083L.GetValue("U_ItemCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_PP083_AddMatrixRow(oMat01.RowCount)
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


Private Function PS_PP083_CreateItems() As Boolean
On Error GoTo PS_PP083_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP083H = oForm01.DataSources.DBDataSources("@PS_PP083H")
    Set oDS_PS_PP083L = oForm01.DataSources.DBDataSources("@PS_PP083L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '//합계수량 sum 해서 보여줌 -선언
    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    oForm01.Items("SumQty").Specific.DataBind.SetBound True, "", "SumQty"
    
    
'    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
'    Call oForm01.DataSources.UserDataSources.Add("WhsCode", dt_SHORT_TEXT, 100)
'    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
'    Call oForm01.Items("WhsCode").Specific.DataBind.SetBound(True, "", "WhsCode")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP083_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP083_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP083_ComboBox_Setting()
On Error GoTo PS_PP083_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP083", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP083", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP083", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP083", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP083", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PP083", "ItemCode")
    
    Call oForm01.Items("BPLId").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code")
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP083_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP083_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP083_CF_ChooseFromList()
On Error GoTo PS_PP083_CF_ChooseFromList_Error
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
PS_PP083_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP083_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP083_FormItemEnabled()
On Error GoTo PS_PP083_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        Call PS_PP083_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        '//Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
        oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
'        oForm01.Items("Mat01").Enabled = True
        oMat01.Columns("ItemCode").Editable = True
        oMat01.Columns("YQty").Editable = True
        oMat01.Columns("WhsCode").Editable = True
        oMat01.Columns("Check").Editable = False
        
        '//수량 Sum
        oForm01.Items("SumQty").Specific.VALUE = 0
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oMat01.Columns("ItemCode").Editable = False
        oMat01.Columns("YQty").Editable = False
        oMat01.Columns("Check").Editable = False
        
        '//수량 Sum
        oForm01.Items("SumQty").Specific.VALUE = 0
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        '//각모드에따른 아이템설정
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("CntcCode").Enabled = False
        oForm01.Items("DocDate").Enabled = False
'        oForm01.Items("Mat01").Enabled = False
        
        oMat01.Columns("ItemCode").Editable = False
        oMat01.Columns("YQty").Editable = False
        oMat01.Columns("WhsCode").Editable = False
        
        If Trim(oDS_PS_PP083H.GetValue("CanCeled", 0)) = "Y" Then
            oMat01.Columns("Check").Editable = False
        Else
            oMat01.Columns("Check").Editable = True
        End If
        
    End If
    
    
    oMat01.AutoResizeColumns
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP083_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP083_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP083_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP083_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP083L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP083L.Offset = oRow
    oDS_PS_PP083L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP083_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP083_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP083_FormClear()
On Error GoTo PS_PP083_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP083'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP083_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP083_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP083_EnableMenus()
On Error GoTo PS_PP083_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP083_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP083_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP083_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP083_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP083_FormItemEnabled
        Call PS_PP083_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_PP083_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP083_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP083_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_PP083_DataValidCheck() As Boolean
On Error GoTo PS_PP083_DataValidCheck_Error
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim RDR1Qty         As Currency
    Dim PP080LQty       As Currency
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    PS_PP083_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP083_FormClear
    End If
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "작성일은 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_PP083_DataValidCheck = False
        Exit Function
    End If
    If oForm01.Items("CntcCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "담당자는 필수입니다.", bmt_Short, True
        oForm01.Items("CntcCode").Click ct_Regular
        PS_PP083_DataValidCheck = False
        Exit Function
    End If
    If oMat01.VisualRowCount <= 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP083_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        
        If (oMat01.Columns("WhsCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "창고코드는 필수입니다.", bmt_Short, True
            oMat01.Columns("WhsCode").Cells(i).Click ct_Regular
            PS_PP083_DataValidCheck = False
            Exit Function
        End If
        If (oMat01.Columns("ItemCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품목코드는 필수입니다.", bmt_Short, True
            oMat01.Columns("ItemCode").Cells(i).Click ct_Regular
            PS_PP083_DataValidCheck = False
            Exit Function
        End If
        If (Val(oMat01.Columns("YQty").Cells(i).Specific.VALUE) <= 0) Then
            Sbo_Application.SetStatusBarMessage "생산수량은 필수입니다.", bmt_Short, True
            oMat01.Columns("YQty").Cells(i).Click ct_Regular
            PS_PP083_DataValidCheck = False
            Exit Function
        End If
        
    Next
    
    If (PS_PP083_Validate("검사01") = False) Then
        PS_PP083_DataValidCheck = False
        Exit Function
    End If
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call oDS_PS_PP083L.RemoveRecord(oDS_PS_PP083L.Size - 1)
    End If
    
    If oForm01.Items("OIGNNo").Specific.VALUE = "" Then
        If (oForm01.Mode = fm_UPDATE_MODE) Then
            Call oDS_PS_PP083L.RemoveRecord(oDS_PS_PP083L.Size - 1)
        End If
    End If
    
    Call oMat01.LoadFromDataSource
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP083_FormClear
    End If
    PS_PP083_DataValidCheck = True
    Set oRecordSet01 = Nothing
    Exit Function
PS_PP083_DataValidCheck_Error:
    Set oRecordSet01 = Nothing
    PS_PP083_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP083_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP083_MTX01()
On Error GoTo PS_PP083_MTX01_Error
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
        GoTo PS_PP083_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_PP083L.InsertRecord (i)
        End If
        oDS_PS_PP083L.Offset = i
        Call oDS_PS_PP083L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_PP083L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_PP083_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_PP083_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP083_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP083_DI_API01() As Boolean
On Error GoTo PS_PP083_DI_API01_Error
    PS_PP083_DI_API01 = True
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
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("YQty").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
        ItemInformationCount = ItemInformationCount + 1
    Next
        
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenEntry)
    oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    oDIObject.UserFields("U_CardCode").VALUE = oForm01.Items("CardCode").Specific.VALUE
    oDIObject.UserFields("U_CardName").VALUE = oForm01.Items("CardName").Specific.VALUE
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.Quantity = ItemInformation(i).Quantity
        
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        oForm01.Items("OIGNNo").Specific.VALUE = ResultDocNum
        Call oDS_PS_PP083H.setValue("U_OIGNNo", 0, ResultDocNum)
        
    Else
        GoTo PS_PP083_DI_API01_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_PP083_DI_API01_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_PP083_DI_API01 = False
    Set oDIObject = Nothing
    Exit Function
PS_PP083_DI_API01_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_PP083_DI_API01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP083_DI_API01 = False
End Function


Private Function PS_PP083_DI_API02() As Boolean
On Error GoTo PS_PP083_DI_API02_Error
    PS_PP083_DI_API02 = True
    Dim i, j As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim ResultDocNum As Long
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Query01 = "SELECT IGN1.ItemCode,IGN1.Quantity,IGN1.WhsCode FROM [OIGN] OIGN LEFT JOIN [IGN1] IGN1 ON OIGN.DocEntry = IGN1.DocEntry WHERE OIGN.DocEntry = '" & oForm01.Items("OIGNNo").Specific.VALUE & "'"
    
    RecordSet01.DoQuery Query01
    
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenExit)
    oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = RecordSet01.Fields("ItemCode").VALUE
        oDIObject.Lines.WarehouseCode = RecordSet01.Fields("WhsCode").VALUE
        oDIObject.Lines.Quantity = RecordSet01.Fields("Quantity").VALUE
        oDIObject.UserFields("U_CancDoc").VALUE = oForm01.Items("OIGNNo").Specific.VALUE '//입고취소 문서번호
        
        
        RecordSet01.MoveNext
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        Call MDC_PS_Common.DoQuery("UPDATE [@PS_PP083H] SET U_OIGENo = '" & ResultDocNum & "' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'")
        RecordSet01.MoveFirst
    Else
        GoTo PS_PP083_DI_API02_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set RecordSet01 = Nothing
    Set oDIObject = Nothing
    Exit Function
PS_PP083_DI_API02_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_PP083_DI_API02 = False
    Set oDIObject = Nothing
    Set RecordSet01 = Nothing
    Exit Function
PS_PP083_DI_API02_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_PP083_DI_API02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP083_DI_API02 = False
    Set RecordSet01 = Nothing
End Function

Private Function PS_PP083_DI_API03() As Boolean
On Error GoTo PS_PP083_DI_API03_Error
    PS_PP083_DI_API03 = True
    
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
                ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
                ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("YQty").Cells(i).Specific.VALUE
                ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
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
    oDIObject.UserFields("U_CancDoc").VALUE = oForm01.Items("OIGNNo").Specific.VALUE
    oDIObject.UserFields("U_CardCode").VALUE = oForm01.Items("CardCode").Specific.VALUE
    oDIObject.UserFields("U_CardName").VALUE = oForm01.Items("CardName").Specific.VALUE
    
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.Quantity = ItemInformation(i).Quantity
        
    Next

    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To ItemInformationCount - 1
            Call MDC_PS_Common.DoQuery("UPDATE [@PS_PP083L] SET U_OIGENum = '" & ResultDocNum & "', U_IGE1Num = '" & i & "', U_Check = 'Y' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "' And LineId = '" & ItemInformation(i).LineNum & "'")
        
        Next i


        '//이부분 코팅해야 함~
'        If (oForm01.Items("OrdGbn").Specific.Selected.VALUE = "101" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "102") Then '//휘팅, 부품 실적추가분 취소처리
'            Call MDC_PS_Common.DoQuery("UPDATE [@PS_PP040H] SET Canceled = 'Y' WHERE DocEntry = '" & oForm01.Items("PP040No").Specific.VALUE & "'")
'        End If
    Else
        GoTo PS_PP083_DI_API03_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Set oDIObject = Nothing
    Exit Function
PS_PP083_DI_API03_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_PP083_DI_API03 = False
    Set oDIObject = Nothing
    Exit Function
PS_PP083_DI_API03_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_PP083_DI_API03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP083_DI_API03 = False
End Function

Private Sub PS_PP083_FormResize()
On Error GoTo PS_PP083_FormResize_Error
    
    Exit Sub
PS_PP083_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP083_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP083_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_PP083_Validate_Error
    PS_PP083_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP083H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
'        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
'        PS_PP083_Validate = False
'        GoTo PS_PP083_Validate_Exit
'    End If
    
    If ValidateType = "검사01" Then
        For i = 1 To oMat01.VisualRowCount - 1 '//입력된 행에 대해
            If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM OITM WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & "'", 0, 1) <= 0 Then
                Call MDC_Com.MDC_GF_Message("품목코드가 존재하지 않습니다.", "W")
                PS_PP083_Validate = False
                GoTo PS_PP083_Validate_Exit
            End If
        Next
    ElseIf ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then
        If oForm01.Mode <> fm_ADD_MODE Then
            MDC_Com.MDC_GF_Message "현재모드는 행삭제가 불가능합니다.", "W"
            PS_PP083_Validate = False
            GoTo PS_PP083_Validate_Exit
        End If

    ElseIf ValidateType = "취소" Then
 
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_PP083_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_PP083_Validate_Error:
    PS_PP083_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_PP083_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



