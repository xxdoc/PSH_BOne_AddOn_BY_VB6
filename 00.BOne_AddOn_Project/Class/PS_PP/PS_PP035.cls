VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP035"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP035.cls
'//  Module         : 생산관리
'//  Desc           : 작업지시목록리스트(전체작지조회)
'//  FormType       : PS_PP035
'//  Create Date    : 2010.10.08
'//  Creator        : Dong sub Kim
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
'Public oMat02 As SAPbouiCOM.Matrix
''Private oDS_PS_PP035H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP035L As SAPbouiCOM.DBDataSource '등록라인

'Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String  '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long    '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMat01Row01 As Long
Private oText(4)    As String
'Private oMat02Row02 As Long

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    LotNo As String
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
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP035.srf")
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
    
    oFormUniqueID01 = "PS_PP035_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP035_CreateItems
    Call PS_PP035_ComboBox_Setting
    Call PS_PP035_Initial_Setting
    Call PS_PP035_CF_ChooseFromList
    Call PS_PP035_EnableMenus
    Call PS_PP035_SetDocument(oFromDocEntry01)
    Call PS_PP035_FormResize
    
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
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_PP035_AddMatrixRow(oMat01.VisualRowCount)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                '//Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_PP035_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_PP035_FormItemEnabled '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_PP035L.RemoveRecord(oDS_PS_PP035L.Size - 1)
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
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_PP035_MTX01
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Btn02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_PP035_SetBaseForm '//부모폼에입력
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "Btn03" Then
            Call Chk_Data1
        End If
        
        If pval.ItemUID = "Btn04" Then
        
            Call Chk_Data
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_PP035_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP035" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP035_FormItemEnabled
'                    Call PS_PP035_FormClear '//UDO방식일때
'                    Call PS_PP035_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_PP035_FormItemEnabled
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "CntcCode") '//사용자값활성
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "Mat01" Then
                If pval.ColUID = "ItemCode" Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
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
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "Mat01" And pval.Row = "0" And pval.ColUID = "CHK" Then
            Dim Chk$
            Dim i As Long
            oForm01.Freeze True
            oMat01.FlushToDataSource
            If Trim(oDS_PS_PP035L.GetValue("U_ColReg01", 0)) = "" Or Trim(oDS_PS_PP035L.GetValue("U_ColReg01", 0)) = "N" Then
                Chk = "Y"
            ElseIf Trim(oDS_PS_PP035L.GetValue("U_ColReg01", 0)) = "Y" Then
                Chk = "N"
            End If
            For i = 0 To oMat01.VisualRowCount - 1
                oDS_PS_PP035L.setValue "U_ColReg01", i, Chk
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
Dim TempForm01 As Variant
    If pval.BeforeAction = True Then
        If oMat01.Columns("DocEntry").Cells(pval.Row).Specific.String <> "" Then
            Set TempForm01 = New PS_PP030
            Call TempForm01.LoadForm(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.String)
            BubbleEvent = False
        Else
            
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
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "DocEntry") Then
                    '//기타작업
'                    Call oDS_PS_PP035L.setValue("U_" & "ColReg02", pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP035L.GetValue("U_" & "ColReg02", pval.Row - 1)) <> "" Then
'                        PS_PP035_AddMatrixRow (pval.Row)
'                    End If
                ElseIf (pval.ColUID = "CntcCode") Then
'                    Query01 = "SELECT U_FULLNAME FROM [OHEM] WHERE U_MSTCOD = '" & Trim(oMat01.Columns("CntcCode").Cells(pval.Row).Specific.VALUE) & "'"
                    Query01 = "SELECT LastName, FirstName FROM [OHEM] WHERE EmpID = '" & Trim(oMat01.Columns("CntcCode").Cells(pval.Row).Specific.VALUE) & "'"
                    RecordSet01.DoQuery (Query01)
                    oMat01.Columns("CntcName").Cells(pval.Row).Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE) & Trim(RecordSet01.Fields(1).VALUE)
'                    Call oDS_PS_PP035L.setValue("U_ColReg14", pval.Row - 1, Trim(RecordSet01.Fields(0).Value) & Trim(RecordSet01.Fields(1).Value))
'                    Call oDS_PS_PP035L.setValue("U_ColReg16", pval.Row - 1, Trim(oMat01.Columns("CntcCode").Cells(pval.Row).Specific.Value))
'                    Call oDS_PS_PP035L.setValue("U_ColReg01", pval.Row - 1, Trim(oMat01.Columns("CHK").Cells(pval.Row).Specific.Checked))
'                    Call oDS_PS_PP035L.setValue("U_ColReg07", pval.Row - 1, Trim(oMat01.Columns("Canceled").Cells(pval.Row).Specific.Value))
'                    Call oDS_PS_PP035L.setValue("U_ColQty02", pval.Row - 1, Trim(oMat01.Columns("SelWt").Cells(pval.Row).Specific.Value))
'                    Call oDS_PS_PP035L.setValue("U_ColDt01", pval.Row - 1, Trim(oMat01.Columns("DocDate").Cells(pval.Row).Specific.Value))
'                    Call oDS_PS_PP035L.setValue("U_ColDt02", pval.Row - 1, Trim(oMat01.Columns("DueDate").Cells(pval.Row).Specific.Value))
                ElseIf (pval.ColUID = "ItemCode") Then
                    Query01 = "Select ItemName From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE) & "'"
                    RecordSet01.DoQuery Query01
                    oMat01.Columns("ItemName").Cells(pval.Row).Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                Else
'                    Call oDS_PS_PP035L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
                End If
            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
            oForm01.Update
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
        Call PS_PP035_FormItemEnabled
        '//Call PS_PP035_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
'        oForm01.Items("Mat01").Top = 50
'        oForm01.Items("Mat01").Height = (oForm01.Height / 2) - 50
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
        If (pval.ItemUID = "ItemCode") Then
            Set oDataTable01 = pval.SelectedObjects
            If oDataTable01 Is Nothing Then
            Else
                oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns("ItemCode").Cells(0).VALUE
                oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns("ItemName").Cells(0).VALUE
            End If
        End If
        oForm01.Update
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
    '        For i = 1 To oMat01.VisualRowCount
    '            oMat01.Columns("COL01").Cells(i).Specific.Value = i
    '        Next i
    '        oMat01.FlushToDataSource
    '        Call oDS_PS_PP035L.RemoveRecord(oDS_PS_PP035L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_PP035_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_PP035L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_PP035_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP035_CreateItems() As Boolean
On Error GoTo PS_PP035_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'   Set oDS_PS_PP035H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_PP035L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oMat01 = oForm01.Items("Mat01").Specific
'   Set oMat02 = oForm01.Items("Mat02").Specific
    oMat01.SelectionMode = ms_NotSupported
'   oMat02.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
'   oMat02.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 20)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    
    Call oForm01.DataSources.UserDataSources.Add("Canceled", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Canceled").Specific.DataBind.SetBound(True, "", "Canceled")
    
    Call oForm01.DataSources.UserDataSources.Add("OrdNum", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdNum").Specific.DataBind.SetBound(True, "", "OrdNum")
    
    Call oForm01.DataSources.UserDataSources.Add("OrdNum1", dt_SHORT_TEXT, 20)
    Call oForm01.Items("OrdNum1").Specific.DataBind.SetBound(True, "", "OrdNum1")
    
    Call oForm01.DataSources.UserDataSources.Add("OrdGbn", dt_SHORT_TEXT, 10)
    Call oForm01.Items("OrdGbn").Specific.DataBind.SetBound(True, "", "OrdGbn")
    
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    
    Call oForm01.DataSources.UserDataSources.Add("WorkDtFr", dt_DATE, 8)
    Call oForm01.Items("WorkDtFr").Specific.DataBind.SetBound(True, "", "WorkDtFr")
    
    Call oForm01.DataSources.UserDataSources.Add("WorkDtTo", dt_DATE, 8)
    Call oForm01.Items("WorkDtTo").Specific.DataBind.SetBound(True, "", "WorkDtTo")
    
    Call oForm01.DataSources.UserDataSources.Add("FrgnName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("FrgnName").Specific.DataBind.SetBound(True, "", "FrgnName")
    
    Call oForm01.DataSources.UserDataSources.Add("Size", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Size").Specific.DataBind.SetBound(True, "", "Size")
    
    '생산완료여부
    Call oForm01.DataSources.UserDataSources.Add("PrdYN", dt_SHORT_TEXT, 20)
    Call oForm01.Items("PrdYN").Specific.DataBind.SetBound(True, "", "PrdYN")
    
    '작업상태변경_체크박스(2012.07.02 송명규 추가)
    Call oForm01.DataSources.UserDataSources.Add("ChkWCon", dt_SHORT_TEXT, 1)
    Call oForm01.Items("ChkWCon").Specific.DataBind.SetBound(True, "", "ChkWCon")
    
    oForm01.Items("Mat01").Enabled = True
'    oForm01.Items("Mat02").Enabled = False
        
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP035_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP035_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP035_ComboBox_Setting()
On Error GoTo PS_PP035_ComboBox_Setting_Error

    Dim i As Long
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    ' 사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)

    ' 작지상태
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP035", "Canceled", "", "N", "계획")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP035", "Canceled", "", "Y", "취소")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Canceled").Specific, "PS_PP035", "Canceled")
    Call oForm01.Items("Canceled").Specific.Select("계획", psk_ByValue)
    
    ' 작지구분
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y'", "101", False, False)
    
    If MDC_PS_Common.User_BPLId() = "1" Then
        Call oForm01.Items("OrdGbn").Specific.Select("104", psk_ByDescription)
    Else
        Call oForm01.Items("OrdGbn").Specific.Select("105", psk_ByDescription)
    End If
    '//Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP035", "oMat01", "OrdGbn", "", "")
    
    ' 작지상태(matrix)
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP035", "oMat01", "Canceled", "N", "계획")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP035", "oMat01", "Canceled", "Y", "취소")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Canceled"), "PS_PP035", "oMat01", "Canceled")
'    Call oMat01.Columns("OrdStat").Cells(oMat01.RowCount).Specific.Select("취소", psk_ByValue)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] order by Code", "", False, True)
'    oForm01.Items("ItemCode").Click ct_Regular
    

    '작업상태(2012.06.30 송명규 추가)
    Dim sQry As String
    
    sQry = "            SELECT      T1.U_Minor,"
    sQry = sQry & "                 T1.U_CdName"
    sQry = sQry & "  FROM       [@PS_SY001H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_SY001L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      T1.Code = 'S003'"

    Call oMat01.Columns("WorkCon").ValidValues.Add("", "")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("WorkCon"), sQry)
    '작업상태(2012.06.30 송명규 추가)
    
    
    Set oCombo = oForm01.Items("TradeType").Specific
    oCombo.ValidValues.Add "1", "전체"
    oCombo.ValidValues.Add "2", "일반"
    oCombo.ValidValues.Add "3", "선생산"
    oCombo.Select "0", psk_Index
    
    '생산완료 구분 추가(2011.11.29 송명규)
    Set oCombo = oForm01.Items("PrdYN").Specific
    oCombo.ValidValues.Add "%", "전체"
    oCombo.ValidValues.Add "Y", "생산완료"
    oCombo.ValidValues.Add "N", "생산미완료"
    oCombo.Select "0", psk_Index
    
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP035_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP035_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP035_Initial_Setting()
On Error GoTo PS_PP035_Initial_Setting_Error
 ' 사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 ' 인수자
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD()
    Exit Sub
PS_PP035_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP035_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP035_CF_ChooseFromList()
On Error GoTo PS_PP035_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

    Set oEdit = oForm01.Items("ItemCode").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = "4"
    oCFLCreationParams.uniqueID = "CFLITEMCD"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oEdit.ChooseFromListUID = "CFLITEMCD"
    oEdit.ChooseFromListAlias = "ItemCode"
    
    Exit Sub
PS_PP035_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP035_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP035_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP035_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP035_FormItemEnabled
        Call PS_PP035_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_PP035_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP035_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP035_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP035_FormItemEnabled()
On Error GoTo PS_PP035_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_PP035_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP035_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP035_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP035_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP035_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP035L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP035L.Offset = oRow
    oDS_PS_PP035L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP035_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP035_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP035_FormClear()
On Error GoTo PS_PP035_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP035'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP035_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP035_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP035_EnableMenus()
On Error GoTo PS_PP035_EnableMenus_Error
    '//메뉴활성화
    Call oForm01.EnableMenu("1281", False)
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP035_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP035_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP035_FormResize()
On Error GoTo PS_PP035_FormResize_Error
    
    Exit Sub
PS_PP035_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP035_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP035_DataValidCheck() As Boolean
On Error GoTo PS_PP035_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_PP035_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_PP035_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_PP035_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_PP035L.RemoveRecord(oDS_PP035L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_PP035_FormClear
    Exit Function
PS_PP035_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP035_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP035_MTX01()
On Error GoTo PS_PP035_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim Param08 As String
    Dim Param09 As String
    Dim Param10 As String
    Dim Param11 As String
    Dim Param12 As String
    Dim Param13 As String
    Dim Param14 As String
    
    Dim Temp01 As String
    Dim Temp02 As String
    Dim Temp03 As String
    
    Param01 = Trim(oForm01.Items("BPLId").Specific.VALUE)
    Param02 = Trim(oForm01.Items("Canceled").Specific.VALUE)
    Param03 = Trim(oForm01.Items("OrdNum").Specific.VALUE)
    Param04 = Trim(oForm01.Items("OrdGbn").Specific.VALUE)
    Param05 = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    Param06 = Trim(oForm01.Items("WorkDtFr").Specific.VALUE)
    Param07 = Trim(oForm01.Items("WorkDtTo").Specific.VALUE)
    Param08 = Trim(oForm01.Items("FrgnName").Specific.VALUE)
    Param09 = Trim(oForm01.Items("Size").Specific.VALUE)
    Param10 = Trim(oForm01.Items("CardCode").Specific.VALUE)
    Param11 = Trim(oForm01.Items("Mark").Specific.VALUE)
    Param12 = Trim(oForm01.Items("OrdNum1").Specific.VALUE)
    Param13 = Trim(oForm01.Items("TradeType").Specific.VALUE)
    Param14 = Trim(oForm01.Items("PrdYN").Specific.VALUE)
    
    If Param06 = "" And Param07 = "" Then
        Param06 = "19000101"
        Param07 = "99991231"
    ElseIf Param06 = "" And Param07 <> "" Then
        Param06 = ""
    ElseIf Param06 <> "" And Param07 = "" Then
        Param07 = Param06
    ElseIf Param06 <> "" And Param07 <> "" Then
    End If
    
    If Param03 = "" And Param12 = "" Then
        Param03 = "0"
        Param12 = "ZZZZZZZZZZZZZZZZZZZ"
    ElseIf Param03 = "" Then
        Param03 = Param12
    ElseIf Param12 = "" Then
        Param12 = Param03
    End If
    
    
    
'    Param07 = Format(Param07, "0000-00-00")
'    Param08 = Format(Param08, "0000-00-00")
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP035_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "', '" & Param08 & "', '" & Param09 & "', '" & Param10 & "', '" & Param11 & "', '" & Param12 & "', '" & Param13 & "','" & Param14 & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        oForm01.Items("Mat01").Enabled = False
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
        GoTo PS_PP035_MTX01_Exit
    Else
'        Call MDC_Com.MDC_GF_Message("조회를 성공하였습니다.", "S")
        oForm01.Items("Mat01").Enabled = True
    End If
    
    
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_PP035L.InsertRecord (i)
        End If
        oDS_PS_PP035L.Offset = i
        Call oDS_PS_PP035L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP035L.setValue("U_ColReg01", i, False)
        Call oDS_PS_PP035L.setValue("U_ColReg02", i, RecordSet01.Fields("DocEntry").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg03", i, RecordSet01.Fields("OrdGbn").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg04", i, RecordSet01.Fields("OrdNum").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg05", i, RecordSet01.Fields("OrdSub1").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg06", i, RecordSet01.Fields("OrdSub2").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg07", i, RecordSet01.Fields("Canceled").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg08", i, RecordSet01.Fields("CardCode").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg09", i, RecordSet01.Fields("CardName").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg10", i, RecordSet01.Fields("ItemCode").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg11", i, RecordSet01.Fields("ItemName").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg12", i, RecordSet01.Fields("Quality").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColNum01", i, RecordSet01.Fields("Unweight").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColNum02", i, RecordSet01.Fields("ReqWt").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColQty01", i, RecordSet01.Fields("SelWt").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg16", i, RecordSet01.Fields("CntcCode").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg14", i, RecordSet01.Fields("CntcName").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg17", i, RecordSet01.Fields("SubName").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg18", i, RecordSet01.Fields("JISNO").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg19", i, RecordSet01.Fields("BatchNum").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg25", i, RecordSet01.Fields("OutSize").VALUE)
        Call oDS_PS_PP035L.setValue("U_ColReg20", i, RecordSet01.Fields("WorkCon").VALUE) '작업상태(2012.06.30 송명규 추가)
        
        If RecordSet01.Fields("DocDate").VALUE = "1900-01-01" Then
            Temp01 = ""
        Else
            Temp01 = RecordSet01.Fields("DocDate").VALUE
        End If
        
        Temp01 = Format(Temp01, "YYYYMMDD")
        Call oDS_PS_PP035L.setValue("U_ColDt01", i, Temp01)
        
        If RecordSet01.Fields("DueDate").VALUE = "1900-01-01" Then
            Temp02 = ""
        Else
            Temp02 = RecordSet01.Fields("DueDate").VALUE
        End If
        
        Temp02 = Format(Temp02, "YYYYMMDD")
        Call oDS_PS_PP035L.setValue("U_ColDt02", i, Temp02)
        
        If RecordSet01.Fields("WalDoc").VALUE = "1900-01-01" Then
            Temp03 = ""
        Else
            Temp03 = RecordSet01.Fields("WalDoc").VALUE
        End If
        
        Temp03 = Format(Temp03, "YYYYMMDD")
        Call oDS_PS_PP035L.setValue("U_ColDt03", i, Temp03)
        
        
'        If RecordSet01.Fields("DocDate").VALUE = "1900-01-01" Then
'            Temp01 = ""
'        Else
'            Temp01 = RecordSet01.Fields("DocDate").VALUE
'        End If
'        Temp01 = Format(Temp01, "YYYYMMDD")
'        Call oDS_PS_PP035L.setValue("U_ColDt02", i, Temp01)
        
        Call oDS_PS_PP035L.setValue("U_ColReg15", i, RecordSet01.Fields("CpName").VALUE)
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    
    If Param04 = "104" Then
        oMat01.Columns("ItemCode").Editable = True
    Else
        oMat01.Columns("ItemCode").Editable = False
    End If
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    Call MDC_Com.MDC_GF_Message("조회를 성공하였습니다.", "S")
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP035_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP035_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP035_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP035_SetBaseForm()
On Error GoTo PS_PP035_SetBaseForm_Error

    Dim i As Long
    Dim ItemCode01 As String

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As Double
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim ItemCode As String
    Dim ItemName As String
    Dim WorkCon As String
    
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    For i = 1 To oMat01.RowCount '//품목선택품목
    
        '재고가 존재하면 닫기(종료) 불가 기능 추가(2012.01.11 송명규 추가)
        Dim QueryString As String
'            If Trim(oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE) = "104" Then '//멀티
'                QueryString = " Select Sum(a.Quantity * (Case When a.Direction = '0' Then 1 Else -1 End)) As Quantity"
'                QueryString = QueryString & "  from IBT1 a Inner Join OITM b On a.ItemCode = b.ItemCode And b.U_ItmBsort = '104' "
'                QueryString = QueryString & "  Where a.BaseType In ('59', '60')"
'                QueryString = QueryString & "    and a.BatchNum = '" & Trim(oMat01.Columns("OrdNum").Cells(i).Specific.VALUE) & "'"
'            Else
'
'                QueryString = "                     SELECT      SUM(A.InQty) - SUM(A.OutQty) AS [StockQty]"
'                QueryString = QueryString & "  FROM       OINM AS A"
'                QueryString = QueryString & "                 INNER JOIN"
'                QueryString = QueryString & "                 OITM As B"
'                QueryString = QueryString & "                     ON A.ItemCode = B.ItemCode"
'                QueryString = QueryString & "  WHERE      B.U_ItmBsort IN ('105','106')"
'                QueryString = QueryString & "                 AND A.ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i).Specific.VALUE) & "'"
'                QueryString = QueryString & "  GROUP BY  A.ItemCode"
'            End If
'
'            If IIf((MDC_PS_Common.GetValue(QueryString, 0, 1)) = "", 0, (MDC_PS_Common.GetValue(QueryString, 0, 1))) > 0 Then
'
'                Call MDC_Com.MDC_GF_Message(i & "행의 작업지시는 현재 재고가 존재합니다. 처리할 수 없습니다.", "W")
''                GoTo PS_PP035_SetBaseForm_Error

        If oMat01.Columns("CHK").Cells(i).Specific.Checked = True Then
        
            '작업상태만 변경 시 재고존재 체크 안함
            If oForm01.Items("ChkWCon").Specific.Checked = True Then
            
                Param01 = oMat01.Columns("DocEntry").Cells(i).Specific.VALUE
                WorkCon = oMat01.Columns("WorkCon").Cells(i).Specific.VALUE
            
                Query01 = "EXEC PS_PP035_80 '" & Param01 & "', '" & WorkCon & "'"
                Call RecordSet01.DoQuery(Query01)
                Call MDC_Com.MDC_GF_Message("작업상태를 수정하였습니다.", "S")
            
            Else '작업상태 이외의 데이터 수정 시는 재고 체크 필요
            
                '재고가 존재하면 닫기(종료) 불가 기능 추가(2012.01.11 송명규 추가)
                If Trim(oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE) = "104" Then '//멀티
                    QueryString = "                     Select      Sum(a.Quantity * (Case When a.Direction = '0' Then 1 Else -1 End)) As Quantity"
                    QueryString = QueryString & "  from       IBT1 a "
                    QueryString = QueryString & "               Inner Join "
                    QueryString = QueryString & "               OITM b "
                    QueryString = QueryString & "                   On a.ItemCode = b.ItemCode "
                    QueryString = QueryString & "                   And b.U_ItmBsort = '104'"
                    QueryString = QueryString & "  Where     a.BaseType In ('59', '60')"
                    QueryString = QueryString & "               and a.BatchNum = '" & Trim(oMat01.Columns("OrdNum").Cells(i).Specific.VALUE) & "'"
                Else
                    QueryString = "                     SELECT      SUM(A.InQty) - SUM(A.OutQty) AS [StockQty]"
                    QueryString = QueryString & "  FROM       OINM AS A"
                    QueryString = QueryString & "                 INNER JOIN"
                    QueryString = QueryString & "                 OITM As B"
                    QueryString = QueryString & "                     ON A.ItemCode = B.ItemCode"
                    QueryString = QueryString & "  WHERE      B.U_ItmBsort IN ('105','106')"
                    QueryString = QueryString & "                 AND A.ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i).Specific.VALUE) & "'"
                    QueryString = QueryString & "  GROUP BY  A.ItemCode"
                End If
                  
                If IIf((MDC_PS_Common.GetValue(QueryString, 0, 1)) = "", 0, (MDC_PS_Common.GetValue(QueryString, 0, 1))) > 0 Then
                
                    Call MDC_Com.MDC_GF_Message(i & "행의 작업지시는 현재 재고가 존재합니다. 처리할 수 없습니다.", "W")
            
                Else
                
    '                Call MDC_Com.MDC_GF_Message("먼저 체크박스(선택)을 누르고 수정하세요!", "E")
                    Param01 = oMat01.Columns("DocEntry").Cells(i).Specific.VALUE
                    Param02 = oMat01.Columns("Canceled").Cells(i).Specific.VALUE
                    Param03 = oMat01.Columns("SelWt").Cells(i).Specific.VALUE
                    Param04 = oMat01.Columns("CntcCode").Cells(i).Specific.VALUE
                    Param05 = oMat01.Columns("CntcName").Cells(i).Specific.VALUE
                    Param06 = oMat01.Columns("DocDate").Cells(i).Specific.VALUE
                    Param07 = oMat01.Columns("DueDate").Cells(i).Specific.VALUE
                    ItemCode = Trim(oMat01.Columns("ItemCode").Cells(i).Specific.VALUE)
                    ItemName = MDC_PS_Common.Make_ItemName(Trim(oMat01.Columns("ItemName").Cells(i).Specific.VALUE))
'                    WorkCon = oMat01.Columns("WorkCon").Cells(i).Specific.VALUE
                    
                    Query01 = "EXEC PS_PP035_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "', '" & ItemCode & "', '" & ItemName & "'"
                    Call RecordSet01.DoQuery(Query01)
                    Call MDC_Com.MDC_GF_Message("데이터를 수정하였습니다.", "S")
                
                End If

            End If

        End If
        
    Next
        
    Exit Sub
    
PS_PP035_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP035_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Chk_Data()
On Error GoTo Chk_Data_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim i As Long
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '문서번호 입력된 임시테이블 삭제
    sQry = "DELETE Temp_LBG10"
    oRecordSet01.DoQuery sQry
    
    '/ 조회화면에서 선택한 문서번호만 임시테이블에 삽입
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        oDS_PS_PP035L.Offset = i
        If Trim(oDS_PS_PP035L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "INSERT INTO Temp_LBG10 (DocEntry) VALUES  (" & Trim(oDS_PS_PP035L.GetValue("U_ColReg02", i)) & ")"
            oRecordSet01.DoQuery sQry
        End If
    Next i
    
    sQry = "EXEC [PS_PP035_03]"
    
    Call ConnectODBC
    'DocNum = oForm01.Items("8").Specific.Value
    WinTitle = "[PS_PP035] 자주검사 CHECK SHEET"
    ReportName = "PS_PP035_03.rpt"
    sQry = sQry
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)

    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Set oRecordSet01 = Nothing
    Exit Sub
Chk_Data_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "Chk_Data_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Chk_Data1()
On Error GoTo Chk_Data1_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim i As Long
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
If oForm01.Items("OrdGbn").Specific.VALUE = "105" Or oForm01.Items("OrdGbn").Specific.VALUE = "106" Then
    
    '문서번호 입력된 임시테이블 삭제
    sQry = "DELETE Temp_LBG11"
    oRecordSet01.DoQuery sQry
    
    '/ 조회화면에서 선택한 문서번호만 임시테이블에 삽입
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        oDS_PS_PP035L.Offset = i
        If Trim(oDS_PS_PP035L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "INSERT INTO Temp_LBG11 (DocEntry) VALUES  (" & Trim(oDS_PS_PP035L.GetValue("U_ColReg02", i)) & ")"
            
            oRecordSet01.DoQuery sQry
        End If
    Next i
    
    sQry = "EXEC [PS_PP035_04] 'M'"
    
    Call ConnectODBC
    'DocNum = oForm01.Items("8").Specific.Value
    WinTitle = "[PS_PP034] 공정작업지시서"
    ReportName = "PS_PP035_04.rpt"

    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(2)
    ReDim gRpt_SRptName(2)
    
    gRpt_SRptSqry(1) = "EXEC [PS_PP035_04] 'S'"
    gRpt_SRptName(1) = "Sub_Report01"
    
    gRpt_SRptSqry(2) = "EXEC [PS_PP035_04] 'E'"
    gRpt_SRptName(2) = "Sub_Report02"
    
    ReDim gRpt_SFormula(2, 2)
    ReDim gRpt_SFormula_Value(2, 2)

ElseIf oForm01.Items("OrdGbn").Specific.VALUE = "101" Or oForm01.Items("OrdGbn").Specific.VALUE = "102" Then
    
    sQry = "DELETE Temp_LBG12"
    oRecordSet01.DoQuery sQry
    
    '/ 조회화면에서 선택한 문서번호만 임시테이블에 삽입
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        oDS_PS_PP035L.Offset = i
        If Trim(oDS_PS_PP035L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "INSERT INTO Temp_LBG12 (DocEntry) VALUES  (" & Trim(oDS_PS_PP035L.GetValue("U_ColReg02", i)) & ")"
            oRecordSet01.DoQuery sQry
        End If
    Next i
    
    sQry = "EXEC [PS_PP035_05] 'M'"
    
    Call ConnectODBC
    'DocNum = oForm01.Items("8").Specific.Value
    WinTitle = "[PS_PP035] 작업지시서"
    ReportName = "PS_PP035_05.rpt"

    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    gRpt_SRptSqry(1) = "EXEC [PS_PP035_05] 'S'"
    gRpt_SRptName(1) = "Sub_Report01"
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)

ElseIf oForm01.Items("OrdGbn").Specific.VALUE = "107" Then
    
    sQry = "DELETE Temp_LBG13"
    oRecordSet01.DoQuery sQry
    
    '/ 조회화면에서 선택한 문서번호만 임시테이블에 삽입
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        oDS_PS_PP035L.Offset = i
        If Trim(oDS_PS_PP035L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "INSERT INTO Temp_LBG13 (DocEntry) VALUES  (" & Trim(oDS_PS_PP035L.GetValue("U_ColReg02", i)) & ")"
            oRecordSet01.DoQuery sQry
        End If
    Next i
    
    sQry = "EXEC [PS_PP035_06] 'M'"
    
    Call ConnectODBC
    'DocNum = oForm01.Items("8").Specific.Value
    WinTitle = "[PS_PP035] End Bearing  공정카드"
    ReportName = "PS_PP035_06.rpt"

    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    gRpt_SRptSqry(1) = "EXEC [PS_PP035_06] 'S'"
    gRpt_SRptName(1) = "Sub_Report01"
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)

ElseIf oForm01.Items("OrdGbn").Specific.VALUE = "104" Then
    
    sQry = "DELETE Temp_LBG14"
    oRecordSet01.DoQuery sQry
    
    '/ 조회화면에서 선택한 문서번호만 임시테이블에 삽입
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        oDS_PS_PP035L.Offset = i
        If Trim(oDS_PS_PP035L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "INSERT INTO Temp_LBG14 (DocEntry) VALUES  (" & Trim(oDS_PS_PP035L.GetValue("U_ColReg02", i)) & ")"
            oRecordSet01.DoQuery sQry
        End If
    Next i
    
    sQry = "EXEC [PS_PP035_07] "
    
    Call ConnectODBC
    WinTitle = "[PS_PP035] M/G 공정카드"
    ReportName = "PS_PP035_07.rpt"

    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
End If

    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Set oRecordSet01 = Nothing
    Exit Sub

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

Chk_Data1_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "Chk_Data1_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
