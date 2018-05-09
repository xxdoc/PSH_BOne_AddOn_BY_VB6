VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM130.cls
'//  Module         : 재고관리 > 외주반출등록
'//  Desc           : MM
'//  FormType       : PS_MM130
'//  Create Date    : 2010.10.13
'//  Creator        : Dong sub Kim
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_MM130H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_MM130L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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

' 입고 DI를 위한 정보를 가지는 구조체
Private Type StockInfos
    CardCode As String            '고객코드
    ItemCode As String            '품목코드
    FromWarehouseCode As String   '창고코드
    ToWarehouseCode As String     '창고코드
    Weight  As Double             '중량
    UnWeight As Double
    BatchNum As String            '배치번호
    BatchWeight As Double         '배치중량
    Qty   As Long                 '수량
    TransNo As String             '재고이전문서번호
    Chk As String
    MatrixRow As Long
    StockTransDocEntry  As String '재고이전문서번호
    StockTransLineNum   As String '재고이전라인번호
    Indate  As String             '전기일
End Type

Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long
Private StockInfo() As StockInfos

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM130.srf")
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
    
    oFormUniqueID01 = "PS_MM130_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01                 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry"    '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_MM130_CreateItems
    Call PS_MM130_ComboBox_Setting
    Call PS_MM130_Initial_Setting
    Call PS_MM130_CF_ChooseFromList
    Call PS_MM130_EnableMenus
    Call PS_MM130_SetDocument(oFromDocEntry01)
    Call PS_MM130_FormResize
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    
    '//커서를 첫번째 ITEM으로 지정
    oForm01.ActiveItem = "BPLId"
    
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
                If (PS_MM130_Validate("취소") = False) Then
                    BubbleEvent = False
                    Exit Sub
                End If
                If oForm01.Mode = fm_ADD_MODE Then
                    BubbleEvent = False
                    Exit Sub
                End If
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
'                Call oDS_PS_MM130H.setValue("U_BPLID", 0, "1")
'                Call oDS_PS_MM130H.setValue("U_CardCode", 0, Trim(Sbo_Company.UserName))
                Call PS_MM130_FormItemEnabled '//UDO방식
            Case "1282": '추가
                oForm01.Freeze True
                
                Call oDS_PS_MM130H.setValue("U_OutGbn", 0, "10")
                Call oDS_PS_MM130H.setValue("U_DocDate", 0, Format(Now, "YYYYMMDD"))
                Call oDS_PS_MM130H.setValue("U_BPLId", 0, "1")
                Call oDS_PS_MM130H.setValue("U_OKYNC", 0, "N")
                
                Call PS_MM130_FormItemEnabled '//UDO방식
                Call PS_MM130_AddMatrixRow(0, True) '//UDO방식
                
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
'    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
'        Select Case BusinessObjectInfo.EventType
'            Case et_FORM_DATA_LOAD: '//33
'            Case et_FORM_DATA_ADD: '//34
'            Case et_FORM_DATA_UPDATE: '//35
'            Case et_FORM_DATA_DELETE: '//36
'        End Select
'    ElseIf (BusinessObjectInfo.BeforeAction = False) Then '//BeforeAction = False
'        Select Case BusinessObjectInfo.EventType
'            Case et_FORM_DATA_LOAD: '//33
'            Case et_FORM_DATA_ADD: '//34
'            Case et_FORM_DATA_UPDATE: '//35
'            Case et_FORM_DATA_DELETE: '//36
'        End Select
'    End If
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
    
        If pval.ItemUID = "PS_MM130" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_MM130_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If PS_MM130_DataInsertCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                '//외주 재고 이동 DI API
                If Trim(oForm01.Items("OKYNC").Specific.VALUE) = "N" And oForm01.Items("STDocNum").Specific.VALUE = "" Then
                    Dim i As Integer
                    Dim j As Integer
                    For i = 0 To oMat01.RowCount - 1
                         If oMat01.Columns("OutGbn").Cells(i + 1).Specific.VALUE = "10" Then
                            If j = 0 Then
                                j = j + 1
                                If PS_MM130_StockTrans() = True Then
                                     Call PS_MM130_UpdateUserField
                                Else
                                    Call PS_MM130_AddMatrixRow(oMat01.VisualRowCount, False)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                         End If
                     Next
                End If
     
            '//외주 재고 이동 취소 DI API
            ElseIf (oForm01.Mode = fm_UPDATE_MODE) And Trim(oForm01.Items("OKYNC").Specific.VALUE) = "C" And oForm01.Items("STDocNum").Specific.VALUE <> "" Then
                    If PS_MM130_DataValidCheck = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If PS_MM130_Cancel_oStockTrans(2) = True Then
                        Call PS_MM130_Update_Cancel
                    Else
                        Call PS_MM130_AddMatrixRow(0, False)
                        BubbleEvent = False
                        Exit Sub
                    End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then   '//창원 문서 수정
                If PS_MM130_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        '// 출고증
        If pval.ItemUID = "Print" Then
            Call Print_Query
        End If
   
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_MM130" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_MM130_FormItemEnabled
                    Call PS_MM130_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
                    oForm01.Items("Rad01").Specific.Selected = True
                    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
                    Call oForm01.Items("BPLId").Specific.Select("0", psk_Index)
                    Call oForm01.Items("OKYNC").Specific.Select("N", psk_ByValue)
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_MM130_FormItemEnabled
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
    Dim TempForm01 As Variant
    Dim RadioGrp As String
    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "Mat01" Then
                If pval.ColUID = "OrdNum" Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = "" Then
        '                oForm01.Items("Rad01").Enabled = False
        '                oForm01.Items("Rad02").Enabled = False
                        Set TempForm01 = New PS_MM131
                        If Trim(oDS_PS_MM130H.GetValue("U_OutGbn", 0)) = "10" Then
                            RadioGrp = "A"
                        ElseIf Trim(oDS_PS_MM130H.GetValue("U_OutGbn", 0)) = "20" Then
                            RadioGrp = "B"
                        Else
                        End If
                        
                        Call TempForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, RadioGrp)
                        Call PS_MM130_AddMatrixRow(0, True) '//UDO방식일때
                        BubbleEvent = False
                    End If
                ElseIf pval.ColUID = "TCpCode" Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                    
            End If
        End If
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "")
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "")
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "OutWhCd")
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "InWhCd")
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "OutItmCd")
    ElseIf pval.BeforeAction = False Then
    
    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
'        If pval.ItemChanged = True Then
'            Call oForm01.Freeze(True)
'            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_MM130L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM130L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_MM130_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_MM130L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                End If
'            Else
'                If (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_MM130H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_MM130H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_MM130H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_MM130H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                End If
'            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
'            Call oForm01.Freeze(False)
'        End If
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
'        If pval.ItemUID = "Rad01" Or pval.ItemUID = "Rad02" Then
'            oDS_PS_MM130L.Clear
'            oMat01.LoadFromDataSource
'            oMat01.FlushToDataSource
'            Call PS_MM130_AddMatrixRow(0, True) '//UDO방식
'        End If
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
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    Call oForm01.Freeze(True)
   
    Dim Query01             As String
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
    Dim Query02             As String
    Dim oRecordset02        As SAPbobsCOM.Recordset
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
    
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "OrdNum") Then
                    '//기타작업
                    Call oDS_PS_MM130L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM130L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_MM130_AddMatrixRow (pval.Row)
                    End If
'                    '//아이템
'                    If Trim(oDS_PS_MM130H.GetValue("U_OutGbn", 0)) = "10" Then
'                       oMat01.Columns("OutItmCd").Editable = True
'                    ElseIf Trim(oDS_PS_MM130H.GetValue("U_OutGbn", 0)) = "20" Then
'                       oMat01.Columns("OutItmCd").Editable = False
'                    Else
'                    End If
                    oMat01.LoadFromDataSource
                ElseIf (pval.ColUID = "OutWhCd") Then
                    Call oDS_PS_MM130L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_MM130L.setValue("U_OUtWhNm", pval.Row - 1, MDC_GetData.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"))
                    oMat01.LoadFromDataSource
                ElseIf (pval.ColUID = "InWhCd") Then
                    Call oDS_PS_MM130L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_MM130L.setValue("U_InWhNm", pval.Row - 1, MDC_GetData.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"))
                    oMat01.LoadFromDataSource
                ElseIf (pval.ColUID = "OutQty") Then
                    Dim Count As Double
                    Count = oMat01.Columns("UnWeight").Cells(pval.Row).Specific.VALUE * oMat01.Columns("OutQty").Cells(pval.Row).Specific.VALUE
                    Call oDS_PS_MM130L.setValue("U_OutQty", pval.Row - 1, oMat01.Columns("OutQty").Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_MM130L.setValue("U_OutWt", pval.Row - 1, Count)
                    oMat01.LoadFromDataSource
                    oMat01.Columns("OutQty").Cells(pval.Row).Click ct_Regular
                     oMat01.FlushToDataSource
                     If oMat01.Columns("OutQty").Cells(pval.Row).Specific.VALUE > 0 Then
                        Query01 = "Select U_ObasUnit FROM OITM WHERE ItemCode = '" & Trim(oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE) & "'"
                        Call oRecordSet01.DoQuery(Query01)
                        Query02 = "Select OnHand, U_Qty FROM OITW WHERE ItemCode = '" & Trim(oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE) & "' AND WhsCode = '101'"
                        Call oRecordset02.DoQuery(Query02)
                        If Left(oRecordSet01.Fields(0).VALUE, 1) = "1" Then
                            If (Trim(oRecordset02.Fields(0).VALUE)) > 0 And oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE <> "" Then
                                oMat01.Columns("OutWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("OutQty").Cells(pval.Row).Specific.VALUE
                            Else
                                oMat01.Columns("OutWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("OutQty").Cells(pval.Row).Specific.VALUE
                            End If
                        ElseIf Left(oRecordSet01.Fields(0).VALUE, 1) = "2" Then
                            If (Trim(oRecordset02.Fields(1).VALUE)) > 0 And oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE <> "" Then
                                oMat01.Columns("OutWt").Cells(pval.Row).Specific.VALUE = (Trim(oRecordset02.Fields(0).VALUE) / Trim(oRecordset02.Fields(1).VALUE)) * oMat01.Columns("OutQty").Cells(pval.Row).Specific.VALUE
                            Else
                                oMat01.Columns("OutWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("OutWt").Cells(pval.Row).Specific.VALUE
                            End If
                        Else
                        End If
                        Call oDS_PS_MM130L.setValue("U_OutQty", pval.Row - 1, oMat01.Columns("OutQty").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_MM130L.setValue("U_OutWt", pval.Row - 1, oMat01.Columns("OutWt").Cells(pval.Row).Specific.VALUE)
                        oMat01.LoadFromDataSource
                        oMat01.Columns("OutWt").Cells(pval.Row).Click ct_Regular
                    End If
                ElseIf (pval.ColUID = "TCpCode") Then
                    Query01 = "Select U_CpName from [@PS_PP001L] Where U_CpCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"
                    Call oRecordSet01.DoQuery(Query01)
                    Call oDS_PS_MM130L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_MM130L.setValue("U_TCpName", pval.Row - 1, oRecordSet01.Fields(0).VALUE)
                    oMat01.LoadFromDataSource
                Else
                    Call oDS_PS_MM130L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    oMat01.LoadFromDataSource
                    oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                End If
            
            ElseIf (pval.ItemUID = "CardCode") Then
                Query01 = "SELECT CardName FROM [OCRD] WHERE CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                Call oRecordSet01.DoQuery(Query01)
                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                Set oRecordSet01 = Nothing
            ElseIf (pval.ItemUID = "CntcCode") Then
                Query01 = "Select U_FULLNAME, U_MSTCOD From [OHEM] Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
                Call oRecordSet01.DoQuery(Query01)
                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                Set oRecordSet01 = Nothing
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_MM130H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
            End If
            
            oMat01.AutoResizeColumns
            oForm01.Update
        End If
    ElseIf pval.BeforeAction = False Then

    End If
    
    Set oRecordSet01 = Nothing
    Set oRecordset02 = Nothing
    
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
        Call PS_MM130_FormItemEnabled
        Call PS_MM130_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        
        Call PS_MM130_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    
    Dim oDataTable01  As SAPbouiCOM.DataTable
    Dim oDataTable02  As SAPbouiCOM.DataTable
    
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        If (pval.ItemUID = "CardCode") Then
            Set oDataTable01 = pval.SelectedObjects
            If oDataTable01 Is Nothing Then
            Else
'               oForm01.Items("CardCode").Specific.Value = oDataTable01.Columns("CardCode").Cells(0).Value
                Call oDS_PS_MM130H.setValue("U_CardCode", 0, oDataTable01.Columns("CardCode").Cells(0).VALUE)
                Call oDS_PS_MM130H.setValue("U_CardName", 0, oDataTable01.Columns("CardName").Cells(0).VALUE)
                ' // 찾기나 문서이동 버튼 클릭 시에 갱신으로 바뀌지 않음
                If oForm01.Mode <> fm_ADD_MODE Then oForm01.Mode = fm_UPDATE_MODE
'               oForm01.DataSources.UserDataSources("CardCode").Value = oDataTable01.Columns("CardCode").Cells(0).Value
'               oForm01.DataSources.UserDataSources("CardName").Value = oDataTable01.Columns("CardName").Cells(0).Value
            End If
        ElseIf (pval.ItemUID = "ShipTo") Then
            Set oDataTable02 = pval.SelectedObjects
            If oDataTable02 Is Nothing Then
            Else
                Call oDS_PS_MM130H.setValue("U_ShipTo", 0, oDataTable02.Columns("CardCode").Cells(0).VALUE)
                Call oDS_PS_MM130H.setValue("U_ShipNm", 0, oDataTable02.Columns("CardName").Cells(0).VALUE)
                If oForm01.Mode <> fm_ADD_MODE Then oForm01.Mode = fm_UPDATE_MODE
'               If oForm01.Mode <> fm_ADD_MODE Then oForm01.Mode = fm_UPDATE_MODE
'               oForm01.DataSources.UserDataSources("ShipTo").Value = oDataTable02.Columns("CardCode").Cells(0).Value
'               oForm01.DataSources.UserDataSources("ShipNm").Value = oDataTable02.Columns("CardName").Cells(0).Value
            End If
        End If
        oForm01.Update
    End If
    Set oDataTable01 = Nothing
    Set oDataTable02 = Nothing
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
'            If (PS_MM130_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_MM130L.RemoveRecord(oDS_PS_MM130L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_MM130_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_MM130L.GetValue("U_OrdNum", oMat01.RowCount - 1)) <> "" Then
                    Call PS_MM130_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_MM130_CreateItems() As Boolean
On Error GoTo PS_MM130_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim optBtn          As SAPbouiCOM.OptionBtn
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM130H = oForm01.DataSources.DBDataSources("@PS_MM130H")
    Set oDS_PS_MM130L = oForm01.DataSources.DBDataSources("@PS_MM130L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    ' // 매트릭스 비활성 상태에서 값이 안 나올 경우
    ' 매트릭스 선택 포맷을 반환하거나 설정
    ' 매트릭스 선택 규칙을 자체적으로 선택할 수 없는 경우 반환되는 값
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns ' 매트릭스의 칼럼 크기를 자동으로 정렬
    
    Set optBtn = oForm01.Items("Rad01").Specific     ' 원재료반출
    optBtn.ValOn = "10"                              ' 옵션버튼이 선택 되었을 때 값을 반환
    optBtn.ValOff = "0"                              ' 옵션버튼이 선택되지 않았을 때의 값을 반환
    optBtn.Selected = True                           ' 옵션버튼이 선택되었는지 여부를 반환
    
    Set optBtn = oForm01.Items("Rad02").Specific     ' 재공반출
    optBtn.ValOn = "20"
    optBtn.ValOff = "0"
    optBtn.GroupWith ("Rad01")                       ' 위의 "Rad01"과 같이 그룹을 지정
    Set optBtn = Nothing
    
    oForm01.Settings.MatrixUID = "Mat01"             ' 서식세팅
    oForm01.Settings.Enabled = True
    oForm01.Settings.EnableRowFormat = True
    
    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
    oForm01.Items("ShipCo").Specific.VALUE = "업체자가"
    
    
    Set optBtn = Nothing
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    
    
    
    Exit Function
PS_MM130_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM130_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_MM130_ComboBox_Setting()
On Error GoTo PS_MM130_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//콤보에 기본값설정
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "Mat01", "OutGbn", "10", "원재료")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "Mat01", "OutGbn", "20", "제공")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("OutGbn"), "PS_MM130", "Mat01", "OutGbn")
      
    Set oCombo = oForm01.Items("OKYNC").Specific
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "OKYNC", "", "N", "재고이동")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "OKYNC", "", "C", "이동취소")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "OKYNC", "", "Y", "승인")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("OKYNC").Specific, "PS_MM130", "OKYNC")
    Call oCombo.Select(0, psk_Index)
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL WHERE BPLId = '1' ORDER BY BPLId", "1", False, False)

    Call oForm01.Freeze(False)
    Exit Sub
PS_MM130_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM130_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM130_CF_ChooseFromList()
On Error GoTo PS_MM130_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFL01                As SAPbouiCOM.ChooseFromList
    Dim oCFL02                As SAPbouiCOM.ChooseFromList
    Dim oCons01               As SAPbouiCOM.Conditions
    Dim oCon01                As SAPbouiCOM.Condition
    Dim oCFLs01               As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLs02               As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLCreationParams01  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oCFLCreationParams02  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit01               As SAPbouiCOM.EditText
    Dim oEdit02               As SAPbouiCOM.EditText


'    Set oEdit01 = oForm01.Items("CardCode").Specific
'    Set oCFLs01 = oForm01.ChooseFromLists
'    Set oCFLCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams01.ObjectType = lf_BusinessPartner
'    oCFLCreationParams01.uniqueID = "CFLCARDCODE"
'    oCFLCreationParams01.MultiSelection = False
'    Set oCFL01 = oCFLs01.Add(oCFLCreationParams01)
'
'    ' Choose from list 에 조건을 줄 경우
'    ' choosefromlist가 화면에 나오면 서식세팅으로 원하는 필드값 추가 가능
'    Set oCons01 = oCFL01.GetConditions()
'    Set oCon01 = oCons01.Add()
'    oCon01.Alias = "CardCode"                                              ' Condition Field
'    oCon01.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL            ' Equal
'    oCon01.CondVal = "6"                                                   ' Condition Value
'    oCFL01.SetConditions oCons01
'
'    oEdit01.ChooseFromListUID = "CFLCARDCODE"
'    oEdit01.ChooseFromListAlias = "CardCode"
    
    Set oEdit02 = oForm01.Items("ShipTo").Specific
    Set oCFLs02 = oForm01.ChooseFromLists
    Set oCFLCreationParams02 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams02.ObjectType = lf_BusinessPartner
    oCFLCreationParams02.uniqueID = "CFLSHIPCODE"
    oCFLCreationParams02.MultiSelection = False
    Set oCFL02 = oCFLs02.Add(oCFLCreationParams02)

    oEdit02.ChooseFromListUID = "CFLSHIPCODE"
    oEdit02.ChooseFromListAlias = "CardCode"                ' 정의된 오브젝트만 사용 가능
    
    Exit Sub
PS_MM130_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM130_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
    
Sub PS_MM130_Initial_Setting()
On Error GoTo PS_MM130_Initial_Setting_Error
    ' 사업장
    Dim lcl_User_BPLId As String
    lcl_User_BPLId = MDC_PS_Common.User_BPLId()
    
    If lcl_User_BPLId = "1" Then '소속사업장이 창원일 때만 콤보박스 Select (2011.09.21 송명규 추가)
        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    End If
 ' 인수자
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
    Exit Sub
PS_MM130_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM130_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM130_FormItemEnabled()
On Error GoTo PS_MM130_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        Call PS_MM130_FormClear                      '//UDO방식
        Call oForm01.EnableMenu("1281", True)        '//찾기
        Call oForm01.EnableMenu("1282", False)       '//추가
        oForm01.Items("Comments").Click ct_Regular
        oForm01.EnableMenu ("1293"), True            '// 행삭제
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("OutDoc").Enabled = False
        oForm01.Items("Print").Enabled = False
        If Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "C" Or Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "Y" Then
            Call oForm01.EnableMenu("1284", False)       '//취소
            Call oForm01.EnableMenu("1286", False)       '//닫기
            Call oForm01.EnableMenu("1293", False)       '//행삭제
        Else
            Call oForm01.EnableMenu("1284", True)       '//취소
            Call oForm01.EnableMenu("1286", True)       '//닫기
            Call oForm01.EnableMenu("1293", True)       '//행삭제
        End If
        
        '외주업체
        If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Or Trim(Sbo_Company.UserName) = "66510" Then
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("Print").Enabled = False
            oForm01.Items("Rad01").Enabled = False
            oForm01.Items("Rad02").Enabled = False
            oForm01.Items("CardCode").Enabled = False
            oForm01.Items("CntcCode").Enabled = False
            oForm01.Items("ShipTo").Enabled = False
            oForm01.Items("CarNo").Enabled = False
            oForm01.Items("ShipCo").Enabled = False
            oForm01.Items("DocDate").Enabled = False
            oForm01.Items("Fare").Enabled = False
            oForm01.Items("ArrivePl").Enabled = False
            oForm01.Items("OKYNC").Enabled = True
            oMat01.Columns("OrdNum").Editable = False
            oMat01.Columns("OutItmCd").Editable = False
            oMat01.Columns("OutQty").Editable = False
            oMat01.Columns("OutWt").Editable = False
            oMat01.Columns("OutWhCd").Editable = False
            oMat01.Columns("OutWhNm").Editable = False
            oMat01.Columns("InWhCd").Editable = False
            oMat01.Columns("InWhNm").Editable = False
        '신동
        Else
'            If Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "C" Or Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "N" Then
'                oForm01.Items("CardCode").Enabled = False
'                oForm01.Items("CntcCode").Enabled = False
'                oForm01.Items("ShipTo").Enabled = False
'                oForm01.Items("CarNo").Enabled = False
'                oForm01.Items("ShipCo").Enabled = False
'                oForm01.Items("DocDate").Enabled = False
'                oForm01.Items("Fare").Enabled = False
'                oForm01.Items("ArrivePl").Enabled = False
'                oMat01.Columns("OrdNum").Editable = False
'                oMat01.Columns("OutQty").Editable = False
'                oMat01.Columns("OutWt").Editable = False
'                oMat01.Columns("OutWhCd").Editable = False
'                oMat01.Columns("OutWhNm").Editable = False
'                oMat01.Columns("InWhCd").Editable = False
'                oMat01.Columns("InWhNm").Editable = False
'                oForm01.Items("Rad01").Enabled = False
'                oForm01.Items("Rad02").Enabled = False
'                oForm01.Items("OKYNC").Enabled = False
'                oMat01.Columns("OrdNum").Editable = True
''            Else
                oForm01.Items("CardCode").Enabled = True
                oForm01.Items("CntcCode").Enabled = True
                oForm01.Items("ShipTo").Enabled = True
                oForm01.Items("CarNo").Enabled = True
                oForm01.Items("ShipCo").Enabled = True
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("Fare").Enabled = True
                oForm01.Items("ArrivePl").Enabled = True
                oMat01.Columns("OrdNum").Editable = True
                oMat01.Columns("OutQty").Editable = True
                oMat01.Columns("OutWt").Editable = True
                oMat01.Columns("OutWhCd").Editable = True
                oMat01.Columns("OutWhNm").Editable = True
                oMat01.Columns("InWhCd").Editable = True
                oMat01.Columns("InWhNm").Editable = True
                oForm01.Items("Rad01").Enabled = True
                oForm01.Items("Rad02").Enabled = True
                oForm01.Items("OKYNC").Enabled = True
                oMat01.Columns("OrdNum").Editable = True
'           End If
        End If
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True)      '//찾기
        Call oForm01.EnableMenu("1282", True)       '//추가
        oForm01.Items("Print").Enabled = False
        oForm01.EnableMenu ("1293"), True           '// 행삭제
        
        If Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "C" Or Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "Y" Then
            Call oForm01.EnableMenu("1284", False)       '//취소
            Call oForm01.EnableMenu("1286", False)       '//닫기
            Call oForm01.EnableMenu("1293", False)       '//행삭제
        Else
            Call oForm01.EnableMenu("1284", True)       '//취소
            Call oForm01.EnableMenu("1286", True)       '//닫기
            Call oForm01.EnableMenu("1293", True)       '//행삭제
        End If
        
      '외주업체
        If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Or Trim(Sbo_Company.UserName) = "66510" Then

            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("CardCode").Specific.VALUE = Trim(Sbo_Company.UserName)
            oForm01.Items("Comments").Click ct_Regular
            oForm01.Items("CardCode").Enabled = False
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("OKYNC").Enabled = True
        Else
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("DocEntry").Enabled = True
            oForm01.Items("OutDoc").Enabled = True
            oMat01.Columns("OrdNum").Editable = True
            oForm01.Items("Rad01").Enabled = False
            oForm01.Items("Rad02").Enabled = False
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("ShipTo").Enabled = True
            oForm01.Items("CarNo").Enabled = True
            oForm01.Items("ShipCo").Enabled = True
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("Fare").Enabled = True
            oForm01.Items("ArrivePl").Enabled = True
            oForm01.Items("OKYNC").Enabled = True
        End If
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True)  '//찾기
        Call oForm01.EnableMenu("1282", True)  '//추가
        Call oForm01.EnableMenu("1293", True)          '// 행삭제
        oForm01.Items("Print").Enabled = True

        If Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "N" Or Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "Y" Then
            Call oForm01.EnableMenu("1284", False)       '//취소
            Call oForm01.EnableMenu("1286", False)       '//닫기
            Call oForm01.EnableMenu("1293", False)       '//행삭제
        Else
            Call oForm01.EnableMenu("1284", True)       '//취소
            Call oForm01.EnableMenu("1286", True)       '//닫기
            Call oForm01.EnableMenu("1293", True)       '//행삭제
        End If
        
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("Comments").Click ct_Regular
        oForm01.Items("OutDoc").Enabled = False
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Rad01").Enabled = False
        oForm01.Items("Rad02").Enabled = False
        oForm01.Items("DocEntry").Enabled = False
        
        '외주업체
        If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Or Trim(Sbo_Company.UserName) = "66510" Then
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("CardCode").Enabled = False
            oForm01.Items("CntcCode").Enabled = False
            oForm01.Items("ShipTo").Enabled = False
            oForm01.Items("CarNo").Enabled = False
            oForm01.Items("ShipCo").Enabled = False
            oForm01.Items("DocDate").Enabled = False
            oForm01.Items("Fare").Enabled = False
            oForm01.Items("ArrivePl").Enabled = False
            If Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "C" Or Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "Y" Then
                oForm01.Items("OKYNC").Enabled = False
            Else
                oForm01.Items("OKYNC").Enabled = True
            End If
            oMat01.Columns("OrdNum").Editable = False
            oMat01.Columns("OutItmCd").Editable = False
            oMat01.Columns("OutQty").Editable = False
            oMat01.Columns("OutWt").Editable = False
            oMat01.Columns("OutWhCd").Editable = False
            oMat01.Columns("OutWhNm").Editable = False
            oMat01.Columns("InWhCd").Editable = False
            oMat01.Columns("InWhNm").Editable = False
        '신동
        Else
            If Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "C" Or Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "Y" Then
                If Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "N" Then
                    oForm01.Items("OKYNC").Enabled = True
                ElseIf Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "C" Or Trim(oDS_PS_MM130H.GetValue("U_OKYNC", 0)) = "Y" Then
                    oForm01.Items("OKYNC").Enabled = False
                End If
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("CntcCode").Enabled = False
                oForm01.Items("ShipTo").Enabled = False
                oForm01.Items("CarNo").Enabled = False
                oForm01.Items("ShipCo").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("Fare").Enabled = False
                oForm01.Items("ArrivePl").Enabled = False
                oMat01.Columns("OrdNum").Editable = False
                oMat01.Columns("OutQty").Editable = False
                oMat01.Columns("OutWt").Editable = False
                oMat01.Columns("OutWhCd").Editable = False
                oMat01.Columns("OutWhNm").Editable = False
                oMat01.Columns("InWhCd").Editable = False
                oMat01.Columns("InWhNm").Editable = False

            Else
                oForm01.Items("Rad01").Enabled = True
                oForm01.Items("Rad02").Enabled = True
                oForm01.Items("CardCode").Enabled = True
                oForm01.Items("CntcCode").Enabled = True
                oForm01.Items("ShipTo").Enabled = True
                oForm01.Items("CarNo").Enabled = True
                oForm01.Items("ShipCo").Enabled = True
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("Fare").Enabled = True
                oForm01.Items("ArrivePl").Enabled = True
                oMat01.Columns("OrdNum").Editable = True
                oMat01.Columns("OutQty").Editable = True
                oMat01.Columns("OutWt").Editable = True
                oMat01.Columns("OutWhCd").Editable = True
                oMat01.Columns("OutWhNm").Editable = True
                oMat01.Columns("InWhCd").Editable = True
                oMat01.Columns("InWhNm").Editable = True
                oForm01.Items("OKYNC").Enabled = True
             End If
        End If
    End If
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM130_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM130_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM130_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_MM130_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM130L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM130L.Offset = oRow
    oDS_PS_MM130L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM130_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM130_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM130_FormClear()
On Error GoTo PS_MM130_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM130'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_MM130_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM130_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM130_EnableMenus()
On Error GoTo PS_MM130_EnableMenus_Error
'    //메뉴활성화
    Call oForm01.EnableMenu("1283", False)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    '//Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_MM130_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM130_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM130_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_MM130_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_MM130_FormItemEnabled
        Call PS_MM130_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_MM130_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_MM130_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM130_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_MM130_DataValidCheck() As Boolean
On Error GoTo PS_MM130_DataValidCheck_Error
    Dim i As Long
    
    oMat01.FlushToDataSource
    PS_MM130_DataValidCheck = False
    
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_MM130_DataValidCheck = False
        Exit Function
    End If
    
    If oForm01.Items("BPLId").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "사업장 코드는 필수입니다.", bmt_Short, True
        oForm01.Items("BPLId").Click ct_Regular
        PS_MM130_DataValidCheck = False
        Exit Function
    
    ElseIf oForm01.Items("CardCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "외주거래처 코드는 필수입니다.", bmt_Short, True
        oForm01.Items("CardCode").Click ct_Regular
        PS_MM130_DataValidCheck = False
        Exit Function
        
    ElseIf oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "전기일자는 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_MM130_DataValidCheck = False
        Exit Function
    End If
    
    If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Or Trim(Sbo_Company.UserName) = "66510" Then
        If Trim(oForm01.Items("OKYNC").Specific.VALUE) <> "Y" Then
            Sbo_Application.SetStatusBarMessage "이동승인상태 Y - 승인만 선택 할 수 있습니다.", bmt_Short, True
            oForm01.Items("OKYNC").Click ct_Regular
            PS_MM130_DataValidCheck = False
            Exit Function
        End If
    Else
'        If oForm01.Items("OKYNC").Specific.VALUE <> "N" And oForm01.Items("OKYNC").Specific.VALUE <> "C" Then
'            Sbo_Application.SetStatusBarMessage "이동승인상태 N - 재고이동 또는 C - 이동취소만 선택 할 수 있습니다.", bmt_Short, True
'            oForm01.Items("OKYNC").Click ct_Regular
'            PS_MM130_DataValidCheck = False
'            Exit Function
'        End If
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("OrdNum").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "작지번호는 필수입니다.", bmt_Short, True
            oMat01.Columns("OrdNum").Cells(i).Click ct_Regular
            PS_MM130_DataValidCheck = False
            Exit Function
        
        ElseIf (oMat01.Columns("OutQty").Cells(i).Specific.VALUE < 0) Or (oMat01.Columns("OutWt").Cells(i).Specific.VALUE < 0) Then
            Sbo_Application.SetStatusBarMessage "반출수(중)량은 필수입니다.", bmt_Short, True
            oMat01.Columns("OutQty").Cells(i).Click ct_Regular
            PS_MM130_DataValidCheck = False
            Exit Function
        ElseIf (oMat01.Columns("TCpCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "외주 마지막공정(외주공정구간)을 선택하셔야 합니다..", bmt_Short, True
            oMat01.Columns("TCpCode").Cells(i).Click ct_Regular
            PS_MM130_DataValidCheck = False
            Exit Function
       End If
    Next
    
    Call oDS_PS_MM130L.RemoveRecord(oDS_PS_MM130L.Size - 1)
    Call oMat01.LoadFromDataSource
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_MM130_FormClear
    End If
    PS_MM130_DataValidCheck = True
    Exit Function
PS_MM130_DataValidCheck_Error:
    PS_MM130_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_MM130_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
Function PS_MM130_DataInsertCheck() As Boolean
On Error GoTo PS_MM130_DataInsertCheck_Error
    PS_MM130_DataInsertCheck = False
    Dim DocEntry As String
    Dim oDate As String
    Dim oOutDoc As String
    Dim i As Integer

    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    sQry = "SELECT ISNULL(MAX(U_OutDoc), 0) FROM [@PS_MM130H] WHERE SubString(U_OutDoc,1,8) = '" & oDate & "'"

    oRecordSet01.DoQuery sQry

    oOutDoc = oRecordSet01.Fields(0).VALUE

    Set oRecordSet01 = Nothing
    
    If oOutDoc = 0 Then
        oOutDoc = oDate + "001"
        Call oDS_PS_MM130H.setValue("U_OutDoc", 0, oOutDoc)
    Else
        oOutDoc = oOutDoc + 1
        Call oDS_PS_MM130H.setValue("U_OutDoc", 0, oOutDoc)
    End If
        Call oDS_PS_MM130H.setValue("U_DocGbn", 0, "반출")
    Dim Count As Double
    
    For i = 1 To oMat01.VisualRowCount
        If oMat01.Columns("UnWeight").Cells(i).Specific.VALUE = 0 Then
            oMat01.Columns("UnWeight").Cells(i).Specific.VALUE = 1
        End If
        
        Call oDS_PS_MM130L.setValue("U_OutDoc", i - 1, oOutDoc)
    Next
    
    If Trim(oDS_PS_MM130H.GetValue("U_OutGbn", 0)) = "10" Then
        Call oDS_PS_MM130H.setValue("U_OutGbn", 0, "10")
    ElseIf Trim(oDS_PS_MM130H.GetValue("U_OutGbn", 0)) = "20" Then
        Call oDS_PS_MM130H.setValue("U_OutGbn", 0, "20")
    End If
    
    oMat01.LoadFromDataSource
    
    PS_MM130_DataInsertCheck = True
            
Exit Function
PS_MM130_DataInsertCheck_Error:
    PS_MM130_DataInsertCheck = False
    Sbo_Application.SetStatusBarMessage "PS_MM130_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_MM130_FormResize()
On Error GoTo PS_MM130_FormResize_Error
    
    Exit Sub
PS_MM130_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM130_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_MM130_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_MM130_Validate_Error
    PS_MM130_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

        If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_MM130H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
            MDC_Com.MDC_GF_Message "이미취소된 문서 입니다. 취소할수 없습니다.", "W"

        End If

    End If

    Set RecordSet01 = Nothing
    Exit Function
PS_MM130_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_MM130_Validate_Error:
    PS_MM130_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_MM130_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub Print_Query()
On Error GoTo Print_Query_Error
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim sQry01                      As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call ConnectODBC
    sQry01 = "SELECT COUNT(*) AS COUNT FROM [@PS_MM130L] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    Call oRecordSet01.DoQuery(sQry01)
    
    
    If Trim(oRecordSet01.Fields(0).VALUE) > 7 Then
        Set oRecordSet01 = Nothing
        WinTitle = "[PS_MM235_20] 레포트"
        ReportName = "PS_MM235_20.rpt"
        sQry = "EXEC PS_MM235_20 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        ReDim gRpt_Formula(1)
        ReDim gRpt_Formula_Value(1)
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        ReDim gRpt_SFormula(1, 1)
        ReDim gRpt_SFormula_Value(1, 1)
    Else
        Set oRecordSet01 = Nothing
        WinTitle = "[PS_MM235_10] 레포트"
        ReportName = "PS_MM235_10.rpt"
        sQry = "EXEC PS_MM235_10 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        ReDim gRpt_Formula(1)
        ReDim gRpt_Formula_Value(1)
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        ReDim gRpt_SFormula(1, 1)
        ReDim gRpt_SFormula_Value(1, 1)

    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM130_Print_Query_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_MM130_StockTrans() As Boolean
On Error GoTo PS_MM130_StockTrans_Error
    Dim RetVal As Long
    Dim ErrCode As Long
    Dim ErrMsg As String
    Dim lQuery As String
    Dim lRecordSet As SAPbobsCOM.Recordset
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim lMaxBatchNum As Long    '해당 품목의 최대 배치번호
    Dim lBatchWeight As Double  '배치별 중량
    Dim lTypeCount As Integer   '전체 StockInfo 구조체배열의 RowCount
    Dim i, j, K, Q, z, r  As Long
    Dim DocCnt        As Long
    Dim Chk1_Val As String
    Dim sCur_ItemCode As String
    Dim sNxt_ItemCode As String
    Dim sCur_TrCardCode As String
    Dim sCur_TrOutWhs As String
    Dim sNxt_TrOutWhs As String
    Dim sCur_TrInWhs As String
    Dim sNxt_TrInWhs As String
    Dim RtnDocNum     As String
    Dim oStockTrans   As SAPbobsCOM.StockTransfer
    Dim oPrgBar       As SAPbouiCOM.ProgressBar
    Dim StockTransLineCounter As Long
    
    Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BatchNum As String
    
    PS_MM130_StockTrans = True
    
    Dim Query01             As String
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
    
    Dim Query02             As String
    Dim oRecordset02        As SAPbobsCOM.Recordset
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
    
    For i = 0 To oMat01.RowCount - 1
        ReDim Preserve StockInfo(lTypeCount)
        
        If oMat01.Columns("OutGbn").Cells(i + 1).Specific.VALUE = "10" Then
            
            oMat01.FlushToDataSource
        
            StockInfo(lTypeCount).ItemCode = Trim(oDS_PS_MM130L.GetValue("U_OutItmCd", i))
            StockInfo(lTypeCount).FromWarehouseCode = Trim(oDS_PS_MM130L.GetValue("U_OutWhCd", i))
            StockInfo(lTypeCount).ToWarehouseCode = Trim(oDS_PS_MM130L.GetValue("U_InWhCd", i))
    '        StockInfo(lTypeCount).BatchNum = Trim(oDS_PS_MM130L.GetValue("U_BatchNum", i))
           
            StockInfo(lTypeCount).Weight = Round(Trim(oDS_PS_MM130L.GetValue("U_OutWt", i)), 2)             'Quantity
            StockInfo(lTypeCount).UnWeight = Round(Trim(oDS_PS_MM130L.GetValue("U_UnWeight", i)), 2)
            StockInfo(lTypeCount).BatchWeight = Round(Trim(oDS_PS_MM130L.GetValue("U_OutQty", i)), 2)
            StockInfo(lTypeCount).Qty = Val(Trim(oDS_PS_MM130L.GetValue("U_OutQty", i)))                    'U_Qty
       
            StockInfo(lTypeCount).TransNo = oForm01.Items("DocEntry").Specific.VALUE & (i + 1)
            StockInfo(lTypeCount).Chk = "N"
            StockInfo(lTypeCount).MatrixRow = (i + 1)
            StockInfo(lTypeCount).Indate = oForm01.Items("DocDate").Specific.VALUE
            lTypeCount = lTypeCount + 1
        End If
    Next
    
    For i = 0 To (UBound(StockInfo))
        StockInfo(i).StockTransDocEntry = ""
    Next
    
    Sbo_Company.StartTransaction
    For i = 0 To (UBound(StockInfo))
    
        Chk1_Val = StockInfo(i).Chk
        
        If Chk1_Val <> "N" Then GoTo Continue_First
        
        sCur_TrOutWhs = StockInfo(i).FromWarehouseCode
        
        Set oStockTrans = Sbo_Company.GetBusinessObject(oStockTransfer)
'        oStockTrans.CardCode = StockInfo(i).CardCode
        oStockTrans.DocDate = Format(StockInfo(i).Indate, "&&&&-&&-&&")
        oStockTrans.FromWarehouse = sCur_TrOutWhs
        oStockTrans.Comments = "재고이전" & oForm01.Items("DocEntry").Specific.VALUE & "."
        
        StockTransLineCounter = -1
        For K = i To (UBound(StockInfo))
            Chk1_Val = StockInfo(K).Chk
            
            If Chk1_Val <> "N" Then GoTo Continue_Second
'            sCur_TrCardCode = StockInfo(K).CardCode
            sNxt_TrOutWhs = StockInfo(K).FromWarehouseCode
            sCur_ItemCode = StockInfo(K).ItemCode
            sCur_TrInWhs = StockInfo(K).ToWarehouseCode
            
            If (sCur_TrOutWhs <> sNxt_TrOutWhs) Then
                GoTo Continue_Second
            End If

            If (i <> K) Then
                oStockTrans.Lines.Add
            End If
            StockTransLineCounter = StockTransLineCounter + 1
            '---------------------------------------------------------------------------< Line >----------
            With oStockTrans.Lines
                
                .ItemCode = StockInfo(K).ItemCode
                .UserFields("U_Qty").VALUE = Trim(StockInfo(K).Qty)                 '// 수량
                .UserFields("U_UnWeight").VALUE = Trim(StockInfo(K).UnWeight)       '// 단중
                .Quantity = Round(StockInfo(K).Weight, 2)                              '// 수량
                .WarehouseCode = StockInfo(K).ToWarehouseCode
                '//ManBatchNum = 'Y' 이면 배치번호를 입력하지 않는다.
'                .UserFields("U_BatchNum").Value = StockInfo(K).BatchNum
'                .BatchNumbers.BatchNumber = StockInfo(K).BatchNum
'                .BatchNumbers.Quantity = Round(StockInfo(K).BatchWeight, 2)
'                .BatchNumbers.Notes = "재고이전(Addon)"
                StockInfo(K).Chk = "Y"   '/ 적용한 라인에 대한 표시
                StockInfo(K).StockTransDocEntry = "Checked"
                StockInfo(K).StockTransLineNum = StockTransLineCounter
                
                For Q = K + 1 To (UBound(StockInfo))
                    Chk1_Val = StockInfo(Q).Chk
                    
                    If Chk1_Val <> "N" Then GoTo Continue_Sixth   '/ 체크2 에 않된건 Skip
                    
                    sNxt_TrOutWhs = StockInfo(Q).FromWarehouseCode
                    sNxt_ItemCode = StockInfo(Q).ItemCode
                    sNxt_TrInWhs = StockInfo(Q).ToWarehouseCode
Continue_Sixth:
                Next
            End With
Continue_Second:
        Next
        '---------------------------------------------------------------------------------------------

        RetVal = oStockTrans.Add
        If RetVal = 0 Then
            DocCnt = DocCnt + 1
            Call Sbo_Company.GetNewObjectCode(RtnDocNum) '//재고이전문서번호
            For r = 0 To UBound(StockInfo)
                If (StockInfo(r).StockTransDocEntry = "Checked") Then
                    StockInfo(r).StockTransDocEntry = RtnDocNum
                End If
            Next
            '// 데이터 업데이트
         Else
            GoTo PS_MM130_StockTrans_Error
        End If
Continue_First:
    Next   '-----------------------------------------------------------------------------------------------< First For End
    
    If (Sbo_Company.InTransaction) Then
        Sbo_Company.EndTransaction (wf_Commit)
    End If
    Call Sbo_Application.SetStatusBarMessage(DocCnt & " 개의 재고이전 문서가 발행되었습니다 !", bmt_Short, False)
    Set oStockTrans = Nothing
    Set lRecordSet = Nothing
    Set oRecordSet = Nothing
    Exit Function
'************Error Process************
PS_MM130_StockTrans_Error:
    If (Sbo_Company.InTransaction) Then
        Sbo_Company.EndTransaction (wf_RollBack)
    End If
    Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
    Call Sbo_Application.SetStatusBarMessage(ErrCode & " : " & ErrMsg, bmt_Short, True)
    PS_MM130_StockTrans = False
    Set oStockTrans = Nothing
    Set lRecordSet = Nothing
    Set oRecordSet = Nothing
'************Error Process************

End Function


Private Function PS_MM130_UpdateUserField() As Boolean
On Error GoTo PS_MM130_UpdateUserField_Error
    Dim i As Long
    Dim lQuery As String
    Dim lRecordSet As SAPbobsCOM.Recordset
    Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oDS_PS_MM130H.setValue("U_STDocNum", 0, (StockInfo(i).StockTransDocEntry))
'    oForm01.Items("StoTrDoc").Specific.Value = StockInfo(i).StockTransDocEntry

    PS_MM130_UpdateUserField = True
    Exit Function
PS_MM130_UpdateUserField_Error:
    PS_MM130_UpdateUserField = False
End Function

Private Function PS_MM130_Update_Cancel() As Boolean
On Error GoTo PS_MM130_Update_Cancel_Error
    Dim i As Long
    Dim lQuery As String
    Dim lRecordSet As SAPbobsCOM.Recordset
    Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oDS_PS_MM130H.setValue("U_STDocNum", 0, "")
'    oForm01.Items("StoTrDoc").Specific.Value = StockInfo(i).StockTransDocEntry

    PS_MM130_Update_Cancel = True
    Exit Function
PS_MM130_Update_Cancel_Error:
    PS_MM130_Update_Cancel = False
End Function

Private Function PS_MM130_Cancel_oStockTrans(ChkType As Integer) As Boolean
On Error GoTo PS_MM130_Cancel_oStockTrans_Error
    Dim oStockTrans  As SAPbobsCOM.StockTransfer
        
    Dim i&
    Dim ErrNum&, ErrCode As Long, ErrMsg$, RetVal&
    
    Dim DocEntry$

    DocEntry = Trim(oDS_PS_MM130H.GetValue("U_STDocNum", 0))
    
    If Trim(oDS_PS_MM130H.GetValue("U_STDocNum", 0)) <> "" Then
        Sbo_Company.StartTransaction
        Set oStockTrans = Nothing
        Set oStockTrans = Sbo_Company.GetBusinessObject(oStockTransfer)
                
        '//완료
        If (oStockTrans.GetByKey(DocEntry) = False) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo PS_MM130_Cancel_oStockTrans_Error
        End If
        RetVal = oStockTrans.Cancel
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            ErrNum = 1
            GoTo PS_MM130_Cancel_oStockTrans_Error
        End If
        
        If ChkType = 1 Then
            Sbo_Company.EndTransaction wf_RollBack
        ElseIf ChkType = 2 Then
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If
    
    Set oStockTrans = Nothing
    PS_MM130_Cancel_oStockTrans = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM130_Cancel_oStockTrans_Error:
    Set oStockTrans = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    PS_MM130_Cancel_oStockTrans = False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "PS_MM130_Cancel_oStockTrans_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "PS_MM130_Cancel_oStockTrans_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function


