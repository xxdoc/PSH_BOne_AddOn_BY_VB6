VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM132"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM132.cls
'//  Module         : 재고관리 > 외주반품등록
'//  Desc           : MM
'//  FormType       : PS_MM132
'//  Create Date    : 2010.12.09
'//  Creator        : Dong sub Kim
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_MM132H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_MM132L As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM132.srf")
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
    
    oFormUniqueID01 = "PS_MM132_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01                 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry"    '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_MM132_CreateItems
    Call PS_MM132_ComboBox_Setting
    Call PS_MM132_Initial_Setting
    Call PS_MM132_CF_ChooseFromList
    Call PS_MM132_EnableMenus
    Call PS_MM132_SetDocument(oFromDocEntry01)
    Call PS_MM132_FormResize
    
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
                If (PS_MM132_Validate("취소") = False) Then
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
                Call PS_MM132_FormItemEnabled '//UDO방식
            Case "1282": '추가
                oForm01.Freeze True
                
                Call oDS_PS_MM132H.setValue("U_DocDate", 0, Format(Now, "YYYYMMDD"))
                Call oDS_PS_MM132H.setValue("U_BPLId", 0, "1")
                Call oDS_PS_MM132H.setValue("U_OKYNC", 0, "N")
                Call PS_MM132_FormItemEnabled '//UDO방식
                Call PS_MM132_AddMatrixRow(0, True) '//UDO방식
                oForm01.Freeze False
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_MM132_FormItemEnabled
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

    Dim Query01             As String
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim OutDoc As String
    Dim OutLine As String
    Dim OKYNC As String
    Dim ReQty As Currency, ReWt As Currency
    
    If pval.BeforeAction = True Then
        Dim i As Integer
        Dim j As Integer
    
        If pval.ItemUID = "PS_MM132" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
            '//외주 재고 이동 DI API
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_MM132_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                 
            
                 
            '//외주 재고 이동 취소 DI API
            ElseIf (oForm01.Mode = fm_UPDATE_MODE) And Trim(oForm01.Items("OKYNC").Specific.VALUE) = "Y" And oForm01.Items("STDocNo").Specific.VALUE = "" Then
                If PS_MM132_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                For i = 0 To oMat01.RowCount - 1
                     If oMat01.Columns("OutGbn").Cells(i + 1).Specific.VALUE = "10" Then
                        If j = 0 Then
                            j = j + 1
                            If PS_MM132_StockTrans() = True Then
                                 Call PS_MM132_UpdateUserField
                            Else
                                Call PS_MM132_AddMatrixRow(oMat01.VisualRowCount, False)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    ElseIf oMat01.Columns("OutGbn").Cells(i + 1).Specific.VALUE = "20" Then
                            
                    End If
                     
                 Next
                 
            ElseIf (oForm01.Mode = fm_UPDATE_MODE) And Trim(oForm01.Items("OKYNC").Specific.VALUE) = "C" Then
                   
                If PS_MM132_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                For i = 0 To oMat01.RowCount - 1
                     If oMat01.Columns("OutGbn").Cells(i + 1).Specific.VALUE = "10" Then
                        If j = 0 Then
                            j = j + 1
                            If PS_MM132_Cancel_oStockTrans(2) = True Then
                            Else
                                Call PS_MM132_AddMatrixRow(0, True)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                     End If
                 Next
            ElseIf oForm01.Mode = fm_UPDATE_MODE And (Trim(oForm01.Items("OKYNC").Specific.VALUE) = "N" And oForm01.Items("STDocNo").Specific.VALUE = "") Then   '//창원 문서 수정
                If PS_MM132_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE And (Trim(oForm01.Items("OKYNC").Specific.VALUE) = "N" And oForm01.Items("STDocNo").Specific.VALUE <> "") Then   '//창원
'                   GoTo Raise_EVENT_ITEM_PRESSED_Error
                   MDC_Com.MDC_GF_Message "이미 승인되었습니다." & Err.Number & " - " & Err.Description, "E"
                  BubbleEvent = False
                  Exit Sub
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
            '/* 반품승인 또는 반품시 기준문서에 반품수량 Update */
            If (oForm01.Mode = fm_UPDATE_MODE) Then
                OKYNC = Trim(oForm01.Items("OKYNC").Specific.VALUE)
                If OKYNC = "Y" Or OKYNC = "C" Then
                    For i = 0 To oMat01.RowCount - 1
                        OutDoc = oMat01.Columns("OutDoc").Cells(i + 1).Specific.VALUE
                        OutLine = oMat01.Columns("OutLine").Cells(i + 1).Specific.VALUE
                        ReQty = oMat01.Columns("ReQty").Cells(i + 1).Specific.VALUE
                        ReWt = oMat01.Columns("ReWt").Cells(i + 1).Specific.VALUE
                        
                        If OKYNC = "C" Then
                           ReQty = ReQty * -1
                           ReWt = ReWt * -1
                        End If
                        
                        Query01 = "Update [@PS_MM130L] set U_ReQty = Isnull(U_ReQty,0) + " & ReQty & ", U_ReWt = Isnull(U_ReWt,0) + " & ReWt & " Where U_OutDoc = '" & OutDoc & "' and LineId = '" & OutLine & "'"
                        Call oRecordSet01.DoQuery(Query01)
                    Next
                End If
            End If
        End If
   
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_MM132" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    
                    
                    
                    Call PS_MM132_FormItemEnabled
                    Call PS_MM132_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
                    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
                    Call oForm01.Items("BPLId").Specific.Select("0", psk_Index)
                    Call oForm01.Items("OKYNC").Specific.Select("Y", psk_ByValue)
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_MM132_FormItemEnabled
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "")
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "OtDocLin" Then
                If PS_MM132_OutDocCheck = True Then
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "OtDocLin")
                End If
            End If
        End If
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "OutWhCd")
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "InWhCd")
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
'                    Call oDS_PS_MM132L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM132L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_MM132_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_MM132L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                End If
'            Else
'                If (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_MM132H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_MM132H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
'                    Call oDS_PS_MM132H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.Value & "'"))
'                Else
'                    Call oDS_PS_MM132H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
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
        If pval.ItemUID = "Rad01" Or pval.ItemUID = "Rad02" Then
            oDS_PS_MM132L.Clear
            oMat01.LoadFromDataSource
            oMat01.FlushToDataSource
            Call PS_MM132_AddMatrixRow(0, True) '//UDO방식
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
                If (pval.ColUID = "OtDocLin") Then
                    '//기타작업
                    Dim i As Double
                    Dim j As Double
                    Dim DocEntry As Double
                    Dim LineId As Double
                    j = InStr(Trim(oMat01.Columns("OtDocLin").Cells(pval.Row).Specific.VALUE), "-")
                    DocEntry = Left(Trim(oMat01.Columns("OtDocLin").Cells(pval.Row).Specific.VALUE), j - 1)
                    LineId = Mid(Trim(oMat01.Columns("OtDocLin").Cells(pval.Row).Specific.VALUE), j + 1)
                    Query01 = "EXEC PS_MM132_01 '" & DocEntry & "', '" & LineId & "'"

                    Call oRecordSet01.DoQuery(Query01)
                    oMat01.FlushToDataSource
                    
                    Call oDS_PS_MM132L.setValue("U_ItemCode", pval.Row - 1, Trim(oRecordSet01.Fields(0).VALUE))
                    Call oDS_PS_MM132L.setValue("U_ItemName", pval.Row - 1, Trim(oRecordSet01.Fields(1).VALUE))
                    Call oDS_PS_MM132L.setValue("U_Size", pval.Row - 1, Trim(oRecordSet01.Fields(2).VALUE))
                    Call oDS_PS_MM132L.setValue("U_Mark", pval.Row - 1, Trim(oRecordSet01.Fields(3).VALUE))
                    Call oDS_PS_MM132L.setValue("U_UnWeight", pval.Row - 1, Trim(oRecordSet01.Fields(4).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OutItmCd", pval.Row - 1, Trim(oRecordSet01.Fields(5).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OutItmNm", pval.Row - 1, Trim(oRecordSet01.Fields(6).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OutGbn", pval.Row - 1, Trim(oRecordSet01.Fields(7).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OutQty", pval.Row - 1, Trim(oRecordSet01.Fields(8).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OutWt", pval.Row - 1, Trim(oRecordSet01.Fields(9).VALUE))
                    Call oDS_PS_MM132L.setValue("U_ReQty", pval.Row - 1, Trim(oRecordSet01.Fields(10).VALUE))
                    Call oDS_PS_MM132L.setValue("U_ReWt", pval.Row - 1, Trim(oRecordSet01.Fields(11).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OutWhCd", pval.Row - 1, Trim(oRecordSet01.Fields(12).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OutWhNm", pval.Row - 1, Trim(oRecordSet01.Fields(13).VALUE))
                    Call oDS_PS_MM132L.setValue("U_InWhCd", pval.Row - 1, Trim(oRecordSet01.Fields(14).VALUE))
                    Call oDS_PS_MM132L.setValue("U_InWhNm", pval.Row - 1, Trim(oRecordSet01.Fields(15).VALUE))
                    Call oDS_PS_MM132L.setValue("U_PP030HNo", pval.Row - 1, Trim(oRecordSet01.Fields(16).VALUE))
                    Call oDS_PS_MM132L.setValue("U_PP030MNo", pval.Row - 1, Trim(oRecordSet01.Fields(17).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OutDoc", pval.Row - 1, Trim(oRecordSet01.Fields(18).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OutLine", pval.Row - 1, Trim(oRecordSet01.Fields(19).VALUE))
                    Call oDS_PS_MM132L.setValue("U_OrdNum", pval.Row - 1, Trim(oRecordSet01.Fields(20).VALUE))
                    Set oRecordSet01 = Nothing

                    Call oDS_PS_MM132L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM132L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_MM132_AddMatrixRow (pval.Row)
                    End If
                    
                ElseIf (pval.ColUID = "OutWhCd") Then
                    Call oDS_PS_MM132L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_MM132L.setValue("U_OUtWhNm", pval.Row - 1, MDC_GetData.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"))
                    oMat01.FlushToDataSource
                ElseIf (pval.ColUID = "InWhCd") Then
                    Call oDS_PS_MM132L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_MM132L.setValue("U_InWhNm", pval.Row - 1, MDC_GetData.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"))
                    oMat01.FlushToDataSource
                ElseIf (pval.ColUID = "ReQty") Then
'                    Dim Count As Double
'                    If oMat01.Columns("UnWeight").Cells(pval.Row).Specific.Value > 0 Then
'                        Count = oMat01.Columns("UnWeight").Cells(pval.Row).Specific.Value * oMat01.Columns("ReQty").Cells(pval.Row).Specific.Value
'                    Else
'                        Count = 1 * oMat01.Columns("ReQty").Cells(pval.Row).Specific.Value
'                    End If

                     oMat01.FlushToDataSource
                     If oMat01.Columns("ReQty").Cells(pval.Row).Specific.VALUE > 0 Then
                        Query01 = "Select U_ObasUnit FROM OITM WHERE ItemCode = '" & Trim(oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE) & "'"
                        Call oRecordSet01.DoQuery(Query01)
                        Query02 = "Select OnHand, U_Qty FROM OITW WHERE ItemCode = '" & Trim(oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE) & "' AND WhsCode = '801'"
                        Call oRecordset02.DoQuery(Query02)
                        If Left(oRecordSet01.Fields(0).VALUE, 1) = "1" Then
                            If (Trim(oRecordset02.Fields(0).VALUE)) > 0 And oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE <> "" Then
                                oMat01.Columns("ReWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("ReQty").Cells(pval.Row).Specific.VALUE
                            Else
                                oMat01.Columns("ReWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("ReQty").Cells(pval.Row).Specific.VALUE
                            End If
                        ElseIf Left(oRecordSet01.Fields(0).VALUE, 1) = "2" Then
                            If (Trim(oRecordset02.Fields(1).VALUE)) > 0 And oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE <> "" Then
                                oMat01.Columns("ReWt").Cells(pval.Row).Specific.VALUE = (Trim(oRecordset02.Fields(0).VALUE) / Trim(oRecordset02.Fields(1).VALUE)) * oMat01.Columns("ReQty").Cells(pval.Row).Specific.VALUE
                            Else
                                oMat01.Columns("ReWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("ReWt").Cells(pval.Row).Specific.VALUE
                            End If
                        Else
                        End If
                        oMat01.LoadFromDataSource
                        Call oDS_PS_MM132L.setValue("U_ReQty", pval.Row - 1, oMat01.Columns("ReQty").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_MM132L.setValue("U_ReWt", pval.Row - 1, oMat01.Columns("ReWt").Cells(pval.Row).Specific.VALUE)
                        oMat01.Columns("ReWt").Cells(pval.Row).Click ct_Regular
                    End If
                    
                Else
                    Call oDS_PS_MM132L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    oMat01.LoadFromDataSource
                    oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                End If
            
            ElseIf (pval.ItemUID = "CardCode") Then
                Query01 = "SELECT CardName FROM [OCRD] WHERE CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                Call oRecordSet01.DoQuery(Query01)
                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                Set oRecordSet01 = Nothing
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_MM132H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
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
        Call PS_MM132_FormItemEnabled
        Call PS_MM132_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_MM132_FormResize
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
                Call oDS_PS_MM132H.setValue("U_CardCode", 0, oDataTable01.Columns("CardCode").Cells(0).VALUE)
                Call oDS_PS_MM132H.setValue("U_CardName", 0, oDataTable01.Columns("CardName").Cells(0).VALUE)
                ' // 찾기나 문서이동 버튼 클릭 시에 갱신으로 바뀌지 않음
                If oForm01.Mode <> fm_ADD_MODE Then oForm01.Mode = fm_UPDATE_MODE
'               oForm01.DataSources.UserDataSources("CardCode").Value = oDataTable01.Columns("CardCode").Cells(0).Value
'               oForm01.DataSources.UserDataSources("CardName").Value = oDataTable01.Columns("CardName").Cells(0).Value
            End If
        ElseIf (pval.ItemUID = "ShipTo") Then
            Set oDataTable02 = pval.SelectedObjects
            If oDataTable02 Is Nothing Then
            Else
  
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
'            If (PS_MM132_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_MM132L.RemoveRecord(oDS_PS_MM132L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_MM132_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_MM132L.GetValue("U_OtDocLin", oMat01.RowCount - 1)) <> "" Then
                    Call PS_MM132_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_MM132_CreateItems() As Boolean
On Error GoTo PS_MM132_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM132H = oForm01.DataSources.DBDataSources("@PS_MM132H")
    Set oDS_PS_MM132L = oForm01.DataSources.DBDataSources("@PS_MM132L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    ' // 매트릭스 비활성 상태에서 값이 안 나올 경우
    ' 매트릭스 선택 포맷을 반환하거나 설정
    ' 매트릭스 선택 규칙을 자체적으로 선택할 수 없는 경우 반환되는 값
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns ' 매트릭스의 칼럼 크기를 자동으로 정렬
    
    oForm01.Settings.MatrixUID = "Mat01"             ' 서식세팅
    oForm01.Settings.Enabled = True
    oForm01.Settings.EnableRowFormat = True
    
    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Or Trim(Sbo_Company.UserName) = "66510" Then
'        Trim(oForm01.Items("CardCode").Specific.Value) = Trim(Sbo_Company.UserName)
    End If
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_MM132_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM132_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_MM132_ComboBox_Setting()
On Error GoTo PS_MM132_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//콤보에 기본값설정
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM132", "Mat01", "OutGbn", "10", "원재료")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM132", "Mat01", "OutGbn", "20", "제공")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("OutGbn"), "PS_MM132", "Mat01", "OutGbn")
    
    Set oCombo = oForm01.Items("OKYNC").Specific
  
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM132", "OKYNC", "", "N", "반품")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM132", "OKYNC", "", "Y", "승인")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM132", "OKYNC", "", "C", "반품취소")
      
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("OKYNC").Specific, "PS_MM132", "OKYNC", False)
    Call oCombo.Select(0, psk_Index)
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL WHERE BPLId = '1'  ORDER BY BPLId", "1", False, False)
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM132_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM132_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM132_CF_ChooseFromList()
On Error GoTo PS_MM132_CF_ChooseFromList_Error
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
    
'    Set oEdit02 = oForm01.Items("ShipTo").Specific
'    Set oCFLs02 = oForm01.ChooseFromLists
'    Set oCFLCreationParams02 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams02.ObjectType = lf_BusinessPartner
'    oCFLCreationParams02.uniqueID = "CFLSHIPCODE"
'    oCFLCreationParams02.MultiSelection = False
'    Set oCFL02 = oCFLs02.Add(oCFLCreationParams02)
'
'    oEdit02.ChooseFromListUID = "CFLSHIPCODE"
'    oEdit02.ChooseFromListAlias = "CardCode"                ' 정의된 오브젝트만 사용 가능
    
    Exit Sub
PS_MM132_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM132_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
    
Sub PS_MM132_Initial_Setting()
On Error GoTo PS_MM132_Initial_Setting_Error
 ' 사업장
'    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 ' 인수자
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD()
  If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Then
    oForm01.Items("CardCode").Specific.VALUE = Trim(Sbo_Company.UserName)
  End If
    Exit Sub
PS_MM132_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM132_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM132_FormItemEnabled()
On Error GoTo PS_MM132_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        Call PS_MM132_FormClear                '//UDO방식
        Call oForm01.EnableMenu("1281", True)  '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        Call oForm01.EnableMenu("1293", True)
        If Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "Y" Then
            Call oForm01.EnableMenu("1284", False)       '//취소
            Call oForm01.EnableMenu("1286", False)       '//닫기
            Call oForm01.EnableMenu("1293", False)       '//행삭제
        Else
            Call oForm01.EnableMenu("1284", True)       '//취소
            Call oForm01.EnableMenu("1286", True)       '//닫기
            Call oForm01.EnableMenu("1293", True)       '//행삭제
        End If
        
            oMat01.Columns("OutWhCd").Editable = True
            oMat01.Columns("OutWhNm").Editable = True
            oMat01.Columns("InWhCd").Editable = True
            oMat01.Columns("InWhNm").Editable = True
            
        oForm01.Items("Comments").Click ct_Regular
                '외주업체
        If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Or Trim(Sbo_Company.UserName) = "66510" Then
            If Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "N" Then
                oForm01.Items("CardCode").Enabled = True
                oForm01.Items("CardCode").Specific.VALUE = Trim(Sbo_Company.UserName)
                oForm01.Items("Comments").Click ct_Regular
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("OKYNC").Enabled = True
                oMat01.Columns("OtDocLin").Editable = True
                oMat01.Columns("ReQty").Editable = True
                oMat01.Columns("ReWt").Editable = True
            End If
        '신동
        
        Else
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("OKYNC").Enabled = True
            oMat01.Columns("OtDocLin").Editable = True
            oMat01.Columns("ReQty").Editable = True
            oMat01.Columns("ReWt").Editable = True

        End If
    
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True)  '//추가
        Call oForm01.EnableMenu("1293", True)
        
        If Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "Y" Then
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
        Else
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("DocEntry").Enabled = True
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("OKYNC").Enabled = True
        End If
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True)  '//찾기
        Call oForm01.EnableMenu("1282", True)  '//추가
        oForm01.EnableMenu ("1293"), True
                
        If Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "Y" Or Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "C" Then
            Call oForm01.EnableMenu("1284", False)      '//취소
            Call oForm01.EnableMenu("1286", False)      '//닫기
            Call oForm01.EnableMenu("1293", False)      '//행삭제
        Else
            Call oForm01.EnableMenu("1284", True)       '//취소
            Call oForm01.EnableMenu("1286", True)       '//닫기
            Call oForm01.EnableMenu("1293", True)       '//행삭제
        End If
        
        oMat01.Columns("OtDocLin").Editable = False
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Comments").Click ct_Regular
        
        '외주업체
        If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Or Trim(Sbo_Company.UserName) = "66510" Then
            If Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "C" Or Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "Y" Then
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("OKYNC").Enabled = False
                oMat01.Columns("OtDocLin").Editable = False
                oMat01.Columns("ReQty").Editable = False
                oMat01.Columns("ReWt").Editable = False
                oMat01.Columns("OutWhCd").Editable = False
                oMat01.Columns("OutWhNm").Editable = False
                oMat01.Columns("InWhCd").Editable = False
                oMat01.Columns("InWhNm").Editable = False
                
            Else
                oForm01.Items("CardCode").Enabled = True
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("OKYNC").Enabled = True
                oMat01.Columns("OtDocLin").Editable = True
                oMat01.Columns("ReQty").Editable = True
                oMat01.Columns("ReWt").Editable = True
            End If
        '신동
        Else
            If Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "N" Or Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "Y" Then
                oForm01.Items("OKYNC").Enabled = True
                oMat01.Columns("OutWhCd").Editable = False
                oMat01.Columns("OutWhNm").Editable = False
                oMat01.Columns("InWhCd").Editable = False
                oMat01.Columns("InWhNm").Editable = False
            ElseIf Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "C" Then
                oForm01.Items("OKYNC").Enabled = False
                
            End If
            oForm01.Items("CardCode").Enabled = False
        End If
        If Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "N" Then
            oMat01.Columns("OutWhCd").Editable = True
            oMat01.Columns("OutWhNm").Editable = True
            oMat01.Columns("InWhCd").Editable = True
            oMat01.Columns("InWhNm").Editable = True
        ElseIf Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "Y" Then
            oMat01.Columns("OutWhCd").Editable = False
            oMat01.Columns("OutWhNm").Editable = False
            oMat01.Columns("InWhCd").Editable = False
            oMat01.Columns("InWhNm").Editable = False
        ElseIf Trim(oDS_PS_MM132H.GetValue("U_OKYNC", 0)) = "C" Then
            oMat01.Columns("OutWhCd").Editable = False
            oMat01.Columns("OutWhNm").Editable = False
            oMat01.Columns("InWhCd").Editable = False
            oMat01.Columns("InWhNm").Editable = False
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM132_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM132_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM132_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_MM132_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM132L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM132L.Offset = oRow
    oDS_PS_MM132L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM132_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM132_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM132_FormClear()
On Error GoTo PS_MM132_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM132'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_MM132_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM132_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM132_EnableMenus()
On Error GoTo PS_MM132_EnableMenus_Error
'    //메뉴활성화
    Call oForm01.EnableMenu("1283", False)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
'    Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
'    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_MM132_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM132_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM132_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_MM132_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_MM132_FormItemEnabled
        Call PS_MM132_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_MM132_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_MM132_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM132_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_MM132_DataValidCheck() As Boolean
On Error GoTo PS_MM132_DataValidCheck_Error
    Dim i As Long
    
    oMat01.FlushToDataSource
    PS_MM132_DataValidCheck = False
    
    If oMat01.VisualRowCount <= 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_MM132_DataValidCheck = False
        Exit Function
    End If
    
    If oForm01.Items("BPLId").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "사업장 코드는 필수입니다.", bmt_Short, True
        oForm01.Items("BPLId").Click ct_Regular
        PS_MM132_DataValidCheck = False
        Exit Function
    
    ElseIf oForm01.Items("CardCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "외주거래처 코드는 필수입니다.", bmt_Short, True
        oForm01.Items("CardCode").Click ct_Regular
        PS_MM132_DataValidCheck = False
        Exit Function
        
    ElseIf oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "전기일자는 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_MM132_DataValidCheck = False
        Exit Function
    End If
    
    
    If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "71090" Or Trim(Sbo_Company.UserName) = "66510" Then
        If Trim(oForm01.Items("OKYNC").Specific.VALUE) <> "N" Then
            Sbo_Application.SetStatusBarMessage "이동승인상태 N- 반품만 선택 할 수 있습니다.", bmt_Short, True
            oForm01.Items("OKYNC").Click ct_Regular
            PS_MM132_DataValidCheck = False
            Exit Function
        End If
    End If
    
    
    
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("OtDocLin").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "반출문서는 필수입니다.", bmt_Short, True
            oMat01.Columns("OtDocLin").Cells(i).Click ct_Regular
            PS_MM132_DataValidCheck = False
            Exit Function
            
        ElseIf (oMat01.Columns("ReQty").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "반품수량은 필수입니다.", bmt_Short, True
            oMat01.Columns("ReQty").Cells(i).Click ct_Regular
            PS_MM132_DataValidCheck = False
            Exit Function
            
        ElseIf (oMat01.Columns("ReWt").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "반품중량은 필수입니다.", bmt_Short, True
            oMat01.Columns("ReWt").Cells(i).Click ct_Regular
            PS_MM132_DataValidCheck = False
            Exit Function
        End If
    Next
    
    Call oDS_PS_MM132L.RemoveRecord(oDS_PS_MM132L.Size - 1)
    Call oMat01.LoadFromDataSource
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_MM132_FormClear
    End If
    PS_MM132_DataValidCheck = True
    Exit Function
PS_MM132_DataValidCheck_Error:
    PS_MM132_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_MM132_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
Function PS_MM132_OutDocCheck() As Boolean
On Error GoTo PS_MM132_OutDocCheck_Error
    If oForm01.Items("BPLId").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "사업장 코드를 먼저 입력하세요.", bmt_Short, True
        oForm01.Items("BPLId").Click ct_Regular
        PS_MM132_OutDocCheck = False
        Exit Function
        
    ElseIf oForm01.Items("CardCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "외주거래처 코드를 먼저 입력하세요.", bmt_Short, True
        oForm01.Items("CardCode").Click ct_Regular
        PS_MM132_OutDocCheck = False
        Exit Function
    End If
    PS_MM132_OutDocCheck = True
    Exit Function
PS_MM132_OutDocCheck_Error:
    PS_MM132_OutDocCheck = False
    Sbo_Application.SetStatusBarMessage "PS_MM132_OutDocCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_MM132_FormResize()
On Error GoTo PS_MM132_FormResize_Error
    
    Exit Sub
PS_MM132_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM132_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_MM132_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_MM132_Validate_Error
    PS_MM132_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "수정" Then
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
'                    PS_MM132_Validate = False
'                    GoTo PS_MM132_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
'        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
'            Else
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(oForm01.Items("8").Specific.Value) & "' AND U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할수 없습니다.", "W"
'                    PS_MM132_Validate = False
'                    GoTo PS_MM132_Validate_Exit
'                End If
'            End If
'        End If
    ElseIf ValidateType = "취소" Then
'//취소가능유무검사
    '/ 반출등록 취소
        If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_MM132H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
            MDC_Com.MDC_GF_Message "이미취소된 문서 입니다. 취소할수 없습니다.", "W"
'            PS_PP030_Validate = False
'            GoTo PS_PP030_Validate_Exit
        End If

    '/ 입고등록 취소에 따른 반출 취소 불가
'        If oForm01.Items("OutDoc").Specific.Selected.Value <> "" Then           '//반출문서
'            If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM140H] WHERE Canceled = 'Y' AND U_OutDoc = '" & oForm01.Items("OutDoc").Specific.Value & "'") Then
'                MDC_Com.MDC_GF_Message "입고 문서가 취소 되어 있습니다.", "W"
'                PS_MM132_Validate = False
'                GoTo PS_PP030_Validate_Exit
'            End If
'        End If
    End If

    Set RecordSet01 = Nothing
    Exit Function
PS_MM132_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_MM132_Validate_Error:
    PS_MM132_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_MM132_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PS_MM132_StockTrans() As Boolean
On Error GoTo PS_MM132_StockTrans_Error
    Dim RetVal As Long
    Dim ErrCode As Long
    Dim ErrMsg As String
    Dim oQuery01 As String
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
    Dim ReQty As Currency, ReWt As Currency
    
    Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BatchNum As String
    
        PS_MM132_StockTrans = True
    
    For i = 0 To oMat01.RowCount - 1
         If oMat01.Columns("OutGbn").Cells(i + 1).Specific.VALUE = "10" Then
        ReDim Preserve StockInfo(lTypeCount)
        'DI API
'        StockInfo(lTypeCount).CardCode = Trim(oDS_PS_MM132H.GetValue("U_CardCode", 0))
        StockInfo(lTypeCount).ItemCode = Trim(oDS_PS_MM132L.GetValue("U_OutItmCd", i))
        StockInfo(lTypeCount).FromWarehouseCode = Trim(oDS_PS_MM132L.GetValue("U_OutWhCd", i))
        StockInfo(lTypeCount).ToWarehouseCode = Trim(oDS_PS_MM132L.GetValue("U_InWhCd", i))
'        StockInfo(lTypeCount).BatchNum = Trim(oDS_PS_MM132L.GetValue("U_BatchNum", i))
       
        StockInfo(lTypeCount).Weight = Round(Trim(oDS_PS_MM132L.GetValue("U_ReWt", i)), 2)
        StockInfo(lTypeCount).UnWeight = Round(Trim(oDS_PS_MM132L.GetValue("U_UnWeight", i)), 2)
        StockInfo(lTypeCount).BatchWeight = Round(Trim(oDS_PS_MM132L.GetValue("U_ReQty", i)), 2)
        StockInfo(lTypeCount).Qty = Val(Trim(oDS_PS_MM132L.GetValue("U_ReQty", i)))
   
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
                .Quantity = Round(StockInfo(K).Weight, 2)                           '// 중량
                .WarehouseCode = StockInfo(K).ToWarehouseCode
                '//ManBatchNum = 'Y' 이면 배치번호를 입력하지 않는다.
'                .UserFields("U_BatchNum").Value = StockInfo(K).BatchNum
'                .BatchNumbers.BatchNumber = StockInfo(K).BatchNum
'                .BatchNumbers.Quantity = Round(StockInfo(K).BatchWeight, 2)

                .BatchNumbers.Notes = "재고이전(Addon)"
                StockInfo(K).Chk = "Y"   '/ 적용한 라인에 대한 표시
                StockInfo(K).StockTransDocEntry = "Checked"
                StockInfo(K).StockTransLineNum = StockTransLineCounter
                
                For Q = K + 1 To (UBound(StockInfo))
                    Chk1_Val = StockInfo(Q).Chk
                    
                    If Chk1_Val <> "N" Then GoTo Continue_Sixth   '/ 체크2 에 않된건 Skip
                    
                    sNxt_TrOutWhs = StockInfo(Q).FromWarehouseCode
                    sNxt_ItemCode = StockInfo(Q).ItemCode
                    sNxt_TrInWhs = StockInfo(Q).ToWarehouseCode
    
'                    If sNxt_TrOutWhs = sCur_TrOutWhs And sCur_ItemCode = sNxt_ItemCode And sCur_TrInWhs = sNxt_TrInWhs Then
'                        '//ManBatchNum = 'Y' 이면 배치번호를 입력하지 않는다.
'                        If MDC_PS_Common.GetValue("SELECT ManBatchNum FROM OITM WHERE ITEMCODE = ''", 0, 1) = "Y" Then
'                            .BatchNumbers.Add
'                            .BatchNumbers.BatchNumber = StockInfo(Q).BatchNum
'                            .BatchNumbers.Quantity = Round(StockInfo(Q).BatchWeight, 2)
'                            .UserFields("Quantity").Value = .UserFields("Quantity").Value + Trim(StockInfo(Q).Qty)           '//수량
'                            .Quantity = .Quantity + Round(StockInfo(Q).Weight, 2)                                            '//중량을 합함
'                            .BatchNumbers.Notes = "재고이전(Addon)"
'                            StockInfo(Q).Chk = "Y"                                                                           '// 적용한 라인에 대한 표시
'                            StockInfo(Q).StockTransDocEntry = "Checked"
'                            StockInfo(Q).StockTransLineNum = StockTransLineCounter
'                        End If
'                    End If
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
            GoTo PS_MM132_StockTrans_Error
        End If
Continue_First:
    Next   '-----------------------------------------------------------------------------------------------< First For End
    
    For i = 0 To oMat01.VisualRowCount - 1
        ReQty = oDS_PS_MM132L.GetValue("U_ReQty", i)
        ReWt = oDS_PS_MM132L.GetValue("U_ReWt", i)
        
        oQuery01 = "Update [@PS_MM130L] "
        oQuery01 = oQuery01 & "Set U_ReQty = IsNull(U_ReQty, 0) + " & ReQty & ", U_ReWt = IsNull(U_ReWt, 0) + " & ReWt & " "
        oQuery01 = oQuery01 & "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry "
        oQuery01 = oQuery01 & "Where b.U_OutDoc = '" & Trim(oDS_PS_MM132L.GetValue("U_OutDoc", i)) & "' "
        oQuery01 = oQuery01 & "And a.U_LineNum = '" & Trim(oDS_PS_MM132L.GetValue("U_OutLine", i)) & "' "
        Call lRecordSet.DoQuery(oQuery01)
    Next i
            
    If (Sbo_Company.InTransaction) Then
        Sbo_Company.EndTransaction (wf_Commit)
    End If
    Call Sbo_Application.SetStatusBarMessage(DocCnt & " 개의 재고이전 문서가 발행되었습니다 !", bmt_Short, False)
    Set oStockTrans = Nothing
    Set lRecordSet = Nothing
    Set oRecordSet = Nothing
    Exit Function
'************Error Process************
PS_MM132_StockTrans_Error:
    If (Sbo_Company.InTransaction) Then
        Sbo_Company.EndTransaction (wf_RollBack)
    End If
    Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
    Call Sbo_Application.SetStatusBarMessage(ErrCode & " : " & ErrMsg, bmt_Short, True)
    PS_MM132_StockTrans = False
    Set oStockTrans = Nothing
    Set lRecordSet = Nothing
    Set oRecordSet = Nothing
'************Error Process************

End Function


Private Function PS_MM132_UpdateUserField() As Boolean
On Error GoTo PS_MM132_UpdateUserField_Error
    Dim i As Long
    Dim lQuery As String
    Dim lRecordSet As SAPbobsCOM.Recordset
    Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oDS_PS_MM132H.setValue("U_STDocNo", 0, (StockInfo(i).StockTransDocEntry))
    
    PS_MM132_UpdateUserField = True
    Exit Function
PS_MM132_UpdateUserField_Error:
    PS_MM132_UpdateUserField = False
End Function

Private Function PS_MM132_Update_Cancel() As Boolean
On Error GoTo PS_MM132_Update_Cancel_Error
    Dim i As Long
    Dim lQuery As String
    Dim lRecordSet As SAPbobsCOM.Recordset
    Set lRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oDS_PS_MM132H.setValue("U_STDocCl", 0, (StockInfo(i).StockTransDocEntry))
'    oForm01.Items("StoTrDoc").Specific.Value = StockInfo(i).StockTransDocEntry

    PS_MM132_Update_Cancel = True
    Exit Function
PS_MM132_Update_Cancel_Error:
    PS_MM132_Update_Cancel = False
End Function

Private Function PS_MM132_Cancel_oStockTrans(ChkType As Integer) As Boolean
On Error GoTo PS_MM132_Cancel_oStockTrans_Error
    Dim oStockTrans  As SAPbobsCOM.StockTransfer
    Dim oQuery01$
    Dim i&
    Dim ErrNum&, ErrCode As Long, ErrMsg$, RetVal&
    Dim RtnDocNum     As String
    Dim ReQty As Currency, ReWt As Currency
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim DocEntry As String
    DocEntry = Trim(oDS_PS_MM132H.GetValue("U_STDocNo", 0))
    
    If Trim(oDS_PS_MM132H.GetValue("U_STDocNo", 0)) <> "" Then
        Sbo_Company.StartTransaction
        Set oStockTrans = Nothing
        Set oStockTrans = Sbo_Company.GetBusinessObject(oStockTransfer)
                
        '//완료
        If (oStockTrans.GetByKey(DocEntry) = False) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            GoTo PS_MM132_Cancel_oStockTrans_Error
        End If
        RetVal = oStockTrans.Cancel
        If (0 <> RetVal) Then
            Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
            ErrNum = 1
            GoTo PS_MM132_Cancel_oStockTrans_Error
        End If
        
        If ChkType = 1 Then
            Sbo_Company.EndTransaction wf_RollBack
        ElseIf ChkType = 2 Then
            Call oDS_PS_MM132H.setValue("U_STDocCl", 0, RtnDocNum)
            Call oDS_PS_MM132H.setValue("Status", 0, "C")
            Call oDS_PS_MM132H.setValue("Canceled", 0, "Y")
            
            oQuery01 = "Select Max(DocEntry) From [OWTR]"
            Call oRecordSet01.DoQuery(oQuery01)
            RtnDocNum = oRecordSet01.Fields(0).VALUE
            Call Sbo_Company.GetNewObjectCode(RtnDocNum)
            
            For i = 0 To oMat01.VisualRowCount - 1
                ReQty = oDS_PS_MM132L.GetValue("U_ReQty", i)
                ReWt = oDS_PS_MM132L.GetValue("U_ReWt", i)
                
                oQuery01 = "Update [@PS_MM130L] "
                oQuery01 = oQuery01 & "Set U_ReQty = IsNull(U_ReQty, 0) - " & ReQty & ", U_ReWt = IsNull(U_ReWt, 0) - " & ReWt & " "
                oQuery01 = oQuery01 & "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry "
                oQuery01 = oQuery01 & "Where b.U_OutDoc = '" & Trim(oDS_PS_MM132L.GetValue("U_OutDoc", i)) & "' "
                oQuery01 = oQuery01 & "And a.U_LineNum = '" & Trim(oDS_PS_MM132L.GetValue("U_OutLine", i)) & "' "
                Call oRecordSet01.DoQuery(oQuery01)
            Next i
            
            Sbo_Company.EndTransaction wf_Commit

            
            oForm01.Items("STDocCl").Specific.VALUE = MDC_GetData.Get_ReData("MAX(DocNum)", "ObjType", "OWTR", "'67'", "")
            
        End If
    End If
    
    Set oStockTrans = Nothing
    Set oRecordSet01 = Nothing
    PS_MM132_Cancel_oStockTrans = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_MM132_Cancel_oStockTrans_Error:
    Set oStockTrans = Nothing
    Set oRecordSet01 = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    PS_MM132_Cancel_oStockTrans = False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "PS_MM132_Cancel_oStockTrans_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "PS_MM132_Cancel_oStockTrans_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function



