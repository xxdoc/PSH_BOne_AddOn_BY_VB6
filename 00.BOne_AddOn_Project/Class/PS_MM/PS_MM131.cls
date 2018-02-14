VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM131"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM131.cls
'//  Module         : 재고관리 > 외주반출등록 서브폼
'//  Desc           : MM
'//  FormType       : PS_MM130
'//  Create Date    : 2010.10.14
'//  Copyright  (c) PoongSan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
'Public oMat02 As SAPbouiCOM.Matrix

Private oDS_PS_MM131L As SAPbouiCOM.DBDataSource '등록라인

Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMat01Row01 As Long
Public oRadioGrp As String

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
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal RadioGrp As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM131.srf")
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
    
    oFormUniqueID01 = "PS_MM131_" & GetTotalFormsCount
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
    oRadioGrp = RadioGrp
'    oBaseTradeType01 = oTradeType02
    
    Call PS_MM131_CreateItems
    Call PS_MM131_ComboBox_Setting
    Call PS_MM131_CF_ChooseFromList
    Call PS_MM131_FormItemEnabled
    Call PS_MM131_EnableMenus
    '//Call PS_MM131_FormClear '//UDO방식일때
    '//Call PS_MM131_AddMatrixRow(0, True) '//UDO방식일때
       
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
                '//Call PS_MM131_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_MM131_FormItemEnabled '//UDO방식
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
        
        If pval.ItemUID = "Btn02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                 Call PS_MM131_Select_Data
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_MM131_SetBaseForm   '//부모폼에입력
                Call oForm01.Close
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_MM131_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_MM131" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_MM131_FormItemEnabled
'                    Call PS_MM131_FormClear '//UDO방식일때
'                    Call PS_MM131_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_MM131_FormItemEnabled
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CpCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItmMsort", "") '//사용자값활성
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
    Dim i&
    Dim Query01             As String
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim oCombo          As SAPbouiCOM.ComboBox
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
    Set oCombo = oForm01.Items("ItmMsort").Specific
    If pval.BeforeAction = True Then
 
    ElseIf pval.BeforeAction = False Then
    If pval.ItemChanged = True Then
        If pval.ItemUID = "ItmBsort" Then
        
        Dim sCount&, sSeq&
            sCount = oCombo.ValidValues.Count
            sSeq = sCount
        For i = 1 To sCount
            oCombo.ValidValues.Remove sSeq - 1, psk_Index
            sSeq = sSeq - 1
        Next i
        Call oForm01.Items("ItmMsort").Specific.ValidValues.Add("", "전체")
        Query01 = "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] Where U_rCode = '" & oForm01.Items("ItmBsort").Specific.VALUE & "'"
        Call oRecordSet01.DoQuery(Query01)
        
        Do Until oRecordSet01.EOF
            oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
            oRecordSet01.MoveNext
        Loop
        End If
    End If
    Set oRecordSet01 = Nothing
    Set oCombo = Nothing
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
        If oForm01.Mode = fm_ADD_MODE Then
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)
                oMat01Row01 = pval.Row
            End If
        ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        ElseIf oForm01.Mode = fm_OK_MODE Then
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
    Dim ItemCode01 As String
    Dim Query01             As String
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
    Dim Query02             As String
    Dim oRecordset02        As SAPbobsCOM.Recordset
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
    
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If pval.ItemUID = "Mat01" Then
                If pval.ColUID = "SelQty" Then
                    If oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE <= 0 Then
                        oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE = 0
                        oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = 0
                    Else
                        ItemCode01 = oMat01.Columns("ItemCode").Cells(pval.Row).Specific.VALUE
                        '//EA자체품
                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101" Or MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "601") Then
                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01))
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = Round(Val(oMat01.Columns("SelQty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0)
                        '//KG입력
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then
                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = 1
'                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "701") Then
'                            oMat01.Columns("SelQty").Cells(pval.Row).Specific.Value = 1
                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = 1
                        Else
                            oMat01.Columns("SelWeight").Cells(pval.Row).Specific.VALUE = 1
                        End If
                    End If
                    oForm01.Update
                ElseIf pval.ColUID = "PosQty" Then
                    If oMat01.Columns("PosQty").Cells(pval.Row).Specific.VALUE > 0 Then
                        Query01 = "Select U_ObasUnit FROM OITM WHERE ItemCode = '" & Trim(oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE) & "'"
                        Call oRecordSet01.DoQuery(Query01)
                        Query02 = "Select OnHand, U_Qty FROM OITW WHERE ItemCode = '" & Trim(oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE) & "' AND WhsCode = '101'"
                        Call oRecordset02.DoQuery(Query02)
                        If Left(oRecordSet01.Fields(0).VALUE, 1) = "1" Then
                            If (Trim(oRecordset02.Fields(0).VALUE)) > 0 And oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE <> "" Then
                                oMat01.Columns("PosWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("PosQty").Cells(pval.Row).Specific.VALUE
                            Else
                                oMat01.Columns("PosWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("PosQty").Cells(pval.Row).Specific.VALUE
                            End If
                        ElseIf Left(oRecordSet01.Fields(0).VALUE, 1) = "2" Then
                            If (Trim(oRecordset02.Fields(1).VALUE)) > 0 And oMat01.Columns("OutItmCd").Cells(pval.Row).Specific.VALUE <> "" Then
                                oMat01.Columns("PosWt").Cells(pval.Row).Specific.VALUE = (Trim(oRecordset02.Fields(0).VALUE) / Trim(oRecordset02.Fields(1).VALUE)) * oMat01.Columns("PosQty").Cells(pval.Row).Specific.VALUE
                            Else
                                oMat01.Columns("PosWt").Cells(pval.Row).Specific.VALUE = oMat01.Columns("PosQty").Cells(pval.Row).Specific.VALUE
                            End If
                        Else
                        End If
                        oMat01.Columns("PosWt").Cells(pval.Row).Click ct_Regular
                    End If
                ElseIf pval.ColUID = "CpCode" Then
                    
                End If
            Else
                Query01 = "Select U_CpName From [@PS_PP001L] Where U_CpCode = '" & Trim(oForm01.Items("CpCode").Specific.VALUE) & "'"
                Call oRecordSet01.DoQuery(Query01)
                oForm01.Items("CpName").Specific.VALUE = oRecordSet01.Fields(0).VALUE
            End If
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
        Call PS_MM131_FormItemEnabled
        '//Call PS_MM131_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then

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
                oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
                oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
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
    '        Call oDS_PS_MM131L.RemoveRecord(oDS_PS_MM131L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_MM131_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_MM131L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_MM131_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_MM131_CreateItems() As Boolean
On Error GoTo PS_MM131_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM131L = oForm01.DataSources.DBDataSources("@PS_USERDS01")

    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    oForm01.Items("Mat01").Enabled = True
        
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    
    Call oForm01.DataSources.UserDataSources.Add("ItmBsort", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItmBsort").Specific.DataBind.SetBound(True, "", "ItmBsort")
    
    Call oForm01.DataSources.UserDataSources.Add("ItmMsort", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItmMsort").Specific.DataBind.SetBound(True, "", "ItmMsort")
    
    Call oForm01.DataSources.UserDataSources.Add("Mark", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Mark").Specific.DataBind.SetBound(True, "", "Mark")
    
    Call oForm01.DataSources.UserDataSources.Add("ItemType", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemType").Specific.DataBind.SetBound(True, "", "ItemType")
    
    Call oForm01.DataSources.UserDataSources.Add("Size", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Size").Specific.DataBind.SetBound(True, "", "Size")
    
    Call oForm01.DataSources.UserDataSources.Add("CpCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CpCode").Specific.DataBind.SetBound(True, "", "CpCode")
    
    Call oForm01.DataSources.UserDataSources.Add("CpName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CpName").Specific.DataBind.SetBound(True, "", "CpName")
        
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_MM131_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM131_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_MM131_ComboBox_Setting()
On Error GoTo PS_MM131_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    

    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    Call oForm01.Items("ItmBsort").Specific.ValidValues.Add("", "전체")
    
    Set oCombo = oForm01.Items("ItmBsort").Specific
    sQry = "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    
    
   ' Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code", "", False, True)
   ' Call oForm01.Items("ItmMsort").Specific.ValidValues.Add("", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT]", "", False, True)

    Call MDC_SetMod.Set_ComboList(oForm01.Items("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] order by Code", "", False, True)
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType").Specific, "SELECT Code, Name FROM [@PSH_SHAPE] order by Code", "", False, True)
    
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM131_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM131_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM131_CF_ChooseFromList()
On Error GoTo PS_MM131_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column
'
    Set oEdit = oForm01.Items("ItemCode").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = "4"
    oCFLCreationParams.uniqueID = "CFLITEMCD"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)
'
'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oEdit.ChooseFromListUID = "CFLITEMCD"
    oEdit.ChooseFromListAlias = "ItemCode"
    Exit Sub
PS_MM131_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM131_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM131_FormItemEnabled()
On Error GoTo PS_MM131_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_MM131_FormClear '//UDO방식
'          oMat01.Columns("CHK").Editable = True
'          oMat01.Columns("PosQty").Editable = True
'          oMat01.Columns("PosWt").Editable = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM131_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM131_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM131_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_MM131_AddMatrixRow_Error
    Call oForm01.Freeze(True)
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_MM131L.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_MM131L.Offset = oRow
'    oDS_PS_MM131L.setValue "U_LineNum", oRow, oRow + 1
'    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM131_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM131_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM131_FormClear()
On Error GoTo PS_MM131_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM131'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_MM131_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM131_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM131_EnableMenus()
On Error GoTo PS_MM131_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_MM131_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM131_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PS_MM131_Select_Data()
On Error GoTo PS_MM131_Select_Data_Error
    Dim i As Long
    Dim j As Long
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim Param08 As String
    
    Dim vReturnValue$
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim oQuery02 As String
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Param01 = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    Param02 = Trim(oForm01.Items("ItmBsort").Specific.VALUE)
    Param03 = Trim(oForm01.Items("ItmMsort").Specific.VALUE)
    Param04 = Trim(oForm01.Items("Mark").Specific.VALUE)
    Param05 = Trim(oForm01.Items("ItemType").Specific.VALUE)
    Param06 = Trim(oForm01.Items("FrgnName").Specific.VALUE)
    Param07 = Trim(oForm01.Items("Size").Specific.VALUE)
    Param08 = Trim(oForm01.Items("CpCode").Specific.VALUE)
    
    If (Left(Param01, 3) = "102" Or Param02 = "102") And oRadioGrp = "B" And Param08 = "" Then
        vReturnValue = Sbo_Application.MessageBox("부품 재공반출은 공정코드를 선택해야 합니다.", 1, "OK")
        Call oForm01.Freeze(False)
        Exit Sub
    End If
    
    If oRadioGrp = "A" Then
        oQuery01 = "EXEC PS_MM131_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "'"
    ElseIf oRadioGrp = "B" Then
        oQuery01 = "EXEC PS_MM131_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "', '" & Param08 & "'"
    End If
    
    oRecordSet01.DoQuery (oQuery01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (oRecordSet01.RecordCount = 0) Then
        oForm01.Items("Mat01").Enabled = False
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
        GoTo PS_MM131_Select_Data_Exit
    Else
        oForm01.Items("Mat01").Enabled = True
    End If
    
    Dim oProgressBar01 As SAPbouiCOM.ProgressBar
    Set oProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    j = 0
    For i = 0 To oRecordSet01.RecordCount - 1
        If oRadioGrp = "A" Then
            If j <> 0 Then
                oDS_PS_MM131L.InsertRecord (j)
            End If
        
            oDS_PS_MM131L.Offset = j
            Call oDS_PS_MM131L.setValue("U_LineNum", j, j + 1)
            Call oDS_PS_MM131L.setValue("U_ColReg01", j, False)
            Call oDS_PS_MM131L.setValue("U_ColReg02", j, oRecordSet01.Fields("OrdNum").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg15", j, oRecordSet01.Fields("JisNo").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg03", j, oRecordSet01.Fields("ItemCode").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg04", j, oRecordSet01.Fields("ItemName").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg13", j, oRecordSet01.Fields("ItemCod2").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg14", j, oRecordSet01.Fields("ItemNam2").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg05", j, oRecordSet01.Fields("Size").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColQty04", j, oRecordSet01.Fields("UnWeight").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg10", j, oRecordSet01.Fields("Mark").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColNum01", j, oRecordSet01.Fields("SelQty").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColQty01", j, oRecordSet01.Fields("SelWt").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColNum02", j, oRecordSet01.Fields("OutQty").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColQty02", j, oRecordSet01.Fields("OutWt").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg06", j, oRecordSet01.Fields("OutWhCd").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg07", j, oRecordSet01.Fields("OutWhNm").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg08", j, oRecordSet01.Fields("InWhCd").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg09", j, oRecordSet01.Fields("InWhNm").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColNum03", j, oRecordSet01.Fields("PosQty").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColQty03", j, oRecordSet01.Fields("PosWt").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg11", j, oRecordSet01.Fields("PP030HNo").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg12", j, oRecordSet01.Fields("PP030MNo").VALUE)
            
            oQuery02 = "Select U_CpCode, U_CpName From [@PS_PP030M] Where DocEntry = '" & oRecordSet01.Fields("PP030HNo").VALUE & "' And U_Sequence = '" & oRecordSet01.Fields("PP030MNo").VALUE & "'"
            
            oRecordset02.DoQuery (oQuery02)
  
            If (oRecordset02.RecordCount > 0) Then
                Call oDS_PS_MM131L.setValue("U_ColReg16", j, oRecordset02.Fields("U_CpCode").VALUE)
                Call oDS_PS_MM131L.setValue("U_ColReg17", j, oRecordset02.Fields("U_CpName").VALUE)
            End If
            
            j = j + 1
            
        ElseIf oRadioGrp = "B" And oRecordSet01.Fields("PosQty").VALUE > 0 Then
            If j <> 0 Then
                oDS_PS_MM131L.InsertRecord (j)
            End If
        
            oDS_PS_MM131L.Offset = j
            Call oDS_PS_MM131L.setValue("U_LineNum", j, j + 1)
            Call oDS_PS_MM131L.setValue("U_ColReg01", j, False)
            Call oDS_PS_MM131L.setValue("U_ColReg02", j, oRecordSet01.Fields("OrdNum").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg15", j, oRecordSet01.Fields("JisNo").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg03", j, oRecordSet01.Fields("ItemCode").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg04", j, oRecordSet01.Fields("ItemName").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg13", j, oRecordSet01.Fields("ItemCode").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg14", j, oRecordSet01.Fields("ItemName").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg05", j, oRecordSet01.Fields("Size").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColQty04", j, oRecordSet01.Fields("UnWeight").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg10", j, oRecordSet01.Fields("Mark").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColNum01", j, oRecordSet01.Fields("SelQty").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColQty01", j, oRecordSet01.Fields("SelWt").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColNum02", j, oRecordSet01.Fields("OutQty").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColQty02", j, oRecordSet01.Fields("OutWt").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg06", j, oRecordSet01.Fields("OutWhCd").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg07", j, oRecordSet01.Fields("OutWhNm").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg08", j, oRecordSet01.Fields("InWhCd").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg09", j, oRecordSet01.Fields("InWhNm").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColNum03", j, oRecordSet01.Fields("PosQty").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColQty03", j, oRecordSet01.Fields("PosWt").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg11", j, oRecordSet01.Fields("PP030HNo").VALUE)
            Call oDS_PS_MM131L.setValue("U_ColReg12", j, oRecordSet01.Fields("PP030MNo").VALUE)
            
            oQuery02 = "Select U_CpCode, U_CpName From [@PS_PP030M] Where DocEntry = '" & oRecordSet01.Fields("PP030HNo").VALUE & "' And U_Sequence = '" & oRecordSet01.Fields("PP030MNo").VALUE & "'"
            
            oRecordset02.DoQuery (oQuery02)
  
            If (oRecordset02.RecordCount > 0) Then
                Call oDS_PS_MM131L.setValue("U_ColReg16", j, oRecordset02.Fields("U_CpCode").VALUE)
                Call oDS_PS_MM131L.setValue("U_ColReg17", j, oRecordset02.Fields("U_CpName").VALUE)
            End If
            j = j + 1
        End If
        
        oRecordSet01.MoveNext
        oProgressBar01.VALUE = oProgressBar01.VALUE + 1
        oProgressBar01.Text = oProgressBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    oProgressBar01.Stop
    Set oProgressBar01 = Nothing
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM131_Select_Data_Exit:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM131_Select_Data_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM311_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_MM131_DataValidCheck() As Boolean
On Error GoTo PS_MM131_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_MM131_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_MM131_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_MM131_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_MM131L.RemoveRecord(oDS_MM131L.Size - 1)
'    Call oMat01.LoadFromDataSource
'    Call PS_MM131_FormClear
    Exit Function
PS_MM131_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM131_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_MM131_SetBaseForm()
On Error GoTo PS_MM131_SetBaseForm_Error
    Dim i As Long
    Dim j As Long
    Dim ItemCode01 As String
    Dim oBaseMat01 As SAPbouiCOM.Matrix
    Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific
    
    Dim oDS_PS_MM130L As SAPbouiCOM.DBDataSource
    Set oDS_PS_MM130L = oBaseForm01.DataSources.DBDataSources("@PS_MM130L")
    If oBaseForm01 Is Nothing Then
        '//DoNothing
    ElseIf oBaseForm01.TypeEx = "PS_MM130" Then
'        For j To oBaseMat01.VisualRowCount
            For i = 1 To oMat01.VisualRowCount '//품목선택품목
            
            If oMat01.Columns("CHK").Cells(i).Specific.Checked = True Then
                ' Call oDS_PS_MM130L.setValue("U_OrdNum", j - 1, oMat01.Columns("OrdNum").Cells(i).Specific.Value)
                ' DBDataSource일 경우 부모폼에 Validate Event가 발생하지 않는다.
                oBaseMat01.Columns("OrdNum").Cells(oBaseMat01.VisualRowCount).Specific.VALUE = oMat01.Columns("OrdNum").Cells(i).Specific.VALUE
'                Call oDS_PS_MM130L.setValue("U_OrdNum", oBaseMat01.VisualRowCount - 1, oMat01.Columns("OrdNum").Cell(i).Specific.Value)
                Call oDS_PS_MM130L.setValue("U_ItemCode", oBaseMat01.VisualRowCount - 2, oMat01.Columns("ItemCode").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_ItemName", oBaseMat01.VisualRowCount - 2, oMat01.Columns("ItemName").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_Size", oBaseMat01.VisualRowCount - 2, oMat01.Columns("Size").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_UnWeight", oBaseMat01.VisualRowCount - 2, oMat01.Columns("UnWeight").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_Qty", oBaseMat01.VisualRowCount - 2, oMat01.Columns("SelQty").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_Weight", oBaseMat01.VisualRowCount - 2, oMat01.Columns("SelWt").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_Mark", oBaseMat01.VisualRowCount - 2, oMat01.Columns("Mark").Cells(i).Specific.VALUE)
                If oRadioGrp = "A" Then
                    Call oDS_PS_MM130L.setValue("U_OutGbn", oBaseMat01.VisualRowCount - 2, "10")
                    Call oDS_PS_MM130L.setValue("U_OutItmCd", oBaseMat01.VisualRowCount - 2, oMat01.Columns("OutItmCd").Cells(i).Specific.VALUE)
                    Call oDS_PS_MM130L.setValue("U_OutItmNm", oBaseMat01.VisualRowCount - 2, oMat01.Columns("OutItmNm").Cells(i).Specific.VALUE)
                ElseIf oRadioGrp = "B" Then
                    Call oDS_PS_MM130L.setValue("U_OutGbn", oBaseMat01.VisualRowCount - 2, "20")
                    Call oDS_PS_MM130L.setValue("U_OutItmCd", oBaseMat01.VisualRowCount - 2, "")
                    Call oDS_PS_MM130L.setValue("U_OutItmNm", oBaseMat01.VisualRowCount - 2, "")
                End If
                Call oDS_PS_MM130L.setValue("U_OutQty", oBaseMat01.VisualRowCount - 2, oMat01.Columns("PosQty").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_OutWt", oBaseMat01.VisualRowCount - 2, oMat01.Columns("PosWt").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_OutWhCd", oBaseMat01.VisualRowCount - 2, oMat01.Columns("OutWhCd").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_OutWhNm", oBaseMat01.VisualRowCount - 2, oMat01.Columns("OutWhNm").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_InWhCd", oBaseMat01.VisualRowCount - 2, oMat01.Columns("InWhCd").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_InWhNm", oBaseMat01.VisualRowCount - 2, oMat01.Columns("InWhNm").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_PP030HNo", oBaseMat01.VisualRowCount - 2, oMat01.Columns("PP030HNo").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_PP030MNo", oBaseMat01.VisualRowCount - 2, oMat01.Columns("PP030MNo").Cells(i).Specific.VALUE)
                
                Call oDS_PS_MM130L.setValue("U_CpCode", oBaseMat01.VisualRowCount - 2, oMat01.Columns("CpCode").Cells(i).Specific.VALUE)
                Call oDS_PS_MM130L.setValue("U_CpName", oBaseMat01.VisualRowCount - 2, oMat01.Columns("CpName").Cells(i).Specific.VALUE)
                j = j + 1
            Else
'                oBaseMat01.Columns("OrdNum").Cells(i).Click ct_Regular
'                oBaseColRow01 = oBaseColRow01 + 1
            End If

        Next
        
        oBaseMat01.LoadFromDataSource
    
    End If
    Exit Sub
PS_MM131_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM131_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



