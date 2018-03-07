VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP095"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP095.cls
'//  Module         : M/G거래명세표등록
'//  Desc           :
'//  FormType       : PS_PP095
'//  Create Date    : 2013.08.29   NGY
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_PP095H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP095L As SAPbouiCOM.DBDataSource '등록라인
Private QREncode As New cQREncode

Private oDocType01 As String

Public oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Public oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Public oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

''//사용자구조체
'Private Type ItemInformations
'    ItemCode As String
'    BatchNum As String
'    Qty As Long '//수량
'    Weight As Double '//중량
'    Currency As String '//통화
'    Price As Double '//단가
'    LineTotal As Double '//총계
'    WhsCode As String '//창고
'    SD030HNum As Long '//출하(선출)문서
'    SD030LNum As Long '//출하(선출)라인
'    SD040HNum As Long '//납품A문서
'    SD040LNum As Long '//납품A라인
'    ORDRNum As Long '//판매오더문서
'    RDR1Num As Long '//판매오더라인
'    Check As Boolean
'    ODLNNum As Long '//납품문서
'    DLN1Num As Long '//납품라인
'    ORDNNum As Long '//반품문서
'    RDN1Num As Long '//반품라인
'End Type
'Private ItemInformation() As ItemInformations
'Private ItemInformationCount As Long
'
'Private Type BatchInformations
'    ItemCode As String '//품목코드
'    WhsCode As String '//창고코드
'    BatchNum As String '//배치번호
'    Weight As Double '//중량
'End Type
'Private BatchInformation() As BatchInformations
'Private BatchInformationCount As Long


'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String, Optional ByVal oFromDocType01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP095.srf")
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
    
    oFormUniqueID01 = "PS_PP095_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    oDocType01 = oFromDocType01
    
    Call PS_PP095_CreateItems
    Call PS_PP095_ComboBox_Setting
    Call PS_PP095_EnableMenus
    Call PS_PP095_SetDocument(oFromDocEntry01)
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
    Dim i As Long
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
'                '//납품취소, 반품
'                If oForm01.Mode = fm_OK_MODE Then
'                    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_SD040H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
'                        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
'                    If oForm01.Items("Opt02").Specific.Selected = True Then '//멀티이면서
'                        For i = 1 To oMat01.VisualRowCount - 1 '//멀티게이지 재고수량검사
'                            If MDC_PS_Common.GetValue("SELECT Quantity FROM [OIBT] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & "' AND BatchNum = '" & oMat01.Columns("LotNo").Cells(i).Specific.VALUE & "' AND WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
'                                Call MDC_Com.MDC_GF_Message("멀티게이지 품목 : " & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & " 의 재고가 존재합니다.", "W")
'                                BubbleEvent = False
'                                Exit Sub
'                            End If
'                        Next
'                    End If
'                    For i = 1 To oMat01.VisualRowCount - 1 '//AR송장처리된 문서존재유무검사
'                        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [INV1] WHERE BaseType = '15' AND BaseEntry = '" & oMat01.Columns("ODLNNum").Cells(i).Specific.VALUE & "' AND BaseLine = '" & oMat01.Columns("DLN1Num").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
'                            Call MDC_Com.MDC_GF_Message("AR송장처리된 문서가 존재합니다.", "W")
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
'                    Next
'                    For i = 1 To oMat01.VisualRowCount - 1 '//반품처리된 문서존재유무검사
'                        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [RDN1] WHERE BaseType = '15' AND BaseEntry = '" & oMat01.Columns("ODLNNum").Cells(i).Specific.VALUE & "' AND BaseLine = '" & oMat01.Columns("DLN1Num").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
'                            Call MDC_Com.MDC_GF_Message("반품처리된 문서가 존재합니다.", "W")
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
'                    Next
'                    If PS_SD040_DI_API_03 = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
'                Else
'                    MDC_Com.MDC_GF_Message "현재 모드에서는 취소할수 없습니다.", "W"
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            Case "1286": '닫기
'                If oForm01.Mode = fm_OK_MODE Then
'                    If MDC_PS_Common.GetValue("SELECT Status FROM [@PS_SD040H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.Value & "'", 0, 1) = "C" Then
'                        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 닫기되었습니다. 작업을 진행할수 없습니다.", "W")
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
'                Else
'                    MDC_Com.MDC_GF_Message "현재 모드에서는 취소할수 없습니다.", "W"
'                    BubbleEvent = False
'                    Exit Sub
'                End If
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
                Call PS_PP095_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
'                Call PS_SD040_FormItemEnabled '//UDO방식
                Call PS_PP095_AddMatrixRow(0, True) '//UDO방식
'                '//2010.12.06 추가
'                '담당자
'                oDS_PS_SD040H.setValue "U_CntcCode", 0, User_MSTCOD()
'                Call oDS_PS_SD040H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items("CntcCode").Specific.VALUE & "'", 0, 1))
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
'                Call PS_SD040_FormItemEnabled '//UDO방식
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
'        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
'                Call Sbo_Application.Menus.RemoveEx("MenuUID")
'        End If
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error

Dim YM As String

    If pval.BeforeAction = True Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
'             If PS_PP095_Weight_Check = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If


                If PS_PP095_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
               
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_PP095_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If PS_PP095_Weight_Check = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_PP095_Print_Report01
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_PP095_Print_Report02
            End If
        End If
        If pval.ItemUID = "Button03" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                
                If PS_PP095_Check_QRCode_PrintYN = "N" Then
                
                    Call PS_PP095_Print_Report03 '일반 패킹리스트 출력
                
                ElseIf PS_PP095_Check_QRCode_PrintYN = "Y" Then
                
                    Call PS_PP095_Print_Report04 'QR코드 패킹리스트 출력
                
                End If
                
            End If
        End If
        If pval.ItemUID = "Button04" Then
            
            Dim ChildForm01 As Variant
            YM = Mid(oForm01.Items("DocDate").Specific.VALUE, 1, 4) + "-" + Mid(oForm01.Items("DocDate").Specific.VALUE, 5, 2)
            
            Set ChildForm01 = New PS_QM041
            Call ChildForm01.LoadForm(YM, oForm01.Items("DocEntry").Specific.VALUE)
                
            
'            If oForm01.Mode = fm_ADD_MODE Then
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'
'            End If
        End If
        If pval.ItemUID = "Button05" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                
                Dim TempForm01 As Variant
                Set TempForm01 = New PS_MM004
                Call TempForm01.LoadForm("PS_PP095", oForm01.Items("DocEntry").Specific.VALUE)
                BubbleEvent = False
                
            End If
        End If
        If pval.ItemUID = "Button06" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                
                Call PS_PP095_Print_Report04
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_PP095_FormItemEnabled
                    '//Call PS_PP095_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
                    oForm01.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1291")
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_PP095_FormItemEnabled
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
    Set ChildForm01 = New PS_PP095S
    
    If pval.BeforeAction = True Then
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "CardCode" Then
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
            If pval.ItemUID = "Mat01" Then
                If pval.ColUID = "OrdNum" Then
                    Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, oForm01.Items("BPLId").Specific.VALUE)
                                
                    BubbleEvent = False
                End If
        
            End If
        End If
        '//Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "OrderNum") '//사용자값활성
        '//Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "WhsCode") '//사용자값활성
'        If pval.ItemUID = "Mat01" Then
'            If pval.ColUID = "PackNo" Then
'                If oForm01.Items("TrType").Specific.Selected.VALUE <> "2" Then
'                    Call MDC_Com.MDC_GF_Message("포장번호조회는 거래형태가 임가공이여야 합니다.", "W")
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                If oForm01.Items("CardCode").Specific.VALUE = "" Then
'                    Call MDC_Com.MDC_GF_Message("고객코드는 필수입니다.", "W")
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                If oForm01.Items("DCardCod").Specific.VALUE = "" Then
'                    Call MDC_Com.MDC_GF_Message("납품처코드는 필수입니다.", "W")
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "PackNo") '//사용자값활성
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

    If pval.BeforeAction = True Then
'        If pval.ItemUID = "DocType" Then
'            If (oForm01.Items("DocType").Specific.Value = "1") Then
'            ElseIf (oForm01.Items("DocType").Specific.Value = "2") Then
'                oForm01.Items("BPLId").Enabled = True
'                Call oForm01.Items("BPLId").Specific.Select("2", psk_ByValue)
'                oForm01.Items("CardCode").Click ct_Regular
'                oForm01.Items("BPLId").Enabled = False
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
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
'        If pval.ItemUID = "Opt01" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                oForm01.Freeze True
'                oForm01.Items("Focus").Click ct_Regular
'                oForm01.Items("CardCode").Enabled = True
'                oForm01.Items("BPLId").Enabled = True
'                oForm01.Items("DCardCod").Enabled = True
'                oForm01.Items("TrType").Enabled = True
'                oMat01.Columns("SD030Num").Visible = True
'                oMat01.Columns("PackNo").Visible = False
'                oMat01.Columns("OrderNum").Visible = True
'                oMat01.Columns("Qty").Editable = True
'                oMat01.Columns("Weight").Editable = True
'                oMat01.Columns("UnWeight").Visible = True
'                oMat01.Columns("Price").Editable = False
'                Call oForm01.Items("TrType").Specific.Select("1", psk_ByValue)
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
'                Call PS_PP095_AddMatrixRow(0, True)
'                oForm01.Freeze False
'            End If
'        End If
'        If pval.ItemUID = "Opt02" Then '//멀티게이지의경우
'            If oForm01.Mode = fm_ADD_MODE Then
'                oForm01.Freeze True
'                oForm01.Items("Focus").Click ct_Regular
'                oForm01.Items("CardCode").Enabled = True
'                oForm01.Items("BPLId").Enabled = True
'                oForm01.Items("DCardCod").Enabled = True
'                oForm01.Items("TrType").Enabled = False
'                oMat01.Columns("SD030Num").Visible = False
'                oMat01.Columns("PackNo").Visible = True
'                oMat01.Columns("OrderNum").Visible = False
'                oMat01.Columns("Qty").Editable = False
'                oMat01.Columns("Weight").Editable = False
'                oMat01.Columns("UnWeight").Visible = False
'                oMat01.Columns("Price").Editable = True
'                Call oForm01.Items("TrType").Specific.Select("2", psk_ByValue)
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
'                Call PS_PP095_AddMatrixRow(0, True)
'                oForm01.Freeze False
'            End If
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
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
'            If pval.ColUID = "SD030Num" Then
'                Set oTempClass = New PS_SD030
'                Call oTempClass.LoadForm(Mid(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, 1, InStr(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, "-") - 1))
'            End If
'            If pval.ColUID = "SD030H" Then
'                Set oTempClass = New PS_SD030
'                Call oTempClass.LoadForm(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'            End If
        End If
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
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Dim RecordSet03 As SAPbobsCOM.Recordset
    Dim ItemCode01 As String
    Dim sQry As String
    Dim SumSjQty As Long, SumSjWt As Currency, SumQty As Long, SumWeight As Currency, HWeight As Currency
    Dim BoxWeight As Double
    
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If pval.ColUID = "OrdNum" Then
                    If oMat01.Columns("OrdNum").Cells(pval.Row).Specific.VALUE = "" Then
                        GoTo Raise_EVENT_VALIDATE_Exit
                    End If
                    
'                    For i = 1 To oMat01.RowCount
'                        If pval.Row <> i Then '//현재 선택되어있는 행이 아니면
'                            If (oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE = oMat01.Columns("SD030Num").Cells(i).Specific.VALUE) Then
'                                Call MDC_Com.MDC_GF_Message("동일한 출하(선출)요청이 존재합니다.", "W")
'                                oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE = ""
'                                GoTo Raise_EVENT_VALIDATE_Exit
'                            End If
'                            If (Mid(oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE, 1, InStr(oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE, "-") - 1) <> _
'                            Mid(oMat01.Columns("SD030Num").Cells(i).Specific.VALUE, 1, InStr(oMat01.Columns("SD030Num").Cells(i).Specific.VALUE, "-") - 1)) Then
'                                If Trim(oForm01.Items("BPLId").Specific.VALUE) = "4" Then '[2011.1.10추가]구로사업소일경우에만 동일하지 않은 출하요청문서 막음
'                                    Call MDC_Com.MDC_GF_Message("동일하지않은 출하요청문서가 존재합니다.", "W")
'                                    oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE = ""
'                                    GoTo Raise_EVENT_VALIDATE_Exit
'                                End If
'                            End If
'                        End If
'                    Next
                    
                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                  
                    Query01 = "SELECT PackNo   = a.U_PackNo, "
                    Query01 = Query01 & " ItemCode = b.U_ItemCode, "
                    Query01 = Query01 & " ItemName = b.U_ItemName, "
                    Query01 = Query01 & " Weight   = b.U_Weight, "
                    Query01 = Query01 & " LWeight  =  isnull((select max(boxwgt) as result from z_packlist where ordnum  in (select u_ordnum from [@ps_pp095l] where u_packno = a.U_PackNo)),0), "
                    Query01 = Query01 & " ProDate = B.U_ProDate "
                    Query01 = Query01 & " FROM [@PS_PP090H] a INNER JOIN [@PS_PP090L] b ON a.DocEntry = b.DocEntry AND a.CanCeled = 'N' "
                    Query01 = Query01 & " WHERE "
                    Query01 = Query01 & " a.U_BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "'"
                    Query01 = Query01 & " AND b.U_LotNo = '" & oMat01.Columns("OrdNum").Cells(pval.Row).Specific.VALUE & "'"
                    
                    Call RecordSet01.DoQuery(Query01)
                    
                    For i = 0 To RecordSet01.RecordCount - 1
                        Call oDS_PS_PP095L.setValue("U_OrdNum", pval.Row - 1, oMat01.Columns("OrdNum").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP095L.setValue("U_PackNo", pval.Row - 1, RecordSet01.Fields("PackNo").VALUE)
                        Call oDS_PS_PP095L.setValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                        Call oDS_PS_PP095L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
                        Call oDS_PS_PP095L.setValue("U_Weight", pval.Row - 1, Val(RecordSet01.Fields("Weight").VALUE))
                        Call oDS_PS_PP095L.setValue("U_ProDate", pval.Row - 1, Format(RecordSet01.Fields("ProDate").VALUE, "YYYYMMDD"))
                        RecordSet01.MoveNext
                    Next
                    
                    
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP095L.GetValue("U_OrdNum", pval.Row - 1)) <> "" Then
                        PS_PP095_AddMatrixRow (pval.Row)
                    End If
'                    PS_PP095_AddMatrixRow (pval.Row)
                    
                    oMat01.LoadFromDataSource
                    oMat01.AutoResizeColumns
                   
                    oForm01.Update
                    Set RecordSet01 = Nothing
                    
                End If
                
                oMat01.LoadFromDataSource
                oMat01.AutoResizeColumns
                oForm01.Update
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                
            Else
                If (pval.ItemUID = "CntcCode") Then
                    Call oDS_PS_PP095H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_PP095H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                End If
                If (pval.ItemUID = "CardCode") Then
                    Call oDS_PS_PP095H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_PP095H.setValue("U_CardName", 0, MDC_PS_Common.GetValue("select cardname from ocrd where cardcode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
                End If
                If (pval.ItemUID = "TWeight") Then
                    If oForm01.Items("TWeight").Specific.VALUE > 0 Then
                        If oForm01.Mode = fm_UPDATE_MODE Or oForm01.Mode = fm_OK_MODE Then
                            BoxWeight = Round((CDbl(oForm01.Items("TWeight").Specific.VALUE) - CDbl(oForm01.Items("S_Weight").Specific.VALUE)) / CDbl(oForm01.Items("BoxCnt").Specific.VALUE), 1)
                            oForm01.Items("Comments").Specific.VALUE = CStr(CInt(oForm01.Items("BoxCnt").Specific.VALUE)) + "EA X " + CStr(BoxWeight) + "Kg = " + CStr(CDbl(oForm01.Items("TWeight").Specific.VALUE) - CDbl(oForm01.Items("S_Weight").Specific.VALUE)) + " Kg"
                        End If
                    Else
                        oForm01.Items("Comments").Specific.VALUE = ""
                    End If
                End If
                
            End If
            
            
        End If
    ElseIf pval.BeforeAction = False Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Exit:
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    Dim i&, S_Weight As Long
    Dim PackNo As String
    Dim sQry As String
    Dim BoxCnt&
    Dim result As Integer
        
    Dim SS_Weight As Long
    Dim PackNo2 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    



    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
        Call PS_PP095_FormItemEnabled
        Call PS_PP095_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        
        PackNo = oMat01.Columns("PackNo").Cells(1).Specific.VALUE
        BoxCnt = 0
        PackNo2 = 0
        If oMat01.VisualRowCount <= 100 Then
            For i = 0 To oMat01.VisualRowCount - 2
                PackNo = Trim(oDS_PS_PP095L.GetValue("U_PackNo", i))
                S_Weight = S_Weight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                
                If i = (oMat01.VisualRowCount - 1) Then
                    PackNo2 = "123"
                Else
                    PackNo2 = Trim(oDS_PS_PP095L.GetValue("U_PackNo", i + 1))
                End If
                If PackNo <> PackNo2 Then

                        sQry = " EXEC [PS_PP095_13] "
                        sQry = sQry & "'" & PackNo & "'"

                        Call oRecordSet01.DoQuery(sQry)

                        result = oRecordSet01.Fields("RESULT").VALUE

                        SS_Weight = SS_Weight + result
                        BoxCnt = BoxCnt + 1
                End If
                
            Next i
            
            oForm01.Items("BoxCnt").Specific.VALUE = BoxCnt
            oForm01.Items("S_Weight").Specific.VALUE = S_Weight
            oForm01.Items("SS_Weight").Specific.VALUE = SS_Weight
            
        Else
            oForm01.Items("BoxCnt").Specific.VALUE = 0
            oForm01.Items("S_Weight").Specific.VALUE = 0
        End If
        
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
    Dim oDataTable01 As SAPbouiCOM.DataTable
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP095H", "U_CardCode,U_CardName")
        End If
'        If (pval.ItemUID = "DCardCod" Or pval.ItemUID = "DCardNam") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP095H", "U_DCardCod,U_DCardNam")
'        End If
        If (pval.ItemUID = "Mat01") Then
'            If (pval.ColUID = "WhsCode") Then
'                If pval.SelectedObjects Is Nothing Then
'                Else
'                    Set oDataTable01 = pval.SelectedObjects
'                    Call oDS_PS_PP095L.setValue("U_WhsCode", pval.Row - 1, oDataTable01.Columns("WhsCode").Cells(0).VALUE)
'                    Call oDS_PS_PP095L.setValue("U_WhsName", pval.Row - 1, oDataTable01.Columns("WhsName").Cells(0).VALUE)
'                    Set oDataTable01 = Nothing
'                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
'                    oMat01.LoadFromDataSource
'                    oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
'                End If
'            End If
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
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//행삭제전 행삭제가능여부검사
'            If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'                MDC_Com.MDC_GF_Message "납품된행은 삭제할수 없습니다.", "W"
'                BubbleEvent = False
'                Exit Sub
'            End If
'            If oForm01.Items("Opt02").Specific.Selected = True Then '//멀티게이지
'                For i = 0 To oDS_PS_PP095L.Size - 1
'                    If i = oDS_PS_PP095L.Size - 1 Then
'                        Exit For
'                    End If
'                    If oDS_PS_PP095L.GetValue("U_PackNo", i) = oDS_PS_PP095L.GetValue("U_PackNo", oLastColRow01 - 1) Then '//선택된행과 같은값을 가지는 모든행
'                        Call oDS_PS_PP095L.RemoveRecord(i)
'                        i = i - 1
'                    End If
'                Next
'                For i = 0 To oDS_PS_PP095L.Size - 1
'                    Call oDS_PS_PP095L.setValue("U_LineNum", i, i + 1)
'                Next
'                oMat01.LoadFromDataSource '//행삭제 정상처리
'
'                If oMat01.RowCount = 0 Then
'                    Call PS_PP095_AddMatrixRow(0)
'                Else
'                    If Trim(oDS_PS_PP095L.GetValue("U_SD030Num", oMat01.RowCount - 1)) <> "" Then
'                        Call PS_PP095_AddMatrixRow(oMat01.RowCount)
'                    End If
'                End If
'                If (oMat01.RowCount > 1) Then
'                    oForm01.Items("Focus").Click ct_Regular
'                    oForm01.Items("CardCode").Enabled = False
'                    oForm01.Items("BPLId").Enabled = False
'                    oForm01.Items("TrType").Enabled = False
'                Else
'                    oForm01.Items("Focus").Click ct_Regular
'                    oForm01.Items("CardCode").Enabled = True
'                    oForm01.Items("BPLId").Enabled = True
'                    oForm01.Items("DCardCod").Enabled = True
'                    oForm01.Items("TrType").Enabled = True
'                End If
'                oForm01.Update
''                MDC_Com.MDC_GF_Message "멀티게이지는 포장번호 단위로 납품가능합니다. 행삭제 할수 없습니다.", "W"
'                BubbleEvent = False '//멀티의경우 이미 행삭제가 되었으므로 행삭제를 하지 않는다.
'                Exit Sub
'            End If
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            
            oMat01.FlushToDataSource
            Call oDS_PS_PP095L.RemoveRecord(oDS_PS_PP095L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_PP095_AddMatrixRow(0)
            Else
'                If Trim(oDS_PS_PP095L.GetValue("U_SD030Num", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_PP095_AddMatrixRow(oMat01.RowCount)
'                End If
            End If
'            If (oMat01.RowCount > 1) Then
'                oForm01.Items("Focus").Click ct_Regular
'                oForm01.Items("CardCode").Enabled = False
'                oForm01.Items("BPLId").Enabled = False
'                oForm01.Items("TrType").Enabled = False
'            Else
'                oForm01.Items("Focus").Click ct_Regular
'                oForm01.Items("CardCode").Enabled = True
'                oForm01.Items("BPLId").Enabled = True
'                oForm01.Items("DCardCod").Enabled = True
'                oForm01.Items("TrType").Enabled = True
'            End If
            oForm01.Update
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP095_CreateItems() As Boolean
On Error GoTo PS_PP095_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP095H = oForm01.DataSources.DBDataSources("@PS_PP095H")
    Set oDS_PS_PP095L = oForm01.DataSources.DBDataSources("@PS_PP095L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("S_Weight", dt_SUM)
    oForm01.Items("S_Weight").Specific.DataBind.SetBound True, "", "S_Weight"
    
    Call oForm01.DataSources.UserDataSources.Add("BoxCnt", dt_SUM)
    oForm01.Items("BoxCnt").Specific.DataBind.SetBound True, "", "BoxCnt"
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
       
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP095_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP095_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP095_ComboBox_Setting()
On Error GoTo PS_PP095_ComboBox_Setting_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP095", "Gubun", "", "1", "내수")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP095", "Gubun", "", "2", "수출")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Gubun").Specific, "PS_PS_PP095", "Gubun")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP095", "Way", "", "1", "편도")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP095", "Way", "", "2", "왕복")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Way").Specific, "PS_PS_PP095", "Way")
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP095", "TCHECK", "", "1", "미확인")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_PP095", "TCHECK", "", "2", "확인")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("TCHECK").Specific, "PS_PS_PP095", "TCHECK")
    
     Set oCombo = oForm01.Items("TCHECK").Specific
    Call oCombo.ValidValues.Add("1", "자동확인")
    Call oCombo.ValidValues.Add("2", "수동확인")
    Call oCombo.Select(0, psk_Index)
    

    Call oForm01.Freeze(False)
    
    Exit Sub
PS_PP095_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP095_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_PP095_FormItemEnabled()
On Error GoTo PS_PP095_FormItemEnabled_Error
    Dim i As Integer

    Call oForm01.Freeze(True)
    oForm01.Items("Focus").Click ct_Regular
    If (oForm01.Mode = fm_ADD_MODE) Then
'        '//각모드에따른 아이템설정
'        oForm01.Items("DocEntry").Enabled = False
'        oForm01.Items("CardCode").Enabled = True
'        oForm01.Items("DCardCod").Enabled = True
'        oForm01.Items("BPLId").Enabled = True
'        oForm01.Items("DocType").Enabled = True
'        oForm01.Items("CntcCode").Enabled = True
'        oForm01.Items("DocDate").Enabled = True
'        oForm01.Items("DueDate").Enabled = True
'        oForm01.Items("Comments").Enabled = True
'        oForm01.Items("TrType").Enabled = True
'        oForm01.Items("Opt01").Enabled = True
'        oForm01.Items("Opt02").Enabled = True
'        oForm01.Items("Mat01").Enabled = True
'
'        oMat01.AutoResizeColumns
'        Call PS_PP095_FormClear '//UDO방식
''        Call oForm01.Items("BPLId").Specific.Select("1", psk_ByValue)
'        Call oForm01.Items("BPLId").Specific.Select(User_BPLId(), psk_ByValue) '//2010.12.06 추가
'        Call oForm01.Items("DocType").Specific.Select("1", psk_ByValue)
'        oForm01.Items("CardCode").Click ct_Regular
'        oForm01.Items("DocType").Enabled = False
'        Call oForm01.Items("TrType").Specific.Select("1", psk_ByValue)
'        Call oForm01.EnableMenu("1281", True) '//찾기
'        Call oForm01.EnableMenu("1282", False) '//추가
'        oForm01.Items("DocDate").Enabled = True
'        oForm01.Items("DueDate").Enabled = True
'        oMat01.Columns("SD030Num").Visible = True
'        oMat01.Columns("PackNo").Visible = True
'        oMat01.Columns("OrderNum").Visible = True
'        oForm01.Items("DocDate").Specific.String = Format(Now, "yyyymmdd")
'        oForm01.Items("DueDate").Specific.String = Format(Now, "yyyymmdd")
'        oForm01.Items("ProgStat").Enabled = True
'        Call oForm01.Items("ProgStat").Specific.Select("3", psk_ByValue)
'        oForm01.Items("CardCode").Click ct_Regular
'        oForm01.Items("ProgStat").Enabled = False
'        oForm01.Items("Opt01").Specific.Selected = True
'
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocEntry").Enabled = True
'        oForm01.Items("CardCode").Enabled = True
'        oForm01.Items("DCardCod").Enabled = True
'        oForm01.Items("BPLId").Enabled = True
'        oForm01.Items("DocType").Enabled = True
'        oForm01.Items("CntcCode").Enabled = True
'        oForm01.Items("DocDate").Enabled = True
'        oForm01.Items("DueDate").Enabled = True
'        oForm01.Items("Comments").Enabled = True
'        oForm01.Items("TrType").Enabled = True
'        oForm01.Items("Opt01").Enabled = True
'        oForm01.Items("Opt02").Enabled = True
'        oForm01.Items("Mat01").Enabled = False
'        oMat01.AutoResizeColumns
'        Call oForm01.EnableMenu("1281", False) '//찾기
'        Call oForm01.EnableMenu("1282", True) '//추가
'        oForm01.Items("DocDate").Enabled = True
'        oForm01.Items("DueDate").Enabled = True
'        oMat01.Columns("SD030Num").Visible = True
'        oMat01.Columns("PackNo").Visible = True
'        oMat01.Columns("OrderNum").Visible = True
'
'        '//각모드에따른 아이템설정
'    ElseIf (oForm01.Mode = fm_OK_MODE) Then
'        Call oForm01.EnableMenu("1281", True) '//찾기
'        Call oForm01.EnableMenu("1282", True) '//추가
'        '//문서상태검사
'        oForm01.Items("DocEntry").Enabled = False
'        oForm01.Items("CardCode").Enabled = False
'        oForm01.Items("DCardCod").Enabled = False
'        oForm01.Items("BPLId").Enabled = False
'        oForm01.Items("DocType").Enabled = False
'        oForm01.Items("CntcCode").Enabled = False
'        oForm01.Items("DocDate").Enabled = False
'        oForm01.Items("DueDate").Enabled = False
'        oForm01.Items("Comments").Enabled = True
'        oForm01.Items("TrType").Enabled = False
'        oForm01.Items("Opt01").Enabled = False
'        oForm01.Items("Opt02").Enabled = False
'        oForm01.Items("Mat01").Enabled = False
'        oMat01.AutoResizeColumns
'        oForm01.Items("DocDate").Enabled = False
'        oForm01.Items("DueDate").Enabled = False
'        oMat01.Columns("SD030Num").Visible = True
'        oMat01.Columns("PackNo").Visible = True
'        oMat01.Columns("OrderNum").Visible = True
'        '//각모드에따른 아이템설정
'
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP095_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP095_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP095_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP095_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP095L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP095L.Offset = oRow
    oDS_PS_PP095L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP095_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP095_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP095_FormClear()
On Error GoTo PS_PP095_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP095'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP095_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP095_EnableMenus()
On Error GoTo PS_PP095_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1293", True)
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, True, True, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP095_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP095_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP095_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP095_FormItemEnabled
        Call PS_PP095_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_PP095_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP095_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Function PS_PP095_DataValidCheck() As Boolean
On Error GoTo PS_PP095_DataValidCheck_Error
    PS_PP095_DataValidCheck = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim OrdNum As String

    If oForm01.Items("CardCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "거래처코드는 필수입니다.", bmt_Short, True
        oForm01.Items("CardCode").Click ct_Regular
        PS_PP095_DataValidCheck = False
        Exit Function
    End If
    
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "전기일은 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_PP095_DataValidCheck = False
        Exit Function
    End If
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP095_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
    
        
        If (oMat01.Columns("OrdNum").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "LOT-NO는 필수입니다.", bmt_Short, True
            oMat01.Columns("OrdNum").Cells(i).Click ct_Regular
            PS_PP095_DataValidCheck = False
            Exit Function
        Else
            If (oForm01.Mode = fm_ADD_MODE) Then
                OrdNum = oMat01.Columns("OrdNum").Cells(i).Specific.VALUE
                Query01 = "select Cnt = Count(U_OrdNum) from [@PS_PP095H] a inner Join [@PS_PP095L] b On a.DocEntry = b.DocEntry and a.Canceled = 'N' "
                Query01 = Query01 + " Where b.U_OrdNum = '" & OrdNum & "'"
                Call RecordSet01.DoQuery(Query01)
                
                If RecordSet01.Fields("Cnt").VALUE > 0 Then
                    Sbo_Application.SetStatusBarMessage OrdNum & " : 등록된 Lot번호입니다.", bmt_Short, True
                    PS_PP095_DataValidCheck = False
                    Exit Function
                End If
            End If
        End If
        
    Next
       
    Call oDS_PS_PP095L.RemoveRecord(oDS_PS_PP095L.Size - 1)
    Call oMat01.LoadFromDataSource
    If oForm01.Mode = fm_ADD_MODE Then
        Call PS_PP095_FormClear
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_PP095_DataValidCheck_Error:
    PS_PP095_DataValidCheck = False
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_PP095_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP095_Print_Report01()
On Error GoTo PS_PP095_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim i As Long
    Dim sQry01                      As String
    Dim Comments                    As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
'        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
'        Exit Sub
'    End If

    Call ConnectODBC
    
    WinTitle = "[PS_PP095] 거래명세서출력"
    
    If oForm01.Items("BPLId").Specific.Selected.VALUE = "2" Then
        ReportName = "PS_PP095_05.rpt" '부산사업장용
    Else
        ReportName = "PS_PP095_01.rpt"
    End If

    sQry = "EXEC PS_PP095_01 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_PP095_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP095_Print_Report02()
On Error GoTo PS_PP095_Print_Report02_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim i As Long
    Dim sQry01                      As String
    Dim Comments                    As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
'        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
'        Exit Sub
'    End If

    Call ConnectODBC
    
    WinTitle = "[PS_PP095_02] 납품명세서출력"
    ReportName = "PS_PP095_02.rpt"
    sQry = "EXEC PS_PP095_02 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_PP095_Print_Report02_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP095_Print_Report02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP095_Print_Report03()
On Error GoTo PS_PP095_Print_Report03_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim i As Long
    Dim Comments                    As String
    Dim sQry01                      As String
    Dim sQry02                      As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim FilePath As String
    FilePath = "\\191.1.1.220\ERP_Project\img\QRCODE" '\\191.1.1.138\b1_shr\QRCODE
    
'    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
'        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
'        Exit Sub
'    End If

    Call ConnectODBC
    
    
    '//QR코드 생성 및 이미지를 테이블에 Insert 2015/11/30 N.G.Y 추가
    'sQry01 = "EXEC PS_PP095_99 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    'Call oRecordSet01.DoQuery(sQry01)
    
'    For i = 0 To oRecordSet01.RecordCount - 1
'        With QREncode.QREncode(New_c.Crypt.VBStringToUTF8(oRecordSet01.Fields(1).VALUE)) 'ImageSurface is the default
'            .WriteContentToJpgFile FilePath & "\" & oRecordSet01.Fields(0).VALUE & ".jpg", 95
'        End With
'
'        sQry02 = "Insert Into ZPS_PP095_QRCODE(DocEntry, PackNo, QRImg) "
'        sQry02 = sQry02 & " Select '" & oForm01.Items("DocEntry").Specific.VALUE & "','" & oRecordSet01.Fields(0).VALUE & "', *"
'        sQry02 = sQry02 & " From OPENROWSET(BULK N'" & FilePath & "\" & oRecordSet01.Fields(0).VALUE & ",jpg', SINGLE_BLOB) As QRImg"
'        Call oRecordSet02.DoQuery(sQry02)
'
'        oRecordSet01.MoveNext
'
'    Next
    
    WinTitle = "[PS_PP095_03] PACKING라벨출력"
    ReportName = "PS_PP095_03.rpt"
    sQry = "EXEC PS_PP095_03 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        'Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
        GoTo PS_PP095_Print_Report03_Error
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Exit Sub
    
PS_PP095_Print_Report03_Error:
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "PS_PP095_Print_Report03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP095_Print_Report04()
On Error GoTo PS_PP095_Print_Report04_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim i As Long
    Dim Comments                    As String
    Dim sQry01                      As String
    Dim sQry02                      As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim oRecordset02 As SAPbobsCOM.Recordset
    Set oRecordset02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim FilePath As String
    Dim j As Long
    Dim Sum As Long
    
'    FilePath = "C:\QRCODEIMG"
    FilePath = "\\191.1.1.220\B1_SHR\QRCODE_PACKING" ' \\191.1.1.138\B1_SHR\QRCODE_PACKING
    
'    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
'        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
'        Exit Sub
'    End If

    Call ConnectODBC
    
    
    '//QR코드 생성 및 이미지를 테이블에 Insert 2015/11/30 N.G.Y 추가
    sQry01 = "EXEC PS_PP095_99 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    Call oRecordSet01.DoQuery(sQry01)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        With QREncode.QREncode(New_c.Crypt.VBStringToUTF8(oRecordSet01.Fields(1).VALUE)) 'ImageSurface is the default
            .WriteContentToJpgFile FilePath & "\" & oRecordSet01.Fields(0).VALUE & ".jpg", 95
        End With
        
'        Application.Wait Now() + TimeSerial(0, 0, 1)
        
        sQry02 = "Insert Into ZPS_PP095_QRCODE(DocEntry, PackNo) "
        sQry02 = sQry02 & " Select '" & oForm01.Items("DocEntry").Specific.VALUE & "','" & oRecordSet01.Fields(0).VALUE & "'"
        Call oRecordset02.DoQuery(sQry02)

        sQry02 = "Update ZPS_PP095_QRCODE Set QRImg = (Select Bulkcolumn From OPENROWSET(BULK N'" & FilePath & "\" & oRecordSet01.Fields(0).VALUE & ".jpg', SINGLE_BLOB) As QRImg)"
        sQry02 = sQry02 + " Where DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "' And PackNo = '" & oRecordSet01.Fields(0).VALUE & "'"
        Call oRecordset02.DoQuery(sQry02)
        
        oRecordSet01.MoveNext

    Next
    
'    sQry01 = "EXEC PS_PP095_99 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
'    Call oRecordSet01.DoQuery(sQry01)
    
'    For i = 0 To oRecordSet01.RecordCount - 1
''        With QREncode.QREncode(New_c.Crypt.VBStringToUTF8(oRecordSet01.Fields(1).VALUE)) 'ImageSurface is the default
''            .WriteContentToJpgFile FilePath & "\" & oRecordSet01.Fields(0).VALUE & ".jpg", 95
''        End With
''
''        Application.Wait Now() + TimeSerial(0, 0, 1)
'
'        sQry02 = "Insert Into ZPS_PP095_QRCODE(DocEntry, PackNo, QRImg) "
'        sQry02 = sQry02 & " Select '" & oForm01.Items("DocEntry").Specific.VALUE & "','" & oRecordSet01.Fields(0).VALUE & "', *"
'        sQry02 = sQry02 & " From OPENROWSET(BULK N'" & FilePath & "\" & oRecordSet01.Fields(0).VALUE & ".jpg', SINGLE_BLOB) As QRImg"
'        Call oRecordset02.DoQuery(sQry02)
'
'        oRecordSet01.MoveNext
'
'    Next
    
    WinTitle = "[PS_PP095_03_QRCODE] PACKING라벨(QRCODE)출력"
    ReportName = "PS_PP095_03_QRCODE.rpt"
    sQry = "EXEC PS_PP095_03 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        'Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
        GoTo PS_PP095_Print_Report04_Error
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Exit Sub
    
PS_PP095_Print_Report04_Error:
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "PS_PP095_Print_Report04_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP095_Check_QRCode_PrintYN() As String
'******************************************************************************
'Function ID : PS_PP095_Check_QRCode_PrintYN()
'해당모듈 : PS_PP095
'기능 : QR Code를 출력해야할 고객인지 체크
'인수 : 없음
'반환값 : "Y":QR Code 출력 함, "N":QR Code 출력 안함
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP095_Check_QRCode_PrintYN_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "                EXEC PS_PP095_10 "
    sQry = sQry & "'" & oForm01.Items("CardCode").Specific.VALUE & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    PS_PP095_Check_QRCode_PrintYN = Trim(oRecordSet01.Fields(0).VALUE)
    Set oRecordSet01 = Nothing

    Exit Function

PS_PP095_Check_QRCode_PrintYN_Error:
    Set oRecordSet01 = Nothing
    PS_PP095_Check_QRCode_PrintYN = "E"
    Call Sbo_Application.SetStatusBarMessage("PS_PP095_Check_QRCode_PrintYN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function
Private Function PS_PP095_Weight_Check() As Boolean
'******************************************************************************
'Function ID : PS_PP095_Check_QRCode_PrintYN()
'해당모듈 : PS_PP095
'기능 : QR Code를 출력해야할 고객인지 체크
'인수 : 없음
'반환값 : "Y":QR Code 출력 함, "N":QR Code 출력 안함
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP095_Weight_Check_Error

    Dim sQry As String
    Dim i&, BoxCnt&
    Dim CardCode As String
    Dim ErrorName As String
    Dim DOEX As String
    Dim ItemType As String
    Dim PackNo As String
    Dim PackNo2 As String
    Dim ItemCode As String
    Dim result As Integer
    Dim result2 As Integer
    Dim U_MINRANGE As Integer
    Dim ErrorCode As String
    Dim TCHECK As String
    Dim ICount As Integer
    Dim maxrange As Integer
    Dim minrange As Integer
    Dim T_Weight, S_Weight As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    T_Weight = Trim(oDS_PS_PP095H.GetValue("U_Tweight", 0))
    TCHECK = Trim(oDS_PS_PP095H.GetValue("U_TCHECK", 0))

    PackNo2 = 0


    For i = 0 To oMat01.VisualRowCount - 1

        PackNo = Trim(oDS_PS_PP095L.GetValue("U_PackNo", i))

        If i = (oMat01.VisualRowCount - 1) Then
            PackNo2 = "123"
        Else

            PackNo2 = Trim(oDS_PS_PP095L.GetValue("U_PackNo", i + 1))
        End If

        If PackNo <> PackNo2 Then

                sQry = " EXEC [PS_PP095_13] "
                sQry = sQry & "'" & PackNo & "'"

                Call oRecordSet01.DoQuery(sQry)

                result = oRecordSet01.Fields("RESULT").VALUE

                S_Weight = S_Weight + result

        End If
    Next i

                sQry = "SELECT U_MINRANGE FROM [@PSH_MULTI_PALLET] WHERE CODE ='STANDARD'"

                oRecordSet01.DoQuery sQry

                result2 = Trim(oRecordSet01.Fields(0).VALUE)

                maxrange = S_Weight + result2
                minrange = S_Weight - result2



     If TCHECK = 1 Then

        If T_Weight > maxrange Or T_Weight < minrange Then       ' 최대값과 최소값 비교하여 참인지 거짓인지 비교

        ErrorCode = "1"
        GoTo PS_PP095_Weight_Check_Error

        End If
     End If
     PS_PP095_Weight_Check = True
          
    Exit Function

PS_PP095_Weight_Check_Error:

    Call PS_PP095_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식

    If ErrorCode = "1" Then
        Set oRecordSet01 = Nothing
        
        Call MDC_Com.MDC_GF_Message("계근 중량을 한번더 확인하십시오.", "E")
        PS_PP095_Weight_Check = False
    
    End If
    If ErrorCode = "2" Then
        Set oRecordSet01 = Nothing
        
        Call MDC_Com.MDC_GF_Message("계근 유무를 확인하십시오.", "E")

        PS_PP095_Weight_Check = False
    End If
End Function
