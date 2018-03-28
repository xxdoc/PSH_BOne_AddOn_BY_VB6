VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END

Attribute VB_Name = "PS_SD040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_SD040.cls
'//  Module         : 판매관리 > 납품처리
'//  Desc           : 납품
'//  FormType       : PS_SD040
'//  Create Date    : 2010.08.20
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_SD040H  As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_SD040L  As SAPbouiCOM.DBDataSource '등록라인

Private oDocType01     As String

Public oLastItemUID01  As String '클래스에서 선택한 마지막 아이템 Uid값
Public oLastColUID01   As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Public oLastColRow01   As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode  As String
    BatchNum  As String
    Qty       As Long '//수량
    Weight    As Double '//중량
    Currency  As String '//통화
    Price     As Double '//단가
    LineTotal As Double '//총계
    WhsCode   As String '//창고
    SD030HNum As Long '//출하(선출)문서
    SD030LNum As Long '//출하(선출)라인
    SD040HNum As Long '//납품A문서
    SD040LNum As Long '//납품A라인
    ORDRNum   As Long '//판매오더문서
    RDR1Num   As Long '//판매오더라인
    Check     As Boolean
    ODLNNum   As Long '//납품문서
    DLN1Num   As Long '//납품라인
    ORDNNum   As Long '//반품문서
    RDN1Num   As Long '//반품라인
End Type
Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

Private Type BatchInformations
    ItemCode As String '//품목코드
    WhsCode  As String '//창고코드
    BatchNum As String '//배치번호
    Weight   As Double '//중량
End Type
Private BatchInformation() As BatchInformations
Private BatchInformationCount As Long

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String, Optional ByVal oFromDocType01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD040.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_SD040_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    oDocType01 = oFromDocType01
    
    Call PS_SD040_CreateItems
    Call PS_SD040_ComboBox_Setting
    Call PS_SD040_CF_ChooseFromList
    Call PS_SD040_EnableMenus
    Call PS_SD040_SetDocument(oFromDocEntry01)
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
    Dim sQry As String
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
                '//납품취소, 반품
                If oForm01.Mode = fm_OK_MODE Then
                    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_SD040H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
                        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If oForm01.Items("Opt02").Specific.Selected = True Then '//멀티이면서
                        For i = 1 To oMat01.VisualRowCount - 1 '//멀티게이지 재고수량검사
                            sQry = "Select Quantity = Sum(a.Quantity) From OIBT a Inner Join OITM b On a.ItemCode = b.ItemCode "
                            sQry = sQry + " Where b.ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & "'"
                            sQry = sQry + " And a.BatchNum = '" & oMat01.Columns("LotNo").Cells(i).Specific.VALUE & "'"
                            sQry = sQry + " AND a.WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.VALUE & "'"
                            If MDC_PS_Common.GetValue(sQry, 0, 1) > 0 Then
                            'If MDC_PS_Common.GetValue("SELECT Quantity FROM [OIBT] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & "' AND BatchNum = '" & oMat01.Columns("LotNo").Cells(i).Specific.VALUE & "' AND WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
                            
                            'If MDC_PS_Common.GetValue("SELECT Sum(Quantity * (Case When Direction = '0' Then 1 Else -1 End))  FROM [IBT1] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & "' AND BatchNum = '" & oMat01.Columns("LotNo").Cells(i).Specific.VALUE & "' AND WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
                                Call MDC_Com.MDC_GF_Message("멀티게이지 품목 : " & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & " 의 재고가 존재합니다.", "W")
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Next
                    End If
                    If oForm01.Items("Opt03").Specific.Selected = True Then '//분말이면서
                        For i = 1 To oMat01.VisualRowCount - 1 '//분말 재고수량검사
                            If MDC_PS_Common.GetValue("SELECT Quantity FROM [OIBT] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & "' AND BatchNum = '" & oMat01.Columns("LotNo").Cells(i).Specific.VALUE & "' AND WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
                                Call MDC_Com.MDC_GF_Message("분말 품목 : " & oMat01.Columns("ItemCode").Cells(i).Specific.VALUE & " 의 재고가 존재합니다.", "W")
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Next
                    End If
                    For i = 1 To oMat01.VisualRowCount - 1 '//AR송장처리된 문서존재유무검사
                        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [INV1] WHERE BaseType = '15' AND BaseEntry = '" & oMat01.Columns("ODLNNum").Cells(i).Specific.VALUE & "' AND BaseLine = '" & oMat01.Columns("DLN1Num").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
                            Call MDC_Com.MDC_GF_Message("AR송장처리된 문서가 존재합니다.", "W")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                    For i = 1 To oMat01.VisualRowCount - 1 '//반품처리된 문서존재유무검사
                        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [RDN1] WHERE BaseType = '15' AND BaseEntry = '" & oMat01.Columns("ODLNNum").Cells(i).Specific.VALUE & "' AND BaseLine = '" & oMat01.Columns("DLN1Num").Cells(i).Specific.VALUE & "'", 0, 1) > 0 Then
                            Call MDC_Com.MDC_GF_Message("반품처리된 문서가 존재합니다.", "W")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                    If PS_SD040_DI_API_03 = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                Else
                    MDC_Com.MDC_GF_Message "현재 모드에서는 취소할수 없습니다.", "W"
                    BubbleEvent = False
                    Exit Sub
                End If
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
                Call PS_SD040_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_SD040_FormItemEnabled '//UDO방식
                Call PS_SD040_AddMatrixRow(0, True) '//UDO방식
                '//2010.12.06 추가
                '담당자
                oDS_PS_SD040H.setValue "U_CntcCode", 0, User_MSTCOD()
                Call oDS_PS_SD040H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items("CntcCode").Specific.VALUE & "'", 0, 1))
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_SD040_FormItemEnabled '//UDO방식
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
    If pval.BeforeAction = True Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_SD040_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
'                납품처리 문서 단독으로 입력할 경우 주석처리 ("납품"문서 생성 안함)_S
                If oForm01.Items("Opt01").Specific.Selected = True Then
                    If PS_SD040_DI_API_01 = False Then '//납품생성
                        Call PS_SD040_AddMatrixRow(oMat01.RowCount) '//UDO방식일때
                        BubbleEvent = False
                        Exit Sub
                    End If
                ElseIf oForm01.Items("Opt02").Specific.Selected = True Then
                    If PS_SD040_DI_API_02 = False Then '//멀티게이지납품생성
                        Call PS_SD040_AddMatrixRow(oMat01.RowCount) '//UDO방식일때
                        BubbleEvent = False
                        Exit Sub
                    End If
                ElseIf oForm01.Items("Opt03").Specific.Selected = True Then
                    If PS_SD040_DI_API_01 = False Then '//분말 납품생성
                        Call PS_SD040_AddMatrixRow(oMat01.RowCount) '//UDO방식일때
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
'                납품처리 문서 단독으로 입력할 경우 주석처리 ("납품"문서 생성 안함)_E
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_SD040_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '----------------------------------------
                '출고가 생성안되었을때
'                If oForm01.Items("Opt01").Specific.Selected = True Then
'                    If PS_SD040_DI_API_01 = False Then '//납품생성
'                        Call PS_SD040_AddMatrixRow(oMat01.RowCount) '//UDO방식일때
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
'                End If
'
'                If oForm01.Items("Opt02").Specific.Selected = True Then
'                    If PS_SD040_DI_API_02 = False Then '//멀티게이지납품생성
'                        Call PS_SD040_AddMatrixRow(oMat01.RowCount) '//UDO방식일때
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
'                End If
                '-----------------------------------------
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_SD040_Print_Report01 '거래명세서
            End If
        End If
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                   If oForm01.Items("Opt03").Specific.Selected = True Then  '분말일경우
                      Call PS_SD040_Print_Report03("N") '거래명세표(수량)
                   Else
                      Call PS_SD040_Print_Report02("N") '출고증
                   End If
            End If
        End If
        If pval.ItemUID = "Button02_1" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                   If oForm01.Items("Opt03").Specific.Selected = True Then  '분말일경우
                      Call PS_SD040_Print_Report03_1("N") '거래명세표(금액)
                   End If
            End If
        End If
        If pval.ItemUID = "Button03" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                
                Dim TempForm01 As Variant
                Set TempForm01 = New PS_MM004
                Call TempForm01.LoadForm("PS_SD040", oForm01.Items("DocEntry").Specific.VALUE)
                BubbleEvent = False
'
                
'                Call PS_SD040_Print_Report02("Y") '출고증(운송장 포함)
            End If
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_SD040_FormItemEnabled
                    '//Call PS_SD040_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
                    oForm01.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1291")
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_SD040_FormItemEnabled
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
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "SD030Num" Then
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    Call MDC_Com.MDC_GF_Message("고객코드는 필수입니다.", "W")
                    BubbleEvent = False
                    Exit Sub
                End If
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "SD030Num") '//사용자값활성
            End If
        End If
        '//Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "OrderNum") '//사용자값활성
        '//Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "WhsCode") '//사용자값활성
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "PackNo" Then
                If oForm01.Items("Opt02").Specific.Selected = True Then
                    If oForm01.Items("TrType").Specific.Selected.VALUE <> "2" Then
                        Call MDC_Com.MDC_GF_Message("포장번호조회는 거래형태가 임가공이여야 합니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    If oForm01.Items("DCardCod").Specific.VALUE = "" Then
                        Call MDC_Com.MDC_GF_Message("납품처코드는 필수입니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    Call MDC_Com.MDC_GF_Message("고객코드는 필수입니다.", "W")
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "PackNo") '//사용자값활성
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
        If pval.ItemUID = "Opt01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                oForm01.Freeze True
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("CardCode").Enabled = True
                oForm01.Items("BPLId").Enabled = True
                oForm01.Items("DCardCod").Enabled = True
                oForm01.Items("TrType").Enabled = True
                oMat01.Columns("SD030Num").Visible = True
                oMat01.Columns("PackNo").Visible = False
                oMat01.Columns("OrderNum").Visible = True
                oMat01.Columns("SD030Num").Editable = True
                oMat01.Columns("Qty").Editable = True
                oMat01.Columns("Weight").Editable = True
                oMat01.Columns("UnWeight").Visible = True
                oMat01.Columns("Price").Editable = False
                Call oForm01.Items("TrType").Specific.Select("1", psk_ByValue)
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Call PS_SD040_AddMatrixRow(0, True)
                
                oForm01.Items("SumSjQty").Specific.VALUE = 0
                oForm01.Items("SumSjWt").Specific.VALUE = 0
                oForm01.Items("SumQty").Specific.VALUE = 0
                oForm01.Items("SumWeight").Specific.VALUE = 0
                oForm01.Items("HWeight").Specific.VALUE = 0
                oForm01.Freeze False
            End If
        End If
        If pval.ItemUID = "Opt02" Then '//멀티게이지의경우
            If oForm01.Mode = fm_ADD_MODE Then
                oForm01.Freeze True
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("CardCode").Enabled = True
                oForm01.Items("BPLId").Enabled = True
                oForm01.Items("DCardCod").Enabled = True
                oForm01.Items("TrType").Enabled = False
                oMat01.Columns("SD030Num").Visible = False
                oMat01.Columns("PackNo").Visible = True
                oMat01.Columns("OrderNum").Visible = False
                oMat01.Columns("Qty").Editable = False
                oMat01.Columns("Weight").Editable = False
                oMat01.Columns("UnWeight").Visible = False
                oMat01.Columns("Price").Editable = True
                Call oForm01.Items("TrType").Specific.Select("2", psk_ByValue)
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Call PS_SD040_AddMatrixRow(0, True)
                
                oForm01.Items("SumSjQty").Specific.VALUE = 0
                oForm01.Items("SumSjWt").Specific.VALUE = 0
                oForm01.Items("SumQty").Specific.VALUE = 0
                oForm01.Items("SumWeight").Specific.VALUE = 0
                oForm01.Items("HWeight").Specific.VALUE = 0
                oForm01.Freeze False
            End If
        End If
            
        If pval.ItemUID = "Opt03" Then '분말
            If oForm01.Mode = fm_ADD_MODE Then
                oForm01.Freeze True
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("CardCode").Enabled = True
                oForm01.Items("BPLId").Enabled = True
                oForm01.Items("DCardCod").Enabled = True
                oForm01.Items("TrType").Enabled = False
                oMat01.Columns("SD030Num").Visible = True
                oMat01.Columns("PackNo").Visible = True
                oMat01.Columns("OrderNum").Visible = True
                oMat01.Columns("Qty").Editable = False
                oMat01.Columns("Weight").Editable = False
                oMat01.Columns("SD030Num").Editable = False
                oMat01.Columns("UnWeight").Visible = False
                oMat01.Columns("Price").Editable = False
                Call oForm01.Items("TrType").Specific.Select("1", psk_ByValue)
                oMat01.Clear
                oMat01.FlushToDataSource
                oMat01.LoadFromDataSource
                Call PS_SD040_AddMatrixRow(0, True)
                
                oForm01.Items("SumSjQty").Specific.VALUE = 0
                oForm01.Items("SumSjWt").Specific.VALUE = 0
                oForm01.Items("SumQty").Specific.VALUE = 0
                oForm01.Items("SumWeight").Specific.VALUE = 0
                oForm01.Items("HWeight").Specific.VALUE = 0
                oForm01.Freeze False
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
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "SD030Num" Then
                Set oTempClass = New PS_SD030
                Call oTempClass.LoadForm(Mid(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, 1, InStr(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE, "-") - 1))
            End If
            If pval.ColUID = "SD030H" Then
                Set oTempClass = New PS_SD030
                Call oTempClass.LoadForm(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
            End If
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
    Dim Work01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Dim RecordSet03 As SAPbobsCOM.Recordset
    Dim ItemCode01 As String
    Dim sQry As String
    Dim SumSjQty As Long, SumSjWt As Currency, SumQty As Long, SumWeight As Currency, HWeight As Currency
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If pval.ColUID = "SD030Num" Then
                    If oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE = "" Then
                        GoTo Raise_EVENT_VALIDATE_Exit
                    End If
                    If oForm01.Items("Opt03").Specific.Selected <> True Then
                        For i = 1 To oMat01.RowCount
                            If pval.Row <> i Then '//현재 선택되어있는 행이 아니면
                                If (oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE = oMat01.Columns("SD030Num").Cells(i).Specific.VALUE) Then
                                    Call MDC_Com.MDC_GF_Message("동일한 출하(선출)요청이 존재합니다.", "W")
                                    oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE = ""
                                    GoTo Raise_EVENT_VALIDATE_Exit
                                End If
                                If (Mid(oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE, 1, InStr(oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE, "-") - 1) <> _
                                Mid(oMat01.Columns("SD030Num").Cells(i).Specific.VALUE, 1, InStr(oMat01.Columns("SD030Num").Cells(i).Specific.VALUE, "-") - 1)) Then
                                    If Trim(oForm01.Items("BPLId").Specific.VALUE) = "4" Then '[2011.1.10추가]구로사업소일경우에만 동일하지 않은 출하요청문서 막음
                                        Call MDC_Com.MDC_GF_Message("동일하지않은 출하요청문서가 존재합니다.", "W")
                                        oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE = ""
                                        GoTo Raise_EVENT_VALIDATE_Exit
                                    End If
                                End If
                            End If
                        Next
                    End If
                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    Query01 = "EXEC PS_SD040_01 '" & oMat01.Columns("SD030Num").Cells(pval.Row).Specific.VALUE & "'"
                    Call RecordSet01.DoQuery(Query01)
                    For i = 0 To RecordSet01.RecordCount - 1
                        Call oDS_PS_SD040L.setValue("U_SD030Num", pval.Row - 1, RecordSet01.Fields("SD030Num").VALUE)
                        Call oDS_PS_SD040L.setValue("U_OrderNum", pval.Row - 1, RecordSet01.Fields("OrderNum").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItemGpCd", pval.Row - 1, RecordSet01.Fields("ItemGpCd").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItmBsort", pval.Row - 1, RecordSet01.Fields("ItmBsort").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItmMsort", pval.Row - 1, RecordSet01.Fields("ItmMsort").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Unit1", pval.Row - 1, RecordSet01.Fields("Unit1").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Size", pval.Row - 1, RecordSet01.Fields("Size").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItemType", pval.Row - 1, RecordSet01.Fields("ItemType").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Quality", pval.Row - 1, RecordSet01.Fields("Quality").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Mark", pval.Row - 1, RecordSet01.Fields("Mark").VALUE)
                        Call oDS_PS_SD040L.setValue("U_SbasUnit", pval.Row - 1, RecordSet01.Fields("SbasUnit").VALUE)
                        Call oDS_PS_SD040L.setValue("U_LotNo", pval.Row - 1, RecordSet01.Fields("LotNo").VALUE)
                        Call oDS_PS_SD040L.setValue("U_SjQty", pval.Row - 1, Val(RecordSet01.Fields("SjQty").VALUE))
                        Call oDS_PS_SD040L.setValue("U_SjWeight", pval.Row - 1, Val(RecordSet01.Fields("SjWeight").VALUE))
                        Call oDS_PS_SD040L.setValue("U_Qty", pval.Row - 1, Val(RecordSet01.Fields("Qty").VALUE))
                        Call oDS_PS_SD040L.setValue("U_UnWeight", pval.Row - 1, Val(RecordSet01.Fields("UnWeight").VALUE))
                        Call oDS_PS_SD040L.setValue("U_Weight", pval.Row - 1, Val(RecordSet01.Fields("Weight").VALUE))
                        Call oDS_PS_SD040L.setValue("U_Currency", pval.Row - 1, RecordSet01.Fields("Currency").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Price", pval.Row - 1, Val(RecordSet01.Fields("Price").VALUE))
                        Call oDS_PS_SD040L.setValue("U_LinTotal", pval.Row - 1, Val(RecordSet01.Fields("LinTotal").VALUE))
                        Call oDS_PS_SD040L.setValue("U_WhsCode", pval.Row - 1, RecordSet01.Fields("WhsCode").VALUE)
                        Call oDS_PS_SD040L.setValue("U_WhsName", pval.Row - 1, RecordSet01.Fields("WhsName").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Comments", pval.Row - 1, RecordSet01.Fields("Comments").VALUE)
                        Call oDS_PS_SD040L.setValue("U_SD030H", pval.Row - 1, RecordSet01.Fields("SD030H").VALUE)
                        Call oDS_PS_SD040L.setValue("U_SD030L", pval.Row - 1, RecordSet01.Fields("SD030L").VALUE)
                        Call oDS_PS_SD040L.setValue("U_TrType", pval.Row - 1, RecordSet01.Fields("TrType").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ORDRNum", pval.Row - 1, RecordSet01.Fields("ORDRNum").VALUE)
                        Call oDS_PS_SD040L.setValue("U_RDR1Num", pval.Row - 1, RecordSet01.Fields("RDR1Num").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Status", pval.Row - 1, RecordSet01.Fields("Status").VALUE)
                        Call oDS_PS_SD040L.setValue("U_LineId", pval.Row - 1, RecordSet01.Fields("LineId").VALUE)
                        RecordSet01.MoveNext
                    Next
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_SD040L.GetValue("U_SD030Num", pval.Row - 1)) <> "" Then
                        PS_SD040_AddMatrixRow (pval.Row)
                    End If
                    oMat01.LoadFromDataSource
                    oMat01.AutoResizeColumns
                    
                    '2011.1.4 추가(yjh) : 출하요청(SD030)의 비고 내용을 납품(SD040)에 넣어준다
                    If oMat01.VisualRowCount > 0 Then
                        Set RecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
                        sQry = "select U_Comments from [@PS_SD030H] where DocEntry = '" & Mid(oMat01.Columns("SD030Num").Cells(1).Specific.VALUE, 1, InStr(oMat01.Columns("SD030Num").Cells(1).Specific.VALUE, "-") - 1) & "'"
                        RecordSet03.DoQuery sQry
                        oForm01.Items("Comments").Specific.String = RecordSet03.Fields("U_Comments").VALUE
                    End If
                    '추가 end------------------------------------------------------------------
                    
                    For i = 0 To oMat01.VisualRowCount - 1
                        If oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE = "" Then
                            SumSjQty = SumSjQty
                        Else
                            SumSjQty = SumSjQty + oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE
                        End If
                        SumSjWt = SumSjWt + oMat01.Columns("SjWeight").Cells(i + 1).Specific.VALUE
                        If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                            SumQty = SumQty
                        Else
                            SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                        End If
                        SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                        HWeight = HWeight + (oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE * oMat01.Columns("UnWeight").Cells(i + 1).Specific.VALUE / 1000)
                    Next i
                    oForm01.Items("SumSjQty").Specific.VALUE = SumSjQty
                    oForm01.Items("SumSjWt").Specific.VALUE = SumSjWt
                    oForm01.Items("SumQty").Specific.VALUE = SumQty
                    oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                    oForm01.Items("HWeight").Specific.VALUE = HWeight
        
                    If (oMat01.RowCount > 1) Then
                        oForm01.Items("Focus").Click ct_Regular
                        oForm01.Items("CardCode").Enabled = False
                        oForm01.Items("BPLId").Enabled = False
                        oForm01.Items("DCardCod").Enabled = True
                        oForm01.Items("TrType").Enabled = False
                    Else
                        oForm01.Items("Focus").Click ct_Regular
                        oForm01.Items("CardCode").Enabled = True
                        oForm01.Items("BPLId").Enabled = True
                        oForm01.Items("DCardCod").Enabled = True
                        oForm01.Items("TrType").Enabled = True
                    End If
                    oForm01.Update
                    Set RecordSet01 = Nothing
                ElseIf pval.ColUID = "PackNo" Then
                    If oMat01.Columns("PackNo").Cells(pval.Row).Specific.VALUE = "" Then
                        GoTo Raise_EVENT_VALIDATE_Exit
                    End If
                    For i = 1 To oMat01.RowCount
                        If pval.Row <> i Then '//현재 선택되어있는 행이 아니면
                            If (oMat01.Columns("PackNo").Cells(pval.Row).Specific.VALUE = oMat01.Columns("PackNo").Cells(i).Specific.VALUE) Then
                                Call MDC_Com.MDC_GF_Message("동일한 포장번호가 존재합니다.", "W")
                                oMat01.Columns("PackNo").Cells(pval.Row).Specific.VALUE = ""
                                GoTo Raise_EVENT_VALIDATE_Exit
                            End If
                        End If
                    Next
                    
                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    If oForm01.Items("Opt02").Specific.Selected = True Then
                        Work01 = "1"
                    ElseIf oForm01.Items("Opt03").Specific.Selected = True Then
                        Work01 = "3"
                    End If
                    Query01 = "EXEC PS_SD040_06 '" & oMat01.Columns("PackNo").Cells(pval.Row).Specific.VALUE & "', '" & Work01 & "', '" & Trim(oForm01.Items("CardCode").Specific.String) & "'"
                    Call RecordSet01.DoQuery(Query01)
                    If RecordSet01.RecordCount <= 0 Then
                        Call MDC_Com.MDC_GF_Message("포장번호정보가 존재하지 않습니다.", "W")
                        oMat01.Columns("PackNo").Cells(pval.Row).Specific.VALUE = ""
                        GoTo Raise_EVENT_VALIDATE_Exit
                    Else
                        '//해당 포장번호로 재고유무확인
                        Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
                        Query01 = "EXEC PS_SD040_07 '" & oMat01.Columns("PackNo").Cells(pval.Row).Specific.VALUE & "'"
                        Call RecordSet02.DoQuery(Query01)
                        If RecordSet02.Fields(0).VALUE = "Enabled" Then
                            '//진행가능
                        ElseIf RecordSet02.Fields(0).VALUE = "Disabled" Then
                            Call MDC_Com.MDC_GF_Message("해당포장번호의 재고가 부족합니다.", "W")
                            oMat01.Columns("PackNo").Cells(pval.Row).Specific.VALUE = ""
                            GoTo Raise_EVENT_VALIDATE_Exit
                        End If
                    End If
                    For i = 0 To RecordSet01.RecordCount - 1
                        Call oDS_PS_SD040L.setValue("U_PackNo", pval.Row - 1 + i, RecordSet01.Fields("PackNo").VALUE)
                        If Work01 = "3" Then
                            Call oDS_PS_SD040L.setValue("U_SD030Num", pval.Row - 1 + i, RecordSet01.Fields("SD030Num").VALUE)
                            Call oDS_PS_SD040L.setValue("U_OrderNum", pval.Row - 1 + i, RecordSet01.Fields("OrderNum").VALUE)
                            Call oDS_PS_SD040L.setValue("U_SD030H", pval.Row - 1 + i, RecordSet01.Fields("SD030H").VALUE)
                            Call oDS_PS_SD040L.setValue("U_SD030L", pval.Row - 1 + i, RecordSet01.Fields("SD030L").VALUE)
                            Call oDS_PS_SD040L.setValue("U_ORDRNum", pval.Row - 1 + i, RecordSet01.Fields("ORDRNum").VALUE)
                            Call oDS_PS_SD040L.setValue("U_RDR1Num", pval.Row - 1 + i, RecordSet01.Fields("RDR1Num").VALUE)
                        End If
                        Call oDS_PS_SD040L.setValue("U_ItemCode", pval.Row - 1 + i, RecordSet01.Fields("ItemCode").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItemName", pval.Row - 1 + i, RecordSet01.Fields("ItemName").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItemGpCd", pval.Row - 1 + i, RecordSet01.Fields("ItemGpCd").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItmBsort", pval.Row - 1 + i, RecordSet01.Fields("ItmBsort").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItmMsort", pval.Row - 1 + i, RecordSet01.Fields("ItmMsort").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Unit1", pval.Row - 1 + i, RecordSet01.Fields("Unit1").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Size", pval.Row - 1 + i, RecordSet01.Fields("Size").VALUE)
                        Call oDS_PS_SD040L.setValue("U_ItemType", pval.Row - 1 + i, RecordSet01.Fields("ItemType").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Quality", pval.Row - 1 + i, RecordSet01.Fields("Quality").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Mark", pval.Row - 1 + i, RecordSet01.Fields("Mark").VALUE)
                        Call oDS_PS_SD040L.setValue("U_SbasUnit", pval.Row - 1 + i, RecordSet01.Fields("SbasUnit").VALUE)
                        Call oDS_PS_SD040L.setValue("U_LotNo", pval.Row - 1 + i, RecordSet01.Fields("LotNo").VALUE)
                        Call oDS_PS_SD040L.setValue("U_SjQty", pval.Row - 1 + i, Val(RecordSet01.Fields("SjQty").VALUE))
                        Call oDS_PS_SD040L.setValue("U_SjWeight", pval.Row - 1 + i, Val(RecordSet01.Fields("SjWeight").VALUE))
                        Call oDS_PS_SD040L.setValue("U_Qty", pval.Row - 1 + i, Val(RecordSet01.Fields("Qty").VALUE))
'                        Call oDS_PS_SD040L.setValue("U_UnWeight", pval.Row - 1 + i, Val(RecordSet01.Fields("UnWeight").Value))
                        Call oDS_PS_SD040L.setValue("U_Weight", pval.Row - 1 + i, Val(RecordSet01.Fields("Weight").VALUE))
                        Call oDS_PS_SD040L.setValue("U_Currency", pval.Row - 1 + i, RecordSet01.Fields("Currency").VALUE)
                        Call oDS_PS_SD040L.setValue("U_Price", pval.Row - 1 + i, Val(RecordSet01.Fields("Price").VALUE))
                        Call oDS_PS_SD040L.setValue("U_LinTotal", pval.Row - 1 + i, Val(RecordSet01.Fields("LinTotal").VALUE))
                        Call oDS_PS_SD040L.setValue("U_WhsCode", pval.Row - 1 + i, RecordSet01.Fields("WhsCode").VALUE)
                        Call oDS_PS_SD040L.setValue("U_WhsName", pval.Row - 1 + i, RecordSet01.Fields("WhsName").VALUE)
'                        Call oDS_PS_SD040L.setValue("U_Comments", pval.Row - 1 + i, RecordSet01.Fields("Comments").Value)
                        Call oDS_PS_SD040L.setValue("U_Status", pval.Row - 1 + i, RecordSet01.Fields("Status").VALUE)
                        Call oDS_PS_SD040L.setValue("U_LineId", pval.Row - 1 + i, RecordSet01.Fields("LineId").VALUE)
                        PS_SD040_AddMatrixRow (pval.Row + i)
'                        If oMat01.RowCount = pval.Row And Trim(oDS_PS_SD040L.GetValue("U_PackNo", pval.Row - 1)) <> "" Then
'                            PS_SD040_AddMatrixRow (pval.Row)
'                        End If
'                        pval.Row = pval.Row + 1
                        RecordSet01.MoveNext
                    Next
                    
                    For i = 0 To oMat01.VisualRowCount - 1
                        If oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE = "" Then
                            SumSjQty = SumSjQty
                        Else
                            SumSjQty = SumSjQty + oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE
                        End If
                        SumSjWt = SumSjWt + oMat01.Columns("SjWeight").Cells(i + 1).Specific.VALUE
                        If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                            SumQty = SumQty
                        Else
                            SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                        End If
                        SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                        HWeight = HWeight + (oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE * oMat01.Columns("UnWeight").Cells(i + 1).Specific.VALUE / 1000)
                    Next i
                    oForm01.Items("SumSjQty").Specific.VALUE = SumSjQty
                    oForm01.Items("SumSjWt").Specific.VALUE = SumSjWt
                    oForm01.Items("SumQty").Specific.VALUE = SumQty
                    oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                    oForm01.Items("HWeight").Specific.VALUE = HWeight
                    
                    
                    If (oMat01.RowCount > 1) Then
                        oForm01.Items("Focus").Click ct_Regular
                        oForm01.Items("CardCode").Enabled = False
                        oForm01.Items("BPLId").Enabled = False
                        oForm01.Items("DCardCod").Enabled = False
                        oForm01.Items("TrType").Enabled = False
                    Else
                        oForm01.Items("Focus").Click ct_Regular
                        oForm01.Items("CardCode").Enabled = True
                        oForm01.Items("BPLId").Enabled = True
                        oForm01.Items("DCardCod").Enabled = True
                        oForm01.Items("TrType").Enabled = True
                    End If
                    oForm01.Update
                    Set RecordSet01 = Nothing
                    Set RecordSet02 = Nothing
                ElseIf pval.ColUID = "Qty" Then
                    If (Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0) Then
                        Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                        Call oDS_PS_SD040L.setValue("U_Weight", pval.Row - 1, 0)
                        Call oDS_PS_SD040L.setValue("U_LinTotal", pval.Row - 1, 0)
                    Else
                        ItemCode01 = Trim(oDS_PS_SD040L.GetValue("U_ItemCode", pval.Row - 1))
                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then '//EA자체품
                            Call oDS_PS_SD040L.setValue("U_Weight", pval.Row - 1, Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE))
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                            Call oDS_PS_SD040L.setValue("U_Weight", pval.Row - 1, Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_Unit1(ItemCode01)))
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                            Call oDS_PS_SD040L.setValue("U_Weight", pval.Row - 1, (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE))
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                            Call oDS_PS_SD040L.setValue("U_Weight", pval.Row - 1, Round(Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0))
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG선택
                        End If
                        Call oDS_PS_SD040L.setValue("U_LinTotal", pval.Row - 1, Trim(oDS_PS_SD040L.GetValue("U_Weight", pval.Row - 1)) * Trim(oDS_PS_SD040L.GetValue("U_Price", pval.Row - 1)))
                        Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    End If
                    
                    For i = 0 To oMat01.VisualRowCount - 1
                        If oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE = "" Then
                            SumSjQty = SumSjQty
                        Else
                            SumSjQty = SumSjQty + oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE
                        End If
                        SumSjWt = SumSjWt + oMat01.Columns("SjWeight").Cells(i + 1).Specific.VALUE
                        If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                            SumQty = SumQty
                        Else
                            SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                        End If
                        SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                        HWeight = HWeight + (oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE * oMat01.Columns("UnWeight").Cells(i + 1).Specific.VALUE / 1000)
                    Next i
                    oForm01.Items("SumSjQty").Specific.VALUE = SumSjQty
                    oForm01.Items("SumSjWt").Specific.VALUE = SumSjWt
                    oForm01.Items("SumQty").Specific.VALUE = SumQty
                    oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                    oForm01.Items("HWeight").Specific.VALUE = HWeight
        
                ElseIf pval.ColUID = "Weight" Then
                    If (Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0) Then
                        Call oDS_PS_SD040L.setValue("U_Qty", pval.Row - 1, 0)
                        Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                        Call oDS_PS_SD040L.setValue("U_LinTotal", pval.Row - 1, 0)
                    Else
                        ItemCode01 = Trim(oDS_PS_SD040L.GetValue("U_ItemCode", pval.Row - 1))
                        If (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "101") Then
                            Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "102") Then '//EAUOM
                            Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "201") Then '//KGSPEC
                            Call oDS_PS_SD040L.setValue("U_Weight", pval.Row - 1, (Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE))
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "202") Then '//KG단중
                            Call oDS_PS_SD040L.setValue("U_Weight", pval.Row - 1, Round(Val(oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE) * Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0))
                        ElseIf (MDC_PS_Common.GetItem_SbasUnit(ItemCode01) = "203") Then '//KG선택
                            Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        End If
                        Call oDS_PS_SD040L.setValue("U_LinTotal", pval.Row - 1, Trim(oDS_PS_SD040L.GetValue("U_Weight", pval.Row - 1)) * Trim(oDS_PS_SD040L.GetValue("U_Price", pval.Row - 1)))
                    End If
                    
                    For i = 0 To oMat01.VisualRowCount - 1
                        If oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE = "" Then
                            SumSjQty = SumSjQty
                        Else
                            SumSjQty = SumSjQty + oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE
                        End If
                        SumSjWt = SumSjWt + oMat01.Columns("SjWeight").Cells(i + 1).Specific.VALUE
                        If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                            SumQty = SumQty
                        Else
                            SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                        End If
                        SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                        HWeight = HWeight + (oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE * oMat01.Columns("UnWeight").Cells(i + 1).Specific.VALUE / 1000)
                    Next i
                    oForm01.Items("SumSjQty").Specific.VALUE = SumSjQty
                    oForm01.Items("SumSjWt").Specific.VALUE = SumSjWt
                    oForm01.Items("SumQty").Specific.VALUE = SumQty
                    oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                    oForm01.Items("HWeight").Specific.VALUE = HWeight
                    
                ElseIf pval.ColUID = "Price" Then
                    If (Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0) Then
                        Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
                        Call oDS_PS_SD040L.setValue("U_LinTotal", pval.Row - 1, 0)
                    Else
                        Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_SD040L.setValue("U_LinTotal", pval.Row - 1, (oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE * oMat01.Columns("Weight").Cells(pval.Row).Specific.VALUE))
                    End If
'                ElseIf pval.ColUID = "WhsCode" Then
'                    Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    Call oDS_PS_SD040L.setValue("U_WhsName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT WhsName FROM [OWHS] WHERE WhsCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value & "'", 0, 1))
                Else
                    oForm01.Items("TranCard").Click ct_Regular
                    Call oDS_PS_SD040L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
                oMat01.LoadFromDataSource
                oMat01.AutoResizeColumns
                oForm01.Update
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_SD040H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "CntcCode") Then
                    Call oDS_PS_SD040H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_SD040H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1))
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
    Dim i&, SumSjQty As Long, SumSjWt As Currency, SumQty As Long, SumWeight As Currency, HWeight As Currency
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
        Call PS_SD040_FormItemEnabled
        Call PS_SD040_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        
        For i = 0 To oMat01.VisualRowCount - 1
            If oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE = "" Then
                SumSjQty = SumSjQty
            Else
                SumSjQty = SumSjQty + oMat01.Columns("SjQty").Cells(i + 1).Specific.VALUE
            End If
            SumSjWt = SumSjWt + oMat01.Columns("SjWeight").Cells(i + 1).Specific.VALUE
            If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                SumQty = SumQty
            Else
                SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
            End If
            SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
            HWeight = HWeight + (oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE * oMat01.Columns("UnWeight").Cells(i + 1).Specific.VALUE / 1000)
        Next i
        oForm01.Items("SumSjQty").Specific.VALUE = SumSjQty
        oForm01.Items("SumSjWt").Specific.VALUE = SumSjWt
        oForm01.Items("SumQty").Specific.VALUE = SumQty
        oForm01.Items("SumWeight").Specific.VALUE = SumWeight
        oForm01.Items("HWeight").Specific.VALUE = HWeight
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
            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_SD040H", "U_CardCode,U_CardName")
        End If
        If (pval.ItemUID = "DCardCod" Or pval.ItemUID = "DCardNam") Then
            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_SD040H", "U_DCardCod,U_DCardNam")
        End If
        If (pval.ItemUID = "Mat01") Then
            If (pval.ColUID = "WhsCode") Then
                If pval.SelectedObjects Is Nothing Then
                Else
                    Set oDataTable01 = pval.SelectedObjects
                    Call oDS_PS_SD040L.setValue("U_WhsCode", pval.Row - 1, oDataTable01.Columns("WhsCode").Cells(0).VALUE)
                    Call oDS_PS_SD040L.setValue("U_WhsName", pval.Row - 1, oDataTable01.Columns("WhsName").Cells(0).VALUE)
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
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//행삭제전 행삭제가능여부검사
            If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
                MDC_Com.MDC_GF_Message "납품된행은 삭제할수 없습니다.", "W"
                BubbleEvent = False
                Exit Sub
            End If
            If oForm01.Items("Opt02").Specific.Selected = True Or oForm01.Items("Opt03").Specific.Selected = True Then '//멀티게이지, 분말
                For i = 0 To oDS_PS_SD040L.Size - 1
                    If i = oDS_PS_SD040L.Size - 1 Then
                        Exit For
                    End If
                    If oDS_PS_SD040L.GetValue("U_PackNo", i) = oDS_PS_SD040L.GetValue("U_PackNo", oLastColRow01 - 1) Then '//선택된행과 같은값을 가지는 모든행
                        Call oDS_PS_SD040L.RemoveRecord(i)
                        i = i - 1
                    End If
                Next
                For i = 0 To oDS_PS_SD040L.Size - 1
                    Call oDS_PS_SD040L.setValue("U_LineNum", i, i + 1)
                Next
                oMat01.LoadFromDataSource '//행삭제 정상처리
                
                If oMat01.RowCount = 0 Then
                    Call PS_SD040_AddMatrixRow(0)
                Else
                    If Trim(oDS_PS_SD040L.GetValue("U_SD030Num", oMat01.RowCount - 1)) <> "" Then
                        Call PS_SD040_AddMatrixRow(oMat01.RowCount)
                    End If
                End If
                If (oMat01.RowCount > 1) Then
                    oForm01.Items("Focus").Click ct_Regular
                    oForm01.Items("CardCode").Enabled = False
                    oForm01.Items("BPLId").Enabled = False
                    oForm01.Items("TrType").Enabled = False
                Else
                    oForm01.Items("Focus").Click ct_Regular
                    oForm01.Items("CardCode").Enabled = True
                    oForm01.Items("BPLId").Enabled = True
                    oForm01.Items("DCardCod").Enabled = True
                    oForm01.Items("TrType").Enabled = True
                End If
                oForm01.Update
'                MDC_Com.MDC_GF_Message "멀티게이지는 포장번호 단위로 납품가능합니다. 행삭제 할수 없습니다.", "W"
                BubbleEvent = False '//멀티의경우 이미 행삭제가 되었으므로 행삭제를 하지 않는다.
                Exit Sub
            End If
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_SD040L.RemoveRecord(oDS_PS_SD040L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_SD040_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_SD040L.GetValue("U_SD030Num", oMat01.RowCount - 1)) <> "" Then
                    Call PS_SD040_AddMatrixRow(oMat01.RowCount)
                End If
            End If
            If (oMat01.RowCount > 1) Then
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("CardCode").Enabled = False
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("TrType").Enabled = False
            Else
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("CardCode").Enabled = True
                oForm01.Items("BPLId").Enabled = True
                oForm01.Items("DCardCod").Enabled = True
                oForm01.Items("TrType").Enabled = True
            End If
            oForm01.Update
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SD040_CreateItems() As Boolean
On Error GoTo PS_SD040_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim optBtn As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_SD040H = oForm01.DataSources.DBDataSources("@PS_SD040H")
    Set oDS_PS_SD040L = oForm01.DataSources.DBDataSources("@PS_SD040L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("SumSjQty", dt_SUM)
    Call oForm01.DataSources.UserDataSources.Add("SumSjWt", dt_QUANTITY)
    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    Call oForm01.DataSources.UserDataSources.Add("SumWeight", dt_QUANTITY)
    Call oForm01.DataSources.UserDataSources.Add("HWeight", dt_QUANTITY)

    Call oForm01.Items("SumSjQty").Specific.DataBind.SetBound(True, "", "SumSjQty")
    Call oForm01.Items("SumSjWt").Specific.DataBind.SetBound(True, "", "SumSjWt")
    Call oForm01.Items("SumQty").Specific.DataBind.SetBound(True, "", "SumQty")
    Call oForm01.Items("SumWeight").Specific.DataBind.SetBound(True, "", "SumWeight")
    Call oForm01.Items("HWeight").Specific.DataBind.SetBound(True, "", "HWeight")
    
    Set optBtn = oForm01.Items("Opt01").Specific
    optBtn.ValOn = "2"
    optBtn.ValOff = "0"
    optBtn.Selected = True
    
    Set optBtn = oForm01.Items("Opt02").Specific
    optBtn.ValOn = "1"
    optBtn.ValOff = "0"
    optBtn.GroupWith ("Opt01")
    
    Set optBtn = oForm01.Items("Opt03").Specific
    optBtn.ValOn = "3"
    optBtn.ValOff = "0"
    optBtn.GroupWith ("Opt02")
    
    Set optBtn = Nothing
    
    'Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    
    '//2010.12.06 추가
    '담당자
    oDS_PS_SD040H.setValue "U_CntcCode", 0, User_MSTCOD()
    Call oDS_PS_SD040H.setValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oForm01.Items("CntcCode").Specific.VALUE & "'", 0, 1))
   
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_SD040_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD040_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SD040_ComboBox_Setting()
On Error GoTo PS_SD040_ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "DocType", "", "1", "출하요청")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "DocType", "", "2", "선출요청")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("DocType").Specific, "PS_PS_SD040", "DocType")
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "TrType", "", "1", "일반")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "TrType", "", "2", "임가공")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("TrType").Specific, "PS_PS_SD040", "TrType")
    
    
    Set oCombo = oForm01.Items("ExportYN").Specific
    oCombo.ValidValues.Add "N", "내수"
    oCombo.ValidValues.Add "Y", "수출"
    oCombo.Select 0, psk_Index
    
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD040", "Mat01", "Status", "O", "미결")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD040", "Mat01", "Status", "C", "완료")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Status"), "PS_SD040", "Mat01", "Status")
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "ProgStat", "", "1", "출하요청")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "ProgStat", "", "2", "선출요청")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "ProgStat", "", "3", "납품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "ProgStat", "", "4", "반품")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("ProgStat").Specific, "PS_PS_SD040", "ProgStat")
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD040", "Mat01", "TrType", "1", "일반")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD040", "Mat01", "TrType", "2", "군납")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("TrType"), "PS_SD040", "Mat01", "TrType")
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PS_SD040", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PS_SD040", "ItemCode")
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemGpCd"), "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmBsort"), "SELECT Code,Name FROM [@PSH_ITMBSORT]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemType"), "SELECT Code,Name FROM [@PSH_SHAPE]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Quality"), "SELECT Code,Name FROM [@PSH_QUALITY]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Mark"), "SELECT Code,Name FROM [@PSH_MARK]")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("SbasUnit"), "SELECT Code,Name FROM [@PSH_UOMORG]")
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD040_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD040_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD040_CF_ChooseFromList()
On Error GoTo PS_SD040_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLs01             As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCons01             As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCon01              As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFL01              As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oCFLCreationParams01  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oEdit01             As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

    Set oEdit = oForm01.Items("DCardCod").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = lf_BusinessPartner
    oCFLCreationParams.uniqueID = "CFLCARDCODE"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

    oEdit.ChooseFromListUID = "CFLCARDCODE"
    oEdit.ChooseFromListAlias = "CardCode"
    
    Set oCons = oCFL.GetConditions()
    Set oCon = oCons.Add()
    oCon.Alias = "CardType"
    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    oCon.CondVal = "C"
    oCFL.SetConditions oCons
    
    
    Set oEdit01 = oForm01.Items("CardCode").Specific
    Set oCFLs01 = oForm01.ChooseFromLists
    Set oCFLCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
    
    oCFLCreationParams01.ObjectType = lf_BusinessPartner
    oCFLCreationParams01.uniqueID = "CFLCARD2CODE"
    oCFLCreationParams01.MultiSelection = False
    Set oCFL01 = oCFLs01.Add(oCFLCreationParams01)

    oEdit01.ChooseFromListUID = "CFLCARD2CODE"
    oEdit01.ChooseFromListAlias = "CardCode"
    
    Set oCons01 = oCFL01.GetConditions()
    Set oCon01 = oCons01.Add()
    oCon01.Alias = "CardType"
    oCon01.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    oCon01.CondVal = "C"
    oCFL01.SetConditions oCons01
    

    Set oColumn = oMat01.Columns("WhsCode")
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = lf_Warehouses
    oCFLCreationParams.uniqueID = "CFLWAREHOUSES"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)
    oColumn.ChooseFromListUID = "CFLWAREHOUSES"
    oColumn.ChooseFromListAlias = "WhsCode"

'    Set oEdit = oForm01.Items("CARDNAME").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_BusinessPartner
'    oCFLCreationParams.uniqueID = "CFLCARDNAME"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)
'
'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLCARDNAME"
'    oEdit.ChooseFromListAlias = "CardName"
    Exit Sub
PS_SD040_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD040_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD040_FormItemEnabled()
On Error GoTo PS_SD040_FormItemEnabled_Error
    Dim i As Integer

    Call oForm01.Freeze(True)
    oForm01.Items("Focus").Click ct_Regular
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DCardCod").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("DocType").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        oForm01.Items("Comments").Enabled = True
        oForm01.Items("TrType").Enabled = True
        oForm01.Items("Opt01").Enabled = True
        oForm01.Items("Opt02").Enabled = True
        oForm01.Items("Opt03").Enabled = True
        oForm01.Items("Mat01").Enabled = True
        
        oMat01.AutoResizeColumns
        Call PS_SD040_FormClear '//UDO방식
'        Call oForm01.Items("BPLId").Specific.Select("1", psk_ByValue)
        Call oForm01.Items("BPLId").Specific.Select(User_BPLId(), psk_ByValue) '//2010.12.06 추가
        Call oForm01.Items("DocType").Specific.Select("1", psk_ByValue)
        oForm01.Items("CardCode").Click ct_Regular
        oForm01.Items("DocType").Enabled = False
        Call oForm01.Items("TrType").Specific.Select("1", psk_ByValue)
        Call oForm01.Items("ExportYN").Specific.Select("N", psk_ByValue)
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        oMat01.Columns("SD030Num").Visible = True
        oMat01.Columns("PackNo").Visible = True
        oMat01.Columns("OrderNum").Visible = True
        oForm01.Items("DocDate").Specific.String = Format(Now, "yyyymmdd")
        oForm01.Items("DueDate").Specific.String = Format(Now, "yyyymmdd")
        oForm01.Items("ProgStat").Enabled = True
        Call oForm01.Items("ProgStat").Specific.Select("3", psk_ByValue)
        oForm01.Items("CardCode").Click ct_Regular
        oForm01.Items("ProgStat").Enabled = False
        oForm01.Items("Opt01").Specific.Selected = True
        
        oForm01.Items("SumSjQty").Specific.VALUE = 0
        oForm01.Items("SumSjWt").Specific.VALUE = 0
        oForm01.Items("SumQty").Specific.VALUE = 0
        oForm01.Items("SumWeight").Specific.VALUE = 0
        oForm01.Items("HWeight").Specific.VALUE = 0
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DCardCod").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("DocType").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        oForm01.Items("Comments").Enabled = True
        oForm01.Items("TrType").Enabled = True
        oForm01.Items("Opt01").Enabled = True
        oForm01.Items("Opt02").Enabled = True
        oForm01.Items("Opt03").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        oMat01.AutoResizeColumns
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DueDate").Enabled = True
        oMat01.Columns("SD030Num").Visible = True
        oMat01.Columns("PackNo").Visible = True
        oMat01.Columns("OrderNum").Visible = True
        
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        '//문서상태검사
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("CardCode").Enabled = False
        oForm01.Items("DCardCod").Enabled = True
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("DocType").Enabled = False
        oForm01.Items("CntcCode").Enabled = False
        oForm01.Items("DocDate").Enabled = False
        oForm01.Items("DueDate").Enabled = False
        oForm01.Items("Comments").Enabled = True
        oForm01.Items("TrType").Enabled = False
        oForm01.Items("Opt01").Enabled = False
        oForm01.Items("Opt02").Enabled = False
        oForm01.Items("Opt03").Enabled = False
        oForm01.Items("Mat01").Enabled = False
        oMat01.AutoResizeColumns
        oForm01.Items("DocDate").Enabled = False
        oForm01.Items("DueDate").Enabled = False
        oMat01.Columns("SD030Num").Visible = True
        oMat01.Columns("PackNo").Visible = True
        oMat01.Columns("OrderNum").Visible = True
        '//각모드에따른 아이템설정
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD040_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD040_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD040_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD040_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_SD040L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD040L.Offset = oRow
    oDS_PS_SD040L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD040_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD040_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD040_FormClear()
On Error GoTo PS_SD040_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD040'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SD040_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD040_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD040_EnableMenus()
On Error GoTo PS_SD040_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1293", True)
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, True, True, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_SD040_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD040_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD040_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_SD040_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_SD040_FormItemEnabled
        Call PS_SD040_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_SD040_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_SD040_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD040_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD040_DataValidCheck() As Boolean
On Error GoTo PS_SD040_DataValidCheck_Error
    PS_SD040_DataValidCheck = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oForm01.Items("CardCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "고객은 필수입니다.", bmt_Short, True
        oForm01.Items("CardCode").Click ct_Regular
        PS_SD040_DataValidCheck = False
        Exit Function
    End If
    
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "전기일은 필수입니다.", bmt_Short, True
        oForm01.Items("DocDate").Click ct_Regular
        PS_SD040_DataValidCheck = False
        Exit Function
    End If
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_SD040_DataValidCheck = False
        Exit Function
    End If
    
    '마감상태 체크_S(2017.11.23 송명규 추가)
    If MDC_PS_Common.Check_Finish_Status(Trim(oForm01.Items("BPLId").Specific.VALUE), oForm01.Items("DocDate").Specific.VALUE, oForm01.TypeEx) = False Then
        Call Sbo_Application.SetStatusBarMessage("마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기일을 확인하고, 회계부서로 문의하세요.", bmt_Short, True)
        PS_SD040_DataValidCheck = False
        Exit Function
    End If
    '마감상태 체크_E(2017.11.23 송명규 추가)
    
    For i = 1 To oMat01.VisualRowCount - 1
    
        If oForm01.Items("TrType").Specific.VALUE = "1" Then '멀티게이지가 아닐 때만 체크
            '전기일 체크_S
            Query01 = "           SELECT  A.U_DocDate"
            Query01 = Query01 & " FROM    [@PS_SD030H] AS A"
            Query01 = Query01 & "         INNER JOIN"
            Query01 = Query01 & "         [@PS_SD030L] AS B"
            Query01 = Query01 & "           ON A.DocEntry = B.DocEntry"
            Query01 = Query01 & " WHERE   A.DocEntry = " & oMat01.Columns("SD030H").Cells(i).Specific.VALUE
            Query01 = Query01 & "         AND B.U_LineNum = " & oMat01.Columns("SD030L").Cells(i).Specific.VALUE
        
            Call RecordSet01.DoQuery(Query01)
            
            '생산완료의 전기일(출하요청의 전기일:생산완료의 전기일은 데이터의 구조상 JOIN을 할 수가 없음)보다 이전으로 입력불가(2011.12.07 송명규 추가)
            If Format(RecordSet01.Fields(0).VALUE, "yyyyMMdd") > oForm01.Items("DocDate").Specific.VALUE Then
            
                Sbo_Application.SetStatusBarMessage "납품처리 전기일이 출하요청(" & oMat01.Columns("SD030Num").Cells(i).Specific.VALUE & ") 의 전기일보다 이전입니다. 관리자에게 문의하시기 바랍니다.", bmt_Short, True
                oMat01.Columns("SD030Num").Cells(i).Click ct_Regular
                
                Set RecordSet01 = Nothing
                
                PS_SD040_DataValidCheck = False
                Exit Function
            
            End If
            '전기일 체크_E
        End If
    
        If oForm01.Items("Opt01").Specific.Selected = True Then
            If (oMat01.Columns("SD030Num").Cells(i).Specific.VALUE = "") Then
                Sbo_Application.SetStatusBarMessage "출하(선출)요청은 필수입니다.", bmt_Short, True
                oMat01.Columns("SD030Num").Cells(i).Click ct_Regular
                PS_SD040_DataValidCheck = False
                Exit Function
            End If
        End If
        If oForm01.Items("Opt02").Specific.Selected = True Then
            If (oMat01.Columns("PackNo").Cells(i).Specific.VALUE = "") Then
                Sbo_Application.SetStatusBarMessage "포장번호는 필수입니다.", bmt_Short, True
                oMat01.Columns("PackNo").Cells(i).Click ct_Regular
                PS_SD040_DataValidCheck = False
                Exit Function
            End If
        End If
        If oForm01.Items("Opt03").Specific.Selected = True Then
            If (oMat01.Columns("PackNo").Cells(i).Specific.VALUE = "") Then
                Sbo_Application.SetStatusBarMessage "포장번호는 필수입니다.", bmt_Short, True
                oMat01.Columns("PackNo").Cells(i).Click ct_Regular
                PS_SD040_DataValidCheck = False
                Exit Function
            End If
        End If
        If oForm01.Items("TrType").Specific.VALUE = "2" Then '//거래형태가 임가공
            If MDC_PS_Common.GetItem_ManBtchNum(oMat01.Columns("ItemCode").Cells(i).Specific.VALUE) = "Y" Then '//해당품목이 배치를 사용하는 품목이면
                If oMat01.Columns("LotNo").Cells(i).Specific.VALUE = "" Then
                    Sbo_Application.SetStatusBarMessage "임가공거래시 LotNo는 필수입니다.", bmt_Short, True
                    oMat01.Columns("LotNo").Cells(i).Click ct_Regular
                    PS_SD040_DataValidCheck = False
                    Exit Function
                End If
            End If
        End If
        If (Val(oMat01.Columns("Weight").Cells(i).Specific.VALUE) <= 0) Then
            Sbo_Application.SetStatusBarMessage "중량(수량)은 필수입니다.", bmt_Short, True
            oMat01.Columns("Weight").Cells(i).Click ct_Regular
            PS_SD040_DataValidCheck = False
            Exit Function
        End If
'        If (Val(oMat01.Columns("LinTotal").Cells(i).Specific.Value) <= 0) Then
'            Sbo_Application.SetStatusBarMessage "총계는 필수입니다.", bmt_Short, True
'            oMat01.Columns("Price").Cells(i).Click ct_Regular
'            PS_SD040_DataValidCheck = False
'            Exit Function
'        End If
        If (oMat01.Columns("WhsCode").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
            oMat01.Columns("WhsCode").Cells(i).Click ct_Regular
            PS_SD040_DataValidCheck = False
            Exit Function
        End If
        For j = i + 1 To oMat01.VisualRowCount - 1
            If oForm01.Items("Opt01").Specific.Selected = True Then
                If (oMat01.Columns("SD030Num").Cells(i).Specific.VALUE = oMat01.Columns("SD030Num").Cells(j).Specific.VALUE) Then
                    Sbo_Application.SetStatusBarMessage "동일한 출하요청문서가 존재합니다.", bmt_Short, True
                    oMat01.Columns("SD030Num").Cells(i).Click ct_Regular
                    PS_SD040_DataValidCheck = False
                    Exit Function
                End If
                If (Mid(oMat01.Columns("SD030Num").Cells(i).Specific.VALUE, 1, InStr(oMat01.Columns("SD030Num").Cells(i).Specific.VALUE, "-") - 1) <> _
                Mid(oMat01.Columns("SD030Num").Cells(j).Specific.VALUE, 1, InStr(oMat01.Columns("SD030Num").Cells(j).Specific.VALUE, "-") - 1)) Then
                    If Trim(oForm01.Items("BPLId").Specific.VALUE) = "4" Then '[2011.1.10추가]구로사업소일경우에만 동일하지 않은 출하요청문서 막음
                        Sbo_Application.SetStatusBarMessage "동일하지않은 출하요청문서가 존재합니다.", bmt_Short, True
                        oMat01.Columns("SD030Num").Cells(i).Click ct_Regular
                        PS_SD040_DataValidCheck = False
                        Exit Function
                    End If
                End If
                
            End If
            If oForm01.Items("Opt02").Specific.Selected = True Then
'                If (oMat01.Columns("PackNo").Cells(i).Specific.Value <> oMat01.Columns("SD030Num").Cells(j).Specific.Value) Then
'                    Sbo_Application.SetStatusBarMessage "동일하지않은 포장번호가 존재합니다.", bmt_Short, True
'                    oMat01.Columns("PackNo").Cells(i).Click ct_Regular
'                    PS_SD040_DataValidCheck = False
'                    Exit Function
'                End If
            End If
        Next
    Next
    '//한도초과 Check
    If oForm01.Items("CardCode").Specific.VALUE <> "12532" Then '//풍산은 체크안함
        If (PS_SD040_ValidateCreditLine = False) Then
            PS_SD040_DataValidCheck = False
            Exit Function
        End If
    End If
    If (PS_SD040_Validate("검사") = False) Then
        PS_SD040_DataValidCheck = False
        Exit Function
    End If
    
    Call oDS_PS_SD040L.RemoveRecord(oDS_PS_SD040L.Size - 1)
    Call oMat01.LoadFromDataSource
    If oForm01.Mode = fm_ADD_MODE Then
        Call PS_SD040_FormClear
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_SD040_DataValidCheck_Error:
    PS_SD040_DataValidCheck = False
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_SD040_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD040_MTX01()
On Error GoTo PS_SD040_MTX01_Error
    '//메트릭스에 데이터 로드
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
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_SD040_MTX01_Exit
    End If
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_SD040L.InsertRecord (i)
        End If
        oDS_PS_SD040L.Offset = i
        Call oDS_PS_SD040L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_SD040L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
    Exit Sub
PS_SD040_MTX01_Exit:
    Set RecordSet01 = Nothing
    Exit Sub
PS_SD040_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_SD040_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SD040_DI_API_01() As Boolean
On Error GoTo PS_SD040_DI_API_01_Error
    PS_SD040_DI_API_01 = True
    Dim i, j, K, m As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim NeedWeight As Double '//필요중량
    Dim RemainWeight As Double '//잔여중량
    Dim SelectedWeight As Double '//선택중량
    Dim LineNumCount As Long
    Dim ResultDocNum As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    '//배치정보
    ReDim BatchInformation(0)
    BatchInformationCount = 0
    '//품목정보
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("LotNo").Cells(i).Specific.VALUE '//배치정보가져가기
        ItemInformation(ItemInformationCount).Qty = oMat01.Columns("Qty").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Weight = oMat01.Columns("Weight").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Currency = oMat01.Columns("Currency").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Price = oMat01.Columns("Price").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).LineTotal = oMat01.Columns("LinTotal").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).SD040HNum = oForm01.Items("DocEntry").Specific.VALUE 'Val(oMat01.Columns("SD030H").Cells(i).Specific.VALUE) '문서번호(2016.10.20 송명규 수정, 출하요청 문서번호>납품처리 문서번호)
        ItemInformation(ItemInformationCount).SD040LNum = i 'Val(oMat01.Columns("LineID").Cells(i).Specific.VALUE) Val(oMat01.Columns("SD030L").Cells(i).Specific.VALUE) '라인번호(2016.10.20 송명규 수정, 출하요청 라인번호>납품처리 라인번호)
        ItemInformation(ItemInformationCount).ORDRNum = Val(oMat01.Columns("ORDRNum").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).RDR1Num = Val(oMat01.Columns("RDR1Num").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next

    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = oForm01.Items("BPLId").Specific.Selected.VALUE
    oDIObject.CardCode = oForm01.Items("CardCode").Specific.VALUE
    oDIObject.UserFields("U_DCardCod").VALUE = oForm01.Items("DCardCod").Specific.VALUE
    oDIObject.UserFields("U_DCardNam").VALUE = oForm01.Items("DCardNam").Specific.VALUE
    oDIObject.UserFields("U_TradeType").VALUE = oForm01.Items("TrType").Specific.Selected.VALUE
    If oForm01.Items("DocDate").Specific.VALUE <> "" Then
        oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    If oForm01.Items("DueDate").Specific.VALUE <> "" Then
        oDIObject.DocDueDate = Format(oForm01.Items("DueDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    
    For i = 0 To ItemInformationCount - 1
        If ItemInformation(i).Check = True Then
            GoTo Continue_First
        End If
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.BaseType = 17
        oDIObject.Lines.BaseEntry = ItemInformation(i).ORDRNum
        oDIObject.Lines.BaseLine = ItemInformation(i).RDR1Num
        oDIObject.Lines.UserFields("U_BaseType").VALUE = "PS_SD040"
        oDIObject.Lines.UserFields("U_BaseEntry").VALUE = ItemInformation(i).SD040HNum
        oDIObject.Lines.UserFields("U_BaseLine").VALUE = ItemInformation(i).SD040LNum

        For j = i To ItemInformationCount - 1
            If ItemInformation(j).Check = True Then
                GoTo Continue_Second
            End If
            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).ORDRNum <> ItemInformation(j).ORDRNum Or ItemInformation(i).RDR1Num <> ItemInformation(j).RDR1Num) Then
                GoTo Continue_Second
            End If
            '//같은것
            oDIObject.Lines.UserFields("U_Qty").VALUE = oDIObject.Lines.UserFields("U_Qty").VALUE + ItemInformation(j).Qty
            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Weight
            oDIObject.Lines.Currency = ItemInformation(j).Currency
            oDIObject.Lines.Price = oDIObject.Lines.Price + ItemInformation(j).Price
            oDIObject.Lines.LineTotal = oDIObject.Lines.LineTotal + ItemInformation(j).LineTotal
            If (oForm01.Items("TrType").Specific.VALUE = "1") And oForm01.Items("Opt03").Specific.Selected <> True Then '//일반
                '//DoNothing
            ElseIf (oForm01.Items("TrType").Specific.VALUE = "2") Or oForm01.Items("Opt03").Specific.Selected = True Then '//임가공, 분말일때
                If MDC_PS_Common.GetItem_ManBtchNum(ItemInformation(j).ItemCode) = "Y" Then
                    NeedWeight = ItemInformation(j).Weight
                    Query01 = "EXEC PS_SD040_03 '" & ItemInformation(j).ItemCode & "','" & ItemInformation(j).BatchNum & "','" & ItemInformation(j).WhsCode & "'"
                    RecordSet01.DoQuery Query01
                    For K = 1 To RecordSet01.RecordCount '//재고배치조회
                        If (NeedWeight <= 0) Then
                            Exit For
                        End If
                        RemainWeight = RecordSet01.Fields("Weight").VALUE
                        For m = 0 To BatchInformationCount - 1
                            If (RecordSet01.Fields("ItemCode").VALUE = BatchInformation(m).ItemCode And RecordSet01.Fields("BatchNum").VALUE = BatchInformation(m).BatchNum And RecordSet01.Fields("WhsCode").VALUE = BatchInformation(m).WhsCode) Then
                                RemainWeight = RemainWeight - BatchInformation(m).Weight
                            End If
                        Next
                        If (RemainWeight > 0) Then '//배치의중량이 남아있으면
                            
                            If MDC_PS_Common.GetItem_ItmBsort(ItemInformation(j).ItemCode) = "102" Or MDC_PS_Common.GetItem_ItmBsort(ItemInformation(j).ItemCode) = "111" Then '//부품
                                '//중량이 선택가능한 상황에서 여러배치를 선택할수 있다.
                                oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet01.Fields("BatchNum").VALUE
                                If (NeedWeight > RemainWeight) Then
                                    SelectedWeight = RemainWeight
                                    oDIObject.Lines.BatchNumbers.Quantity = SelectedWeight
                                ElseIf (NeedWeight <= RemainWeight) Then
                                    SelectedWeight = NeedWeight
                                    oDIObject.Lines.BatchNumbers.Quantity = SelectedWeight
                                End If
                                oDIObject.Lines.BatchNumbers.Add
                                NeedWeight = NeedWeight - RemainWeight
                                ReDim Preserve BatchInformation(BatchInformationCount)
                                BatchInformation(BatchInformationCount).ItemCode = RecordSet01.Fields("ItemCode").VALUE
                                BatchInformation(BatchInformationCount).BatchNum = RecordSet01.Fields("BatchNum").VALUE
                                BatchInformation(BatchInformationCount).WhsCode = RecordSet01.Fields("WhsCode").VALUE
                                BatchInformation(BatchInformationCount).Weight = SelectedWeight
                                BatchInformationCount = BatchInformationCount + 1
                            ElseIf MDC_PS_Common.GetItem_ItmBsort(ItemInformation(j).ItemCode) = "104" _
                            Or MDC_PS_Common.GetItem_ItmBsort(ItemInformation(j).ItemCode) = "302" Then  '//MULTI
                                If (NeedWeight = RemainWeight) Then '//필요중량과 배치의 중량이 같아야 선택가능, 다르면 선택불가
                                    SelectedWeight = NeedWeight
                                    oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet01.Fields("BatchNum").VALUE
                                    oDIObject.Lines.BatchNumbers.Quantity = SelectedWeight
                                    oDIObject.Lines.BatchNumbers.Add
                                    NeedWeight = NeedWeight - RemainWeight
                                    ReDim Preserve BatchInformation(BatchInformationCount)
                                    BatchInformation(BatchInformationCount).ItemCode = RecordSet01.Fields("ItemCode").VALUE
                                    BatchInformation(BatchInformationCount).BatchNum = RecordSet01.Fields("BatchNum").VALUE
                                    BatchInformation(BatchInformationCount).WhsCode = RecordSet01.Fields("WhsCode").VALUE
                                    BatchInformation(BatchInformationCount).Weight = SelectedWeight
                                    BatchInformationCount = BatchInformationCount + 1
                                End If
                            End If
                        Else '//배치가 이미 선택된경우
                            '//DoNothing 다음배치로이동
                        End If
                        RecordSet01.MoveNext '//다음배치
                    Next
                    If NeedWeight > 0 Then
                        '//MDC_Com.MDC_GF_Message "배치가 모두 선택되지 않았습니다. 재고를 확인하시기 바랍니다.", "W"
                    End If
                End If
            End If
            ItemInformation(j).DLN1Num = LineNumCount
            ItemInformation(j).Check = True
Continue_Second:
        Next
        LineNumCount = LineNumCount + 1
Continue_First:
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To ItemInformationCount - 1
            Call oDS_PS_SD040L.setValue("U_ODLNNum", i, ResultDocNum)
            Call oDS_PS_SD040L.setValue("U_DLN1Num", i, ItemInformation(i).DLN1Num)
'            MDC_PS_Common.DoQuery ("UPDATE [@PS_SD030H] SET U_ProgStat = '3' WHERE DocEntry = '" & ItemInformation(i).SD030HNum & "'") '//문서상태를 납품으로 변경
'            MDC_PS_Common.DoQuery ("UPDATE [@PS_SD030L] SET U_Status = 'C' WHERE DocEntry = '" & ItemInformation(i).SD030HNum & "' AND LineId = '" & ItemInformation(i).SD030LNum & "'") '//일부라도 납품되면 행이 클로즈된다.
        Next
    Else
        GoTo PS_SD040_DI_API_01_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    Set RecordSet01 = Nothing
    Set oDIObject = Nothing
    Exit Function
PS_SD040_DI_API_01_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_SD040_DI_API_01 = False
    Set oDIObject = Nothing
    Set RecordSet01 = Nothing
    Exit Function
PS_SD040_DI_API_01_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_SD040_DI_API_01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_SD040_DI_API_01 = False
    Set oDIObject = Nothing
    Set RecordSet01 = Nothing
End Function

Private Function PS_SD040_DI_API_02() As Boolean
On Error GoTo PS_SD040_DI_API_02_Error
    PS_SD040_DI_API_02 = True
    Dim i, j, K, m As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim NeedWeight As Double '//필요중량
    Dim RemainWeight As Double '//잔여중량
    Dim SelectedWeight As Double '//선택중량
    Dim LineNumCount As Long
    Dim ResultDocNum As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    '//배치정보
    ReDim BatchInformation(0)
    BatchInformationCount = 0
    '//품목정보
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("LotNo").Cells(i).Specific.VALUE '//배치정보가져가기
        ItemInformation(ItemInformationCount).Qty = oMat01.Columns("Qty").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Weight = oMat01.Columns("Weight").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Currency = oMat01.Columns("Currency").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Price = oMat01.Columns("Price").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).LineTotal = oMat01.Columns("LinTotal").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).WhsCode = oMat01.Columns("WhsCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).SD030HNum = Val(oMat01.Columns("SD030H").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).SD030LNum = Val(oMat01.Columns("SD030L").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).ORDRNum = Val(oMat01.Columns("ORDRNum").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).RDR1Num = Val(oMat01.Columns("RDR1Num").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = oForm01.Items("BPLId").Specific.Selected.VALUE
    oDIObject.CardCode = oForm01.Items("CardCode").Specific.VALUE
    oDIObject.UserFields("U_DCardCod").VALUE = oForm01.Items("DCardCod").Specific.VALUE
    oDIObject.UserFields("U_DCardNam").VALUE = oForm01.Items("DCardNam").Specific.VALUE
    oDIObject.UserFields("U_TradeType").VALUE = oForm01.Items("TrType").Specific.Selected.VALUE
    If oForm01.Items("DocDate").Specific.VALUE <> "" Then
        oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    If oForm01.Items("DueDate").Specific.VALUE <> "" Then
        oDIObject.DocDueDate = Format(oForm01.Items("DueDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        
        oDIObject.Lines.UserFields("U_BaseType").VALUE = "PS_SD040" '//별로 의미가 없을듯..
        oDIObject.Lines.UserFields("U_Qty").VALUE = ItemInformation(i).Qty
        oDIObject.Lines.Quantity = ItemInformation(i).Weight
        oDIObject.Lines.Currency = ItemInformation(i).Currency
        oDIObject.Lines.Price = ItemInformation(i).Price
        oDIObject.Lines.LineTotal = ItemInformation(i).LineTotal
        oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(i).BatchNum
        oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(i).Weight
        oDIObject.Lines.BatchNumbers.Add
        ItemInformation(i).DLN1Num = LineNumCount
        ItemInformation(i).Check = True
        LineNumCount = LineNumCount + 1
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To ItemInformationCount - 1
            Call oDS_PS_SD040L.setValue("U_ODLNNum", i, ResultDocNum)
            Call oDS_PS_SD040L.setValue("U_DLN1Num", i, ItemInformation(i).DLN1Num)
        Next
    Else
        GoTo PS_SD040_DI_API_02_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    Set RecordSet01 = Nothing
    Set oDIObject = Nothing
    Exit Function
PS_SD040_DI_API_02_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_SD040_DI_API_02 = False
    Set oDIObject = Nothing
    Set RecordSet01 = Nothing
    Exit Function
PS_SD040_DI_API_02_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_SD040_DI_API_02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_SD040_DI_API_02 = False
    Set oDIObject = Nothing
    Set RecordSet01 = Nothing
End Function

Private Function PS_SD040_DI_API_03() As Boolean
On Error GoTo PS_SD040_DI_API_03_Error
    PS_SD040_DI_API_03 = True
    Dim i, j, K, m As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim NeedWeight As Double '//필요중량
    Dim RemainWeight As Double '//잔여중량
    Dim SelectedWeight As Double '//선택중량
    Dim LineNumCount As Long
    Dim ResultDocNum As Long
    Dim BatchYN As String
    
    Dim Query01 As String
    Dim Query02 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    '//배치정보
    ReDim BatchInformation(0)
    BatchInformationCount = 0
    '//품목정보
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    
    
    '//배치번호 Y/N LotNo에 번호가 있으면 Batch = "Y"
    If oMat01.Columns("LotNo").Cells(1).Specific.VALUE = "" Then
        BatchYN = "N"
    Else
        BatchYN = "Y"
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
    
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ODLNNum = Val(oMat01.Columns("ODLNNum").Cells(i).Specific.VALUE) '//납품
        ItemInformation(ItemInformationCount).DLN1Num = Val(oMat01.Columns("DLN1Num").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).SD040HNum = Val(oForm01.Items("DocEntry").Specific.VALUE) '//납품A
        ItemInformation(ItemInformationCount).SD040LNum = Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oReturns) '//반품DI '//납품의 취소시 반품, AR송장시, AR대변메모, AR송장이 지급시
    oDIObject.BPL_IDAssignedToInvoice = oForm01.Items("BPLId").Specific.Selected.VALUE
    oDIObject.CardCode = oForm01.Items("CardCode").Specific.VALUE
    oDIObject.UserFields("U_DCardCod").VALUE = oForm01.Items("DCardCod").Specific.VALUE
    oDIObject.UserFields("U_DCardNam").VALUE = oForm01.Items("DCardNam").Specific.VALUE
    oDIObject.UserFields("U_TradeType").VALUE = oForm01.Items("TrType").Specific.Selected.VALUE
    If oForm01.Items("DocDate").Specific.VALUE <> "" Then
        oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    If oForm01.Items("DueDate").Specific.VALUE <> "" Then
        oDIObject.DocDueDate = Format(oForm01.Items("DueDate").Specific.VALUE, "&&&&-&&-&&")
    End If
    
    For i = 0 To ItemInformationCount - 1
    
        If i <> 0 Then
            '생산완료 등록 시 분할하여 등록한 데이터를 기준으로 납품처리 후 취소처리 할 때 판매오더의
            '이전 라인의 문서번호와 행번호를 체크하여 같으면 DI의 라인을 추가하지 않는 로직이 필요
            '해당 로직이 없으면 취소시 DI 에러 발생(2012.07.17 송명규 추가)
            '==>MG생산의 납품 취소시 에러가 남 아래 조건을 주석처리하면 반품처리가 됨. 배치번호문제인것 같음.
            '   배치제품이 안닌것만 조건을 걸어야 하지 않을까? 합니다. 2012.07.29 노근용
            If BatchYN = "N" Then
                '//배치관리가 아닐때
                If ItemInformation(i - 1).ODLNNum <> ItemInformation(i).ODLNNum And ItemInformation(i - 1).DLN1Num <> ItemInformation(i).DLN1Num Then
                    oDIObject.Lines.Add
                End If
            Else
                If oForm01.Items("Opt03").Specific.Selected = True Then
                    If ItemInformation(i - 1).ODLNNum <> ItemInformation(i).ODLNNum And ItemInformation(i - 1).DLN1Num <> ItemInformation(i).DLN1Num Then
                        '//배치번호일때
                        oDIObject.Lines.Add
                    End If
                Else
                    oDIObject.Lines.Add
                End If
            End If
        End If
        '//쿼리로 값구하기, 해당 납품라인에 대해 처리하기
        Query01 = "SELECT "
        Query01 = Query01 & " ItemCode AS ItemCode,"
        Query01 = Query01 & " WhsCode AS WhsCode,"
        Query01 = Query01 & " U_Qty AS Qty,"
        Query01 = Query01 & " Quantity AS Weight,"
        Query01 = Query01 & " DocEntry AS ODLNNum,"
        Query01 = Query01 & " LineNum AS DLN1Num"
        Query01 = Query01 & " FROM"
        Query01 = Query01 & " [DLN1]"
        Query01 = Query01 & " WHERE"
        Query01 = Query01 & " DocEntry = '" & ItemInformation(i).ODLNNum & "'"
        Query01 = Query01 & " AND LineNum = '" & ItemInformation(i).DLN1Num & "'"
        '//Query01 = Query01 & " AND LineStatus = 'O'"
        RecordSet01.DoQuery Query01
        
        oDIObject.Lines.ItemCode = RecordSet01.Fields("ItemCode").VALUE
        oDIObject.Lines.WarehouseCode = RecordSet01("WhsCode").VALUE
        oDIObject.Lines.UserFields("U_BaseType").VALUE = "PS_SD040" '//별로 의미가 없을듯..
        oDIObject.Lines.UserFields("U_Qty").VALUE = RecordSet01.Fields("Qty").VALUE
        oDIObject.Lines.Quantity = RecordSet01.Fields("Weight").VALUE
        oDIObject.Lines.BaseType = "15"
        oDIObject.Lines.BaseEntry = RecordSet01.Fields("ODLNNum").VALUE
        oDIObject.Lines.BaseLine = RecordSet01.Fields("DLN1Num").VALUE
        If MDC_PS_Common.GetItem_ManBtchNum(RecordSet01.Fields("ItemCode").VALUE) = "Y" Then
            '//해당납품으로 출고된 배치내역 조회
            Query02 = "SELECT "
            Query02 = Query02 & " BatchNum, Quantity"
            'Query02 = Query02 & " FROM [IBT1] "
            Query02 = Query02 & " FROM [IBT1_LINK] "
            Query02 = Query02 & " WHERE BaseType = '15'"
            Query02 = Query02 & " AND BaseEntry = '" & RecordSet01.Fields("ODLNNum").VALUE & "'"
            Query02 = Query02 & " AND BaseLinNum = '" & RecordSet01.Fields("DLN1Num").VALUE & "'"
            RecordSet02.DoQuery Query02
            For j = 0 To RecordSet02.RecordCount - 1
                oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet02.Fields("BatchNum").VALUE
                oDIObject.Lines.BatchNumbers.Quantity = RecordSet02.Fields("Quantity").VALUE
                oDIObject.Lines.BatchNumbers.Add
                RecordSet02.MoveNext
            Next
        End If
        ItemInformation(i).RDN1Num = LineNumCount
        LineNumCount = LineNumCount + 1
        RecordSet01.MoveNext
    Next

    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        
'        MDC_PS_Common.DoQuery ("UPDATE [@PS_SD040H] SET U_ProgStat = '4' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'")
'        'Call RecordSet03.DoQuery("UPDATE [@PS_SD040H] SET U_ProgStat = '4' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'")
'
'        For i = 0 To ItemInformationCount - 1 '//납품문서번호 업데이트
'            MDC_PS_Common.DoQuery ("UPDATE [@PS_SD040L] SET U_ORDNNum = '" & ResultDocNum & "', U_RDN1Num = '" & ItemInformation(i).RDN1Num & "' WHERE DocEntry = '" & ItemInformation(i).SD040HNum & "' AND LineId = '" & ItemInformation(i).SD040LNum & "'")
'            'Call RecordSet03.DoQuery("UPDATE [@PS_SD040L] SET U_ORDNNum = '" & ResultDocNum & "', U_RDN1Num = '" & ItemInformation(i).RDN1Num & "' WHERE DocEntry = '" & ItemInformation(i).SD040HNum & "' AND LineId = '" & ItemInformation(i).SD040LNum & "'")
'        Next
       
    Else
        GoTo PS_SD040_DI_API_03_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
        
        '트랜젝션 속에 UPDATE 구문이 있어서 장시간 트랜젝션이 걸릴 경우 "시스템 호출 오류"가 발생하는 듯 함
        '트랜젝션 종료 후 정상적으로 Commit 되었을 때 상태 UPDATE 및 반품문서번호 UPDATE 되게 수정
        '2013.10.07 송명규 수정
        MDC_PS_Common.DoQuery ("UPDATE [@PS_SD040H] SET U_ProgStat = '4' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'")
        'Call RecordSet03.DoQuery("UPDATE [@PS_SD040H] SET U_ProgStat = '4' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'")

        For i = 0 To ItemInformationCount - 1 '//납품문서번호 업데이트
            MDC_PS_Common.DoQuery ("UPDATE [@PS_SD040L] SET U_ORDNNum = '" & ResultDocNum & "', U_RDN1Num = '" & ItemInformation(i).RDN1Num & "' WHERE DocEntry = '" & ItemInformation(i).SD040HNum & "' AND LineId = '" & ItemInformation(i).SD040LNum & "'")
            'Call RecordSet03.DoQuery("UPDATE [@PS_SD040L] SET U_ORDNNum = '" & ResultDocNum & "', U_RDN1Num = '" & ItemInformation(i).RDN1Num & "' WHERE DocEntry = '" & ItemInformation(i).SD040HNum & "' AND LineId = '" & ItemInformation(i).SD040LNum & "'")
        Next
        
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Set oDIObject = Nothing
    Exit Function
PS_SD040_DI_API_03_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_SD040_DI_API_03 = False
    Set oDIObject = Nothing
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Exit Function
PS_SD040_DI_API_03_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_SD040_DI_API_03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_SD040_DI_API_03 = False
    Set oDIObject = Nothing
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
End Function

Function PS_SD040_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_SD040_Validate_Error
    PS_SD040_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_SD040H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
        PS_SD040_Validate = False
        GoTo PS_SD040_Validate_Exit
    End If
    
    If ValidateType = "검사" Then
        '//여신한도체크
'        If MDC_PS_Common.GetValue("EXEC PS_SD040_04 '" & oForm01.Items("CardCode").Specific.VALUE & "'", 0, 1) <= 0 Then
'            Call MDC_Com.MDC_GF_Message("여신한도 부족합니다.", "W")
'            PS_SD040_Validate = False
'            GoTo PS_SD040_Validate_Exit
'        End If
                
        For i = 1 To oMat01.VisualRowCount - 1 '//입력된 행에 대해
            If oForm01.Items("Opt01").Specific.Selected = True Then
                If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_SD030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_SD030L.LineId) = '" & oMat01.Columns("SD030Num").Cells(i).Specific.VALUE & "'", 0, 1) <= 0 Then
                    Call MDC_Com.MDC_GF_Message("출하(선출)요청문서가 존재하지 않습니다.", "W")
                    PS_SD040_Validate = False
                    GoTo PS_SD040_Validate_Exit
                End If
            End If
            If oForm01.Items("Opt02").Specific.Selected = True Then
            
            End If
        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_SD040_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_SD040_Validate_Error:
    PS_SD040_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_SD040_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD040_Print_Report01()
On Error GoTo PS_SD040_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim i As Long
    Dim sQry01                      As String
    Dim Comments                    As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
        Exit Sub
    End If
    If oForm01.Items("BPLId").Specific.Selected.VALUE <> "2" Then
        Call MDC_Com.MDC_GF_Message("사업장이 동래가 아닙니다.", "W")
        Exit Sub
    End If
    For i = 1 To oMat01.VisualRowCount - 1
        If oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.VALUE <> "105" And oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.VALUE <> "106" Then
            Call MDC_Com.MDC_GF_Message("품목이 기계공구,몰드가 아닙니다.", "W")
            Exit Sub
        End If
    Next
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    sQry01 = "SELECT COUNT(*) AS COUNT FROM [@PS_SD040L] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    Call oRecordSet01.DoQuery(sQry01)
    If Trim(oRecordSet01.Fields(0).VALUE) > 7 Then
    
        WinTitle = "[PS_PP540_60] 레포트"
        ReportName = "PS_PP540_60.rpt"
        sQry = "EXEC PS_PP540_60 '납품','" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        ReDim gRpt_Formula(1)
        ReDim gRpt_Formula_Value(1)
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        ReDim gRpt_SFormula(1, 1)
        ReDim gRpt_SFormula_Value(1, 1)
        
    Else
        
        WinTitle = "[PS_PP540_10] 레포트"
        ReportName = "PS_PP540_10.rpt"
        sQry = "EXEC PS_PP540_10 '납품','" & oForm01.Items("DocEntry").Specific.VALUE & "'"
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
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Exit Sub
PS_SD040_Print_Report01_Error:
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "PS_SD040_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD040_Print_Report02(ByVal Chk As String)
On Error GoTo PS_SD040_Print_Report02_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim Sub_sQry                    As String
    Dim i As Long
    Dim sQry01                      As String
    Dim Comments                    As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
        Exit Sub
    End If
    If oForm01.Items("BPLId").Specific.Selected.VALUE <> "1" And oForm01.Items("BPLId").Specific.Selected.VALUE <> "4" Then
        Call MDC_Com.MDC_GF_Message("사업장이 창원,서울이 아닙니다.", "W")
        Exit Sub
    End If
'    For i = 1 To oMat01.VisualRowCount - 1
'        If oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.Value <> "101" And oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.Value <> "102" Then
'            Call MDC_Com.MDC_GF_Message("품목이 휘팅,부품이 아닙니다.", "W")
'            Exit Sub
'        End If
'    Next
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    sQry01 = "SELECT COUNT(*) AS COUNT FROM [@PS_SD040L] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    Call oRecordSet01.DoQuery(sQry01)
    If Trim(oRecordSet01.Fields(0).VALUE) > 10 Then
        Set oRecordSet01 = Nothing
        WinTitle = "[PS_SD220_20] 레포트"
        ReportName = "PS_SD220_20.rpt"
        sQry = "EXEC PS_SD220_20 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        ReDim gRpt_Formula(4)
        ReDim gRpt_Formula_Value(4)
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        ReDim gRpt_SFormula(1, 1)
        ReDim gRpt_SFormula_Value(1, 1)
        
        Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        sQry01 = "Select BPLName FROM [OBPL] WHERE BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "'"
        Call oRecordSet01.DoQuery(sQry01)
        gRpt_Formula(1) = "BPLName"
        gRpt_Formula(2) = "CardName"
        gRpt_Formula(3) = "DCardNam"
        gRpt_Formula(4) = "DocDate"
'        Rpt_Formula(5) = "DocDate"
        
        gRpt_Formula_Value(1) = Trim(oRecordSet01.Fields(0).VALUE)
        gRpt_Formula_Value(2) = oForm01.Items("CardName").Specific.VALUE
        gRpt_Formula_Value(3) = oForm01.Items("DCardNam").Specific.VALUE
        gRpt_Formula_Value(4) = Format(oForm01.Items("DocDate").Specific.VALUE, "####-##-##")
    Else
        WinTitle = "[PS_SD220_10] 레포트"
        ReportName = "PS_SD220_10.rpt"
        sQry = "EXEC PS_SD220_10 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        ReDim gRpt_Formula(1)
        ReDim gRpt_Formula_Value(1)
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        ReDim gRpt_SFormula(1, 1)
        ReDim gRpt_SFormula_Value(1, 1)
        
        gRpt_Formula(1) = "Chk"
        gRpt_Formula_Value(1) = Chk
        
        Sub_sQry = "EXEC PS_SD220_11 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
        gRpt_SRptSqry(1) = Sub_sQry
        gRpt_SRptName(1) = "PS_SD220_SUB1"
        
    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Exit Sub
PS_SD040_Print_Report02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "PS_SD040_Print_Report02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD040_Print_Report03(ByVal Chk As String) ' 거래명세서(수량)
On Error GoTo PS_SD040_Print_Report03_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim Sub_sQry                    As String
    Dim i As Long
    Dim sQry01                      As String
    Dim Comments                    As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
        Exit Sub
    End If
    If oForm01.Items("BPLId").Specific.Selected.VALUE <> "1" And oForm01.Items("BPLId").Specific.Selected.VALUE <> "4" Then
        Call MDC_Com.MDC_GF_Message("사업장이 창원,서울이 아닙니다.", "W")
        Exit Sub
    End If
    
    WinTitle = "분말거래명세표(수량)[PS_SD220_50]"
    ReportName = "PS_SD220_50.rpt"
    sQry = "EXEC PS_SD040_50 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Exit Sub
PS_SD040_Print_Report03_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "PS_SD040_Print_Report03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD040_Print_Report03_1(ByVal Chk As String) ' 거래명세서(금액)
On Error GoTo PS_SD040_Print_Report03_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim Sub_sQry                    As String
    Dim i As Long
    Dim sQry01                      As String
    Dim Comments                    As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
    If oForm01.Items("ProgStat").Specific.Selected.VALUE <> "3" Then
        Call MDC_Com.MDC_GF_Message("문서상태가 납품이 아닙니다.", "W")
        Exit Sub
    End If
    If oForm01.Items("BPLId").Specific.Selected.VALUE <> "1" And oForm01.Items("BPLId").Specific.Selected.VALUE <> "4" Then
        Call MDC_Com.MDC_GF_Message("사업장이 창원,서울이 아닙니다.", "W")
        Exit Sub
    End If
    
    WinTitle = "분말거래명세표(금액)[PS_SD220_51]"
    ReportName = "PS_SD220_51.rpt"
    sQry = "EXEC PS_SD040_50 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Exit Sub
PS_SD040_Print_Report03_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "PS_SD040_Print_Report03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_SD040_ValidateCreditLine() As Boolean '//여신한도체크
On Error GoTo PS_SD040_ValidateCreditLine_Error
    PS_SD040_ValidateCreditLine = True
    
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    
    Dim i As Long
    Dim OCRDCreditLine As Currency '//고객여신한도
    Dim SD080CreditLine As Currency '//추가여신한도
    Dim OCRDBalance As Currency '//계정잔액
    Dim OCRDDNotesBal As Currency '//납품액
    Dim CurrentLineSum As Currency '//현재문서총계
    Dim OutPreP As Currency '출고예정금액
    
    Dim CardCode As String
    Dim DocDate As String

'    If Trim(oDS_PS_SD040H.GetValue("U_BPLId", 0)) = "4" Then
    If oForm01.Items("Opt03").Specific.Selected = True Then '//분말제품
    'If oMat01.Columns("WhsCode").Cells(1).Specific.VALUE = "104" And (oMat01.Columns("ItmBsort").Cells(1).Specific.VALUE = "101" Or oMat01.Columns("ItmBsort").Cells(1).Specific.VALUE = "202") Then
        'If MDC_PS_Common.GetValue("SELECT Isnull(U_PrtYn,'N')  FROM [@PS_SD030H] WHERE DocEntry = '" & oMat01.Columns("SD030H").Cells(1).Specific.VALUE & "'", 0, 1) = "N" Then
        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
        CardCode = oForm01.Items("CardCode").Specific.VALUE
        DocDate = oForm01.Items("DocDate").Specific.VALUE
        
        Query01 = "EXEC [S139_hando] '" & CardCode & "', '" & DocDate & "'"
        Call RecordSet01.DoQuery(Query01)

        If RecordSet01.RecordCount > 0 Then
            If RecordSet01.Fields("OverAmt").VALUE < 0 Then '//여신한도금액
                Call MDC_Com.MDC_GF_Message("여신한도를 초과했습니다.", "W")
                PS_SD040_ValidateCreditLine = False
                Exit Function
            End If
        End If
    End If

'    If Trim(oDS_PS_SD040H.GetValue("U_BPLId", 0)) = "4" Then
'        OCRDCreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT CreditLine FROM [OCRD] WHERE CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        SD080CreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Sum(PS_SD080L.U_RequestP) FROM [@PS_SD080H] PS_SD080H LEFT JOIN [@PS_SD080L] PS_SD080L ON PS_SD080H.DocEntry = PS_SD080L.DocEntry WHERE PS_SD080H.Canceled = 'N' And PS_SD080H.U_OkYN = 'Y' AND PS_SD080H.U_DocDate = '" & oForm01.Items("DocDate").Specific.VALUE & "' AND PS_SD080L.U_CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
''        OCRDBalance = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Balance FROM [OCRD] WHERE CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        OCRDBalance = MDC_PS_Common.GetValue("SELECT ISNULL((Select Sum(Debit - Credit) from JDT1 where ShortName = '" & oForm01.Items("CardCode").Specific.VALUE & "' And Account = '11104010'),0)", 0, 1)
'        OCRDDNotesBal = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT DNotesBal FROM [OCRD] WHERE CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
''        OutPreP = MDC_PS_Common.GetValue("Select IsNull((Select Sum(LineTotal) From  [ORDR] a Inner Join [RDR1] b On a.DocEntry = b.DocEntry Where  b.LineStatus = 'O' And a.DocStatus = 'O' And a.CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        CurrentLineSum = 0
'        For i = 1 To oMat01.VisualRowCount - 1
'            CurrentLineSum = CurrentLineSum + Val(oMat01.Columns("LinTotal").Cells(i).Specific.VALUE)
'        Next
''        If ((OCRDCreditLine + SD080CreditLine) - (OCRDBalance + OCRDDNotesBal) < CurrentLineSum) Then
'        If ((OCRDCreditLine + SD080CreditLine) - (OCRDBalance + OCRDDNotesBal + OutPreP) < CurrentLineSum) Then
'            Call MDC_Com.MDC_GF_Message("여신한도가 부족합니다.", "W")
'            PS_SD040_ValidateCreditLine = False
'            Exit Function
'        End If
'    End If
    
'    If Trim(oDS_PS_SD040H.GetValue("U_BPLId", 0)) = "4" Then
'        OCRDCreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT CreditLine FROM [OCRD] WHERE CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        SD080CreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Sum(PS_SD080L.U_RequestP) FROM [@PS_SD080H] PS_SD080H LEFT JOIN [@PS_SD080L] PS_SD080L ON PS_SD080H.DocEntry = PS_SD080L.DocEntry WHERE PS_SD080H.Canceled = 'N' And PS_SD080H.U_OkYN = 'Y' AND PS_SD080H.U_DocDate = '" & oForm01.Items("DocDate").Specific.VALUE & "' AND PS_SD080L.U_CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        OCRDBalance = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Balance FROM [OCRD] WHERE CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        OCRDDNotesBal = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT DNotesBal FROM [OCRD] WHERE CardCode = '" & oForm01.Items("CardCode").Specific.VALUE & "'),0)", 0, 1)
'        CurrentLineSum = 0
'        For i = 1 To oMat01.VisualRowCount - 1
'            CurrentLineSum = CurrentLineSum + Val(oMat01.Columns("LinTotal").Cells(i).Specific.VALUE)
'        Next
'        If ((OCRDCreditLine + SD080CreditLine) - (OCRDBalance + OCRDDNotesBal) < CurrentLineSum) Then
'            Call MDC_Com.MDC_GF_Message("여신한도가 부족합니다.", "W")
'            PS_SD040_ValidateCreditLine = False
'            Exit Function
'        End If
'    End If
    Exit Function
PS_SD040_ValidateCreditLine_Error:
   PS_SD040_ValidateCreditLine = False
   Sbo_Application.SetStatusBarMessage "PS_SD040_ValidateCreditLine_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PS_SD040_CheckDate(ByVal pBaseEntry As String) As Boolean
'******************************************************************************
'Function ID : PS_SD040_CheckDate()
'해당모듈    : PS_SD040
'기능        : 선행프로세스와 일자 비교
'인수        : pBaseEntry : 기준문서번호
'반환값      : True-선행프로세스보다 일자가 같거나 느릴 경우, False-선행프로세스보다 일자가 빠를 경우
'특이사항    : 구현만 함, 사용은 하지 않음
'******************************************************************************
On Error GoTo PS_SD040_CheckDate_Error

    Dim Query01      As String
    Dim loopCount    As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BaseEntry    As String
    Dim BaseLine     As String
    Dim DocType      As String
    Dim CurDocDate   As String
    
    Dim Entry()      As String
    
    
    Entry = Split(pBaseEntry, "-")
    BaseEntry = Entry(0)
    BaseLine = Entry(1)
    DocType = "PS_SD040"
    CurDocDate = oMat01.Columns("DocDate").Cells(loopCount).Specific.VALUE

    Query01 = "         EXEC PS_Z_CHECK_DATE '"
    Query01 = Query01 & BaseEntry & "','"
    Query01 = Query01 & BaseLine & "','"
    Query01 = Query01 & DocType & "','"
    Query01 = Query01 & CurDocDate & "'"

    Call oRecordSet01.DoQuery(Query01)

    If oRecordSet01.Fields("ReturnValue").VALUE = "False" Then
        PS_SD040_CheckDate = False
    Else
        PS_SD040_CheckDate = True
    End If
    
    Set oRecordSet01 = Nothing

    Exit Function
PS_SD040_CheckDate_Error:
    PS_SD040_CheckDate = False
    Call Sbo_Application.SetStatusBarMessage("PS_SD040_CheckDate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function
