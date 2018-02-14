VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP043"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_PP043.cls
'//  Module         : 생산관리 > 작업일보등록(포장사업팀 공정)
'//  Desc           : 작업일보등록
'//  FormType       : PS_PP043
'//  Create Date    : 2013.04.11
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix
Public oMat03 As SAPbouiCOM.Matrix
Private oDS_PS_PP043H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP043L As SAPbouiCOM.DBDataSource '등록라인
Private oDS_PS_PP043M As SAPbouiCOM.DBDataSource '등록라인
Private oDS_PS_PP043N As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMat01Row01 As Long
Private oMat02Row02 As Long
Private oMat03Row03 As Long

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

Private oDocType01 As String
Private oDocEntry01 As String
Private oFormMode01 As SAPbouiCOM.BoFormMode
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP043.srf")
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
    
    oFormUniqueID01 = "PS_PP043_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
    
    oForm01.Freeze True
    
    Call PS_PP043_CreateItems
    Call PS_PP043_ComboBox_Setting
    Call PS_PP043_CF_ChooseFromList
    Call PS_PP043_EnableMenus
    Call PS_PP043_SetDocument(oFromDocEntry01)
    Call PS_PP043_FormResize
    
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
                If oForm01.Mode = fm_OK_MODE Then
                    If (PS_PP043_Validate("취소") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If Sbo_Application.MessageBox("정말로 취소하시겠습니까?", "1", "예", "아니오") <> "1" Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    '//2010.12.30 yjh 추가---------------------------------------------------
                    If Insert_oInventoryGenEntry(2) = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                  '//2010.12.30 yjh 추가 end-----------------------------------------------
                Else
                    MDC_Com.MDC_GF_Message "현재 모드에서는 취소할수 없습니다.", "W"
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call Raise_EVENT_RECORD_MOVE(FormUID, pval, BubbleEvent)
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_PP043_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_PP043_FormItemEnabled '//UDO방식
                Call PS_PP043_AddMatrixRow01(0, True) '//UDO방식
                Call PS_PP043_AddMatrixRow02(0, True) '//UDO방식
                Call oForm01.Items("DocType").Specific.Select("20", psk_ByValue)
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call Raise_EVENT_RECORD_MOVE(FormUID, pval, BubbleEvent)
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
                If (oForm01.Mode <> fm_ADD_MODE) Then
                    If (PS_PP043_FindValidateDocument("@PS_PP040H") = False) Then
                        If Sbo_Application.Menus.Item("1281").Enabled = True Then '//찾기메뉴 활성화일때 수행
                            Sbo_Application.ActivateMenuItem ("1281")
                        Else
                            Sbo_Application.SetStatusBarMessage "관리자에게 문의바랍니다", bmt_Short, True
                        End If
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
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
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Or pval.ItemUID = "Mat03" Then
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
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oMat01Row01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat02" Then
        If pval.Row > 0 Then
            oMat02Row02 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat03" Then
        If pval.Row > 0 Then
            oMat03Row03 = pval.Row
        End If
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error
    Dim vReturnValue$
    
    Dim tot_time As Double
    Dim UnitTime As Double
    Dim UnitRemainTime As Double
    Dim i As Integer
    
    
    If pval.BeforeAction = True Then
        If pval.ItemUID = "PS_PP043" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_PP043_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//2010.12.30 yjh 추가---------------------------------------------------
'                vReturnValue = Sbo_Application.MessageBox("이 데이터를 취소한 후에는 변경할 수 없습니다. 계속하겠습니까?", 1, "&확인", "&취소")
'                Select Case vReturnValue
'                    Case 1:
                        If Insert_oInventoryGenExit(2) = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
'                            Sbo_Application.MessageBox ("선택저장 작업이 완료되었습니다.")
                        End If
                        
'                    Case 2:
'                        BubbleEvent = False
'                        Exit Sub
'                End Select
                '//2010.12.30 yjh 추가 end-----------------------------------------------
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_PP043_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                '//해야할일 작업
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "2" Then '//취소버튼 누를시 저장할 자료가 있으면 메시지 표시
            If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                If oMat01.VisualRowCount > 1 Then
                    vReturnValue = Sbo_Application.MessageBox("저장하지 않는 자료가 있습니다. 취소하시겠습니까?", 2, "&확인", "&취소")
                    Select Case vReturnValue
                        Case 1:
                        Case 2:
                                BubbleEvent = False
                                Exit Sub
                    End Select
                End If
            End If
        End If
        
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                Call PS_PP043_OrderInfoLoad
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                'Call PS_PP040_OrderInfoLoad
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "Button02" Then
            If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_OK_MODE Then
'                If Trim(oForm01.Items("SCpCode").Specific.VALUE) <> "CP50101" Then '//VM공정 외는 공수배부(VMILL 제외 기능 주석 2012.06.26 송명규)
                    For i = 1 To oMat02.VisualRowCount - 1
                        tot_time = tot_time + Val(oMat02.Columns("YTime").Cells(i).Specific.VALUE)
                    Next
                    
                    If Val(tot_time) > 0 Then
                        '//소숫점 2자리
                        UnitTime = Format(tot_time / (oMat01.VisualRowCount - 1), "#.##")
 
                        'UnitTime = trunc(tot_time / (oMat01.VisualRowCount - 1), 2)
                        UnitRemainTime = Format(Val(tot_time) - Val(UnitTime * (oMat01.VisualRowCount - 1)), "##0.##") 'Format 함수를 이용해서 강제로 소수점 2자리로 변경(2012.07.19 송명규)
                        For i = 0 To oMat01.VisualRowCount - 2
                            If i <> oMat01.VisualRowCount - 2 Then
                                'oMat01.Columns("WorkTime").Cells(i).Specific.VALUE = UnitTime
                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime)
                            Else
                                'oMat01.Columns("WorkTime").Cells(i).Specific.VALUE = UnitTime + UnitRemainTime
                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime + UnitRemainTime)
                            End If
                        Next
                    End If
                    oMat01.LoadFromDataSource
                    
                    If oForm01.Mode = fm_OK_MODE Then
                       oForm01.Mode = fm_UPDATE_MODE
                    End If
                    
'                End If
            End If
        End If
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_PP043" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_PP043_FormItemEnabled
                    Call PS_PP043_AddMatrixRow02(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    If (oFormMode01 = fm_UPDATE_MODE) Then
                        oFormMode01 = fm_OK_MODE
                        oForm01.Mode = fm_FIND_MODE
                        Call PS_PP043_FormItemEnabled
                        oForm01.Items("DocEntry").Specific.VALUE = oDocEntry01
                        oForm01.Items("1").Click ct_Regular
                    End If
                    Call PS_PP043_FormItemEnabled
                End If
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
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
        If pval.ItemUID = "Mat01" Then
            If pval.ColUID = "OrdMgNum" Then
                If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//일반,조정
                    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then
                        Call MDC_Com.MDC_GF_Message("작업구분이 선택되지 않았습니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Items("CpCode").Specific.VALUE = "" Then
                        Call MDC_Com.MDC_GF_Message("공정이 선택되지 않았습니다.", "W")
                        BubbleEvent = False
                        Exit Sub
                    Else
                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "108" Or oForm01.Items("OrdGbn").Specific.Selected.VALUE = "109" Then
                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                            BubbleEvent = False
                        Else
                            Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "OrdMgNum") '//사용자값활성
                        End If
                    End If
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//지원
                    
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//외주
                    
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//실적
                    
                End If
            
            End If
        End If
        If pval.ItemUID = "Mat02" Then
            If pval.ColUID = "WorkCode" Then
                If Val(oForm01.Items("BaseTime").Specific.VALUE) = 0 Then
                    Call MDC_Com.MDC_GF_Message("기준시간을 입력하지 않았습니다.", "W")
                    oForm01.Items("BaseTime").Click ct_Regular
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
        End If
        If pval.ItemUID = "SCpCode" Then
            If oForm01.Items("SOrdGbn").Specific.Selected.VALUE = "선택" Then
                Call MDC_Com.MDC_GF_Message("작업구분이 선택되지 않았습니다.", "W")
                BubbleEvent = False
                Exit Sub
            End If
        End If
        
        If pval.ItemUID = "SMoldNo" Then
            If oForm01.Items("SCpCode").Specific.VALUE = "" Then
                Call MDC_Com.MDC_GF_Message("공정이 선택되지 않았습니다.", "W")
                BubbleEvent = False
                Exit Sub
            End If
            If oForm01.Items("SMoldNo").Specific.VALUE = "" Then
                Sbo_Application.ActivateMenuItem ("7425")
                BubbleEvent = False
            End If
        End If
        
        If pval.ItemUID = "UseMCode" Then
            If oForm01.Items("UseMCode").Specific.VALUE = "" Then
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "UseMCode", "") '//사용자값활성
            End If
        End If
        
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat02", "WorkCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "SCpCode", "") '//사용자값활성
        
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat02", "NCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat03", "FailCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat03", "CsCpCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat03", "CsWkCode") '원인공정작업자명 포맷서치 설정(2012.03.20 송명규)
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "MachCode") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "MoldNo") '//사용자값활성
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
                    Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP043L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        'PS_PP043_AddMatrixRow (pval.Row)
                    End If
                Else
                    Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                End If
            ElseIf (pval.ItemUID = "Mat02") Then
                If (pval.ColUID = "특정컬럼") Then
                    '//기타작업
                    Call oDS_PS_PP043M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                    If oMat02.RowCount = pval.Row And Trim(oDS_PS_PP043M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        'PS_PP043_AddMatrixRow (pval.Row)
                    End If
                Else
                    Call oDS_PS_PP043M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                End If
            ElseIf (pval.ItemUID = "Mat03") Then
                If (pval.ColUID = "특정컬럼") Then
                Else
                    Call oDS_PS_PP043N.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                End If
            Else
                If (pval.ItemUID = "특정컬럼") Then
                    Call oDS_PS_PP043H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                ElseIf (pval.ItemUID = "SBPLId" Or pval.ItemUID = "SOrdGbn") Then
                    
                Else
                    Call oDS_PS_PP043H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oMat02.LoadFromDataSource
            oMat02.AutoResizeColumns
            oMat03.LoadFromDataSource
            oMat03.AutoResizeColumns
            oForm01.Update
            If pval.ItemUID = "Mat01" Then
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Collapsed
            ElseIf pval.ItemUID = "Mat02" Then
                oMat02.Columns(pval.ColUID).Cells(pval.Row).Click ct_Collapsed
            ElseIf pval.ItemUID = "Mat03" Then
                oMat03.Columns(pval.ColUID).Cells(pval.Row).Click ct_Collapsed
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
    If pval.BeforeAction = True Then
        If pval.ItemUID = "Opt01" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat02"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Opt02" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat03"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
        If pval.ItemUID = "Opt03" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oMat01.AutoResizeColumns
            oMat02.AutoResizeColumns
            oMat03.AutoResizeColumns
            oForm01.Freeze False
        End If
'        If pval.ItemUID = "Mat01" Then
'            If pval.Row > 0 Then
'                Call oMat01.SelectRow(pval.Row, True, False)
'            End If
'        End If
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)
                oMat01Row01 = pval.Row
            End If
            
        End If
        If pval.ItemUID = "Mat02" Then
            If pval.Row > 0 Then
                Call oMat02.SelectRow(pval.Row, True, False)
                oMat02Row02 = pval.Row
            End If
        End If
        If pval.ItemUID = "Mat03" Then
            If pval.Row > 0 Then
                Call oMat03.SelectRow(pval.Row, True, False)
                oMat03Row03 = pval.Row
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
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 일반,조정
                    If oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE = "" Then
                    
                    Else
                        If oMat03.VisualRowCount = 0 Then
                            Call PS_PP043_AddMatrixRow03(0, True)
                        Else
                            Call PS_PP043_AddMatrixRow03(oMat03.VisualRowCount)
                        End If
                        Call oDS_PS_PP043N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP043N.setValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns("CpCode").Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP043N.setValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns("CpName").Cells(pval.Row).Specific.VALUE)
                        oMat03.LoadFromDataSource
                        oMat03.AutoResizeColumns
                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
                        oMat03.FlushToDataSource
                    End If
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT지원인경우
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
                ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
                End If
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
    Dim oTempClass As Variant
    If pval.BeforeAction = True Then
        
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim Weight As Double
    
    Dim UnitTime As Long
    Dim UnitRemainTime As Double
    Dim Time As Double
    Dim Hour As Long
    Dim Minute As Long
    Dim LineTime As Currency
    
    Dim BPLID As String
    Dim CpCode As String
    Dim CpName As String
    Dim OrdGbn As String
    
    Dim ErrNum As String '에러 번호
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (PS_PP043_Validate("수정01") = False) Then
                    Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, Trim(oDS_PS_PP043L.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                Else
                    If (pval.ColUID = "OrdMgNum") Then
                        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                        
                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then '//작업구분에 값이 없으면 작업지시가 불러오기전
                            Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                        Else '//작업지시가 선택된상태
                            If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 일반,조정
                                BPLID = oForm01.Items("BPLId").Specific.VALUE
                                CpCode = oForm01.Items("CpCode").Specific.VALUE
                                CpName = oForm01.Items("CpName").Specific.VALUE
                                OrdGbn = oForm01.Items("OrdGbn").Specific.VALUE
                                
                                
                                For i = 1 To oMat01.RowCount
                                    '//현재 입력한 값이 이미 입력되어 있는경우
                                    If oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.VALUE And i <> pval.Row Then
                                        Call MDC_Com.MDC_GF_Message("이미 입력한 공정입니다.", "W")
                                        Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                                        GoTo Continue
                                    End If
'                                        '//공정라인의 공정순서가 앞공정보다 높으면
'                                        If Val(oMat01.Columns("Sequence").Cells(i).Specific.Value) >= MDC_PS_Common.GetValue("SELECT PS_PP030M.U_Sequence FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE CONVERT(NVARCHAR,PS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" & oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.Value & "'") Then
'                                            Call MDC_Com.MDC_GF_Message("공정순서가 올바르지 않습니다.", "W")
'                                            Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
'                                            GoTo Continue
'                                        End If
                                Next
                                Query01 = "Select a.ItemCode, a.ItemName From OITM a Where a.ItmsGrpCod = '102' "
                                Query01 = Query01 + "And Not Exists (Select * from [@PS_MM002H] b Where a.ItemCode = b.U_ItemCode And b.U_Type = '2')  " '원재료 판매분이 아닌것
                                Query01 = Query01 + "And a.ItemCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"
                                
                                'Query01 = "EXEC PS_PP043_02 '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"
                                RecordSet01.DoQuery Query01
                                If RecordSet01.RecordCount = 0 Then
                                    Call MDC_Com.MDC_GF_Message("제품코드가 아니거나 없는 원재료 판매 제품입니다.", "E")
                                    Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
                                Else
                                    Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                                    Call oDS_PS_PP043L.setValue("U_Sequence", pval.Row - 1, Right(CpCode, 1))
                                    Call oDS_PS_PP043L.setValue("U_CpCode", pval.Row - 1, CpCode)
                                    Call oDS_PS_PP043L.setValue("U_CpName", pval.Row - 1, CpName)
                                    Call oDS_PS_PP043L.setValue("U_OrdGbn", pval.Row - 1, OrdGbn)
                                    Call oDS_PS_PP043L.setValue("U_BPLId", pval.Row - 1, BPLID)
                                    Call oDS_PS_PP043L.setValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                                    Call oDS_PS_PP043L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
                                    Call oDS_PS_PP043L.setValue("U_PSum", pval.Row - 1, 0)
                                    Call oDS_PS_PP043L.setValue("U_BQty", pval.Row - 1, 0)
                                    
                                    Call oDS_PS_PP043L.setValue("U_PQty", pval.Row - 1, 0)
                                    Call oDS_PS_PP043L.setValue("U_YQty", pval.Row - 1, 0)
                                    
                                    Call oDS_PS_PP043L.setValue("U_PWeight", pval.Row - 1, 0)
                                    Call oDS_PS_PP043L.setValue("U_YWeight", pval.Row - 1, 0)
                                    Call oDS_PS_PP043L.setValue("U_NQty", pval.Row - 1, 0)
                                    Call oDS_PS_PP043L.setValue("U_NWeight", pval.Row - 1, 0)
                                    Call oDS_PS_PP043L.setValue("U_LineId", pval.Row - 1, "")
                                    Call oDS_PS_PP043L.setValue("U_WorkTime", pval.Row - 1, 0)
                                    Call oDS_PS_PP043L.setValue("U_PP030HNo", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
                                    Call oDS_PS_PP043L.setValue("U_PP030MNo", pval.Row - 1, Right(CpCode, 1))
                                    
                                    '//불량코드테이블
                                    If oMat03.VisualRowCount = 0 Then
                                        Call PS_PP043_AddMatrixRow03(0, True)
                                    Else
                                        Call PS_PP043_AddMatrixRow03(oMat03.VisualRowCount)
                                    End If
                                    Call oDS_PS_PP043N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields("ItemCode").VALUE)
                                    Call oDS_PS_PP043N.setValue("U_CpCode", oMat03.VisualRowCount - 1, CpCode)
                                    Call oDS_PS_PP043N.setValue("U_CpName", oMat03.VisualRowCount - 1, CpName)
                                End If
                            ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT지원

                            ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주

                            ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적

                            End If
Continue:
                            If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP043L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                                Call PS_PP043_AddMatrixRow01(pval.Row)
                            End If
                        End If
                        Set RecordSet01 = Nothing
                    ElseIf pval.ColUID = "PQty" Then
                        'If oMat01.Columns("OrdGbn").Cells(pval.Row).Specific.VALUE = "104" Then '//멀티이면
                            If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0 Then
                                Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, 0) '불량일경우
                                Call oDS_PS_PP043L.setValue("U_YQty", pval.Row - 1, 0) '합격수량에 0
                                Call oDS_PS_PP043L.setValue("U_NQty", pval.Row - 1, Val(oMat01.Columns("BQty").Cells(pval.Row).Specific.VALUE)) '전량 불량처리
                                Call oDS_PS_PP043L.setValue("U_ScrapWt", pval.Row - 1, 0) '스크랩 0
                                'Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP043L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                            'ElseIf Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) > Val(oMat01.Columns("BQty").Cells(pval.Row).Specific.VALUE) Then
'                                Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP043L.GetValue("U_" & pval.ColUID, pval.Row - 1))
'                            Else
'                                Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
'                                Call oDS_PS_PP043L.setValue("U_YQty", pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
'                                Call oDS_PS_PP043L.setValue("U_NQty", pval.Row - 1, 0) '불량 0
'                                Call oDS_PS_PP043L.setValue("U_ScrapWt", pval.Row - 1, Val(oMat01.Columns("BQty").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            End If
                            'Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, oDS_PS_PP043L.GetValue("U_" & pval.ColUID, pval.Row - 1))
                            Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            Call oDS_PS_PP043L.setValue("U_YQty", pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            
                            
                        'End If
                        
                    ElseIf pval.ColUID = "NQty" Then
                        If Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) <= 0 Then
                            Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oDS_PS_PP043L.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                        ElseIf Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) > Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) Then
                            Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oDS_PS_PP043L.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                        Else
                            Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                            Call oDS_PS_PP043L.setValue("U_YQty", pval.Row - 1, Val(oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE) - Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                        End If
                        
                    ElseIf pval.ColUID = "WorkTime" Then
                        Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    ElseIf pval.ColUID = "MachCode" Then
                        Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        Call oDS_PS_PP043L.setValue("U_MachName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                        
                    Else
                        Call oDS_PS_PP043L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    End If
                End If
            ElseIf (pval.ItemUID = "Mat02") Then
                If (pval.ColUID = "WorkCode") Then
                    '//기타작업
                    Call oDS_PS_PP043M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_PP043M.setValue("U_WorkName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" & oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                    '//기준시간을 작업시간에 입력
                    Call oDS_PS_PP043M.setValue("U_YTime", pval.Row - 1, oForm01.Items("BaseTime").Specific.VALUE)
                    Call oDS_PS_PP043M.setValue("U_TTime", pval.Row - 1, oForm01.Items("BaseTime").Specific.VALUE)
                    If oMat02.RowCount = pval.Row And Trim(oDS_PS_PP043M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        Call PS_PP043_AddMatrixRow02(pval.Row)
                    End If
                ElseIf pval.ColUID = "NStart" Then
                    Call oDS_PS_PP043M.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    If Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE) = 0 Or Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) = 0 Then
                        Call oDS_PS_PP043M.setValue("U_NTime", pval.Row - 1, 0)
                        Call oDS_PS_PP043M.setValue("U_YTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE))
                        Call oDS_PS_PP043M.setValue("U_TTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE))
                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티일때만..
                            If oMat02.VisualRowCount > 1 Then
                                If (MDC_PS_Common.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" & oMat01.Columns("PP030HNo").Cells(1).Specific.VALUE & "' ORDER BY U_Sequence ASC", 0, 1) <> _
                                oMat01.Columns("CpCode").Cells(1).Specific.VALUE) Then
                                    '//해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                    If Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) > 0 Then
                                        UnitTime = Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) / (oMat01.VisualRowCount - 1)
                                        UnitRemainTime = Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) - (UnitTime * (oMat01.VisualRowCount - 1))
                                        For i = 0 To oMat01.VisualRowCount - 2
                                            If i <> oMat01.VisualRowCount - 2 Then
                                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime)
                                            Else
                                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime + UnitRemainTime)
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE) <= Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) Then
                            Time = Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) - Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE)
                        Else
                            Time = (2400 - Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE)) + Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE)
                        End If
                        Hour = Fix(Time / 100)
                        Minute = Time Mod 100
                        Time = Hour
                        If Minute > 0 Then
                            Time = Time + 0.5
                        End If
                        Call oDS_PS_PP043M.setValue("U_NTime", pval.Row - 1, Time)
                        Call oDS_PS_PP043M.setValue("U_YTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE) - Time)
                        Call oDS_PS_PP043M.setValue("U_TTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE) - Time)
                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티일때만..
                            If oMat02.VisualRowCount > 1 Then
                                If (MDC_PS_Common.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" & oMat01.Columns("PP030HNo").Cells(1).Specific.VALUE & "' ORDER BY U_Sequence ASC", 0, 1) <> _
                                oMat01.Columns("CpCode").Cells(1).Specific.VALUE) Then
                                    '//해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                    If Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) > 0 Then
                                        UnitTime = Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) / (oMat01.VisualRowCount - 1)
                                        UnitRemainTime = Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) - (UnitTime * (oMat01.VisualRowCount - 1))
                                        For i = 0 To oMat01.VisualRowCount - 2
                                            If i <> oMat01.VisualRowCount - 2 Then
                                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime)
                                            Else
                                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime + UnitRemainTime)
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        End If
                    End If
                ElseIf pval.ColUID = "NEnd" Then
                    Call oDS_PS_PP043M.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    If Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE) = 0 Or Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) = 0 Then
                        Call oDS_PS_PP043M.setValue("U_NTime", pval.Row - 1, 0)
                        Call oDS_PS_PP043M.setValue("U_YTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE))
                        Call oDS_PS_PP043M.setValue("U_TTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE))
                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티일때만..
                            If oMat02.VisualRowCount > 1 Then
                                If (MDC_PS_Common.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" & oMat01.Columns("PP030HNo").Cells(1).Specific.VALUE & "' ORDER BY U_Sequence ASC", 0, 1) <> _
                                oMat01.Columns("CpCode").Cells(1).Specific.VALUE) Then
                                    '//해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                    If Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) > 0 Then
                                        UnitTime = Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) / (oMat01.VisualRowCount - 1)
                                        UnitRemainTime = Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) - (UnitTime * (oMat01.VisualRowCount - 1))
                                        For i = 0 To oMat01.VisualRowCount - 2
                                            If i <> oMat01.VisualRowCount - 2 Then
                                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime)
                                            Else
                                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime + UnitRemainTime)
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE) <= Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) Then
                            Time = Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE) - Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE)
                        Else
                            Time = (2400 - Val(oMat02.Columns("NStart").Cells(pval.Row).Specific.VALUE)) + Val(oMat02.Columns("NEnd").Cells(pval.Row).Specific.VALUE)
                        End If
                        Hour = Fix(Time / 100)
                        Minute = Time Mod 100
                        Time = Hour
                        If Minute > 0 Then
                            Time = Time + 0.5
                        End If
                        Call oDS_PS_PP043M.setValue("U_NTime", pval.Row - 1, Time)
                        Call oDS_PS_PP043M.setValue("U_YTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE) - Time)
                        Call oDS_PS_PP043M.setValue("U_TTime", pval.Row - 1, Val(oForm01.Items("BaseTime").Specific.VALUE) - Time)
                        If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티일때만..
                            If oMat02.VisualRowCount > 1 Then
                                If (MDC_PS_Common.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" & oMat01.Columns("PP030HNo").Cells(1).Specific.VALUE & "' ORDER BY U_Sequence ASC", 0, 1) <> _
                                oMat01.Columns("CpCode").Cells(1).Specific.VALUE) Then
                                    '//해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                    If Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) > 0 Then
                                        UnitTime = Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) / (oMat01.VisualRowCount - 1)
                                        UnitRemainTime = Val(oDS_PS_PP043M.GetValue("U_YTime", pval.Row - 1)) - (UnitTime * (oMat01.VisualRowCount - 1))
                                        For i = 0 To oMat01.VisualRowCount - 2
                                            If i <> oMat01.VisualRowCount - 2 Then
                                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime)
                                            Else
                                                Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime + UnitRemainTime)
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        End If
                    End If
                ElseIf pval.ColUID = "YTime" Then
                    Call oDS_PS_PP043M.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    Call oDS_PS_PP043M.setValue("U_TTime", pval.Row - 1, Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE))
                    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "104" Then '//멀티일때만..
                        If oMat02.VisualRowCount > 1 Then
                            If (MDC_PS_Common.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" & oMat01.Columns("PP030HNo").Cells(1).Specific.VALUE & "' ORDER BY U_Sequence ASC", 0, 1) <> _
                            oMat01.Columns("CpCode").Cells(1).Specific.VALUE) Then
                                '//해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                If Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) > 0 Then
                                    UnitTime = Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) / (oMat01.VisualRowCount - 1)
                                    UnitRemainTime = Val(oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) - (UnitTime * (oMat01.VisualRowCount - 1))
                                    For i = 0 To oMat01.VisualRowCount - 2
                                        If i <> oMat01.VisualRowCount - 2 Then
                                            Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime)
                                        Else
                                            Call oDS_PS_PP043L.setValue("U_WorkTime", i, UnitTime + UnitRemainTime)
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                Else
                    Call oDS_PS_PP043M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            ElseIf (pval.ItemUID = "Mat03") Then
                If (pval.ColUID = "FailCode") Then
                    Call oDS_PS_PP043N.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_PP043N.setValue("U_FailName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" & oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                ElseIf (pval.ColUID = "CsCpCode") Then
                    Call oDS_PS_PP043N.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_PP043N.setValue("U_CsCpName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" & oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                ElseIf (pval.ColUID = "CsWkCode") Then '원인공정 작업자 정보 추가(2012.03.20 송명규)
                    Call oDS_PS_PP043N.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    Call oDS_PS_PP043N.setValue("U_CsWkName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT T0.lastName+T0.firstName FROM OHEM T0 WHERE T0.U_MSTCOD = '" & oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", 0, 1))
                Else
                    Call oDS_PS_PP043N.setValue("U_" & pval.ColUID, pval.Row - 1, oMat03.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_PP043H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "BaseTime") Then
                    Call oDS_PS_PP043H.setValue("U_" & pval.ItemUID, 0, Val(oForm01.Items(pval.ItemUID).Specific.VALUE))
                ElseIf (pval.ItemUID = "UseMCode") Then
                    oForm01.Items("UseMName").Specific.VALUE = MDC_PS_Common.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1)
                    '장비의 금형번호
                    oForm01.Items("SMoldNo").Specific.VALUE = MDC_PS_Common.GetValue("SELECT U_MoldNo FROM [@PS_PP130H] WHERE U_MachCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1)
                    oForm01.Items("SMoldNm").Specific.VALUE = MDC_PS_Common.GetValue("SELECT b.U_Item + '[' + b.U_Callsize +']' FROM [@PS_PP130H] a Inner join [@PS_PP190H] b On a.U_MoldNo = b.Code WHERE a.U_MachCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1)
                ElseIf (pval.ItemUID = "SCpCode") Then
                    oForm01.Items("SCpName").Specific.VALUE = MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1)
                    If oForm01.Mode = fm_ADD_MODE Then
                        Call PS_PP043_OrderInfoLoad
                    End If
                    
                ElseIf (pval.ItemUID = "SMoldNo") Then
                    '금형번호의 이름
                    oForm01.Items("SMoldNm").Specific.VALUE = MDC_PS_Common.GetValue("SELECT U_Item + '[' + U_Callsize +']' FROM [@PS_PP190H] WHERE Code = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", 0, 1)
                
                Else
                    Call oDS_PS_PP043H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oMat02.LoadFromDataSource
            oMat02.AutoResizeColumns
            oMat03.LoadFromDataSource
            oMat03.AutoResizeColumns
            oForm01.Update
            If pval.ItemUID = "Mat01" Then
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            ElseIf pval.ItemUID = "Mat02" Then
                oMat02.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
            ElseIf pval.ItemUID = "Mat03" Then
                oMat03.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
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
    
    If ErrNum = "1" Then
        Call MDC_Com.MDC_GF_Message("품질부적합 등록건(등록상태)입니다. 품질보증팀(검사자)에 문의하십시오.", "W")
    ElseIf ErrNum = "2" Then
        Call MDC_Com.MDC_GF_Message("품질부적합 등록건(1차 해제상태)입니다. 품질보증팀(팀장)에 문의하십시오.", "W")
    Else
        Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_PP043_FormItemEnabled
        If pval.ItemUID = "Mat01" Then
            Call PS_PP043_AddMatrixRow01(oMat01.VisualRowCount) '//UDO방식
        ElseIf pval.ItemUID = "Mat02" Then
            Call PS_PP043_AddMatrixRow02(oMat02.VisualRowCount) '//UDO방식
        End If
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP043_FormResize
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
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP040H", "U_CardCode,U_CardName")
'        End If
'        If (pval.ItemUID = "Mat02") Then
'            If (pval.ColUID = "WorkCode") Then
'                If pval.SelectedObjects Is Nothing Then
'                Else
'                    Set oDataTable01 = pval.SelectedObjects
'                    Call oDS_PS_PP043M.setValue("U_WorkCode", pval.Row - 1, oDataTable01.Columns("empID").Cells(0).Value)
'                    Call oDS_PS_PP043M.setValue("U_WorkName", pval.Row - 1, oDataTable01.Columns("firstName").Cells(0).Value & oDataTable01.Columns("lastName").Cells(0).Value)
'                    If oMat02.RowCount = pval.Row And Trim(oDS_PS_PP043M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        Call PS_PP043_AddMatrixRow02(pval.Row)
'                    End If
'                    Set oDataTable01 = Nothing
'                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
'                    oMat02.LoadFromDataSource
'                    oMat02.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
'                End If
'            End If
'        End If
    End If
    Exit Sub
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error
    If pval.ItemUID = "Mat01" Or pval.ItemUID = "Mat02" Or pval.ItemUID = "Mat03" Then
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
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oMat01Row01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat02" Then
        If pval.Row > 0 Then
            oMat02Row02 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat03" Then
        If pval.Row > 0 Then
            oMat03Row03 = pval.Row
        End If
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
    Dim i, j As Long
    Dim Exist As Boolean
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '2011.01.24 추가(yjh)
            If Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "104" And Trim(oForm01.Items("CpCode").Specific.VALUE) = "CP50101" Then '멀티의 첫공정일 경우
                MDC_Com.MDC_GF_Message "멀티의 첫공정은 행삭제 할수 없습니다.", "E"
                BubbleEvent = False
                Exit Sub
            ElseIf Trim(oForm01.Items("OrdGbn").Specific.VALUE) = "107" And Trim(oForm01.Items("CpCode").Specific.VALUE) = "CP10101" Then '엔드베어링의 첫공정일 경우
                MDC_Com.MDC_GF_Message "엔드베어링의 첫공정은 행삭제 할수 없습니다.", "E"
                BubbleEvent = False
                Exit Sub
            End If '추가 End

            If oLastItemUID01 = "Mat01" Then
                If (PS_PP043_Validate("행삭제01") = False) Then
                    BubbleEvent = False
                    Exit Sub
                End If
Continue:
                For i = 1 To oMat03.RowCount
                    If oMat01.Columns("OrdMgNum").Cells(oLastColRow01).Specific.VALUE = oMat03.Columns("OrdMgNum").Cells(i).Specific.VALUE Then
                        oDS_PS_PP043N.RemoveRecord (i - 1)
                        oMat03.DeleteRow (i)
                        oMat03.FlushToDataSource
                        GoTo Continue
                    End If
                Next
            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            If oLastItemUID01 = "Mat01" Then
                For i = 1 To oMat01.VisualRowCount
                    oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                oMat01.FlushToDataSource
                Call oDS_PS_PP043L.RemoveRecord(oDS_PS_PP043L.Size - 1)
                oMat01.LoadFromDataSource
                If oMat01.RowCount = 0 Then
                    Call PS_PP043_AddMatrixRow01(0)
                Else
                    If Trim(oDS_PS_PP043L.GetValue("U_OrdMgNum", oMat01.RowCount - 1)) <> "" Then
                        Call PS_PP043_AddMatrixRow01(oMat01.RowCount)
                    End If
                End If
            ElseIf oLastItemUID01 = "Mat02" Then
                For i = 1 To oMat02.VisualRowCount
                    oMat02.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                oMat02.FlushToDataSource
                Call oDS_PS_PP043M.RemoveRecord(oDS_PS_PP043M.Size - 1)
                oMat02.LoadFromDataSource
                If oMat02.RowCount = 0 Then
                    Call PS_PP043_AddMatrixRow02(0)
                Else
                    If Trim(oDS_PS_PP043M.GetValue("U_WorkCode", oMat02.RowCount - 1)) <> "" Then
                        Call PS_PP043_AddMatrixRow02(oMat02.RowCount)
                    End If
                End If
            ElseIf oLastItemUID01 = "Mat03" Then
                For i = 1 To oMat03.VisualRowCount
                    oMat03.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                oMat03.FlushToDataSource
                If oDS_PS_PP043N.Size = 1 Then '//사이즈가 0일때는 행을 빼주면 oMat03.VisualRowCount 가 0 으로 변경되어서 문제가 생김
                Else
                    Call oDS_PS_PP043N.RemoveRecord(oDS_PS_PP043N.Size - 1)
                End If
                oMat03.LoadFromDataSource
                
                For i = 1 To oMat01.RowCount - 1 '//공정 테이블에는 있는데 불량 테이블에 존재하지 않는값이 있는경우 불량테이블에 값을 추가함
                    Exist = False
                    For j = 1 To oMat03.RowCount
                        If oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat03.Columns("OrdMgNum").Cells(j).Specific.VALUE Then
                            Exist = True
                        End If
                    Next
                    If Exist = False Then '//불량코드테이블에 값이 존재하지 않으면
                        If oMat03.VisualRowCount = 0 Then
                            Call PS_PP043_AddMatrixRow03(0, True)
                        Else
                            Call PS_PP043_AddMatrixRow03(oMat03.VisualRowCount)
                        End If
                        Call oDS_PS_PP043N.setValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE)
                        Call oDS_PS_PP043N.setValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns("CpCode").Cells(i).Specific.VALUE)
                        Call oDS_PS_PP043N.setValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns("CpName").Cells(i).Specific.VALUE)
                        oMat03.LoadFromDataSource
                        oMat03.AutoResizeColumns
                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
                        oMat03.FlushToDataSource
                    End If
                Next
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RECORD_MOVE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RECORD_MOVE_Error
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim DocEntry As String
    Dim DocEntryNext As String
    DocEntry = oForm01.Items("DocEntry").Specific.VALUE '//원본문서
    DocEntryNext = oForm01.Items("DocEntry").Specific.VALUE '//다음문서
    
    If pval.MenuUID = "1288" Then '//다음
        If pval.BeforeAction = True Then
            If oForm01.Mode = fm_ADD_MODE Then
                Sbo_Application.ActivateMenuItem ("1290")
                BubbleEvent = False
                Exit Sub
            ElseIf oForm01.Mode = fm_FIND_MODE Then
                If (oForm01.Items("DocEntry").Specific.VALUE = "") Then
                    Sbo_Application.ActivateMenuItem ("1290")
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
            If PS_PP043_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    ElseIf pval.MenuUID = "1289" Then '//이전
        If pval.BeforeAction = True Then
            If oForm01.Mode = fm_ADD_MODE Then
                Sbo_Application.ActivateMenuItem ("1291")
                BubbleEvent = False
                Exit Sub
            ElseIf oForm01.Mode = fm_FIND_MODE Then
                If (oForm01.Items("DocEntry").Specific.VALUE = "") Then
                    Sbo_Application.ActivateMenuItem ("1291")
                    BubbleEvent = False
                    Exit Sub
                End If
            End If
            If PS_PP043_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    ElseIf pval.MenuUID = "1290" Then '//첫번째레코드로이동
        If pval.BeforeAction = True Then
            Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry DESC" '//가장마지막행을 부여
            RecordSet01.DoQuery Query01
            DocEntry = Val(RecordSet01.Fields(0).VALUE) '//원본문서
            DocEntryNext = Val(RecordSet01.Fields(0).VALUE) '//다음문서
            Set RecordSet01 = Nothing
            If PS_PP043_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    ElseIf pval.MenuUID = "1291" Then '//마지막문서로이동
        If pval.BeforeAction = True Then
            Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry ASC" '//가장 첫행을 부여
            RecordSet01.DoQuery Query01
            DocEntry = Val(RecordSet01.Fields(0).VALUE) '//원본문서
            DocEntryNext = Val(RecordSet01.Fields(0).VALUE) '//다음문서
            Set RecordSet01 = Nothing
            If PS_PP043_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") = False Then
                BubbleEvent = False
                Exit Sub
            End If
        End If
    End If
    Set RecordSet01 = Nothing
    Exit Sub
Raise_EVENT_RECORD_MOVE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RECORD_MOVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP043_CreateItems() As Boolean
On Error GoTo PS_PP043_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_PP043H = oForm01.DataSources.DBDataSources("@PS_PP040H")
    Set oDS_PS_PP043L = oForm01.DataSources.DBDataSources("@PS_PP040L")
    Set oDS_PS_PP043M = oForm01.DataSources.DBDataSources("@PS_PP040M")
    Set oDS_PS_PP043N = oForm01.DataSources.DBDataSources("@PS_PP040N")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
    
    Set oMat03 = oForm01.Items("Mat03").Specific
    oMat03.SelectionMode = ms_NotSupported
    oMat03.AutoResizeColumns
    
    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("Opt03", dt_SHORT_TEXT, 10)
    Call oForm01.Items("Opt01").Specific.DataBind.SetBound(True, "", "Opt01")
    Call oForm01.Items("Opt02").Specific.DataBind.SetBound(True, "", "Opt02")
    Call oForm01.Items("Opt03").Specific.DataBind.SetBound(True, "", "Opt03")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt03")
    
    
'    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
'    Call oForm01.DataSources.UserDataSources.Add("WhsCode", dt_SHORT_TEXT, 100)
'    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
'    Call oForm01.Items("WhsCode").Specific.DataBind.SetBound(True, "", "WhsCode")
    
    Call oForm01.DataSources.UserDataSources.Add("SBPLId", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("SOrdGbn", dt_SHORT_TEXT, 100)
    Call oForm01.DataSources.UserDataSources.Add("SCpCode", dt_SHORT_TEXT, 30)
    Call oForm01.DataSources.UserDataSources.Add("SCpName", dt_SHORT_TEXT, 30)
'    Call oForm01.DataSources.UserDataSources.Add("SMoldNo", dt_SHORT_TEXT, 30)
'    Call oForm01.DataSources.UserDataSources.Add("SMoldNm", dt_SHORT_TEXT, 30)
    Call oForm01.Items("SBPLId").Specific.DataBind.SetBound(True, "", "SBPLId")
    Call oForm01.Items("SOrdGbn").Specific.DataBind.SetBound(True, "", "SOrdGbn")
    Call oForm01.Items("SCpCode").Specific.DataBind.SetBound(True, "", "SCpCode")
    Call oForm01.Items("SCpName").Specific.DataBind.SetBound(True, "", "SCpName")
'    Call oForm01.Items("SMoldNo").Specific.DataBind.SetBound(True, "", "SMoldNo")
'    Call oForm01.Items("SMoldNm").Specific.DataBind.SetBound(True, "", "SMoldNm")
    
    Call oForm01.DataSources.UserDataSources.Add("EmpChk", dt_SHORT_TEXT, 10)
    Call oForm01.Items("EmpChk").Specific.DataBind.SetBound(True, "", "EmpChk")
    
    oDocType01 = "작업일보등록(공정)"
    If (oDocType01 = "작업일보등록(작지)") Then
        Call oForm01.Items("DocType").Specific.Select("10", psk_ByValue)
    ElseIf (oDocType01 = "작업일보등록(공정)") Then
        Call oForm01.Items("DocType").Specific.Select("20", psk_ByValue)
    End If
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_PP043_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP043_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP043_ComboBox_Setting()
On Error GoTo PS_PP043_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_PP043", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_PP043", "ItemCode")

'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "Mat02", "LTime", "Y", "Y")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "Mat02", "LTime", "N", "N")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat02.Columns("LTime"), "PS_PP043", "Mat02", "LTime")

    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "10", "일반")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "20", "PSMT지원")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "30", "외주")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "40", "실적")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "50", "일반조정")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "60", "외주조정")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("OrdType").Specific, "PS_PP043", "OrdType")
    
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "DocType", "", "10", "작지기준")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_PP043", "DocType", "", "20", "공정기준")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("DocType").Specific, "PS_PP043", "DocType")
    
    Call oForm01.Items("SOrdGbn").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SOrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN('108','109') order by Code", "", False, False)
    Call oForm01.Items("SBPLId").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
    Call oForm01.Items("OrdGbn").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN('108','109') order by Code", "", False, False)
    Call oForm01.Items("BPLId").Specific.ValidValues.Add("선택", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP043_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP043_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP043_CF_ChooseFromList()
On Error GoTo PS_PP043_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column


    Exit Sub
PS_PP043_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP043_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP043_FormItemEnabled()
On Error GoTo PS_PP043_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("OrdType").Enabled = True
'        oForm01.Items("UseMCode").Enabled = True
'        oForm01.Items("MoldCode").Enabled = False
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("Mat01").Enabled = True
        oForm01.Items("Mat02").Enabled = True
'        oMat02.Columns("NStart").Editable = False  '//비가동시작시간 사용안함
'        oMat02.Columns("NEnd").Editable = False    '//비가동종료시간 사용안함
        oMat02.Columns("NTime").Editable = True    '//비가동시간만 사용
        oForm01.Items("Mat03").Enabled = True
        oForm01.Items("Button01").Enabled = True
        oForm01.Items("1").Enabled = True
        
        Call oForm01.Items("OrdType").Specific.Select(0, psk_Index) '//항상일반타입이여야함
        Call oForm01.Items("OrdGbn").Specific.Select(0, psk_Index) '//선택인상태
        Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        Call oForm01.Items("SOrdGbn").Specific.Select(0, psk_Index)
        Call oForm01.Items("SBPLId").Specific.Select(0, psk_Index)
        oForm01.Items("SCpCode").Specific.VALUE = ""
'        oForm01.Items("UseMCode").Specific.VALUE = ""
'        oForm01.Items("SMoldNo").Specific.VALUE = ""
'        oForm01.Items("MoldCode").Specific.VALUE = ""
        Call PS_PP043_FormClear '//UDO방식
        If (oDocType01 = "작업일보등록(작지)") Then
            Call oForm01.Items("DocType").Specific.Select("10", psk_ByValue)
        ElseIf (oDocType01 = "작업일보등록(공정)") Then
            Call oForm01.Items("DocType").Specific.Select("20", psk_ByValue)
        End If
        oForm01.Items("DocDate").Specific.VALUE = Format(Now - 1, "YYYYMMDD")
        'Call oForm01.Items("SBPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
        Call oForm01.Items("SBPLId").Specific.Select("3", psk_ByValue)
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("Focus").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = True
'        oForm01.Items("UseMCode").Enabled = True
'        oForm01.Items("MoldCode").Enabled = False
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        oForm01.Items("Mat02").Enabled = False
        oForm01.Items("Mat03").Enabled = False
        oForm01.Items("Button01").Enabled = True
        oForm01.Items("1").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        '//각모드에따른 아이템설정
        If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" & Trim(oDS_PS_PP043H.GetValue("DocEntry", 0)) & "'", 0, 1) = "Y" Then
            oForm01.Items("Focus").Click ct_Regular
            oForm01.Items("DocEntry").Enabled = False
'            oForm01.Items("UseMCode").Enabled = False
'            oForm01.Items("MoldCode").Enabled = False
            oForm01.Items("DocDate").Enabled = False
            oForm01.Items("Mat01").Enabled = False
            oForm01.Items("Mat02").Enabled = False
            oForm01.Items("Mat03").Enabled = False
            oForm01.Items("Button01").Enabled = False
            oForm01.Items("1").Enabled = False
        Else
            If Trim(oDS_PS_PP043H.GetValue("U_OrdType", 0)) = "20" Then '//PSMT
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("DocEntry").Enabled = False
'                oForm01.Items("UseMCode").Enabled = False
'                oForm01.Items("MoldCode").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("Mat01").Enabled = False
                oForm01.Items("Mat02").Enabled = False
                oForm01.Items("Mat03").Enabled = False
                oForm01.Items("Button01").Enabled = False
                oForm01.Items("1").Enabled = False
            ElseIf Trim(oDS_PS_PP043H.GetValue("U_OrdType", 0)) = "30" Then '//외주
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("DocEntry").Enabled = False
'                oForm01.Items("UseMCode").Enabled = False
'                oForm01.Items("MoldCode").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("Mat01").Enabled = False
                oForm01.Items("Mat02").Enabled = False
                oForm01.Items("Mat03").Enabled = False
                oForm01.Items("Button01").Enabled = False
                oForm01.Items("1").Enabled = False
            ElseIf Trim(oDS_PS_PP043H.GetValue("U_OrdType", 0)) = "40" Then '//실적
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("DocEntry").Enabled = False
'                oForm01.Items("UseMCode").Enabled = False
'                oForm01.Items("MoldCode").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("Mat01").Enabled = False
                oForm01.Items("Mat02").Enabled = False
                oForm01.Items("Mat03").Enabled = False
                oForm01.Items("Button01").Enabled = False
                oForm01.Items("1").Enabled = False
            ElseIf Trim(oDS_PS_PP043H.GetValue("U_OrdType", 0)) = "10" Or Trim(oDS_PS_PP043H.GetValue("U_OrdType", 0)) = "50" Then '//일반,조정
                If Trim(oDS_PS_PP043H.GetValue("U_OrdGbn", 0)) = "104" Then '//멀티이면
                    Dim NextCpInfo As String
                    Dim i As Long
                    Dim Query01 As String
                    Dim RecordSet01 As SAPbobsCOM.Recordset
                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
                    Query01 = "SELECT "
                    Query01 = Query01 & " PS_PP040H.DocEntry,PS_PP040L.LineId,"
                    Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,"
                    Query01 = Query01 & " PS_PP040L.U_OrdGbn AS OrdGbn,"
                    Query01 = Query01 & " PS_PP040L.U_PP030HNo AS PP030HNo,"
                    Query01 = Query01 & " PS_PP040L.U_PP030MNo AS PP030MNo, "
                    Query01 = Query01 & " PS_PP040L.U_OrdMgNum AS OrdMgNum "
                    Query01 = Query01 & " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry "
                    Query01 = Query01 & " WHERE "
                    Query01 = Query01 & " PS_PP040H.Canceled = 'N' AND PS_PP040L.DocEntry = '" & oDS_PS_PP043H.GetValue("DocEntry", 0) & "'"
                    RecordSet01.DoQuery Query01
                    '//2010.12.08
                    '//2번문서 노대리님이 오픈당시 입력한 기존데이터인 관계로 멀티게이지의  작업지시 문서관계가
                    '//올바르지 못해서 오류가 발생하기때문에 해당문서는 검사하지 않게 처리하였습니다.
                    If oDS_PS_PP043H.GetValue("DocEntry", 0) <> "2" Then
                        For i = 0 To RecordSet01.RecordCount - 1
                            If RecordSet01.Fields("OrdGbn").VALUE = "104" Then '//멀티이면
                                '//마지막공정에 항상 실적포인트가 되어있음
                                If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1) = "Y" Then
                                    '//실적 '//문서의타입필요
                                    If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP080L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
                                        '//비활성상태
                                        GoTo Continue_First
                                    End If
                                End If
                            End If

                            If RecordSet01.Fields("OrdGbn").VALUE = "104" Then '//멀티
                                '//다음공정이 존재하면
                                NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1)
                                If NextCpInfo <> "" Then
                                    If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" & NextCpInfo & "'", 0, 1) > 0 Then
                                        '//다음공정으로 생성된 실적조회 만약 존재한다면 수정불가능
                                        '//비활성상태
                                        GoTo Continue_First
                                    End If
                                Else
                                    '//다음공정이 존재하지 않으면 마지막 공정임, 마지막공정일때는 실적등록여부로 적용여부 판정
                                End If
                            End If
                            RecordSet01.MoveNext
                        Next
                    End If
                    Set RecordSet01 = Nothing
                    GoTo Continue_Second
Continue_First:

                    Dim SuperUserYN As String
                    SuperUserYN = MDC_PS_Common.GetValue("SELECT SUPERUSER FROM OUSR WHERE USER_CODE = '" & Trim(Sbo_Company.UserName) & "'", 0, 1)

                    oForm01.Items("Focus").Click ct_Regular

                    If SuperUserYN = "N" Then

                        oForm01.Items("DocEntry").Enabled = False
'                        oForm01.Items("UseMCode").Enabled = False
'                        oForm01.Items("MoldCode").Enabled = False
                        oForm01.Items("DocDate").Enabled = False
                        oForm01.Items("Mat01").Enabled = False
                        oForm01.Items("Mat02").Enabled = False
                        oForm01.Items("Mat03").Enabled = False
                        oForm01.Items("Button01").Enabled = False

                        oForm01.Items("1").Enabled = False

                    ElseIf SuperUserYN = "Y" Then

                        oForm01.Items("DocEntry").Enabled = False
'                        oForm01.Items("UseMCode").Enabled = False
'                        oForm01.Items("MoldCode").Enabled = False
                        oForm01.Items("DocDate").Enabled = False
                        oForm01.Items("Mat01").Enabled = True
                        oForm01.Items("Mat02").Enabled = False
                        oForm01.Items("Mat03").Enabled = False
                        oForm01.Items("Button01").Enabled = False

                        oForm01.Items("1").Enabled = True

                    End If
                    GoTo Continue_Third
Continue_Second:
                    oForm01.Items("Focus").Click ct_Regular
                    oForm01.Items("DocEntry").Enabled = False
'                    oForm01.Items("UseMCode").Enabled = True
'                    oForm01.Items("MoldCode").Enabled = False
                    oForm01.Items("DocDate").Enabled = True
                    oForm01.Items("Mat01").Enabled = True
                    oForm01.Items("Mat02").Enabled = True
                    oForm01.Items("Mat03").Enabled = True
                    oForm01.Items("Button01").Enabled = True
                    oForm01.Items("1").Enabled = True
                    GoTo Continue_Third
Continue_Third:
                Else
                    oForm01.Items("Focus").Click ct_Regular
                    oForm01.Items("DocEntry").Enabled = False
'                    oForm01.Items("UseMCode").Enabled = True
'                    oForm01.Items("MoldCode").Enabled = False
                    oForm01.Items("DocDate").Enabled = True
                    oForm01.Items("Mat01").Enabled = True
                    oForm01.Items("Mat02").Enabled = True
                    oForm01.Items("Mat03").Enabled = True
                    oForm01.Items("Button01").Enabled = True
                    oForm01.Items("1").Enabled = True
                End If
                '//멀티의경우 수량관계
                oMat01.Columns("BQty").Visible = True
                oMat01.Columns("PSum").Visible = False
                oMat01.Columns("PWeight").Visible = False
                oMat01.Columns("YWeight").Visible = False
                oMat01.Columns("NWeight").Visible = False
            Else '//멀티를 제외한 경우, 엔드베어링
                oForm01.Items("Focus").Click ct_Regular
                oForm01.Items("DocEntry").Enabled = False
'                oForm01.Items("UseMCode").Enabled = True
'                oForm01.Items("MoldCode").Enabled = False
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("Mat01").Enabled = True
                oForm01.Items("Mat02").Enabled = True
                oForm01.Items("Mat03").Enabled = True
                oForm01.Items("Button01").Enabled = True
                oForm01.Items("1").Enabled = True
                '//엔드베어링의경우 수량관계
                oMat01.Columns("BQty").Visible = False
                oMat01.Columns("PSum").Visible = True
                oMat01.Columns("PWeight").Visible = True
                oMat01.Columns("YWeight").Visible = True
                oMat01.Columns("NWeight").Visible = True
            End If
        End If
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP043_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP043_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP043_AddMatrixRow01(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP043_AddMatrixRow01_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP043L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP043L.Offset = oRow
    oDS_PS_PP043L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP043_AddMatrixRow01_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP043_AddMatrixRow01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP043_AddMatrixRow02(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP043_AddMatrixRow02_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP043M.InsertRecord (oRow)
    End If
    oMat02.AddRow
    oDS_PS_PP043M.Offset = oRow
    oDS_PS_PP043M.setValue "U_LineNum", oRow, oRow + 1
    oMat02.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP043_AddMatrixRow02_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP043_AddMatrixRow02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP043_AddMatrixRow03(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP043_AddMatrixRow03_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP043N.InsertRecord (oRow)
    End If
    oMat03.AddRow
    oDS_PS_PP043N.Offset = oRow
    oDS_PS_PP043N.setValue "U_LineNum", oRow, oRow + 1
    oMat03.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP043_AddMatrixRow03_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP043_AddMatrixRow03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP043_FormClear()
On Error GoTo PS_PP043_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP040'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP043_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP043_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP043_EnableMenus()
On Error GoTo PS_PP043_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP043_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP043_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP043_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP043_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP043_FormItemEnabled
        Call PS_PP043_AddMatrixRow01(0, True) '//UDO방식일때
        Call PS_PP043_AddMatrixRow02(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_PP043_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP043_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP043_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_PP043_DataValidCheck() As Boolean
On Error GoTo PS_PP043_DataValidCheck_Error
    PS_PP043_DataValidCheck = False
    Dim i, j As Long
    Dim FailQty As Double
    
   
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP043_FormClear
    End If
    
    '//공수배부
    
    For i = 1 To oMat02.VisualRowCount - 1
        If Val(oMat02.Columns("YTime").Cells(i).Specific.VALUE) <= 0 Then
            Sbo_Application.SetStatusBarMessage "작업시간은 필수입니다.", bmt_Short, True
            oMat02.Columns("YTime").Cells(i).Click ct_Regular
            PS_PP043_DataValidCheck = False
            Exit Function
        End If
    Next
    
    
    
    If oForm01.Items("OrdType").Specific.Selected.VALUE <> "10" And oForm01.Items("OrdType").Specific.Selected.VALUE <> "50" Then
        Sbo_Application.SetStatusBarMessage "작업타입이 일반,조정이 아닙니다.", bmt_Short, True
        PS_PP043_DataValidCheck = False
        Exit Function
    End If
    
    If oForm01.Items("OrdGbn").Specific.Selected.VALUE = "선택" Then
        Sbo_Application.SetStatusBarMessage "작업구분이 선택되지 않았습니다.", bmt_Short, True
        PS_PP043_DataValidCheck = False
        Exit Function
    End If
    
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "공정정보 라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP043_DataValidCheck = False
        Exit Function
    End If
    If oMat02.VisualRowCount = 1 Then
        If Trim(oForm01.Items("OrdGbn").Specific.Selected.VALUE) = "107" Then
            oMat02.FlushToDataSource
            oDS_PS_PP043M.setValue "U_WorkCode", 0, "9999999"
            oDS_PS_PP043M.setValue "U_WorkName", 0, "조정"
            oDS_PS_PP043M.setValue "U_YTime", 0, 1
            Call PS_PP043_AddMatrixRow02(1, False)
            oMat02.LoadFromDataSource
        Else
            Sbo_Application.SetStatusBarMessage "작업자정보 라인이 존재하지 않습니다.", bmt_Short, True
            PS_PP043_DataValidCheck = False
            Call oMat02.SelectRow(oMat02.VisualRowCount, True, False)
            Exit Function
        End If
    End If
    If oMat03.VisualRowCount = 0 Then
        Sbo_Application.SetStatusBarMessage "불량정보 라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP043_DataValidCheck = False
        Exit Function
    End If
'    If oMat02.VisualRowCount > 2 Then '//멀티,엔드베어링 작업자 1명 입력가능
'        Sbo_Application.SetStatusBarMessage "작업자정보 한명만 입력할수 있습니다.", bmt_Short, True
'        PS_PP043_DataValidCheck = False
'        Exit Function
'    End If
    
    
    
    For i = 1 To oMat01.VisualRowCount - 1
        If (oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "작지문서번호는 필수입니다.", bmt_Short, True
            oMat01.Columns("OrdMgNum").Cells(i).Click ct_Regular
            PS_PP043_DataValidCheck = False
            Exit Function
        End If
'        If (Val(oMat01.Columns("YQty").Cells(i).Specific.VALUE) + Val(oMat01.Columns("NQty").Cells(i).Specific.VALUE) + Val(oMat01.Columns("ScrapWt").Cells(i).Specific.VALUE) <= 0) Then
'            Sbo_Application.SetStatusBarMessage "생산수량또는 불량수량 또는 스크랩이 없습니다.", bmt_Short, True
'            oMat01.Columns("YQty").Cells(i).Click ct_Regular
'            PS_PP043_DataValidCheck = False
'            Exit Function
'        End If
        '// 엔드베어링은 생산수량 = 합격수량 + 불량수량 이어야 한다.
        If Trim(oForm01.Items("SOrdGbn").Specific.VALUE) = "107" Then
            If Val(oMat01.Columns("PQty").Cells(i).Specific.VALUE) <> Val(oMat01.Columns("YQty").Cells(i).Specific.VALUE) + Val(oMat01.Columns("NQty").Cells(i).Specific.VALUE) Then
                Sbo_Application.SetStatusBarMessage "생산수량 = 합격수량 + 불량수량 이어야 합니다.", bmt_Short, True
                oMat01.Columns("YQty").Cells(i).Click ct_Regular
                PS_PP043_DataValidCheck = False
                Exit Function
            End If
        End If
        
        If (Val(oMat01.Columns("WorkTime").Cells(i).Specific.VALUE) <= 0) Then
            If Trim(oForm01.Items("SOrdGbn").Specific.VALUE) = "104" And Trim(oForm01.Items("SCpCode").Specific.VALUE) <> "CP50107" Then '//멀티 포장공정외는 공수필수
                Sbo_Application.SetStatusBarMessage "실동시간은 필수입니다.", bmt_Short, True
                oMat01.Columns("WorkTime").Cells(i).Click ct_Regular
                PS_PP043_DataValidCheck = False
                Exit Function
            End If
        End If
        
        FailQty = 0
        For j = 1 To oMat03.VisualRowCount
            '//불량코드를 입력했는지 check
            If Val(oMat03.Columns("FailQty").Cells(j).Specific.VALUE) <> 0 And oMat03.Columns("FailCode").Cells(j).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "불량수량이 입력되었을 때는 불량코드는 필수입니다.", bmt_Short, True
                oMat03.Columns("FailCode").Cells(i).Click ct_Regular
                PS_PP043_DataValidCheck = False
                Exit Function
            End If
            
'            If Val(oMat03.Columns("FailQty").Cells(j).Specific.VALUE) = 0 And oMat03.Columns("FailCode").Cells(j).Specific.VALUE <> "" Then
'                Sbo_Application.SetStatusBarMessage "불량코드를 확인하세요.", bmt_Short, True
'                oMat03.Columns("FailCode").Cells(i).Click ct_Regular
'                PS_PP043_DataValidCheck = False
'                Exit Function
'            End If
            
            If (oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE = oMat03.Columns("OrdMgNum").Cells(j).Specific.VALUE) Then
                FailQty = FailQty + Val(oMat03.Columns("FailQty").Cells(j).Specific.VALUE)
            End If
        Next
        If Val(oMat01.Columns("NQty").Cells(i).Specific.VALUE) <> FailQty Then
            Sbo_Application.SetStatusBarMessage "공정리스트의 불량수량과 불량정보의 불량수량이 일치하지 않습니다.", bmt_Short, True
            PS_PP043_DataValidCheck = False
            Exit Function
        End If
        '//불량수량 검사
    Next
    
    '비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_S
    For i = 1 To oMat02.VisualRowCount - 1
    
        If (oMat02.Columns("NCode").Cells(i).Specific.VALUE <> "") Then
            
            If oMat02.Columns("NTime").Cells(i).Specific.VALUE = "" Then
            
                Sbo_Application.SetStatusBarMessage "비가동코드가 입력되었을 때는 비가동시간은 필수입니다.", bmt_Short, True
                oMat02.Columns("NTime").Cells(i).Click ct_Regular
                PS_PP043_DataValidCheck = False
                Exit Function
            
            End If
        
        End If
        
        If (oMat02.Columns("NTime").Cells(i).Specific.VALUE <> "") Then
            
            If oMat02.Columns("NCode").Cells(i).Specific.VALUE = "" Then
            
                Sbo_Application.SetStatusBarMessage "비가동시간이 입력되었을 때는 비가동코드는 필수입니다.", bmt_Short, True
                oMat02.Columns("NCode").Cells(i).Click ct_Regular
                PS_PP043_DataValidCheck = False
                Exit Function
            
            End If
        
        End If
    
    Next
    '비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_E
    
    '불량정보 입력 체크_S (2012.03.20 송명규 추가)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim SuperUserYN As String
    
    SuperUserYN = MDC_PS_Common.GetValue("SELECT SUPERUSER FROM OUSR WHERE USER_CODE = '" & Trim(Sbo_Company.UserName) & "'", 0, 1)
    
    If SuperUserYN = "N" Then '슈퍼유저가 아니면 불령정보 필수 입력
    
        For i = 1 To oMat03.VisualRowCount - 1
        
            '해당 작업지시의 재작업 여부 조회
            oQuery01 = "SELECT U_ReWorkYN FROM [@PS_PP030M] WHERE Convert(Nvarchar(50),DocEntry) +" & "'-'" & "+ Convert(Nvarchar(20),U_LineId) = '" & oMat03.Columns("OrdMgNum").Cells(i).Specific.VALUE & "'"
            Call oRecordSet01.DoQuery(oQuery01)
            
            If oRecordSet01.Fields("U_ReWorkYN").VALUE = "Y" Then
            
                '불량코드
                If (oMat03.Columns("FailCode").Cells(i).Specific.VALUE = "") Then
                    Sbo_Application.SetStatusBarMessage "재작업 시 불량정보는 필수입니다. 확인하십시오.", bmt_Short, True
                    oMat03.Columns("FailCode").Cells(i).Click ct_Regular
                    PS_PP043_DataValidCheck = False
                    Exit Function
                End If
                
                '원인공정코드
                If (oMat03.Columns("CsCpCode").Cells(i).Specific.VALUE = "") Then
                    Sbo_Application.SetStatusBarMessage "재작업 시 원인공정정보는 필수입니다. 확인하십시오.", bmt_Short, True
                    oMat03.Columns("CsCpCode").Cells(i).Click ct_Regular
                    PS_PP043_DataValidCheck = False
                    Exit Function
                End If
                
                '작업자코드
                If (oMat03.Columns("CsWkCode").Cells(i).Specific.VALUE = "") Then
                    Sbo_Application.SetStatusBarMessage "재작업 시 작업자정보는 필수입니다. 확인하십시오.", bmt_Short, True
                    oMat03.Columns("CsWkCode").Cells(i).Click ct_Regular
                    PS_PP043_DataValidCheck = False
                    Exit Function
                End If
                
            End If
    
        Next
        
    Else

    End If
    
    Set oRecordSet01 = Nothing
    '불량정보 입력 체크_E (2012.03.20 송명규 추가)
    
    
    If (PS_PP043_Validate("검사01") = False) Then
        PS_PP043_DataValidCheck = False
        Exit Function
    End If
    
    Call oDS_PS_PP043L.RemoveRecord(oDS_PS_PP043L.Size - 1)
    Call oMat01.LoadFromDataSource
    Call oDS_PS_PP043M.RemoveRecord(oDS_PS_PP043M.Size - 1)
    Call oMat02.LoadFromDataSource
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP043_FormClear
    End If
    PS_PP043_DataValidCheck = True
    Exit Function
PS_PP043_DataValidCheck_Error:
    PS_PP043_DataValidCheck = False
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_PP043_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP043_MTX01()
On Error GoTo PS_PP043_MTX01_Error
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
        GoTo PS_PP043_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_PP043L.InsertRecord (i)
        End If
        oDS_PS_PP043L.Offset = i
        Call oDS_PS_PP043L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_PP043L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_PP043_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_PP043_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP043_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP043_DI_API() As Boolean
On Error GoTo PS_PP043_DI_API_Error
    PS_PP043_DI_API = True
    Dim i, j As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim LineNumCount As Long
    Dim ResultDocNum As Long
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    ReDim ItemInformation(0)
    ItemInformationCount = 0
    For i = 1 To oMat01.VisualRowCount
        ReDim Preserve ItemInformation(ItemInformationCount)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).LotNo = oMat01.Columns("LotNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = oForm01.Items("BPLId").Specific.Selected.VALUE
    oDIObject.CardCode = oForm01.Items("CardCode").Specific.VALUE
    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.VALUE, "&&&&-&&-&&")
    For i = 0 To ItemInformationCount - 1
        If ItemInformation(i).Check = True Then
            GoTo Continue_First
        End If
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = oForm01.Items("WhsCode").Specific.VALUE
        oDIObject.Lines.BaseType = "22"
        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
        For j = i To UBound(ItemInformation)
            If ItemInformation(j).Check = True Then
                GoTo Continue_Second
            End If
            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
                GoTo Continue_Second
            End If
            '//같은것
            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).LotNo
            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.Add
            ItemInformation(j).PDN1No = LineNumCount
            ItemInformation(j).Check = True
Continue_Second:
        Next
        LineNumCount = LineNumCount + 1
Continue_First:
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To UBound(ItemInformation)
            Call oDS_PS_PP043L.setValue("U_OPDNNo", i, ResultDocNum)
            Call oDS_PS_PP043L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        Next
    Else
        GoTo PS_PP043_DI_API_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_PP043_DI_API_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_PP043_DI_API = False
    Set oDIObject = Nothing
    Exit Function
PS_PP043_DI_API_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP043_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_PP043_DI_API = False
End Function

Private Sub PS_PP043_FormResize()
On Error GoTo PS_PP043_FormResize_Error
    oForm01.Items("Mat02").Top = 170
    oForm01.Items("Mat02").Left = 7
    oForm01.Items("Mat02").Height = ((oForm01.Height - 170) / 3 * 1) - 20
    oForm01.Items("Mat02").Width = oForm01.Width / 2 - 14
    
    oForm01.Items("Mat03").Top = 170
    oForm01.Items("Mat03").Left = oForm01.Width / 2
    oForm01.Items("Mat03").Height = ((oForm01.Height - 170) / 3 * 1) - 20
    oForm01.Items("Mat03").Width = oForm01.Width / 2 - 14
    
    oForm01.Items("Mat01").Top = oForm01.Items("Mat03").Top + oForm01.Items("Mat03").Height + 40
    oForm01.Items("Mat01").Left = 7
    oForm01.Items("Mat01").Height = ((oForm01.Height - 170) / 3 * 2) - 80
    oForm01.Items("Mat01").Width = oForm01.Width - 21
    
    oForm01.Items("Opt01").Left = 10
    oForm01.Items("Opt02").Left = oForm01.Width / 2
    oForm01.Items("Opt03").Left = 10
    oForm01.Items("Opt03").Top = oForm01.Items("Mat03").Top + oForm01.Items("Mat03").Height + 20
    oForm01.Items("Button02").Left = 365
    oForm01.Items("Button02").Top = oForm01.Items("Opt03").Top 'oForm01.Items("Mat03").Top + oForm01.Items("Mat03").Height + 20
    Exit Sub
PS_PP043_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP043_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP043_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_PP043_Validate_Error
    PS_PP043_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    Dim PrevDBCpQty As Long
    Dim PrevMATRIXCpQty As Long
    Dim CurrentDBCpQty As Long
    Dim CurrentMATRIXCpQty As Long
    Dim NextDBCpQty As Long
    Dim NextMATRIXCpQty As Long
    Dim PrevCpInfo As String
    Dim CurrentCpInfo As String
    Dim NextCpInfo As String
    
    Dim OrdMgNum As String
    Dim Exist As Boolean
        
    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
        PS_PP043_Validate = False
        GoTo PS_PP043_Validate_Exit
    End If
        
    If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 일반,조정인경우
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT지원인경우
        MDC_Com.MDC_GF_Message "해당작업타입은 변경이 불가능합니다.", "W"
        PS_PP043_Validate = False
        GoTo PS_PP043_Validate_Exit
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
        MDC_Com.MDC_GF_Message "해당작업타입은 변경이 불가능합니다.", "W"
        PS_PP043_Validate = False
        GoTo PS_PP043_Validate_Exit
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
        MDC_Com.MDC_GF_Message "해당작업타입은 변경이 불가능합니다.", "W"
        PS_PP043_Validate = False
        GoTo PS_PP043_Validate_Exit
    End If
        
    If ValidateType = "검사01" Then
        For i = 1 To oMat01.VisualRowCount - 1 '//입력된 행에 대해
            If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [OITM] WHERE ItemCode = '" & oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE & "'", 0, 1) <= 0 Then
                Call MDC_Com.MDC_GF_Message("제품코드가 존재하지 않습니다.", "W")
                PS_PP043_Validate = False
                GoTo PS_PP043_Validate_Exit
            End If
        Next
        If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Then '//작업타입이 일반인경우
'            '//삭제된 행에 대한처리
'            Query01 = "SELECT "
'            Query01 = Query01 & " PS_PP040H.DocEntry,PS_PP040L.LineId,"
'            Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,"
'            Query01 = Query01 & " PS_PP040L.U_OrdGbn AS OrdGbn,"
'            Query01 = Query01 & " PS_PP040L.U_PP030HNo AS PP030HNo,"
'            Query01 = Query01 & " PS_PP040L.U_PP030MNo AS PP030MNo, "
'            Query01 = Query01 & " PS_PP040L.U_OrdMgNum AS OrdMgNum "
'            Query01 = Query01 & " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry "
'            Query01 = Query01 & " WHERE "
'            Query01 = Query01 & " PS_PP040H.Canceled = 'N' AND PS_PP040L.DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
'            RecordSet01.DoQuery Query01
'            For i = 0 To RecordSet01.RecordCount - 1
'                Exist = False
'                For j = 1 To oMat01.VisualRowCount - 1 '//기존에 있는 행에대한처리
'                    If (oMat01.Columns("LineId").Cells(j).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
'                    Else
'                        '//라인번호가 같고, 문서번호가 같으면 존재하는행
'                        If Val(RecordSet01.Fields(0).VALUE) = Val(oForm01.Items("DocEntry").Specific.VALUE) And Val(RecordSet01.Fields(1).VALUE) = Val(oMat01.Columns("LineId").Cells(j).Specific.VALUE) Then
'                            Exist = True
'                        End If
'                    End If
'                Next
'                If Exist = False Then '//삭제된 행중 수량관계를 알아본다.
'                    If RecordSet01.Fields("OrdGbn").VALUE = "107" Then '//엔트베어링이면서
'                        '//실적포인트가 체크된 공정인경우..
'                        If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1) = "Y" Then
'                            '//실적 '//문서의타입필요
'                            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP080L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
'                                MDC_Com.MDC_GF_Message "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
'                                PS_PP043_Validate = False
'                                GoTo PS_PP043_Validate_Exit
'                            End If
'                        End If
'                    ElseIf RecordSet01.Fields("OrdGbn").VALUE = "104" Then '//멀티이면
'                        '//마지막공정에 항상 실적포인트가 되어있음
'                        If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1) = "Y" Then
'                            '//실적 '//문서의타입필요
'                            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP080L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
'                                MDC_Com.MDC_GF_Message "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
'                                PS_PP043_Validate = False
'                                GoTo PS_PP043_Validate_Exit
'                            End If
'                        End If
'                    End If
'
'                    If RecordSet01.Fields("OrdGbn").VALUE = "104" Then '//멀티
'                        '//다음공정이 존재하면
'                        NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1)
'                        If NextCpInfo <> "" Then
'                            If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" & NextCpInfo & "'", 0, 1) > 0 Then
'                                '//다음공정으로 생성된 실적조회 만약 존재한다면 취소불가능
'                                MDC_Com.MDC_GF_Message "후행공정이 입력된 문서입니다.", "W"
'                                PS_PP043_Validate = False
'                                GoTo PS_PP043_Validate_Exit
'                            End If
'                        Else
'                            '//다음공정이 존재하지 않으면 마지막 공정임, 마지막공정일때는 실적등록여부로 적용여부 판정
'                        End If
'                    ElseIf RecordSet01.Fields("OrdGbn").VALUE = "107" Then '//엔드베어링
'                        '//삭제된 행에 대한 검사..
'                        OrdMgNum = RecordSet01.Fields("OrdMgNum").VALUE '// DocEntry + '-' + LineId
'                        CurrentCpInfo = OrdMgNum
'
'                        PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_02 '" & OrdMgNum & "'")
'                        If PrevCpInfo = "" Then
'                            '//해당공정이 첫공정이면 입력되어도 상관없다.
'                        Else
'                            PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & PrevCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040H.Canceled = 'N'")
'                            PrevMATRIXCpQty = 0
'                            For j = 1 To oMat01.VisualRowCount - 1
'                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = PrevCpInfo) Then
'                                    PrevMATRIXCpQty = PrevMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
'                                End If
'                            Next
'                            CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040H.Canceled = 'N'")
'                            CurrentMATRIXCpQty = 0
'                            For j = 1 To oMat01.VisualRowCount - 1
'                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = CurrentCpInfo) Then
'                                    CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
'                                End If
'                            Next
'                            If ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
'                                Sbo_Application.SetStatusBarMessage "삭제된 공정의 선행공정의 생산수량이 삭제된 공정의 생산수량을 미달합니다.", bmt_Short, True
'                                PS_PP043_Validate = False
'                                GoTo PS_PP043_Validate_Exit
'                            End If
'                        End If
'
''                       If oForm01.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
''                            NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & OrdMgNum & "'")
''                            If NextCpInfo = "" Then
''                                '//해당공정이 마지막공정이면 삭제되어도 상관없다.
''                            Else
''                                NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
''                                NextMATRIXCpQty = 0
''                                For j = 1 To oMat01.VisualRowCount - 1
''                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = NextCpInfo) Then
''                                        NextMATRIXCpQty = NextMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
''                                    End If
''                                Next
''                                CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
''                                CurrentMATRIXCpQty = 0
''                                For j = 1 To oMat01.VisualRowCount - 1 '//현재공정은 삭제되었으므로.. 매트릭스에 존재하지 않는다.
''                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = CurrentCpInfo) Then
''                                        CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
''                                    End If
''                                Next
''                                If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
''                                    Sbo_Application.SetStatusBarMessage "삭제된 공정의 후행공정의 생산수량이 삭제된 공정의 생산수량을 초과합니다.", bmt_Short, True
''                                    PS_PP043_Validate = False
''                                    GoTo PS_PP043_Validate_Exit
''                                End If
''                            End If
''                        End If
'                    End If
'                End If
'                RecordSet01.MoveNext
'            Next
            
            
'            For i = 1 To oMat01.VisualRowCount - 1
'                If (oMat01.Columns("LineId").Cells(i).Specific.VALUE = "") Then '//새로추가된 행인경우, 검사할필요없다
'                Else
'                    If oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE = "104" Then '//멀티
'                        '//수량이 수정되면 뒷공정이 존재한다면 수정할수 없다.
'                        NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE & "'", 0, 1)
'                        If NextCpInfo <> "" Then
'                            If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" & NextCpInfo & "'", 0, 1) > 0 Then
'                                '//다음공정으로 생성된 실적조회 만약 존재한다면 취소불가능
'                                '//작업일보등록된문서중에 수정이 된문서를 구함
'                                Query01 = "SELECT "
'                                Query01 = Query01 & " PS_PP040L.U_OrdMgNum,"
'                                Query01 = Query01 & " PS_PP040L.U_Sequence,"
'                                Query01 = Query01 & " PS_PP040L.U_CpCode,"
'                                Query01 = Query01 & " PS_PP040L.U_ItemCode,"
'                                Query01 = Query01 & " PS_PP040L.U_PP030HNo,"
'                                Query01 = Query01 & " PS_PP040L.U_PP030MNo,"
'                                Query01 = Query01 & " PS_PP040L.U_PQty,"
'                                Query01 = Query01 & " PS_PP040L.U_NQty,"
'                                Query01 = Query01 & " PS_PP040L.U_ScrapWt"
''                                Query01 = Query01 & " PS_PP040L.U_WorkTime" '작업시간은 수정이 가능하도록 체크 기능에서 제외(2012.06.18 송명규)
'                                Query01 = Query01 & " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry"
'                                Query01 = Query01 & " WHERE PS_PP040H.DocEntry = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "'"
'                                Query01 = Query01 & " AND PS_PP040L.LineId = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'"
'                                Query01 = Query01 & " AND PS_PP040H.Canceled = 'N'"
'                                RecordSet01.DoQuery Query01
'                                '작업시간은 수정이 가능하도록 체크 기능에서 제외(2012.06.18 송명규)
''                                If (RecordSet01.Fields(0).VALUE = oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE _
''                                And RecordSet01.Fields(1).VALUE = oMat01.Columns("Sequence").Cells(i).Specific.VALUE _
''                                And RecordSet01.Fields(2).VALUE = oMat01.Columns("CpCode").Cells(i).Specific.VALUE _
''                                And RecordSet01.Fields(3).VALUE = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE _
''                                And RecordSet01.Fields(4).VALUE = oMat01.Columns("PP030HNo").Cells(i).Specific.VALUE _
''                                And RecordSet01.Fields(5).VALUE = oMat01.Columns("PP030MNo").Cells(i).Specific.VALUE _
''                                And RecordSet01.Fields(6).VALUE = oMat01.Columns("PQty").Cells(i).Specific.VALUE _
''                                And RecordSet01.Fields(7).VALUE = oMat01.Columns("NQty").Cells(i).Specific.VALUE _
''                                And RecordSet01.Fields(8).VALUE = oMat01.Columns("ScrapWt").Cells(i).Specific.VALUE _
''                                And RecordSet01.Fields(9).VALUE = oMat01.Columns("WorkTime").Cells(i).Specific.VALUE _
''                                ) Then
'                                If (Trim(RecordSet01.Fields(0).VALUE) = oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE _
'                                And Trim(RecordSet01.Fields(1).VALUE) = oMat01.Columns("Sequence").Cells(i).Specific.VALUE _
'                                And Trim(RecordSet01.Fields(2).VALUE) = oMat01.Columns("CpCode").Cells(i).Specific.VALUE _
'                                And Trim(RecordSet01.Fields(3).VALUE) = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE _
'                                And Trim(RecordSet01.Fields(4).VALUE) = oMat01.Columns("PP030HNo").Cells(i).Specific.VALUE _
'                                And Trim(RecordSet01.Fields(5).VALUE) = oMat01.Columns("PP030MNo").Cells(i).Specific.VALUE _
'                                And CDbl(Trim(RecordSet01.Fields(6).VALUE)) = oMat01.Columns("PQty").Cells(i).Specific.VALUE _
'                                And CDbl(Trim(RecordSet01.Fields(7).VALUE)) = oMat01.Columns("NQty").Cells(i).Specific.VALUE _
'                                And CDbl(Trim(RecordSet01.Fields(8).VALUE)) = oMat01.Columns("ScrapWt").Cells(i).Specific.VALUE _
'                                ) Then
'                                Else '//값이 변경된 행의경우
'                                    MDC_Com.MDC_GF_Message "후행공정이 입력된 문서입니다. 수정할수 없습니다.", "W"
'                                    PS_PP043_Validate = False
'                                    GoTo PS_PP043_Validate_Exit
'                                End If
'                            End If
'                        Else
'                            '//다음공정이 존재하지 않으면 마지막 공정임, 마지막공정일때는 실적등록여부로 적용여부 판정
'                        End If
'
'                        '//실적포인트면
'                        If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "-" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1) = "Y" Then  '//현재 공정이 바렐 앞공정이면..
'                            '//실적테이블에 값이 존재하는지 검사
'                            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP080L.U_PP030MNo = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 Then
'                                '//작업일보등록된문서중에 수정이 된문서를 구함
'                                Query01 = "SELECT "
'                                Query01 = Query01 & " PS_PP040L.U_OrdMgNum,"
'                                Query01 = Query01 & " PS_PP040L.U_Sequence,"
'                                Query01 = Query01 & " PS_PP040L.U_CpCode,"
'                                Query01 = Query01 & " PS_PP040L.U_ItemCode,"
'                                Query01 = Query01 & " PS_PP040L.U_PP030HNo,"
'                                Query01 = Query01 & " PS_PP040L.U_PP030MNo,"
'                                Query01 = Query01 & " PS_PP040L.U_PQty,"
'                                Query01 = Query01 & " PS_PP040L.U_NQty,"
'                                Query01 = Query01 & " PS_PP040L.U_ScrapWt,"
'                                Query01 = Query01 & " PS_PP040L.U_WorkTime"
'                                Query01 = Query01 & " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry"
'                                Query01 = Query01 & " WHERE PS_PP040H.DocEntry = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "'"
'                                Query01 = Query01 & " AND PS_PP040L.LineId = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'"
'                                Query01 = Query01 & " AND PS_PP040H.Canceled = 'N'"
'                                RecordSet01.DoQuery Query01
'                                If (RecordSet01.Fields(0).VALUE = oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(1).VALUE = oMat01.Columns("Sequence").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(2).VALUE = oMat01.Columns("CpCode").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(3).VALUE = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(4).VALUE = oMat01.Columns("PP030HNo").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(5).VALUE = oMat01.Columns("PP030MNo").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(6).VALUE = oMat01.Columns("PQty").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(7).VALUE = oMat01.Columns("NQty").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(8).VALUE = oMat01.Columns("ScrapWt").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(9).VALUE = oMat01.Columns("WorkTime").Cells(i).Specific.VALUE _
'                                ) Then
'                                Else '//값이 변경된 행의경우
'                                    MDC_Com.MDC_GF_Message "생산실적이 등록된 행은 수정할수 없습니다.", "W"
'                                    PS_PP043_Validate = False
'                                    GoTo PS_PP043_Validate_Exit
'                                End If
'                            End If
'                        End If
'                    ElseIf oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE = "107" Then '//엔드베어링이면서
'                        '//실적포인트면
'                        If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "-" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1) = "Y" Then  '//현재 공정이 바렐 앞공정이면..
'                            '//휘팅벌크포장,휘팅실적
'                            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP070L.U_PP030MNo = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 _
'                            Or (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND PS_PP080L.U_PP030MNo = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'", 0, 1)) > 0 Then
'                                '//작업일보등록된문서중에 수정이 된문서를 구함
'                                Query01 = "SELECT "
'                                Query01 = Query01 & " PS_PP040L.U_OrdMgNum,"
'                                Query01 = Query01 & " PS_PP040L.U_Sequence,"
'                                Query01 = Query01 & " PS_PP040L.U_CpCode,"
'                                Query01 = Query01 & " PS_PP040L.U_ItemCode,"
'                                Query01 = Query01 & " PS_PP040L.U_PP030HNo,"
'                                Query01 = Query01 & " PS_PP040L.U_PP030MNo,"
'                                Query01 = Query01 & " PS_PP040L.U_PQty,"
'                                Query01 = Query01 & " PS_PP040L.U_NQty,"
'                                Query01 = Query01 & " PS_PP040L.U_ScrapWt,"
'                                Query01 = Query01 & " PS_PP040L.U_WorkTime"
'                                Query01 = Query01 & " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry"
'                                Query01 = Query01 & " WHERE PS_PP040H.DocEntry = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "'"
'                                Query01 = Query01 & " AND PS_PP040L.LineId = '" & Val(oMat01.Columns("LineId").Cells(i).Specific.VALUE) & "'"
'                                Query01 = Query01 & " AND PS_PP040H.Canceled = 'N'"
'                                RecordSet01.DoQuery Query01
'                                If (RecordSet01.Fields(0).VALUE = oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(1).VALUE = oMat01.Columns("Sequence").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(2).VALUE = oMat01.Columns("CpCode").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(3).VALUE = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(4).VALUE = oMat01.Columns("PP030HNo").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(5).VALUE = oMat01.Columns("PP030MNo").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(6).VALUE = oMat01.Columns("PQty").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(7).VALUE = oMat01.Columns("NQty").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(8).VALUE = oMat01.Columns("ScrapWt").Cells(i).Specific.VALUE _
'                                And RecordSet01.Fields(9).VALUE = oMat01.Columns("WorkTime").Cells(i).Specific.VALUE _
'                                ) Then
'                                Else '//값이 변경된 행의경우
'                                    MDC_Com.MDC_GF_Message "생산실적이 등록된 행은 수정할수 없습니다.", "W"
'                                    PS_PP043_Validate = False
'                                    GoTo PS_PP043_Validate_Exit
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'            Next
            
'            For i = 1 To oMat01.VisualRowCount - 1 '//입력된 모든행에 대해 입력가능성 검사
'                If oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE = "104" Then '//멀티
'                    '//상단에서 후행공정입력여부검사 및 실적등록여부 검사 수정되었다면 앞공정보다 많게는 입력할수 없다(Validate에서처리)
'                ElseIf oMat01.Columns("OrdGbn").Cells(i).Specific.VALUE = "107" Then '//엔드베어링
'                    OrdMgNum = oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE
'                    CurrentCpInfo = oMat01.Columns("OrdMgNum").Cells(i).Specific.VALUE
'
'                    PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_02 '" & OrdMgNum & "'")
'                    If PrevCpInfo = "" Then
'                        '//해당공정이 첫공정이면 입력되어도 상관없다.
'                    Else
'                        PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & PrevCpInfo & "' AND PS_PP040H.DocEntry <> '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND PS_PP040H.Canceled = 'N'")
'                        PrevMATRIXCpQty = 0
'                        For j = 1 To oMat01.VisualRowCount - 1
'                            If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = PrevCpInfo) Then
'                                PrevMATRIXCpQty = PrevMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
'                            End If
'                        Next
'                        CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & oForm01.Items("DocEntry").Specific.VALUE & "' AND PS_PP040H.Canceled = 'N'")
'                        CurrentMATRIXCpQty = 0
'                        For j = 1 To oMat01.VisualRowCount - 1
'                            If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = CurrentCpInfo) Then
'                                CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
'                            End If
'                        Next
'                        If ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
'                            Sbo_Application.SetStatusBarMessage "선행공정의 생산수량이 현공정의 생산수량에 미달 합니다.", bmt_Short, True
'                            Call oMat01.SelectRow(i, True, False)
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
'
''                   If oForm01.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
''                        NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & OrdMgNum & "'")
''                        If NextCpInfo = "" Then
''                            '//해당공정이 마지막공정이면 삭제되어도 상관없다.
''                        Else
''                            NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP040H.DocEntry <> '" & oForm01.Items("DocEntry").Specific.Value & "' AND PS_PP040H.Canceled = 'N'")
''                            NextMATRIXCpQty = 0
''                            For j = 1 To oMat01.VisualRowCount - 1
''                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = NextCpInfo) Then
''                                    NextMATRIXCpQty = NextMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
''                                End If
''                            Next
''                            CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & oForm01.Items("DocEntry").Specific.Value & "' AND PS_PP040H.Canceled = 'N'")
''                            CurrentMATRIXCpQty = 0
''                            For j = 1 To oMat01.VisualRowCount - 1 '//현재공정은 삭제되었으므로.. 매트릭스에 존재하지 않는다.
''                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = CurrentCpInfo) Then
''                                    CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
''                                End If
''                            Next
''                            If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
''                                Sbo_Application.SetStatusBarMessage "후행공정의 생산수량이 현공정의 생산수량을 초과 합니다.", bmt_Short, True
''                                PS_PP043_Validate = False
''                                GoTo PS_PP043_Validate_Exit
''                            End If
''                        End If
''                    End If
'                End If
'            Next
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 조정인경우
        End If
    ElseIf ValidateType = "행삭제01" Then
        If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Then '//작업타입이 일반인경우
            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_MM095H] a Inner JOIN [@PS_MM095L] b ON a.DocEntry = b.DocEntry WHERE a.Canceled = 'N' AND b.U_PP040Doc = '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "' AND b.U_PP040Lin = '" & Val(oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
                MDC_Com.MDC_GF_Message "원재료 불출된 행입니다. 행 삭제할 수 없습니다. 원재료 불출 취소 후 행 삭제 바랍니다.", "W"
                PS_PP043_Validate = False
                GoTo PS_PP043_Validate_Exit
            End If

            '//행삭제전 행삭제가능여부검사
'            If (oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
'            Else
'                If oMat01.Columns("OrdGbn").Cells(oMat01Row01).Specific.VALUE = "107" Then '//엔드베어링이면서
'                    '//실적포인트이면
'                    If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "-" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1) = "Y" Then
'                        '//휘팅벌크포장
'                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "' AND PS_PP070L.U_PP030MNo = '" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
'                            MDC_Com.MDC_GF_Message "생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                        '//휘팅실적
'                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "' AND PS_PP080L.U_PP030MNo = '" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
'                            MDC_Com.MDC_GF_Message "생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
'                ElseIf oMat01.Columns("OrdGbn").Cells(oMat01Row01).Specific.VALUE = "104" Then '//멀티이면서
'                    '//실적포인트이면
'                    If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "-" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1) = "Y" Then
'                        '//실적
'                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "' AND PS_PP080L.U_PP030MNo = '" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
'                            MDC_Com.MDC_GF_Message "생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
'                    '//뒷공정이 존재한다면 수정할수 없다.
'                    NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & oMat01.Columns("OrdMgNum").Cells(oMat01Row01).Specific.VALUE & "'", 0, 1)
'                    If NextCpInfo <> "" Then
'                        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" & NextCpInfo & "'", 0, 1) > 0 Then
'                            '//다음공정으로 생성된 실적조회 만약 존재한다면 취소불가능
'                            MDC_Com.MDC_GF_Message "후행공정이 존재합니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
'                End If
'            End If
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 조정인경우
        End If
    ElseIf ValidateType = "수정01" Then
        If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Then '//작업타입이 일반인경우
            '//수정전 수정가능여부검사
'            If (oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE = "") Then '//새로추가된 행인경우, 수정하여도 무방하다
'            Else
'                If oMat01.Columns("OrdGbn").Cells(oMat01Row01).Specific.VALUE = "107" Then '//엔드베어링이면서
'                    '//실적포인트이면
'                    If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "-" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1) = "Y" Then
'                        '//휘팅벌크포장
'                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "' AND PS_PP070L.U_PP030MNo = '" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
'                            MDC_Com.MDC_GF_Message "생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                        '//휘팅실적
'                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & Val(oMat01.Columns("PP030HNo").Cells(oMat01Row01).Specific.VALUE) & "' AND PS_PP080L.U_PP030MNo = '" & Val(oMat01.Columns("PP030MNo").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1)) > 0 Then
'                            MDC_Com.MDC_GF_Message "생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
'                ElseIf oMat01.Columns("OrdGbn").Cells(oMat01Row01).Specific.VALUE = "104" Then '//멀티이면서
'                    '//실적포인트이면
'                    If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & Val(oForm01.Items("DocEntry").Specific.VALUE) & "-" & Val(oMat01.Columns("LineId").Cells(oMat01Row01).Specific.VALUE) & "'", 0, 1) = "Y" Then
'                        '//실적
'                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP080L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
'                            MDC_Com.MDC_GF_Message "생산실적 등록된 행입니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
'
'                    '//뒷공정이 존재한다면 수정할수 없다.
'                    NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & oMat01.Columns("OrdMgNum").Cells(oMat01Row01).Specific.VALUE & "'", 0, 1)
'                    If NextCpInfo <> "" Then
'                        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" & NextCpInfo & "'", 0, 1) > 0 Then
'                            '//다음공정으로 생성된 실적조회 만약 존재한다면 취소불가능
'                            MDC_Com.MDC_GF_Message "후행공정이 존재합니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
'                End If
'            End If
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 조정인경우
        End If
    ElseIf ValidateType = "취소" Then
        If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
            MDC_Com.MDC_GF_Message "이미취소된 문서 입니다. 취소할수 없습니다.", "W"
            PS_PP043_Validate = False
            GoTo PS_PP043_Validate_Exit
        End If
        
        If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Then '//작업타입이 일반인경우
'            '//삭제된 행에 대한처리
'            Query01 = "SELECT "
'            Query01 = Query01 & " PS_PP040H.DocEntry,PS_PP040L.LineId,"
'            Query01 = Query01 & " CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,"
'            Query01 = Query01 & " PS_PP040L.U_OrdGbn AS OrdGbn,"
'            Query01 = Query01 & " PS_PP040L.U_PP030HNo AS PP030HNo,"
'            Query01 = Query01 & " PS_PP040L.U_PP030MNo AS PP030MNo, "
'            Query01 = Query01 & " PS_PP040L.U_OrdMgNum AS OrdMgNum "
'            Query01 = Query01 & " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry "
'            Query01 = Query01 & " WHERE "
'            Query01 = Query01 & " PS_PP040H.Canceled = 'N' AND PS_PP040L.DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
'            RecordSet01.DoQuery Query01
'            For i = 0 To RecordSet01.RecordCount - 1
'                If RecordSet01.Fields("OrdGbn").VALUE = "107" Then '//엔트베어링이면서
'                    '//실적포인트가 체크된 공정인경우..
'                    If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1) = "Y" Then
'                        '//실적 '//문서의타입필요
'                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP080L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
'                            MDC_Com.MDC_GF_Message "생산실적 등록된 문서입니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
'                ElseIf RecordSet01.Fields("OrdGbn").VALUE = "104" Then '//멀티이면
'                    '//마지막공정에 항상 실적포인트가 되어있음
'                    If MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1) = "Y" Then
'                        '//실적 '//문서의타입필요
'                        If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").VALUE & "' AND PS_PP080L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").VALUE & "'", 0, 1)) > 0 Then
'                            MDC_Com.MDC_GF_Message "생산실적 등록된 문서입니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
'                End If
'
'                If RecordSet01.Fields("OrdGbn").VALUE = "104" Then '//멀티
'                    '//다음공정이 존재하면
'                    NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & RecordSet01.Fields("OrdMgNum").VALUE & "'", 0, 1)
'                    If NextCpInfo <> "" Then
'                        If MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" & NextCpInfo & "'", 0, 1) > 0 Then
'                            '//다음공정으로 생성된 실적조회 만약 존재한다면 취소불가능
'                            MDC_Com.MDC_GF_Message "후행공정이 입력된 문서입니다. 적용할수 없습니다.", "W"
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    Else
'                        '//다음공정이 존재하지 않으면 마지막 공정임, 마지막공정일때는 실적등록여부로 적용여부 판정
'                    End If
'                ElseIf RecordSet01.Fields("OrdGbn").VALUE = "107" Then '//엔드베어링
'                    '//삭제된 행에 대한 검사..
'                    OrdMgNum = RecordSet01.Fields("OrdMgNum").VALUE '// DocEntry + '-' + LineId
'                    CurrentCpInfo = OrdMgNum
'
'                    PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_02 '" & OrdMgNum & "'")
'                    If PrevCpInfo = "" Then
'                        '//해당공정이 첫공정이면 입력되어도 상관없다.
'                    Else
'                        PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & PrevCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040H.Canceled = 'N'")
'                        PrevMATRIXCpQty = 0
'                        For j = 1 To oMat01.VisualRowCount - 1
'                            If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = PrevCpInfo) Then
'                                PrevMATRIXCpQty = PrevMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
'                            End If
'                        Next
'                        CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).VALUE & "' AND PS_PP040H.Canceled = 'N'")
'                        CurrentMATRIXCpQty = 0
'                        For j = 1 To oMat01.VisualRowCount - 1
'                            If (oMat01.Columns("OrdMgNum").Cells(j).Specific.VALUE = CurrentCpInfo) Then
'                                CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.VALUE)
'                            End If
'                        Next
'                        If ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
'                            Sbo_Application.SetStatusBarMessage "취소문서의 선행공정의 생산수량이 취소문서의 생산수량을 미달합니다.", bmt_Short, True
'                            PS_PP043_Validate = False
'                            GoTo PS_PP043_Validate_Exit
'                        End If
'                    End If
''                   If oForm01.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
''                        NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & OrdMgNum & "'")
''                        If NextCpInfo = "" Then
''                            '//해당공정이 마지막공정이면 삭제되어도 상관없다.
''                        Else
''                            NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP040H.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
''                            NextMATRIXCpQty = 0
''                            For j = 1 To oMat01.VisualRowCount - 1
''                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = NextCpInfo) Then
''                                    NextMATRIXCpQty = NextMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
''                                End If
''                            Next
''                            CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP040L.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP040H.Canceled = 'N'")
''                            CurrentMATRIXCpQty = 0
''                            For j = 1 To oMat01.VisualRowCount - 1 '//현재공정은 삭제되었으므로.. 매트릭스에 존재하지 않는다.
''                                If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = CurrentCpInfo) Then
''                                    CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
''                                End If
''                            Next
''                            If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
''                                Sbo_Application.SetStatusBarMessage "취소문서의 후행공정의 생산수량이 취소문서의 생산수량을 초과합니다.", bmt_Short, True
''                                PS_PP043_Validate = False
''                                GoTo PS_PP043_Validate_Exit
''                            End If
''                        End If
''                    End If
'                End If
'                RecordSet01.MoveNext
'            Next
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입이 PSMT인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입이 외주인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입이 실적인경우
        ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입이 조정인경우
        End If
    End If
    
    Set RecordSet01 = Nothing
    Exit Function
PS_PP043_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_PP043_Validate_Error:
    PS_PP043_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_PP043_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP043_OrderInfoLoad()
On Error GoTo PS_PP043_OrderInfoLoad_Error
    Call oForm01.Freeze(True)
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    If oForm01.Items("OrdType").Specific.Selected.VALUE = "10" Or oForm01.Items("OrdType").Specific.Selected.VALUE = "50" Then '//작업타입 일반,조정
        If oForm01.Items("SCpCode").Specific.VALUE = "" Then '//공정코드
            Call MDC_Com.MDC_GF_Message("작업지시 공정을 입력하지 않습니다.", "W")
            GoTo PS_PP043_OrderInfoLoad_Exit
        Else
            '//작업구분선택
            oForm01.Items("CpCode").Specific.VALUE = oForm01.Items("SCpCode").Specific.VALUE
            oForm01.Items("CpName").Specific.VALUE = oForm01.Items("SCpName").Specific.VALUE
            
            If oForm01.Items("SOrdGbn").Specific.Selected.VALUE = "선택" Then '//값이 선택되어 있지 않다면 기본으로
                Call oForm01.Items("OrdGbn").Specific.Select(MDC_PS_Common.GetValue("SELECT U_ItmBsort FROM [@PS_PP001L] WHERE U_CpCode = '" & oForm01.Items("SCpCode").Specific.VALUE & "'", 0, 1), psk_ByValue)
            Else '//값이 선택되었다면 선택된 값으로..
                Call oForm01.Items("OrdGbn").Specific.Select(oForm01.Items("SOrdGbn").Specific.Selected.VALUE, psk_ByValue)
            End If
            '//사업장
            If oForm01.Items("SBPLId").Specific.Selected.VALUE = "선택" Then
                Call oForm01.Items("BPLId").Specific.Select(oForm01.Items("SBPLId").Specific.Selected.VALUE, psk_ByValue)
            Else
                Call oForm01.Items("BPLId").Specific.Select(oForm01.Items("SBPLId").Specific.Selected.VALUE, psk_ByValue)
            End If
            
            '//공정코드에 따라 매트릭스 변경
            If MDC_PS_Common.GetValue("SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = '" & oForm01.Items("CpCode").Specific.VALUE & "'", 0, 1) = "CP101" Then '//엔드베어링
                oMat01.Columns("BQty").Visible = False
                oMat01.Columns("PSum").Visible = True
                oMat01.Columns("PWeight").Visible = True
                oMat01.Columns("YWeight").Visible = True
                oMat01.Columns("NWeight").Visible = True
            ElseIf MDC_PS_Common.GetValue("SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = '" & oForm01.Items("CpCode").Specific.VALUE & "'", 0, 1) = "CP501" Then '//멀티
                oMat01.Columns("BQty").Visible = True
                oMat01.Columns("PSum").Visible = False
                oMat01.Columns("PWeight").Visible = False
                oMat01.Columns("YWeight").Visible = False
                oMat01.Columns("NWeight").Visible = False
            End If
            
            '//매트릭스삭제
            oMat01.Clear
            oMat01.FlushToDataSource
            oMat01.LoadFromDataSource
            Call PS_PP043_AddMatrixRow01(0, True)
            oMat02.Clear
            oMat02.FlushToDataSource
            oMat02.LoadFromDataSource
            Call PS_PP043_AddMatrixRow02(0, True)
            oMat03.Clear
            oMat03.FlushToDataSource
            oMat03.LoadFromDataSource
            oForm01.Update
        End If
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "20" Then '//작업타입 PSMT
        Call MDC_Com.MDC_GF_Message("PSMT지원은 입력할수 없습니다.", "W")
        GoTo PS_PP043_OrderInfoLoad_Exit
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "30" Then '//작업타입 외주
        Call MDC_Com.MDC_GF_Message("외주은 입력할수 없습니다.", "W")
        GoTo PS_PP043_OrderInfoLoad_Exit
    ElseIf oForm01.Items("OrdType").Specific.Selected.VALUE = "40" Then '//작업타입 실적
        Call MDC_Com.MDC_GF_Message("실적은 입력할수 없습니다.", "W")
        GoTo PS_PP043_OrderInfoLoad_Exit
    End If
    
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
        Set ProgressBar01 = Nothing
    End If
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP043_OrderInfoLoad_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
        Set ProgressBar01 = Nothing
    End If
    Exit Sub
PS_PP043_OrderInfoLoad_Error:
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
        Set ProgressBar01 = Nothing
    End If
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP043_OrderInfoLoad_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Function PS_PP043_FindValidateDocument(ByVal ObjectType As String) As Boolean
On Error GoTo PS_PP043_FindValidateDocument_Error
    PS_PP043_FindValidateDocument = True
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim Query02 As String
    Dim RecordSet02 As SAPbobsCOM.Recordset
    
    Dim i As Long
    Dim DocEntry As String
    DocEntry = oForm01.Items("DocEntry").Specific.VALUE '//원본문서
    
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Query01 = " SELECT DocEntry"
    Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry = "
    Query01 = Query01 & DocEntry
    If (oDocType01 = "작업일보등록(작지)") Then
        Query01 = Query01 & " AND U_DocType = '10'"
    ElseIf (oDocType01 = "작업일보등록(공정)") Then
        Query01 = Query01 & " AND U_DocType = '20'"
    End If
    RecordSet01.DoQuery Query01
    If (RecordSet01.RecordCount = 0) Then
        If (oDocType01 = "작업일보등록(작지)") Then
            Sbo_Application.SetStatusBarMessage "작업일보등록(공정)문서 이거나 존재하지 않는 문서입니다.", bmt_Short, True
        ElseIf (oDocType01 = "작업일보등록(공정)") Then
            Sbo_Application.SetStatusBarMessage "작업일보등록(작지)문서 이거나 존재하지 않는 문서입니다.", bmt_Short, True
        End If
        PS_PP043_FindValidateDocument = False
        Set RecordSet01 = Nothing
        Set RecordSet02 = Nothing
        Exit Function
    End If
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Exit Function
PS_PP043_FindValidateDocument_Error:
    Sbo_Application.SetStatusBarMessage Err.Number & Err.Description, bmt_Short, True
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_PP043_FindValidateDocument = False
End Function

Public Function PS_PP043_DirectionValidateDocument(ByVal DocEntry As String, ByVal DocEntryNext As String, ByVal Direction As String, ByVal ObjectType As String) As Boolean
On Error GoTo PS_PP043_DirectionValidateDocument_Error
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Dim Query02 As String
    Dim RecordSet02 As SAPbobsCOM.Recordset
    
    Dim i As Long
    Dim MaxDocEntry As String
    Dim MinDocEntry As String
    Dim DoNext As Boolean
    Dim IsFirst As Boolean '//시작유무
    DoNext = True
    IsFirst = True
    
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Do While (DoNext = True)
        If (IsFirst <> True) Then
            If (DocEntry = DocEntryNext) Then '//문서전체를 경유하고도 유효값을 찾지못했다면
                Sbo_Application.SetStatusBarMessage "유효한문서가 존재하지 않습니다", bmt_Short, True
                PS_PP043_DirectionValidateDocument = False
                Set RecordSet01 = Nothing
                Set RecordSet02 = Nothing
                Exit Function
            End If
        End If
        If (Direction = "Next") Then
            Query01 = " SELECT TOP 1 DocEntry"
            Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry > "
            Query01 = Query01 & DocEntryNext
            If (oDocType01 = "작업일보등록(작지)") Then
                Query01 = Query01 & " AND U_DocType = '10'"
            ElseIf (oDocType01 = "작업일보등록(공정)") Then
                Query01 = Query01 & " AND U_DocType = '20'"
            End If
            Query01 = Query01 & " ORDER BY DocEntry ASC"
        ElseIf (Direction = "Prev") Then
            Query01 = " SELECT TOP 1 DocEntry"
            Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry < "
            Query01 = Query01 & DocEntryNext
            If (oDocType01 = "작업일보등록(작지)") Then
                Query01 = Query01 & " AND U_DocType = '10'"
            ElseIf (oDocType01 = "작업일보등록(공정)") Then
                Query01 = Query01 & " AND U_DocType = '20'"
            End If
            Query01 = Query01 & " ORDER BY DocEntry DESC"
        End If
        RecordSet01.DoQuery Query01
        If (RecordSet01.Fields(0).VALUE = 0) Then '//해당문서가 마지막문서라면
            If (Direction = "Next") Then
                Query02 = " SELECT TOP 1 DocEntry FROM [" & ObjectType & "]"
                If (oDocType01 = "작업일보등록(작지)") Then
                    Query02 = Query02 & " WHERE U_DocType = '10'"
                ElseIf (oDocType01 = "작업일보등록(공정)") Then
                    Query02 = Query02 & " WHERE U_DocType = '20'"
                End If
                Query02 = Query02 & " ORDER BY DocEntry ASC"
            ElseIf (Direction = "Prev") Then
                Query02 = " SELECT TOP 1 DocEntry FROM [" & ObjectType & "]"
                If (oDocType01 = "작업일보등록(작지)") Then
                    Query02 = Query02 & " WHERE U_DocType = '10'"
                ElseIf (oDocType01 = "작업일보등록(공정)") Then
                    Query02 = Query02 & " WHERE U_DocType = '20'"
                End If
                Query02 = Query02 & " ORDER BY DocEntry DESC"
            End If
            RecordSet02.DoQuery Query02
            If (RecordSet02.RecordCount = 0) Then '//문서가 아예 존재하지 않는다면
                Sbo_Application.SetStatusBarMessage "유효한문서가 존재하지 않습니다", bmt_Short, True
                Set RecordSet01 = Nothing
                Set RecordSet02 = Nothing
                PS_PP043_DirectionValidateDocument = False
                Exit Function
            Else
                If (Direction = "Next") Then
                    DocEntryNext = Val(RecordSet02.Fields(0).VALUE) - 1
                    Query01 = " SELECT TOP 1 DocEntry"
                    Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry > "
                    Query01 = Query01 & DocEntryNext
                    If (oDocType01 = "작업일보등록(작지)") Then
                        Query01 = Query01 & " AND U_DocType = '10'"
                    ElseIf (oDocType01 = "작업일보등록(공정)") Then
                        Query01 = Query01 & " AND U_DocType = '20'"
                    End If
                    Query01 = Query01 & " ORDER BY DocEntry ASC"
                    RecordSet01.DoQuery Query01
                ElseIf (Direction = "Prev") Then
                    DocEntryNext = Val(RecordSet02.Fields(0).VALUE) + 1
                    Query01 = " SELECT TOP 1 DocNum"
                    Query01 = Query01 & " FROM [" & ObjectType & "] Where DocEntry < "
                    Query01 = Query01 & DocEntryNext
                    If (oDocType01 = "작업일보등록(작지)") Then
                        Query01 = Query01 & " AND U_DocType = '10'"
                    ElseIf (oDocType01 = "작업일보등록(공정)") Then
                        Query01 = Query01 & " AND U_DocType = '20'"
                    End If
                    Query01 = Query01 & " ORDER BY DocEntry DESC"
                    RecordSet01.DoQuery Query01
                End If
            End If
        End If
        If (oDocType01 = "작업일보등록(작지)") Then
            DoNext = False
            If (Direction = "Next") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) - 1
            ElseIf (Direction = "Prev") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) + 1
            End If
        ElseIf (oDocType01 = "작업일보등록(공정)") Then
            DoNext = False
            If (Direction = "Next") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) - 1
            ElseIf (Direction = "Prev") Then
                DocEntryNext = Val(RecordSet01.Fields(0).VALUE) + 1
            End If
        End If
        IsFirst = False
    Loop
    If (DocEntry = DocEntryNext) Then '//다음문서가 유효하다면 그냥 넘어가고
        Call PS_PP043_FormItemEnabled '//UDO방식
    Else '//다음문서가 유효하지 않다면
        oForm01.Freeze True
        oForm01.Mode = fm_FIND_MODE
        Call PS_PP043_FormItemEnabled '//UDO방식
        If oForm01.Items("DocEntry").Enabled = True Then '//문서번호 필드가 입력이 가능하다면
            If (Direction = "Next") Then
                oForm01.Items("DocEntry").Specific.VALUE = Val(DocEntryNext + 1)
            ElseIf (Direction = "Prev") Then
                oForm01.Items("DocEntry").Specific.VALUE = Val(DocEntryNext - 1)
            End If
            oForm01.Items("1").Click ct_Regular
        End If
        oForm01.Freeze False
        PS_PP043_DirectionValidateDocument = False
        Set RecordSet01 = Nothing
        Set RecordSet02 = Nothing
        Exit Function
    End If
    PS_PP043_DirectionValidateDocument = True
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    Exit Function
PS_PP043_DirectionValidateDocument_Error:
    Sbo_Application.SetStatusBarMessage Err.Number & Err.Description, bmt_Short, True
    PS_PP043_DirectionValidateDocument = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
End Function

Private Function Insert_oInventoryGenExit(ChkType As Integer) As Boolean
    On Error GoTo error_Message
'===================================================================================
'   출고 DI
'   멀티,엔드베어링의 경우 첫공정이면서 처음 작업일보 등록시 투입자재를 출고 시킨다.
'===================================================================================
    Dim DI_oInventoryGenExit As SAPbobsCOM.Documents    '//재고입고 문서 객체
    Dim RetVal&
    Dim VIDocNum$, VODocNum$
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim i&, K&, oRow&
    Dim ErrNum          As Integer
    Dim Cnt&
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    For oRow = 0 To oDS_PS_PP043L.Size - 1
        If Trim(oDS_PS_PP043L.GetValue("U_OrdGbn", oRow)) = "104" Or Trim(oDS_PS_PP043L.GetValue("U_OrdGbn", oRow)) = "107" Then '멀티 & 엔드베어링일경우
            '107010002(END BEARING #44),107010004(END BEARING #2) 일경우에는 원자재가 스크랩을 이용이므로, 불출될 원자재가 없음.
            If Trim(oDS_PS_PP043L.GetValue("U_ItemCode", oRow)) <> "107010002" And Trim(oDS_PS_PP043L.GetValue("U_ItemCode", oRow)) <> "107010004" Then

                If Trim(oDS_PS_PP043L.GetValue("U_Sequence", oRow)) = "1" Then '첫공정일 경우
                    
                    sQry = "select b.docentry from [@PS_PP040L] a inner join [@PS_PP040H] b on a.docentry=b.docentry "
                    sQry = sQry & "where a.U_OrdGbn in ('104','107') and b.canceled <> 'Y' "
                    sQry = sQry & "and a.U_PP030HNo = '" & Trim(oDS_PS_PP043L.GetValue("U_PP030HNo", oRow)) & "' "
                    sQry = sQry & "and a.U_Sequence = '" & Trim(oDS_PS_PP043L.GetValue("U_Sequence", oRow)) & "' "
                    oRecordSet.DoQuery sQry
                    
                    If oRecordSet.RecordCount < 1 Then '처음 작업일보 등록시
                        Cnt = Cnt + 1
                    End If
                End If
            End If
        End If
    Next oRow
    If Cnt < 1 Then
        Insert_oInventoryGenExit = True
        Exit Function '출고 DI를 태울 조건에 해당하는 Data가 없을시.
    End If
        
    '/////////////////////////////////////////////////////////////////////////////////
    '출고
    '/////////////////////////////////////////////////////////////////////////////////
    Sbo_Company.StartTransaction
    Set DI_oInventoryGenExit = Nothing
    Set DI_oInventoryGenExit = Sbo_Company.GetBusinessObject(oInventoryGenExit)   '// 문서타입(입고)
    
    i = 1
    With DI_oInventoryGenExit
        
        '/ Header
        .DocDate = Format(Trim(oForm01.Items("DocDate").Specific.VALUE), "0000-00-00")
        .TaxDate = Format(Trim(oForm01.Items("DocDate").Specific.VALUE), "0000-00-00")
        .Comments = "작업일보등록(" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & ") 출고"
    
        '/ Line
        For oRow = 0 To oDS_PS_PP043L.Size - 1
            If Trim(oDS_PS_PP043L.GetValue("U_OrdGbn", oRow)) = "104" Or Trim(oDS_PS_PP043L.GetValue("U_OrdGbn", oRow)) = "107" Then '멀티 & 엔드베어링일경우
                If Trim(oDS_PS_PP043L.GetValue("U_Sequence", oRow)) = "1" Then '첫공정일 경우
                    
                    sQry = "select b.docentry from [@PS_PP040L] a inner join [@PS_PP040H] b on a.docentry=b.docentry "
                    sQry = sQry & "where a.U_OrdGbn in ('104','107') and b.canceled <> 'Y' "
                    sQry = sQry & "and a.U_PP030HNo = '" & Trim(oDS_PS_PP043L.GetValue("U_PP030HNo", oRow)) & "' "
                    sQry = sQry & "and a.U_Sequence = '" & Trim(oDS_PS_PP043L.GetValue("U_Sequence", oRow)) & "' "
                    oRecordSet.DoQuery sQry
                    
                    If oRecordSet.RecordCount < 1 Then '처음 작업일보 등록시
                        If .Lines.Count < i Then
                           .Lines.Add
                           .Lines.BatchNumbers.Add
                        End If
                        
                        sQry = "select U_ItemCode,U_ItemName,U_BatchNum,U_Weight from [@PS_PP030L] where docentry = '" & Trim(oDS_PS_PP043L.GetValue("U_PP030HNo", oRow)) & "'"
                        oRecordSet.DoQuery sQry
                        
                        .Lines.SetCurrentLine i - 1
                        .Lines.ItemCode = Trim(oRecordSet.Fields("U_ItemCode").VALUE)
                        .Lines.ItemDescription = Trim(oRecordSet.Fields("U_ItemName").VALUE)
                        .Lines.BatchNumbers.BatchNumber = Trim(oRecordSet.Fields("U_BatchNum").VALUE)
                        .Lines.BatchNumbers.Quantity = Trim(oRecordSet.Fields("U_Weight").VALUE)
                        .Lines.Quantity = Val(oRecordSet.Fields("U_Weight").VALUE)
                        
                        sQry = "select TOP 1 WhsCode from [OIBT] where Quantity <> 0 "
                        sQry = sQry & "and ItemCode = '" & Trim(oRecordSet.Fields("U_ItemCode").VALUE) & "' "
                        sQry = sQry & "and BatchNum = '" & Trim(oRecordSet.Fields("U_BatchNum").VALUE) & "'"
                        oRecordSet.DoQuery sQry
                        .Lines.WarehouseCode = Trim(oRecordSet.Fields("WhsCode").VALUE)
                       
        '                .Lines.PriceAfterVAT = Val(oDS_PS_PP077H.GetValue("U_Price", oRow))
        '                .Lines.UserFields("U_Number").Value = Val(oDS_PS_PP077H.GetValue("U_Number", oRow))
        '                .Lines.AccountCode = Trim(ZPSH_SetFunc.Get_ReData("StkInTnAct", "WhsCode", "OWHS", "'" & Trim(oDS_PS_PP077H.GetValue("U_AWhsCode", oRow)) & "'"))
        '                .Lines.Currency = "MYR"
                    
                        i = i + 1
                    End If
                End If
            End If
        Next oRow
        
    End With

    '/-------------------------------------------------------------------------/
    '/ 출고전송
'    If Trim(oForm01.Items("IDocNum").Specific.Value) = "" Then
        RetVal = DI_oInventoryGenExit.Add
        If (0 <> RetVal) Then
             Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
             GoTo error_Message  '//저장시 에러 발생
        End If
        If ChkType <> 2 Then
            Sbo_Company.EndTransaction wf_RollBack
        Else
            Sbo_Company.GetNewObjectCode VIDocNum
'            oForm01.Items("IDocNum").Specific.Value = VIDocNum
            
            i = 1
            For oRow = 0 To oDS_PS_PP043L.Size - 1
                If Trim(oDS_PS_PP043L.GetValue("U_OrdGbn", oRow)) = "104" Or Trim(oDS_PS_PP043L.GetValue("U_OrdGbn", oRow)) = "107" Then '멀티 & 엔드베어링일경우
                    If Trim(oDS_PS_PP043L.GetValue("U_Sequence", oRow)) = "1" Then '첫공정일 경우
                       oDS_PS_PP043L.setValue "U_OutDoc", oRow, VIDocNum
                       oDS_PS_PP043L.setValue "U_OutLin", oRow, i
                       
                       i = i + 1
                    End If
                End If
            Next oRow
            oMat01.LoadFromDataSource
            
            Sbo_Company.EndTransaction wf_Commit
        End If
'     End If
     
    Insert_oInventoryGenExit = True
    Set DI_oInventoryGenExit = Nothing
    Set oRecordSet = Nothing
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
error_Message:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Insert_oInventoryGenExit = False
    
    If ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "Save_Data_Error:" & Err.Number & " - " & Err.Description, "E"
    Else
'        MDC_Com.MDC_GF_Message "Insert_oInventoryGenExit_Error:" & Err.Number & " - " & Err.Description, "E"
        MDC_Com.MDC_GF_Message "Insert_oInventoryGenExit_Error:" & ErrCode & " - " & ErrMsg, "E"
    End If
        
    Set DI_oInventoryGenExit = Nothing
    Set oRecordSet = Nothing
End Function
Private Function Insert_oInventoryGenEntry(ChkType As Integer) As Boolean
On Error GoTo error_Message
'============================================================================
'   입고 DI
'   출고 취소시 입고
'============================================================================
    Dim DI_oInventoryGenEntry As SAPbobsCOM.Documents    '//재고입고 문서 객체
    Dim RetVal&
    Dim VIDocNum$, VODocNum$
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim i&, K&, oRow&
    Dim ErrNum          As Integer
    Dim Cnt&
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    For oRow = 0 To oDS_PS_PP043L.Size - 1
        If Trim(oDS_PS_PP043L.GetValue("U_OutDoc", oRow)) <> "" Then '출고되어진 문서가 있는경우
            Cnt = Cnt + 1
        End If
    Next oRow
    If Cnt < 1 Then
        Insert_oInventoryGenEntry = True
        Exit Function '입고 DI를 태울 조건에 해당하는 Data가 없을시.
    End If
   
    '/////////////////////////////////////////////////////////////////////////////////
    '입고
    '/////////////////////////////////////////////////////////////////////////////////
    Sbo_Company.StartTransaction
    Set DI_oInventoryGenEntry = Nothing
    Set DI_oInventoryGenEntry = Sbo_Company.GetBusinessObject(oInventoryGenEntry)   '// 문서타입(입고)

    i = 1
    With DI_oInventoryGenEntry
        
        '/ Header
        .DocDate = Format(Trim(oForm01.Items("DocDate").Specific.VALUE), "0000-00-00")
        .TaxDate = Format(Trim(oForm01.Items("DocDate").Specific.VALUE), "0000-00-00")
        .Comments = "작업일보등록(" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & ") 출고 취소"
    
        '/ Line
        For oRow = 0 To oDS_PS_PP043L.Size - 1
            If Trim(oDS_PS_PP043L.GetValue("U_OutDoc", oRow)) <> "" Then '출고되어진 문서가 있는경우
           
                If .Lines.Count < i Then
                   .Lines.Add
                   .Lines.BatchNumbers.Add
                End If
                
                sQry = "select U_ItemCode,U_ItemName,U_BatchNum,U_Weight from [@PS_PP030L] where docentry = '" & Trim(oDS_PS_PP043L.GetValue("U_PP030HNo", oRow)) & "'"
                oRecordSet.DoQuery sQry
                
                .Lines.SetCurrentLine i - 1
                .Lines.ItemCode = Trim(oRecordSet.Fields("U_ItemCode").VALUE)
                .Lines.ItemDescription = Trim(oRecordSet.Fields("U_ItemName").VALUE)
                .Lines.BatchNumbers.BatchNumber = Trim(oRecordSet.Fields("U_BatchNum").VALUE)
                .Lines.BatchNumbers.Quantity = Trim(oRecordSet.Fields("U_Weight").VALUE)
                .Lines.Quantity = Val(oRecordSet.Fields("U_Weight").VALUE)
                
                '출고된 창고 select
                sQry = "select WhsCode from [IGE1]"
                sQry = sQry & "where docentry = '" & Trim(oDS_PS_PP043L.GetValue("U_OutDoc", oRow)) & "' "
                sQry = sQry & "and linenum = '" & Trim(oDS_PS_PP043L.GetValue("U_OutLin", oRow)) - 1 & "'"
                oRecordSet.DoQuery sQry
                .Lines.WarehouseCode = Trim(oRecordSet.Fields("WhsCode").VALUE)
                
'                .Lines.PriceAfterVAT = Val(oDS_PS_PP077H.GetValue("U_Price", oRow))
'                .Lines.UserFields("U_Number").Value = Val(oDS_PS_PP077H.GetValue("U_Number", oRow))
'                .Lines.AccountCode = Trim(ZPSH_SetFunc.Get_ReData("StkInTnAct", "WhsCode", "OWHS", "'" & Trim(oDS_PS_PP077H.GetValue("U_AWhsCode", oRow)) & "'"))
'                .Lines.Currency = "MYR"
                
                i = i + 1
            End If
        Next oRow
    End With

    '/-------------------------------------------------------------------------/
    '/ 입고전송
'    If Trim(oForm01.Items("IDocNum").Specific.Value) = "" Then
        RetVal = DI_oInventoryGenEntry.Add
        If (0 <> RetVal) Then
             Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
             GoTo error_Message  '//저장시 에러 발생
        End If
        If ChkType <> 2 Then
            Sbo_Company.EndTransaction wf_RollBack
        Else
            Sbo_Company.GetNewObjectCode VIDocNum
'            oForm01.Items("IDocNum").Specific.Value = VIDocNum
            
            i = 1
            For oRow = 0 To oDS_PS_PP043L.Size - 1
                If Trim(oDS_PS_PP043L.GetValue("U_OutDoc", oRow)) <> "" Then '출고되어진 문서가 있는경우
'                       oDS_PS_PP043L.setValue "U_OutDocC", oRow, VIDocNum
'                       oDS_PS_PP043L.setValue "U_OutLinC", oRow, oRow + 1
                       
                       '//update
                       sQry = "Update [@PS_PP040L] set U_OutDocC = '" & VIDocNum & "'"
                       sQry = sQry & ", U_OutLinC = '" & i & "' "
                       sQry = sQry & "where docentry = '" & Trim(oForm01.Items("DocEntry").Specific.VALUE) & "' "
                       sQry = sQry & "and visorder = '" & oRow & "'"
                       oRecordSet.DoQuery sQry
                       
                       i = i + 1
                End If
            Next oRow
            oMat01.LoadFromDataSource

             Sbo_Company.EndTransaction wf_Commit
        End If
'     End If
     
    Insert_oInventoryGenEntry = True
    Set DI_oInventoryGenEntry = Nothing
    Set oRecordSet = Nothing
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
error_Message:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Insert_oInventoryGenEntry = False
    
    If ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "Add_PS_PP040_Error:" & Err.Number & " - " & Err.Description, "E"
    ElseIf ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Save_Data_Error:" & Err.Number & " - " & Err.Description, "E"
    Else
        MDC_Com.MDC_GF_Message "Insert_oInventoryGenEntry_Error:" & ErrCode & " - " & ErrMsg, "E"
    End If
        
    Set DI_oInventoryGenEntry = Nothing
    Set oRecordSet = Nothing
End Function



