VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM090"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM090.cls
'//  Module         : 자재기타출고등록
'//  Desc           : 재고관리하지 않는 출고분에대한 출력문서 등록후 출력
'//  FormType       : PS_MM090
'//  Create Date    : 2010.11.01
'//  Editor         : 박성수
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_MM090H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_MM090L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    Dscription As String
    LotNo As String
    Quantity As Long
    OPORNo As Long
    POR1No As Long
    Check As Boolean
    OPDNNo As Long
    PDN1No As Long
    LineTotal As Long
    
End Type

Private oOutMan As String

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM090.srf")
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
    
    oFormUniqueID01 = "PS_MM090_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_MM090_CreateItems
    Call PS_MM090_ComboBox_Setting
    Call PS_MM090_CF_ChooseFromList
    Call PS_MM090_EnableMenus
    Call PS_MM090_SetDocument(oFromDocEntry01)
    Call PS_MM090_FormResize
    
    oForm01.Items("empty").Visible = False
    
    oForm01.Items("CardCode").Click
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
                    If (PS_MM090_Validate("취소") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
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
            Case "1288", "1289", "1290", "1291": '레코드이동버
                    
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
        
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_MM090_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_MM090_FormItemEnabled '//UDO방식
                Call PS_MM090_AddMatrixRow(0, True) '//UDO방식 추가하고 나면 메트릭스에 한행추가한다.
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_MM090_FormItemEnabled
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
        If pval.ItemUID = "PS_MM090" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_MM090_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oOutMan = Trim(oForm01.Items("OutMan").Specific.VALUE)
                

                '//해야할일 작업 (출고증문서번호를 순번해서 이문서에 업데이트해준다.)
'                Dim sPackDate, sIndex, sPackNo As String
'                Dim i As Integer
'
'                sPackDate = Trim(oForm01.Items("InDate").Specific.VALUE)
'                sIndex = MDC_PS_Common.GetValue("EXEC PS_MM090_01 '" & oForm01.Items("InDate").Specific.VALUE & "'")
'                sPackNo = sPackDate & sIndex
'
'                'Call oDS_PS_MM090H.setValue("U_PackNo", 0, sPackNo & MDC_PS_Common.GetValue("EXEC PS_MM090_01 '" & oForm01.Items("InDate").Specific.Value & "'"))
'                oDS_PS_MM090H.setValue "U_OutNum", 0, sPackNo
                               
                
                'UpdateToPP090 (sPackNo)
'                For i = 0 To (oMat01.VisualRowCount - 1)
'                    Call oDS_PS_MM090L.setValue("U_OutNum", i, sPackNo)
'                    oMat01.Columns("PackNo").Cells(i + 1).Specific.Value = sPackNo
'                Next
    
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_MM090_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
                oOutMan = Trim(oForm01.Items("OutMan").Specific.VALUE)
            ElseIf oForm01.Mode = fm_OK_MODE Then
                
            ElseIf oForm01.Mode = fm_FIND_MODE Then
                 
            End If
        ElseIf (pval.ItemUID = "Btn1") Then '출고증출력시 출력으로 액션
        
            Call PS_SMM090_Print_Report01
             
        ElseIf (pval.ItemUID = "Btn02") Then '거래명세표출력시 출력으로 액션
        
             Call PS_SMM090_Print_Report02
        
        ElseIf pval.ItemUID = "Opt01" Then
        
            Call oForm01.Freeze(True)
            oMat01.Columns("DocLin").Visible = False
            Call oMat01.AutoResizeColumns
            Call oForm01.Freeze(False)
            
        ElseIf pval.ItemUID = "Opt02" Then
        
            Call oForm01.Freeze(True)
            oMat01.Columns("DocLin").Visible = True
            Call oMat01.AutoResizeColumns
            Call oForm01.Freeze(False)
             
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_MM090" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_MM090_FormItemEnabled
                    Call PS_MM090_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_MM090_FormItemEnabled
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
    
'        If pval.ItemUID = "Mat01" Then
'            If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "작지번호를 먼저 선택하세요", bmt_Short, True
'                BubbleEvent = False
'                Exit Sub
'            End If
'        End If
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "Mat01" Then
                If pval.ColUID = "ItemName" Then
                    If oMat01.Columns("ItemName").Cells(pval.Row).Specific.VALUE = "" Then
                        '포장이고 bom을 사용할때(현재는 포장만 bom사용)
                        If (oForm01.Items("BPLId").Specific.Selected.VALUE = "3" Or oForm01.Items("BPLId").Specific.Selected.VALUE = "5") And oForm01.Items("BOM_CHECK").Specific.Checked = True Then
                            '//창원 휘팅일때
                            Dim ChildForm01 As Variant
                            Set ChildForm01 = New PS_SM030 '포장제품BOM
                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, "")
                            Set ChildForm01 = Nothing
                        End If
                    End If
                ElseIf pval.ColUID = "DocLin" Then
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "DocLin")
                End If
                    
    
            End If
        End If
        '헤더에 질의연결이벤트
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "")        '//거래처코드
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OutMan", "")        '//반출자조회
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "TranCard", "")        '//차량조회
        'Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "")        '//품목코드
        
        '메트릭스에 질의연결이벤트
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "LotNo")     '//Lot번호
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "CpCode")     '//공정코드
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "WorkMan")    '//작업자정보
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode")    '//품목코드조회
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    If (pval.ItemUID = "OutCls") Then
        If oForm01.Items("OutCls").Specific.VALUE = "09" Then
            Sbo_Application.SetStatusBarMessage "반출구분이 'MG스크랩'일때는 Line 출고구분행을 'MG스크랩출고'로 등록해주세요."
        End If
    End If
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemChanged = True Then
            Call oForm01.Freeze(True)
            If (pval.ItemUID = "Mat01") Then
                Call oMat01.FlushToDataSource
'                Call oMat01.LoadFromDataSource
'                Call oMat01.AutoResizeColumns
            Else
                If (pval.ItemUID = "Tonnage") Then
                    '창원사업장은 톤수에 따라 운송비를 입력(MG스크랩 운송비때문에) 2016/2/28 N.G.Y 추가
                    If oForm01.Items("BPLId").Specific.VALUE = "1" Then
                        sQry = " Select U_RelCd From [@PS_SY001L] WHERE Code = 'M009' and U_Minor = '" & oForm01.Items("Tonnage").Specific.Selected.VALUE & "'"
                        Call oRecordSet01.DoQuery(sQry)
                        Call oDS_PS_MM090H.setValue("U_TranCost", 0, Trim(oRecordSet01.Fields(0).VALUE))
                        
                        Call oDS_PS_MM090H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    End If
                Else
                    Call oDS_PS_MM090H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                End If
            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
            oForm01.Update
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
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                Call oMat01.SelectRow(pval.Row, True, False)    '메트릭스 한줄선택시 반전시켜주는 구문
            End If
        
        ElseIf pval.ItemUID = "1" Then
            
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "1" Then
            Call oForm01.EnableMenu("1281", True) '찾기하고 다시 찾기아이콘활성화처리
        End If
    
'        BubbleEvent = False
'         Exit Sub
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Call oForm01.Freeze(False)
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
    Dim sReHour As Long
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Mat01") Then
                If (pval.ColUID = "ItemName") Then
                    '//기타작업
                    Call oDS_PS_MM090L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) '//자기자신먼저 Flush처리
                    
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM090L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_MM090_AddMatrixRow (pval.Row)
                    End If

                    'Call oDS_PS_MM090L.setValue("U_ItemCode", pval.Row - 1, MDC_PS_Common.GetValue("SELECT Distinct ItemCode FROM OIBT WHERE BatchNum='" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value & "'"))
                    'Call oDS_PS_MM090L.setValue("U_ItemName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT Distinct ItemName FROM OIBT WHERE BatchNum='" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value & "'"))
                    'Call oDS_PS_MM090L.setValue("U_Qty", pval.Row - 1, 1)
                    'Call oDS_PS_MM090L.setValue("U_Weight", pval.Row - 1, MDC_PS_Common.GetValue("SELECT Distinct Quantity FROM OIBT WHERE BatchNum='" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value & "'"))
                    'Call oDS_PS_MM090L.setValue("U_ProDate", pval.Row - 1, Format(MDC_PS_Common.GetValue("SELECT Distinct CreateDate FROM OIBT WHERE BatchNum='" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value & "'"), "YYYYMMDD"))
                    'Call oDS_PS_MM090L.setValue("U_ItmBsort", pval.Row - 1, "104")
                ElseIf (pval.ColUID = "JakName") Then
                    '//기타작업
                    Call oDS_PS_MM090L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) '//자기자신먼저 Flush처리
                    
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM090L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_MM090_AddMatrixRow (pval.Row)
                    End If
                    
                    sQry = "           SELECT   FrgnName AS [ItemName],"
                    sQry = sQry & "             U_Size AS [Size],"
                    sQry = sQry & "             InvntryUom AS [Unit],"
                    sQry = sQry & "             0 AS [Qty]"
                    sQry = sQry & " FROM    OITM"
                    sQry = sQry & " WHERE   ItemCode = '" & Trim(oMat01.Columns("JakName").Cells(pval.Row).Specific.VALUE) & "'"
                    
                    Call oRecordSet01.DoQuery(sQry)
                    oMat01.Columns("ItemName").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("ItemName").VALUE)
                    oMat01.Columns("Size").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("Size").VALUE)
                    oMat01.Columns("Unit").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("Unit").VALUE)
                    'oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("Qty").VALUE)
                
                ElseIf pval.ColUID = "DocLin" Then
                    '//기타작업
                    Call oDS_PS_MM090L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) '//자기자신먼저 Flush처리
                    
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM090L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_MM090_AddMatrixRow (pval.Row)
                    End If
                    
                    sQry = " EXEC [PS_MM090_90] '" & Trim(oMat01.Columns("DocLin").Cells(pval.Row).Specific.VALUE) & "'"
                    
                    Call oRecordSet01.DoQuery(sQry)
                    
                    Call oDS_PS_MM090L.setValue("U_JakName", pval.Row - 1, Trim(oRecordSet01.Fields("ItemCode").VALUE))
                    Call oDS_PS_MM090L.setValue("U_ItemName", pval.Row - 1, Trim(oRecordSet01.Fields("ItemName").VALUE))
                    Call oDS_PS_MM090L.setValue("U_Size", pval.Row - 1, Trim(oRecordSet01.Fields("OutSize").VALUE))
                    Call oDS_PS_MM090L.setValue("U_Unit", pval.Row - 1, Trim(oRecordSet01.Fields("OutUnit").VALUE))
                    Call oDS_PS_MM090L.setValue("U_Qty", pval.Row - 1, Trim(oRecordSet01.Fields("Qty").VALUE))
                    Call oDS_PS_MM090L.setValue("U_Weight", pval.Row - 1, Trim(oRecordSet01.Fields("Qty").VALUE))
                    
                    'oMat01.Columns("JakName").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("ItemCode").VALUE)
                    'oMat01.Columns("ItemName").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("ItemName").VALUE)
                    'oMat01.Columns("Size").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("OutSize").VALUE)
                    'oMat01.Columns("Unit").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("OutUnit").VALUE)
                    'oMat01.Columns("Qty").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("Qty").VALUE)
                    'oMat01.Columns("Weight").Cells(pval.Row).Specific.VALUE = Trim(oRecordSet01.Fields("Qty").VALUE)
                    
                Else
                    Call oDS_PS_MM090L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_MM090H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "CardCode") Then
                   ' Call oDS_PS_MM090H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
                    Call oDS_PS_MM090H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
                Else
                    'Call oDS_PS_MM090H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oForm01.Update
            If pval.ItemUID = "Mat01" Then oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
        End If
    
    ElseIf pval.BeforeAction = False Then
    
        If (pval.ItemUID = "Mat01") Then
        
            If (pval.ColUID = "Qty") Then
                Call PS_MM090_CalculateAmount(pval.ColUID, pval.Row)
            ElseIf (pval.ColUID = "Price") Then
                Call PS_MM090_CalculateAmount(pval.ColUID, pval.Row)
            ElseIf (pval.ColUID = "Amt") Then
                Call PS_MM090_CalculateAmount(pval.ColUID, pval.Row)
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
        Call PS_MM090_FormItemEnabled
        Call PS_MM090_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        Call oMat01.AutoResizeColumns
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_MM090_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "ItemCode") Then
'            Dim oDataTable01 As SAPbouiCOM.DataTable
'            Set oDataTable01 = pval.SelectedObjects
'            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_MM090H", "U_CardCode,U_CardName")
'        End If
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
        If pval.BeforeAction = True Then '//Matrix 행삭제전 행삭제가능여부검사타기
            If (PS_MM090_Validate("행삭제") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
            
        ElseIf pval.BeforeAction = False Then '//Matrix 행삭제후 다시 행번채번등처리
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i '//행을 다시 순서대로정렬해서 행번에넣고(VisualCount값은 줄어든상태)
            Next i
            
            '//Matrix의 RowCount값 갯수도 줄어든 수만큼 갱신처리 해주며
            oMat01.FlushToDataSource
            
            Call oDS_PS_MM090L.RemoveRecord(oDS_PS_MM090L.Size - 1) '//줄어든 행수만큼 DataSources값 갱신한 뒤
            oMat01.LoadFromDataSource   '//그후 다시 데이터소스를 읽어와 화면완성을 한다.
            
            If oMat01.RowCount = 0 Then '행이 없으면 한줄추가
                Call PS_MM090_AddMatrixRow(0)
            Else
                '현재행삭제한 행의PorNum값이 있는행지우면 넘어가고 없는 마지막행값지우면 한행추가
                If Trim(oDS_PS_MM090L.GetValue("U_ItemName", oMat01.RowCount - 1)) <> "" Then
                    Call PS_MM090_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_MM090_CreateItems() As Boolean
On Error GoTo PS_MM090_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM090H = oForm01.DataSources.DBDataSources("@PS_MM090H")
    Set oDS_PS_MM090L = oForm01.DataSources.DBDataSources("@PS_MM090L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported '메트릭스 한행선택반전기능사용시 이구문과   'Call oMat01.SelectRow(pval.Row, True, False) 가 클릭이벤트에서 처리해야함
    oMat01.AutoResizeColumns
    
'    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
'    Call oForm01.DataSources.UserDataSources.Add("WhsCode", dt_SHORT_TEXT, 100)
'    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
'    Call oForm01.Items("WhsCode").Specific.DataBind.SetBound(True, "", "WhsCode")
'    Call oForm01.DataSources.UserDataSources.Add("ApDate", dt_DATE) '1.송장전기일
'    Call oForm01.Items("ApDate").Specific.DataBind.SetBound(True, "", "ApDate") '2.송장전기일

'    oForm01.DataSources.UserDataSources.Item("FrDate").ValueEx = Format(Now, "yyyymm") & "01"
'    oForm01.DataSources.UserDataSources.Item("ToDate").ValueEx = Format(Now, "yyyymmdd")
    oQuery01 = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & MDC_PS_Common.User_MSTCOD & "'"
    oRecordSet01.DoQuery oQuery01

    oForm01.Items("OutMan").Specific.VALUE = oRecordSet01.Fields(0).VALUE
    oOutMan = oRecordSet01.Fields(0).VALUE
    
    oForm01.Items("InDate").Specific.VALUE = Format(Now, "yyyymmdd")
'    oForm01.Items("CntcCode").Specific.Value = "dd"
    
    

    Call oForm01.DataSources.UserDataSources.Add("BOM_CHECK", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BOM_CHECK").Specific.DataBind.SetBound(True, "", "BOM_CHECK")
    
    
    '//////////라디오버튼//////////S
'    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 10)
'    Call oForm01.Items("Opt01").Specific.DataBind.SetBound(True, "", "Opt01")
'
'    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 10)
'    Call oForm01.Items("Opt02").Specific.DataBind.SetBound(True, "", "Opt02")
'
    Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    Call oForm01.Items("Opt01").Click
    '//////////라디오버튼//////////E
    
    
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_MM090_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM090_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_MM090_ComboBox_Setting()
On Error GoTo PS_MM090_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    
    Dim sQry As String
    Dim CntcCode As String
    
    sQry = "        SELECT  U_Minor AS [Code],"
    sQry = sQry & "         U_CdName AS [Name]"
    sQry = sQry & " FROM    [@PS_SY001L]"
    sQry = sQry & " WHERE   Code = 'M008'"
    sQry = sQry & "         AND U_UseYN = 'Y'"
    
    CntcCode = MDC_PS_Common.User_MSTCOD
    
    If CntcCode = "0074514" Or CntcCode = "0672514" Or CntcCode = "1178504" Or CntcCode = "2050802" Then
    
        sQry = sQry & " UNION ALL"
        sQry = sQry & " SELECT  U_Minor AS [Code],"
        sQry = sQry & "         U_CdName AS [Name]"
        sQry = sQry & " FROM    [@PS_SY001L]"
        sQry = sQry & " WHERE   Code = 'M008'"
        sQry = sQry & "         AND U_UseYN = 'N'"
    
    End If
    
'    sQry = sQry & " ORDER BY U_Seq"
    
    Call oForm01.Items("OutCls").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("OutCls").Specific, sQry, "", False, False)
    
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM090", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM090", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_MM090", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM090", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM090", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_MM090", "ItemCode")

    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM090", "Title", "", "01", "송장")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM090", "Title", "", "02", "거래명세서")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Title").Specific, "PS_MM090", "Title")
    Call oForm01.Items("Title").Specific.Select(0, psk_Index)
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL]  order by BPLId", "1", False, False)
    
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Tonnage").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] Where Code = 'M009' and U_UseYN = 'Y'  order by U_Seq", "", True, True)
    
    
    oMat01.Columns("OutGbn").ValidValues.Add "", ""
    oMat01.Columns("OutGbn").ValidValues.Add "01", "기타출고"
    oMat01.Columns("OutGbn").ValidValues.Add "10", "MG스크랩출고"
    oMat01.Columns("OutGbn").ValidValues.Add "20", "MG원재료반납"
    oMat01.Columns("OutGbn").ValidValues.Add "30", "MG LOSS"
    oMat01.Columns("OutGbn").ValidValues.Add "40", "내부거래"
    oMat01.Columns("OutGbn").ValidValues.Add "50", "MG스크랩기타출고"
    
    '검사분류(기계)
    Call oMat01.Columns("ChkType").ValidValues.Add("", "")
    Call oMat01.Columns("ChkType").ValidValues.Add("PS_QM450", "검사여부등록")
    Call oMat01.Columns("ChkType").ValidValues.Add("PS_QM410", "검사성적서(A)등록")
    
    'Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("COL01"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")

    Call oForm01.Freeze(False)
    Exit Sub
PS_MM090_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM090_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM090_CF_ChooseFromList()
On Error GoTo PS_MM090_CF_ChooseFromList_Error
'    '//ChooseFromList 설정
'    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
'    Dim oCons               As SAPbouiCOM.Conditions
'    Dim oCon                As SAPbouiCOM.Condition
'    Dim oCFL                As SAPbouiCOM.ChooseFromList
'    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
'    Dim oEdit               As SAPbouiCOM.EditText
'    Dim oColumn             As SAPbouiCOM.Column
'
'    Set oEdit = oForm01.Items("CARDCODE").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_BusinessPartner
'    oCFLCreationParams.uniqueID = "CFLCARDCODE"
'    oCFLCreationParams.MultiSelection = False
''    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons
'
'    oEdit.ChooseFromListUID = "CFLCARDCODE"
'    oEdit.ChooseFromListAlias = "CardCode"
'
'    Set oEdit = oForm01.Items("CARDNAME").Specific
'    Set oCFLs = oForm01.ChooseFromLists
'    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
'
'    oCFLCreationParams.ObjectType = lf_BusinessPartner
'    oCFLCreationParams.uniqueID = "CFLCARDNAME"
'    oCFLCreationParams.MultiSelection = False
'    Set oCFL = oCFLs.Add(oCFLCreationParams)

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
PS_MM090_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM090_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM090_FormItemEnabled()
On Error GoTo PS_MM090_FormItemEnabled_Error

    Dim i As Integer
    Dim Query01 As String
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim oCombo01       As SAPbouiCOM.ComboBox
    Set oCombo = oForm01.Items("BPLId").Specific    '//사업장콤보
    Set oCombo01 = oForm01.Items("Title").Specific  '//타이틀콤보

    Dim User_BPLId As String
    User_BPLId = MDC_PS_Common.User_BPLId

    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        'Call oCombo.Select(0, psk_Index)        '//콤보기본선택
        Call oCombo.Select(User_BPLId, psk_ByValue) '//2010.12.06 추가
        Call oCombo01.Select(0, psk_Index)      '//콤보기본선택
        
        If User_BPLId = "2" Then '소속사업장이 부산사업장이면 반출구분을 필수로 선택
            Call oForm01.Items("OutCls").Specific.Select(0, psk_Index)
        Else
            Call oForm01.Items("OutCls").Specific.Select(1, psk_Index)
        End If
        
        '//반출자, 일자 표시
        oForm01.Items("OutMan").Specific.VALUE = oOutMan
        oForm01.Items("InDate").Specific.VALUE = Format(Now, "yyyymmdd")
        
        Call PS_MM090_FormClear '//UDO방식
        
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        
        oForm01.Items("CardCode").Click
        oForm01.Items("Btn1").Enabled = False
        oForm01.Items("Btn02").Enabled = False
        oForm01.Items("DocEntry").Enabled = False   '//비활성   문서번호
        oForm01.Items("CardCode").Enabled = True    '//활성     업체코드
        oForm01.Items("Mat01").Enabled = True       '//활성     메트릭스
        oForm01.Items("InDate").Enabled = True      '//활성     작성일
        oForm01.Items("BPLId").Enabled = True       '//활성     사업장
        oForm01.Items("OutMan").Enabled = True      '//활성     반출자
        oForm01.Items("RecMan").Enabled = True      '//활성     인수자
        oForm01.Items("PurPose").Enabled = True     '//활성     목적
        oForm01.Items("Destin").Enabled = True      '//활성     도착지
        oForm01.Items("TranCard").Enabled = True    '//활성     운송편
        oForm01.Items("TranCode").Enabled = True    '//활성     차량번호
        oForm01.Items("TranCost").Enabled = True    '//활성     운임
        oForm01.Items("Title").Enabled = True       '//활성     출력타이틀
        oForm01.Items("Opt01").Enabled = True
        oForm01.Items("Opt02").Enabled = True
        
        Call oForm01.Items("Opt01").Click
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        oForm01.Items("DocEntry").Enabled = True '//문서번호활성화
        oForm01.Items("CardCode").Enabled = True    '//활성     업체코드
        'oForm01.Items("Mat01").Enabled = True       '//활성     메트릭스
        oForm01.Items("InDate").Enabled = True      '//활성     작성일
        'oForm01.Items("BPLId").Enabled = True       '//활성     사업장
        oForm01.Items("OutMan").Enabled = True      '//활성     반출자
        'oForm01.Items("RecMan").Enabled = True      '//활성     인수자
        'oForm01.Items("PurPose").Enabled = True     '//활성     목적
        'oForm01.Items("Destin").Enabled = True      '//활성     도착지
        oForm01.Items("TranCard").Enabled = True    '//활성     운송편
        'oForm01.Items("TranCode").Enabled = True    '//활성     차량번호
        'oForm01.Items("TranCost").Enabled = True    '//활성     운임
        'oForm01.Items("Title").Enabled = True       '//활성     출력타이틀
        oForm01.Items("Opt01").Enabled = False
        oForm01.Items("Opt02").Enabled = False
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
'        Query01 = ""
'        Query01 = "Select Distinct Quantity FROM OIBT WHERE BatchNum = '" & oMat01.Columns("LotNo").Cells(1).Specific.Value & "'"
'        RecordSet01.DoQuery Query01
        'RecordSet01.Fields("Quantity").Value
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        If oDS_PS_MM090H.GetValue("Canceled", 0) = "Y" Then
            oForm01.Items("Btn1").Enabled = False '//출력버튼비활성
            oForm01.Items("Btn02").Enabled = False '//출력버튼비활성
            oForm01.Items("DocEntry").Enabled = False
            oForm01.Items("CardCode").Enabled = False
            oForm01.Items("Mat01").Enabled = False
            oForm01.Items("InDate").Enabled = False
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("OutMan").Enabled = False
            oForm01.Items("RecMan").Enabled = False
            oForm01.Items("PurPose").Enabled = False
            oForm01.Items("Destin").Enabled = False
            oForm01.Items("TranCard").Enabled = False
            oForm01.Items("TranCode").Enabled = False
            oForm01.Items("TranCost").Enabled = False
            oForm01.Items("Title").Enabled = False
            oForm01.Items("Opt01").Enabled = False
            oForm01.Items("Opt02").Enabled = False
        
        Else
            oForm01.Items("Btn1").Enabled = True
            oForm01.Items("Btn02").Enabled = True
            oForm01.Items("DocEntry").Enabled = False '찾기하고나면 문서비활성화처리
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("Mat01").Enabled = True
            oForm01.Items("InDate").Enabled = True
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("OutMan").Enabled = True
            oForm01.Items("RecMan").Enabled = True
            oForm01.Items("PurPose").Enabled = True
            oForm01.Items("Destin").Enabled = True
            oForm01.Items("TranCard").Enabled = True
            oForm01.Items("TranCode").Enabled = True
            oForm01.Items("TranCost").Enabled = True
            oForm01.Items("Title").Enabled = True
            oForm01.Items("Opt01").Enabled = False
            oForm01.Items("Opt02").Enabled = False
        End If
        
        If Trim(oDS_PS_MM090H.GetValue("U_Opt02", 0)) = "1" Then
            oMat01.Columns("DocLin").Visible = True
        ElseIf Trim(oDS_PS_MM090H.GetValue("U_Opt02", 0)) = "2" Then
            oMat01.Columns("DocLin").Visible = False
        End If
        

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM090_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM090_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM090_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_MM090_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM090L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM090L.Offset = oRow
    oDS_PS_MM090L.setValue "U_LineNum", oRow, oRow + 1
    '                        '필드명  '행  '기본값

    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM090_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM090_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM090_FormClear()
On Error GoTo PS_MM090_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM090'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_MM090_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM090_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM090_EnableMenus()
On Error GoTo PS_MM090_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    'Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_MM090_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM090_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM090_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_MM090_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_MM090_FormItemEnabled
        Call PS_MM090_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_MM090_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_MM090_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM090_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'//데이터 필수값등등 체크
Function PS_MM090_DataValidCheck() As Boolean
On Error GoTo PS_MM090_DataValidCheck_Error
    PS_MM090_DataValidCheck = False
    Dim i As Long

    If oForm01.Items("CardCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "고객코드는 필수입니다.", bmt_Short, True
        oForm01.Items("CardCode").Click ct_Regular
        PS_MM090_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("OutMan").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "반출자는 필수입니다.", bmt_Short, True
        oForm01.Items("OutMan").Click ct_Regular
        PS_MM090_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("InDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "반출일은 필수입니다.", bmt_Short, True
        oForm01.Items("InDate").Click ct_Regular
        PS_MM090_DataValidCheck = False
        Exit Function
        
     ElseIf oForm01.Items("OutCls").Specific.VALUE = "%" Then
        Sbo_Application.SetStatusBarMessage "반출구분은 필수입니다.", bmt_Short, True
        oForm01.Items("OutCls").Click ct_Regular
        PS_MM090_DataValidCheck = False
        Exit Function
        
    ElseIf oForm01.Items("PurPose").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "목적은 필수입니다.", bmt_Short, True
        oForm01.Items("PurPose").Click ct_Regular
        PS_MM090_DataValidCheck = False
        Exit Function
    
    End If
    


    
    If oMat01.VisualRowCount = 0 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_MM090_DataValidCheck = False
        Exit Function
    Else
        '값이 한줄들어있을때 한줄삭제후 갱신한다거나한다면
        If oMat01.Columns("ItemName").Cells(1).Specific.VALUE = "" Then
            Sbo_Application.SetStatusBarMessage "Matrix값이 한줄이상은 있어야합니다.", bmt_Short, True
            PS_MM090_DataValidCheck = False
            Exit Function
        End If
        
    End If
    For i = 1 To (oMat01.VisualRowCount - 1)
        If (oMat01.Columns("ItemName").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품명은 필수입니다.", bmt_Short, True
            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
            PS_MM090_DataValidCheck = False
            Exit Function
        End If
        If (oMat01.Columns("OutGbn").Cells(i).Specific.VALUE = "50") Then
            If Trim(oForm01.Items("PurPose").Specific.VALUE) = "" Then
                Sbo_Application.SetStatusBarMessage "스크랩기타출고는 목적(사유)를 필수입력해야합니다.", bmt_Short, True
                oForm01.Items("PurPose").Click ct_Regular
                PS_MM090_DataValidCheck = False
                Exit Function
            End If
        End If
        '내부거래시 적요란에 사용처입력해야함
        If (oMat01.Columns("OutGbn").Cells(i).Specific.VALUE = "40") Then
            If Trim(oMat01.Columns("LineNote").Cells(i).Specific.VALUE) = "" Then
                Sbo_Application.SetStatusBarMessage "내부거래시 적요란에 사용처를 필수입력해야합니다.", bmt_Short, True
                oMat01.Columns("LineNote").Cells(i).Click ct_Regular
                PS_MM090_DataValidCheck = False
                Exit Function
            End If
        End If
        
        If oForm01.Items("BPLId").Specific.VALUE <> "2" Then '기계사업부는 라인의 출고구분 필수 미적용
            If (oMat01.Columns("OutGbn").Cells(i).Specific.VALUE = "") Then
                Sbo_Application.SetStatusBarMessage "출고구분은 필수 입니다.", bmt_Short, True
                oMat01.Columns("OutGbn").Cells(i).Click ct_Regular
                PS_MM090_DataValidCheck = False
                Exit Function
            End If
        End If
        
        '기계사업부의 납품등록 시 필수 정보 체크(작번, 검사분류, 검사문서번호) 2016.06.17 송명규
        If oForm01.Items("BPLId").Specific.VALUE = "2" And (oForm01.Items("OutCls").Specific.VALUE = "01" Or oForm01.Items("OutCls").Specific.VALUE = "02" Or oForm01.Items("OutCls").Specific.VALUE = "06") Then '납품(제품), 납품(분할납품), 납품(미발주)
        
            Dim CheckData As String
            CheckData = PS_MM090_CheckChkData _
                             ( _
                                oForm01.Items("InDate").Specific.VALUE, _
                                oMat01.Columns("JakName").Cells(i).Specific.VALUE, _
                                oMat01.Columns("ChkType").Cells(i).Specific.VALUE, _
                                oMat01.Columns("ChkDoc").Cells(i).Specific.VALUE, _
                                oMat01.Columns("ChkNo").Cells(i).Specific.VALUE, _
                                oForm01.Items("DocEntry").Specific.VALUE _
                             )
        
            If (oMat01.Columns("JakName").Cells(i).Specific.VALUE) = "" Then
                Call Sbo_Application.SetStatusBarMessage(i & "행 작번을 입력하십시오.", bmt_Short, True)
                Call oMat01.Columns("JakName").Cells(i).Click(ct_Regular)
                PS_MM090_DataValidCheck = False
                Exit Function
            ElseIf (oMat01.Columns("ChkType").Cells(i).Specific.VALUE) = "" Then
                Call Sbo_Application.SetStatusBarMessage(i & "행 검사분류를 선택하십시오.", bmt_Short, True)
                Call oMat01.Columns("ChkType").Cells(i).Click(ct_Regular)
                PS_MM090_DataValidCheck = False
                Exit Function
            ElseIf (oMat01.Columns("ChkDoc").Cells(i).Specific.VALUE) = "" Then
                Call Sbo_Application.SetStatusBarMessage(i & "행 검사문서번호를 입력하십시오.", bmt_Short, True)
                Call oMat01.Columns("ChkDoc").Cells(i).Click(ct_Regular)
                PS_MM090_DataValidCheck = False
                Exit Function
            ElseIf (oMat01.Columns("ChkNo").Cells(i).Specific.VALUE) = "" Then
                Call Sbo_Application.SetStatusBarMessage(i & "행 성적서번호를 입력하십시오.", bmt_Short, True)
                Call oMat01.Columns("ChkNo").Cells(i).Click(ct_Regular)
                PS_MM090_DataValidCheck = False
                Exit Function
            ElseIf CheckData <> "True" Then '검사정보 체크
                Call Sbo_Application.SetStatusBarMessage(i & "행 작번[" & oMat01.Columns("JakName").Cells(i).Specific.VALUE & "]의 자료는 " & CheckData, bmt_Short, True)
                PS_MM090_DataValidCheck = False
                Exit Function
            End If
            
        End If
    Next
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_MM090L.RemoveRecord(oDS_PS_MM090L.Size - 1) 'UDO방식일때 저장되기전에(BeforeAction구문이 완결되기전) 한줄빼고 넣어주는 구문
    Call oMat01.LoadFromDataSource
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_MM090_FormClear
    End If
    PS_MM090_DataValidCheck = True
    Exit Function
PS_MM090_DataValidCheck_Error:
    PS_MM090_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_MM090_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_MM090_MTX01()
On Error GoTo PS_MM090_MTX01_Error
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
        GoTo PS_MM090_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_MM090L.InsertRecord (i)
        End If
        oDS_PS_MM090L.Offset = i
        Call oDS_PS_MM090L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_MM090L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_MM090_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_MM090_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM090_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_MM090_DI_API() As Boolean
On Error GoTo PS_MM090_DI_API_Error
    PS_MM090_DI_API = True
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
        '작지번호를 묶어서 내역에 넣기위함
        ItemInformation(ItemInformationCount).Dscription = oMat01.Columns("PorNum").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).LineTotal = Val(oMat01.Columns("Total").Cells(i).Specific.VALUE)
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE '지원받은품목
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("ReQty").Cells(i).Specific.VALUE    '지원받은수량
        'ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseInvoices)
    '문서헤더상태정의
    oDIObject.HandWritten = TNO
    oDIObject.DocType = dDocument_Service '서비스송장
    oDIObject.BPL_IDAssignedToInvoice = 2 '동래기본박아넣음 'oForm01.Items("BPLId").Specific.Selected.Value
    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)                       '공급업체
    oDIObject.CardName = Trim(oForm01.Items("CardName").Specific.VALUE)                       '공급처이름
    oDIObject.DocDate = Format(oForm01.Items("ApDate").Specific.VALUE, "&&&&-&&-&&")    '전기일
    oDIObject.DocDueDate = Format(oForm01.Items("ApDate").Specific.VALUE, "&&&&-&&-&&") '만기일
    oDIObject.Comments = "기준PSMT내역문서 : " & oForm01.Items("DocEntry").Specific.VALUE
    
    For i = 0 To ItemInformationCount - 1
        If ItemInformation(i).Check = True Then
            GoTo Continue_First
        End If
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemDescription = ItemInformation(i).Dscription                 '내역에 작지번호저장
        oDIObject.Lines.AccountCode = "73020010"                                        'G/L 계정 값 (외주임가공비)
        oDIObject.Lines.TaxCode = "V2"                                                  'VAT그룹코드
        oDIObject.Lines.LineTotal = ItemInformation(i).LineTotal                        '공정단가로계산된총금액
        oDIObject.Lines.UserFields("U_sItemCode").VALUE = ItemInformation(i).ItemCode   '지원받은품목
        oDIObject.Lines.UserFields("U_sQty").VALUE = ItemInformation(i).Quantity        '지원받은수량
        
Continue_First:
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        
        oForm01.Items("ApDocNum").Specific.VALUE = ResultDocNum '생성된송장문서를 현재문서에 담는다.
        'Call oDS_PS_MM090L.setValue("U_ApDocNum", 1, ResultDocNum)
'        For i = 0 To UBound(ItemInformation)
'            Call oDS_PS_MM090L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_MM090L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
    Else
        GoTo PS_MM090_DI_API_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_MM090_DI_API_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_MM090_DI_API = False
    Set oDIObject = Nothing
    Exit Function
PS_MM090_DI_API_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM090_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_MM090_DI_API = False
End Function

Private Sub PS_MM090_FormResize()
On Error GoTo PS_MM090_FormResize_Error
    
    Exit Sub
PS_MM090_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM090_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_MM090_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_MM090_Validate_Error
    PS_MM090_Validate = True
    Dim i, j As Long
    Dim sDocTotal As Long '문서총계
    
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
'                    PS_MM090_Validate = False
'                    GoTo PS_MM090_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
            If (oMat01.Columns("LineNum").Cells(oLastColRow01).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
            Else
'                If oForm01.Items("ApDocNum").Specific.Value <> "" Then
'                    MDC_Com.MDC_GF_Message "AP전기된문서입니다. 행삭제를 할수 없습니다.", "W"
'                    PS_MM090_Validate = False
'                    GoTo PS_MM090_Validate_Exit
'                End If
                
                If oForm01.Items("Canceled").Specific.VALUE = "Y" Then
                    MDC_Com.MDC_GF_Message "취소된문서는 수정할수 없습니다.", "W"
                    PS_MM090_Validate = False
                    GoTo PS_MM090_Validate_Exit
                End If

            End If
        End If
    ElseIf ValidateType = "취소" Then
'        If oForm01.Items("ApDocNum").Specific.Value <> "" Then
'            MDC_Com.MDC_GF_Message "AP전기된문서입니다. 취소할수 없습니다.", "W"
'            PS_MM090_Validate = False
'            GoTo PS_MM090_Validate_Exit
'        End If
        If oForm01.Items("Canceled").Specific.VALUE = "Y" Then
                    MDC_Com.MDC_GF_Message "이미취소된문서입니다.", "W"
                    PS_MM090_Validate = False
                    GoTo PS_MM090_Validate_Exit
        End If
        
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
'                PS_MM090_Validate = False
'                GoTo PS_MM090_Validate_Exit
'            End If
'            RecordSet01.MoveNext
'        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_MM090_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_MM090_Validate_Error:
    PS_MM090_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_MM090_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'DB데이터소스에스는 아래와같이 메트릭스에 것을 읽어와서 수정하거나 저장한다.
Public Function UpdateToPP090(Doc)
On Error GoTo UpdateToDBData_Error
    UpdateToPP090 = True
    Dim i As Integer
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String


  
        DocEntry = Trim(oDS_PS_MM090H.GetValue("DocEntry", 0))

        
        Query01 = "UPDATE"
        Query01 = Query01 & " [@PS_MM090H]"
        Query01 = Query01 & " SET"
        Query01 = Query01 & " U_PackNo = '" & Doc & "'"
'        Query01 = Query01 & " , U_ApDate = '" & Format(oDS_PS_MM090H.GetValue("U_ApDate", 0), "&&&&-&&-&&") & "'"
        Query01 = Query01 & " WHERE DocEntry = '" & DocEntry & "'"
        
        RecordSet01.DoQuery Query01
        
        '처리하고난뒤 화면을 잠금처리한다.
        oForm01.Mode = fm_OK_MODE
        oForm01.Items("Btn1").Enabled = False
        oForm01.Items("Btn02").Enabled = False
        oForm01.Items("OutMan").Enabled = False
        oForm01.Items("CardCode").Enabled = False
        oForm01.Items("InDate").Enabled = False
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("Mat01").Enabled = False
    
    
    Sbo_Application.SetStatusBarMessage "처리 완료되었습니다.", bmt_Short, False
    Exit Function
UpdateToDBData_Error:
    UpdateToPP090 = False
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage Err.Number & Err.Description, bmt_Short, True
End Function

Private Sub PS_SMM090_Print_Report01()
'******************************************************************************
'Function ID        : PS_SMM090_Print_Report01
'해당모듈           : PS_MM090
'기        능       : 레포트 출력
'인        수       : 없음
'반  환  값         : 없음
'특이사항           : 없음
'******************************************************************************
On Error GoTo PS_SMM090_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
    DocNum = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    WinTitle = "[PS_MM090] 기타자재출고증출력"
    ReportName = "PS_MM090_01.rpt"
    sQry = "EXEC PS_MM090_Print '" & DocNum & "'"
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
PS_SMM090_Print_Report01_Error:
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "PS_SMM090_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SMM090_Print_Report02()
'******************************************************************************
'Function ID        : PS_SMM090_Print_Report01
'해당모듈           : PS_MM090
'기        능       : 레포트 출력
'인        수       : 없음
'반  환  값         : 없음
'특이사항           : 없음
'******************************************************************************
On Error GoTo PS_SMM090_Print_Report02_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    DocNum = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    WinTitle = "[PS_MM090] 거래명세서"
    
    If oForm01.Items("BPLId").Specific.VALUE = "2" Then
        ReportName = "PS_MM090_03.rpt"
    Else
        ReportName = "PS_MM090_02.rpt"
    End If
    sQry = "EXEC PS_MM090_02 '" & DocNum & "'"
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
PS_SMM090_Print_Report02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Sbo_Application.SetStatusBarMessage "PS_SMM090_Print_Report02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM090_CalculateAmount(ByVal pColUID As String, ByVal pRowID As Integer)
'******************************************************************************
'Function ID : PS_MM090_CalculateAmount
'해당 모듈 : PS_MM090
'기       능 : 수량, 단가, 금액 입력 시 단가 및 입력 계산
'인       수 : pColUID : Validate Event가 발생되는 ColUID, pRowID : 행ID
'반  환  값 : 없음
'특이 사항 : 없음
'******************************************************************************
On Error GoTo PS_MM090_CalculateAmount_Error
    
    Dim l_ErrNum As Integer '에러번호
    
    Dim l_Qty As Double '수량
    Dim l_Price As Double '단가
    Dim l_Amount As Double '금액
       
    oMat01.FlushToDataSource
    
    l_Qty = Val(oMat01.Columns("Qty").Cells(pRowID).Specific.VALUE)
    l_Price = oMat01.Columns("Price").Cells(pRowID).Specific.VALUE
    l_Amount = oMat01.Columns("Amt").Cells(pRowID).Specific.VALUE
    
    If pColUID = "Qty" Then
        
        Call oDS_PS_MM090L.setValue("U_Amt", pRowID - 1, l_Qty * l_Price)
        
    ElseIf pColUID = "Price" Then
    
        Call oDS_PS_MM090L.setValue("U_Amt", pRowID - 1, l_Qty * l_Price)
    
    ElseIf pColUID = "Amt" Then
        
'        If l_Qty = 0 Then
'            l_ErrNum = 1
'            GoTo PS_MM090_CalculateAmount_Error
'        End If
        If l_Qty > 0 Then
            Call oDS_PS_MM090L.setValue("U_Price", pRowID - 1, l_Amount / l_Qty)
        End If
    
    End If
    
    oMat01.LoadFromDataSource
        
    Exit Sub
        
PS_MM090_CalculateAmount_Error:
    If l_ErrNum = 1 Then
        Sbo_Application.SetStatusBarMessage "수량을 0으로 입력하셨습니다. 확인하십시오.", bmt_Short, True
    Else
        Sbo_Application.SetStatusBarMessage "PS_MM090_CalculateAmount_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub

Private Function PS_MM090_CheckChkData(ByVal pDocDate As String, ByVal pOrdNum As String, ByVal pChkType As String, ByVal pChkDoc, ByVal pChkNo, ByVal pDocEntry) As String
'******************************************************************************
'Function ID : PS_MM090_CheckChkData
'해당 모듈 : PS_MM090
'기       능 : 반출정보 VS 검사정보 비교
'인       수 : pDocDate(반출일자), pOrdNum(작번), pChkType(검사분류), pChkDoc(검사문서번호), pChkNo(검사성적서번호)
'반  환  값 : 해당 오류 리턴
'특이 사항 : 없음
'******************************************************************************
On Error GoTo PS_MM090_CheckChkData_Error

    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim ReturnValue As String

    Query01 = "              EXEC PS_MM090_91 '"
    Query01 = Query01 & pDocDate & "','"
    Query01 = Query01 & pOrdNum & "','"
    Query01 = Query01 & pChkType & "','"
    Query01 = Query01 & pChkDoc & "','"
    Query01 = Query01 & pChkNo & "','"
    Query01 = Query01 & pDocEntry & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    PS_MM090_CheckChkData = RecordSet01.Fields(0).VALUE
    
    Set RecordSet01 = Nothing
    Exit Function

PS_MM090_CheckChkData_Error:
    Set RecordSet01 = Nothing
    Call Sbo_Application.SetStatusBarMessage("CheckDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function



