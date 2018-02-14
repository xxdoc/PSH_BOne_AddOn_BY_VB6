VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM092"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM092.cls
'//  Module         : 분말샘플출고등록
'//  Desc           : 분말 팩킹등록 후 분말 샘플출고
'//  FormType       : PS_MM092
'//  Create Date    : 2016.5.30
'//  Editor         : NGY
'//  Copyright  (c) Morning Data
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_MM092H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_MM092L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    BatchNum As String
    Quantity As Long
    WhsCode As String
    LineNum As Long
    
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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM092.srf")
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
    
    oFormUniqueID01 = "PS_MM092_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_MM092_CreateItems
    Call PS_MM092_ComboBox_Setting
    Call PS_MM092_CF_ChooseFromList
    Call PS_MM092_EnableMenus
    Call PS_MM092_SetDocument(oFromDocEntry01)
    Call PS_MM092_FormResize
    
    'oForm01.Items("empty").Visible = False
    
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
                    If (PS_MM092_Validate("취소") = False) Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If Sbo_Application.MessageBox("정말로 취소하시겠습니까?", "1", "예", "아니오") <> "1" Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    If PS_MM092_DI_API02 = False Then
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
                Call PS_MM092_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_MM092_FormItemEnabled '//UDO방식
                Call PS_MM092_AddMatrixRow(0, True) '//UDO방식 추가하고 나면 메트릭스에 한행추가한다.
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_MM092_FormItemEnabled
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
        If pval.ItemUID = "PS_MM092" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_MM092_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                '//출고
                'Addon만등록 시 주석_S
                If PS_MM092_DI_API01 = False Then
                    Call PS_MM092_AddMatrixRow(oMat01.VisualRowCount)
                    BubbleEvent = False
                    Exit Sub
                End If
                'Addon만등록 시 주석_E
                
                oOutMan = Trim(oForm01.Items("OutMan").Specific.VALUE)
                
    
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_MM092_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
                oOutMan = Trim(oForm01.Items("OutMan").Specific.VALUE)
            ElseIf oForm01.Mode = fm_OK_MODE Then
                
            ElseIf oForm01.Mode = fm_FIND_MODE Then
                 
            End If
        ElseIf (pval.ItemUID = "Btn1") Then '출고증출력시 출력으로 액션
        
            Call PS_SMM092_Print_Report01
             
        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_MM092" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_MM092_FormItemEnabled
                    Call PS_MM092_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_MM092_FormItemEnabled
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
            If pval.ItemUID = "CardCode" Then
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "OutMan" Then
                If oForm01.Items("OutMan").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "TranCard" Then
                If oForm01.Items("TranCard").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
        '헤더에 질의연결이벤트
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "")        '//거래처코드
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OutMan", "")        '//반출자조회
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "TranCard", "")        '//차량조회
        'Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "")        '//품목코드
        
        '메트릭스에 질의연결이벤트
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "LotNo")     '//Lot번호
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "CpCode")     '//공정코드
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "WorkMan")    '//작업자정보
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "PP092No")    '//분말팩킹번호
        
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
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        If pval.ItemChanged = True Then
            Call oForm01.Freeze(True)
            If (pval.ItemUID = "Mat01") Then
                
            Else
                If (pval.ItemUID = "Tonnage") Then
                    '창원사업장은 톤수에 따라 운송비를 입력(MG스크랩 운송비때문에) 2016/2/28 N.G.Y 추가
                    If oForm01.Items("BPLId").Specific.VALUE = "1" Then
                        sQry = " Select U_RelCd From [@PS_SY001L] WHERE Code = 'M009' and U_Minor = '" & oForm01.Items("Tonnage").Specific.Selected.VALUE & "'"
                        Call oRecordSet01.DoQuery(sQry)
                        Call oDS_PS_MM092H.setValue("U_TranCost", 0, Trim(oRecordSet01.Fields(0).VALUE))
                        
                        Call oDS_PS_MM092H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
                    End If
                Else
                    Call oDS_PS_MM092H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Selected.VALUE)
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
                If pval.ColUID = "PP092No" Then
                    '//기타작업
                    Call oDS_PS_MM092L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) '//자기자신먼저 Flush처리
                    
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_MM092L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        PS_MM092_AddMatrixRow (pval.Row)
                    End If
                    
                    sQry = " select  a.DocEntry,"
                    sQry = sQry + " b.LineId,"
                    sQry = sQry + " b.U_ItemCode,"
                    sQry = sQry + " c.FrgnName,"
                    sQry = sQry + " c.U_Size,"
                    sQry = sQry + " c.SalUnitMsr,"
                    sQry = sQry + " b.U_LotNo,"
                    sQry = sQry + " b.U_Qty,"
                    sQry = sQry + " B.U_Weight"
                    sQry = sQry + " from [@PS_PP092H] a Inner Join [@PS_PP092L] b ON a.DocEntry = b.DocEntry and a.Canceled = 'N'"
                    sQry = sQry + " Inner Join OITM c On b.U_ItemCode = c.ItemCode"
                    sQry = sQry + " Where b.U_OutGbn = '20'"
                    sQry = sQry + " AND Convert(Nvarchar(10),a.DocEntry) + '-' + Convert(Nvarchar(10),b.LineId) = '" & Trim(oMat01.Columns("PP092No").Cells(pval.Row).Specific.VALUE) & "'"
  
                    Call oRecordSet01.DoQuery(sQry)
                    
                    Call oDS_PS_MM092L.setValue("U_PP092HNo", pval.Row - 1, Trim(oRecordSet01.Fields(0).VALUE))
                    Call oDS_PS_MM092L.setValue("U_PP092LNo", pval.Row - 1, Trim(oRecordSet01.Fields(1).VALUE))
                    Call oDS_PS_MM092L.setValue("U_ItemCode", pval.Row - 1, Trim(oRecordSet01.Fields(2).VALUE))
                    Call oDS_PS_MM092L.setValue("U_ItemName", pval.Row - 1, Trim(oRecordSet01.Fields(3).VALUE))
                    Call oDS_PS_MM092L.setValue("U_Size", pval.Row - 1, Trim(oRecordSet01.Fields(4).VALUE))
                    Call oDS_PS_MM092L.setValue("U_Unit", pval.Row - 1, Trim(oRecordSet01.Fields(5).VALUE))
                    Call oDS_PS_MM092L.setValue("U_LotNo", pval.Row - 1, Trim(oRecordSet01.Fields(6).VALUE))
                    Call oDS_PS_MM092L.setValue("U_Qty", pval.Row - 1, Trim(oRecordSet01.Fields(7).VALUE))
                    Call oDS_PS_MM092L.setValue("U_Weight", pval.Row - 1, Trim(oRecordSet01.Fields(8).VALUE))
                    
                Else
                    Call oDS_PS_MM092L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                End If
            Else
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_MM092H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "CardCode") Then
                    Call oDS_PS_MM092H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
                Else
                    'Call oDS_PS_MM092H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.Value)
                End If
            End If
            oMat01.LoadFromDataSource
            oMat01.AutoResizeColumns
            oForm01.Update
            If pval.ItemUID = "Mat01" Then oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
        End If
    
    ElseIf pval.BeforeAction = False Then
    
        If (pval.ItemUID = "Mat01") Then
                        
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
        Call PS_MM092_FormItemEnabled
        Call PS_MM092_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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
        Call PS_MM092_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_MM092H", "U_CardCode,U_CardName")
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
            If (PS_MM092_Validate("행삭제") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
            
        ElseIf pval.BeforeAction = False Then '//Matrix 행삭제후 다시 행번채번등처리
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i '//행을 다시 순서대로정렬해서 행번에넣고(VisualCount값은 줄어든상태)
            Next i
            
            '//Matrix의 RowCount값 갯수도 줄어든 수만큼 갱신처리 해주며
            oMat01.FlushToDataSource
            
            Call oDS_PS_MM092L.RemoveRecord(oDS_PS_MM092L.Size - 1) '//줄어든 행수만큼 DataSources값 갱신한 뒤
            oMat01.LoadFromDataSource   '//그후 다시 데이터소스를 읽어와 화면완성을 한다.
            
            If oMat01.RowCount = 0 Then '행이 없으면 한줄추가
                Call PS_MM092_AddMatrixRow(0)
            Else
                '현재행삭제한 행의PorNum값이 있는행지우면 넘어가고 없는 마지막행값지우면 한행추가
                If Trim(oDS_PS_MM092L.GetValue("U_PP092No", oMat01.RowCount - 1)) <> "" Then
                    Call PS_MM092_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_MM092_CreateItems() As Boolean
On Error GoTo PS_MM092_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM092H = oForm01.DataSources.DBDataSources("@PS_MM092H")
    Set oDS_PS_MM092L = oForm01.DataSources.DBDataSources("@PS_MM092L")
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
    
    
    '//////////라디오버튼//////////S
'    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 10)
'    Call oForm01.Items("Opt01").Specific.DataBind.SetBound(True, "", "Opt01")
'
'    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 10)
'    Call oForm01.Items("Opt02").Specific.DataBind.SetBound(True, "", "Opt02")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_MM092_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM092_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_MM092_ComboBox_Setting()
On Error GoTo PS_MM092_ComboBox_Setting_Error
    Call oForm01.Freeze(True)
    
    Dim sQry As String
    Dim CntcCode As String
    
    
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM092", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM092", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_MM092", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM092", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM092", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_MM092", "ItemCode")

    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM092", "Title", "", "01", "송장")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM092", "Title", "", "02", "거래명세서")
    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Title").Specific, "PS_MM092", "Title")
    Call oForm01.Items("Title").Specific.Select(0, psk_Index)
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL]  order by BPLId", "1", False, False)
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM092_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM092_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM092_CF_ChooseFromList()
On Error GoTo PS_MM092_CF_ChooseFromList_Error
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
PS_MM092_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM092_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM092_FormItemEnabled()
On Error GoTo PS_MM092_FormItemEnabled_Error

    Dim i As Integer
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
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
        
        '//담당자 일자 표시
        oForm01.Items("OutMan").Specific.VALUE = oOutMan
        oForm01.Items("InDate").Specific.VALUE = Format(Now, "yyyymmdd")
        
        Call PS_MM092_FormClear '//UDO방식
        
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        
        'oForm01.Items("empty").Click
        oForm01.Items("Btn1").Enabled = False
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
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
'        Query01 = ""
'        Query01 = "Select Distinct Quantity FROM OIBT WHERE BatchNum = '" & oMat01.Columns("LotNo").Cells(1).Specific.Value & "'"
'        RecordSet01.DoQuery Query01
        'RecordSet01.Fields("Quantity").Value
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        If oDS_PS_MM092H.GetValue("Canceled", 0) = "Y" Then
            oForm01.Items("Btn1").Enabled = False '//출력버튼비활성
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
        
        Else
            oForm01.Items("Btn1").Enabled = True
            oForm01.Items("DocEntry").Enabled = False '찾기하고나면 문서비활성화처리
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("Mat01").Enabled = True
            oForm01.Items("InDate").Enabled = False
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("OutMan").Enabled = True
            oForm01.Items("RecMan").Enabled = True
            oForm01.Items("PurPose").Enabled = True
            oForm01.Items("Destin").Enabled = True
            oForm01.Items("TranCard").Enabled = True
            oForm01.Items("TranCode").Enabled = True
            oForm01.Items("TranCost").Enabled = True
            oForm01.Items("Title").Enabled = True
        End If
        
    End If
     
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM092_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM092_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM092_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_MM092_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM092L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM092L.Offset = oRow
    oDS_PS_MM092L.setValue "U_LineNum", oRow, oRow + 1
    '                        '필드명  '행  '기본값

    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_MM092_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM092_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_MM092_FormClear()
On Error GoTo PS_MM092_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM092'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_MM092_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM092_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM092_EnableMenus()
On Error GoTo PS_MM092_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    'Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_MM092_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM092_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_MM092_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_MM092_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_MM092_FormItemEnabled
        Call PS_MM092_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_MM092_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_MM092_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM092_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'//데이터 필수값등등 체크
Function PS_MM092_DataValidCheck() As Boolean
On Error GoTo PS_MM092_DataValidCheck_Error
    PS_MM092_DataValidCheck = False
    Dim i As Long

    If oForm01.Items("CardCode").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "고객코드는 필수입니다.", bmt_Short, True
        oForm01.Items("CardCode").Click ct_Regular
        PS_MM092_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("OutMan").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "반출자는 필수입니다.", bmt_Short, True
        oForm01.Items("OutMan").Click ct_Regular
        PS_MM092_DataValidCheck = False
        Exit Function
    ElseIf oForm01.Items("InDate").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "반출일은 필수입니다.", bmt_Short, True
        oForm01.Items("InDate").Click ct_Regular
        PS_MM092_DataValidCheck = False
        Exit Function
     
    ElseIf oForm01.Items("PurPose").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "목적은 필수입니다.", bmt_Short, True
        oForm01.Items("PurPose").Click ct_Regular
        PS_MM092_DataValidCheck = False
        Exit Function
    
    End If
    


    
    If oMat01.VisualRowCount = 0 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_MM092_DataValidCheck = False
        Exit Function
    Else
        '값이 한줄들어있을때 한줄삭제후 갱신한다거나한다면
        If oMat01.Columns("ItemCode").Cells(1).Specific.VALUE = "" Then
            Sbo_Application.SetStatusBarMessage "Matrix값이 한줄이상은 있어야합니다.", bmt_Short, True
            PS_MM092_DataValidCheck = False
            Exit Function
        End If
        
    End If
    For i = 1 To (oMat01.VisualRowCount - 1)
        If (oMat01.Columns("ItemName").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "품명은 필수입니다.", bmt_Short, True
            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
            PS_MM092_DataValidCheck = False
            Exit Function
        End If
        
    Next
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_MM092L.RemoveRecord(oDS_PS_MM092L.Size - 1) 'UDO방식일때 저장되기전에(BeforeAction구문이 완결되기전) 한줄빼고 넣어주는 구문
    Call oMat01.LoadFromDataSource
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_MM092_FormClear
    End If
    PS_MM092_DataValidCheck = True
    Exit Function
PS_MM092_DataValidCheck_Error:
    PS_MM092_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_MM092_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PS_MM092_MTX01()
On Error GoTo PS_MM092_MTX01_Error
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
        GoTo PS_MM092_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_MM092L.InsertRecord (i)
        End If
        oDS_PS_MM092L.Offset = i
        Call oDS_PS_MM092L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_MM092L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_MM092_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_MM092_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_MM092_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Private Sub PS_MM092_FormResize()
On Error GoTo PS_MM092_FormResize_Error
    
    Exit Sub
PS_MM092_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_MM092_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_MM092_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_MM092_Validate_Error
    PS_MM092_Validate = True
    Dim i, j As Long
    Dim sDocTotal As Long '문서총계
    
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
            If (oMat01.Columns("LineNum").Cells(oLastColRow01).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
            Else
                
                If oForm01.Items("Canceled").Specific.VALUE = "Y" Then
                    MDC_Com.MDC_GF_Message "취소된문서는 수정할수 없습니다.", "W"
                    PS_MM092_Validate = False
                    GoTo PS_MM092_Validate_Exit
                End If

            End If
        End If
    ElseIf ValidateType = "취소" Then
        If oForm01.Items("Canceled").Specific.VALUE = "Y" Then
                    MDC_Com.MDC_GF_Message "이미취소된문서입니다.", "W"
                    PS_MM092_Validate = False
                    GoTo PS_MM092_Validate_Exit
        End If
        
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_MM092_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_MM092_Validate_Error:
    PS_MM092_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_MM092_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PS_SMM092_Print_Report01()
'******************************************************************************
'Function ID        : PS_SMM090_Print_Report01
'해당모듈           : PS_MM092
'기        능       : 레포트 출력
'인        수       : 없음
'반  환  값         : 없음
'특이사항           : 없음
'******************************************************************************
On Error GoTo PS_SMM092_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    
    
    
    Call ConnectODBC
    
    DocNum = oForm01.Items("DocEntry").Specific.VALUE
    WinTitle = "[PS_MM092] 기타자재출고증출력"
    ReportName = "PS_MM092_01.rpt"
    sQry = "EXEC PS_MM092_01 '" & DocNum & "'"
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
PS_SMM092_Print_Report01_Error:
    
    
    Sbo_Application.SetStatusBarMessage "PS_SMM092_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


'//샘플출고
Private Function PS_MM092_DI_API01() As Boolean
On Error GoTo PS_MM092_DI_API01_Error
    PS_MM092_DI_API01 = True
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
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Weight").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).WhsCode = "101"
        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("LotNo").Cells(i).Specific.VALUE
        ItemInformationCount = ItemInformationCount + 1
    Next
        
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenExit)
    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.VALUE, "&&&&-&&-&&")
    oDIObject.UserFields("U_CardCode").VALUE = oForm01.Items("CardCode").Specific.VALUE
    oDIObject.UserFields("U_CardName").VALUE = oForm01.Items("CardName").Specific.VALUE
    oDIObject.UserFields("U_IssueTyp").VALUE = "4" '//샘플
    For i = 0 To ItemInformationCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = ItemInformation(i).WhsCode
        oDIObject.Lines.Quantity = ItemInformation(i).Quantity
        If MDC_PS_Common.GetItem_ManBtchNum(ItemInformation(i).ItemCode) = "Y" Then '//배치사용품목이면
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(i).BatchNum
            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(i).Quantity
            oDIObject.Lines.BatchNumbers.Add
        End If
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        'oForm01.Items("OIGENo").Specific.VALUE = ResultDocNum
        Call oDS_PS_MM092H.setValue("U_OIGENo", 0, ResultDocNum)
        
    Else
        GoTo PS_MM092_DI_API01_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_MM092_DI_API01_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_MM092_DI_API01 = False
    Set oDIObject = Nothing
    Exit Function
PS_MM092_DI_API01_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_MM092_DI_API01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_MM092_DI_API01 = False
End Function


'//샘플출고취소(입고)
Private Function PS_MM092_DI_API02() As Boolean
On Error GoTo PS_MM092_DI_API02_Error
    PS_MM092_DI_API02 = True
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
    
    Query01 = "SELECT IGE1.ItemCode,IGE1.Quantity,IGE1.WhsCode,(SELECT BatchNum FROM [IBT1_LINK] WHERE BaseType = '60' AND BaseEntry = OIGE.DocEntry AND BaseLinNum = IGE1.LineNum) AS BatchNum "
    Query01 = Query01 & " FROM [OIGE] OIGE LEFT JOIN [IGE1] IGE1 ON OIGE.DocEntry = IGE1.DocEntry WHERE OIGE.DocEntry = '" & oForm01.Items("OIGENo").Specific.VALUE & "'"
    
    RecordSet01.DoQuery Query01
    
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenEntry)
    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.VALUE, "&&&&-&&-&&")
    
    oDIObject.UserFields("U_CardCode").VALUE = oForm01.Items("CardCode").Specific.VALUE
    oDIObject.UserFields("U_CardName").VALUE = oForm01.Items("CardName").Specific.VALUE
    oDIObject.UserFields("U_IssueTyp").VALUE = "4" '//샘플
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = RecordSet01.Fields("ItemCode").VALUE
        oDIObject.Lines.WarehouseCode = RecordSet01.Fields("WhsCode").VALUE
        oDIObject.Lines.Quantity = RecordSet01.Fields("Quantity").VALUE
        oDIObject.UserFields("U_CancDoc").VALUE = oForm01.Items("OIGENo").Specific.VALUE '//입고(출고취소) 문서번호
        
        If MDC_PS_Common.GetItem_ManBtchNum(RecordSet01.Fields("ItemCode").VALUE) = "Y" Then '//배치관리품목이면
            oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet01.Fields("BatchNum").VALUE
            oDIObject.Lines.BatchNumbers.Quantity = RecordSet01.Fields("Quantity").VALUE
            oDIObject.Lines.BatchNumbers.Add
        End If
        
        RecordSet01.MoveNext
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        Call MDC_PS_Common.DoQuery("UPDATE [@PS_MM092H] SET U_OIGNNo = '" & ResultDocNum & "' WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'")
        RecordSet01.MoveFirst
    Else
        GoTo PS_MM092_DI_API02_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set RecordSet01 = Nothing
    Set oDIObject = Nothing
    Exit Function
PS_MM092_DI_API02_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_MM092_DI_API02 = False
    Set oDIObject = Nothing
    Set RecordSet01 = Nothing
    Exit Function
PS_MM092_DI_API02_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_MM092_DI_API02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_MM092_DI_API02 = False
    Set RecordSet01 = Nothing
End Function
