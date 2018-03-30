VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP560"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP560.cls
'//  Module         : 생산관리 > 애드온 > 장비이익률관리(기계)
'//  Desc           : 장비별 계획금액 등록
'//  FormType       : PS_PP560
'//  Create Date    : 2015.10.21
'//  Modified Date  :
'//  Creator        : Song Myoung gyu
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_PP560H  As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP560L  As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
'Private Type ItemInformations
'    ItemCode As String
'    BatchNum As String
'    Quantity As Long
'    OPORNo As Long
'    POR1No As Long
'    Check As Boolean
'    OPDNNo As Long
'    PDN1No As Long
'End Type
'Private ItemInformation() As ItemInformations
'Private ItemInformationCount As Long

Private oDocType01  As String
Private oDocEntry01 As String
Private oFormMode01 As SAPbouiCOM.BoFormMode

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP560.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP560_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP560_CreateItems
    Call PS_PP560_ComboBox_Setting
    Call PS_PP560_CF_ChooseFromList
    Call PS_PP560_EnableMenus
    Call PS_PP560_SetDocument(oFromDocEntry01)
    Call PS_PP560_FormResize
    
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
                Call PS_PP560_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_PP560_FormItemEnabled '//UDO방식
                Call PS_PP560_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_PP560_FormItemEnabled
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

    Dim ErrorCode As String

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "PS_PP560" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        End If
        
        If pval.ItemUID = "1" Then
        
            If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
        
                If Sbo_Application.MessageBox("계획금액은 등록 후 0원 이외의 금액에 대해서 수정이 불가능합니다. 정말로 저장하시겠습니까?", "1", "예", "아니오") = "1" Then
                
                    '작번 중복 체크
                    If oForm01.Mode = fm_ADD_MODE Then '최초 입력할 때만 체크
                    
                        If PS_PP560_DuplicateCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    
                    End If
                
                    If PS_PP560_DataValidCheck = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    '//해야할일 작업
                    oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                    oFormMode01 = oForm01.Mode
                
                Else

                    BubbleEvent = False
                    Exit Sub

                End If
            
            ElseIf oForm01.Mode = fm_OK_MODE Then
            
            End If
        
        ElseIf pval.ItemUID = "BtnSrch" Then
        
            If oForm01.Items("OrdNum").Specific.VALUE <> "" Then
            
                Call PS_PP560_MTX01
                Call PS_PP560_AddMatrixRow(oMat01.RowCount)
                
            Else
            
                ErrorCode = "1"
                GoTo Raise_EVENT_ITEM_PRESSED_Error
                BubbleEvent = False
            
            End If
         
        End If
        
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "PS_PP560" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        End If
        
        If pval.ItemUID = "1" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If pval.ActionSuccess = True Then
                
                    Call PS_PP560_FormItemEnabled
                    Call PS_PP560_AddMatrixRow(0, True) '//UDO방식일때
                    
                End If
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            
                If pval.ActionSuccess = True Then
                    Call PS_PP560_FormItemEnabled
                End If
                
            End If
        End If
        
        If pval.ItemUID = "Mat01" Then
        
            If oForm01.Mode <> fm_ADD_MODE Then
            
                If oMat01.RowCount >= pval.Row Then     '빈 Select 필드에 클릭했을 때 생기는 오류 수정을 위한 구문
        
                    Call PS_PP560_UpdateCheck(oForm01.Items("DocEntry").Specific.VALUE, oMat01.Columns("LineNum").Cells(pval.Row).Specific.VALUE, pval.ColUID)
                
                End If
            
            End If
        End If
        
    End If
    
    Exit Sub
    
Raise_EVENT_ITEM_PRESSED_Error:
    
    If ErrorCode = "1" Then

        Call MDC_Com.MDC_GF_Message("조회하려는 작번을 입력하세요.", "E")
        Call oForm01.Items("OrdNum").Click

'    ElseIf ErrCode = "2" Then '작번 중복 에러

    Else
        Call Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
    End If
    
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.ColUID = "BatchNum" Then
            
                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "BatchNum")
            
            End If
        
        Else
        
            If pval.ItemUID = "OrdNum" Then
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OrdNum", "")
            
            End If
        
        End If
    
'        Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "BatchNum")
    
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
    
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemChanged = True Then

            If (pval.ItemUID = "Mat01") Then

                Call oMat01.FlushToDataSource

                If (pval.ColUID = "AmtCls") Then
                    '//기타작업
                    Call oDS_PS_PP560L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)

                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP560L.GetValue("U_AmtCls", pval.Row - 1)) <> "" Then
                        PS_PP560_AddMatrixRow (pval.Row)
                    End If
                Else
                    Call oDS_PS_PP560L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
                End If

                Call oMat01.LoadFromDataSource
                Call oMat01.AutoResizeColumns

            Else

            End If

            Call oForm01.Update

'            If pval.ItemUID = "Mat01" Then
'                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
'            Else
'                oForm01.Items(pval.ItemUID).Click ct_Regular
'            End If

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
'        If pval.ItemUID = "Mat01" Then
'            If pval.Row > 0 Then
'                Call oMat01.SelectRow(pval.Row, True, False)
'            End If
'        End If
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat01.SelectRow(pval.Row, True, False)
            End If
        Else
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = ""
            oLastColRow01 = 0
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
            If pval.ColUID = "OrdNum" Then
                Set oTempClass = New PS_PP030
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

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
        
            If pval.ItemUID = "Mat01" Then
            
                Call oMat01.FlushToDataSource
            
                If pval.ColUID = "AmtCls" Then
                
'                    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'
'                    Query01 = "EXEC PS_PP560_01 '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'"
'                    Call RecordSet01.DoQuery(Query01)
'
'                    Call oDS_PS_PP560L.setValue("U_BatchNum", pval.Row - 1, RecordSet01.Fields("BatchNum").VALUE)
'                    Call oDS_PS_PP560L.setValue("U_OrdNum", pval.Row - 1, RecordSet01.Fields("OrdNum").VALUE)
'                    Call oDS_PS_PP560L.setValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields("ItemCode").VALUE)
'                    Call oDS_PS_PP560L.setValue("U_ItemName", pval.Row - 1, RecordSet01.Fields("ItemName").VALUE)
'                    Call oDS_PS_PP560L.setValue("U_Wgt", pval.Row - 1, RecordSet01.Fields("Wgt").VALUE)

                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP560L.GetValue("U_AmtType", pval.Row - 1)) <> "" Then '비용구분을 선택하면
                        Call PS_PP560_AddMatrixRow(pval.Row)
                    End If

                ElseIf pval.ColUID = "QuotAmt" Then

                    Call PS_PP560_SUM("QuotAmt")

                ElseIf pval.ColUID = "OrdAmt" Then

                    Call PS_PP560_SUM("OrdAmt")

                ElseIf pval.ColUID = "TrgtAmt" Then

                    Call PS_PP560_SUM("TrgtAmt")
                
                End If
                        
                Call oMat01.LoadFromDataSource
                Call oMat01.AutoResizeColumns
                
            Else
            
                If (pval.ItemUID = "DocEntry") Then
                    Call oDS_PS_PP560H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                ElseIf (pval.ItemUID = "OrdNum") Then
                    Call oDS_PS_PP560H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                    Call oDS_PS_PP560H.setValue("U_ItemName", 0, MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")) '품명
                    Call oDS_PS_PP560H.setValue("U_ItemSpec", 0, MDC_GetData.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")) '규격
                Else
                    Call oDS_PS_PP560H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
                End If
                
                Call oForm01.Items(pval.ItemUID).Click(ct_Regular)
            
            End If

            oForm01.Update
            
'            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_PP560L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_PP560L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_PP560_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_PP560L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
'                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
'            Else
'                If (pval.ItemUID = "DocEntry") Then
'                    Call oDS_PS_PP560H.setValue(pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
'                ElseIf (pval.ItemUID = "CardCode") Then
'                    Call oDS_PS_PP560H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
'                    Call oDS_PS_PP560H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
'                Else
'                    Call oDS_PS_PP560H.setValue("U_" & pval.ItemUID, 0, oForm01.Items(pval.ItemUID).Specific.VALUE)
'                End If
'                oForm01.Items(pval.ItemUID).Click ct_Regular
'            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
'            If pval.ItemUID = "Mat01" Then
'                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
'            Else
'                oForm01.Items(pval.ItemUID).Click ct_Regular
'            End If
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
        Call PS_PP560_FormItemEnabled
        Call PS_PP560_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP560_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP560H", "U_CardCode,U_CardName")
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
        If pval.BeforeAction = True Then
            If (PS_PP560_Validate("행삭제") = False) Then
                BubbleEvent = False
                Exit Sub
            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
        
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            
            oMat01.FlushToDataSource
            Call oDS_PS_PP560L.RemoveRecord(oDS_PS_PP560L.Size - 1)
            Call PS_PP560_SUM
            oMat01.LoadFromDataSource
            Call oForm01.Update
            
            If oMat01.RowCount = 0 Then
                Call PS_PP560_AddMatrixRow(0)
            Else
            
                If Trim(oDS_PS_PP560L.GetValue("U_AmtCls", oMat01.RowCount - 1)) <> "" Then
                    Call PS_PP560_AddMatrixRow(oMat01.RowCount)
                End If
                
            End If
            
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_PP560_CreateItems() As Boolean
On Error GoTo PS_PP560_CreateItems_Error
   
    Set oDS_PS_PP560H = oForm01.DataSources.DBDataSources("@PS_PP560H")
    Set oDS_PS_PP560L = oForm01.DataSources.DBDataSources("@PS_PP560L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
'    oMat01.AutoResizeColumns
    
    Exit Function
PS_PP560_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP560_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP560_ComboBox_Setting()
On Error GoTo PS_PP560_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
'    Set oCombo = oForm01.Items("BPLId").Specific
'    sQry = "SELECT BPLId, BPLName From [OBPL]"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
    
    '//////////Matrix//////////_S
    '비용구분
    sQry = "        SELECT      T0.U_Minor, "
    sQry = sQry & "             T0.U_CdName"
    sQry = sQry & " FROM        [@PS_SY001L] T0"
    sQry = sQry & " WHERE       T0.Code = 'P210'"
    sQry = sQry & "             AND T0.U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    T0.U_Seq"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("AmtCls"), sQry)
    '//////////Matrix//////////_E
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_PP560_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP560_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP560_CF_ChooseFromList()
On Error GoTo PS_PP560_CF_ChooseFromList_Error
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
PS_PP560_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP560_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP560_FormItemEnabled()
On Error GoTo PS_PP560_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정

        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        oForm01.Items("BtnSrch").Enabled = True
        Call PS_PP560_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        
'        oForm01.Items("DocDate").Specific.VALUE = Format(Now - 1, "YYYYMMDD")
'        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        oForm01.Items("BtnSrch").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        oForm01.Items("BtnSrch").Enabled = False
        
    End If
    
    Call oMat01.AutoResizeColumns
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP560_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP560_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP560_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP560_AddMatrixRow_Error

    Dim MaxLineUID As Integer
    Dim loopCount As Integer

    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP560L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    
    oDS_PS_PP560L.Offset = oRow
    Call oDS_PS_PP560L.setValue("U_LineNum", oRow, oRow + 1)
    
    If oForm01.Mode = fm_ADD_MODE Then
    
        Call oDS_PS_PP560L.setValue("U_LineUID", oRow, oRow + 1)
        
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
    
        MaxLineUID = PS_PP560_MatrixMaxColumnValue 'Max LineUID 값 조회(DB가 아닌 현재 화면 기준)
        
        Call oDS_PS_PP560L.setValue("U_LineUID", oRow - 1, MaxLineUID + 1)
        
    End If
    
    oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP560_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP560_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP560_FormClear()
On Error GoTo PS_PP560_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP560'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP560_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP560_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP560_EnableMenus()
On Error GoTo PS_PP560_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_PP560_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP560_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP560_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP560_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP560_FormItemEnabled
        Call PS_PP560_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_PP560_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_PP560_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP560_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_PP560_DataValidCheck() As Boolean
On Error GoTo PS_PP560_DataValidCheck_Error
    PS_PP560_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP560_FormClear
    End If
    
    '작번 미입력 시
    If oForm01.Items("OrdNum").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("작번을 입력하지 않았습니다.", bmt_Short, True)
        PS_PP560_DataValidCheck = False
        Exit Function
    End If
    
    '라인정보 미입력 시
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_PP560_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
'        If (oMat01.Columns("BatchNum").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "온산LotNo는 필수입니다.", bmt_Short, True
'            oMat01.Columns("BatchNum").Cells(i).Click ct_Regular
'            PS_PP560_DataValidCheck = False
'            Exit Function
'        End If
        
'        If (oMat01.Columns("BatchNum").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "온산LotNo는 필수입니다.", bmt_Short, True
'            oMat01.Columns("BatchNum").Cells(i).Click ct_Regular
'            PS_PP560_DataValidCheck = False
'            Exit Function
'        End If
    Next
    
    
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_PP560_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_PP560_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_PP560_DataValidCheck = False
'            Exit Function
'        End If
'    Next

    Call oDS_PS_PP560L.RemoveRecord(oDS_PS_PP560L.Size - 1)
    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_PP560_FormClear
    End If
    
    PS_PP560_DataValidCheck = True
    Exit Function
PS_PP560_DataValidCheck_Error:
    PS_PP560_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP560_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP560_MTX01()
On Error GoTo PS_PP560_MTX01_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    
    Dim loopCount   As Long
    Dim Query01     As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim OrdNum As String
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
    
    Query01 = "         EXEC [PS_PP560_01] '"
    Query01 = Query01 & OrdNum & "'" '작번
    
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP560_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
    
        If loopCount <> 0 Then
            Call oDS_PS_PP560L.InsertRecord(loopCount)
        End If
        oDS_PS_PP560L.Offset = loopCount
        
        Call oDS_PS_PP560L.setValue("U_LineNum", loopCount, loopCount + 1)
        Call oDS_PS_PP560L.setValue("U_AmtCls", loopCount, RecordSet01.Fields("AmtCls").VALUE) '비용구분
        Call oDS_PS_PP560L.setValue("U_OrdNum", loopCount, RecordSet01.Fields("OrdNum").VALUE) '작번
        Call oDS_PS_PP560L.setValue("U_OrdSub1", loopCount, RecordSet01.Fields("OrdSub1").VALUE) '서브작번1
        Call oDS_PS_PP560L.setValue("U_OrdSub2", loopCount, RecordSet01.Fields("OrdSub2").VALUE) '서즈작번2
        Call oDS_PS_PP560L.setValue("U_OrdName", loopCount, RecordSet01.Fields("OrdName").VALUE) '작명
        Call oDS_PS_PP560L.setValue("U_QuotAmt", loopCount, RecordSet01.Fields("QuotAmt").VALUE) '견적금액
        Call oDS_PS_PP560L.setValue("U_OrdAmt", loopCount, RecordSet01.Fields("OrdAmt").VALUE) '수주금액
        Call oDS_PS_PP560L.setValue("U_TrgtAmt", loopCount, RecordSet01.Fields("TrgtAmt").VALUE) '목표금액
        Call oDS_PS_PP560L.setValue("U_Comment", loopCount, RecordSet01.Fields("Comment").VALUE) '비고
        Call oDS_PS_PP560L.setValue("U_LineUID", loopCount, loopCount + 1) 'Line Unique ID
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    '합계금액 초기화
    Call oDS_PS_PP560H.setValue("U_TQuotAmt", 0, 0) '견적금액
    Call oDS_PS_PP560H.setValue("U_TOrdAmt", 0, 0) '수주금액
    Call oDS_PS_PP560H.setValue("U_TTrgtAmt", 0, 0) '목표금액
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    ProgressBar01.Stop
    
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
    
PS_PP560_MTX01_Exit:

    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    
    Exit Sub
    
PS_PP560_MTX01_Error:

    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Sbo_Application.SetStatusBarMessage "PS_PP560_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Sub

'Private Function PS_PP560_DI_API() As Boolean
'On Error GoTo PS_PP560_DI_API_Error
'    PS_PP560_DI_API = True
'    Dim i, j As Long
'    Dim oDIObject As SAPbobsCOM.Documents
'    Dim RetVal As Long
'    Dim LineNumCount As Long
'    Dim ResultDocNum As Long
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Company.StartTransaction
'
'    ReDim ItemInformation(0)
'    ItemInformationCount = 0
'    For i = 1 To oMat01.VisualRowCount
'        ReDim Preserve ItemInformation(ItemInformationCount)
'        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
'        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.VALUE
'        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.VALUE
'        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.VALUE
'        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.VALUE
'        ItemInformation(ItemInformationCount).Check = False
'        ItemInformationCount = ItemInformationCount + 1
'    Next
'
'    LineNumCount = 0
'    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
'    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
'    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
'    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.VALUE, "&&&&-&&-&&")
'    For i = 0 To ItemInformationCount - 1
'        If ItemInformation(i).Check = True Then
'            GoTo Continue_First
'        End If
'        If i <> 0 Then
'            oDIObject.Lines.Add
'        End If
'        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
'        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
'        oDIObject.Lines.BaseType = "22"
'        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
'        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
'        For j = i To UBound(ItemInformation)
'            If ItemInformation(j).Check = True Then
'                GoTo Continue_Second
'            End If
'            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
'                GoTo Continue_Second
'            End If
'            '//같은것
'            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
'            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
'            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
'            oDIObject.Lines.BatchNumbers.Add
'            ItemInformation(j).PDN1No = LineNumCount
'            ItemInformation(j).Check = True
'Continue_Second:
'        Next
'        LineNumCount = LineNumCount + 1
'Continue_First:
'    Next
'    RetVal = oDIObject.Add
'    If RetVal = 0 Then
'        ResultDocNum = Sbo_Company.GetNewObjectKey
'        For i = 0 To UBound(ItemInformation)
'            Call oDS_PS_PP560L.setValue("U_OPDNNo", i, ResultDocNum)
'            Call oDS_PS_PP560L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
'        Next
'    Else
'        GoTo PS_PP560_DI_API_Error
'    End If
'
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_Commit
'    End If
'    oMat01.LoadFromDataSource
'    oMat01.AutoResizeColumns
'
'    Set oDIObject = Nothing
'    Exit Function
'PS_PP560_DI_API_DI_Error:
'    If Sbo_Company.InTransaction = True Then
'        Sbo_Company.EndTransaction wf_RollBack
'    End If
'    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
'    PS_PP560_DI_API = False
'    Set oDIObject = Nothing
'    Exit Function
'PS_PP560_DI_API_Error:
'    Sbo_Application.SetStatusBarMessage "PS_PP560_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    PS_PP560_DI_API = False
'End Function

Private Sub PS_PP560_FormResize()
On Error GoTo PS_PP560_FormResize_Error
    
    Exit Sub
PS_PP560_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP560_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_PP560_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_PP560_Validate_Error
    PS_PP560_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP560H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
'        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
'        PS_PP560_Validate = False
'        GoTo PS_PP560_Validate_Exit
'    End If
'
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
'                    PS_PP560_Validate = False
'                    GoTo PS_PP560_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
            If (oMat01.Columns("LineNum").Cells(oLastColRow01).Specific.VALUE = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
            Else
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(oForm01.Items("8").Specific.VALUE) & "' AND U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.VALUE) & "'", 0, 1)) > 0 Then
                    Call MDC_Com.MDC_GF_Message("해당 내용은 행삭제할수 없습니다.", "W")
                    PS_PP560_Validate = False
                    GoTo PS_PP560_Validate_Exit
'                End If
            End If
        End If
    ElseIf ValidateType = "취소" Then
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
'                PS_PP560_Validate = False
'                GoTo PS_PP560_Validate_Exit
'            End If
'            RecordSet01.MoveNext
'        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_PP560_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_PP560_Validate_Error:
    PS_PP560_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_PP560_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP560_SUM(Optional ByVal pAmtType As String)
'******************************************************************************
'Function ID : PS_PP560_SUM()
'해당모듈    : PS_PP560
'기능        : 금액 합계
'인수        : 합계를 구할 금액 종류(QuotAmt : 견적금액, OrdAmt : 수주금액, TrgtAmt : 목표금액)
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP560_SUM_Error

    Dim loopCount As Integer
    Dim TotalQuotAmt As Double '견적금액
    Dim TotalOrdAmt As Double '수주금액
    Dim TotalTrgtAmt As Double '목표금액
    
    For loopCount = 1 To oMat01.VisualRowCount - 1
    
        If pAmtType = "QuotAmt" Then
        
            TotalQuotAmt = TotalQuotAmt + oDS_PS_PP560L.GetValue("U_QuotAmt", loopCount - 1) '견적금액
        
        ElseIf pAmtType = "OrdAmt" Then
        
            TotalOrdAmt = TotalOrdAmt + oDS_PS_PP560L.GetValue("U_OrdAmt", loopCount - 1) '수주금액
        
        ElseIf pAmtType = "TrgtAmt" Then
        
            TotalTrgtAmt = TotalTrgtAmt + oDS_PS_PP560L.GetValue("U_TrgtAmt", loopCount - 1) '목표금액
        
        ElseIf pAmtType = "" Then
        
            TotalQuotAmt = TotalQuotAmt + oDS_PS_PP560L.GetValue("U_QuotAmt", loopCount - 1) '견적금액
            TotalOrdAmt = TotalOrdAmt + oDS_PS_PP560L.GetValue("U_OrdAmt", loopCount - 1) '수주금액
            TotalTrgtAmt = TotalTrgtAmt + oDS_PS_PP560L.GetValue("U_TrgtAmt", loopCount - 1) '목표금액
        
        End If
        
    Next
    
    If pAmtType = "QuotAmt" Then
    
        Call oDS_PS_PP560H.setValue("U_TQuotAmt", 0, TotalQuotAmt) '견적금액
        
    ElseIf pAmtType = "OrdAmt" Then
    
        Call oDS_PS_PP560H.setValue("U_TOrdAmt", 0, TotalOrdAmt) '수주금액
        
    ElseIf pAmtType = "TrgtAmt" Then
    
        Call oDS_PS_PP560H.setValue("U_TTrgtAmt", 0, TotalTrgtAmt) '목표금액
        
    ElseIf pAmtType = "" Then
        
        Call oDS_PS_PP560H.setValue("U_TQuotAmt", 0, TotalQuotAmt) '견적금액
        Call oDS_PS_PP560H.setValue("U_TOrdAmt", 0, TotalOrdAmt) '수주금액
        Call oDS_PS_PP560H.setValue("U_TTrgtAmt", 0, TotalTrgtAmt) '목표금액
        
    End If
    
    Exit Sub

PS_PP560_SUM_Error:
    
    Call Sbo_Application.SetStatusBarMessage("PS_PP560_SUM_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Private Sub PS_PP560_UpdateCheck(ByVal pDocEntry As Integer, ByVal pLineNum As Integer, ByVal pColUID As String)
'******************************************************************************
'Function ID : PS_PP560_UpdateCheck()
'해당모듈    : PS_PP560
'기능        : 견적, 수주, 목표금액 수정 불가
'인수        : pDocEntry : 문서번호, pLineNum : 라인Num, pColUID : 클릭한 컬럼ID
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP560_UpdateCheck_Error

    Dim sQry        As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim ErrorCode   As String
    
    Dim DocEntry    As Integer
    Dim LineNum     As Integer
    Dim ColUID      As String
    
    DocEntry = pDocEntry
    LineNum = pLineNum
    ColUID = pColUID
    
    sQry = "      EXEC [PS_PP560_91] '"
    sQry = sQry & DocEntry & "','" '문서번호
    sQry = sQry & LineNum & "','" '라인Num
    sQry = sQry & ColUID & "'" '컬럼ID
    
    Call RecordSet01.DoQuery(sQry)
    
    If RecordSet01.Fields("ReturnValue").VALUE > 0 Then
    
        Call MDC_Com.MDC_GF_Message("해당금액은 수정이 불가능합니다.", "E")
        Call oForm01.Items("Comment").Click
        Exit Sub
    
    End If
    
    Set RecordSet01 = Nothing
    
    Exit Sub

PS_PP560_UpdateCheck_Error:

    Set RecordSet01 = Nothing
    Call Sbo_Application.SetStatusBarMessage("PS_PP560_UpdateCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
    
End Sub

Function PS_PP560_DuplicateCheck() As Boolean
'******************************************************************************
'Function ID : PS_PP560_DuplicateCheck()
'해당모듈    : PS_PP560
'기능        : 작번 중복 체크
'인수        : 없음
'반환값      : 작번중복 있을 시 False, 작번중복 없을 시 : True
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP560_DuplicateCheck_Error

    PS_PP560_DuplicateCheck = False
    
    Dim sQry        As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim OrdNum As String
    
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
    
    sQry = "      EXEC [PS_PP560_02] '"
    sQry = sQry & OrdNum & "'" '작번
        
    Call RecordSet01.DoQuery(sQry)
    
    If RecordSet01.Fields("Cnt").VALUE > 0 Then
    
        Call Sbo_Application.SetStatusBarMessage("해당 작번은 이미 등록되어 있습니다. 확인하세요.", bmt_Short, True)
        PS_PP560_DuplicateCheck = False
        Exit Function
    
    End If
        
    Set RecordSet01 = Nothing
    PS_PP560_DuplicateCheck = True
    Exit Function
    
PS_PP560_DuplicateCheck_Error:

    Set RecordSet01 = Nothing
    PS_PP560_DuplicateCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP560_DuplicateCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Function

Function PS_PP560_MatrixMaxColumnValue() As Integer
'******************************************************************************
'Function ID : PS_PP560_MatrixMaxColumnValue()
'해당모듈    : PS_PP560
'기능        : 현재 매트릭스의 특정 컬럼에 최댓값을 반환
'인수        : 없음
'반환값      : 컬럼내의 최댓값
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP560_MatrixMaxColumnValue_Error

    Dim loopCount   As Integer
    
    Dim LargeValue  As Integer
    Dim MaxValue    As Integer
    
    MaxValue = 0

    For loopCount = 1 To oMat01.VisualRowCount - 1
    
        If loopCount > 1 Then
        
            If oMat01.Columns("LineUID").Cells(loopCount - 1).Specific.VALUE < oMat01.Columns("LineUID").Cells(loopCount).Specific.VALUE Then
                LargeValue = oMat01.Columns("LineUID").Cells(loopCount).Specific.VALUE
            Else
                LargeValue = oMat01.Columns("LineUID").Cells(loopCount - 1).Specific.VALUE
            End If
            
            If LargeValue > MaxValue Then
                MaxValue = LargeValue
            End If
            
        End If
        
    Next

    PS_PP560_MatrixMaxColumnValue = MaxValue
    
    Exit Function
    
PS_PP560_MatrixMaxColumnValue_Error:

    PS_PP560_MatrixMaxColumnValue = 0
    Call Sbo_Application.SetStatusBarMessage("PS_PP560_MatrixMaxColumnValue_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)

End Function
