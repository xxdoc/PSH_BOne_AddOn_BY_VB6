VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM066"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_QM066.cls
'//  Module : 품질관리>(부품) 제품검사서 등록
'//  FormType : PS_QM066
'//  Create Date : 2014.01.08
'//  Modified Date :
'//  Creator : NGY
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM066H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_QM066L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    BatchNum As String
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
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM066.srf")
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
    
    oFormUniqueID01 = "PS_QM066_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM066_CreateItems
    Call PS_QM066_ComboBox_Setting
    Call PS_QM066_Initial_Setting
    Call PS_QM066_CF_ChooseFromList
    Call PS_QM066_EnableMenus
    Call PS_QM066_SetDocument(oFromDocEntry01)
    Call PS_QM066_FormResize
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1286"), True        '// 닫기
    oForm01.EnableMenu ("1284"), True        '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    
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

    Dim i As Integer

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
                Call PS_QM066_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_QM066_FormItemEnabled '//UDO방식
                Call PS_QM066_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_QM066_FormItemEnabled
            Case "1287"
            
                oForm01.Freeze True
                Call oDS_PS_QM066H.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat01.VisualRowCount - 1
                    Call oMat01.FlushToDataSource
                    Call oDS_PS_QM066H.setValue("DocEntry", i, "")
                    Call oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
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
        If pval.ItemUID = "PS_QM066" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_QM066_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_QM066_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Button01" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                Call PS_QM066_Print_Report01
            End If
        End If
        
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM066" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_QM066_FormItemEnabled
                    Call PS_QM066_AddMatrixRow(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_QM066_FormItemEnabled
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
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "CardCode" Then
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
    '
            ElseIf pval.ItemUID = "ODLNNum" Then
                
                If oForm01.Items("ODLNNum").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
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
    
        If pval.ItemChanged = True Then
            
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
    Dim sQry As String
    Dim ODLNNum As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim i As Long
    Dim YQty As Long
    Dim AQLYN As String
    
    Dim ItmMsort As String
    Dim OutSize As String
    
    Dim PS_PH030() As String

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then

            If (pval.ItemUID = "Mat01") Then
            
                If pval.ColUID = "FailQty" Then '작지문서
                
                    oMat01.FlushToDataSource
                    
                    Call oDS_PS_QM066L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    
                    YQty = oMat01.Columns("PQty").Cells(pval.Row).Specific.VALUE - oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE
                    
                    Call oDS_PS_QM066L.setValue("U_YQty", pval.Row - 1, YQty)  '합격수량
                    
                    
                    For i = pval.Row To oMat01.VisualRowCount - 1
                        If Trim(oDS_PS_QM066L.GetValue("U_InspItem", i)) <> "" Then
                            Call oDS_PS_QM066L.setValue("U_PQty", i, YQty)  '검사수량
                            Call oDS_PS_QM066L.setValue("U_YQty", i, YQty)  '합격수량
                        End If
                    Next
                    
                    oMat01.LoadFromDataSource
                    
                End If
                
'                Call oMat01.Columns(pval.ColUID).Cells(pval.Row).Click(ct_Regular)
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                oMat01.AutoResizeColumns
                
            ElseIf (pval.ItemUID = "CardCode") Then  '납품처
            
                Call oDS_PS_QM066H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("cardname", "cardcode", "[ocrd]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
                
            ElseIf (pval.ItemUID = "ODLNNum") Then  '납품문서번호
                oForm01.Freeze True
                
                ODLNNum = oForm01.Items("ODLNNum").Specific.VALUE
                
                sQry = " select a.ItemCode, a.Dscription, b.U_Spec4, b.U_ItmMsort from [DLN1] a Inner join OITM b On a.ItemCode = b.ItemCode "
                sQry = sQry + " where a.DocEntry = left('" & ODLNNum & "', charindex('-', '" & ODLNNum & "') -1) "
                sQry = sQry + " and a.LineNum  = Substring('" & ODLNNum & "', charindex('-', '" & ODLNNum & "') + 1, 10) "
                RecordSet01.DoQuery sQry
                oDS_PS_QM066H.setValue "U_ItemCode", 0, Trim(RecordSet01.Fields(0).VALUE)
                oDS_PS_QM066H.setValue "U_ItemName", 0, Trim(RecordSet01.Fields(1).VALUE)
                oDS_PS_QM066H.setValue "U_OutSize", 0, Trim(RecordSet01.Fields(2).VALUE)
                
                OutSize = Trim(RecordSet01.Fields(2).VALUE)
                ItmMsort = Trim(RecordSet01.Fields(3).VALUE)
                
                sQry = " select DocEntry, U_SizeNo, U_Sample1, U_Sample2, U_TAQL1, U_TAQL2 from [@PS_QM065H] "
                sQry = sQry + " where U_ItmMsort = '" & ItmMsort & "' and U_OutSize = '" & OutSize & "'"
                RecordSet01.DoQuery sQry
                
                               
                If RecordSet01.RecordCount >= 1 Then
                    oDS_PS_QM066H.setValue "U_SizeNo", 0, Trim(RecordSet01.Fields(1).VALUE)
                    oDS_PS_QM066H.setValue "U_Sample1", 0, Trim(RecordSet01.Fields(2).VALUE)
                    oDS_PS_QM066H.setValue "U_Sample2", 0, Trim(RecordSet01.Fields(3).VALUE)
                    oDS_PS_QM066H.setValue "U_TAQL1", 0, Trim(RecordSet01.Fields(4).VALUE)
                    oDS_PS_QM066H.setValue "U_TAQL2", 0, Trim(RecordSet01.Fields(5).VALUE)
                
                    Call LoadData(Trim(RecordSet01.Fields(0).VALUE))
                Else
                    sQry = " select DocEntry, U_SizeNo, U_Sample1, U_Sample2, U_TAQL1, U_TAQL2 from [@PS_QM065H] "
                    sQry = sQry + " where U_ItmMsort = '" & ItmMsort & "'"
                    RecordSet01.DoQuery sQry
                    
                    If RecordSet01.RecordCount >= 1 Then
                        oDS_PS_QM066H.setValue "U_SizeNo", 0, Trim(RecordSet01.Fields(1).VALUE)
                        oDS_PS_QM066H.setValue "U_Sample1", 0, Trim(RecordSet01.Fields(2).VALUE)
                        oDS_PS_QM066H.setValue "U_Sample2", 0, Trim(RecordSet01.Fields(3).VALUE)
                        oDS_PS_QM066H.setValue "U_TAQL1", 0, Trim(RecordSet01.Fields(4).VALUE)
                        oDS_PS_QM066H.setValue "U_TAQL2", 0, Trim(RecordSet01.Fields(5).VALUE)
                        Call LoadData(Trim(RecordSet01.Fields(0).VALUE))
                    End If
                End If
                
                oForm01.Freeze False
            ElseIf (pval.ItemUID = "PQty") Then  '검사수량
                oForm01.Freeze True
                
                oMat01.FlushToDataSource
                    
                    For i = 0 To oMat01.VisualRowCount - 1
                        If Trim(oDS_PS_QM066L.GetValue("U_InspItNm", i)) <> "" Then
                            Call oDS_PS_QM066L.setValue("U_Sample", i, oForm01.Items("PQty").Specific.VALUE)  '시료수
                            Call oDS_PS_QM066L.setValue("U_PQty", i, oForm01.Items("PQty").Specific.VALUE)  '검사수량
                            Call oDS_PS_QM066L.setValue("U_YQty", i, oForm01.Items("PQty").Specific.VALUE)  '합격수량
                        End If
                    Next
                    
                    oMat01.LoadFromDataSource
                
                oForm01.Freeze False
            ElseIf (pval.ItemUID = "ReqNo") Then '의뢰번호
                oForm01.Items("LotNo").Specific.VALUE = oForm01.Items("ReqNo").Specific.VALUE
            End If
        End If
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_QM066_FormItemEnabled
        Call PS_QM066_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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
        Call PS_QM066_FormResize
    End If
    Exit Sub
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
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
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_QM066L.RemoveRecord(oDS_PS_QM066L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_QM066_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_QM066L.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_QM066_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_QM066_CreateItems() As Boolean
On Error GoTo PS_QM066_CreateItems_Error
   
    Set oDS_PS_QM066H = oForm01.DataSources.DBDataSources("@PS_QM066H")
    Set oDS_PS_QM066L = oForm01.DataSources.DBDataSources("@PS_QM066L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Exit Function
PS_QM066_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM066_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_QM066_ComboBox_Setting()
On Error GoTo PS_QM066_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_QM066", "Mat01", "UseYN", "Y", "Y")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_QM066", "Mat01", "UseYN", "N", "N")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("UseYN"), "PS_QM066", "Mat01", "UseYN")

    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
PS_QM066_ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM066_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM066_CF_ChooseFromList()
On Error GoTo PS_QM066_CF_ChooseFromList_Error
    Exit Sub
PS_QM066_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM066_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM066_FormItemEnabled()
On Error GoTo PS_QM066_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
'
        'oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        Call PS_QM066_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM066_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM066_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM066_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM066_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM066L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_QM066L.Offset = oRow
    oDS_PS_QM066L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM066_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM066_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM066_FormClear()
On Error GoTo PS_QM066_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM066'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_QM066_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM066_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM066_Initial_Setting()
On Error GoTo PS_QM066_Initial_Setting_Error

    ' 사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

    '납품처
    'oForm01.Items("CardCode").Specific.VALUE = "12521"
       
    'oForm01.Items("AQL1").Specific.VALUE = "100" '중결점
    'oForm01.Items("AQL2").Specific.VALUE = "100" '경결점
    
    '일자
    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")

    Exit Sub
PS_QM066_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM066_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM066_EnableMenus()
On Error GoTo PS_QM066_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_QM066_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM066_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM066_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM066_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM066_FormItemEnabled
        Call PS_QM066_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_QM066_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_QM066_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM066_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_QM066_DataValidCheck() As Boolean
On Error GoTo PS_QM066_DataValidCheck_Error
    PS_QM066_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_QM066_FormClear
    End If
    
'    '검사자 미입력시
'    If oForm01.Items("ChkWkCd").Specific.VALUE = "" Then
'        Call Sbo_Application.SetStatusBarMessage("검사자가 입력되지 않았습니다.", bmt_Short, True)
'        PS_QM066_DataValidCheck = False
'        Exit Function
'    End If
    
    '일자 미입력시
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("일자가 입력되지 않았습니다.", bmt_Short, True)
        PS_QM066_DataValidCheck = False
        Exit Function
    End If
    
    '라인정보 미입력 시
    If oMat01.VisualRowCount <= 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM066_DataValidCheck = False
        Exit Function
    End If
    
    If oForm01.Items("YQty").Specific.VALUE <> oMat01.Columns("YQty").Cells(oMat01.VisualRowCount - 1).Specific.VALUE Then
        Sbo_Application.SetStatusBarMessage "합격수량이 일치하지 않습니다.", bmt_Short, True
        PS_QM066_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        
        '검사항목
        If (oMat01.Columns("InspItNm").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "검사항목명은 필수입니다.", bmt_Short, True
            oMat01.Columns("InspItem").Cells(i).Click ct_Regular
            PS_QM066_DataValidCheck = False
            Exit Function
        End If
'
'        '검진구분
'        If (oMat01.Columns("ChkCls").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "검진구분은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ChkCls").Cells(i).Click ct_Regular
'            PS_QM066_DataValidCheck = False
'            Exit Function
'        End If
        
    Next
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_QM066L.RemoveRecord(oDS_PS_QM066L.Size - 1)
    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_QM066_FormClear
    End If
    
    PS_QM066_DataValidCheck = True
    Exit Function
PS_QM066_DataValidCheck_Error:
    PS_QM066_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_QM066_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PS_QM066_MTX01()
On Error GoTo PS_QM066_MTX01_Error
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
        GoTo PS_QM066_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_QM066L.InsertRecord (i)
        End If
        oDS_PS_QM066L.Offset = i
        Call oDS_PS_QM066L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_QM066L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_QM066_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_QM066_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM066_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_QM066_FormResize()
On Error GoTo PS_QM066_FormResize_Error
    
    Exit Sub
PS_QM066_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM066_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_QM066_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_QM066_Validate_Error
    PS_QM066_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If ValidateType = "수정" Then
        '//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
    ElseIf ValidateType = "취소" Then
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_QM066_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_QM066_Validate_Error:
    PS_QM066_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_QM066_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Public Sub LoadData(ByVal oCode As String)
On Error GoTo LoadData_Error
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim i As Integer
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select b.U_Page, b.U_InspItem, b.U_InspItNm, b.U_InspSpec"
    sQry = sQry & " From [@PS_QM065H] a INNER JOIN [@PS_QM065L] b ON a.Code = b.Code "
    sQry = sQry & "Where a.Code = '" & oCode & "' Order By b.U_Seqno "
    oRecordSet01.DoQuery sQry
    
    oDS_PS_QM066L.Clear
    oMat01.Clear
    oMat01.FlushToDataSource
    
    i = 0
    Do Until oRecordSet01.EOF
        oDS_PS_QM066L.InsertRecord i
        oDS_PS_QM066L.setValue "U_LineNum", i, i + 1
        oDS_PS_QM066L.setValue "U_Page", i, Trim(oRecordSet01.Fields(0).VALUE)
        oDS_PS_QM066L.setValue "U_InspItem", i, Trim(oRecordSet01.Fields(1).VALUE)
        oDS_PS_QM066L.setValue "U_InspItNm", i, Trim(oRecordSet01.Fields(2).VALUE)
        oDS_PS_QM066L.setValue "U_InspSpec", i, Trim(oRecordSet01.Fields(3).VALUE)
        i = i + 1
        oRecordSet01.MoveNext
    Loop
    
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub PS_QM066_Print_Report01()
On Error GoTo PS_QM066_Print_Report01_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim i As Long
    Dim sQry01                      As String
    Dim TAQL1                       As String
    Dim TAQL2                       As String
    Dim CardCode                    As String
    Dim ItmMsort                    As String
    Dim Comments                    As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call ConnectODBC
    
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    ItmMsort = Left(oForm01.Items("ItemCode").Specific.VALUE, 5)
    
    sQry01 = "Select Isnull(U_TAQL1,''), Isnull(U_TAQL2,'') From [@PS_QM066H] Where DocEntry = '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
    oRecordSet01.DoQuery sQry01
    
    TAQL1 = oRecordSet01.Fields(0).VALUE
    TAQL2 = oRecordSet01.Fields(1).VALUE
    
    If CardCode = "12575" Then
       If ItmMsort = "10807" Then
          WinTitle = "[PS_QM066_10] 검사기록서출력_동래철선묶음상자"
          ReportName = "PS_QM066_10.rpt"
          sQry = "EXEC PS_QM066_10 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
       ElseIf ItmMsort = "10802" Then
              WinTitle = "[PS_QM066_11] 검사기록서출력_동래파래트"
              ReportName = "PS_QM066_11.rpt"
              sQry = "EXEC PS_QM066_11 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
       ElseIf ItmMsort = "10801" Then
              WinTitle = "[PS_QM066_12] 검사기록서출력_동래목상자"
              ReportName = "PS_QM066_12.rpt"
              sQry = "EXEC PS_QM066_12 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
       End If
    Else
       If TAQL1 = "" And TAQL2 = "" Then
          WinTitle = "[PS_QM066_01] 검사기록서 출력"
          ReportName = "PS_QM066_01.rpt"
          sQry = "EXEC PS_QM066_01 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
       Else
          WinTitle = "[PS_QM066_02] 검사기록서 출력"
          ReportName = "PS_QM066_02.rpt"
          sQry = "EXEC PS_QM066_01 '" & oForm01.Items("DocEntry").Specific.VALUE & "'"
       End If
    End If
    
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
PS_QM066_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM066_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub





