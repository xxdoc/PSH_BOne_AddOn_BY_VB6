VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR412"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_HR412.cls
'//  Module         : 인사관리
'//  Desc           : 전문직평가->전문직 평가 그룹핑
'//  FormType       : PS_HR412
'//  Create Date    : 2012.10.19
'//  Creator        : N.G.Y
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
'Public oMat02 As SAPbouiCOM.Matrix
''Private oDS_PS_HR412H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_HR412L As SAPbouiCOM.DBDataSource '등록라인

'Public oBaseForm01 As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01 As String
Public oBaseColUID01 As String
Public oBaseColRow01 As Long
Public oBaseTradeType01 As String

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String  '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long    '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oMat01Row01 As Long
Private oText(4)    As String
'Private oMat02Row02 As Long

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
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR412.srf")
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
    
    oFormUniqueID01 = "PS_HR412_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_HR412_CreateItems
    Call PS_HR412_ComboBox_Setting
    Call PS_HR412_Initial_Setting
'    Call PS_HR412_CF_ChooseFromList
    Call PS_HR412_EnableMenus
    Call PS_HR412_SetDocument(oFromDocEntry01)
    Call PS_HR412_FormResize
    
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
                '//Call PS_HR412_FormItemEnabled '//UDO방식
            Case "1282": '추가
                '//Call PS_HR412_FormItemEnabled '//UDO방식
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
        If pval.ItemUID = "Btn01" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PasswordChk(pval) = False Then
                    Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
                    oForm01.Items("PassWd").Click ct_Regular
                Else
                    Call PS_HR412_MTX01
                End If
                    
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "Btn02" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PasswordChk(pval) = False Then
                    Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
                    oForm01.Items("PassWd").Click ct_Regular
                Else
                    Call PS_HR412_SetBaseForm("수정") '//수정
                    Call PS_HR412_MTX01
                End If
                
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "Btn03" Then '//평가자 그룹핑 삭제
            If PasswordChk(pval) = False Then
                Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
                oForm01.Items("PassWd").Click ct_Regular
            Else
                Call PS_HR412_SetBaseForm("삭제") '//삭제
                Call PS_HR412_MTX01
            End If
            
        End If
        
        
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If PS_HR412_DataValidCheck = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'                '//해야할일 작업
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'            End If
'        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_HR412" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_HR412_FormItemEnabled
'                    Call PS_HR412_FormClear '//UDO방식일때
'                    Call PS_HR412_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_HR412_FormItemEnabled
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
        If pval.CharPressed = 9 Then
            If pval.ItemUID = "Year" Then
                If oForm01.Items("Year").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "MSTCOD" Then
                If oForm01.Items("MSTCOD").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "TeamCode" Then
                If oForm01.Items("TeamCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "RspCode" Then
                If oForm01.Items("RspCode").Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.ItemUID = "ClsCode" Then
                If oForm01.Items("ClsCode").Specific.VALUE = "" Then
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
        If pval.ItemUID = "Group" Then
            Call oForm01.Items("SGroup").Specific.Select(oForm01.Items("Group").Specific.VALUE, psk_ByValue)
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
                Call oMat01.SelectRow(pval.Row, True, False)
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
        If pval.ItemUID = "Mat01" And pval.Row = "0" And pval.ColUID = "CHK" Then
            Dim Chk$
            Dim i As Long
            oForm01.Freeze True
            oMat01.FlushToDataSource
            If Trim(oDS_PS_HR412L.GetValue("U_ColReg01", 0)) = "" Or Trim(oDS_PS_HR412L.GetValue("U_ColReg01", 0)) = "N" Then
                Chk = "Y"
            ElseIf Trim(oDS_PS_HR412L.GetValue("U_ColReg01", 0)) = "Y" Then
                Chk = "N"
            End If
            For i = 0 To oMat01.VisualRowCount - 1
                oDS_PS_HR412L.setValue "U_ColReg01", i, Chk
            Next i
            oMat01.LoadFromDataSource
            oForm01.Freeze False
        End If
    End If
    Exit Sub
Raise_EVENT_DOUBLE_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_DOUBLE_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LINK_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LINK_PRESSED_Error
Dim TempForm01 As Variant
    If pval.BeforeAction = True Then
        If oMat01.Columns("DocEntry").Cells(pval.Row).Specific.String <> "" Then
'            Set TempForm01 = New PS_PP030
'            Call TempForm01.LoadForm(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.String)
'            BubbleEvent = False
        Else
            
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
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
            If (pval.ItemUID = "Year") Then
                If Trim(oForm01.Items("Year").Specific.VALUE) <> "" Then
                    
                    Query01 = "select U_Number from [@PS_HR410H] a"
                    Query01 = Query01 + " Where Isnull(a.U_OpenYN,'N') = 'Y' and isnull(a.U_CloseYN,'N') = 'N' "
                    Query01 = Query01 + " and a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "' "
                    Query01 = Query01 + " and a.U_Year = '" & Trim(oForm01.Items("Year").Specific.VALUE) & "' "
                    RecordSet01.DoQuery Query01
                    
                    If Trim(RecordSet01.Fields(0).VALUE) <> "" Then
                        Call oForm01.Items("Number").Specific.Select(Trim(RecordSet01.Fields(0).VALUE), psk_ByValue)
                    End If
                End If
            ElseIf (pval.ItemUID = "MSTCOD") Then
                Query01 = "Select FULLNAME = t.U_FULLNAME, "
                Query01 = Query01 + " JigNm = Isnull((Select U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b on a.Code = b.Code  "
                Query01 = Query01 + " Where a.Name = '직급코드' And b.U_Code = t.U_JIGCOD ),''), "
                Query01 = Query01 + " CallName = Isnull((Select U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b on a.Code = b.Code  "
                Query01 = Query01 + " Where a.Name = '전문직호칭' And b.U_Code = t.U_CallName ),''), "
                Query01 = Query01 + " JicNm = Isnull((SELECT name FROM [OHPS]  "
                Query01 = Query01 + " Where posID = t.U_position ),'') "
                Query01 = Query01 + " From [@PH_PY001A] t Where Code =  '" & Trim(oForm01.Items("MSTCOD").Specific.VALUE) & "' "
                'Query01 = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("MSTCOD").Specific.VALUE) & "'"
                RecordSet01.DoQuery Query01
                
                oForm01.Items("FULLNAME").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                oForm01.Items("JigNm").Specific.VALUE = Trim(RecordSet01.Fields(1).VALUE)
                oForm01.Items("CallName").Specific.VALUE = Trim(RecordSet01.Fields(2).VALUE)
                oForm01.Items("JicNm").Specific.VALUE = Trim(RecordSet01.Fields(3).VALUE)
                
                    '//기타작업
'                    Call oDS_PS_HR412L.setValue("U_" & "ColReg02", pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_HR412L.GetValue("U_" & "ColReg02", pval.Row - 1)) <> "" Then
'                        PS_HR412_AddMatrixRow (pval.Row)
'                    End If
                
            
            ElseIf (pval.ItemUID = "TeamCode") Then
                Query01 = "Select b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code Where a.Name = '부서' and b.U_Code = '" & oForm01.Items("TeamCode").Specific.VALUE & "'"
                RecordSet01.DoQuery Query01
                
                oForm01.Items("TeamNm").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                
                
            ElseIf (pval.ItemUID = "RspCode") Then
                Query01 = "Select b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code Where a.Name = '담당' and b.U_Code = '" & oForm01.Items("RspCode").Specific.VALUE & "'"
                RecordSet01.DoQuery Query01
                
                oForm01.Items("RspNm").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
            ElseIf (pval.ItemUID = "ClsCode") Then
                Query01 = "Select b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code Where a.Name = '반' and b.U_Code = '" & oForm01.Items("ClsCode").Specific.VALUE & "'"
                RecordSet01.DoQuery Query01
                
                oForm01.Items("ClsNm").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
            Set RecordSet01 = Nothing
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
        Call PS_HR412_FormItemEnabled
        '//Call PS_HR412_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then

    ElseIf pval.BeforeAction = False Then
'        oForm01.Items("Mat01").Top = 50
'        oForm01.Items("Mat01").Height = (oForm01.Height / 2) - 50
'        oForm01.Items("Mat01").Left = 7
'        oForm01.Items("Mat01").Width = oForm01.Width - 21
'        oForm01.Items("Mat02").Top = (oForm01.Height / 2) + 10
'        oForm01.Items("Mat02").Height = (oForm01.Height / 2) - 75
'        oForm01.Items("Mat02").Left = 7
'        oForm01.Items("Mat02").Width = oForm01.Width - 21
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
                oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns("ItemCode").Cells(0).VALUE
                oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns("ItemName").Cells(0).VALUE
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
    '        Call oDS_PS_HR412L.RemoveRecord(oDS_PS_HR412L.Size - 1)
    '        oMat01.LoadFromDataSource
    '        If oMat01.RowCount = 0 Then
    '            Call PS_HR412_AddMatrixRow(0)
    '        Else
    '            If Trim(oDS_PP035L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
    '                Call PS_HR412_AddMatrixRow(oMat01.RowCount)
    '            End If
    '        End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_HR412_CreateItems() As Boolean
On Error GoTo PS_HR412_CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'   Set oDS_PS_HR412H = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_HR412L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oMat01 = oForm01.Items("Mat01").Specific

    oMat01.SelectionMode = ms_NotSupported

    oMat01.AutoResizeColumns

    '//사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 1)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    '//년도
    Call oForm01.DataSources.UserDataSources.Add("Year", dt_SHORT_TEXT, 4)
    Call oForm01.Items("Year").Specific.DataBind.SetBound(True, "", "Year")
    
    '//차수
    Call oForm01.DataSources.UserDataSources.Add("Number", dt_SHORT_TEXT, 1)
    Call oForm01.Items("Number").Specific.DataBind.SetBound(True, "", "Number")
    
    '//평가권한
    Call oForm01.DataSources.UserDataSources.Add("Evaluate", dt_SHORT_TEXT, 1)
    Call oForm01.Items("Evaluate").Specific.DataBind.SetBound(True, "", "Evaluate")
    
    '//사번
    Call oForm01.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    Call oForm01.Items("MSTCOD").Specific.DataBind.SetBound(True, "", "MSTCOD")
    '//성명
    Call oForm01.DataSources.UserDataSources.Add("FULLNAME", dt_SHORT_TEXT, 20)
    Call oForm01.Items("FULLNAME").Specific.DataBind.SetBound(True, "", "FULLNAME")
    '//직급명
    Call oForm01.DataSources.UserDataSources.Add("JigNm", dt_SHORT_TEXT, 20)
    Call oForm01.Items("JigNm").Specific.DataBind.SetBound(True, "", "JigNm")
    
'    Call oForm01.DataSources.UserDataSources.Add("C", dt_DATE, 8)
'    Call oForm01.Items("WorkDtFr").Specific.DataBind.SetBound(True, "", "WorkDtFr")
'
'    Call oForm01.DataSources.UserDataSources.Add("WorkDtTo", dt_DATE, 8)
'    Call oForm01.Items("WorkDtTo").Specific.DataBind.SetBound(True, "", "WorkDtTo")
    '//호칭
    Call oForm01.DataSources.UserDataSources.Add("CallName", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CallName").Specific.DataBind.SetBound(True, "", "CallName")
    '//직책
    Call oForm01.DataSources.UserDataSources.Add("JicNm", dt_SHORT_TEXT, 20)
    Call oForm01.Items("JicNm").Specific.DataBind.SetBound(True, "", "JicNm")
    
    '팀코드
    Call oForm01.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("TeamCode").Specific.DataBind.SetBound(True, "", "TeamCode")
    
    '팀명
    Call oForm01.DataSources.UserDataSources.Add("TeamNm", dt_SHORT_TEXT, 20)
    Call oForm01.Items("TeamNm").Specific.DataBind.SetBound(True, "", "TeamNm")
    
    
    '담당코드
    Call oForm01.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("RspCode").Specific.DataBind.SetBound(True, "", "RspCode")
    
    '담당명
    Call oForm01.DataSources.UserDataSources.Add("RspNm", dt_SHORT_TEXT, 20)
    Call oForm01.Items("RspNm").Specific.DataBind.SetBound(True, "", "RspNm")
    
    
    '반코드
    Call oForm01.DataSources.UserDataSources.Add("ClsCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ClsCode").Specific.DataBind.SetBound(True, "", "ClsCode")
    
    '반명
    Call oForm01.DataSources.UserDataSources.Add("ClsNm", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ClsNm").Specific.DataBind.SetBound(True, "", "ClsNm")
    
    '평가대상자그룹
    Call oForm01.DataSources.UserDataSources.Add("Group", dt_SHORT_TEXT, 1)
    Call oForm01.Items("Group").Specific.DataBind.SetBound(True, "", "Group")
    
    '평가대상자그룹(조회용)
    Call oForm01.DataSources.UserDataSources.Add("SGroup", dt_SHORT_TEXT, 1)
    Call oForm01.Items("SGroup").Specific.DataBind.SetBound(True, "", "SGroup")
    
    oForm01.Items("Mat01").Enabled = True
'    oForm01.Items("Mat02").Enabled = False
        
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
PS_HR412_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR412_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_HR412_ComboBox_Setting()
On Error GoTo PS_HR412_ComboBox_Setting_Error

    Dim i As Long
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Call oForm01.Freeze(True)
    '//콤보에 기본값설정
    ' 사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)

    ' 평가차수
    Set oCombo = oForm01.Items("Number").Specific
    oCombo.ValidValues.Add "1", "1차"
    oCombo.ValidValues.Add "2", "2차"
    
    '평가권한
    Set oCombo = oForm01.Items("Evaluate").Specific
    oCombo.ValidValues.Add "1", "1차평가자"
    oCombo.ValidValues.Add "2", "2차평가자"
    oCombo.ValidValues.Add "3", "종합평가자"
    
    '평가대상자그룹
    Set oCombo = oForm01.Items("Group").Specific
    oCombo.ValidValues.Add "1", "반장"
    oCombo.ValidValues.Add "2", "사원"
    
    '평가대상자그룹(조회용)
    Set oCombo = oForm01.Items("SGroup").Specific
    oCombo.ValidValues.Add "1", "반장"
    oCombo.ValidValues.Add "2", "사원"
    
    '//Call MDC_PS_Common.Combo_ValidValues_Insert("PS_HR412", "oMat01", "OrdGbn", "", "")
    
    
    
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR412_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR412_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR412_Initial_Setting()
On Error GoTo PS_HR412_Initial_Setting_Error
 ' 사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 ' 인수자
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD()
    Exit Sub
PS_HR412_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR412_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR412_CF_ChooseFromList()
On Error GoTo PS_HR412_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
    Dim oCons               As SAPbouiCOM.Conditions
    Dim oCon                As SAPbouiCOM.Condition
    Dim oCFL                As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
    Dim oEdit               As SAPbouiCOM.EditText
    Dim oColumn             As SAPbouiCOM.Column

    Set oEdit = oForm01.Items("MSTCOD").Specific
    Set oCFLs = oForm01.ChooseFromLists
    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

    oCFLCreationParams.ObjectType = "4"
    oCFLCreationParams.uniqueID = "CFLITEMCD"
    oCFLCreationParams.MultiSelection = False
    Set oCFL = oCFLs.Add(oCFLCreationParams)

'    Set oCons = oCFL.GetConditions()
'    Set oCon = oCons.Add()
'    oCon.Alias = "CardType"
'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
'    oCon.CondVal = "C"
'    oCFL.SetConditions oCons

    oEdit.ChooseFromListUID = "CFLITEMCD"
    oEdit.ChooseFromListAlias = "MSTCOD"
    
    Exit Sub
PS_HR412_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR412_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR412_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_HR412_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_HR412_FormItemEnabled
        Call PS_HR412_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_HR412_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_HR412_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR412_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR412_FormItemEnabled()
On Error GoTo PS_HR412_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        '//Call PS_HR412_FormClear '//UDO방식
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR412_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR412_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR412_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_HR412_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_HR412L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_HR412L.Offset = oRow
    oDS_PS_HR412L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR412_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR412_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_HR412_FormClear()
On Error GoTo PS_HR412_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_HR412'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_HR412_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR412_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR412_EnableMenus()
On Error GoTo PS_HR412_EnableMenus_Error
    '//메뉴활성화
    Call oForm01.EnableMenu("1281", False)
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_HR412_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR412_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_HR412_FormResize()
On Error GoTo PS_HR412_FormResize_Error
    
    Exit Sub
PS_HR412_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR412_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_HR412_DataValidCheck() As Boolean
On Error GoTo PS_HR412_DataValidCheck_Error
    Dim i As Long
'    If oForm01.Items("WhsCode").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "창고는 필수입니다.", bmt_Short, True
'        oForm01.Items("WhsCode").Click ct_Regular
'        PS_HR412_DataValidCheck = False
'        Exit Function
'    End If
'    If oMat01.VisualRowCount = 0 Then
'        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
'        PS_HR412_DataValidCheck = False
'        Exit Function
'    End If
'    For i = 1 To oMat01.VisualRowCount
'        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
'            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
'            PS_HR412_DataValidCheck = False
'            Exit Function
'        End If
'    Next
'    Call oDS_PP035L.RemoveRecord(oDS_PP035L.Size - 1)
'    Call oMat01.LoadFromDataSource
    Call PS_HR412_FormClear
    Exit Function
PS_HR412_DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR412_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_HR412_MTX01()
On Error GoTo PS_HR412_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim Param08 As String
    Dim Param09 As String
    
    Param01 = Trim(oForm01.Items("BPLId").Specific.VALUE) '//사업장
    Param02 = Trim(oForm01.Items("Year").Specific.VALUE)  '//년도
    Param03 = Trim(oForm01.Items("Number").Specific.VALUE) '//평가차수 1차, 2차평가
    Param04 = Trim(oForm01.Items("Evaluate").Specific.VALUE) '//평가권한 1차,2차,3차권한
    Param05 = Trim(oForm01.Items("TeamCode").Specific.VALUE) '//팀코드
    Param06 = Trim(oForm01.Items("RspCode").Specific.VALUE) '//담당코드
    Param07 = Trim(oForm01.Items("ClsCode").Specific.VALUE) '//반코드
    Param08 = Trim(oForm01.Items("SGroup").Specific.VALUE) '//그룹 1:반장, 2:사원
    Param09 = Trim(oForm01.Items("MSTCOD").Specific.VALUE) '//평가자사번
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_HR412_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "', '" & Param08 & "', '" & Param09 & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        oForm01.Items("Mat01").Enabled = False
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
        GoTo PS_HR412_MTX01_Exit
    Else
'        Call MDC_Com.MDC_GF_Message("조회를 성공하였습니다.", "S")
        oForm01.Items("Mat01").Enabled = True
    End If
    
    
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_HR412L.InsertRecord (i)
        End If
        oDS_PS_HR412L.Offset = i
        Call oDS_PS_HR412L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_HR412L.setValue("U_ColReg01", i, False)
        Call oDS_PS_HR412L.setValue("U_ColReg02", i, RecordSet01.Fields("MSTCOD").VALUE) '//사번
        Call oDS_PS_HR412L.setValue("U_ColReg03", i, RecordSet01.Fields("FULLNAME").VALUE) '//성명
        Call oDS_PS_HR412L.setValue("U_ColReg04", i, RecordSet01.Fields("TeamNm").VALUE) '//팀명
        Call oDS_PS_HR412L.setValue("U_ColReg05", i, RecordSet01.Fields("RspNm").VALUE) '//담당명
        Call oDS_PS_HR412L.setValue("U_ColReg06", i, RecordSet01.Fields("JigNm").VALUE) '//직급
        Call oDS_PS_HR412L.setValue("U_ColReg07", i, RecordSet01.Fields("CallName").VALUE) '//호칭
        Call oDS_PS_HR412L.setValue("U_ColReg08", i, RecordSet01.Fields("JicNm").VALUE) '//직책
        Call oDS_PS_HR412L.setValue("U_ColReg15", i, RecordSet01.Fields("PeakYN").VALUE) '//피크사원여
        Call oDS_PS_HR412L.setValue("U_ColReg09", i, RecordSet01.Fields("MSTCOD1").VALUE) '//1차평가자사번
        Call oDS_PS_HR412L.setValue("U_ColReg10", i, RecordSet01.Fields("FULLNAM1").VALUE) '//1차평가자성명
        Call oDS_PS_HR412L.setValue("U_ColReg11", i, RecordSet01.Fields("MSTCOD2").VALUE) '//2차평가자사번
        Call oDS_PS_HR412L.setValue("U_ColReg12", i, RecordSet01.Fields("FULLNAM2").VALUE) '//2차평가자성명
        Call oDS_PS_HR412L.setValue("U_ColReg13", i, RecordSet01.Fields("MSTCOD3").VALUE)  '//종합평가자사번
        Call oDS_PS_HR412L.setValue("U_ColReg14", i, RecordSet01.Fields("FULLNAM3").VALUE) '//종합평가자성명
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    oForm01.Update
    
    Call MDC_Com.MDC_GF_Message("조회를 성공하였습니다.", "S")
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR412_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR412_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR412_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
'//update
Private Sub PS_HR412_SetBaseForm(ByVal oStatus As String)
On Error GoTo PS_HR412_SetBaseForm_Error

    Dim i As Long
    Dim ItemCode01 As String

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As Double
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim Param08 As String
    Dim Param09 As String
    
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    For i = 1 To oMat01.RowCount '//품목선택품목
    
        
        Dim QueryString As String

        If oMat01.Columns("CHK").Cells(i).Specific.Checked = True Then
            
            Param01 = oMat01.Columns("MSTCOD").Cells(i).Specific.VALUE
            Param02 = oForm01.Items("BPLId").Specific.VALUE '//사업장
            Param03 = oForm01.Items("Year").Specific.VALUE '//평가년도
            Param04 = oForm01.Items("Number").Specific.VALUE '//평가차수
            Param05 = oForm01.Items("Evaluate").Specific.VALUE '//평가권한
            Param06 = oForm01.Items("Group").Specific.VALUE '//반장, 사원구분
            Param07 = oForm01.Items("MSTCOD").Specific.VALUE '//평가자사번
            Param08 = oForm01.Items("FULLNAME").Specific.VALUE '//평가자성명
            
            If oForm01.Items("Evaluate").Specific.VALUE = "1" And oForm01.Items("Group").Specific.VALUE = "2" Then  '//1차평가자 이고 반장
                Param09 = oForm01.Items("CallName").Specific.VALUE '//직책명
            Else
                Param09 = oForm01.Items("JigNm").Specific.VALUE '//직급명
            End If
            
            Query01 = "EXEC PS_HR412_02 '" & oStatus & "', '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "', '" & Param08 & "', '" & Param09 & "'"
            
            Call RecordSet01.DoQuery(Query01)
            Call MDC_Com.MDC_GF_Message("데이터를 수정하였습니다.", "S")
            

        End If
        
    Next
        
    Exit Sub
    
PS_HR412_SetBaseForm_Error:
    Sbo_Application.SetStatusBarMessage "PS_HR412_SetBaseForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'// 패스워드 Check
Public Function PasswordChk(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo PasswordChk_Error
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim MSTCOD$, PassWd$
        
    MSTCOD = Trim(oForm01.Items("MSTCOD").Specific.VALUE)
    PassWd = Trim(oForm01.Items("PassWd").Specific.VALUE)
    
    If Trim(MSTCOD) = "" Then
        MDC_Com.MDC_GF_Message "사번이 없습니다. 입력바랍니다!", "E"
        PasswordChk = False
        Exit Function
    End If
        
    sQry = "Select Count(*) From Z_PS_HRPASS Where MSTCOD = '" & Trim(oForm01.Items("MSTCOD").Specific.VALUE) & "'"
    sQry = sQry & " And  BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
    sQry = sQry & " And  PassWd = '" & oForm01.Items("PassWd").Specific.VALUE & "' "
    RecordSet01.DoQuery sQry
    
    If Trim(RecordSet01.Fields(0).VALUE) <= 0 Then
        PasswordChk = False
    Else
        PasswordChk = True
    End If
        
    Set RecordSet01 = Nothing
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PasswordChk_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PasswordChk_Error:" & Err.Number & " - " & Err.Description, "E"
End Function
